$listOfClientFolders = @("_NetSuite automatically creates Opportunity & Project folders","Background","Non-specific BusDev")

#Get Terms from Managed Metadata Store
$sharePointAdmin = "kimblebot@anthesisgroup.com"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\KimbleBot.txt") 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com" -Credentials $adminCreds
$pnpTermGroup = "Kimble"
$pnpTermSet = "Clients"
$allClientTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet

#Filter these client-side (CSOM, eh?) to get only the changes since this script last completed successfully
[datetime]$lastProcessed = $(Get-PnPTerm -TermGroup "Anthesis" -TermSet "IT" -Identity "LastModified" -Includes CustomProperties).CustomProperties.ClientSiteDriveCreation
$clientTermsToCheck = $allClientTerms | ? {$_.LastModifiedDate -gt $lastProcessed -and ![string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteId)}
$clientTermsToCheck | % {
    Add-Member -InputObject $_ -MemberType NoteProperty -Name Name2 -Value $($_.Name).Replace("&","").Replace("＆","").Replace("  "," ") -Force
    Add-Member -InputObject $_ -MemberType NoteProperty -Name DriveId -Value $($_.CustomProperties.GraphDriveId) -Force
    }

#Get the Drives from Graph to compare against
$sharePointBotDetails = get-graphAppClientCredentials -appName SharePointBot
$tokenResponseSharePointBot = get-graphTokenResponse -aadAppCreds $sharePointBotDetails
$clientSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,faed84bc-70be-4e35-bfbf-cdab31aeeb99"
$supplierSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,9fb8ecd6-c87d-485d-a488-26fd18c62303"
$devSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,8ba7475f-dad0-4d16-bdf5-4f8787838809"
$allClientDrives = get-graphDrives -tokenResponse $tokenResponseSharePointBot -siteGraphId $clientSiteId
$allClientDrives | % {
    Add-Member -InputObject $_ -MemberType NoteProperty -Name Name2 -Value $($_.name).Replace("&","").Replace("＆","").Replace("  "," ") -Force
    Add-Member -InputObject $_ -MemberType NoteProperty -Name DriveId -Value $($_.id) -Force
    }

#############################
#Update any Clients Drives that have changed their names in NetSuite
#############################
$matchedGraphId = Compare-Object -ReferenceObject $clientTermsToCheck -DifferenceObject $allClientDrives -Property DriveId -IncludeEqual -ExcludeDifferent -PassThru #We find out which $clientTermsToCheck records already have valid GraphDriveId values
$matchedGraphIdReversed = Compare-Object -ReferenceObject $allClientDrives -DifferenceObject $matchedGraphId -Property DriveId -IncludeEqual -ExcludeDifferent -PassThru #We then use $matchedGraphId to filter only the Drive objects with corresponding $clientTermsToCheck records
$deltaName = Compare-Object -ReferenceObject $matchedGraphId -DifferenceObject $matchedGraphIdReversed -Property DriveId,Name2 -PassThru -IncludeEqual #We compare the two equal sets on both DriveId and Name2 to see which pairs have mismatched Name values
$clientsWithChangedNames = $deltaName | ? {$_.SideIndicator -eq "<="} #Anything on this side has a different Name2 in NetSuite
$clientsWithChangedNames | % {
    $thisUpdatedClient = $_
    Write-Verbose "Company name [$($thisUpdatedClient.Name)][$($thisUpdatedClient.id)] seems to have changed. Investigating further."
    if([string]::IsNullOrWhiteSpace($thisUpdatedClient.CustomProperties.DocLibId)){ #If it's missing it's DocLibID, try to fix it
        Write-Verbose "[$($thisUpdatedClient.Name)][$($thisUpdatedClient.id)] is missing its .CustomProperties.DocLibId value - attempting repair"
        $graphList = get-graphList -tokenResponse $tokenResponseSharePointBot -graphSiteId $clientSiteId -graphDriveId $thisUpdatedClient.CustomProperties.GraphDriveId
        if($graphList){
            $thisUpdatedClient.SetCustomProperty("DocLibId",$graphList.id)
            try{
                Write-Verbose "`tTrying: [$($thisUpdatedClient.Name)].SetCustomProperty(`"DocLibId`",$($graphList.id))"
                $thisUpdatedClient.Context.ExecuteQuery()
                $thisUpdatedClient = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Identity $thisUpdatedClient.Id -Includes CustomProperties
                }
            catch{
                Write-Error "Error setting [$($thisUpdatedClient.Name)].SetCustomProperty(`"DocLibId`",$($graphList.id)) on Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisUpdatedClient.Name)] in sync-netsuiteManagedMetaDataToSharePoint()"
                }
            }
        }
    if($thisUpdatedClient.CustomProperties.DocLibId){
        Write-Verbose "[$($thisUpdatedClient.Name)][$($thisUpdatedClient.id)] has .CustomProperties.DocLibId value [$($thisUpdatedClient.CustomProperties.DocLibId)] - attempting to rename List to match Term name"
        try{
            Write-Verbose "`tTrying: set-graphList -graphSiteId $clientSiteId -graphListId $($thisUpdatedClient.CustomProperties.DocLibId) -listPropertyHash @{displayName=$($thisUpdatedClient.Name)}"
            $updatedGraphList = set-graphList -tokenResponse $tokenResponseSharePointBot -graphSiteId $clientSiteId -graphListId $thisUpdatedClient.CustomProperties.DocLibId -listPropertyHash @{displayName=$thisUpdatedClient.Name}
            }
        catch{
            Write-Error "Error setting List [$($clientSiteId)][$($thisUpdatedClient.CustomProperties.DocLibId)] DisplayName to [$($thisUpdatedClient.Name)] in sync-netsuiteManagedMetaDataToSharePoint()"
            }
        
        if($updatedGraphList.displayName -ne $thisUpdatedClient.Name){#If this didn;t work, it might be because of a DisplayName collision, but it'll need investigating by a human for now as no errors are returned for us to handle
            Write-Warning "Failed to update List [$($updatedGraphList.displayName)][$($updatedGraphList.CustomProperties.DocLibId)][$($updatedGraphList.id)] DisplayName to [$($thisUpdatedClient.Name)] in sync-netsuiteManagedMetaDataToSharePoint()"
            [array]$problemChilds += ,@($thisUpdatedClient,$updatedGraphList,"Failed to update List DisplayName via graph") 
            }
        else{
            Write-Verbose "`tSuccess! List renamed to [$($thisUpdatedClient.Name)]"
            }
        }
    else{
        Write-Error "Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisUpdatedClient.Name)] does not have a .CustomProperties.DocLibId value, and one could not be determined - cannot update DisplayName."
        [array]$problemChilds += ,@($thisUpdatedClient,$updatedGraphList,"Failed to update List DisplayName via graph") 
        }
    }


#############################
#Create new Prospects/Clients
#############################
#We haven't changed any IDs or created any new Drives, so no need to refresh
$missingFromSpo = Compare-Object -ReferenceObject $clientTermsToCheck -DifferenceObject $allClientDrives -Property DriveId -PassThru | ? {$_.SideIndicator -eq "<="}
$missingFromSpo | % {
    $thisNewClient = $_
    Write-Verbose "Creating new GraphList [$($thisNewClient.Name)]"
    $newGraphList = new-graphList -tokenResponse $tokenResponseSharePointBot -siteGraphId $clientSiteId -listDisplayName $thisNewClient.Name -listType documentLibrary #Graph doesn't support creating Drives, so we need to create a List
    Write-Verbose "Getting new GraphDrive from GraphList [$($newGraphList.name)][$($newGraphList.id)]"
    $newGraphListDrive = get-graphDrives -tokenResponse $tokenResponseSharePointBot -siteGraphId $clientSiteId -listGraphId $newGraphList.id #Then get the new Drive object form the List
    Write-Verbose "Creating standard Client folders in Drive [$($newGraphListDrive.name)][$($newGraphListDrive.id)]"
    $newFolders = add-graphArrayOfFoldersToDrive -graphDriveId $newGraphListDrive.id -foldersAndSubfoldersArray $listOfClientFolders -tokenResponse $tokenResponseSharePointBot -conflictResolution Replace
    Write-Verbose "Updating Term [$($thisNewClient.Name)] with CustomProperties @{DocLibId=$($newGraphList.id);GraphDriveId=$($newGraphListDrive.id)}"
    $thisNewClient.SetCustomProperty("DocLibId",$newGraphList.id)
    $thisNewClient.SetCustomProperty("GraphDriveId",$newGraphListDrive.id)
    try{
        Write-Verbose "`tTrying to update Term [$($thisNewClient.Name)] with CustomProperties @{DocLibId=$($newGraphList.id);GraphDriveId=$($newGraphListDrive.id)}"
        $thisNewClient.Context.ExecuteQuery()
        }
    catch{
        Write-Error "Error updating Term [$($thisNewClient.Name)] with CustomProperties @{DocLibId=$($newGraphList.id);GraphDriveId=$($newGraphListDrive.id)} in sync-netsuiteManagedMetaDataToSharePoint()"
        }
    }
