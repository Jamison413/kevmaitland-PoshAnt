function get-clientDrives(){
    #Get the Drives from Graph to compare against
    $sharePointBotDetails = get-graphAppClientCredentials -appName SharePointBot
    $tokenResponseSharePointBot = get-graphTokenResponse -aadAppCreds $sharePointBotDetails
    $clientSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,faed84bc-70be-4e35-bfbf-cdab31aeeb99"
    #$supplierSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,9fb8ecd6-c87d-485d-a488-26fd18c62303"
    #$devSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,8ba7475f-dad0-4d16-bdf5-4f8787838809"
    $allClientDrives = get-graphDrives -tokenResponse $tokenResponseSharePointBot -siteGraphId $clientSiteId
    $allClientDrives | % {
        Add-Member -InputObject $_ -MemberType NoteProperty -Name Name2 -Value $($_.name).Replace("&","").Replace("＆","").Replace("  "," ") -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name DriveId -Value $($_.id) -Force
        }
    $allClientDrives
    }
if($PSCommandPath){
    $InformationPreference = 2
    $VerbosePreference = 0
    $logFileLocation = "C:\ScriptLogs\"
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))`_Transcript_$(Get-Date -Format "yyyy-MM-dd").log"
    Start-Transcript $transcriptLogName -Append
    }

$listOfClientFolders = @("_NetSuite automatically creates Opportunity & Project folders","Background","Non-specific BusDev")
$listOfLeadProjSubFolders = @("Admin & contracts", "Analysis","Data & refs","Meetings","Proposal","Reports","Summary (marketing) - end of project")
$now = Get-Date

$sharePointBotDetails = get-graphAppClientCredentials -appName SharePointBot
$tokenResponseSharePointBot = get-graphTokenResponse -aadAppCreds $sharePointBotDetails
$clientSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,faed84bc-70be-4e35-bfbf-cdab31aeeb99"
#$supplierSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,9fb8ecd6-c87d-485d-a488-26fd18c62303"
#$devSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,8ba7475f-dad0-4d16-bdf5-4f8787838809"

$sharePointAdmin = "kimblebot@anthesisgroup.com"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\KimbleBot.txt") 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com" -Credentials $adminCreds

Write-Information "sync-netsuiteManagedMetaDataToSharePoint started at [$(Get-Date -Format s)]"
$fullSyncTime = Measure-Command {
    [datetime]$lastSpoSyncRun = $(Get-PnPTerm -TermGroup "Anthesis" -TermSet "IT" -Identity "LastModified" -Includes CustomProperties).CustomProperties.LastSpoSyncRun

    #region Clients
    $pnpTermGroup = "Kimble"
    $pnpTermSet = "Clients"
    $allClientTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes CustomProperties | ? {$_.IsDeprecated -eq $false}

    #Filter these client-side (CSOM, eh?) to get only the changes since this script last completed successfully
    [array]$clientTermsToCheck = $allClientTerms | ? {($_.LastModifiedDate -gt $lastSpoSyncRun -or $_.CustomProperties.flagForReprocessing -eq $true) -and ![string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteId)}
    Write-Information "Processing [$($clientTermsToCheck.Count)] Clients"
    if($clientTermsToCheck){
        $clientTermsToCheck | % {
            Add-Member -InputObject $_ -MemberType NoteProperty -Name Name2 -Value $($_.Name).Replace("&","").Replace("＆","").Replace("  "," ") -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteId -Value $($_.CustomProperties.GraphDriveId) -Force
            }
        if(!$allClientDrives){$allClientDrives = get-clientDrives} #Only load the Client Drives via graph if there is work to do and we haven't already got them

        #############################
        #Create new Prospects/Clients
        #############################
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
                [array]$flagForReprocessing += $thisNewClient
                }
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
        }

    #endregion

    #region Opportunities
    $pnpTermGroup = "Kimble"
    $pnpTermSet = "Opportunities"
    $allOppTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes CustomProperties | ? {$_.IsDeprecated -eq $false}
    $allOppTerms | % {
        Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteId -Value $_.CustomProperties.NetSuiteOppId -Force
        }
    #Filter these client-side (CSOM, eh?) to get only the changes since this script last completed successfully
    [array]$oppTermsToCheck = $allOppTerms | ? {($_.LastModifiedDate -gt $lastSpoSyncRun -or $_.CustomProperties.flagForReprocessing -eq $true) -and ![string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteOppId) -and [string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteProjId)}
    Write-Information "Processing [$($oppTermsToCheck.Count)] Opportunities"
    if($oppTermsToCheck){
        $oppTermsToCheck | % {
            Add-Member -InputObject $_ -MemberType NoteProperty -Name DriveId -Value $($_.CustomProperties.GraphDriveId) -Force
            }
        if(!$allClientDrives){$allClientDrives = get-clientDrives} #Only load the Client Drives via graph if there is work to do and we haven't already got them

        #############################
        #Process Opps en-masse as don't want to query every DriveItem in every Drive in advance
        #############################
        #Check for DriveItemId 
            #If DriveItemId, update folder (Opps with NetSuiteProjId are already filtered out above, so once an Opp is converted to a Project, the Opp Term no longer controls the folder)
            #If no DriveItemId, Create new folders

        [array]$flagForReprocessing = @()
        $oppTermsToCheck | % {
            $thisOppTerm = $_
            Write-Information "Checking Opp Term [$($thisOppTerm.Name)] for CustomProperties.DriveItemId"
            if($thisOppTerm.CustomProperties.DriveItemId){
                $thisClientTerm = $allClientTerms | ? {$_.CustomProperties.NetSuiteId -eq $thisOppTerm.CustomProperties.NetSuiteClientId}
                Write-Information "`tOpp Term [$($thisOppTerm.Name)].CustomProperties.DriveItemId is [$($thisOppTerm.CustomProperties.DriveItemId)] - setting Folder name to [$($thisOppTerm.Name)]"
                try{
                    $updatedFolder = set-graphDriveItem -tokenResponse $tokenResponseSharePointBot -driveId $thisClientTerm.CustomProperties.GraphDriveId -driveItemId $thisOppTerm.CustomProperties.DriveItemId -driveItemPropertyHash @{name=$thisOppTerm.Name}
                    }
                catch{
                    $_
                    #There's a whole lot that could wrong here ($thisClientTerm.DriveId could be $null, the Drive could be missing or the DriveItem could be missing. Log the error for further investigation
                    }
                }
            else{#If no Opp, Create new folders
                Write-Information "`tNo DriveItemId [$($thisOppTerm.CustomProperties.DriveItemId)] found - creating new set of Project folders"
                $thisClientTerm = $allClientTerms | ? {$_.CustomProperties.NetSuiteId -eq $thisOppTerm.CustomProperties.NetSuiteClientId}
                [array]$customisedFolderList = $thisOppTerm.Name
                $customisedFolderList += $listOfLeadProjSubFolders | % {"$($thisOppTerm.Name)\$_"}
                $newOppFolders = add-graphArrayOfFoldersToDrive -graphDriveId $thisClientTerm.CustomProperties.GraphDriveId -foldersAndSubfoldersArray $customisedFolderList -tokenResponse $tokenResponseSharePointBot -conflictResolution Fail #-ErrorAction SilentlyContinue
                $thisOppTerm.SetCustomProperty("DriveItemId",$newOppFolders[1].id)
                try{
                    Write-Verbose "`tTrying: [$($thisOppTerm.Name)].SetCustomProperty(DriveItemId,[$($updatedFolder.id)])"
                    $thisOppTerm.Context.ExecuteQuery()
                    }
                catch{
                    Write-Error "Error updating DriveItemId CustomProperty [$($thisOppTerm.CustomProperties.DriveItemId))] to [$($updatedFolder.id)] for Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisOppTerm.Name)] in sync-netsuiteToManagedMetaData()"
                    [array]$flagForReprocessing += $thisNewClient
                    }
                }
            }

        #############################
        #Update LastModifiedDate
        #############################
        $oppTermsToCheck | % {
            $thisOppToUpdate = $_
            if($flagForReprocessing -notcontains $thisOppToUpdate){ #If the process above worked as expected, update SharePointLastModifiedDate to prevent it from being re-processed next time
                Write-Information "[$($thisOppToUpdate.Name)] was processed successfully - updating SharePointLastModifiedDate to [$($now)]"
                $thisOppToUpdate.SetCustomProperty("SharePointLastModifiedDate",$now)
                }
            else{
                Write-Information "Something went wrong with [$($thisOppToUpdate.Name)] - flagging for reprocessing"
                $thisOppToUpdate.SetCustomProperty("flagForReprocessing",$true)
                }
            try{
                Write-Verbose "`tTrying: [$($thisOppToUpdate.Name)][$($thisOppToUpdate.Id)].SetCustomProperty(SharePointLastModifiedDate,$($now))"
                $thisOppToUpdate.Context.ExecuteQuery()
                }
            catch{
                Write-Error "Error setting CustomProperty SharePointLastModifiedDate on Term [$($updatedTerm.Name)][$($updatedTerm.Id)] in sync-netsuiteToManagedMetaData()"
                $_
                }
            }
        }
    #endregion

    #region Projects
    $pnpTermGroup = "Kimble"
    $pnpTermSet = "Projects"
    $allProjTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes CustomProperties | ? {$_.IsDeprecated -eq $false}
    $allProjTerms | % {
        Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteId -Value $_.CustomProperties.NetSuiteProjId -Force
        }
    #Filter these client-side (CSOM, eh?) to get only the changes since this script last completed successfully
    [array]$projTermsToCheck = $allProjTerms | ? {($_.LastModifiedDate -gt $lastSpoSyncRun -or $_.CustomProperties.flagForReprocessing -eq $true) -and ![string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteProjId) -and ![string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteClientId)}
    Write-Information "Processing [$($projTermsToCheck.Count)] Projects"
    if($projTermsToCheck){
        $projTermsToCheck | % {
            Add-Member -InputObject $_ -MemberType NoteProperty -Name DriveId -Value $($_.CustomProperties.GraphDriveId) -Force
            }
        if(!$allClientDrives){$allClientDrives = get-clientDrives} #Only load the Client Drives via graph if there is work to do and we haven't already got them

        #############################
        #Process Projects en-masse as don't want to query every DriveItem in every Drive in advance
        #############################
        #Check for DriveItemId
            #If DriveItemId, update folder
            #If no DriveItemId, Check for Opp
                #If Opp, update DriveItemId and add to flagForReprocessing (to include in next run)
                #If no Opp, Create new folders
        [array]$flagForReprocessing = @()
                                                                                                                                                                                                                                                                        $projTermsToCheck | % {
        $thisProjTerm = $_
        Write-Information "Checking Project Term [$($thisProjTerm.Name)] for CustomProperties.DriveItemId"
        if($thisProjTerm.CustomProperties.DriveItemId){        #If DriveItemId, update folder
            try{
                $thisClientTerm = $allClientTerms | ? {$_.CustomProperties.NetSuiteId -eq $thisProjTerm.CustomProperties.NetSuiteClientId}
                Write-Information "`tProject Term [$($thisProjTerm.Name)].CustomProperties.DriveItemId is [$($thisProjTerm.CustomProperties.DriveItemId)] - setting Folder name to [$($thisProjTerm.Name)]"
                try{
                    $updatedFolder = set-graphDriveItem -tokenResponse $tokenResponseSharePointBot -driveId $thisClientTerm.CustomProperties.GraphDriveId -driveItemId $thisProjTerm.CustomProperties.DriveItemId -driveItemPropertyHash @{name=$thisProjTerm.Name}
                    }
                catch{
                    if($_.Exception -match "(409) Conflict"){
                        Write-Warning "`tPotential duplicate Project folder found for [$($thisClientTerm.Name)][$($thisProjTerm.Name)]"
                        $duplicateNetProjectFolder = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $thisClientTerm.CustomProperties.GraphDriveId | ? {$_.name -eq $thisProjTerm.Name}
                        if($duplicateNetProjectFolder.size -eq 0 -and $duplicateNetProjectFolder.id -ne $thisProjTerm.CustomProperties.DriveItemId){
                            Write-Warning "`tDeleting empty duplicate project folder"
                            delete-graphDriveItem -tokenResponse $tokenResponseSharePointBot -graphDriveId $thisClientTerm.CustomProperties.GraphDriveId -graphDriveItemId $duplicateNetProjectFolder.id -eTag $duplicateNetProjectFolder.eTag -Verbose
                            $updatedFolder = set-graphDriveItem -tokenResponse $tokenResponseSharePointBot -driveId $thisClientTerm.CustomProperties.GraphDriveId -driveItemId $thisProjTerm.CustomProperties.DriveItemId -driveItemPropertyHash @{name=$thisProjTerm.Name}
                            }
                        }
                    if($_.exception -eq "Cannot bind argument to parameter 'graphDriveId' because it is an empty string."){
                        [array]$noDriveForYou += $thisProjTerm
                        #Check whether this is an InterCompany ("pretend") client, and ignore it if it is.
                        }
                    }
                }
            catch{
                $_
                #There's a whole lot that could wrong here ($thisClientTerm.DriveId could be $null, the Drive could be missing or the DriveItem could be missing. Log the error for further investigation
                }
            }
        else{#If no DriveItemId, Check for Opp
            Write-Information "`tProject Term [$($thisProjTerm.Name)].CustomProperties.DriveItemId is missing - checking for Opportunity with NetSuiteProjectId -eq [$($thisProjTerm.CustomProperties.NetSuiteProjId)]"
            $thisOppTerm = $allOppTerms | ? {$_.CustomProperties.NetSuiteProjectId -eq $thisProjTerm.Id}
            if(![string]::IsNullOrWhiteSpace($thisOppTerm.CustomProperties.DriveItemId)){#If Opp, update DriveItemId and add to flagForReprocessing (to include in next run)
                Write-Information "`t`tOpportunity Term [$($thisOppTerm.Name)] found with .CustomProperties.DriveItemId [$($thisOppTerm.CustomProperties.DriveItemId)] - updating Project Term [$($thisProjTerm.Name)].CustomProperties.DriveItemId to match"
                $thisProjTerm.SetCustomProperty("DriveItemId",$thisOppTerm.CustomProperties.DriveItemId)
                try{
                    Write-Verbose "`tTrying: [$($thisProjTerm.Name)].SetCustomProperty(DriveItemId,[$($updatedFolder.id)])"
                    $thisProjTerm.Context.ExecuteQuery()
                    [array]$flagForReprocessing += $thisProjTerm
                    }
                catch{
                    Write-Error "Error updating DriveItemId CustomProperty [$($thisProjTerm.CustomProperties.DriveItemId))] to [$($updatedFolder.id)] for Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisProjTerm.Name)] in sync-netsuiteToManagedMetaData()"
                    [array]$flagForReprocessing += $thisNewClient
                    }
                }
            else{#If no Opp, Create new folders
                Write-Information "`t`tNo corresponding Opportunity Term [$($thisOppTerm.Name)] or DriveItemId [$($thisOppTerm.CustomProperties.DriveItemId)] found - creating new set of Project folders"
                $thisClientTerm = $allClientTerms | ? {$_.CustomProperties.NetSuiteId -eq $thisProjTerm.CustomProperties.NetSuiteClientId}
                [array]$customisedFolderList = $thisProjTerm.Name
                $customisedFolderList += $listOfLeadProjSubFolders | % {"$($thisProjTerm.Name)\$_"}
                $newProjectFolders = add-graphArrayOfFoldersToDrive -graphDriveId $thisClientTerm.CustomProperties.GraphDriveId -foldersAndSubfoldersArray $customisedFolderList -tokenResponse $tokenResponseSharePointBot -conflictResolution Fail #-ErrorAction SilentlyContinue
                $thisProjTerm.SetCustomProperty("DriveItemId",$newProjectFolders[1].id)
                try{
                    Write-Verbose "`tTrying: [$($thisProjTerm.Name)].SetCustomProperty(DriveItemId,[$($updatedFolder.id)])"
                    $thisProjTerm.Context.ExecuteQuery()
                    }
                catch{
                    Write-Error "Error updating DriveItemId CustomProperty [$($thisProjTerm.CustomProperties.DriveItemId))] to [$($updatedFolder.id)] for Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisProjTerm.Name)] in sync-netsuiteToManagedMetaData()"
                    [array]$flagForReprocessing += $thisNewClient
                    }
                }
            }
        }
        #############################
        #Update flagForReprocessing
        #############################
        $projTermsToCheck | % {
            $thisProjToUpdate = $_
            if($flagForReprocessing -notcontains $thisProjToUpdate){ #If the process above worked as expected, update SharePointLastModifiedDate to prevent it from being re-processed next time
                Write-Information "[$($thisProjToUpdate.Name)] was processed successfully - updating flagForReprocessing to [$false]"
                $thisProjToUpdate.SetCustomProperty("flagForReprocessing",$false)
                }
            else{
                Write-Warning "Something went wrong with [$($thisProjToUpdate.Name)] - flagging for reprocessing"
                $thisProjToUpdate.SetCustomProperty("flagForReprocessing",$true)
                }
            try{
                Write-Verbose "`tTrying: [$($thisProjToUpdate.Name)][$($thisProjToUpdate.Id)].SetCustomProperty(SharePointLastModifiedDate,$($now))"
                $thisProjToUpdate.Context.ExecuteQuery()
                }
            catch{
                Write-Error "Error setting CustomProperty SharePointLastModifiedDate on Term [$($thisProjToUpdate.Name)][$($thisProjToUpdate.Id)] in sync-netsuiteToManagedMetaData()"
                $_
                }
            }
        }

    #endregion


    ###########################################
    #If the script hasn't borked completely, update the LastSpoSyncRun timestamp
    Write-Information "Setting Term [Anthesis][IT][LastModified] CustomProperty LastSpoSyncRun = [$(Get-Date $now -f s)]"
    $lastProcessedTerm = Get-PnPTerm -TermGroup "Anthesis" -TermSet "IT" -Identity "LastModified" -Includes CustomProperties
    $lastProcessedTerm.SetCustomProperty("LastSpoSyncRun",$(Get-Date $now -f s))
    try{
        $lastProcessedTerm.Context.ExecuteQuery()
        }
    catch{
        #Pfft.
        }


    }

Write-Information "sync-netsuiteManagedMetaDataToSharePoint completed in [$($fullSyncTime.TotalSeconds)] seconds"

Stop-Transcript