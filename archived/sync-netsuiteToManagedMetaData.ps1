if($PSCommandPath){
    $InformationPreference = 2
    $VerbosePreference = 0
    $logFileLocation = "C:\ScriptLogs\"
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))`_Transcript_$(Get-Date -Format "yyyy-MM-dd").log"
    Start-Transcript $transcriptLogName -Append
    }

$sharePointAdmin = "kimblebot@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Downloads\KimbleBot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com" -Credentials $adminCreds

Write-Host "sync-netsuiteToManagedMetaData started at $(Get-Date -Format s)"
$fullSyncTime = Measure-Command {
    #region Prospects/Clients
    $pnpTermGroup = "Kimble"
    $pnpTermSet = "Clients"
    $allClientTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes CustomProperties | ? {$_.IsDeprecated -eq $false}
    $allClientTerms | % {
        Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteId -Value $_.CustomProperties.NetSuiteId -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name Name2 -Value $($_.Name).Replace("&","").Replace("＆","").Replace("  "," ").Trim() -Force
        }

    [datetime]$lastProcessed = $($allClientTerms | sort {$_.CustomProperties.NetSuiteLastModifiedDate} | select -Last 1).CustomProperties.NetSuiteLastModifiedDate

    $netQuery =  "?q=companyName CONTAIN_NOT `"Anthesis`"" #Excludes any Companies with "Anthesis" in the companyName
    $netQuery += " AND companyName CONTAIN_NOT `"intercompany project`"" #Excludes any Companies with "(intercompany project)" in the companyName
    $netQuery += " AND companyName START_WITH_NOT `"x `"" #Excludes any Companies that begin with "x " in the companyName
    $netQuery += " AND entityStatus ANY_OF_NOT [6, 7]" #Excludes LEAD-Unqualified and LEAD-Qualified (https://XXX.app.netsuite.com/app/crm/sales/customerstatuslist.nl?whence=)
    $netQuery += " AND lastModifiedDate ON_OR_AFTER `"$($(Get-Date $lastProcessed -Format g))`"" #Excludes any Companies that haven;t been updated since X
    [array]$clientsToCheck = get-netSuiteClientsFromNetSuite -query $netQuery -netsuiteParameters $(get-netSuiteParameters -connectTo Production)
    $clientsToCheck = $clientsToCheck | ? {$(Get-Date $_.lastModifiedDate) -ge $(Get-Date $lastProcessed)} #We can currently only filter by date in NetSuite, so filter again client-side to exclude all the other Clients we've processed earlier today
    [array]$processedAtExactlyLastTimestamp = $clientsToCheck | ? {$(Get-Date $_.lastModifiedDate) -ge $(Get-Date $lastProcessed)} #Find how many Clients match the $lastProcessed timestamp exactly
    if($processedAtExactlyLastTimestamp.Count -eq 1){$clientsToCheck = $clientsToCheck | ? {$clientsToCheck.id -notcontains $processedAtExactlyLastTimestamp[0].id}} #If it's exactly one, exclude it from processing (as we've already processed it on a previous cycle)
    
    Write-Host "Processing [$($clientsToCheck.Count)] Clients"
    #$clientsToCheck = $clientsToCheck | ? {$_.entityStatus.refName -notmatch "LEAD"} #Filter out Leads
    $clientsToCheck | Select-Object | % { #Set the objects up so they are easy to compare
        Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteId -Value $_.id -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name Name2 -Value $($_.companyName).Replace("&","").Replace("＆","").Replace("  "," ").Trim() -Force #We've inherited shitly-named Client records
        }

    #Match all updated NetSuite records against Managed Metadata using NetSuiteId
        #If no Id match, then check if a Term already exists due to shitty data
            #If Term exists, back up any out-of-date NetSuiteId value and re-use the Term and flag for reprocessing
            #If Term does not exist, create new Client Term and flag for reprocessing
        #If Id match then check if Name has changed
            #If Name has not changed, do nothing
            #if Name has changed, rename the Term and flag for reprocessing
        #If nothing went wrong, update NetSuiteLastModifiedDate on the Term to exclude it from the future cycles (untilthe NetSuite record is updated again)

    #############################
    #Create new Prospects/Clients
    #############################
    [array]$doNotUpdateLastModified = @() #If anything goes wrong processing a Client, we don't want to update the NetSuiteProjLastModifiedDate CustomProperty on the Term as the mismatch means it will get picked up in the next Full Reconcile
    $deltaClientId = Compare-Object -ReferenceObject @($clientsToCheck | Select-Object) -DifferenceObject $allClientTerms -Property NetSuiteId -PassThru -IncludeEqual #Match all updated NetSuite records against Managed Metadata using NetSuiteId
    [array]$missingFromTermStore = $deltaClientId | ? {$_.SideIndicator -eq "<="}
    Write-Host "`tProcessing [$($missingFromTermStore.Count)] new Clients"
    @($missingFromTermStore | Select-Object) |  % { #If no Id match, then create new Client Term and flag for reprocessing
        $thisNewClient = $_
        Write-Host "`t`tProcessing new Client [$($thisNewClient.companyName)]"
        $testForCollision = $allClientTerms | ? {$_.Name2 -eq $thisNewClient.Name2}
        if($testForCollision){#If Term exists, back up any out-of-date NetSuiteId value and re-use the Term and flag for reprocessing
            Write-Warning "`tThere is already a term with the same default label and parent term [$($thisNewClient.companyName)] - cannot create new Client Term."
            if(![string]::IsNullOrWhiteSpace($testForCollision.CustomProperties.NetSuiteId) -and $testForCollision.CustomProperties.NetSuiteId -ne $thisNewClient.id){ #If the Term already has a _different_ NetSuiteId then somthing has gone badly wrong. We need to preserve this information so we can unpick it later, so we'll preserve the old NetSuiteId by suffixing it with _overwritten$i
                Write-Warning "`t`t`t`tTerm [$($testForCollision.Name)][$($testForCollision.Id)] already has a NetSuiteId of [$($testForCollision.CustomProperties.NetSuiteId)] - backing this up and overwriting with NetSuiteId [$($thisNewClient.id)]"
                while(![string]::IsNullOrWhiteSpace($testForCollision.CustomProperties."NetSuiteId_overwritten$i")){ #Find the lowest number for merging without overwriting any pre-existing CustomProperties
                    $i++
                    }
                $testForCollision.SetCustomProperty("NetSuiteId_overwritten$i",$testForCollision.CustomProperties.NetSuiteId) #Add this CustomProperty back into the CustomProperties as a pseudo-backup
                $testForCollision.SetCustomProperty("NetSuiteId",$thisNewClient.id) #Set the correct NEtsuiteId
                $testForCollision.SetCustomProperty("flagForReprocessing",$true) #Set the flag for reprocessing so this Term gets processed into SharePoint
                try{
                    Write-Host "`t`t`t`t`tReusing existing Term [$($testForCollision.Name)][$($testForCollision.Id)] and setting [NetSuiteId_overwritten$i] to [$($testForCollision.CustomProperties.NetSuiteId)], and [NetSuiteId] to [$($thisNewClient.id)]"
                    $testForCollision.Context.ExecuteQuery()
                    }
                catch{
                    Write-Error "Error `"backing up`" an old NetSuiteId value [$($testForCollision.CustomProperties.NetSuiteId))] to [NetSuiteId_overwritten$i] for Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisNewClient.companyName)] in sync-netsuiteToManagedMetaData()"
                    Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                    [array]$doNotUpdateLastModified += $thisNewClient
                    return #If we can't backup the odl NetSuiteId value, skip over of this Client
                    }
                }
            else{#Just update NetSuiteId with the correct value
                $testForCollision.SetCustomProperty("NetSuiteId",$thisNewClient.id) #Set the correct NEtsuiteId
                $testForCollision.SetCustomProperty("NetSuiteLastModifiedDate",$thisNewClient.lastModifiedDate)
                $testForCollision.SetCustomProperty("flagForReprocessing",$true) #Set the flag for reprocessing so this Term gets processed into SharePoint
                try{
                    Write-Host "`t`t`tReusing existing Term [$($testForCollision.Name)] and setting [NetSuiteId]=[$($thisNewClient.id)] & [flagForReprocessing]=[$true])"
                    $testForCollision.Context.ExecuteQuery()
                    }
                catch{
                    Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                    [array]$doNotUpdateLastModified += $thisNewClient
                    return #If we can't backup the odl NetSuiteId value, skip over of this Client
                    }
                }
            }
        else{#If Term does not exist, create new Client Term and flag for reprocessing
            try{
                Write-Host "`t`t`tCreating new Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisNewClient.companyName)][@{NetSuiteId=$($thisNewClient.id);NetSuiteLastModifiedDate=$($thisNewClient.lastModifiedDate);flagForReprocessing=$true]"
                $newTerm = New-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Name $thisNewClient.companyName -Lcid 1033 -CustomProperties @{NetSuiteId=$thisNewClient.id;NetSuiteLastModifiedDate=$thisNewClient.lastModifiedDate;flagForReprocessing=$true}
                }
            catch{
                Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                [array]$doNotUpdateLastModified += $thisNewClient
                }
            }
        }
    #############################
    #Update any Prospects/Clients 
    #############################
    $matchedId = $deltaClientId | ? {$_.SideIndicator -eq "=="}
    $matchedIdReversed = Compare-Object -ReferenceObject $allClientTerms -DifferenceObject @($matchedId | Select-Object) -Property NetSuiteId -PassThru -IncludeEqual -ExcludeDifferent #We then use $matchedId to filter only the Terms with corresponding $clientsToCheck records
        <#Sanity check - these should produce identical results, (but weirdly you have to run them separately). CSOM, eh?:
        $matchedId | sort NetSuiteId | select companyName -First 10
        $matchedIdReversed | sort NetSuiteId | select Name -First 10
        #>
    $deltaName = Compare-Object -ReferenceObject @($matchedId | Select-Object) -DifferenceObject @($matchedIdReversed | Select-Object) -Property NetSuiteId,Name2 -PassThru #We compare the two equal sets on both NetSuiteId and Name2 to see which pairs have mismatched Name values
    [array]$clientsWithChangedNames = $deltaName | ? {$_.SideIndicator -eq "<="}
    Write-Host "`tProcessing [$($clientsWithChangedNames.Count)] renamed Clients"
    @($clientsWithChangedNames | Select-Object) | % {
        $thisUpdatedClient = $_
        Write-Host "`t`tProcessing renames Company [$($thisUpdatedClient.companyName)][$($thisUpdatedClient.id)]"
        $termWithWrongName = $matchedIdReversed | ? {$_.NetSuiteId -eq $thisUpdatedClient.NetSuiteId}
        if ($termWithWrongName.Count -eq 1){
            Write-Host "`t`t`tRenaming Term `t[$($termWithWrongName.Name)][$($termWithWrongName.Id)]"
            Write-Host "`t`t`tto:`t`t`t`t[$($thisUpdatedClient.companyName)]"
            $termWithWrongName_originalName = $termWithWrongName.Name
            $termWithWrongName.Name = $thisUpdatedClient.companyName
            $termWithWrongName.SetCustomProperty("flagForReprocessing",$true)
            try{
                Write-Verbose "`tTrying: [$($termWithWrongName_originalName)].Name = [$($thisUpdatedClient.companyName)]"
                $termWithWrongName.Context.ExecuteQuery()
                }
            catch {
                if($_.Exception -match "TermStoreErrorCodeEx:There is already a term with the same default label and parent term."){
                    Write-Warning "`t`tThere is already a term with the same default label and parent term [$($termWithWrongName_originalName)]->[$($thisUpdatedClient.companyName)]"
                    #If there is already a Term with the same name, merge the would-be-collision into this Term and preserve any conflicting CustomProperties by suffixing them with _merged$i
                    $termWithWrongName.Name = $termWithWrongName_originalName #We need to set this back in case something went wrong with a previous .Merge() and we need mess about with Labels
                    $duffTermToMergeIntoGoodTerm = $allClientTerms | ? {$_.Name2 -eq $thisUpdatedClient.Name2 -and $_.Id -ne $termWithWrongName.id}
                    if($duffTermToMergeIntoGoodTerm){ #If there's another Term, merge them
                        try{
                            Write-Host "`t`t`t`tMerging Terms -termToBeRetained [$($termWithWrongName.Name)] -termToBeMerged [$($duffTermToMergeIntoGoodTerm.Name)] -pnpTermGroup $pnpTermGroup -pnpTermSet $pnpTermSet"
                            merge-pnpTerms -termToBeRetained $termWithWrongName -termToBeMerged $duffTermToMergeIntoGoodTerm -setDefaultLabelTo Merged -pnpTermGroup $pnpTermGroup -pnpTermSet $pnpTermSet -Verbose:$VerbosePreference
                            }
                        catch{
                            Write-Error "Error merging Term [$($pnpTermGroup)][$($pnpTermSet)][$($duffTermToMergeIntoGoodTerm.Name)] into [$($termWithWrongName.Name)] in sync-netsuiteToManagedMetaData()"
                            Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                            [array]$doNotUpdateLastModified += $thisUpdatedClient
                            }
                        }
                    else{#If there isn't another Term, they've probably already been merged, so try relabelling it.
                        Write-Host "`t`t`t`tSetting default Label to [$($thisUpdatedClient.companyName)] for Term [$($termWithWrongName.Name)][$($termWithWrongName.Id)]"
                        $i=0
                        do{ #CSOM Voodoo 
                            if($i -eq 0){$updatedTerm = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Identity $termWithWrongName.Id -Includes CustomProperties, Labels} #Refresh the Term to ensure we've got the correct Labels
                            else{
                                Write-Verbose "Term name is still [$($updatedTerm.Name)] on iteration [$($i)]  - sleeping for another 5 seconds and dancing widdershins around the Term"
                                Start-Sleep -Seconds 5
                                }
                            $updatedTerm.Labels | ? {$_.Value -eq $thisUpdatedClient.companyName} | Out-Null
                            $($updatedTerm.Labels | ? {$_.Value -eq $thisUpdatedClient.companyName}) | fl # .SetAsDefaultForLanguage() only works if the relevant Label has been enumerated to the screen. WTF. CSOM, eh?
                            $correctLabel = $updatedTerm.Labels | ? {$_.Value -eq $thisUpdatedClient.companyName}
                            $correctLabel.SetAsDefaultForLanguage()
                            $updatedTerm.SetCustomProperty("flagForReprocessing",$true)
                            $updatedTerm.SetCustomProperty("NetSuiteLastModifiedDate",$thisUpdatedClient.lastModifiedDate)
                            try{
                                Write-Verbose "`tTrying: [$($updatedTerm.Name)].[$($correctLabel.Value)].SetAsDefaultForLanguage()"
                                $updatedTerm.Context.ExecuteQuery()
                                }
                            catch{
                                Write-Error "Error setting Default Label to [$($correctLabel.Value)] on Term [$($pnpTermGroup)][$($pnpTermSet)][$($updatedTerm.Name)] in sync-netsuiteToManagedMetaData()"
                                Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                                [array]$doNotUpdateLastModified += $thisUpdatedClient
                                }
                            $i++
                            $updatedTerm = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Identity $duffTermToMergeIntoGoodTerm.Id -Includes CustomProperties, Labels
                            }
                        while($updatedTerm.Name -eq $termWithWrongName_originalName)
                        }
                    }
                elseif($_.Exception -match "TermStoreEx:Failed to read from or write to database. Refresh and try again. If the problem persists, please contact the administrator."){
                    Write-Warning "Failed to read from or write to database. Refresh and try again. [$($termWithWrongName_originalName)]->[$($thisUpdatedClient.companyName)]"
                    #If there isn't another Term, they've probably already been merged, so try relabelling it.
                    Write-Host "Setting default Label to [$($thisUpdatedClient.companyName)] for Term [$($termWithWrongName.Name)][$($termWithWrongName.Id)]"
                    $i=0
                    do{#CSOM Voodoo
                        if($i -eq 0){$updatedTerm = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Identity $termWithWrongName.Id -Includes CustomProperties, Labels} #Refresh the Term to ensure we've got the correct Labels
                        else{
                            Write-Verbose "Term name is still [$($updatedTerm.Name)] on iteration [$($i)]  - sleeping for another 5 seconds and dancing widdershins around the Term"
                            Start-Sleep -Seconds 5
                            }
                        $updatedTerm.Labels | ? {$_.Value -eq $thisUpdatedClient.companyName} | Out-Null
                        $($updatedTerm.Labels | ? {$_.Value -eq $thisUpdatedClient.companyName}) | fl # .SetAsDefaultForLanguage() only works if the relevant Label has been enumerated to the screen. WTF. CSOM, eh?
                        $correctLabel = $updatedTerm.Labels | ? {$_.Value -eq $thisUpdatedClient.companyName}
                        $correctLabel.SetAsDefaultForLanguage()
                        $updatedTerm.SetCustomProperty("flagForReprocessing",$true)
                        $updatedTerm.SetCustomProperty("NetSuiteLastModifiedDate",$thisUpdatedClient.lastModifiedDate)
                        try{
                            Write-Verbose "`tTrying: [$($updatedTerm.Name)].[$($correctLabel.Value)].SetAsDefaultForLanguage()"
                            $updatedTerm.Context.ExecuteQuery()
                            }
                        catch{
                            Write-Error "Error setting Default Label to [$($correctLabel.Value)] on Term [$($pnpTermGroup)][$($pnpTermSet)][$($updatedTerm.Name)] in sync-netsuiteToManagedMetaData()"
                            Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                            [array]$doNotUpdateLastModified += $thisUpdatedClient
                            }
                        $i++
                        $updatedTerm = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Identity $termWithWrongName.Id -Includes CustomProperties, Labels
                        }
                    while($updatedTerm.Name -eq $termWithWrongName_originalName)
                    }
                else{
                    Write-Error "Error renaming Term [$($termWithWrongName_originalName)] to [$($thisUpdatedClient.companyName)] in sync-netsuiteToManagedMetaData()"
                    Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                    [array]$doNotUpdateLastModified += $thisUpdatedClient
                    }
                }
            }
        else{
             Write-Warning "Could not find corresponding Term for updated NetSuite Client [$($thisUpdatedClient.companyName)][$($thisUpdatedClient.id)]"
             [array]$doNotUpdateLastModified += $thisUpdatedClient
            }
        }


    #############################
    #Update LastModifiedDate
    #############################
    Write-Host "`tRemaining [$($clientsToCheck.Count - $missingFromTermStore.Count -$clientsWithChangedNames.Count)] Clients must have had minor/irrelevant updates."
    Write-Host "`tUpdating lastmodified timestamps (if everything worked)"
    @($clientsToCheck | Select-Object) | % {
        $thisClientToUpdate = $_
        if($doNotUpdateLastModified.id -notcontains $thisClientToUpdate.id){ #If the rename/merge didn't explictly fail, update the NetSuiteLastModified CustomProperty. This will update NetSuiteLastModified for all successful updates, all new Terms and all Clients that were updated in NetSuite but didn;t have Name changes
            $updatedTerm = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Identity $thisClientToUpdate.companyName -Includes CustomProperties, Labels
            if($updatedTerm){ 
                $updatedTerm.SetCustomProperty("NetSuiteLastModifiedDate",$thisClientToUpdate.lastModifiedDate)
                Write-Host "`t`t[$($thisClientToUpdate.companyName)] was processed successfully - updating NetSuiteLastModifiedDate to [$($thisClientToUpdate.lastModifiedDate)]"
                try{
                    Write-Verbose "`tTrying: [$($updatedTerm.Name)][$($updatedTerm.Id)].SetCustomProperty(NetSuiteLastModifiedDate,$($thisClientToUpdate.lastModifiedDate))"
                    $updatedTerm.Context.ExecuteQuery()
                    }
                catch{
                    Write-Error "Error setting CustomProperty NetSuiteLastModifiedDate = [$($thisClientToUpdate.lastModifiedDate)] on Term [$($updatedTerm.Name)][$($updatedTerm.Id)] in sync-netsuiteToManagedMetaData()"
                    Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                    }
                }
            }
        else{
            Write-Warning "Something went wrong with [$($thisClientToUpdate.Name)] - not updating NetSuiteLastModifiedDate"
            }
        }



    #endregion

    #region Opportunities
    Write-Host ""
    $pnpTermGroup = "Kimble"
    $pnpTermSet = "Opportunities"
    $allOppTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes CustomProperties | ? {$_.IsDeprecated -eq $false}
    $allOppTerms | Select-Object | % { #Set the objects up so they are easy to compare. compare-object was struggling with $nulls, "" and whitespaces, so we're standardising on $null here
        if([string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteOppId)){Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteId -Value $null -Force}
        else{Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteId -Value $_.CustomProperties.NetSuiteOppId -Force}
        if([string]::IsNullOrWhiteSpace($_.Name)){Add-Member -InputObject $_ -MemberType NoteProperty -Name Name2 -Value $null -Force}
        else{Add-Member -InputObject $_ -MemberType NoteProperty -Name Name2 -Value $_.Name -Force}
        if([string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteClientId)){Add-Member -InputObject $_ -MemberType NoteProperty -Name ClientId -Value $null -Force}
        else{Add-Member -InputObject $_ -MemberType NoteProperty -Name ClientId -Value $_.CustomProperties.NetSuiteClientId -Force}
        if([string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteProjectId)){Add-Member -InputObject $_ -MemberType NoteProperty -Name ProjectId -Value $null -Force}
        else{Add-Member -InputObject $_ -MemberType NoteProperty -Name ProjectId -Value $_.CustomProperties.NetSuiteProjectId -Force}
        }

    [datetime]$lastProcessed = $($allOppTerms | sort {$_.CustomProperties.NetSuiteOppLastModifiedDate} | select -Last 1).CustomProperties.NetSuiteOppLastModifiedDate

    $netQuery =  "?q=lastModifiedDate ON_OR_AFTER `"$($(Get-Date $lastProcessed -Format g).Split(" ")[0])`"" #Excludes any Opps that haven;t been updated since X
    #$netQuery =  "?q=lastModifiedDate ON_OR_AFTER `"$(get-dateInIsoFormat $lastProcessed -precision Milliseconds)`"" #Excludes any Opps that haven;t been updated since X ?q=lastModifiedDate ON_OR_AFTER "15/01/2021 11:24:00"
    [array]$oppsToCheck = get-netSuiteOpportunityFromNetSuite -query $netQuery -netsuiteParameters $(get-netSuiteParameters -connectTo Production)
    $oppsToCheck = $oppsToCheck | ? {$(Get-Date $_.lastModifiedDate) -ge $(Get-Date $lastProcessed)} #We can currently only filter by date in NetSuite, so filter again client-side to exclude all the other Opportunities we've processed earlier today
    [array]$processedAtExactlyLastTimestamp = $oppsToCheck | ? {$(Get-Date $_.lastModifiedDate) -ge $(Get-Date $lastProcessed)} #Find how many Opportunities match the $lastProcessed timestamp exactly
    if($processedAtExactlyLastTimestamp.Count -eq 1){$oppsToCheck = $oppsToCheck | ? {$oppsToCheck.id -notcontains $processedAtExactlyLastTimestamp[0].id}} #If it's exactly one, exclude it from processing (as we've already processed it on a previous cycle)
    
    Write-Host "Processing [$($oppsToCheck.Count)] Opportunities"
    #$oppsToCheck = get-netSuiteOpportunityFromNetSuite -netsuiteParameters $(get-netSuiteParameters -connectTo Production) 
    $oppsToCheck | % { #Set the objects up so they are easy to compare. compare-object was struggling with $nulls, "" and whitespaces, so we're standardising on $null here
        if([string]::IsNullOrWhiteSpace($_.id)){Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteId -Value $null -Force}
        else{Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteId -Value $_.id -Force}
        if([string]::IsNullOrWhiteSpace("$($_.tranId) $($_.title)")){Add-Member -InputObject $_ -MemberType NoteProperty -Name Name2 -Value $null -Force}
        else{Add-Member -InputObject $_ -MemberType NoteProperty -Name Name2 -Value "$($_.tranId) $($_.title)" -Force}
        if([string]::IsNullOrWhiteSpace($_.entity.id)){Add-Member -InputObject $_ -MemberType NoteProperty -Name ClientId -Value $null -Force}
        else{Add-Member -InputObject $_ -MemberType NoteProperty -Name ClientId -Value $_.entity.id -Force}
        if([string]::IsNullOrWhiteSpace($_.custbody_project_created.id)){Add-Member -InputObject $_ -MemberType NoteProperty -Name ProjectId -Value $null -Force}
        else{Add-Member -InputObject $_ -MemberType NoteProperty -Name ProjectId -Value $_.custbody_project_created.id -Force}
        }

    #Match all updated NetSuite records against Managed Metadata using NetSuiteId
        #If no Id match, then create new Opp Term and flag for reprocessing
        #If Id match then check if Name has changed
            #If Name has not changed, do nothing
            #if Name has changed, rename the Term and flag for reprocessing
        #If Id match then check if associated ClientId has changed
            #If ClientId has not changed, do nothing
            #if ClientId has changed, change the NetSuiteClientId for the Term and flag for reprocessing
        #If nothing went wrong, update NetSuiteLastModifiedDate on the Term to exclude it from the future cycles (untilthe NetSuite record is updated again)


    #############################
    #Create new Opps
    #############################
    [array]$doNotUpdateLastModified = @() #If anything goes wrong processing an Opp, we don't want to update the NetSuiteProjLastModifiedDate CustomProperty on the Term as the mismatch means it will get picked up in the next Full Reconcile
    $deltaOppId = Compare-Object -ReferenceObject @($oppsToCheck | Select-Object) -DifferenceObject $allOppTerms -Property NetSuiteId -PassThru -IncludeEqual #Match all updated NetSuite records against Managed Metadata using NetSuiteId
    [array]$missingFromTermStore = $deltaOppId | ? {$_.SideIndicator -eq "<="}
    Write-Host "`tProcessing [$($missingFromTermStore.Count)] new Opportunities"
    @($missingFromTermStore | Select-Object) | % {
        $thisNewOpp = $_
        $thisOppLabel = "$($thisNewOpp.tranId) $($thisNewOpp.title)"
        Write-Host "`t`tProcessing new Opportunity [$thisOppLabel][$($thisNewOpp.id)] for Client ID [$($thisNewOpp.entity.id)]"
        $testForCollision = $allOppTerms | ? {$_.Name -eq $thisNewOpp.Name}
        if($testForCollision){
            Write-Warning "`tThere is already a term with the same default label and parent term [$($thisOppLabel)] - cannot create new Opportunity Term."
            #If there is already a Term with the same name, re-use it
            if(![string]::IsNullOrWhiteSpace($testForCollision.CustomProperties.NetSuiteOppId) -and $testForCollision.CustomProperties.NetSuiteOppId -ne $thisNewOpp.id){ #If the Term already has a _different_ NetSuiteOppId then somthing has gone badly wrong. We need to preserve this information so we can unpick it later, so we'll preserve the old NetSuiteOppId by suffixing it with _overwritten$i
                while(![string]::IsNullOrWhiteSpace($testForCollision.CustomProperties."NetSuiteOppId_overwritten$i")){ #Find the lowest number for merging without overwriting any pre-existing CustomProperties
                    $i++
                    }
                $testForCollision.SetCustomProperty("NetSuiteOppId_overwritten$i",$testForCollision.CustomProperties.NetSuiteOppId) #Add this CustomProperty back into the CustomProperties as a pseudo-backup
                $testForCollision.SetCustomProperty("NetSuiteOppId",$thisNewOpp.id) #Set the correct NetSuiteOppId
                $testForCollision.SetCustomProperty("flagForReprocessing",$true) #Set the flag for reprocessing so this Term gets processed into SharePoint
                try{
                    Write-Host "`t`t`t`t`tReusing existing Term [$($testForCollision.Name)][$($testForCollision.Id)] and setting [NetSuiteOppId_overwritten$i] to [$($testForCollision.CustomProperties.NetSuiteOppId)], and [NetSuiteOppId] to [$($thisNewOpp.id)]"
                    $testForCollision.Context.ExecuteQuery()
                    }
                catch{
                    Write-Error "Error `"backing up`" an old NetSuiteOppId value [$($testForCollision.CustomProperties.NetSuiteOppId))] to [NetSuiteOppId_overwritten$i] for Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisOppLabel)] in sync-netsuiteToManagedMetaData()"
                    Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                    [array]$doNotUpdateLastModified += $thisNewOpp
                    return #If we can't backup the old NetSuiteOppId value, skip over this Opp
                    }
                }
            }
        else{
            try{
                Write-Host "`t`t`tCreating new Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisOppLabel)][@{NetSuiteOppId=$($thisNewOpp.id);NetSuiteOppLastModifiedDate=$($thisNewOpp.lastModifiedDate);flagForReprocessing=$true]"
                Write-Verbose "`tTrying: New-PnPTerm -TermGroup [$pnpTermGroup] -TermSet [$pnpTermSet] -Name [$($thisOppLabel)] -Lcid 1033 -CustomProperties @{NetSuiteOppId=$($thisNewOpp.id);NetSuiteOppLastModifiedDate=$($thisNewOpp.lastModifiedDate);NetSuiteClientId=$($thisNewOpp.entity.id);NetSuiteProjectId=$($thisNewOpp.custbody_project_created.id);flagForReprocessing=$true}"
                $newTerm = New-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Name $thisOppLabel -Lcid 1033 -CustomProperties @{NetSuiteOppId=$thisNewOpp.id;NetSuiteOppLastModifiedDate=$thisNewOpp.lastModifiedDate;NetSuiteClientId=$thisNewOpp.entity.id;NetSuiteProjectId=$thisNewOpp.custbody_project_created.id;flagForReprocessing=$true}
                }
            catch{
                Write-Error "Error creating new Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisOppLabel)][@{NetSuiteOppId=$($thisNewOpp.id);NetSuiteOppLastModifiedDate=$($thisNewOpp.lastModifiedDate);NetSuiteClientId=$($thisNewOpp.entity.id);NetSuiteProjectId=$($thisNewOpp.custbody_project_created.id)}] in sync-NetsuiteTpManagedMetaData()"
                Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                [array]$doNotUpdateLastModified += $thisNewOpp
                }
            }
        }


    #############################
    #Update Opps (Names & ClientIDs)
    #############################
    $matchedId = $deltaOppId | ? {$_.SideIndicator -eq "=="}
    $matchedIdReversed = Compare-Object -ReferenceObject $allOppTerms -DifferenceObject @($matchedId | Select-Object) -Property NetSuiteId -PassThru -IncludeEqual -ExcludeDifferent #We then use $matchedId to filter only the Terms with corresponding $clientsToCheck records
        <#Sanity check - these should produce identical results, (but weirdly you have to run them separately). CSOM, eh?:
        $matchedId | sort NetSuiteId | select Name2 -Last 10
        $matchedIdReversed | sort NetSuiteId | select Name -Last 10
        #>
    $deltaName = Compare-Object -ReferenceObject @($matchedId | Select-Object) -DifferenceObject @($matchedIdReversed | Select-Object) -Property NetSuiteId,Name2 -PassThru -IncludeEqual #We compare the two equal sets on both NetSuiteId and Name2 to see which pairs have mismatched Name values
    [array]$oppsWithChangedNames = $deltaName | ? {$_.SideIndicator -eq "<="}
    Write-Host "`tProcessing [$($oppsWithChangedNames.Count)] Opportunities with changed names"
    @($oppsWithChangedNames | Select-Object) | % {
        $thisUpdatedOpp = $_
        Write-Host "`t`tProcessing updated Opportunity [$($thisUpdatedOpp.tranId) $($thisUpdatedOpp.title)][$($thisUpdatedOpp.id)]"
        $termWithWrongName = $matchedIdReversed | ? {$_.NetSuiteId -eq $thisUpdatedOpp.NetSuiteId}
        if ($termWithWrongName.Count -eq 1){
            Write-Host "`t`t`tRenaming Term `t[$($termWithWrongName.Name)][$($termWithWrongName.Id)]"
            Write-Host "`t`t`tto:`t`t`t`t[$($thisUpdatedOpp.tranId) $($thisUpdatedOpp.title)]"
            $termWithWrongName_originalName = $termWithWrongName.Name
            $termWithWrongName.Name = "$($thisUpdatedOpp.tranId) $($thisUpdatedOpp.title)"
            try{
                $termWithWrongName.Context.ExecuteQuery()
                }
            catch {
                Write-Error "Error renaming Term [$($termWithWrongName_originalName)] to [$($termWithWrongName.Name)][$($termWithWrongName.Id)] in sync-netsuiteToManagedMetaData()"
                Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                [array]$doNotUpdateLastModified += $thisUpdatedOpp
                }
            }
        else{
             Write-Warning "`tCould not find corresponding Term for updated NetSuite Opp [$($termWithWrongName.Name)][$($termWithWrongName.Id)][$($thisUpdatedOpp.id)]"
             [array]$doNotUpdateLastModified += $thisUpdatedOpp
            }    
        }

    $deltaClientId = Compare-Object -ReferenceObject @($matchedId | Select-Object) -DifferenceObject @($matchedIdReversed | Select-Object) -Property NetSuiteId,ClientId -PassThru #We compare the two equal sets on both NetSuiteId and ClientId to see which pairs have mismatched Name values
    [array]$oppsWithChangedClient = $deltaClientId | ? {$_.SideIndicator -eq "<="}
    Write-Host "`tProcessing [$($oppsWithChangedClient.Count)] Opportunities with changed Clients"
    @($oppsWithChangedClient | Select-Object) | % {
        $thisUpdatedOpp = $_
        Write-Host "`t`tProcessing updated Opportunity [$($thisUpdatedOpp.tranId) $($thisUpdatedOpp.title)][$($thisUpdatedOpp.id)]"
        $termWithWrongClient = $matchedIdReversed | ? {$_.NetSuiteId -eq $thisUpdatedOpp.NetSuiteId}
        if ($termWithWrongClient.Count -eq 1){
            Write-Host "`t`t`tReassigning Opportunity Term [$($termWithWrongClient.Name)][$($termWithWrongClient.Id)] to Client [$($thisUpdatedOpp.parent.id)]"
            while(![string]::IsNullOrWhiteSpace($termWithWrongClient.CustomProperties."NetSuiteClientId_previous$i")){ #Find the lowest number for merging without overwriting anything
                $i++
                }
            $termWithWrongClient.SetCustomProperty("NetSuiteClientId_previous$i",$termWithWrongClient.CustomProperties.NetSuiteClientId)
            $termWithWrongClient.SetCustomProperty("NetSuiteClientId",$thisUpdatedOpp.parent.id)
            try{
                Write-Host "`t`t`t`t`tSetting [(NetSuiteClientId_previous$i] to [$($termWithWrongClient.CustomProperties.NetSuiteClientId)], and [NetSuiteClientId] to [$($thisUpdatedOpp.parent.id)] for Opp [$($thisUpdatedOpp.tranId) $($thisUpdatedOpp.title)][$($thisUpdatedOpp.id)]"
                $termWithWrongClient.Context.ExecuteQuery()
                }
            catch {
                Write-Error "Error reassigning Opportunity Term [$($termWithWrongClient.Name)] to Client [$($thisUpdatedOpp.parent.id)] in sync-netsuiteToManagedMetaData()"
                Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                [array]$doNotUpdateLastModified += $thisUpdatedOpp
                }
            }
        else{
             Write-Warning "Could not find corresponding Term for updated NetSuite Opportunity [$($termWithWrongClient.Name)][$($termWithWrongClient.Id)][$($thisUpdatedOpp.id)]"
             [array]$doNotUpdateLastModified += $thisUpdatedOpp
            }    
        }

    $deltaProjectId = Compare-Object -ReferenceObject @($matchedId | Select-Object) -DifferenceObject @($matchedIdReversed | Select-Object) -Property NetSuiteId,ProjectId -PassThru #We compare the two equal sets on both NetSuiteId and NetSuiteProjectId to see which pairs have mismatched NetSuiteProjectId values
    [array]$oppsWithChangedProject = $deltaProjectId | ? {$_.SideIndicator -eq "<="} | ? {![string]::IsNullOrWhiteSpace($_.ProjectId)} 
    Write-Host "`tProcessing [$($oppsWithChangedProject.Count)] Opportunities with changed Projects"
    @($oppsWithChangedProject | Select-Object) | % {
        $thisUpdatedOpp = $_
        Write-Host "`t`tProcessing updated Opportunity [$($thisUpdatedOpp.tranId) $($thisUpdatedOpp.title)][$($thisUpdatedOpp.id)]"
        $termWithWrongProject = $matchedIdReversed | ? {$_.NetSuiteId -eq $thisUpdatedOpp.NetSuiteId}
        if ($termWithWrongProject.Count -eq 1){
            Write-Host "`t`t`tSetting [NetSuiteProjectId] = [$($thisUpdatedOpp.ProjectId)] (previously [$($termWithWrongProject.CustomProperties.NetSuiteProjectId)]) for Opportunity Term [$($termWithWrongClient.Name)][$($termWithWrongClient.Id)]"
            $termWithWrongProject.SetCustomProperty("NetSuiteProjectId",$thisUpdatedOpp.ProjectId)
            try{
                Write-Host "`t`t`t`t`tSetting [(NetSuiteProjectId_previous$i] to [$($termWithWrongClient.CustomProperties.NetSuiteProjectId)], and [NetSuiteProjectId] to [$($thisUpdatedOpp.parent.id)] for Opp [$($thisUpdatedOpp.tranId) $($thisUpdatedOpp.title)][$($thisUpdatedOpp.id)]"
                Write-Verbose "`tTrying: [$($termWithWrongProject.Name)][$($termWithWrongProject.Id)].CustomProperties.NetSuiteProjectId = [$($thisUpdatedOpp.ProjectId)]"
                $termWithWrongProject.Context.ExecuteQuery()
                }
            catch {
                Write-Error "Error updating Term [$($termWithWrongProject.Name)][$($termWithWrongProject.Id)].CustomProperties.NetSuiteProjectId = [$($thisUpdatedOpp.ProjectId)] in sync-netsuiteToManagedMetaData()"
                Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                [array]$doNotUpdateLastModified += $thisUpdatedOpp
                }
            }
        else{
             Write-Warning "Could not find corresponding Term for updated NetSuite Opp [$($termWithWrongProject.Name)][$($termWithWrongProject.Id)][$($thisUpdatedOpp.id)]"
             [array]$doNotUpdateLastModified += $thisUpdatedOpp
            }    
 
        }

    #############################
    #Update LastModifiedDate
    #############################
    Write-Host "`tRemaining ~[$($oppsToCheck.Count - $missingFromTermStore.Count - $oppsWithChangedNames.Count - $oppsWithChangedClient.Count - $oppsWithChangedProject.Count)] Opportunities must have had minor/irrelevant updates."
    Write-Host "`tUpdating lastmodified timestamps (if everything worked)"
    @($oppsToCheck | Select-Object) | % {
        $thisOppToUpdate = $_
        $thisOppLabel = "$($thisOppToUpdate.tranId) $($thisOppToUpdate.title)"
        if($doNotUpdateLastModified.Id -notcontains $thisOppToUpdate.Id){ #If the rename/merge didn't explictly fail, update the NetSuiteOppLastModifiedDate CustomProperty. This will update NetSuiteOppLastModifiedDate for all successful updates, all new Terms and all Opps that were updated in NetSuite but didn;t have Name changes
            $updatedTerm = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Identity $thisOppLabel -Includes CustomProperties, Labels
            Write-Host "`t`t[$($thisOppToUpdate.tranId) $($thisOppToUpdate.title)][$($thisOppToUpdate.id)] was processed successfully - updating NetSuiteLastModifiedDate to [$($thisOppToUpdate.lastModifiedDate)]"
            if($updatedTerm){ 
                $updatedTerm.SetCustomProperty("NetSuiteOppLastModifiedDate",$thisOppToUpdate.lastModifiedDate)
                Write-Verbose "`tTrying: [$($updatedTerm.Name)][$($updatedTerm.Id)].SetCustomProperty(NetSuiteOppLastModifiedDate,$($thisOppToUpdate.lastModifiedDate))"
                try{$updatedTerm.Context.ExecuteQuery()}
                catch{
                    Write-Error "Error setting CustomProperty NetSuiteOppLastModifiedDate = [$($thisOppToUpdate.lastModifiedDate)] on Term [$($updatedTerm.Name)][$($updatedTerm.Id)] in sync-netsuiteToManagedMetaData()"
                    Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                    }
                }
            }
        else{
            Write-Warning "Something went wrong with [$($thisOppToUpdate.Name)] - not updating NetSuiteLastModifiedDate"
            }
        }

    #endregion

    #region Projects
    Write-Host ""
    $pnpTermGroup = "Kimble"
    $pnpTermSet = "Projects"
    $allProjTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes CustomProperties | ? {$_.IsDeprecated -eq $false -and $(![string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteClientId))}
    $allProjTerms | Select-Object | % {
        Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteId -Value $_.CustomProperties.NetSuiteProjId -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name Name2 -Value $_.Name -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name ClientId -Value $_.CustomProperties.NetSuiteClientId -Force
        }

    [datetime]$lastProcessed = $($allProjTerms | sort {$_.CustomProperties.NetSuiteProjLastModifiedDate} | select -Last 1).CustomProperties.NetSuiteProjLastModifiedDate

    $netQuery =  "?q=lastModifiedDate ON_OR_AFTER `"$($(Get-Date $lastProcessed -Format g).Split(" ")[0])`"" #Excludes any Companies that haven;t been updated since X
    #$netQuery += " AND custentity_ant_projectsector IS_NOT `"Intercompany`"" #Excludes any Companies with "(intercompany project)" in the companyName
    [array]$projToCheck = get-netSuiteProjectFromNetSuite -query $netQuery -netsuiteParameters $(get-netSuiteParameters -connectTo Production) 
    $projToCheck = $projToCheck | ? {$_.custentity_ant_projectsector -ne "Intercompany"}   #Fix this after Go LIVE
    $projToCheck = $projToCheck | ? {$(Get-Date $_.lastModifiedDate) -ge $(Get-Date $lastProcessed)} #We can currently only filter by date in NetSuite, so filter again client-side to exclude all the other Projects we've processed earlier today
    [array]$processedAtExactlyLastTimestamp = $projToCheck | ? {$(Get-Date $_.lastModifiedDate) -ge $(Get-Date $lastProcessed)} #Find how many Opportunities match the $lastProcessed timestamp exactly
    if($processedAtExactlyLastTimestamp.Count -eq 1){$projToCheck = $projToCheck | ? {$projToCheck.id -notcontains $processedAtExactlyLastTimestamp[0].id}} #If it's exactly one, exclude it from processing (as we've already processed it on a previous cycle)
    
    Write-Host "Processing [$($projToCheck.Count)] Projects"
    #$projToCheck = get-netSuiteProjectFromNetSuite -netsuiteParameters $(get-netSuiteParameters -connectTo Production)    ##GET ALL PROJECTS
    $projToCheck | Select-Object | % { #Set the objects up so they are easy to compare-object
        Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteId -Value $_.id -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name Name2 -Value $_.entityid -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name ClientId -Value $_.parent.id -Force
        }

    #Match all updated NetSuite records against Managed Metadata using NetSuiteId
        #If no Id match, then create new Proj Term and flag for reprocessing
        #If Id match then check if Name has changed
            #If Name has not changed, do nothing
            #if Name has changed, rename the Term and flag for reprocessing
        #If Id match then check if associated ClientId has changed
            #If ClientId has not changed, do nothing
            #if ClientId has changed, change the NetSuiteClientId for the Term and flag for reprocessing
        #If nothing went wrong, update NetSuiteLastModifiedDate on the Term to exclude it from the future cycles (untilthe NetSuite record is updated again)

    [array]$doNotUpdateLastModified = @() #If anything goes wrong processing a Project, we don't want to update the NetSuiteProjLastModifiedDate CustomProperty on the Term as the mismatch means it will get picked up in the next Full Reconcile
    $deltaProjId = Compare-Object -ReferenceObject @($projToCheck | Select-Object) -DifferenceObject $allProjTerms -Property NetSuiteId -PassThru -IncludeEqual #Match all updated NetSuite records against Managed Metadata using NetSuiteId
    [array]$missingFromTermStore = $deltaProjId | ? {$_.SideIndicator -eq "<="}
    Write-Host "`tProcessing [$($missingFromTermStore.Count)] new Projects"
    @($missingFromTermStore | Select-Object) | % {
        $thisNewProj = $_
        $thisProjLabel = $thisNewProj.entityId
        Write-Host "`t`tProcessing new Project [$thisProjLabel][$($thisNewProj.id)] for Client ID [$($thisNewProj.parent.id)]"
        $testForCollision = $allProjTerms | ? {$_.Name -eq $thisNewProj.Name}
        if($testForCollision){
            Write-Warning "`tThere is already a term with the same default label and parent term [$($thisProjLabel)] - cannot create new Project Term."
            #If there is already a Term with the same name, re-use it
            if(![string]::IsNullOrWhiteSpace($testForCollision.CustomProperties.NetSuiteProjId) -and $testForCollision.CustomProperties.NetSuiteProjId -ne $thisNewProj.id){ #If the Term already has a _different_ NetSuiteProjId then somthing has gone badly wrong. We need to preserve this information so we can unpick it later, so we'll preserve the old NetSuiteOppId by suffixing it with _overwritten$i
                while(![string]::IsNullOrWhiteSpace($testForCollision.CustomProperties."NetSuiteProjId_overwritten$i")){ #Find the lowest number for merging without overwriting any pre-existing CustomProperties
                    $i++
                    }
                $testForCollision.SetCustomProperty("NetSuiteProjId_overwritten$i",$testForCollision.CustomProperties.NetSuiteProjId) #Add this CustomProperty back into the CustomProperties as a pseudo-backup
                $testForCollision.SetCustomProperty("NetSuiteProjId",$thisNewProj.id) #Set the correct NetSuiteOppId
                $testForCollision.SetCustomProperty("flagForReprocessing",$true) #Set the flag for reprocessing so this Term gets processed into SharePoint
                try{
                    Write-Host "`t`t`t`t`tReusing existing Term [$($testForCollision.Name)][$($testForCollision.Id)] and setting [NetSuiteProjId_overwritten$i] to [$($testForCollision.CustomProperties.NetSuiteProjId)], and [NetSuiteProjId] to [$($thisNewProj.id)]"
                    $testForCollision.Context.ExecuteQuery()
                    }
                catch{
                    Write-Error "Error `"backing up`" an old NetSuiteOppId value [$($testForCollision.CustomProperties.NetSuiteProjId))] to [NetSuiteOppId_overwritten$i] for Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisProjLabel)] in sync-netsuiteToManagedMetaData()"
                    Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                    [array]$doNotUpdateLastModified += $thisNewProj
                    }
                }
            }
        else{
            try{
                Write-Host "`t`t`tCreating new Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisProjLabel)][@{NetSuiteProjId=$($thisNewProj.id);NetSuiteProjLastModifiedDate=$($thisNewProj.lastModifiedDate);NetSuiteClientId=$($thisNewProj.parent.id)]"
                Write-Verbose "`tTrying: New-PnPTerm -TermGroup [$pnpTermGroup] -TermSet [$pnpTermSet] -Name [$($thisProjLabel)] -Lcid 1033 -CustomProperties @{NetSuiteProjId=$($thisNewProj.id);NetSuiteProjLastModifiedDate=$($thisNewProj.lastModifiedDate);NetSuiteClientId=$($thisNewProj.parent.id)}"
                $newTerm = New-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Name $thisProjLabel -Lcid 1033 -CustomProperties @{NetSuiteProjId=$thisNewProj.id;NetSuiteProjLastModifiedDate=$thisNewProj.lastModifiedDate;NetSuiteClientId=$thisNewProj.parent.id}
                }
            catch{
                Write-Error "Error creating new Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisProjLabel)][@{NetSuiteProjId=$($thisNewProj.id);NetSuiteProjLastModifiedDate=$($thisNewProj.lastModifiedDate);NetSuiteClientId=$($thisNewProj.parent.id)}] in sync-NetsuiteTpManagedMetaData()"
                Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                [array]$doNotUpdateLastModified += $thisNewProj
                }
            }
        }

    #############################
    #Update Projects (Names & ClientIDs)
    #############################
    $matchedId = $deltaProjId | ? {$_.SideIndicator -eq "=="}
    $matchedIdReversed = Compare-Object -ReferenceObject $allProjTerms -DifferenceObject @($matchedId | Select-Object) -Property NetSuiteId -PassThru -IncludeEqual -ExcludeDifferent #We then use $matchedId to filter only the Terms with corresponding $clientsToCheck records
        <#Sanity check - these should produce identical results, (but weirdly you have to run them separately). CSOM, eh?:
        $matchedId | sort NetSuiteId | select entityid -First 10
        $matchedIdReversed | sort NetSuiteId | select Name -First 10
        #>
    $deltaName = Compare-Object -ReferenceObject @($matchedId | Select-Object) -DifferenceObject @($matchedIdReversed | Select-Object) -Property NetSuiteId,Name2 -PassThru #We compare the two equal sets on both NetSuiteId and Name2 to see which pairs have mismatched Name values
    [array]$projectsWithChangedNames = $deltaName | ? {$_.SideIndicator -eq "<="}
    Write-Host "`tProcessing [$($projectsWithChangedNames.Count)] Projects with changed names"
    @($projectsWithChangedNames | Select-Object) | % {
        $thisUpdatedProject = $_
        Write-Host "`t`tProcessing updated Project [$($thisUpdatedProject.entityid)][$($thisUpdatedProject.id)]"
        $termWithWrongName = $matchedIdReversed | ? {$_.NetSuiteId -eq $thisUpdatedProject.NetSuiteId}
        if ($termWithWrongName.Count -eq 1){
            Write-Host "`t`t`tRenaming Term `t[$($termWithWrongName.Name)][$($termWithWrongName.Id)]"
            Write-Host "`t`t`tto:`t`t`t`t[$($thisUpdatedProject.entityid)]"
            $termWithWrongName_originalName = $termWithWrongName.Name
            $termWithWrongName.Name = $thisUpdatedProject.entityid
            try{
                $termWithWrongName.Context.ExecuteQuery()
                }
            catch {
                Write-Error "Error renaming Term [$($termWithWrongName_originalName)] to [$($thisUpdatedProject.entityid)] in sync-netsuiteToManagedMetaData()"
                Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                [array]$doNotUpdateLastModified += $thisUpdatedProject
                }
            }
        else{
             Write-Warning "Could not find corresponding Term for updated NetSuite Project [$($thisUpdatedProject.entityid)][$($thisUpdatedProject.id)]"
             [array]$doNotUpdateLastModified += $thisUpdatedProject
            }    
        }

    $deltaClientId = Compare-Object -ReferenceObject @($matchedId | Select-Object) -DifferenceObject @($matchedIdReversed | Select-Object) -Property NetSuiteId,ClientId -PassThru #We compare the two equal sets on both NetSuiteId and Name2 to see which pairs have mismatched Name values
    [array]$projectsWithChangedClient = $deltaClientId | ? {$_.SideIndicator -eq "<="}
    Write-Host "`tProcessing [$($projectsWithChangedClient.Count)] Projects with changed Clients"
    @($projectsWithChangedClient | Select-Object) | % {
        $thisUpdatedProject = $_
        Write-Host "`t`tProcessing updated Project [$($thisUpdatedProject.entityid)][$($thisUpdatedProject.id)]"
        $termWithWrongClient = $matchedIdReversed | ? {$_.NetSuiteId -eq $thisUpdatedProject.NetSuiteId}
        if ($termWithWrongClient.Count -eq 1){
            Write-Host "`t`t`tReassigning Project Term [$($termWithWrongClient.Name)][$($termWithWrongClient.Id)] to Client [$($thisUpdatedProject.parent.id)]"
            while(![string]::IsNullOrWhiteSpace($termWithWrongClient.CustomProperties."NetSuiteClientId_previous$i")){ #Find the lowest number for merging without overwriting anything
                $i++
                }
            $termWithWrongClient.SetCustomProperty("NetSuiteClientId_previous$i",$termWithWrongClient.CustomProperties.NetSuiteClientId)
            $termWithWrongClient.SetCustomProperty("NetSuiteClientId",$thisUpdatedProject.parent.id)
            try{
                Write-Host "`t`t`t`t`tSetting [(NetSuiteClientId_previous$i] to [$($termWithWrongClient.CustomProperties.NetSuiteClientId)], and [NetSuiteClientId] to [$($thisUpdatedProject.parent.id)] for Project [$($thisUpdatedProject.entityId)][$($thisUpdatedProject.id)]"
                $termWithWrongClient.Context.ExecuteQuery()
                }
            catch {
                Write-Error "Error reassigning Project Term [$($termWithWrongClient.Name)] to Client [$($thisUpdatedProject.parent.id)] in sync-netsuiteToManagedMetaData()"
                Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                [array]$doNotUpdateLastModified += $thisUpdatedProject
                }
            }
        else{
             Write-Warning "Could not find corresponding Term for updated NetSuite Project [$($thisUpdatedProject.entityid)][$($thisUpdatedProject.id)]"
             [array]$doNotUpdateLastModified += $thisUpdatedProject
            }    
        }



    #############################
    #Update LastModifiedDate
    #############################
    Write-Host "`tRemaining ~[$($projToCheck.Count - $missingFromTermStore.Count - $projectsWithChangedNames.Count - $projectsWithChangedClient.Count)] Projects must have had minor/irrelevant updates."
    Write-Host "`tUpdating lastmodified timestamps (if everything worked)"
    $projToCheck | % {
        $thisProjToUpdate = $_
        if($doNotUpdateLastModified -notcontains $thisProjToUpdate){ #If the rename/merge didn't explictly fail, update the NetSuiteOppLastModifiedDate CustomProperty. This will update NetSuiteOppLastModifiedDate for all successful updates, all new Terms and all Opps that were updated in NetSuite but didn;t have Name changes
            $updatedTerm = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Identity $thisProjToUpdate.entityId -Includes CustomProperties, Labels
            Write-Host "`t`t[$($thisProjToUpdate.entityid)][$($thisProjToUpdate.id)] was processed successfully - updating NetSuiteLastModifiedDate to [$($thisProjToUpdate.lastModifiedDate)]"
            if($updatedTerm){ 
                $updatedTerm.SetCustomProperty("NetSuiteProjLastModifiedDate",$thisProjToUpdate.lastModifiedDate)
                Write-Verbose "`tTrying: [$($updatedTerm.Name)][$($updatedTerm.Id)].SetCustomProperty(NetSuiteProjLastModifiedDate,$($thisProjToUpdate.lastModifiedDate))"
                try{$updatedTerm.Context.ExecuteQuery()}
                catch{
                    Write-Error "Error setting CustomProperty NetSuiteProjLastModifiedDate = [$($thisProjToUpdate.lastModifiedDate)] on Term [$($updatedTerm.Name)][$($updatedTerm.Id)] in sync-netsuiteToManagedMetaData()"
                    Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                    }
                }
            }
        else{
            Write-Warning "Something went wrong with [$($thisProjToUpdate.Name)] - not updating NetSuiteLastModifiedDate"
            }
        }
    
    #endregion

    }
Write-Host "sync-netsuiteToManagedMetaData completed in [$($fullSyncTime.TotalSeconds)] seconds"
Write-Host "***************************************************************************"
Stop-Transcript