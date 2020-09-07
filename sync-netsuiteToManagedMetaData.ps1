if($PSCommandPath){
    $InformationPreference = 2
    $VerbosePreference = 0
    $logFileLocation = "C:\ScriptLogs\"
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))`_Transcript_$(Get-Date -Format "yyyy-MM-dd").log"
    Start-Transcript $transcriptLogName -Append
    }

$sharePointAdmin = "kimblebot@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\KimbleBot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com" -Credentials $adminCreds

Write-Information "sync-netsuiteToManagedMetaData started at $(Get-Date -Format s)"
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
    $netQuery += " AND companyName CONTAIN_NOT `"(intercompany project)`"" #Excludes any Companies with "(intercompany project)" in the companyName
    $netQuery += " AND companyName START_WITH_NOT `"x `"" #Excludes any Companies that begin with "x " in the companyName
    $netQuery += " AND entityStatus ANY_OF_NOT [6, 7]" #Excludes LEAD-Unqualified and LEAD-Qualified (https://XXX.app.netsuite.com/app/crm/sales/customerstatuslist.nl?whence=)
    $netQuery += " AND lastModifiedDate ON_OR_AFTER `"$($(Get-Date $lastProcessed -Format g).Split(" ")[0])`"" #Excludes any Companies that haven;t been updated since X
    [array]$clientsToCheck = get-netSuiteClientsFromNetSuite -query $netQuery -netsuiteParameters $(get-netSuiteParameters -connectTo Production)
    Write-Information "Processing [$($clientsToCheck.Count)] Clients"
    #$clientsToCheck = $clientsToCheck | ? {$_.entityStatus.refName -notmatch "LEAD"} #Filter out Leads
    $clientsToCheck | % { #Set the objects up so they are easy to compare
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
    $missingFromMmd = $deltaClientId | ? {$_.SideIndicator -eq "<="}
    $missingFromMmd | % { #If no Id match, then create new Client Term and flag for reprocessing
        $thisNewClient = $_
        Write-Information "Creating new Client Term for [$($thisNewClient.companyName)]"
        $testForCollision = $allClientTerms | ? {$_.Name2 -eq $thisNewClient.Name2}
        if($testForCollision){#If Term exists, back up any out-of-date NetSuiteId value and re-use the Term and flag for reprocessing
            Write-Warning "There is already a term with the same default label and parent term [$($thisNewClient.companyName)] - cannot create new Client Term."
            if(![string]::IsNullOrWhiteSpace($testForCollision.CustomProperties.NetSuiteId) -and $testForCollision.CustomProperties.NetSuiteId -ne $thisNewClient.id){ #If the Term already has a _different_ NetSuiteId then somthing has gone badly wrong. We need to preserve this information so we can unpick it later, so we'll preserve the old NetSuiteId by suffixing it with _overwritten$i
                while(![string]::IsNullOrWhiteSpace($testForCollision.CustomProperties."NetSuiteId_overwritten$i")){ #Find the lowest number for merging without overwriting any pre-existing CustomProperties
                    $i++
                    }
                $testForCollision.SetCustomProperty("NetSuiteId_overwritten$i",$testForCollision.CustomProperties.NetSuiteId) #Add this CustomProperty back into the CustomProperties as a pseudo-backup
                $testForCollision.SetCustomProperty("NetSuiteId",$thisNewClient.id) #Set the correct NEtsuiteId
                $testForCollision.SetCustomProperty("flagForReprocessing",$true) #Set the flag for reprocessing so this Term gets processed into SharePoint
                try{
                    Write-Information "Reusing existing Term [$($testForCollision.Name)]and updating CCustomProperties"
                    Write-Verbose "`tTrying: [$($testForCollision.Name)].SetCustomProperty(NetSuiteId_overwritten$i,[$($testForCollision.CustomProperties.NetSuiteId)]) & SetCustomProperty(NetSuiteId,$($thisNewClient.id)) & SetCustomProperty(flagForReprocessing,$true)"
                    $testForCollision.Context.ExecuteQuery()
                    }
                catch{
                    Write-Error "Error `"backing up`" an old NetSuiteId value [$($testForCollision.CustomProperties.NetSuiteId))] to [NetSuiteId_overwritten$i] for Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisNewClient.companyName)] in sync-netsuiteToManagedMetaData()"
                    [array]$doNotUpdateLastModified += $thisNewClient
                    [array]$bigProblems += ,@($thisNewClient,$testForCollision,"Error `"backing up`" an old NetSuiteId value [$($testForCollision.CustomProperties.NetSuiteId))] to [NetSuiteId_overwritten$i] for Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisNewClient.companyName)] in sync-netsuiteToManagedMetaData()")
                    $_
                    continue #If we can't backup the odl NetSuiteId value, skip over of this Client
                    }
                }
            }
        else{#If Term does not exist, create new Client Term and flag for reprocessing
            try{
                Write-Information "Creating new Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisNewClient.companyName)][@{NetSuiteId=$($thisNewClient.id);NetSuiteLastModifiedDate=$($thisNewClient.lastModifiedDate);flagForReprocessing=$true]"
                Write-Verbose "`tTrying: New-PnPTerm -TermGroup [$pnpTermGroup] -TermSet [$pnpTermSet] -Name [$($thisNewClient.companyName)] -Lcid 1033 -CustomProperties @{NetSuiteId=$($thisNewClient.id);NetSuiteLastModifiedDate=$($thisNewClient.lastModifiedDate);flagForReprocessing=$true"
                $newTerm = New-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Name $thisNewClient.companyName -Lcid 1033 -CustomProperties @{NetSuiteId=$thisNewClient.id;NetSuiteLastModifiedDate=$thisNewClient.lastModifiedDate;flagForReprocessing=$true}
                }
            catch{
                Write-Error "Error creating new Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisNewClient.companyName)][@{NetSuiteId=$($thisNewClient.id);NetSuiteLastModifiedDate=$($thisNewClient.lastModifiedDate)] in sync-NetsuiteTpManagedMetaData()"
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
    $clientsWithChangedNames = $deltaName | ? {$_.SideIndicator -eq "<="}
    $clientsWithChangedNames | % {
        $thisUpdatedClient = $_
        Write-Information "Company name [$($thisUpdatedClient.companyName)][$($thisUpdatedClient.id)] seems to have changed. Investigating further."
        $termWithWrongName = $matchedIdReversed | ? {$_.NetSuiteId -eq $thisUpdatedClient.NetSuiteId}
        if ($termWithWrongName.Count -eq 1){
            Write-Information "`tRenaming Term [$($termWithWrongName.Name)][$($termWithWrongName.Id)] to [$($thisUpdatedClient.companyName)]"
            $termWithWrongName_originalName = $termWithWrongName.Name
            $termWithWrongName.Name = $thisUpdatedClient.companyName
            try{
                Write-Verbose "`tTrying: [$($termWithWrongName_originalName)].Name = [$($thisUpdatedClient.companyName)]"
                $termWithWrongName.Context.ExecuteQuery()
                }
            catch {
                if($_.Exception -match "TermStoreErrorCodeEx:There is already a term with the same default label and parent term."){
                    Write-Warning "There is already a term with the same default label and parent term [$($termWithWrongName_originalName)]->[$($thisUpdatedClient.companyName)]"
                    #If there is already a Term with the same name, merge the would-be-collision into this Term and preserve any conflicting CustomProperties by suffixing them with _merged$i
                    $termWithWrongName.Name = $termWithWrongName_originalName #We need to set this back in case something went wrong with a previous .Merge() and we need mess about with Labels
                    $duffTermToMergeIntoGoodTerm = $allClientTerms | ? {$_.Name2 -eq $thisUpdatedClient.Name2 -and $_.Id -ne $termWithWrongName.id}
                    if($duffTermToMergeIntoGoodTerm){ #If there's another Term, merge them
                        try{
                            Write-Information "`t`tMerging Terms -termToBeRetained [$($termWithWrongName.Name)] -termToBeMerged [$($duffTermToMergeIntoGoodTerm.Name)] -pnpTermGroup $pnpTermGroup -pnpTermSet $pnpTermSet"
                            Write-Verbose "`tTrying: merge-pnpTerms -termToBeRetained [$($termWithWrongName.Name)] -termToBeMerged [$($duffTermToMergeIntoGoodTerm.Name)] -setDefaultLabelTo Merged -pnpTermGroup $pnpTermGroup -pnpTermSet $pnpTermSet -Verbose:$VerbosePreference"
                            merge-pnpTerms -termToBeRetained $termWithWrongName -termToBeMerged $duffTermToMergeIntoGoodTerm -setDefaultLabelTo Merged -pnpTermGroup $pnpTermGroup -pnpTermSet $pnpTermSet -Verbose:$VerbosePreference
                            }
                        catch{
                            Write-Error "Error merging Term [$($pnpTermGroup)][$($pnpTermSet)][$($duffTermToMergeIntoGoodTerm.Name)] into [$($termWithWrongName.Name)] in sync-netsuiteToManagedMetaData()"
                            $_
                            [array]$doNotUpdateLastModified += $thisUpdatedClient
                            }
                        }
                    else{#If there isn't another Term, they've probably already been merged, so try relabelling it.
                        Write-Information "Setting default Label to [$($thisUpdatedClient.companyName)] for Term [$($termWithWrongName.Name)][$($termWithWrongName.Id)]"
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
                            try{
                                Write-Verbose "`tTrying: [$($updatedTerm.Name)].[$($correctLabel.Value)].SetAsDefaultForLanguage()"
                                $updatedTerm.Context.ExecuteQuery()
                                }
                            catch{
                                Write-Error "Error setting Default Label to [$($correctLabel.Value)] on Term [$($pnpTermGroup)][$($pnpTermSet)][$($updatedTerm.Name)] in sync-netsuiteToManagedMetaData()"
                                $_
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
                    Write-Information "Setting default Label to [$($thisUpdatedClient.companyName)] for Term [$($termWithWrongName.Name)][$($termWithWrongName.Id)]"
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
                        try{
                            Write-Verbose "`tTrying: [$($updatedTerm.Name)].[$($correctLabel.Value)].SetAsDefaultForLanguage()"
                            $updatedTerm.Context.ExecuteQuery()
                            }
                        catch{
                            Write-Error "Error setting Default Label to [$($correctLabel.Value)] on Term [$($pnpTermGroup)][$($pnpTermSet)][$($updatedTerm.Name)] in sync-netsuiteToManagedMetaData()"
                            $_
                            [array]$doNotUpdateLastModified += $thisUpdatedClient
                            }
                        $i++
                        $updatedTerm = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Identity $termWithWrongName.Id -Includes CustomProperties, Labels
                        }
                    while($updatedTerm.Name -eq $termWithWrongName_originalName)
                    }
                else{
                    Write-Error "Error renaming Term [$($termWithWrongName_originalName)] to [$($thisUpdatedClient.companyName)] in sync-netsuiteToManagedMetaData()"
                    #$_
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
    $clientsToCheck | % {
        $thisClientToUpdate = $_
        if($doNotUpdateLastModified.id -notcontains $thisClientToUpdate.id){ #If the rename/merge didn't explictly fail, update the NetSuiteLastModified CustomProperty. This will update NetSuiteLastModified for all successful updates, all new Terms and all Clients that were updated in NetSuite but didn;t have Name changes
            $updatedTerm = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Identity $thisClientToUpdate.companyName -Includes CustomProperties, Labels
            if($updatedTerm){ 
                $updatedTerm.SetCustomProperty("NetSuiteLastModifiedDate",$thisClientToUpdate.lastModifiedDate)
                Write-Information "[$($thisClientToUpdate.companyName)] was processed successfully - updating NetSuiteLastModifiedDate to [$($thisClientToUpdate.lastModifiedDate)]"
                try{
                    Write-Verbose "`tTrying: [$($updatedTerm.Name)][$($updatedTerm.Id)].SetCustomProperty(NetSuiteLastModifiedDate,$($thisClientToUpdate.lastModifiedDate))"
                    $updatedTerm.Context.ExecuteQuery()
                    }
                catch{
                    Write-Error "Error setting CustomProperty NetSuiteLastModifiedDate = [$($thisClientToUpdate.lastModifiedDate)] on Term [$($updatedTerm.Name)][$($updatedTerm.Id)] in sync-netsuiteToManagedMetaData()"
                    $_
                    }
                }
            }
        else{
            Write-Warning "Something went wrong with [$($thisClientToUpdate.Name)] - not updating NetSuiteLastModifiedDate"
            }
        }



    #endregion

    #region Opportunities
    $pnpTermGroup = "Kimble"
    $pnpTermSet = "Opportunities"
    $allOppTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes CustomProperties | ? {$_.IsDeprecated -eq $false}
    $allOppTerms | % {
        Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteId -Value $_.CustomProperties.NetSuiteOppId -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name Name2 -Value $_.Name -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name ClientId -Value $_.CustomProperties.NetSuiteClientId -Force
        }

    [datetime]$lastProcessed = $($allOppTerms | sort {$_.CustomProperties.NetSuiteOppLastModifiedDate} | select -Last 1).CustomProperties.NetSuiteOppLastModifiedDate

    $netQuery =  "?q=lastModifiedDate ON_OR_AFTER `"$($(Get-Date $lastProcessed -Format g).Split(" ")[0])`"" #Excludes any Opps that haven;t been updated since X
    [array]$oppsToCheck = get-netSuiteOpportunityFromNetSuite -query $netQuery -netsuiteParameters $(get-netSuiteParameters -connectTo Production) 
    Write-Information "Processing [$($oppsToCheck.Count)] Opportunities"
    #$oppsToCheck = get-netSuiteOpportunityFromNetSuite -netsuiteParameters $(get-netSuiteParameters -connectTo Production) 
    $oppsToCheck | % { #Set the objects up so they are easy to compare
        Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteId -Value $_.id -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name Name2 -Value "$($_.tranId) $($_.title)" -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name ClientId -Value $_.entity.id -Force
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
    $missingFromMmd = $deltaOppId | ? {$_.SideIndicator -eq "<="}
    $missingFromMmd | % {
        $thisNewOpp = $_
        $thisOppLabel = "$($thisNewOpp.tranId) $($thisNewOpp.title)"
        $testForCollision = $allOppTerms | ? {$_.Name -eq $thisNewOpp.Name}
        if($testForCollision){
            Write-Warning "There is already a term with the same default label and parent term [$($thisOppLabel)] - cannot create new Opportunity Term."
            #If there is already a Term with the same name, re-use it
            if(![string]::IsNullOrWhiteSpace($testForCollision.CustomProperties.NetSuiteOppId) -and $testForCollision.CustomProperties.NetSuiteOppId -ne $thisNewOpp.id){ #If the Term already has a _different_ NetSuiteOppId then somthing has gone badly wrong. We need to preserve this information so we can unpick it later, so we'll preserve the old NetSuiteOppId by suffixing it with _overwritten$i
                while(![string]::IsNullOrWhiteSpace($testForCollision.CustomProperties."NetSuiteOppId_overwritten$i")){ #Find the lowest number for merging without overwriting any pre-existing CustomProperties
                    $i++
                    }
                $testForCollision.SetCustomProperty("NetSuiteOppId_overwritten$i",$testForCollision.CustomProperties.NetSuiteOppId) #Add this CustomProperty back into the CustomProperties as a pseudo-backup
                try{
                    Write-Verbose "`tTrying: [$($testForCollision.Name)].SetCustomProperty(NetSuiteOppId_overwritten$i,[$($testForCollision.CustomProperties.NetSuiteOppId)])"
                    $testForCollision.Context.ExecuteQuery()
                    }
                catch{
                    Write-Error "Error `"backing up`" an old NetSuiteOppId value [$($testForCollision.CustomProperties.NetSuiteOppId))] to [NetSuiteOppId_overwritten$i] for Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisOppLabel)] in sync-netsuiteToManagedMetaData()"
                    [array]$doNotUpdateLastModified += $thisNewOpp
                    $_
                    }
                }
            }
        else{
            try{
                Write-Information "Creating new Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisOppLabel)][@{NetSuiteOppId=$($thisNewOpp.id);NetSuiteOppLastModifiedDate=$($thisNewOpp.lastModifiedDate);flagForReprocessing=$true]"
                Write-Verbose "`tTrying: New-PnPTerm -TermGroup [$pnpTermGroup] -TermSet [$pnpTermSet] -Name [$($thisOppLabel)] -Lcid 1033 -CustomProperties @{NetSuiteOppId=$($thisNewOpp.id);NetSuiteOppLastModifiedDate=$($thisNewOpp.lastModifiedDate);NetSuiteClientId=$($thisNewOpp.entity.id);NetSuiteProjectId=$($thisNewOpp.custbody_project_created.id);flagForReprocessing=$true}"
                $newTerm = New-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Name $thisOppLabel -Lcid 1033 -CustomProperties @{NetSuiteOppId=$thisNewOpp.id;NetSuiteOppLastModifiedDate=$thisNewOpp.lastModifiedDate;NetSuiteClientId=$thisNewOpp.entity.id;NetSuiteProjectId=$thisNewOpp.custbody_project_created.id;flagForReprocessing=$true}
                }
            catch{
                Write-Error "Error creating new Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisOppLabel)][@{NetSuiteOppId=$($thisNewOpp.id);NetSuiteOppLastModifiedDate=$($thisNewOpp.lastModifiedDate);NetSuiteClientId=$($thisNewOpp.entity.id);NetSuiteProjectId=$($thisNewOpp.custbody_project_created.id)}] in sync-NetsuiteTpManagedMetaData()"
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
        $matchedId | sort NetSuiteId | select entityid -First 10
        $matchedIdReversed | sort NetSuiteId | select Name -First 10
        #>
    $deltaName = Compare-Object -ReferenceObject @($matchedId | Select-Object) -DifferenceObject @($matchedIdReversed | Select-Object) -Property NetSuiteId,Name2 -PassThru #We compare the two equal sets on both NetSuiteId and Name2 to see which pairs have mismatched Name values
    $oppsWithChangedNames = $deltaName | ? {$_.SideIndicator -eq "<="}
    $oppsWithChangedNames | % {
        $thisUpdatedOpp = $_
        Write-Information "Project name [$($thisUpdatedOpp.entityid)][$($thisUpdatedOpp.id)] seems to have changed. Investigating further."
        $termWithWrongName = $matchedIdReversed | ? {$_.NetSuiteId -eq $thisUpdatedOpp.NetSuiteId}
        if ($termWithWrongName.Count -eq 1){
            Write-Verbose "Renaming Term [$($termWithWrongName.Name)][$($termWithWrongName.Id)] to [$($thisUpdatedOpp.entityid)]"
            $termWithWrongName_originalName = $termWithWrongName.Name
            $termWithWrongName.Name = $thisUpdatedOpp.entityid
            try{
                Write-Verbose "`tTrying: [$($termWithWrongName_originalName)].Name = [$($thisUpdatedOpp.entityid)]"
                $termWithWrongName.Context.ExecuteQuery()
                }
            catch {
                Write-Error "Error renaming Term [$($termWithWrongName_originalName)] to [$($thisUpdatedOpp.entityid)] in sync-netsuiteToManagedMetaData()"
                [array]$doNotUpdateLastModified += $thisUpdatedOpp
                }
            }
        else{
             Write-Warning "Could not find corresponding Term for updated NetSuite Opp [$($thisUpdatedOpp.entityid)][$($thisUpdatedOpp.id)]"
             [array]$doNotUpdateLastModified += $thisUpdatedOpp
            }    
        }

    $deltaClientId = Compare-Object -ReferenceObject @($matchedId | Select-Object) -DifferenceObject @($matchedIdReversed | Select-Object) -Property NetSuiteId,ClientId -PassThru #We compare the two equal sets on both NetSuiteId and Name2 to see which pairs have mismatched Name values
    $oppsWithChangedClient = $deltaClientId | ? {$_.SideIndicator -eq "<="}
    $oppsWithChangedClient | % {
        $thisUpdatedOpp = $_
        Write-Verbose "Project [$($thisUpdatedOpp.entityid)][$($thisUpdatedOpp.id)] seems to have been assigned to a new Client. Investigating further."
        $termWithWrongClient = $matchedIdReversed | ? {$_.NetSuiteId -eq $thisUpdatedOpp.NetSuiteId}
        if ($termWithWrongClient.Count -eq 1){
            Write-Verbose "Reassigning Project Term [$($termWithWrongClient.Name)][$($termWithWrongClient.Id)] to Client [$($thisUpdatedOpp.parent.id)]"
            while(![string]::IsNullOrWhiteSpace($termWithWrongClient.CustomProperties."NetSuiteClientId_previous$i")){ #Find the lowest number for merging without overwriting anything
                $i++
                }
            $termWithWrongClient.SetCustomProperty("NetSuiteClientId_previous$i",$termWithWrongClient.CustomProperties.NetSuiteClientId)
            $termWithWrongClient.SetCustomProperty("NetSuiteClientId",$thisUpdatedOpp.parent.id)
            try{
                Write-Verbose "`tTrying: `$termWithWrongClient.SetCustomProperty(NetSuiteClientId_previous$i,$($termWithWrongClient.CustomProperties.NetSuiteClientId)) & `$termWithWrongClient.SetCustomProperty(NetSuiteClientId,$($thisUpdatedOpp.parent.id))"
                $termWithWrongClient.Context.ExecuteQuery()
                }
            catch {
                Write-Error "Error reassigning Project Term [$($termWithWrongClient.Name)] to Client [$($thisUpdatedOpp.parent.id)] in sync-netsuiteToManagedMetaData()"
                [array]$doNotUpdateLastModified += $thisUpdatedOpp
                }
            }
        else{
             Write-Warning "Could not find corresponding Term for updated NetSuite Project [$($thisUpdatedOpp.entityid)][$($thisUpdatedOpp.id)]"
             [array]$doNotUpdateLastModified += $thisUpdatedOpp
            }    
        }
    #############################
    #Update LastModifiedDate
    #############################
    $oppsToCheck | % {
        $thisOppToUpdate = $_
        $thisOppLabel = "$($thisOppToUpdate.tranId) $($thisOppToUpdate.title)"
        if($doNotUpdateLastModified.Id -notcontains $thisOppToUpdate.Id){ #If the rename/merge didn't explictly fail, update the NetSuiteOppLastModifiedDate CustomProperty. This will update NetSuiteOppLastModifiedDate for all successful updates, all new Terms and all Opps that were updated in NetSuite but didn;t have Name changes
            $updatedTerm = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Identity $thisOppLabel -Includes CustomProperties, Labels
            if($updatedTerm){ 
                $updatedTerm.SetCustomProperty("NetSuiteOppLastModifiedDate",$thisOppToUpdate.lastModifiedDate)
                Write-Verbose "`tTrying: [$($updatedTerm.Name)][$($updatedTerm.Id)].SetCustomProperty(NetSuiteOppLastModifiedDate,$($thisOppToUpdate.lastModifiedDate))"
                try{$updatedTerm.Context.ExecuteQuery()}
                catch{
                    Write-Error "Error setting CustomProperty NetSuiteOppLastModifiedDate = [$($thisOppToUpdate.lastModifiedDate)] on Term [$($updatedTerm.Name)][$($updatedTerm.Id)] in sync-netsuiteToManagedMetaData()"
                    $_
                    }
                }
            }
    
        }

    #endregion

    #region Projects
    $pnpTermGroup = "Kimble"
    $pnpTermSet = "Projects"
    $allProjTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes CustomProperties | ? {$_.IsDeprecated -eq $false -and $(![string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteClientId))}
    $allProjTerms | % {
        Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteId -Value $_.CustomProperties.NetSuiteProjId -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name Name2 -Value $_.Name -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name ClientId -Value $_.CustomProperties.NetSuiteClientId -Force
        }

    [datetime]$lastProcessed = $($allProjTerms | sort {$_.CustomProperties.NetSuiteProjLastModifiedDate} | select -Last 1).CustomProperties.NetSuiteProjLastModifiedDate

    $netQuery =  "?q=lastModifiedDate ON_OR_AFTER `"$($(Get-Date $lastProcessed -Format g).Split(" ")[0])`"" #Excludes any Companies that haven;t been updated since X
    #$netQuery += " AND custentity_ant_projectsector IS_NOT `"Intercompany`"" #Excludes any Companies with "(intercompany project)" in the companyName
    [array]$projToCheck = get-netSuiteProjectFromNetSuite -query $netQuery -netsuiteParameters $(get-netSuiteParameters -connectTo Production) 
    Write-Information "Processing [$($projToCheck.Count)] Projects"
    #$projToCheck = get-netSuiteProjectFromNetSuite -netsuiteParameters $(get-netSuiteParameters -connectTo Production)    ##GET ALL PROJECTS
    $projToCheck = $projToCheck | ? {$_.custentity_ant_projectsector -ne "Intercompany"}   #Fix this after Go LIVE
    $projToCheck | % { #Set the objects up so they are easy to compare-object
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
    $missingFromMmd = $deltaProjId | ? {$_.SideIndicator -eq "<="}
    $missingFromMmd | % {
        $thisNewProj = $_
        $thisProjLabel = $thisNewProj.entityId
        $testForCollision = $allProjTerms | ? {$_.Name -eq $thisNewProj.Name}
        if($testForCollision){
            Write-Warning "There is already a term with the same default label and parent term [$($thisProjLabel)] - cannot create new Project Term."
            #If there is already a Term with the same name, re-use it
            if(![string]::IsNullOrWhiteSpace($testForCollision.CustomProperties.NetSuiteProjId) -and $testForCollision.CustomProperties.NetSuiteProjId -ne $thisNewProj.id){ #If the Term already has a _different_ NetSuiteProjId then somthing has gone badly wrong. We need to preserve this information so we can unpick it later, so we'll preserve the old NetSuiteOppId by suffixing it with _overwritten$i
                while(![string]::IsNullOrWhiteSpace($testForCollision.CustomProperties."NetSuiteProjId_overwritten$i")){ #Find the lowest number for merging without overwriting any pre-existing CustomProperties
                    $i++
                    }
                $testForCollision.SetCustomProperty("NetSuiteProjId_overwritten$i",$testForCollision.CustomProperties.NetSuiteProjId) #Add this CustomProperty back into the CustomProperties as a pseudo-backup
                try{
                    Write-Verbose "`tTrying: [$($testForCollision.Name)].SetCustomProperty(NetSuiteProjId_overwritten$i,[$($testForCollision.CustomProperties.NetSuiteProjId)])"
                    $testForCollision.Context.ExecuteQuery()
                    }
                catch{
                    Write-Error "Error `"backing up`" an old NetSuiteOppId value [$($testForCollision.CustomProperties.NetSuiteProjId))] to [NetSuiteOppId_overwritten$i] for Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisProjLabel)] in sync-netsuiteToManagedMetaData()"
                    [array]$doNotUpdateLastModified += $thisNewProj
                    $_
                    }
                }
            }
        else{
            try{
                Write-Verbose "Creating new Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisProjLabel)][@{NetSuiteProjId=$($thisNewProj.id);NetSuiteProjLastModifiedDate=$($thisNewProj.lastModifiedDate);NetSuiteClientId=$($thisNewProj.parent.id)]"
                Write-Verbose "`tTrying: New-PnPTerm -TermGroup [$pnpTermGroup] -TermSet [$pnpTermSet] -Name [$($thisProjLabel)] -Lcid 1033 -CustomProperties @{NetSuiteProjId=$($thisNewProj.id);NetSuiteProjLastModifiedDate=$($thisNewProj.lastModifiedDate);NetSuiteClientId=$($thisNewProj.parent.id)}"
                $newTerm = New-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Name $thisProjLabel -Lcid 1033 -CustomProperties @{NetSuiteProjId=$thisNewProj.id;NetSuiteProjLastModifiedDate=$thisNewProj.lastModifiedDate;NetSuiteClientId=$thisNewProj.parent.id}
                }
            catch{
                Write-Error "Error creating new Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisProjLabel)][@{NetSuiteProjId=$($thisNewProj.id);NetSuiteProjLastModifiedDate=$($thisNewProj.lastModifiedDate);NetSuiteClientId=$($thisNewProj.parent.id)}] in sync-NetsuiteTpManagedMetaData()"
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
    $projectsWithChangedNames = $deltaName | ? {$_.SideIndicator -eq "<="}
    $projectsWithChangedNames | % {
        $thisUpdatedProject = $_
        $doNotProceed = $false
        Write-Verbose "Project name [$($thisUpdatedProject.entityid)][$($thisUpdatedProject.id)] seems to have changed. Investigating further."
        $termWithWrongName = $matchedIdReversed | ? {$_.NetSuiteId -eq $thisUpdatedProject.NetSuiteId}
        if ($termWithWrongName.Count -eq 1){
            Write-Verbose "Renaming Term [$($termWithWrongName.Name)][$($termWithWrongName.Id)] to [$($thisUpdatedProject.entityid)]"
            $termWithWrongName_originalName = $termWithWrongName.Name
            $termWithWrongName.Name = $thisUpdatedProject.entityid
            try{
                Write-Verbose "`tTrying: [$($termWithWrongName_originalName)].Name = [$($thisUpdatedProject.entityid)]"
                $termWithWrongName.Context.ExecuteQuery()
                }
            catch {
                Write-Error "Error renaming Term [$($termWithWrongName_originalName)] to [$($thisUpdatedProject.entityid)] in sync-netsuiteToManagedMetaData()"
                [array]$doNotUpdateLastModified += $thisUpdatedProject
                }
            }
        else{
             Write-Warning "Could not find corresponding Term for updated NetSuite Project [$($thisUpdatedProject.entityid)][$($thisUpdatedProject.id)]"
             [array]$doNotUpdateLastModified += $thisUpdatedProject
            }    
        }

    $deltaClientId = Compare-Object -ReferenceObject @($matchedId | Select-Object) -DifferenceObject @($matchedIdReversed | Select-Object) -Property NetSuiteId,ClientId -PassThru #We compare the two equal sets on both NetSuiteId and Name2 to see which pairs have mismatched Name values
    $projectsWithChangedClient = $deltaClientId | ? {$_.SideIndicator -eq "<="}
    $projectsWithChangedClient | % {
        $thisUpdatedProject = $_
        Write-Verbose "Project [$($thisUpdatedProject.entityid)][$($thisUpdatedProject.id)] seems to have been assigned to a new Client. Investigating further."
        $termWithWrongClient = $matchedIdReversed | ? {$_.NetSuiteId -eq $thisUpdatedProject.NetSuiteId}
        if ($termWithWrongClient.Count -eq 1){
            Write-Verbose "Reassigning Project Term [$($termWithWrongClient.Name)][$($termWithWrongClient.Id)] to Client [$($thisUpdatedProject.parent.id)]"
            while(![string]::IsNullOrWhiteSpace($termWithWrongClient.CustomProperties."NetSuiteClientId_previous$i")){ #Find the lowest number for merging without overwriting anything
                $i++
                }
            $termWithWrongClient.SetCustomProperty("NetSuiteClientId_previous$i",$termWithWrongClient.CustomProperties.NetSuiteClientId)
            $termWithWrongClient.SetCustomProperty("NetSuiteClientId",$thisUpdatedProject.parent.id)
            try{
                Write-Verbose "`tTrying: `$termWithWrongClient.SetCustomProperty(NetSuiteClientId_previous$i,$($termWithWrongClient.CustomProperties.NetSuiteClientId)) & `$termWithWrongClient.SetCustomProperty(NetSuiteClientId,$($thisUpdatedProject.parent.id))"
                $termWithWrongClient.Context.ExecuteQuery()
                }
            catch {
                Write-Error "Error reassigning Project Term [$($termWithWrongClient.Name)] to Client [$($thisUpdatedProject.parent.id)] in sync-netsuiteToManagedMetaData()"
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
    $projToCheck | % {
        $thisProjToUpdate = $_
        if($doNotUpdateLastModified -notcontains $thisProjToUpdate){ #If the rename/merge didn't explictly fail, update the NetSuiteOppLastModifiedDate CustomProperty. This will update NetSuiteOppLastModifiedDate for all successful updates, all new Terms and all Opps that were updated in NetSuite but didn;t have Name changes
            $updatedTerm = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Identity $thisProjToUpdate.entityId -Includes CustomProperties, Labels
            if($updatedTerm){ 
                $updatedTerm.SetCustomProperty("NetSuiteProjLastModifiedDate",$thisProjToUpdate.lastModifiedDate)
                Write-Verbose "`tTrying: [$($updatedTerm.Name)][$($updatedTerm.Id)].SetCustomProperty(NetSuiteProjLastModifiedDate,$($thisProjToUpdate.lastModifiedDate))"
                try{$updatedTerm.Context.ExecuteQuery()}
                catch{
                    Write-Error "Error setting CustomProperty NetSuiteProjLastModifiedDate = [$($thisProjToUpdate.lastModifiedDate)] on Term [$($updatedTerm.Name)][$($updatedTerm.Id)] in sync-netsuiteToManagedMetaData()"
                    $_
                    }
                }
            }
    
        }
    
    #endregion

    }
Write-Information "sync-netsuiteToManagedMetaData completed in [$($fullSyncTime.TotalSeconds)] seconds"
Stop-Transcript