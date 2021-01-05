
#region Get HubSpot Companies
$apiKey = get-hubSpotApiKey
$filterCompanyFlaggedForSync = [ordered]@{
    propertyName="netsuite_sync_company_"
    operator="HAS_PROPERTY"
    }
$filterExcludeCompaniesCalledAnthesis = [ordered]@{
    propertyName="name"
    operator="NOT_CONTAINS_TOKEN"
    value="Anthesis"
    }
$hubSortMaxlastmodifiedinhubspotbysync = [ordered]@{
    propertyName = "lastmodifiedinhubspot"
    direction = "DESCENDING"
    }
$hubspotCompanyMaxlastmodifiedinhubspotbysync = get-hubSpotObjects -apiKey $apiKey.HubApiKey -objectType companies -filterGroup1 @{filters=@($filterCompanyFlaggedForSync)} -sortPropertyNameAndDirection $hubSortMaxlastmodifiedinhubspotbysync -pageSize 1 -firstPageOnly

$filterCompanyUpdatedSinceLastSync = [ordered]@{
    propertyName="hs_lastmodifieddate"
    operator="GT"
    #value = [Math]::Floor([decimal](Get-Date(Get-Date "2000-10-20T08:34:48.887Z").ToUniversalTime()-uformat "%s"))*1000 #Convert to UNIX Epoch time and add Milliseconds
    value = [Math]::Floor([decimal](Get-Date(Get-Date $hubspotCompanyMaxlastmodifiedinhubspotbysync.properties.lastmodifiedinhubspot).ToUniversalTime()-uformat "%s"))*1000 #Convert to UNIX Epoch time and add Milliseconds
    }
[array]$hubSpotCompaniesToCheck = get-hubspotObjects -apiKey $apiKey.HubApiKey -objectType companies -filterGroup1 @{filters=@($filterCompanyFlaggedForSync,$filterCompanyUpdatedSinceLastSync,$filterExcludeCompaniesCalledAnthesis)} -pageSize 100 #-firstPageOnly 
$hubSpotCompaniesToCheck | Select-Object | % { #Prep the objects for compare-object later
    Add-Member -InputObject $_ -MemberType NoteProperty -Name HubSpotId -Value $_.Id
    Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteId -Value $_.properties.netsuiteid
    }
#endregion

#region Get NetSuite Companies
$hubSortLastModifiedInNetSuite = [ordered]@{
    propertyName = "lastmodifiedinnetsuite"
    direction = "DESCENDING"
    }
$hubspotCompanyMaxLastModifiedInNetSuite = get-hubSpotObjects -apiKey $apiKey.HubApiKey -objectType companies -filterGroup1 @{filters=@($filterCompanyFlaggedForSync)} -sortPropertyNameAndDirection $hubSortLastModifiedInNetSuite -pageSize 1 -firstPageOnly

$netQuery =  "?q=companyName CONTAIN_NOT `"Anthesis`"" #Excludes any Companies with "Anthesis" in the companyName
$netQuery += " AND companyName CONTAIN_NOT `"intercompany project`"" #Excludes any Companies with "(intercompany project)" in the companyName
$netQuery += " AND companyName START_WITH_NOT `"x `"" #Excludes any Companies that begin with "x " in the companyName
$netQuery += " AND lastModifiedDate ON_OR_AFTER `"$($(Get-Date $hubspotCompanyMaxLastModifiedInNetSuite.properties.lastmodifiedinnetsuite -Format g))`"" #Excludes any Companies that haven't been updated since X
$netSuiteParameters = get-netSuiteParameters -connectTo Production
[array]$netSuiteCompaniesToCheck = get-netSuiteClientsFromNetSuite -query $netQuery -netsuiteParameters $netSuiteParameters #-Verbose
$netSuiteCompaniesToCheck | % { #Prep the objects for compare-object later
    Add-Member -InputObject $_ -MemberType NoteProperty -Name HubSpotId -Value $_.custentitycustentity_hubspotid
    Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteId -Value $_.Id
    }
#endregion


#region Process changes, starting with NetSuite companies
$netSuiteCompaniesToCheck | % {
    $thisNetSuiteCompany = $_
    if([string]::IsNullOrWhiteSpace($thisNetSuiteCompany.HubSpotId)){ #$netSuite.HubSpotId -eq $null
        #***Create new record in HubSpot
        #Try to match to generic company e-mail address (and link to that)
        if([string]::IsNullOrWhiteSpace($thisNetSuiteCompany.email) -or $thisNetSuiteCompany.email -match "@anthesisgroup.com" -or $thisNetSuiteCompany.email -match "@lavola.com"){
            Write-Host -ForegroundColor Yellow "Unlinked NetSuite company found [$($thisNetSuiteCompany.companyName)][$($thisNetSuiteCompany.id)], but no generic e-mail address set. No meaningful way to identify this company in HubSpot."
            $matchedHubspotCompany = $null
            }
        else{
            Write-Host -ForegroundColor Yellow "Unlinked NetSuite company found [$($thisNetSuiteCompany.companyName)][$($thisNetSuiteCompany.id)] - looking for match in HubSpot"
            $matchedHubspotCompany = $hubSpotCompaniesToCheck | ? {$_.properties.generic_email_address__c -eq $thisNetSuiteCompany.email}
            if($matchedHubspotCompany){Write-Host -ForegroundColor DarkYellow "`tGeneric e-mail address [$($thisNetSuiteCompany.email)] found in HubSpot records marked for Sync [$($matchedHubspotCompany.properties.name)][$($matchedHubspotCompany.id)]"}
            else{
                Write-Host -ForegroundColor DarkYellow "`tGeneric e-mail address [$($thisNetSuiteCompany.email)] not found in HubSpot records marked for Sync - searching all HubSpot records"
                $filterCompanyGenericEmailEq = [ordered]@{
                    propertyName="generic_email_address__c"
                    operator="EQ"
                    value=$thisNetSuiteCompany.email
                    }
                $matchedHubspotCompany = get-hubSpotObjects -apiKey $apiKey.HubApiKey -objectType companies -filterGroup1 @{filters=@($filterCompanyGenericEmailEq)}

                #Try to match to any contact e-mail address and deduce Company (and link to that)
                if($matchedHubspotCompany){Write-Host -ForegroundColor DarkYellow "`tGeneric e-mail address [$($thisNetSuiteCompany.email)] found in HubSpot [$($matchedHubspotCompany.properties.name)][$($matchedHubspotCompany.id)]"}
                else{
                    Write-Host -ForegroundColor DarkYellow "`tGeneric e-mail address [$($thisNetSuiteCompany.email)] not found in HubSpot either - searching Contacts"
                    $filterContactEmailEq = [ordered]@{
                        propertyName="email"
                        operator="EQ"
                        value=$thisNetSuiteCompany.email
                        }
                    $matchedHubspotContact = get-hubSpotObjects -apiKey $apiKey.HubApiKey -objectType contacts -filterGroup1 @{filters=@($filterContactEmailEq)}
                    if(![string]::IsNullOrWhiteSpace($matchedHubspotContact.associatedcompanyid)){
                        $filterCompanyIdEq = [ordered]@{
                            propertyName="hs_object_id"
                            operator="EQ"
                            value=$matchedHubspotContact.associatedcompanyid
                            }
                        $matchedHubspotCompany = get-hubSpotObjects -apiKey $apiKey.HubApiKey -objectType companies -filterGroup1 @{filters=@($filterCompanyIdEq)}
                        }

                    #Have one last stab via domain name
                    if($matchedHubspotCompany){Write-Host -ForegroundColor DarkYellow "`tGeneric e-mail address [$($thisNetSuiteCompany.email)] matched to a Contact in HubSpot. Contact's Company is: [$($matchedHubspotCompany.properties.name)][$($matchedHubspotCompany.id)]"}
<#                    elseif(![string]::IsNullOrWhiteSpace($thisNetSuiteCompany.email) -and $thisNetSuiteCompany.email -match "@"){
                        Write-Host -ForegroundColor DarkYellow "`tGeneric e-mail address [$($thisNetSuiteCompany.email)] not found in any HubSpot Contacts either - seaching domain names"
                        $filterCompanyDomainEq = [ordered]@{
                            propertyName="domain"
                            operator="EQ"
                            value=$($thisNetSuiteCompany.email.Split("@")[1])
                            }
                        [array]$wildlyOptimisticSearch = get-hubSpotObjects -apiKey $apiKey.HubApiKey -objectType companies -filterGroup1 @{filters=@($filterCompanyDomainEq)}
                        if($wildlyOptimisticSearch.Count -eq 1){
                            $matchedHubspotCompany = $wildlyOptimisticSearch[0]
                            Write-Host -ForegroundColor DarkYellow "`tMatched to HubSpot company [$($matchedHubspotCompany.properties.name)][$($matchedHubspotCompany.id)] on the rather flimsy basis that the generic e-mail address [$($thisNetSuiteCompany.email)] in NetSuite matches the domain name [$($matchedHubspotCompany.properties.domain)] in HubSpot"
                            } 
                        else{
                            $matchedHubspotCompany = $null
                            Write-Host -ForegroundColor DarkYellow "`tNo match for domain names either :("
                            }#If we've not matched _exactly_ one company, abort this attempt
                        }#>
                    }
                }
            }
                
        if($matchedHubspotCompany){#Update and cross-reference both companies (if we've matched them)
            Write-Host -ForegroundColor Yellow "`tCross-referencing NetSuite Company [$($thisNetSuiteCompany.companyName)][$($thisNetSuiteCompany.id)] with HubSpot Company [$($matchedHubspotCompany.properties.name)][$($matchedHubspotCompany.id)]"
            try{$updatedHubSpotCompany = update-hubSpotObject -apiKey $apiKey.HubApiKey -objectType companies -objectId $matchedHubspotCompany.id -fieldHash @{netsuiteid=$thisNetSuiteCompany.id; lastmodifiedinnetsuite=$(get-dateInIsoFormat -dateTime $(Get-Date $thisNetSuiteCompany.lastModifiedDate).AddYears(-100) -precision Seconds);lastmodifiedinhubspot=$(get-dateInIsoFormat -dateTime $(Get-Date) -precision Ticks)}}
            catch{Write-Error $_}
            try{$updatedNetSuiteCompany = update-netSuiteClientInNetSuite -netSuiteClientId $thisNetSuiteCompany.id -fieldHash @{custentitycustentity_hubspotid = $matchedHubspotCompany.id} -netsuiteParameters $netSuiteParameters}
            catch [System.Net.WebException]{
                $json = ConvertFrom-Json -InputObject $_.ErrorDetails.Message
                if($json.status -eq 400){
                    Write-Host -ForegroundColor Red "`t$($json.'o:errorDetails'.detail) [$($thisNetSuiteCompany.companyName)][$($thisNetSuiteCompany.id)]"
                    }
                    #}
                else{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
                }
            catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
            }
        else{ #Create new Company in HubSpot (and link to that)
            Write-Host -ForegroundColor Yellow "`tCorresponding HubSpot Company for [$($thisNetSuiteCompany.companyName)][$($thisNetSuiteCompany.id)] could not be identified - creating a new HubSpot Company"
            try{$newHubspotCompany = new-hubspotCompanyFromNetsuiteCompany -apiKey $apiKey.HubApiKey -netSuiteCompany $thisNetSuiteCompany}
            catch{Write-Error $_}
            try{$updatedNetSuiteCompany = update-netSuiteClientInNetSuite -netSuiteClientId $thisNetSuiteCompany.id -fieldHash @{custentitycustentity_hubspotid = $newHubspotCompany.id} -netsuiteParameters $netSuiteParameters}
            catch [System.Net.WebException]{
                $json = ConvertFrom-Json -InputObject $_.ErrorDetails.Message
                if($json.status -eq 400){
                    Write-Host -ForegroundColor Red "`t$($json.'o:errorDetails'.detail) [$($thisNetSuiteCompany.companyName)][$($thisNetSuiteCompany.id)]"
                    }
                    #}
                else{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
                }
            catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
            }

        }
    else{#Match to HubSpot on ForeignKey (HubSpotID)
        Write-Host -ForegroundColor Cyan "Linked NetSuite company found [$($thisNetSuiteCompany.companyName)][$($thisNetSuiteCompany.id)] - retrieving HubSpot Company with Id [$($thisNetSuiteCompany.HubSpotId)]"
        $correspondingHubSpotCompany = Compare-Object -ReferenceObject $hubSpotCompaniesToCheck -DifferenceObject $thisNetSuiteCompany -Property HubSpotId -ExcludeDifferent -IncludeEqual -PassThru
        if($correspondingHubSpotCompany){
            Write-Host -ForegroundColor DarkCyan "`tCorresponding HubSpot Company [$($correspondingHubSpotCompany.properties.name)][$($correspondingHubSpotCompany.id)] retrieved from Companies flagged to Sync"
            }
        else{ #Error checking
            Write-Host -ForegroundColor DarkCyan "`tHubSpot Company not retrieved from Companies flagged to Sync - trying HubSpot"
            $hubspotFilterThisId = new-hubSpotFilterById -hubSpotId $thisNetSuiteCompany.HubSpotId
            $correspondingHubSpotCompany = get-hubSpotObjects -apiKey $apiKey.HubApiKey -objectType companies -filterGroup1 @{filters=@($hubspotFilterThisId)}
            if($correspondingHubSpotCompany){Write-Host -ForegroundColor DarkCyan "`tCorresponding HubSpot Company [$($correspondingHubSpotCompany.properties.name)][$($correspondingHubSpotCompany.id)] retrieved from HubSpot"}
            elseif(!$correspondingHubSpotCompany){
                Write-Warning "HubSpot company [$($thisNetSuiteCompany.HubSpotId)] (NetSuite name:[$($thisNetSuiteCompany.companyName)][$($thisNetSuiteCompany.NetSuiteId)]) could not be retrieved from HubSpot - it may have been deleted from HubSpot?"
                return #Break out of current $thisNetSuiteCompany loop if no $correspondingHubSpotCompany exists
                }
            elseif($correspondingHubSpotCompany.properties.netsuite_sync_company_ -eq $false){
                Write-Warning "HubSpot company [$($correspondingHubSpotCompany.properties.name)][$($thisNetSuiteCompany.HubSpotId)] (NetSuite name:[$($thisNetSuiteCompany.companyName)][$($thisNetSuiteCompany.NetSuiteId)]) is not flagged with {netsuite_sync_company_} - not syncing chamges back to HubSpot"
                return #Break out of current $thisNetSuiteCompany loop if $correspondingHubSpotCompany is not flagged netsuite_sync_company_
                }
            }
        if($thisNetSuiteCompany.lastModifiedDate -ne $correspondingHubSpotCompany.properties.lastmodifiedinnetsuite){ #Has the NetSuite object been updated since the last sync?
            Write-Host -ForegroundColor DarkCyan "`tNetSuite company has been modified."
            if($correspondingHubSpotCompany.updatedAt -ne $correspondingHubSpotCompany.properties.lastmodifiedinhubspot){ #Has the HubSpot object been updated since the last sync?
                Write-Host -ForegroundColor DarkCyan "`tHubSpot company has also been modified."
                #Both objects modified - in the event of a conflict, Leads are overwritten by HubSpot and Prospects/Clients are overwritten by NetSuite
                if($thisNetSuiteCompany.entityStatus.refName -match "LEAD"){
                    #***Update NetSuite record based on HubSpot data
                    Write-Host -ForegroundColor Cyan "NetSuite Client [$($thisNetSuiteCompany.companyName)][$($thisNetSuiteCompany.id)] is set to [$($thisNetSuiteCompany.entityStatus.refName)] - updating NetSuite object based on HubSpot object"
                    try{
                        $updatedNetSuiteCompany = update-netSuiteClientFromHubSpotObject -hubSpotCompanyObject $correspondingHubSpotCompany -netsuiteParameters $netSuiteParameters -hubSpotApiKey $apiKey.HubApiKey
                        }
                    catch [System.Net.WebException]{
                        $json = ConvertFrom-Json -InputObject $_.ErrorDetails.Message
                        if($json.status -eq 400){
                            Write-Host -ForegroundColor Red "`t$($json.'o:errorDetails'.detail) [$($thisNetSuiteCompany.companyName)][$($thisNetSuiteCompany.id)]"
                            }
                            #}
                        else{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
                        }
                    catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
                    $hubSpotCompaniesToCheck = $hubSpotCompaniesToCheck | ? {$_.id -ne $correspondingHubSpotCompany.id} #Remove this company from $hubSpotCompaniesToCheck as we've already processed it
                    }
                elseif($thisNetSuiteCompany.entityStatus.refName -match "PROSPECT" -or $thisNetSuiteCompany.entityStatus.refName -match "CLIENT"){
                    Write-Host -ForegroundColor Cyan "`tNetSuite Client [$($thisNetSuiteCompany.companyName)][$($thisNetSuiteCompany.id)] is set to [$($thisNetSuiteCompany.entityStatus.refName)] - updating HubSpot object based on NetSuite object"
                    #***Update HubSpot record based on NetSuite data
                    try{
                        $updatedHubSpotCompany = update-hubSpotObjectFromNetSuiteObject -apiKey $apiKey.HubApiKey -objectType companies -netSuiteObject $thisNetSuiteCompany
                        }
                    catch [System.Net.WebException]{
                        $json = ConvertFrom-Json -InputObject $_.ErrorDetails.Message
                        if($json.status -eq 400){
                            Write-Host -ForegroundColor Red "`t$($json.'o:errorDetails'.detail) [$($thisNetSuiteCompany.companyName)][$($thisNetSuiteCompany.id)]"
                            }
                            #}
                        else{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
                        }
                    catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
                    }
                else{
                    Write-Error "NetSuite Company [$($thisNetSuiteCompany.companyName)][$($thisNetSuiteCompany.NetSuiteId)] has conflicting updates with HubSpot, and doesn't seem to be a LEAD, PROSPECT /or/ CLIENT. Looks like someone left a sponge in the patient."
                    continue #Break out of current $thisNetSuiteCompany loop as we've got no idea how to proceed!
                    }
                }
            else{
                Write-Host -ForegroundColor Cyan "`tUpdating HubSpot object [$($correspondingHubSpotCompany.properties.name)][$($correspondingHubSpotCompany.id)] based on NetSuite object [$($thisNetSuiteCompany.companyName)][$($thisNetSuiteCompany.id)]"
                #No conflict - only NetSuite object has been updated, so update HubSpot record based on NetSuite data
                try{
                    $updatedHubSpotCompany = update-hubSpotObjectFromNetSuiteObject -apiKey $apiKey.HubApiKey -objectType companies -netSuiteObject $thisNetSuiteCompany
                    }
                catch{
                    Write-Error $_
                    }
                }
            }
        else{Write-Host -ForegroundColor Cyan "Weird - [$($thisNetSuiteCompany.companyName)][$($thisNetSuiteCompany.id)] doesn't look like it's been updated after all - ignoring it."}
        }
    }

$hubSpotCompaniesToCheck | ? { #Check the remaining HubSpot companies to see whether any need updating/creating
    $thisHubSpotCompany = $_
    if([string]::IsNullOrWhiteSpace($thisHubSpotCompany.NetSuiteId)){ #$netSuite.NetSuiteId -eq $null
        #Check this HUbSpotId isn't already in NetSuite (as we imported a load during the migration)
        $correspondingNetSuiteCompany = get-netSuiteClientsFromNetSuite -query "?q=custentitycustentity_hubspotid IS $($thisHubSpotCompany.id)" -netsuiteParameters $netSuiteParameters
        if($correspondingNetSuiteCompany){
            Write-Host -ForegroundColor Yellow "Unlinked HubSpot company found [$($thisHubSpotCompany.properties.name)][$($thisHubSpotCompany.id)], but corresponding company found in NetSuite [$($correspondingNetSuiteCompany.companyName)][$($correspondingNetSuiteCompany.id)] (probably due to the migration)"
            Write-Host -f DarkYellow "`tupdating NetSuiteId in HubSpot"
            try{$updatedHubSpotCompany = update-hubSpotObject -apiKey $apiKey.HubApiKey -objectType companies -objectId $thisHubSpotCompany.id -fieldHash @{netsuiteid=$correspondingNetSuiteCompany.id}}
            catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
            }
        else{
            #***Create new record in NetSuite
            try{ 
                Write-Host -ForegroundColor Yellow "`tAdding [$($thisHubSpotCompany.properties.name)][$($thisHubSpotCompany.id)] to NetSuite"
                $neNetSuiteClient = add-netSuiteClientToNetSuiteFromHubSpotObject -hubSpotCompanyObject $thisHubSpotCompany -hubSpotApiKey $apiKey.HubApiKey -netsuiteParameters $netSuiteParameters
                }
            catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
            }
        }
    else{#Match to HubSpot on ForeignKey (NetSuiteId)
        Write-Host -ForegroundColor Cyan "Linked HubSpot company found [$($thisHubSpotCompany.properties.name)][$($thisHubSpotCompany.id)] with NetSuiteId [$($thisHubSpotCompany.NetSuiteId)] - getting corresponding Client from cache"
        $correspondingNetSuiteCompany = Compare-Object -ReferenceObject $netSuiteCompaniesToCheck -DifferenceObject $thisHubSpotCompany -Property NetSuiteId -ExcludeDifferent -IncludeEqual -PassThru
        if($correspondingNetSuiteCompany){Write-Host -ForegroundColor DarkCyan "`tCorresponding NetSuite company found [$($correspondingNetSuiteCompany.companyName)][$($correspondingNetSuiteCompany.id)]"}
        else{#Error checking
            Write-Host -ForegroundColor DarkCyan "`tCorresponding NetSuite company not found in cache - checking NetSuite"
            $correspondingNetSuiteCompany = get-netSuiteClientsFromNetSuite -clientId $thisHubSpotCompany.NetSuiteId -netsuiteParameters $netSuiteParameters
            if($correspondingNetSuiteCompany){Write-Host -ForegroundColor DarkCyan "`tCorresponding NetSuite company found [$($correspondingNetSuiteCompany.companyName)][$($correspondingNetSuiteCompany.id)]"}
            else{
                Write-Warning "NetSuite company [$($thisHubSpotCompany.NetSuiteId)] (HubSpot name:[$($thisHubSpotCompany.properties.name)][$($thisHubSpotCompany.HubSpotId)]) could not be retrieved from NetSuite - it may have been deleted from NetSuite?"
                return #Break out of current $thisHubSpotCompany loop if no $correspondingNetSuiteCompany exists
                }
            }
        if([string]::IsNullOrWhiteSpace($thisHubSpotCompany.properties.lastmodifiedinhubspot) -or [Math]::Abs(([datetime]$thisHubSpotCompany.updatedAt - [datetime]$thisHubSpotCompany.properties.lastmodifiedinhubspot).TotalSeconds) -gt 5){ #Has the HubSpot object been updated since the last sync? Specifically: is lastmodifiedinhubspot missing a value (suggesting that it's never been synced) or is the value more than 5 seconds either side of updatedAt (suggesting that it's been edited in HubSpot since the last sync)? The reason we can't compare with -eq here is because lastmodifiedinhubspot and updatedat can never match exactly: whenever we write the current value of updatedat into lastmodifiedinhubspotbysync, it updates the HubSpot record and generates a new value for updatedat (which no longer matches the value we've just writted to lastmodifiedinhubspotbysync). We have to be a little fuzzy and allow the timestamps to be "close enough". We have to ensure that the last time we update the HubSpot record, we get lastmodifiedinhubspot and updatedat within this window.
            Write-Host -ForegroundColor DarkCyan "`t[$($thisHubSpotCompany.properties.name)][$($thisHubSpotCompany.id)] has been updated"
            if($correspondingNetSuiteCompany.entityStatus.refName -match "LEAD"){
                #***Update HubSpot record based on NetSuite data
                Write-Host -ForegroundColor Cyan "`tNetSuite Client [$($correspondingNetSuiteCompany.companyName)][$($correspondingNetSuiteCompany.id)] is set to [$($correspondingNetSuiteCompany.entityStatus.refName)] - updating NetSuite object based on HubSpot object"
                try{$updatedHubSpotCompany = update-netSuiteClientFromHubSpotObject -hubSpotCompanyObject $thisHubSpotCompany -netsuiteParameters $netSuiteParameters -hubSpotApiKey $apiKey.HubApiKey}
                catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
                }
            elseif($correspondingNetSuiteCompany.entityStatus.refName -match "PROSPECT" -or $correspondingNetSuiteCompany.entityStatus.refName -match "CLIENT"){
                #***Update NetSuite record based on HubSpot data
                if([string]::IsNullOrWhiteSpace($correspondingNetSuiteCompany.HubSpotId)){ #There can be edge-cases where the NetSuite Company hasn't been linked to the HubSpot company yet. Link the two first if necessary, then run the update.
                    try{
                        Write-Host -ForegroundColor Cyan "`tNetSuite Client [$($correspondingNetSuiteCompany.companyName)][$($correspondingNetSuiteCompany.id)] hasn't been linked to HubSpot company [$($thisHubSpotCompany.properties.name)][$($thisHubSpotCompany.id)] - updating NetSuite object first"
                        update-netSuiteClientInNetSuite -netSuiteClientId $correspondingNetSuiteCompany.Id -fieldHash @{HubSpotId=$thisHubSpotCompany.id} -netsuiteParameters $netSuiteParameters
                        $correspondingNetSuiteCompany = get-netSuiteClientsFromNetSuite -clientId $correspondingNetSuiteCompany.Id -netsuiteParameters $netSuiteParameters
                        }
                    catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
                    }
                else{
                    Write-Host -ForegroundColor Cyan "`tNetSuite Client [$($correspondingNetSuiteCompany.companyName)][$($correspondingNetSuiteCompany.id)] is set to [$($correspondingNetSuiteCompany.entityStatus.refName)] - we're not allowed to write changes to these, so updating HubSpot object based on NetSuite object to get everything else back in Sync"
                    try{update-hubSpotObjectFromNetSuiteObject -apiKey $apiKey.HubApiKey -objectType companies -netSuiteObject $correspondingNetSuiteCompany}
                    catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
                    }
                }
            else{
                Write-Error "NetSuite Company [$($correspondingNetSuiteCompany.companyName)][$($correspondingNetSuiteCompany.NetSuiteId)] has conflicting updates with HubSpot, and doesn't seem to be a LEAD, PROSPECT /or/ CLIENT. Looks like someone left a sponge in the patient."
                return #Break out of current $thisHubSpotCompany loop as we've got no idea how to proceed!
                }
            #***Update $correspondingNetSuiteCompany.properties.lastmodifiedinnetsuite to $thisHubSpotCompany.lastModifiedDate to exclude it from future syncs (until it is updated again) /*-+This should be part of the Update X record based on Y data functinos/*-+
            }
        }
    }



    #False
        #Check counterpart exists
            #False - Flag as archived/deleted / do nothing
            #True
                #Check $hubSpot.LastNetSuiteModified -le $netSuite.LastModified
                    #True - do nothing
                    #False
                        #Check $netSuite.LastHubSpotModified -eq $hubSpot.LastModified
                            #True - Update $hubSpot with values from $netSuite
                            #False - collision ###Decide which system wins###
    #True
        #Match to HubSpot records
            #NoMatch - Create new HubSpot record
            #Match - Update $netSuote.HubSpotId


#$netSuite.HubSpotId -eq $null
    #False
        #Check counterpart exists
            #False - Flag as archived/deleted / do nothing
            #True
                #Check $hubSpot.LastNetSuiteModified -le $netSuite.LastModified
                    #True - do nothing
                    #False
                        #Check $netSuite.LastHubSpotModified -eq $hubSpot.LastModified
                            #True - Update $hubSpot with values from $netSuite
                            #False - collision ###Decide which system wins###
    #True
        #Match to HubSpot records
            #NoMatch - Create new HubSpot record
            #Match - Update $netSuote.HubSpotId


#endregion

#region Get NetSuite Contacts
$netSuiteContactsToCheck = 
#endregion

#region Process Contacts
$hubSpotCompaniesToCheck | Select-Object | % {
    $thisHubSpotCompany = $_
    if($thisHubSpotCompany.properties.num_associated_contacts -gt 0){
        $theseHubSpotContacts = get-hubSpotContactsFromCompanyId -apiKey $apiKey.HubApiKey -hubspotCompanyId $thisHubSpotCompany.id
        $theseHubSpotContacts | Select-Object | % {
            $thisHubSpotContact = $_
            #Does this Hubspot Contact have a NetSuiteId?
                #No - Does this Contact's HubSpotId appear in NetSuite?
                    #Yes - Cross-reference the Contacts
                    #No  - Does this Contact's e-mail address appear in NetSuite?
                        #Yes - Cross-reference the Contacts
                        #No  - Create a new NetSuite Contact based ont he HubSpot Contact
                #Yes - Has the client been updated more recently in NetSuite?
                    #Yes - Update NetSuite > HubSpot
                    #No  - Update HubSpot > NetSuite

            Write-Host -ForegroundColor Green "Processing HubSpot Contact [$($thisHubSpotContact.properties.firstname)][$($thisHubSpotContact.properties.lastname)][$($thisHubSpotContact.id)]"
            #Does this Hubspot Contact have a NetSuiteId?
            if([string]::IsNullOrWhiteSpace($thisHubSpotContact.properties.netsuiteid)){
                #No - Does this Contact's HubSpotId appear in NetSuite?
                Write-Host -ForegroundColor DarkGreen "`tHubSpot Contact [$($thisHubSpotContact.properties.firstname)][$($thisHubSpotContact.properties.lastname)][$($thisHubSpotContact.id)] has no NetSuiteId - searching for Contact in NetSuite"
                $correspondingNetSuiteContact = get-netSuiteContactFromNetSuite -query "?q=custentitycustentity_hubspotid IS $($thisHubSpotContact.id)" -netsuiteParameters $netSuiteParameters
                if($correspondingNetSuiteContact){
                    #Yes - Cross-reference the Contacts
                    Write-Host -ForegroundColor Green "`tCorresponding NetSuite Contact [$($correspondingNetSuiteContact.entityId)][$($correspondingNetSuiteContact.id)] found by HubSpotId - CROSS-REFERRENCING"
                    try{$updatedHubSpotContact = update-hubSpotObject -apiKey $apiKey.HubApiKey -objectType contacts -objectId $thisHubSpotContact.id -fieldHash @{netsuiteid=$correspondingNetSuiteContact.id}}
                    catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
                    }
                else{
                    #No  - Does this Contact's e-mail address appear in NetSuite?
                    $correspondingNetSuiteContact = get-netSuiteContactFromNetSuite -query "?q=email IS `"$($thisHubSpotContact.properties.email)`"" -netsuiteParameters $netSuiteParameters
                    if($correspondingNetSuiteContact){
                        #Yes - Cross-reference the Contacts
                        Write-Host -ForegroundColor Green "`tCorresponding NetSuite Contact [$($correspondingNetSuiteContact.entityId)][$($correspondingNetSuiteContact.id)] found by email address - CROSS-REFERRENCING"
                        try{$updatedHubSpotContact = update-hubSpotObject -apiKey $apiKey.HubApiKey -objectType contacts -objectId $thisHubSpotContact.id -fieldHash @{netsuiteid=$correspondingNetSuiteContact.id}}
                        catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
                        if([string]::IsNullOrWhiteSpace($correspondingNetSuiteContact.custentitycustentity_hubspotid)){
                            Write-Host -ForegroundColor DarkGreen "`tCorresponding NetSuite Contact [$($correspondingNetSuiteContact.entityId)][$($correspondingNetSuiteContact.id)] was missing custentitycustentity_hubspotid - UPDATING that too"
                            try{$updatedNetSuiteContact = update-netSuiteContactInNetSuite -netSuiteContactId $correspondingNetSuiteContact.id -fieldHash @{custentitycustentity_hubspotid=$thisHubSpotContact.id}}
                            catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
                            }
                        }
                    else{
                        #No  - Create a new NetSuite Contact based on the HubSpot Contact
                        Write-Host -ForegroundColor Green "`tCould not match HubSpot Contact [$($thisHubSpotContact.properties.firstname)][$($thisHubSpotContact.properties.lastname)][$($thisHubSpotContact.id)] to NetSuite - CREATING new contact"
                        if([string]::IsNullOrWhiteSpace($thisHubSpotCompany.properties.netsuiteid)){
                            $filterById = new-hubSpotFilterById -hubSpotId $thisHubSpotCompany.id
                            $thisHubSpotCompany = get-hubSpotObjects -apiKey $apiKey.HubApiKey -objectType companies -filterGroup1 @{filters=@($filterById)} #Refresh the HubSpot Company in case we've added a NetSuiteId above
                            if([string]::IsNullOrWhiteSpace($thisHubSpotCompany.properties.netsuiteid)){
                                Write-Warning "HubSpot Company [$($thisHubSpotCompany.properties.name)][$($thisHubSpotCompany.id)] has no NetSuiteId - cannot create orphaned NetSuite Contacts"
                                return #Exit this foreach-object iteration early
                                }
                            }
                        if([string]::IsNullOrWhiteSpace($thisHubSpotCompany.properties.netsuite_subsidiary)){
                            Write-Warning "HubSpot Company [$($thisHubSpotCompany.properties.name)][$($thisHubSpotCompany.id)] has no netsuite_subsidiary - cannot create NetSuite Contacts without a Subsidiary"
                            return #Exit this foreach-object iteration early
                            }

                        try{
                            $newNetSuiteContact = add-netSuiteContactToNetSuiteFromHubSpotObject -hubSpotContactObject $thisHubSpotContact -hubSpotApiKey $apiKey.HubApiKey -companyNetSuiteId $thisHubSpotCompany.properties.netsuiteid -subsidiary $thisHubSpotCompany.properties.netsuite_subsidiary -netsuiteParameters $netSuiteParameters
                            $newNetSuiteContact = get-netSuiteContactFromNetSuite -query "?q=custentitycustentity_hubspotid IS $($thisHubSpotContact.id)" -netsuiteParameters $netSuiteParameters
                            Write-Host -ForegroundColor DarkGreen "`tNew NetSuite Contact [$($newNetSuiteContact.entityId)][$($newNetSuiteContact.id)] CREATED"
                            }
                        catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
                        
                        }
                    }
                }
            else{
                $correspondingNetSuiteContact = get-netSuiteContactFromNetSuite -contactId $thisHubSpotContact.properties.netsuiteid -netsuiteParameters $netSuiteParameters
                if($correspondingNetSuiteContact){
                    #Yes - Has the client been updated more recently in NetSuite?
                    Write-Host -ForegroundColor DarkGreen "`tCorresponding NetSuite Contact [$($correspondingNetSuiteContact.entityId)][$($correspondingNetSuiteContact.id)] found by NetSuiteId"
                    if((Get-Date $correspondingNetSuiteContact.lastModifiedDate) -gt (Get-Date $thisHubSpotContact.updatedAt)){
                        #Yes - Update NetSuite > HubSpot
                        Write-Host -ForegroundColor Green "`tNetSuite Contact [$($correspondingNetSuiteContact.entityId)][$($correspondingNetSuiteContact.id)] updated more recently than HubSpot Contact [$($thisHubSpotContact.properties.firstname)][$($thisHubSpotContact.properties.lastname)][$($thisHubSpotContact.id)] - UPDATING NetSuite -> HubSpot"
                        if([string]::IsNullOrWhiteSpace($correspondingNetSuiteContact.custentitycustentity_hubspotid)){
                            Write-Warning "NetSuite Contact [$($correspondingNetSuiteContact.entityId)][$($correspondingNetSuiteContact.id)] is missing its custentitycustentity_hubspotid - fixing this first"
                            $updatedNetSuiteContact = update-netSuiteContactInNetSuite -netSuiteContactId $correspondingNetSuiteContact -fieldHash @{custentitycustentity_hubspotid = $thisHubSpotContact.id}
                            $correspondingNetSuiteContact = get-netSuiteContactFromNetSuite -contactId $correspondingNetSuiteContact.id -netsuiteParameters $netSuiteParameters
                            }
                        try{
                            $updatedHubSpotContact = update-hubSpotObjectFromNetSuiteObject -apiKey $apiKey.HubApiKey -objectType contacts -netSuiteObject $correspondingNetSuiteContact
                            Write-Host -ForegroundColor DarkGreen "`tHubSpot Contact [$($thisHubSpotContact.properties.firstname)][$($thisHubSpotContact.properties.lastname)][$($thisHubSpotContact.id)] updated"
                            #Pop this update from $netSuiteContactsToCheck to prevent it the update running again when we 
                            $netSuiteContactsToCheck = $netSuiteContactsToCheck | ? {$_.id -ne $correspondingNetSuiteContact.id}
                            }
                        catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
                        
                        }
                    elseif([Math]::Abs(([datetime]$thisHubSpotContact.updatedAt - [datetime]$thisHubSpotContact.properties.lastmodifiedinhubspot).TotalSeconds) -gt 5){ #Has the HubSpot object been updated since the last sync? Specifically: is lastmodifiedinhubspot missing a value (suggesting that it's never been synced) or is the value more than 5 seconds either side of updatedAt (suggesting that it's been edited in HubSpot since the last sync)? The reason we can't compare with -eq here is because lastmodifiedinhubspot and updatedat can never match exactly: whenever we write the current value of updatedat into lastmodifiedinhubspotbysync, it updates the HubSpot record and generates a new value for updatedat (which no longer matches the value we've just writted to lastmodifiedinhubspotbysync). We have to be a little fuzzy and allow the timestamps to be "close enough". We have to ensure that the last time we update the HubSpot record, we get lastmodifiedinhubspot and updatedat within this window.
                        #No  - Update HubSpot > NetSuite
                        Write-Host -ForegroundColor Green "`tHubSpot Contact [$($thisHubSpotContact.properties.firstname)][$($thisHubSpotContact.properties.lastname)][$($thisHubSpotContact.id)] updated more recently than NetSuite Contact [$($correspondingNetSuiteContact.entityId)][$($correspondingNetSuiteContact.id)] - UPDATING HubSpot -> NetSuite"
                        try{
                            $updatedNetSuiteContact = update-netSuiteContactInNetSuiteFromHubSpotObject -hubSpotContactObject $thisHubSpotContact -hubSpotApiKey $apiKey.HubApiKey -netsuiteParameters $netSuiteParameters
                            }
                        catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
                        }
                    else{
                        #HubSpot Contact has not been updated at all! (Because we've just grabbed all contacts at this Company, we'll probably find quite a few of these)
                        }
                    }
                else{
                    Write-Warning "NetSuite Contact with NetSuiteId [$($thisHubSpotContact.properties.netsuiteid)] is missing from NetSuite (probably deleted) - REMOVING NetSuiteId from HubSpot Contact [$($thisHubSpotContact.properties.firstname)][$($thisHubSpotContact.properties.lastname)][$($thisHubSpotContact.id)]"
                    try{update-hubSpotObject -apiKey $apiKey.HubApiKey -objectType contacts -objectId $thisHubSpotContact.id -fieldHash @{netsuiteid=$null}}
                    catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
                    }
                }
            }
        }
    
    }

#endregion

   
#Match company by FK
    #Match
        #Check counterpart exists
            #False - Flag as archived/deleted
            #True
                #Check $hubSpot.LastNetSuiteModified -le $netSuite.LastModified
                    #True - do nothing
                    #False
                        #Check $netSuite.LastHubSpotModified -eq $hubSpot.LastModified
                            #True - Update $hubSpot with values from $netSuite
                            #False - collision ###Decide which system wins###
                #Check $netSuite.LastHubSpotModified -le $hubSpot.LastModified
                    #True - do nothing
                    #False
                        #Check $netSuite.LastHubSpotModified -eq $hubSpot.LastModified
                            #True - Update $hubSpot with values from $netSuite
                            #False - collision ###Decide which system wins###
    #NoMatch
        #Check $hubSpot.NetSuiteId
            #-eq $null - Create new Netsuite record
            #-ne $null - Update $netSuite.HubSpotId = $hubSpot.id
        #Check $netSuite.hubSpotId
            #-eq $null - Create new HubSpot record
            #-ne $null - $netSuite.HubSpotId = $hubSpot.id













#$netSuite.HubSpotId -eq $null
    #False
        #Check counterpart exists
            #False - Flag as archived/deleted / do nothing
            #True
                #Check $hubSpot.LastNetSuiteModified -le $netSuite.LastModified
                    #True - do nothing
                    #False
                        #Check $netSuite.LastHubSpotModified -eq $hubSpot.LastModified
                            #True - Update $hubSpot with values from $netSuite
                            #False - collision ###Decide which system wins###
    #True
        #Match to HubSpot records
            #NoMatch - Create new HubSpot record
            #Match - Update $netSuote.HubSpotId


   
#Match company by FK
    #Match
        #Check counterpart exists
            #False - Flag as archived/deleted
            #True
                #Check $hubSpot.LastNetSuiteModified -le $netSuite.LastModified
                    #True - do nothing
                    #False
                        #Check $netSuite.LastHubSpotModified -eq $hubSpot.LastModified
                            #True - Update $hubSpot with values from $netSuite
                            #False - collision ###Decide which system wins###
                #Check $netSuite.LastHubSpotModified -le $hubSpot.LastModified
                    #True - do nothing
                    #False
                        #Check $netSuite.LastHubSpotModified -eq $hubSpot.LastModified
                            #True - Update $hubSpot with values from $netSuite
                            #False - collision ###Decide which system wins###
    #NoMatch
        #Check $hubSpot.NetSuiteId
            #-eq $null - Create new Netsuite record
            #-ne $null - Update $netSuite.HubSpotId = $hubSpot.id
        #Check $netSuite.hubSpotId
            #-eq $null - Create new HubSpot record
            #-ne $null - $netSuite.HubSpotId = $hubSpot.id