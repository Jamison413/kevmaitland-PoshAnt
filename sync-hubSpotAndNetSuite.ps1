[cmdletbinding()]
param(
    [Parameter(Mandatory = $false, Position = 0)]
        [string]$deltaSync = $true #Specifies whether we are doing a full or incremental sync.
    )

if($PSCommandPath){
    $InformationPreference = 2
    $VerbosePreference = 0
    $logFileLocation = "C:\ScriptLogs\"
    if($deltaSync -eq $true){$suffix = "_deltaSync"}
    else{$suffix = "_fullSync"}
    #$suffix = "_fullSync"
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))$suffix`_Transcript_$(Get-Date -Format "yyyy-MM-dd").log"
    Start-Transcript $transcriptLogName -Append
    }
function test-hubSpotTimeStampIsCloseEnough(){
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
            [datetime]$updatedAt
        ,[Parameter(Mandatory = $true, Position = 1)]
            [AllowNull()]
            [nullable[datetime]]$lastmodifiedinhubspot
        )

    if($lastmodifiedinhubspot -eq $null){return $false}

    $allowedDiscrepencyInSeconds = 5

    if([Math]::Abs(([datetime]$updatedAt - [datetime]$lastmodifiedinhubspot).TotalSeconds) -lt $allowedDiscrepencyInSeconds){$true}
    else{$false}

    }

#region Get HubSpot Companies
$apiKey = get-hubSpotApiKey
$filterCompanyFlaggedForSync = [ordered]@{
    propertyName="netsuite_sync_company_"
    operator="HAS_PROPERTY"
    }
$filterCompanyNotBroken = [ordered]@{
    propertyName="netsuite_company_has_been_merged_or_deleted"
    operator="NOT_HAS_PROPERTY"
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
$hubspotCompanyMaxlastmodifiedinhubspotbysync = get-hubSpotObjects -apiKey $apiKey.HubApiKey -objectType companies -filterGroup1 @{filters=@($filterCompanyFlaggedForSync,$filterCompanyNotBroken)} -sortPropertyNameAndDirection $hubSortMaxlastmodifiedinhubspotbysync -pageSize 1 -firstPageOnly

$filterCompanyUpdatedSinceLastSync = [ordered]@{
    propertyName="hs_lastmodifieddate"
    operator="GT"
    #value = [Math]::Floor([decimal](Get-Date(Get-Date "2000-10-20T08:34:48.887Z").ToUniversalTime()-uformat "%s"))*1000 #Convert to UNIX Epoch time and add Milliseconds
    value = [Math]::Floor([decimal](Get-Date(Get-Date $hubspotCompanyMaxlastmodifiedinhubspotbysync.properties.lastmodifiedinhubspot).ToUniversalTime()-uformat "%s"))*1000 #Convert to UNIX Epoch time and add Milliseconds
    }
if($deltaSync -eq $false){
    [array]$hubSpotCompaniesToCheck = get-hubspotObjects -apiKey $apiKey.HubApiKey -objectType companies -filterGroup1 @{filters=@($filterCompanyFlaggedForSync,$filterCompanyNotBroken,$filterExcludeCompaniesCalledAnthesis)} -pageSize 100 #-firstPageOnly 
    if($hubSpotCompaniesToCheck.Count -gt 0){export-encryptedCache -arrayOfObjects $hubSpotCompaniesToCheck -fileName hubCompanies.csv }
    }
else{
    [array]$hubSpotCompaniesToCheck = get-hubspotObjects -apiKey $apiKey.HubApiKey -objectType companies -filterGroup1 @{filters=@($filterCompanyFlaggedForSync,$filterCompanyUpdatedSinceLastSync,$filterExcludeCompaniesCalledAnthesis)} -pageSize 100 #-firstPageOnly 
    $hubSpotCompaniesToCheck = $hubSpotCompaniesToCheck | ? {$_.properties.netsuite_company_has_been_merged_or_deleted -ne $true} #We can only provide 3 Filters in the group above, so we have to do any additional filtering Client-Side
    }
$hubSpotCompaniesToCheck | Select-Object | % { #Prep the objects for compare-object later
    Add-Member -InputObject $_ -MemberType NoteProperty -Name HubSpotId -Value $_.Id
    Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteId -Value $_.properties.netsuiteid
    }
$hubSpotCompaniesToCheckForContacts = $hubSpotCompaniesToCheck
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
if($deltaSync -eq $true){$netQuery += " AND lastModifiedDate ON_OR_AFTER `"$($(Get-Date $hubspotCompanyMaxLastModifiedInNetSuite.properties.lastmodifiedinnetsuite -Format g))`""} #Excludes any Companies that haven't been updated since X

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
    $correspondingHubSpotCompany = $null
    #***Create new record in HubSpot
    #First, try to find a HubSpot Company with the corresponding NetSuiteID
    $filterCompanyNetSuiteIdEq = [ordered]@{
        propertyName="netsuiteid"
        operator="EQ"
        value=$thisNetSuiteCompany.NetSuiteId
        }
    $hubSortOldestCreatedInHubSpot = [ordered]@{
        propertyName = "createdate"
        direction = "ASCENDING"
        }
    $oldestHubspotCompanyByNetSuiteId = get-hubSpotObjects -apiKey $apiKey.HubApiKey -objectType companies -filterGroup1 @{filters=@($filterCompanyNetSuiteIdEq)} -sortPropertyNameAndDirection $hubSortOldestCreatedInHubSpot -pageSize 1 -firstPageOnly
    if(![string]::IsNullOrWhiteSpace($oldestHubspotCompanyByNetSuiteId)){
        Write-Host -ForegroundColor Cyan "Corresponding HubSpot Company [$($oldestHubspotCompanyByNetSuiteId.properties.name)][$($oldestHubspotCompanyByNetSuiteId.id)] retrieved from HubSpot based on NetSuiteId [$($thisNetSuiteCompany.NetSuiteId)]"
        $correspondingHubSpotCompany = $oldestHubspotCompanyByNetSuiteId
        }
    if([string]::IsNullOrWhiteSpace($oldestHubspotCompanyByNetSuiteId) -and ![string]::IsNullOrWhiteSpace($thisNetSuiteCompany.HubSpotId)){
    #If we can't find the corresponding HubSpot company by NetSuiteId, try by the HubSpotId on the NetSuite object (if there is one)
        Write-Host -ForegroundColor Cyan "Linked NetSuite company found [$($thisNetSuiteCompany.companyName)][$($thisNetSuiteCompany.id)] - retrieving HubSpot Company with Id [$($thisNetSuiteCompany.HubSpotId)]"
        $hubspotCompanyByHubSpotId = Compare-Object -ReferenceObject $hubSpotCompaniesToCheck -DifferenceObject $thisNetSuiteCompany -Property HubSpotId -ExcludeDifferent -IncludeEqual -PassThru
        if($hubspotCompanyByHubSpotId){
            Write-Host -ForegroundColor DarkCyan "`tCorresponding HubSpot Company [$($hubspotCompanyByHubSpotId.properties.name)][$($hubspotCompanyByHubSpotId.id)] retrieved from Companies flagged to Sync"
            $correspondingHubSpotCompany = $hubspotCompanyByHubSpotId
            }
        else{ #Error checking
            Write-Host -ForegroundColor DarkCyan "`tHubSpot Company not retrieved from Companies flagged to Sync - trying HubSpot"
            $hubspotFilterThisId = new-hubSpotFilterById -hubSpotId $thisNetSuiteCompany.HubSpotId
            $hubspotCompanyByHubSpotId = get-hubSpotObjects -apiKey $apiKey.HubApiKey -objectType companies -filterGroup1 @{filters=@($hubspotFilterThisId)}
            if($hubspotCompanyByHubSpotId){
                Write-Host -ForegroundColor DarkCyan "`tCorresponding HubSpot Company [$($hubspotCompanyByHubSpotId.properties.name)][$($hubspotCompanyByHubSpotId.id)] retrieved from HubSpot"
                $correspondingHubSpotCompany = $hubspotCompanyByHubSpotId
                }
            elseif(!$hubspotCompanyByHubSpotId){
                Write-Warning "HubSpot company [$($thisNetSuiteCompany.HubSpotId)] (NetSuite name:[$($thisNetSuiteCompany.companyName)][$($thisNetSuiteCompany.NetSuiteId)]) could not be retrieved from HubSpot - it may have been deleted from HubSpot?"
                return #Break out of current $thisNetSuiteCompany loop if no $hubspotCompanyByHubSpotId exists
                }
            elseif($hubspotCompanyByHubSpotId.properties.netsuite_sync_company_ -eq $false){
                Write-Warning "HubSpot company [$($hubspotCompanyByHubSpotId.properties.name)][$($thisNetSuiteCompany.HubSpotId)] (NetSuite name:[$($thisNetSuiteCompany.companyName)][$($thisNetSuiteCompany.NetSuiteId)]) is not flagged with {netsuite_sync_company_} - not syncing chamges back to HubSpot"
                return #Break out of current $thisNetSuiteCompany loop if $hubspotCompanyByHubSpotId is not flagged netsuite_sync_company_
                }
            }
        }
    if([string]::IsNullOrWhiteSpace($oldestHubspotCompanyByNetSuiteId) -and [string]::IsNullOrWhiteSpace($hubspotCompanyByHubSpotId)){
    #If we can't find a corresponding HubSpot company by either Id, try by e-mail address
        if([string]::IsNullOrWhiteSpace($thisNetSuiteCompany.email) -or $thisNetSuiteCompany.email -match "@anthesisgroup.com" -or $thisNetSuiteCompany.email -match "@lavola.com"){
            Write-Host -ForegroundColor Yellow "Unlinked NetSuite company found [$($thisNetSuiteCompany.companyName)][$($thisNetSuiteCompany.id)], but no generic e-mail address set. No meaningful way to identify this company in HubSpot."
            $hubspotCompanyByEmail = $null
            }
        else{
            Write-Host -ForegroundColor Yellow "Unlinked NetSuite company found [$($thisNetSuiteCompany.companyName)][$($thisNetSuiteCompany.id)] - looking for match in HubSpot"
            $hubspotCompanyByEmail = $hubSpotCompaniesToCheck | ? {$_.properties.generic_email_address__c -eq $thisNetSuiteCompany.email}
            if($hubspotCompanyByEmail){
                Write-Host -ForegroundColor DarkYellow "`tGeneric e-mail address [$($thisNetSuiteCompany.email)] found in HubSpot records marked for Sync [$($hubspotCompanyByEmail.properties.name)][$($hubspotCompanyByEmail.id)]"
                $correspondingHubSpotCompany = $hubspotCompanyByEmail
                }
            else{
                Write-Host -ForegroundColor DarkYellow "`tGeneric e-mail address [$($thisNetSuiteCompany.email)] not found in HubSpot records marked for Sync - searching all HubSpot records"
                $filterCompanyGenericEmailEq = [ordered]@{
                    propertyName="generic_email_address__c"
                    operator="EQ"
                    value=$thisNetSuiteCompany.email
                    }
                $hubspotCompanyByEmail = get-hubSpotObjects -apiKey $apiKey.HubApiKey -objectType companies -filterGroup1 @{filters=@($filterCompanyGenericEmailEq)}

                #Try to match to any contact e-mail address and deduce Company (and link to that)
                if($hubspotCompanyByEmail){
                    Write-Host -ForegroundColor DarkYellow "`tGeneric e-mail address [$($thisNetSuiteCompany.email)] found in HubSpot [$($hubspotCompanyByEmail.properties.name)][$($hubspotCompanyByEmail.id)]"
                    $correspondingHubSpotCompany = $hubspotCompanyByEmail
                    }
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
                        $hubspotCompanyByEmail = get-hubSpotObjects -apiKey $apiKey.HubApiKey -objectType companies -filterGroup1 @{filters=@($filterCompanyIdEq)}
                        if($hubspotCompanyByEmail){
                            Write-Host -ForegroundColor DarkYellow "`tContact [$($matchedHubspotContact.properties.firstname)][$($matchedHubspotContact.properties.lastname)][$($matchedHubspotContact.id)] e-mail address [$($thisNetSuiteCompany.email)] found in HubSpot [$($hubspotCompanyByEmail.properties.name)][$($hubspotCompanyByEmail.id)]"
                            $correspondingHubSpotCompany = $hubspotCompanyByEmail
                            }
                        }
                    }
                }
            }
        if($hubspotCompanyByEmail){#Update and cross-reference both companies (if we've matched them)
            Write-Host -ForegroundColor Yellow "`tCross-referencing NetSuite Company [$($thisNetSuiteCompany.companyName)][$($thisNetSuiteCompany.id)] with HubSpot Company [$($hubspotCompanyByEmail.properties.name)][$($hubspotCompanyByEmail.id)]"
            try{$updatedHubSpotCompany = update-hubSpotObject -apiKey $apiKey.HubApiKey -objectType companies -objectId $hubspotCompanyByEmail.id -fieldHash @{netsuiteid=$thisNetSuiteCompany.id; lastmodifiedinnetsuite=$(get-dateInIsoFormat -dateTime $(Get-Date $thisNetSuiteCompany.lastModifiedDate).AddYears(-100) -precision Seconds);lastmodifiedinhubspot=$(get-dateInIsoFormat -dateTime $(Get-Date) -precision Ticks)}}
            catch{Write-Error $_}
            try{$updatedNetSuiteCompany = update-netSuiteClientInNetSuite -netSuiteClientId $thisNetSuiteCompany.id -fieldHash @{custentitycustentity_hubspotid = $hubspotCompanyByEmail.id} -netsuiteParameters $netSuiteParameters}
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

    if($correspondingHubSpotCompany){
    #If we've found a corresponding HubSpot company using any of the methods above, see if it needs updating
        if($thisNetSuiteCompany.lastModifiedDate -ne $correspondingHubSpotCompany.properties.lastmodifiedinnetsuite){ #Has the NetSuite object been updated since the last sync?
            Write-Host -ForegroundColor DarkCyan "`tNetSuite company has been modified."
            if($(test-hubSpotTimeStampIsCloseEnough -updatedAt $correspondingHubSpotCompany.updatedAt -lastmodifiedinhubspot $correspondingHubSpotCompany.properties.lastmodifiedinhubspot) -eq $false){ #Has the HubSpot object been updated since the last sync?
                Write-Host -ForegroundColor DarkCyan "`tHubSpot company has also been modified."
                #Both objects modified - in the event of a conflict, Leads are overwritten by HubSpot and Prospects/Clients are overwritten by NetSuite
                if($thisNetSuiteCompany.entityStatus.refName -match "LEAD"){
                    #***Update NetSuite record based on HubSpot data
                    Write-Host -ForegroundColor Cyan "`t`tNetSuite Client [$($thisNetSuiteCompany.companyName)][$($thisNetSuiteCompany.id)] is set to [$($thisNetSuiteCompany.entityStatus.refName)] - updating NetSuite object based on HubSpot object"
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
                    Write-Host -ForegroundColor Cyan "`t`tNetSuite Client [$($thisNetSuiteCompany.companyName)][$($thisNetSuiteCompany.id)] is set to [$($thisNetSuiteCompany.entityStatus.refName)] - updating HubSpot object based on NetSuite object"
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

$hubSpotCompaniesToCheck | % { #Check the remaining HubSpot companies to see whether any need updating/creating
    $thisHubSpotCompany = $_
    if([string]::IsNullOrWhiteSpace($thisHubSpotCompany.NetSuiteId)){ #$netSuite.NetSuiteId -eq $null
        #Check this HUbSpotId isn't already in NetSuite (as we imported a load during the migration)
        $correspondingNetSuiteCompany = get-netSuiteClientsFromNetSuite -query "?q=custentitycustentity_hubspotid IS $($thisHubSpotCompany.id)" -netsuiteParameters $netSuiteParameters #Match by immutable Id first
        if( [string]::IsNullOrEmpty($correspondingNetSuiteCompany.id)){$correspondingNetSuiteCompany = get-netSuiteClientsFromNetSuite -query "?q=companyName IS `"$($thisHubSpotCompany.properties.name)`"" -netsuiteParameters $netSuiteParameters} #Try again by companyName as this would collide when we try to create a new Client anyway
        if(![string]::IsNullOrEmpty($correspondingNetSuiteCompany.id)){
            Write-Host -ForegroundColor Yellow "Unlinked HubSpot company found [$($thisHubSpotCompany.properties.name)][$($thisHubSpotCompany.id)], but corresponding company found in NetSuite [$($correspondingNetSuiteCompany.companyName)][$($correspondingNetSuiteCompany.id)] (probably due to the migration)"
            Write-Host -f DarkYellow "`tupdating NetSuiteId in HubSpot"
            try{$updatedHubSpotCompany = update-hubSpotObject -apiKey $apiKey.HubApiKey -objectType companies -objectId $thisHubSpotCompany.id -fieldHash @{netsuiteid=$correspondingNetSuiteCompany.id}}
            catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
            }
        else{
            #***Create new record in NetSuite
            try{ 
                Write-Host -ForegroundColor Yellow "[$($thisHubSpotCompany.properties.name)][$($thisHubSpotCompany.id)] not found - Adding to NetSuite"
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
            try{
                $correspondingNetSuiteCompany = get-netSuiteClientsFromNetSuite -clientId $thisHubSpotCompany.NetSuiteId -netsuiteParameters $netSuiteParameters -ErrorAction Stop
                if($correspondingNetSuiteCompany){Write-Host -ForegroundColor DarkCyan "`tCorresponding NetSuite company found [$($correspondingNetSuiteCompany.companyName)][$($correspondingNetSuiteCompany.id)]"}
                }
            catch{
                if($_.Exception -match "404" -or $_.InnerException -match "404"){
                    Write-Warning "NetSuite company [$($thisHubSpotCompany.NetSuiteId)] (HubSpot name:[$($thisHubSpotCompany.properties.name)][$($thisHubSpotCompany.HubSpotId)]) could not be retrieved from NetSuite - it may have been deleted from NetSuite?"
                    #Write-Host -ForegroundColor DarkCyan "`tRemoving invalid NetSuiteId from HubSpot Company [$($thisHubSpotCompany.properties.name)][$($thisHubSpotCompany.HubSpotId)]"
                    #$updatedHubSpotCompany = update-hubSpotObject -apiKey $apiKey.HubApiKey -objectType companies -objectId $thisHubSpotCompany.id -fieldHash @{netsuiteid=""} #HubSpot won't let us $null this
                    #Nah - this'll just recreate the merged Company in NetSuite and piss everyone off.
                    Write-Host -ForegroundColor DarkCyan "`tFlagging HubSpot Company [$($thisHubSpotCompany.properties.name)][$($thisHubSpotCompany.HubSpotId)] as having an invalid NetSuiteId"
                    $updatedHubSpotCompany = update-hubSpotObject -apiKey $apiKey.HubApiKey -objectType companies -objectId $thisHubSpotCompany.id -fieldHash @{netsuite_company_has_been_merged_or_deleted=$true} 
                    $hubSpotCompaniesToCheck = $hubSpotCompaniesToCheck |? {$hubSpotCompaniesToCheck.id -notcontains $_.id} #Remove this duffer from the array
                    }
                else{Write-Host -f Red $(get-errorSummary $_)}
                return #Break out of current $thisHubSpotCompany loop if no $correspondingNetSuiteCompany exists
                }
            
            }
        if([string]::IsNullOrWhiteSpace($thisHubSpotCompany.properties.lastmodifiedinhubspot) -or $(test-hubSpotTimeStampIsCloseEnough -updatedAt $thisHubSpotCompany.updatedAt -lastmodifiedinhubspot $thisHubSpotCompany.properties.lastmodifiedinhubspot) -eq $false){ #Has the HubSpot object been updated since the last sync? Specifically: is lastmodifiedinhubspot missing a value (suggesting that it's never been synced) or is the value more than 5 seconds either side of updatedAt (suggesting that it's been edited in HubSpot since the last sync)? The reason we can't compare with -eq here is because lastmodifiedinhubspot and updatedat can never match exactly: whenever we write the current value of updatedat into lastmodifiedinhubspotbysync, it updates the HubSpot record and generates a new value for updatedat (which no longer matches the value we've just writted to lastmodifiedinhubspotbysync). We have to be a little fuzzy and allow the timestamps to be "close enough". We have to ensure that the last time we update the HubSpot record, we get lastmodifiedinhubspot and updatedat within this window.
            Write-Host -ForegroundColor DarkCyan "`t[$($thisHubSpotCompany.properties.name)][$($thisHubSpotCompany.id)] has been updated"
            if($correspondingNetSuiteCompany.entityStatus.refName -match "LEAD"){
                #***Update HubSpot record based on NetSuite data
                Write-Host -ForegroundColor Cyan "`t`tNetSuite Client [$($correspondingNetSuiteCompany.companyName)][$($correspondingNetSuiteCompany.id)] is set to [$($correspondingNetSuiteCompany.entityStatus.refName)] - updating NetSuite object based on HubSpot object"
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
                    try{$updatedHubSpotCompany = update-hubSpotObjectFromNetSuiteObject -apiKey $apiKey.HubApiKey -objectType companies -netSuiteObject $correspondingNetSuiteCompany}
                    catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
                    }
                }
            else{
                Write-Error "NetSuite Company [$($correspondingNetSuiteCompany.companyName)][$($correspondingNetSuiteCompany.NetSuiteId)] has conflicting updates with HubSpot, and doesn't seem to be a LEAD, PROSPECT /or/ CLIENT. Looks like someone left a sponge in the patient."
                return #Break out of current $thisHubSpotCompany loop as we've got no idea how to proceed!
                }
            #***Update $correspondingNetSuiteCompany.properties.lastmodifiedinnetsuite to $thisHubSpotCompany.lastModifiedDate to exclude it from future syncs (until it is updated again) /*-+This should be part of the Update X record based on Y data functinos/*-+
            }
        else{Write-Host -f DarkCyan "`t`t[$($thisHubSpotCompany.properties.name)][$($thisHubSpotCompany.id)] is unchanged and did not require updating"}
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


#region Process Contacts
#region Get NetSuite Contacts
$netContactQuery = "?q=email EMPTY_NOT"
if($deltaSync -eq $true){
    $dummyFilter = [ordered]@{
        propertyName="id"
        operator="HAS_PROPERTY"
        }
    $hubspotContactMaxLastModifiedInNetSuite = get-hubSpotObjects -apiKey $apiKey.HubApiKey -objectType contacts -filterGroup1 @{filters=@($dummyFilter)} -sortPropertyNameAndDirection $hubSortLastModifiedInNetSuite -pageSize 1 -firstPageOnly #Sorting only works alongside a Filter :/
    $netContactQuery +=  " AND lastModifiedDate ON_OR_AFTER `"$($(Get-Date $hubspotContactMaxLastModifiedInNetSuite.properties.lastmodifiedinnetsuite -Format g))`"" #Excludes any Contacts that haven't been updated since X
    }
$netSuiteContactsToCheck = get-netSuiteContactFromNetSuite -netsuiteParameters $netSuiteParameters -query $netContactQuery 
if($deltaSync -eq $false -and $netSuiteContactsToCheck.Count -gt 0){export-encryptedCache -arrayOfObjects $netSuiteContactsToCheck -fileName netContacts.csv}
#endregion

#region Process HubSpot Contacts
@($hubSpotCompaniesToCheckForContacts | ? {$_.properties.num_associated_contacts -gt 0} | Select-Object) | % {
    $thisHubSpotCompany = $_
    if([string]::IsNullOrWhiteSpace($hubspotContactMaxLastModifiedInNetSuite.properties.lastmodifiedinhubspot)){
        $theseHubSpotContacts = get-hubSpotContactsFromCompanyId -apiKey $apiKey.HubApiKey -hubspotCompanyId $thisHubSpotCompany.id -includeContactsWithNoTimeStamp #-Verbose
        } 
    else{
        $theseHubSpotContacts = get-hubSpotContactsFromCompanyId -apiKey $apiKey.HubApiKey -hubspotCompanyId $thisHubSpotCompany.id -updatedAfter $hubspotContactMaxLastModifiedInNetSuite.properties.lastmodifiedinhubspot -includeContactsWithNoTimeStamp #-Verbose
        } 
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
        $thisHubSpotContactsEvents = get-hubSpotEvents -apiKey $apiKey.HubApiKey -hubspotContactId $thisHubSpotContact.id
        $thisHubSpotContact = add-hubSpoteventDataToContact -contactObject $thisHubSpotContact -eventsArray $thisHubSpotContactsEvents
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
                            try{$updatedNetSuiteContact = update-netSuiteContactInNetSuite -netSuiteContactId $correspondingNetSuiteContact.id -fieldHash @{custentitycustentity_hubspotid=$thisHubSpotContact.id} -netsuiteParameters $netSuiteParameters} 
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
                            $newNetSuiteContact = add-netSuiteContactToNetSuiteFromHubSpotObject -hubSpotContactObject $thisHubSpotContact -hubSpotApiKey $apiKey.HubApiKey -companyNetSuiteId $thisHubSpotCompany.properties.netsuiteid -subsidiary $thisHubSpotCompany.properties.netsuite_subsidiary -netsuiteParameters $netSuiteParameters #-Verbose
                            Write-Host -ForegroundColor DarkGreen "`tNew NetSuite Contact [$($newNetSuiteContact.entityId)][$($newNetSuiteContact.id)][$($newNetSuiteContact.company.refName)][$($newNetSuiteContact.company.id)] CREATED"
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
                        $netSuiteContactsToCheck = $netSuiteContactsToCheck | ? {$_.id -ne $correspondingNetSuiteContact.id} #Pop this update from $netSuiteContactsToCheck to prevent it being updated again when we process $netSuiteContactsToCheck in a moment
                        }
                    catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
                        
                    }
                elseif($(test-hubSpotTimeStampIsCloseEnough -updatedAt $thisHubSpotContact.updatedAt -lastmodifiedinhubspot $thisHubSpotContact.properties.lastmodifiedinhubspot) -eq $false){ #Has the HubSpot object been updated since the last sync? Specifically: is lastmodifiedinhubspot missing a value (suggesting that it's never been synced) or is the value more than 5 seconds either side of updatedAt (suggesting that it's been edited in HubSpot since the last sync)? The reason we can't compare with -eq here is because lastmodifiedinhubspot and updatedat can never match exactly: whenever we write the current value of updatedat into lastmodifiedinhubspotbysync, it updates the HubSpot record and generates a new value for updatedat (which no longer matches the value we've just writted to lastmodifiedinhubspotbysync). We have to be a little fuzzy and allow the timestamps to be "close enough". We have to ensure that the last time we update the HubSpot record, we get lastmodifiedinhubspot and updatedat within this window.
                    #No  - Update HubSpot > NetSuite
                    Write-Host -ForegroundColor Green "`tHubSpot Contact [$($thisHubSpotContact.properties.firstname)][$($thisHubSpotContact.properties.lastname)][$($thisHubSpotContact.id)] updated more recently than NetSuite Contact [$($correspondingNetSuiteContact.entityId)][$($correspondingNetSuiteContact.id)] - UPDATING HubSpot -> NetSuite"
                    try{
                        $updatedNetSuiteContact = update-netSuiteContactInNetSuiteFromHubSpotObject -hubSpotContactObject $thisHubSpotContact -hubSpotApiKey $apiKey.HubApiKey -netsuiteParameters $netSuiteParameters
                        $netSuiteContactsToCheck = $netSuiteContactsToCheck | ? {$_.id -ne $updatedNetSuiteContact.id} #Pop this update from $netSuiteContactsToCheck to prevent it being updated again when we process $netSuiteContactsToCheck in a moment
                        }
                    catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
                    }
                else{
                    #HubSpot Contact has not been updated at all! (now we filter Contacts based on lastmodified
                    Write-Host -ForegroundColor DarkGreen "`tHubSpot Contact [$($thisHubSpotContact.properties.firstname)][$($thisHubSpotContact.properties.lastname)][$($thisHubSpotContact.id)] did not require updating!"
                    }
                }
            else{
                Write-Warning "NetSuite Contact with NetSuiteId [$($thisHubSpotContact.properties.netsuiteid)] is missing from NetSuite (probably deleted) - REMOVING NetSuiteId from HubSpot Contact [$($thisHubSpotContact.properties.firstname)][$($thisHubSpotContact.properties.lastname)][$($thisHubSpotContact.id)]"
                try{update-hubSpotObject -apiKey $apiKey.HubApiKey -objectType contacts -objectId $thisHubSpotContact.id -fieldHash @{netsuiteid=""}} #HUbSpot won't let us $null this
                catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
                }
            }
        }
    
    }

#endregion

#region Process remaining NetSuite Contacts
@($netSuiteContactsToCheck | Select-Object) | % {
    #Does this NetSuite Contact have a HubSpotId?
        #No - Does this Contact's NetSuiteId appear in HubSpot?
            #Yes - Cross-reference the Contacts
            #No  - Does this Contact's e-mail address appear in HubSpot?
                #Yes - Cross-reference the Contacts
                #No  - Create a new HubSpot Contact based on the NetSuite Contact
        #Yes - Has the client been updated more recently in HubSpot?
            #Yes - Update HubSpot > NetSuite
            #No  - Update NetSuite > HubSpot
    $thisNetSuiteContact = $_
    #Does this NetSuite Contact have a HubSpotId?
    Write-Host -ForegroundColor Green "Processing NetSuite Contact [$($thisNetSuiteContact.email)][$($thisNetSuiteContact.id)][$($thisNetSuiteContact.company.refName)][$($thisNetSuiteContact.company.id)]"
    if([string]::IsNullOrEmpty($thisNetSuiteContact.custentitycustentity_hubspotid)){
        #No - Does this Contact's NetSuiteId appear in HubSpot?
        $hubContactIdFilter = [ordered]@{
            propertyName="netsuiteid"
            operator="EQ"
            value=$thisNetSuiteContact.id
            }
        $correspondingHubSpotContact = get-hubSpotObjects -apiKey $apiKey.HubApiKey -objectType contacts -filterGroup1 @{filters=@($hubContactIdFilter)}
        if(![string]::IsNullOrEmpty($correspondingHubSpotContact.id)){
            #Yes - Cross-reference the Contacts
            Write-Host -f DarkGreen "`tHubSpot Contact [$($correspondingHubSpotContact.properties.email)][$($correspondingHubSpotContact.id)] matched by Id"
            try{
                $updatedNetSuiteContact = update-netSuiteContactInNetSuite -netSuiteContactId $thisNetSuiteContact.id -fieldHash @{custentitycustentity_hubspotid=$correspondingHubSpotContact.id} -netsuiteParameters $netSuiteParameters -ErrorAction Stop
                Write-Host -f Green "`tHubSpot Contact [$($correspondingHubSpotContact.properties.email)][$($correspondingHubSpotContact.id)] and NetSuite Contact [$($thisNetSuiteContact.email)][$($thisNetSuiteContact.id)][$($thisNetSuiteContact.company.refName)][$($thisNetSuiteContact.company.id)] CROSS-REFERENCED by HubSpotId"
                }
            catch{Write-Host -f Red $(get-errorSummary -errorToSummarise $_)}
            }
            #No  - Does this Contact's e-mail address appear in HubSpot?
        if( [string]::IsNullOrEmpty($correspondingHubSpotContact.id)){ 
            $hubContactEmailFilter = [ordered]@{
                propertyName="email"
                operator="EQ"
                value=$thisNetSuiteContact.email
                }
            $correspondingHubSpotContact = get-hubSpotObjects -apiKey $apiKey.HubApiKey -objectType contacts -filterGroup1 @{filters=@($hubContactEmailFilter)}
                #Yes - Cross-reference the Contacts
            if(![string]::IsNullOrEmpty($correspondingHubSpotContact.id)){
                Write-Host -f DarkGreen "`tHubSpot Contact [$($correspondingHubSpotContact.properties.email)][$($correspondingHubSpotContact.id)] matched by Email"
                try{
                    $updatedNetSuiteContact = update-netSuiteContactInNetSuite -netSuiteContactId $thisNetSuiteContact.id -fieldHash @{custentitycustentity_hubspotid=$correspondingHubSpotContact.id} -netsuiteParameters $netSuiteParameters -ErrorAction Stop 
                    Write-Host -f Green "`tHubSpot Contact [$($correspondingHubSpotContact.properties.email)][$($correspondingHubSpotContact.id)] and NetSuite Contact [$($thisNetSuiteContact.email)][$($thisNetSuiteContact.id)][$($thisNetSuiteContact.company.refName)][$($thisNetSuiteContact.company.id)] CROSS-REFERENCED by email"
                    }
                catch{Write-Host -f Red $(get-errorSummary -errorToSummarise $_)}
                }
                #No  - Create a new HubSpot Contact based on the NetSuite Contact
            else{
                Write-Host -ForegroundColor Green "`tCould not match NetSuite Contact [$($thisNetSuiteContact.email)][$($thisNetSuiteContact.id)][$($thisNetSuiteContact.company.refName)][$($thisNetSuiteContact.company.id)] to any HubSpot Contact: CREATING new HubSpot Contact"
                $parentHubSpotCompanyId = $($netSuiteCompaniesToCheck | ? {$_.id -eq $thisNetSuiteContact.company.id}).HubSpotId #Try to save a query to NetSuite by checking the cache. This could be improved with a compare-object
                if([string]::IsNullOrEmpty($parentHubSpotCompanyId)){
                    try{
                        $newHubSpotContact = new-hubspotContactFromNetsuiteContact -apiKey $apiKey.HubApiKey -netSuiteContact $thisNetSuiteContact -netSuiteParams $netSuiteParameters -ErrorAction Stop
                        Write-Host -f Green "`t`tNew HubSpot Contact [$($newHubSpotContact.properties.email)][$($newHubSpotContact.id)] created!"
                        $null = update-netSuiteContactInNetSuite -netSuiteContactId $thisNetSuiteContact.id -fieldHash @{custentitycustentity_hubspotid = $newHubSpotContact.id; custentity_marketing_originalsourcesyste = "NetSuite"} -netsuiteParameters $netSuiteParameters
                        $thisNetSuiteContact = get-netSuiteContactFromNetSuite -contactId $thisNetSuiteContact.id -netsuiteParameters $netSuiteParameters
                        Write-Host -f DarkGreen "`t`tNetSuite Contact [$($thisNetSuiteContact.email)][$($thisNetSuiteContact.id)][$($thisNetSuiteContact.company.refName)][$($thisNetSuiteContact.company.id)] updated with new HobSpotId [$($newHubSpotContact.id)]"
                        }
                    catch{Write-Host -f Red $(get-errorSummary -errorToSummarise $_)}
                    
                    }
                else{
                    try{
                        $newHubSpotContact = new-hubspotContactFromNetsuiteContact -apiKey $apiKey.HubApiKey -netSuiteContact $thisNetSuiteContact -hubSpotCompanyId $parentHubSpotCompanyId -ErrorAction Stop
                        Write-Host -f Green "`t`tNew HubSpot Contact [$($newHubSpotContact.properties.email)][$($newHubSpotContact.id)] created!"
                        $null = update-netSuiteContactInNetSuite -netSuiteContactId $thisNetSuiteContact.id -fieldHash @{custentitycustentity_hubspotid = $newHubSpotContact.id} -netsuiteParameters $netSuiteParameters
                        $thisNetSuiteContact = get-netSuiteContactFromNetSuite -contactId $thisNetSuiteContact.id -netsuiteParameters $netSuiteParameters
                        Write-Host -f DarkGreen "`t`tNetSuite Contact [$($thisNetSuiteContact.email)][$($thisNetSuiteContact.id)][$($thisNetSuiteContact.company.refName)][$($thisNetSuiteContact.company.id)] updated with new HobSpotId [$($newHubSpotContact.id)]"
                        }
                    catch{Write-Host -f Red $(get-errorSummary -errorToSummarise $_)}
                    }
                
                }
            }
        }
    else{
        #Yes - Has the Contact been updated more recently in HubSpot?
        $correspondingHubSpotContact = get-hubSpotContactById -apiKey $apiKey.HubApiKey -contactId $thisNetSuiteContact.custentitycustentity_hubspotid -ErrorAction SilentlyContinue
        if(![string]::IsNullOrWhiteSpace($correspondingHubSpotContact.id)){
            Write-Host -ForegroundColor DarkGreen "`tCorresponding HubSpot Contact [$($correspondingHubSpotContact.properties.email)][$($correspondingHubSpotContact.id)] found by HubSpotId"
            if((Get-Date $correspondingHubSpotContact.updatedAt) -gt (Get-Date $thisNetSuiteContact.lastModifiedDate)){
            #Yes - Update HubSpot > NetSuite
                Write-Host -ForegroundColor Green "`tHubSpot Contact [$($correspondingHubSpotContact.properties.email)][$($correspondingHubSpotContact.id)] updated more recently than NetSuite Contact [$($thisNetSuiteContact.email)][$($thisNetSuiteContact.id)][$($thisNetSuiteContact.company.refName)][$($thisNetSuiteContact.company.id)] - UPDATING HubSpot -> NetSuite"
                try{
                    $updatedNetSuiteContact = update-netSuiteContactInNetSuiteFromHubSpotObject -hubSpotContactObject $correspondingHubSpotContact -hubSpotApiKey $apiKey.HubApiKey -netsuiteParameters $netSuiteParameters -ErrorAction Stop #-Verbose
                    Write-Host -ForegroundColor DarkGreen "`tNetSuite Contact [$($thisNetSuiteContact.email)][$($thisNetSuiteContact.id)][$($thisNetSuiteContact.company.refName)][$($thisNetSuiteContact.company.id)] updated"
                    }
                catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
                }

            #No  - Update NetSuite > HubSpot
            elseif(($thisNetSuiteContact.lastModifiedDate -gt $correspondingHubSpotContact.properties.lastmodifiedinnetsuite) -and (Get-Date $thisNetSuiteContact.lastModifiedDate) -gt (Get-Date $correspondingHubSpotContact.updatedAt)){
                Write-Host -ForegroundColor Green "`tNetSuite Contact [$($thisNetSuiteContact.email)][$($thisNetSuiteContact.id)][$($thisNetSuiteContact.company.refName)][$($thisNetSuiteContact.company.id)] updated more recently than HubSpot Contact [$($correspondingHubSpotContact.properties.email)][$($correspondingHubSpotContact.id)] - UPDATING NetSuite -> HubSpot"
                try{
                    $updatedHubSpotContact = update-hubSpotObjectFromNetSuiteObject -apiKey $apiKey.HubApiKey -objectType contacts -netSuiteObject $thisNetSuiteContact -ErrorAction Stop
                    Write-Host -ForegroundColor DarkGreen "`tHubSpot Contact [$($correspondingHubSpotContact.properties.email)][$($correspondingHubSpotContact.id)] updated"
                    }
                catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
                }
            else{
                Write-Host -f DarkGreen "NetSuite Contact [$($thisNetSuiteContact.email)][$($thisNetSuiteContact.id)][$($thisNetSuiteContact.company.refName)][$($thisNetSuiteContact.company.id)] doesn't seem to have changed."
                }
            }

        else{
            Write-Warning "HubSpot Contact with HubSpotId [$($thisNetSuiteContact.custentitycustentity_hubspotid)] is missing from HubSpot (probably deleted) - REMOVING HubSpotId from NetSuite Contact [$($thisNetSuiteContact.email)][$($thisNetSuiteContact.id)][$($thisNetSuiteContact.company.refName)][$($thisNetSuiteContact.company.id)]"
            try{update-netSuiteContactInNetSuite -netSuiteContactId $thisNetSuiteContact.id -fieldHash @{custentitycustentity_hubspotid=$null} -netsuiteParameters $netSuiteParameters}
            catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
            }
        
        }
    }
#endregion

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

Stop-Transcript