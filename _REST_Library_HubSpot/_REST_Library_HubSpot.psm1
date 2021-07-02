function add-hubSpotEventDataToContact(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [PSCustomObject]$contactObject
        ,[parameter(Mandatory = $false)]
            [Object[]]$eventsArray
        )
    $eventsArray | Sort-Object occurredAt -Descending | % {
        if($_.eventType -eq "e_visited_page"){$pageViewHistory += "$($_.occurredAt.Split("T")[0])`t$((($_.properties.hs_url -split '\?utm') -split '&utm')[0])`r`n"}
        if($_.eventType -eq "e_submitted_form"){$formSubmissionHistory += "$($_.occurredAt.Split("T")[0])`t$((($_.properties.hs_url -split '\?utm') -split '&utm')[0])`r`n"} #The "Message" part of the Form submission isn't currenlty returned via the REST API: https://developers.hubspot.com/docs/api/events/web-analytics#Event%20type%20selection%20and%20filters#:~:text=Event type selection and filters
        if($_.eventType -eq "e_attended_marketing_event"){$webinarHistory += "$($_.occurredAt.Split("T")[0])`t$($_.properties.hs_marketing_event)`r`n"} #The "Message" part of the Form submission isn't currenlty returned via the REST API: https://developers.hubspot.com/docs/api/events/web-analytics#Event%20type%20selection%20and%20filters#:~:text=Event type selection and filters
        }
    if(![string]::IsNullOrWhiteSpace($pageViewHistory))      {$contactObject | Add-Member -MemberType NoteProperty -Name pageViewHistory       -Value $pageViewHistory -Force}
    if(![string]::IsNullOrWhiteSpace($formSubmissionHistory)){$contactObject | Add-Member -MemberType NoteProperty -Name formSubmissionHistory -Value $formSubmissionHistory -Force}
    if(![string]::IsNullOrWhiteSpace($webinarHistory))       {$contactObject | Add-Member -MemberType NoteProperty -Name webinarHistory        -Value $webinarHistory  -Force}
    $contactObject
    }
function get-hubSpotApiKey(){
    [cmdletbinding()]
    param()
    $encryptedCredsFile = "HubSync.txt"
    $placesToLook = @( #Figure out where to look
        "$env:USERPROFILE\Downloads\$encryptedCredsFile"
        "$env:USERPROFILE\Desktop\$encryptedCredsFile"
        ,"$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\$encryptedCredsFile"
        )

    for($i=0; $i -lt $placesToLook.Count; $i++){ #Look for the file in each location until we find it
        if(Test-Path $placesToLook[$i]){
            $pathToEncryptedCsv = $placesToLook[$i]
            break
            }
        }
    if([string]::IsNullOrWhiteSpace($pathToEncryptedCsv)){ #Break if we can't find it
        Write-Error "Encrypted ApiKey file [$encryptedCredsFile] not found in any of these locations: $($placesToLook -join ", ")"
        break
        }
    else{ #Otherwise, import the file
        $apiKey = import-encryptedCsv -pathToEncryptedCsv $pathToEncryptedCsv
        }
    $apiKey
    }
function get-hubSpotCompanyById(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [string]$apiKey
        ,[parameter(Mandatory = $true)]
            [string]$companyId
        )

    $filterCompanyIdEquals = [ordered]@{
        propertyName="hs_object_id"
        operator="EQ"
        value=$companyId
        }
    get-hubSpotObjects -apiKey $apiKey -objectType companies -filterGroup1 @{filters=@($filterCompanyIdEquals)} -Verbose:$VerbosePreference

    }
function get-hubSpotContactById(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [string]$apiKey
        ,[parameter(Mandatory = $true)]
            [string]$contactId
        )

    $filterContactIdEquals = [ordered]@{
        propertyName="hs_object_id"
        operator="EQ"
        value=$contactId
        }
    get-hubSpotObjects -apiKey $apiKey -objectType contacts -filterGroup1 @{filters=@($filterContactIdEquals)} -Verbose:$VerbosePreference

    }
function get-hubSpotContactsFromCompanyId(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [string]$apiKey
        ,[parameter(Mandatory = $true)]
            [string]$hubspotCompanyId
        ,[parameter(Mandatory = $false)]
            [datetime]$updatedAfter
        ,[parameter(Mandatory = $false)]
            [switch]$includeContactsWithNoTimeStamp
        )
    if($updatedAfter){
        $filterContactUpdatedSinceLastSync = [ordered]@{
            propertyName="lastmodifiedinhubspot"
            operator="GT"
            #value = [Math]::Floor([decimal](Get-Date(Get-Date "2000-10-20T08:34:48.887Z").ToUniversalTime()-uformat "%s"))*1000 #Convert to UNIX Epoch time and add Milliseconds
            value = [Math]::Floor([decimal](Get-Date(Get-Date $updatedAfter).ToUniversalTime() -uformat "%s"))*1000 #Convert to UNIX Epoch time and add Milliseconds
            }
        }
    
    if($includeContactsWithNoTimeStamp){
        $filterContactsWithNoTimestamp = [ordered]@{
            propertyName="lastmodifiedinhubspot"
            operator="NOT_HAS_PROPERTY"
            }
        }
    
    $associations = invoke-hubSpotGet -apiKey $apiKey -query "/objects/companies/$hubspotCompanyId`?associations=contacts&paginateAssociations=true" -returnEntireResponse
    #$contactIds = Invoke-RestMethod -Uri "https://api.hubapi.com/crm-associations/v1/associations/$hubspotCompanyId/HUBSPOT_DEFINED/2?`&hapikey=$($apiKey.HubApiKey)" -ContentType "application/json; charset=utf-8" -Method GET -Verbose:$VerbosePreference #Old-skool

    [array]$contacts = @()
    $associations.associations.contacts.results.id | Select-Object | % {
        #$contacts += invoke-hubSpotGet -apiKey $apiKey.HubApiKey -query "/objects/contacts/1304117" -returnEntireResponse
        $hubspotFilterById = [ordered]@{
            propertyName="hs_object_id"
            operator="EQ"
            value = $_
            }
        if($updatedAfter){
            $contacts += get-hubSpotObjects -apiKey $apiKey -objectType contacts -filterGroup1 @{filters=@($hubspotFilterById,$filterContactUpdatedSinceLastSync)}
            if($includeContactsWithNoTimeStamp){$contacts += get-hubSpotObjects -apiKey $apiKey -objectType contacts -filterGroup1 @{filters=@($hubspotFilterById,$filterContactsWithNoTimestamp)}}
            }
        else{
            $contacts += get-hubSpotObjects -apiKey $apiKey -objectType contacts -filterGroup1 @{filters=@($hubspotFilterById)}
            }
        }
    $contacts
    }
function get-hubSpotEvents(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [string]$apiKey
        ,[parameter(Mandatory = $false)]
            [string]$hubspotContactId
        ,[parameter(Mandatory = $false)]
            [datetime]$occurredAfter
        ,[parameter(Mandatory = $false)]
            [datetime]$occurredBefore
        ,[parameter(Mandatory = $false)]
            [switch]$onlyGetIntegrationEvents
        )
    
    $eventQuery = "/events?objectType=contact"
    if($hubspotContactId){$eventQuery += "&objectId=$($hubspotContactId)"}
    if($occurredAfter){$eventQuery += "&occurredAfter=$(get-dateInIsoFormat -dateTime $occurredAfter -precision Milliseconds)"}
    if($occurredBefore){$eventQuery += "&occurredBefore=$(get-dateInIsoFormat -dateTime $occurredBefore -precision Milliseconds)"}
    if($onlyGetIntegrationEvents){$eventQuery += "&eventType=e_visited_page,e_submitted_form,e_registered_marketing_event"}
    [array]$events = invoke-hubSpotGet -apiKey $apiKey -query $eventQuery -api events
    $events
    }
function get-hubSpotMarketingEvents(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [string]$apiKey
        ,[parameter(Mandatory = $false)]
            [string]$hubspotContactId
        ,[parameter(Mandatory = $false)]
            [datetime]$occurredAfter
        ,[parameter(Mandatory = $false)]
            [datetime]$occurredBefore
        )
    
    $eventQuery = "marketing-events-beta/events/"
    if($hubspotContactId){$eventQuery += "&objectId=$($hubspotContactId)"}
    if($occurredAfter){$eventQuery += "&occurredAfter=$(get-dateInIsoFormat -dateTime $occurredAfter -precision Milliseconds)"}
    if($occurredBefore){$eventQuery += "&occurredBefore=$(get-dateInIsoFormat -dateTime $occurredBefore -precision Milliseconds)"}
    [array]$marketingEvents = invoke-hubSpotGet -apiKey $apiKey -query $eventQuery -api marketing
    $marketingEvents
   
    }
function get-hubSpotObjectProperties(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [string]$apiKey
        ,[parameter(Mandatory = $true)]
            [ValidateSet("companies", "contacts")]
            [string]$objectType
        )
    invoke-hubSpotGet -apiKey $apiKey -query "/properties/$objectType"
    }
function get-hubSpotObjects(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [string]$apiKey
        ,[parameter(Mandatory = $true)]
            [ValidateSet("companies", "contacts")]
            [string]$objectType
        ,[parameter(Mandatory = $true,ParameterSetName = "0FilterGroups")]
            [switch]$noFilter
        ,[parameter(Mandatory = $true,ParameterSetName = "1FilterGroups")]
            [parameter(Mandatory = $true,ParameterSetName = "2FilterGroups")]
            [parameter(Mandatory = $true,ParameterSetName = "3FilterGroups")]
            [ValidateCount(1,3)]
            [hashtable[]]$filterGroup1
        ,[parameter(Mandatory = $true,ParameterSetName = "2FilterGroups")]
            [parameter(Mandatory = $true,ParameterSetName = "3FilterGroups")]
            [ValidateCount(1,3)]
            [hashtable[]]$filterGroup2
        ,[parameter(Mandatory = $true,ParameterSetName = "3FilterGroups")]
            [ValidateCount(1,3)]
            [hashtable[]]$filterGroup3
        ,[parameter(Mandatory = $false,ParameterSetName = "1FilterGroups")]
            [parameter(Mandatory = $false,ParameterSetName = "2FilterGroups")]
            [parameter(Mandatory = $false,ParameterSetName = "3FilterGroups")]
            [hashtable]$sortPropertyNameAndDirection
        ,[parameter(Mandatory = $false)]
            [string[]]$propertiesToReturn
        ,[parameter(Mandatory = $false)]
            [switch]$firstPageOnly = $false
        ,[parameter(Mandatory = $false)]
            [switch]$returnEntireResponse
        ,[parameter(Mandatory = $false)]
            [int]$pageSize
        )
    
    if(!$propertiesToReturn){ #If unspecified, override HubSpot's default properties for each objecttype with sync-related properties
        switch ($objectType){
            "companies" {
                $propertiesToReturn = @(
                    "hs_object_id"
                    ,"name"
                    ,"hs_lastmodifieddate"
                    ,"netsuiteid"
                    ,"netsuite_sync_company_"
                    ,"netsuite_subsidiary"
                    ,"netsuite_sector"
                    ,"owneremail"
                    ,"generic_email_address__c"
                    ,"address"
                    ,"address2"
                    ,"city"
                    ,"state"
                    ,"zip"
                    ,"country"
                    ,"lastmodifiedinnetsuite"
                    ,"lastmodifiedinhubspot"
                    ,"domain"
                    ,"num_associated_contacts"
                    ,"netsuite_company_has_been_merged_or_deleted"
                    ,"hs_analytics_source"
                    ,"hs_analytics_num_page_views"
                    ,"hs_analytics_last_timestamp"
                    ,"first_conversion_event_name"
                    ,"recent_conversion_event_name"
                    ,"recent_conversion_date"
                    )
                }
            "contacts"  {
                $propertiesToReturn = @(
                    "salutation"
                    ,"firstname"
                    ,"lastname"
                    ,"jobtitle"
                    ,"email"
                    ,"other_email__c"
                    ,"phone"
                    ,"mobilephone"
                    ,"address"
                    ,"address2"
                    ,"city"
                    ,"state"
                    ,"zip"
                    ,"country"
                    ,"associatedcompanyid"
                    ,"netsuiteid"
                    ,"lastmodifiedinnetsuite"
                    ,"lastmodifiedinhubspot"
                    ,"hs_analytics_source"
                    ,"hs_analytics_source_data_1"
                    ,"hs_analytics_source_data_2"
                    ,"hs_analytics_first_referrer"
                    ,"recent_conversion_event_name"
                    ,"recent_conversion_date"
                    ,"opted_out_of_some_marketing_emails"
                    ,"last_webinar_attended_date"
                    ,"message"
                    ,"hs_analytics_num_page_views"
                    ,"hs_analytics_last_visit_timestamp"
                    ,"first_conversion_event_name"
                    ,"first_conversion_date"
                    ,"createdAt"
                    ,"event_or_webinar"
                    ,"timezone"
                    )
                }
            "events"  {
                $propertiesToReturn = @(
                    "salutation"
                    ,"firstname"
                    ,"lastname"
                    ,"jobtitle"
                    ,"email"
                    ,"other_email__c"
                    ,"phone"
                    ,"mobilephone"
                    ,"address"
                    ,"address2"
                    ,"city"
                    ,"state"
                    ,"zip"
                    ,"country"
                    ,"associatedcompanyid"
                    ,"netsuiteid"
                    ,"lastmodifiedinnetsuite"
                    ,"lastmodifiedinhubspot"
                    ,"hs_analytics_source"
                    ,"hs_analytics_source_data_1"
                    ,"hs_analytics_source_data_2"
                    ,"hs_analytics_first_referrer"
                    ,"recent_conversion_event_name"
                    ,"recent_conversion_date"
                    ,"opted_out_of_some_marketing_emails"
                    ,"last_webinar_attended_date"
                    ,"message"
                    ,"hs_analytics_num_page_views"
                    ,"hs_analytics_last_visit_timestamp"
                    ,"first_conversion_event_name"
                    ,"first_conversion_date"
                    )
                }
            }
        }
    switch ($PsCmdlet.ParameterSetName){
        "0FilterGroups" {
            $query = "/objects/$objectType`?archived=false&properties=$($propertiesToReturn -join ",")"
            invoke-hubSpotGet -apiKey $apiKey -query $query -firstPageOnly:$firstPageOnly -pageSize $pageSize
            }
        default {
            $query = "/objects/$objectType/search?archived=false"
            $filter = @($filterGroup1)
            if($filterGroup2){$filter += $filterGroup2}
            if($filterGroup3){$filter += $filterGroup3}
            $bodyHashTable = @{
                filterGroups=$filter
                properties=$propertiesToReturn
                }
            if($sortPropertyNameAndDirection){
                $bodyHashTable.Add("sorts",@($sortPropertyNameAndDirection))
                }
            invoke-hubSpotPost -apiKey $apiKey -query $query -bodyHashtable $bodyHashTable -firstPageOnly:$firstPageOnly -pageSize $pageSize
            }
        }

    }
function invoke-hubSpotGet(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [string]$apiKey
        ,[parameter(Mandatory = $true)]
            [string]$query
        ,[parameter(Mandatory = $false)]
            [switch]$firstPageOnly
        ,[parameter(Mandatory = $false)]
            [switch]$returnEntireResponse
        ,[parameter(Mandatory = $false)]
            [int]$pageSize = 100
        ,[parameter(Mandatory = $false)]
            [ValidateSet("v1", "v2", "v3")]
            [string]$apiVersion = "v3"
        ,[parameter(Mandatory = $false)]
            [ValidateSet("analytics", "crm", "events", "reports","marketing","content/api","forms","form-integrations")]
            [string]$api = "crm"
        )
    $sanitisedataQuery = $query.Trim("/")
    if(!$sanitisedataQuery.Contains("?")){$sanitisedataQuery+="?"}
    $backOff = 1
    #if($sanitisedataQuery -notmatch "limit="){$sanitisedataQuery+="&limit=$pageSize"}
    do{
        Write-Verbose "https://api.hubapi.com/$api/$apiVersion/$sanitisedataQuery`&hapikey=$apiKey"
        try{
            $response = Invoke-RestMethod -Uri "https://api.hubapi.com/$api/$apiVersion/$sanitisedataQuery`&hapikey=$apiKey" -ContentType "application/json; charset=utf-8" -Method GET -Verbose:$VerbosePreference
            $results += $response.results
            Write-Verbose "[$($response.results.count)] results returned on this cycle, [$([int]$results.count)] in total"
            }
        catch{
            if(![string]::IsNullOrWhiteSpace($_.ErrorDetails.Message)){
                #if (($_.ErrorDetails.Message | ConvertFrom-Json).Category -eq "RATE_LIMITS"){
                if ($_.ErrorDetails.Message -match "RATE_LIMITS"){
                    $backOff++
                    Write-Warning "HubSpot Rate Limit reached - backing off for [$backOff] seconds"
                    Start-Sleep -Seconds $backOff
                    }
                else{Write-Error $_}
                }
            else{Write-Error $_}
            }
        
        if($firstPageOnly){break}
        if(![string]::IsNullOrWhiteSpace($response.paging.next)){
            if($apiVersion -eq "v1"){
                $sanitisedataQuery = $sanitisedataQuery -replace '\?after=[\w]+',""
                $sanitisedataQuery = $sanitisedataQuery -replace '\&limit=[\w]+',""
                $sanitisedataQuery += $response.paging.next.link}
            else{$sanitisedataQuery = $response.paging.next.link.Replace("https://api.hubapi.com/$api/$apiVersion/","")}
            }
        }
    #while($response.value.count -gt 0)
    while($response.paging.next.link)
    if($returnEntireResponse){$response}
    else{$results}
    }
function invoke-hubSpotPatch(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [string]$apiKey
        ,[parameter(Mandatory = $true)]
            [string]$query
        ,[parameter(Mandatory = $true)]
            [Hashtable]$bodyHashtable
        )

    $sanitisedataQuery = $query.Trim("/")
    $bodyJson = ConvertTo-Json -InputObject $bodyHashtable -Depth 10
    Write-Verbose "[https://api.hubapi.com/crm/v3/$sanitisedataQuery`&hapikey=$apiKey] [$bodyJson]"
    Invoke-RestMethod -Uri "https://api.hubapi.com/crm/v3/$sanitisedataQuery`&hapikey=$apiKey" -ContentType "application/json; charset=utf-8" -Method PATCH -Body $bodyJson -Verbose:$VerbosePreference
    }
function invoke-hubSpotPost(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [string]$apiKey
        ,[parameter(Mandatory = $true)]
            [string]$query
        ,[parameter(Mandatory = $true)]
            [Hashtable]$bodyHashtable
        ,[parameter(Mandatory = $false)]
            [int]$pageSize
        ,[parameter(Mandatory = $false)]
            [switch]$firstPageOnly
        ,[parameter(Mandatory = $false)]
            [switch]$returnEntireResponse
        )
    $sanitisedataQuery = $query.Trim("/")
    if($bodyHashtable.Keys -notcontains "limit" -and $pageSize){$bodyHashtable.Add("limit",$pageSize)}
    $bodyJson = ConvertTo-Json -InputObject $bodyHashtable -Depth 10

    Write-Verbose "[https://api.hubapi.com/crm/v3/$sanitisedataQuery`&hapikey=$apiKey] [$bodyJson]"
    do{
        try{
            $response = Invoke-RestMethod -Uri "https://api.hubapi.com/crm/v3/$sanitisedataQuery`&hapikey=$apiKey" -ContentType "application/json; charset=utf-8" -Method POST -Body $bodyJson -Verbose:$VerbosePreference
            $results += $response.results
            Write-Verbose "[$($response.results.count)] results returned on this cycle, [$([int]$results.count)]/[$([int]$response.total)] in total"
            }
        catch{
            if (($_.ErrorDetails.Message | ConvertFrom-Json).Category -eq "RATE_LIMITS"){
                $backOff++
                Write-Warning "HubSpot Rate Limit reached - backing off for [$backOff] seconds"
                Start-Sleep -Seconds $backOff
                }
            }
        if($firstPageOnly){break}
        #Write-Verbose $response.paging
        if($bodyHashtable.Keys -contains "after"){$bodyHashtable["after"] = $response.paging.next.after}
        else{$bodyHashtable.Add("after",$response.paging.next.after)}
        $bodyJson = ConvertTo-Json -InputObject $bodyHashtable -Depth 10
        }
    #while($response.value.count -gt 0)
    while(![string]::IsNullOrWhiteSpace($response.paging.next.after))
    if($returnEntireResponse){$response}
    else{$results}
    }
function new-hubspotCompany(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [string]$apiKey
        ,[parameter(Mandatory = $true)]
            [string]$companyName
        ,[parameter(Mandatory = $false)]
            [hashtable]$fieldHash
        )

    if($fieldHash.Keys -notcontains "name"){$fieldHash.Add("name",$companyName)}
    else{$fieldHash["name"] = $companyName}

    invoke-hubSpotPost -apiKey $apiKey -query "/objects/companies?" -bodyHashtable @{"properties"=$fieldHash} -returnEntireResponse
    }
function new-hubspotCompanyFromNetsuiteCompany(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [string]$apiKey
        ,[parameter(Mandatory = $true)]
            [psobject]$netSuiteCompany
        )
    
    $fieldHash = [ordered]@{
        name = $netSuiteCompany.companyName
        netsuite_sector = $netSuiteCompany.custentity_ant_clientsector.refName
        netsuite_subsidiary = $netSuiteCompany.subsidiary.refName
        netsuiteid = $netSuiteCompany.id
        generic_email_address__c = $netSuiteCompany.email
        netsuite_sync_company_ = $true
        address = $netSuiteCompany.shipAddr1
        address2 = $netSuiteCompany.shipAddr2
        city = $netSuiteCompany.shipCity
        country = $netSuiteCompany.shipCountry
        state = $netSuiteCompany.shipState
        zip = $netSuiteCompany.shipZip
        lastmodifiedinnetsuite = $netSuiteCompany.lastModifiedDate
        lastmodifiedinhubspot = $(Get-Date (Get-Date).ToUniversalTime() -Format o)
        }
    new-hubspotCompany -apiKey $apiKey -companyName $netSuiteCompany.companyName -fieldHash $fieldHash
    }
function new-hubspotContact(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [string]$apiKey
        ,[parameter(Mandatory = $true)]
            [string]$email
        ,[parameter(Mandatory = $true)]
            [string]$associatedcompanyid
        #,[parameter(Mandatory = $true)]
        #    [string]$netsuiteSubsidiary
        ,[parameter(Mandatory = $false)]
            [hashtable]$fieldHash
        )

    if($fieldHash.Keys -notcontains "email"){$fieldHash.Add("email",$email)}
    else{$fieldHash["email"] = $email}
    if($fieldHash.Keys -notcontains "associatedcompanyid"){$fieldHash.Add("associatedcompanyid",$associatedcompanyid)}
    else{$fieldHash["associatedcompanyid"] = $associatedcompanyid}
    #if($fieldHash.Keys -notcontains "netsuite_subsidiary"){$fieldHash.Add("netsuite_subsidiary",$netsuiteSubsidiary)}
    #else{$fieldHash["netsuite_subsidiary"] = $netsuiteSubsidiary}

    
    invoke-hubSpotPost -apiKey $apiKey -query "/objects/contacts?" -bodyHashtable @{"properties"=$fieldHash} -returnEntireResponse
    }
function new-hubspotContactFromNetsuiteContact(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [string]$apiKey
        ,[parameter(Mandatory = $true)]
            [psobject]$netSuiteContact
        ,[parameter(Mandatory = $true,ParameterSetName="IdProvided")]
            [string]$hubSpotCompanyId
        ,[parameter(Mandatory = $true,ParameterSetName="FindParentId")]
            [psobject]$netSuiteParams
        )

    switch ($PsCmdlet.ParameterSetName){
        "FindParentId"    {
            try{
                $hubSpotFilterByNetSuiteId = [ordered]@{
                        propertyName="netsuiteid"
                        operator="EQ"
                        value=$netSuiteContact.company.id
                        }
                $thisContactsHubSpotCompany  = get-hubSpotObjects -apiKey $apiKey -objectType companies -filterGroup1 @{filters=@($hubSpotFilterByNetSuiteId)}
                if(![string]::IsNullOrEmpty($thisContactsHubSpotCompany.id)){
                    $hubSpotCompanyId = $thisContactsHubSpotCompany.id
                    }
                else{
                    $thisContactsNetSuiteCompany = get-netSuiteClientsFromNetSuite -clientId $netSuiteContact.company.id -netsuiteParameters $netSuiteParams -ErrorAction Stop
                    $hubSpotCompanyId = $thisContactsNetSuiteCompany.HubSpotId
                    }
                }
            catch{
                #Write-Error "Error retrieving NetSuite Company with Id [$($netSuiteContact.company.id)] for NetSuite Contact [$($netSuiteContact.entityId)][$($netSuiteContact.id)]"
                $_
                return
                }
            }
        }
    
    if([string]::IsNullOrEmpty($netSuiteContact.firstName)){$firstName = $netSuiteContact.entityId.Split(" ")[0].Trim(" ")}
    else{$firstName = $netSuiteContact.firstName.Trim(" ")}
    if([string]::IsNullOrEmpty($netSuiteContact.lastName) -and $netSuiteContact.entityId.Split(" ").Count -gt 1){$lastName = $netSuiteContact.entityId.Split(" ")[1].Trim(" ")} #Taking 2nd name rather than _Last_ name to please the Spaniards.
    else{$lastName = $netSuiteContact.lastName.Trim(" ")}

    $fieldHash = [ordered]@{
        firstname = $firstName
        lastname = $lastName
        email = $netSuiteContact.email
        jobtitle = $netSuiteContact.title
        netsuiteid = $netSuiteContact.id
        associatedcompanyid = $hubSpotCompanyId
        #netsuite_subsidiary = $netSuiteContact.subsidiary.refName
        lastmodifiedinnetsuite = $netSuiteContact.lastModifiedDate
        lastmodifiedinhubspot = $(Get-Date (Get-Date).ToUniversalTime() -Format o)
        }

    new-hubspotContact -apiKey $apiKey -email $netSuiteContact.email -associatedcompanyid $hubSpotCompanyId -fieldHash $fieldHash -Verbose:$VerbosePreference
     
    }
function new-hubSpotFilterById(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [string]$hubSpotId
        )
    [ordered]@{
        propertyName="hs_object_id"
        operator="EQ"
        value = $hubSpotId
        }
    }
function test-hubSpotFilter(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [ValidateCount(1,3)]
            [hashtable]$filter
        )
    $requiredKeys = @("propertyName","operator","value")
    $requiredKeys | % {
        if(!$filter.ContainsKey($_)){[array]$missing+=$_}
        }
    if($missing){Write-Warning "HubSpot filter is missing the requred "}
    }
function update-hubSpotObject(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [string]$apiKey
        ,[parameter(Mandatory = $true)]
            [ValidateSet("companies", "contacts")]
            [string]$objectType
        ,[parameter(Mandatory = $true)]
            [string]$objectId
        ,[parameter(Mandatory = $true)]
            [hashtable]$fieldHash = @{}
        )

    invoke-hubSpotPatch -apiKey $apiKey -query "/objects/$objectType/$objectId`?" -bodyHashtable @{properties=$fieldHash} -Verbose:$VerbosePreference
    }
function update-hubSpotObjectFromNetSuiteObject(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [string]$apiKey
        ,[parameter(Mandatory = $true)]
            [ValidateSet("companies", "contacts")]
            [string]$objectType
        ,[parameter(Mandatory = $true)]
            [psobject]$netSuiteObject
        )
    
    if([string]::IsNullOrEmpty($netSuiteObject.custentitycustentity_hubspotid)){
        Write-Error "NetSuite Object [$($netSuiteObject.id)] is missing a HubSpotId - cannot update HubSpot [$objectType]"
        return
        }

    switch($objectType){
        "companies" {
            $fieldHash = [ordered]@{
                name = $netSuiteObject.companyName
                netsuite_sector = $netSuiteObject.custentity_ant_clientsector.refName
                netsuite_subsidiary = $netSuiteObject.subsidiary.refName
                netsuiteid = $netSuiteObject.id
                generic_email_address__c = $netSuiteObject.email
                netsuite_sync_company_ = $true
                address = $netSuiteObject.shipAddr1
                address2 = $netSuiteObject.shipAddr2
                city = $netSuiteObject.shipCity
                country = $netSuiteObject.shipCountry
                state = $netSuiteObject.shipState
                zip = $netSuiteObject.shipZip
                lastmodifiedinnetsuite = $netSuiteObject.lastModifiedDate
                lastmodifiedinhubspot = $(Get-Date (Get-Date).ToUniversalTime() -Format o)
                companystatusinnetsuite = $netSuiteObject.entityStatus.refName
                }
            if([string]::IsNullOrWhiteSpace($fieldHash["generic_email_address__c"])){
                $contacts = get-hubSpotContactsFromCompanyId -apiKey $apiKey -hubspotCompanyId $netSuiteObject.HubSpotId
                $mostRecentlyCreatedContact = $contacts | Sort-Object createdAt -Descending | select -First 1
                $fieldHash["generic_email_address__c"] = $mostRecentlyCreatedContact.properties.email
                }
            }
        "contacts"  {
            if([string]::IsNullOrEmpty($netSuiteObject.firstName)){$firstName = $netSuiteObject.entityId.Split(" ")[0].Trim(" ")}
            else{$firstName = $netSuiteObject.firstName.Trim(" ")}
            if([string]::IsNullOrEmpty($netSuiteObject.lastName) -and $netSuiteObject.entityId.Split(" ").Count -gt 1){$lastName = $netSuiteObject.entityId.Split(" ")[$($netSuiteObject.entityId.Split(" ").Count)-1].Trim(" ")}
            else{$lastName = $netSuiteObject.lastName.Trim(" ")}

            $fieldHash = [ordered]@{
                firstname = $firstname
                lastname = $lastName
                email = $netSuiteObject.email
                jobtitle = $netSuiteObject.title
                mobilephone = $netSuiteObject.mobilePhone
                phone = $netSuiteObject.officePhone
                lastmodifiedinnetsuite = $netSuiteObject.lastModifiedDate
                lastmodifiedinhubspot = $(Get-Date (Get-Date).ToUniversalTime() -Format o)
                }
            }
        }

        update-hubSpotObject -apiKey $apiKey -objectType $objectType -fieldHash $fieldHash -objectId $netSuiteObject.custentitycustentity_hubspotid
    }
