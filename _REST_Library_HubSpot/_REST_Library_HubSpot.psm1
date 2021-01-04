function get-hubSpotApiKey(){
    [cmdletbinding()]
    param()
    $encryptedCredsFile = "HubSync.txt"
    $placesToLook = @( #Figure out where to look
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
        ,[parameter(Mandatory = $false)]
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
                    )
                }
            }
        }
    switch ($PsCmdlet.ParameterSetName){
        "0FilterGroups" {
            $query = "/objects/$objectType`?archived=false&properties=$($propertiesToReturn -join ",")"
            invoke-hubSpotGet -apiKey $apiKey -query $query -firstPageOnly:$firstPageOnly
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
function get-hubSpotContactsFromCompanyId(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [string]$apiKey
        ,[parameter(Mandatory = $true)]
            [string]$hubspotCompanyId
        )
    
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
        $contacts += get-hubSpotObjects -apiKey $apiKey -objectType contacts -filterGroup1 @{filters=@($hubspotFilterById)}
        }
    $contacts
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
            [string]$apiVersion = "v3"
        )
    $sanitisedataQuery = $query.Trim("/")
    if(!$sanitisedataQuery.Contains("?")){$sanitisedataQuery+="?"}
    #if($sanitisedataQuery -notmatch "limit="){$sanitisedataQuery+="&limit=$pageSize"}
    do{
        Write-Verbose "https://api.hubapi.com/crm/$apiVersion/$sanitisedataQuery`&hapikey=$apiKey"
        try{
            $response = Invoke-RestMethod -Uri "https://api.hubapi.com/crm/v3/$sanitisedataQuery`&hapikey=$apiKey" -ContentType "application/json; charset=utf-8" -Method GET -Verbose:$VerbosePreference
            $results += $response.results
            Write-Verbose "[$($response.results.count)] results returned on this cycle, [$([int]$results.count)] in total"
            }
        catch{
            if(![string]::IsNullOrWhiteSpace($_.ErrorDetails.Message)){
                if (($_.ErrorDetails.Message | ConvertFrom-Json).Category -eq "RATE_LIMITS"){
                    $backOff++
                    Write-Warning "HubSpot Rate Limit reached - backing off for [$backOff] seconds"
                    Start-Sleep -Seconds $backOff
                    }
                }
            else{Write-Error $_}
            }
        
        if($firstPageOnly){break}
        if(![string]::IsNullOrWhiteSpace($response.paging.next)){$sanitisedataQuery = $response.paging.next.link.Replace("https://api.hubapi.com/crm/v3/","")}
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
    
    if([string]::IsNullOrEmpty($netSuiteObject.HubSpotId)){
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
                }
            if([string]::IsNullOrWhiteSpace($fieldHash["generic_email_address__c"])){
                $contacts = get-hubSpotContactsFromCompanyId -apiKey $apiKey -hubspotCompanyId $netSuiteObject.HubSpotId
                $mostRecentlyCreatedContact = $contacts | Sort-Object createdAt -Descending | select -First 1
                $fieldHash["generic_email_address__c"] = $mostRecentlyCreatedContact.properties.email
                }
            }
        "contacts"  {
            $firstName = $netSuiteObject.entityId.Split(" ")[0].Trim(" ")
            $fieldHash = [ordered]@{
                firstname = $firstName
                email = $netSuiteObject.email
                jobtitle = $netSuiteObject.title
                mobilephone = $netSuiteObject.mobilePhone
                phone = $netSuiteObject.officePhone
                lastmodifiedinnetsuite = $netSuiteObject.lastModifiedDate
                }
            if($netSuiteObject.entityId.Split(" ").Count -gt 1){
                $lastName = $netSuiteObject.entityId.Substring($firstName.length,$netSuiteObject.entityId.length - $firstName.length).Trim(" ")
                $fieldHash.Add("lastname",$lastName)
                }
            }
        }

        update-hubSpotObject -apiKey $apiKey -objectType $objectType -fieldHash $fieldHash -objectId $netSuiteObject.HubSpotId
    }
