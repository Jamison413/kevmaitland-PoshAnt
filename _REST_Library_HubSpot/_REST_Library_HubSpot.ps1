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
            [string[]]$propertiesToReturn
        ,[parameter(Mandatory = $false)]
            [switch]$firstPageOnly = $false
        ,[parameter(Mandatory = $false)]
            [switch]$returnEntireResponse
        ,[parameter(Mandatory = $false)]
            [int]$pageSize = 100
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
            invoke-hubSpotPost -apiKey $apiKey -query $query -bodyHashtable @{filterGroups=$filter;properties=$propertiesToReturn} -firstPageOnly:$firstPageOnly
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
        )
    $sanitisedataQuery = $query.Trim("/")
    if(!$sanitisedataQuery.EndsWith("?")){$sanitisedataQuery+="?"}
    if($sanitisedataQuery -notmatch "limit="){$sanitisedataQuery+="&limit=$pageSize"}
    do{
        Write-Verbose "https://api.hubapi.com/crm/v3/$sanitisedataQuery`&hapikey=$apiKey"
        try{
            $response = Invoke-RestMethod -Uri "https://api.hubapi.com/crm/v3/$sanitisedataQuery`&hapikey=$apiKey" -ContentType "application/json; charset=utf-8" -Method GET -Verbose:$VerbosePreference
            $results += $response.results
            Write-Verbose "[$($response.results.count)] results returned on this cycle, [$([int]$results.count)] in total"
            }
        catch{
            if (($_.ErrorDetails.Message | ConvertFrom-Json).Category -eq "RATE_LIMITS"){
                $backOff++
                Write-Warning "Rate Limit reached - backing off for [$backOff] seconds"
                Start-Sleep -Seconds $backOff
                }
            }
        
        if($firstPageOnly){break}
        if(![string]::IsNullOrWhiteSpace($response.paging.next)){$sanitisedataQuery = $response.paging.next.link.Replace("https://api.hubapi.com/crm/v3/","")}
        }
    #while($response.value.count -gt 0)
    while($response.paging.next.link)
    if($returnEntireResponse){$response}
    else{$results}
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
            [int]$pageSize = 100
        ,[parameter(Mandatory = $false)]
            [switch]$firstPageOnly
        )
    $sanitisedataQuery = $query.Trim("/")
    if($bodyHashtable.Keys -notcontains "limit"){$bodyHashtable.Add("limit",$pageSize)}
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
                Write-Warning "Rate Limit reached - backing off for [$backOff] seconds"
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

