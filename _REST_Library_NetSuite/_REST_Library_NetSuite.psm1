add-type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint, X509Certificate certificate,
        WebRequest request, int certificateProblem) {
        return true;
    }
}
"@
$AllProtocols = [System.Net.SecurityProtocolType]'Ssl3,Tls,Tls11,Tls12'
[System.Net.ServicePointManager]::SecurityProtocol = $AllProtocols
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

function add-netsuiteAccountToSqlCache{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [PSCustomObject]$netsuiteAccount 
        ,[parameter(Mandatory = $true)]
        [ValidateSet("Client","Supplier")]
        [string]$accountType
        ,[parameter(Mandatory = $true)]
        [System.Data.Common.DbConnection]$dbConnection
        )
    Write-Verbose "add-netsuiteAccountToSqlCache [$($netsuiteAccount.companyName)]"
    $sql = "SELECT TOP 1 AccountName, NsInteralId, LastModified FROM ACCOUNTS WHERE NsInteralId = '$($netsuiteAccount.Id)' ORDER BY LastModified Desc"
    Write-Verbose "`t$sql"
    $alreadyPresent = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $dbConnection
    if($netsuiteAccount.companyName -eq $alreadyPresent.AccountName){
        if($(Get-Date $netsuiteAccount.lastModifiedDate) -ne $(Get-Date $alreadyPresent.LastModified)){
            Write-Verbose "`tNsInteralId [$($netsuiteAccount.Id)] has been updated, but the name has not changed. Updating LastModified for existing record."
            $sql =  "UPDATE ACCOUNTS "
            $sql += "SET LastModified = $(sanitise-forSqlValue -value $netsuiteAccount.lastModifiedDate -dataType Date) "
            $sql += "WHERE NsInteralId = $(sanitise-forSqlValue -value $netsuiteAccount.id -dataType String) "
            $sql += "AND LastModified = $(sanitise-forSqlValue -value $alreadyPresent.LastModified -dataType Date)"
            }
        else{
            Write-Verbose "`tNsInteralId [$($netsuiteAccount.Id)] doesn't seem to have changed (probably caused by a lack of granularity in NetSuite's REST WHERE clauses). Not updating anything."
            }
        }
    else{
        if(!$alreadyPresent){Write-Verbose "`tNsInteralId [$($netsuiteAccount.Id)] not present in SQL, adding to [ACCOUNTS]"}
        else{Write-Verbose "`tNsInteralId [$($netsuiteAccount.Id)] CompanyName has changed from [$($alreadyPresent.AccountName)] to [$($netsuiteAccount.companyName)], adding new record to [ACCOUNTS]"}
        $now = $(Get-Date)
        $sql = "INSERT INTO ACCOUNTS (NsInteralId,NsExternalId,RecordType,AccountName,entityStatus,DateCreated,LastModified,IsDirty,DateCreatedInSql,DateModifiedInSql) VALUES ("
        $sql += $(sanitise-forSqlValue -value $netsuiteAccount.id -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $netsuiteAccount.accountNumber -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $accountType -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $netsuiteAccount.companyName -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $netsuiteAccount.entityStatus.refName -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $netsuiteAccount.dateCreated -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $netsuiteAccount.lastModifiedDate -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $true -dataType Boolean)
        $sql += ","+$(sanitise-forSqlValue -value $now -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $now -dataType Date)
        $sql += ")"
        Write-Verbose "`t$sql"
        $result = Execute-SQLQueryOnSQLDB -query $sql -queryType nonquery -sqlServerConnection $dbConnection
        if($result -eq 1){Write-Verbose "`t`tSUCCESS!"}
        else{Write-Verbose "`t`tFAILURE :( - Code: $result"}
        $result
        }
    }
function get-allNetSuiteClients(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $false)]
        [ValidatePattern('^?[\w+][=][\w+]')]
        [string]$query

        ,[parameter(Mandatory=$false)]
        [hashtable]$netsuiteParameters
        )

    Write-Verbose "`tget-allNetSuiteClients([$($query)])"
    if([string]::IsNullOrWhiteSpace($netsuiteParameters)){$netsuiteParameters = get-netsuiteParameters}

    $customers = invoke-netsuiteRestMethod -requestType GET -url "https://3487287-sb1.suitetalk.api.netsuite.com/rest/platform/v1/record/customer$query" -netsuiteParameters $netsuiteParameters #-Verbose 
    $customersEnumerated = [psobject[]]::new($customers.count)
    for ($i=0; $i -lt $customers.count;$i++) {
        $customersEnumerated[$i] = invoke-netsuiteRestMethod -requestType GET -url $customers.items[$i].links[0].href -netsuiteParameters $netsuiteParameters 
        }
    $customersEnumerated
    }
function get-netsuiteAuthHeaders(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [ValidateSet("GET","POST")]
        [string]$requestType
        
        ,[parameter(Mandatory = $true)]
        [ValidatePattern("http")]
        [string]$url
        
        ,[parameter(Mandatory=$true)]
        [hashtable]$oauthParameters

        ,[parameter(Mandatory=$true)]
        [string]$oauth_consumer_secret

        ,[parameter(Mandatory=$false)]
        [string]$oauth_token_secret

        ,[parameter(Mandatory=$true)]
        [string]$realm
        )

    Write-Verbose "get-netsuiteAuthHeaders()"
    $oauth_signature = get-oauthSignature -requestType $requestType -url $url -oauthParameters $oauthParameters -oauth_consumer_secret $oauth_consumer_secret -oauth_token_secret $oauth_token_secret

    #Irritatingly, we only include some predetermined oAuthParameters in the AuthHeader:
    $authHeaderString = ($oauthParameters.Keys | Sort-Object | ? {@("oauth_nonce","oauth_timestamp","oauth_consumer_key","oauth_token","oauth_signature_method","oauth_version") -contains $_} | % {
        "$_=`"$([uri]::EscapeDataString($oauthParameters[$_]))`""
        }) -join ","
    $authHeaderString += ",realm=`"$([uri]::EscapeDataString($realm))`""
    $authHeaderString += ",oauth_signature=`"$([uri]::EscapeDataString($oauth_signature))`""
    $authHeaders = @{"Authorization"="OAuth $authHeaderString"
        ;"Cache-Control"="no-cache"
#        ;"Accept"="application/swagger+json"
#        ;"Accept-Encoding"="gzip, deflate"
        }
    Write-Verbose "`$authHeaders = $($(
        $authHeaders.Keys | Sort-Object | % {
            "$_=$($authHeaders[$_])"
            }
        ) -join "&")"
    $authHeaders
    }
function get-netsuiteParameters(){
    [cmdletbinding()]
    Param()
    Write-Verbose "get-netsuiteParameters()"
    $placesToLook = @(
        "$env:USERPROFILE\Desktop\netsuite.txt"
        ,"$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\netsuite.txt"
        )
    for($i=0; $i -lt $placesToLook.Count; $i++){
        if(Test-Path $placesToLook[$i]){
            $pathToEncryptedCsv = $placesToLook[$i]
            continue
            }
        }
    if([string]::IsNullOrWhiteSpace($pathToEncryptedCsv)){
        Write-Error "NetSuite Paramaters CSV file not found in any of these locations: $($placesToLook -join ", ")"
        break
        }
    else{
        Write-Verbose "Importing NetSuite Paramaters fvrom [$pathToEncryptedCsv]"
        import-encryptedCsv $pathToEncryptedCsv
        }
    }
function get-oauthSignature(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [ValidateSet("GET","POST")]
        [string]$requestType
        
        ,[parameter(Mandatory = $true)]
        [ValidatePattern("http")]
        [string]$url
        
        ,[parameter(Mandatory=$true)]
        [hashtable]$oauthParameters

        ,[parameter(Mandatory=$true)]
        [string]$oauth_consumer_secret

        ,[parameter(Mandatory=$false)]
        [string]$oauth_token_secret
        )
    Write-Verbose "get-oauthSignature()"
    $requestType = $requestType.ToUpper()
                           
    $encodedUrl = [uri]::EscapeDataString($url.ToLower())

    $oAauthParamsString = (
        $oauthParameters.Keys | Sort-Object | % {
            if(@("realm","oauth_signature") -notcontains $_){
                "$_=$($oauthParameters[$_])"
                }
            }
        ) -join "&"
    $encodedOAuthParamsString = [uri]::EscapeDataString($oAauthParamsString)

    Write-Verbose "`tUnencoded base_string: [$($requestType + "&" + $url + "&" + $oAauthParamsString)]"
    $base_string = $requestType + "&" + $encodedUrl + "&" + $encodedOAuthParamsString
    $key = $oauth_consumer_secret + "&" + $oauth_token_secret
    Write-Verbose "`tEncoded base_string: [$base_string]"

    Switch($oauthParameters["oauth_signature_method"]){
        "HMAC-SHA1" {
            $cryptoFunction = new-object System.Security.Cryptography.HMACSHA1
            }
        "HMAC-SHA256" {
            $cryptoFunction = new-object System.Security.Cryptography.HMACSHA256
            }
        "HMAC-SHA384" {
            $cryptoFunction = new-object System.Security.Cryptography.HMACSHA384
            }
        "HMAC-SHA512" {
            $cryptoFunction = new-object System.Security.Cryptography.HMACSHA512
            }
        default {
            Write-Error "Unsupported oauth_signature_method [$_]"
            break
            }
        }

    $cryptoFunction.Key = [System.Text.Encoding]::ASCII.GetBytes($key)
    $oauth_signature = [System.Convert]::ToBase64String($cryptoFunction.ComputeHash([System.Text.Encoding]::ASCII.GetBytes($base_string)))
    Write-Verbose "`t`$oauth_signature = [$oauth_signature]"
    $oauth_signature
    }
function invoke-netsuiteRestMethod(){
    [cmdletbinding()]
    Param(
        [parameter(Mandatory = $true)]
        [ValidateSet("GET","POST")]
        [string]$requestType
        
        ,[parameter(Mandatory = $true)]
        [ValidatePattern("http")]
        [string]$url

        ,[parameter(Mandatory=$false)]
        [hashtable]$netsuiteParameters
        )

    if(!$netsuiteParameters){$netsuiteParameters = get-netsuiteParameters}
    
    if($url -match "\?"){
        $parameters = $url.Split("?")[1]
        $hostUrl = $url.Split("?")[0]
        }
    else{
        $hostUrl=$url
        $parameters = ""
        }

    $oauth_nonce = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes([System.DateTime]::Now.Ticks.ToString()))
    $oauth_timestamp = [int64](([datetime]::UtcNow)-(Get-Date "1970-01-01")).TotalSeconds

    $oAuthParamsForSigning = @{}
    #Add standard oAuth 1.0 parameters
    $oAuthParamsForSigning.Add("oauth_nonce",$oauth_nonce)
    $oAuthParamsForSigning.Add("oauth_timestamp",$oauth_timestamp)
    $oAuthParamsForSigning.Add("oauth_consumer_key",$netsuiteParameters.oauth_consumer_key)
    $oAuthParamsForSigning.Add("oauth_token",$netsuiteParameters.oauth_token)
    $oAuthParamsForSigning.Add("oauth_signature_method",$netsuiteParameters.oauth_signature_method)
    $oAuthParamsForSigning.Add("oauth_version",$netsuiteParameters.oauth_version)
    #Add parameters from url
    $parameters.Split("&") | % {
        if(![string]::IsNullOrWhiteSpace($_.Split("=")[0])){
            $oAuthParamsForSigning.Add([uri]::EscapeDataString($_.Split("=")[0]),[uri]::EscapeDataString($_.Split("=")[1])) #Weirdly, these extra paramaters have to be Encoded twice...
            #write-host -f Green "$([uri]::EscapeDataString($_.Split("=")[0]),[uri]::EscapeDataString($_.Split("=")[1]))"
            }
        }
    
    $netsuiteRestHeaders = get-netsuiteAuthHeaders -requestType $requestType -url $hostUrl -oauthParameters $oAuthParamsForSigning  -oauth_consumer_secret $netsuiteParameters["oauth_consumer_secret"] -oauth_token_secret $netsuiteParameters["oauth_token_secret"] -realm $netsuiteParameters["realm"]
    
    #Write-Host -f Green "Invoke-RestMethod -Uri $([uri]::EscapeUriString($url)) -Headers $(stringify-hashTable $netsuiteRestHeaders) -Method $requestType -Verbose -ContentType application/swagger+json"
    $response = Invoke-RestMethod -Uri $([uri]::EscapeUriString($url)) -Headers $netsuiteRestHeaders -Method $requestType -Verbose -ContentType "application/swagger+json"
    $response            
    }
