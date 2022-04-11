function add-userToPowerBIWorkspace(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true,ParameterSetName = "aadObjectId")]
        [parameter(Mandatory = $true,ParameterSetName = "userPrincipalName")]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true,ParameterSetName = "aadObjectId")]
         [parameter(Mandatory = $true,ParameterSetName = "userPrincipalName")]
            [string]$workspaceID
        ,[parameter(Mandatory = $true,ParameterSetName = "aadObjectId")]
            [string]$aadObjectId
        ,[parameter(Mandatory = $true,ParameterSetName = "userPrincipalName")]
            [string]$userPrincipalName
        ,[parameter(Mandatory = $true,ParameterSetName = "aadObjectId")]
         [parameter(Mandatory = $true,ParameterSetName = "userPrincipalName")]
            [validateSet("Admin","Contributor","Member","Viewer")]$groupUserAccessRight
        ,[parameter(Mandatory = $true,ParameterSetName = "aadObjectId")]
         [parameter(Mandatory = $true,ParameterSetName = "userPrincipalName")]
            [validateSet("User","Group","App")]$PrincipalType

        )
    #Write-Verbose "https://api.powerbi.com/v1.0/myorg/groups/$($workspaceID)/users"
    Switch ($PsCmdlet.ParameterSetName){
        "aadObjectId" {$identifier = $aadObjectId}
        "userPrincipalName" {$identifier = $userPrincipalName}
        }
    Try{
        $result = invoke-powerBIPost -tokenResponse $tokenResponse -powerBIQuery "admin/groups/$($workspaceID)/users" -powerBIBodyHashtable @{"identifier"="$($identifier)";"groupUserAccessRight" = "$($groupUserAccessRight)";"principalType" = $PrincipalType} -Verbose
        }
    Catch{
        get-errorSummary $_
        }
    $result
    }
function get-powerBiAuthCode() {
     [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [string]$clientID
        ,[parameter(Mandatory = $true)]
            [string]$redirectUri
        ,[parameter(Mandatory = $false)]
            [string]$scope
        )

    $clientIDEncoded = [System.Web.HttpUtility]::UrlEncode($clientID)
    $redirectUriEncoded =  [System.Web.HttpUtility]::UrlEncode($redirectUri)
    $resourceEncoded = [System.Web.HttpUtility]::UrlEncode("https://analysis.windows.net/powerbi/api")
    $scopeEncoded = [System.Web.HttpUtility]::UrlEncode($scope) #"https://outlook.office.com/user.readwrite.all" "https://outlook.office.com/Directory.AccessAsUser.All"

    Add-Type -AssemblyName System.Windows.Forms
    if($scope){$url = "https://login.windows.net/common/oauth2/authorize/?response_type=code&redirect_uri=$redirectUriEncoded&client_id=$clientID&resource=$resourceEncoded&prompt=admin_consent&scope=$scopeEncoded"}
    #else{$url = "https://login.microsoftonline.com/common/oauth2/authorize?response_type=code&redirect_uri=$redirectUriEncoded&client_id=$clientID&resource=$resourceEncoded&prompt=admin_consent"}
    else{$url = "https://login.windows.net/common/oauth2/authorize/?response_type=code&redirect_uri=$redirectUriEncoded&client_id=$clientID&prompt=admin_consent"}
    #else{$url = "https://login.windows.net/common/oauth2/authorize/?response_type=code&redirect_uri=$redirectUriEncoded&client_id=$clientID&resource=$resourceEncoded&prompt=admin_consent"}
    $form = New-Object -TypeName System.Windows.Forms.Form -Property @{Width=440;Height=640}
    $web  = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width=420;Height=600;Url=($url -f ($Scope -join "%20")) }
    $docComp  = {
        $uri = $web.Url.AbsoluteUri        
        if ($uri -match "error=[^&]*|code=[^&]*") {$form.Close() }
        }
    $web.ScriptErrorsSuppressed = $true
    $web.Add_DocumentCompleted($docComp)
    $form.Controls.Add($web)
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() | Out-Null
    $queryOutput = [System.Web.HttpUtility]::ParseQueryString($web.Url.Query)
    $output = @{}
    foreach($key in $queryOutput.Keys){
        $output["$key"] = $queryOutput[$key]
        }
    $output
    }
function get-powerBITokenResponse{
     [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [PSCustomObject]$aadAppCreds
        ,[parameter(Mandatory = $false)]
            [ValidateSet(“client_credentials”,”authorization_code”,"device_code")]
            [string]$grant_type = "client_credentials"
        ,[parameter(Mandatory = $false)]
            [string]$resource = "https://analysis.windows.net/powerbi/api"
        )
        $scope = "https://analysis.windows.net/powerbi/api/.default" #
    switch($grant_type){
        "authorization_code" {if(!$scope){$scope = "https://graph.microsoft.com/.default"}
            $authCode = get-powerBiAuthCode -clientID $aadAppCreds.ClientID -redirectUri $aadAppCreds.Redirect #-scope $scope
            $ReqTokenBody = @{
                Grant_Type    = "authorization_code"
                Scope         = $scope
                client_Id     = $aadAppCreds.ClientID
                Client_Secret = $aadAppCreds.Secret
                redirect_uri  = $aadAppCreds.Redirect
                code          = $authCode.code
                resource      = "https://analysis.windows.net/powerbi/api"
                }
            #if($resource){$ReqTokenBody.Add("resource",$resource)}
            }
        "client_credentials" {
            $ReqTokenBody = @{
                Grant_Type    = "client_credentials"
                Scope         = $scope #could be https://analysis.windows.net/powerbi/api/.default - authorisation url to retrieve token
                client_Id     = $aadAppCreds.ClientID
                Client_Secret = $aadAppCreds.Secret
                }
            if($resource){$ReqTokenBody.Add("resource",$resource)}
            }
        "device_code" {
            $tenant = "anthesisllc.onmicrosoft.com"
            $authUrl = "https://login.microsoftonline.com/$tenant"
            $postParams = @{
                client_id = $aadAppCreds.ClientId
                Client_Secret = $aadAppCreds.Secret
                }
            if($resource){$postParams.Add("resource",$resource)}
            else{$postParams.Add("resource","https://graph.microsoft.com/")}
            $response = Invoke-RestMethod -Method POST -Uri "$authurl/oauth2/devicecode" -Body $postParams
            $code = ($response.message -split "code " | Select-Object -Last 1) -split " to authenticate."
            Set-Clipboard -Value $code

            Add-Type -AssemblyName System.Windows.Forms
            $form = New-Object -TypeName System.Windows.Forms.Form -Property @{ Width = 440; Height = 640 }
            $web = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{ Width = 440; Height = 600; Url = "https://www.microsoft.com/devicelogin" }
            $web.Add_DocumentCompleted($DocComp)
            $web.DocumentText
            $form.Controls.Add($web)
            $form.Add_Shown({ $form.Activate() })
            $web.ScriptErrorsSuppressed = $true
            $form.AutoScaleMode = 'Dpi'
            $form.text = "Graph API Authentication"
            $form.ShowIcon = $False
            $form.AutoSizeMode = 'GrowAndShrink'
            $Form.StartPosition = 'CenterScreen'
            $form.ShowDialog() | Out-Null     

    write-host "Hello ReqTokenBody:"
            $ReqTokenBody = @{
                grant_type    = "device_code"
                client_Id     = $aadAppCreds.ClientID
                client_secret = $aadAppCreds.Secret
                code          = $response.device_code
                }
    Write-Host $(stringify-hashTable -hashtable $ReqTokenBody -interlimiter "=" -delimiter "; ")
            $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$($aadAppCreds.TenantId)/oauth2/devicecode" -Method POST -Body $ReqTokenBody
            $tokenResponse | Add-Member -MemberType NoteProperty -Name OriginalExpiryTime -Value $((Get-Date).AddSeconds($tokenResponse.expires_in))
            return $tokenResponse
            }
        }

    write-host "Hello2!"
    Set-Variable dummy -Value $ReqTokenBody -Scope Global
    Write-Host $(stringify-hashTable -hashtable $ReqTokenBody -interlimiter "=" -delimiter "; ")
    $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$($aadAppCreds.TenantId)/oauth2/token" -Method POST -Body $ReqTokenBody
    $tokenResponse | Add-Member -MemberType NoteProperty -Name OriginalExpiryTime -Value $((Get-Date).AddSeconds($tokenResponse.expires_in))
    return $tokenResponse

    }
function get-usersFromPowerBIWorkspace(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$workspaceID
        )
    Write-Verbose "https://api.powerbi.com/v1.0/myorg/admin/groups/$($workspaceID)/users"
Try{
    $result = invoke-powerBIGet -tokenResponse $tokenResponse -powerBIQuery "admin/groups/$($workspaceID)/users" -Verbose
}
Catch{
    $error[0]
}
$result
    }
function invoke-powerBIGet(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$powerBIQuery
        ,[parameter(Mandatory = $false)]
            [switch]$firstPageOnly
        ,[parameter(Mandatory = $false)]
            [switch]$returnEntireResponse
        ,[parameter(Mandatory = $false)]
            [hashtable]$additionalHeaders
        )
    $endpoint = "v1.0/myorg" # only one version of app, there is a different REST API for dedicated capacity tenants
    $sanitisedpowerBIQuery = $powerBIQuery.Trim("/")
    $headers = @{Authorization = "Bearer $($tokenResponse.access_token)"}
    if($additionalHeaders){
        $additionalHeaders.GetEnumerator() | %{
            $headers.Add($_.Key,$_.Value)
            }
        }
    #Write-Verbose $(stringify-hashTable -hashtable $headers -interlimiter "=" -delimiter ";")
    do{
        Write-Verbose "https://api.powerbi.com/$endpoint/$sanitisedpowerBIQuery"
        $response = Invoke-RestMethod -Uri "https://api.powerbi.com/$endpoint/$sanitisedpowerBIQuery" -ContentType "application/json; charset=utf-8" -Headers $headers -Method GET -Verbose:$VerbosePreference
        if($response.value){
            $results += $response.value
            Write-Verbose "[$([int]$response.value.count)] results returned on this cycle, [$([int]$results.count)] in total"
            }
        elseif([string]::IsNullOrWhiteSpace($response) `
            -or ($response.'@odata.context' -and [string]::IsNullOrWhiteSpace($response.value) -and [string]::IsNullOrWhiteSpace($response.id))){ #If $response is $null, or if we get a response with a $null value
            Write-Verbose "[0] results returned on this cycle, [$([int]$results.count)] in total"
            }
        else{
            $results += $response
            Write-Verbose "[1] results returned on this cycle, [$([int]$results.count)] in total"
            }
        
        if($firstPageOnly){break}
        if(![string]::IsNullOrWhiteSpace($response.'@odata.nextLink')){$sanitisedpowerBIQuery = $response.'@odata.nextLink'.Replace("https://api.powerbi.com/$endpoint/","")}
        }
    #while($response.value.count -gt 0)
    while($response.'@odata.nextLink')
    if($returnEntireResponse){$response}
    else{$results}
    }
function invoke-powerBIPost(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$powerBIQuery
        ,[parameter(Mandatory = $true)]
            [Hashtable]$powerBIBodyHashtable
        ,[parameter(Mandatory = $false)]
            [hashtable]$additionalHeaders
        )

    $sanitisedpowerBIQuery = $powerBIQuery.Trim("/")
    $headers = @{Authorization = "Bearer $($tokenResponse.access_token)"}
    if($additionalHeaders){
        $additionalHeaders.GetEnumerator() | %{
            $headers.Add($_.Key,$_.Value)
            }
        }
    Write-Verbose "https://api.powerbi.com/v1.0/myorg/$sanitisedpowerBIQuery"
        
    $powerBIBodyJson = ConvertTo-Json -InputObject $powerBIBodyHashtable -Depth 10
    Write-Verbose $powerBIBodyJson
    $powerBIBodyJsonEncoded = [System.Text.Encoding]::UTF8.GetBytes($powerBIBodyJson)
    write-host "Hello there :)"
    Invoke-RestMethod -Uri "https://api.powerbi.com/v1.0/myorg/$sanitisedpowerBIQuery" -Body $powerBIBodyJsonEncoded -ContentType "application/json; charset=utf-8" -Headers $headers -Method Post
    }
function invoke-powerBIPut(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true,ParameterSetName = "BinaryFileStream")]
            [parameter(Mandatory = $true,ParameterSetName = "NormalRequest")]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true,ParameterSetName = "BinaryFileStream")]
            [parameter(Mandatory = $true,ParameterSetName = "NormalRequest")]
            [string]$powerBIQuery
        ,[parameter(Mandatory = $true,ParameterSetName = "BinaryFileStream")]
            $binaryFileStream
        ,[parameter(Mandatory = $true,ParameterSetName = "NormalRequest")]
            [Hashtable]$powerBIBodyHashtable
        ,[parameter(Mandatory = $false)]
            [hashtable]$additionalHeaders
        )

    $sanitisedpowerBIQuery = $powerBIQuery.Trim("/")
    $headers = @{Authorization = "Bearer $($tokenResponse.access_token)"}
    if($additionalHeaders){
        $additionalHeaders.GetEnumerator() | %{
            $headers.Add($_.Key,$_.Value)
            }
        }
    Write-Verbose "https://api.powerbi.com/v1.0/myorg/$($powerBIQuery)"
        
    if($binaryFileStream){
        $contentType = "text/plain"
        $bodyData = $binaryFileStream
        }
    elseif($powerBIBodyHashtable){
        $contentType = "application/json; charset=utf-8"
        $powerBIBodyJson = ConvertTo-Json -InputObject $powerBIBodyHashtable
        Write-Verbose $powerBIBodyJson
        $bodyData = [System.Text.Encoding]::UTF8.GetBytes($powerBIBodyJson)
        }

    Invoke-RestMethod -Uri "https://api.powerbi.com/v1.0/myorg/$($powerBIQuery)" -Body $bodyData -ContentType $contentType -Headers $headers -Method Put
    }
function invoke-powerBIPatch(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$powerBIQuery
        ,[parameter(Mandatory = $true)]
            [Hashtable]$powerBIBodyHashtable
        ,[parameter(Mandatory = $false)]
            [hashtable]$additionalHeaders
        )

    $sanitisedpowerBIQuery = $powerBIQuery.Trim("/")
    $headers = @{Authorization = "Bearer $($tokenResponse.access_token)"}
    if($additionalHeaders){
        $additionalHeaders.GetEnumerator() | %{
            $headers.Add($_.Key,$_.Value)
            }
        }
    Write-Verbose "https://api.powerbi.com/v1.0/myorg/$sanitisedGraphQuery"
        
    $powerBIBodyJson = ConvertTo-Json -InputObject $powerBIBodyHashtable
    Write-Verbose $powerBIBodyJson
    $powerBIBodyJsonEncoded = [System.Text.Encoding]::UTF8.GetBytes($powerBIBodyJson)
    
    Invoke-RestMethod -Uri "https://api.powerbi.com/v1.0/myorg/$sanitisedpowerBIQuery" -Body $powerBIBodyJsonEncoded -ContentType "application/json; charset=utf-8" -Headers $headers -Method Patch
    }
function invoke-powerBIDelete(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$powerBIQuery
        ,[parameter(Mandatory = $false)]
            [string]$powerBIBodyHashtable
        ,[parameter(Mandatory = $false)]
            [hashtable]$additionalHeaders
        )
    $sanitisedpowerBIQuery = $powerBIQuery.Trim("/")
    Write-Verbose "https://api.powerbi.com/v1.0/myorg/$sanitisedpowerBIQuery"

    $headers = @{Authorization = "Bearer $($tokenResponse.access_token)"}
    if($additionalHeaders){
        $additionalHeaders.GetEnumerator() | %{
            $headers.Add($_.Key,$_.Value)
            }
        }

    if($powerBIBodyHashtable){
        $powerBIBodyJson = ConvertTo-Json -InputObject $powerBIBodyHashtable
        Write-Verbose $powerBIBodyJson
        $powerBIBodyJsonEncoded = [System.Text.Encoding]::UTF8.GetBytes($powerBIBodyJson)
        $response = Invoke-RestMethod -Uri "https://api.powerbi.com/v1.0/myorg/$sanitisedpowerBIQuery" -ContentType "application/json; charset=utf-8" -Headers $headers -Method DELETE -Body $powerBIBodyJsonEncoded
        }
    else{
        $response = Invoke-RestMethod -Uri "https://api.powerbi.com/v1.0/myorg/$sanitisedPowerBIQuery" -ContentType "application/json; charset=utf-8" -Headers $headers -Method DELETE
        }
    if($response.value){$response.value}
    else{$response}
    }
function new-powerBIWorkspace(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$workspacename
        ,[parameter(Mandatory = $false)]
            [validateSet("v1","v2")]$version
        )
Switch($version){
    "v1" {$versionToApply = "groups"}
    "v2" {$versionToApply = "groups?workspaceV2=True"}
}
    Write-Verbose "https://api.powerbi.com/v1.0/myorg/$($versionToApply)"
Try{
    $result = invoke-powerBIPost -tokenResponse $powerBIBottokenResponse -powerBIQuery $versionToApply -powerBIBodyHashtable @{"name"="$($workspacename)"} -Verbose
}
Catch{
    $error[0]
}
$result
    }
function refresh-powerBIWorkspaceUserPermissions(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        )
    Write-Verbose "https://api.powerbi.com/v1.0/myorg/RefreshUserPermissions"
    $headers = @{Authorization = "Bearer $($tokenResponse.access_token)"}
Try{
    $result = Invoke-RestMethod -Uri "https://api.powerbi.com/v1.0/myorg/RefreshUserPermissions" -Headers $headers -Method Post
    Write-Warning "It takes about two minutes for the permissions to get refreshed. Before calling other APIs, wait for two minutes. User can call this API once per hour."
}
Catch{
    $error[0]
}
$result
    }
function remove-userFromPowerBIWorkspace(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true,ParameterSetName = "aadObjectId")]
        [parameter(Mandatory = $true,ParameterSetName = "userPrincipalName")]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true,ParameterSetName = "aadObjectId")]
        [parameter(Mandatory = $true,ParameterSetName = "userPrincipalName")]
            [string]$workspaceID
        ,[parameter(Mandatory = $true,ParameterSetName = "aadObjectId")]
            [string]$aadObjectId
        ,[parameter(Mandatory = $true,ParameterSetName = "userPrincipalName")]
            [string]$userPrincipalName
        )
#Small note: the other calls use "identifier" to cover email addresses and aadObjectIds, this call explicitly requires the "user" property, but this can be any principal type
Switch ($PsCmdlet.ParameterSetName){
    "aadObjectId" {$identifier = $aadObjectId}
    "userPrincipalName" {$identifier = $userPrincipalName}
}
Write-Verbose "https://api.powerbi.com/v1.0/myorg/groups/$($workspaceID)/users/$($identifier)"
Try{
    $result = invoke-powerBIDelete -tokenResponse $powerBIBottokenResponse -powerBIQuery "groups/$($workspaceID)/users/$($identifier)" -Verbose
}
Catch{
    $error[0]
}
$result
    }
function update-powerBIWorkspaceUserPermissions(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true,ParameterSetName = "aadObjectId")]
        [parameter(Mandatory = $true,ParameterSetName = "userPrincipalName")]
            [psobject]$tokenResponse 
       ,[parameter(Mandatory = $true,ParameterSetName = "aadObjectId")]
        [parameter(Mandatory = $true,ParameterSetName = "userPrincipalName")]
            [string]$workspaceID
       ,[parameter(Mandatory = $true,ParameterSetName = "aadObjectId")]
            [string]$aadObjectId
       ,[parameter(Mandatory = $true,ParameterSetName = "userPrincipalName")]
            [string]$userPrincipalName
       ,[parameter(Mandatory = $true,ParameterSetName = "aadObjectId")]
        [parameter(Mandatory = $true,ParameterSetName = "userPrincipalName")]
            [validateSet("Admin","Contributor","Member","Viewer")]$groupUserAccessRight
       ,[parameter(Mandatory = $true,ParameterSetName = "aadObjectId")]
        [parameter(Mandatory = $true,ParameterSetName = "userPrincipalName")]
            [validateSet("User","Group","App")]$PrincipalType
        )
    #Small note on this call, the documentation isn't clear on what is needed for this. You will need the identifier (upn/objectID), groupUserAccessRight (permission level), and principalType (user, group or app)
    Write-Verbose "https://api.powerbi.com/v1.0/myorg/groups/$($workspaceID)/users"
Switch ($PsCmdlet.ParameterSetName){
    "aadObjectId" {$identifier = $aadObjectId}
    "userPrincipalName" {$identifier = $userPrincipalName}
}
Try{
    $result = invoke-powerBIPut -tokenResponse $tokenResponse -powerBIQuery "groups/$($workspaceID)/users" -powerBIBodyHashtable @{"identifier"= "$($identifier)"; "groupUserAccessRight" = "$($groupUserAccessRight)";"principalType" = $PrincipalType} -Verbose
    Write-Warning "Permission changes may take a few minutes to reflect in the api."
}
Catch{
    $error[0]
}
$result
    }

