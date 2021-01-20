function get-atpTokenResponse{
     [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [PSCustomObject]$aadAppCreds
        ,[parameter(Mandatory = $false)]
            [ValidateSet(“client_credentials”,”authorization_code”,"device_code")]
            [string]$grant_type = "client_credentials"
        ,[parameter(Mandatory = $false)]
            [string]$resource = "https://api.securitycenter.windows.com"
        )
    switch($grant_type){
        "authorization_code" {if(!$scope){$scope = "https://graph.microsoft.com/.default"}
            $authCode = get-graphAuthCode -clientID $aadAppCreds.ClientID -redirectUri $aadAppCreds.Redirect -scope $scope
            $ReqTokenBody = @{
                Grant_Type    = "authorization_code"
                #Scope         = $scope
                client_Id     = $aadAppCreds.ClientID
                Client_Secret = $aadAppCreds.Secret
                redirect_uri  = $aadAppCreds.Redirect
                code          = $authCode
                resource      = "https://graph.microsoft.com"
                }
            if($resource){$ReqTokenBody.Add("resource",$resource)}
            }
        "client_credentials" {
            $ReqTokenBody = @{
                Grant_Type    = "client_credentials"
                #Scope         = $scope
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

            $ReqTokenBody = @{
                grant_type    = "device_code"
                client_Id     = $aadAppCreds.ClientID
                code          = $response.device_code
                }

            }
        }

    $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$($aadAppCreds.TenantId)/oauth2/token" -Method POST -Body $ReqTokenBody
    $tokenResponse | Add-Member -MemberType NoteProperty -Name OriginalExpiryTime -Value $((Get-Date).AddSeconds($tokenResponse.expires_in))
    $tokenResponse
    }
function get-atpMachines(){
     [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        )

    invoke-atpGet -tokenResponse $tokenResponse -atpQuery "/machines" -Verbose:$VerbosePreference
    }
function get-atpSoftware(){
     [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        )

    invoke-atpGet -tokenResponse $tokenResponseIntuneBotAtp -atpQuery "/software" -Verbose:$VerbosePreference
    }
function invoke-atpGet(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$atpQuery
        ,[parameter(Mandatory = $false)]
            [switch]$firstPageOnly
        ,[parameter(Mandatory = $false)]
            [switch]$returnEntireResponse
        ,[parameter(Mandatory = $false)]
            [hashtable]$additionalHeaders
        ,[parameter(Mandatory = $false)]
            [switch]$useBetaEndpoint
        )
    if($useBetaEndpoint){$endpoint = "beta"}
    else{$endpoint = "v1.0"}
    $sanitisedatpQuery = $atpQuery.Trim("/")
    $headers = @{Authorization = "Bearer $($tokenResponse.access_token)"}
    if($additionalHeaders){
        $additionalHeaders.GetEnumerator() | %{
            $headers.Add($_.Key,$_.Value)
            }
        }
    #Write-Verbose $(stringify-hashTable -hashtable $headers -interlimiter "=" -delimiter ";")
    do{
        Write-Verbose "https://api.securitycenter.windows.com/api/$endpoint/$sanitisedatpQuery"
        $response = Invoke-RestMethod -Uri "https://api.securitycenter.windows.com/api/$endpoint/$sanitisedatpQuery" -ContentType "application/json; charset=utf-8" -Headers $headers -Method GET -Verbose:$VerbosePreference
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
        if(![string]::IsNullOrWhiteSpace($response.'@odata.nextLink')){$sanitisedatpQuery = $response.'@odata.nextLink'.Replace("https://atp.microsoft.com/$endpoint/","")}
        }
    #while($response.value.count -gt 0)
    while($response.'@odata.nextLink')
    if($returnEntireResponse){$response}
    else{$results}
    }
