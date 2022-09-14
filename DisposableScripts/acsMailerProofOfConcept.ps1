function get-graphTokenResponse{
    [cmdletbinding()]
   param(
       [parameter(Mandatory = $true)]
           [PSCustomObject]$aadAppCreds
       ,[parameter(Mandatory = $false)]
           [ValidateSet(“client_credentials”,”authorization_code”,"device_code","certificate")]
           [string]$grant_type = "client_credentials"
       ,[parameter(Mandatory = $false)]
           [string]$scope = "https://graph.microsoft.com/.default"
       ,[parameter(Mandatory = $false)]
           [System.Security.Cryptography.X509Certificates.X509Certificate]$cert
       )
   switch($grant_type){
       "authorization_code" {if(!$scope){$scope = "https://graph.microsoft.com/.default"}
           $authCode = get-graphAuthCode -clientID $aadAppCreds.ClientID -redirectUri $aadAppCreds.Redirect -scope $scope
           $ReqTokenBody = @{
               Grant_Type    = "authorization_code"
               Scope         = $scope
               client_Id     = $aadAppCreds.ClientID
               Client_Secret = $aadAppCreds.Secret
               redirect_uri  = $aadAppCreds.Redirect
               code          = $authCode
               #resource      = "https://graph.microsoft.com"
               }
           }
       "client_credentials" {
           $ReqTokenBody = @{
               Grant_Type    = "client_credentials"
               Scope         = $scope
               client_Id     = $aadAppCreds.ClientID
               Client_Secret = $aadAppCreds.Secret
               }
           }
       "device_code" {
           $tenant = "anthesisllc.onmicrosoft.com"
           $authUrl = "https://login.microsoftonline.com/$tenant"
           $postParams = @{
               resource = "https://graph.microsoft.com/"
               client_id = $aadAppCreds.ClientId
               }
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
       "certificate" {
           if($null -eq $cert){$cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object {$_.Subject -match $(whoami /upn)}}
           $clientAssertion = new-clientAssertionWithCertificate -X509cert $cert -clientId $aadAppCreds.ClientID -tenantId $aadAppCreds.TenantId -resource "https://graph.microsoft.com" -loginEndpoint "https://login.microsoftonline.com"
           #Write-Verbose "`$clientAssertion = [$($clientAssertion)]"
           #This also works, but's different from the way our other otkens work and it's more opaque:
           <#$ClientApplicationBuilder = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($aadAppCreds.ClientID)
           [void]$ClientApplicationBuilder.WithAuthority($("https://login.microsoftonline.com/$($aadAppCreds.TenantId)"))
           [void]$ClientApplicationBuilder.WithCertificate($myCert)
           $confidentialClientApplication = $ClientApplicationBuilder.Build()
           #[Microsoft.Identity.Client.AuthenticationResult] $authResult  = $null
           $AquireTokenParameters = $confidentialClientApplication.AcquireTokenForClient([string[]]$scope)
           try {
               [Microsoft.Identity.Client.AuthenticationResult]$authResult = $AquireTokenParameters.ExecuteAsync().GetAwaiter().GetResult()
               $authResult | Add-Member -MemberType NoteProperty -Name access_token -Value $authResult.AccessToken
           }
           catch {
               get-errorSummary $_
           }
           #>
           $ReqTokenBody = @{
               Grant_Type    = "client_credentials"
               Scope         = $scope
               client_Id     = $aadAppCreds.ClientID
               client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
               client_assertion = $clientAssertion
               }
           Write-Verbose "`$ReqTokenBody:"
           $ReqTokenBody.Keys | ForEach-Object {
                   Write-Verbose "`t[$_]`t`t[$($ReqTokenBody[$_])]"
           }
           }
       }

   $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$($aadAppCreds.TenantId)/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody
   $tokenResponse | Add-Member -MemberType NoteProperty -Name OriginalExpiryTime -Value $((Get-Date).AddSeconds($tokenResponse.expires_in))
   $tokenResponse
   }
function new-clientAssertionWithCertificate{
[cmdletbinding()]
param(
    [parameter(Mandatory = $true)]
    [System.Security.Cryptography.X509Certificates.X509Certificate]$X509cert
    ,[parameter(Mandatory = $true)]
    [string]$clientId
    ,[parameter(Mandatory = $true)]
    [string]$tenantId
    ,[parameter(Mandatory = $false)]
    [string]$resource='https://graph.microsoft.com'
    ,[parameter(Mandatory = $false)]
    [string]$loginEndpoint='https://login.microsoftonline.com'
    )

<#
.SYNOPSIS
Generates a signed client_assertion using the provided certificate, ready to present to Graph API for authentication.
**REQUIRES PowerShell 5**

.DESCRIPTION
#Adapted from https://gist.github.com/jformacek/aecc4f379b88b3a330ee19b045252462
There are alternative methods for authenticating with Graph using certificates (see get-graphTokenResponse), but this is in-keping with our 5.1 codebase.

.PARAMETER X509cert
A certificate with a Private Key (to sign the assertion)
.PARAMETER clientId
The ClientId/ApplicationId of the App Registration to authenticate with
.PARAMETER tenantId
The Id of Azure tenant to authenticate with
.PARAMETER resource
Future compatibility to support alternative resources (default is [https://graph.microsoft.com])
.PARAMETER loginEndpoint
Future compatibility to support alternative auth endpoints (default is [https://login.microsoftonline.com])
#>
    

#load required assembly that is not loaded by default by PowerShell
[System.Reflection.Assembly]::LoadWithPartialName('system.identitymodel') | Out-Null

#retrieve certificate from cert store
#$cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object {$_.Thumbprint -eq $certThumbprint}[0]

# Create base64 hash of the certificate
$certHash = [System.Convert]::ToBase64String($X509cert.GetCertHash())

#JWT expiration timestamp - valid for 5 minutes - just to allow token request to complete
$StartDate = (Get-Date "1970-01-01T00:00:00Z" ).ToUniversalTime()
$JWTExpiration = [int]((New-TimeSpan -Start $StartDate -End (Get-Date).ToUniversalTime().AddMinutes(5)).TotalSeconds)

#JWT Start timestamp - optional
$NotBefore = [int]((New-TimeSpan -Start $StartDate -End ((Get-Date).ToUniversalTime())).TotalSeconds)

# Create JWT header
$JWTHeader = @{
    alg = "RS256"
    typ = "JWT"
    x5t = ($certHash -replace '\+','-' -replace '/','_' -replace '=')
    }

#create request payload - notice that we do not include nbf (but we could, if needed)
$JWTPayLoad = @{
    aud = "$loginEndpoint/$tenantId/oauth2/token"
    exp = $JWTExpiration
    iss = $clientId
    jti = [guid]::NewGuid()
    #nbf = $NotBefore
    sub = $clientId
    }

# Convert header and payload to base64 and create JWT assertion
$EncodedHeader = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(($JWTHeader | ConvertTo-Json)))
$EncodedPayload = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(($JWTPayload | ConvertTo-Json)))
$JWT = $EncodedHeader + "." + $EncodedPayload

#now sign the assertion
$dataToSign = [byte[]] [System.Text.Encoding]::UTF8.GetBytes($JWT)
#original RSA CSP from cert is not guaranteed to support SHA256
#so we need to create new CSP with proper params and with private key from cert for signing
$algo = (new-object System.IdentityModel.Tokens.X509AsymmetricSecurityKey($X509cert)).GetAsymmetricAlgorithm("http://www.w3.org/2001/04/xmldsig-more#rsa-sha256", $true) -as [System.Security.Cryptography.RSA]

if($algo -is [System.Security.Cryptography.RSACryptoServiceProvider]){
    #cert uses CryptoAPI CSP
    if(($algo.CspKeyContainerInfo.ProviderType -ne 1) -and ($algo.CspKeyContainerInfo.ProviderType -ne 12) -or $algo.CspKeyContainerInfo.HardwareDevice){
        #we have SHA256 compatible provider, just use it
        $csp = $algo -as [System.Security.Cryptography.RSACryptoServiceProvider]
        }
    else {
        #we have to create new compatible CSP with key from cert
        $cspParams = new-object System.Security.Cryptography.CspParameters
        $cspParams.ProviderType=24 #MS Enhanced RSA and AES CSP - this supports SHA256; see Computer\HKLM\SOFTWARE\Microsoft\Cryptography\Defaults\Provider Types\Type 024
        $cspParams.KeyContainerName=$algo.CspKeyContainerInfo.KeyContainerName
        $cspParams.KeyNumber = $algo.CspKeyContainerInfo.KeyNumber
        $cspParams.Flags = 'UseExistingkey'
        if($algo.CspKeyContainerInfo.MachineKeyStore) {$cspParams.Flags = $cspParams.Flags -bor 'UseMachineKeyStore'}

        $csp = new-object System.Security.Cryptography.RSACryptoServiceProvider($cspParams)
        }

    $sha256 = new-object System.Security.Cryptography.SHA256Cng

    # Create a signature of the JWT
    $Signature = [Convert]::ToBase64String($csp.SignData($dataToSign,$sha256))
    }
else{
    #we will use CNG - use the provider
    $csp = $algo -as [System.Security.Cryptography.RsaCng]
    $sha256 = new-object System.Security.Cryptography.SHA256Cng
    #and create a signature
    $hash = $sha256.ComputeHash($dataToSign)
    $Signature = [Convert]::ToBase64String($csp.SignHash($hash,[System.Security.Cryptography.HashAlgorithmName]::SHA256,[System.Security.Cryptography.RSASignaturePadding]::Pkcs1))
    }

# add signature to assertion
$JWT = $JWT + "." + ($Signature -replace '\+','-' -replace '/','_' -replace '=')
return $JWT
    }
function send-graphMailMessage(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [ValidatePattern("@")]
            [string]$fromUpn
        ,[parameter(Mandatory = $true)]
            [ValidatePattern("@")]
            [string[]]$toAddresses
        ,[parameter(Mandatory = $false)]
            [ValidatePattern("@")]
            [string[]]$ccAddresses
        ,[parameter(Mandatory = $false)]
            [ValidatePattern("@")]
            [string[]]$bccAddresses
        ,[parameter(Mandatory = $true)]
            [string]$subject
        ,[parameter(Mandatory = $true,ParameterSetName = "text")]
            [string]$bodyText
        ,[parameter(Mandatory = $true,ParameterSetName = "HTML")]
            [string]$bodyHtml
        ,[parameter(Mandatory = $false)]
            [bool]$saveToSentItems = $true
        ,[parameter(Mandatory = $false)]
            [ValidateSet ("low","normal","high")]
            [string]$priority = "normal"
        )

    [array]$formattedToAddresses = $toAddresses | % {
        @{emailAddress=@{'address'=$_}}
        }
    [array]$formattedFromAddresses = $fromUpn | % {
        @{emailAddress=@{'address'=$_}}
        }
    $message = @{
        toRecipients = $formattedToAddresses
        subject = $subject
        importance=$priority
        #from = $formattedFromAddresses
        #sender = $formattedFromAddresses
        }

    if($ccAddresses){
        [array]$formattedCcAddresses = $ccAddresses | % {
            @{emailAddress=@{'address'=$_}}
            }
        $message.Add("ccRecipients",$formattedCcAddresses)
        }
    if($bccAddresses){
        [array]$formattedBccAddresses = $bccAddresses | % {
            @{emailAddress=@{'address'=$_}}
            }
        $message.Add("bccRecipients",$formattedBccAddresses)
        }
    if($bodyText){$message.Add("body",@{"contentType"="Text";"content"=$bodyText})}
    if($bodyHtml){$message.Add("body",@{"contentType"="HTML";"content"=$bodyHtml})}

    invoke-graphPost -tokenResponse $tokenResponse -graphQuery "/users/$fromUpn/sendMail" -graphBodyHashtable @{"message"=$message;"saveToSentItems"=$saveToSentItems}
    }

#$acsAppCreds is a custom object with 2 relevant properties here (TenantId & ClientId). The above is hardcoded as a proof-of-concept, but should also be stored in a Key Vault.
$acsAppCreds = New-Object -TypeName psobject -Property @{TenantId='271df584-ab64-437f-85b6-80ff9bef6c9f';ClientID='b32a2aee-b0bd-46c8-83c8-88b726b180f7'}

#$acsCert is the certificate (with private key) retrieved from the local certificate store
$acsCert = Get-ChildItem Cert:\CurrentUser\My | Where-Object {$_.Subject -match "acs-mailer@anthesisgroup.com"}

#$acsToken is the authtoken required to interact with the Graph API:
$acsToken = get-graphTokenResponse -aadAppCreds $acsAppCreds -grant_type certificate -cert $cert
#$acsToken looks like this:
#token_type         : Bearer
#expires_in         : 3599
#ext_expires_in     : 3599
#access_token       : eyJ0eXAiOiJKV1QiLCJub25jZSI6ImdBZlQycWNaY05fazF1cXdaOF9nM0FZMWlNRlVEelBaekRraEYyY3NRc1UiLCJhbGciOiJSUzI1NiIsIng1dCI6IjJaUXBKM1VwYmpBWVhZR2FYRUpsOGxWMFR 
#                     PSSIsImtpZCI6IjJaUXBKM1VwYmpBWVhZR2FYRUpsOGxWMFRPSSJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC8yNzFkZ 
#                     jU4NC1hYjY0LTQzN2YtODViNi04MGZmOWJlZjZjOWYvIiwiaWF0IjoxNjYyNjM4OTcwLCJuYmYiOjE2NjI2Mzg5NzAsImV4cCI6MTY2MjY0Mjg3MCwiYWlvIjoiRTJaZ1lIaG9aeit2b3BYdFZ2VmZJ 
#                     YStWb1RLVkFBPT0iLCJhcHBfZGlzcGxheW5hbWUiOiJBQ1MgTWFpbGVyIiwiYXBwaWQiOiJiMzJhMmFlZS1iMGJkLTQ2YzgtODNjOC04OGI3MjZiMTgwZjciLCJhcHBpZGFjciI6IjIiLCJpZHAiOiJ 
#                     odHRwczovL3N0cy53aW5kb3dzLm5ldC8yNzFkZjU4NC1hYjY0LTQzN2YtODViNi04MGZmOWJlZjZjOWYvIiwiaWR0eXAiOiJhcHAiLCJvaWQiOiIxODQxZmM5YS04ZTY3LTRkMDMtYWIwYi0xNjg3MT 
#                     ViNDk2NDEiLCJyaCI6IjAuQVc0QWhQVWRKMlNyZjBPRnRvRF9tLTlzbndNQUFBQUFBQUFBd0FBQUFBQUFBQUJ1QUFBLiIsInJvbGVzIjpbIk1haWwuUmVhZFdyaXRlIiwiTWFpbC5TZW5kIl0sInN1Y 
#                     iI6IjE4NDFmYzlhLThlNjctNGQwMy1hYjBiLTE2ODcxNWI0OTY0MSIsInRlbmFudF9yZWdpb25fc2NvcGUiOiJOQSIsInRpZCI6IjI3MWRmNTg0LWFiNjQtNDM3Zi04NWI2LTgwZmY5YmVmNmM5ZiIs 
#                     InV0aSI6IndKSHFKSXpucUVxMl9KV0VmM1JMQVEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbIjA5OTdhMWQwLTBkMWQtNGFjYi1iNDA4LWQ1Y2E3MzEyMWU5MCJdLCJ4bXNfdGNkdCI6MTM3NzA1MzI4MH0 
#                     .FyoATNd-nDJwxDRPII3fPJG0_pH2zChqPcmtWmn3TnmSsKCHZstH_d9wwBhfWkAWojgj4pRSM8WpbI42E30_t099LZX5HYxYZ9KniVP1w9RNTon4EeDIbz5GGeu0B04NTmslKk7j0Stw-1ypSfTT8a 
#                     0wKYFRgDJzhhgnucNz5eYiA2chjZEPxzPeWK0nHS7116vM0XpHyORtYK37qri0cYNjmwQ3tr2Hwci1F-r87ucNS0gbYerZ4euaj7qoEyB-OKVnAO4I0twPAL1gBriRhoppYteLf9J80I50Iwiob7U6z 
#                     6jEkb-HeX7kVqtSs2W031pjpA90Ua3MxraAPSm3VA
#OriginalExpiryTime : 08/09/2022 14:14:28

#Then we can test sending mail from any of the mailboxes that ACS-Mailer is permitted to send from (https://portal.azure.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/Members/groupId/a307c61f-c5c0-42c1-a68e-76ea35f6cdf3)
send-graphMailMessage -tokenResponse $acsToken -fromUpn 'acs.demov1@anthesisgroup.com' -toAddresses 'michael.malate@anthesisgroup.com' -subject "Hello Mike" -bodyText "Hurrah!"