$ClientID           = "{client_id}"
$loginURL           = "https://login.microsoftonline.com"
$tenantdomain       = "{tenant_id}"
$CertPassWord       = "{password_for_cert}"
$certPath           = "C:\temp\Certs\testCert_01.pfx"
 
[string[]] $Scopes  = "https://graph.microsoft.com/.default"
 
Function Load-MSAL {
    if ($PSVersionTable.PSVersion.Major -gt 5)
    { 
        $core = $true
        $foldername =  "netcoreapp2.1"
    }
    else
    { 
        $core = $false
        $foldername = "net45"
    }
 
    # Download MSAL.Net module to a local folder if it does not exist there
    if ( ! (Get-ChildItem $HOME/MSAL/lib/Microsoft.Identity.Client.* -erroraction ignore) ) {
        install-package -Source nuget.org -ProviderName nuget -SkipDependencies -Name Microsoft.Identity.Client -Destination $HOME/MSAL/lib -force -forcebootstrap | out-null
    }
   
    # Load the MSAL assembly -- needed once per PowerShell session
    [System.Reflection.Assembly]::LoadFrom((Get-ChildItem $HOME/MSAL/lib/Microsoft.Identity.Client.*/lib/$foldername/Microsoft.Identity.Client.dll).fullname) | out-null
  }
  
Function Get-GraphAccessTokenFromMSAL {
 
    Load-MSAL
 
    $global:app = $null
 
    $x509cert = [System.Security.Cryptography.X509Certificates.X509Certificate2] (GetX509Certificate_FromPfx -CertificatePath $certPath -CertificatePassword $CertPassWord)
    write-host "Cert = {$x509cert}"
 
    $ClientApplicationBuilder = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($ClientID)
        [void]$ClientApplicationBuilder.WithAuthority($("$loginURL/$tenantdomain"))
        [void]$ClientApplicationBuilder.WithCertificate($x509cert)
    $global:app = $ClientApplicationBuilder.Build()
 
    [Microsoft.Identity.Client.AuthenticationResult] $authResult  = $null
    $AquireTokenParameters = $global:app.AcquireTokenForClient($Scopes)
    try {
        $authResult = $AquireTokenParameters.ExecuteAsync().GetAwaiter().GetResult()
    }
    catch {
        $ErrorMessage = $_.Exception.Message
        Write-Host $ErrorMessage
    }
     
    return $authResult
}
 
function GetX509Certificate_FromPfx($CertificatePath, $CertificatePassword){
    #write-host "Path: '$CertificatePath'"
    
    if(![System.IO.Path]::IsPathRooted($CertificatePath))
    {
        $LocalPath = Get-Location
        $CertificatePath = "$LocalPath\$CertificatePath"
    }
 
    #Write-Host "Looking for '$CertificatePath'"
 
    $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath, $CertificatePassword)
     
    Return $certificate
}
 
$myvar = Get-GraphAccessTokenFromMSAL
Write-Host "Access Token: " $myvar.AccessToken



Import-Module Microsoft.Identity.Client
$loginURL           = "https://login.microsoftonline.com"
[string[]] $Scopes  = "https://graph.microsoft.com/.default"

$reportBotDetails = $(get-graphAppClientCredentials -appName UserBot)
#$teamBotDetails = $(get-graphAppClientCredentials -appName TeamsBot)
$myCert = Get-ChildItem Cert:\CurrentUser\My | ? {$_.Subject -match $env:USERNAME}
$ClientApplicationBuilder = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($reportBotDetails.ClientID)
[void]$ClientApplicationBuilder.WithAuthority($("$loginURL/$($reportBotDetails.TenantId)"))
[void]$ClientApplicationBuilder.WithCertificate($myCert)
$confidentialClientApplication = $ClientApplicationBuilder.Build()

    #[Microsoft.Identity.Client.AuthenticationResult] $authResult  = $null
    $AquireTokenParameters = $confidentialClientApplication.AcquireTokenForClient($Scopes)
    try {
        [Microsoft.Identity.Client.AuthenticationResult]$authResult = $AquireTokenParameters.ExecuteAsync().GetAwaiter().GetResult()
        $authResult | Add-Member -MemberType NoteProperty -Name access_token -Value $authResult.AccessToken
    }
    catch {
        get-errorSummary $_
    }
     

Get-Package -Name Microsoft.Identity.Client -ProviderName nuget

