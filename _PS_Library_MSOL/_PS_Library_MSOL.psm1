
function set-MsolCredentials($username, $password){
    <#
    .Synopsis
        Gets and sets Global Admin credentials for connecting to MSOL services
    .DESCRIPTION
        Allows the user to capture Global Admin credentials for authenticating with different 
        MSOL services (Azure AD, Exchange Online, SharePoint Online, etc.). If no username is 
        supplied, the current user context is assumed.
    .EXAMPLE
       Set-MsolCredentials
    .EXAMPLE
       Set-MsolCredentials -username kevin.maitland@anthesisgroup.com -password MyPasswordAsPlainText
    #>
    if ($username -eq $null -or $username -eq ""){$username = Read-Host -Prompt "Enter Office 365 Global Administrator username (blank for $($env:USERNAME)@anthesisgroup.com)"}
    if ($username -eq $null -or $username -eq ""){$username = "$($env:USERNAME)@anthesisgroup.com"}
    if ($password -eq $null -or $password -eq ""){$password = Read-Host -Prompt "Password for $username" -AsSecureString}
        else{ConvertTo-SecureString ($password) -AsPlainText -Force}
    $credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $password
    $credential
    }
function connect-ToMsol($credential){
    <#
    .Synopsis
        Provides a standardised (and simplifed) way to connect to MSOL services
    .DESCRIPTION
        Provides a standardised (and simplifed) way to connect to MSOL services.
        If no credentials are supplied, set-MsolCredentials is called.
    .EXAMPLE
       connect-ToMsol
    .EXAMPLE
       connect-ToMsol -credential $creds
    #>
    if ($credential -eq $null){$credential = set-MsolCredentials}
    Import-Module MSOnline
    Connect-MsolService -Credential $credential
    }
function connect-toAAD($credential){
    if ($(Get-Module -ListAvailable AzureAD) -ne $null){
        Import-Module AzureAD
        Connect-AzureAD -Credential $credential
        }
    if ($(Get-Module -ListAvailable AzureADPreview) -ne $null){
        Import-Module AzureADPreview
        Connect-AzureAD -Credential $credential
        }
    }
function connect-ToExo($credential){
    <#
    .Synopsis
        Provides a standardised (and simplifed) way to connect to MSOL services
    .DESCRIPTION
        Provides a standardised (and simplifed) way to connect to MSOL services.
        If no credentials are supplied, set-MsolCredentials is called.
    .EXAMPLE
       connect-ToExo
    .EXAMPLE
       connect-ToExo -credential $creds
    #>
    if ($credential -eq $null){$credential = set-MsolCredentials}
    $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
    Import-Module (Import-PSSession $ExchangeSession -AllowClobber) -Global
    }

function connect-ToSpo($credential){
    <#
    .Synopsis
        Provides a standardised (and simplifed) way to connect to MSOL services
    .DESCRIPTION
        Provides a standardised (and simplifed) way to connect to MSOL services.
        If no credentials are supplied, set-MsolCredentials is called.
    .EXAMPLE
       connect-ToSpo
    .EXAMPLE
       connect-ToSpo -credential $creds
    #>
    if ($credential -eq $null){$credential = set-MsolCredentials}
    Import-Module Microsoft.Online.Sharepoint.PowerShell
    Connect-SPOService -url 'https://anthesisllc-admin.sharepoint.com' -Credential $credential
    }
function connect-to365(){
    $msolCredentials = set-MsolCredentials 
    connect-ToMsol $msolCredentials
    connect-toAAD $msolCredentials
    connect-ToExo $msolCredentials
    connect-ToSpo $msolCredentials
    $csomCredentials = new-csomCredentials -username $msolCredentials.UserName -password $msolCredentials.Password
    $restCredentials = new-spoCred -username $msolCredentials.UserName -securePassword $msolCredentials.Password
    }
