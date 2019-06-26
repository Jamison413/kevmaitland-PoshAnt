
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
    Write-Host -f Yellow Connecting to MSOL services
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
    Write-Host -f DarkYellow "Executing Connect-MsolService"
    Connect-MsolService -Credential $credential
    }
function connect-toAAD($credential){
    Write-Host -f Yellow Connecting to AAD services
    if ($(Get-Module -ListAvailable AzureAD) -ne $null){
        Write-Host -f DarkYellow "Importing AzureAD (_not_ Preview)"
        Import-Module AzureAD
        Write-Host -f DarkYellow "Executing Connect-AzureAD"
        }
    if ($(Get-Module -ListAvailable AzureADPreview) -ne $null){
        Write-Host -f DarkYellow "Importing AzureADPreview"
        Import-Module AzureADPreview
        Write-Host -f DarkYellow "Executing Connect-AzureAD"
        }
    try{Connect-AzureAD -Credential $credential -ErrorAction Stop -WarningAction Stop -InformationAction Stop}
    catch{
        Write-Host -ForegroundColor DarkRed "MFA might be required"
        Connect-AzureAD
        }
    }
function connect-ToExo($credential){
    Write-Host -f Yellow Connecting to EXO services
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
    Import-Module Microsoft.Exchange.Management.ExoPowershellModule
    Write-Host -f DarkYellow "Initiating New-PSSession"
    try {
        $ExchangeSession = New-ExoPSSession -UserPrincipalName $Credential.Username -ConnectionUri 'https://outlook.office365.com/PowerShell-LiveId' -AzureADAuthorizationEndpointUri 'https://login.windows.net/common' -Credential $Credential -ErrorAction Stop -WarningAction Stop -InformationAction Stop
        }
    catch{
        Write-Host -ForegroundColor DarkRed "MFA might be required"
        $ExchangeSession = New-ExoPSSession -UserPrincipalName $Credential.Username -ConnectionUri 'https://outlook.office365.com/PowerShell-LiveId' -AzureADAuthorizationEndpointUri 'https://login.windows.net/common'
        }
    Write-Host -f DarkYellow "Importing New-PSSession"
    Import-Module (Import-PSSession $ExchangeSession -AllowClobber) -Global
    }
function connect-ToSecCom($credential){
    Write-Host -f Yellow Connecting to Security `& Compliance services
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
    Import-Module Microsoft.Exchange.Management.ExoPowershellModule
    Write-Host -f DarkYellow "Initiating New-PSSession"
    try {
        $SecComSession = New-ExoPSSession -UserPrincipalName $Credential.Username -ConnectionUri 'https://ps.compliance.protection.outlook.com/powershell-liveid/' -AzureADAuthorizationEndpointUri 'https://login.windows.net/common' -Credential $Credential -ErrorAction Stop -WarningAction Stop -InformationAction Stop
        }
    catch{
        Write-Host -ForegroundColor DarkRed "MFA might be required"
        $SecComSession = New-ExoPSSession -UserPrincipalName $Credential.Username -ConnectionUri 'https://ps.compliance.protection.outlook.com/powershell-liveid/' -AzureADAuthorizationEndpointUri 'https://login.windows.net/common'
        }
    Write-Host -f DarkYellow "Importing New-PSSession"
    Import-Module (Import-PSSession $SecComSession -AllowClobber) -Global
    }

function connect-ToSpo($credential){
    Write-Host -f Yellow Connecting to SPO services
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
    if ($credential -eq $null){
        Write-Host -f DarkYellow "Credentials not provided, requesting now."
        $credential = set-MsolCredentials
        }
    
    Write-Host -f DarkYellow "Importing Microsoft.Online.Sharepoint.PowerShell"
    Import-Module Microsoft.Online.Sharepoint.PowerShell
    Write-Host -f DarkYellow "Executing Connect-SPOService"
    Write-Host -f DarkYellow "Credential: $($credential.UserName) $($credential.Password)"
    Connect-SPOService -url 'https://anthesisllc-admin.sharepoint.com' -Credential $credential
    Write-Host -f DarkYellow "Executing Connect-PnPOnline"
    Connect-PnPOnline –Url https://anthesisllc.sharepoint.com -Credentials  $credential
    }
function connect-to365(){
    Write-Host -f Yellow "Importing Modules"
    Write-Host -f DarkYellow "_PS_Library_GeneralFunctionality"
    Import-Module _PS_Library_GeneralFunctionality
    Write-Host -f DarkYellow "_PS_Library_Groups"
    Import-Module _PS_Library_Groups
    Write-Host -f DarkYellow "_CSOM_Library-SPO"
    Import-Module _CSOM_Library-SPO
    Write-Host -f DarkYellow "_REST_Library-SPO"
    Import-Module _REST_Library-SPO
    Write-Host -f DarkYellow "_REST_Library-Kimble"
    Import-Module _REST_Library-Kimble

    Write-Host -f Yellow Connecting to 365 services
    Write-Host -f DarkYellow "Executing set-MsolCredentials"
    $msolCredentials = set-MsolCredentials 
    Write-Host -f DarkYellow "Executing connect-ToMsol"
    connect-ToMsol $msolCredentials
    Write-Host -f DarkYellow "Executing connect-toAAD"
    connect-toAAD $msolCredentials
    Write-Host -f DarkYellow "Executing connect-ToExo"
    connect-ToExo $msolCredentials
    #Write-Host -f DarkYellow "connect-ToSpo"
    #connect-ToSpo $msolCredentials
    #$csomCredentials = new-csomCredentials -username $msolCredentials.UserName -password $msolCredentials.Password
    #$restCredentials = new-spoCred -username $msolCredentials.UserName -securePassword $msolCredentials.Password
    }
