﻿
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
    if ($username -notmatch "@anthesisgroup.com"){$username = "$username@anthesisgroup.com"}
    if ($password -eq $null -or $password -eq ""){$password = Read-Host -Prompt "Password for $username" -AsSecureString}
        else{ConvertTo-SecureString ($password) -AsPlainText -Force}
    $credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $password
    $credential
    }
function bodge-exo(){
    [cmdletbinding()]
    param()
    add-registryValue -registryPath "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WinRM\Client" -registryKey "AllowBasic" -registryValue "1" -registryType DWord -Verbose
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
    try{
        Get-MsolDomain -ErrorAction Stop | Out-Null
        Write-Host -ForegroundColor Yellow "Already connected to MSOL services"
        }
    catch{
        Write-Host -ForegroundColor Yellow "Connecting to MSOL services"
        Import-Module MSOnline
        try{Connect-MsolService -Credential $credential}
        catch{
            Write-Warning "Couldn't connect to MSOL non-interactively, trying interactively."
            Connect-MsolService
            }
        }
    }
function connect-toAzureRm{
    Param (
        [parameter(Mandatory = $false)]
        [pscredential]$aadCreds
        )
    Write-Host -f Yellow Connecting to AzureRM services
    Import-Module AzureRM.Profile
    Try {
        Login-AzureRmAccount -Credential $aadCreds -ErrorAction Stop | Out-Null
        } 
    Catch {
        Write-Warning "Couldn't connect to Azure RM non-interactively, trying interactively."
        Login-AzureRmAccount -TenantId $(($aadCreds.UserName.Split("@"))[1]) -ErrorAction Stop | Out-Null
        }
    }
function connect-toAAD($credential){
    try{
        Get-AzureADTenantDetail -ErrorAction Stop | Out-Null
        Write-Host -f Yellow "Already connected to AAD services"
        }
    catch{
        Write-Host -f Yellow "Connecting to AAD services"
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
    }
function connect-ToExo($credential){
    <#
    .Synopsis
        Provides a standardised (and simplifed) way to connect to MSOL services
    .DESCRIPTION
        Provides a standardised (and simplifed) way to connect to MSOL services.
        If no credentials are supplied, built-in Microsoft login window opens to sign in (included in Exchange v2 module).
    .EXAMPLE
        connect-ToExo
    .EXAMPLE
        connect-ToExo -credential $creds
    #>
    switch ($(Get-PSSession | ? {$_.ComputerName -eq "outlook.office365.com" -and $_.Availability -eq "Available" -and $_.State -eq "Opened"}).Count){
        0 {
            Write-Host -f Yellow Connecting to EXO services
            #(v1)if ($credential -eq $null){$credential = set-MsolCredentials}
            #Import-Module Microsoft.Exchange.Management.ExoPowershellModule
            Write-Host -f DarkYellow "Initiating New-PSSession"
            
            If($credential -eq $null){
            try {
                #bodge-exo 
                #$ExchangeSession = New-ExoPSSession -UserPrincipalName $Credential.Username -ConnectionUri 'https://outlook.office365.com/PowerShell-LiveId?SerializationData=Full' -AzureADAuthorizationEndpointUri 'https://login.windows.net/common' -Credential $Credential -ErrorAction Stop -WarningAction Stop -InformationAction Stop
                #(v1)$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Authentication Basic -AllowRedirection -Credential $credential  -ErrorAction Stop
                $ExchangeSession = Connect-ExchangeOnline
                }
            catch{
                Write-Host -f DarkRed "Something went wrong connecting to EXO :/"
                #(v1)Write-Host -ForegroundColor DarkRed "MFA might be required"
                #$ExchangeSession = New-ExoPSSession -UserPrincipalName $Credential.Username -ConnectionUri 'https://outlook.office365.com/PowerShell-LiveId?SerializationData=Full' -AzureADAuthorizationEndpointUri 'https://login.windows.net/common'
                #(v1)$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Authentication Basic -AllowRedirection -ErrorAction Stop
                }
            }
            Else{
            Try{
               $ExchangeSession = Connect-ExchangeOnline -Credential $credential
               }
            Catch{
                Write-Host -f DarkRed "Something went wrong connecting to EXO with credentials :/"
               }
            }    
            Write-Host -f DarkYellow "Importing New-PSSession"
            #Import-Module (Import-PSSession $ExchangeSession -DisableNameChecking -AllowClobber -CommandName *) -Global        
            #(v1)Import-Module (Import-PSSession $ExchangeSession -DisableNameChecking -AllowClobber) -Global
            
            }
        1 {
            if((Get-Module | ? {$_.ExportedCommands.Keys -contains "Get-Mailbox"}).Count -gt 0)
                {
                Write-Host -f Yellow "Already connected to EXO services"
                }
            else{
                Import-Module (Import-PSSession $(Get-PSSession | ? {$_.ComputerName -eq "outlook.office365.com" -and $_.Availability -eq "Available" -and $_.State -eq "Opened"}) -AllowClobber -CommandName * -DisableNameChecking) -Global
                }
            }
        default {
            Write-Host -f DarkRed "Something went wrong connecting to EXO :/"
            }
        }
    }
function connect-ToS4b($credential){
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
    switch ($(Get-PSSession | ? {$_.ComputerName -eq "outlook.office365.com" -and $_.Availability -eq "Available" -and $_.State -eq "Opened"}).Count){
        0 {
            Write-Host -f Yellow Connecting to S4B services
            if ($credential -eq $null){$credential = set-MsolCredentials}
            Import-Module SkypeOnlineConnector
            Write-Host -f DarkYellow "Initiating New-PSSession"
            try {
                #bodge-exo 
                $s4bSession = New-CsOnlineSession -Credential $credential
                }
            catch{
                Write-Host -ForegroundColor DarkRed "MFA might be required"
                $s4bSession = New-CsOnlineSession -UserName $Credential.Username
                }
            Write-Host -f DarkYellow "Importing New-PSSession"
            Import-Module (Import-PSSession $s4bSession -AllowClobber) -Global            
            }
        1 {
            if((Get-Module | ? {$_.ExportedCommands.Keys -contains "Get-CsAudioConferencingProvider"}).Count -gt 0)
                {
                Write-Host -f Yellow "Already connected to S4B services"
                }
            else{
                Import-Module (Import-PSSession $(Get-PSSession | ? {$_.ComputerName -match "online.lync.com" -and $_.Availability -eq "Available" -and $_.State -eq "Opened"}) -AllowClobber) -Global
                }
            }
        default {
            Write-Host -f DarkRed "Something went wrong connecting to S4B :/"
            }
        }
    }
function connect-ToScc($credential){
    <#
    .Synopsis
        Provides a standardised (and simplifed) way to connect to MSOL services
    .DESCRIPTION
        Provides a standardised (and simplifed) way to connect to MSOL services.
        If no credentials are supplied, set-MsolCredentials is called.
    .EXAMPLE
        connect-ToScc
    .EXAMPLE
        connect-ToScc -credential $creds
    #>
    Write-Host -f Yellow Connecting to Scc services
    if ($credential -eq $null){$credential = set-MsolCredentials}
    Import-Module Microsoft.Exchange.Management.ExoPowershellModule
    Write-Host -f DarkYellow "Initiating New-PSSession"
    try {
        #bodge-Scc 
        $ExchangeSession = New-ExoPSSession -UserPrincipalName $Credential.Username -ConnectionUri 'https://ps.compliance.protection.outlook.com/powershell-liveid' -AzureADAuthorizationEndpointUri 'https://login.windows.net/common' -Credential $Credential -ErrorAction Stop -WarningAction Stop -InformationAction Stop
        }
    catch{
        Write-Host -ForegroundColor DarkRed "MFA might be required"
        $ExchangeSession = New-ExoPSSession -UserPrincipalName $Credential.Username -ConnectionUri 'https://ps.compliance.protection.outlook.com/powershell-liveid' -AzureADAuthorizationEndpointUri 'https://login.windows.net/common'
        }
    Write-Host -f DarkYellow "Importing New-PSSession"
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
    if ($credential -eq $null){
        Write-Host -f DarkYellow "Credentials not provided, requesting now."
        $credential = set-MsolCredentials
        }
    
    #Write-Host -f DarkYellow "Importing Microsoft.Online.Sharepoint.PowerShell"
    #Import-Module Microsoft.Online.Sharepoint.PowerShell
    Write-Verbose "Executing Connect-SPOService"
    
    try{
        Get-SPOTenant -ErrorAction Stop | Out-Null
        Write-Host -f Yellow "Already connected to SPO services"
        }
    catch{
        Write-Host -f Yellow "Connecting to SPO [https://anthesisllc-admin.sharepoint.com]"
        try{Connect-SPOService -url 'https://anthesisllc-admin.sharepoint.com' -Credential $credential -ErrorAction Stop -WarningAction Stop -InformationAction Stop}
        catch{
            Write-Host -ForegroundColor DarkRed "MFA might be required"
            Connect-SPOService -url 'https://anthesisllc-admin.sharepoint.com'
            }
        }

    try{
        Get-PnPConnection -ErrorAction Stop | Out-Null
        Write-Host -f Yellow "Already connected to PNP services"
        }
    catch{
        Write-Host -f Yellow "Connecting to PNP [https://anthesisllc.sharepoint.com]"
        try{Connect-PnPOnline –Url https://anthesisllc.sharepoint.com -Credentials  $credential -ErrorAction Stop -WarningAction Stop -InformationAction Stop}
        catch{
            Write-Host -ForegroundColor DarkRed "MFA might be required"
            Connect-PnPOnline –Url https://anthesisllc.sharepoint.com
            }
        }

    }
function connect-toTeams(){
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$false)]
            [pscredential]$credential
        )
    if([string]::IsNullOrEmpty($credential)){
        Write-Verbose "[$credential] is $null"
        $credential = set-MsolCredentials
        }
    try{
        Get-Team -ErrorAction Stop | Out-Null
        Write-Host -f Yellow "Already connected to Teams services"
        }
    catch{
        Write-Host -f Yellow "Connecting to Teams services"
        try{Connect-MicrosoftTeams -Credential $credential -ErrorAction Stop -WarningAction Stop -InformationAction Stop | out-null}
        catch{
            Write-Host -ForegroundColor DarkRed "MFA might be required"
            Connect-MicrosoftTeams
            }
        }
    }
function connect-to365(){
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$false)]
            [pscredential]$credential
        )
    Write-Verbose "Connecting to 365 services"
    if(!$credential){
        Write-Verbose "Executing set-MsolCredentials"
        $credential = set-MsolCredentials
        }
    Write-Verbose "Executing connect-ToMsol"
    connect-ToMsol $credential
    Write-Verbose "Executing connect-toAAD"
    connect-toAAD $credential
    Write-Verbose "Executing connect-ToExo"
    connect-ToExo $credential
    Write-Verbose "Executing connect-ToTeams"
    connect-toTeams -credential $credential
    Write-Verbose "Executing connect-ToSpo"
    connect-ToSpo -credential $credential
    #$csomCredentials = new-csomCredentials -username $msolCredentials.UserName -password $msolCredentials.Password
    #$restCredentials = new-spoCred -username $msolCredentials.UserName -securePassword $msolCredentials.Password
    $credential
    }