<#
    .DESCRIPTION
        Template for PowerShell Runbook using certificate-based authentication to utilise App Registrations.
        Managed Identity is used to for the Azure runs, which gets a certificate from Key Vault and uses it to authenticate with a dedicated App Registration for this Azure Automation account.
        Developers run the commented-out Auth section that uses a local certificate to authenticate with their own dedicated App Registration for this Azure Automation account.
        Either way, we end up with an access_token for the Graph API with the same permissions/scopes, using certificates and no secrets.

    .NOTES
        AUTHOR: Anthesis IT Team 
        LASTEDIT: 2022-05-05 (Kev Maitland)
#>


#region auth
<# Local Auth - run this section manually if you are writing code locally in VS Code 
    $tokenMe = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName MyBot) -grant_type certificate
#>

# Ensures you do not inherit an AzContext in your runbook
Disable-AzContextAutosave -Scope Process
# Connect to Azure with system-assigned managed identity
$AzureContext = (Connect-AzAccount -Identity).context

# set and store context
try {
    $AzureContext = Set-AzContext -SubscriptionName $AzureContext.Subscription -DefaultProfile $AzureContext -ErrorAction Stop
}
catch {
    Write-Output "If you see [this.Client.SubscriptionId cannot be null] errors, you need to assign a Role to the Azure Automation account"
    get-errorSummary $_
    exit
}

#Get AppCreds stored as Automation variables - we _could_ move these to Key Vault (as we're reliant on Key Vault for the certificate), but this is handy as a code example
$myAppCreds = [PSCustomObject]@{
    ClientID = $(Get-AzAutomationVariable -ResourceGroupName "Automation" -AutomationAccountName "azant-userbot" -Name "auto-userbot-clientid").Value
    TenantId = $(Get-AzAutomationVariable -ResourceGroupName "Automation" -AutomationAccountName "azant-userbot" -Name "tenantId").Value
}
Write-Verbose "`$myAppCreds = $myAppCreds"

#Get certificate (inc. private key for signing) from Key Vault and convert to X509Certificate2
Write-Verbose "Getting Secret from Key Vault"
$secret = Get-AzKeyVaultSecret -VaultName "azant-auto-userbot" -Name "auto-userbot" -AsPlainText
$secretBytes = [System.Convert]::FromBase64String($secret)
$x509Cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::New($secretBytes)
Write-Verbose "`$x509Cert = $x509Cert"

#Get access_token for Graph using the X509Certificate2
$tokenMe = get-graphTokenResponse -aadAppCreds $myAppCreds -grant_type certificate -cert $x509Cert

#Uncomment to connect to Exchange
    #$exchangeOrg = $(Get-AzAutomationVariable -ResourceGroupName "Automation" -AutomationAccountName "azant-userbot" -Name "exchangeOrg").Value
    #Connect-ExchangeOnline -AppId $myAppCreds.ClientID -Certificate $x509Cert -Organization $exchangeOrg


#endregion


#region Do all the things





#endregion


#region Tidyup
Disconnect-ExchangeOnline -Confirm:$FALSE -ErrorAction SilentlyContinue
#endregion