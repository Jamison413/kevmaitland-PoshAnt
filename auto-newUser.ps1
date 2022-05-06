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

<#
Local Auth
Import-Module @("_PS_Library_Graph", "Az.Accounts", "Az.Automation", "Az.ManagedServiceIdentity", "Az.Compute")
$tokenMe = get-graphTokenResponse `
		-aadAppCreds $(get-graphAppClientCredentials -appName MyBot) `
		-grant_type certificate
#>

# Sign in to your Azure subscription

# Ensures you do not inherit an AzContext in your runbook
Disable-AzContextAutosave -Scope Process | Out-Null
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

#Get AppCreds stored as Automation variables - we _could_ move these to Key Vault
$myAppCreds = [PSCustomObject]@{
    ClientID = $(Get-AzAutomationVariable -ResourceGroupName "Automation" -AutomationAccountName "azant-userbot" -Name "auto-userbot-clientid").Value
    TenantId = $(Get-AzAutomationVariable -ResourceGroupName "Automation" -AutomationAccountName "azant-userbot" -Name "tenantId").Value
}

#Get certificate (inc. private key for signing) from Key Vault and convert to X509Certificate2
Write-Output "Getting Secret from Key Vault"
$secret = Get-AzKeyVaultSecret -VaultName "azant-auto-userbot" -Name "auto-userbot" -AsPlainText
$secretBytes = [System.Convert]::FromBase64String($secret)
$x509Cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::New($secretBytes)

#Get access_token for Graph using the X509Certificate2
$tokenMe = get-graphTokenResponse -aadAppCreds $myAppCreds -grant_type certificate -cert $x509Cert
#Uncomment to connect to Exchange
    $exchangeOrg = $(Get-AzAutomationVariable -ResourceGroupName "Automation" -AutomationAccountName "azant-userbot" -Name "exchangeOrg").Value
    Connect-ExchangeOnline -AppId $myAppCreds.ClientID -Certificate $x509Cert -Organization $exchangeOrg
#endregion




#region Do all the things
$hrSiteId = $(Get-AzAutomationVariable -ResourceGroupName "Automation" -AutomationAccountName "azant-userbot" -Name "oldHrSiteId").Value
#$hrSiteId = $(get-graphSite -tokenResponse $tokenMe -serverRelativeUrl "https://anthesisllc.sharepoint.com/teams/hr/").id
$hrListId = $(Get-AzAutomationVariable -ResourceGroupName "Automation" -AutomationAccountName "azant-userbot" -Name "oldHrNewUserListId").Value
#$hrListId = $(get-graphList -tokenResponse $tokenMe -graphSiteId $site.id -listName "New User Requests").id
$unfulfilledRequests = get-graphListItems -tokenResponse $tokenMe -graphSiteId $hrSiteId -listId $hrListId -expandAllFields -filterQuery "fields/GraphUserGUID eq null and fields/Start_x0020_Date gt '2022-01-01T00:00:00Z'"
Write-Output "[$([int]$unfulfilledRequests.Count)] New User Requests without GraphUserGUIDs retrieved"
Write-Output $unfulfilledRequests | Sort-Object {$_.fields.Start_x0020_Date} | Select-Object {$_.fields.Title}, {$_.fields.Start_x0020_Date}, {$_.fields.Finance_x0020_Cost_x0020_Attribu.Label}, {$_.fields.Current_x0020_Status} | Format-Table


Write-Output "Retrieving Mailbox..."
get-exomailbox "kevin.maitland@anthesisgroup.com"


#endregion




#region Tidyup
Disconnect-ExchangeOnline -Confirm:$FALSE -ErrorAction SilentlyContinue
#endregion