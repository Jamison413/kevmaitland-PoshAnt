#set-defaultSecurityAllTeamSites
$logFileLocation = "C:\ScriptLogs\"
if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)){
    $fullLogPathAndName = $logFileLocation+"set-defaultSecurityAllTeamSites_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = $logFileLocation+"set-defaultSecurityAllTeamSites_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
else{
    $fullLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
if($PSCommandPath){
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_Transcript_$(Get-Date -Format "yyMMdd").log"
    Start-Transcript $transcriptLogName -Append
    }

Import-Module SharePointPnPPowerShellOnline
Import-Module _PNP_Library_SPO

$teamBotDetails = Import-Csv "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\teambotdetails.txt"
$resource = "https://graph.microsoft.com"
$tenantId = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.TenantId)
$clientId = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.ClientID)
$redirect = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.Redirect)
$secret   = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.Secret)

$ReqTokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    client_Id     = $clientID
    Client_Secret = $secret
    } 
$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody


$groupAdmin = "groupbot@anthesisgroup.com"
#convertTo-localisedSecureString ""
$groupAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\GroupBot.txt) 
$exoCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $groupAdmin, $groupAdminPass

$sharePointAdmin = "kimblebot@anthesisgroup.com"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\KimbleBot.txt) 
$sharePointCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
#$sharePointCreds = set-MsolCredentials

connect-ToExo -credential $exoCreds
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com" -Credentials $sharePointCreds

$allUnifiedGroups = Get-UnifiedGroup
$excludeThese = @("https://anthesisllc.sharepoint.com/teams/Energy_%26_Carbon_Consulting_Analysts_%26_Software_ECCAST_Community_","https://anthesisllc.sharepoint.com/sites/AccountsPayable","https://anthesisllc.sharepoint.com/sites/anthesisnorthamerica","https://anthesisllc.sharepoint.com/sites/apparel","https://anthesisllc.sharepoint.com/sites/bdcontacts42","https://anthesisllc.sharepoint.com/teams/BusinessDevelopmentTeam-GBR-","https://anthesisllc.sharepoint.com/teams/PreSalesTeam","https://anthesisllc.sharepoint.com/teams/teamstestingteam","https://anthesisllc.sharepoint.com/sites/sparke","https://anthesisllc.sharepoint.com/sites/supplychainsym")

$groupsToProcess = $allUnifiedGroups | ? {$excludeThese -notcontains $_.SharePointSiteUrl -and $_.Displayname -notmatch "Confidential"}
$groupsToProcess | % {
    $thisUnifiedGroup = $_
    #Write-Host $thisUnifiedGroup.Url
    if([string]::IsNullOrWhiteSpace($thisUnifiedGroup.SharePointSiteUrl)){
        Write-Verbose "Site [$($thisUnifiedGroup.DisplayName)] is not provisioned yet..."
        }
    else{
        #set-standardTeamSitePermissions -teamSiteAbsoluteUrl $thisUnifiedGroup.SharePointSiteUrl -fullArrayOfUnifiedGroups $allUnifiedGroups -Verbose
        Write-Verbose ""
        set-standardSitePermissions -unifiedGroupObject $thisUnifiedGroup -tokenResponse $tokenResponse -pnpCreds $sharePointCreds -Verbose
        }
    }

