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
    $thisTeamSite = $_
    #Write-Host $thisTeamSite.Url
    if([string]::IsNullOrWhiteSpace($thisTeamSite.SharePointSiteUrl)){
        Write-Verbose "Site [$($thisTeamSite.DisplayName)] is not provisioned yet. having a pop at it, but don't hold your breath."
        #$web = Invoke-WebRequest -Uri "https://outlook.office365.com/owa/$($thisTeamSite.PrimarySmtpAddress)/groupsubscription.ashx?realm=anthesisgroup.com&source=WelcomeEmail&action=files" -Credential $sharePointCreds -SessionVariable thisSession
        #$web2 = Invoke-WebRequest -Uri "https://anthesisllc.sharepoint.com/_layouts/15/groupstatus.aspx?id=$($thisTeamSite.ExternalDirectoryObjectId)&target=documents" -Credential $sharePointCreds -SessionVariable thisSession -Method Get
        #$web3 = Invoke-WebRequest -Uri "https://anthesisllc.sharepoint.com/_layouts/15/groupstatus.aspx?id=$($thisTeamSite.ExternalDirectoryObjectId)&target=documents" -Credential $sharePointCreds -SessionVariable $thisSession
        }
    else{
        Write-Verbose "Setting security defaults for [$($thisTeamSite.DisplayName)]"
        Set-PnPTenantSite -Url $thisTeamSite.SharePointSiteUrl -Owners $((Get-PnPConnection).PSCredential.UserName)
        Connect-PnPOnline -Url $thisTeamSite.SharePointSiteUrl -Credentials $sharePointCreds
        set-standardTeamSitePermissions -teamSiteAbsoluteUrl $thisTeamSite.SharePointSiteUrl -fullArrayOfUnifiedGroups $allUnifiedGroups -Verbose
        }
    }


