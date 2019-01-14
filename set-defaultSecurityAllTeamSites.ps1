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

$teamSites = Get-PnPTenantSite -Template GROUP#0

$excludeThese = @("https://anthesisllc.sharepoint.com/teams/Energy_%26_Carbon_Consulting_Analysts_%26_Software_ECCAST_Community_","https://anthesisllc.sharepoint.com/sites/AccountsPayable","https://anthesisllc.sharepoint.com/sites/anthesisnorthamerica","https://anthesisllc.sharepoint.com/sites/apparel","https://anthesisllc.sharepoint.com/sites/bdcontacts42","https://anthesisllc.sharepoint.com/teams/BusinessDevelopmentTeam-GBR-","https://anthesisllc.sharepoint.com/teams/PreSalesTeam","https://anthesisllc.sharepoint.com/teams/teamstestingteam","https://anthesisllc.sharepoint.com/sites/sparke","https://anthesisllc.sharepoint.com/sites/supplychainsym")

$teamSites | ? {$excludeThese -notcontains $_.Url -and $_.Url -notmatch "Confidential"} | % {
    $thisTeamSite = $_
    #Write-Host $thisTeamSite.Url
    #Set-PnPTenantSite -Url $thisTeamSite.Url -Owners (Get-PnPConnection).PSCredential.UserName # This will be automatically removed in the set-standardTeamSitePermissions script
    #set-standardTeamSitePermissions -teamSiteAbsoluteUrl $thisTeamSite.Url -adminCredentials $sharePointCreds -verboseLogging $verboseLogging -fullLogPathAndName $fullLogPathAndName -errorLogPathAndName $errorLogPathAndName 
    Connect-PnPOnline -Url $thisTeamSite.Url -Credentials $sharePointCreds
    Remove-PnPSiteCollectionAdmin -Owners (Get-PnPConnection).PSCredential.UserName
    }

$url = "https://anthesisllc.sharepoint.com/teams/Sustainable_Chemistry_Team_All_365"
Connect-PnPOnline -Url $Url -Credentials $msolCredentials
Set-PnPTenantSite -Url $Url -Owners (Get-PnPConnection).PSCredential.UserName # This will be automatically removed in the set-standardTeamSitePermissions script
Remove-PnPSiteCollectionAdmin -Owners (Get-PnPConnection).PSCredential.UserName
