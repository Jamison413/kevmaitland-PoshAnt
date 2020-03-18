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

$teamBotDetails = Import-Csv "$env:USERPROFILE\Desktop\teambotdetails.txt"
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
$groupAdmins = get-groupAdminRoleEmailAddresses

$groupsToProcess = $allUnifiedGroups | ? {$excludeThese -notcontains $_.SharePointSiteUrl -and $_.Displayname -notmatch "Confidential"}
$groupsToProcess | % {
    $thisUnifiedGroup = $_
    #Write-Host $thisUnifiedGroup.Url
    $combinedDg = Get-DistributionGroup -Identity $thisUnifiedGroup.CustomAttribute4
    #Standardise the names
    if($combinedDg){
        if($combinedDg.DisplayName -ne $thisUnifiedGroup.DisplayName){
            Set-UnifiedGroup -Identity $thisUnifiedGroup.ExternalDirectoryObjectId -DisplayName $combinedDg.DisplayName
            $groupOwners = Get-UnifiedGroupLinks -LinkType Owners -Identity $thisUnifiedGroup.ExternalDirectoryObjectId
            $groupOwnersFirstNames = $($($groupOwners.FirstName | Sort-Object FirstName) -join ", ")
            $groupOwnersFirstNames = $groupOwnersFirstNames -replace "(.*),(.*)", "`$1 &`$2"
            $body = "<HTML><FONT FACE=`"Calibri`">Hello $groupOwnersFirstNames`r`n`r`n<BR><BR>"
            $body += "Sorry, I'm rolling this name change back - our group names need to adhere to our  <A HREF=`"https://anthesisllc.sharepoint.com/sites/Resources-IT/_layouts/15/DocIdRedir.aspx?ID=HXX7CE52TSD2-1759992947-11`">Naming Conventions</A> to ensure everyone in Anthesis is talking a common language.`r`n`r`n<BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>The Helpful Groups Robot</FONT></HTML>"
            Send-MailMessage -From groupbot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Team name [$($combinedDg.DisplayName)] reverted from [$($thisUnifiedGroup.DisplayName)]" -BodyAsHtml $body -To "kevin.maitland@anthesisgroup.com" #$($groupOwners.WindowsLiveID) -Cc $groupAdmins
            }
        }
    if([string]::IsNullOrWhiteSpace($thisUnifiedGroup.SharePointSiteUrl)){
        Write-Verbose "Site [$($thisUnifiedGroup.DisplayName)] is not provisioned yet..."
        }
    else{
        #set-standardTeamSitePermissions -teamSiteAbsoluteUrl $thisUnifiedGroup.SharePointSiteUrl -fullArrayOfUnifiedGroups $allUnifiedGroups -Verbose
        Write-Verbose ""
        set-standardSitePermissions -unifiedGroupObject $thisUnifiedGroup -tokenResponse $tokenResponse -pnpCreds $sharePointCreds -Verbose
        }
    }

