$logFileLocation = "C:\ScriptLogs\"
$transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_Transcript_$(Get-Date -Format "yyMMdd").log"
if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)){
    $fullLogPathAndName = $logFileLocation+"record-licensingUsage_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = $logFileLocation+"record-licensingUsage_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
else{
    $fullLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
Start-Transcript $transcriptLogName -Append

Import-Module _PS_Library_MSOL.psm1
Import-Module _PS_Library_GeneralFunctionality
Import-Module _REST_Library-SPO.psm1

[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client") 
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Sharing") 
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Taxonomy") 
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.UserProfiles")

$logFile = $fullLogPathAndName
$errorLogFile = $errorLogPathAndName
$smtpServer = "anthesisgroup-com.mail.protection.outlook.com"

$upnSMA = "sustainmailboxaccess@anthesisgroup.com"
#$passSMA = ConvertTo-SecureString -String '' -AsPlainText -Force | ConvertFrom-SecureString
$passSMA =  ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\SustainMailboxAccess.txt) 
$msolCredentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $upnSMA, $passSMA
$restCredentials = new-spoCred -username $msolCredentials.UserName -securePassword $msolCredentials.Password
#$csomCredentials = new-csomCredentials -username $msolCredentials.UserName -password $msolCredentials.Password
connect-ToMsol -credential $msolCredentials
connect-ToExo -credential $msolCredentials
#connect-toAAD -credential $msolCredentials
#connect-ToSpo -credential $msolCredentials

$sharePointServerUrl = "https://anthesisllc.sharepoint.com"
$ITSite = "/teams/IT_Team_All_365"
$licensingListName = "365 Licensing Logs"
$itSiteDigest = new-spoDigest -serverUrl $sharePointServerUrl -sitePath $ITSite -restCreds $restCredentials -logFile $logFile -verboseLogging $true

$msolUsers = Get-MsolUser -all | ?{$_.Licenses.Count -gt 0}
$mailUsers = Get-Mailbox | ?{$msolUsers.UserPrincipalName -contains $_.MicrosoftOnlineServicesID}

$userHash = [ordered]@{}
$msolUsers | % {$userHash.Add($_.UserPrincipalName,@($_,$null))}
$mailUsers | % {$userHash[$_.MicrosoftOnlineServicesID][1] = $_}

$targetList = get-list -serverUrl $sharePointServerUrl -sitePath $ITSite -listName $licensingListName -restCreds $restCredentials -verboseLogging $verboseLogging -logFile $logFile
$timeStamp = Get-Date
$prettyLicenseNames = @{"AnthesisLLC:ENTERPRISEPACK" = "E3";"AnthesisLLC:EXCHANGEDESKLESS"="Kiosk";"AnthesisLLC:PROJECTPROFESSIONAL"="Project";"AnthesisLLC:STANDARDPACK"="E1";"AnthesisLLC:VISIOCLIENT"="Visio";"AnthesisLLC:WACONEDRIVESTANDARD"="OneDrive";"AnthesisLLC:ATP_ENTERPRISE"="AdvancedSpam";"AnthesisLLC:POWER_BI_STANDARD"="PowerBI";"AnthesisLLC:EMS"="Security"}

foreach($upn in $userHash.Keys){
    foreach($license in $userHash[$upn][0].Licenses){
        if($prettyLicenseNames.Keys -contains $license.AccountSkuId){$licenseName = $prettyLicenseNames[$license.AccountSkuId]}
        else{$licenseName = $license.AccountSkuId}
        if($userHash[$upn][1].CustomAttribute1 -ne ""){$businessEntity = $userHash[$upn][1].CustomAttribute1}
        else{$businessEntity = "Unknown"}
        $itemToAdd = @{Title=$userHash[$upn][0].DisplayName;UserPrincipalName=$upn;LicenseName=$licenseName;BusinessUnit=$businessEntity;TimeStamp=$timeStamp;Community=$userHash[$upn][0].Department;Country=$userHash[$upn][0].Country}
        new-itemInList -serverUrl $sharePointServerUrl -sitePath $ITSite -listName $licensingListName -predeterminedItemType $targetList.ListItemEntityTypeFullName -hashTableOfItemData $itemToAdd -restCreds $restCredentials -digest $itSiteDigest -verboseLogging $verboseLogging -logFile $logFile
        }
    }
Stop-Transcript