$logFileLocation = "C:\ScriptLogs\"
$transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_Transcript_$(Get-Date -Format "yyMMdd").log"
if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)){
    $fullLogPathAndName = $logFileLocation+"sync-kimbleClientsToSpo_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = $logFileLocation+"sync-kimbleClientsToSpo_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
else{
    $fullLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
Start-Transcript $transcriptLogName -Append

Import-Module _PS_Library_GeneralFunctionality
Import-Module _CSOM_Library-SPO.psm1
Import-Module _REST_Library-Kimble.psm1
Import-Module _REST_Library-SPO.psm1
Import-Module SharePointPnPPowerShellOnline

########################################
$webUrl = "https://anthesisllc.sharepoint.com" 
$sitePath = "/clients"
$listName = "Kimble Clients"
$smtpServer = "anthesisgroup-com.mail.protection.outlook.com"
#$mailFrom = "scriptrobot@sustain.co.uk"
#$mailTo = "kevin.maitland@anthesisgroup.com"
$sharePointAdmin = "kimblebot@anthesisgroup.com"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\KimbleBot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
$restCreds2 = new-spoCred -Credential -username $sharePointAdmin -securePassword $sharePointAdminPass
$csomCreds = new-csomCredentials -username $adminCreds.UserName -password $adminCreds.Password
########################################
Connect-PnPOnline –Url $($webUrl+$sitePath) –Credentials $adminCreds
########################################


function reconcile-clientsInSpo(){
    $clientList = Get-PnPList -Identity $listName -Includes ContentTypes
    $clientListContentType = $clientList.ContentTypes | ? {$_.Name -eq "Item"}
    $clientListItems = Get-PnPListItem -List $listName -PageSize 1000 -Fields "Title","GUID","KimbleId","ClientDescription","IsDirty","IsDeleted","Modified","LastModifiedDate","PreviousName","PreviousDescription","Id"
    $clientFoldersAlreadyCreated = Get-PnPList
    $clientFoldersAlreadyCreated = get-allLists -serverUrl $webUrl -sitePath $sitePath -restCreds $restCreds2 -logFile $fullLogPathAndName -verboseLogging $true 
    $dummy2 = get-itemsInList -serverUrl $webUrl -sitePath $sitePath -listName "ITCoreNet" -restCreds $restCreds2 -logFile $fullLogPathAndName -verboseLogging $true
    }