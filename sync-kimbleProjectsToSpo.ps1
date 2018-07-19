$logFileLocation = "C:\ScriptLogs\"
$transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_Transcript_$(Get-Date -Format "yyMMdd").log"
if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)){
    $fullLogPathAndName = $logFileLocation+"sync-kimbleProjectsToSpo_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = $logFileLocation+"sync-kimbleProjectsToSpo_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
else{
    $fullLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
Start-Transcript $transcriptLogName -Append

Import-Module _PS_Library_GeneralFunctionality
#Import-Module _CSOM_Library-SPO.psm1
Import-Module _REST_Library-Kimble.psm1
#Import-Module _REST_Library-SPO.psm1
Import-Module _PNP_Library_SPO

##################################
#
#Get ready
#
##################################
$webUrl = "https://anthesisllc.sharepoint.com" 
$sitePath = "/clients"
$listName = "Kimble Clients"
$sharePointAdmin = "kimblebot@anthesisgroup.com"
$smtpServer = "anthesisgroup-com.mail.protection.outlook.com"
$mailFrom = "scriptrobot@sustain.co.uk"
$mailTo = "kevin.maitland@anthesisgroup.com"
#convertTo-localisedSecureString ""
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\KimbleBot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
$csomCreds = new-csomCredentials -username $adminCreds.UserName -password $adminCreds.Password
$restCreds = new-spoCred -Credential -username $adminCreds.UserName -securePassword $adminCreds.Password
$kimbleCreds = Import-Csv "$env:USERPROFILE\Desktop\Kimble.txt"
$standardKimbleHeaders = get-kimbleHeaders -clientId $kimbleCreds.clientId -clientSecret $kimbleCreds.clientSecret -username $kimbleCreds.username -password $kimbleCreds.password -securityToken $kimbleCreds.securityToken -connectToLiveContext $true -verboseLogging $true
$standardKimbleQueryUri = get-kimbleQueryUri

Connect-PnPOnline -Url $($webUrl+$sitePath) -Credentials $adminCreds


##################################
#
#Do Stuff
#
##################################

#region Kimble Sync
#Get the last Project modified in /Projects/lists/Kimble Projects to minimise the number of records to process
try{
    log-action -myMessage "Getting [Kimble Projects] to minimise the number of records to process" -logFile $fullLogPathAndName 
    $kp = Get-PnPList -Identity "Kimble Projects" -Includes ContentTypes, LastItemModifiedDate
    if($kp){log-result -myMessage "SUCCESS: List retrieved" -logFile $fullLogPathAndName}
    else{log-result -myMessage "FAILURE: List could not be retrieved" -logFile $fullLogPathAndName}
    }
catch{log-error -myError $_ -myFriendlyMessage "Error retrieving List: [$listName]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}

#Get the Kimble Projects that have been modifed since the last update
$cutoffDate = (Get-Date (Get-Date $kp.LastItemModifiedDate).AddHours(-1) -Format s) #Look one hour behind just in case there is ever a delay between polling Kimble and updating SharePoint
#$cutoffDate = (Get-Date (Get-Date $kp.LastItemModifiedDate).AddYears(-1) -Format s) #Bodge this once for the initial Sync
$soqlQuery = "SELECT Name,Id,KimbleOne__Account__c,LastModifiedDate,SystemModStamp,CreatedDate,IsDeleted,Community__c,Project_Type__c FROM KimbleOne__DeliveryGroup__c WHERE LastModifiedDate > $cutoffDate`Z"
try{
    log-action -myMessage "Getting Kimble Project data" -logFile $fullLogPathAndName
    $kimbleModifiedProjects = Get-KimbleSoqlDataset -queryUri $standardKimbleQueryUri -soqlQuery $soqlQuery -restHeaders $standardKimbleHeaders
    if($kimbleModifiedProjects){log-result -myMessage "SUCCESS: $($kimbleModifiedProjects.Count) records retrieved!" -logFile $fullLogPathAndName}
    elseif($kimbleModifiedProjects -eq $null){log-result -myMessage "SUCCESS: Connected, but no records to update." -logFile $fullLogPathAndName}
    else{log-result -myMessage "FAILED: Unable to retrieve data!" -logFile $fullLogPathAndName}
    }
catch{log-error -myError $_ -myFriendlyMessage "Error retrieving Kimble Project data" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
$kimbleChangedProjects = $kimbleModifiedProjects | ?{$_.LastModifiedDate -ge $cutoffDate} #-and $_.CreatedDate -lt $cutoffDate} These can be both Created and Modified
$kimbleNewProjects = $kimbleModifiedProjects | ?{$_.CreatedDate -ge $cutoffDate}


foreach($kimbleChangedProject in $kimbleChangedProjects){
    log-action -myMessage "CHANGED PROJECT:`t[$($kimbleChangedProject.Name)] needs updating!" -logFile $fullLogPathAndName
    try{
        $updatedProject = update-spoKimbleProjectItem -kimbleProjectObject $kimbleChangedProject -pnpProjectList $kp -fullLogPathAndName $fullLogPathAndName -verboseLogging $verboseLogging
        }
    catch{
        log-error -myError $_ -myFriendlyMessage "Error updating Project [$($kimbleChangedProject.Name)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
        }
    if($updatedProject){log-result -myMessage "SUCCESS: Looks like that worked!" -logFile $fullLogPathAndName}
    else{
        log-result -myMessage "FAILED: Looks like Project [$($kimbleChangedProject.Name)] didn't update correctly - will send it fcor re-creation" -logFile $fullLogPathAndName
        $kimbleNewProjects += $kimbleChangedProject
        }
    }
foreach ($kimbleNewProject in $kimbleNewProjects){
    log-action -myMessage "NEW PROJECT:`t[$($kimbleNewProject.Name)] needs creating!" -logFile $fullLogPathAndName
    try{
        $newProject = new-spoKimbleProjectItem -kimbleProjectObject $kimbleNewProject -pnpProjectList $kp -fullLogPathAndName $fullLogPathAndName -verboseLogging $verboseLogging
        }
    catch{
        log-error -myError $_ -myFriendlyMessage "Error creating Project [$($kimbleNewProject.Name)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
        }
    if($newProject){log-result -myMessage "SUCCESS: Looks like that worked!" -logFile $fullLogPathAndName}
    else{
        log-result -myMessage "FAILED: Looks like Project [$($kimbleNewProject.Name)] didn't create correctly :(  - that's a bit of a problem!" -logFile $fullLogPathAndName
        }
    }

Stop-Transcript