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

#region Variables
##################################
#
#Get ready
#
##################################
#Don't change these unless the Kimble account or App changes
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$callbackUri = "https://login.salesforce.com/services/oauth2/token" #"https://test.salesforce.com/services/oauth2/token"
$grantType = "password"
$myInstance = "https://eu5.salesforce.com"
$queryUri = "$myInstance/services/data/v39.0/query/?q="
$querySuffixStub = " -H `"Authorization: Bearer "
$kimbleLogin = Import-Csv "$env:USERPROFILE\Desktop\Kimble.txt"
$clientId = $kimbleLogin.clientId
$clientSecret = $kimbleLogin.clientSecret
$username = $kimbleLogin.username
$password = $kimbleLogin.password
$securityToken = $kimbleLogin.securityToken
########################################
#Change these as required
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
$restCreds = new-spoCred -Credential -username $adminCreds.UserName -securePassword $adminCreds.Password
$csomCreds = new-csomCredentials -username $adminCreds.UserName -password $adminCreds.Password
########################################
$oAuthReqBody = Get-KimbleAuthorizationTokenWithUsernamePasswordFlowRequestBody -client_id $clientId -client_secret $clientSecret -user_name $username -pass_word $password -security_token $securityToken
try{$kimbleAccessToken=Invoke-RestMethod -Method Post -Uri $callbackUri -Body $oAuthReqBody} catch {Failure}
$kimbleRestHeaders = @{Authorization = "Bearer " + $kimbleAccessToken.access_token}
#endregion

#region Functions
function try-newListItem($webUrl, $sitePath, $newSpoItemData, $spoListToAddTo, $restCreds, $clientsDigest, $fullLogPathAndName){
    try{
        log-action -myMessage "Creating new SPO List item [$($newSpoItemData["Title"])] in [$($spoListToAddTo.Title)]" -logFile $fullLogPathAndName
        $newItem = new-itemInList -serverUrl $webUrl -sitePath $sitePath -listName $spoListToAddTo.Title -predeterminedItemType $spoListToAddTo.ListItemEntityTypeFullName -hashTableOfItemData $newSpoItemData -restCreds $restCreds -digest $clientsDigest -verboseLogging $true -logFile $fullLogPathAndName
        #Check it's worked
        if($newItem){log-result -myMessage "SUCCESS: SPO [$($spoListToAddTo.Title)] item $($newItem.Title) created!" -logFile $fullLogPathAndName}
        else{
            log-result -myMessage "FAILED: SPO [$($spoListToAddTo.Title)] item $($newSpoItemData["Title"]) did not create!" -logFile $fullLogPathAndName
            #Bodge this with an e-mail alert as we don't want Projects going missing
            Send-MailMessage -SmtpServer $smtpServer -To $mailTo -From $mailFrom -Subject "[$($spoListToAddTo.Title)].[$($newSpoItemData["Title"])] could not be created in SPO" -Body "[$($spoListToAddTo.Title)]: $($newSpoItemData["KimbleId"])"
            }
        }
    catch{log-error -myError $_ -myFriendlyMessage "Failed to create new [Kimble Leads].$($kimbleLeadObject.Name) with @{$($($newSpoLeadData.Keys | % {$_+":"+$newSpoLeadData[$_]+","}) -join "`r")}" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -smtpServer $smtpServer -mailTo $mailTo -mailFrom $mailFrom}
    
    }
function new-spoClient($kimbleClientObject, $webUrl, $sitePath, $spoClientList, $restCreds, $clientsDigest, $fullLogPathAndName){
    log-action -myMessage "CREATING NEW CLIENT:`t[$($kimbleClientObject.Name)]" -logFile $fullLogPathAndName
    $newSpoClientData = @{KimbleId=$kimbleClientObject.Id;Title=$kimbleClientObject.Name;IsDeleted=$kimbleClientObject.IsDeleted;IsDirty=$true}
    #Create the new List item
    try-newListItem -webUrl $webUrl -sitePath $sitePath -newSpoItemData $newSpoClientData -spoListToAddTo $spoClientList -restCreds $restCreds -clientDigest $clientsDigest -fullLogPathAndName $fullLogPathAndName
    }
function new-spoLead($kimbleLeadObject, $webUrl, $sitePath, $spoLeadsList, $restCreds, $clientsDigest, $fullLogPathAndName){
    log-action -myMessage "CREATING NEW LEAD:`t[$($kimbleLeadObject.Name)]" -logFile $fullLogPathAndName
    $newSpoLeadData = @{KimbleId=$kimbleLeadObject.Id;Title=$kimbleLeadObject.Name;IsDeleted=$kimbleLeadObject.IsDeleted;IsDirty=$true}
    #Create the new List item
    try-newListItem -webUrl $webUrl -sitePath $sitePath -newSpoItemData $newSpoLeadData -spoListToAddTo $spoLeadsList -restCreds $restCreds -clientDigest $clientsDigest -fullLogPathAndName $fullLogPathAndName
    }
function new-spoProject($kimbleProjectObject, $webUrl, $sitePath, $spoProjectList, $restCreds, $clientsDigest, $fullLogPathAndName){
    log-action -myMessage "CREATING NEW PROJECT:`t[$($kimbleProjectObject.Name)]" -logFile $fullLogPathAndName
    $newSpoProjectData = @{KimbleId=$kimbleProjectObject.Id;Title=$kimbleProjectObject.Name;IsDeleted=$kimbleProjectObject.IsDeleted;IsDirty=$true}
    #Create the new List item
    try-newListItem -webUrl $webUrl -sitePath $sitePath -newSpoItemData $newSpoProjectData -spoListToAddTo $spoProjectList -restCreds $restCreds -clientDigest $clientsDigest -fullLogPathAndName $fullLogPathAndName
    }
function reconcile-leads(){
    #Get the full list of Kimble Leads
    $soqlQuery = "SELECT Name,Id,KimbleOne__Account__c,LastModifiedDate,SystemModStamp,CreatedDate,IsDeleted,Community__c,Project_Type__c FROM KimbleOne__SalesOpportunity__c"
    try{
        log-action -myMessage "Getting Kimble Lead data from SalesForce" -logFile $fullLogPathAndName
        $allKimbleLeads = Get-KimbleSoqlDataset -queryUri $queryUri -soqlQuery $soqlQuery -restHeaders $kimbleRestHeaders
        if($allKimbleLeads){log-result -myMessage "SUCCESS: $($kimbleModifiedLeads.Count) records retrieved!" -logFile $fullLogPathAndName}
        else{log-result -myMessage "FAILED: Unable to retrieve data!" -logFile $fullLogPathAndName}
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving Kimble Lead data" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
    
    #Get the full list of SPO Leads
    try{
        log-action -myMessage "Getting List Items: [Kimble Leads]" -logFile $fullLogPathAndName
        $spoLeadsItems = get-itemsInList -serverUrl $webUrl  -sitePath $sitePath -listName "Kimble Leads" -restCreds $restCreds -logFile $fullLogPathAndName
        if($spoLeadsItems){log-result -myMessage "SUCCESS: List retrieved!" -logFile $fullLogPathAndName}
        else{log-result -myMessage "FAILED: Unable to retrieve list" -logFile $fullLogPathAndName}
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving List: [Kimble Leads]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}

    #Get the List to make creating new Items easier
    try{
        log-action -myMessage "Getting List: [Kimble Leads]" -logFile $fullLogPathAndName
        $spoLeadsList = get-list -serverUrl $webUrl  -sitePath $sitePath -listName "Kimble Leads" -restCreds $restCreds
        if($spoLeadsList){log-result -myMessage "SUCCESS: List retrieved!" -logFile $fullLogPathAndName}
        else{log-result -myMessage "FAILED: Unable to retrieve list" -logFile $fullLogPathAndName}
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving List: [$listName]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}

    #Get a digest so we can write stuff back
    try{
        log-action -myMessage "Getting new Digest for https://anthesisllc.sharepoint.com/clients" -logFile $fullLogPathAndName
        $clientsDigest = new-spoDigest -serverUrl $webUrl -sitePath $sitePath -restCreds $restCreds
        if($clientsDigest){log-result -myMessage "SUCCESS: New digest expires at $($clientsDigest.expiryTime)" -logFile $fullLogPathAndName}
        else{log-result -myMessage "FAILED: Unable to retrieve digest" -logFile $fullLogPathAndName}
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving digest for https://anthesisllc.sharepoint.com/clients" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}

    #Work out what's missing and create any omissions
    $allKimbleLeads | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name KimbleId -Value $_.Id}
    $missingLeads = Compare-Object -ReferenceObject $allKimbleLeads -DifferenceObject $spoLeadsItems -Property "KimbleId" -PassThru -CaseSensitive:$false
    $missingLeads | % {
        if ($_.SideIndicator -eq "<="){new-spoLead -kimbleLeadObject $_ -webUrl $webUrl -sitePath $sitePath -spoLeadsList $spoLeadsList -restCreds $restCreds -clientDigest $clientsDigest -fullLogPathAndName $fullLogPathAndName}
        }

    }
function reconcile-projects(){
    #Get the full list of Kimble Projects
    $soqlQuery = "SELECT Name,Id,KimbleOne__Account__c,LastModifiedDate,SystemModStamp,CreatedDate,IsDeleted,Community__c,Project_Type__c FROM KimbleOne__DeliveryGroup__c"
    try{
        log-action -myMessage "Getting Kimble Project data from SalesForce" -logFile $fullLogPathAndName
        $allKimbleProjects = Get-KimbleSoqlDataset -queryUri $queryUri -soqlQuery $soqlQuery -restHeaders $kimbleRestHeaders
        if($allKimbleProjects){log-result -myMessage "SUCCESS: $($kimbleModifiedProjects.Count) records retrieved!" -logFile $fullLogPathAndName}
        else{log-result -myMessage "FAILED: Unable to retrieve data!" -logFile $fullLogPathAndName}
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving Kimble Project data" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
    
    #Get the full list of SPO Projects
    try{
        log-action -myMessage "Getting List Items: [Kimble Projects]" -logFile $fullLogPathAndName
        $spoProjectsItems = get-itemsInList -serverUrl $webUrl  -sitePath $sitePath -listName "Kimble Projects" -restCreds $restCreds -logFile $fullLogPathAndName
        if($spoProjectsItems){log-result -myMessage "SUCCESS: List retrieved!" -logFile $fullLogPathAndName}
        else{log-result -myMessage "FAILED: Unable to retrieve list" -logFile $fullLogPathAndName}
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving List: [Kimble Projects]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}

    #Get the List to make creating new Items easier
    try{
        log-action -myMessage "Getting List: [Kimble Projects]" -logFile $fullLogPathAndName
        $spoProjectsList = get-list -serverUrl $webUrl  -sitePath $sitePath -listName "Kimble Projects" -restCreds $restCreds
        if($spoProjectsList){log-result -myMessage "SUCCESS: List retrieved!" -logFile $fullLogPathAndName}
        else{log-result -myMessage "FAILED: Unable to retrieve list" -logFile $fullLogPathAndName}
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving List: [$listName]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}

    #Get a digest so we can write stuff back
    try{
        log-action -myMessage "Getting new Digest for https://anthesisllc.sharepoint.com/clients" -logFile $fullLogPathAndName
        $clientsDigest = new-spoDigest -serverUrl $webUrl -sitePath $sitePath -restCreds $restCreds
        if($clientsDigest){log-result -myMessage "SUCCESS: New digest expires at $($clientsDigest.expiryTime)" -logFile $fullLogPathAndName}
        else{log-result -myMessage "FAILED: Unable to retrieve digest" -logFile $fullLogPathAndName}
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving digest for https://anthesisllc.sharepoint.com/clients" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}

    #Work out what's missing and create any omissions
    $allKimbleProjects | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name KimbleId -Value $_.Id}
    $missingProjects = Compare-Object -ReferenceObject $allKimbleProjects -DifferenceObject $spoProjectsItems -Property "KimbleId" -PassThru -CaseSensitive:$false
    $missingProjects | % {
        if ($_.SideIndicator -eq "<="){new-spoProject -kimbleProjectObject $_ -webUrl $webUrl -sitePath $sitePath -spoProjectList $spoProjectsList -restCreds $restCreds -clientDigest $clientsDigest -fullLogPathAndName $fullLogPathAndName}
        }

    }

function reconcile-clients(){
    #Get the full list of Kimble Clients 
    $soqlQuery = "SELECT Name,Id,Description,Type,KimbleOne__IsCustomer__c,LastModifiedDate,SystemModStamp,CreatedDate,IsDeleted FROM account WHERE ((KimbleOne__IsCustomer__c = TRUE) OR (Type = 'Client') OR (Type = 'Potential Client'))"
    try{
        log-action -myMessage "Getting Kimble Client data" -logFile $fullLogPathAndName
        $allKimbleClients = Get-KimbleSoqlDataset -queryUri $queryUri -soqlQuery $soqlQuery -restHeaders $kimbleRestHeaders
        if($allKimbleClients){log-result -myMessage "SUCCESS: $($kimbleModifiedClients.Count) records retrieved!" -logFile $fullLogPathAndName}
        else{log-result -myMessage "FAILED: Unable to retrieve data!" -logFile $fullLogPathAndName}
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving Kimble Client data" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
    #Get the full list of SPO Clients
    try{
        log-action -myMessage "Getting new Digest for https://anthesisllc.sharepoint.com/clients" -logFile $fullLogPathAndName
        $clientsDigest = new-spoDigest -serverUrl $webUrl -sitePath $sitePath -restCreds $restCreds
        if($clientsDigest){log-result -myMessage "SUCCESS: New digest expires at $($clientsDigest.expiryTime)" -logFile $fullLogPathAndName}
        else{log-result -myMessage "FAILED: Unable to retrieve digest" -logFile $fullLogPathAndName}
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving digest for https://anthesisllc.sharepoint.com/clients" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}

    try{
        log-action -myMessage "Getting List: [Kimble Clients]" -logFile $fullLogPathAndName
        $spoClients = get-itemsInList -serverUrl $webUrl  -sitePath $sitePath -listName "Kimble Clients" -restCreds $restCreds -logFile $fullLogPathAndName 
        if($spoClients){log-result -myMessage "SUCCESS: List retrieved!" -logFile $fullLogPathAndName}
        else{log-result -myMessage "FAILED: Unable to retrieve list" -logFile $fullLogPathAndName}
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving List: [Kimble Clients]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
    
    try{
        log-action -myMessage "Getting List: [$listName]" -logFile $fullLogPathAndName
        $kp = get-list -serverUrl $webUrl  -sitePath $sitePath -listName $listName -restCreds $restCreds
        if($kp){log-result -myMessage "SUCCESS: List retrieved!" -logFile $fullLogPathAndName}
        else{log-result -myMessage "FAILED: Unable to retrieve list" -logFile $fullLogPathAndName}
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving List: [$listName]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}

    $allKimbleClients | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name KimbleId -Value $_.Id}
    $missingClients = Compare-Object -ReferenceObject $allKimbleClients -DifferenceObject $spoClients -Property "KimbleId" -PassThru -CaseSensitive:$false
    $missingClients | % {
        if ($_.SideIndicator -eq "<="){new-spoClient -kimbleClientObject $_ -webUrl $webUrl -sitePath $sitePath -spoClientList $kp -restCreds $restCreds -clientDigest $clientsDigest -fullLogPathAndName $fullLogPathAndName}
        }
    }
#endregion


##################################
#
#Do Stuff
#
##################################
