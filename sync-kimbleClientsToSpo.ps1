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
#convertTo-localisedSecureString "kimbleBotPasswordHere"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\KimbleBot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
$restCreds = new-spoCred -Credential -username $adminCreds.UserName -securePassword $adminCreds.Password
$csomCreds = new-csomCredentials -username $adminCreds.UserName -password $adminCreds.Password


########################################
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



##################################
#
#Do Stuff
#
##################################
$oAuthReqBody = Get-KimbleAuthorizationTokenWithUsernamePasswordFlowRequestBody -client_id $clientId -client_secret $clientSecret -user_name $username -pass_word $password -security_token $securityToken
try{$kimbleAccessToken=Invoke-RestMethod -Method Post -Uri $callbackUri -Body $oAuthReqBody} catch {Failure}
$kimbleRestHeaders = @{Authorization = "Bearer " + $kimbleAccessToken.access_token}


#region Kimble Sync
#Get the last Project modified in /Projects/lists/Kimble Projects to minimise the number of records to process
try{
    log-action -myMessage "Getting new Digest for https://anthesisllc.sharepoint.com/clients" -logFile $fullLogPathAndName
    $clientDigest = new-spoDigest -serverUrl $webUrl -sitePath $sitePath -restCreds $restCreds
    if($clientDigest){log-result -myMessage "SUCCESS: New digest expires at $($clientDigest.expiryTime)" -logFile $fullLogPathAndName}
    else{log-result -myMessage "FAILED: Unable to retrieve digest" -logFile $fullLogPathAndName}
    }
catch{log-error -myError $_ -myFriendlyMessage "Error retrieving digest for https://anthesisllc.sharepoint.com/clients" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
try{
    log-action -myMessage "Getting List: [$listName]" -logFile $fullLogPathAndName
    $kp = get-list -serverUrl $webUrl  -sitePath $sitePath -listName $listName -restCreds $restCreds
    if($kp){log-result -myMessage "SUCCESS: List retrieved!" -logFile $fullLogPathAndName}
    else{log-result -myMessage "FAILED: Unable to retrieve list" -logFile $fullLogPathAndName}
    }
catch{log-error -myError $_ -myFriendlyMessage "Error retrieving List: [$listName]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}

#Get the Kimble Clients that have been modifed since the last update
$cutoffDate = (Get-Date (Get-Date $kp.LastItemModifiedDate).AddHours(-1) -Format s) #Look one hour behind just in case there is ever a delay between polling Kimble and updating SharePoint
#$cutoffDate = (Get-Date (Get-Date $kp.LastItemModifiedDate).AddYears(-1) -Format s) #Bodge this once for the initial Sync
$soqlQuery = "SELECT Name,Id,Description,Type,KimbleOne__IsCustomer__c,LastModifiedDate,SystemModStamp,CreatedDate,IsDeleted FROM account WHERE ((LastModifiedDate > $cutoffDate`Z) AND ((KimbleOne__IsCustomer__c = TRUE) OR (Type = 'Client') OR (Type = 'Potential Client')))"
try{
    log-action -myMessage "Getting Kimble Client data" -logFile $fullLogPathAndName
    $kimbleModifiedClients = Get-KimbleSoqlDataset -queryUri $queryUri -soqlQuery $soqlQuery -restHeaders $kimbleRestHeaders
    if($kimbleModifiedClients){log-result -myMessage "SUCCESS: $($kimbleModifiedClients.Count) records retrieved!" -logFile $fullLogPathAndName}
    elseif($kimbleModifiedClients -eq $null){log-result -myMessage "SUCCESS: Connected, but no records to update." -logFile $fullLogPathAndName}
    else{log-result -myMessage "FAILED: Unable to retrieve data!" -logFile $fullLogPathAndName}
    }
catch{log-error -myError $_ -myFriendlyMessage "Error retrieving Kimble Client data" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
$kimbleChangedClients = $kimbleModifiedClients | ?{$_.LastModifiedDate -ge $cutoffDate -and $_.CreatedDate -lt $cutoffDate}
$kimbleNewClients = $kimbleModifiedClients | ?{$_.CreatedDate -ge $cutoffDate}
#Check any other Clients for changes
#At what point does it become more efficent to dump the whole [Kimble Clients] List from SP, rather than query individual items?
#SP pages results back in 100s, so when $spClient.Count/100 -gt $kimbleChangedClients.Count, it will take fewer requests to query each $kimbleChangedClients individually. This ought to happen most of the time (unless there is a batch update of Clients)

<# Use this is a batch update...
$spClients = get-itemsInList -sitePath $sitePath -listName "Kimble Clients"
foreach($kimbleChangedClient in $kimbleChangedClients){
    $spClient = $null
    $spClient = $spClients | ?{$_.KimbleId -eq $kimbleChangedClient.Id}
    if($spClient){
        #Check whether spClient.Title = modClient.Name and update and mark IsDirty if necessary ;PreviousName=
        #if($spClient)
        }
    else{#Client is missing from SP, so add it
        $kimbleNewClients += $kimbleChangedClient
        }
    }
#>
#Otherwise, use this:
foreach($kimbleChangedClient in $kimbleChangedClients){
    log-action -myMessage "CHANGED CLIENT:`t[$($kimbleChangedClient.Name)] needs updating!" -logFile $fullLogPathAndName
    try{
        log-action -myMessage "Retrieving existing Client from SPO" -logFile $fullLogPathAndName
        $kpListItem = get-itemsInList -serverUrl $webUrl -sitePath $sitePath -listName "Kimble Clients" -oDataQuery "?&`$filter=KimbleId eq `'$($kimbleChangedClient.Id)`'" -restCreds $restCreds
        if($kpListItem){
            log-result -myMessage "SUCCESS: list item [$($kpListItem.Title)] retrieved!" -logFile $fullLogPathAndName
            #Check whether the data has changed
            if($kpListItem.Title -ne $kimbleChangedClient.Name -or $kpListItem.KimbleId -ne $kimbleChangedClient.Id -or $kpListItem.IsDeleted -ne $kimbleChangedClient.IsDeleted){
                #If it has, update the entry in [Kimble Clients]
                $updateData = @{PreviousName=$kpListItem.ClientName;PreviousDescription=$kpListItem.ClientDescription;Title=$kimbleChangedClient.Name;ClientDescription=$kimbleChangedClient.Description;IsDeleted=$kimbleChangedClient.IsDeleted;IsDirty=$true}
                try{
                    log-action -myMessage "Updating SPO [Kimble Client] item $($kpListItem.Title)" -logFile $fullLogPathAndName
                    update-itemInList -serverUrl $webUrl -sitePath $sitePath -listName "Kimble Clients" -predeterminedItemType $kp.ListItemEntityTypeFullName -itemId $kpListItem.Id -hashTableOfItemData $updateData -restCreds $restCreds -digest $clientDigest
                    #Check it's worked
                    try{
                        $updatedItem = get-itemsInList -serverUrl $webUrl -sitePath $sitePath -listName "Kimble Clients" -oDataQuery "?`$filter=Id eq $($kpListItem.Id)" -restCreds $restCreds
                        if($updateData.IsDirty -eq $true -and $updateData.Title -eq $kimbleChangedClient.Name){log-result -myMessage "SUCCESS: SPO [Kimble Client] item $($kpListItem.Title) updated!" -logFile $fullLogPathAndName}
                        else{log-result -myMessage "FAILED: SPO [Kimble Client] item $($kpListItem.Title) did not update correctly!" -logFile $fullLogPathAndName}
                        }
                    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving updated SPO [Kimble Client] item $($kpListItem.Title)" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
                    }
                catch{log-error -myError $_ -myFriendlyMessage "Failed to update [Kimble Clients].$($kimbleChangedClient.Id) with @{$($($updateData.Keys | % {$_+":"+$updateData[$_]+","}) -join "`r")}" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
                }            
            else{
                log-result -myMessage "WARNING: SPO [Kimble Clients].[$($kpListItem.Title)] has changed, but I can't work out what needs changing (this might be because this Client has already been processed, or because the changes don't affect the SPO object)." -logFile $fullLogPathAndName
                $kimbleNewProjects += $kimbleChangedProject 
                }
            }
        else{
            log-result -myMessage "FAILED: Unable to retrieve SPO List Item for Kimble Client [$($kimbleChangedClient.Name)]" -logFile $fullLogPathAndName
            #The List Item doesn't exist, so add it to the "to-be-created" array, even though we were expecting it to exist
            $kimbleNewClients += $kimbleChangedClient 
            }
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving SPO List Item for Kimble Client [$($kimbleChangedClient.Name)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
    }


#Add the new Clients
foreach ($kimbleNewClient in $kimbleNewClients){
    log-action -myMessage "NEW CLIENT:`t[$($kimbleNewClient.Name)] needs creating!" -logFile $fullLogPathAndName
    $kimbleNewClientData = @{KimbleId=$kimbleNewClient.Id;Title=$kimbleNewClient.Name;IsDeleted=$kimbleNewClient.IsDeleted;IsDirty=$true}
    #Create the new List item
    try{
        log-action -myMessage "Creating new SPO List item [$($kimbleNewClient.Name)]" -logFile $fullLogPathAndName
        $newItem = new-itemInList -serverUrl $webUrl -sitePath $sitePath -listName "Kimble Clients" -predeterminedItemType $kp.ListItemEntityTypeFullName -hashTableOfItemData $kimbleNewClientData -restCreds $restCreds -digest $clientDigest -verboseLogging $true -logFile "$env:USERPROFILE\Desktop\Log.txt"
        #Check it's worked
        if($newItem){log-result -myMessage "SUCCESS: SPO [Kimble Client] item $($newItem.Title) created!" -logFile $fullLogPathAndName}
        else{
            log-result -myMessage "FAILED: SPO [Kimble Client] item $($kimbleNewClient.Name) did not create!" -logFile $fullLogPathAndName
            #Bodge this with an e-mail alert as we don't want Projects going missing
            Send-MailMessage -SmtpServer $smtpServer -To $mailTo -From $mailFrom -Subject "Kimble Client [$($kimbleNewClient.Name)] could not be created in SPO" -Body "Project: $($kimbleNewClient.Id)"
            }
        }
    catch{log-error -myError $_ -myFriendlyMessage "Failed to create new [Kimble Clients].$($kimbleNewClient.Name) with @{$($($kimbleNewClientData.Keys | % {$_+":"+$kimbleNewClientData[$_]+","}) -join "`r")}" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -smtpServer $smtpServer -mailTo $mailTo -mailFrom $mailFrom}
    }

#endregion






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
        $clientDigest = new-spoDigest -serverUrl $webUrl -sitePath $sitePath -restCreds $restCreds
        if($clientDigest){log-result -myMessage "SUCCESS: New digest expires at $($clientDigest.expiryTime)" -logFile $fullLogPathAndName}
        else{log-result -myMessage "FAILED: Unable to retrieve digest" -logFile $fullLogPathAndName}
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving digest for https://anthesisllc.sharepoint.com/clients" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}

    try{
        log-action -myMessage "Getting List: [$listName]" -logFile $fullLogPathAndName
        $kp = get-itemsInList -serverUrl $webUrl  -sitePath $sitePath -listName $listName -restCreds $restCreds -logFile $logFileLocation 
        if($kp){log-result -myMessage "SUCCESS: List retrieved!" -logFile $fullLogPathAndName}
        else{log-result -myMessage "FAILED: Unable to retrieve list" -logFile $fullLogPathAndName}
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving List: [$listName]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}

    }





<#
Old stuff
$logFileLocation = "C:\ScriptLogs\"
$transcriptLogName = "$($logFileLocation+$($MyInvocation.PSCommandPath.Split("\")[$MyInvocation.PSCommandPath.Split("\").Count-1]))_Transcript_$(Get-Date -Format "yyMMdd").log"
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

##################################
#
#Get ready
#
##################################
$o365user = "kevin.maitland@anthesisgroup.com"
$o365Pass = ConvertTo-SecureString (Get-Content 'C:\New Text Document.txt') -AsPlainText -Force
$credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $o365user, $o365Pass

$logfile = "C:\Scripts\Logs\sync-kimbleClientsToSpo_$(Get-Date -Format "yyMMdd").log"
$logErrors = $true
$logMethodMain = $true
$logFunctionCalls = $true
Set-SPORestCredentials -Credential $credential

$oAuthReqBody = Get-KimbleAuthorizationTokenWithUsernamePasswordFlowRequestBody -client_id $clientId -client_secret $clientSecret -user_name $username -pass_word $password -security_token $securityToken
try{$kimbleAccessToken=Invoke-RestMethod -Method Post -Uri $callbackUri -Body $oAuthReqBody} catch {Failure}
$kimbleRestHeaders = @{Authorization = "Bearer " + $kimbleAccessToken.access_token}

##################################
#
#Do Stuff
#
##################################

#region Kimble Sync
#Get the last Client modified in /clients/lists/Kimble Clients to minimise the number of records to process
$serverUrl = "https://anthesisllc.sharepoint.com" 
$sitePath = "/clients"
$listName = "Kimble Clients"
get-newDigest -serverUrl $serverUrl -sitePath $sitePath
$kc = get-list -sitePath $sitePath -listName $listName
#$lastModifiedClient = get-lastModifiedItemInList -sitePath $sitePath -listName "Kimble Clients"

#Get the Kimble Clients that have been modifed since the last update
$cutoffDate = (Get-Date (Get-Date $kc.LastItemModifiedDate).AddHours(-1) -Format s) #Look one hour behind just in case there is ever a delay between polling Kimble and updating SharePoint
#$cutoffDate = (Get-Date (Get-Date $kc.LastItemModifiedDate).AddYears(-1) -Format s) #Bodge this once for the initial Sync
$soqlQuery = "SELECT Name,Id,Description,Type,KimbleOne__IsCustomer__c,LastModifiedDate,SystemModStamp,CreatedDate,IsDeleted FROM account WHERE ((LastModifiedDate > $cutoffDate`Z) AND ((KimbleOne__IsCustomer__c = TRUE) OR (Type = 'Client')))"
$kimbleModifiedClients = Get-KimbleSoqlDataset -queryUri $queryUri -soqlQuery $soqlQuery -restHeaders $kimbleRestHeaders
$kimbleChangedClients = $kimbleModifiedClients | ?{$_.LastModifiedDate -ge $cutoffDate}
$kimbleNewClients = $kimbleModifiedClients | ?{$_.CreatedDate -ge $cutoffDate}

#Check any other Clients for changes
#At what point does it become more efficent to dump the whole [Kimble Clients] List from SP, rather than query individual items?
#SP pages results back in 100s, so when $spClient.Count/100 -gt $kimbleChangedClients.Count, it will take fewer requests to query each $kimbleChangedClients individually. This ought to happen most of the time (unless there is a batch update of Clients)

<# Use this is a batch update...
$spClients = get-itemsInList -sitePath $sitePath -listName "Kimble Clients"

#foreach($kimbleChangedClient in $kimbleChangedClients){
#$kimbleNewClients = @()
for($j=0; $j -lt $kimbleChangedClients.Count;$j++){
    $spClient = $null
    if ($j -lt $kimbleChangedClients.Count/2){
        for ($i=0 ; $i -lt $spClients.Count;$i++){
            if ($spClient -ne $null){break}
            else{if($spClients[$i].KimbleId -eq $kimbleChangedClients[$j].Id){$spClient = $spClients[$i]}}
            }
        }
    else{
        for ($i=$spClients.Count-1 ; $i -ge 0;$i--){
            if ($spClient -ne $null){break}
            else{if($spClients[$i].KimbleId -eq $kimbleChangedClients[$j].Id){$spClient = $spClients[$i]}}
            }
        }
    if($spClient -eq $null){$kimbleNewClients += $kimbleChangedClients[$j]}
    if($j%100 -eq 0){Write-Host -ForegroundColor Magenta "$j / $($kimbleChangedClients.Count) _ $($kimbleNewClients.Count)"}
    }


#Otherwise, use this:
foreach($kimbleChangedClient in $kimbleChangedClients){
    $kCListItem = get-itemsInList -sitePath $sitePath -listName "Kimble Clients" -oDataQuery "?&`$filter=KimbleId eq `'$($kimbleChangedClient.Id)`'"
    if($kCListItem){
        #Check whether the data has changed
        if($kCListItem.Title -ne $kimbleChangedClient.Name `
            -or $kCListItem.ClientDescription -ne $kimbleChangedClient.Description `
            -or $kCListItem.IsDeleted -ne $kimbleChangedClient.IsDeleted){
            #Update the entry in [Kimble Clients]
            $updateData = @{PreviousName=$kCListItem.ClientName;PreviousDescription=$kCListItem.ClientDescription;Title=$kimbleChangedClient.Name;ClientDescription=$kimbleChangedClient.Description;IsDeleted=$kimbleChangedClient.IsDeleted;IsDirty=$true}
            try{update-itemInList -serverUrl $serverUrl -sitePath $sitePath -listNameOrGuid "Kimble Clients" -predeterminedItemType $kc.ListItemEntityTypeFullName -itemId $kCListItem.Id -hashTableOfItemData $updateData}
            catch{$false;log-error -myError $Error[0] -myFriendlyMessage "Failed to update [Kimble Clients].$($kimbleChangedClient.Id) with $updateData"}
            }
        }
    else{$kimbleNewClients += $kimbleChangedClient} #The Library doesn't exist, so add it to the "to-be-created" array, even though we were expecting it to exist
    }


#Add the new Clients to the SPO List
foreach ($kimbleNewClient in $kimbleNewClients){
    $kimbleNewClientData = @{KimbleId=$kimbleNewClient.Id;Title=$kimbleNewClient.Name;IsDeleted=$kimbleNewClient.IsDeleted;IsDirty=$true}
    if($kimbleNewClient.Description){$kimbleNewClientData.Add("ClientDescription","$($kimbleNewClient.Description)")}
    try{new-itemInList -sitePath $sitePath -listName "Kimble Clients" -predeterminedItemType $kc.ListItemEntityTypeFullName -hashTableOfItemData $kimbleNewClientData}
    catch{$false;log-error -myError $Error[0] -myFriendlyMessage "Failed to create new [Kimble Clients].$($kimbleNewClient.Id) with $kimbleNewClientData"}
    }


#endregion



<##############################
#For building the initial Sync
###############################


$spClients = get-itemsInList -sitePath $sitePath -listName "Kimble Clients" 
$remainingKimbleClients = $kimbleModifiedClients | ?{($spClients | %{$_.KimbleId}) -notcontains $_.Id}

$remainingKimbleClients = ,@();$j=0
foreach ($c in $kimbleModifiedClients){
    $foundIt = $false
    foreach ($createdClient in $spClients){
        if($c.Id -eq $createdClient.KimbleId){$foundIt= $true;break}
        }
    if(!$foundIt){$remainingKimbleClients += $c}
    $j++
    if($j%100 -eq 0){$j}
    }

foreach ($kimbleNewClient in $remainingKimbleClients){
#foreach ($kimbleNewClient in $kimbleNewClients){
    $kimbleNewClientData = @{KimbleId=$kimbleNewClient.Id;Title=$kimbleNewClient.Name;IsDeleted=$kimbleNewClient.IsDeleted;IsDirty=$true}
    switch ($kimbleNewClient.Description.Length){
        0 {break}
        {$_ -lt 255} {$kimbleNewClientData.Add("ClientDescription","$($kimbleNewClient.Description)");break}
        default {$kimbleNewClientData.Add("ClientDescription","$($kimbleNewClient.Description.Substring(0,254))")}
        }
    new-itemInList -sitePath $sitePath -listName "Kimble Clients" -predeterminedItemType $kc.ListItemEntityTypeFullName -hashTableOfItemData $kimbleNewClientData
    }

$kimbleModifiedClients.Count
$spClients.Count
$remainingKimbleClients.Count

#>
Stop-Transcript
#$kimbleModifiedClients | ?{$_.Name -match "Link"} | Select Name, Type, KimbleOne__IsCustomer__c | sort Name
#$kimbleModifiedClients | Select Name, Type, KimbleOne__IsCustomer__c | sort Name
#$kimbleModifiedClients.Count | ?{$_.Name -match "Linked"}
