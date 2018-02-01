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
#convertTo-localisedSecureString ""
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
#Get the last Client modified in [/lists/Kimble Clients] to minimise the number of records to process
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
        $kpListItem = get-itemsInList -serverUrl $webUrl -sitePath $sitePath -listName "Kimble Clients" -oDataQuery "?&`$filter=KimbleId eq `'$($kimbleChangedClient.Id)`'" -restCreds $restCreds -logFile $fullLogPathAndName
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
                        $updatedItem = get-itemsInList -serverUrl $webUrl -sitePath $sitePath -listName "Kimble Clients" -oDataQuery "?`$filter=Id eq $($kpListItem.Id)" -restCreds $restCreds -logFile $fullLogPathAndName
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
    new-spoClient -kimbleClientObject $kimbleNewClient -webUrl $webUrl -sitePath $sitePath -spoClientList $kp -restCreds $restCreds -clientDigest $clientDigest -fullLogPathAndName $fullLogPathAndName
    }

#endregion

Stop-Transcript
