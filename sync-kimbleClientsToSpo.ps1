$logFileLocation = "C:\ScriptLogs\"
if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)){
    $fullLogPathAndName = $logFileLocation+"sync-kimbleClientsToSpo_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = $logFileLocation+"sync-kimbleClientsToSpo_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
else{
    $fullLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
if($PSCommandPath){
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_Transcript_$(Get-Date -Format "yyMMdd").log"
    Start-Transcript $transcriptLogName -Append
    }

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

#region Kimble Sync
#Get the last Client modified in [/lists/Kimble Clients] to minimise the number of records to process
try{
    log-action -myMessage "Getting [Kimble Clients] to minimise the number of records to process" -logFile $fullLogPathAndName 
    $kc = Get-PnPList -Identity "Kimble Clients" -Includes ContentTypes, LastItemModifiedDate
    if($kc){log-result -myMessage "SUCCESS: List retrieved" -logFile $fullLogPathAndName}
    else{log-result -myMessage "FAILURE: List could not be retrieved" -logFile $fullLogPathAndName}
    }
catch{log-error -myError $_ -myFriendlyMessage "Could not retrieve [Kimble Clients]" -fullLogFile $fullLogPathAndName -errorLogFile -doNotLogToEmail $true}

#Get the Kimble Clients that have been modifed since the last update
Get-PnPListItem -List "Kimble Clients" -Query "<View><Query> <OrderBy> <FieldRef Name='LastModifiedDate' Ascending='False' /> </OrderBy> </Query> </View>" -PageSize 10 -ErrorAction SilentlyContinue | % {if($dummyArray){rv dummyArray};[array]$dummyArray += $_;break} #Get the list item with the most recent LastModifedDate (from Kimble)
$cutoffDate = Get-Date $dummyArray[0].FieldValues.LastModifiedDate -Format s
#$cutoffDate = (Get-Date (Get-Date $kc.LastItemModifiedDate).AddHours(-1) -Format s) #Look one hour behind just in case there is ever a delay between polling Kimble and updating SharePoint
#$cutoffDate = (Get-Date (Get-Date $kp.LastItemModifiedDate).AddYears(-1) -Format s) #Bodge this once for the initial Sync
$soqlQuery = "SELECT Name,Id,Description,Type,KimbleOne__IsCustomer__c,LastModifiedDate,SystemModStamp,CreatedDate,IsDeleted FROM account WHERE ((LastModifiedDate > $cutoffDate`Z) AND ((KimbleOne__IsCustomer__c = TRUE) OR (Type = 'Client') OR (Type = 'Potential Client')))"
try{
    log-action -myMessage "Getting Kimble Client data" -logFile $fullLogPathAndName
    $kimbleModifiedClients = Get-KimbleSoqlDataset -queryUri $standardKimbleQueryUri -soqlQuery $soqlQuery -restHeaders $standardKimbleHeaders
    if($kimbleModifiedClients){log-result -myMessage "SUCCESS: $($kimbleModifiedClients.Count) records retrieved!" -logFile $fullLogPathAndName}
    elseif($kimbleModifiedClients -eq $null){log-result -myMessage "SUCCESS: Connected, but no records to update." -logFile $fullLogPathAndName}
    else{log-result -myMessage "FAILED: Unable to retrieve data!" -logFile $fullLogPathAndName}
    }
catch{log-error -myError $_ -myFriendlyMessage "Error retrieving Kimble Client data" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
$kimbleChangedClients = $kimbleModifiedClients | ?{$_.LastModifiedDate -ge $cutoffDate -and $_.CreatedDate -lt $cutoffDate}
$kimbleNewClients = $kimbleModifiedClients | ?{$_.CreatedDate -ge $cutoffDate}

foreach($kimbleChangedClient in $kimbleChangedClients){
    log-action -myMessage "CHANGED CLIENT:`t[$($kimbleChangedClient.Name)] needs updating!" -logFile $fullLogPathAndName
    try{
        $updatedClient = update-spoKimbleClientItem -kimbleClientObject $kimbleChangedClient -pnpClientList $kc -fullLogPathAndName $fullLogPathAndName -verboseLogging $verboseLogging
        }
    catch{
        log-error -myError $_ -myFriendlyMessage "Error updating Client [$($kimbleChangedClient.Name)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
        }
    if($updatedClient){log-result -myMessage "SUCCESS: Looks like that worked!" -logFile $fullLogPathAndName}
    else{
        log-result -myMessage "FAILED: Looks like Client [$($kimbleChangedClient.Name)] didn't update correctly - will send it fcor re-creation" -logFile $fullLogPathAndName
        $kimbleNewClients += $kimbleChangedClient
        }
    }
<#        log-action -myMessage "Retrieving existing Client from SPO" -logFile $fullLogPathAndName
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
#>

#Add the new Clients
foreach ($kimbleNewClient in $kimbleNewClients){
    #new-spoClient -kimbleClientObject $kimbleNewClient -webUrl $webUrl -sitePath $sitePath -spoClientList $kp -restCreds $restCreds -clientDigest $clientDigest -fullLogPathAndName $fullLogPathAndName
    log-action -myMessage "NEW CLIENT:`t[$($kimbleNewClient.Name)] needs creating!" -logFile $fullLogPathAndName
    try{
        $newClient = new-spoKimbleClientItem -kimbleClientObject $kimbleNewClient -pnpClientList $kc -fullLogPathAndName $fullLogPathAndName -verboseLogging $verboseLogging
        }
    catch{
        log-error -myError $_ -myFriendlyMessage "Error creating Client [$($kimbleNewClient.Name)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
        }
    if($newClient){log-result -myMessage "SUCCESS: Looks like that worked!" -logFile $fullLogPathAndName}
    else{
        log-result -myMessage "FAILED: Looks like Client [$($kimbleNewClient.Name)] didn't create correctly :(  - that's a bit of a problem!" -logFile $fullLogPathAndName
        }
    }

#endregion

Stop-Transcript
