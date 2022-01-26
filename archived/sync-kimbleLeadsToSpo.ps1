$logFileLocation = "C:\ScriptLogs\"
$transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_Transcript_$(Get-Date -Format "yyMMdd").log"
if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)){
    $fullLogPathAndName = $logFileLocation+"sync-kimbleLeadsToSpo_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = $logFileLocation+"sync-kimbleLeadsToSpo_ErrorLog_$(Get-Date -Format "yyMMdd").log"
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
$listName = "Kimble Leads"
$sharePointAdmin = "kimblebot@anthesisgroup.com"
$smtpServer = "anthesisgroup-com.mail.protection.outlook.com"
$mailFrom = "scriptrobot@sustain.co.uk"
$mailTo = "kevin.maitland@anthesisgroup.com"
#convertTo-localisedSecureString "kimbleBotPasswordHere"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Downloads\KimbleBot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
$restCreds = new-spoCred -Credential -username $adminCreds.UserName -securePassword $adminCreds.Password
$csomCreds = new-csomCredentials -username $adminCreds.UserName -password $adminCreds.Password
$verboseLogging = $false

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
#Get the last Lead modified in /lists/Kimble Leads to minimise the number of records to process
try{
    log-action -myMessage "Getting new Digest for https://anthesisllc.sharepoint.com/clients" -logFile $fullLogPathAndName
    $clientDigest = new-spoDigest -serverUrl $webUrl -sitePath $sitePath -restCreds $restCreds -verboseLogging $verboseLogging -logFile $fullLogPathAndName
    if($clientDigest){log-result -myMessage "SUCCESS: New digest expires at $($clientDigest.expiryTime)" -logFile $fullLogPathAndName}
    else{log-result -myMessage "FAILED: Unable to retrieve digest" -logFile $fullLogPathAndName}
    }
catch{log-error -myError $_ -myFriendlyMessage "Error retrieving digest for https://anthesisllc.sharepoint.com/clients" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
try{
    log-action -myMessage "Getting List: [$listName]" -logFile $fullLogPathAndName
    $kp = get-list -serverUrl $webUrl  -sitePath $sitePath -listName $listName -restCreds $restCreds -verboseLogging $verboseLogging -logFile $fullLogPathAndName
    if($kp){log-result -myMessage "SUCCESS: List retrieved!" -logFile $fullLogPathAndName}
    else{log-result -myMessage "FAILED: Unable to retrieve list" -logFile $fullLogPathAndName}
    }
catch{log-error -myError $_ -myFriendlyMessage "Error retrieving List: [$listName]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}

#Get the Kimble Leads that have been modifed since the last update
$cutoffDate = (Get-Date (Get-Date $kp.LastItemModifiedDate).AddHours(-1) -Format s) #Look one hour behind just in case there is ever a delay between polling Kimble and updating SharePoint
#$cutoffDate = (Get-Date (Get-Date $kp.LastItemModifiedDate).AddYears(-1) -Format s) #Bodge this once for the initial Sync
$soqlQuery = "SELECT Name,Id,KimbleOne__Account__c,LastModifiedDate,SystemModStamp,CreatedDate,IsDeleted,Community__c,Project_Type__c FROM KimbleOne__SalesOpportunity__c WHERE LastModifiedDate > $cutoffDate`Z"
try{
    log-action -myMessage "Getting Kimble Lead data" -logFile $fullLogPathAndName
    $kimbleModifiedLeads = Get-KimbleSoqlDataset -queryUri $queryUri -soqlQuery $soqlQuery -restHeaders $kimbleRestHeaders
    if($kimbleModifiedLeads){log-result -myMessage "SUCCESS: $($kimbleModifiedLeads.Count) records retrieved!" -logFile $fullLogPathAndName}
    elseif($kimbleModifiedLeads -eq $null){log-result -myMessage "SUCCESS: Connected, but no records to update." -logFile $fullLogPathAndName}
    else{log-result -myMessage "FAILED: Unable to retrieve data!" -logFile $fullLogPathAndName}
    }
catch{log-error -myError $_ -myFriendlyMessage "Error retrieving Kimble Lead data" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
$kimbleChangedLeads = $kimbleModifiedLeads | ?{$_.LastModifiedDate -ge $cutoffDate} # -and $_.CreatedDate -lt $cutoffDate} #These can be both Created AND Modified if the user creates, then immediately updates the record
$kimbleNewLeads = $kimbleModifiedLeads | ?{$_.CreatedDate -ge $cutoffDate}
#Check any other Leads for changes
#At what point does it become more efficent to dump the whole [Kimble Leads] List from SP, rather than query individual items?
#SP pages results back in 100s, so when $spLead.Count/100 -gt $kimbleChangedLeads.Count, it will take fewer requests to query each $kimbleChangedLeads individually. This ought to happen most of the time (unless there is a batch update of Leads)

<# Use this is a batch update...
$spLeads = get-itemsInList -sitePath $sitePath -listName "Kimble Leads"
foreach($kimbleChangedLead in $kimbleChangedLeads){
    $spLead = $null
    $spLead = $spLeads | ?{$_.KimbleId -eq $kimbleChangedLead.Id}
    if($spLead){
        #Check whether spLead.Title = modLead.Name and update and mark IsDirty if necessary ;PreviousName=
        #if($spLead)
        }
    else{#Lead is missing from SP, so add it
        $kimbleNewLeads += $kimbleChangedLead
        }
    }
#>
#Otherwise, use this:
foreach($kimbleChangedLead in $kimbleChangedLeads){
    log-action -myMessage "CHANGED LEAD:`t[$($kimbleChangedLead.Name)] needs updating!" -logFile $fullLogPathAndName
    try{
        log-action -myMessage "Retrieving existing Lead from SPO" -logFile $fullLogPathAndName
        $kpListItem = get-itemsInList -serverUrl $webUrl -sitePath $sitePath -listName "Kimble Leads" -oDataQuery "?&`$filter=KimbleId eq `'$($kimbleChangedLead.Id)`'" -restCreds $restCreds -verboseLogging $verboseLogging -logFile $fullLogPathAndName
        if($kpListItem){
            log-result -myMessage "SUCCESS: list item [$($kpListItem.Title)] retrieved!" -logFile $fullLogPathAndName
            #Check whether the data has changed
            if($kpListItem.Title -ne $kimbleChangedLead.Name -or $kpListItem.KimbleClientId -ne $kimbleChangedLead.KimbleOne__Account__c -or $kpListItem.IsDeleted -ne $kimbleChangedLead.IsDeleted){
                #If it has, update the entry in [Kimble Leads]
                #SusChem don't want folders set up for specific sorts of Leads
                if(($kimbleChangedLead.Community__c -eq "UK - Sustainable Chemistry" -and ($kimbleChangedLead.Project_Type__c -eq "Only Representative (including TPR)" -or $kimbleChangedLead.Project_Type__c -eq "Registration Consortia"))){$doNotProcess = $true} #Exemption for specific SusChem projects
                    else{$doNotProcess = $false} #Everyone else wants Lead folders set up
                $updateData = @{PreviousName=$kpListItem.LeadName;PreviousKimbleClientId=$kpListItem.KimbleClientId;Title=$kimbleChangedLead.Name;KimbleClientId=$kimbleChangedLead.KimbleOne__Account__c;IsDeleted=$kimbleChangedLead.IsDeleted;IsDirty=$true;DoNotProcess=$doNotProcess}
                try{
                    log-action -myMessage "Updating SPO [Kimble Lead] item $($kpListItem.Title)" -logFile $fullLogPathAndName
                    update-itemInList -serverUrl $webUrl -sitePath $sitePath -listName "Kimble Leads" -predeterminedItemType $kp.ListItemEntityTypeFullName -itemId $kpListItem.Id -hashTableOfItemData $updateData -restCreds $restCreds -digest $clientDigest -verboseLogging $verboseLogging -logFile $fullLogPathAndName
                    #Check it's worked
                    try{
                        $updatedItem = get-itemsInList -serverUrl $webUrl -sitePath $sitePath -listName "Kimble Leads" -oDataQuery "?`$filter=Id eq $($kpListItem.Id)" -restCreds $restCreds -verboseLogging $verboseLogging -logFile $fullLogPathAndName
                        if($updateData.IsDirty -eq $true -and $updateData.Title -eq $kimbleChangedLead.Name){log-result -myMessage "SUCCESS: SPO [Kimble Lead] item $($kpListItem.Title) updated!" -logFile $fullLogPathAndName}
                        else{log-result -myMessage "FAILED: SPO [Kimble Lead] item $($kpListItem.Title) did not update correctly!" -logFile $fullLogPathAndName}
                        }
                    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving updated SPO [Kimble Lead] item $($kpListItem.Title)" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
                    }
                catch{log-error -myError $_ -myFriendlyMessage "Failed to update [Kimble Leads].$($kimbleChangedLead.Id) with @{$($($updateData.Keys | % {$_+":"+$updateData[$_]+","}) -join "`r")}" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
                }            
            else{
                log-result -myMessage "WARNING: SPO [Kimble Leads].[$($kpListItem.Title)] has changed, but I can't work out what needs changing (this might be because this Lead has alrady been processed, or because the changes don't affect the SPO object)." -logFile $fullLogPathAndName
                #$kimbleNewLeads += $kimbleChangedLead  #Only uncomment this to reprocess borked Leads
                }
            }
        else{
            log-result -myMessage "FAILED: Unable to retrieve SPO List Item for Kimble Lead [$($kimbleChangedLead.Name)]" -logFile $fullLogPathAndName
            #The List Item doesn't exist, so add it to the "to-be-created" array, even though we were expecting it to exist
            $kimbleNewLeads += $kimbleChangedLead 
            }
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving SPO List Item for Kimble Lead [$($kimbleChangedLead.Name)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
    }


#Add the new Leads
foreach ($kimbleNewLead in $kimbleNewLeads){
    log-action -myMessage "NEW LEAD:`t[$($kimbleNewLead.Name)] needs creating!" -logFile $fullLogPathAndName
    #SusChem don't want folders set up for specific types of Lead
    if(($kimbleNewLead.Community__c -eq "UK - Sustainable Chemistry" -and ($kimbleNewLead.Project_Type__c -eq "Only Representative (including TPR)" -or $kimbleNewLead.Project_Type__c -eq "Registration Consortia"))){$doNotProcess = $true} #Exemption for specific SusChem Leads
        else{$doNotProcess = $false} #Everyone else wants Lead folders set up
    $kimbleNewLeadData = @{KimbleId=$kimbleNewLead.Id;Title=$kimbleNewLead.Name;KimbleClientId=$kimbleNewLead.KimbleOne__Account__c;IsDeleted=$kimbleNewLead.IsDeleted;IsDirty=$true;DoNotProcess=$doNotProcess}
    #Create the new List item
    try{
        log-action -myMessage "Creating new SPO List item [$($kimbleNewLead.Name)]" -logFile $fullLogPathAndName
        $newItem = new-itemInList -serverUrl $webUrl -sitePath $sitePath -listName "Kimble Leads" -predeterminedItemType $kp.ListItemEntityTypeFullName -hashTableOfItemData $kimbleNewLeadData -restCreds $restCreds -digest $clientDigest -verboseLogging $TRUE -logFile $fullLogPathAndName
        #Check it's worked
        if($newItem){log-result -myMessage "SUCCESS: SPO [Kimble Lead] item $($newItem.Title) created!" -logFile $fullLogPathAndName}
        else{
            log-result -myMessage "FAILED: SPO [Kimble Lead] item $($kimbleNewLead.Name) did not create!" -logFile $fullLogPathAndName
            #Bodge this with an e-mail alert as we don't want Leads going missing
            Send-MailMessage -SmtpServer $smtpServer -To $mailTo -From $mailFrom -Subject "Kimble Lead [$($kimbleNewLead.Name)] could not be created in SPO" -Body "Lead: $($kimbleNewLead.Id)"
            }
        }
    catch{log-error -myError $_ -myFriendlyMessage "Failed to create new [Kimble Leads].$($kimbleNewLead.Name) with @{$($($kimbleNewLeadData.Keys | % {$_+":"+$kimbleNewLeadData[$_]+","}) -join "`r")}" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -smtpServer $smtpServer -mailTo $mailTo -mailFrom $mailFrom}
    }

#endregion



<##############################
#For building the initial Sync
###############################


$spLeads = get-itemsInList -sitePath $sitePath -listName "Kimble Leads" 
$remainingKimbleLeads = $kimbleModifiedLeads | ?{($spLeads | %{$_.KimbleId}) -notcontains $_.Id}

$remainingKimbleLeads = ,@();$j=0
foreach ($c in $kimbleModifiedLeads){
    $foundIt = $false
    foreach ($createdLead in $spLeads){
        if($c.Id -eq $createdLead.KimbleId){$foundIt= $true;break}
        }
    if(!$foundIt){$remainingKimbleLeads += $c}
    $j++
    if($j%100 -eq 0){$j}
    }

foreach ($kimbleNewLead in $remainingKimbleLeads){
#foreach ($kimbleNewLead in $kimbleNewLeads){
    $kimbleNewLeadData = @{KimbleId=$kimbleNewLead.Id;Title=$kimbleNewLead.Name;IsDeleted=$kimbleNewLead.IsDeleted;IsDirty=$true}
    switch ($kimbleNewLead.Description.Length){
        0 {break}
        {$_ -lt 255} {$kimbleNewLeadData.Add("LeadDescription","$($kimbleNewLead.Description)");break}
        default {$kimbleNewLeadData.Add("LeadDescription","$($kimbleNewLead.Description.Substring(0,254))")}
        }
    new-itemInList -sitePath $sitePath -listName "Kimble Leads" -predeterminedItemType $kp.ListItemEntityTypeFullName -hashTableOfItemData $kimbleNewLeadData
    }

$kimbleModifiedLeads.Count
$spLeads.Count
$remainingKimbleLeads.Count

#>
Stop-Transcript