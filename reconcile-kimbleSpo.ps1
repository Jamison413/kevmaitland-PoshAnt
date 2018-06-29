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

#region Variables
##################################
#
#Get ready
#
########################################
$webUrl = "https://anthesisllc.sharepoint.com" 
$sitePath = "/clients"
$listName = "Kimble Clients"
$smtpServer = "anthesisgroup-com.mail.protection.outlook.com"
$mailFrom = "scriptrobot@sustain.co.uk"
$mailTo = "kevin.maitland@anthesisgroup.com"
#convertTo-localisedSecureString ""
$sharePointAdmin = "kimblebot@anthesisgroup.com"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\KimbleBot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
$restCreds = new-spoCred -Credential -username $adminCreds.UserName -securePassword $adminCreds.Password
$csomCreds = new-csomCredentials -username $adminCreds.UserName -password $adminCreds.Password
########################################
$kimbleCreds = Import-Csv "$env:USERPROFILE\Desktop\Kimble.txt"
$standardKimbleHeaders = get-kimbleHeaders -clientId $kimbleCreds.clientId -clientSecret $kimbleCreds.clientSecret -username $kimbleCreds.username -password $kimbleCreds.password -securityToken $kimbleCreds.securityToken -connectToLiveContext $true -verboseLogging $true
$standardKimbleQueryUri = get-kimbleQueryUri
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
function reconcile-leadsBetweenKimbleAndSpo(){
    #Get the full list of Kimble Leads
    $allKimbleLeads = get-allKimbleLeads -pQueryUri $standardKimbleQueryUri -pRestHeaders $standardKimbleHeaders
    
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
function reconcile-clientsBetweenKimbleAndSpo(){
    #Get the full list of Kimble Clients
    $allKimbleClients = get-allKimbleAccounts -pQueryUri $standardKimbleQueryUri -pRestHeaders $standardKimbleHeaders -pWhereStatement "WHERE ((KimbleOne__IsCustomer__c = TRUE) OR (Type = 'Client') OR (Type = 'Potential Client'))" 
    #Get the full list of SPO Clients
    try{
        log-action -myMessage "Getting new Digest for https://anthesisllc.sharepoint.com/clients" -logFile $fullLogPathAndName
        $clientsDigest = new-spoDigest -serverUrl $webUrl -sitePath $sitePath -restCreds $restCreds -logFile $fullLogPathAndName -verboseLogging $true
        if($clientsDigest){log-result -myMessage "SUCCESS: New digest expires at $($clientsDigest.expiryTime)" -logFile $fullLogPathAndName}
        else{log-result -myMessage "FAILED: Unable to retrieve digest" -logFile $fullLogPathAndName}
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving digest for $webUrl$sitePath" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
    
    Connect-PnPOnline –Url $($webUrl+$sitePath) –Credentials $adminCreds
    $clientList = Get-PnPList -Identity "Kimble Clients" -Includes ContentTypes
    $clientListContentType = $clientList.ContentTypes | ? {$_.Name -eq "Item"}
    $clientListItems2 = Get-PnPListItem -List "Kimble Clients" -PageSize 1000 -Fields "Title","GUID","KimbleId","ClientDescription","IsDirty","IsDeleted","Modified","LastModifiedDate","PreviousName","PreviousDescription","Id"
    $clientListItems2.FieldValues | %{
        $thisClient = $_
        [array]$allSpoClients += New-Object psobject -Property $([ordered]@{"Id"=$thisClient["KimbleId"];"Name"=$thisClient["Title"];"GUID"=$thisClient["GUID"];"SPListItemID"=$thisClient["ID"];"IsDirty"=$thisClient["IsDirty"];"IsDeleted"=$thisClient["IsDeleted"];"LastModifiedDate"=$thisClient["LastModifiedDate"];"PreviousName"=$thisClient["PreviousName"];"PreviousDescription"=$thisClient["PreviousDescription"]})
        }

    $missingSpoClients = Compare-Object -ReferenceObject $allKimbleClients -DifferenceObject $allSpoClients -Property "Id" -PassThru -CaseSensitive:$false
    
    $missingSpoClients | ?{$_.SideIndicator -eq "<="} | %{
        $missingClient = $_
        $newItem = Add-PnPListItem -List $clientList.Id -ContentType $clientListContentType.Id.StringValue -Values @{"Title"=$missingClient.Name;"KimbleId"=$missingClient.Id;"ClientDescription"=$missingClient.Description;"IsDirty"=$true;"IsDeleted"=$missingClient.IsDeleted;"LastModifiedDate"=$(Get-Date $missingClient.LastModifiedDate -Format "MM/dd/yyyy hh:mm")}
        }

    $updatedKimbleClients = Compare-Object -ReferenceObject $allKimbleClients -DifferenceObject $allSpoClients -Property @("Id","LastModifiedDate") -PassThru -CaseSensitive:$false
    $updatedKimbleClients | ?{$_.SideIndicator -eq "<="} | % {
        $updatedClient = $_
        $spoClient = $allSpoClients | ? {$_.Id -eq $updatedClient.Id}
        $updatedValues = @{"IsDeleted"=$updatedClient.IsDeleted}
        if($updatedClient.LastModifiedDate -ne $null){
            $updatedValues.Add("LastModifiedDate",$(Get-Date $updatedClient.LastModifiedDate -Format "yyyy/MM/dd hh:mm:ss"))
            }
        if($updatedClient.Name -ne $spoClient.Name){
            $updatedValues.Add("Title",$updatedClient.Name)
            $updatedValues.Add("PreviousName",$spoClient.Name)
            $updatedValues.Add("IsDirty",$true)
            #Write-Host "$($spoClient.Name) renamed to $($updatedClient.Name)"
            $testName = $updatedValues
            }
        if((sanitise-stripHtml $updatedClient.Description) -ne $(sanitise-stripHtml $spoClient.Description)){
            $updatedValues.Add("ClientDescription",$(sanitise-stripHtml $updatedClient.Description))
            $updatedValues.Add("PreviousDescription", $spoClient.Description)
            #Write-Host "$($spoClient.Name) description change from $($spoClient.Description) to $($updatedClient.des)"
            if($updatedValues.Keys -notcontains "IsDirty"){$($updatedValues.Add("IsDirty",$true))}
            $testDesc = $updatedValues
            }
        Set-PnPListItem -List $clientList.Id -Identity $spoClient.SPListItemID -Values $updatedValues
        #Re-run for borked ClientDescription field
        <#if($updatedValues.Keys -contains "ClientDescription"){
            Write-Host -ForegroundColor Yellow "Updating $($spoClient.Name)"
            Set-PnPListItem -List $clientList.Id -Identity $spoClient.SPListItemID -Values $updatedValues
            }#>
        }
    }
function reconcile-projectsBetweenKimbleAndSpo(){
    #Get the full list of Kimble Projects
    $allKimbleProjects = get-allKimbleProjects -pQueryUri $standardKimbleQueryUri -pRestHeaders $standardKimbleHeaders
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
    


#endregion


##################################
#
#Do Stuff
#
##################################

reconcile-clients
reconcile-leads
reconcile-projects