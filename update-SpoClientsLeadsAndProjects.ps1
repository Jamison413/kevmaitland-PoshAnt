﻿$logFileLocation = "C:\ScriptLogs\"
$transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_Transcript_$(Get-Date -Format "yyMMdd").log"
if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)){
    $fullLogPathAndName = $logFileLocation+"update-SpoClientsLeadsAndProjects_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = $logFileLocation+"update-SpoClientsLeadsAndProjects_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
else{
    $fullLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
$debugLog = "$env:USERPROFILE\Desktop\debugdump.log"
Start-Transcript $transcriptLogName -Append

Import-Module _PS_Library_GeneralFunctionality
Import-Module _CSOM_Library-SPO
Import-Module _REST_Library-SPO


$webUrl = "https://anthesisllc.sharepoint.com"
$clientSite = "/clients"
$listOfClientFolders = @("_Kimble automatically creates Lead & Project folders","Background","Non-specific BusDev")
$listOfLeadProjSubFolders = @("Admin & contracts", "Analysis","Data & refs","Meetings","Proposal","Reports","Summary (marketing) - end of project")
$smtpServer = "anthesisgroup-com.mail.protection.outlook.com"
$mailFrom = "scriptrobot@sustain.co.uk"
$mailTo = "kevin.maitland@anthesisgroup.com"

$sharePointAdmin = "kimblebot@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
$sharePointAdminPass = ConvertTo-SecureString "01000000d08c9ddf0115d1118c7a00c04fc297eb01000000392cb8f8735d884c82c0932b5782960b0000000002000000000003660000c0000000100000001106ea74b4a38baa299968a1a66276830000000004800000a0000000100000000e9cfebd739622b6a7a5ab5dc7ea090120000000aee0b1e143e4f5bd5b18e0e5a6aefe9114e83a20069acb2cba2342cce5cca27c14000000e3535184af5408c2425bbec440f59aff355f69e1"
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
$restCreds = new-spoCred -Credential -username $adminCreds.UserName -securePassword $adminCreds.Password
$csomCreds = set-csomCredentials -username $adminCreds.UserName -password $adminCreds.Password

$cacheFilePath = "$env:USERPROFILE\KimbleCache\"
$clientsCacheFile = "kimbleClients.csv"
$projectsCacheFile = "kimbleProjects.csv"
$leadsCacheFile = "kimbleLeads.csv"

#region functions
function cache-kimbleClients(){
    try{
        log-action -myMessage "Getting [Kimble Clients] to check whether it needs recaching" -logFile $fullLogPathAndName 
        $kc = get-list -serverUrl $webUrl -sitePath $clientSite -listName "Kimble Clients" -restCreds $restCreds 
        if($kc){log-result -myMessage "SUCCESS: List retrieved" -logFile $fullLogPathAndName}
        else{log-result -myMessage "FAILURE: List could not be retrieved" -logFile $fullLogPathAndName}
        }
    catch{log-error -myError $_ -myFriendlyMessage "Could not retrieve [Kimble Clients] to check whether it needs recaching" -fullLogFile $fullLogPathAndName -errorLogFile -doNotLogToEmail $true}
    $kcCacheFile = Get-Item $cacheFilePath$clientsCacheFile
    if((get-date $kc.LastItemModifiedDate).AddMinutes(-5) -gt $kcCacheFile.LastWriteTimeUtc){#This is bodged so we don't miss any new clients added during the time it takes to actually download the full client list
        try{
            log-action -myMessage "[Kimble Clients] needs recaching - downloading full list" -logFile $fullLogPathAndName 
            $spClients = get-itemsInList -sitePath $clientSite -listName "Kimble Clients" -serverUrl $webUrl -restCreds $restCreds
            if($spClients){
                log-result -myMessage "SUCCESS: $($spClients.Count) Kimble Client records retrieved!" -logFile $fullLogPathAndName
                $spClients | Export-Csv -Path $cacheFilePath$clientsCacheFile -Force -NoTypeInformation -Encoding UTF8
                }
            else{log-result -myMessage "FAILURE: [Kimble Clients] items could not be retrieved" -logFile $fullLogPathAndName}
            }
        catch{log-error -myError $_ -myFriendlyMessage "Could not retrieve [Kimble Clients] items to recache the local copy" -fullLogFile $fullLogPathAndNamel -errorLogFile -doNotLogToEmail $true}
        }
    else{log-result -myMessage "SUCCESS: Cache is up-to-date and does not require refreshing" -logFile $fullLogPathAndName}
    $clientCache = Import-Csv $cacheFilePath$clientsCacheFile
    $clientCache
    }
#endregion

#Retrieve (and update if necessary) the full Clients cache as we'll need it to set up any new Leads/Projects
$clientCache = cache-kimbleClients
#Build a hashtable so we can look up Client name by it's KimbleId
$kimbleClientHashTable = @{}
foreach ($spClient in $clientCache){$kimbleClientHashTable.Add($spClient.KimbleId,$(sanitise-forSharePointListName $spClient.Title))}


#region Create folders for any new Clients
$dirtyClients = $clientCache | ?{$_.IsDirty -eq $true}
$clientDigest = new-spoDigest -serverUrl $webUrl -sitePath $clientSite -restCreds $restCreds

foreach($dirtyClient in $dirtyClients){
    log-action -myMessage "************************************************************************" -logFile $fullLogPathAndName
    log-action -myMessage "CLIENT [$($dirtyClient.Title)] needs updating!" -logFile $fullLogPathAndName
    #Check if the Client needs creating
    if((!$dirtyClient.PreviousName -and !$dirtyClient.PreviousDescription) -OR $recreateAllFolders -eq $true){
        #Create a new Library and subfolders
        try{
            log-action "new-library /$($dirtyClient.Title) Description: $((sanitise-stripHtml $dirtyClient.ClientDescription).SubString(0,20))" -logFile $fullLogPathAndName
            $newLibrary = new-library -serverUrl $webUrl -sitePath $clientSite -libraryName $dirtyClient.Title -libraryDesc $dirtyClient.ClientDescription -restCreds $restCreds -digest $clientDigest
            if($newLibrary){#If the new Library has been created, make the subfolders and update the List Item
                log-result "SUCCESS: $($dirtyClient.Title) is there!" -logFile $fullLogPathAndName
                #Try to create the subfolders
                foreach($sub in $listOfClientFolders){ 
                    try{
                        log-action "new-FolderInLibrary /$($dirtyClient.Title)/$sub" -logFile $fullLogPathAndName
                        $newFolder = new-FolderInLibrary -serverUrl $webUrl -site $clientSite -libraryName ("/"+$dirtyClient.Title) -folderPathAndOrName $sub 
                        }
                    catch{log-error $_ -myFriendlyMessage "Failed to create new subfolder $($dirtyClient.Title)/$sub" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
                    if ($newFolder){log-result "SUCCESS: $($dirtyClient.Title)\$sub created!" -logFile $fullLogPathAndName}
                    else{log-result "FAILURE: $($dirtyClient.Title)\$sub was not created/retrievable!" -logFile $fullLogPathAndName}
                    }
                #If we've got this far, try to update the IsDirty property on the Client in [Kimble Clients]
                try{
                    log-action "update-itemInList Kimble Clients | $($dirtyClient.Title) [$($dirtyClient.Id) @{IsDirty=$false}]" -logFile $fullLogPathAndName
                    update-itemInList -serverUrl $webUrl -sitePath $clientSite -listName "Kimble Clients" -predeterminedItemType $dirtyClient.__metadata.type -itemId $dirtyClient.Id -hashTableOfItemData @{IsDirty=$false} -restCreds $restCreds -digest $clientDigest | Out-Null
                    try{
                        $updatedItem = get-itemsInList -serverUrl $webUrl -sitePath $clientSite -listName "Kimble Clients" -oDataQuery "?`$filter=Id eq $($dirtyClient.Id)" -restCreds $restCreds
                        if($updatedItem.IsDirty -eq $false){log-result "SUCCESS: $($dirtyClient.Title) updated!" -logFile $fullLogPathAndName}
                        else{log-result "FAILED: Could not set Client [$($dirtyClient.Title)].IsDirty = `$true " -logFile $fullLogPathAndName}
                        }
                    catch{log-error -myError $_ -myFriendlyMessage "Failed to set IsDirty=`$true for Client [$($dirtyClient.Title)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
                    }
                catch{log-error $_ -myFriendlyMessage "Failed to update $($dirtyClient.Title) in [Kimble Clients] List - this will stay as IsDirty=true forever :(" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
                }
            else{log-result "FAILURE: $($dirtyClient.Title) was not created/retrievable!" -logFile $fullLogPathAndName}
            }
        catch{log-error $_ -myFriendlyMessage "Failed to create new Library for $($dirtyClient.Title)" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -smtpServer $smtpServer -mailTo $mailTo -mailFrom $mailFrom}
        #Now try to add the new ClientName to the TermStore
        try{
            log-action "add-termToStore: Kimble | Clients | $($dirtyClient.Title)" -logFile $fullLogPathAndName
            add-termToStore -pGroup "Kimble" -pSet "Clients" -pTerm $($dirtyClient.Title) -credentials $csomCreds -webUrl $webUrl -siteCollection "/"
            log-result "SUCCESS: $($dirtyClient.Title) (probably) added to Managed MetaData Term Store" -logFile $fullLogPathAndName
            }
        catch{log-error $_ -myFriendlyMessage "Failed to add $($dirtyClient.Title) to Term Store" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -smtpServer $smtpServer -mailFrom $mailFrom -mailTo $mailTo}
        
        }
    #Check if the Client Name needs updating
    elseif(!([string]::IsNullOrEmpty($dirtyClient.PreviousName)) -and ($dirtyClient.PreviousName -ne $dirtyClient.Title)){
        #Update the folder name
        try{
            log-action "update-list $($dirtyClient.PreviousName) > @{Title=$($dirtyClient.Title)}" -logFile $fullLogPathAndName
            update-list -serverUrl $webUrl -sitePath $clientSite -listName $dirtyClient.PreviousName -hashTableOfUpdateData @{Title=$dirtyClient.Title} -restCreds $restCreds -digest $clientDigest | Out-Null
            #Update the Client in [Kimble Clients]
            try{
                if((get-list -sitePath $clientSite -listName $dirtyClient.Title -serverUrl $webUrl -restCreds $restCreds) -ne $false){ #If it's worked, set the IsDirty flag to $false to prevent it reprocessing
                    log-result "SUCCESS: $($dirtyClient.PreviousName) updated to $($dirtyClient.Title)" -logFile $fullLogPathAndName
                    log-action "update-itemInList Kimble Clients | $($dirtyClient.Title) ($($dirtyClient.Id) @{IsDirty=$false})" -logFile $fullLogPathAndName
                    update-itemInList -serverUrl $webUrl -sitePath $clientSite -listName "Kimble Clients" -predeterminedItemType ($dirtyClient.__metadata -split "; " | ?{$_.Substring(0,5) -imatch "type="}).Replace("type=","").Replace("@{","").Replace("}","") -itemId $dirtyClient.Id -hashTableOfItemData @{IsDirty=$false} -restCreds $restCreds -digest $clientDigest | Out-Null
                    try{
                        $updatedItem = get-itemsInList -serverUrl $webUrl -sitePath $clientSite -listName "Kimble Clients" -oDataQuery "?`$filter=Id eq $($dirtyClient.Id)" -restCreds $restCreds
                        if($updatedItem.IsDirty -eq $false){log-result "SUCCESS: $($dirtyClient.Title) updated!" -logFile $fullLogPathAndName}
                        else{log-result "FAILED: Could not set Client [$($dirtyClient.Title)].IsDirty = `$true " -logFile $fullLogPathAndName}
                        }
                    catch{log-error -myError $_ -myFriendlyMessage "Failed to set IsDirty=`$true for Client [$($dirtyClient.Title)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
                    }
                else{log-result -myMessage "FAILURE: I couldn't retrieve the Library I (allegedly) just updated: [$($dirtyClient.Title)] this will stay as IsDirty=true forever :(" -logFile $fullLogPathAndName}
                }
            catch{log-error $_ -myFriendlyMessage "Failed to update $($dirtyClient.Title) in [Kimble Clients] List - this will stay as IsDirty=true forever :(" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
            }
        catch{log-error $_ -myFriendlyMessage "Failed to update Library Title $($dirtyClient.PreviousName) to $($dirtyClient.Title)" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
        #Update the Managed MetaData in the TermStore
        try{
            log-action "rename-termInStore Kimble | Clients | $($dirtyClient.PreviousName) > $($dirtyClient.Title)" -logFile $fullLogPathAndName
            rename-termInStore -pGroup "Kimble" -pSet "Clients" -pOldTerm $($dirtyClient.PreviousName) -pNewTerm $($dirtyClient.Title) -credentials $csomCreds -webUrl $webUrl -siteCollection "/"
            log-result "SUCCESS: Term $($dirtyClient.PreviousName) renamed to $($dirtyClient.Title)" -logFile $fullLogPathAndName
            }
        catch{log-error $_ -myFriendlyMessage "Failed to rename ManagedMetadata term $($dirtyClient.PreviousName) to $($dirtyClient.Title)" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
        }
    #Check if the Client Description needs updating
    elseif(((sanitise-stripHtml $dirtyClient.PreviousDescription) -ne (sanitise-stripHtml $dirtyClient.ClientDescription)) -or ((sanitise-stripHtml $dirtyClient.ClientDescription) -ne ($dirtyClient.ClientDescription))){
        #Update the Library's Description
        try{
            log-action -myMessage "update-list [$($dirtyClient.Title)].Description `"$($dirtyClient.PreviousDescription.Substring(0,20))...`" > `"$((sanitise-stripHtml $dirtyClient.ClientDescription).Substring(0,20))...`"" -logFile $fullLogPathAndName
            update-list -serverUrl $webUrl -sitePath $clientSite -restCreds $restCreds -digest $clientDigest -listName $dirtyClient.Title -hashTableOfUpdateData @{Description=$(sanitise-stripHtml $dirtyClient.ClientDescription)} 
            #If it's worked, update the IsDirty property on the Client
            if($(get-list -serverUrl $webUrl -sitePath $clientSite -listName $dirtyClient.Title -restCreds $restCreds).Description -eq $(sanitise-stripHtml $dirtyClient.ClientDescription)){ #If it's worked, set the IsDirty flag to $false to prevent it reprocessing
                log-result -myMessage "SUCCESS: [$($dirtyClient.Title)].Description updated to `"$((sanitise-stripHtml $dirtyClient.ClientDescription).Substring(0,20))...`"" -logFile $fullLogPathAndName
                try{
                    log-action "update-itemInList Kimble Clients | $($dirtyClient.Title) ($($dirtyClient.Id) @{IsDirty=$false})" -logFile $fullLogPathAndName
                    update-itemInList -serverUrl $webUrl -sitePath $clientSite -restCreds $restCreds -digest $clientDigest -listName "Kimble Clients" -predeterminedItemType $dirtyClient.__metadata.type -itemId $dirtyClient.Id -hashTableOfItemData @{IsDirty=$false}  | Out-Null
                    try{
                        $updatedItem = get-itemsInList -serverUrl $webUrl -sitePath $clientSite -listName "Kimble Clients" -oDataQuery "?`$filter=Id eq $($dirtyClient.Id)" -restCreds $restCreds
                        if($updatedItem.IsDirty -eq $false){log-result "SUCCESS: $($dirtyClient.Title) updated!" -logFile $fullLogPathAndName}
                        else{log-result "FAILED: Could not set Client [$($dirtyClient.Title)].IsDirty = `$true " -logFile $fullLogPathAndName}
                        }
                    catch{log-error -myError $_ -myFriendlyMessage "Failed to set IsDirty=`$true for Client [$($dirtyClient.Title)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
                    }
                catch{log-error $_ -myFriendlyMessage "Failed to update $($dirtyClient.Title) in [Kimble Clients] List - this will stay as IsDirty=true forever :(" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
                }
            else{log-result -myMessage "FAILURE: I couldn't retrieve the Library I (allegedly) just created: [$($dirtyClient.Title)] this will stay as IsDirty=true forever :(" -logFile $fullLogPathAndName}
            }
        catch{log-error $_ -myFriendlyMessage "Failed to update Library Description for $($dirtyClient.Title)" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
        }
    #Otherwise, the Client is flagged IsDirty, but it's not going to get processed
    else{log-action -myMessage "WARNING: CLIENT [$($dirtyClient.Title)] IsDirty, but I can't work out why :/" -logFile $fullLogPathAndName}
    }
#endregion

#region Create folders for any new Leads
#Get the items in [Kimble Leads] that need processing
try{
    log-action -myMessage "Retrieving [Kimble Leads] flagged IsDirty" -logFile $fullLogPathAndName
    $dirtyLeads = get-itemsInList -serverUrl $webUrl -sitePath $clientSite -listName "Kimble Leads" -restCreds $restCreds -oDataQuery "?&`$filter=IsDirty eq 1" 
    if($dirtyLeads.Count -gt 0){log-result -myMessage "SUCCESS: $($dirtyLeads.Count) [Kimble Leads] items need processing" -logFile $fullLogPathAndName}
    elseif([string]::IsNullOrEmpty($dirtyLeads)){log-result -myMessage "SUCCESS: [Kimble Leads] retrieved successfully, but no records need processing!" -logFile $fullLogPathAndName}
    else{log-result -myMessage "FAILED: Unable to retrieve [Kimble Leads] items" -logFile $fullLogPathAndName}
    }
catch{log-error -myError $_ -myFriendlyMessage "Error retrieving [Kimble Leads] items flagged as IsDirty" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}

#Process any [Kimble Leads] flagged as IsDirty
foreach($dirtyLead in $dirtyLeads){
    log-action -myMessage "************************************************************************" -logFile $fullLogPathAndName
    log-action -myMessage "LEAD [$($dirtyLead.Title)] for client [$($kimbleClientHashTable[$dirtyLead.KimbleClientId])] needs processing!" -logFile $fullLogPathAndName
    $leadFolderName = "BD_"+$dirtyLead.Title
    #Check if the Lead needs creating
    if(!$dirtyLead.PreviousName -and (!$dirtyLead.PreviousKimbleClientId -or $dirtyLead.PreviousKimbleClientId -eq $dirtyLead.KimbleClientId)){
        #Create a new folder tree under the Client Library
        log-action -myMessage "LEAD [$($dirtyLead.Title)] for client [$($kimbleClientHashTable[$dirtyLead.KimbleClientId])] needs creating!" -logFile $fullLogPathAndName
        try{
            log-action "new-folderInLibrary $("/"+$kimbleClientHashTable[$dirtyLead.KimbleClientId])/$leadFolderName" -logFile $fullLogPathAndName
            #Check that the corresponding Client Name can be looked up from the [Kimble Clients] cache (we need this to know which Client Library to add the Lead folders into)
            if ($kimbleClientHashTable[$dirtyLead.KimbleClientId] -eq $null){
                log-result "FAILURE: Client with Id [$($dirtyLead.KimbleClientId)]could not be found in [Kimble Clients]"
                #Bodge this with an e-mail alert until we can automatically update the Client in Kimble
                Send-MailMessage -SmtpServer $smtpServer -To $mailTo -From $mailFrom -Subject "Client with ID $($dirtyLead.KimbleClientId) is not a Kimble Client" -Body "Lead: $($dirtyLead.Title)"
                continue
                }
            #Check that the Client Library is retrievable
            try{
                log-action -myMessage "Trying to create Folder: $clientSite/$($kimbleClientHashTable[$dirtyLead.KimbleClientId])/$leadFolderName" -logFile $fullLogPathAndName
                $newLeadLibraryFolder = new-folderInLibrary -serverUrl $webUrl -sitePath $clientSite -libraryName ("/"+ $kimbleClientHashTable[$dirtyLead.KimbleClientId]) -folderPathAndOrName $leadFolderName -restCreds $restCreds -digest $clientDigest #-logFile Out-Null -verboseLogging $true
                #If the new Folder has been created, make the subfolders and update the List Item
                if($newLeadLibraryFolder.__metadata){
                    #Create the subfolders
                    log-result "SUCCESS: $("/"+$kimbleClientHashTable[$dirtyLead.KimbleClientId])/$leadFolderName is retrievable!" -logFile $fullLogPathAndName
                    foreach($sub in $listOfLeadProjSubFolders){
                        try{
                            log-action "new-folderInLibrary $("/"+$kimbleClientHashTable[$dirtyLead.KimbleClientId])/$leadFolderName/$sub" -logFile $fullLogPathAndName
                            $newLeadLibrarySubfolder = new-FolderInLibrary -serverUrl $webUrl -site $clientSite -libraryName ("/"+$kimbleClientHashTable[$dirtyLead.KimbleClientId]) -folderPathAndOrName ("/"+$leadFolderName.Replace("/","")+"/"+$sub) -restCreds $restCreds -digest $clientDigest
                            }
                        catch{log-error -myError $_ -myFriendlyMessage "Failed to create the subfolder $("/"+$kimbleClientHashTable[$dirtyLead.KimbleClientId])/$leadFolderName/$sub"}
                        #Validate that new-FolderInLibrary returned *something* (we're not validating that each subfolder gets created - if the main Lead folder created correctly, we'll just assume that they all will)
                        if($newLeadLibrarySubfolder.__metadata){log-result "SUCCESS: $($kimbleClientHashTable[$dirtyLead.KimbleClientId]+"\"+$leadFolderName)\$sub is retrievable!" -logFile $fullLogPathAndName}
                        else{log-result "FAILURE: SubFolder $sub was not created/retrievable" -logFile $fullLogPathAndName}
                        }
                    #If we've got this far, try to update the IsDirty property on the Lead
                    try{
                        log-action "update-itemInList Kimble Leads | $($dirtyLead.Title) [$($dirtyLead.Id) @{IsDirty=$false}]" -logFile $fullLogPathAndName
                        update-itemInList -serverUrl $webUrl -sitePath $clientSite -listName "Kimble Leads" -predeterminedItemType $dirtyLead.__metadata.type -itemId $dirtyLead.Id -hashTableOfItemData @{IsDirty=$false} -restCreds $restCreds -digest $clientDigest | Out-Null
                        #Validate that the change was actually made
                        try{
                            $updatedItem = get-itemsInList -serverUrl $webUrl -sitePath $clientSite -listName "Kimble Leads" -oDataQuery "?`$filter=Id eq $($dirtyLead.Id)" -restCreds $restCreds
                            if($updatedItem.IsDirty -eq $false){log-result "SUCCESS: $($dirtyLead.Title) updated!" -logFile $fullLogPathAndName}
                            else{log-result "FAILED: Could not set Lead [$($dirtyLead.Title)].IsDirty = `$true " -logFile $fullLogPathAndName}
                            }
                        catch{log-error -myError $_ -myFriendlyMessage "Failed to set IsDirty=`$true for Lead [$($dirtyLead.Title)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
                        }
                    catch{log-error $_ -myFriendlyMessage "Failed to update $($dirtyLead.Title) in [Kimble Leads] List - this will stay as IsDirty=true forever :(" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
                    }
                else{log-result -myMessage "FAILURE: Folder could not be created/retrieved" -logFile $fullLogPathAndName}
                }
            catch{log-error -myError $_ -myFriendlyMessage "Error creating Lead Folder $clientSite/$($kimbleClientHashTable[$dirtyLead.KimbleClientId])/$leadFolderName" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
            }
        catch{log-error $_ -myFriendlyMessage "Failed to create new Folder: $("/"+$kimbleClientHashTable[$dirtyLead.KimbleClientId])/$leadFolderName" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
        }
    #Check if the Lead needs updating
    elseif(($dirtyLead.PreviousName) -and ($dirtyLead.PreviousName -ne $dirtyLead.Title)){
        log-action -myMessage "LEAD [$($dirtyLead.Title)] for client [$($kimbleClientHashTable[$dirtyLead.KimbleClientId])] needs renaming!" -logFile $fullLogPathAndName
        try{
            #Try to get the folder first as we'll need its Id to update it
            log-action -myMessage "get-folderInLibrary /$($kimbleClientHashTable[$dirtyLead.KimbleClientId])/BD_$($dirtyLead.PreviousName)" -logFile $fullLogPathAndName
            $clientLibraryFolder = get-folderInLibrary -serverUrl $webUrl -sitePath $clientSite -libraryName $($kimbleClientHashTable[$dirtyLead.KimbleClientId]) -folderPathAndOrName "/BD_$($dirtyLead.PreviousName)" -restCreds $restCreds
            if($clientLibraryFolder.__metadata){
                try{
                    log-action -myMessage "update-itemInList [$($kimbleClientHashTable[$dirtyLead.KimbleClientId])] | $($dirtyLead.PreviousName) > @{Title=$leadFolderName;FileLeafRef=$leadFolderName}" -logFile $fullLogPathAndName
                    update-itemInList -serverUrl $webUrl -sitePath $clientSite -listName $($kimbleClientHashTable[$dirtyLead.KimbleClientId]) -predeterminedItemType $clientLibraryFolder.__metadata.type -itemId $clientLibraryFolder.Id -hashTableOfItemData @{Title=$leadFolderName;FileLeafRef=$leadFolderName} -restCreds $restCreds -digest $clientDigest #| Out-Null
                    try{
                        $updatedItem = get-folderInLibrary -serverUrl $webUrl -sitePath $clientSite -libraryName $($kimbleClientHashTable[$dirtyLead.KimbleClientId]) -folderPathAndOrName "/$leadFolderName" -restCreds $restCreds
                        if($updatedItem.__metadata){log-result "SUCCESS: $($dirtyLead.PreviousName) updated!" -logFile $fullLogPathAndName}
                        else{log-result "FAILED: Could not retrieve folder for Lead $($kimbleClientHashTable[$dirtyLead.KimbleClientId])/$leadFolderName - rename failed" -logFile $fullLogPathAndName}
                        }
                    catch{log-error -myError $_ -myFriendlyMessage "Failed to rename folder for Lead [$($dirtyLead.PreviousName)] to [$leadFolderName]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
                    }
                catch{log-error -myError $_ -myFriendlyMessage "Failed to rename Lead folder /$($kimbleClientHashTable[$dirtyLead.KimbleClientId])/BD_$($dirtyLead.PreviousName) > $leadFolderName" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
            else{log-result -myMessage "FAILED: Could not retrieve folder: /$($kimbleClientHashTable[$dirtyLead.KimbleClientId])/BD_$($dirtyLead.PreviousName) (so cannot rname it)" -logFile $fullLogPathAndName}
            }
            if((get-folderInLibrary -sitePath $clientSite -libraryName $kimbleClientHashTable[$dirtyLead.KimbleClientId] -folderName $leadFolderName) -ne $false){ #If it's worked, set the IsDirty flag to $false to prevent it reprocessing
                update-itemInList -serverUrl $webUrl -sitePath $clientSite -listName "Kimble Leads" -predeterminedItemType $dirtyLead.__metadata.type -itemId $dirtyLead.Id -hashTableOfItemData @{IsDirty=$false}
                }
            else{log-result -myMessage "Uh-oh, I couldn't find the Library I (allegedly) just updated: [$($dirtyLead.Title)] this will stay as IsDirty=true forever :(" -logFile $fullLogPathAndName}
            }
        catch{log-error $_ -myFriendlyMessage "Failed to update Library Title $($dirtyLead.PreviousName) to $($dirtyLead.Title)" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
        }
    #Check if the Lead needs moving to another Client
    elseif($dirtyLead.PreviousKimbleClientId -ne $dirtyLead.KimbleClientId){
        #Move the folder to the new Client
        log-action -myMessage "LEAD [$($dirtyLead.Title)] for client [$($kimbleClientHashTable[$dirtyLead.KimbleClientId])] needs moving from Client [$($kimbleClientHashTable[$dirtyLead.PreviousKimbleClientId])] to [$($kimbleClientHashTable[$dirtyLead.KimbleClientId])]!" -logFile $fullLogPathAndName
        try{
            #Yeah Kev, you actually need to write some code to *do* this. Move $kimbleClientHashTable[$dirtyLead.PreviousKimbleClientId]/$dirtyLead.Title to $kimbleClientHashTable[$dirtyLead.KimbleClientId]
            }
        catch{log-error $_ -myFriendlyMessage "Failed to move Lead from Client X to Client Y" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
        }
    #Otherwise, the Lead is flagged IsDirty, but it's not going to get processed
    else{log-action -myMessage "WARNING: LEAD [$($dirtyLead.Title)] for client [$($kimbleClientHashTable[$dirtyLead.KimbleClientId])] IsDirty, but I can't work out why :/" -logFile $fullLogPathAndName}
    }

#endregion

#region Create folders for any new Projects
#Get the items in [Kimble Projects] that need processing
try{
    log-action -myMessage "Retrieving [Kimble Projects] flagged IsDirty" -logFile $fullLogPathAndName
    $dirtyProjects = get-itemsInList -serverUrl $webUrl -sitePath $clientSite -listName "Kimble Projects" -restCreds $restCreds -oDataQuery "?&`$filter=IsDirty eq 1" 
    if($dirtyProjects.Count -gt 0){log-result -myMessage "SUCCESS: $($dirtyProjects.Count) [Kimble Projects] items need processing" -logFile $fullLogPathAndName}
    elseif([string]::IsNullOrEmpty($dirtyProjects)){log-result -myMessage "SUCCESS: [Kimble Projects] retrieved successfully, but no records need processing!" -logFile $fullLogPathAndName}
    else{log-result -myMessage "FAILED: Unable to retrieve [Kimble Projects] items" -logFile $fullLogPathAndName}
    }
catch{log-error -myError $_ -myFriendlyMessage "Error retrieving [Kimble Projects] items flagged as IsDirty" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}

#Process any [Kimble Projects] flagged as IsDirty
foreach($dirtyProject in $dirtyProjects){
    log-action -myMessage "************************************************************************" -logFile $fullLogPathAndName
    log-action -myMessage "PROJECT [$($dirtyProject.Title)] for client [$($kimbleClientHashTable[$dirtyProject.KimbleClientId])] needs updating!" -logFile $fullLogPathAndName
    $projectFolderName = $dirtyProject.Title

    #Check if the Project needs creating
    if(!$dirtyProject.PreviousName -and (!$dirtyProject.PreviousKimbleClientId -or $dirtyProject.PreviousKimbleClientId -eq $dirtyProject.KimbleClientId)){
        #Create a new folder tree under the Client Library
        log-action -myMessage "PROJECT [$($dirtyProject.Title)] for client [$($kimbleClientHashTable[$dirtyProject.KimbleClientId])] needs creating!" -logFile $fullLogPathAndName
        try{
            log-action "new-folderInLibrary $("/"+$kimbleClientHashTable[$dirtyProject.KimbleClientId])/$projectFolderName" -logFile $fullLogPathAndName
            #Check that the corresponding Client Name can be looked up from the [Kimble Clients] cache (we need this to know which Client Library to add the Project folders into)
            if ($kimbleClientHashTable[$dirtyProject.KimbleClientId] -eq $null){
                log-result "FAILURE: Client with Id [$($dirtyProject.KimbleClientId)]could not be found in [Kimble Clients]"
                #Bodge this with an e-mail alert until we can automatically update the Client in Kimble
                Send-MailMessage -SmtpServer $smtpServer -To $mailTo -From $mailFrom -Subject "Client with ID $($dirtyProject.KimbleClientId) is not a Kimble Client" -Body "Project: $($dirtyProject.Title)"
                continue
                }
            #Check that the Client Library is retrievable
            try{
                log-action -myMessage "Trying to create Folder: $clientSite/$($kimbleClientHashTable[$dirtyProject.KimbleClientId])/$projectFolderName" -logFile $fullLogPathAndName
                $newProjectLibraryFolder = new-folderInLibrary -serverUrl $webUrl -sitePath $clientSite -libraryName ("/"+ $kimbleClientHashTable[$dirtyProject.KimbleClientId]) -folderPathAndOrName $projectFolderName -restCreds $restCreds -digest $clientDigest -logFile Out-Null -verboseLogging $true
                #If the new Folder has been created, make the subfolders and update the List Item
                if($newProjectLibraryFolder.__metadata){
                    #Create the subfolders
                    log-result "SUCCESS: $("/"+$kimbleClientHashTable[$dirtyProject.KimbleClientId])/$projectFolderName is retrievable!" -logFile $fullLogPathAndName
                    foreach($sub in $listOfLeadProjSubFolders){
                        try{
                            log-action "new-folderInLibrary $("/"+$kimbleClientHashTable[$dirtyProject.KimbleClientId])/$projectFolderName/$sub" -logFile $fullLogPathAndName
                            $newProjectLibrarySubfolder = new-FolderInLibrary -serverUrl $webUrl -site $clientSite -libraryName ("/"+$kimbleClientHashTable[$dirtyProject.KimbleClientId]) -folderPathAndOrName ("/"+$projectFolderName.Replace("/","")+"/"+$sub) -restCreds $restCreds -digest $clientDigest
                            }
                        catch{log-error -myError $_ -myFriendlyMessage "Failed to create the subfolder $("/"+$kimbleClientHashTable[$dirtyProject.KimbleClientId])/$projectFolderName/$sub"}
                        #Validate that new-FolderInLibrary returned *something* (we're not validating that each subfolder gets created - if the main Project folder created correctly, we'll just assume that they all will)
                        if($newProjectLibrarySubfolder.__metadata){log-result "SUCCESS: $($kimbleClientHashTable[$dirtyProject.KimbleClientId]+"\"+$projectFolderName)\$sub is retrievable!" -logFile $fullLogPathAndName}
                        else{log-result "FAILURE: SubFolder $sub was not created/retrievable" -logFile $fullLogPathAndName}
                        }
                    #If we've got this far, try to update the IsDirty property on the Project
                    try{
                        log-action "update-itemInList Kimble Projects | $($dirtyProject.Title) [$($dirtyProject.Id) @{IsDirty=$false}]" -logFile $fullLogPathAndName
                        update-itemInList -serverUrl $webUrl -sitePath $clientSite -listName "Kimble Projects" -predeterminedItemType $dirtyProject.__metadata.type -itemId $dirtyProject.Id -hashTableOfItemData @{IsDirty=$false} -restCreds $restCreds -digest $clientDigest | Out-Null
                        #Validate that the change was actually made
                        try{
                            $updatedItem = get-itemsInList -serverUrl $webUrl -sitePath $clientSite -listName "Kimble Projects" -oDataQuery "?`$filter=Id eq $($dirtyProject.Id)" -restCreds $restCreds
                            if($updatedItem.IsDirty -eq $false){log-result "SUCCESS: $($dirtyProject.Title) updated!" -logFile $fullLogPathAndName}
                            else{log-result "FAILED: Could not set Project [$($dirtyProject.Title)].IsDirty = `$true " -logFile $fullLogPathAndName}
                            }
                        catch{log-error -myError $_ -myFriendlyMessage "Failed to set IsDirty=`$true for Project [$($dirtyProject.Title)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
                        }
                    catch{log-error $_ -myFriendlyMessage "Failed to update $($dirtyProject.Title) in [Kimble Projects] List - this will stay as IsDirty=true forever :(" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
                    }
                else{log-result -myMessage "FAILURE: Folder could not be created/retrieved" -logFile $fullLogPathAndName}
                }
            catch{log-error -myError $_ -myFriendlyMessage "Error creating Project Folder $clientSite/$($kimbleClientHashTable[$dirtyProject.KimbleClientId])/$projectFolderName" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
            }
        catch{log-error $_ -myFriendlyMessage "Failed to create new Folder: $("/"+$kimbleClientHashTable[$dirtyProject.KimbleClientId])/$($dirtyProject.Title)" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
        }
    #Check if the Project needs updating
    elseif(($dirtyProject.PreviousName) -and ($dirtyProject.PreviousName -ne $dirtyProject.Title)){
        #Update the folder name
        log-action -myMessage "PROJECT [$($dirtyProject.Title)] for client [$($kimbleClientHashTable[$dirtyProject.KimbleClientId])] needs renaming!" -logFile $fullLogPathAndName
        try{
            #Try to get the folder first as we'll need its Id to update it
            log-action -myMessage "get-folderInLibrary /$($kimbleClientHashTable[$dirtyProject.KimbleClientId])/$($dirtyProject.PreviousName)" -logFile $fullLogPathAndName
            $clientLibraryFolder = get-folderInLibrary -serverUrl $webUrl -sitePath $clientSite -libraryName $($kimbleClientHashTable[$dirtyProject.KimbleClientId]) -folderPathAndOrName "/$($dirtyProject.PreviousName)" -restCreds $restCreds
            if($clientLibraryFolder.__metadata){
                try{
                    log-action -myMessage "update-itemInList [$($kimbleClientHashTable[$dirtyProject.KimbleClientId])] | $($dirtyProject.PreviousName) > @{Title=$dirtyProject.Title;FileLeafRef=$dirtyProject.Title}" -logFile $fullLogPathAndName
                    update-itemInList -serverUrl $webUrl -sitePath $clientSite -listName $($kimbleClientHashTable[$dirtyProject.KimbleClientId]) -predeterminedItemType $clientLibraryFolder.__metadata.type -itemId $clientLibraryFolder.Id -hashTableOfItemData @{Title=$dirtyProject.Title;FileLeafRef=$dirtyProject.Title} -restCreds $restCreds -digest $clientDigest #| Out-Null
                    try{
                        $updatedItem = get-folderInLibrary -serverUrl $webUrl -sitePath $clientSite -libraryName $($kimbleClientHashTable[$dirtyProject.KimbleClientId]) -folderPathAndOrName "/$projectFolderName" -restCreds $restCreds
                        if($updatedItem.__metadata){log-result "SUCCESS: $($dirtyProject.PreviousName) updated!" -logFile $fullLogPathAndName}
                        else{log-result "FAILED: Could not retrieve folder for Project $($kimbleClientHashTable[$dirtyProject.KimbleClientId])/$projectFolderName - rename failed" -logFile $fullLogPathAndName}
                        }
                    catch{log-error -myError $_ -myFriendlyMessage "Failed to rename folder for Project [$($dirtyProject.PreviousName)] to [$projectFolderName]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
                    }
                catch{log-error -myError $_ -myFriendlyMessage "Failed to rename Project folder /$($kimbleClientHashTable[$dirtyProject.KimbleClientId])/$($dirtyProject.PreviousName) > $projectFolderName" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
            else{log-result -myMessage "FAILED: Could not retrieve folder: /$($kimbleClientHashTable[$dirtyProject.KimbleClientId])/$($dirtyProject.PreviousName) (so cannot rename it)" -logFile $fullLogPathAndName}
            }
            if((get-folderInLibrary -sitePath $clientSite -libraryName $kimbleClientHashTable[$dirtyProject.KimbleClientId] -folderName $dirtyProject.Title) -ne $false){ #If it's worked, set the IsDirty flag to $false to prevent it reprocessing
                update-itemInList -serverUrl $webUrl -sitePath $clientSite -listName "Kimble Projects" -predeterminedItemType $dirtyProject.__metadata.type -itemId $dirtyProject.Id -hashTableOfItemData @{IsDirty=$false}
                }
            else{log-result -myMessage "Uh-oh, I couldn't find the Library I (allegedly) just updated: [$($dirtyProject.Title)] this will stay as IsDirty=true forever :(" -logFile $fullLogPathAndName}
            }
        catch{log-error $_ -myFriendlyMessage "Failed to update Library Title $($dirtyProject.PreviousName) to $projectFolderName" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
        }
    #Check if the Project needs moving to another Client
    elseif($dirtyProject.PreviousKimbleClientId -ne $dirtyProject.KimbleClientId){
        #Move the folder to the new Client
        log-action -myMessage "PROJECT [$($dirtyProject.Title)] for client [$($kimbleClientHashTable[$dirtyProject.KimbleClientId])] needs moving from Client [$($kimbleClientHashTable[$dirtyProject.PreviousKimbleClientId])] to [$($kimbleClientHashTable[$dirtyProject.KimbleClientId])]!" -logFile $fullLogPathAndName
        try{
            #Yeah Kev, you actually need to write some code to *do* this. Move $kimbleClientHashTable[$dirtyProject.PreviousKimbleClientId]/$dirtyProject.Title to $kimbleClientHashTable[$dirtyProject.KimbleClientId]
            }
        catch{log-error $_ -myFriendlyMessage "Failed to move Project from Client X to Client Y" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
        }
    #Otherwise, the Project is flagged IsDirty, but it's not going to get processed
    else{log-action -myMessage "WARNING: PROJECT [$($dirtyProject.Title)] for client [$($kimbleClientHashTable[$dirtyProject.KimbleClientId])] IsDirty, but I can't work out why :/" -logFile $fullLogPathAndName}
    }
#endregion

<#
        try{
            $newProjectLibraryFolder = new-folderInLibrary -serverUrl $webUrl -sitePath $clientSite -libraryName ("/"+ $kimbleClientHashTable[$dirtyProject.KimbleClientId]) -folderPathAndOrName $dirtyProject.Title -restCreds $restCreds -digest $clientDigest
            if($newProjectLibraryFolder){#If the new Folder has been created, make the subfolders and update the List Item
                #Create the subfolders
                log-result "SUCCESS: $("/"+$kimbleClientHashTable[$dirtyProject.KimbleClientId])/$($dirtyProject.Title) is retrievable!"
                foreach($sub in $listOfLeadProjSubFolders){
                    try{
                        log-action "new-folderInLibrary $("/"+$kimbleClientHashTable[$dirtyProject.KimbleClientId])/$($dirtyProject.Title)/$sub" -logFile $fullLogPathAndName
                        $newProjectLibrarySubfolder = new-FolderInLibrary -serverUrl $webUrl -site $clientSite -libraryName ("/"+$kimbleClientHashTable[$dirtyProject.KimbleClientId]) -folderPathAndOrName ("/"+$dirtyProject.Title.Replace("/","")+"/"+$sub) -restCreds $restCreds -digest $clientDigest
                        }
                    catch{log-error -myError $_ -myFriendlyMessage "Failed to create the subfolder $("/"+$kimbleClientHashTable[$dirtyProject.KimbleClientId])/$dirtyProject.Title/$sub"}
                    if($newProjectLibrarySubfolder){log-result "SUCCESS: $($kimbleClientHashTable[$dirtyProject.KimbleClientId]+"\"+$dirtyProject.Title)\$sub is retrievable!"}
                    else{log-result "FAILURE: SubFolder $sub was not created/retrievable" -logFile $fullLogPathAndName}
                    }
                #If we've got this far, try to update the IsDirty property on the Project
                try{
                    log-action "update-itemInList Kimble Projects | $($dirtyProject.Title) [$($dirtyProject.Id) @{IsDirty=$false}]" -logFile $fullLogPathAndName
                    update-itemInList -serverUrl $webUrl -sitePath $clientSite -listName "Kimble Projects" -predeterminedItemType $dirtyProject.__metadata.type -itemId $dirtyProject.Id -hashTableOfItemData @{IsDirty=$false} -restCreds $restCreds -digest $clientDigest | Out-Null
                    try{
                        $updatedItem = get-itemsInList -serverUrl $webUrl -sitePath $clientSite -listName "Kimble Projects" -oDataQuery "?`$filter=Id eq $($dirtyProject.Id)" -restCreds $restCreds
                        if($updatedItem.IsDirty -eq $false){log-result "SUCCESS: $($dirtyProject.Title) updated!" -logFile $fullLogPathAndName}
                        else{log-result "FAILED: Could not set Project [$($dirtyProject.Title)].IsDirty = `$true " -logFile $fullLogPathAndName}
                        }
                    catch{log-error -myError $_ -myFriendlyMessage "Failed to set IsDirty=`$true for Project [$($dirtyProject.Title)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
                    }
                catch{log-error $_ -myFriendlyMessage "Failed to update $($dirtyProject.Title) in [Kimble Projects] List - this will stay as IsDirty=true forever :(" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
                }
            else{log-result "FAILURE: Folder $("/"+$kimbleClientHashTable[$dirtyProject.KimbleClientId])/$($dirtyProject.Title) was not created/retrievable" -logFile $fullLogPathAndName}
            }

#>