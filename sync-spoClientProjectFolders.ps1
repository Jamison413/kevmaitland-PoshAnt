$logFileLocation = "C:\ScriptLogs\"
$transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_Transcript_$(Get-Date -Format "yyMMdd").log"
if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)){
    $fullLogPathAndName = $logFileLocation+"sync-spoClientsProjectsFolders_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = $logFileLocation+"sync-spoClientsProjectsFolders_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
else{
    $fullLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
$debugLog = "$env:USERPROFILE\Desktop\debugdump.log"
Start-Transcript $transcriptLogName -Append

Import-Module _PS_Library_GeneralFunctionality
#Import-Module _CSOM_Library-SPO
#Import-Module _REST_Library-SPO


$webUrl = "https://anthesisllc.sharepoint.com"
$clientSite = "/clients"
$listOfClientFolders = @("_Kimble automatically creates Project folders","Background","Non-specific BusDev")
$listOfLeadProjSubFolders = @("Admin & contracts", "Analysis","Data & refs","Meetings","Proposal","Reports","Summary (marketing) - end of project")
$defaultProjectFilesToCopy = @(@{"fromList"="/teams/communities/heathandsafetyteam";"from"="/teams/communities/heathandsafetyteam/Shared Documents/RAs/Projects/Anthesis UK Project Risk Assessment.xlsx";"to"="/Admin & contracts/Anthesis UK Project Risk Assessment.xlsx";"conditions"="UK"})

$smtpServer = "anthesisgroup-com.mail.protection.outlook.com"
$mailFrom = "$(split-path $PSCommandPath -Leaf)_netmon@sustain.co.uk"
$mailTo = "kevin.maitland@anthesisgroup.com"
$recreateAllFolders = $false

$sharePointAdmin = "kimblebot@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\KimbleBot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass

$cacheFilePath = "$env:USERPROFILE\KimbleCache\"
$clientsCacheFile = "kimbleClients.csv"
$projectsCacheFile = "kimbleProjects.csv"
$leadsCacheFile = "kimbleLeads.csv"


Connect-PnPOnline –Url $($webUrl+$clientSite) –Credentials $adminCreds

#region functions
function cache-kimbleClients($pnpKimbleClientsList, $kimbleClientsCachePathAndFileName){
    $kcCacheFile = Get-Item $kimbleClientsCachePathAndFileName
    if((get-date $pnpKimbleClientsList.LastItemModifiedDate).AddMinutes(-5) -gt $kcCacheFile.LastWriteTimeUtc){#This is bodged so we don't miss any new clients added during the time it takes to actually download the full client list
        try{
            log-action -myMessage "[Kimble Clients] needs recaching - downloading full list" -logFile $fullLogPathAndName 
            $duration2 = Measure-Command {$spClients = get-spoKimbleClientListItems -spoCredentials $adminCreds}
            if($spClients){
                log-result -myMessage "SUCCESS: $($spClients.Count) Kimble Client records retrieved [$($duration2.Seconds) secs]!" -logFile $fullLogPathAndName
                if(!(Test-Path -Path $cacheFilePath)){New-Item -Path $cacheFilePath -ItemType Directory}
                $spClients | Export-Csv -Path $cacheFilePath$clientsCacheFile -Force -NoTypeInformation -Encoding UTF8
                }
            else{log-result -myMessage "FAILURE: [Kimble Clients] items could not be retrieved" -logFile $fullLogPathAndName}
            }
        catch{log-error -myError $_ -myFriendlyMessage "Could not retrieve [Kimble Clients] items to recache the local copy" -fullLogFile $fullLogPathAndNamel -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
        }
    else{log-result -myMessage "SUCCESS: Cache is up-to-date and does not require refreshing" -logFile $fullLogPathAndName}
    $clientCache = Import-Csv $cacheFilePath$clientsCacheFile
    $clientCache
    }
function cache-kimbleProjects($pnpKimbleProjectsList, $kimbleProjectsCachePathAndFileName){
    $projCacheFile = Get-Item $kimbleProjectsCachePathAndFileName
    if((get-date $pnpKimbleProjectsList.LastItemModifiedDate).AddMinutes(-5) -gt $projCacheFile.LastWriteTimeUtc){#This is bodged so we don't miss any new clients added during the time it takes to actually download the full client list
        try{
            log-action -myMessage "[Kimble Projects] needs recaching - downloading full list" -logFile $fullLogPathAndName 
            $duration = Measure-Command {$spoProjects = get-spoKimbleProjectListItems -spoCredentials $adminCreds}
            if($spoProjects){
                log-result -myMessage "SUCCESS: $($spoProjects.Count) Kimble Project records retrieved [$($duration.Seconds) secs]!" -logFile $fullLogPathAndName
                if(!(Test-Path -Path $cacheFilePath)){New-Item -Path $cacheFilePath -ItemType Directory}
                $spoProjects | Export-Csv -Path $cacheFilePath$projectsCacheFile -Force -NoTypeInformation -Encoding UTF8
                }
            else{log-result -myMessage "FAILURE: [Kimble Projects] items could not be retrieved" -logFile $fullLogPathAndName}
            }
        catch{log-error -myError $_ -myFriendlyMessage "Error retrieving [Kimble Projects] items to recache the local copy" -fullLogFile $fullLogPathAndNamel -errorLogFile -doNotLogToEmail $true}
        }
    else{log-result -myMessage "SUCCESS: Cache is up-to-date and does not require refreshing" -logFile $fullLogPathAndName}
    $projCache = Import-Csv $cacheFilePath$clientsCacheFile
    $projCache
    }
function new-clientFolder($spoKimbleClientList, $spoKimbleClientListItem, $arrayOfClientSubfolders, $recreateSubFolderOverride, $adminCreds, $fullLogPathAndName){
    log-action "new-clientFolder [$($spoKimbleClientListItem.Name)] Description: $(sanitise-stripHtml $spoKimbleClientListItem.ClientDescription)" -logFile $fullLogPathAndName
    $duration = Measure-Command {$newLibrary = new-spoClientLibrary -clientName $spoKimbleClientListItem.Name -clientDescription $spoKimbleClientListItem.ClientDescription -spoCredentials $adminCreds -verboseLogging $verboseLogging}
    if($newLibrary){#If the new Library has been created, make the subfolders and update the List Item
        log-result "SUCCESS: $($newLibrary.RootFolder.ServerRelativeUrl) is there [$($duration.Seconds) seconds]!" -logFile $fullLogPathAndName
        #Try to create the subfolders
        log-action "new-clientFolder $($newLibrary.RootFolder.ServerRelativeUrl) [subfolders]: $($arrayOfClientSubfolders -join ", ")" -logFile $fullLogPathAndName
        $formattedArrayOfClientSubfolders = @()
        $arrayOfClientSubfolders | % {$formattedArrayOfClientSubfolders += $($newLibrary.RootFolder.ServerRelativeUrl)+"/"+$_}
        $duration = Measure-Command {$lastNewSubfolder = add-spoLibrarySubfolders -pnpList $newLibrary -arrayOfSubfolderNames $formattedArrayOfClientSubfolders -recreateIfNotEmpty $recreateSubFolderOverride -spoCredentials $adminCreds -verboseLogging $verboseLogging}
        if($lastNewSubfolder){        
            log-result "SUCCESS: $($lastNewSubfolder) is there [$($duration.Seconds) seconds]!" -logFile $fullLogPathAndName
            #If we've got this far, try to update the IsDirty property on the Client in [Kimble Clients]
            $updatedValues = @{"IsDirty"=$false;"LibraryGUID"=$newLibrary.id.Guid}
            log-action "Set-PnPListItem Kimble Clients | $($spoKimbleClientListItem.Name) [$($spoKimbleClientListItem.Id) @{$(stringify-hashTable $updatedValues)}]" -logFile $fullLogPathAndName
            $duration = Measure-Command {$updatedItem = Set-PnPListItem -List $spoKimbleClientList.Id -Identity $spoKimbleClientListItem.SPListItemID -Values $updatedValues}
            if($updatedItem.FieldValues.IsDirty -eq $false){
                log-result "SUCCESS: Kimble Clients | $($spoKimbleClientListItem.Name) is no longer Dirty [$($duration.Seconds) seconds]" -logFile $fullLogPathAndName
                }
            else{log-result "FAILED: Could not set Client [$($spoKimbleClientListItem.Name)].IsDirty = `$false " -logFile $fullLogPathAndName}
            }
        else{log-result "FAILED: $($newLibrary.RootFolder.ServerRelativeUrl) [subfolders]: $($arrayOfClientSubfolders -join ", ") were not created properly" -logFile $fullLogPathAndName}
        }
    else{log-result "FAILED: Library for $($spoKimbleClientListItem.Name) was not created/retrievable!" -logFile $fullLogPathAndName}    
    }
function new-projectFolder($spoKimbleProjectList, $spoKimbleProjectListItem, $clientCacheHashTable, $arrayOfProjectSubfolders, $recreateSubFolderOverride, $adminCreds, $fullLogPathAndName){
    if($verboseLogging){Write-Host -ForegroundColor Cyan "new-projectFolder($($spoKimbleProjectList.Title), $($spoKimbleProjectListItem.Name), `$clientCacheHashTable, $($arrayOfProjectSubfolders -join ", "), $recreateSubFolderOverride=`$recreateSubFolderOverride)"}
    try{
        log-action -myMessage "Retrieving Client Library [$($clientCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["Name"])]" -logFile $fullLogPathAndName
        $clientLibrary = get-spoClientLibrary -clientName $($clientCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["Name"]) -clientLibraryGuid $($clientCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["LibraryId"]) -fullLogPathAndName $fullLogPathAndName -adminCreds $adminCreds -verboseLogging $verboseLogging
        if(!$clientLibrary){
            #Do something clever and add the Client Library. Or we can do nothing and wait for this to fix itself on the next run once ClientCache is updated again.
            }
        else{
            #Create the Project Folder
            log-result -myMessage "SUCCESS: $($clientLibrary.RootFolder.ServerRelativeUrl) retrieved" -logFile $fullLogPathAndName
            log-action -myMessage "Creating Project Folder [$($spoKimbleProjectListItem.Name)] in Client Library [$($clientCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["Name"])]" -logFile $fullLogPathAndName
            if($verboseLogging){Write-Host -ForegroundColor DarkCyan "add-spoLibrarySubfolders -pnpList $($clientLibrary.Title) -arrayOfSubfolderNames @($($spoKimbleProjectListItem.Name)) -recreateIfNotEmpty $recreateSubFolderOverride"}
            $projectFolder = add-spoLibrarySubfolders -pnpList $clientLibrary -arrayOfSubfolderNames @($clientLibrary.RootFolder.ServerRelativeUrl+"/"+$(sanitise-forPnpSharePoint $spoKimbleProjectListItem.Name)) -recreateIfNotEmpty $recreateSubFolderOverride -spoCredentials $adminCreds -verboseLogging $verboseLogging
            #Create the Project Subfolders
            if($projectFolder){
                log-result "SUCCESS: Project Folder [$projectFolder] created" -logFile $fullLogPathAndName
                log-action -myMessage "Creating Subfolders [$($arrayOfProjectSubfolders -join ",")] in [$($spoKimbleProjectListItem.Name)] in Client Library [$($clientCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["Name"])]" -logFile $fullLogPathAndName
                $subFolders =@()
                $arrayOfProjectSubfolders | % {$subFolders+= "$($clientLibrary.RootFolder.ServerRelativeUrl)/$(sanitise-forPnpSharePoint $spoKimbleProjectListItem.Name)/$_"}
                $lastSubfolder = add-spoLibrarySubfolders -pnpList $clientLibrary -arrayOfSubfolderNames $subFolders -recreateIfNotEmpty $recreateSubFolderOverride -spoCredentials $adminCreds -verboseLogging $verboseLogging
            #Populate any default files
                if($lastSubfolder){
                    log-result "SUCCESS: Project Subfolder [$lastSubfolder] created" -logFile $fullLogPathAndName
                    log-action -myMessage "Creating default files" -logFile $fullLogPathAndName
                    $defaultProjectFilesToCopy | % {
                        if($verboseLogging){Write-Host -ForegroundColor DarkCyan "copy-spoFile -fromList $($_.fromList) -from $($_.from) -to $($projectFolder.FieldValues.FileRef+$_.to)"}
                        copy-spoFile -fromList $_.fromList -from $_.from -to $($projectFolder.FieldValues.FileRef+$_.to) -spoCredentials $adminCreds
                        }
            #Update the List Item
                    try{
                        log-action -myMessage "Updating [Kimble Projects].[$($spoKimbleProjectListItem.Name)]" -logFile $fullLogPathAndName
                        $updatedValues = @{"IsDirty"=$false;"FolderGUID"=$($newProjectFolder.FieldValues.UniqueId)}
                        log-action "Set-PnPListItem Kimble Projects | $($spoKimbleProjectListItem.Name) [$($spoKimbleProjectListItem.Id)] @{$(stringify-hashTable $updatedValues)}]" -logFile $fullLogPathAndName
                        $duration = Measure-Command {$updatedItem = Set-PnPListItem -List $spoKimbleProjectList.Id -Identity $spoKimbleProjectListItem.SPListItemID -Values $updatedValues}
                        if($updatedItem.FieldValues.IsDirty -eq $false){
                            log-result "SUCCESS: Kimble Projects | $($spoKimbleProjectListItem.Name) is no longer Dirty [$($duration.Seconds) seconds]" -logFile $fullLogPathAndName
                            }
                        else{log-result "FAILED: Could not set Project [$($spoKimbleProjectListItem.Name)].IsDirty = `$false " -logFile $fullLogPathAndName}
                        }
                    catch{
                        #Error Updating list item
                        log-error -myError $_ -myFriendlyMessage "Error updating [Kimble Projects].[$($spoKimbleProjectListItem.Name)] in new-projectFolder" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
                        }
                    $projectFolder #return the new ProjectFolder to show it's worked
                    }
                else{log-result "FAILED: Project Subfolder [$($subFolders[$subFolders.Length-1])] for $($clientCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["Name"]) was not created" -logFile $fullLogPathAndName}
                }
            else{log-result "FAILED: Project Folder [$($spoKimbleProjectListItem.Name)] for $($clientCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["Name"]) was not created" -logFile $fullLogPathAndName}
            }
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error in new-projectFolder" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName}
    }
function update-clientFolder($spoKimbleClientList, $spoKimbleClientListItem, $arrayOfClientSubfolders, $recreateSubFolderOverride, $adminCreds, $fullLogPathAndName){
    log-action "update-clientFolder [$($spoKimbleClientListItem.Name)] - looking for existing Library" -logFile $fullLogPathAndName
    try{
        $duration = Measure-Command {
            if([string]::IsNullOrWhiteSpace($spoKimbleClientListItem.LibraryGUID)){
                if($verboseLogging){Write-Host -ForegroundColor DarkCyan "Looking for Client Library (no LibraryGUID, trying OldName): Get-PnPList -Identity $($spoKimbleClientListItem.PreviousName)"}
                $existingLibrary = Get-PnPList -Identity $spoKimbleClientListItem.PreviousName
                }
            else{
                if($verboseLogging){Write-Host -ForegroundColor DarkCyan "Looking for Client Library using LibraryGUID: Get-PnPList -Identity $($spoKimbleClientListItem.LibraryGUID)"}
                $existingLibrary = Get-PnPList -Identity $spoKimbleClientListItem.LibraryGUID
                }
            }
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving Client Library in update-clientFolder [$($spoKimbleClientListItem.Name)][$($spoKimbleClientListItem.LibraryGUID)] $($Error[0].Exception.InnerException.Response)" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName}
    #If we can't find it by the old name, try finding it by the New name (as it might just be the Description that's changed)
    if(!$existingLibrary){
        try{
            $duration2 = Measure-Command {
                if($verboseLogging){Write-Host -ForegroundColor DarkCyan "Looking for Client Library (no LibraryGUID, trying NewName): Get-PnPList -Identity $($spoKimbleClientListItem.PreviousName)"}
                $existingLibrary = Get-PnPList -Identity $spoKimbleClientListItem.Name
                }
            }
        catch{log-error -myError $_ -myFriendlyMessage "Error retrieving Client Library in update-clientFolder [$($spoKimbleClientListItem.Name)][$($spoKimbleClientListItem.LibraryGUID)] $($Error[0].Exception.InnerException.Response)" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName}
        }
    if($existingLibrary){
        log-result -myMessage "SUCCESS: [$($existingLibrary.RootFolder.ServerRelativeUrl)] found (GUID:[$($existingLibrary.Id.Guid)] [$($duration.Seconds + $duration2.Seconds) seconds])" -logFile $fullLogPathAndName
        log-action -myMessage "Updating Client Library [$($existingLibrary.RootFolder.ServerRelativeUrl)]" -logFile $fullLogPathAndName
        try{
            #Update the Library
            if($verboseLogging){Write-Host -ForegroundColor DarkCyan "Updating Library [$($existingLibrary.RootFolder.ServerRelativeUrl)] Description:[$(sanitise-stripHtml $spoKimbleClientListItem.ClientDescription)]"}
            $duration = Measure-Command {
                $existingLibrary.Description = $(sanitise-stripHtml $spoKimbleClientListItem.ClientDescription)
                $existingLibrary.Update()
                if($verboseLogging){Write-Host -ForegroundColor DarkCyan "Updating Library [$($existingLibrary.RootFolder.ServerRelativeUrl)] Title:[$($spoKimbleClientListItem.Name)]: Set-PnPList -Identity $($existingLibrary.Id.Guid) -Title $($spoKimbleClientListItem.Name)"}
                Set-PnPList -Identity $existingLibrary.Id -Title $spoKimbleClientListItem.Name
                $updatedLibrary = Get-PnPList -Identity $spoKimbleClientListItem.Name
                }
            if($updatedLibrary){
                log-result -myMessage "SUCCESS: Client Library [$($existingLibrary.RootFolder.ServerRelativeUrl)] updated successfully (no error on update) [$($duration.Seconds) secs]" -logFile $fullLogPathAndName
                try{
                    #Update the List Item
                    $duration = Measure-Command {
                        $updatedValues =@{"LibraryGUID"=$existingLibrary.Id.Guid;"IsDirty"=$false}
                        if($verboseLogging){Write-Host -ForegroundColor DarkCyan "Updating [Kimble Client] List item [$($spoKimbleClientListItem.Name)]: Set-PnPListItem -List $($spoKimbleClientList.Title) -Identity $($spoKimbleClientListItem.SPListItemID) `$updatedValues = @{$(stringify-hashTable $updatedValues)}"}
                        $updatedListItem = Set-PnPListItem -List $spoKimbleClientList -Identity $spoKimbleClientListItem.SPListItemID -Values $updatedValues
                        }
                    if($updatedListItem.FieldValues.IsDirty -eq $false){log-result -myMessage "SUCCESS: [Kimble Clients].[$($spoKimbleClientListItem.Name)] updated successfully (no error on update) [$($duration.Seconds) secs]" -logFile $fullLogPathAndName}
                    else{log-result -myMessage "FAILED: [Kimble Clients].[$($spoKimbleClientListItem.Name)] was not updated" -logFile $fullLogPathAndName}
                    }
                catch{
                    #Failed to update SPListItem
                    log-error -myError $_ -myFriendlyMessage "Error updating [Kimble Clients].[$($spoKimbleClientListItem.Name)] - this is still marked as IsDirty=`$true :( [$($Error[0].Exception.InnerException.Response)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
                    }
                }
            }
        catch{
            #Failed to update Client Library
            log-result -myMessage "FAILED: Client Library [$($existingLibrary.Title)] was not updated" -logFile $fullLogPathAndName
            log-error -myError $_ -myFriendlyMessage "Error updating Client Library [$($existingLibrary.Title)] - this is still marked as IsDirty=`$true :( [$($Error[0].Exception.InnerException.Response)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
            }
        }
    else{
        #Couldn't find the Library, so try creating a new one to paper over the cracks. #WCGW
        log-result -myMessage "FAILED: Could not retrieve a Client Library for [Kimble Clients].[$($spoKimbleClientListItem.Name)] - sending it back for re-creation :/" -logFile $fullLogPathAndName
        log-action -myMessage "Sending [Kimble Clients].[$($spoKimbleClientListItem.Name)] back for re-creation as it has mysteriously disappeared" -logFile $fullLogPathAndName
        if($verboseLogging){Write-Host -ForegroundColor DarkCyan "new-clientFolder -spoKimbleClientList $($spoKimbleClientList.Title) -spoKimbleClientListItem $($spoKimbleClientListItem.Name) -arrayOfClientSubfolders @($($arrayOfClientSubfolders -join ",")) -recreateSubFolderOverride `$false"}
        try{
            $duration = Measure-Command {$newLibrary = new-clientFolder -spoKimbleClientList $spoKimbleClientList -spoKimbleClientListItem $spoKimbleClientListItem -arrayOfClientSubfolders $arrayOfClientSubfolders -recreateSubFolderOverride $false -adminCreds $adminCreds -fullLogPathAndName $fullLogPathAndName}
            if($newLibrary){log-result -myMessage "SUCCESS: Weirdly unfindable Client Library [$($newLibrary.RootFolder.ServerRelativeUrl)] was recreated  [$($duration.Seconds) secs]" -logFile $fullLogPathAndName}
            else{
                log-result -myMessage "FAILED: Someone left a sponge in the patient - I couldn't retrieve a Library for [$($spoKimbleClientListItem.Name)] and I couldn't create a new one either..." -logFile $fullLogPathAndName
                log-error -myError $null -myFriendlyMessage "Borked Client Library update [$($spoKimbleClientListItem.Name)]" -smtpServer $smtpServer -mailTo $mailTo -mailFrom $mailFrom
                }
            }
        catch{log-error -myError $_ -myFriendlyMessage "Error: Borked Client Library update [$($spoKimbleClientListItem.Name)]" -smtpServer $smtpServer -mailTo $mailTo -mailFrom $mailFrom}
        }
    }
function update-projectFolder($spoKimbleProjectList, $spoKimbleProjectListItem, $clientCacheHashTable, $arrayOfProjectSubfolders, $recreateSubFolderOverride, $adminCreds, $fullLogPathAndName){
    #Get the ClientLibrary
    log-action -myMessage "Retrieving Client Library [$($clientCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["Name"])]" -logFile $fullLogPathAndName
    $clientLibrary = get-spoClientLibrary -clientName $($clientCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["Name"]) -clientLibraryGuid $($clientCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["LibraryId"]) -fullLogPathAndName $fullLogPathAndName -adminCreds $adminCreds -verboseLogging $verboseLogging

    #Check for Client change
    if(![string]::IsNullOrWhiteSpace($spoKimbleProjectListItem.PreviousKimbleClientId) -and ($spoKimbleProjectListItem.KimbleClientId -ne $spoKimbleProjectListItem.PreviousKimbleClientId)){
        log-result -myMessage "Project has PreviousKimbleClientId - checking whether it needs moving" -logFile $fullLogPathAndName
        $oldClientName = $($clientCacheHashTable[$spoKimbleProjectListItem.PreviousKimbleClientId]["Name"])
        if($oldClientName){
            log-action -myMessage "Retrieving Previous Client Library [$($clientCacheHashTable[$spoKimbleProjectListItem.PreviousKimbleClientId]["Name"])]" -logFile $fullLogPathAndName
            $oldClientLibrary = get-spoClientLibrary -clientName $($clientCacheHashTable[$spoKimbleProjectListItem.PreviousKimbleClientId]["Name"]) -clientLibraryGuid $($clientCacheHashTable[$spoKimbleProjectListItem.PreviousKimbleClientId]["LibraryId"]) -fullLogPathAndName $fullLogPathAndName -adminCreds $adminCreds -verboseLogging $verboseLogging
            if($oldClientLibrary){
    #Move the folder to the new client
                log-result -myMessage "SUCCESS: Previous Client Library [$($clientCacheHashTable[$spoKimbleProjectListItem.PreviousKimbleClientId]["Name"])] retrieved" -logFile $fullLogPathAndName
                log-action -myMessage "Looking for Project folder in Previous Client Library [$($clientCacheHashTable[$spoKimbleProjectListItem.PreviousKimbleClientId]["Name"])]" -logFile $fullLogPathAndName
                $misplacedProjectFolder = get-spoFolder -pnpList $oldClientLibrary -folderServerRelativeUrl $(format-asServerRelativeUrl -serverRelativeUrl $oldClientLibrary.RootFolder.ServerRelativeUrl -stringToFormat $(sanitise-forPnpSharePoint $spoKimbleProjectListItem.PreviousName)) -folderGuid $spoKimbleProjectListItem.FolderGUID -verboseLogging $verboseLogging
                if(!$misplacedProjectFolder){$misplacedProjectFolder = get-spoFolder -pnpList $oldClientLibrary -folderServerRelativeUrl $(format-asServerRelativeUrl -serverRelativeUrl $oldClientLibrary.RootFolder.ServerRelativeUrl -stringToFormat $(sanitise-forPnpSharePoint $spoKimbleProjectListItem.Name)) -verboseLogging $verboseLogging}
                if($misplacedProjectFolder -and $clientLibrary){
                    log-result -myMessage "SUCCESS: Project folder [$($misplacedProjectFolder.FieldValues.FileRef)] found in Previous Client [$($clientCacheHashTable[$spoKimbleProjectListItem.PreviousKimbleClientId]["Name"])]" -logFile $fullLogPathAndName
                    log-action -myMessage "Moving Project folder [$($misplacedProjectFolder.FieldValues.FileRef)] from [$($clientCacheHashTable[$spoKimbleProjectListItem.PreviousKimbleClientId]["Name"])] to [$($clientCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["Name"])]" -logFile $fullLogPathAndName
                    $libraryRelavtiveUrl = $misplacedProjectFolder.FieldValues.FileRef.Replace("/clients/","") #Yeah, this isn't a great way to handle it.
                    $movedFolder = Move-PnPFolder -Folder $libraryRelavtiveUrl -TargetFolder $clientLibrary.RootFolder.ServerRelativeUrl
                    if($movedFolder){log-result -myMessage "SUCCESS: Project folder [$($movedFolder.Name)] moved from [$($clientCacheHashTable[$spoKimbleProjectListItem.PreviousKimbleClientId]["Name"])] to [$($clientCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["Name"])]" -logFile $fullLogPathAndName}
                    else{log-result -myMessage "FAILED: Project folder [$($misplacedProjectFolder.FieldValues.FileRef)] could not be moved from [$($clientCacheHashTable[$spoKimbleProjectListItem.PreviousKimbleClientId]["Name"])] to [$($clientCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["Name"])]" -logFile $fullLogPathAndName}
                    }
                else{log-result -myMessage "Project folder [$($spoKimbleProjectListItem.Name)] not found in Previous Client [$($clientCacheHashTable[$spoKimbleProjectListItem.PreviousKimbleClientId]["Name"])] - it may already have been moved/deleted" -logFile $fullLogPathAndName}
                }
            }
        }

    if($clientLibrary){
    #Get the ProjectFolder
        log-result -myMessage "SUCCESS: $($clientLibrary.RootFolder.ServerRelativeUrl) retrieved" -logFile $fullLogPathAndName
        log-action -myMessage "Retrieving Project Folder [$($spoKimbleProjectListItem.Name)] in Client Library [$($clientCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["Name"])]" -logFile $fullLogPathAndName
        $misnamedProjectFolder = get-spoFolder -pnpList $clientLibrary -folderServerRelativeUrl $(format-asServerRelativeUrl -serverRelativeUrl $clientLibrary.RootFolder.ServerRelativeUrl -stringToFormat $(sanitise-forPnpSharePoint $spoKimbleProjectListItem.PreviousName)) -folderGuid $spoKimbleProjectListItem.FolderGUID -verboseLogging $verboseLogging
        if($misnamedProjectFolder){
    #Rename the ProjectFolder
            log-result -myMessage "SUCCESS: Misnamed Project Folder [$($misnamedProjectFolder.FieldValues.FileRef)] retrieved - will rename to $(sanitise-forPnpSharePoint $spoKimbleProjectListItem.Name)" -logFile $fullLogPathAndName
            $misnamedProjectFolder.ParseAndSetFieldValue("Title",$(sanitise-forPnpSharePoint $spoKimbleProjectListItem.Name))
            $misnamedProjectFolder.ParseAndSetFieldValue("FileLeafRef",$(sanitise-forPnpSharePoint $spoKimbleProjectListItem.Name))
            $misnamedProjectFolder.Update()
            $misnamedProjectFolder.Context.ExecuteQuery()
            }
        $correctlyNamedProjectFolder = get-spoFolder -pnpList $clientLibrary -folderServerRelativeUrl $(format-asServerRelativeUrl -serverRelativeUrl $clientLibrary.RootFolder.ServerRelativeUrl -stringToFormat $(sanitise-forPnpSharePoint $spoKimbleProjectListItem.Name)) -verboseLogging $verboseLogging
        if($correctlyNamedProjectFolder){
    #Update the ListItem
            log-result -myMessage "SUCCESS: Correctly-named Project Folder [$($correctlyNamedProjectFolder.FieldValues.FileRef)] retrieved" -logFile $fullLogPathAndName
            try{
                #Update LIstItem
                $updatedValues = @{"IsDirty"=$false;"FolderGUID"=$($correctlyNamedProjectFolder.FieldValues.UniqueId)}
                log-action "Set-PnPListItem Kimble Projects | $($spoKimbleProjectListItem.Name) [$($spoKimbleProjectListItem.Id)] @{$(stringify-hashTable $updatedValues)}]" -logFile $fullLogPathAndName
                $duration = Measure-Command {$updatedItem = Set-PnPListItem -List $spoKimbleProjectList.Id -Identity $spoKimbleProjectListItem.SPListItemID -Values $updatedValues}
                if($updatedItem.FieldValues.IsDirty -eq $false){
                    log-result "SUCCESS: Kimble Projects | $($spoKimbleProjectListItem.Name) is no longer Dirty [$($duration.Seconds) seconds]" -logFile $fullLogPathAndName
                    }
                else{log-result "FAILED: Could not set Projects [$($spoKimbleProjectListItem.Name)].IsDirty = `$false " -logFile $fullLogPathAndName}
                }
            catch{
                #Error Updating list item
                log-error -myError $_ -myFriendlyMessage "Error updating [Kimble Projects].[$($spoKimbleProjectListItem.Name)] in update-projectFolder" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
                }
            $correctlyNamedProjectFolder #Return the updated project folder to show it worked
            }
        else{
            log-result -myMessage "FAILED: Folder for project [$($spoKimbleProjectListItem.Name)] could not be retrieved. That'll learn you for trying to guess ServerRelativeUrls." -logFile $fullLogPathAndName
            #It's also possible that there was a problem setting up the project originally, so I've we've set an override, try recreating the folders:
            if($recreateSubFolderOverride){
                log-action -myMessage "Override set - recreating Project Folder [$($spoKimbleProjectListItem.Name)] in Client Library [$($clientCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["Name"])]" -logFile $fullLogPathAndName
                $newProjectFolder = new-projectFolder -spoKimbleProjectList $spoKimbleProjectList -spoKimbleProjectListItem $spoKimbleProjectListItem -clientCacheHashTable $clientCacheHashTable -arrayOfProjectSubfolders $arrayOfProjectSubfolders -recreateSubFolderOverride $recreateSubFolderOverride -adminCreds $adminCreds -fullLogPathAndName $fullLogPathAndName
                if($newProjectFolder){log-result -myMessage "SUCCESS: Folder for project [$($spoKimbleProjectListItem.Name)] Recreated" -logFile $fullLogPathAndName}
                else{log-result -myMessage "FAILED: recreating Folder for project [$($spoKimbleProjectListItem.Name)]. This one's properly borked." -logFile $fullLogPathAndName}
                $newProjectFolder
                }
            }
        }
    else{
        #No Client Library found. Probably best /not/ to create a new for a "update" as there's probably something wrong.
        }
    }
#endregion





#region Process Clients
#Get [Kimble Clients] List from SPO
try{
    log-action -myMessage "Getting [Kimble Clients] to check whether it needs recaching" -logFile $fullLogPathAndName 
    $kc = Get-PnPList -Identity "Kimble Clients" -Includes ContentTypes, LastItemModifiedDate
    if($kc){log-result -myMessage "SUCCESS: List retrieved" -logFile $fullLogPathAndName}
    else{log-result -myMessage "FAILURE: List could not be retrieved" -logFile $fullLogPathAndName}
    }
catch{log-error -myError $_ -myFriendlyMessage "Could not retrieve [Kimble Clients] to check whether it needs recaching" -fullLogFile $fullLogPathAndName -errorLogFile -doNotLogToEmail $true}

#Get the list of Projects to update *before* we get the list of Clients, so we don't create a race condition where any Projects created while we're processing the Clients incorrectly appear orphaned
$dirtyProjects = get-spoKimbleProjectListItems -camlQuery "<View><Query><Where><Eq><FieldRef Name='IsDirty'/><Value Type='Boolean'>1</Value></Eq></Where></Query></View>" -spoCredentials $adminCreds -verboseLogging $verboseLogging
#Retrieve (and update if necessary) the full Clients Cache as we'll need it to set up any new Leads/Projects
$clientCache = cache-kimbleClients -pnpKimbleClientsList $kc -kimbleClientsCachePathAndFileName $($cacheFilePath+$clientsCacheFile)
#Build a hashtable so we can look up Client name by it's KimbleId
$kimbleClientHashTable = @{}
$clientCache | % {$kimbleClientHashTable.Add($_.Id, @{"Name"=$_.Name;"LibraryId"=$_.LibraryGUID})}

#Process any [Kimble Clients] flagged as IsDirty
$dirtyClients = $clientCache | ?{$_.IsDirty -eq $true -and $_.IsDeleted -eq $false}
$i = 1
$dirtyClients | % {
    Write-Progress -Id 1000 -Status "Processing DirtyClients" -Activity "$i/$($dirtyClients.Count)" -PercentComplete ($i/$dirtyClients.Count) #Display the overall progress
    $dirtyClient = $_
    $duration = Measure-Command {
        log-action -myMessage "************************************************************************" -logFile $fullLogPathAndName
        log-action -myMessage "CLIENT [$($dirtyClient.Name)][$i/$($dirtyClients.Count)] isDirty!" -logFile $fullLogPathAndName
        #Check if the Client needs creating
        if(([string]::IsNullOrEmpty($dirtyClient.PreviousName) -and [string]::IsNullOrEmpty($dirtyClient.PreviousDescription)) -OR $recreateAllFolders -eq $true){
            log-action -myMessage "CLIENT [$($dirtyClient.Name)] looks new - creating new Library" -logFile $fullLogPathAndName
            #Create a new Library and subfolders
            try{
                new-clientFolder -spoKimbleClientList $kc -spoKimbleClientListItem $dirtyClient -arrayOfClientSubfolders $listOfClientFolders -recreateSubFolderOverride $recreateAllFolders -adminCreds $adminCreds -fullLogPathAndName $fullLogPathAndName 
                }
            catch{log-error $_ -myFriendlyMessage "Failed to create new Library for $($dirtyClient.Name)" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -smtpServer $smtpServer -mailTo $mailTo -mailFrom $mailFrom}

            #Now try to add the new ClientName to the TermStore
            try{
                log-action "add-termToStore: Kimble | Clients | $($dirtyClient.Name)" -logFile $fullLogPathAndName
                $duration = Measure-Command {$newTerm = add-spoTermToStore -termGroup "Kimble" -termSet "Clients" -term $($dirtyClient.Name) -kimbleId $dirtyClient.Id -verboseLogging $verboseLogging}
                if($newTerm){log-result "SUCCESS: Kimble | Clients | $($dirtyClient.Name) added to Managed MetaData Term Store [$($duration.Seconds) secs]" -logFile $fullLogPathAndName}
                }
            catch{log-error $_ -myFriendlyMessage "Failed to add $($dirtyClient.Title) to Term Store" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -smtpServer $smtpServer -mailFrom $mailFrom -mailTo $mailTo}
            }

        #Otherwise try to update it
        else{
            log-action -myMessage "CLIENT [$($dirtyClient.Name)] doesn't look new, so I'm going to try updating it" -logFile $fullLogPathAndName
            try{
                update-clientFolder -spoKimbleClientList $kc -spoKimbleClientListItem $dirtyClient -arrayOfClientSubfolders $arrayOfClientSubfolders -recreateSubFolderOverride $recreateAllFolders -adminCreds $adminCreds -fullLogPathAndName $fullLogPathAndName
                }
            catch{log-error $_ -myFriendlyMessage "Error updating Client [$($dirtyClient.Name)]"}

            #Then try updating the Managed Metadata
            try{
                log-action -myMessage "Updating Managed Metadata for $($dirtyClient.Name)" -logFile $fullLogPathAndName
                if($verboseLogging){Write-Host -ForegroundColor DarkCyan "update-spoTerm -termGroup 'Kimble' -termSet 'Clients' -oldTerm $($dirtyClient.PreviousName) -newTerm $($dirtyClient.Name) -kimbleId $($dirtyClient.Id)"}
                $duration = Measure-Command {$updatedTerm = update-spoTerm -termGroup "Kimble" -termSet "Clients" -oldTerm $($dirtyClient.PreviousName) -newTerm $($dirtyClient.Name) -kimbleId $($dirtyClient.Id) -verboseLogging $verboseLogging}
                if($updatedTerm){log-result "SUCCESS: Kimble | Clients | [$($dirtyClient.PreviousName)] updated to [$($dirtyClient.Name)] in Managed MetaData Term Store [$($duration.Seconds) secs]" -logFile $fullLogPathAndName}
                }
            catch{
                #Failed to update Managed Metadata
                log-error -myError $_ -myFriendlyMessage "Error updating Managed Metadata Term [$($dirtyClient.PreviousName)] to [$($dirtyClient.Name)] [$($Error[0].Exception.InnerException.Response)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
                }
            if($(sanitise-forTermStore $updatedTerm.Name) -eq $(sanitise-forTermStore $dirtyClient.Name)){log-result -myMessage "SUCCESS: Managed Metadata Term [$($dirtyClient.PreviousName)] updated to [$($dirtyClient.Name)] [$($duration.Seconds) secs]" -logFile $fullLogPathAndName}
            else{log-result -myMessage "FAILED: Managed Metadata Term [$($dirtyClient.PreviousName)] did not update to [$($dirtyClient.Name)]" -logFile $fullLogPathAndName}
            }
        $i++
        }
    log-result "DirtyClient [$($dirtyClient.Name)] proccessed in $($duration.Seconds) seconds" -logFile $fullLogPathAndName
    }
#endregion

#region Process Projects
#Get [Kimble Projects] List from SPO
try{
    log-action -myMessage "Getting [Kimble Projects]" -logFile $fullLogPathAndName 
    $kp = Get-PnPList -Identity "Kimble Projects" -Includes ContentTypes, LastItemModifiedDate
    if($kp){log-result -myMessage "SUCCESS: List retrieved" -logFile $fullLogPathAndName}
    else{log-result -myMessage "FAILURE: List could not be retrieved" -logFile $fullLogPathAndName}
    }
catch{log-error -myError $_ -myFriendlyMessage "Could not retrieve [Kimble Projects]" -fullLogFile $fullLogPathAndName -errorLogFile -doNotLogToEmail $true}

#Process any [Kimble Projects] flagged as IsDirty
#We got $dirtyProjects before the Clients to avoid a race condition
$dirtyProjects | % {
    $dirtyProject = $_
    log-action -myMessage "************************************************************************" -logFile $fullLogPathAndName
    log-action -myMessage "Project [$($dirtyProject.Name)] IsDirty" -dirtyProject $fullLogPathAndName -logFile $fullLogPathAndName
    log-action -myMessage "Checking that Client with Id [$($dirtyProject.KimbleClientId)] is in the Cache" -logFile $fullLogPathAndName
    try{
        if ($kimbleClientHashTable[$dirtyProject.KimbleClientId]){
            log-result "SUCCESS: Client [$($kimbleClientHashTable[$dirtyProject.KimbleClientId]["Name"])] found in cache LibraryId:[$($kimbleClientHashTable[$dirtyProject.KimbleClientId]["LibraryId"])]" -logFile $fullLogPathAndName
            }
        else{
            log-result "FAILED: Project Folder [$($dirtyProject.Name)] could not be created because I couldn't identify the Client with Id [$($dirtyProject.KimbleClientId)]" -logFile $fullLogPathAndName
            #This will flood the error logs indefinitely, so mark it as IsDirty = $false. If the user wants it recreated, they can fix the Client then update the Project
            try{
                log-action -myMessage "Updating [Kimble Projects].[$($dirtyProject.Name)]" -logFile $fullLogPathAndName
                $updatedValues = @{"IsDirty"=$false}
                log-action "Set-PnPListItem Kimble Projects | $($dirtyProject.Name) [$($dirtyProject.Id)] @{$(stringify-hashTable $updatedValues)}]" -logFile $fullLogPathAndName
                $duration = Measure-Command {$updatedItem = Set-PnPListItem -List $kp.Id -Identity $dirtyProject.SPListItemID -Values $updatedValues}
                if($updatedItem.FieldValues.IsDirty -eq $false){
                    log-result "SUCCESS: Kimble Projects  | $($dirtyProject.Name) is no longer Dirty [$($duration.Seconds) seconds]" -logFile $fullLogPathAndName
                    }
                else{log-result "FAILED: Could not set Project [$($dirtyProject.Name)].IsDirty = `$false " -logFile $fullLogPathAndName}
                return
                }
            catch{
                #Error Updating list item
                log-error -myError $_ -myFriendlyMessage "Error updating [Kimble Projects].[$($dirtyProject.Name)] in new-projectFolder" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
                }
            }
        }
    catch{
        #Couldn't look up Client Name in Hash Table
        log-error -myError $_ -myFriendlyMessage "Error looking up Client ID [$($dirtyProject.KimbleClientId)]. Project is [$($dirtyProject.Name)] for further troubleshooting. " -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
        }

    if($dirtyProject.DoNotProcess){#Some Projects shouldn't have folders set up. Just mark it as IsDirty = $false
        try{
            #Update LIstItem
            $updatedValues = @{"IsDirty"=$false}
            log-action "Set-PnPListItem Kimble Projects | $($dirtyProject.Name) [$($dirtyProject.Id)] @{$(stringify-hashTable $updatedValues)}]" -logFile $fullLogPathAndName
            $duration = Measure-Command {$updatedItem = Set-PnPListItem -List $kp.Id -Identity $dirtyProject.SPListItemID -Values $updatedValues}
            if($updatedItem.FieldValues.IsDirty -eq $false){
                log-result "SUCCESS: Kimble Projects | $($dirtyProject.Name) is no longer Dirty [$($duration.Seconds) seconds]" -logFile $fullLogPathAndName
                }
            else{log-result "FAILED: Could not set Projects [$($dirtyProject.Name)].IsDirty = `$true " -logFile $fullLogPathAndName}
            }
        catch{
            #Error Updating list item
            log-error -myError $_ -myFriendlyMessage "Error updating [Kimble Projects].[$($dirtyProject.Name)] in update-projectFolder" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
            }

        }
    elseif(!$dirtyProject.PreviousName -and (!$dirtyProject.PreviousKimbleClientId -or $dirtyProject.PreviousKimbleClientId -eq $dirtyProject.KimbleClientId)){
        #Create a new folder tree under the Client Library
        if($($kimbleClientHashTable[$dirtyProject.KimbleClientId])){log-action -myMessage "PROJECT [$($dirtyProject.Name)] for client [$($kimbleClientHashTable[$dirtyProject.KimbleClientId]["Name"])] looks new - creating subfolders!" -logFile $fullLogPathAndName}
        else{log-action -myMessage "PROJECT [$($dirtyProject.Name)] looks new, but I can't work out the Client. Creating subfolders anyway, but this probbaly won't work" -logFile $fullLogPathAndName}
        
        try{
            if($verboseLogging){Write-Host -ForegroundColor DarkCyan "new-projectFolder -spoKimbleProjectList $($kp.Title) -spoKimbleProjectListItem $($dirtyProject.Name) -clientCacheHashTable `$kimbleClientHashTable -arrayOfProjectSubfolders $($listOfLeadProjSubFolders -join ", ")"}
            $duration = Measure-Command {$newProjectFolder = new-projectFolder -spoKimbleProjectList $kp -spoKimbleProjectListItem $dirtyProject -clientCacheHashTable $kimbleClientHashTable -arrayOfProjectSubfolders $listOfLeadProjSubFolders -adminCreds $adminCreds -fullLogPathAndName $fullLogPathAndName }
            if($newProjectFolder){
                log-result -myMessage "SUCCESS: Project Folder [$($newProjectFolder.FieldValues.FileRef)] created successfully [$($duration.Seconds) secs]!" -logFile $fullLogPathAndName}
            else{log-result "FAILED: Project Folder [$($dirtyProject.Name)] for [$($kimbleClientHashTable[$dirtyProject.KimbleClientId]["Name"])] was not created" -logFile $fullLogPathAndName}
            }
        catch{
            #Failed to create new Project Folder
            log-error -myError $_ -myFriendlyMessage "Error creating new Project Folder [$($dirtyProject.Name)] for [$($kimbleClientHashTable[$dirtyProject.KimbleClientId]["Name"])]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
            }
        }
    else{#Otherwise try updating it
        try{
            log-action -myMessage "PROJECT [$($dirtyProject.Name)] for client [$($kimbleClientHashTable[$dirtyProject.KimbleClientId]["Name"])] looks like it needs updating!" -logFile $fullLogPathAndName
            $duration = Measure-Command {$updatedProjectFolder = update-projectFolder -spoKimbleProjectList $kp -spoKimbleProjectListItem $dirtyProject -clientCacheHashTable $kimbleClientHashTable -arrayOfProjectSubfolders $listOfLeadProjSubFolders -recreateSubFolderOverride $recreateAllFolders -adminCreds $adminCreds -fullLogPathAndName $fullLogPathAndName}
            if($updatedProjectFolder){
                log-result -myMessage "SUCCESS - Project Folder [$($dirtyProject.Name)] updated successfully [$($duration.Seconds) secs]!" -logFile $fullLogPathAndName
                }
            else{log-result "FAILED - Project Folder [$($dirtyProject.Name)] for [$($kimbleClientHashTable[$dirtyProject.KimbleClientId]["Name"])] was not updated" -logFile $fullLogPathAndName}
            }
        catch{
            #Error updating Project Folder
            log-error -myError $_ -myFriendlyMessage "Error updating project folder [$($dirtyProject.Name)] for [$($kimbleClientHashTable[$dirtyProject.KimbleClientId]["Name"]))]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
            }
        }
    }
#endregion

<#
function reconcile-spoClients(){
    try{
        log-action -myMessage "Getting all Client Libraries:" -logFile $fullLogPathAndName
        $spoClientLibraries = get-allLists -serverUrl $webUrl -sitePath "/clients" -restCreds $restCreds -logFile $fullLogPathAndName -verboseLogging $true
        if($spoClients){log-result -myMessage "SUCCESS: Libraries retrieved!" -logFile $fullLogPathAndName}
        else{log-result -myMessage "FAILED: Unable to retrieve list of Libraries" -logFile $fullLogPathAndName}
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving List: [$listName]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}

    try{
        log-action -myMessage "Getting List: [Kimble Clients]" -logFile $fullLogPathAndName
        $spoClients = get-itemsInList -serverUrl $webUrl  -sitePath $sitePath -listName "Kimble Clients" -restCreds $restCreds -logFile $logFileLocation 
        if($spoClients){log-result -myMessage "SUCCESS: List retrieved!" -logFile $fullLogPathAndName}
        else{log-result -myMessage "FAILED: Unable to retrieve list" -logFile $fullLogPathAndName}
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving List: [Kimble Clients]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}

    $missingClientLibraries = Compare-Object -ReferenceObject $spoClients -DifferenceObject $spoClientLibraries.results -Property "Title" -PassThru
    $missingClientLibraries | %{
        if($_.SideIndicator -eq "<="){new-clientFolder -clientName $($_.Title) -clientDescription $($_.ClientDescription) -listofClientSubfolders $listOfClientFolders -webUrl $webUrl -restCreds $restCreds -fullLogPathAndName $fullLogPathAndName -errorLogPathAndName $errorLogPathAndName -digest $clientDigest}
        }
    }
#>


<#
$theseProjects = get-spoKimbleProjectListItems -camlQuery "<View><Query><Where><Eq><FieldRef Name='KimbleClientId'/><Value Type='Text'>0012400000TVyffAAD</Value></Eq></Where></Query></View>"
$theseProjects | % {
    $thisProject = $_
    try{
        #Update LIstItem
        $updatedValues = @{"DoNotProcess"=$true;"IsDirty"=$false}
        log-action "Set-PnPListItem Kimble Projects | $($thisProject.Name) [$($thisProject.Id)] @{$(stringify-hashTable $updatedValues)}]" -logFile $fullLogPathAndName
        $duration = Measure-Command {$updatedItem = Set-PnPListItem -List $kp.Id -Identity $thisProject.SPListItemID -Values $updatedValues}
        if($updatedItem.FieldValues.IsDirty -eq $false){
            log-result "SUCCESS: Kimble Projects | $($thisProject.Name) is no longer Dirty [$($duration.Seconds) seconds]" -logFile $fullLogPathAndName
            }
        else{log-result "FAILED: Could not set Projects [$($thisProject.Name)].IsDirty = `$true " -logFile $fullLogPathAndName}
        }
    catch{
        #Error Updating list item
        log-error -myError $_ -myFriendlyMessage "Error updating [Kimble Projects].[$($thisProject.Name)] in update-projectFolder" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
        }

    }
#>