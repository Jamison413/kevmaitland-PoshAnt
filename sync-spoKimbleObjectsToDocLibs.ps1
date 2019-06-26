param(
    # Specifies whether we are updating Clients or Suppliers.
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Clients", "Suppliers","Projects","ClientsProjects")]
    [string]$whatToSync
    )
    $verboseLogging = $true

$logFileLocation = "C:\ScriptLogs\"
$fullLogPathAndName = $logFileLocation+"sync-spoKimbleObjectsToDocLibs.ps1_$whatToSync`_FullLog_$(Get-Date -Format "yyMMdd").log"
$errorLogPathAndName = $logFileLocation+"sync-spoKimbleObjectsToDocLibs.ps1_$whatToSync`_ErrorLog_$(Get-Date -Format "yyMMdd").log"
if($PSCommandPath){
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_$whatToSync`_Transcript_$(Get-Date -Format "yyMMdd").log"
    Start-Transcript $transcriptLogName -Append
    }

Import-Module _PS_Library_GeneralFunctionality
Import-Module _PNP_Library_SPO
Import-Module SharePointPnPPowerShellOnline
#Import-Module _CSOM_Library-SPO
#Import-Module _REST_Library-SPO


$webUrl = "https://anthesisllc.sharepoint.com"
$listOfClientFolders = @("_Kimble automatically creates Project folders","Background","Non-specific BusDev")
$listOfSupplierFolders = @("_Kimble automatically creates Supplier & Subcontractor folders","Background")
$listOfLeadProjSubFolders = @("Admin & contracts", "Analysis","Data & refs","Meetings","Proposal","Reports","Summary (marketing) - end of project")
$defaultProjectFilesToCopy = @(@{"fromList"="/sites/Resources-HealthSafetyGBR";"from"="/sites/Resources-HealthSafetyGBR/Shared Documents/Risk assessments, Safe Systems of Work & emergency plans/Risk Assessments/Anthesis UK Project Risk Assessment.xlsx";"to"="/Admin & contracts/";"conditions"="UK"})

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
$suppliersCacheFile = "kimbleSuppliers.csv"

#Set Variables based on what we're syncing
if($whatToSync -match "Clients" -or $whatToSync -match "Project"){
    $spoSite = "/clients"
    $spoListName = "Kimble Clients"
    $spoProjectListName = "Kimble Projects"
    $accountsCacheFile = $clientsCacheFile
    $arrayOfSubfolders = $listOfClientFolders
    $termSetName = "Clients"
    }
elseif($whatToSync -match "Suppliers"){
    $spoSite = "/subs"
    $spoListName = "Kimble Suppliers"
    $accountsCacheFile = $suppliersCacheFile
    $arrayOfSubfolders = $listOfSupplierFolders
    $termSetName = "Subcontractors"
    }
else{}

Connect-PnPOnline –Url $($webUrl+$spoSite) –Credentials $adminCreds #-RequestTimeout 7200000

#region functions
function new-projectFolder($spoKimbleProjectList, $spoKimbleProjectListItem, $accountsCacheHashTable, $arrayOfProjectSubfolders, $recreateSubFolderOverride, $adminCreds, $fullLogPathAndName){
    if($verboseLogging){Write-Host -ForegroundColor Cyan "new-projectFolder($($spoKimbleProjectList.Title), $($spoKimbleProjectListItem.Name), `$accountsCacheHashTable, $($arrayOfProjectSubfolders -join ", "), $recreateSubFolderOverride=`$recreateSubFolderOverride)"}
    try{
        log-action -myMessage "Retrieving Client Library [$($accountsCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["Name"])]" -logFile $fullLogPathAndName
        $clientLibrary = get-spoClientLibrary -clientName $($accountsCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["Name"]) -clientLibraryGuid $($accountsCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["LibraryId"]) -fullLogPathAndName $fullLogPathAndName -adminCreds $adminCreds -verboseLogging $verboseLogging
        if(!$clientLibrary){
            #Do something clever and add the Client Library. Or we can do nothing and wait for this to fix itself on the next run once ClientCache is updated again.
            }
        else{
            #Create the Project Folder
            log-result -myMessage "SUCCESS: $($clientLibrary.RootFolder.ServerRelativeUrl) retrieved" -logFile $fullLogPathAndName
            log-action -myMessage "Creating Project Folder [$($spoKimbleProjectListItem.Name)] in Client Library [$($accountsCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["Name"])]" -logFile $fullLogPathAndName
            if($verboseLogging){Write-Host -ForegroundColor DarkCyan "add-spoLibrarySubfolders -pnpList $($clientLibrary.Title) -arrayOfSubfolderNames @($($spoKimbleProjectListItem.Name)) -recreateIfNotEmpty $recreateSubFolderOverride"}
            $projectFolder = add-spoLibrarySubfolders -pnpList $clientLibrary -arrayOfSubfolderNames @($clientLibrary.RootFolder.ServerRelativeUrl+"/"+$(sanitise-forPnpSharePoint $spoKimbleProjectListItem.Name)) -recreateIfNotEmpty $recreateSubFolderOverride -spoCredentials $adminCreds -verboseLogging $verboseLogging
            #Create the Project Subfolders
            if($projectFolder){
                log-result "SUCCESS: Project Folder [$projectFolder] created" -logFile $fullLogPathAndName
                log-action -myMessage "Creating Subfolders [$($arrayOfProjectSubfolders -join ",")] in [$($spoKimbleProjectListItem.Name)] in Client Library [$($accountsCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["Name"])]" -logFile $fullLogPathAndName
                $subFolders =@()
                $arrayOfProjectSubfolders | % {$subFolders+= "$($clientLibrary.RootFolder.ServerRelativeUrl)/$(sanitise-forPnpSharePoint $spoKimbleProjectListItem.Name)/$_"}
                $lastSubfolder = add-spoLibrarySubfolders -pnpList $clientLibrary -arrayOfSubfolderNames $subFolders -recreateIfNotEmpty $recreateSubFolderOverride -spoCredentials $adminCreds -verboseLogging $verboseLogging
            #Populate any default files
                if($lastSubfolder){
                    log-result "SUCCESS: Project Subfolder [$lastSubfolder] created" -logFile $fullLogPathAndName
                    log-action -myMessage "Creating default files" -logFile $fullLogPathAndName
                    $defaultProjectFilesToCopy | % {
                        if($verboseLogging){Write-Host -ForegroundColor DarkCyan "copy-spoFile -fromList $($_.fromList) -from $($_.from) -to $($projectFolder.ServerRelativeUrl+$_.to)"}
                        copy-spoFile -fromList $_.fromList -from $_.from -to $($projectFolder.ServerRelativeUrl+$_.to) -spoCredentials $adminCreds
                        }
            #Update the List Item
                    try{
                        log-action -myMessage "Updating [Kimble Projects].[$($spoKimbleProjectListItem.Name)]" -logFile $fullLogPathAndName
                        $updatedValues = @{"IsDirty"=$false;"FolderGUID"=$($newProjectFolder.FieldValues.UniqueId)}
                        log-action "Set-PnPListItem Kimble Projects | $($spoKimbleProjectListItem.Name) [$($spoKimbleProjectListItem.Id)] @{$(stringify-hashTable $updatedValues)}]" -logFile $fullLogPathAndName
                        $duration = Measure-Command {$updatedItem = Set-PnPListItem -List $spoKimbleProjectList.Id -Identity $spoKimbleProjectListItem.SPListItemID -Values $updatedValues}
                        if($updatedItem.FieldValues.IsDirty -eq $false){
                            log-result "SUCCESS: Kimble Projects | $($spoKimbleProjectListItem.Name) is no longer Dirty [$($duration.TotalSeconds) seconds]" -logFile $fullLogPathAndName
                            }
                        else{log-result "FAILED: Could not set Project [$($spoKimbleProjectListItem.Name)].IsDirty = `$false " -logFile $fullLogPathAndName}
                        }
                    catch{
                        #Error Updating list item
                        log-error -myError $_ -myFriendlyMessage "Error updating [Kimble Projects].[$($spoKimbleProjectListItem.Name)] in new-projectFolder" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
                        }
                    $projectFolder #return the new ProjectFolder to show it's worked
                    }
                else{log-result "FAILED: Project Subfolder [$($subFolders[$subFolders.Length-1])] for $($accountsCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["Name"]) was not created" -logFile $fullLogPathAndName}
                }
            else{log-result "FAILED: Project Folder [$($spoKimbleProjectListItem.Name)] for $($accountsCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["Name"]) was not created" -logFile $fullLogPathAndName}
            }
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error in new-projectFolder" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName}
    }
function update-projectFolder($spoKimbleProjectList, $spoKimbleProjectListItem, $accountsCacheHashTable, $arrayOfProjectSubfolders, $recreateSubFolderOverride, $adminCreds, $fullLogPathAndName){
    #Get the ClientLibrary
    log-action -myMessage "Retrieving Client Library [$($accountsCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["Name"])]" -logFile $fullLogPathAndName
    $clientLibrary = get-spoClientLibrary -clientName $($accountsCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["Name"]) -clientLibraryGuid $($accountsCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["LibraryId"]) -fullLogPathAndName $fullLogPathAndName -adminCreds $adminCreds -verboseLogging $verboseLogging

    #Check for Client change
    if(![string]::IsNullOrWhiteSpace($spoKimbleProjectListItem.PreviousKimbleClientId) -and ($spoKimbleProjectListItem.KimbleClientId -ne $spoKimbleProjectListItem.PreviousKimbleClientId)){
        log-result -myMessage "Project has PreviousKimbleClientId - checking whether it needs moving" -logFile $fullLogPathAndName
        $oldClientName = $($accountsCacheHashTable[$spoKimbleProjectListItem.PreviousKimbleClientId]["Name"])
        if($oldClientName){
            log-action -myMessage "Retrieving Previous Client Library [$($accountsCacheHashTable[$spoKimbleProjectListItem.PreviousKimbleClientId]["Name"])]" -logFile $fullLogPathAndName
            $oldClientLibrary = get-spoClientLibrary -clientName $($accountsCacheHashTable[$spoKimbleProjectListItem.PreviousKimbleClientId]["Name"]) -clientLibraryGuid $($accountsCacheHashTable[$spoKimbleProjectListItem.PreviousKimbleClientId]["LibraryId"]) -fullLogPathAndName $fullLogPathAndName -adminCreds $adminCreds -verboseLogging $verboseLogging
            if($oldClientLibrary){
    #Move the folder to the new client
                log-result -myMessage "SUCCESS: Previous Client Library [$($accountsCacheHashTable[$spoKimbleProjectListItem.PreviousKimbleClientId]["Name"])] retrieved" -logFile $fullLogPathAndName
                log-action -myMessage "Looking for Project folder in Previous Client Library [$($accountsCacheHashTable[$spoKimbleProjectListItem.PreviousKimbleClientId]["Name"])]" -logFile $fullLogPathAndName
                $misplacedProjectFolder = get-spoProjectFolder -pnpList $oldClientLibrary -kimbleEngagementCodeToLookFor $(get-kimbleEngagementCodeFromString $spoKimbleProjectListItem.Name -verboseLogging $verboseLogging) -adminCreds $adminCreds -verboseLogging $verboseLogging
                #Hopefully these shouldn't be needed any more as get-spoProjectFolder uses the (hopefully) immutable Kimble Engagement Code to identify the correct folder
                if(!$misplacedProjectFolder){$misplacedProjectFolder = get-spoFolder -pnpList $oldClientLibrary -folderServerRelativeUrl $(format-asServerRelativeUrl -serverRelativeUrl $oldClientLibrary.RootFolder.ServerRelativeUrl -stringToFormat $(sanitise-forPnpSharePoint $spoKimbleProjectListItem.PreviousName)) -folderGuid $spoKimbleProjectListItem.FolderGUID -verboseLogging $verboseLogging -adminCreds $adminCreds}
                if(!$misplacedProjectFolder){$misplacedProjectFolder = get-spoFolder -pnpList $oldClientLibrary -folderServerRelativeUrl $(format-asServerRelativeUrl -serverRelativeUrl $oldClientLibrary.RootFolder.ServerRelativeUrl -stringToFormat $(sanitise-forPnpSharePoint $spoKimbleProjectListItem.Name)) -verboseLogging $verboseLogging -adminCreds $adminCreds} 
                
                if($misplacedProjectFolder -and $clientLibrary){
                    log-result -myMessage "SUCCESS: Project folder [$($misplacedProjectFolder.FieldValues.FileRef)] found in Previous Client [$($accountsCacheHashTable[$spoKimbleProjectListItem.PreviousKimbleClientId]["Name"])]" -logFile $fullLogPathAndName
                    log-action -myMessage "Moving Project folder [$($misplacedProjectFolder.FieldValues.FileRef)] from [$($accountsCacheHashTable[$spoKimbleProjectListItem.PreviousKimbleClientId]["Name"])] to [$($accountsCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["Name"])]" -logFile $fullLogPathAndName
                    $libraryRelavtiveUrl = $misplacedProjectFolder.FieldValues.FileRef.Replace("/clients/","") #Yeah, this isn't a great way to handle it.
                    $movedFolder = Move-PnPFolder -Folder $libraryRelavtiveUrl -TargetFolder $clientLibrary.RootFolder.ServerRelativeUrl
                    if($movedFolder){log-result -myMessage "SUCCESS: Project folder [$($movedFolder.Name)] moved from [$($accountsCacheHashTable[$spoKimbleProjectListItem.PreviousKimbleClientId]["Name"])] to [$($accountsCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["Name"])]" -logFile $fullLogPathAndName}
                    else{log-result -myMessage "FAILED: Project folder [$($misplacedProjectFolder.FieldValues.FileRef)] could not be moved from [$($accountsCacheHashTable[$spoKimbleProjectListItem.PreviousKimbleClientId]["Name"])] to [$($accountsCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["Name"])]" -logFile $fullLogPathAndName}
                    }
                else{log-result -myMessage "Project folder [$($spoKimbleProjectListItem.Name)] not found in Previous Client [$($accountsCacheHashTable[$spoKimbleProjectListItem.PreviousKimbleClientId]["Name"])] - it may already have been moved/deleted" -logFile $fullLogPathAndName}
                }
            }
        }

    if($clientLibrary){
    #Get the ProjectFolder
        log-result -myMessage "SUCCESS: $($clientLibrary.RootFolder.ServerRelativeUrl) retrieved" -logFile $fullLogPathAndName
        log-action -myMessage "Retrieving Project Folder [$($spoKimbleProjectListItem.Name)] in Client Library [$($accountsCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["Name"])]" -logFile $fullLogPathAndName
        $currentProjectFolder = get-spoProjectFolder -pnpList $clientLibrary -kimbleEngagementCodeToLookFor $(get-kimbleEngagementCodeFromString $spoKimbleProjectListItem.Name)
        if(($currentProjectFolder.FieldValues.FileLeafRef -ne $spoKimbleProjectListItem.Name) -and $currentProjectFolder){
        #$misnamedProjectFolder = get-spoFolder -pnpList $clientLibrary -folderServerRelativeUrl $(format-asServerRelativeUrl -serverRelativeUrl $clientLibrary.RootFolder.ServerRelativeUrl -stringToFormat $(sanitise-forPnpSharePoint $spoKimbleProjectListItem.PreviousName)) -folderGuid $spoKimbleProjectListItem.FolderGUID -verboseLogging $verboseLogging
        #if($misnamedProjectFolder){
    #Rename the ProjectFolder
            log-result -myMessage "SUCCESS: Misnamed Project Folder [$($currentProjectFolder.FieldValues.FileRef)] retrieved - will rename to $(sanitise-forPnpSharePoint $spoKimbleProjectListItem.Name)" -logFile $fullLogPathAndName
            $currentProjectFolder.ParseAndSetFieldValue("Title",$(sanitise-forPnpSharePoint $spoKimbleProjectListItem.Name))
            $currentProjectFolder.ParseAndSetFieldValue("FileLeafRef",$(sanitise-forPnpSharePoint $spoKimbleProjectListItem.Name))
            $currentProjectFolder.Update()
            $currentProjectFolder.Context.ExecuteQuery()
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
                    log-result "SUCCESS: Kimble Projects | $($spoKimbleProjectListItem.Name) is no longer Dirty [$($duration.TotalSeconds) seconds]" -logFile $fullLogPathAndName
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
                log-action -myMessage "Override set - recreating Project Folder [$($spoKimbleProjectListItem.Name)] in Client Library [$($accountsCacheHashTable[$spoKimbleProjectListItem.KimbleClientId]["Name"])]" -logFile $fullLogPathAndName
                $newProjectFolder = new-projectFolder -spoKimbleProjectList $spoKimbleProjectList -spoKimbleProjectListItem $spoKimbleProjectListItem -clientCacheHashTable $accountsCacheHashTable -arrayOfProjectSubfolders $arrayOfProjectSubfolders -recreateSubFolderOverride $recreateSubFolderOverride -adminCreds $adminCreds -fullLogPathAndName $fullLogPathAndName
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





#region Process Accounts

#If we need it, get the list of Projects to update *before* we get the list of Clients, so we don't create a race condition where any Projects created while we're processing the Clients queue incorrectly appear to be orphaned
if($whatToSync -match "Projects"){
    Write-Host "If we need it, get the list of Projects to update *before* we get the list of Clients, so we don't create a race condition where any Projects created while we're processing the Clients queue incorrectly appear to be orphaned"
    $dirtyProjects = get-spoKimbleProjectListItems -camlQuery "<View><Query><Where><Eq><FieldRef Name='IsDirty'/><Value Type='Boolean'>1</Value></Eq></Where></Query></View>" -spoCredentials $adminCreds -verboseLogging $verboseLogging
    }

#Get the appropriate [Kimble XXX] List from SPO to see whether our existing cache is out-of-date
try{
    log-action -myMessage "Getting [$spoListName] to check whether it needs recaching" -logFile $fullLogPathAndName 
    $pnpList = Get-PnPList -Identity $spoListName -Includes ContentTypes, LastItemModifiedDate
    if($pnpList){log-result -myMessage "SUCCESS: List retrieved" -logFile $fullLogPathAndName}
    else{log-result -myMessage "FAILURE: List could not be retrieved" -logFile $fullLogPathAndName}
    }
catch{log-error -myError $_ -myFriendlyMessage "Could not retrieve [$spoListName] to check whether it needs recaching" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}

#Retrieve (and update if necessary) the full Clients Cache as we'll need it to set up any new Leads/Projects
Write-Host "Retrieve (and update if necessary) the full Clients Cache as we'll need it to set up any new Leads/Projects"
$accountsCache = cache-spoKimbleAccountsList -pnpList $pnpList -kimbleListCachePathAndFileName $($cacheFilePath+$accountsCacheFile) -fullLogPathAndName $fullLogPathAndName -errorLogPathAndName $errorLogPathAndName -verboseLogging $verboseLogging
    
#Build a hashtable so we can look up Client name by it's KimbleId
Write-Host "Build a hashtable so we can look up Client name by it's KimbleId"
$kimbleAccountHashTable = @{}
$accountsCache | % {$kimbleAccountHashTable.Add($_.Id, @{"Name"=$_.Name;"LibraryId"=$_.LibraryGUID})}


#Process any [Kimble Clients] flagged as IsDirty
Write-Host "Process any [Kimble Clients] flagged as IsDirty"
$dirtyAccounts = $accountsCache | ?{$_.IsDirty -eq $true -and -not ($_.IsDeleted -eq $true -or $_.isMisclassifed -eq $true -or $_.IsOrphaned -eq $true)}
$i = 1
Write-Host "Process [$($dirtyAccounts.Count)] [Kimble Clients] flagged as IsDirty"
$dirtyAccounts | % {
    Write-Progress -Id 1000 -Status "Processing DirtyClients" -Activity "$i/$($dirtyAccounts.Count)" -PercentComplete ($i*100/$dirtyAccounts.Count) #Display the overall progress
    $dirtyAccount = $_
    $duration = Measure-Command {
        log-action -myMessage "************************************************************************" -logFile $fullLogPathAndName
        log-action -myMessage "$whatToSync [$($dirtyAccount.Name)][$i/$($dirtyAccounts.Count)] isDirty!" -logFile $fullLogPathAndName 
        #Check if the Client needs creating
        if(([string]::IsNullOrEmpty($dirtyAccount.PreviousName) -and [string]::IsNullOrEmpty($dirtyAccount.PreviousDescription)) -OR $recreateAllFolders -eq $true){
            log-action -myMessage "$whatToSync [$($dirtyAccount.Name)] looks new - creating new Library" -logFile $fullLogPathAndName
            #Create a new Library and subfolders
            try{new-spoDocumentLibraryAndSubfoldersFromPnpKimbleListItem -pnpList $pnpList -pnpListItem $dirtyAccount -arrayOfSubfolders $arrayOfSubfolders -recreateSubFolderOverride $recreateAllFolders -adminCreds $adminCreds -fullLogPathAndName $fullLogPathAndName}
            catch{log-error $_ -myFriendlyMessage "Failed to create new Library for $($dirtyAccount.Name)" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -smtpServer $smtpServer -mailTo $mailTo -mailFrom $mailFrom}

            #Now try to add the new ClientName to the TermStore
            try{
                log-action "add-termToStore: Kimble | $termSetName | $($dirtyAccount.Name)" -logFile $fullLogPathAndName
                $duration2 = Measure-Command {$newTerm = add-spoTermToStore -termGroup "Kimble" -termSet $termSetName -term $($dirtyAccount.Name) -kimbleId $dirtyAccount.Id -verboseLogging $verboseLogging}
                if($newTerm){log-result "SUCCESS: Kimble | $termSetName | $($dirtyAccount.Name) added to Managed MetaData Term Store [$($duration2.TotalSeconds) secs]" -logFile $fullLogPathAndName}
                }
            catch{log-error $_ -myFriendlyMessage "Failed to add $($dirtyAccount.Title) to Term Store" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -smtpServer $smtpServer -mailFrom $mailFrom -mailTo $mailTo}
            }

        #Otherwise try to update it
        else{
            log-action -myMessage "$whatToSync [$($dirtyAccount.Name)] doesn't look new, so I'm going to try updating it" -logFile $fullLogPathAndName
            try{update-spoDocumentLibraryAndSubfoldersFromPnpKimbleListItem -pnpList $pnpList -pnpListItem $dirtyAccount -arrayOfSubfolders $arrayOfSubfolders -recreateSubFolderOverride $recreateAllFolders -adminCreds $adminCreds -fullLogPathAndName $fullLogPathAndName -verboseLogging $verboseLogging}
            catch{log-error $_ -myFriendlyMessage "Error updating Client [$($dirtyAccount.Name)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -smtpServer $smtpServer -mailFrom $mailFrom -mailTo $mailTo}

            #Then try updating the Managed Metadata
            try{
                log-action -myMessage "Updating Managed Metadata for $($dirtyAccount.Name)" -logFile $fullLogPathAndName
                if($verboseLogging){Write-Host -ForegroundColor DarkCyan "update-spoTerm -termGroup [Kimble] -termSet [$termSetName] -oldTerm [$($dirtyAccount.PreviousName)] -newTerm [$($dirtyAccount.Name)] -kimbleId [$($dirtyAccount.Id)]"}
                $duration2 = Measure-Command {$updatedTerm = update-spoTerm -termGroup "Kimble" -termSet $termSetName -oldTerm $($dirtyAccount.PreviousName) -newTerm $($dirtyAccount.Name) -kimbleId $($dirtyAccount.Id) -verboseLogging $verboseLogging}
                if($updatedTerm){log-result "SUCCESS: Kimble | $termSetName | [$($dirtyAccount.PreviousName)] updated to [$($dirtyAccount.Name)] in Managed MetaData Term Store [$($duration2.TotalSeconds) secs]" -logFile $fullLogPathAndName}
                }
            catch{
                #Failed to update Managed Metadata
                log-error -myError $_ -myFriendlyMessage "Error updating Managed Metadata Term [$($dirtyAccount.PreviousName)] to [$($dirtyAccount.Name)] [$($Error[0].Exception.InnerException.Response)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
                }
            if($(sanitise-forTermStore $updatedTerm.Name) -eq $(sanitise-forTermStore $dirtyAccount.Name)){log-result -myMessage "SUCCESS: Managed Metadata Term [$($dirtyAccount.PreviousName)] updated to [$($dirtyAccount.Name)] [$($duration.TotalSeconds) secs]" -logFile $fullLogPathAndName}
            else{log-result -myMessage "FAILED: Managed Metadata Term [$($dirtyAccount.PreviousName)] did not update to [$($dirtyAccount.Name)]" -logFile $fullLogPathAndName}
            }
        }
    log-result "DirtyClient [$($dirtyAccount.Name)] processed in $($duration.TotalSeconds) seconds" -logFile $fullLogPathAndName
    $i++
    }
#endregion



#region Process Projects
if($whatToSync -match "Projects"){
    #Get [$spoProjectListName] List from SPO
    try{
        log-action -myMessage "Getting [$spoProjectListName]" -logFile $fullLogPathAndName 
        $pnpProjectList = Get-PnPList -Identity $spoProjectListName -Includes ContentTypes, LastItemModifiedDate
        if($pnpProjectList){log-result -myMessage "SUCCESS: List retrieved" -logFile $fullLogPathAndName}
        else{log-result -myMessage "FAILURE: List could not be retrieved" -logFile $fullLogPathAndName}
        }
    catch{log-error -myError $_ -myFriendlyMessage "Could not retrieve [$spoProjectListName]" -fullLogFile $fullLogPathAndName -errorLogFile -doNotLogToEmail $true}

    #Process any [$spoProjectListName] flagged as IsDirty
    #We got $dirtyProjects before the Clients to avoid a race condition
    $i = 1
    $dirtyProjects | % {
        Write-Progress -Id 1000 -Status "Processing DirtyProjects" -Activity "$i/$($dirtyProjects.Count)" -PercentComplete ($i*100/$dirtyProjects.Count) #Display the overall progress

        $dirtyProject = $_
        log-action -myMessage "************************************************************************" -logFile $fullLogPathAndName
        log-action -myMessage "PROJECT [$($dirtyProject.Name)][$i/$($dirtyProjects.Count)] IsDirty" -logFile $fullLogPathAndName
        log-action -myMessage "Checking that Client with Id [$($dirtyProject.KimbleClientId)] is in the Cache" -logFile $fullLogPathAndName
        #First, check that we can identify the corresponding client
        try{
            if ($kimbleAccountHashTable[$dirtyProject.KimbleClientId]){
                log-result "SUCCESS: Client [$($kimbleAccountHashTable[$dirtyProject.KimbleClientId]["Name"])] found in cache LibraryId:[$($kimbleAccountHashTable[$dirtyProject.KimbleClientId]["LibraryId"])]" -logFile $fullLogPathAndName
                }
            else{
                log-result "FAILED: Project Folder [$($dirtyProject.Name)] could not be created because I couldn't identify the Client with Id [$($dirtyProject.KimbleClientId)]" -logFile $fullLogPathAndName
                #This will flood the error logs indefinitely, so mark it as IsDirty = $false. If the user wants it recreated, they can fix the Client then update the Project
                try{
                    log-action -myMessage "Updating [$spoProjectListName].[$($dirtyProject.Name)]" -logFile $fullLogPathAndName
                    $updatedValues = @{"IsDirty"=$false}
                    log-action "Set-PnPListItem [$spoProjectListName] | $($dirtyProject.Name) [$($dirtyProject.Id)] @{$(stringify-hashTable $updatedValues)}]" -logFile $fullLogPathAndName
                    $duration = Measure-Command {$updatedItem = Set-PnPListItem -List $pnpProjectList.Id -Identity $dirtyProject.SPListItemID -Values $updatedValues}
                    if($updatedItem.FieldValues.IsDirty -eq $false){
                        log-result "SUCCESS: [$spoProjectListName]  | $($dirtyProject.Name) is no longer Dirty [$($duration.TotalSeconds) seconds]" -logFile $fullLogPathAndName
                        }
                    else{log-result "FAILED: Could not set Project [$($dirtyProject.Name)].IsDirty = `$false" -logFile $fullLogPathAndName}
                    return
                    }
                catch{
                    #Error Updating list item
                    log-error -myError $_ -myFriendlyMessage "Error updating [$spoProjectListName] | [$($dirtyProject.Name)] to .isDirty = `$false after failing to identify the associated Client (ClientId [$($dirtyProject.KimbleClientId)]). That's a lot of problem. :/" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
                    }
                }
            }
        catch{
            #Couldn't look up Client Name in Hash Table
            log-error -myError $_ -myFriendlyMessage "Error looking up Client ID [$($dirtyProject.KimbleClientId)]. Project is [$($dirtyProject.Name)] for further troubleshooting. " -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
            }

#This is where we add special rules for different sortos of Projects
        if($dirtyProject.DoNotProcess){#Some Projects shouldn't have folders set up. Just mark it as IsDirty = $false
            try{
                #Update LIstItem
                $updatedValues = @{"IsDirty"=$false}
                log-action "Set-PnPListItem [$spoProjectListName] | $($dirtyProject.Name) [$($dirtyProject.Id)] @{$(stringify-hashTable $updatedValues)}]" -logFile $fullLogPathAndName
                $duration = Measure-Command {$updatedItem = Set-PnPListItem -List $pnpProjectList.Id -Identity $dirtyProject.SPListItemID -Values $updatedValues}
                if($updatedItem.FieldValues.IsDirty -eq $false){
                    log-result "SUCCESS: [$spoProjectListName] | $($dirtyProject.Name) is no longer Dirty [$($duration.TotalSeconds) seconds]" -logFile $fullLogPathAndName
                    }
                else{log-result "FAILED: Could not set Projects [$($dirtyProject.Name)].IsDirty = `$true " -logFile $fullLogPathAndName}
                }
            catch{
                #Error Updating list item
                log-error -myError $_ -myFriendlyMessage "Error updating [$spoProjectListName].[$($dirtyProject.Name)] in if(`$dirtyProject.DoNotProcess)" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
                }

            }
#This is the standard behaviour for Projects
        elseif(!$dirtyProject.PreviousName -and (!$dirtyProject.PreviousKimbleClientId -or $dirtyProject.PreviousKimbleClientId -eq $dirtyProject.KimbleClientId)){ 
            #If it looks like a new Project, create a new folder tree under the Client Library
            if($($kimbleAccountHashTable[$dirtyProject.KimbleClientId])){log-action -myMessage "PROJECT [$($dirtyProject.Name)] for client [$($kimbleAccountHashTable[$dirtyProject.KimbleClientId]["Name"])] looks new - creating subfolders!" -logFile $fullLogPathAndName}
            else{log-action -myMessage "PROJECT [$($dirtyProject.Name)] looks new, but I can't work out the Client. Creating subfolders anyway, but this probably won't work" -logFile $fullLogPathAndName}
        
            try{
                if($verboseLogging){Write-Host -ForegroundColor DarkCyan "new-projectFolder -spoKimbleProjectList $($pnpProjectList.Title) -spoKimbleProjectListItem $($dirtyProject.Name) -clientCacheHashTable `$kimbleAccountHashTable -arrayOfProjectSubfolders $($listOfLeadProjSubFolders -join ", ")"}
                $duration = Measure-Command {$newProjectFolder = new-projectFolder -spoKimbleProjectList $pnpProjectList -spoKimbleProjectListItem $dirtyProject -accountsCacheHashTable $kimbleAccountHashTable -arrayOfProjectSubfolders $listOfLeadProjSubFolders -adminCreds $adminCreds -fullLogPathAndName $fullLogPathAndName }
                if($newProjectFolder){
                    log-result -myMessage "SUCCESS: Project Folder [$($newProjectFolder.FieldValues.FileRef)] created successfully [$($duration.TotalSeconds) secs]!" -logFile $fullLogPathAndName}
                else{log-result "FAILED: Project Folder [$($dirtyProject.Name)] for [$($kimbleAccountHashTable[$dirtyProject.KimbleClientId]["Name"])] was not created" -logFile $fullLogPathAndName}
                }
            catch{
                #Failed to create new Project Folder
                log-error -myError $_ -myFriendlyMessage "Error creating new Project Folder [$($dirtyProject.Name)] for [$($kimbleAccountHashTable[$dirtyProject.KimbleClientId]["Name"])]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
                }
            }
        else{#Otherwise try updating it
            try{
                log-action -myMessage "PROJECT [$($dirtyProject.Name)] for client [$($kimbleAccountHashTable[$dirtyProject.KimbleClientId]["Name"])] looks like it needs updating!" -logFile $fullLogPathAndName
                $duration = Measure-Command {$updatedProjectFolder = update-projectFolder -spoKimbleProjectList $pnpProjectList -spoKimbleProjectListItem $dirtyProject -accountsCacheHashTable $kimbleAccountHashTable -arrayOfProjectSubfolders $listOfLeadProjSubFolders -recreateSubFolderOverride $recreateAllFolders -adminCreds $adminCreds -fullLogPathAndName $fullLogPathAndName}
                if($updatedProjectFolder){
                    log-result -myMessage "SUCCESS - Project Folder [$($dirtyProject.Name)] updated successfully [$($duration.TotalSeconds) secs]!" -logFile $fullLogPathAndName
                    }
                else{log-result "FAILED - Project Folder [$($dirtyProject.Name)] for [$($kimbleAccountHashTable[$dirtyProject.KimbleClientId]["Name"])] was not updated" -logFile $fullLogPathAndName}
                }
            catch{
                #Error updating Project Folder
                log-error -myError $_ -myFriendlyMessage "Error updating project folder [$($dirtyProject.Name)] for [$($kimbleAccountHashTable[$dirtyProject.KimbleClientId]["Name"]))]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
                }
            }
        $i++
        }
    }


#endregion
Stop-Transcript