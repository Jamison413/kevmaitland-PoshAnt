param(
    # Specifies whether we are updating Clients or Suppliers.
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Clients", "Suppliers","Projects","ClientsProjects")]
    [string]$whatToSync
    )
    $verboseLogging = $true

$logFileLocation = "C:\ScriptLogs\"
$logFileName = "sync-kimbleSqlObjectsToSpoObjects"
$fullLogPathAndName = $logFileLocation+$logFileName+"_$whatToSync`_FullLog_$(Get-Date -Format "yyMMdd").log"
$errorLogPathAndName = $logFileLocation+$logFileName+"_$whatToSync`_ErrorLog_$(Get-Date -Format "yyMMdd").log"
if($PSCommandPath){
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_$whatToSync`_Transcript_$(Get-Date -Format "yyMMdd").log"
    Start-Transcript $transcriptLogName -Append
    }

Import-Module _PS_Library_GeneralFunctionality
Import-Module _PNP_Library_SPO
Import-Module SharePointPnPPowerShellOnline

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

$sqlAccountsTableName = "SUS_Kimble_Accounts"
$sqlEngagementsTableName = "SUS_Kimble_Engagements"

#Set Variables based on what we're syncing
if($whatToSync -match "Clients" -or $whatToSync -match "Project"){
    $spoSite = "/clients"
    $arrayOfSubfolders = $listOfClientFolders
    $termSetName = "Clients"
    }
elseif($whatToSync -match "Suppliers"){
    $spoSite = "/subs"
    $arrayOfSubfolders = $listOfSupplierFolders
    $termSetName = "Subcontractors"
    }
else{}

Connect-PnPOnline –Url $($webUrl+$spoSite) –Credentials $adminCreds #-RequestTimeout 7200000
$sqlDbConn = connect-toSqlServer -SQLServer "sql.sustain.co.uk" -SQLDBName "SUSTAIN_LIVE"


#region Process Accounts
log-action -myMessage "" -logFile $fullLogPathAndName
log-action -myMessage "" -logFile $fullLogPathAndName
log-action -myMessage "************************************************************************" -logFile $fullLogPathAndName
log-action -myMessage "Starting new synchronisation run for $whatToSync" -logFile $fullLogPathAndName
log-action -myMessage "************************************************************************" -logFile $fullLogPathAndName

#If we need it, get the list of Projects to update *before* we get the list of Clients, so we don't create a race condition where any Projects created while we're processing the Clients queue incorrectly appear to be orphaned
if($whatToSync -match "Projects"){
    log-action -myMessage "Retrieving DirtyProjects" -logFile $fullLogPathAndName
    Write-Verbose "If we need it, get the list of Projects to update *before* we get the list of Clients, so we don't create a race condition where any Projects created while we're processing the Clients queue incorrectly appear to be orphaned"
    $dirtyProjects = get-allFocalPointCachedKimbleEngagements -dbConnection $sqlDbConn -pWhereStatement "WHERE IsDirty = 1 AND SuppressFolderCreation = 0"
    $dirtyProjects = $dirtyProjects | Sort-Object KimbleOne__Reference__c -Descending #Prioritise the latest ones first as they're the least likely to have been created
    log-result "[$($dirtyProjects.Count)] DirtyProjects retrieved!" -logFile $fullLogPathAndName
    #$dirtyProjects = get-spoKimbleProjectListItems -camlQuery "<View><Query><Where><Eq><FieldRef Name='IsDirty'/><Value Type='Boolean'>1</Value></Eq></Where></Query></View>" -spoCredentials $adminCreds -verboseLogging $verboseLogging
    }
if($whatToSync -match "Clients"){
    log-action -myMessage "Retrieving DirtyAccounts (Clients)" -logFile $fullLogPathAndName
    $dirtyAccounts = get-allFocalPointCachedKimbleAccounts -dbConnection $sqlDbConn -pWhereStatement "WHERE IsDirty = 1 AND (Type LIKE '%Client%' OR KimbleOne__IsCustomer__c = 1)"
    log-result "[$($dirtyAccounts.Count)] DirtyAccounts retrieved!" -logFile $fullLogPathAndName
    #$dirtyProjects = get-spoKimbleProjectListItems -camlQuery "<View><Query><Where><Eq><FieldRef Name='IsDirty'/><Value Type='Boolean'>1</Value></Eq></Where></Query></View>" -spoCredentials $adminCreds -verboseLogging $verboseLogging
    }
elseif($whatToSync -match "Supplier" -or $whatToSync -match "Sub"){
    log-action -myMessage "Retrieving DirtyAccounts (Suppliers)" -logFile $fullLogPathAndName
    $dirtyAccounts = get-allFocalPointCachedKimbleAccounts -dbConnection $sqlDbConn -pWhereStatement "WHERE IsDirty = 1 AND (Type LIKE '%Supplier%' OR Type LIKE '%Sub%')"
    log-result "[$($dirtyAccounts.Count)] DirtyAccounts retrieved!" -logFile $fullLogPathAndName
    }


#Process any [SUS_Kimble_Accounts] flagged as IsDirty
Write-Verbose "Process any [SUS_Kimble_Accounts] flagged as IsDirty"
$i = 1
Write-Verbose "Process [$($dirtyAccounts.Count)] [Kimble Clients] flagged as IsDirty"
$dirtyAccounts | % {
    Write-Progress -Id 1000 -Status "Processing DirtyAccounts" -Activity "$i/$($dirtyAccounts.Count)" -PercentComplete ($i*100/$dirtyAccounts.Count) #Display the overall progress
    $dirtyAccount = $_
    $duration = Measure-Command {
        log-action -myMessage "************************************************************************" -logFile $fullLogPathAndName
        log-action -myMessage "$whatToSync [$($dirtyAccount.Name)][$i/$($dirtyAccounts.Count)] isDirty!" -logFile $fullLogPathAndName 
        #Check if the Client needs creating
        #if(([string]::IsNullOrEmpty($dirtyAccount.PreviousName) -and [string]::IsNullOrEmpty($dirtyAccount.PreviousDescription)) -OR $recreateAllFolders -eq $true){
        if([string]::IsNullOrEmpty($dirtyAccount.DocumentLibraryGuid) -OR ($recreateAllFolders -eq $true)){
            log-action -myMessage "$whatToSync [$($dirtyAccount.Name)] looks new - creating new Library" -logFile $fullLogPathAndName
            #Create a new Library and subfolders
            try{$newLibrary = new-spoDocumentLibraryAndSubfoldersFromSqlKimbleListItem -sqlDbConn $sqlDbConn -sqlKimbleAccount $dirtyAccount -arrayOfSubfolders $arrayOfSubfolders -recreateSubFolderOverride $recreateAllFolders -adminCreds $adminCreds -fullLogPathAndName $fullLogPathAndName -errorLogPathAndName $errorLogPathAndName -Verbose}
            catch{log-error $_ -myFriendlyMessage "Error creating new Library for [$($dirtyAccount.Name)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -smtpServer $smtpServer -mailTo $mailTo -mailFrom $mailFrom}

            if($newLibrary){ #If that worked, mark the Account as IsDirty = $false
                $sql = "UPDATE SUS_Kimble_Accounts SET IsDirty = 0, DocumentLibraryGuid = '$($newLibrary.id.Guid)' WHERE ID = '$($dirtyAccount.Id)'"
                $subDuration = Measure-Command {$result = Execute-SQLQueryOnSQLDB -query $sql -queryType NonQuery -sqlServerConnection $sqlDbConn}
                if($result -eq 1){
                    log-result "SUCCESS: [SUS_Kimble_Accounts] | [$($dirtyAccount.Name)] is no longer Dirty [$($subDuration.TotalSeconds) seconds]" -logFile $fullLogPathAndName
                    }
                else{log-result "FAILED: Could not set [SUS_Kimble_Accounts] | [$($dirtyAccount.Name)].IsDirty = `$false [`$result = $result] [$($subDuration.TotalSeconds) seconds]" -logFile $fullLogPathAndName}
                }

            #Now try to add the new ClientName to the TermStore
            try{
                log-action "add-termToStore: [Kimble] | [$termSetName] | [$($dirtyAccount.Name)]" -logFile $fullLogPathAndName
                $subDuration = Measure-Command {$newTerm = add-spoTermToStore -termGroup "Kimble" -termSet $termSetName -term $($dirtyAccount.Name) -kimbleId $dirtyAccount.Id -verboseLogging $verboseLogging}
                }
            catch{log-error $_ -myFriendlyMessage "Failed to add $($dirtyAccount.Title) to Term Store" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -smtpServer $smtpServer -mailFrom $mailFrom -mailTo $mailTo}
            if($newTerm){log-result "SUCCESS: [Kimble] | [$termSetName] | [$($dirtyAccount.Name)] added to Managed MetaData Term Store [$($subDuration.TotalSeconds) secs]" -logFile $fullLogPathAndName}
            else{log-result "FAILED: [Kimble] | [$termSetName] | [$($dirtyAccount.Name)] added to Managed MetaData Term Store [$($subDuration.TotalSeconds) secs]" -logFile $fullLogPathAndName}
            }

        #Otherwise try to update it
        else{
            log-action -myMessage "$whatToSync [$($dirtyAccount.Name)] doesn't look new, so I'm going to try updating it" -logFile $fullLogPathAndName
            try{$updatedLibrary = update-spoDocumentLibraryAndSubfoldersFromSqlKimbleListItem -sqlKimbleAccount $dirtyAccount -sqlDbConn $sqlDbConn -arrayOfSubfolders $arrayOfSubfolders -recreateSubFolderOverride $recreateAllFolders -adminCreds $adminCreds -fullLogPathAndName $fullLogPathAndName -errorLogPathAndName $errorLogPathAndName}
            catch{log-error $_ -myFriendlyMessage "Error updating Client [$($dirtyAccount.Name)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -smtpServer $smtpServer -mailFrom $mailFrom -mailTo $mailTo}

            if($updatedLibrary){
                try{
                    #Update the List Item
                    $sql = "UPDATE SUS_Kimble_Accounts SET IsDirty = 0, DocumentLibraryGuid = '$($updatedLibrary.id.Guid)'"
                    if(![string]::IsNullOrWhiteSpace($dirtyAccount.PreviousName) -and ($(sanitise-forSqlValue -value $dirtyAccount.Name -dataType String) -ne $(sanitise-forSqlValue -value $dirtyAccount.PreviousName -dataType String))){$sql += ", PreviousName = '$(sanitise-forSqlValue -value $dirtyAccount.Name -dataType String)'"} #If the Name has changed, overwrite the old one with the new one to indicate that this has been processed (a trigger on [SUS_Kimble_Accounts] preserves this data)
                    if(![string]::IsNullOrWhiteSpace($dirtyAccount.PreviousDescription) -and ((sanitise-forSqlValue -value $dirtyAccount.Description -dataType HTML) -ne (sanitise-forSqlValue -value $dirtyAccount.Description -dataType HTML))){$sql += ", PreviousName = '$(sanitise-forSqlValue -value $dirtyAccount.Description -dataType HTML)'"} #If the Description has changed, overwrite the old one with the new one to indicate that this has been processed (a trigger on [SUS_Kimble_Accounts] preserves this data)
                    $sql +=" WHERE ID = '$($dirtyAccount.Id)'"
                    $result = Execute-SQLQueryOnSQLDB -query $sql -queryType NonQuery -sqlServerConnection $sqlDbConn
                    if($result -eq 1){log-result "SUCCESS: [SUS_Kimble_Accounts] | [$($dirtyAccount.Name)] is no longer Dirty [$($duration.TotalSeconds) seconds]" -logFile $fullLogPathAndName}
                    else{log-result "FAILED: Could not set [SUS_Kimble_Accounts] | [$($dirtyAccount.Name)].IsDirty = `$false [`$result = $result]" -logFile $fullLogPathAndName}
                    }
                catch{
                    #Failed to update SPListItem
                    log-error -myError $_ -myFriendlyMessage "Error updating [$($pnpList.Title)] | [$($pnpListItem.Name)] - this is still marked as IsDirty=`$true :( [$($Error[0].Exception.InnerException.Response)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
                    }
                }
            #Then try updating the Managed Metadata if the name might have changed
            if(![string]::IsNullOrWhiteSpace($dirtyAccount.PreviousName) -and $dirtyAccount.Name -ne $dirtyAccount.PreviousName){
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
        }
    log-result "DirtyClient [$($dirtyAccount.Name)] processed in $($duration.TotalSeconds) seconds" -logFile $fullLogPathAndName
    $i++
    }
#endregion



#region Process [SUS_Kimble_Engagements]
if($whatToSync -match "Projects"){
    #Process any [$spoProjectListName] flagged as IsDirty
    #We got $dirtyProjects before the Clients to avoid a race condition
    $i = 1
    $dirtyProjects | % {
        Write-Progress -Id 1000 -Status "Processing DirtyProjects" -Activity "$i/$($dirtyProjects.Count)" -PercentComplete ($i*100/$dirtyProjects.Count) #Display the overall progress

        $dirtyProject = $_
        log-action -myMessage "************************************************************************" -logFile $fullLogPathAndName
        log-action -myMessage "PROJECT [$($dirtyProject.Name)][$i/$($dirtyProjects.Count)] IsDirty" -logFile $fullLogPathAndName
        log-action -myMessage "Checking that Client with Id [$($dirtyProject.KimbleOne__Account__c)] is in [SUS_Kimble_Accounts]" -logFile $fullLogPathAndName
        #First, check that we can identify the corresponding client
        try{
            $sql = "SELECT Id, Name, DocumentLibraryGuid FROM SUS_Kimble_Accounts WHERE Id = '$($dirtyProject.KimbleOne__Account__c)'"
            $clientForThisProject = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $sqlDbConn
            }
        catch{
            #Couldn't look up Client Name in SQL Table
            log-error -myError $_ -myFriendlyMessage "Error looking up Client ID [$($dirtyProject.KimbleClientId)]. Project is [$($dirtyProject.Name)] for further troubleshooting. " -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
            }

        if(!$clientForThisProject){
            log-result "FAILED: Project Folder [$($dirtyProject.Name)] could not be created because I couldn't identify the Client with Id [$($dirtyProject.KimbleOne__Account__c)]" -logFile $fullLogPathAndName
            #This will flood the error logs indefinitely, so mark it as IsDirty = $false. If the user wants it recreated, they can fix the Client then update the Project
            try{
                log-action -myMessage "Updating [SUS_Kimble_Engagements].[$($dirtyProject.Name)]" -logFile $fullLogPathAndName
                $sql = "UPDATE SUS_Kimble_Engagements SET IsDirty = 0 WHERE Id = '$($dirtyProject.Id)'"
                $result =  Execute-SQLQueryOnSQLDB -query $sql -queryType NonQuery -sqlServerConnection $sqlDbConn
                if($result -eq 1){
                    log-result "SUCCESS: [SUS_Kimble_Engagements] | $($dirtyProject.Name) is no longer Dirty" -logFile $fullLogPathAndName
                    }
                else{log-result "FAILED: Could not set [SUS_Kimble_Engagements] | [$($dirtyProject.Name)].IsDirty = `$false" -logFile $fullLogPathAndName}
                return
                }
            catch{
                #Error Updating list item
                log-error -myError $_ -myFriendlyMessage "Error updating [SUS_Kimble_Engagements] | [$($dirtyProject.Name)] to .isDirty = `$false after failing to identify the associated Client (ClientId [$($dirtyProject.KimbleOne__Account__c)]). That's a lot of problem. :/" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
                return
                }
            }
        else{log-result "SUCCESS: Account [$($clientForThisProject.Name)] for [$($dirtyProject.Name)] found in [SUS_Kimble_Engagements]" -logFile $fullLogPathAndName}

        log-action -myMessage "Retrieving Client DocLib for [$($clientForThisProject.Name )]" -logFile $fullLogPathAndName
        try{$clientLibrary = get-spoClientLibrary -clientName $clientForThisProject.Name -clientLibraryGuid $clientForThisProject.DocumentLibraryGuid}
        catch{log-error -myError $_ -myFriendlyMessage "Error retrieving Client DocLib [$($clientForThisProject.Name)]. Project is [$($dirtyProject.Name)] for further troubleshooting. " -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName}
        if(!$clientLibrary){
            #If we can't find the Client Library, we're not going to be able to do _anything_ (see !$clientForThisProject above)
            log-result "FAILED: Project Folder [$($dirtyProject.Name)] could not be created because I couldn't find a DocLib for Client [$($clientForThisProject.Name)]" -logFile $fullLogPathAndName
            try{
                log-action -myMessage "Updating [SUS_Kimble_Engagements].[$($dirtyProject.Name)]" -logFile $fullLogPathAndName
                $sql = "UPDATE SUS_Kimble_Engagements SET IsDirty = 0 WHERE Id = '$($dirtyProject.Id)'"
                $result =  Execute-SQLQueryOnSQLDB -query $sql -queryType NonQuery -sqlServerConnection $sqlDbConn
                if($result -eq 1){
                    log-result "SUCCESS: [SUS_Kimble_Engagements] | $($dirtyProject.Name) is no longer Dirty" -logFile $fullLogPathAndName
                    }
                else{log-result "FAILED: Could not set [SUS_Kimble_Engagements] | [$($dirtyProject.Name)].IsDirty = `$false" -logFile $fullLogPathAndName}
                return
                }
            catch{
                #Error Updating list item
                log-error -myError $_ -myFriendlyMessage "Error updating [SUS_Kimble_Engagements] | [$($dirtyProject.Name)] to .isDirty = `$false after failing to find DocLib for Client $($clientForThisProject.Name)" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
                return
                }
            }
        else{log-result "SUCCESS: Client DocLib [$($clientLibrary.RootFolder.ServerRelativeUrl)] for [$($dirtyProject.Name)] retrieved" -logFile $fullLogPathAndName}
#This is where we add special rules for different sorts of Projects
#
# We no longer look for SuppressFolderCreation here (because simply don't process aby Engagements marked SuppressFolderCreation = $true)
# We could check for ohter criteria and _set_ SuppressFolderCreation = $true though...
#

#This is the standard behaviour for Projects
        #if([string]::IsNullOrWhiteSpace($dirtyProject.PreviousName) -and ([string]::IsNullOrWhiteSpace($dirtyProject.PreviousKimbleClientId) -or ($dirtyProject.PreviousKimbleClientId -eq $dirtyProject.KimbleOne__Account__c))){ 
        if([string]::IsNullOrWhiteSpace($dirtyProject.FolderGuid) -or $recreateAllFolders -eq $true){ 
            #If it looks like a new Project, create a new folder tree under the Client Library
            log-action -myMessage "PROJECT [$($dirtyProject.Name)] for client [$($clientForThisProject.Name)] looks new - creating main Project Folder!" -logFile $fullLogPathAndName
            $duration = Measure-Command {$newProjectFolder = add-spoLibrarySubfolders -pnpList $clientLibrary -arrayOfSubfolderNames @($clientLibrary.RootFolder.ServerRelativeUrl+"/"+$(sanitise-forPnpSharePoint $dirtyProject.Name)) -recreateIfNotEmpty $recreateSubFolderOverride -spoCredentials $adminCreds -verboseLogging $verboseLogging}
            if($newProjectFolder){
                #Now create the subfolders
                log-result "SUCCESS: Project Folder [$($newProjectFolder.ServerRelativeUrl)] created" -logFile $fullLogPathAndName
                log-action -myMessage "Creating Subfolders [$($listOfLeadProjSubFolders -join ",")] in [$($dirtyProject.Name)] in Client Library [$($clientForThisProject.Name)]" -logFile $fullLogPathAndName
                $subFolders =@()
                $listOfLeadProjSubFolders | % {$subFolders+= "$($clientLibrary.RootFolder.ServerRelativeUrl)/$(sanitise-forPnpSharePoint $dirtyProject.Name)/$_"}
                try{$lastSubfolder = add-spoLibrarySubfolders -pnpList $clientLibrary -arrayOfSubfolderNames $subFolders -recreateIfNotEmpty $recreateSubFolderOverride -spoCredentials $adminCreds -verboseLogging $verboseLogging}
                catch{log-error -myError $_ -myFriendlyMessage "Error creating new main Project Folder [$($dirtyProject.Name)] for [$($clientForThisProject.Name)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName}
                #Populate any default files
                if($lastSubfolder){
                    log-result "SUCCESS: Project Subfolder [$($lastSubfolder.ServerRelativeUrl)] created" -logFile $fullLogPathAndName
                    log-action -myMessage "Creating default files" -logFile $fullLogPathAndName
                    $defaultProjectFilesToCopy | % {
                        try{
                            if($verboseLogging){Write-Host -ForegroundColor DarkCyan "copy-spoFile -fromList $($_.fromList) -from $($_.from) -to $($newProjectFolder.ServerRelativeUrl+$_.to)"}
                            $result = copy-spoFile -fromList $_.fromList -from $_.from -to $($newProjectFolder.ServerRelativeUrl+$_.to) -spoCredentials $adminCreds
                            }
                        catch{
                            if($_.Exception.ServerErrorCode -eq -2130575257){log-result "FAILED: (but that's okay): $($_.Exception)" -logFile $fullLogPathAndName}
                            else{log-error -myError $_ -myFriendlyMessage "Error creating new Project Folders [$($dirtyProject.Name)] for [$($clientForThisProject.Name)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName}
                            }
                        if($result){log-result -myMessage "SUCCESS: File [$(Split-Path $_.from -Leaf)] is in [$($newProjectFolder.ServerRelativeUrl+$_.to)]" -logFile $fullLogPathAndName}
                        else{log-result -myMessage "FAILED: File [$($_.from)] was not copied to [$($newProjectFolder.ServerRelativeUrl+$_.to)]" -logFile $fullLogPathAndName}
                        }
                    }
                else{log-result "FAILED: Project Subfolders [$($dirtyProject.Name)] for [$($clientForThisProject.Name)] were not created" -logFile $fullLogPathAndName}
                }
            else{log-result "FAILED: Main Project Folder [$($dirtyProject.Name)] for [$($clientForThisProject.Name)] was not created" -logFile $fullLogPathAndName}
            $finalProjectFolder = $newProjectFolder #We need the Project Folder later regardless of whether it's new or updated
            }

        else{#Otherwise try updating it
            #Check for Client change as moving these is really mportant to avoid duplicates
            if($dirtyProject.PreviousKimbleClientId){
                log-action -myMessage "Retrieving Previous Client Library [$($previousClientForThisProject.Name)]" -logFile $fullLogPathAndName
                $sql = "SELECT Id, Name, DocumentLibraryGuid FROM SUS_Kimble_Accounts WHERE Id = '$($dirtyProject.PreviousKimbleClientId)'"
                $previousClientForThisProject = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $sqlDbConn
                $previousClientLibrary = get-spoClientLibrary -clientName $previousClientForThisProject.Name
                if($previousClientLibrary){
                    #Look for the Project folder on the old DocLib first
                    log-result -myMessage "SUCCESS: Previous Client Library [$($previousClientForThisProject.Name)] retrieved" -logFile $fullLogPathAndName
                    log-action -myMessage "Looking for Project folder in Previous Client Library [$($previousClientForThisProject.Name)]" -logFile $fullLogPathAndName
                    $misplacedProjectFolder = get-spoProjectFolder -pnpList $previousClientLibrary -folderGuid $dirtyProject.FolderGuid -kimbleEngagementCodeToLookFor $(get-kimbleEngagementCodeFromString $dirtyProject.Name -verboseLogging $verboseLogging) -adminCreds $adminCreds -verboseLogging $verboseLogging
                    if($misplacedProjectFolder){
                        #Everything looks good - let's move the Project Folder to the new Client
                        log-result -myMessage "SUCCESS: Project folder [$($misplacedProjectFolder.FieldValues.FileRef)] found in Previous Client [$($previousClientForThisProject.Name)]" -logFile $fullLogPathAndName
                        log-action -myMessage "Moving Project folder [$($misplacedProjectFolder.FieldValues.FileRef)] from [$($previousClientForThisProject.Name)] to [$($clientForThisProject.Name)]" -logFile $fullLogPathAndName
                        $libraryRelavtiveUrl = $misplacedProjectFolder.FieldValues.FileRef.Replace("/clients/","") #Yeah, this isn't a great way to handle it.
                        try{$movedFolder = Move-PnPFolder -Folder $libraryRelavtiveUrl -TargetFolder $clientLibrary.RootFolder.ServerRelativeUrl}
                        catch{log-error -myError $_ -myFriendlyMessage "Error moving Project Folder [$libraryRelavtiveUrl] from [$($previousClientForThisProject.Name)] to [$($clientForThisProject.Name)] " -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName}
                        if($movedFolder){log-result -myMessage "SUCCESS: Project folder [$($movedFolder.Name)] moved from [$($previousClientForThisProject.Name)] to [$($clientForThisProject.Name)]" -logFile $fullLogPathAndName}
                        else{log-result -myMessage "FAILED: Project folder [$($misplacedProjectFolder.FieldValues.FileRef)] could not be moved from [$($previousClientForThisProject.Name)] to [$($clientForThisProject.Name)] - something went wrong with the Move-PnpFolder process" -logFile $fullLogPathAndName}
                        }
                    else{log-result -myMessage "Project folder [$($dirtyProject.Name)] not found in Previous Client [$($previousClientForThisProject.Name)] - this is probably a good thing as it may already have been moved/deleted" -logFile $fullLogPathAndName}
                    }
                else{
                    #Well, we can't find a Client DocLib for the old Client, so we can't check whether the Project folder needs moving. Not much more we can do here.
                    log-result -myMessage "A DocLib for Previous Client [$($previousClientForThisProject.Name)] could not be found. Cannot determine whether there is a legacy folder that needs moving." -logFile $fullLogPathAndName
                    }
                }
            #Now we've moved any Project folders that have been assigned to a new client, process any name changes
            log-action -myMessage "Retrieving Project Folder [$($dirtyProject.Name)] in Client Library [$($clientForThisProject.Name)]" -logFile $fullLogPathAndName
            try{$currentProjectFolder = get-spoProjectFolder -pnpList $clientLibrary -folderGuid $dirtyProject.FolderGuid -kimbleEngagementCodeToLookFor $(get-kimbleEngagementCodeFromString $dirtyProject.Name)}
            catch{log-error -myError $_ -myFriendlyMessage "Error retrieving original main Project Folder [$($dirtyProject.Name)] from [$($clientLibrary.RootFolder.ServerRelativeUrl)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName}
            if($currentProjectFolder){
                log-result -myMessage "SUCCESS: [$($currentProjectFolder.FieldValues.FileRef)] retrieved" -logFile $fullLogPathAndName
                if($currentProjectFolder.FieldValues.FileLeafRef -ne $dirtyProject.Name){
                    #Rename the ProjectFolder
                    log-action -myMessage "Updating name of Main Project Folder from [$($currentProjectFolder.FieldValues.FileLeafRef)] to [$($dirtyProject.Name)] in Client Library [$($clientForThisProject.Name)]" -logFile $fullLogPathAndName
                    try{
                        $currentProjectFolder.ParseAndSetFieldValue("Title",$(sanitise-forPnpSharePoint $dirtyProject.Name))
                        $currentProjectFolder.ParseAndSetFieldValue("FileLeafRef",$(sanitise-forPnpSharePoint $dirtyProject.Name))
                        $currentProjectFolder.Update()
                        $currentProjectFolder.Context.ExecuteQuery()
                        }
                    catch{log-error -myError $_ -myFriendlyMessage "Error updating main Project Folder [$($currentProjectFolder.FieldValues.FileRef)] to [$(sanitise-forPnpSharePoint $dirtyProject.Name)] in [$($clientLibrary.RootFolder.ServerRelativeUrl)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName}
                    try{$updatedProjectFolder = get-spoProjectFolder -pnpList $clientLibrary -folderGuid $dirtyProject.FolderGuid -kimbleEngagementCodeToLookFor $(get-kimbleEngagementCodeFromString $dirtyProject.Name)}
                    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving updated main Project Folder [$($dirtyProject.Name)] from [$($clientLibrary.RootFolder.ServerRelativeUrl)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName}
    
                    if($updatedProjectFolder){
                        if($updatedProjectFolder.FieldValues.FileLeafRef -eq $currentProjectFolder.FieldValues.FileLeafRef){log-result -myMessage "SUCCESS: Misnamed Project Folder [$($currentProjectFolder.FieldValues.FileRef)] updated to [$($updatedProjectFolder.FieldValues.FileRef)]" -logFile $fullLogPathAndName}
                        else{log-result -myMessage "FAILED: Misnamed Project Folder [$($currentProjectFolder.FieldValues.FileRef)] did not update to [$($updatedProjectFolder.FieldValues.FileRef)]" -logFile $fullLogPathAndName}
                        }
                    else{log-result -myMessage "FAILED: Could not retrieve Project Folder [$($currentProjectFolder.FieldValues.FileRef)] after updating it to [$(sanitise-forPnpSharePoint $dirtyProject.Name)]. Sorry - I don't know where it went :/" -logFile $fullLogPathAndName}
                    $currentProjectFolder = $updatedProjectFolder #This will null $currentProjectFolder if there was a problem above
                    }
                else{
                    #The Main Project folder name is as expected (some other property update must have triggered this). Nothing to do.
                    }
                }
            else{log-result -myMessage "FAILED: Could not find Main Project folder for [$($dirtyProject.Name)] in [$($clientLibrary.RootFolder.ServerRelativeUrl)]" -logFile $fullLogPathAndName}

            $finalProjectFolder = $currentProjectFolder #We need the Project Folder later regardless of whether it's new or updated
            }

#Update the SqlObject
        log-action -myMessage "Updating [SUS_Kimble_Engagements].[$($dirtyProject.Name)]" -logFile $fullLogPathAndName
        if($finalProjectFolder){
            #The GUID might be in different locations depending whether we've got a PnPListItem or PnPFolder object
            switch($finalProjectFolder.GetType().Name){
                "Folder" {
                    if([string]::IsNullOrWhiteSpace($finalProjectFolder.ListItemAllFields.FieldValues.GUID)){$finalProjectFolder = Get-PnPFolder -Url $finalProjectFolder.ServerRelativeUrl -Includes ListItemAllFields}
                    $finalProjectFolderGuid = $finalProjectFolder.ListItemAllFields.FieldValues.GUID
                    }
                "ListItemCollection" {$finalProjectFolderGuid = $finalProjectFolder.FieldValues.GUID}
                }

            try{
                $sql = "UPDATE SUS_Kimble_Engagements SET IsDirty = 0, FolderGuid = $(sanitise-forSqlValue -value $finalProjectFolderGuid -dataType Guid) WHERE Id = '$($dirtyProject.Id)'"
                $result =  Execute-SQLQueryOnSQLDB -query $sql -queryType NonQuery -sqlServerConnection $sqlDbConn
                if($result -eq 1){
                    log-result "SUCCESS: [SUS_Kimble_Engagements] | $($dirtyProject.Name) is no longer Dirty" -logFile $fullLogPathAndName
                    }
                else{log-result "FAILED: Could not set [SUS_Kimble_Engagements] | [$($dirtyProject.Name)].IsDirty = `$false" -logFile $fullLogPathAndName}
                }
            catch{
                #Error Updating list item
                log-error -myError $_ -myFriendlyMessage "Error updating [SUS_Kimble_Engagements] | [$($dirtyProject.Name)] to .isDirty = `$false after failing to find DocLib for Client $($clientForThisProject.Name)" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
                return
                }
            log-result "SUCCESS: Project Folder [$($finalProjectFolder.ServerRelativeUrl)] was created" -logFile $fullLogPathAndName
            }
        else{
            #We've got a non-specific problem here, so it's _probably_ better to leave the SQL object marked as IsDirty = $true and investigate the underlying cause rather than just mark it as IsDirty = $false and ignore the problem
            }
        $i++
        }
    }

$sqlDbConn.close()
#endregion
Stop-Transcript
