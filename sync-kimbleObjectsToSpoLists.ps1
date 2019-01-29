param(
    # Specifies whether we are updating Clients or Suppliers.
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Clients", "Suppliers","Projects")]
    [string]$objectType 
    )

$logFileLocation = "C:\ScriptLogs\"
if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)){
    $fullLogPathAndName = $logFileLocation+"sync-kimbleObjectsToSpoLists_$objectType`_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = $logFileLocation+"sync-kimbleObjectsToSpoLists_$objectType`_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
else{
    $fullLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_$objectType`_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_$objectType`_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
if($PSCommandPath){
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_$objectType`_Transcript_$(Get-Date -Format "yyMMdd").log"
    Start-Transcript $transcriptLogName -Append
    }

Import-Module _PS_Library_GeneralFunctionality
Import-Module _REST_Library-Kimble.psm1
Import-Module _PNP_Library_SPO

function get-pnpKimbleList($spoListName,$fullLogPathAndName){
    #Get the [/lists/listName] pnpList so we can CRUD things 
    try{
        log-action -myMessage "Getting [$spoListName] so we can CRUD things later" -logFile $fullLogPathAndName 
        $pnpKimbleObjectList = Get-PnPList -Identity $spoListName -Includes ContentTypes, LastItemModifiedDate
        if($pnpKimbleObjectList){log-result -myMessage "SUCCESS: List retrieved" -logFile $fullLogPathAndName}
        else{log-result -myMessage "FAILURE: List could not be retrieved" -logFile $fullLogPathAndName}
        }
    catch{log-error -myError $_ -myFriendlyMessage "Could not retrieve [$spoListName]" -fullLogFile $fullLogPathAndName -errorLogFile -doNotLogToEmail $true}
    $pnpKimbleObjectList
    }

#Get 1st round of Client/Supplier-specific values
if($objectType -imatch "client"){
    $spoSitePath = "/clients"
    $spoListName = "Kimble Clients"
    $soqlQuery = "SELECT Name,Id,Description,Type,KimbleOne__IsCustomer__c,LastModifiedDate,SystemModStamp,CreatedDate,IsDeleted FROM account WHERE ((LastModifiedDate > **PLACEHOLDER**`Z) AND ((KimbleOne__IsCustomer__c = TRUE) OR (Type = 'Client') OR (Type = 'Potential Client')))"
    }
elseif($objectType -imatch "project"){
    $spoSitePath = "/clients"
    $spoListName = "Kimble Projects"
    $soqlQuery = "SELECT Name,Id,KimbleOne__Account__c,LastModifiedDate,SystemModStamp,CreatedDate,IsDeleted,Community__c,Project_Type__c FROM KimbleOne__DeliveryGroup__c WHERE LastModifiedDate > **PLACEHOLDER**`Z"
    }
elseif($objectType -match "supplier"){
    $spoSitePath = "/subs"
    $spoListName = "Kimble Suppliers"
    $soqlQuery = "SELECT Name,Id,Description,Type,KimbleOne__IsCustomer__c,LastModifiedDate,SystemModStamp,CreatedDate,IsDeleted FROM account WHERE (((LastModifiedDate > **PLACEHOLDER**`Z) AND ((Is_Partner__c = TRUE) OR (Type = 'Partner') OR (Type = 'Partner/subcontractor') OR (Type = 'Supplier'))) AND (LastModifiedById <> '00524000001qHf8AAE'))"
    }

##################################
#
#Get ready
#
##################################
$spoWebUrl = "https://anthesisllc.sharepoint.com" 
$sharePointAdmin = "kimblebot@anthesisgroup.com"
$smtpServer = "anthesisgroup-com.mail.protection.outlook.com"
$mailFrom = "scriptrobot@sustain.co.uk"
$mailTo = "kevin.maitland@anthesisgroup.com"
#convertTo-localisedSecureString ""
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\KimbleBot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
$kimbleCreds = Import-Csv "$env:USERPROFILE\Desktop\Kimble.txt"
$standardKimbleHeaders = get-kimbleHeaders -clientId $kimbleCreds.clientId -clientSecret $kimbleCreds.clientSecret -username $kimbleCreds.username -password $kimbleCreds.password -securityToken $kimbleCreds.securityToken -connectToLiveContext $true -verboseLogging $true
$standardKimbleQueryUri = get-kimbleQueryUri

Connect-PnPOnline -Url $($spoWebUrl+$spoSitePath) -Credentials $adminCreds

#region Kimble Sync
#Get the most recent LastModifiedDate value in the SharePoint List, so we only retrieve the relevant Kimble Objects (this is *not* the LastItemModifiedDate/Modified/SMLastModifiedDate property on the spoListItemObject - it is the Kimble "LastModifiedDate" timestamp that has been synchronised across as a separate property, and this removes any ambiguity caused by the time taken for the sync process to run)
Get-PnPListItem -List $spoListName -Query "<View><Query> <OrderBy> <FieldRef Name='LastModifiedDate' Ascending='False' /> </OrderBy> </Query> </View>" -PageSize 10 -ErrorAction SilentlyContinue | % {if($dummyArray){rv dummyArray};[array]$dummyArray += $_;break} #Get the list item with the most recent LastModifedDate (from Kimble)
$cutoffDate = Get-Date $dummyArray[0].FieldValues.LastModifiedDate -Format s
#$cutoffDate = (Get-Date $pnpKimbleObjectList.LastItemModifiedDate -Format s).AddMinutes(-5) #Bodged as it takes some time for the sync process to run #This caused ambiguities and inefficiencies
$soqlQuery = $soqlQuery.Replace('**PLACEHOLDER**',$cutoffDate)

#Get the Kimble Objects from SalesForce
try{
    log-action -myMessage "Getting [$spoListName] data" -logFile $fullLogPathAndName
    $kimbleModifiedObjects = Get-KimbleSoqlDataset -queryUri $standardKimbleQueryUri -soqlQuery $soqlQuery -restHeaders $standardKimbleHeaders
    if($kimbleModifiedObjects){log-result -myMessage "SUCCESS: $($kimbleModifiedObjects.Count) records retrieved!" -logFile $fullLogPathAndName}
    elseif($kimbleModifiedObjects -eq $null){log-result -myMessage "SUCCESS: Connected, but no records to update." -logFile $fullLogPathAndName}
    else{log-result -myMessage "FAILED: Unable to retrieve data!" -logFile $fullLogPathAndName}
    }
catch{log-error -myError $_ -myFriendlyMessage "Error retrieving [$spoListName] data" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}

#Sort the Kimble Objects for processign based on their relative timestamps
$kimbleChangedObjects = $kimbleModifiedObjects | ?{$_.LastModifiedDate -ge $cutoffDate -and $_.CreatedDate -lt $cutoffDate}
$kimbleNewObjects = $kimbleModifiedObjects | ?{$_.CreatedDate -ge $cutoffDate}

#Process the "Changed" ones first, as we might need to send some for recreation
foreach($kimbleChangedObject in $kimbleChangedObjects){
    if(!$pnpKimbleObjectList){$pnpKimbleObjectList = get-pnpKimbleList -spoListName $spoListName -fullLogPathAndName $fullLogPathAndName} #Now we need it, get the pnpList for CRUD operations (if we don't already have it)
    log-action -myMessage "CHANGED [$spoListName]:`t[$($kimbleChangedObject.Name)] needs updating!" -logFile $fullLogPathAndName
    try{
        $updatedPnpKimbleListItem = update-spoKimbleObjectListItem -kimbleObject $kimbleChangedObject -pnpKimbleObjectList $pnpKimbleObjectList -fullLogPathAndName $fullLogPathAndName -verboseLogging $verboseLogging
        }
    catch{
        log-error -myError $_ -myFriendlyMessage "Error updating [$spoListName] [$($kimbleChangedObject.Name)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
        }
    if($updatedPnpKimbleListItem){log-result -myMessage "SUCCESS: Looks like that worked!" -logFile $fullLogPathAndName}
    else{
        log-result -myMessage "FAILED: Looks like [$spoListName] [$($kimbleChangedObject.Name)] didn't update correctly - will send it for re-creation" -logFile $fullLogPathAndName
        $kimbleNewObjects += $kimbleChangedObject
        }
    }

#Add the new Accounts
foreach ($kimbleNewObject in $kimbleNewObjects){
    if(!$pnpKimbleObjectList){$pnpKimbleObjectList = get-pnpKimbleList -spoListName $spoListName -fullLogPathAndName $fullLogPathAndName} #Now we need it, get the pnpList for CRUD operations (if we don't already have it)
    log-action -myMessage "NEW [$spoListName]:`t[$($kimbleNewObject.Name)] needs creating!" -logFile $fullLogPathAndName
    try{
        $newAccount = new-spoKimbleObjectListItem -kimbleObject $kimbleNewObject -pnpKimbleObjectList $pnpKimbleObjectList -fullLogPathAndName $fullLogPathAndName -verboseLogging $verboseLogging
        }
    catch{
        log-error -myError $_ -myFriendlyMessage "Error creating [$spoListName] [$($kimbleNewObject.Name)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
        }
    if($newAccount){log-result -myMessage "SUCCESS: Looks like that worked!" -logFile $fullLogPathAndName}
    else{
        log-result -myMessage "FAILED: Looks like [$spoListName] [$($kimbleNewObject.Name)] didn't create correctly :(  - that's a bit of a problem!" -logFile $fullLogPathAndName
        }
    }

#endregion



Stop-Transcript
