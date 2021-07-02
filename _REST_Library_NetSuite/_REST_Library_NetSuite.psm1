add-type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint, X509Certificate certificate,
        WebRequest request, int certificateProblem) {
        return true;
    }
}
"@
$AllProtocols = [System.Net.SecurityProtocolType]'Ssl3,Tls,Tls11,Tls12'
[System.Net.ServicePointManager]::SecurityProtocol = $AllProtocols
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

function add-netSuiteAccountToSharePoint{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [PSCustomObject]$sqlNetsuiteAccount 
        ,[parameter(Mandatory = $true)]
        [System.Data.Common.DbConnection]$dbConnection
        ,[parameter(Mandatory = $true)]
        [psobject]$tokenResponse
        )
    Write-Verbose "add-netSuiteAccountToSharePoint [$($sqlNetsuiteAccount.AccountName)]"

    $clientSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,faed84bc-70be-4e35-bfbf-cdab31aeeb99"
    $supplierSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,9fb8ecd6-c87d-485d-a488-26fd18c62303"
    $devSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,8ba7475f-dad0-4d16-bdf5-4f8787838809"

    #Switch to set correct SiteId
    switch($sqlNetsuiteAccount.RecordType){
        "Client"   {$correctSiteId = $clientSiteId}
        "Supplier" {$correctSiteId = $supplierSiteId}
        default    {Write-Error "SqlNetSuiteAccount [$($sqlNetsuiteAccount.AccountName)][$($sqlNetsuiteAccount.NsInternalId)] is neither flagged as a 'Client' nor a 'Supplier' [$($sqlNetsuiteAccount.RecordType)]";break}
        }
    $correctSiteId = $devSiteId

    if(![string]::IsNullOrWhiteSpace($sqlNetsuiteAccount.SharePointDocLibGraphDriveId)){    #Check whether the DocLib Exists already
        Write-Verbose "Looking for /drive by Id [$($sqlNetsuiteAccount.SharePointDocLibGraphDriveId)]"
        $graphDrive = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/drives/$($sqlNetsuiteAccount.SharePointDocLibGraphDriveId)"
        }
    if(!$graphDrive){ #If we don't have a Graph DriveId, or the one we do have doesn't work (e.g. it's been deleted and manually re-created), have a rummage and try to find it by DisplayName(name)
        Write-Verbose "Couldn't find the /drive by Id, looking for Name in case it's been deleted and manually recreated. "
        $allDrivesInSite = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/sites/$correctSiteId/drives" -Verbose  #/drives does not support $filter (as of 2020-01-21)
        $graphDrive = $allDrivesInSite.value | ? {(sanitise-forSharePointGroupName (remove-diacritics $_.name)) -eq (sanitise-forSharePointGroupName (remove-diacritics $sqlNetsuiteAccount.AccountName))}
        if($graphDrive.Count -gt 1){
            Write-Error "Multiple potential Graph /drive matches found with displayName [$($sqlNetsuiteAccount.AccountName)]:`r`n`t$($graphDrive.webUrl -join '`r`n`t')`r`nCannot continue"
            break
            }
        if($graphDrive -eq $null){Write-Verbose "Couldn't find the /drive by Name either. Will try to create a new one. "}
        }

    if($graphDrive){ #Check Name -eq AccountName
        if((sanitise-forSharePointGroupName (remove-diacritics $graphDrive.name)) -ne (sanitise-forSharePointGroupName (remove-diacritics $sqlNetsuiteAccount.AccountName))){ #Update if different
            $docLibNameUpdateHash = @{"displayName"="$(sanitise-forSharePointGroupName $sqlNetsuiteAccount.AccountName)"}
            #Get List Ids from /drive object
            $graphList = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/drives/$($graphDrive.id)/list"
            $updatedGraphList = invoke-graphPatch -tokenResponse $tokenResponse -graphQuery "/sites/$($graphList.parentReference.siteId)/lists/$($graphList.id)" -graphBodyHashtable $docLibNameUpdateHash
            Write-Verbose "$($updatedGraphList.list.template) [$($updatedGraphList.webUrl)][$($updatedGraphList.parentReference) | $($updatedGraphList.id)] changed displayName from [$($graphList.displayName)] to [$($updatedGraphList.displayName)]"
            $updateSqlRecord = $true
            }
        else{$updateSqlRecord = $true} #We don't need to process anything in SharePoint if the Displayname hasn't changed, just prevent this record from re-processings on the next cycle
        }
    else{ #If we can't find a /drives object, create a new one
        $docLibInnerHash = @{"template"="documentLibrary"}
        $docLibOuterHash = @{"displayName"="$(sanitise-forSharePointGroupName $sqlNetsuiteAccount.AccountName)";"list"=$docLibInnerHash}
        $newGraphList = invoke-graphPost -tokenResponse $tokenResponse -graphQuery "/sites/$correctSiteId/lists" -graphBodyHashtable $docLibOuterHash
        $graphDrive = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/sites/$($newGraphList.parentReference.siteId)/lists/$($newGraphList.id)/drive"
        Write-Verbose "$($newGraphList.list.template) [$($newGraphList.webUrl)][$($newGraphList.parentReference) | $($newGraphList.id)] created with displayName [$($newGraphList.displayName)]"
        $updateSqlRecord = $true
        }
    
    if($graphDrive){ #If we've got a /drive object now, try creating the standard folders
        $standardClientFolders = @(
            "_These Client Document Libraries are created automatically by NetSuite"
            ,"_These Client Document Libraries are created automatically by NetSuite\_That's clever!"
            ,"_These Client Document Libraries are created automatically by NetSuite\Create a Client in NetSuite and see"
            )
        add-graphArrayOfFoldersToDrive -graphDriveId $graphDrive.id -foldersAndSubfoldersArray $standardClientFolders -tokenResponse $tokenResponse -conflictResolution Fail
        }

    if($updateSqlRecord){#If we think we should update this record to prevent re-processing on the next cycle
        Write-Verbose "Updating SQL record after successful proccesing"
        $sqlNetsuiteAccount.SharePointDocLibGraphDriveId = $graphDrive.id
        $sqlNetsuiteAccount.DateModifiedInSql = Get-Date
        $updateResult = update-netSuiteAccountInSqlCache -sqlNetsuiteAccount $sqlNetsuiteAccount -dbConnection $dbConnection -isNotDirty
        Write-Verbose "Update Result: [$($updateResult)]"
        #One day, we'll write something to write the URL of the DocLib back to NetSuite...
        }

    $graphDrive #Return the /drives object (if found)
    }
function add-netsuiteAccountToSqlCache{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [PSCustomObject]$nsNetsuiteAccount 
        ,[parameter(Mandatory = $true)]
        [ValidateSet("Client","Supplier")]
        [string]$accountType
        ,[parameter(Mandatory = $true)]
        [System.Data.Common.DbConnection]$dbConnection
        )
    Write-Verbose "add-netsuiteAccountToSqlCache [$($nsNetsuiteAccount.companyName)]"
    $sql = "SELECT TOP 1 AccountName, NsInternalId, LastModified FROM t_ACCOUNTS WHERE NsInternalId = '$($nsNetsuiteAccount.Id)' ORDER BY LastModified Desc"
    Write-Verbose "`t$sql"
    $alreadyPresent = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $dbConnection
    if((remove-diacritics $nsNetsuiteAccount.companyName) -eq (remove-diacritics $alreadyPresent.AccountName)){
        if($(Get-Date $nsNetsuiteAccount.lastModifiedDate) -ne $(Get-Date $alreadyPresent.LastModified)){
            Write-Verbose "`tNsInternalId [$($nsNetsuiteAccount.Id)] has been updated, but the name has not changed. Updating LastModified for existing record adn flagging as IsDirty = $false."
            update-netSuiteAccountInSqlCache -nsNetsuiteAccount $nsNetsuiteAccount -dbConnection $dbConnection -isNotDirty
            }
        else{
            Write-Verbose "`tNsInternalId [$($nsNetsuiteAccount.Id)] doesn't seem to have changed (probably caused by a lack of granularity in NetSuite's REST WHERE clauses). Not updating anything."
            }
        }
    else{
        if(!$alreadyPresent){Write-Verbose "`tNsInternalId [$($nsNetsuiteAccount.Id)] not present in SQL, adding to [ACCOUNTS]"}
        else{Write-Verbose "`tNsInternalId [$($nsNetsuiteAccount.Id)] CompanyName has changed from [$($alreadyPresent.AccountName)] to [$($nsNetsuiteAccount.companyName)], adding new record to [ACCOUNTS]"}
        $now = $(Get-Date)
        $sql = "INSERT INTO t_ACCOUNTS (NsInternalId,NsExternalId,RecordType,AccountName,CustomerNumber,entityId,entityStatus,DateCreated,LastModified,IsDirty,DateCreatedInSql,DateModifiedInSql) VALUES ("
        $sql += $(sanitise-forSqlValue -value $nsNetsuiteAccount.id -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteAccount.accountNumber -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $accountType -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteAccount.companyName -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $($nsNetsuiteAccount.entityId.Split(" ")[0]) -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteAccount.entityId -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteAccount.entityStatus.refName -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteAccount.dateCreated -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteAccount.lastModifiedDate -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $true -dataType Boolean)
        $sql += ","+$(sanitise-forSqlValue -value $now -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $now -dataType Date)
        $sql += ")"
        Write-Verbose "`t$sql"
        $result = Execute-SQLQueryOnSQLDB -query $sql -queryType nonquery -sqlServerConnection $dbConnection
        if($result -eq 1){Write-Verbose "`t`tSUCCESS!"}
        else{Write-Verbose "`t`tFAILURE :( - Code: $result"}
        $result
        }
    }
function add-netSuiteClientToNetSuite{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
            [ValidateSet("HubSpot")]
            [string]$newDataOriginatedFrom
        ,[parameter(Mandatory = $true)]
            [string]$companyName 
        ,[parameter(Mandatory = $true)]
            [string]$externalId
        ,[parameter(Mandatory = $true)]
            [ValidatePattern(".[@].")]
            [string]$genericEmail
        ,[parameter(Mandatory = $true)]
            #[ValidateSet("Anthesis (UK) Ltd","Anthesis Canada Inc.","Anthesis Consulting (USA) Inc.","Anthesis Consulting Group Limited","Anthesis Consulting UK Ltd","Anthesis Consultoria Ambiental ltda","Anthesis Energy UK Ltd","Anthesis Enveco AB","Anthesis Finland OY","Anthesis GmBh","Anthesis Ireland Ltd","Anthesis LLC","Anthesis Middle East","Anthesis Philippines Inc.","Caleb Management Services Ltd","Lavola 1981 SAU","Lavola Andora SA","Lavola Columbia","The Goodbrand Works Ltd","X-Elimination ACUS","X-Elimination AUK","X-Elimination LSA","X-Elimination PC")]
            [string]$subsidiary 
        ,[parameter(Mandatory = $true)]
            #[ValidateSet("LEAD-Qualified","LEAD-Unqualified","CLIENT-Closed Won","CLIENT-Renewal")]
            [string]$status 
        ,[parameter(Mandatory = $false)]
            [AllowNull()]
            #[ValidateSet("Aerospace & Defense","Agriculture","Apparel","Biotechnology","Business & Trade Organization","Business Services","Chemicals & Raw Materials","Construction & Architecture","Consultancy","Containers & Packaging","Distribution & Logistics","Education & Academia","Energy","Engineering & Engineering Services","FMCG - Non-Food","Financial Services & Insurance","Food & Beverage","Forestry, Timber & Paper","Government & Public Services","Health & Pharmaceutical","Hospitality","Information & Communications Technology","Intercompany","Legal Services","Machinery","Manufacturing","Media, Entertainment & Sport","Metals & Mining","NGO & Not for profit","Oil Gas & Renewables","Property & Facilities Management","Retail","Transport & Automotive","Utilities","Waste Disposal & Recycling")]
            [string]$sector 
        ,[parameter(Mandatory = $false)]
            [AllowNull()]
            #[ValidateSet("Government","NGO","Private Company","Public Company","Public Sector")]
            [string]$clientType #= "Private Company"
        ,[parameter(Mandatory = $false)]
            [AllowNull()]
            #[ValidateSet("A – T3 Key Client","B – High Potential","C – Medium Potential","D – Low Potential")]
            [string]$clientRating = "D – Low Potential"
        ,[parameter(Mandatory=$false)]
            [psobject]$netsuiteParameters
        )
    Write-Verbose "add-netSuiteClientToNetSuite [$($companyName)]"

    try{$subsidiaryId = convert-netSuiteSubsidiaryToId -subsidiary $subsidiary}
    catch{Write-Error $_;return}
    try{$statusId = convert-netSuiteStatusToId -status $status}
    catch{Write-Error $_;return}
    if(@("13","15") -contains $statusId){$customFormId = convert-netSuiteCustomFormToId -formType 'Anthesis | Client Form'}
    else{$customFormId = convert-netSuiteCustomFormToId -formType 'Anthesis | Lead Form'}
    if($sector){ #Not mandatory
        try{$sectorId = convert-netSuiteSectorToId -sector $sector}
        catch{Write-Error $_;return}
        }
    if($clientType){ #Not mandatory
        try{$clientTypeId = convert-netSuiteClientTypeToId -clientType $clientType}
        catch{Write-Error $_;return}
        }
    if($clientRating){ #Not mandatory
        try{$clientRatingId = convert-netSuiteclientRatingToId -clientRating $clientRating}
        catch{Write-Error $_;return}
        }
    if([string]::IsNullOrWhiteSpace($netsuiteParameters)){
        $netsuiteParameters = get-netsuiteParameters -connectTo Sandbox
        Write-Warning "NetSuite environment unspecified - connecting to Sandbox"
        }
    
    $bodyHash = @{
        companyName = $companyName
        custentity_2663_email_address_notif = $genericEmail
        custentity_clientrating = @{id=$clientRatingId}
        email = $genericEmail
        entityStatus = @{id=$statusId}
        externalId = $externalId
        subsidiary = @{id=$subsidiaryId}
        #customForm = @{id=$customFormId}
        }
    switch($newDataOriginatedFrom){
        "HubSpot" {
            $bodyHash.Add("custentitycustentity_hubspotid",$externalId)
            $bodyHash.Add("custentity_marketing_originalsourcesyste","HubSpot") 
            }
        }
    
    invoke-netSuiteRestMethod -requestType POST -url "$($netsuiteParameters.uri)/customer" -netsuiteParameters $netsuiteParameters -requestBodyHashTable $bodyHash

    }
function add-netSuiteClientToNetSuiteFromHubSpotObject{
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$hubSpotCompanyObject
        ,[parameter(Mandatory = $true)]
            [string]$hubSpotApiKey
        ,[parameter(Mandatory=$false)]
            [psobject]$netsuiteParameters
        )

    if([string]::IsNullOrWhiteSpace($hubSpotCompanyObject.properties.generic_email_address__c)){
        #No generic e-mail address set - grabbing a Contact and using theirs
        $contacts = get-hubSpotContactsFromCompanyId -apiKey $hubSpotApiKey -hubspotCompanyId $hubSpotCompanyObject.id
        $mostRecentlyCreatedContact = $contacts | Sort-Object createdAt -Descending | select -First 1
        $genericEmailAddress = $mostRecentlyCreatedContact.properties.email
        }
    else{
        $genericEmailAddress = $hubSpotCompanyObject.properties.generic_email_address__c
        }

    $newNetSuiteCompany = add-netSuiteClientToNetSuite `
        -newDataOriginatedFrom HubSpot `
        -companyName $hubSpotCompanyObject.properties.name `
        -externalId $hubSpotCompanyObject.id `
        -genericEmail $genericEmailAddress `
        -subsidiary $hubSpotCompanyObject.properties.netsuite_subsidiary `
        -status LEAD-Unqualified `
        -sector $hubSpotCompanyObject.properties.netsuite_sector `
        -clientRating 'D – Low Potential' `
        -netsuiteParameters $netSuiteParameters
    if(!$newNetSuiteCompany){$newNetSuiteCompany = get-netSuiteClientsFromNetSuite -query "?q=custentitycustentity_hubspotid IS $($hubSpotCompanyObject.id)" -netsuiteParameters $netSuiteParameters}
    $updatedHubSpotCompany = update-hubSpotObject -apiKey $hubSpotApiKey -objectType companies -objectId $hubSpotCompanyObject.id -fieldHash @{netsuiteid=$newNetSuiteCompany.id}
    $newNetSuiteCompany
    }
function add-netSuiteContactToNetSuite{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
            [ValidateSet("HubSpot")]
            [string]$newDataOriginatedFrom
        ,[parameter(Mandatory = $true)]
            [string]$companyNetSuiteId 
        ,[parameter(Mandatory = $true)]
            [AllowNull()]
            [string]$externalId
        ,[parameter(Mandatory = $true)]
            [ValidatePattern(".[@].")]
            [string]$email
        ,[parameter(Mandatory = $true)]
            [string]$contactFirstName 
        ,[parameter(Mandatory = $true)]
            [string]$contactLastName 
        ,[parameter(Mandatory = $false)]
            [string]$mainPhone
        ,[parameter(Mandatory = $false)]
            [string]$mobilePhone
        ,[parameter(Mandatory = $false)]
            [string]$officePhone
        ,[parameter(Mandatory = $false)]
            [string]$jobTitle
        ,[parameter(Mandatory = $true)]
            #[ValidateSet("Anthesis (UK) Ltd","Anthesis Canada Inc.","Anthesis Consulting (USA) Inc.","Anthesis Consulting Group Limited","Anthesis Consulting UK Ltd","Anthesis Consultoria Ambiental ltda","Anthesis Energy UK Ltd","Anthesis Enveco AB","Anthesis Finland OY","Anthesis GmBh","Anthesis Ireland Ltd","Anthesis LLC","Anthesis Middle East","Anthesis Philippines Inc.","Caleb Management Services Ltd","Lavola 1981 SAU","Lavola Andora SA","Lavola Columbia","The Goodbrand Works Ltd","X-Elimination ACUS","X-Elimination AUK","X-Elimination LSA","X-Elimination PC")]
            [string]$subsidiary 
        ,[parameter(Mandatory=$false)]
            [string]$originalSourceOfContact
        ,[parameter(Mandatory=$false)]
            [psobject]$netsuiteParameters
        )
    $fullContactName = "$($hubSpotContactObject.properties.firstname) $($hubSpotContactObject.properties.lastname)".Trim(" ")
    Write-Verbose "add-netSuiteContactToNetSuite [$($contactFullName)]"

    try{$subsidiaryId = convert-netSuiteSubsidiaryToId -subsidiary $subsidiary}
    catch{Write-Error $_;return}

    if([string]::IsNullOrWhiteSpace($netsuiteParameters)){
        $netsuiteParameters = get-netsuiteParameters -connectTo Sandbox
        Write-Warning "NetSuite environment unspecified - connecting to Sandbox"
        }

    
    $bodyHash = [ordered]@{
        entityId = $fullContactName
        externalId = $externalId
        email = $email
        phone = $mainPhone
        mobilePhone = $mobilePhone
        officePhone = $officePhone
        title = $jobTitle
        subsidiary = @{id=$subsidiaryId}
        company = @{id=$companyNetSuiteId}
        }
    switch($newDataOriginatedFrom){
        "HubSpot" {
            $bodyHash.Add("custentitycustentity_hubspotid",$externalId)
            $bodyHash.Add("custentity_marketing_originalsourcesyste","HubSpot") 
            }
        }
    
    $result = invoke-netSuiteRestMethod -requestType POST -url "$($netsuiteParameters.uri)/contact" -netsuiteParameters $netsuiteParameters -requestBodyHashTable $bodyHash

    $newNetSuiteContact = get-netSuiteContactFromNetSuite -query "?q=externalId IS $($externalId)" -netsuiteParameters $netSuiteParameters #invoke above doesn't return the new object :(
    $newNetSuiteContact
    }
function add-netSuiteContactToNetSuiteFromHubSpotObject{
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$hubSpotContactObject
        ,[parameter(Mandatory = $true)]
            [string]$hubSpotApiKey
        ,[parameter(Mandatory = $true)]
            [string]$companyNetSuiteId 
        ,[parameter(Mandatory = $true)]
            #[ValidateSet("Anthesis (UK) Ltd","Anthesis Canada Inc.","Anthesis Consulting (USA) Inc.","Anthesis Consulting Group Limited","Anthesis Consulting UK Ltd","Anthesis Consultoria Ambiental ltda","Anthesis Energy UK Ltd","Anthesis Enveco AB","Anthesis Finland OY","Anthesis GmBh","Anthesis Ireland Ltd","Anthesis LLC","Anthesis Middle East","Anthesis Philippines Inc.","Caleb Management Services Ltd","Lavola 1981 SAU","Lavola Andora SA","Lavola Columbia","The Goodbrand Works Ltd","X-Elimination ACUS","X-Elimination AUK","X-Elimination LSA","X-Elimination PC")]
            [string]$subsidiary 
        ,[parameter(Mandatory=$false)]
            [psobject]$netsuiteParameters
        )
    Write-Verbose "add-netSuiteContactToNetSuiteFromHubSpotObject [$($hubSpotContactObject.properties.firstname)][$($hubSpotContactObject.properties.lastname)][$($hubSpotContactObject.properties.id)]"

    try{$subsidiaryId = convert-netSuiteSubsidiaryToId -subsidiary $subsidiary}
    catch{Write-Error $_;return}

    if([string]::IsNullOrWhiteSpace($netsuiteParameters)){
        $netsuiteParameters = get-netsuiteParameters -connectTo Sandbox
        Write-Warning "NetSuite environment unspecified - connecting to Sandbox"
        }

    $fullContactName = "$($hubSpotContactObject.properties.firstname) $($hubSpotContactObject.properties.lastname)".Trim(" ")
    if(![string]::IsNullOrWhiteSpace($hubSpotContactObject.properties.phone)){$mainPhone = $hubSpotContactObject.properties.phone}
    if(![string]::IsNullOrWhiteSpace($hubSpotContactObject.properties.mobilephone)){$mainPhone = $hubSpotContactObject.properties.mobilephone} #Prefer mobiles over landlines as primary phone
    
    #Create a Stub record in NetSuite
    $newNetSuiteContact = add-netSuiteContactToNetSuite `
        -newDataOriginatedFrom HubSpot `
        -contactFirstName $($hubSpotContactObject.properties.firstname).Trim() `
        -contactLastName $($hubSpotContactObject.properties.lastname).Trim() `
        -companyNetSuiteId $companyNetSuiteId `
        -externalId $hubSpotContactObject.id `
        -email $hubSpotContactObject.properties.email `
        -mainPhone $mainPhone `
        -mobilePhone $hubSpotContactObject.properties.mobilephone `
        -officePhone $hubSpotContactObject.properties.phone `
        -jobTitle $hubSpotContactObject.properties.jobtitle `
        -subsidiary $subsidiary `
        -netsuiteParameters $netsuiteParameters

    #Then update it with all the Marketing Bells & Whistles
    $hubSpotContactObject.properties.netsuiteid = $newNetSuiteContact.id
    $updatedNetSuiteContact = update-netSuiteContactInNetSuiteFromHubSpotObject -hubSpotContactObject $hubSpotContactObject -hubSpotApiKey $hubSpotApiKey -companyNetSuiteId $companyNetSuiteId -netsuiteParameters $netsuiteParameters 
    $updatedHubSpotContact = update-hubSpotObject -apiKey $hubSpotApiKey -objectType contacts -objectId $hubSpotContactObject.id -fieldHash @{netsuiteid=$newNetSuiteContact.id; lastmodifiedinnetsuite=$newNetSuiteContact.lastModifiedDate; lastmodifiedinhubspot=$(get-dateInIsoFormat -dateTime $(Get-Date) -precision Ticks)}
    $newNetSuiteContact
    #$updatedNetSuiteContact 
    }
function add-netSuiteProjectToSharePoint{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [PSCustomObject]$sqlNetsuiteProject 
        ,[parameter(Mandatory = $true)]
        [System.Data.Common.DbConnection]$dbConnection
        ,[parameter(Mandatory = $true)]
        [psobject]$tokenResponse
        )
    Write-Verbose "add-netSuiteProjectToSharePoint [$($sqlNetsuiteProject.entityId)]"

    #Check if the folder already exists
    #Get Drive Id
    $sqlNetSuiteClient = get-netSuiteClientFromSqlCache -dbConnection $dbConnection -sqlWhereClause "WHERE NsInternalId = '$($sqlNetsuiteProject.AccountNsInternalId)'"
    if(![string]::IsNullOrWhiteSpace($sqlNetSuiteClient.SharePointDocLibGraphDriveId)){
        Write-Verbose "Looking for /drive/{drive-id}/items/{item-id} by Id [$($sqlNetsuiteAccount.SharePointDocLibGraphDriveId)]"
        $graphDriveItem = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/drives/$($sqlNetSuiteClient.SharePointDocLibGraphDriveId)/items/$($sqlNetsuiteProject.SharePointDriveItemId)" -ErrorAction SilentlyContinue
        }
    else{Write-Error "Unable to find SharePointDocLibGraphDriveId for Client [$($sqlNetsuiteProject.AccountNsInternalId)][$($sqlNetSuiteClient.AccountName)]. Cannot attempt to create Project Folders.";break}

    if(!$graphDriveItem){#If we don't have a Graph DriveItemId, or the one we do have doesn't work (e.g. it's been deleted and manually re-created), have a rummage and try to find it by DisplayName(name)
        Write-Verbose "Couldn't find the /drive/{drive-id}/items/{item-id} by Id, looking for Name in case it's been deleted and manually recreated. "
        $allGraphDriveRootItems = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/drives/$($sqlNetSuiteClient.SharePointDocLibGraphDriveId)/root/children" #/drives does not support $filter (as of 2020-01-21)
        $graphDriveItem = $allGraphDriveRootItems.value | ? {$(sanitise-forSharePointGroupName (remove-diacritics $_.name)) -eq $(sanitise-forSharePointGroupName (remove-diacritics $sqlNetsuiteProject.entityId))}
        }

    if($graphDriveItem){#If we've found an existing item, check whether it needs updating
        if($(sanitise-forSharePointGroupName (remove-diacritics $graphDriveItem.name)) -ne $(sanitise-forSharePointGroupName (remove-diacritics $sqlNetsuiteProject.entityId))){ #Update the Project/folder name if it's changed
            $folderUpdateHash = @{"name"="$(sanitise-forSharePointGroupName $sqlNetsuiteProject.entityId)"}
            Write-Verbose "Updating name from [$($graphDriveItem.name)] to [$(sanitise-forSharePointGroupName $sqlNetsuiteProject.entityId)] for Project [$($graphDriveItem.webUrl)] | $($graphDriveItem.id)]"
            $updatedGraphDriveItem = invoke-graphPatch -tokenResponse $tokenResponse -graphQuery "/drives/$($sqlNetSuiteClient.SharePointDocLibGraphDriveId)/items/$($sqlNetsuiteProject.SharePointDriveItemId)" -graphBodyHashtable $folderUpdateHash
            Write-Verbose "[$($updatedGraphDriveItem.webUrl)] | $($updatedGraphDriveItem.id)] changed displayName from [$($graphDriveItem.name)] to [$($updatedGraphDriveItem.name)]"
            $updateSqlRecord = $true
            }
        else{$updateSqlRecord = $true} #We don't need to process anything in SharePoint if the Displayname hasn't changed, just prevent this record from re-processings on the next cycle
        }
    else{#If we still can't find an existing item, create a new one
        $arrayOfProjectFolders = @(
            "$($sqlNetsuiteProject.entityId)"
            ,"$($sqlNetsuiteProject.entityId)\Admin & contracts"
            ,"$($sqlNetsuiteProject.entityId)\Analysis"
            ,"$($sqlNetsuiteProject.entityId)\Data & refs"
            ,"$($sqlNetsuiteProject.entityId)\Meetings"
            ,"$($sqlNetsuiteProject.entityId)\Proposal"
            ,"$($sqlNetsuiteProject.entityId)\Reports"
            ,"$($sqlNetsuiteProject.entityId)\Summary (marketing) - end of project"
            )
        $newProjectFolders = add-graphArrayOfFoldersToDrive -graphDriveId $sqlNetSuiteClient.SharePointDocLibGraphDriveId -foldersAndSubfoldersArray $arrayOfProjectFolders -tokenResponse $tokenResponse -conflictResolution Fail
        $graphDriveItem = $newProjectFolders | ? {$_.name -eq (sanitise-forSharePointGroupName $sqlNetsuiteProject.entityId)}
        Write-Verbose "[$($graphDriveItem.webUrl)] | $($graphDriveItem.id)]  created with displayName [$($graphDriveItem.name)]"
        $updateSqlRecord = $true
        }

    if($updateSqlRecord){#If we think we should update this record to prevent re-processing on the next cycle
        Write-Verbose "Updating SQL Project record after successful proccesing"
        $sqlNetsuiteProject.SharePointDriveItemId = $graphDriveItem.id
        $sqlNetsuiteProject.DateModifiedInSql = Get-Date
        $updateResult = update-netSuiteProjectInSqlCache -sqlNetSuiteProject $sqlNetSuiteProject -dbConnection $dbConnection -isNotDirty
        Write-Verbose "Update Result: [$($updateResult)]"
        #One day, we'll write something to write the URL of the Folder back to NetSuite...
        }

    $graphDriveItem #Return the /drives object (if found)
    }
function add-netsuiteProjectToSqlCache{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [PSCustomObject]$nsNetsuiteProject 
        ,[parameter(Mandatory = $true)]
        [System.Data.Common.DbConnection]$dbConnection
        )
    Write-Verbose "add-netsuiteProjectToSqlCache [$($nsNetsuiteProject.ProjectName)]"
    $sql = "SELECT TOP 1 ProjectName, NsInternalId, LastModified FROM t_PROJECTS WHERE NsInternalId = '$($nsNetsuiteProject.Id)' ORDER BY LastModified Desc"
    Write-Verbose "`t$sql"
    $alreadyPresent = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $dbConnection
    if((remove-diacritics $nsNetsuiteProject.companyName) -eq (remove-diacritics $alreadyPresent.ProjectName)){
        if($(Get-Date $nsNetsuiteProject.lastModifiedDate) -ne $(Get-Date $alreadyPresent.LastModified)){
            Write-Verbose "`tNsInternalId [$($nsNetsuiteProject.Id)] has been updated, but the name has not changed. Updating LastModified for existing record and flagging as IsDirty = $false."
            update-netSuiteProjectInSqlCache -nsNetsuiteProject $nsNetsuiteProject -dbConnection $dbConnection -isNotDirty
            }
        else{
            Write-Verbose "`tNsInternalId [$($nsNetsuiteProject.Id)] doesn't seem to have changed (probably caused by a lack of granularity in NetSuite's REST WHERE clauses). Not updating anything."
            }
        }
    else{
        if(!$alreadyPresent){Write-Verbose "`tNsInternalId [$($nsNetsuiteProject.Id)] not present in SQL, adding to [t_PROJECTS]"}
        else{Write-Verbose "`tNsInternalId [$($nsNetsuiteProject.Id)] ProjectName has changed from [$($alreadyPresent.ProjectName)] to [$($nsNetsuiteProject.companyName)], adding new record to [t_PROJECTS]"}
        $now = $(Get-Date)
        $sql = "INSERT INTO t_PROJECTS (NsInternalId, NsExternalId, AccountNsInternalId, ProjectName, ProjectNumber, entityId, entityStatus, custentity_atlas_svcs_mm_department, custentity_ant_projectsector, custentity_ant_projectsource, custentity_atlas_svcs_mm_location, custentity_atlas_svcs_mm_projectmngr, jobType, subsidiary, DateCreated, LastModified, IsDirty, DateCreatedInSql, DateModifiedInSql) VALUES ("
        $sql += $(sanitise-forSqlValue -value $nsNetsuiteProject.id -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteProject.entityId -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteProject.parent.id -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteProject.companyName -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $($nsNetsuiteProject.entityId.Split(" ")[0]) -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteProject.entityId -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteProject.entityStatus.refName -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteProject.custentity_atlas_svcs_mm_department.refName -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteProject.custentity_ant_projectsector.refName -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteProject.custentity_ant_projectsource.refName -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteProject.custentity_atlas_svcs_mm_location.refName -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteProject.custentity_atlas_svcs_mm_projectmngr.refName -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteProject.jobType.refName -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteProject.subsidiary.refName -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteProject.dateCreated -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteProject.lastModifiedDate -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $true -dataType Boolean)
        $sql += ","+$(sanitise-forSqlValue -value $now -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $now -dataType Date)
        $sql += ")"
        Write-Verbose "`t$sql"
        $result = Execute-SQLQueryOnSQLDB -query $sql -queryType nonquery -sqlServerConnection $dbConnection
        if($result -eq 1){Write-Verbose "`t`tSUCCESS!"}
        else{Write-Verbose "`t`tFAILURE :( - Code: $result"}
        $result
        }
    }
function add-netsuiteOpportunityToSqlCache{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [PSCustomObject]$nsNetsuiteOpportunity 
        ,[parameter(Mandatory = $true)]
        [System.Data.Common.DbConnection]$dbConnection
        )
    Write-Verbose "add-netsuiteOpportunityToSqlCache [$($nsNetsuiteOpportunity.title)]"
    $sql = "SELECT TOP 1 OpportunityName, NsInternalId, LastModified FROM t_OPPORTUNITIES WHERE NsInternalId = '$($nsNetsuiteOpportunity.Id)' ORDER BY LastModified Desc"
    Write-Verbose "`t$sql"
    $alreadyPresent = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $dbConnection
    if((remove-diacritics $nsNetsuiteOpportunity.title) -eq (remove-diacritics $alreadyPresent.OpportunityName) -and $alreadyPresent){
        if($(Get-Date $nsNetsuiteOpportunity.lastModifiedDate) -ne $(Get-Date $alreadyPresent.LastModified)){
            Write-Verbose "`tNsInternalId [$($nsNetsuiteOpportunity.Id)] has been updated, but the name has not changed. Updating LastModified for existing record and flagging as IsDirty = $false."
            update-netSuiteOpportunityInSqlCache -nsNetSuiteOpportunity $nsNetsuiteOpportunity -dbConnection $dbConnection -isNotDirty
            }
        else{
            Write-Verbose "`tNsInternalId [$($nsNetsuiteOpportunity.Id)] doesn't seem to have changed (probably caused by a lack of granularity in NetSuite's REST WHERE clauses). Not updating anything."
            }
        }
    else{
        if(!$alreadyPresent){Write-Verbose "`tNsInternalId [$($nsNetsuiteOpportunity.Id)] not present in SQL, adding to [t_OPPORTUNITIES]"}
        else{Write-Verbose "`tNsInternalId [$($nsNetsuiteOpportunity.Id)] Title has changed from [$($alreadyPresent.ProjectName)] to [$($nsNetsuiteOpportunity.companyName)], adding new record to [t_PROJECTS]"}
        $now = $(Get-Date)
        $sql = "INSERT INTO t_OPPORTUNITIES (NsInternalId, NsExternalId, AccountNsInternalId, ProjectNsInternalId, OpportunityName, OpportunityNumber, entityId, entityStatus, entityNexus, custbody_project_template, tranId, status, probability, custbody_industry, subsidiary, DateCreated, LastModified, DateCreatedInSql, DateModifiedInSql, IsDirty) VALUES ("
                                             
        $sql += $(sanitise-forSqlValue -value $nsNetsuiteOpportunity.id -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteOpportunity.id -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteOpportunity.entity.id -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteOpportunity.custbody_project_created.id -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteOpportunity.title -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteOpportunity.tranId -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value "$($nsNetsuiteOpportunity.tranId) $($nsNetsuiteOpportunity.title)" -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteOpportunity.entityStatus.refName -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteOpportunity.entityNexus.refName -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteOpportunity.custbody_project_template.refName -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteOpportunity.tranId -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteOpportunity.status -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteOpportunity.probability -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteOpportunity.custbody_industry -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteOpportunity.subsidiary.refName -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteOpportunity.createdDate -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteOpportunity.lastModifiedDate -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $now -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $now -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $true -dataType Boolean)
        $sql += ")"
        Write-Verbose "`t$sql"
        $result = Execute-SQLQueryOnSQLDB -query $sql -queryType nonquery -sqlServerConnection $dbConnection
        if($result -eq 1){Write-Verbose "`t`tSUCCESS!"}
        else{Write-Verbose "`t`tFAILURE :( - Code: $result"}
        $result
        }
    }
function convert-netSuiteClientRatingToId(){
    [cmdletbinding()]
    Param (    [parameter(Mandatory = $true)]
        [ValidateSet("A – T3 Key Client","B – High Potential","C – Medium Potential","D – Low Potential")]
        [string]$clientRating
        )
    #clientRating Validation:($(get-netSuiteCustomListValues -objectType clientRating -netsuiteParameters $(get-netSuiteParameters -connectTo Production)) | ? {$_.isInactive -eq $false}).name -join '","'
    #                         $(get-netSuiteCustomListValues -objectType clientRating -netsuiteParameters $(get-netSuiteParameters -connectTo Production)) | ? {$_.isInactive -eq $false} | % {"`"$($_.Name)`"`t{`$clientRatingId = $($_.id)}"}

    switch($clientRating){
        "A – T3 Key Client"	{$clientRatingId = 1}
        "B – High Potential"	{$clientRatingId = 2}
        "C – Medium Potential"	{$clientRatingId = 3}
        "D – Low Potential"	{$clientRatingId = 4}
        }
    $clientRatingId
    }
function convert-netSuiteClientTypeToId(){
    [cmdletbinding()]
    Param (    [parameter(Mandatory = $true)]
        [ValidateSet("Government","NGO","Private Company","Public Company","Public Sector")]
        [string]$clientType
        )
    #ClientType Validation:  ($(get-netSuiteCustomListValues -objectType clientType -netsuiteParameters $(get-netSuiteParameters -connectTo Production)) | ? {$_.isInactive -eq $false}).name -join '","'
    #                         $(get-netSuiteCustomListValues -objectType clientType -netsuiteParameters $(get-netSuiteParameters -connectTo Production)) | ? {$_.isInactive -eq $false} | % {"`"$($_.Name)`"`t{`$clientTypeId = $($_.id)}"}

    switch($clientType){
        "Government"	{$clientTypeId = 3}
        "NGO"	{$clientTypeId = 101}
        "Private Company"	{$clientTypeId = 2}
        "Public Company"	{$clientTypeId = 1}
        "Public Sector"	{$clientTypeId = 102}
        }
    $clientTypeId
    }
function convert-netSuiteCustomFormToId(){
    [cmdletbinding()]
    Param (    [parameter(Mandatory = $true)]
        [ValidateSet("Anthesis | Client Form","Anthesis | Lead Form","Anthesis Contact Form","REST HubSpot Integration| Client Form","REST HubSpot Integration | Contact Form")]
        [string]$formType
        )
    #ClientType Validation:  ($(get-netSuiteCustomListValues -objectType clientType -netsuiteParameters $(get-netSuiteParameters -connectTo Production)) | ? {$_.isInactive -eq $false}).name -join '","'
    #                         $(get-netSuiteCustomListValues -objectType clientType -netsuiteParameters $(get-netSuiteParameters -connectTo Production)) | ? {$_.isInactive -eq $false} | % {"`"$($_.Name)`"`t{`$clientTypeId = $($_.id)}"}

    switch($formType){
        "Anthesis | Client Form"	{$formTypeId = 141}
        "Anthesis Contact Form"	{$formTypeId = 107}
        "Anthesis | Lead Form"	    {$formTypeId = 101}
        "REST HubSpot Integration| Client Form"	{$formTypeId = 171}
        "REST HubSpot Integration | Contact Form"	{$formTypeId = 172}
        }
    $formTypeId
    }
function convert-netSuiteSectorToId(){
    [cmdletbinding()]
    Param (    [parameter(Mandatory = $true)]
        [ValidateSet("Aerospace & Defense","Agriculture","Apparel","Biotechnology","Business & Trade Organization","Business Services","Chemicals & Raw Materials","Construction & Architecture","Consultancy","Containers & Packaging","Distribution & Logistics","Education & Academia","Energy","Engineering & Engineering Services","FMCG - Non-Food","Financial Services & Insurance","Food & Beverage","Forestry, Timber & Paper","Government & Public Services","Health & Pharmaceutical","Hospitality","Information & Communications Technology","Intercompany","Legal Services","Machinery","Manufacturing","Media, Entertainment & Sport","Metals & Mining","NGO & Not for profit","Oil Gas & Renewables","Property & Facilities Management","Retail","Transport & Automotive","Utilities","Waste Disposal & Recycling")]
        [string]$sector
        )
    #Sector Validation:      ($(get-netSuiteCustomListValues -objectType clientSector -netsuiteParameters $(get-netSuiteParameters -connectTo Production)) | ? {$_.isInactive -eq $false}).name -join '","'
    #                         $(get-netSuiteCustomListValues -objectType clientSector -netsuiteParameters $(get-netSuiteParameters -connectTo Production)) | ? {$_.isInactive -eq $false} | % {"`"$($_.Name)`"`t{`$sectorId = $($_.id)}"}

    switch($sector){
        "Aerospace & Defense"	{$sectorId = 1}
        "Agriculture"	{$sectorId = 2}
        "Apparel"	{$sectorId = 3}
        "Biotechnology"	{$sectorId = 33}
        "Business & Trade Organization"	{$sectorId = 4}
        "Business Services"	{$sectorId = 5}
        "Chemicals & Raw Materials"	{$sectorId = 7}
        "Construction & Architecture"	{$sectorId = 8}
        "Consultancy"	{$sectorId = 9}
        "Containers & Packaging"	{$sectorId = 10}
        "Distribution & Logistics"	{$sectorId = 11}
        "Education & Academia"	{$sectorId = 12}
        "Energy"	{$sectorId = 34}
        "Engineering & Engineering Services"	{$sectorId = 13}
        "FMCG - Non-Food"	{$sectorId = 15}
        "Financial Services & Insurance"	{$sectorId = 14}
        "Food & Beverage"	{$sectorId = 16}
        "Forestry, Timber & Paper"	{$sectorId = 17}
        "Government & Public Services"	{$sectorId = 18}
        "Health & Pharmaceutical"	{$sectorId = 19}
        "Hospitality"	{$sectorId = 36}
        "Information & Communications Technology"	{$sectorId = 21}
        "Intercompany"	{$sectorId = 38}
        "Legal Services"	{$sectorId = 22}
        "Machinery"	{$sectorId = 37}
        "Manufacturing"	{$sectorId = 23}
        "Media, Entertainment & Sport"	{$sectorId = 24}
        "Metals & Mining"	{$sectorId = 25}
        "NGO & Not for profit"	{$sectorId = 26}
        "Oil Gas & Renewables"	{$sectorId = 27}
        "Property & Facilities Management"	{$sectorId = 28}
        "Retail"	{$sectorId = 29}
        "Transport & Automotive"	{$sectorId = 30}
        "Utilities"	{$sectorId = 31}
        "Waste Disposal & Recycling"	{$sectorId = 32}
        }
    $sectorId
    }
function convert-netSuiteStatusToId(){
    [cmdletbinding()]
    Param (    [parameter(Mandatory = $true)]
        [ValidateSet("LEAD-Qualified","LEAD-Unqualified","CLIENT-Closed Won","CLIENT-Renewal")]
        [string]$status
        )
    #Status Validation:      ($(get-netSuiteCustomListValues -objectType customerstatus -netsuiteParameters $(get-netSuiteParameters -connectTo Production)) | ? {$_.isInactive -eq $false -and $_.stage -ne "JOB"} | Sort-Object probability,name |  % {"`"$($_.stage)-$($_.name)`""}
    #                         $(get-netSuiteCustomListValues -objectType customerstatus -netsuiteParameters $(get-netSuiteParameters -connectTo Production)) | ? {$_.isInactive -eq $false -and $_.stage -ne "JOB"} | Sort-Object probability,name |  % {"`"$($_.stage)-$($_.name)`"`t{`$statusId = $($_.id)}"}

    switch($status){
        "PROSPECT-Closed Lost"	{$statusId = 14}
        "LEAD-Unqualified"	{$statusId = 6}
        "LEAD-Qualified"	{$statusId = 7}
        "PROSPECT-Identified Opportunity"	{$statusId = 8}
        "PROSPECT-Initial Discussion"	{$statusId = 9}
        "PROSPECT-RFP"	{$statusId = 20}
        "PROSPECT-Proposal Submitted"	{$statusId = 10}
        "PROSPECT-Positive Proposal Response"	{$statusId = 19}
        "PROSPECT-Detailed Opportunity Discussion"	{$statusId = 21}
        "PROSPECT-In Negotiation"	{$statusId = 11}
        "PROSPECT-Verbal Agreement"	{$statusId = 12}
        "CUSTOMER-Closed Won"	{$statusId = 13}
        "CUSTOMER-Renewal"	{$statusId = 15}
        }
    $statusId
    }
function convert-netSuiteSubsidiaryToId(){
    [cmdletbinding()]
    Param (    [parameter(Mandatory = $true)]
        [ValidateSet("Anthesis (UK) Ltd","Anthesis Canada Inc.","Anthesis Consulting (USA) Inc.","Anthesis Consulting Group Limited","Anthesis Consulting UK Ltd","Anthesis Consultoria Ambiental ltda","Anthesis Energy UK Ltd","Anthesis Enveco AB","Anthesis Finland OY","Anthesis GmBh","Anthesis Ireland Ltd","Anthesis LLC","Anthesis Middle East","Anthesis Philippines Inc.","Caleb Management Services Ltd","Lavola 1981 SAU","Lavola Andora SA","Lavola Columbia","The Goodbrand Works Ltd","X-Elimination ACUS","X-Elimination AUK","X-Elimination LSA","X-Elimination PC")]
        [string]$subsidiary
        )
    #Subsidiary Validation:  ($(get-netSuiteCustomListValues -objectType subsidiary -netsuiteParameters $(get-netSuiteParameters -connectTo Production)) | ? {$_.isInactive -eq $false}).name -join '","'
    #                         $(get-netSuiteCustomListValues -objectType subsidiary -netsuiteParameters $(get-netSuiteParameters -connectTo Production)) | ? {$_.isInactive -eq $false} | % {"`"$($_.Name)`"`t{`$subsidiaryId = $($_.id)}"}

    switch($subsidiary){
        "Anthesis (UK) Ltd"	{$subsidiaryId = 6}
        "Anthesis Canada Inc."	{$subsidiaryId = 43}
        "Anthesis Consulting (USA) Inc."	{$subsidiaryId = 41}
        "Anthesis Consulting Group Limited"	{$subsidiaryId = 1}
        "Anthesis Consulting UK Ltd"	{$subsidiaryId = 33}
        "Anthesis Consultoria Ambiental ltda"	{$subsidiaryId = 57}
        "Anthesis Energy UK Ltd"	{$subsidiaryId = 7}
        "Anthesis Enveco AB"	{$subsidiaryId = 52}
        "Anthesis Finland OY"	{$subsidiaryId = 55}
        "Anthesis GmBh"	{$subsidiaryId = 49}
        "Anthesis Ireland Ltd"	{$subsidiaryId = 23}
        "Anthesis LLC"	{$subsidiaryId = 42}
        "Anthesis Middle East"	{$subsidiaryId = 46}
        "Anthesis Philippines Inc."	{$subsidiaryId = 44}
        "Caleb Management Services Ltd"	{$subsidiaryId = 34}
        "Lavola 1981 SAU"	{$subsidiaryId = 4}
        "Lavola Andora SA"	{$subsidiaryId = 40}
        "Lavola Columbia"	{$subsidiaryId = 47}
        "The Goodbrand Works Ltd"	{$subsidiaryId = 3}
        "X-Elimination ACUS"	{$subsidiaryId = 45}
        "X-Elimination AUK"	{$subsidiaryId = 32}
        "X-Elimination LSA"	{$subsidiaryId = 48}
        "X-Elimination PC"	{$subsidiaryId = 12}
        default {Write-Error "Subsidiary [$subsidiary] does not map to a known Subsidiary ID"}
        }
    $subsidiaryId
    }
function convert-nsNetSuiteAccountToSqlNetSuiteAccount(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
            [PSCustomObject]$nsNetsuiteAccount 
        )
    switch($nsNetsuiteAccount.entityStatus.refName){
        {$_ -match "Client"} {$recordType = "Client"}
        {$_ -match "Supplier"} {$recordType = "Supplier"}
        default {$recordType = $null}
        }

    <#$pretendSqlNetSuiteAccount = New-Object psobject -Property @{
        "AccountName"                  = $(sanitise-forSqlValue -value $nsNetsuiteAccount.companyName -dataType String)
        "NsInternalId"                 = $(sanitise-forSqlValue -value $nsNetsuiteAccount.id -dataType String)
        "NsExternalId"                 = $(sanitise-forSqlValue -value $nsNetsuiteAccount.accountNumber -dataType String)
        "RecordType"                   = $recordType
        "entityStatus"                 = $(sanitise-forSqlValue -value $nsNetsuiteAccount.entityStatus.refName -dataType String)
        "DateCreated"                  = $(sanitise-forSqlValue -value $nsNetsuiteAccount.dateCreated -dataType Date)
        "LastModified"                 = $(sanitise-forSqlValue -value $nsNetsuiteAccount.lastModifiedDate -dataType Date)
        "DateCreatedInSql"             = $null
        "DateModifiedInSql"            = $null
        "IsDirty"                      = $null
        "SharePointDocLibGraphListId"  = $null
        "SharePointDocLibGraphDriveId" = $null    
        }#>
    $pretendSqlNetSuiteAccount = New-Object psobject -Property @{
        "AccountName"                  = $nsNetsuiteAccount.companyName
        "NsInternalId"                 = $nsNetsuiteAccount.id
        "NsExternalId"                 = $nsNetsuiteAccount.accountNumber
        "RecordType"                   = $recordType
        "entityStatus"                 = $nsNetsuiteAccount.entityStatus.refName
        "DateCreated"                  = $nsNetsuiteAccount.dateCreated
        "LastModified"                 = $nsNetsuiteAccount.lastModifiedDate
        "DateCreatedInSql"             = $null
        "DateModifiedInSql"            = $null
        "IsDirty"                      = $null
        "SharePointDocLibGraphListId"  = $null
        "SharePointDocLibGraphDriveId" = $null    
        }
    $pretendSqlNetSuiteAccount
    }
function convert-nsNetSuiteOpportunityToSqlNetSuiteOpportunity(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
            [PSCustomObject]$nsNetsuiteOpportunity 
        )

    $pretendSqlNetSuiteOpportunity = New-Object psobject -Property @{
        "NsInternalId"              = $nsNetsuiteOpportunity.id
        "NsExternalId"              = $null
        "AccountNsInternalId"       = $nsNetsuiteOpportunity.entity.id
        "ProjectNsInternalId"       = $nsNetsuiteOpportunity.custbody_project_created.id
        "OpportunityName"           = $nsNetsuiteOpportunity.title
        "OpportunityNumber"         = $nsNetsuiteOpportunity.tranId
        "entityId"                  = "$($nsNetsuiteOpportunity.tranId) $($nsNetsuiteOpportunity.title)"
        "entityStatus"              = $nsNetsuiteOpportunity.entityStatus.refName
        "entityNexus"               = $nsNetsuiteOpportunity.entityNexus.refName
        "custbody_project_template" = $nsNetsuiteOpportunity.custbody_project_template.refName
        "tranId"                    = $nsNetsuiteOpportunity.tranId
        "status"                    = $nsNetsuiteOpportunity.status
        "probability"               = $nsNetsuiteOpportunity.probability
        "custbody_industry"         = $nsNetsuiteOpportunity.custbody_industry
        "subsidiary"                = $nsNetsuiteOpportunity.subsidiary.refName
        "DateCreated"               = $nsNetsuiteOpportunity.createdDate
        "LastModified"              = $nsNetsuiteOpportunity.lastModifiedDate
        "DateCreatedInSql"          = $null
        "DateModifiedInSql"         = $null
        "IsDirty"                   = $null
        "SharePointDriveItemId"     = $null
        }
    $pretendSqlNetSuiteOpportunity
    }
function convert-nsNetSuiteProjectToSqlNetSuiteProject(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
            [PSCustomObject]$nsNetsuiteProject 
        )

    $pretendSqlNetSuiteProject = New-Object psobject -Property @{
        "NsInternalId"                         = $nsNetsuiteProject.id
        "NsExternalId"                         = $null
        "AccountNsInternalId"                  = $null
        "ProjectName"                          = $nsNetsuiteProject.companyName
        "ProjectNumber"                        = $($nsNetsuiteProject.entityId.Split(" ")[0])
        "entityId"                             = $nsNetsuiteProject.entityId
        "custentity_atlas_svcs_mm_department"  = $nsNetsuiteProject.entityStatus.refName
        "entityStatus"                         = $nsNetsuiteProject.entityStatus.refName
        "custentity_ant_projectsector"         = $nsNetsuiteProject.custentity_atlas_svcs_mm_department.refName
        "custentity_ant_projectsource"         = $nsNetsuiteProject.custentity_ant_projectsource.refName
        "custentity_atlas_svcs_mm_location"    = $nsNetsuiteProject.custentity_atlas_svcs_mm_location.refName
        "custentity_atlas_svcs_mm_projectmngr" = $nsNetsuiteProject.custentity_atlas_svcs_mm_projectmngr.refName
        "jobType"                              = $nsNetsuiteProject.jobType.refName
        "subsidiary"                           = $nsNetsuiteProject.subsidiary.refName
        "DateCreated"                          = $nsNetsuiteProject.dateCreated
        "LastModified"                         = $nsNetsuiteProject.lastModifiedDate
        "DateCreatedInSql"                     = $null
        "DateModifiedInSql"                    = $null
        "IsDirty"                              = $null
        "SharePointSiteId"                     = $null
        "SharePointListId"                     = $null
        "SharePointDriveItemId"                = $null
        }
    $pretendSqlNetSuiteProject
    }
function delete-netSuiteContactFromNetSuite(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [string]$id

        ,[parameter(Mandatory=$false)]
        [psobject]$netsuiteParameters
        )

    invoke-netsuiteRestMethod -requestType DELETE -url "$($netsuiteParameters.uri)/contact/$id" -netsuiteParameters $netsuiteParameters #-Verbose 
    }
function get-netSuiteAuthHeaders(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [ValidateSet("DELETE","GET","POST","PATCH")]
        [string]$requestType
        
        ,[parameter(Mandatory = $true)]
        [ValidatePattern("http")]
        [string]$url
        
        ,[parameter(Mandatory=$true)]
        [hashtable]$oauthParameters

        ,[parameter(Mandatory=$true)]
        [string]$oauth_consumer_secret

        ,[parameter(Mandatory=$false)]
        [string]$oauth_token_secret

        ,[parameter(Mandatory=$true)]
        [string]$realm
        )

    Write-Verbose "get-netsuiteAuthHeaders()"
    $oauth_signature = get-oauthSignature -requestType $requestType -url $url -oauthParameters $oauthParameters -oauth_consumer_secret $oauth_consumer_secret -oauth_token_secret $oauth_token_secret

    #Irritatingly, we only include some predetermined oAuthParameters in the AuthHeader:
    $authHeaderString = ($oauthParameters.Keys | Sort-Object | ? {@("oauth_nonce","oauth_timestamp","oauth_consumer_key","oauth_token","oauth_signature_method","oauth_version") -contains $_} | % {
        "$_=`"$([uri]::EscapeDataString($oauthParameters[$_]))`""
        }) -join ","
    $authHeaderString += ",realm=`"$([uri]::EscapeDataString($realm))`""
    $authHeaderString += ",oauth_signature=`"$([uri]::EscapeDataString($oauth_signature))`""
    $authHeaders = @{"Authorization"="OAuth $authHeaderString"
        ;"Cache-Control"="no-cache"
#        ;"Accept"="application/swagger+json"
#        ;"Accept-Encoding"="gzip, deflate"
        }
    Write-Verbose "`$authHeaders = $($(
        $authHeaders.Keys | Sort-Object | % {
            "$_=$($authHeaders[$_])"
            }
        ) -join "&")"
    $authHeaders
    }
function get-netSuiteClientsFromNetSuite(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true,ParameterSetName="Query")]
            [ValidatePattern('^?[\w+][=][\w+]')]
            [string]$query
        ,[parameter(Mandatory=$true,ParameterSetName="Id")]
            [string]$clientId
        ,[parameter(Mandatory=$false,ParameterSetName="Query")]
            [parameter(Mandatory = $false,ParameterSetName="GetAll")]
            [parameter(Mandatory = $false,ParameterSetName="Id")]
            [psobject]$netsuiteParameters
        )

    Write-Verbose "`tget-netSuiteClientsFromNetSuite([$($query)])"
    if([string]::IsNullOrWhiteSpace($netsuiteParameters)){
        $netsuiteParameters = get-netsuiteParameters -connectTo Sandbox
        Write-Warning "NetSuite environment unspecified - connecting to Sandbox"
        }

    switch ($PsCmdlet.ParameterSetName){
        "Id"    {
            $customersEnumerated = invoke-netsuiteRestMethod -requestType GET -url "$($netsuiteParameters.uri)/customer/$clientId" -netsuiteParameters $netsuiteParameters 
            }
        default {
            $customers = invoke-netsuiteRestMethod -requestType GET -url "$($netsuiteParameters.uri)/customer$query" -netsuiteParameters $netsuiteParameters #-Verbose 
            #$customersEnumerated = [psobject[]]::new($customers.count)
            [array]$customersEnumerated = @($null) * $customers.count
            for ($i=0; $i -lt $customers.count;$i++) {
                write-progress -activity "Retrieving NetSuite Client details..." -Status "[$($i)]/[$($customers.count)]" -PercentComplete $(($i*100)/$customers.count)
                if($i%100 -eq 0){Write-Verbose "[$($i)]/[$($customers.count)] ($($i / $customers.count)%)"}
                $customersEnumerated[$i] = invoke-netsuiteRestMethod -requestType GET -url "$($customers.items[$i].links[0].href)/?expandSubResources=$true" -netsuiteParameters $netsuiteParameters 
                }
            }
        }


    $customersEnumerated
    }
function get-netSuiteClientFromSqlCache{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $false)]
        [ValidatePattern('^[WHERE]')]
        [string]$sqlWhereClause
        ,[parameter(Mandatory = $true)]
        [System.Data.Common.DbConnection]$dbConnection
        )
    Write-Verbose "get-netSuiteClientFromSqlCache [$sqlWhereClause]"
    <#$sql = "SELECT  a.AccountName, a.NsInternalId, a.NsExternalId, a.RecordType, a.entityStatus, a.DateCreated, a.LastModified, a.DateCreatedInSql, a.DateModifiedInSql, a.IsDirty, a.SharePointDocLibGraphDriveId FROM t_ACCOUNTS a
            INNER JOIN (
                SELECT  AccountName, MAX(DateModifiedInSql) AS MaxDate
                FROM t_ACCOUNTS
                GROUP BY AccountName) am
                ON a.AccountName = am.AccountName 
                    AND a.DateModifiedInSql = am.MaxDate 
            $sqlWhereClause" #>
    $sql = "SELECT a.AccountName, a.NsInternalId, a.NsExternalId, a.RecordType, a.entityStatus, a.DateCreated, a.LastModified, a.DateCreatedInSql, a.DateModifiedInSql, a.IsDirty, a.SharePointDocLibGraphDriveId FROM v_ACCOUNTS_Current a $sqlWhereClause"
    Write-Verbose "`t$sql"
    $result = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $dbConnection
    $result
    }
function get-netSuiteContactFromNetSuite(){
[cmdletbinding()]
Param (
    [parameter(Mandatory = $false,ParameterSetName="Query")]
        [ValidatePattern('^?[\w+][=][\w+]|^$')]
        [string]$query
    ,[parameter(Mandatory=$true,ParameterSetName="Id")]
        [string]$contactId
    ,[parameter(Mandatory=$false,ParameterSetName="Query")]
        [parameter(Mandatory = $false,ParameterSetName="GetAll")]
        [parameter(Mandatory = $false,ParameterSetName="Id")]
        [psobject]$netsuiteParameters
    )

    Write-Verbose "`tget-netSuiteContactFromNetSuite([$($query)])"
    if([string]::IsNullOrWhiteSpace($netsuiteParameters)){
        $netsuiteParameters = get-netsuiteParameters -connectTo Sandbox
        Write-Warning "NetSuite environment unspecified - connecting to Sandbox"
        }

    switch ($PsCmdlet.ParameterSetName){
        "Id"    {
            $contactsEnumerated = invoke-netsuiteRestMethod -requestType GET -url "$($netsuiteParameters.uri)/contact/$contactId" -netsuiteParameters $netsuiteParameters 
            }
        default {
            $contacts = invoke-netsuiteRestMethod -requestType GET -url "$($netsuiteParameters.uri)/contact$query" -netsuiteParameters $netsuiteParameters #-Verbose 
            [array]$contactsEnumerated = [psobject[]]::new($contacts.count)
            for ($i=0; $i -lt $contacts.count;$i++) {
                write-progress -activity "Retrieving NetSuite Contact details..." -Status "[$($i)]/[$($contactsEnumerated.count)]" -PercentComplete $(($i*100)/$contactsEnumerated.count)
                $url = "$($contacts.items[$i].links[0].href)" + "/?expandSubResources=True"
                if($i%100 -eq 0){Write-Verbose "[$($i)]/[$($contacts.count)] ($($i / $contacts.count)%)"}
                $contactsEnumerated[$i] = invoke-netsuiteRestMethod -requestType GET -url $url -netsuiteParameters $netsuiteParameters 
                }            
            }
        }

    $contactsEnumerated
    }
function get-netSuiteCustomListValues(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
            [ValidateSet("customer","contact","clientSector","clientType","clientRating","customerstatus","subsidiary")]
            [string]$objectType

        ,[parameter(Mandatory=$false)]
            [psobject]$netsuiteParameters
        )

    Write-Verbose "`tget-netSuiteProjectFromNetSuite([$($query)])"
    if([string]::IsNullOrWhiteSpace($netsuiteParameters)){
        $netsuiteParameters = get-netsuiteParameters -connectTo Sandbox
        Write-Warning "NetSuite environment unspecified - connecting to Sandbox"
        }

    switch($objectType){
        "customer"       {$endpoint = "customer"}
        "contact"        {$endpoint = "contact"}
        "clientSector"   {$endpoint = "customlist_ant_clientsector"}
        "clientType"     {$endpoint = "customlist_clienttype"}
        "clientRating"   {$endpoint = "customlist_clientrating"}
        "customerstatus" {$endpoint = "customerstatus"}
        "subsidiary"     {$endpoint = "subsidiary"}
        }

    try{
        $customListValues = invoke-netsuiteRestMethod -requestType GET -url "$($netsuiteParameters.uri)/$endpoint" -netsuiteParameters $netsuiteParameters #-Verbose 
        $customListValuesEnumerated = [psobject[]]::new($customListValues.count)
        for ($i=0; $i -lt $customListValues.count;$i++) {
            if($i%100 -eq 0){Write-Verbose "[$($i)]/[$($customListValues.count)] ($($i / $customListValues.count)%)"}
            $customListValuesEnumerated[$i] = invoke-netsuiteRestMethod -requestType GET -url "$($customListValues.items[$i].links[0].href)/?expandSubResources=$true" -netsuiteParameters $netsuiteParameters 
            }
        }
    catch{
        #Weird 405 error on customerStatus prevents us listing the available values, so we have to trial-and-error
        if((ConvertFrom-Json $_.ErrorDetails).status -eq 405){
            while($consecutiveErrors -le 2){
                try{
                    $i++
                    [array]$customListValuesEnumerated += invoke-netsuiteRestMethod -requestType GET -url "$($netsuiteParameters.uri)/$endpoint/$i" -netsuiteParameters $netsuiteParameters -ErrorAction SilentlyContinue
                    $consecutiveErrors = 0
                    }
                catch{$consecutiveErrors++}
                }
            }
        }
    $customListValuesEnumerated
    }
function get-netSuiteEmployeesFromNetSuite(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $false)]
        [ValidatePattern('^?[\w+][=][\w+]')]
        [string]$query
        ,[parameter(Mandatory=$false)]
        [psobject]$netsuiteParameters
        ,[parameter(Mandatory=$false)]
        [ValidateSet('True','False')]
        [string]$allNetsuiteEmployees
        ,[parameter(Mandatory=$false)]
        [ValidateSet('Actively Employed','Probation','Terminated')]
        [string]$employeestatus
        )
    Write-Verbose "`tget-netSuiteEmployeesFromNetSuite([$($query)])"
    if([string]::IsNullOrWhiteSpace($netsuiteParameters)){
        $netsuiteParameters = get-netsuiteParameters -connectTo Sandbox
        Write-Warning "NetSuite environment unspecified - connecting to Sandbox"
        }
    Write-Host "Getting all @anthesisgroup.com employees - please use allNetsuiteEmployees switch to pull everything back with no filters" -ForegroundColor Cyan
    
    #employeestatus
    If($employeestatus){
    $Netquery = "?q=email CONTAIN `"@anthesisgroup.com`""
    $Netquery += "AND email CONTAIN_NOT `"netsuitebot@anthesisgroup.com`""
    $Netquery += "AND email CONTAIN_NOT `"purchasing@anthesisgroup.com`""
    $Netquery += "AND email CONTAIN_NOT `"noemail@anthesisgroup.com`""
    $Netquery += "AND email CONTAIN_NOT `"group.finance@anthesisgroup.com`""
    $Netquery += "AND email CONTAIN_NOT `"pieface@anthesisgroup.com`""
    $Netquery += "AND employeestatus CONTAIN `"$($employeestatus)`""

    $employees = invoke-netsuiteRestMethod -requestType GET -url "$($netsuiteParameters.uri)/employee/$Netquery" -netsuiteParameters $netsuiteParameters #-Verbose 
    $employeesEnumerated = [psobject[]]::new($employees.count)
    for ($i=0; $i -lt $employees.count;$i++) {
        $employeesEnumerated[$i] = invoke-netsuiteRestMethod -requestType GET -url $employees.items[$i].links[0].href -netsuiteParameters $netsuiteParameters
        }
    $employeesEnumerated
    }

    #customquery
    If($query){
    $Netquery = "?q=email CONTAIN `"@anthesisgroup.com`""
    $Netquery += "AND email CONTAIN_NOT `"netsuitebot@anthesisgroup.com`""
    $Netquery += "AND email CONTAIN_NOT `"purchasing@anthesisgroup.com`""
    $Netquery += "AND email CONTAIN_NOT `"noemail@anthesisgroup.com`""
    $Netquery += "AND email CONTAIN_NOT `"group.finance@anthesisgroup.com`""
    $Netquery += "AND email CONTAIN_NOT `"pieface@anthesisgroup.com`""
    $Netquery += " AND $($query)`""

    $employees = invoke-netsuiteRestMethod -requestType GET -url "$($netsuiteParameters.uri)/employee/$Netquery" -netsuiteParameters $netsuiteParameters #-Verbose 
    $employeesEnumerated = [psobject[]]::new($employees.count)
    for ($i=0; $i -lt $employees.count;$i++) {
        $employeesEnumerated[$i] = invoke-netsuiteRestMethod -requestType GET -url $employees.items[$i].links[0].href -netsuiteParameters $netsuiteParameters
        }
    $employeesEnumerated
    }

    #allAnthesisemployees
    If(!($query) -and !($employeestatus)){
    $Netquery = "?q=email CONTAIN `"@anthesisgroup.com`""
    $Netquery += "AND email CONTAIN_NOT `"netsuitebot@anthesisgroup.com`""
    $Netquery += "AND email CONTAIN_NOT `"purchasing@anthesisgroup.com`""
    $Netquery += "AND email CONTAIN_NOT `"noemail@anthesisgroup.com`""
    $Netquery += "AND email CONTAIN_NOT `"group.finance@anthesisgroup.com`""
    $Netquery += "AND email CONTAIN_NOT `"pieface@anthesisgroup.com`""
    $employees = invoke-netsuiteRestMethod -requestType GET -url "$($netsuiteParameters.uri)/employee/$Netquery" -netsuiteParameters $netsuiteParameters #-Verbose 
    $employeesEnumerated = [psobject[]]::new($employees.count)
    for ($i=0; $i -lt $employees.count;$i++) {
        $employeesEnumerated[$i] = invoke-netsuiteRestMethod -requestType GET -url $employees.items[$i].links[0].href -netsuiteParameters $netsuiteParameters
        }
    $employeesEnumerated
    }
    
    #allNetsuiteemployees
    If($allNetsuiteEmployees){
        $employees = invoke-netsuiteRestMethod -requestType GET -url "$($netsuiteParameters.uri)/employee" -netsuiteParameters $netsuiteParameters #-Verbose 
    $employeesEnumerated = [psobject[]]::new($employees.count)
    for ($i=0; $i -lt $employees.count;$i++) {
        $employeesEnumerated[$i] = invoke-netsuiteRestMethod -requestType GET -url $employees.items[$i].links[0].href -netsuiteParameters $netsuiteParameters
        }
    $employeesEnumerated
    }
}
function get-netSuiteMetadata(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $false)]
        [ValidateSet("customer","contact")]
        [string]$objectType

        ,[parameter(Mandatory=$false)]
        [psobject]$netsuiteParameters
        )

    Write-Verbose "get-netSuiteMetadata [$objectType]"
    if([string]::IsNullOrWhiteSpace($netsuiteParameters)){
        $netsuiteParameters = get-netsuiteParameters -connectTo Sandbox
        Write-Warning "NetSuite environment unspecified - connecting to Sandbox"
        }

    $metadata = invoke-netsuiteRestMethod -requestType GET -url "$($netsuiteParameters.uri)/metadata-catalog/$objectType" -netsuiteParameters $netsuiteParameters  #-Verbose 
    #$metadata = invoke-netsuiteRestMethod -requestType GET -url "$($netsuiteParameters.uri)/metadata-catalog?select=$objectType" -netsuiteParameters $netsuiteParameters #-Verbose 
    $metadata 
    }
function get-netSuiteOpportunityFromNetSuite(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $false)]
        [AllowEmptyString()] 
        [ValidatePattern('^?[\w+][=][\w+]|^$')]
        [string]$query

        ,[parameter(Mandatory=$false)]
        [psobject]$netsuiteParameters
        )

    Write-Verbose "`tget-netSuiteProjectFromNetSuite([$($query)])"
    if([string]::IsNullOrWhiteSpace($netsuiteParameters)){
        $netsuiteParameters = get-netsuiteParameters -connectTo Sandbox
        Write-Warning "NetSuite environment unspecified - connecting to Sandbox"
        }

    $opportunities = invoke-netsuiteRestMethod -requestType GET -url "$($netsuiteParameters.uri)/opportunity$query" -netsuiteParameters $netsuiteParameters #-Verbose 
    #$opportunitiesEnumerated = [psobject[]]::new($opportunities.count)
    [array]$opportunitiesEnumerated = @($null) * $opportunities.count
    for ($i=0; $i -lt $opportunities.count;$i++) {
        write-progress -activity "Retrieving NetSuite Opportunity details..." -Status "[$($i)]/[$($opportunities.count)]" -PercentComplete $(($i*100)/$opportunities.count)
        $opportunitiesEnumerated[$i] = invoke-netsuiteRestMethod -requestType GET -url $opportunities.items[$i].links[0].href -netsuiteParameters $netsuiteParameters 
        }
    $opportunitiesEnumerated
    }
function get-netSuiteOpportunityFromSqlCache{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $false)]
        [ValidatePattern('^[WHERE]')]
        [string]$sqlWhereClause
        ,[parameter(Mandatory = $true)]
        [System.Data.Common.DbConnection]$dbConnection
        )
    Write-Verbose "get-netSuiteOpportunityFromSqlCache [$sqlWhereClause]"
    $sql = "SELECT NsInternalId, NsExternalId, AccountNsInternalId, OpportunityName, OpportunityNumber, entityId, entityStatus, entityNexus, custbody_project_template, tranId, status, probability, custbody_industry, subsidiary, DateCreated, LastModified, DateCreatedInSql, DateModifiedInSql, IsDirty, SharePointDriveItemId
                FROM v_OPPOPTUNITIES_Current 
            $sqlWhereClause"
    Write-Verbose "`t$sql"
    $result = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $dbConnection
    if($result -eq 1){Write-Verbose "`t`tSUCCESS!"}
    else{Write-Verbose "`t`tFAILURE :( - Code: $result"}
    $result
    }
function get-netSuitePaddedCode(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [string]$unpaddedCode
        ,[parameter(Mandatory = $false)]
        [string]$padToXDigits = 7
        )

    if($unpaddedCode -match '^[a-zA-Z]+'){$prefix = $Matches[0]}
    else{Write-Error "Prefix not found";break}

    if($unpaddedCode -match '[0-9]+$'){$suffix = $Matches[0]}
    else{Write-Error "Suffix not found";break}

    $prefix + $("{0:d$padToXDigits}" -f [int]$suffix)    
    }
function get-netSuiteParameters(){
    [cmdletbinding()]
    Param([parameter(Mandatory = $false)]
        [ValidateSet("Production","Sandbox")]
        [string]$connectTo = "Sandbox"
        )
    Write-Verbose "get-netsuiteParameters()"
    if($connectTo -eq "Production"){
        $placesToLook = @(
            "$env:USERPROFILE\Desktop\netsuite_live.txt"
            "$env:USERPROFILE\Downloads\netsuite_live.txt"
            ,"$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\netsuite_live.txt"
            )
        }
    else{
        $placesToLook = @(
            "$env:USERPROFILE\Desktop\netsuite.txt"
            "$env:USERPROFILE\Downloads\netsuite_sandbox.txt"
            ,"$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\netsuite.txt"
            )
        
        }
    for($i=0; $i -lt $placesToLook.Count; $i++){
        if(Test-Path $placesToLook[$i]){
            $pathToEncryptedCsv = $placesToLook[$i]
            continue
            }
        }
    if([string]::IsNullOrWhiteSpace($pathToEncryptedCsv)){
        Write-Error "NetSuite Paramaters CSV file not found in any of these locations: $($placesToLook -join ", ")"
        break
        }
    else{
        Write-Verbose "Importing NetSuite Paramaters fvrom [$pathToEncryptedCsv]"
        $importedParameters = import-encryptedCsv $pathToEncryptedCsv
        $importedParameters.oauth_consumer_key = $importedParameters.oauth_consumer_key.ToUpper()
        $importedParameters.oauth_consumer_secret = $importedParameters.oauth_consumer_secret.ToLower()
        $importedParameters.oauth_token = $importedParameters.oauth_token.ToUpper()
        $importedParameters.oauth_token_secret = $importedParameters.oauth_token_secret.ToLower()
        $importedParameters
        }
    }
function get-netSuiteProjectFromNetSuite(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true,ParameterSetName="Query")]
            [ValidatePattern('^?[\w+][=][\w+]|^$')]
            [string]$query
        ,[parameter(Mandatory=$true,ParameterSetName="Id")]
            [string]$projectId
        ,[parameter(Mandatory=$false,ParameterSetName="Query")]
            [parameter(Mandatory = $false,ParameterSetName="GetAll")]
            [parameter(Mandatory = $false,ParameterSetName="Id")]
            [psobject]$netsuiteParameters
        )

    Write-Verbose "`tget-netSuiteProjectFromNetSuite([$($query)])"
    if([string]::IsNullOrWhiteSpace($netsuiteParameters)){
        $netsuiteParameters = get-netsuiteParameters -connectTo Sandbox
        Write-Warning "NetSuite environment unspecified - connecting to Sandbox"
        }

    switch ($PsCmdlet.ParameterSetName){
        "Id"    {
            $projectsEnumerated = invoke-netsuiteRestMethod -requestType GET -url "$($netsuiteParameters.uri)/job/$projectId" -netsuiteParameters $netsuiteParameters 
            }
        default {
            $projects = invoke-netsuiteRestMethod -requestType GET -url "$($netsuiteParameters.uri)/job$query" -netsuiteParameters $netsuiteParameters #-Verbose 
            #$projectsEnumerated = [psobject[]]::new($projects.count)
            [array]$projectsEnumerated = @($null) * $projects.Count
                                    for ($i=0; $i -lt $projects.count;$i++) {
        write-progress -activity "Retrieving NetSuite Project details..." -Status "[$($i)]/[$($projects.count)]" -PercentComplete $(($i*100)/$projects.count)
        $projectsEnumerated[$i] = invoke-netsuiteRestMethod -requestType GET -url $projects.items[$i].links[0].href -netsuiteParameters $netsuiteParameters 
        }
            }
        }

    $projectsEnumerated
    }
function get-netSuiteProjectFromSqlCache{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $false)]
        [ValidatePattern('^[WHERE]')]
        [string]$sqlWhereClause
        ,[parameter(Mandatory = $true)]
        [System.Data.Common.DbConnection]$dbConnection
        )
    Write-Verbose "get-netSuiteProjectFromSqlCache [$sqlWhereClause]"
    $sql = "SELECT  p.NsInternalId, p.NsExternalId, p.AccountNsInternalId, p.ProjectName, p.ProjectNumber, p.entityId, p.entityStatus, p.custentity_atlas_svcs_mm_department, p.custentity_ant_projectsector, p.custentity_ant_projectsource, p.custentity_atlas_svcs_mm_location, p.custentity_atlas_svcs_mm_projectmngr, p.jobType, p.subsidiary, p.DateCreated, p.LastModified, p.IsDirty, p.DateCreatedInSql, p.DateModifiedInSql, p.SharePointDriveItemId FROM t_PROJECTS p
            INNER JOIN (
                SELECT  entityId, MAX(DateModifiedInSql) AS MaxDate
                FROM t_PROJECTS
                GROUP BY entityId) pm
                ON p.entityId = pm.entityId 
                    AND p.DateModifiedInSql = pm.MaxDate 
            $sqlWhereClause"
    Write-Verbose "`t$sql"
    $result = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $dbConnection
    if($result -eq 1){Write-Verbose "`t`tSUCCESS!"}
    else{Write-Verbose "`t`tFAILURE :( - Code: $result"}
    $result
    }
function get-oAuthSignature(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [ValidateSet("DELETE","GET","POST","PATCH")]
        [string]$requestType
        
        ,[parameter(Mandatory = $true)]
        [ValidatePattern("http")]
        [string]$url
        
        ,[parameter(Mandatory=$true)]
        [hashtable]$oauthParameters

        ,[parameter(Mandatory=$true)]
        [string]$oauth_consumer_secret

        ,[parameter(Mandatory=$false)]
        [string]$oauth_token_secret
        )
    Write-Verbose "get-oauthSignature()"
    $requestType = $requestType.ToUpper()
                           
    $encodedUrl = [uri]::EscapeDataString($url.ToLower())

    $oAauthParamsString = (
        $oauthParameters.Keys | Sort-Object | % {
            if(@("realm","oauth_signature") -notcontains $_){
                "$_=$($oauthParameters[$_])"
                }
            }
        ) -join "&"
    $encodedOAuthParamsString = [uri]::EscapeDataString($oAauthParamsString)

    Write-Verbose "`tUnencoded base_string: [$($requestType + "&" + $url + "&" + $oAauthParamsString)]"
    $base_string = $requestType + "&" + $encodedUrl + "&" + $encodedOAuthParamsString
    $key = $oauth_consumer_secret + "&" + $oauth_token_secret
    Write-Verbose "`tEncoded base_string: [$base_string]"

    Switch($oauthParameters["oauth_signature_method"]){
        "HMAC-SHA1" {
            $cryptoFunction = new-object System.Security.Cryptography.HMACSHA1
            }
        "HMAC-SHA256" {
            $cryptoFunction = new-object System.Security.Cryptography.HMACSHA256
            }
        "HMAC-SHA384" {
            $cryptoFunction = new-object System.Security.Cryptography.HMACSHA384
            }
        "HMAC-SHA512" {
            $cryptoFunction = new-object System.Security.Cryptography.HMACSHA512
            }
        default {
            Write-Error "Unsupported oauth_signature_method [$_]"
            break
            }
        }

    $cryptoFunction.Key = [System.Text.Encoding]::ASCII.GetBytes($key)
    $oauth_signature = [System.Convert]::ToBase64String($cryptoFunction.ComputeHash([System.Text.Encoding]::ASCII.GetBytes($base_string)))
    Write-Verbose "`t`$oauth_signature = [$oauth_signature]"
    $oauth_signature
    }
function invoke-netSuiteRestMethod(){
    [cmdletbinding()]
    Param(
        [parameter(Mandatory = $true)]
        [ValidateSet("DELETE","GET","POST","PATCH")]
        [string]$requestType
        ,[parameter(Mandatory = $true)]
        [ValidatePattern("http")]
        [string]$url
        ,[parameter(Mandatory=$false)]
        [psobject]$netsuiteParameters
        ,[parameter(Mandatory=$false)]
        [hashtable]$requestBodyHashTable
        )
    if(!$netsuiteParameters){$netsuiteParameters = get-netsuiteParameters}
    
    if($url -match "\?"){
        $parameters = $url.Split("?")[1]
        $hostUrl = $url.Split("?")[0]
        }
    else{
        $hostUrl=$url
        $parameters = ""
        }

    $oauth_nonce = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes([System.DateTime]::Now.Ticks.ToString()))
    $oauth_timestamp = [int64](([datetime]::UtcNow)-(Get-Date "1970-01-01")).TotalSeconds

    $oAuthParamsForSigning = @{}
    #Add standard oAuth 1.0 parameters
    $oAuthParamsForSigning.Add("oauth_nonce",$oauth_nonce)
    $oAuthParamsForSigning.Add("oauth_timestamp",$oauth_timestamp)
    $oAuthParamsForSigning.Add("oauth_consumer_key",$netsuiteParameters.oauth_consumer_key)
    $oAuthParamsForSigning.Add("oauth_token",$netsuiteParameters.oauth_token)
    $oAuthParamsForSigning.Add("oauth_signature_method",$netsuiteParameters.oauth_signature_method)
    $oAuthParamsForSigning.Add("oauth_version",$netsuiteParameters.oauth_version)
    #$oAuthParamsForSigning.Add([uri]::EscapeDataString("ignoreMandatoryFields"),[uri]::EscapeDataString($true))
    #Add parameters from url
    $parameters.Split("&") | % {
        if(![string]::IsNullOrWhiteSpace($_.Split("=")[0])){
            $oAuthParamsForSigning.Add([uri]::EscapeDataString($_.Split("=")[0]),[uri]::EscapeDataString($_.Split("=")[1])) #Weirdly, these extra paramaters have to be Encoded twice...
            #write-host -f Green "$([uri]::EscapeDataString($_.Split("=")[0]),[uri]::EscapeDataString($_.Split("=")[1]))"
            }
        }
    
    $netsuiteRestHeaders = get-netsuiteAuthHeaders -requestType $requestType -url $hostUrl -oauthParameters $oAuthParamsForSigning  -oauth_consumer_secret $netsuiteParameters.oauth_consumer_secret -oauth_token_secret $netsuiteParameters.oauth_token_secret -realm $netsuiteParameters.realm

    Write-Verbose "Invoke-RestMethod -Uri $([uri]::EscapeUriString($url)) -Headers $(stringify-hashTable $netsuiteRestHeaders) -Method $requestType -ContentType application/swagger+json -Body $(stringify-hashTable $requestBodyHashTable)"
    if($requestType -eq "GET"){
        try{
            $partialDataset = Invoke-RestMethod -Uri $([uri]::EscapeUriString($url)) -Headers $netsuiteRestHeaders -Method $requestType -ContentType "application/swagger+json" #-Proxy 'http://127.0.0.1:8888'
            if($partialDataset.totalResults -ne $partialDataset.count){ #If the query has been paginated
                if($partialDataset.offset -eq 0){
                    $fullDataSet = New-Object object[] $partialDataSet.totalResults
                    Write-Verbose "`$fullDataSet.count = [$($fullDataSet.Count)]"
                    }
                do{
                    for($i = 0; $i -lt $partialDataset.count; $i++){ #Fill $fullDataset with the contents of $partialDataset
                        $fullDataset[$i+$partialDataset.offset] = $partialDataset.items[$i]
                        if($i%100 -eq 0){
                            Write-Verbose "[$($i+$partialDataset.offset)]/[$($partialDataSet.totalResults)] ($([System.Math]::Floor(($i+$partialDataset.offset)*100 / $partialDataSet.totalResults))%)"
                            write-progress -activity "Retrieving results from [$($url)]" -Status "[$($i+$partialDataset.offset)]/[$($partialDataSet.totalResults)]" -PercentComplete $(($i+$partialDataset.offset*100)/$partialDataSet.totalResults)
                            }
                        }
                    $nextUrl = [uri]::EscapeUriString($($partialDataset.links | ? {$_.rel -eq "next"}).href) #Check if there are more results to retrieve
                    if([string]::IsNullOrWhiteSpace($nextUrl)){}#$partialDataset.links.rel | % {Write-Verbose $_}}
                    else{
                        Write-Verbose "`tNext URL: [$($nextUrl)] (parameters [$($parameters)])"
                        if($nextUrl -match "limit=" -and $parameters -match "limit="){$parameters = $($parameters -replace '(?<=(^limit=))\w*(?=(&))','').Replace("limit=","") -replace '^&',''} #Trim off any leading limit parameter from the previous iteration
                        if($nextUrl -match "limit=" -and $parameters -match "offset="){$parameters = $($parameters -replace '(?<=(^offset=))\w*(?=(&))','').Replace("offset=","") -replace '^&',''} #Trim off any leading offset parameter from the previous iteration
                        if(![string]::IsNullOrWhiteSpace($parameters)){$nextUrl = "$nextUrl&$parameters"} #Weirldy links.next.href doesn't include the original query, so this will [Index was outside the bounds of the array.] if we don't resupply it manually
                        Write-Verbose "`tUpdated URL: [$($nextUrl)] (parameters [$($parameters)])"
                        $partialDataset = invoke-netSuiteRestMethod -requestType $requestType -url $nextUrl -netsuiteParameters $netsuiteParameters
                        }
                    }
                while($partialDataset.hasMore -eq $true)

                $partialDataset.items = $fullDataset
                $partialDataset.count = $partialDataset.items.count
                }
            $partialDataset
            }
        catch{
            if($_.Exception -match "401"){Write-Warning "401: Unauthorised access attempt to [$($url)]"}
            else{Write-Error $_}
            }
        }
    else{
        if($requestType -ne "DELETE"){
            $bodyJson = ConvertTo-Json -InputObject $requestBodyHashTable
            Write-Verbose $bodyJson
            $bodyJsonEncoded = [System.Text.Encoding]::UTF8.GetBytes($bodyJson)
            }
        Invoke-RestMethod -Uri $([uri]::EscapeUriString($url)) -Headers $netsuiteRestHeaders -Method $requestType -ContentType "application/json" -Body $bodyJsonEncoded
        }
    }
function sync-netSuiteClientsFromNetSuiteToSql(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $false)]
        [PSCustomObject]$netSuiteParams 
        ,[parameter(Mandatory = $true)]
        [ValidateSet("Full","Delta")]
        [string]$sync
        ,[parameter(Mandatory = $true)]
        [System.Data.Common.DbConnection]$sqlDbConn
        )

    if([string]::IsNullOrWhiteSpace($netsuiteParameters)){
        $netsuiteParameters = get-netsuiteParameters -connectTo Sandbox
        Write-Warning "NetSuite environment unspecified - connecting to Sandbox"
        }

    switch ($sync) {
        "Delta" {
            $sql = "SELECT MAX(LastModified) FROM v_ACCOUNTS_Current"
            $mostRecentModifiedDateInSqlCache = Execute-SQLQueryOnSQLDB -query $sql -queryType Scalar -sqlServerConnection $sqlDbConn

            $query = "?q=lastModifiedDate ON_OR_AFTER `"$(Get-Date -UFormat "%d/%m/%y" $mostRecentModifiedDateInSqlCache)`""
            $nsClients = get-netSuiteClientsFromNetSuite -netsuiteParameters $netSuiteParams -query $query
            }
        "Full" {
            $nsClients = get-netSuiteClientsFromNetSuite -netsuiteParameters $netSuiteParams 
            }
        }

    $nsClients | % {
        $thisClient = $_
        try{
            add-netsuiteAccountToSqlCache -nsNetsuiteAccount $_ -accountType Client -dbConnection $sqlDbConn
            }
        catch{
            [array]$problemClients += $thisClient
            }
        }

    #E-mail $problemClients to IT
    $problemClients
    }
function sync-netSuiteOpportunitiesFromNetSuiteToSql(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $false)]
        [PSCustomObject]$netSuiteParams 
        ,[parameter(Mandatory = $true)]
        [ValidateSet("Full","Delta")]
        [string]$sync
        ,[parameter(Mandatory = $true)]
        [System.Data.Common.DbConnection]$sqlDbConn
        )

    if([string]::IsNullOrWhiteSpace($netsuiteParameters)){
        $netsuiteParameters = get-netsuiteParameters -connectTo Sandbox
        Write-Warning "NetSuite environment unspecified - connecting to Sandbox"
        }

    switch ($sync) {
        "Delta" {
            $sql = "SELECT MAX(LastModified) FROM v_OPPORTUNITIES_Current"
            $mostRecentModifiedDateInSqlCache = Execute-SQLQueryOnSQLDB -query $sql -queryType Scalar -sqlServerConnection $sqlDbConn

            $query = "?q=lastModifiedDate ON_OR_AFTER `"$(Get-Date -UFormat "%d/%m/%y" $mostRecentModifiedDateInSqlCache)`""
            $nsOpps = get-netSuiteOpportunityFromNetSuite -netsuiteParameters $netSuiteParams -query $query
            }
        "Full" {
            $nsOpps = get-netSuiteOpportunityFromNetSuite -netsuiteParameters $netSuiteParams
            }
        }

    $nsOpps | % {
        $thisOpp = $_
        try{
            add-netsuiteOpportunityToSqlCache -nsNetsuiteOpportunity $_ -dbConnection $sqlDbConn
            }
        catch{
            [array]$problemOpps += @($thisOpp,$_)
            }
        }

    #E-mail $problemClients to IT
    $problemOpps
    }
function sync-netSuiteProjectsFromNetSuiteToSql(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $false)]
        [PSCustomObject]$netSuiteParams 
        ,[parameter(Mandatory = $true)]
        [ValidateSet("Full","Delta")]
        [string]$sync
        ,[parameter(Mandatory = $true)]
        [System.Data.Common.DbConnection]$sqlDbConn
        )
    
    if([string]::IsNullOrWhiteSpace($netsuiteParameters)){
        $netsuiteParameters = get-netsuiteParameters -connectTo Sandbox
        Write-Warning "NetSuite environment unspecified - connecting to Sandbox"
        }

    switch ($sync) {
        "Delta" {
            $sql = "SELECT MAX(LastModified) FROM v_PROJECTS_Current"
            $mostRecentModifiedDateInSqlCache = Execute-SQLQueryOnSQLDB -query $sql -queryType Scalar -sqlServerConnection $sqlDbConn

            $query = "?q=lastModifiedDate ON_OR_AFTER `"$(Get-Date -UFormat "%d/%m/%y" $mostRecentModifiedDateInSqlCache)`""
            $nsProjects = get-netSuiteProjectFromNetSuite -netsuiteParameters $netSuiteParams -query $query
            }
        "Full" {
            $nsProjects = get-netSuiteProjectFromNetSuite -netsuiteParameters $netSuiteParams
            }
        }

    $nsProjects | % {
        $thisProject = $_
        try{
            add-netsuiteProjectToSqlCache -nsNetsuiteProject $thisProject -dbConnection $sqlDbConn
            }
        catch{
            [array]$problemProjects += @($thisProject,$_)
            }
        }

    #E-mail $problemClients to IT
    $problemProjects
    }
function update-netSuiteAccountInSqlCache(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true, ParameterSetName="nsNetSuiteAccount")]
            [PSCustomObject]$nsNetsuiteAccount 
        ,[parameter(Mandatory = $true, ParameterSetName="sqlNetSuiteAccount")]
            [PSCustomObject]$sqlNetsuiteAccount 
        ,[parameter(Mandatory = $true, ParameterSetName="nsNetSuiteAccount")]
            [parameter(Mandatory = $true, ParameterSetName="sqlNetSuiteAccount")]
            [System.Data.Common.DbConnection]$dbConnection
        ,[parameter(Mandatory = $false, ParameterSetName="nsNetSuiteAccount")]
            [parameter(Mandatory = $false, ParameterSetName="sqlNetSuiteAccount")]
            [switch]$isDirty
        ,[parameter(Mandatory = $false, ParameterSetName="nsNetSuiteAccount")]
            [parameter(Mandatory = $false, ParameterSetName="sqlNetSuiteAccount")]
            [switch]$isNotDirty
        )
    switch ($PsCmdlet.ParameterSetName){
        'nsNetSuiteAccount' {$sqlNetsuiteAccount = convert-nsNetSuiteAccountToSqlNetSuiteAccount -nsNetsuiteAccount $nsNetsuiteAccount}
        }

    Write-Verbose "update-netSuiteAccountInSqlCache [$($sqlNetsuiteAccount.AccountName)]"
    #Check record exists in SQL
    $sql = "SELECT TOP 1 AccountName, NsInternalId, LastModified FROM t_ACCOUNTS WHERE NsInternalId = '$($sqlNetsuiteAccount.NsInternalId)' ORDER BY LastModified Desc"
    $preExistingRecord = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $dbConnection
    
    if($preExistingRecord){
        if([string]::IsNullOrWhiteSpace($preExistingRecord.NsInternalId)){
            Write-Error "No NsInsternalId found on sql record [$($sqlNetsuiteAccount.AccountName)]. Cannot identify unique record, so cannot UPDATE."
            break
            }
        else{
            #Generate SQL statement
            $fieldsToUpdate = $sqlNetsuiteAccount.PSObject.Properties | ? {$_.Value -ne $null}
            if($fieldsToUpdate){
                $sql = "UPDATE t_ACCOUNTS "
                $sql += "SET "
                $fieldsToUpdate | % {
                    $thisField = $_
                    switch($_.Name){
                        "AccountName"                  {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetsuiteAccount.AccountName -dataType String), "}
                        "NsInternalId"                 {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetsuiteAccount.NsInternalId -dataType String), "}
                        "NsExternalId"                 {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetsuiteAccount.NsExternalId -dataType String), "}
                        "RecordType"                   {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetsuiteAccount.RecordType -dataType String), "}
                        "entityStatus"                 {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetsuiteAccount.entityStatus -dataType String), "}
                        "DateCreated"                  {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetsuiteAccount.dateCreated -dataType Date), "}
                        "LastModified"                 {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetsuiteAccount.LastModified -dataType Date), "}
                        "DateCreatedInSql"             {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetsuiteAccount.DateCreatedInSql -dataType Date), "}
                        "DateModifiedInSql"            {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetsuiteAccount.DateModifiedInSql -dataType Date), "}
                        "IsDirty"                      {
                            if($isDirty)                    {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $true -dataType Boolean), "}
                            elseif($isNotDirty)             {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $false -dataType Boolean), "}
                            else                            {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetsuiteAccount.IsDirty -dataType Boolean), "}
                                                        }
                        "SharePointDocLibGraphListId"  {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetsuiteAccount.SharePointDocLibGraphListId -dataType String), "}
                        "SharePointDocLibGraphDriveId" {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetsuiteAccount.SharePointDocLibGraphDriveId -dataType String), "}    
                        }
                    }
                $sql = $sql.TrimEnd(", ")
                $sql += " WHERE NsInternalId = $(sanitise-forSqlValue -value $preExistingRecord.NsInternalId -dataType String) "
                $sql += "AND LastModified = $(sanitise-forSqlValue -value $preExistingRecord.LastModified -dataType Date) "
                Write-Verbose "`t$sql"
                Execute-SQLQueryOnSQLDB -query $sql -queryType NonQuery -sqlServerConnection $dbConnection
                }
            }
        }
    else{Write-Error "Record with NsInsternalId [$($sqlNetsuiteAccount.NsExternalId)] does not exist in database. Cannot UPDATE.";break}
    }
function update-netSuiteOpportunityInSqlCache(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true, ParameterSetName="nsNetSuiteOpportunity")]
            [PSCustomObject]$nsNetSuiteOpportunity 
        ,[parameter(Mandatory = $true, ParameterSetName="sqlNetSuiteOpportunity")]
            [PSCustomObject]$sqlNetSuiteOpportunity 
        ,[parameter(Mandatory = $true, ParameterSetName="nsNetSuiteOpportunity")]
            [parameter(Mandatory = $true, ParameterSetName="sqlNetSuiteOpportunity")]
            [System.Data.Common.DbConnection]$dbConnection
        ,[parameter(Mandatory = $false, ParameterSetName="nsNetSuiteOpportunity")]
            [parameter(Mandatory = $false, ParameterSetName="sqlNetSuiteOpportunity")]
            [switch]$isDirty
        ,[parameter(Mandatory = $false, ParameterSetName="nsNetSuiteOpportunity")]
            [parameter(Mandatory = $false, ParameterSetName="sqlNetSuiteOpportunity")]
            [switch]$isNotDirty
        )
    switch ($PsCmdlet.ParameterSetName){
        'nsNetSuiteOpportunity' {$sqlNetSuiteOpportunity = convert-nsNetSuiteOpportunityToSqlNetSuiteOpportunity -nsNetSuiteOpportunity $nsNetSuiteOpportunity}
        }

    Write-Verbose "update-netSuiteOpportunityInSqlCache [$($sqlNetSuiteOpportunity.entityId)]"
    #Check record exists in SQL
    $sql = "SELECT TOP 1 OpportunityName, NsInternalId, LastModified FROM t_OPPORTUNITIES WHERE NsInternalId = '$($sqlNetSuiteOpportunity.NsInternalId)' ORDER BY LastModified Desc"
    $preExistingRecord = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $dbConnection
    
    if($preExistingRecord){
        if([string]::IsNullOrWhiteSpace($preExistingRecord.NsInternalId)){
            Write-Error "No NsInsternalId found on sql record [$($sqlNetSuiteOpportunity.entityId)]. Cannot identify unique record, so cannot UPDATE."
            break
            }
        else{
            #Generate SQL statement
            $fieldsToUpdate = $sqlNetSuiteOpportunity.PSObject.Properties | ? {$_.Value -ne $null}
            if($fieldsToUpdate){
                $sql = "UPDATE t_OPPORTUNITIES "
                $sql += "SET "
                $fieldsToUpdate | % {
                    $thisField = $_
                    switch($_.Name){
                        "NsInternalId"                 {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteOpportunity.NsInternalId -dataType String), "}
                        "NsExternalId"                 {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteOpportunity.NsExternalId -dataType String), "}
                        "AccountNsInternalId"          {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteOpportunity.AccountNsInternalId -dataType String), "}
                        "OpportunityName"              {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteOpportunity.OpportunityName -dataType String), "}
                        "OpportunityNumber"            {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteOpportunity.OpportunityNumber -dataType String), "}
                        "entityId"                     {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteOpportunity.entityId -dataType String), "}
                        "entityStatus"                 {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteOpportunity.entityStatus -dataType String), "}
                        "entityNexus"                  {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteOpportunity.entityNexus -dataType String), "}
                        "custbody_project_template"    {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteOpportunity.custbody_project_template -dataType String), "}
                        "tranId"                       {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteOpportunity.tranId -dataType String), "}
                        "status"                       {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteOpportunity.status -dataType String), "}
                        "probability"                  {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteOpportunity.probability -dataType String), "}
                        "custbody_industry"            {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteOpportunity.custbody_industry -dataType String), "}
                        "subsidiary"                   {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteOpportunity.subsidiary -dataType String), "}
                        "DateCreated"                  {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteOpportunity.dateCreated -dataType Date), "}
                        "LastModified"                 {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteOpportunity.LastModified -dataType Date), "}
                        "DateCreatedInSql"             {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteOpportunity.DateCreatedInSql -dataType Date), "}
                        "DateModifiedInSql"            {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteOpportunity.DateModifiedInSql -dataType Date), "}
                        "IsDirty"                      {
                            if($isDirty)                    {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $true -dataType Boolean), "}
                            elseif($isNotDirty)             {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $false -dataType Boolean), "}
                            else                            {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteOpportunity.IsDirty -dataType Boolean), "}
                                                        }
                        "SharePointDriveItemId"  {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteOpportunity.SharePointDocLibGraphListId -dataType String), "}
                        }
                    }
                $sql = $sql.TrimEnd(", ")
                $sql += " WHERE NsInternalId = $(sanitise-forSqlValue -value $preExistingRecord.NsInternalId -dataType String) "
                $sql += "AND LastModified = $(sanitise-forSqlValue -value $preExistingRecord.LastModified -dataType Date) "
                Write-Verbose "`t$sql"
                Execute-SQLQueryOnSQLDB -query $sql -queryType NonQuery -sqlServerConnection $dbConnection
                }
            }
        }
    else{Write-Error "Record with NsInsternalId [$($sqlNetSuiteOpportunity.NsExternalId)] does not exist in database. Cannot UPDATE.";break}
    }
function update-netSuiteProjectInSqlCache(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true, ParameterSetName="nsNetSuiteProject")]
            [PSCustomObject]$nsNetSuiteProject 
        ,[parameter(Mandatory = $true, ParameterSetName="sqlNetSuiteProject")]
            [PSCustomObject]$sqlNetSuiteProject 
        ,[parameter(Mandatory = $true, ParameterSetName="nsNetSuiteProject")]
            [parameter(Mandatory = $true, ParameterSetName="sqlNetSuiteProject")]
            [System.Data.Common.DbConnection]$dbConnection
        ,[parameter(Mandatory = $false, ParameterSetName="nsNetSuiteProject")]
            [parameter(Mandatory = $false, ParameterSetName="sqlNetSuiteProject")]
            [switch]$isDirty
        ,[parameter(Mandatory = $false, ParameterSetName="nsNetSuiteProject")]
            [parameter(Mandatory = $false, ParameterSetName="sqlNetSuiteProject")]
            [switch]$isNotDirty
        )
    switch ($PsCmdlet.ParameterSetName){
        'nsNetSuiteProject' {$sqlNetSuiteProject = convert-nsNetSuiteProjectToSqlNetSuiteProject -nsNetSuiteProject $nsNetSuiteProject}
        }

    Write-Verbose "update-netSuiteProjectInSqlCache [$($sqlNetSuiteProject.entityId)]"
    #Check record exists in SQL
    $sql = "SELECT TOP 1 ProjectName, NsInternalId, LastModified FROM t_PROJECTS WHERE NsInternalId = '$($sqlNetSuiteProject.NsInternalId)' ORDER BY LastModified Desc"
    $preExistingRecord = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $dbConnection
    
    if($preExistingRecord){
        if([string]::IsNullOrWhiteSpace($preExistingRecord.NsInternalId)){
            Write-Error "No NsInsternalId found on sql record [$($sqlNetSuiteProject.entityId)]. Cannot identify unique record, so cannot UPDATE."
            break
            }
        else{
            #Generate SQL statement
            $fieldsToUpdate = $sqlNetSuiteProject.PSObject.Properties | ? {$_.Value -ne $null}
            if($fieldsToUpdate){
                $sql = "UPDATE t_PROJECTS "
                $sql += "SET "
                $fieldsToUpdate | % {
                    $thisField = $_
                    switch($_.Name){
                        "NsInternalId"                 {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteProject.NsInternalId -dataType String), "}
                        "NsExternalId"                 {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteProject.NsExternalId -dataType String), "}
                        "AccountNsInternalId"          {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteProject.AccountNsInternalId -dataType String), "}
                        "ProjectName"                  {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteProject.ProjectName -dataType String), "}
                        "ProjectNumber"                {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteProject.ProjectNumber -dataType String), "}
                        "entityId"                     {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteProject.entityId -dataType String), "}
                        "entityStatus"                 {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteProject.entityStatus -dataType String), "}
                        "custentity_atlas_svcs_mm_department"  {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteProject.custentity_atlas_svcs_mm_department -dataType String), "}
                        "custentity_ant_projectsector"         {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteProject.custentity_ant_projectsector -dataType String), "}
                        "custentity_ant_projectsource"         {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteProject.custentity_ant_projectsource -dataType String), "}
                        "custentity_atlas_svcs_mm_location"    {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteProject.custentity_atlas_svcs_mm_location -dataType String), "}
                        "custentity_atlas_svcs_mm_projectmngr" {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteProject.custentity_atlas_svcs_mm_projectmngr -dataType String), "}
                        "jobType"                      {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteProject.jobType -dataType String), "}
                        "subsidiary"                   {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteProject.subsidiary -dataType String), "}
                        "DateCreated"                  {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteProject.dateCreated -dataType Date), "}
                        "LastModified"                 {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteProject.LastModified -dataType Date), "}
                        "DateCreatedInSql"             {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteProject.DateCreatedInSql -dataType Date), "}
                        "DateModifiedInSql"            {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteProject.DateModifiedInSql -dataType Date), "}
                        "IsDirty"                      {
                            if($isDirty)                    {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $true -dataType Boolean), "}
                            elseif($isNotDirty)             {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $false -dataType Boolean), "}
                            else                            {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteProject.IsDirty -dataType Boolean), "}
                                                        }
                        "SharePointDriveItemId"  {$sql += "$($thisField.Name) = $(sanitise-forSqlValue -value $sqlNetSuiteProject.SharePointDocLibGraphListId -dataType String), "}
                        }
                    }
                $sql = $sql.TrimEnd(", ")
                $sql += " WHERE NsInternalId = $(sanitise-forSqlValue -value $preExistingRecord.NsInternalId -dataType String) "
                $sql += "AND LastModified = $(sanitise-forSqlValue -value $preExistingRecord.LastModified -dataType Date) "
                Write-Verbose "`t$sql"
                Execute-SQLQueryOnSQLDB -query $sql -queryType NonQuery -sqlServerConnection $dbConnection
                }
            }
        }
    else{Write-Error "Record with NsInsternalId [$($sqlNetSuiteProject.NsExternalId)] does not exist in database. Cannot UPDATE.";break}
    }
function update-netSuiteClientInNetSuite(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [string]$netSuiteClientId
        ,[parameter(Mandatory = $true)]
            [hashtable]$fieldHash = @{}
        ,[parameter(Mandatory=$false)]
            [psobject]$netsuiteParameters
        )
    if([string]::IsNullOrWhiteSpace($netsuiteParameters)){
        $netsuiteParameters = get-netsuiteParameters -connectTo Sandbox
        Write-Warning "NetSuite environment unspecified - connecting to Sandbox"
        }

    #Validate Dropdown fields if they've not been provided in the correct format
    if($fieldHash.Keys -contains "custentity_clientrating" -and $fieldHash["custentity_clientrating"] -isnot [hashtable]){
        try{$ratingId = convert-netSuiteClientRatingToId -clientRating $fieldHash["custentity_clientrating"]}
        catch{return} #Errors reported by validation cmdlet. Just exit early.
        $fieldHash["custentity_clientrating"] = @{id=$ratingId}
        }
    if($fieldHash.Keys -contains "custentity_clienttype" -and $fieldHash["custentity_clienttype"] -isnot [hashtable]){
        try{$typeId = convert-netSuiteClientTypeToId -clientType $fieldHash["custentity_clienttype"]}
        catch{return} #Errors reported by validation cmdlet. Just exit early.
        $fieldHash["custentity_clienttype"] = @{id=$typeId}
        }
    if($fieldHash.Keys -contains "custentity_ant_clientsector" -and $fieldHash["custentity_ant_clientsector"] -isnot [hashtable]){
        try{$sectorId = convert-netSuiteSectorToId -sector $fieldHash["custentity_ant_clientsector"]}
        catch{return}
        $fieldHash["custentity_ant_clientsector"] = @{id=$sectorId}
        }
    if($fieldHash.Keys -contains "entityStatus" -and $fieldHash["entityStatus"] -isnot [hashtable]){
        try{$statusId = convert-netSuiteStatusToId -status $fieldHash["entityStatus"]}
        catch{return}
        $fieldHash["entityStatus"] = @{id=$statusId}
        }
    if($fieldHash.Keys -contains "subsidiary" -and $fieldHash["subsidiary"] -isnot [hashtable]){
        try{$subsidiaryId = convert-netSuiteSubsidiaryToId -subsidiary $fieldHash["subsidiary"]}
        catch{return}
        $fieldHash["subsidiary"] = @{id=$subsidiaryId}
        }
#    if($fieldHash.Keys -contains "customForm" -and $fieldHash["customForm"] -isnot [hashtable]){
#        try{$customFormId = convert-netSuiteCustomFormToId -formType $fieldHash["customForm"]}
#        catch{return}
#        $fieldHash["customForm"] = @{id=$customFormId}
#        }
#    else{
#        if(@("13","15") -contains $statusId){$customFormId = convert-netSuiteCustomFormToId -formType 'Anthesis | Client Form'}
#        else{$customFormId = convert-netSuiteCustomFormToId -formType 'Anthesis | Lead Form'}
#        }

    invoke-netSuiteRestMethod -requestType PATCH -url "$($netsuiteParameters.uri)/customer/$netSuiteClientId" -netsuiteParameters $netsuiteParameters -requestBodyHashTable $fieldHash

    }
function update-netSuiteClientFromHubSpotObject(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$hubSpotCompanyObject
        ,[parameter(Mandatory = $true)]
            [string]$hubSpotApiKey
        ,[parameter(Mandatory=$false)]
            [psobject]$netsuiteParameters
        )

    $fieldHash = @{
        companyName = $hubSpotCompanyObject.properties.name
        email = $hubSpotCompanyObject.properties.generic_email_address__c
        shipAddr1 = $hubSpotCompanyObject.properties.address
        shipAddr2 = $hubSpotCompanyObject.properties.address2
        shipCity = $hubSpotCompanyObject.properties.city
        shipCountry = $hubSpotCompanyObject.properties.country
        shipState = $hubSpotCompanyObject.properties.state
        shipZip = $hubSpotCompanyObject.properties.zip
        custentity_marketing_originalsourcetype  = $hubSpotCompanyObject.properties.hs_analytics_source
        custentity_marketing_numberofpageviews   = [int]$hubSpotCompanyObject.properties.hs_analytics_num_page_views
        custentity_marketing_timeoflastsession   = $hubSpotCompanyObject.properties.hs_analytics_last_timestamp
        custentity_marketing_firstconversion     = $hubSpotCompanyObject.properties.first_conversion_event_name
        custentity_marketing_mostrecentconversio = $hubSpotCompanyObject.properties.recent_conversion_event_name 
        custentity_marketing_mostrecentconvdate  = $hubSpotCompanyObject.properties.recent_conversion_date
        }

        if(![string]::IsNullOrEmpty($hubSpotCompanyObject.properties.netsuite_sector)){
            
            }
        if(![string]::IsNullOrEmpty($hubSpotCompanyObject.properties.netsuite_subsidiary)){
            
            }
        #Don't update any other fields (e.g. Status, Susbsidiary,etc. based on HubSpot data)
        #custentity_clientrating = ""
        #custentity_clienttype = ""
        #custentity_ant_clientsector = ""
        #entityStatus = ""
        #subsidiary = ""
    
    if(![string]::IsNullOrEmpty($hubSpotCompanyObject.properties.netsuiteid)){
        $netClient = get-netSuiteClientsFromNetSuite -clientId $hubSpotCompanyObject.properties.netsuiteid -netsuiteParameters $netsuiteParameters #Just double-check that we're allowed to update this record
        if($netClient.entityStatus -match "LEAD"){
            $updatedNetSuiteClient = update-netSuiteClientInNetSuite -netSuiteClientId $hubSpotCompanyObject.properties.netsuiteid -fieldHash $fieldHash -netsuiteParameters $netsuiteParameters
            $updatedNetSuiteClient = get-netSuiteClientsFromNetSuite -clientId $hubSpotCompanyObject.properties.netsuiteid -netsuiteParameters $netsuiteParameters
            $updatedHubSpotClient = update-hubSpotObject -apiKey $hubSpotApiKey -objectType companies -objectId $hubSpotCompanyObject.id -fieldHash @{lastmodifiedinnetsuite=$updatedNetSuiteClient.lastModifiedDate; lastmodifiedinhubspot=$(get-dateInIsoFormat -dateTime $(Get-Date) -precision Ticks)}
            $updatedNetSuiteClient
            }
        else{Write-Error "NetSuite company [$($netClient.companyName)][$($hubSpotCompanyObject.properties.netsuiteid)] is set to [$($netClient.entityStatus.refName)] cannot update this object using a HubSpot object"}
        }
    }
function update-netSuiteContactInNetSuite(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [string]$netSuiteContactId
        ,[parameter(Mandatory = $true)]
            [hashtable]$fieldHash = @{}
        ,[parameter(Mandatory=$false)]
            [psobject]$netsuiteParameters
        )
    if([string]::IsNullOrWhiteSpace($netsuiteParameters)){
        $netsuiteParameters = get-netsuiteParameters -connectTo Sandbox
        Write-Warning "NetSuite environment unspecified - connecting to Sandbox"
        }

    if($fieldHash.Keys -contains "subsidiary" -and $fieldHash["subsidiary"] -isnot [hashtable]){
        try{$subsidiaryId = convert-netSuiteSubsidiaryToId -subsidiary $fieldHash["subsidiary"]}
        catch{return}
        $fieldHash["subsidiary"] = @{id=$subsidiaryId}
        }
    if($fieldHash.Keys -contains "customForm" -and $fieldHash["customForm"] -isnot [hashtable]){
        try{$customFormId = convert-netSuiteCustomFormToId -formType $fieldHash["customForm"]}
        catch{return}
        $fieldHash["customForm"] = @{id=$customFormId}
        }

    invoke-netSuiteRestMethod -requestType PATCH -url "$($netsuiteParameters.uri)/contact/$netSuiteContactId" -netsuiteParameters $netsuiteParameters -requestBodyHashTable $fieldHash

    }
function update-netSuiteContactInNetSuiteFromHubSpotObject(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$hubSpotContactObject
        ,[parameter(Mandatory = $true)]
            [string]$hubSpotApiKey
        ,[parameter(Mandatory=$false)]
            [string]$companyNetSuiteId
        ,[parameter(Mandatory=$false)]
            [psobject]$netsuiteParameters
        )
    if([string]::IsNullOrWhiteSpace($netsuiteParameters)){
        $netsuiteParameters = get-netsuiteParameters -connectTo Sandbox
        Write-Warning "NetSuite environment unspecified - connecting to Sandbox"
        }

    $fullContactName = "$($hubSpotContactObject.properties.firstname) $($hubSpotContactObject.properties.lastname)".Trim(" ")
    if(![string]::IsNullOrWhiteSpace($hubSpotContactObject.properties.phone)){$mainPhone = $hubSpotContactObject.properties.phone}
    if(![string]::IsNullOrWhiteSpace($hubSpotContactObject.properties.mobilephone)){$mainPhone = $hubSpotContactObject.properties.mobilephone} #Prefer mobiles over landlines as primary phone

    $bodyHash = @{
        entityId = $fullContactName #Is NetSuite really going to let us update this? Looks like it...
        firstname = $hubSpotContactObject.properties.firstname
        lastname = $hubSpotContactObject.properties.lastname
        email = $hubSpotContactObject.properties.email
        phone = $mainPhone
        mobilePhone = $hubSpotContactObject.properties.mobilephone
        officePhone = $hubSpotContactObject.properties.phone
        title = $hubSpotContactObject.properties.jobtitle
        }

    if(![string]::IsNullOrWhiteSpace($companyNetSuiteId)){
        $bodyHash.Add("company",@{id=$companyNetSuiteId})
        }
    if(![string]::IsNullOrWhiteSpace($hubSpotContactObject.properties.first_conversion_event_name)){
        $bodyHash.Add("custentity_marketing_firstconversion",$hubSpotContactObject.properties.first_conversion_event_name)
        }
    if(![string]::IsNullOrWhiteSpace($hubSpotContactObject.properties.hs_analytics_first_referrer)){
        $bodyHash.Add("custentity_marketing_firstreferringsite",$hubSpotContactObject.properties.hs_analytics_first_referrer)
        }
    if(![string]::IsNullOrWhiteSpace($hubSpotContactObject.formSubmissionHistory)){
        $bodyHash.Add("custentity_marketing_formsubmission",$hubSpotContactObject.formSubmissionHistory)
        }
    if(![string]::IsNullOrWhiteSpace($hubSpotContactObject.properties.last_webinar_attended_date)){
        $bodyHash.Add("custentity_marketing_lastwebinardate",$(get-dateInIsoFormat -dateTime $hubSpotContactObject.properties.last_webinar_attended_date -precision Milliseconds))
        }
    if(![string]::IsNullOrWhiteSpace($hubSpotContactObject.properties.message)){
        $bodyHash.Add("custentity_marketing_message",$hubSpotContactObject.properties.message)
        }
    if(![string]::IsNullOrWhiteSpace($hubSpotContactObject.properties.recent_conversion_date)){
        $bodyHash.Add("custentity_marketing_mostrecentconvdate",$(get-dateInIsoFormat -dateTime $hubSpotContactObject.properties.recent_conversion_date -precision Milliseconds))
        }
    if(![string]::IsNullOrWhiteSpace($hubSpotContactObject.properties.recent_conversion_event_name)){
        $bodyHash.Add("custentity_marketing_mostrecentconversio",$hubSpotContactObject.properties.recent_conversion_event_name)
        }
    if(![string]::IsNullOrWhiteSpace($hubSpotContactObject.properties.hs_analytics_num_page_views)){
        $bodyHash.Add("custentity_marketing_numberofpageviews",$hubSpotContactObject.properties.hs_analytics_num_page_views)
        }
    if(![string]::IsNullOrWhiteSpace($hubSpotContactObject.properties.opted_out_of_some_marketing_emails)){
        $bodyHash.Add("custentity_marketing_optoutany",$(if($hubSpotContactObject.properties.opted_out_of_some_marketing_emails -eq $true){$true}else{$false}))
        }
    if(![string]::IsNullOrWhiteSpace($hubSpotContactObject.properties.hs_analytics_source_data_1)){
        $bodyHash.Add("custentity_marketing_originalsourcedril1",$hubSpotContactObject.properties.hs_analytics_source_data_1)
        }
    if(![string]::IsNullOrWhiteSpace($hubSpotContactObject.properties.hs_analytics_source_data_2)){
        $bodyHash.Add("custentity_marketing_originalsourcedril2",$hubSpotContactObject.properties.hs_analytics_source_data_2)  #This is stored ina a different format!!
        }
    if(![string]::IsNullOrWhiteSpace($hubSpotContactObject.properties.hs_analytics_source)){
        $bodyHash.Add("custentity_marketing_originalsourcetype",$hubSpotContactObject.properties.hs_analytics_source)
        }
    if(![string]::IsNullOrWhiteSpace($hubSpotContactObject.pageViewHistory)){
        $bodyHash.Add("custentity_marketing_pageview",$hubSpotContactObject.pageViewHistory)
        }
    if(![string]::IsNullOrWhiteSpace($hubSpotContactObject.properties.hs_analytics_last_visit_timestamp)){
        $bodyHash.Add("custentity_marketing_timeoflastsession",$(get-dateInIsoFormat -dateTime $hubSpotContactObject.properties.hs_analytics_last_visit_timestamp -precision Milliseconds))
        }
    if(![string]::IsNullOrWhiteSpace($hubSpotContactObject.properties.webinarHistory)){
        $bodyHash.Add("custentity_marketing_webinarhistory",$hubSpotContactObject.properties.webinarHistory)
        }

    
    if([string]::IsNullOrEmpty($hubSpotContactObject.properties.netsuiteid)){
        Write-Error "HubSpot Contact [$($hubSpotContactObject.properties.firstname)][$($hubSpotContactObject.properties.lastname)][$($hubSpotContactObject.id)] has no NetSuiteId. Cannot update this object using a HubSpot object"
        }
    else{
        $updatedNetSuiteContact = update-netSuiteContactInNetSuite -netSuiteContactId $hubSpotContactObject.properties.netsuiteid -fieldHash $bodyHash -netsuiteParameters $netsuiteParameters
        $updatedNetSuiteContact = get-netSuiteContactFromNetSuite -contactId $hubSpotContactObject.properties.netsuiteid -netsuiteParameters $netsuiteParameters
        $updatedHubSpotContact = update-hubSpotObject -apiKey $hubSpotApiKey -objectType contacts -objectId $hubSpotContactObject.id -fieldHash @{lastmodifiedinnetsuite=$updatedNetSuiteContact.lastModifiedDate; lastmodifiedinhubspot=$(get-dateInIsoFormat -dateTime $(Get-Date) -precision Ticks)} -Verbose:$VerbosePreference
        $updatedNetSuiteContact
        }

    }

#$clientStauses = invoke-netSuiteRestMethod -requestType GET -url "$((get-netsuiteParameters -connectTo Production).uri)/customerstatus/20" -netsuiteParameters $(get-netsuiteParameters -connectTo Production)