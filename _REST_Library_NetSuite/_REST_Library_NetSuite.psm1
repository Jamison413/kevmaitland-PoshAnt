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
            update-netSuiteOpportunityInSqlCache -nsNetsuiteProject $nsNetsuiteOpportunity -dbConnection $dbConnection -isNotDirty
            }
        else{
            Write-Verbose "`tNsInternalId [$($nsNetsuiteOpportunity.Id)] doesn't seem to have changed (probably caused by a lack of granularity in NetSuite's REST WHERE clauses). Not updating anything."
            }
        }
    else{
        if(!$alreadyPresent){Write-Verbose "`tNsInternalId [$($nsNetsuiteOpportunity.Id)] not present in SQL, adding to [t_OPPORTUNITIES]"}
        else{Write-Verbose "`tNsInternalId [$($nsNetsuiteOpportunity.Id)] Title has changed from [$($alreadyPresent.ProjectName)] to [$($nsNetsuiteOpportunity.companyName)], adding new record to [t_PROJECTS]"}
        $now = $(Get-Date)
        $sql = "INSERT INTO t_OPPORTUNITIES (NsInternalId, NsExternalId, AccountNsInternalId, OpportunityName, OpportunityNumber, entityId, entityStatus, entityNexus, custbody_project_template, tranId, status, probability, custbody_industry, subsidiary, DateCreated, LastModified, DateCreatedInSql, DateModifiedInSql, IsDirty) VALUES ("
                                             
        $sql += $(sanitise-forSqlValue -value $nsNetsuiteOpportunity.id -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteOpportunity.id -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $nsNetsuiteOpportunity.entity.id -dataType String)
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
function get-netSuiteAuthHeaders(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [ValidateSet("GET","POST")]
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
        [parameter(Mandatory = $false)]
        [ValidatePattern('^?[\w+][=][\w+]')]
        [string]$query

        ,[parameter(Mandatory=$false)]
        [psobject]$netsuiteParameters
        )

    Write-Verbose "`tget-allNetSuiteClients([$($query)])"
    if([string]::IsNullOrWhiteSpace($netsuiteParameters)){$netsuiteParameters = get-netsuiteParameters}

    $customers = invoke-netsuiteRestMethod -requestType GET -url "https://3487287-sb1.suitetalk.api.netsuite.com/rest/platform/v1/record/customer$query" -netsuiteParameters $netsuiteParameters #-Verbose 
    $customersEnumerated = [psobject[]]::new($customers.count)
    for ($i=0; $i -lt $customers.count;$i++) {
        $customersEnumerated[$i] = invoke-netsuiteRestMethod -requestType GET -url $customers.items[$i].links[0].href -netsuiteParameters $netsuiteParameters 
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
        [parameter(Mandatory = $false)]
        [ValidatePattern('^?[\w+][=][\w+]')]
        [string]$query

        ,[parameter(Mandatory=$false)]
        [psobject]$netsuiteParameters
        )

    Write-Verbose "`tget-netSuiteContactFromNetSuite([$($query)])"
    if([string]::IsNullOrWhiteSpace($netsuiteParameters)){$netsuiteParameters = get-netsuiteParameters}

    $contacts = invoke-netsuiteRestMethod -requestType GET -url "https://3487287-sb1.suitetalk.api.netsuite.com/rest/platform/v1/record/contact$query" -netsuiteParameters $netsuiteParameters #-Verbose 
    $contactsEnumerated = [psobject[]]::new($contacts.count)
    for ($i=0; $i -lt $contacts.count;$i++) {
        $contactsEnumerated[$i] = invoke-netsuiteRestMethod -requestType GET -url $contacts.items[$i].links[0].href -netsuiteParameters $netsuiteParameters 
        }
    $contactsEnumerated
    }
function get-netSuiteOpportunityFromNetSuite(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $false)]
        [ValidatePattern('^?[\w+][=][\w+]')]
        [string]$query

        ,[parameter(Mandatory=$false)]
        [psobject]$netsuiteParameters
        )

    Write-Verbose "`tget-netSuiteProjectFromNetSuite([$($query)])"
    if([string]::IsNullOrWhiteSpace($netsuiteParameters)){$netsuiteParameters = get-netsuiteParameters}

    $opportunities = invoke-netsuiteRestMethod -requestType GET -url "https://3487287-sb1.suitetalk.api.netsuite.com/rest/platform/v1/record/opportunity$query" -netsuiteParameters $netsuiteParameters #-Verbose 
    $opportunitiesEnumerated = [psobject[]]::new($opportunities.count)
    for ($i=0; $i -lt $opportunities.count;$i++) {
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
    Param()
    Write-Verbose "get-netsuiteParameters()"
    $placesToLook = @(
        "$env:USERPROFILE\Desktop\netsuite.txt"
        ,"$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\netsuite.txt"
        )
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
        [parameter(Mandatory = $false)]
        [ValidatePattern('^?[\w+][=][\w+]')]
        [string]$query

        ,[parameter(Mandatory=$false)]
        [psobject]$netsuiteParameters
        )

    Write-Verbose "`tget-netSuiteProjectFromNetSuite([$($query)])"
    if([string]::IsNullOrWhiteSpace($netsuiteParameters)){$netsuiteParameters = get-netsuiteParameters}

    $projects = invoke-netsuiteRestMethod -requestType GET -url "https://3487287-sb1.suitetalk.api.netsuite.com/rest/platform/v1/record/customer$query" -netsuiteParameters $netsuiteParameters #-Verbose 
    $projectsEnumerated = [psobject[]]::new($projects.count)
    for ($i=0; $i -lt $projects.count;$i++) {
        $projectsEnumerated[$i] = invoke-netsuiteRestMethod -requestType GET -url $projects.items[$i].links[0].href -netsuiteParameters $netsuiteParameters 
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
        [ValidateSet("GET","POST")]
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
        [ValidateSet("GET","POST")]
        [string]$requestType
        
        ,[parameter(Mandatory = $true)]
        [ValidatePattern("http")]
        [string]$url

        ,[parameter(Mandatory=$false)]
        [psobject]$netsuiteParameters
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
    #Add parameters from url
    $parameters.Split("&") | % {
        if(![string]::IsNullOrWhiteSpace($_.Split("=")[0])){
            $oAuthParamsForSigning.Add([uri]::EscapeDataString($_.Split("=")[0]),[uri]::EscapeDataString($_.Split("=")[1])) #Weirdly, these extra paramaters have to be Encoded twice...
            #write-host -f Green "$([uri]::EscapeDataString($_.Split("=")[0]),[uri]::EscapeDataString($_.Split("=")[1]))"
            }
        }
    
    $netsuiteRestHeaders = get-netsuiteAuthHeaders -requestType $requestType -url $hostUrl -oauthParameters $oAuthParamsForSigning  -oauth_consumer_secret $netsuiteParameters.oauth_consumer_secret -oauth_token_secret $netsuiteParameters.oauth_token_secret -realm $netsuiteParameters.realm
    
    Write-Verbose "Invoke-RestMethod -Uri $([uri]::EscapeUriString($url)) -Headers $(stringify-hashTable $netsuiteRestHeaders) -Method $requestType -ContentType application/swagger+json"
    $response = Invoke-RestMethod -Uri $([uri]::EscapeUriString($url)) -Headers $netsuiteRestHeaders -Method $requestType -ContentType "application/swagger+json"
    $response            
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
    $sql = "SELECT TOP 1 OpportunityName, NsInternalId, LastModified FROM t_OpportunityS WHERE NsInternalId = '$($sqlNetSuiteOpportunity.NsInternalId)' ORDER BY LastModified Desc"
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
                $sql = "UPDATE t_OpportunityS "
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
