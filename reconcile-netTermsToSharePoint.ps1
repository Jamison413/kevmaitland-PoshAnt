[cmdletbinding()]
param(
    [Parameter(Mandatory = $false, Position = 0)]
        [string]$deltaSync = $true #Specifies whether we are doing a full or incremental sync.
    )

if($PSCommandPath){
    $InformationPreference = 2
    $VerbosePreference = 0
    $logFileLocation = "C:\ScriptLogs\"
    if($deltaSync -eq $true){$suffix = "_deltaSync"}
    else{$suffix = "_fullSync"}
    #$suffix = "_fullSync"
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))$suffix`_Transcript_$(Get-Date -Format "yyyy-MM-dd").log"
    Start-Transcript $transcriptLogName -Append
    }

function new-clientDocLib(){
    [cmdletbinding()]
    param(
         [Parameter(Mandatory = $true,Position=0)]
            [psobject]$tokenResponse
        ,[Parameter(Mandatory = $true, Position = 2)]
            [Microsoft.SharePoint.Client.Taxonomy.TermSetItem]$clientTerm
        )

    $listOfClientFolders = @("_NetSuite automatically creates Opportunity & Project folders","Background","Non-specific BusDev")
    $clientSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,faed84bc-70be-4e35-bfbf-cdab31aeeb99"

    Write-Host "`t`tCreating new DocLib for Term [$($clientTerm.Name)][$($clientTerm.CustomProperties.NetSuiteId)]"
    #if($(test-validNameForSharePointFolder $clientTerm.Name) -eq $false){Write-Warning "`t`t`tTerm Name [$($clientTerm.Name)] contains illegal characters. This won't work, so I'm not going to try.";return} #DocLibs support any old rubbish because they have separate DisplayName and Name properties. Name is automatically stripped of problematic characters.

    try{ #Graph doesn't support creating Drives, so we need to create a List
        $newClientList = new-graphList -tokenResponse $tokenResponseSharePointBot -siteGraphId $clientSiteId -listType documentLibrary -listDisplayName $clientTerm.Name #$(sanitise-forSql $clientTerm.Name) #Should this really be sanitised? NO! 
        } 
    catch{
        if($_.Exception -match "409" -or $_.InnerException -match "409"){ #Already exists
            $newClientList = get-graphList -tokenResponse $tokenResponseSharePointBot -graphSiteId $clientSiteId -listName $(sanitise-forSql $clientTerm.Name) #Should this really be sanitised? YES! 
            }
        else{
            Write-Host -ForegroundColor Red "`t`t$(get-errorSummary -errorToSummarise $_)"
            return
            }
        }

    try{
        $newClientDrive = get-graphDrives -tokenResponse $tokenResponseSharePointBot -siteGraphId $clientSiteId -listGraphId $newClientList.id -ErrorAction Stop
        try{
            $newFolders = add-graphArrayOfFoldersToDrive -graphDriveId $newClientDrive.id -foldersAndSubfoldersArray $listOfClientFolders -tokenResponse $tokenResponse -conflictResolution Fail
            return $newClientDrive
            }
        catch{Write-Host -ForegroundColor Red "`t`t$(get-errorSummary -errorToSummarise $_)"}
        }
    catch{#Failed to retrieve Drive
        Write-Host -ForegroundColor Red "`t`t$(get-errorSummary -errorToSummarise $_)"
        }

    
    }
function new-oppProjFolders(){
    [cmdletbinding()]
    param(
         [Parameter(Mandatory = $true,Position=0)]
            [psobject]$tokenResponse
        ,[Parameter(Mandatory = $true, Position = 1)]
            [Microsoft.SharePoint.Client.Taxonomy.TermSetItem]$oppProjTermWithClientInfo
        )

    $listOfLeadProjSubFolders = @("Admin & contracts", "Analysis","Data & refs","Meetings","Proposal","Reports","Summary (marketing) - end of project")

    if(![string]::IsNullOrWhiteSpace($oppProjTermWithClientInfo.UniversalOppName) -and ![string]::IsNullOrWhiteSpace($oppProjTermWithClientInfo.NetSuiteProjectId)){
        Write-Warning "Opportunity [$($oppProjTermWithClientInfo.UniversalOppName)][$($oppProjTermWithClientInfo.NetSuiteOppId)] for [$($oppProjTermWithClientInfo.UniversalClientName)][$($oppProjTermWithClientInfo.DriveClientId)] has already been converted to a Project. Not recreating the Opp Folders."
        return
        }
    
    if($(test-validNameForSharePointFolder $oppProjTermWithClientInfo.Name) -eq $false){Write-Warning "Term Name [$($oppProjTermWithClientInfo.Name)] contains illegal characters. This won't work, so I'm not going to try.";return}

    Write-Host "`tCreating new Folders for [$($oppProjTermWithClientInfo.Name)] in [$($oppProjTermWithClientInfo.UniversalClientName)][$($oppProjTermWithClientInfo.NetSuiteClientId)][$($oppProjTermWithClientInfo.DriveClientId)]"
    [array]$customisedFolderList = $oppProjTermWithClientInfo.Name
    $customisedFolderList += $listOfLeadProjSubFolders | % {"$($oppProjTermWithClientInfo.Name)\$_"}
    try{
        $newFolders = add-graphArrayOfFoldersToDrive -graphDriveId $oppProjTermWithClientInfo.DriveClientId -foldersAndSubfoldersArray $customisedFolderList -tokenResponse $tokenResponseSharePointBot -conflictResolution Fail -ErrorAction Continue
        #else{Write-Host "`t`tadd-graphArrayOfFoldersToDrive didn't return the new Folders, but didn't produce an error either :/"}
        }
    catch{
        Write-Host -ForegroundColor Red "`t`t$(get-errorSummary -errorToSummarise $_)"
        #Write-Error "Error creating $($thisIsA)Folders for [$($oppProjTermWithClientInfo.Name)][$($oppProjTermWithClientInfo.id)] for Client [$($clientTerm.Name)][$($clientTerm.CustomProperties.GraphDriveId)] | Retrying with Verbose"
        #add-graphArrayOfFoldersToDrive -graphDriveId $clientTerm.CustomProperties.GraphDriveId -foldersAndSubfoldersArray $customisedFolderList -tokenResponse $tokenResponseSharePointBot -conflictResolution Fail -Verbose
        }
    $newFolders   
    }
function process-comparison(){
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
            [AllowNull()]
            [array]$subsetOfNetObjects 
        ,[Parameter(Mandatory = $true, Position = 1)]
            [AllowNull()]
            [array]$allTermObjects 
        ,[Parameter(Mandatory = $true, Position = 2)]
            [string]$idInCommon 
        ,[Parameter(Mandatory = $true, Position = 3)]
            [string]$propertyToTest
        ,[Parameter(Mandatory = $false, Position = 4)]
            [switch]$validate
        )

    #compare-object jiggery-pokery documented with pictures on IT Site: https://anthesisllc.sharepoint.com/sites/Resources-IT/_layouts/15/DocIdRedir.aspx?ID=HXX7CE52TSD2-1759992947-266
    #Prerequisite: $subsetOfNetObjects, $allTermObjects
    [array]$correspondingSubsetOfTermObjects = Compare-Object -ReferenceObject @($allTermObjects | Select-Object) -DifferenceObject @($subsetOfNetObjects | Select-Object) -Property $idInCommon -PassThru -IncludeEqual -ExcludeDifferent
    [array]$comparisonOfPropertyToTest = Compare-Object -ReferenceObject @($subsetOfNetObjects | Select-Object) -DifferenceObject @($correspondingSubsetOfTermObjects | Select-Object) -Property $idInCommon,$propertyToTest -PassThru -IncludeEqual
    [array]$netObjectsWithMismatchedProperty = $comparisonOfPropertyToTest | ? {$_.SideIndicator -eq "<="} | Sort-Object $idInCommon
    [array]$correspondingTermObjectsWithMismatchedProperty = $comparisonOfPropertyToTest | ? {$_.SideIndicator -eq "=>"} | Sort-Object $idInCommon
    [array]$netObjectsWithMatchingProperty = $comparisonOfPropertyToTest | ? {$_.SideIndicator -eq "=="} 
    
    Write-Verbose "subsetOfNetObjects.Count = `t`t`t`t[$($subsetOfNetObjects.Count)]";Write-Verbose "correspondingSubsetOfTermObjects.Count = `t[$($correspondingSubsetOfTermObjects.Count)] (should be equal)";Write-Verbose "comparisonOfPropertyToTest.Count = `t`t[$($comparisonOfPropertyToTest.Count)] (<=[$(($netObjectsWithMismatchedProperty).Count)]  ==[$($netObjectsWithMatchingProperty.Count)]  =>[$(($correspondingTermObjectsWithMismatchedProperty).Count)]) (<= should equal =>)"
    if($validate){
        if($netObjectsWithMismatchedProperty.Count -ne $correspondingTermObjectsWithMismatchedProperty.Count){
            Write-Verbose "`"<=`" array Count [$($netObjectsWithMismatchedProperty.Count)] does not equal `"=>`" array Count [$($correspondingTermObjectsWithMismatchedProperty.Count)]: Invalid output"
            $invalid = $true
            }
        for($i=0; $i -lt $correspondingTermObjectsWithMismatchedProperty.Count; $i++){
            if($correspondingTermObjectsWithMismatchedProperty[$i]."$idInCommon" -ne $netObjectsWithMismatchedProperty[$i]."$idInCommon"){
                Write-Verbose "Property [$propertyToTest] for array `"<=`" item [$i] [$($netObjectsWithMismatchedProperty[$i]."$idInCommon")] does not equal `"=>`" array item [$($correspondingTermObjectsWithMismatchedProperty[$i]."$idInCommon")]: Invalid output"
                $invalid = $true
                }
            }
        if($invalid -eq $true){return $false} #Return $false instead of the invalid comparison data
        }
    @{"<=" = $netObjectsWithMismatchedProperty
        "==" = $netObjectsWithMatchingProperty
        "=>" = $correspondingTermObjectsWithMismatchedProperty
        }
    }
function process-docLibs(){
    [cmdletbinding()]
    param(
         [Parameter(Mandatory = $true,Position=0)]
            [psobject]$tokenResponse
        ,[Parameter(Mandatory = $true, Position = 1)]
            [PSCustomObject]$standardisedSourceDocLib
        #,[Parameter(Mandatory = $true, Position = 2,ParameterSetName="merge")]
        #    [PSCustomObject]$mergeInto 
        ,[Parameter(Mandatory = $true, Position = 2,ParameterSetName="rename")]
            [string]$renameAs 
        ,[Parameter(Mandatory = $true, Position = 2,ParameterSetName="delete")]
            [switch]$confirmDeleteEmptyDocLib
        )
    $clientSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,faed84bc-70be-4e35-bfbf-cdab31aeeb99"

    switch ($PsCmdlet.ParameterSetName){
        'delete' {
            if($standardisedSourceDocLib.quota.used -ne 0){
                Write-Error -Message "[$($standardisedSourceDocLib.DriveClientName)][$($standardisedSourceDocLib.DriveClientId)] contains data. NOT removing DocLib containing data." -TargetObject $standardisedSourceDocLib
                return $standardisedSourceDocLib
                }
            elseif($standardisedSourceDocLib.quota.used -eq 0){
                try{
                    Write-Host "`t`t`t`tDeleting empty DocLib [$($standardisedSourceDocLib.DriveClientName)][$($standardisedSourceDocLib.DriveClientId)] using PNP"
                    #Deleting via Graph works, but it bypasses the Recycle Bin (which ios too dangerous). Using PNP instead until the Graph API supports the -Recycle function
                    $list = get-graphList -tokenResponse $tokenResponse -graphDriveId $($standardisedSourceDocLib.DriveClientId) -ErrorAction Stop
                    #invoke-graphDelete -tokenResponse $tokenResponse -graphQuery "/sites/$($list.parentReference.siteId)/lists/$($list.id)" -ErrorAction Stop
                    if(![string]::IsNullOrWhiteSpace($list.id)){
                        Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/clients" -Credentials $adminCreds -ErrorAction Stop
                        Remove-PnPList -Identity $list.id -Recycle -Force 
                        Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com" -Credentials $adminCreds
                        return $true
                        }
                    else{return $false}
                    }
                catch{
                    Write-Host -ForegroundColor Red "`t`t`t`t$(get-errorSummary $_)"
                    return $false
                    }
                }
            }
        'rename' {
            #if($(test-validNameForSharePointFolder $renameAs) -eq $false){Write-Warning "New Name [$renameAs] for Drive [$($standardisedSourceDocLib.DriveClientName)][$($standardisedSourceDocLib.DriveClientId)] contains illegal characters. This won't work, so I'm not going to try.";return} #DocLibs support any old rubbish because they have separate DisplayName and Name properties. Name is automatically stripped of problematic characters.
            Write-Host "`t`t`t`tRenaming DocLib [$($standardisedSourceDocLib.DriveClientName)][$($standardisedSourceDocLib.DriveClientId)] to [$renameAs]"
            try{
                $correspondingList = get-graphList -tokenResponse $tokenResponse -graphDriveId $($standardisedSourceDocLib.DriveClientId) -ErrorAction Stop
                try{
                    [array]$result = set-graphList -tokenResponse $tokenResponseSharePointBot -graphSiteId $clientSiteId -graphListId $correspondingList.id -listPropertyHash @{displayName=$renameAs} -ErrorAction Stop
                    }
                catch{Write-Host -ForegroundColor Red "`t`t`t`t$(get-errorSummary $_)"}
                }
            catch{
                Write-Host -ForegroundColor Red "`t`t`t`t$(get-errorSummary $_)"
                }
            if($result[0].name -eq $renameAs -or $result[0].name -eq $(sanitise-forPnpSharePoint $renameAs)){return $true}
            else{return $false}
            }
        }    
    }
function process-folders(){
    [cmdletbinding()]
    param(
         [Parameter(Mandatory = $true,Position=0)]
            [psobject]$tokenResponse
        ,[Parameter(Mandatory = $true, Position = 1)]
            [PSCustomObject]$standardisedSourceFolder 
        ,[Parameter(Mandatory = $true, Position = 2,ParameterSetName="merge")]
            [PSCustomObject]$mergeInto 
        ,[Parameter(Mandatory = $true, Position = 2,ParameterSetName="rename")]
            [string]$renameAs 
        ,[Parameter(Mandatory = $true, Position = 2,ParameterSetName="delete")]
            [switch]$confirmDeleteEmptyFolders 
        )
    switch ($PsCmdlet.ParameterSetName){
        "delete" {
            if($standardisedSourceFolder.DriveItemSize -ne 0){
                Write-Warning -Message "[$($standardisedSourceFolder.DriveClientName)][$($standardisedSourceFolder.DriveClientId)][$($standardisedSourceFolder.DriveItemId)] contains data - NOT removing folder containing data."
                return $standardisedSourceFolder
                }
            elseif($standardisedSourceFolder.DriveItemSize -eq 0){
                try{
                    Write-Host "`t`t`t`tDeleting empty folder [$($standardisedSourceFolder.DriveItemName)][$($standardisedSourceFolder.DriveItemId)][$($standardisedSourceFolder.DriveClientName)][$($standardisedSourceFolder.DriveClientId)]"
                    invoke-graphDelete -tokenResponse $tokenResponse -graphQuery "/drives/$($standardisedSourceFolder.DriveClientId)/items/$($standardisedSourceFolder.DriveItemId)" -ErrorAction Stop
                    return $true
                    }
                catch{
                    Write-Host -ForegroundColor Red "`t`t`t`t$(get-errorSummary $_)"
                    return $false
                    }
                }
            }
        "merge" {
            #Move DriveItem
            [array]$childDriveItems = get-graphDriveItems -tokenResponse $tokenResponse -driveGraphId $standardisedSourceFolder.DriveClientId -itemGraphId $standardisedSourceFolder.DriveItemId -returnWhat Children
            if($([bool]$mergeInto.PSobject.Properties.name -match "webUrl") -eq $false -and ![string]::IsNullOrWhiteSpace($mergeInto.DriveItemUrl)){$mergeInto | Add-Member -MemberType NoteProperty -Name webUrl -Value $mergeInto.DriveItemUrl -Force}#If we've been passed a Standardised destination folder, re-set the raw properties so this function works for both raw DriveItems *and* standardised folders
            if($([bool]$mergeInto.PSobject.Properties.name -match "id") -eq $false -and ![string]::IsNullOrWhiteSpace($mergeInto.DriveItemId)){$mergeInto | Add-Member -MemberType NoteProperty -Name id -Value $mergeInto.DriveItemId -Force}#If we've been passed a Standardised destination folder, re-set the raw properties so this function works for both raw DriveItems *and* standardised folders
            if($([bool]$mergeInto.PSobject.Properties.name -match "parentReference.driveId") -eq $false -and ![string]::IsNullOrWhiteSpace($mergeInto.DriveClientId)){$mergeInto | Add-Member -MemberType NoteProperty -Name "parentReference" -Value @(New-Object psobject -ArgumentList @{driveId=$mergeInto.DriveClientId}) -Force}#If we've been passed a Standardised destination folder, re-set the raw properties so this function works for both raw DriveItems *and* standardised folders
            if($([bool]$mergeInto.PSobject.Properties.name -match "name") -eq $false -and ![string]::IsNullOrWhiteSpace($mergeInto.DriveItemName)){$mergeInto | Add-Member -MemberType NoteProperty -Name "name" -Value $mergeInto.DriveItemName -Force}#If we've been passed a Standardised destination folder, re-set the raw properties so this function works for both raw DriveItems *and* standardised folders
            @($childDriveItems | Select-Object) | % {
                $thisChildDriveItem = $_
                try{
                    Write-Host "`t`t`t`tMoving [$($standardisedSourceFolder.DriveItemName+" | "+$thisChildDriveItem.name)][$($thisChildDriveItem.webUrl)] to [$($mergeInto.webUrl)]"
                    [array]$movedDriveItems += move-graphDriveItem -tokenResponse $tokenResponse -driveGraphIdSource $thisChildDriveItem.parentReference.driveId -itemGraphIdSource $thisChildDriveItem.id -driveGraphIdDestination $mergeInto.parentReference.driveId -parentItemGraphIdDestination $mergeInto.id
                    }
                catch{
                    if($_.Exception -match "(409)"){
                        #If there is a conflict, try moving the childitems. If all childitems are moved, delete this item.
                        Write-Warning "`tFolder [$($thisChildDriveItem.name)] exists in destination - attempting to move subfolders instead"
                        if($([bool]$thisChildDriveItem.PSobject.Properties.name -match "DriveClientId") -eq $false){$thisChildDriveItem | Add-Member -MemberType NoteProperty -Name DriveClientId -Value $thisChildDriveItem.parentReference.driveId -Force}#As we now have a raw source folder, set the standardised properties so we can resubmit it to this function
                        if($([bool]$mergeInto.PSobject.Properties.name -match "parentReference.driveId") -eq $false -and ![string]::IsNullOrWhiteSpace($mergeInto.DriveClientId)){$mergeInto | Add-Member -MemberType NoteProperty -Name "parentReference" -Value @(New-Object psobject -ArgumentList @{driveId=$mergeInto.DriveClientId}) -Force}#If we've been passed a Standardised destination folder, re-set the raw properties so this function works for both raw DriveItems *and* standardised folders
                        if($([bool]$thisChildDriveItem.PSobject.Properties.name -match "DriveItemId") -eq $false){$thisChildDriveItem | Add-Member -MemberType NoteProperty -Name DriveItemId -Value $thisChildDriveItem.id -Force}#As we now have a raw source folder, set the standardised properties so we can resubmit it to this function
                        $mergeIntoSubfolder = get-graphDriveItems -tokenResponse $tokenResponse -driveGraphId $mergeInto.parentReference.driveId -itemGraphId $mergeInto.id -returnWhat Children -filterNameRegex $([regex]::Escape($thisChildDriveItem.name))
                        process-folders -tokenResponse $tokenResponse -standardisedSourceFolder $thisChildDriveItem -mergeInto $mergeIntoSubfolder -Verbose:$VerbosePreference
                        $refreshedThisChildDriveItem = get-graphDriveItems -tokenResponse $tokenResponse -driveGraphId $thisChildDriveItem.parentReference.driveId -itemGraphId $thisChildDriveItem.id -returnWhat Item
                        if($([bool]$refreshedThisChildDriveItem.PSobject.Properties.name -match "DriveClientId")   -eq $false){$refreshedThisChildDriveItem | Add-Member -MemberType NoteProperty -Name DriveClientId   -Value $refreshedThisChildDriveItem.parentReference.driveId -Force}#As we now have a raw source folder, set the standardised properties so we can resubmit it to this function
                        if($([bool]$refreshedThisChildDriveItem.PSobject.Properties.name -match "DriveClientName") -eq $false){$refreshedThisChildDriveItem | Add-Member -MemberType NoteProperty -Name DriveClientName -Value $standardisedSourceFolder.DriveClientName -Force}#As we now have a raw source folder, set the standardised properties so we can resubmit it to this function
                        if($([bool]$refreshedThisChildDriveItem.PSobject.Properties.name -match "DriveItemId")     -eq $false){$refreshedThisChildDriveItem | Add-Member -MemberType NoteProperty -Name DriveItemId     -Value $refreshedThisChildDriveItem.id -Force}#As we now have a raw source folder, set the standardised properties so we can resubmit it to this function
                        if($([bool]$refreshedThisChildDriveItem.PSobject.Properties.name -match "DriveItemName")   -eq $false){$refreshedThisChildDriveItem | Add-Member -MemberType NoteProperty -Name DriveItemName   -Value $refreshedThisChildDriveItem.name -Force}#As we now have a raw source folder, set the standardised properties so we can resubmit it to this function
                        if($([bool]$refreshedThisChildDriveItem.PSobject.Properties.name -match "DriveItemSize")   -eq $false){$refreshedThisChildDriveItem | Add-Member -MemberType NoteProperty -Name DriveItemSize   -Value $refreshedThisChildDriveItem.size -Force}#As we now have a raw source folder, set the standardised properties so we can resubmit it to this function
                        process-folders -tokenResponse $tokenResponse -standardisedSourceFolder $refreshedThisChildDriveItem -confirmDeleteEmptyFolders
                        }
                    else{
                        Write-Host -ForegroundColor Red "`t`t`t`t$(get-errorSummary $_)"
                        [array]$failedDriveItems += $thisChildDriveItem
                        }
                    }
                }
            if($failedDriveItems.Count -gt 0){return $failedDriveItems}
            else{return $movedDriveItems}
            }
        "rename" {
            if($(test-validNameForSharePointFolder $renameAs) -eq $false){Write-Warning "New Name [$renameAs] for DriveItem [$($standardisedSourceFolder.DriveItemName)][$($standardisedSourceFolder.DriveItemId)][$($standardisedSourceFolder.DriveClientName)][$($standardisedSourceFolder.DriveClientId)] contains illegal characters. This won't work, so I'm not going to try.";return}
            Write-Host "`t`t`t`tRenaming folder [$($standardisedSourceFolder.DriveItemName)][$($standardisedSourceFolder.DriveItemId)][$($standardisedSourceFolder.DriveClientName)][$($standardisedSourceFolder.DriveClientId)] to [$renameAs]"
            try{
                [array]$result = set-graphDriveItem -tokenResponse $tokenResponse -driveId $standardisedSourceFolder.DriveClientId -driveItemId $standardisedSourceFolder.DriveItemId -driveItemPropertyHash @{name=$renameAs} -ErrorAction Stop
                }
            catch{
                if($_.Exception -match "409"){#If the target already exists, if our original is empty then delete it
                    if($standardisedSourceFolder.DriveItemSize -eq 0){process-folders -tokenResponse $tokenResponse -standardisedSourceFolder $standardisedSourceFolder -confirmDeleteEmptyFolders}
                    }
                else{Write-Host -ForegroundColor Red "`t`t`t`t$(get-errorSummary $_)"}
                }
            if($result[0].name -eq $renameAs -or $result[0].name -eq $(sanitise-forPnpSharePoint $renameAs)){return $true}
            else{return $false}
            }
        }
    
    }
function set-standardisedClientDriveProperties(){
    [cmdletbinding()]
    param(
         [Parameter(Mandatory = $true,Position=0,ParameterSetName="clientDrive")]
            [psobject]$rawClientDrive
         ,[Parameter(Mandatory = $true,Position=0,ParameterSetName="opp")]
            [psobject]$rawOppOrProjTerm
         ,[Parameter(Mandatory = $true,Position=1,ParameterSetName="opp")]
            [psobject]$allClientTerms
        )
    switch ($PsCmdlet.ParameterSetName){
        "clientDrive" {
            $rawClientDrive | % {
                Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalClientName -Value $_.name -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalClientNameSanitised -Value $(sanitise-forNetsuiteIntegration $_.name) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name DriveClientId -Value $_.id -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name DriveClientName -Value $_.name -Force
                }
            $rawClientDrive
            }
        
        "opp" {
            $correspondingClient = Compare-Object -ReferenceObject $allClientTerms -DifferenceObject $rawOppOrProjTerm -Property NetSuiteClientId -IncludeEqual -ExcludeDifferent -PassThru
            $rawOppOrProjTerm | % {
                Add-Member -InputObject $_ -MemberType NoteProperty -Name "DriveClientId" -Value $correspondingClient.DriveClientId -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name "UniversalClientName" -Value $correspondingClient.UniversalClientName -Force
                }
            $rawOppOrProjTerm
            }

        }

    }
function test-validNameForSharePointFolder(){
    [cmdletbinding()]
    param(
         [Parameter(Mandatory = $true,Position=0)]
            [string]$stringToTest
        )
    if($stringToTest -eq $(sanitise-forSharePointFolderName $stringToTest)){$true}
    else{$false}
    }
$timeForFullCycle = Measure-Command {
cls
    
    #region GetData
    #region getDriveData
    $appCredsSharePointBot = $(get-graphAppClientCredentials -appName SharePointBot)
    $tokenResponseSharePointBot = get-graphTokenResponse -aadAppCreds $appCredsSharePointBot
    if($deltaSync -eq $false){
        $driveClientRetrieval = Measure-Command {
            $clientSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,faed84bc-70be-4e35-bfbf-cdab31aeeb99"
            $allClientDrives = get-graphDrives -tokenResponse $tokenResponseSharePointBot -siteGraphId $clientSiteId
            $allClientDrives = $allClientDrives | % {set-standardisedClientDriveProperties -rawClientDrive $_}
            }
        Write-Host "[$($allClientDrives.Count)] Client Drives retrieved from SharePoint in [$($driveClientRetrieval.TotalSeconds)] seconds ([$($driveClientRetrieval.totalMinutes)] minutes)"

        $now = $(Get-Date -f FileDateTimeUniversal) #USed to create a temp file to speed up the enumeration
        $topLevelFolderRetrieval = Measure-Command {
            for($i=0; $i-lt $allClientDrives.Count; $i++){
                write-progress -activity "Enumerating Drives contents" -Status "[$i/$($allClientDrives.count)]" -PercentComplete ($i/ $allClientDrives.count *100)
                $thisClientDrive = $allClientDrives[$i]
                $tokenResponseSharePointBot = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponseSharePointBot -renewTokenExpiringInSeconds 60 -aadAppCreds $appCredsSharePointBot
                try{
                    $theseTopLevelFolders = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $thisClientDrive.DriveClientId -returnWhat Children
                    }
                catch{
                    write-warning "`tCould not retrieve DriveItems for Client [$($thisClientDrive.NetSuiteClientName)][$($thisClientDrive.NetSuiteClientId)]"
                    Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                    }
                #$thisCombinedClient = $combinedClients | ? {$_.NetSuiteClientId -eq }
                @($theseTopLevelFolders | Select-Object) | % {
                    $folderObject = New-Object PSObject -Property ([ordered]@{
                        #NetSuiteClientId = $thisClientDrive.NetSuiteClientId
                        #NetSuiteClientName = $thisClientDrive.NetSuiteClientName
                        #TermClientId = $thisClientDrive.TermClientId
                        #TermClientName = $thisClientDrive.TermClientName
                        DriveClientId = $thisClientDrive.DriveClientId
                        DriveClientName = $thisClientDrive.DriveClientName
                        DriveClientUrl = $thisClientDrive.DriveClientUrl
                        DriveItemName = $_.name
                        DriveItemId = $_.Id
                        DriveItemUrl = $_.weburl
                        DriveItemCreatedDateTime = $_.createdDateTime
                        DriveItemLastModifiedDateTime = $_.lastModifiedDateTime
                        DriveItemSize = $_.size
                        DriveItemChildCountForFolders = $_.folder.childCount
                        DriveItemFirstWord = $null
                        })
                    $folderObject.DriveItemFirstWord = ([uri]::UnescapeDataString($(Split-Path $folderObject.DriveItemUrl -Leaf)) -split " ")[0]
                    if($folderObject.DriveItemFirstWord -match "^O-"){$folderObject | add-member -MemberType NoteProperty -Name UniversalOppName -Value $($_.name) -Force}
                    elseif($folderObject.DriveItemFirstWord -match "^P-"){$folderObject | add-member -MemberType NoteProperty -Name UniversalProjName -Value $($_.name) -Force}
                    $folderObject | Export-Csv -Path "$env:TEMP\NetRec_AllFolders_$now.csv" -Append -NoTypeInformation -Encoding UTF8 -Force #There are going to be a _lot_ of these, but the number is unknown. Rather than += an array (which will get very inefficient at large numbers), append the data to a CSV and import the CSV once the enumeration is complete
                    }
                }
            $topLevelFolders = import-csv "$env:TEMP\NetRec_AllFolders_$now.csv"
            }
        Write-Host "[$($topLevelFolders.count)] ClientDrive top-level folders enumerated in [$($topLevelFolderRetrieval.TotalMinutes)] minutes ([$($allClientDrives.count / $topLevelFolderRetrieval.TotalMinutes)] per minute)"

        $driveItemsOppFolders = $topLevelFolders | ? {$_.DriveItemFirstWord -match '^O-'} 
        $driveItemsOppFolders | % {
            Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalOppName -Value $_.DriveItemName -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalOppCode -Value $($_.DriveItemName -split " ")[0] -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalOppNameSanitised -Value $(sanitise-forNetsuiteIntegration $_.DriveItemName) -Force
            }

        $driveItemsProjFolders = $topLevelFolders | ? {$_.DriveItemFirstWord -match '^P-'} 
        $driveItemsProjFolders | % {
            Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalProjName -Value $_.DriveItemName -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalProjCode -Value $($_.DriveItemName -split " ")[0] -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalProjNameSanitised -Value $(sanitise-forNetsuiteIntegration $_.DriveItemName) -Force
            }

        }
    #endregion

    #region getTermData
        $sharePointAdmin = "kimblebot@anthesisgroup.com"
        #convertTo-localisedSecureString "KimbleBotPasswordHere"
        try{$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Downloads\KimbleBot.txt)}
        catch{
            if($_.Exception -match "Key not valid for use in specified state"){
                Write-Error "[$env:USERPROFILE\Downloads\KimbleBot.txt] Key not valid for use in specified state."
                exit
                }
            else{get-errorSummary -errorToSummarise $_}
            }
        $adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
        Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com" -Credentials $adminCreds


        #region getProjectData
    $termProjRetrieval = Measure-Command {
        $pnpTermGroup = "Kimble"
        $pnpTermSet = "Projects"
        $allProjTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes TermSet,TermSet.Group,TermStore,CustomProperties | ? {$_.IsDeprecated -eq $false}
        $allProjTerms | % {
            Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteProjectId -Value $($_.CustomProperties.NetSuiteProjId) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteClientId -Value $($_.CustomProperties.NetSuiteClientId) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteProjLastModifiedDate -Value $($_.CustomProperties.NetSuiteProjLastModifiedDate) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name TermProjName -Value $($_.name) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name TermProjCode -Value $(($_.name -split " ")[0]) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name TermProjId -Value $($_.Id) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name DriveItemId -Value $($_.CustomProperties.DriveItemId) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalProjCode -Value $(($_.name -split " ")[0]) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalProjName -Value $($_.name) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalProjNameSanitised -Value $(sanitise-forNetsuiteIntegration $_.name) -Force
            }
        }
    Write-Host "[$($allProjTerms.Count)] Projects retrieved from TermStore in [$($termProjRetrieval.TotalSeconds)] seconds"

    [array]$allProjTermsWithoutOvertlyDuffNames = $allProjTerms | ? {$(test-validNameForSharePointFolder -stringToTest $_.UniversalProjName) -eq $true}
    if($allProjTerms.Count -ne $allProjTermsWithoutOvertlyDuffNames.count){
        Write-Host "`t[$($allProjTerms.Count -$allProjTermsWithoutOvertlyDuffNames.Count)] of these contain illegal characters for Folders, so I'll just process the remaining [$($allProjTermsWithoutOvertlyDuffNames.Count)]"
        $allProjTerms = $allProjTermsWithoutOvertlyDuffNames
        }




        #endregion

        #region getOppData
    $termOppRetrieval = Measure-Command {
        $pnpTermGroup = "Kimble"
        $pnpTermSet = "Opportunities"
        $allOppTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes TermSet,TermSet.Group,TermStore,CustomProperties | ? {$_.IsDeprecated -eq $false}
        $allOppTerms | % {
            Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteOppId -Value $($_.CustomProperties.NetSuiteOppId) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteClientId -Value $($_.CustomProperties.NetSuiteClientId) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteOppLastModifiedDate -Value $($_.CustomProperties.NetSuiteOppLastModifiedDate) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name TermOppLabel -Value $($_.name) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalOppCode -Value $(($_.name -split " ")[0]) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name TermProjId -Value $($_.CustomProperties.NetSuiteProjectId) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name DriveItemId -Value $($_.CustomProperties.DriveItemId) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalOppName -Value $($_.name) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalOppNameSanitised -Value $(sanitise-forNetsuiteIntegration $_.name) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteProjectId -Value $($_.CustomProperties.NetSuiteProjectId) -Force
            }
        }
    Write-Host "[$($allOppTerms.Count)] Opportunities retrieved from TermStore in [$($termOppRetrieval.TotalSeconds)] seconds"

    [array]$allOppTermsWithoutOvertlyDuffNames = $allOppTerms | ? {$(test-validNameForSharePointFolder -stringToTest $_.UniversalOppName) -eq $true}
    if($allOppTerms.Count -ne $allOppTermsWithoutOvertlyDuffNames.count){
        Write-Host "`t[$($allOppTerms.Count -$allOppTermsWithoutOvertlyDuffNames.Count)] of these contain illegal characters for Folders, so I'll just process the remaining [$($allOppTermsWithoutOvertlyDuffNames.Count)]"
        $allOppTerms = $allOppTermsWithoutOvertlyDuffNames
        }

        #endregion

        #region getClientData
    $termClientRetrieval = Measure-Command {
        $pnpTermGroup = "Kimble"
        $pnpTermSet = "Clients"
        $allClientTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes TermSet,TermSet.Group,TermStore,CustomProperties | ? {$_.IsDeprecated -eq $false}
        @($allClientTerms | Select-Object) | % {
            Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteClientId -Value $($_.CustomProperties.NetSuiteId) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name DriveClientId -Value $($_.CustomProperties.GraphDriveId) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name TermClientId -Value $($_.Id) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name TermClientName -Value $($_.name) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteLastModifiedDate -Value $($_.CustomProperties.NetSuiteLastModifiedDate) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalClientName -Value $($_.Name) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalClientNameSanitised -Value $(sanitise-forNetsuiteIntegration $_.Name) -Force #This helps to avoid weird encoding, diacritic and special character problems when comparing strings
            }
        }
    Write-Host "[$($allClientTerms.Count)] Clients Terms retrieved from TermStore in [$($termClientRetrieval.TotalSeconds)] seconds"

    #DocLibs accept any old rubbish because they have both a Name and DisplayName properties. DisplayName allows any characters, and Name is generated automatically by stripping out illegal characters.
    <#[array]$allClientTermsWithoutOvertlyDuffNames = $allClientTerms | ? {$(test-validNameForSharePointFolder -stringToTest $_.UniversalClientName) -eq $true} 
    if($allClientTerms.Count -ne $allClientTermsWithoutOvertlyDuffNames.count){
        Write-Host "`t[$($allClientTerms.Count -$allClientTermsWithoutOvertlyDuffNames.Count)] of these contain illegal characters for Folders, so I'll just process the remaining [$($allClientTermsWithoutOvertlyDuffNames.Count)]"
        $allClientTerms = $allClientTermsWithoutOvertlyDuffNames
        }#>

        #endregion

    #endregion
    #endregion

    #region Prepare data
    #region Prepare Clients datasets

    if($deltaSync -eq $true){
        [array]$newClients = $allClientTerms | ? {[string]::IsNullOrEmpty($_.DriveClientId)}
        [array]$existingClients = $allClientTerms | ? {![string]::IsNullOrEmpty($_.DriveClientId) -and $_.CustomProperties.flagForReprocessing -ne $false}
        }

    if($deltaSync -eq $false){
        $clientComparison = Compare-Object -ReferenceObject @($allClientTerms | Select-Object) -DifferenceObject @($allClientDrives | Select-Object) -Property "DriveClientId" -IncludeEqual -PassThru
        [array]$newClients = $clientComparison | ? {$_.SideIndicator -eq "<=" -and [string]::IsNullOrWhiteSpace($_.NetSuiteProjectId)} #Exclude any Opps already converted to a Project
        [array]$existingClients = $clientComparison | ? {$_.SideIndicator -eq "=="}
        #[array]$orphanedClientDocLibs = $clientComparison | ? {$_.SideIndicator -eq "=>"}
        }

        #endregion
    #region Prepare Opps datasets
    $matchingOppsToClients = Measure-Command {
        if($deltaSync -eq $true){
            [array]$newOpps = $allOppTerms | ? {[string]::IsNullOrEmpty($_.DriveItemId)}
            [array]$existingOpps = $allOppTerms | ? {![string]::IsNullOrEmpty($_.DriveItemId) -and $_.CustomProperties.flagForReprocessing -ne $false}
            }

        if($deltaSync -eq $false){
            $oppComparison = Compare-Object -ReferenceObject @($allOppTerms | Select-Object) -DifferenceObject @($driveItemsOppFolders | Select-Object) -Property "DriveItemId" -IncludeEqual -PassThru
            [array]$newOpps = $oppComparison | ? {$_.SideIndicator -eq "<=" -and [string]::IsNullOrWhiteSpace($_.NetSuiteProjectId)} #Exclude any Opps already converted to a Project
            [array]$existingOpps = $oppComparison | ? {$_.SideIndicator -eq "=="}
            #[array]$orphanedOppFolders = $oppComparison | ? {$_.SideIndicator -eq "=>"}
            }

        $orphanedOppFolders = @($orphanedOppFolders | Select-Object) | % {set-standardisedClientDriveProperties -rawOppOrProjTerm $_ -allClientTerms $allClientTerms}
        $newOpps            = @($newOpps            | Select-Object) | % {set-standardisedClientDriveProperties -rawOppOrProjTerm $_ -allClientTerms $allClientTerms}
        $existingOpps       = @($existingOpps       | Select-Object) | % {set-standardisedClientDriveProperties -rawOppOrProjTerm $_ -allClientTerms $allClientTerms}
        [array]$misplacedOpps = $orphanedOppFolders | ? {[string]::IsNullOrWhiteSpace($_.DriveClientId)}
        [array]$misplacedOpps += $newOpps           | ? {[string]::IsNullOrWhiteSpace($_.DriveClientId)}
        [array]$misplacedOpps += $existingOpps      | ? {[string]::IsNullOrWhiteSpace($_.DriveClientId)}
        if($($misplacedOpps.Count) -gt 0){
            if($deltaSync -eq $false){@($misplacedOpps | Select-Object) | % {Write-Host "`t`t`t[$($_.UniversalOppName)][$($_.NetSuiteOppId)][$($_.NetSuiteClientId)]"}}
            $orphanedOppFolders = $orphanedOppFolders | ? {$misplacedOpps.id -notcontains $_.id}
            $newOpps            = $newOpps            | ? {$misplacedOpps.id -notcontains $_.id}
            $existingOpps       = $existingOpps       | ? {$misplacedOpps.id -notcontains $_.id}
            }
        }
    Write-Host "`t[$($orphanedOppFolders.Count+$newOpps.Count+$existingOpps.Count-$misplacedOpps.Count)]/[$($orphanedOppFolders.Count+$newOpps.Count+$existingOpps.Count)] Opps matched to Client Terms ([$($($orphanedOppFolders.Count+$newOpps.Count+$existingOpps.Count-$misplacedOpps.Count)*100/$($orphanedOppFolders.Count+$newOpps.Count+$existingOpps.Count))]%) in [$($matchingOppsToClients.TotalSeconds)] seconds. [$($misplacedOpps.Count)] Opps don't have a corresponding Client Term (there's probably a duplicate Prospect/Client in NetSuite blocking creation of the Term)"
        #endregion
    #region Prepare Projs datasets
    $matchingProjsToClients = Measure-Command {
        if($deltaSync -eq $true){
            [array]$newProjs = $allProjTerms | ? {[string]::IsNullOrEmpty($_.DriveItemId)}
            [array]$existingProjs = $allProjTerms | ? {![string]::IsNullOrEmpty($_.DriveItemId) -and $_.CustomProperties.flagForReprocessing -ne $false}
            }


        if($deltaSync -eq $false){
            $projComparison = Compare-Object -ReferenceObject @($allProjTerms | Select-Object) -DifferenceObject @($driveItemsProjFolders | Select-Object) -Property "DriveItemId" -IncludeEqual -PassThru
            [array]$newProjs = $allProjTerms | ? {$_.SideIndicator -eq "<="} 
            [array]$existingProjs = $allProjTerms | ? {$_.SideIndicator -eq "=="}
            #[array]$orphanedProjFolders = $projComparison | ? {$_.SideIndicator -eq "=>"}
            }

        $orphanedProjFolders = @($orphanedProjFolders | Select-Object) | % {set-standardisedClientDriveProperties -rawOppOrProjTerm $_ -allClientTerms $allClientTerms}
        $newProjs            = @($newProjs            | Select-Object) | % {set-standardisedClientDriveProperties -rawOppOrProjTerm $_ -allClientTerms $allClientTerms}
        $existingProjs       = @($existingProjs       | Select-Object) | % {set-standardisedClientDriveProperties -rawOppOrProjTerm $_ -allClientTerms $allClientTerms}
        [array]$misplacedProjs = $orphanedProjFolders | ? {[string]::IsNullOrWhiteSpace($_.DriveClientId)}
        [array]$misplacedProjs += $newProjs           | ? {[string]::IsNullOrWhiteSpace($_.DriveClientId)}
        [array]$misplacedProjs += $existingProjs      | ? {[string]::IsNullOrWhiteSpace($_.DriveClientId)}
        }
    Write-Host "`t[$($orphanedProjFolders.Count+$newProjs.Count+$existingProjs.Count-$misplacedProjs.Count)]/[$($orphanedProjFolders.Count+$newProjs.Count+$existingProjs.Count)] Projs matched to Client Terms ([$($($orphanedProjFolders.Count+$newProjs.Count+$existingProjs.Count-$misplacedProjs.Count)*100/$($orphanedProjFolders.Count+$newProjs.Count+$existingProjs.Count))]%) in [$($matchingProjsToClients.TotalSeconds)] seconds. [$($misplacedProjs.Count)] Projs don't have a corresponding Client Term (there's probably a duplicate Prospect/Client in NetSuite blocking creation of the Term)"
    if($($misplacedProjs.Count) -gt 0){
        if($deltaSync -eq $false){@($misplacedProjs | Select-Object) | % {Write-Host "`t`t`t[$($_.UniversalProjName)][$($_.NetSuiteProjId)][$($_.NetSuiteClientId)]"}}
        $orphanedProjFolders = $orphanedProjFolders | ? {$misplacedProjs.id -notcontains $_.id}
        $newProjs            = $newProjs            | ? {$misplacedProjs.id -notcontains $_.id}
        $existingProjs       = $existingProjs       | ? {$misplacedProjs.id -notcontains $_.id}
        }

        #endregion    
    #endregion

    #region ProcessData

    #region ProcessClientsData

        #region Orphaned Client DocLibs
        $tokenResponseSharePointBot = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponseSharePointBot -renewTokenExpiringInSeconds 3000 -aadAppCreds $appCredsSharePointBot
        Write-Host "`tProcessing [$($orphanedClientDocLibs.Count)] orphaned Client DocLibs"
        $emptyDocLibs = $orphanedClientDocLibs | ? {$_.quota.used -eq 0}
        Write-Host "`t`tOnly [$($emptyDocLibs.Count)] of these orphaned Client DocLibs are empty though, so we'll ignore the the remaining [$($orphanedClientDocLibs.Count-$emptyDocLibs.Count)]"
        @($emptyDocLibs | Select-Object) | % {
            $thisOrphanedDocLib = $_#orphanedClientDocLibs[0]
            $result = process-docLibs -tokenResponse $tokenResponseSharePointBot -standardisedSourceDocLib $thisOrphanedDocLib -confirmDeleteEmptyDocLib
            if($result -eq $true){$emptyDocLibs = $emptyDocLibs | ? {$_.DriveClientId -notcontains $thisOrphanedDocLib.DriveClientId}}
            }
        if($emptyDocLibs.Count -ge 1){
            Write-Host "`t`t[$($orphanedClientDocLibs.Count)] Orphaned Opportunity folders failed to process"
            [array]$nonEmptyOppFolders = $($($orphanedClientDocLibs | Group-Object -Property {$_.DriveItemSize -gt 0}) | ? {$_.Name -eq "True"}).Group
            Write-Host "`t`t`t[$($nonEmptyOppFolders.Count)] Orphaned Opportunity folders contain data and will need resolving manually:"
            $orphanedClientDocLibs | % {Write-Host "`t`t`t`t[$($_.DriveItemName)][$($_.DriveItemId)][$($_.DriveItemUrl)][$($_.DriveClientName)][$($_.DriveClientId)]"}
            #Report this via e-mail too
            }
        #endregion

        #region New Clients
        Write-Host "`tProcessing [$($newClients.Count)] new Clients"
        @($newClients| Select-Object) | % {
            $tokenResponseSharePointBot = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponseSharePointBot -renewTokenExpiringInSeconds 30 -aadAppCreds $appCredsSharePointBot
            $thisNewClientTerm = $_
            try{
                $newClientDrive = new-clientDocLib -tokenResponse $tokenResponseSharePointBot -clientTerm $thisNewClientTerm -ErrorAction Stop
                $thisNewClientTerm.SetCustomProperty("GraphDriveId",$newClientDrive.id)
                try{
                    Write-Verbose "`t`t`tTrying to update Term [$($thisNewClientTerm.Name)][$($thisNewClientTerm.CustomProperties.NetSuiteId)] with CustomProperties @{GraphDriveId=$($newGraphListDrive.id)}"
                    $thisNewClientTerm.Context.ExecuteQuery()
                    $thisNewClientTerm.SetCustomProperty("flagForReprocessing",$false) #If the previous ExecuteQuery() worked, deflag the Term so it doesn;t get processed next time
                    $thisNewClientTerm.Context.ExecuteQuery()
                    [array]$newClients = $newClients | ? {$_.Id -notcontains $thisNewClientTerm.Id} #Pop this Term for the to-process stack so we can see any failures at the end
                    }
                catch{
                    if($deltaSync -eq $false -and $_.Exception -match "Term update failed because of save conflict"){
                        #Do nothing - a deltaSync=$true iteration has probably already processed this
                        }
                    else{Write-Host -ForegroundColor Red "`t`t$(get-errorSummary -errorToSummarise $_)"}
                    }
                }
            catch{
                Write-Host -ForegroundColor Red "`t`t$(get-errorSummary -errorToSummarise $_)"
                }
            }

        if($newClients.Count -ge 1){
            Write-Host "`t`t[$($newClients.Count)] New Client DocLibs failed to create:"
            $newClients | % {Write-Host "`t`t`t[$($_.UniversalClientName)][$($_.NetSuiteClientId)]"}
            #Report this via e-mail too
            }    
        #endregion

        #region Existing Clients
        if($deltaSync -eq $true){
            $appCredsSharePointBot = $(get-graphAppClientCredentials -appName SharePointBot)
            $tokenResponseSharePointBot = get-graphTokenResponse -aadAppCreds $appCredsSharePointBot
            Write-Host "`tProcessing [$($thisExistingClient.Count)] existing Clients"
            $existingClients | % {
                $thisExistingClient = $_
                try{
                    $thisClientDrive = get-graphDrives -tokenResponse $tokenResponseSharePointBot -driveId $thisExistingClient.DriveClientId -ErrorAction Stop
                    $thisClientDrive = set-standardisedClientDriveProperties -rawClientDrive $thisClientDrive
                    if($thisExistingClient.UniversalClientNameSanitised -ne $thisClientDrive.UniversalClientNameSanitised){
                        Write-Host "`t`t`Updating DriveClientName `t[$($thisClientDrive.DriveClientName)] for Drive [$($thisClientDrive.DriveClientId)][$($thisClientDrive.webUrl)]"
                        Write-Host "`t`tto:`t`t`t`t`t`t`t[$($thisExistingClient.UniversalClientName)] from Term [$($thisExistingClient.NetSuiteClientId)][$($thisExistingClient.Id)]"
                        try{
                            $docLibUpdatedCorrectly = process-docLibs -tokenResponse $tokenResponseSharePointBot -standardisedSourceDocLib $thisClientDrive -renameAs $thisExistingClient.UniversalClientName -ErrorAction Stop
                            if($docLibUpdatedCorrectly -eq $true){
                                $thisExistingClient.SetCustomProperty("flagForReprocessing",$false)
                                try{
                                    Write-Verbose "`tTrying to deflag processed Client [$($thisExistingClient.UniversalClientName)]"
                                    $thisExistingClient.Context.ExecuteQuery()
                                    }
                                catch{Write-Host -ForegroundColor Red "`t`t$(get-errorSummary -errorToSummarise $_)"}
                                }
                            }
                        catch{
                            Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                            [array]$duffUpdatedOpps += @($thisClientDrive,$(get-errorSummary -errorToSummarise $_))
                            }
                        }
                    else{
                        Write-Host "`t`tThe Client Name for [$($thisExistingClient.UniversalClientName)][$($thisExistingClient.NetSuiteClientId)] doesn't seem to have changed. Not sure why this was flagForReprocessing, but deflagging now."
                        $thisExistingClient.SetCustomProperty("flagForReprocessing",$false)
                        try{
                            Write-Verbose "`tTrying to deflag processed Client [$($thisExistingClient.UniversalClientName)]"
                            $thisExistingClient.Context.ExecuteQuery()
                            }
                        catch{Write-Host -ForegroundColor Red "`t`t$(get-errorSummary -errorToSummarise $_)"}
                        }
                    }
                catch{return}

                }
            }

        if($deltaSync -eq $false){        
            $existingClientsNameComparison = process-comparison -subsetOfNetObjects $existingClients -allTermObjects $allClientDrives -idInCommon DriveClientId -propertyToTest UniversalClientNameSanitised -validate 
            [array]$existingTermClientsWithChangedName  = $existingClientsNameComparison["<="]
            [array]$existingDriveClientsWithChangedName = $existingClientsNameComparison["=>"]
                        #Yes: Update the DriveItemName, & set flagForReproccessing = $false
            $tokenResponseSharePointBot = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponseSharePointBot -renewTokenExpiringInSeconds 600 -aadAppCreds $appCredsSharePointBot
            Write-Host "`tProcessing [$($existingDriveClientsWithChangedName.Count)] existing Clients with changed Names"
            for($i=0;$i -lt $existingDriveClientsWithChangedName.Count; $i++){
                Write-Host "`t`t`Updating DriveClientName `t[$($existingDriveClientsWithChangedName[$i].DriveClientName)] for Drive [$($existingDriveClientsWithChangedName[$i].DriveClientId)][$($existingDriveClientsWithChangedName[$i].webUrl)]"
                Write-Host "`t`tto:`t`t`t`t`t`t`t[$($existingTermClientsWithChangedName[$i].UniversalClientName)] from Term [$($existingTermClientsWithChangedName[$i].NetSuiteClientId)][$($existingTermClientsWithChangedName[$i].Id)]"
                try{
                    $docLibUpdatedCorrectly = process-docLibs -tokenResponse $tokenResponseSharePointBot -standardisedSourceDocLib $existingDriveClientsWithChangedName[$i] -renameAs $existingTermClientsWithChangedName[$i].UniversalClientName -ErrorAction Stop
                    if($docLibUpdatedCorrectly -eq $true){
                        $existingTermClientsWithChangedName[$i].SetCustomProperty("flagForReprocessing",$false)
                        try{
                            Write-Verbose "`tTrying to deflag processed Client [$($existingTermClientsWithChangedName[$i].UniversalClientName)]"
                            $existingTermClientsWithChangedName[$i].Context.ExecuteQuery()
                            }
                        catch{
                            if($deltaSync -eq $false -and $_.Exception -match "Term update failed because of save conflict"){
                                #Do nothing - a deltaSync=$true iteration has probably already processed this
                                }
                            else{
                                Write-Host -ForegroundColor Red "`t`t$(get-errorSummary -errorToSummarise $_)"
                                [array]$duffUpdatedClients += @($existingDriveClientsWithChangedName[$i],$(get-errorSummary -errorToSummarise $_))
                                }
                            }
                        }
                    }
                catch{
                    Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                    [array]$duffUpdatedOpps += @($existingDriveClientsWithChangedName[$i],$(get-errorSummary -errorToSummarise $_))
                    }
                }

                        #No: Set flagForReproccessing = $false
            [array]$existingTermClientsWithOriginalName = $existingClientsNameComparison["=="] #We'll updated these once we've finished the deltaClients ones too.
            [array]$existingClientsIncorrectlyFlaggedForProcessing = $existingTermClientsWithOriginalName | ? {$_.CustomProperties.flagForReprocessing -ne $false}
            if($existingClientsIncorrectlyFlaggedForProcessing.Count -gt 0){
                Write-Host "`t`t[$($existingClientsIncorrectlyFlaggedForProcessing.Count)] seems to have been flagged for reprocessing, but they don't seem to have changed. Deflagging them."
                $existingClientsIncorrectlyFlaggedForProcessing | % {
                    $thisIncorrectlyFlaggedClient = $_
                    $thisIncorrectlyFlaggedClient.SetCustomProperty("flagForReprocessing",$false)
                    try{
                        Write-Verbose "`tTrying to deflag processed Client [$($thisIncorrectlyFlaggedClient.UniversalClientName)]"
                        Write-Host "`t`t`tDeflagging [$($thisIncorrectlyFlaggedClient.UniversalClientName)][$($thisIncorrectlyFlaggedClient.NetSuiteClientId)]"
                        $thisIncorrectlyFlaggedClient.Context.ExecuteQuery()
                        }
                    catch{
                        if($deltaSync -eq $false -and $_.Exception -match "Term update failed because of save conflict"){
                            #Do nothing - a deltaSync=$true iteration has probably already processed this
                            }
                        else{
                            Write-Host -ForegroundColor Red "`t`t$(get-errorSummary -errorToSummarise $_)"
                            [array]$duffUpdatedClients += @($thisIncorrectlyFlaggedClient,$(get-errorSummary -errorToSummarise $_))
                            }
                        }
                    }
                }
            }
        #endregion

        #Does Term have a DriveClientId?
            #No: 
                #Create a new Drive
                #Update the Term DriveId & deflag
            #Yes: 
                #Has the Name changed?
                    #Yes: Update the DriveName
                    #No: Deflag the Term

    #endregion

    #region ProcessOpportunities

        #region orphanedOpps
    $tokenResponseSharePointBot = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponseSharePointBot -renewTokenExpiringInSeconds 3000 -aadAppCreds $appCredsSharePointBot
    Write-Host "`tProcessing [$($orphanedOppFolders.Count)] orphaned Opportunities"
    @($orphanedOppFolders | Select-Object) | % {
        $thisOrphanedFolder = $_#orphanedOppFolders[0]
        $result = process-folders -tokenResponse $tokenResponseSharePointBot -standardisedSourceFolder $thisOrphanedFolder -confirmDeleteEmptyFolders
        if($result -eq $true){$orphanedOppFolders = $orphanedOppFolders | ? {$_.DriveItemId -notcontains $thisOrphanedFolder.DriveItemId}}
        }
    if($orphanedOppFolders.Count -ge 1){
        Write-Host "`t`t[$($orphanedOppFolders.Count)] Orphaned Opportunity folders failed to process"
        [array]$nonEmptyOppFolders = $($($orphanedOppFolders | Group-Object -Property {$_.DriveItemSize -gt 0}) | ? {$_.Name -eq "True"}).Group
        Write-Host "`t`t`t[$($nonEmptyOppFolders.Count)] Orphaned Opportunity folders contain data and will need resolving manually:"
        $orphanedOppFolders | % {Write-Host "`t`t`t`t[$($_.DriveItemName)][$($_.DriveItemId)][$($_.DriveItemUrl)][$($_.DriveClientName)][$($_.DriveClientId)]"}
        #Report this via e-mail too
        }
        #endregion

        #region newOpps
    Write-Host "`tProcessing [$($newOpps.Count)] new Opportunities"
    [array]$newOppsWithProjectIds = $newOpps | ? {![string]::IsNullOrWhiteSpace($_.NetSuiteProjectId)}
    if($newOppsWithProjectIds.Count -gt 0){
        Write-Host "`t`tExcluding [$($newOppsWithProjectIds.Count)] Opps because they already have Projects (not recreating Opp Folders), [$($newOpps.Count - $newOppsWithProjectIds.Count)] Opps remaining"
        [array]$newOpps = $newOpps | ? {$newOppsWithProjectIds.id -notcontains $_.id}
        }
    [array]$newOppsWithoutClientDriveIds = $newOpps | ? {[string]::IsNullOrWhiteSpace($_.DriveClientId)}
    if($newOppsWithoutClientDriveIds.Count -gt 0){
        Write-Host "`t`tExcluding [$($newOppsWithoutClientDriveIds.Count)] Opps because they do not have a ClientDriveId (the Client Term creating is _probably_ being blocked by a duplicate name in NetSuite), [$($newOpps.Count - $newOppsWithoutClientDriveIds.Count)] Opps remaining"
        [array]$newOpps = $newOpps | ? {$newOppsWithoutClientDriveIds.id -notcontains $_.id}
        }
    @($newOpps| Select-Object) | % {
        $tokenResponseSharePointBot = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponseSharePointBot -renewTokenExpiringInSeconds 30 -aadAppCreds $appCredsSharePointBot
        $thisNewOppTerm = $_
        try{
            [array]$newOppFolders = new-oppProjFolders -tokenResponse $tokenResponseSharePointBot -oppProjTermWithClientInfo $thisNewOppTerm #-Verbose
            if($newOppFolders.Count -ge 1 -and ![string]::IsNullOrWhiteSpace($newOppFolders[0].id)){
                $thisNewOppTerm.SetCustomProperty("DriveItemId",$newOppFolders[0].id)
                $thisNewOppTerm.SetCustomProperty("flagForReprocessing",$false)
                try{
                    $thisNewOppTerm.Context.ExecuteQuery()
                    [array]$newOpps = $newOpps | ? {$_.DriveItemId -notcontains $thisNewOppTerm.DriveItemId}
                    }
                catch{get-errorSummary -errorToSummarise $_}
                }
            }
        catch{get-errorSummary -errorToSummarise $_}

        }

    if($newOpps.Count -ge 1){
        Write-Host "`t`t[$($newOpps.Count)] New Opportunity folders failed to create:"
        $newOpps | % {Write-Host "`t`t`t[$($_.UniversalOppName)][$($_.Id)][$($_.NetSuiteOppId)] for NetSuiteClientId [$($_.NetSuiteClientId)]"}
        #Report this via e-mail too
        }

    if($deltaSync -eq $false){
        #Deflag the stragglers
        Write-Host "`t[$($newOppsWithProjectIds.Count)] Opps were excluded because they already have Projects (not recreating Opp Folders):"
        $newOppsWithProjectIds | % {
            Write-Host "`t`tDeflagging [$($_.UniversalOppName)][$($_.UniversalClientName)]"
            $_.SetCustomProperty("flagForReprocessing",$false)
            try{$_.Context.ExecuteQuery()}
            catch{get-errorSummary -errorToSummarise $_}
            }
        Write-Host "`t[$($newOppsWithoutClientDriveIds.Count)] Opps were excluded because they have no Client, or their Client has no ClientDriveId (the Client Term creating is _probably_ being blocked by a duplicate name in NetSuite)(cannot create Opp Folders) :"
        $newOppsWithoutClientDriveIds | % {
            Write-Host "`t`tDeflagging [$($_.UniversalOppName)][$($_.UniversalClientName)][$($_.NetSuiteClientId)]"
            $_.SetCustomProperty("flagForReprocessing",$false)
            try{$_.Context.ExecuteQuery()}
            catch{get-errorSummary -errorToSummarise $_}
            }
        }
        #endregion

        #region existingOpps
        #Does Term have a TermProjId?
            #Yes: Do nothing to Opps that have been won & set flagForReproccessing = $false
            #No:
                #Has the Name changed?
                    #Yes: Update the DriveItemName, & set flagForReproccessing = $false
                    #No: Set flagForReproccessing = $false
                #Has the Client changed?
                    #Yes: Update the NetSuiteClientId, & set flagForReproccessing = $false
                    #No: Dedupe & set flagForReproccessing = $false

        #Does Term have a TermProjId?
        $existingOppTermsWithProject = $existingOpps    | ? {![string]::IsNullOrWhiteSpace($_.NetSuiteProjectId) -and $_.CustomProperty.flagForReprocessing -ne $false}
            #Yes: Do nothing to Opps that have been won & set flagForReproccessing = $false
            #Actually, if $deltaSync -eq $false, check whether the Opp folder still exists _and_ the Proj folder exists. If both > migrate Opp folder data to Proj folder and delete Opp folder
        Write-Host "`tProcessing [$($existingOppTermsWithProject.Count)] existing Opportunities with Projects"
        @($existingOppTermsWithProject | Select-Object) | % {
            $thisWonOppTerm = $_
            #Actually, if $deltaSync -eq $false, check whether the Opp folder still exists _and_ the Proj folder exists. If both > migrate Opp folder data to Proj folder and delete Opp folder
            if($deltaSync -eq $false){ #This relies on us having enumerated $topLevelFolders
                $thisWonOppFolder = Compare-Object $topLevelFolders -DifferenceObject $thisWonOppTerm -Property DriveClientId,DriveItemId -ExcludeDifferent -IncludeEqual -PassThru #Check whether there is still an OppFolder
                if(![string]::IsNullOrWhiteSpace($thisWonOppFolder)){ #If the Opp folder hasn't been migrated, find the corresponding Proj folder and try again
                    $thisWonProjTerm = Compare-Object $allProjTerms -DifferenceObject $thisWonOppTerm -Property NetSuiteProjectId -ExcludeDifferent -IncludeEqual -PassThru
                    if(![string]::IsNullOrWhiteSpace($thisWonProjTerm)){
                        $thisWonProjFolder =  Compare-Object $topLevelFolders -DifferenceObject $thisWonProjTerm -Property DriveClientId,DriveItemId -ExcludeDifferent -IncludeEqual -PassThru
                        if(![string]::IsNullOrWhiteSpace($thisWonProjFolder)){ #If we have both $thisWonOppFolder _and_ $thisWonProjFolder, try to merge the OppFolder into the ProjFolder
                            try{
                                Write-Host "`t`tOpp [$($thisWonOppTerm.UniversalOppName)][$($thisWonOppTerm.NetSuiteOppId)][$($thisWonOppTerm.Id)] for [$($thisWonOppTerm.UniversalClientName)][$($thisWonOppTerm.NetSuiteClientId)] has an orphaned Opp folder - attempting to tidy it up automatically"
                                $didItWork = process-folders -tokenResponse $tokenResponseSharePointBot -standardisedSourceFolder $thisWonOppFolder -mergeInto $thisWonProjFolder
                                if($didItWork.weburl -match $thisWonProjFolder.DriveItemUrl -or [string]::IsNullOrWhiteSpace($didItWork)){
                                    Write-Host "`t`t`tSuccess: [$(@($didItWork.id | Select-Object).Count)] folders with data moved. Attempting to delete orphaned Opp Folder"
                                    $isTheOppFolderGone = process-folders -tokenResponse $tokenResponseSharePointBot -standardisedSourceFolder $thisWonOppFolder -confirmDeleteEmptyFolders
                                    if($isTheOppFolderGone -eq $true){Write-Host "`t`t`t`tSuccess: Opp Folder [$($thisWonOppFolder.DriveItemUrl)] was empty - removed successfully!"}
                                    else{Write-Host "`t`t`t`Failed to remove Opp Folder [$($thisWonOppFolder.DriveItemUrl)]"}
                                    }
                                }
                            catch{get-errorSummary $_}
                            }
                        else{<#This is an error - there _should_ be a Proj Folder for all Projs because *new* Projs are processed after this region#>}
                        }
                    else{<#This is an error - there _should_ be a Proj Term for all Opps with a ProjId because *new* Projs are processed after this region#>}
                    }
                else{<#Do nothing - if there is no OppFolder, it must have been process correctly already#>}
                
                }
            
            Write-Host "`t`tDeflagging Opp [$($thisWonOppTerm.UniversalOppName)][$($thisWonOppTerm.NetSuiteOppId)][$($thisWonOppTerm.Id)] for [$($thisWonOppTerm.UniversalClientName)][$($thisWonOppTerm.NetSuiteClientId)] (Opps no longer control Folders once they have been converted to Projects)"
            $thisWonOppTerm.SetCustomProperty("flagForReprocessing",$false)
            try{$thisWonOppTerm.Context.ExecuteQuery()}
            catch{Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"}
            }

            #No:
                #Has the Name changed?
        $existingOppTermsWithoutProject = $existingOpps | ? { [string]::IsNullOrWhiteSpace($_.NetSuiteProjectId)}                  
        if($deltaSync -eq $true){
            Write-Host "`t[$($existingOppTermsWithoutProject.Count)] existing Opps need examining to see if anything has changed"
            $tokenResponseSharePointBot = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponseSharePointBot -renewTokenExpiringInSeconds 600 -aadAppCreds $appCredsSharePointBot
            @($existingOppTermsWithoutProject | Select-Object) | % {
                $thisExistingOpp = $_
                try{
                    try{$thisExistingOppDriveItem = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $thisExistingOpp.DriveClientId -itemGraphId $thisExistingOpp.DriveItemId -returnWhat Item -ErrorAction Stop}         #Try to get the link DriveItem so we can test whether it needs updating. -ErrorAction SilentlyContinue is ignored inside the outer Try/Catch block,s o we need another one just for this command :/
                    catch{$thisExistingOppDriveItem = $null} #This is required to prevent the next itreration getting cross-linked with this DriveItem
                    if([string]::IsNullOrEmpty($thisExistingOppDriveItem.id)){Write-Warning "`t`tOppDriveItem [$($thisExistingOpp.UniversalOppName)][$($thisExistingOpp.NetSuiteOppId)][$($thisExistingOpp.Id)] for [$($thisExistingOpp.UniversalClientName)][$($thisExistingOpp.NetSuiteClientId)] is missing. It might have been assigned to a different Client (which will be fixed on the next Full Reconcile), or it may have been manually moved/deleted."}
                    elseif($(sanitise-forNetsuiteIntegration $thisExistingOppDriveItem.name) -ne $thisExistingOpp.UniversalOppNameSanitised){
                        $thisExistingOppDriveItem | Add-Member -MemberType NoteProperty -Name DriveItemName -Value $thisExistingOppDriveItem.name -Force
                        $thisExistingOppDriveItem | Add-Member -MemberType NoteProperty -Name DriveItemId -Value $thisExistingOppDriveItem.id -Force
                        $thisExistingOppDriveItem | Add-Member -MemberType NoteProperty -Name DriveClientName -Value $thisExistingOpp.UniversalClientName -Force
                        $thisExistingOppDriveItem | Add-Member -MemberType NoteProperty -Name DriveClientId -Value $thisExistingOpp.DriveClientId -Force
                        Write-Host "`tUpdating OppDriveItem Name`t[$($thisExistingOppDriveItem.DriveItemName)][$($thisExistingOppDriveItem.DriveItemId)][$($thisExistingOppDriveItem.DriveClientId)] "
                        Write-Host "`tTo:`t`t`t`t`t`t`t[$($thisExistingOpp.UniversalOppName)][$($thisExistingOpp.NetSuiteOppId)][$($thisExistingOpp.Id)] for [$($thisExistingOpp.UniversalClientName)][$($thisExistingOpp.NetSuiteClientId)]"
                        try{$updatedFolder = process-folders -tokenResponse $tokenResponseSharePointBot -standardisedSourceFolder $thisExistingOppDriveItem -renameAs $thisExistingOpp.UniversalOppName -Verbose -ErrorAction Stop}
                        catch{Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"}
                        if($updatedFolder.name -eq $thisExistingOpp.UniversalOppName){
                            $thisExistingOpp.SetCustomProperty("flagForReprocessing",$false)
                            try{$thisExistingOpp.Context.ExecuteQuery()}
                            catch{Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"}
                            }
                        }
                    else{
                        Write-Host "`t`tWeird - it doesn't look like Opp [$($thisExistingOpp.UniversalOppName)][$($thisExistingOpp.NetSuiteOppId)][$($thisExistingOpp.Id)] for [$($thisExistingOpp.UniversalClientName)][$($thisExistingOpp.NetSuiteClientId)] needed updating..."
                        }
                    }
                catch{Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"}
                }


        
            }
        if($deltaSync -eq $false){ #Full reconcile
            $existingOppWithoutProjectNameComparison = process-comparison -subsetOfNetObjects $existingOppTermsWithoutProject -allTermObjects $driveItemsOppFolders -idInCommon DriveItemId -propertyToTest UniversalOppNameSanitised -validate
            [array]$existingTermOppsWithoutProjectWithChangedName  = $existingOppWithoutProjectNameComparison["<="]
            [array]$existingDriveOppsWithoutProjectWithChangedName = $existingOppWithoutProjectNameComparison["=>"]
                        #Yes: Update the DriveItemName, & set flagForReproccessing = $false
            Write-Host "`tProcessing [$($existingDriveOppsWithoutProjectWithChangedName.Count)] existing Opportunities with changed Names"
            for($i=0;$i -lt $existingDriveOppsWithoutProjectWithChangedName.Count; $i++){
                Write-Host "`t`t`Updating DriveItemName `t[$($existingDriveOppsWithoutProjectWithChangedName[$i].DriveItemName)] for DriveItem [$($existingDriveOppsWithoutProjectWithChangedName[$i].DriveItemId)][$($existingDriveOppsWithoutProjectWithChangedName[$i].DriveItemUrl)][$($existingDriveOppsWithoutProjectWithChangedName[$i].DriveClientName)][$($existingDriveOppsWithoutProjectWithChangedName[$i].DriveClientId)]"
                Write-Host "`t`tto:`t`t`t`t`t`t[$($existingTermOppsWithoutProjectWithChangedName[$i].UniversalOppName)] from Term `t[$($existingTermOppsWithoutProjectWithChangedName[$i].NetSuiteOppId)][$($existingTermOppsWithoutProjectWithChangedName[$i].Id)][$($existingTermOppsWithoutProjectWithChangedName[$i].NetSuiteClientId)]"
                try{
                    $updatedFolder = process-folders -tokenResponse $tokenResponseSharePointBot -standardisedSourceFolder $existingDriveOppsWithoutProjectWithChangedName[$i] -renameAs $existingTermOppsWithoutProjectWithChangedName[$i].UniversalOppName -ErrorAction Stop
                    $existingTermOppsWithoutProjectWithChangedName[$i].SetCustomProperty("flagForReprocessing",$false)
                    try{
                        Write-Verbose "`tTrying to deflag processed Opp [$($existingTermOppsWithoutProjectWithChangedName[$i].UniversalOppName)]"
                        $existingTermOppsWithoutProjectWithChangedName[$i].Context.ExecuteQuery()
                        }
                    catch{
                        Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                        [array]$duffUpdatedOpps += @($existingDriveOppsWithoutProjectWithChangedName[$i],$(get-errorSummary -errorToSummarise $_))
                        }
                    }
                catch{
                    if($_.Exception -match "cannot contain any illegal characters"){
                        Write-Warning "Illegal characters in OppName [$($existingTermOppsWithoutProjectWithChangedName[$i].UniversalOppName)][$($existingTermOppsWithoutProjectWithChangedName[$i].NetSuiteOppId)][$($existingTermOppsWithoutProjectWithChangedName[$i].Id)] for [$($existingTermOppsWithoutProjectWithChangedName[$i].UniversalClientName)][$($existingTermOppsWithoutProjectWithChangedName[$i].NetSuiteClientId)]"
                        [array]$objectsWithIllegalCharacters += $existingTermOppsWithoutProjectWithChangedName[$i]
                        }
                    else{Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"}
                    [array]$duffUpdatedOpps += @($existingDriveOppsWithoutProjectWithChangedName[$i],$(get-errorSummary -errorToSummarise $_))
                    }
                }
                        #No: Set flagForReproccessing = $false
            [array]$existingDriveOppsWithoutProjectWithOriginalName = $existingOppWithoutProjectNameComparison["=="] #We'll updated these once we've finished the deltaClients ones too.

                #Has the Client changed?
            $oppExpectedDriveIdComparion = process-comparison -subsetOfNetObjects $existingOppTermsWithoutProject -allTermObjects $driveItemsOppFolders -idInCommon "DriveItemId" -propertyToTest "DriveClientId" -validate 
            $existingOppTermsWithMismatchedClients = $oppExpectedDriveIdComparion["<="]
            $existingOppDrivesWithMismatchedClients = $oppExpectedDriveIdComparion["=>"]
            Write-Host "`tProcessing [$($existingOppDrivesWithMismatchedClients.Count)] existing Opportunities with changed Clients"
            $tokenResponseSharePointBot = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponseSharePointBot -renewTokenExpiringInSeconds 30 -aadAppCreds $appCredsSharePointBot
            for($i=0;$i -lt $existingOppDrivesWithMismatchedClients.Count; $i++){
                    #Yes: Move the DriveItems, Update the DriveItemId, & set flagForReproccessing = $false
                Write-Host "`t`tMoving existing Opp folder from`t[$($existingOppDrivesWithMismatchedClients[$i].DriveClientName)][$($existingOppDrivesWithMismatchedClients[$i].DriveClientId)] (OppFolder is [$($existingOppDrivesWithMismatchedClients[$i].UniversalOppName)][$($existingOppTermsWithMismatchedClients[$i].NetSuiteOppId)][$($existingOppDrivesWithMismatchedClients[$i].DriveItemUrl)]"
                Write-Host "`t`tto`t`t`t`t`t`t`t`t[$($existingOppTermsWithMismatchedClients[$i].UniversalClientName)][$($existingOppTermsWithMismatchedClients[$i].DriveClientId)][$($existingOppTermsWithMismatchedClients[$i].NetSuiteClientId)] (OppTerm is [$($existingOppTermsWithMismatchedClients[$i].UniversalOppName)][$($existingOppTermsWithMismatchedClients[$i].NetSuiteOppId)]"
                try{
                    try{$newDestinationFolder = add-graphFolderToDrive -graphDriveId $existingOppTermsWithMismatchedClients[$i].DriveClientId -folderName $existingOppTermsWithMismatchedClients[$i].UniversalOppName -tokenResponse $tokenResponseSharePointBot -conflictResolution Fail}
                    catch{Write-Verbose = "Condonable fail - if this errors, the target folder has already been created on a previous iteration"}                    
                    $movedFolders = process-folders -tokenResponse $tokenResponseSharePointBot -standardisedSourceFolder $existingOppDrivesWithMismatchedClients[$i] -mergeInto $newDestinationFolder -ErrorAction Continue
                    if($movedFolders[0].parentReference.driveId -eq $existingOppDrivesWithMismatchedClients[$i].DriveClientId){
                        Write-Host "`t`t`tFailed to move these [$($movedFolders.count)] folders:"
                        @($movedFolders | Select-Object) | % {Write-Host "`t`t`t[$($_.name)][$($_.weburl)]"}
                        }
                    else{
                        $existingOppTermsWithMismatchedClients[$i].SetCustomProperty("DriveItemId",$newDestinationFolder.id)
                        $existingOppTermsWithMismatchedClients[$i].SetCustomProperty("flagForReprocessing",$false)
                        try{$existingOppTermsWithMismatchedClients[$i].Context.ExecuteQuery()}
                        catch{Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"}
                        try{
                            $result = process-folders -tokenResponse $tokenResponseSharePointBot -standardisedSourceFolder $existingOppDrivesWithMismatchedClients[$i] -confirmDeleteEmptyFolders #Finally, try to delete any empty folder
                            if($result -ne $true){Write-Warning "Failed to delete (hopefully) empty folder [$($existingOppDrivesWithMismatchedClients[$i].DriveItemUrl)]"}
                            }
                        catch{Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"}
                        }
                    }
                catch{
                    Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                    }
                }

                    #No: Dedupe & set flagForReprocessing = $false
            [array]$existingDriveOppsWithoutProjectWithOriginalClient = $oppExpectedDriveIdComparion["=="] 
            $dedupedOppsWithOriginalClientAndOriginalName = [System.Collections.Generic.Hashset[Microsoft.SharePoint.Client.Taxonomy.TermSetItem]] ($existingDriveOppsWithoutProjectWithOriginalName + $existingDriveOppsWithoutProjectWithOriginalClient)
            [array]$dedupedOppsStillFlaggedForProcessing = $dedupedOppsWithOriginalClientAndOriginalName | ? {$_.CustomProperties.flagForReprocessing -ne $false}
            if($dedupedOppsStillFlaggedForProcessing.Count -gt 0){
                Write-Host "`t[$($dedupedOppsStillFlaggedForProcessing.Count)] Opportunity Terms were flagged for reprocessing, but they don't seem to have any changes. This isn't specifically a _problem_, but it's an indication that reconcile-netSuiteToTermStore() is incorrectly flagging Opportunities as requiring processing when they don't"
                $dedupedOppsStillFlaggedForProcessing | % {
                    $thisdedupedOppStillFlaggedForProcessing = $_
                    Write-Host "`t`t[$($thisdedupedOppStillFlaggedForProcessing.UniversalOppName)][$($thisdedupedOppStillFlaggedForProcessing.NetSuiteOppId)][$($thisdedupedOppStillFlaggedForProcessing.id)] for [$($thisdedupedOppStillFlaggedForProcessing.UniversalClientName)][$($thisdedupedOppStillFlaggedForProcessing.NetSuiteClientId)]"
                    $thisdedupedOppStillFlaggedForProcessing.SetCustomProperty("flagForReprocessing",$false)
                    try{$thisdedupedOppStillFlaggedForProcessing.Context.ExecuteQuery()}
                    catch{Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"}
                    }
                }

            }




        #endregion
    #endregion

    #region ProcessProjectsData


        #region Orphaned Projects
    if($deltaSync -eq $false){
        $tokenResponseSharePointBot = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponseSharePointBot -renewTokenExpiringInSeconds 3000 -aadAppCreds $appCredsSharePointBot

        Write-Host "`tProcessing [$($orphanedProjFolders.Count)] orphaned Projects"
        @($orphanedProjFolders | Select-Object) | % {
            $thisOrphanedFolder = $_#orphanedProjFolders[0]
            $result = process-folders -tokenResponse $tokenResponseSharePointBot -standardisedSourceFolder $thisOrphanedFolder -confirmDeleteEmptyFolders
            if($result -eq $true){$orphanedProjFolders = $orphanedProjFolders | ? {$_.DriveItemId -notcontains $thisOrphanedFolder.DriveItemId}}#
            }
        if($orphanedProjFolders.Count -ge 1){
            Write-Host "`t`t[$($orphanedProjFolders.Count)] Orphaned Projects folders failed to process"
            [array]$nonEmptyOppFolders = $($($orphanedProjFolders | Group-Object -Property {$_.DriveItemSize -gt 0}) | ? {$_.Name -eq "True"}).Group
            Write-Host "`t`t`t[$($nonEmptyOppFolders.Count)] Orphaned Projects folders contain data and will need resolving manually:"
            $orphanedProjFolders | % {Write-Host "`t`t`t`t[$($_.DriveItemName)][$($_.DriveItemId)][$($_.DriveItemUrl)][$($_.DriveClientName)][$($_.DriveClientId)]"}
            #Report this via e-mail too
            }

        <#Do some clever self-healing next
        $projFoldersWithMatchingCodes = Compare-Object -ReferenceObject $driveItemsProjFolders -DifferenceObject $allProjTerms -Property UniversalProjCode -PassThru -IncludeEqual -ExcludeDifferent
    
        #dedupe from above $projComparison | ? {$_.SideIndicator -eq "=>"}
        $projFoldersWithMatchingCodes2 = $projFoldersWithMatchingCodes | ? {$($projComparison | ? {$_.SideIndicator -eq "=>"}).id -notcontains $_.id}
        $projFolderCodeComparison = process-comparison -subsetOfNetObjects $projFoldersWithMatchingCodes -allTermObjects $allProjTerms -idInCommon UniversalProjCode -propertyToTest DriveItemId -validate
        $additionalOrphanedProjFolders = $projFolderCodeComparison["<="]
        $additionalOrphanedProjTerms   = $projFolderCodeComparison["=>"]
        $tokenResponseSharePointBot = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponseSharePointBot -renewTokenExpiringInSeconds 600 -aadAppCreds $appCredsSharePointBot
        for($i=0;$i -lt $additionalOrphanedProjTerms.Count;$i++){
           #if($additionalOrphanedProjTerms[$i].name -match "O-1002467"){Write-Host -f Yellow $i;break}
            if($additionalOrphanedProjTerms[$i].DriveClientId -eq $additionalOrphanedProjFolders[$i].DriveClientId){
                if([string]::IsNullOrEmpty($additionalOrphanedProjTerms[$i].DriveItemId)){
                    #Link
                    Write-Host "1-[$($additionalOrphanedProjFolders[$i].DriveItemName)][$($additionalOrphanedProjFolders[$i].DriveItemId)] is in the correct Drive [$($additionalOrphanedProjTerms[$i].UniversalClientName)][$($additionalOrphanedProjTerms[$i].NetSuiteClientId)][$($additionalOrphanedProjTerms[$i].DriveClientId)], and the Term has no DriveItemId - linking to this folder"
                    #$additionalOrphanedProjTerms[$i].SetCustomProperty("DriveItemId",$additionalOrphanedProjFolders[$i].DriveItemId)
                    #$additionalOrphanedProjTerms[$i].Context.ExecuteQuery()
                    }
                else{
                    $testPath = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $additionalOrphanedProjTerms[$i].DriveClientId -itemGraphId $additionalOrphanedProjTerms[$i].DriveItemId -returnWhat Item -ErrorAction SilentlyContinue
                    #Test & 
                    if([string]::IsNullOrEmpty($testPath)){
                        Write-Host "2-[$($additionalOrphanedProjFolders[$i].DriveItemName)][$($additionalOrphanedProjFolders[$i].DriveItemId)] is in the correct Drive [$($additionalOrphanedProjTerms[$i].UniversalClientName)][$($additionalOrphanedProjTerms[$i].NetSuiteClientId)][$($additionalOrphanedProjTerms[$i].DriveClientId)], and the Term's current DriveItemId is invalid - linking to this folder"
                        #$additionalOrphanedProjTerms[$i].SetCustomProperty("DriveItemId",$additionalOrphanedProjFolders[$i].DriveItemId)
                        #$additionalOrphanedProjTerms[$i].Context.ExecuteQuery()
                        }
                    else{
                        Write-Host "3-[$($additionalOrphanedProjFolders[$i].DriveItemName)][$($additionalOrphanedProjFolders[$i].DriveItemId)] is in the correct Drive [$($additionalOrphanedProjTerms[$i].UniversalClientName)][$($additionalOrphanedProjTerms[$i].NetSuiteClientId)][$($additionalOrphanedProjTerms[$i].DriveClientId)], but the Term's current DriveItemId is valid - deleting this incorrect folder"
                        #$result = process-folders -tokenResponse $tokenResponseSharePointBot -standardisedSourceFolder $additionalOrphanedProjFolders[$i] -confirmDeleteEmptyFolders
                        }
                    }
                }
            else{
                Write-Host "4-[$($additionalOrphanedProjFolders[$i].DriveItemName)][$($additionalOrphanedProjFolders[$i].DriveItemId)] is in the wrong Drive [$($additionalOrphanedProjTerms[$i].UniversalClientName)][$($additionalOrphanedProjTerms[$i].NetSuiteClientId)][$($additionalOrphanedProjTerms[$i].DriveClientId)] - deleting this incorrect folder"
                #$result = process-folders -tokenResponse $tokenResponseSharePointBot -standardisedSourceFolder $additionalOrphanedProjFolders[$i] -confirmDeleteEmptyFolders
                }
            }#>
        }

        #endregion

        #region New Projects
        #Does the Term have a DriveItemId?
            #No: 
                #Can we find a corresponding Opp?
                    #Yes: Re-use DriveItemId, rename folder & set flagForReprocessing = $false
                    #No: Create a new DriveItem, & set flagForReprocessing = $false
        Write-Host "`tProcessing [$($newProjs.Count)] new Projects"
        @($newProjs | Select-Object) | % {
            $thisNewProjTerm = $_
            $tokenResponseSharePointBot = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponseSharePointBot -renewTokenExpiringInSeconds 30 -aadAppCreds $appCredsSharePointBot
                #Can we find a corresponding Opp?
            $correspondingOpp = Compare-Object -ReferenceObject $allOppTerms -DifferenceObject $thisNewProjTerm -Property NetSuiteProjectId -ExcludeDifferent -IncludeEqual -PassThru
            if(![string]::IsNullOrWhiteSpace($correspondingOpp)){$correspondingOpp = set-standardisedClientDriveProperties -rawOppOrProjTerm $correspondingOpp -allClientTerms $allClientTerms} #Set the friendly property names in case we haven;t already done this
            if(![string]::IsNullOrEmpty($correspondingOpp.DriveItemId)){
                Write-Host "`t`tCorresponding Opp [$($correspondingOpp.UniversalOppName)][$($correspondingOpp.DriveClientId)][$($correspondingOpp.DriveItemId)] found for [$($thisNewProjTerm.UniversalProjName)][$($thisNewProjTerm.NetSuiteProjectId)][$($thisNewProjTerm.NetSuiteClientId)]"
                try{
                    $existingOppDriveItem = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $correspondingOpp.DriveClientId -itemGraphId $correspondingOpp.DriveItemId -returnWhat Item -ErrorAction Stop
                    #Yes: Re-use DriveItemId, rename folder & set flagForReprocessing = $false   <---Do this
                    Write-Host "`t`t`tValid OppFolder retrieved from Drive: Updating Project to re-use OppFolder, renaming Folder, deflagging Project Term"
                    $thisNewProjTerm.SetCustomProperty("DriveItemId",$correspondingOpp.DriveItemId)
                    try{
                        $thisNewProjTerm.Context.ExecuteQuery()
                        try{
                            $existingOppDriveItem | Add-Member -MemberType NoteProperty -Name DriveItemName -Value $existingOppDriveItem.name -Force
                            $existingOppDriveItem | Add-Member -MemberType NoteProperty -Name DriveItemId -Value $existingOppDriveItem.id -Force
                            $existingOppDriveItem | Add-Member -MemberType NoteProperty -Name DriveClientName -Value $correspondingOpp.UniversalClientName -Force
                            $existingOppDriveItem | Add-Member -MemberType NoteProperty -Name DriveClientId -Value $correspondingOpp.DriveClientId -Force
                            $didItWork = process-folders -tokenResponse $tokenResponseSharePointBot -standardisedSourceFolder $existingOppDriveItem -renameAs $thisNewProjTerm.UniversalProjName -Verbose
                            if($didItWork -eq $true){
                                $thisNewProjTerm.SetCustomProperty("flagForReprocessing",$false)
                                $thisNewProjTerm.Context.ExecuteQuery()
                                [array]$newProjs = $newProjs | ? {$_.id -notcontains $thisNewProjTerm.Id} #If it worked, pop this Proj from the to-do list
                                }
                            }
                        catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
                        }
                    catch{
                        Write-Error "Error updating DriveItemId for Proj folder [$($thisNewProjTerm.UniversalProjName)][$($thisNewProjTerm.NetSuiteProjectId)][$($thisNewProjTerm.NetSuiteClientId)] to [$($correspondingOpp.DriveItemId)] (from Opp [$($correspondingOpp.DriveClientId)][$($correspondingOpp.DriveClientId)])"
                        Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)
                        }
                    }
                catch{
                    if($_.Exception -match "404" -or $_.InnerException -match "404"){$thisNewProjTerm | Add-Member -MemberType NoteProperty -Name RecreateFolders -Value $true -Force} #If the folder doesn't exist, mark the Project as needing folders to be created
                    else{
                        Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)
                        #Write-Error "Error retrieving Opp folder [$($correspondingOpp.UniversalOppName)][$($correspondingOpp.DriveClientId)][$($correspondingOpp.DriveItemId)] for Proj [$($thisNewProjTerm.UniversalProjName)][$($thisNewProjTerm.NetSuiteProjectId)][$($thisNewProjTerm.NetSuiteClientId)]"
                        }
                    }
                }
            if( [string]::IsNullOrEmpty($correspondingOpp.DriveItemId) -or $thisNewProjTerm.RecreateFolders -eq $true){ #If the Opp didn't have a DriveItemId set, or if did but the folder doesn't exist. 
                    #No: Create a new DriveItem, & set flagForReprocessing = $false
                if([string]::IsNullOrEmpty($correspondingOpp.DriveItemId)){
                    if([string]::IsNullOrEmpty($correspondingOpp.id)){Write-Host "`t`tNo corresponding Opp found for [$($thisNewProjTerm.UniversalProjName)][$($thisNewProjTerm.NetSuiteProjectId)][$($thisNewProjTerm.NetSuiteClientId)]: Creating new Project folders & deflagging Project"}
                    else{Write-Host "`t`tCorresponding Opp, but no DriveItemId [$($correspondingOpp.UniversalOppName)] found for [$($thisNewProjTerm.UniversalProjName)][$($thisNewProjTerm.NetSuiteProjectId)][$($thisNewProjTerm.NetSuiteClientId)]: Creating new Project folders, back-referencing Opp & deflagging Project"}
                    }
                if($thisNewProjTerm.RecreateFolders -eq $true){Write-Host "`t`tCorresponding Opp found, but OppFolder [$($correspondingOpp.UniversalOppName)][$($correspondingOpp.DriveClientId)][$($correspondingOpp.DriveItemId)] does not exist: Creating new Project folders, back-referencing Opp & deflagging Project"}
                try{
                    [array]$newProjFolders = new-oppProjFolders -tokenResponse $tokenResponseSharePointBot -oppProjTermWithClientInfo $thisNewProjTerm
                    if($newProjFolders.Count -ge 1 -and ![string]::IsNullOrWhiteSpace($newProjFolders[0].id)){
                        $thisNewProjTerm.SetCustomProperty("DriveItemId",$newProjFolders[0].id)
                        $thisNewProjTerm.SetCustomProperty("flagForReprocessing",$false)
                        try{
                            $thisNewProjTerm.Context.ExecuteQuery()
                            if(![string]::IsNullOrEmpty($correspondingOpp.id)){ #If there is an Opp, back-date it with the new Project folders 
                                $correspondingOpp.SetCustomProperty("DriveItemId",$newProjFolders[0].id)
                                $correspondingOpp.Context.ExecuteQuery()
                                }
                            $newProjs = $newProjs | ? {$_.DriveItemId -notcontains $thisNewProjTerm.DriveItemId} #If it worked, pop this Proj from the to-do list
                            }
                        catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
                        }
                    }
                catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
                }
            }

        if($newProjs.Count -ge 1){
            Write-Host "`t`t[$($newProjs.Count)] New Project folders failed to create:"
            $newProjs | % {Write-Host "`t`t`t[$($_.UniversalProjName)][$($_.Id)][$($_.NetSuiteProjectId)] for NetSuiteClientId [$($_.NetSuiteClientId)]"}
            #Report this via e-mail too
            }

        #endregion

        #region Existing Projects
        #Does the Term have a DriveItemId?
            #Yes:
                #Has the Name changed?
                    #Yes: Update the DriveItemName, & set flagForReprocessing = $false
                    #No: Set flagForReproccessing = $false
                #Has the Client changed?
                    #Yes: Update the NetSuiteClientId, & set flagForReprocessing = $false
                    #No: Dedupe & set flagForReproccessing = $false
    
        if($deltaSync -eq $true){
            Write-Host "`t[$($existingProjs.Count)] existing Projects need examining to see if anything has changed"
            $tokenResponseSharePointBot = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponseSharePointBot -renewTokenExpiringInSeconds 600 -aadAppCreds $appCredsSharePointBot
            @($existingProjs | Select-Object) | % {
                $thisExistingProj = $_
                try{
                    try{$thisExistingProjDriveItem = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $thisExistingProj.DriveClientId -itemGraphId $thisExistingProj.DriveItemId -returnWhat Item -ErrorAction Stop}        #Try to get the link DriveItem so we can test whether it needs updating
                    catch{$thisExistingProjDriveItem = $null} #This is required to prevent the next itreration getting cross-linked with this DriveItem
                    if([string]::IsNullOrEmpty($thisExistingProjDriveItem.id)){Write-Warning "`t`tOppDriveItem [$($thisExistingProj.UniversalProjName)][$($thisExistingProj.NetSuiteProjectId)] for [$($thisExistingProj.UniversalClientName)][$($thisExistingProj.NetSuiteClientId)] is missing. It might have been assigned to a different Client (which will be fixed on the next Full Reconcile), or it may have been manually moved/deleted."}
                    elseif($(sanitise-forNetsuiteIntegration $thisExistingProjDriveItem.name) -ne $thisExistingProj.UniversalProjNameSanitised){
                        $thisExistingProjDriveItem | Add-Member -MemberType NoteProperty -Name DriveItemName -Value $thisExistingProjDriveItem.name -Force
                        $thisExistingProjDriveItem | Add-Member -MemberType NoteProperty -Name DriveItemId -Value $thisExistingProjDriveItem.id -Force
                        $thisExistingProjDriveItem | Add-Member -MemberType NoteProperty -Name DriveClientName -Value $thisExistingProj.UniversalClientName -Force
                        $thisExistingProjDriveItem | Add-Member -MemberType NoteProperty -Name DriveClientId -Value $thisExistingProj.DriveClientId -Force
                        Write-Host "`tUpdating OppDriveItem Name`t[$($thisExistingProjDriveItem.DriveItemName)][$($thisExistingProjDriveItem.DriveItemId)][$($thisExistingProjDriveItem.DriveClientId)] "
                        Write-Host "`tTo:`t`t`t`t`t`t`t[$($thisExistingProj.UniversalProjName)][$($thisExistingProj.NetSuiteProjectId)][$($thisExistingProj.Id)] for [$($thisExistingProj.UniversalClientName)][$($thisExistingProj.NetSuiteClientId)]"
                        try{$updatedFolder = process-folders -tokenResponse $tokenResponseSharePointBot -standardisedSourceFolder $thisExistingProjDriveItem -renameAs $thisExistingProj.UniversalProjName -Verbose -ErrorAction Stop}
                        catch{Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"}
                        if($updatedFolder.name -eq $thisExistingProj.UniversalProjName){
                            $thisExistingProj.SetCustomProperty("flagForReprocessing",$false)
                            try{$thisExistingProj.Context.ExecuteQuery()}
                            catch{Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"}
                            }
                        }
                    else{
                        #We can't definitely tell what happened to a missing DriveItem without a full reconcile, but if we've already returned the current DriveItem is still valid then it definitely hasn't Client
                        Write-Host "`t`tWeird - it doesn't look like Proj [$($thisExistingProj.UniversalProjName)][$($thisExistingProj.NetSuiteProjectId)][$($thisExistingProj.Id)] for [$($thisExistingProj.UniversalClientName)][$($thisExistingProj.NetSuiteClientId)] needed updating..."
                        }
                    }
                catch{Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"}
                }#>
            }

        if($deltaSync -eq $false){ #Full reconcile
            $existingProjsNameComparison = process-comparison -subsetOfNetObjects $existingProjs -allTermObjects $driveItemsProjFolders -idInCommon DriveItemId -propertyToTest UniversalProjNameSanitised -validate
            [array]$existingTermProjsWithChangedName  = $existingProjsNameComparison["<="]
            [array]$existingDriveProjsWithChangedName = $existingProjsNameComparison["=>"]
                        #Yes: Update the DriveItemName, & set flagForReproccessing = $false
            $tokenResponseSharePointBot = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponseSharePointBot -renewTokenExpiringInSeconds 600 -aadAppCreds $appCredsSharePointBot
            Write-Host "`tProcessing [$($existingDriveProjsWithChangedName.Count)] existing Projects with changed Names"
            for($i=0;$i -lt $existingDriveProjsWithChangedName.Count; $i++){
                Write-Host "`t`t`Updating DriveItemName `t[$($existingDriveProjsWithChangedName[$i].DriveItemName)] for DriveItem [$($existingDriveProjsWithChangedName[$i].DriveItemId)][$($existingDriveProjsWithChangedName[$i].DriveItemUrl)][$($existingDriveProjsWithChangedName[$i].DriveClientName)][$($existingDriveProjsWithChangedName[$i].DriveClientId)]"
                Write-Host "`t`tto:`t`t`t`t`t`t[$($existingTermProjsWithChangedName[$i].UniversalProjName)] from Term [$($existingTermProjsWithChangedName[$i].NetSuiteProjectId)][$($existingTermProjsWithChangedName[$i].Id)][$($existingTermProjsWithChangedName[$i].NetSuiteClientId)]"
                try{
                    $updatedFolder = process-folders -tokenResponse $tokenResponseSharePointBot -standardisedSourceFolder $existingDriveProjsWithChangedName[$i] -renameAs $existingTermProjsWithChangedName[$i].UniversalProjName -ErrorAction Stop
                    $existingTermProjsWithChangedName[$i].SetCustomProperty("flagForReprocessing",$false)
                    try{
                        Write-Verbose "`tTrying to deflag processed Opp [$($existingTermProjsWithChangedName[$i].UniversalProjName)]"
                        $existingTermProjsWithChangedName[$i].Context.ExecuteQuery()
                        }
                    catch{
                        Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                        [array]$duffUpdatedProjs += @($existingDriveProjsWithChangedName[$i],$(get-errorSummary -errorToSummarise $_))
                        }
                    }
                catch{
                    Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                    [array]$duffUpdatedOpps += @($existingDriveProjsWithChangedName[$i],$(get-errorSummary -errorToSummarise $_))
                    }
                }

                        #No: Set flagForReproccessing = $false
            [array]$existingTermProjsWithOriginalName = $existingProjsNameComparison["=="] #We'll updated these once we've finished the deltaClients ones too.

                #Has the Client changed?
            $projExpectedDriveIdComparion = process-comparison -subsetOfNetObjects $existingProjs -allTermObjects $driveItemsProjFolders -idInCommon "DriveItemId" -propertyToTest "DriveClientId" -validate 
            $existingProjTermsWithMismatchedClients = $projExpectedDriveIdComparion["<="]
            $existingProjDrivesWithMismatchedClients = $projExpectedDriveIdComparion["=>"]
            $tokenResponseSharePointBot = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponseSharePointBot -renewTokenExpiringInSeconds 60 -aadAppCreds $appCredsSharePointBot
            Write-Host "`t`tProcessing [$($existingProjDrivesWithMismatchedClients.Count)] existing Projects with changed Clients"
            for($i=0;$i -lt $existingProjDrivesWithMismatchedClients.Count; $i++){
                    #Yes: Move the DriveItems, Update the DriveItemId, & set flagForReproccessing = $false
                Write-Host "`t`tMoving existing Proj folder from`t[$($existingProjDrivesWithMismatchedClients[$i].DriveClientName)][$($existingProjDrivesWithMismatchedClients[$i].DriveClientId)] (ProjFolder is [$($existingProjDrivesWithMismatchedClients[$i].UniversalProjName)][$($existingProjTermsWithMismatchedClients[$i].NetSuiteProjId)][$($existingProjDrivesWithMismatchedClients[$i].DriveItemUrl)][$($existingProjDrivesWithMismatchedClients[$i].DriveItemId)]"
                Write-Host "`t`tto:`t`t`t`t`t`t`t`t`t[$($existingProjTermsWithMismatchedClients[$i].UniversalClientName)][$($existingProjTermsWithMismatchedClients[$i].DriveClientId)][$($existingProjTermsWithMismatchedClients[$i].NetSuiteClientId)] (ProjTerm is [$($existingProjTermsWithMismatchedClients[$i].UniversalProjName)][$($existingProjTermsWithMismatchedClients[$i].NetSuiteProjectId)][$($existingProjTermsWithMismatchedClients[$i].DriveItemId)]"
                try{
                    $newDestinationFolder = add-graphFolderToDrive -graphDriveId $existingProjTermsWithMismatchedClients[$i].DriveClientId -folderName $existingProjTermsWithMismatchedClients[$i].UniversalProjName -tokenResponse $tokenResponseSharePointBot -conflictResolution Fail
                    $movedFolders = process-folders -tokenResponse $tokenResponseSharePointBot -standardisedSourceFolder $existingProjDrivesWithMismatchedClients[$i] -mergeInto $newDestinationFolder -ErrorAction Continue
                    if($movedFolders[0].parentReference.driveId -eq $existingProjDrivesWithMismatchedClients[$i].DriveClientId){
                        Write-Host "`t`t`tFailed to move these [$($movedFolders.count)] folders:"
                        @($movedFolders | Select-Object) | % {Write-Host "`t`t`t[$($_.name)][$($_.weburl)]"}
                        }
                    else{
                        $existingProjTermsWithMismatchedClients[$i].SetCustomProperty("DriveItemId",$newDestinationFolder.id)
                        $existingProjTermsWithMismatchedClients[$i].SetCustomProperty("flagForReprocessing",$false)
                        try{$existingProjTermsWithMismatchedClients[$i].Context.ExecuteQuery()}
                        catch{Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"}
                        try{
                            $result = process-folders -tokenResponse $tokenResponseSharePointBot -standardisedSourceFolder $existingProjDrivesWithMismatchedClients[$i] -confirmDeleteEmptyFolders #Finally, try to delete any empty folder
                            if($result -ne $true){Write-Warning "Failed to delete (hopefully) empty folder [$($existingProjDrivesWithMismatchedClients[$i].DriveItemUrl)]"}
                            }
                        catch{Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"}
                        }
                    }
                catch{
                    Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                    }
                }

                    #No: Dedupe & set flagForReprocessing = $false
            [array]$existingTermProjsWithOriginalClient = $projExpectedDriveIdComparion["=="]
            $dedupedProjectsWithOriginalClientAndOriginalName = [System.Collections.Generic.Hashset[Microsoft.SharePoint.Client.Taxonomy.TermSetItem]] ($existingTermProjsWithOriginalName + $existingTermProjsWithOriginalClient)
            [array]$dedupedProjectsStillFlaggedForProcessing = $dedupedProjectsWithOriginalClientAndOriginalName | ? {$_.CustomProperties.flagForReprocessing -ne $false}
            if($dedupedProjectsStillFlaggedForProcessing.Count -gt 0){
                Write-Host "`t[$($dedupedProjectsStillFlaggedForProcessing.Count)] Project Terms were flagged for reprocessing, but they don't seem to have any changes. This isn't specifically a _problem_, but it's an indication that reconcile-netSuiteToTermStore() is incorrectly flagging Projects as requiring processing when they don't"
                $dedupedProjectsStillFlaggedForProcessing | % {
                    $thisDedupedProjectsStillFlaggedForProcessing = $_
                    Write-Host "`t`t[$($thisDedupedProjectsStillFlaggedForProcessing.UniversalProjName)][$($thisDedupedProjectsStillFlaggedForProcessing.NetSuiteProjectId)][$($thisDedupedProjectsStillFlaggedForProcessing.id)] for [$($thisDedupedProjectsStillFlaggedForProcessing.UniversalClientName)][$($thisDedupedProjectsStillFlaggedForProcessing.NetSuiteClientId)]"
                    $thisDedupedProjectsStillFlaggedForProcessing.SetCustomProperty("flagForReprocessing",$false)
                    try{$thisDedupedProjectsStillFlaggedForProcessing.Context.ExecuteQuery()}
                    catch{Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"}
                    }
                }

            }


        #endregion

    #endregion
    #endregion
    }

Write-Host "Processing complete at [$(get-date -Format s)] in [$($timeForFullCycle.TotalMinutes)] minutes ([$($timeForFullCycle.TotalSeconds)] seconds)"

Stop-Transcript