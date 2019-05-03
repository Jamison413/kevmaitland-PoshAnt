function add-spoLibrarySubfolders($pnpList, $arrayOfSubfolderNames, $recreateIfNotEmpty, $spoCredentials, $verboseLogging){
    #$arrayOfSubfolderNames - I think these are supposed to be serverRelativeUrls
    if($verboseLogging){Write-Host -ForegroundColor Magenta "add-spoLibrarySubfolders($($pnpList.Title), $($arrayOfSubfolderNames -join ", "), `$recreateIfNotEmpty=$recreateIfNotEmpty"}
    if($(Get-PnPConnection).Url -notmatch $pnpList.ParentWebUrl){
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Connected to wrong site - connecting to $($pnpList.RootFolder.Context.Url)"}
        Connect-PnPOnline –Url $($pnpList.RootFolder.Context.Url) –Credentials $spoCredentials
        }
    #[array]$formattedArrayOfSubfolderNames = $arrayOfSubfolderNames | % {format-asServerRelativeUrl -serverRelativeUrl $pnpList.RootFolder.ServerRelativeUrl -stringToFormat $_}
    #Get the site-relative Url by comparing the List's ServerRelativeUrl with the Site's ServerRelativeUrl and eliminating any overlap e.g. "/clients/MyClient" becomes "/MyClient"
    $checkForServerSiteOverlap = [regex]::Match($pnpList.RootFolder.ServerRelativeUrl,"^$($pnpList.RootFolder.Context.Web.ServerRelativeUrl)(.+)*")
    if($checkForServerSiteOverlap.Success){$siteRelativeUrlPrefix = $checkForServerSiteOverlap.Groups[$checkForServerSiteOverlap.Groups.Count-1].Value}
    else{$siteRelativeUrlPrefix = $pnpList.RootFolder.Context.Web.ServerRelativeUrl}
    
    #[array]$formattedArrayOfSiteRelativeSubfolderNames = $arrayOfSubfolderNames | % {$siteRelativeUrlPrefix+$_.Replace($pnpList.RootFolder.ServerRelativeUrl,"")}
    #Changed [KM] 2019-03-14 As Client DocLibs weren't beign created properly (missing the trailing / on the site relative path: /JUUL_Kimble automatically creates Project folders)
    [array]$formattedArrayOfSiteRelativeSubfolderNames = $arrayOfSubfolderNames | % {$($siteRelativeUrlPrefix+"/"+$_.Replace($pnpList.RootFolder.ServerRelativeUrl,"")).Replace("//","/")}

    try{
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "get-spoFolder -pnpList $($pnpList.Title) -folderServerRelativeUrl $($formattedArrayOfSubfolderNames[$formattedArrayOfSubfolderNames.Length-1])"}
        #$hasItems = get-spoFolder -pnpList $pnpList -folderServerRelativeUrl $($formattedArrayOfSubfolderNames[$formattedArrayOfSubfolderNames.Length-1]) -adminCreds $adminCreds -verboseLogging $verboseLogging
        #$hasItems = Get-PnPListItem -List $pnpList -Query "<View><RowLimit>5</RowLimit></View>" #This RowLimit doesn't work at the moment, but hopefully it'll get fixed in the future and this'll be efficient https://github.com/SharePoint/PnP-PowerShell/issues/879
        #$hasItems = Get-PnPListItem -List $pnpList -Query "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='Text'>DummyOp5 (E003941)</Value></Eq></Where></Query></View>" 
        #$hasItems = Get-PnPListItem -List $pnpList -Query "<View><Query><Where><Eq><FieldRef Name='FileRef'/><Value Type='Text'>/clients/DummyCo Ltd/DummyOp5 (E003941)</Value></Eq></Where></Query></View>" 
        #$hasItems = Get-PnPListItem -List $pnpList -Query "<View><Query><Where><Eq><FieldRef Name='FileRef'/><Value Type='Text'>/clients/DummyCo Ltd/DummyOp5 (E003941)/Analysis</Value></Eq></Where></Query></View>" 
        #$hasItems = Get-PnPListItem -List $pnpList #-Query "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='Text'>$($arrayOfSubfolderNames[0])</Value></Eq></Where></Query></View>" 
        #$hasItems = $hasItems | ? {$_.FieldValues.FileRef -eq "$($arrayOfSubfolderNames[$arrayOfSubfolderNames.Length-1])"}
        $hasItems = Get-PnPFolder -Url $formattedArrayOfSiteRelativeSubfolderNames[$formattedArrayOfSiteRelativeSubfolderNames.Count-1] -ErrorAction Stop -Includes ListItemAllFields
        }
    catch{
        #Meh.
        }
    if(!$hasItems -or $recreateIfNotEmpty){
        if($verboseLogging){
            if(!$hasItems){Write-Host -ForegroundColor DarkMagenta "$($pnpList.RootFolder.ServerRelativeUrl) has no conflicting item - creating subfolder/s"}
            else{Write-Host -ForegroundColor DarkMagenta "$($pnpList.RootFolder.ServerRelativeUrl) has items, but override set - creating subfolders"}
            }
        <#$formattedArrayOfSubfolderNames | % {
            #We have to search for these using ServerRelativeUrls, but create them using LibraryRelativeUrls. Oh no we fucking don't. 
            $libraryRelativePath = $_.Replace($pnpList.RootFolder.ServerRelativeUrl,"")
            if($libraryRelativePath.Substring(0,1) -eq "/"){$libraryRelativePath = $libraryRelativePath.Substring(1,$libraryRelativePath.Length-1)} #Trim any leading "/"
            if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Add-PnPDocumentSet -List $($pnpList.Id) [$($pnpList.Title)] -Name [$libraryRelativePath] -ContentType ""Document Set"""}
            $newFolderUrl = Add-PnPDocumentSet -List $pnpList.Id -Name $libraryRelativePath -ContentType "Document Set"
            }
        $newFolder = get-spoFolder -pnpList $pnpList -folderServerRelativeUrl $newFolderUrl.Replace("https://anthesisllc.sharepoint.com","") -adminCreds $spoCredentials -verboseLogging $verboseLogging #>
        $formattedArrayOfSiteRelativeSubfolderNames | % {
            $folderName = Split-Path $_ -Leaf
            $folderPath = $_.Substring(0,$_.Length-$folderName.Length)
            if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Add-PnPFolder -Folder [$($folderPath)] -Name [$($folderName)]"}
            Add-PnPFolder -Folder $folderPath -Name $folderName            
            }

        $newFolder = Get-PnPFolder $formattedArrayOfSiteRelativeSubfolderNames[$formattedArrayOfSiteRelativeSubfolderNames.Count-1] -Includes ListItemAllFields
        $newFolder #Return last folder created (we have to do this separately as Add-PnPDocumentSet only returns the Absolute URL)
        }
    else{
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "$($pnpList.RootFolder.ServerRelativeUrl) has items and no override set - *NOT* recreating subfolders"}
        $hasItems
        }
    }
function add-spoTermToStore($termGroup,$termSet,$term,$kimbleId,$verboseLogging){
    $cleanTerm = sanitise-forTermStore $term
    if($verboseLogging){Write-Host -ForegroundColor Magenta "add-spoTermToStore($termGroup,$termSet,$cleanTerm,$kimbleId)"}
    try{
        $pnpTermGroup = Get-PnPTermGroup $termGroup 
        $pnpTermSet = Get-PnPTermSet -TermGroup $pnpTermGroup -Identity $termSet
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Get-PnPTerm -TermGroup $($pnpTermGroup.Name) -TermSet $($pnpTermSet.Name) -Identity $cleanTerm -ErrorAction Stop"}
        #$pnpTerm = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Identity $cleanTerm -ErrorAction Stop #Weirdly, Get-PnPTerm throws a non-terminating exception if the Term isn't found. We want an exception, so that catch{} returns $null value
        #2019-03-14 [KM] Retrieving all Terms now as it's bizarrely faster than retrieving an individual term and we're hitting a 30 second timeout.
        #$alreadyInStore = Get-PnPTaxonomyItem -TermPath "$termGroup|$termSet|$term"
        $allTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet
        $pnpTerm = $allTerms | ? {$_.Name -eq $cleanTerm}
        }
    catch{
        #Meh.
        }
    if($pnpTerm){
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "$termGroup | $termSet | $cleanTerm already exists - not creating duplicate"}
        if(![string]::IsNullOrEmpty($kimbleId)){#If we've got a KimbleId, try to add it as there's loads missing
            $pnpTerm.SetCustomProperty("KimbleId",$kimbleId)
            $pnpTerm.Context.ExecuteQuery()
            }
        $pnpTerm
        }
    else{
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "$termGroup | $termSet | $cleanTerm does not exist - creating new term"}
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "New-PnPTerm -TermGroup $($pnpTermGroup.Id) -TermSet $($pnpTermSet.Id) -Name $cleanTerm -Lcid 1033"}
        if(![string]::IsNullOrEmpty($kimbleId)){$customProps = @{"KimbleId"=$kimbleId}}
        $newPnpTerm = New-PnPTerm -TermGroup $pnpTermGroup.Id -TermSet $pnpTermSet.Id -Name $cleanTerm -Lcid 1033 -CustomProperties $customProps
        $newPnpTerm
        }
    }
function cache-spoKimbleAccountsList($pnpList, $kimbleListCachePathAndFileName, $fullLogPathAndName, $errorLogPathAndName, $verboseLogging){
    $listCacheFile = Get-Item $kimbleListCachePathAndFileName
    if((get-date $pnpList.LastItemModifiedDate).AddMinutes(-5) -gt $listCacheFile.LastWriteTimeUtc){#This is bodged so we don't miss any new List added during the time it takes to actually download the full Account list
        try{
            log-action -myMessage "[$($pnpList.Title)] needs recaching - downloading full list" -logFile $fullLogPathAndName 
            $duration = Measure-Command {$spList = get-spoKimbleAccountListItems -pnpList $pnpList -spoCredentials $adminCreds }
            if($spList){
                log-result -myMessage "SUCCESS: $($spList.Count) [$($pnpList.Title)] records retrieved [$($duration.TotalSeconds) secs]!" -logFile $fullLogPathAndName
                if(!(Test-Path -Path $cacheFilePath)){New-Item -Path $cacheFilePath -ItemType Directory}
                $spList | Export-Csv -Path $kimbleListCachePathAndFileName -Force -NoTypeInformation -Encoding UTF8
                }
            else{log-result -myMessage "FAILURE: [$($pnpList.Title)] items could not be retrieved" -logFile $fullLogPathAndName}
            }
        catch{log-error -myError $_ -myFriendlyMessage "Could not retrieve [$($pnpList.Title)] items to recache the local copy" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -doNotLogToEmail $true}
        }
    else{log-result -myMessage "SUCCESS: [$($pnpList.Title)] Cache is up-to-date and does not require refreshing" -logFile $fullLogPathAndName}
    $listCache = Import-Csv $kimbleListCachePathAndFileName
    $listCache
    }
function copy-spoFile($fromList,$from,$to,$spoCredentials){
    if($verboseLogging){Write-Host -ForegroundColor Magenta "copy-spoFile($fromList,$from,$to"}
    if($fromList.Substring(0,1) -ne "/"){$fromList = "/"+$fromList}
    if($(Split-Path $from -Leaf) -eq $(Split-Path $to -Leaf)){$to = $to.SubString(0,$($to.Length - $(split-path $to -leaf).Length) -1)} #Specififying a file name is broken for (presumably) Sites with large numbers of Libraries/Files
    if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Copy-PnPFile -SourceUrl $from -TargetUrl $to -Force (but not -OverwriteIfAlreadyExists)"}
    Copy-PnPFile -SourceUrl $from -TargetUrl $to -Force
    Get-PnPFile -Url "$to$(Split-Path $from -Leaf)"
    }
function format-asServerRelativeUrl($serverRelativeUrl,$stringToFormat){
    $formattedString = $stringToFormat
    if([string]::IsNullOrWhiteSpace($formattedString)){$formattedString = "/"}
    if($formattedString -notmatch $serverRelativeUrl){
        if($formattedString.Substring(0,1) -ne "/"){
            $formattedString = "/"+$formattedString
            }
        $formattedString = $($serverRelativeUrl+$formattedString).Replace("//","/")
        }
    $formattedString
    }
function format-asServerRelativeUrls($serverRelativeUrl,$arrayOfStringToFormat){
    $arrayOfStringsToFormat | % {
        $thisThing = $_ 
        if([string]::IsNullOrWhiteSpace($thisThing)){$thisThing = "/"}
        if($thisThing -notmatch $serverRelativeUrl){
            if($thisThing.Substring(0,1) -ne "/"){
                $thisThing = "/"+$thisThing
                }
            $thisThing = $($serverRelativeUrl+$thisThing).Replace("//","/")
            }
        [array]$formattedThings+=$thisThing
        }
    #if($formattedThings.Count -eq 1){$formattedThings[0]} #If $thingsToFormat was just a single string, return a string
    #else{$formattedThings} #If $thingsToFormat was an array, return an array
    $formattedArrayOfClientSubfolders #Change of plan - always return an array
    }
function get-spoClientLibrary($clientName, $clientLibraryGuid, $adminCreds, $verboseLogging){
    #Check that the Client Library is retrievable
    try{
        if($verboseLogging){Write-Host -ForegroundColor Magenta "get-spoClientLibrary($clientName, $clientLibraryGuid)"}
        if(![string]::IsNullOrWhiteSpace($clientLibraryGuid)){
            if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Found LibraryGUID for Client - trying that!"}
            try{
                $thisClientLibrary = Get-PnPList -Identity $($clientLibraryGuid) 
                if($verboseLogging){if(!$thisClientLibrary){Write-Host -ForegroundColor DarkMagenta "`tDidn't work :("}}
                }
            catch{<#Meh.#>}
            }
        if(!$thisClientLibrary){
            $sanitisedClientName = $(sanitise-forPnpSharePoint $clientName)
            if($clientName.SubString($clientName.Length-1,1) -eq "."){$sanitisedClientName+="."} #Trailing fullstops /are/ allowed in this context
            if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Trying to retrieve Library by Client Name: Get-PnPList -Identity [$sanitisedClientName]"}
            try{$thisClientLibrary = Get-PnPList -Identity $sanitisedClientName}
            catch{<#Meh.#>}
            }
        $thisClientLibrary
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving Client Library in get-spoClientLibrary" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName}
    }
function get-spoDocumentLibrary($docLibName, $docLibGuid, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Magenta "get-spoDocumentLibrary($docLibName, $docLibGuid)"}
    #Check that the Client Library is retrievable
    try{
        if(![string]::IsNullOrWhiteSpace($docLibGuid)){
            if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Found LibraryGUID for DocLib - trying that!"}
            try{
                $thisDocumentLibrary = Get-PnPList -Identity $($docLibGuid) 
                if($verboseLogging){if(!$thisDocumentLibrary){Write-Host -ForegroundColor DarkMagenta "`tDidn't work :("}}
                }
            catch{<#Meh.#>}
            }
        if(!$thisDocumentLibrary -and ![string]::IsNullOrWhiteSpace($docLibName)){
            $sanitisedDocLibName = $(sanitise-forPnpSharePoint $docLibName)
            if($docLibName.SubString($docLibName.Length-1,1) -eq "."){$sanitisedDocLibName+="."} #Trailing fullstops /are/ allowed in this context
            if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Trying to retrieve Document Library by Name: Get-PnPList -Identity [$sanitisedDocLibName]"}
            try{$thisDocumentLibrary = Get-PnPList -Identity $sanitisedDocLibName}
            catch{<#Meh.#>}
            }
        $thisDocumentLibrary
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving Document Library in get-spoDocumentLibrary" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName}
    }
function get-spoFolder($pnpList, $folderServerRelativeUrl, $folderGuid, $adminCreds, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Magenta "get-spoFolder($($pnpList.Title), $folderServerRelativeUrl)"}
    #$hasItems = Get-PnPListItem -List $pnpList -Query "<View><RowLimit>5</RowLimit></View>" #This RowLimit doesn't work at the moment, but hopefully it'll get fixed in the future and this'll be efficient https://github.com/SharePoint/PnP-PowerShell/issues/879
    #$hasItems = Get-PnPListItem -List $pnpList -Query "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='Text'>DummyOp5 (E003941)</Value></Eq></Where></Query></View>" 
    #$hasItems = Get-PnPListItem -List $pnpList -Query "<View><Query><Where><Eq><FieldRef Name='FileRef'/><Value Type='Text'>/clients/DummyCo Ltd/DummyOp5 (E003941)</Value></Eq></Where></Query></View>" 
    #$hasItems = Get-PnPListItem -List $pnpList -Query "<View><Query><Where><Eq><FieldRef Name='FileRef'/><Value Type='Text'>/clients/DummyCo Ltd/DummyOp5 (E003941)/Analysis</Value></Eq></Where></Query></View>" 
    if(![string]::IsNullOrWhiteSpace($folderGuid)){
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Get-PnPListItem -List $($pnpList.Title) -UniqueId $folderGuid"}
        $pnpListItem = Get-PnPListItem -List $pnpList -UniqueId $folderGuid
        }
    if($pnpListItem.Count -eq 0){
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Get-PnPFolder -Url $folderServerRelativeUrl -Includes UniqueId,ServerRelativeUrl,ServerRelativePath,ListItemAllFields"}# -Query <Where><Eq><FieldRef Name='FSObjType' /><Value Type='int'>1</Value></Eq></Where>"}
        try{
            $pnpFolder = Get-PnPFolder -Url $folderServerRelativeUrl -Includes UniqueId,ServerRelativeUrl,ServerRelativePath,ListItemAllFields -ErrorAction Stop
            }
        catch{
            #Weirdly, Get-PnPFolder throws a non-terminating Exception if it can't find the folder. We don't want that, we either want it to return $null, or Stop so we can return $null from the catch{} block like this
            }
        if($pnpFolder.ListItemAllFields.Id){
            if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Get-PnPListItem -List $($pnpList.Title) -Id $($pnpFolder.ListItemAllFields.Id)"}# -Query <Where><Eq><FieldRef Name='FSObjType' /><Value Type='int'>1</Value></Eq></Where>"}
            $pnpListItem = Get-PnPListItem -List $pnpList -Id $($pnpFolder.ListItemAllFields.Id)
            }
        #$test = Get-PnPListItem -List $pnpList -Query "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='ID'/><Value Type='Integer'>$($hasItems.ListItemAllFields.Id)</Value></Eq></Where></Query></View>"
        #$hasItems2 = Get-PnPListItem -List $pnpList #-Query "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='FSObjType' /><Value Type='int'>1</Value></Eq></Where></Query></View>" #FileRef CAML shit doesn't work for >5000 list items
        #$hasItems3 = $hasItems2 | ? {$_.FieldValues.FileRef -eq $folderServerRelativeUrl} 
        }
    if($verboseLogging){
        if($pnpListItem){Write-Host -ForegroundColor DarkMagenta "Found $($pnpListItem.Count) items: $($pnpListItem.FieldValues.FileRef)"}# -Query <Where><Eq><FieldRef Name='FSObjType' /><Value Type='int'>1</Value></Eq></Where>"}
        else{Write-Host -ForegroundColor DarkMagenta "No item found"}
        }
    $pnpListItem
    }
function get-spoProjectFolder($pnpList, $kimbleEngagementCodeToLookFor, $folderGuid, $adminCreds, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Magenta "get-spoProjectFolder($($pnpList.Title), $kimbleEngagementCodeToLookFor)"}
    if($folderGuid){
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "get-PnPListItem -list [$($pnpList.Title)] -UniqueId [$folderGuid]"}
        $pnpListItem = Get-PnPListItem -List $pnpList -UniqueId $folderGuid
        }
    else{
        if(!$pnpListItem -and ![string]::IsNullOrWhiteSpace($kimbleEngagementCodeToLookFor)){}
        #$pnpQuery = "<View><Query><Where><Contains><FieldRef Name='Title'/><Value Type='Text'>$kimbleEngagementCodeToLookFor</Value></Eq></Where></Query></View>"
        $pnpQuery = "<View><Query><Where><Contains><FieldRef Name='FileLeafRef'/><Value Type='Text'>$kimbleEngagementCodeToLookFor</Value></Eq></Where></Query></View>" #Changed to FileLeafRef because Title property is not always populated
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "get-PnPListItem -list [$($pnpList.Title)] -Query [$pnpQuery]"}
        $pnpListItem = Get-PnPListItem -List $pnpList -Query $pnpQuery
        }
    if($verboseLogging){
        if($pnpListItem){Write-Host -ForegroundColor DarkMagenta "Found $($pnpListItem.Count) items: $($pnpListItem.FieldValues.FileRef)"}# -Query <Where><Eq><FieldRef Name='FSObjType' /><Value Type='int'>1</Value></Eq></Where>"}
        else{Write-Host -ForegroundColor DarkMagenta "No item found"}
        }

    $pnpListItem
    }
function get-spoSupplierLibrary($SupplierName, $SupplierLibraryGuid, $adminCreds, $verboseLogging){
    #Check that the Supplier Library is retrievable
    try{
        if($verboseLogging){Write-Host -ForegroundColor Magenta "get-spoSupplierLibrary($SupplierName, $SupplierLibraryGuid)"}
        if(![string]::IsNullOrWhiteSpace($SupplierLibraryGuid)){
            if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Found LibraryGUID for Supplier - trying that!"}
            try{
                $thisSupplierLibrary = Get-PnPList -Identity $($SupplierLibraryGuid) 
                if($verboseLogging){if(!$thisSupplierLibrary){Write-Host -ForegroundColor DarkMagenta "`tDidn't work :("}}
                }
            catch{<#Meh.#>}
            }
        if(!$thisSupplierLibrary){
            $sanitisedSupplierName = $(sanitise-forPnpSharePoint $SupplierName)
            if($SupplierName.SubString($SupplierName.Length-1,1) -eq "."){$sanitisedSupplierName+="."} #Trailing fullstops /are/ allowed in this context
            if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Trying to retrieve Library by Supplier Name: Get-PnPList -Identity [$sanitisedSupplierName]"}
            try{$thisSupplierLibrary = Get-PnPList -Identity $sanitisedSupplierName}
            catch{<#Meh.#>}
            }
        $thisSupplierLibrary
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving Supplier Library in get-spoSupplierLibrary" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName}
    }
function get-allSpoListItemsWithUniquePermissions($pnpList,$adminCredentials, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Magenta "get-allSpoListItemsWithUniquePermissions($($pnpList.Title))"}
    try{Get-PnPConnection | Out-Null}
    catch{
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "No Connect-PnPOnline connection available. Creating new Connect-PnpOnline to [$($pnpList.Context.Url)]"}
        Connect-PnPOnline -Url $pnpList.Context.Url -Credentials $adminCredentials
        }
        $tempConnection = Get-PnPConnection
    if((Get-PnPConnection).Url -eq $pnpList.Context.Url){$tempConnection = Get-PnPConnection}
    else{
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Current Connect-PnPOnline connection [$((Get-PnPConnection).Url)] does not match target site. Creating temporary Connect-PnpOnline to [$($pnpList.Context.Url)]"}
        Write-Warning "Current Connect-PnPOnline connection [$((Get-PnPConnection).Url)] does not match target site. Creating new Connect-PnpOnline to [$($pnpList.Context.Url)]"
        $oldPnPConnection = (Get-PnPConnection).Url
        $tempConnection = Connect-PnPOnline -Url $pnpList.Context.Url -ReturnConnection -Credentials $adminCredentials
        }
    #Best to enumrate everything and test
    if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Get-PnPListItem -List $($pnpList.Title) -Query '<View><Query><Where><IsNotNull><FieldRef Name='SharedWithDetails' /></IsNotNull></Where></Query></View>'"}
    try{
        $results = Get-PnPListItem -List $pnpList.Id -Query "<View><Query><Where><IsNotNull><FieldRef Name='SharedWithDetails' /></IsNotNull></Where></Query></View>" -ErrorAction stop
        $results | ? {$_.FieldValues["SharedWithUsers"]} #Remove any results that have been shared (creating the SharedWithDetails field), but then unshared (removing all entries from the SharedWithUsers field)
        }
    catch{
        $false
        if($_.Exception -eq "One or more field types are not installed properly. Go to the list settings page to delete these fields."){write-warning "Error in get-allSpoListItemsWithUniquePermissions searching for ListItems with SharedWithDetails - this *probably* just mean that there were none"}
        else{Write-Error $_}
        }
    if($oldPnPConnection){
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Disconecting temporary Connect-PnpOnline to [$($pnpList.Context.Url)] and reconnecting to [$oldPnPConnection]"}
        Connect-PnPOnline -ur $oldPnPConnection -Credentials $adminCredentials
        #Disconnect-PnPOnline -Connection $tempConnection
        }
    }
function get-allSpoListsWithItemsWithUniquePermissions($siteAbsoluteUrl,$adminCredentials, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Magenta "get-allSpoListsWithItemsWithUniquePermissions($siteAbsoluteUrl)"}
    $siteServerRelativeUrl = ([System.Uri]$siteAbsoluteUrl).LocalPath
    try{Get-PnPConnection | Out-Null}
    catch{Connect-PnPOnline -Url $siteAbsoluteUrl -Credentials $adminCredentials}
    if(([System.Uri](Get-PnPConnection).Url).LocalPath -eq $siteServerRelativeUrl){$tempConnection = Get-PnPConnection}
    else{
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Current Connect-PNPOnline connection [$((Get-PnPConnection).Url)] does not match target site. Creating temporary Connect-PnpOnline to [$siteAbsoluteUrl]"}
        $tempConnection = Connect-PnPOnline -Url $siteAbsoluteUrl -ReturnConnection -Credentials $adminCredentials
        $tempConnectionInitiated = $true
        }
    #Setting unique permissions on a list item seems to add a flag to the List XML too. Presumably this is how /_layouts/15/uniqperm.aspx works so quickly?
    if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Get-PnPList -Includes SchemaXml | ? {$_.SchemaXML -match 'SharedWithDetails'"}
    Get-PnPList -Includes SchemaXml -Connection $thisConnection | ? {$_.SchemaXML -match "SharedWithDetails"}
    if($tempConnectionInitiated){Disconnect-PnPOnline -Connection $tempConnection}
    }
function get-spoKimbleAccountListItems($pnpList,$spoCredentials, $verboseLogging){
    if($pnpList.Title -match "Clients"){get-spoKimbleClientListItems -spoCredentials $spoCredentials -verboseLogging $verboseLogging}
    elseif($pnpList.Title -match "Projects"){get-spoKimbleProjectListItems -spoCredentials $spoCredentials -verboseLogging $verboseLogging}
    elseif($pnpList.Title -match "Suppliers"){get-spoKimbleSupplierListItems -spoCredentials $spoCredentials -verboseLogging $verboseLogging}
    else{}
    }
function get-spoKimbleClientListItems($spoCredentials, $verboseLogging){
    if($(Get-PnPConnection).Url -ne "https://anthesisllc.sharepoint.com/clients"){
        Connect-PnPOnline –Url $("https://anthesisllc.sharepoint.com/clients") –Credentials $spoCredentials
        }
    if($verboseLogging){Write-Host -ForegroundColor Magenta 'Get-PnPListItem -List "Kimble Clients" -PageSize 5000 -Fields "Title","GUID","KimbleId","ClientDescription","IsDirty","IsDeleted","Modified","LastModifiedDate","PreviousName","PreviousDescription","Id","LibraryGUID","IsOrphaned","isMisclassified"'}
    $clientListItems = Get-PnPListItem -List "Kimble Clients" -PageSize 5000 -Fields "Title","GUID","KimbleId","ClientDescription","IsDirty","IsDeleted","Modified","LastModifiedDate","PreviousName","PreviousDescription","Id","LibraryGUID","IsOrphaned","isMisclassified"
    $clientListItems.FieldValues | %{
        $thisClient = $_
        [array]$allSpoClients += New-Object psobject -Property $([ordered]@{"Id"=$thisClient["KimbleId"];"Name"=$thisClient["Title"];"GUID"=$thisClient["GUID"];"SPListItemID"=$thisClient["ID"];"IsDirty"=$thisClient["IsDirty"];"IsDeleted"=$thisClient["IsDeleted"];"LastModifiedDate"=$thisClient["LastModifiedDate"];"PreviousName"=$thisClient["PreviousName"];"ClientDescription"=$(sanitise-stripHtml $thisClient["ClientDescription"]);"PreviousDescription"=$thisClient["PreviousDescription"];"LibraryGUID"=$thisClient["LibraryGUID"];"IsOrphaned"=$thisClient["IsOrphaned"];"isMisclassified"=$thisClient["isMisclassified"]})
        }
    $allSpoClients
    }
function get-spoKimbleProjectListItems($camlQuery, $spoCredentials, $verboseLogging){
    if($(Get-PnPConnection).Url -ne "https://anthesisllc.sharepoint.com/clients"){
        Connect-PnPOnline –Url $("https://anthesisllc.sharepoint.com/clients") –Credentials $spoCredentials
        }
    if($verboseLogging){Write-Host -ForegroundColor Magenta "Get-PnPListItem -List ""Kimble Projects"" -Query $camlQuery -PageSize 5000 "}
    #-Fields "Title","GUID","KimbleId","IsDirty","IsDeleted","Modified","LastModifiedDate","PreviousName","KimbleClientId","PreviousKimbleClientId","Id","DoNotProcess","BusinessUnit","FolderGUID"'}
    $projectListItems = Get-PnPListItem -List "Kimble Projects" -Query $camlQuery -PageSize 5000 
    #-Fields "Title","GUID","KimbleId","IsDirty","IsDeleted","Modified","LastModifiedDate","PreviousName","KimbleClientId","PreviousKimbleClientId","Id","DoNotProcess","BusinessUnit","FolderGUID" 
    $projectListItems.FieldValues | %{
        $thisProject = $_
        [array]$allSpoProjects += New-Object psobject -Property $([ordered]@{"Id"=$thisProject["KimbleId"];"Name"=$thisProject["Title"];"GUID"=$thisProject["GUID"];"SPListItemID"=$thisProject["ID"];"IsDirty"=$thisProject["IsDirty"];"IsDeleted"=$thisProject["IsDeleted"];"LastModifiedDate"=$thisProject["LastModifiedDate"];"PreviousName"=$thisProject["PreviousName"];"KimbleClientId"=$thisProject["KimbleClientId"];"PreviousKimbleClientId"=$thisProject["PreviousKimbleClientId"];"DoNotProcess"=$thisProject["DoNotProcess"];"BusinessUnit"=$thisProject["BusinessUnit"];"FolderGUID"=$thisProject["FolderGUID"]})
        }
    $allSpoProjects
    }
function get-spoKimbleSupplierListItems($spoCredentials, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Magenta 'Get-PnPListItem -List "Kimble Suppliers" -PageSize 5000 -Fields "Title","GUID","KimbleId","SupplierDescription","IsDirty","IsDeleted","Modified","LastModifiedDate","PreviousName","PreviousDescription","Id","LibraryGUID"'}
    $SupplierListItems = Get-PnPListItem -List "Kimble Suppliers" -PageSize 5000 -Fields "Title","GUID","KimbleId","SupplierDescription","IsDirty","IsDeleted","Modified","LastModifiedDate","PreviousName","PreviousDescription","Id","LibraryGUID"
    $SupplierListItems.FieldValues | %{
        $thisSupplier = $_
        [array]$allSpoSuppliers += New-Object psobject -Property $([ordered]@{"Id"=$thisSupplier["KimbleId"];"Name"=$thisSupplier["Title"];"GUID"=$thisSupplier["GUID"];"SPListItemID"=$thisSupplier["ID"];"IsDirty"=$thisSupplier["IsDirty"];"IsDeleted"=$thisSupplier["IsDeleted"];"LastModifiedDate"=$thisSupplier["LastModifiedDate"];"PreviousName"=$thisSupplier["PreviousName"];"SupplierDescription"=$(sanitise-stripHtml $thisSupplier["SupplierDescription"]);"PreviousDescription"=$thisSupplier["PreviousDescription"];"LibraryGUID"=$thisSupplier["LibraryGUID"]})
        }
    $allSpoSuppliers
    }
function new-spoClientLibrary($clientName, $clientDescription, $spoCredentials, $verboseLogging){
    #
    # Deprecated - use new-spoDocumentLibrary
    #
    #
    if($verboseLogging){Write-Host -ForegroundColor Magenta "new-spoClientLibrary($clientName, $clientDescription)"}
    if($(Get-PnPConnection).Url -ne "https://anthesisllc.sharepoint.com/clients"){
        Connect-PnPOnline –Url $("https://anthesisllc.sharepoint.com/clients") –Credentials $spoCredentials
        }
    try{
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Get-PnPList -Identity $(sanitise-forSql $clientName)"}
        $clientLibrary = get-spoClientLibrary -clientName $clientName
        }
    catch{<#Meh.#>}
    if($clientLibrary){if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Existing Library for $clientName found - not creating another"}}
    else{
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Existing Library for $clientName not found - creating a new one"}
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "New-PnPList -Title $(sanitise-forSql $clientName) -Template DocumentLibrary"}
        $clientLibrary = New-PnPList -Title $(sanitise-forSql $clientName) -Template DocumentLibrary
        if($clientLibrary){
            if(![string]::IsNullOrWhiteSpace($clientDescription)){
                if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "$clientLibrary.Description = $(sanitise-stripHtml $clientDescription)"}
                $clientLibrary.Description = sanitise-stripHtml $clientDescription
                $clientLibrary.Update()
                }
            }
        else{if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Summat went wrong creating the Client Library"}}
        }
    $clientLibrary
    }
function new-spoDocumentLibrary{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [string]$docLibName

        ,[parameter(Mandatory = $true)]
        [PSCredential]$spoCredentials

        ,[parameter(Mandatory = $false)]
        [string]$docLibDescription
        )
    Write-Verbose "new-spoDocumentLibrary($docLibName, $docLibDescription)"
    try{
        Write-Verbose "`tget-spoDocumentLibrary -docLibName [$(sanitise-forSql $docLibName)]"
        $documentLibrary = get-spoDocumentLibrary -docLibName $docLibName
        }
    catch{
        Write-Verbose "`tError trying to retrieve DocLib [$docLibName]"
        $_
        }
    if($documentLibrary){Write-Verbose "Existing Library for $docLibName FOUND - not creating duplicate!"}
    else{
        Write-Verbose "`tExisting Library for $docLibName not found - creating a new one"
        Write-Verbose "`tNew-PnPList -Title [$(sanitise-forSql $docLibName)] -Template DocumentLibrary"
        try{$documentLibrary = New-PnPList -Title $(sanitise-forSql $docLibName) -Template DocumentLibrary}
        catch{
            Write-Verbose "`t`tError trying to create DocLib [$docLibName]"
            $_
            }
         try{
            #Weirdly, New-PnPList doesn't seem to return the new object, so we have to go looking for it again...
            Write-Verbose "`tget-spoDocumentLibrary -docLibName [$(sanitise-forSql $docLibName)] (after creation)"
            $documentLibrary = get-spoDocumentLibrary -docLibName $docLibName
            }
        catch{
            Write-Verbose "`tError trying to retrieve DocLib [$docLibName]"
            $_
            }
        if($documentLibrary){
            Write-Verbose "`t`tSuccess! DocLib [$($documentLibrary.RootFolder.ServerRelativeUrl)] created!"
            if(![string]::IsNullOrWhiteSpace($docLibDescription)){
                Write-Verbose "`t$($documentLibrary.Name).Description = [$(sanitise-stripHtml $docLibDescription)]"
                $documentLibrary.Description = sanitise-stripHtml $docLibDescription
                $documentLibrary.Update()
                $documentLibrary.Context.ExecuteQuery()
                }
            }
        else{Write-Verbose "Summat went wrong creating the Document Library: New-PnpList didn't return an object"}
        }
    $documentLibrary
    }
function new-spoDocumentLibraryAndSubfoldersFromPnpKimbleListItem($pnpList, $pnpListItem, $arrayOfSubfolders, $recreateSubFolderOverride, $adminCreds, $fullLogPathAndName){
    log-action "new-spoDocumentLibraryAndSubfoldersFromPnpKimbleListItem [$($pnpList.Title)] | [$($pnpListItem.Name)] " -logFile $fullLogPathAndName
    #Bodge to capture Descriptions for Clients & Suppliers
    if(![string]::IsNullOrWhiteSpace($pnpListItem.ClientDescription)){$docLibDescription = $pnpListItem.ClientDescription}
    elseif(![string]::IsNullOrWhiteSpace($pnpListItem.SupplierDescription)){$docLibDescription = $pnpListItem.SupplierDescription}
    elseif(![string]::IsNullOrWhiteSpace($pnpListItem.Description)){$docLibDescription = $pnpListItem.Description} #Who knows - there /might/ be a Description property...
    else{$docLibDescription = $null}

    $duration = Measure-Command {$newLibrary = new-spoDocumentLibrary -docLibName $pnpListItem.Name -docLibDescription $docLibDescription -spoCredentials $adminCreds -verboseLogging $verboseLogging}
    if($newLibrary){#If the new Library has been created, make the subfolders and update the List Item
        log-result "SUCCESS: $($newLibrary.RootFolder.ServerRelativeUrl) is there [$($duration.TotalSeconds) seconds]!" -logFile $fullLogPathAndName
        #Try to create the subfolders
        log-action "new-spoDocumentLibraryAndSubfoldersFromPnpKimbleListItem $($newLibrary.RootFolder.ServerRelativeUrl) [subfolders]: $($arrayOfSubfolders -join ", ")" -logFile $fullLogPathAndName
        $formattedArrayOfSubfolders = @()
        $arrayOfSubfolders | % {$formattedArrayOfSubfolders += $($newLibrary.RootFolder.ServerRelativeUrl)+"/"+$_}
        $duration = Measure-Command {$lastNewSubfolder = add-spoLibrarySubfolders -pnpList $newLibrary -arrayOfSubfolderNames $formattedArrayOfSubfolders -recreateIfNotEmpty $recreateSubFolderOverride -spoCredentials $adminCreds -verboseLogging $verboseLogging}
        if($lastNewSubfolder){        
            log-result "SUCCESS: $($lastNewSubfolder) is there [$($duration.TotalSeconds) seconds]!" -logFile $fullLogPathAndName
            #If we've got this far, try to update the (assumed) IsDirty property on the $pnpListItem in $pnpList
            $updatedValues = @{"IsDirty"=$false;"LibraryGUID"=$newLibrary.id.Guid}
            log-action "Set-PnPListItem [$($pnpList.Title)] | $($pnpListItem.Name) [$($pnpListItem.Id) @{$(stringify-hashTable $updatedValues)}]" -logFile $fullLogPathAndName
            $duration = Measure-Command {$updatedItem = Set-PnPListItem -List $pnpList.Id -Identity $pnpListItem.SPListItemID -Values $updatedValues}
            if($updatedItem.FieldValues.IsDirty -eq $false){
                log-result "SUCCESS: [$($pnpList.Title)] | $($pnpListItem.Name) is no longer Dirty [$($duration.TotalSeconds) seconds]" -logFile $fullLogPathAndName
                }
            else{log-result "FAILED: Could not set [$($pnpList.Title)] | [$($pnpListItem.Name)].IsDirty = `$false " -logFile $fullLogPathAndName}
            }
        else{log-result "FAILED: $($newLibrary.RootFolder.ServerRelativeUrl) [subfolders]: $($arrayOfSubfolders -join ", ") were not created properly" -logFile $fullLogPathAndName}
        }
    else{log-result "FAILED: Library [$($pnpList.Title)] for $($pnpListItem.Name) was not created/retrievable!" -logFile $fullLogPathAndName}    
    }
function new-spoDocumentLibraryAndSubfoldersFromSqlKimbleListItem{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [PSCustomObject]$sqlKimbleAccount,

        [parameter(Mandatory = $true)]
        [array]$arrayOfSubfolders,

        [parameter(Mandatory = $true)]
        [PSCredential]$adminCreds,

        [parameter(Mandatory = $true)]
        [string]$fullLogPathAndName,

        [parameter(Mandatory = $true)]
        [string]$errorLogPathAndName,

        [parameter(Mandatory = $true)]
        [System.Data.Common.DbConnection]$sqlDbConn,

        [parameter(Mandatory = $false)]
        [bool]$recreateSubFolderOverride
        )
    Write-Verbose "new-spoDocumentLibraryAndSubfoldersFromSqlKimbleListItem [SUS_Kimble_Accounts] | [$($sqlKimbleAccount.Name)] "
    log-action "new-spoDocumentLibraryAndSubfoldersFromSqlKimbleListItem [SUS_Kimble_Accounts] | [$($sqlKimbleAccount.Name)] " -logFile $fullLogPathAndName

    try{$duration = Measure-Command {$newLibrary = new-spoDocumentLibrary -docLibName $sqlKimbleAccount.Name -docLibDescription $sqlKimbleAccount.Description -spoCredentials $adminCreds}}
    catch{log-error -myError $_ -myFriendlyMessage "Error creating Document Library for Account [$($sqlKimbleAccount.Name)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName}
    if($newLibrary){#If the new Library has been created, make the subfolders and update the List Item
        log-result "SUCCESS: $($newLibrary.RootFolder.ServerRelativeUrl) is there [$($duration.TotalSeconds) seconds]!" -logFile $fullLogPathAndName
        #Try to create the subfolders
        log-action "new-spoDocumentLibraryAndSubfoldersFromSqlKimbleListItem $($newLibrary.RootFolder.ServerRelativeUrl) [subfolders]: $($arrayOfSubfolders -join ", ")" -logFile $fullLogPathAndName
        $formattedArrayOfSubfolders = @()
        $arrayOfSubfolders | % {$formattedArrayOfSubfolders += $($newLibrary.RootFolder.ServerRelativeUrl)+"/"+$_}
        Write-Verbose "`$formattedArrayOfSubfolders: [$formattedArrayOfSubfolders]"
        try{$duration = Measure-Command {$lastNewSubfolder = add-spoLibrarySubfolders -pnpList $newLibrary -arrayOfSubfolderNames $formattedArrayOfSubfolders -recreateIfNotEmpty $recreateSubFolderOverride -spoCredentials $adminCreds -verboseLogging $verboseLogging}}
        catch{log-error -myError $_ -myFriendlyMessage "Error creating subfolders [$($arrayOfSubfolders -join ", ")] for Account [$($sqlKimbleAccount.Name)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName}
        if($lastNewSubfolder){        
            log-result "SUCCESS: $($lastNewSubfolder.ServerRelativeUrl) is there [$($duration.TotalSeconds) seconds]!" -logFile $fullLogPathAndName
            $newLibrary
            }
        else{log-result "FAILED: $($newLibrary.RootFolder.ServerRelativeUrl) [subfolders]: $($arrayOfSubfolders -join ", ") were not created properly [$($duration.TotalSeconds) seconds]" -logFile $fullLogPathAndName}
        }
    else{log-result "FAILED: Document Library for [$($sqlKimbleAccount.Name)] was not created/retrievable [$($duration.TotalSeconds) seconds]!" -logFile $fullLogPathAndName}    
    }
function new-spoKimbleObjectListItem($kimbleObject, $pnpKimbleObjectList, $fullLogPathAndName,$verboseLogging){
    #Create the new List item
    if($verboseLogging){Write-Host -ForegroundColor Magenta "new-spoKimbleAccountItem($($kimbleObject.Name), $($pnpKimbleObjectList.Title)"}
    #Check that PNP is connected to Accounts Site
    #Check that the list is valid
    #Get the Content Type
    $contentType = $pnpKimbleObjectList.ContentTypes | ? {$_.Name -eq "Item"}
    $updateValues = @{"Title"=$kimbleObject.Name;"KimbleId"=$kimbleObject.Id;"IsDirty"=$true;"IsDeleted"=$kimbleObject.IsDeleted}
    #Different $updateValues required for Client vs Supplier
    if($pnpKimbleObjectList.Title -match "Client"){$updateValues.Add("ClientDescription",$(sanitise-stripHtml $kimbleObject.Description))}
    elseif($pnpKimbleObjectList.Title -match "Project"){$updateValues.Add("KimbleClientId",$kimbleObject.KimbleOne__Account__c)}
    elseif($pnpKimbleObjectList.Title -match "Supplier"){$updateValues.Add("SupplierDescription",$(sanitise-stripHtml $kimbleObject.Description))}
    else{}
    if($kimbleObject.LastModifiedDate){$updateValues.Add("LastModifiedDate",$(Get-Date $kimbleObject.LastModifiedDate -Format "MM/dd/yyyy hh:mm"))}
    try{
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "`tAdd-PnPListItem -List $($pnpKimbleObjectList.Title) -ContentType $($contentType.Id.StringValue) -Values @{$(stringify-hashTable $updateValues)}"}
        $newItem = Add-PnPListItem -List $pnpKimbleObjectList.Id -ContentType $($contentType.Id.StringValue) -Values $updateValues
        }
    catch{
        log-error -myError $_ -myFriendlyMessage "Error creating new [$($pnpKimbleObjectList.Title)] list item [$($kimbleObject.Name)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
        }
    if($newItem){if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "`tSUCCESS: Item [$($kimbleObject.Name)] created in [$($pnpKimbleObjectList.Title)]"}}
    else{Write-Host -ForegroundColor DarkRed "`tFAILED: Item NOT [$($kimbleObject.Name)] created in [$($pnpKimbleObjectList.Title)] :("}
    $newItem
    }
function new-spoSupplierLibrary($SupplierName, $SupplierDescription, $spoCredentials, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Magenta "new-spoSupplierLibrary($SupplierName, $SupplierDescription)"}
    if($(Get-PnPConnection).Url -ne "https://anthesisllc.sharepoint.com/Subs"){
        Connect-PnPOnline –Url $("https://anthesisllc.sharepoint.com/Subs") –Credentials $spoCredentials
        }
    try{
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Get-PnPList -Identity $(sanitise-forSql $SupplierName)"}
        $SupplierLibrary = get-spoSupplierLibrary -SupplierName $SupplierName
        }
    catch{<#Meh.#>}
    if($SupplierLibrary){if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Existing Library for $SupplierName found - not creating another"}}
    else{
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Existing Library for $SupplierName not found - creating a new one"}
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "New-PnPList -Title $(sanitise-forSql $SupplierName) -Template DocumentLibrary"}
        $SupplierLibrary = New-PnPList -Title $(sanitise-forSql $SupplierName) -Template DocumentLibrary
        if($SupplierLibrary){
            if(![string]::IsNullOrWhiteSpace($SupplierDescription)){
                if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "$SupplierLibrary.Description = $(sanitise-stripHtml $SupplierDescription)"}
                $SupplierLibrary.Description = sanitise-stripHtml $SupplierDescription
                $SupplierLibrary.Update()
                }
            }
        else{if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Summat went wrong creating the Supplier Library"}}
        }
    $SupplierLibrary
    }
function report-itemsWithUniquePermissions($pnpListItems,$permissionsHaveBeenReset,$verboseLogging){
    
    $managers = Get-UnifiedGroupLinks -LinkType Owners -Identity $(Split-Path $pnpListItems[0].Context.Url -Leaf)
    $web = $pnpListItems[0].Context.Web
    $pnpListItems[0].Context.ExecuteQuery()
    $siteTitle = $web.Title
    $siteUrl = $web.Url
    $pnpListItems | %{
        $thisItem = $_
        if($thisItem.FieldValues.FSObjType -eq 0){$iAmA = "File"}
        elseif($thisItem.FieldValues.FSObjType -eq 1){$iAmA = "Folder"}
        else{$iAmA = "Unknown Object"}
        $thisItem | Add-Member -MemberType NoteProperty -Name ItemType -Value $iAmA
        [array]$itemsToReport += $thisItem
        }

    send-itemsWithUniquePermissionsReport -arrayOfManagerMailboxes $managers -arrayOfItemsToReport $itemsToReport -siteName $siteTitle -siteUrl $siteUrl -permissionsHaveBeenReset $permissionsHaveBeenReset -verboseLogging $verboseLogging
    }
function send-itemsWithUniquePermissionsReport($arrayOfManagerMailboxes,$arrayOfItemsToReport,$siteName,$siteUrl,$permissionsHaveBeenReset,$verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Magenta "send-itemsWithUniquePermissionsReport"}
    $subject = "Non-standard sharing activity in $siteName Site"
    $arrayOfManagerMailboxes | % {
        [array]$to += $_.PrimarySmtpAddress
        $names += $_.FirstName+", "
        $finalName = $_.FirstName
        }
    if($names.Split(",").Count -gt 2){$names = $names.Replace(", $finalName"," & $finalName")}
    $body = "<HTML><FONT FACE=`"Calibri`">Hello $names`r`n`r`n<BR><BR>"
    $body += "The following items have been Shared with specific users in the <A HREF=`"$siteUrl`">$siteName</A> Site, which isn't a good way of managing access to your data (partly because it's not very transparent to see who-has-access-to-what, and partly because these unique permissions will prevent newly-added Team Members from accessing these items). "
    if($permissionsHaveBeenReset){$body += "I've reset these permissions for you, so there are no actions to take unless you want to speak with the sharer and remind them of best practices."}
    $body += "`r`n`r`n<BR><BR><UL>"
    if($arrayOfItemsToReport){
        $arrayOfItemsToReport | % {
            if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "`tSharesWithDetails: $($_.FieldValues.SharedWithDetails)"}
            $with = $_.FieldValues.SharedWithDetails.Split("{")[1].Replace("`"i:0#.f|membership|","").Replace("`":","")
            $on = $_.FieldValues.SharedWithDetails.Split("{")[2].Split(",")[0].Replace(")\/`"","")
            $on = $on.Substring($on.Length-13,13)
            $on = ([datetime]'1/1/1970').AddSeconds([int]($on / 1000))
            $by = $_.FieldValues.SharedWithDetails.Split("{")[2].Split(",")[1].Split(":")[1].Replace('"','').Replace("}","")
            $body += "<LI>$($_.ItemType)`t<B>$($_.FieldValues.FileLeafRef)</B>`tin <A HREF=`"https://anthesisllc.sharepoint.com$($_.FieldValues.FileRef)`">$($_.FieldValues.FileRef)</A> was shared with $with by $by on $on</LI>"
            }
        }
    $body += "</UL>"
    $body += "As an owner, you can manage the membership of this group (and there is a <A HREF=`"https://anthesisllc.sharepoint.com/sites/Resources-IT/SitePages/Group-membership-management-(for-Team-Managers).aspx`">guide available to help you</A>) with this and other tips for best practise, or you can contact the IT team for your region,`r`n`r`n<BR><BR>"
    $body += "Love,`r`n`r`n<BR><BR>The Helpful Groups Robot</FONT></HTML>"
    Send-MailMessage -To "kevin.maitland@anthesisgroup.com" -From "thehelpfulgroupsrobot@anthesisgroup.com" -cc "kevin.maitland@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8


    }
function set-standardTeamSitePermissions($teamSiteAbsoluteUrl, $adminCredentials, $verboseLogging,$fullLogPathAndName,$errorLogPathAndName){
    #$teamSiteAbsoluteUrl = "https://anthesisllc.sharepoint.com/teams/Energy_Engineering_Team_All_365/"
    #$teamSiteAbsoluteUrl = "https://anthesisllc.sharepoint.com/teams/Waste_&_Resource_Sustainability_WRS_Team_All_365"
   if($verboseLogging){Write-Host -ForegroundColor Magenta "set-standardTeamSitePermissions($teamSiteAbsoluteUrl, $($adminCredentials.Username))"}
    if([string]::IsNullOrWhiteSpace($teamSiteAbsoluteUrl)){
        $false
        Write-Error "Null or Empty value passed to set-standardTeamSitePermissions() for `$teamSiteAbsoluteUrl"
        }
    else{
        $teamSiteAbsoluteUrl = $teamSiteAbsoluteUrl.TrimEnd("/")
        if(!(test-pnpConnectionMatchesResource -resourceUrl $teamSiteAbsoluteUrl -verboseLogging $verboseLogging)){
            Write-Warning "Connect-PnPOnline connection mismatch - connecting to [$teamSiteAbsoluteUrl]"
            Connect-PnPOnline -Url $teamSiteAbsoluteUrl -Credentials $adminCredentials
            }

        if((test-pnpConnectionMatchesResource -resourceUrl $teamSiteAbsoluteUrl -verboseLogging $verboseLogging)){
            #Find the 365 Group associated with this Team Site
            log-action "Finding 365 group associated with [$(Split-Path $teamSiteAbsoluteUrl -Leaf)]" -logFile $fullLogPathAndName
            try{
                $ownersSpoGroup = Get-PnPGroup -AssociatedOwnerGroup 
                #Temporarily add this user to Site Collection Admins
                Set-PnPTenantSite -Url $teamSiteAbsoluteUrl -Owners (Get-PnPConnection).PSCredential.UserName
                $owner365Group = $ownersSpoGroup.Users | ? {$_.LoginName -match "federateddirectoryclaimprovider"}
                if(Get-PnPProperty -ClientObject $owner365Group -Property AadObjectId){log-result "SUCCESS: [$($owner365Group.Title)] [$($owner365Group.AadObjectId.NameId)] owns [$(Split-Path $teamSiteAbsoluteUrl -Leaf)]" -logFile $fullLogPathAndName}
                else{log-result "FAILED: Could not identify Guid for [$($owner365Group.Title)]" -logFile $fullLogPathAndName}
                }
            catch{log-error -myError $_ -myFriendlyMessage "Error finding 365 group associated with [$(Split-Path $teamSiteAbsoluteUrl -Leaf)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName}
        
            #Get the corresponding Mail-Enabeld Security Groups from AAD
            log-action "Finding the AAD groups associated with [$($owner365Group.Title)] [$($owner365Group.AadObjectId.NameId)]" -logFile $fullLogPathAndName
            try{
               if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "`$unifiedGroup = Get-UnifiedGroup -Identity `$owner365Group.AadObjectId.NameId [$($owner365Group.AadObjectId.NameId)]"}
                $unifiedGroup = Get-UnifiedGroup -Identity $owner365Group.AadObjectId.NameId
               if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "`$aadManagersGroup = Get-DistributionGroup -Identity `$unifiedGroup.CustomAttribute2 [$($unifiedGroup.CustomAttribute2)]"}
                $aadManagersGroup = Get-DistributionGroup -Identity $unifiedGroup.CustomAttribute2
               if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "`$aadMembersGroup = Get-DistributionGroup -Identity `$unifiedGroup.CustomAttribute3 [$($unifiedGroup.CustomAttribute3)]"}
                $aadMembersGroup = Get-DistributionGroup -Identity $unifiedGroup.CustomAttribute3
               if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "`$aadOverallGroup = Get-DistributionGroup -Identity `$unifiedGroup.CustomAttribute4 [$($unifiedGroup.CustomAttribute4)]"}
                $aadOverallGroup = Get-DistributionGroup -Identity $unifiedGroup.CustomAttribute4
                if($unifiedGroup -and $aadManagersGroup -and $aadMembersGroup -and $aadOverallGroup){log-result "SUCCESS: For [$($unifiedGroup.DisplayName)], the Managers group is [$($aadManagersGroup.DisplayName)], the Members Group is [$($aadMembersGroup.DisplayName)] and the combined Group is [$($aadOverallGroup.DisplayName)]" -logFile $fullLogPathAndName}
                else{log-result "FAILED: For [$($unifiedGroup.DisplayName)], the Managers group is [$($aadManagersGroup.DisplayName)], the Members Group is [$($aadMembersGroup.DisplayName)] and the combined Group is [$($aadOverallGroup.DisplayName)]"}
                }
            catch{log-error -myError $_ -myFriendlyMessage "Error finding AAD groups associated with [$(Split-Path $teamSiteAbsoluteUrl -Leaf)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName}

            if([string]::IsNullOrWhiteSpace($aadMembersGroup)){
                #Notify someone that there is no Members Group associated with this 365 Group
                }
            if([string]::IsNullOrWhiteSpace($aadManagersGroup)){
                #Notify someone that there is no Managers Group associated with this 365 Group
                }


            #Add Managers group to Site Coll Admins & Site Owners Group
            if($aadManagersGroup){
                #Add the AAD Managers group to the Site Owners Group #I'm not sure we want to do this :/
                #Add-PnPUserToGroup -EmailAddress $aadManagersGroup.PrimarySmtpAddress -Identity $ownersSpoGroup.Id -SendEmail:$false
                #Get the SPO version of the AAD Managers Group (as we need the SharePoint LoginName)
                $managersSpoObject = Get-PnPUser | ? {$_.Email -eq $($aadManagersGroup.PrimarySmtpAddress)}
                #If we didn;t find it, we need to add it like this:
                if(!$managersSpoObject){$managersSpoObject = New-PnPUser -LoginName $($aadManagersGroup.PrimarySmtpAddress)}
                #Add the Managers group as a Site Collection Administrator
                if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Set-PnPTenantSite -Url $teamSiteAbsoluteUrl -Owners $($managersSpoObject.LoginName)"}
                Set-PnPTenantSite -Url $teamSiteAbsoluteUrl -Owners $managersSpoObject.LoginName
                if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Set-PnPTenantSite -Url $teamSiteAbsoluteUrl -Owners 'kimblebot@anthesisgroup.com'"}
                Set-PnPTenantSite -Url $teamSiteAbsoluteUrl -Owners "kimblebot@anthesisgroup.com"
                }

            #Check the Site Collection Administrators
            $siteCollectionAdmins = Get-PnPSiteCollectionAdmin
            if($aadMembersGroup){
                $siteCollectionAdmins | ? {$_.Email -eq $aadMembersGroup.PrimarySmtpAddress} | % {
                    #Remove the AAD Members Group from Site Collection admins (if it's there)
                    if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Remove-PnPSiteCollectionAdmin -Owners $($_.Email)"}
                    Remove-PnPSiteCollectionAdmin -Owners $_.LoginName
                    }
                }
            if($aadOverallGroup){
                $siteCollectionAdmins | ? {$_.Email -eq $aadOverallGroup.PrimarySmtpAddress} | % {
                    #Remove the AAD Members Group from Site Collection admins (if it's there)
                    if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Remove-PnPSiteCollectionAdmin -Owners $($_.Email)"}
                    Remove-PnPSiteCollectionAdmin -Owners $_.LoginName
                    }
                }
            if($siteCollectionAdmins.Title -notcontains "Kimble Bot"){Write-Warning "KimbleBot is not a Site Collection Administrator"}
            if($siteCollectionAdmins.Email -notcontains $managersGroup.Email){
                if($managersGroup){Write-Warning "$($managersGroup.Title) was not added as a Site Collection Administrator"}
                }

            #Block all external sharing
            if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Blocking external Sharing: Set-PnPTenantSite -Url $teamSiteAbsoluteUrl -Sharing Disabled"}
            Set-PnPTenantSite -Url $teamSiteAbsoluteUrl -Sharing Disabled

            #Enable the DocID service
            if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Enabling Document ID Service Feature on Site Collection"}
            $site = Get-PnPSite
            $site.Features.Add([guid]"b50e3104-6812-424f-a011-cc90e6327318",$false,[Microsoft.SharePoint.Client.FeatureDefinitionScope]::None)
            $site.Context.ExecuteQuery()
                    
            #Untick Members can share boxes 
            #***************************************************************************************************************************
            # Requires temporary elevation to Site Owners Group (assumes Site Collection administrator rights already granted)
            #***************************************************************************************************************************
            if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Restricting internal Sharing: (MembersCanShare & AllowMembersEditMembership = `$false)"}
            Add-PnPUserToGroup -EmailAddress (Get-PnPConnection).PSCredential.UserName -Identity $ownersSpoGroup.Id -SendEmail:$false
            $thisWeb = Get-PnPWeb -Includes MembersCanShare, AssociatedMemberGroup.AllowMembersEditMembership
            $thisWeb.MembersCanShare = $false
            $thisWeb.AssociatedMemberGroup.AllowMembersEditMembership = $false
            $thisWeb.AssociatedMemberGroup.Update()
            $thisWeb.Update()
            $thisWeb.Context.ExecuteQuery()
            if((Get-PnPConnection).PSCredential.UserName -eq "kimblebot@anthesisgroup.com"){Remove-PnPUserFromGroup -LoginName "i:0#.f|membership|kimblebot@anthesisgroup.com" -Identity $ownersSpoGroup.Id} #Special case for KimbleBot as it (intentionally) doesn't have an E1 license
            else{#Remove the current user from the Site Owners and Site Collection Admins
                Remove-PnPUserFromGroup -LoginName (Get-PnPConnection).PSCredential.UserName -Identity $ownersSpoGroup.Id
                Remove-PnPSiteCollectionAdmin -Owners (Get-PnPConnection).PSCredential.UserName
                }
        
            <#
            #Break inheritance on Documents folder and prevent Owners from sharing contents
            $standardDocumentLibrary = Get-PnPList -Includes FirstUniqueAncestorSecurableObject,HasUniqueRoleAssignments -Identity "Shared Documents"
            #if($standardDocumentLibrary.FirstUniqueAncestorSecurableObject.Id -eq $standardDocumentLibrary.Id){
            if($standardDocumentLibrary){
                if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Resetting permissions on Documents Library"}
                $standardDocumentLibrary.ResetRoleInheritance()
                $standardDocumentLibrary.Update()
                $standardDocumentLibrary.Context.ExecuteQuery()
                $standardDocumentLibrary.BreakRoleInheritance($true,$true)
                $standardDocumentLibrary.Update()
                $standardDocumentLibrary.Context.ExecuteQuery()
                Set-PnPListPermission -Identity "Documents" -Group $ownersSpoGroup -AddRole "Edit" -RemoveRole "Full Control"
                #E-mail Managers to let them know that content had been shared.
                }
            #Check whether any items in the Documents have unique permissions on them
            if ((get-allSpoListsWithItemsWithUniquePermissions -siteAbsoluteUrl $teamSiteAbsoluteUrl -adminCredentials $adminCredentials -verboseLogging $verboseLogging).Title -contains $standardDocumentLibrary.Title){
                if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Custom permissions found on LIst Items - resetting them"}
                [array]$itemsWithUniquePermissions = get-allSpoListItemsWithUniquePermissions -pnpList $standardDocumentLibrary -adminCredentials $adminCredentials -verboseLogging $verboseLogging
                if($itemsWithUniquePermissions){
                    $itemsWithUniquePermissions | % {
                        $thisItem = $_
                        $thisItem.ResetRoleInheritance()
                        $thisItem.Update()
                        $thisItem.BreakRoleInheritance($true,$true)
                        $thisItem.Update()
                        $thisItem.ResetRoleInheritance()
                        $thisItem.FieldValues["SharedWithUsers"].SetValue([Microsoft.SharePoint.Client.FieldLookupValue]::new())
                    
                        $thisItem.Update()
                        $thisItem.Context.ExecuteQuery()
                        #Set-PnPListItemPermission -List $standardDocumentLibrary.Id -Identity $thisItem.Id -InheritPermissions
                        }
                    $itemsWithUniquePermissions[0].Context.ExecuteQuery()
                    report-itemsWithUniquePermissions -pnpListItems $itemsWithUniquePermissions -permissionsHaveBeenReset $true -verboseLogging $verboseLogging
                    }
                }
                #>
            if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "All finished"}
            }
        else{Write-Error "Could not connect to Site"}
        }
    }
function test-pnpConnectionMatchesResource($resourceUrl, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Magenta "test-pnpConnectionMatchesResource($resourceUrl, $($adminCredentials.UserName)"}
    try{Get-PnPConnection | Out-Null}
    catch{
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "No Connect-PnPOnline connection available. Creating new Connect-PnpOnline to [$resourceUrl]"}
        $false
        break
        }
    if((split-path ([System.Uri](Get-PnPConnection).Url).LocalPath -Leaf) -eq (Split-Path $resourceUrl -Leaf)){
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Connect-PnPOnline connection matches [$resourceUrl]"}
        $true
        }
    else{
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Connect-PnPOnline connection [$([System.Uri](Get-PnPConnection).Url).LocalPath)] does not match [$resourceUrl]"}
        $false
        }
    }
function update-spoDocumentLibraryAndSubfoldersFromPnpKimbleListItem($pnpList, $pnpListItem, $arrayOfSubfolders, $recreateSubFolderOverride, $adminCreds, $fullLogPathAndName, $verboseLogging){
    log-action "update-spoDocumentLibraryAndSubfoldersFromPnpKimbleListItem [$($pnpListItem.Name)] - looking for existing Library" -logFile $fullLogPathAndName
    try{
        $duration = Measure-Command {
            #Try to get the Document Library by GUID (most accurate), then by PreviousName (next most likely), then by Name (least likely)
            $existingLibrary = get-spoDocumentLibrary -docLibName $pnpListItem.PreviousName -docLibGuid $pnpListItem.LibraryGUID
            if(!$existingLibrary){$existingLibrary = get-spoDocumentLibrary -docLibName $pnpListItem.Name}
            }
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving Document Library in update-spoDocumentLibraryAndSubfoldersFromPnpKimbleListItem [$($pnpListItem.Name)][$($pnpListItem.LibraryGUID)] $($Error[0].Exception.InnerException.Response)" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName}

    if($existingLibrary){
        log-result -myMessage "SUCCESS: [$($existingLibrary.RootFolder.ServerRelativeUrl)] found (GUID:[$($existingLibrary.Id.Guid)] [$($duration.TotalSeconds) seconds])" -logFile $fullLogPathAndName
        log-action -myMessage "Updating Document Library [$($existingLibrary.RootFolder.ServerRelativeUrl)]" -logFile $fullLogPathAndName
        #Bodge to capture Descriptions for Clients & Suppliers
        if(![string]::IsNullOrWhiteSpace($pnpListItem.ClientDescription)){$docLibDescription = $pnpListItem.ClientDescription}
        elseif(![string]::IsNullOrWhiteSpace($pnpListItem.SupplierDescription)){$docLibDescription = $pnpListItem.SupplierDescription}
        elseif(![string]::IsNullOrWhiteSpace($pnpListItem.Description)){$docLibDescription = $pnpListItem.Description} #Who knows - there /might/ be a Description property...
        else{$docLibDescription = $null}

        try{
            #Update the Library
            if($verboseLogging){Write-Host -ForegroundColor DarkCyan "Updating Library [$($existingLibrary.RootFolder.ServerRelativeUrl)] Description:[$(sanitise-stripHtml $docLibDescription)]"}
            $duration = Measure-Command {
                $existingLibrary.Description = $(sanitise-stripHtml $docLibDescription)
                $existingLibrary.Update()
                $existingLibrary.Context.ExecuteQuery()
                if($verboseLogging){Write-Host -ForegroundColor DarkCyan "Updating Library [$($existingLibrary.RootFolder.ServerRelativeUrl)] Title:[$($pnpListItem.Name)]: Set-PnPList -Identity $($existingLibrary.Id.Guid) -Title $($pnpListItem.Name)"}
                Set-PnPList -Identity $existingLibrary.Id -Title $pnpListItem.Name
                $updatedLibrary = Get-PnPList -Identity $existingLibrary.Id #The Id property is constant between $existingLibrary and $updatedLibrary 
                }
            #Check the update worked
            if($updatedLibrary.Title -eq $pnpListItem.Name -and $(sanitise-stripHtml $updatedLibrary.Description) -eq $(sanitise-stripHtml $docLibDescription)){
                log-result -myMessage "SUCCESS: Client Library [$($existingLibrary.RootFolder.ServerRelativeUrl)] updated successfully [$($duration.TotalSeconds) secs]" -logFile $fullLogPathAndName
                try{
                    #Update the List Item
                    $duration = Measure-Command {
                        $updatedValues =@{"LibraryGUID"=$existingLibrary.Id.Guid;"IsDirty"=$false}
                        if($verboseLogging){Write-Host -ForegroundColor DarkCyan "Updating [$($pnpList.Title)] | [$($pnpListItem.Name)]: Set-PnPListItem -List $($pnpList.Id) -Identity $($pnpListItem.SPListItemID) `$updatedValues = @{$(stringify-hashTable $updatedValues)}"}
                        $updatedListItem = Set-PnPListItem -List $pnpList.Id -Identity $pnpListItem.SPListItemID -Values $updatedValues
                        }
                    if($updatedListItem.FieldValues.IsDirty -eq $false){log-result -myMessage "SUCCESS: [Kimble Clients].[$($pnpListItem.Name)] updated successfully (no error on update) [$($duration.TotalSeconds) secs]" -logFile $fullLogPathAndName}
                    else{log-result -myMessage "FAILED: [$($pnpList.Title)] | [$($pnpListItem.Name)] was not updated" -logFile $fullLogPathAndName}
                    }
                catch{
                    #Failed to update SPListItem
                    log-error -myError $_ -myFriendlyMessage "Error updating [$($pnpList.Title)] | [$($pnpListItem.Name)] - this is still marked as IsDirty=`$true :( [$($Error[0].Exception.InnerException.Response)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
                    }
                }
            }
        catch{
            #Failed to update Client Library
            log-result -myMessage "FAILED: Document Library [$($existingLibrary.Title)] was found, but not updated" -logFile $fullLogPathAndName
            log-error -myError $_ -myFriendlyMessage "Error updating Document Library [$($existingLibrary.Title)] - this is still marked as IsDirty=`$true :( [$($Error[0].Exception.InnerException.Response)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
            }
        }
    else{
        #Couldn't find the Library, so try creating a new one to paper over the cracks. #WCGW
        log-result -myMessage "FAILED: Could not retrieve a Document Library for [$($pnpList.Title)] | [$($pnpListItem.Name)] - sending it back for re-creation :/" -logFile $fullLogPathAndName
        log-action -myMessage "Sending [$($pnpList.Title)] | [$($pnpListItem.Name)] back for re-creation as it has mysteriously disappeared" -logFile $fullLogPathAndName
        if($verboseLogging){Write-Host -ForegroundColor DarkCyan "new-clientFolder -spoKimbleClientList $($pnpList.Title) -spoKimbleClientListItem $($pnpListItem.Name) -arrayOfClientSubfolders @($($arrayOfSubfolders -join ",")) -recreateSubFolderOverride `$false"}
        try{
            $duration = Measure-Command {$newLibrary = new-spoDocumentLibraryAndSubfoldersFromPnpKimbleListItem -pnpList $pnpList -pnpListItem $pnpListItem -arrayOfSubfolders $arrayOfSubfolders -recreateSubFolderOverride $false -adminCreds $adminCreds -fullLogPathAndName $fullLogPathAndName}
            if($newLibrary){log-result -myMessage "SUCCESS: Weirdly unfindable Client Library [$($newLibrary.RootFolder.ServerRelativeUrl)] was recreated [$($duration.TotalSeconds) secs]" -logFile $fullLogPathAndName}
            else{
                log-result -myMessage "FAILED: Someone left a sponge in the patient - I couldn't retrieve a Document Library for [$($pnpList.Title)] | [$($pnpListItem.Name)] and I couldn't create a new one either..." -logFile $fullLogPathAndName
                log-error -myError $null -myFriendlyMessage "Borked update-spoDocumentLibraryAndSubfoldersFromPnpKimbleListItem [$($pnpList.Title)] | [$($pnpListItem.Name)]" -smtpServer "anthesisgroup-com.mail.protection.outlook.com" -mailTo "kevin.maitland@anthesisgroup.com" -mailFrom "$(split-path $PSCommandPath -Leaf)_netmon@sustain.co.uk"
                }
            }
        catch{log-error -myError $_ -myFriendlyMessage "Error: Borked update-spoDocumentLibraryAndSubfoldersFromPnpKimbleListItem [$($pnpList.Title)] | [$($pnpListItem.Name)]" -smtpServer "anthesisgroup-com.mail.protection.outlook.com" -mailTo "kevin.maitland@anthesisgroup.com" -mailFrom "$(split-path $PSCommandPath -Leaf)_netmon@sustain.co.uk"}
        }
    }
function update-spoDocumentLibraryAndSubfoldersFromSqlKimbleListItem{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [PSCustomObject]$sqlKimbleAccount

        ,[parameter(Mandatory = $true)]
        [array]$arrayOfSubfolders

        ,[parameter(Mandatory = $true)]
        [PSCredential]$adminCreds

        ,[parameter(Mandatory = $true)]
        [string]$fullLogPathAndName

        ,[parameter(Mandatory = $true)]
        [string]$errorLogPathAndName

        ,[parameter(Mandatory = $true)]
        [System.Data.Common.DbConnection]$sqlDbConn

        ,[parameter(Mandatory = $false)]
        [bool]$recreateSubFolderOverride
        )
    log-action "update-spoDocumentLibraryAndSubfoldersFromSqlKimbleListItem [$($sqlKimbleAccount.Name)] - looking for existing Library" -logFile $fullLogPathAndName
    try{
        $duration = Measure-Command {
            #Try to get the Document Library by GUID (most accurate), then by PreviousName (next most likely), then by Name (least likely)
            $existingLibrary = get-spoDocumentLibrary -docLibName $sqlKimbleAccount.PreviousName -docLibGuid $sqlKimbleAccount.DocumentLibraryGuid
            if(!$existingLibrary){$existingLibrary = get-spoDocumentLibrary -docLibName $sqlKimbleAccount.Name}
            }
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving Document Library in update-spoDocumentLibraryAndSubfoldersFromSqlKimbleListItem [$($pnpListItem.Name)][$($pnpListItem.DocumentLibraryGuid)] $($Error[0].Exception.InnerException.Response)" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName}

    if($existingLibrary){
        log-result -myMessage "SUCCESS: [$($existingLibrary.RootFolder.ServerRelativeUrl)] found (GUID:[$($existingLibrary.Id.Guid)] [$($duration.TotalSeconds) seconds])" -logFile $fullLogPathAndName
        log-action -myMessage "Updating Document Library [$($existingLibrary.RootFolder.ServerRelativeUrl)]" -logFile $fullLogPathAndName

        try{
            #Update the Library
            if($verboseLogging){Write-Host -ForegroundColor DarkCyan "Updating Library [$($existingLibrary.RootFolder.ServerRelativeUrl)] Description:[$(sanitise-forSqlValue -value $sqlKimbleAccount.Description -dataType HTML) ]"}
            $duration = Measure-Command {
                $existingLibrary.Description = $(sanitise-stripHtml $sqlKimbleAccount.Description)
                $existingLibrary.Update()
                $existingLibrary.Context.ExecuteQuery()
                if($verboseLogging){Write-Host -ForegroundColor DarkCyan "Updating Library [$($existingLibrary.RootFolder.ServerRelativeUrl)] Title:[$($sqlKimbleAccount.Name)]: Set-PnPList -Identity [$($existingLibrary.Id.Guid)] -Title [$($sqlKimbleAccount.Name)]"}
                Set-PnPList -Identity $existingLibrary.Id -Title $sqlKimbleAccount.Name
                $updatedLibrary = Get-PnPList -Identity $existingLibrary.Id #The Id property is constant between $existingLibrary and $updatedLibrary 
                }
            #Check the update worked
            if($updatedLibrary.Title -eq $sqlKimbleAccount.Name -and $(sanitise-stripHtml $updatedLibrary.Description) -eq $(sanitise-stripHtml $sqlKimbleAccount.Description)){
                log-result -myMessage "SUCCESS: Client Library [$($existingLibrary.RootFolder.ServerRelativeUrl)] updated successfully [$($duration.TotalSeconds) secs]" -logFile $fullLogPathAndName
                $updatedLibrary
                }
            }
        catch{
            #Failed to update Client Library
            log-result -myMessage "FAILED: Document Library [$($existingLibrary.Title)] was found, but not updated" -logFile $fullLogPathAndName
            log-error -myError $_ -myFriendlyMessage "Error updating Document Library [$($existingLibrary.Title)] - this is still marked as IsDirty=`$true :( [$($Error[0].Exception.InnerException.Response)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
            }
        }
    else{
        #Couldn't find the Library, so try creating a new one to paper over the cracks. #WCGW
        log-result -myMessage "FAILED: Could not retrieve a Document Library for [$($sqlKimbleAccount.Name)] - sending it back for re-creation :/" -logFile $fullLogPathAndName
        log-action -myMessage "Sending [$($sqlKimbleAccount.Name)] back for re-creation as it has mysteriously disappeared" -logFile $fullLogPathAndName
        if($verboseLogging){Write-Host -ForegroundColor DarkCyan "new-spoDocumentLibraryAndSubfoldersFromSqlKimbleListItem -sqlKimbleAccount $($sqlKimbleAccount.Name) -sqlDbConn $($sqlDbConn.DataSource) -arrayOfClientSubfolders @($($arrayOfSubfolders -join ",")) -recreateSubFolderOverride `$false"}
        try{
            $duration = Measure-Command {$newLibrary = new-spoDocumentLibraryAndSubfoldersFromSqlKimbleListItem -sqlKimbleAccount $sqlKimbleAccount -sqlDbConn $sqlDbConn -arrayOfSubfolders $arrayOfSubfolders -recreateSubFolderOverride $false -adminCreds $adminCreds -fullLogPathAndName $fullLogPathAndName -errorLogPathAndName $errorLogPathAndName}
            if($newLibrary){log-result -myMessage "SUCCESS: Weirdly unfindable Client Library [$($newLibrary.RootFolder.ServerRelativeUrl)] was recreated [$($duration.TotalSeconds) secs]" -logFile $fullLogPathAndName}
            else{
                log-result -myMessage "FAILED: Someone left a sponge in the patient - I couldn't retrieve a Document Library for [$($sqlKimbleAccount.Name)] and I couldn't create a new one either..." -logFile $fullLogPathAndName
                log-error -myError $null -myFriendlyMessage "Borked update-spoDocumentLibraryAndSubfoldersFromSqlKimbleListItem [$($sqlKimbleAccount.Name)]" -smtpServer "anthesisgroup-com.mail.protection.outlook.com" -mailTo "kevin.maitland@anthesisgroup.com" -mailFrom "$(split-path $PSCommandPath -Leaf)_netmon@sustain.co.uk"
                }
            }
        catch{log-error -myError $_ -myFriendlyMessage "Error: Borked update-spoDocumentLibraryAndSubfoldersFromSqlKimbleListItem [$($sqlKimbleAccount.Name)]" -smtpServer "anthesisgroup-com.mail.protection.outlook.com" -mailTo "kevin.maitland@anthesisgroup.com" -mailFrom "$(split-path $PSCommandPath -Leaf)_netmon@sustain.co.uk"}
        }
    }
function update-spoKimbleObjectListItem($kimbleObject, $pnpKimbleObjectList, $overrideIsDirtyTrue, $overrideIsDirtyFalse, $overrideIsOrphanedTrue, $overrideIsOrphanedFalse, $overrideIsMisclassified, $fullLogPathAndName,$verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Magenta "update-spoKimbleAccountItem($($kimbleObject.Name), $($pnpKimbleObjectList.Title)"}
    #Bodge the KimbleId value if it's not present (this happens when a SalesForce object is submitted, rather than a pnpListItem)
    if([string]::IsNullOrWhiteSpace($kimbleObject.KimbleId) -and $kimbleObject.Id.Length -eq 18){
        $kimbleObject | Add-Member -MemberType NoteProperty -Name KimbleId -Value $kimbleObject.Id
        }
    #Retrieve the existing item
    try{
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Get-PnPListItem -List $($pnpKimbleObjectList.Title) -Query <View><Query><Where><Eq><FieldRef Name='KimbleId'/><Value Type='Text'>$($kimbleObject.KimbleId)</Value></Eq></Where></Query></View> -ErrorAction Stop"}
        $existingPnpListItem = Get-PnPListItem -List $pnpKimbleObjectList -Query "<View><Query><Where><Eq><FieldRef Name='KimbleId'/><Value Type='Text'>$($kimbleObject.KimbleId)</Value></Eq></Where></Query></View>" -ErrorAction Stop
        }
    catch{
        log-error -myError $_ -myFriendlyMessage "Error retrieving [$($pnpKimbleObjectList.Title)] list item [$($kimbleObject.Name)] in update-spoKimbleAccountItem()" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
        }

    #Update it
    if(!$existingPnpListItem){
        if($verboseLogging){Write-Host -ForegroundColor DarkRed "`tFAILED: Existing item [$($kimbleObject.Name)] in [$($pnpKimbleObjectList.Title)] not found"}
        }
    else{
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "`tSUCCESS: Existing item [$($existingPnpListItem.FieldValues.Title)][$($existingPnpListItem.Id)] in [$($pnpKimbleObjectList.Title)] found"}
        #We need to test whether any of the properties *that we're interested in* have been updated as it's really expensive to query even individual Document Libraries in the Clients Site, so we need to minimise the number of [Kimble XYZ] records we flag as isDirty = $true
        #$updateValues = @{"Title"=$kimbleObject.Name;"KimbleId"=$kimbleObject.Id;"IsDirty"=$true;"IsDeleted"=$kimbleObject.IsDeleted}
        $updateValues = @{}
        #Compare Names and update if changed
        if($kimbleObject.Name -ne $existingPnpListItem.FieldValues.Title){
            $updateValues.Add("Title",$kimbleObject.Name)
            $updateValues.Add("PreviousName",$existingPnpListItem.FieldValues.Title)
            } 
        #Compare IsDeleted and update if changed
        if($kimbleObject.IsDeleted -ne $existingPnpListItem.FieldValues.IsDeleted){$updateValues.Add("IsDeleted",$kimbleObject.IsDeleted)} 
        #Split out Clients & Suppliers as we stupidly gave them different field names
        if($pnpKimbleObjectList.Title -match "Client"){
            if($(sanitise-stripHtml $kimbleObject.Description) -ne $(sanitise-stripHtml $existingPnpListItem.FieldValues.ClientDescription)){$updateValues.Add("ClientDescription",$(sanitise-stripHtml $kimbleObject.Description))}#Compare Descriptions and update if changed
            }
        elseif($pnpKimbleObjectList.Title -match "Project"){
            if($kimbleObject.KimbleOne__Account__c -ne $existingPnpListItem.FieldValues.KimbleClientId){
                $updateValues.Add("KimbleClientId",$kimbleObject.KimbleOne__Account__c)
                $updateValues.Add("PreviousKimbleClientId",$existingPnpListItem.FieldValues.KimbleClientId)
                }
            }
        elseif($pnpKimbleObjectList.Title -match "Supplier"){
            if($(sanitise-stripHtml $kimbleObject.Description) -ne $(sanitise-stripHtml $existingPnpListItem.FieldValues.SupplierDescription)){$updateValues.Add("SupplierDescription",$(sanitise-stripHtml $kimbleObject.Description))}#Compare Descriptions and update if changed
            }
        else{}#Just in case we accidentally pass anything other than a Client, Project or Supplier through
        if($updateValues.Count -gt 0){$updateValues.Add("IsDirty",$true)} #If something notable (and only if) has changed, flag as IsDirty
        #else{$updateValues.Add("IsDirty",$false)} #We don't want to automatically mark items as IsDirty = $false because we might not have processed them yet. If something goes bonkers again and marks thousands of records as IsDirty, we've go the override function below to acheive this that we can call from a reconcile-XXX function
        if($kimbleObject.LastModifiedDate){
            if($(get-date $kimbleObject.LastModifiedDate) -ne $(get-date $existingPnpListItem.FieldValues.LastModifiedDate)){$updateValues.Add("LastModifiedDate",$(Get-Date $kimbleObject.LastModifiedDate -Format "MM/dd/yyyy HH:mm:ss"))}
            }

        #Now handle overrides
        if($overrideIsDirtyTrue){
            if($updateValues.ContainsKey("IsDirty")){$updateValues["IsDirty"] = $true}
            else{$updateValues.Add("IsDirty",$true)}
            }
        if($overrideIsDirtyFalse){
            if($updateValues.ContainsKey("IsDirty")){$updateValues["IsDirty"] = $false}
            else{$updateValues.Add("IsDirty",$false)}
            }
        if($overrideIsOrphanedTrue){
            if($updateValues.ContainsKey("IsOrphaned")){$updateValues["IsOrphaned"] = $true}
            else{$updateValues.Add("IsOrphaned",$true)}
            }
        if($overrideIsOrphanedFalse){
            if($updateValues.ContainsKey("IsOrphaned")){$updateValues["IsOrphaned"] = $false}
            else{$updateValues.Add("IsOrphaned",$false)}
            }
        if($overrideIsMisclassified){
            if($updateValues.ContainsKey("isMisclassified")){$updateValues["isMisclassified"] = $true}
            else{$updateValues.Add("isMisclassified",$true)}
            }

        if($updateValues){ #If there's nothing to update, there's no need to waste time talking to SharePoint
            try{
                if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Set-PnPListItem -List $($pnpKimbleObjectList.Id) -Identity $($existingPnpListItem.Id) -Values @{$(stringify-hashTable $updateValues)}"}
                $updatedItem = Set-PnPListItem -List $pnpKimbleObjectList.Id -Identity $existingPnpListItem.Id -Values $updateValues -ErrorAction Stop
                }
                    catch{
            log-error -myError $_ -myFriendlyMessage "Error updating [$($pnpKimbleObjectList.Title)] list item [$($existingPnpListItem.FieldValues.Title)] in update-spoKimbleAccountItem()" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
            }
            if($updatedItem){if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "`tSUCCESS: Item [$($updatedItem.FieldValues.Title)] updated in [$($pnpKimbleObjectList.Title)]"}}
            else{Write-Host -ForegroundColor DarkRed "`tFAILED: Item [$($existingPnpListItem.FieldValues.Title)] NOT updated in [$($pnpKimbleObjectList.Title)], even though I found it to update :("}
            }
        else{if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "[$($pnpKimbleObjectList.Title)] | [$($existingPnpListItem.FieldValues.Title)] didn't need updating after all, so I've left it in peace"}}
        }
    #Return it
    $updatedItem
    }
function update-spoTerm($termGroup,$termSet,$oldTerm,$newTerm,$kimbleId,$verboseLogging){
     if($verboseLogging){Write-Host -ForegroundColor Magenta "update-spoTerm($termGroup,$termSet,$oldTerm,$newTerm)"}
     $cleanOldTerm = $(sanitise-forTermStore $oldTerm)
     $cleanNewTerm = $(sanitise-forTermStore $newTerm)
    try{
        $pnpTermGroup = Get-PnPTermGroup $termGroup 
        $pnpTermSet = Get-PnPTermSet -TermGroup $pnpTermGroup -Identity $termSet
        #$pnpOldTerm = Get-PnPTerm -Identity $cleanOldTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet 
        #if(!$pnpOldTerm){Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Identity $oldTerm} #Try the dirty version if we can't find the clean version
        #$pnpNewTerm = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Identity $cleanNewTerm
        #2019-03-14 [KM] Retrieving all Terms now as it's bizarrely faster than retrieving an individual term and we're hitting a 30 second timeout.
        $allTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet
        $pnpOldTerm = $allTerms | ? {$_.Name -eq $cleanOldTerm}
        if(!$pnpOldTerm){$allTerms | ? {$_.Name -eq $oldTerm}} #Try the dirty version if we can't find the clean version
        $pnpNewTerm = $allTerms | ? {$_.Name -eq $cleanNewTerm}
        }
    catch{
        #Meh.
        }
    if($pnpOldTerm -and !$pnpNewTerm){
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "$termGroup | $termSet | [$cleanOldTerm]  found - updating to [$cleanNewTerm]"}
        $pnpOldTerm.Name = $cleanNewTerm
        if(![string]::IsNullOrEmpty($kimbleId)){$pnpOldTerm.SetCustomProperty("KimbleId",$kimbleId)}
        $pnpOldTerm.Context.ExecuteQuery()
        $pnpOldTerm
        }
    elseif($pnpNewTerm -and $pnpOldTerm){
        #Deprecate the old term as the new one has already been created. we don't delete in case it's in use anywhere
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "$termGroup | $termSet | [$cleanNewTerm] is already present - deprecating old term: [$cleanOldTerm]"}
        $pnpOldTerm.Deprecate($true)
        $pnpOldTerm.Context.ExecuteQuery()
        if(![string]::IsNullOrEmpty($kimbleId)){
            $pnpNewTerm.SetCustomProperty("KimbleId",$kimbleId)
            $pnpNewTerm.Context.ExecuteQuery()
            }
        $pnpNewTerm
        }
    else{
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "$termGroup | $termSet | [$cleanOldTerm]  not found - creating new term: [$newTerm]"}
        add-spoTermToStore -termGroup $termGroup -termSet $termSet -term $newTerm -kimbleId $kimbleId -verboseLogging $verboseLogging #We are deliberately not sending the $cleanNewTerm
        }
    }