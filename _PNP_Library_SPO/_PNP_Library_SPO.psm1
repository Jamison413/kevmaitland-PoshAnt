function add-spoLibrarySubfolders($pnpList, $arrayOfSubfolderNames, $recreateIfNotEmpty, $spoCredentials, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Magenta "add-spoLibrarySubfolders($($pnpList.Title), $($arrayOfSubfolderNames -join ", "), `$recreateIfNotEmpty=$recreateIfNotEmpty"}
    if($(Get-PnPConnection).Url -notmatch $pnpList.ParentWebUrl){
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Connected to wrong site - connecting to $($pnpList.RootFolder.Context.Url)"}
        Connect-PnPOnline –Url $($pnpList.RootFolder.Context.Url) –Credentials $spoCredentials
        }
    [array]$formattedArrayOfSubfolderNames = $arrayOfSubfolderNames | % {format-asServerRelativeUrl -serverRelativeUrl $pnpList.RootFolder.ServerRelativeUrl -stringToFormat $_}
    try{
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "get-spoFolder -pnpList $($pnpList.Title) -folderServerRelativeUrl $($formattedArrayOfSubfolderNames[$formattedArrayOfSubfolderNames.Length-1])"}
        $hasItems = get-spoFolder -pnpList $pnpList -folderServerRelativeUrl $($formattedArrayOfSubfolderNames[$formattedArrayOfSubfolderNames.Length-1]) -adminCreds $adminCreds -verboseLogging $verboseLogging
        #$hasItems = Get-PnPListItem -List $pnpList -Query "<View><RowLimit>5</RowLimit></View>" #This RowLimit doesn't work at the moment, but hopefully it'll get fixed in the future and this'll be efficient https://github.com/SharePoint/PnP-PowerShell/issues/879
        #$hasItems = Get-PnPListItem -List $pnpList -Query "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='Text'>DummyOp5 (E003941)</Value></Eq></Where></Query></View>" 
        #$hasItems = Get-PnPListItem -List $pnpList -Query "<View><Query><Where><Eq><FieldRef Name='FileRef'/><Value Type='Text'>/clients/DummyCo Ltd/DummyOp5 (E003941)</Value></Eq></Where></Query></View>" 
        #$hasItems = Get-PnPListItem -List $pnpList -Query "<View><Query><Where><Eq><FieldRef Name='FileRef'/><Value Type='Text'>/clients/DummyCo Ltd/DummyOp5 (E003941)/Analysis</Value></Eq></Where></Query></View>" 
        #$hasItems = Get-PnPListItem -List $pnpList #-Query "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='Text'>$($arrayOfSubfolderNames[0])</Value></Eq></Where></Query></View>" 
        #$hasItems = $hasItems | ? {$_.FieldValues.FileRef -eq "$($arrayOfSubfolderNames[$arrayOfSubfolderNames.Length-1])"}
        }
    catch{
        #Meh.
        }
    if(!$hasItems -or $recreateIfNotEmpty){
        if($verboseLogging){
            if(!$hasItems){Write-Host -ForegroundColor DarkMagenta "$($pnpList.RootFolder.ServerRelativeUrl) has no conflicting item - creating subfolder/s"}
            else{Write-Host -ForegroundColor DarkMagenta "$($pnpList.RootFolder.ServerRelativeUrl) has items, but override set - creating subfolders"}
            }
        $formattedArrayOfSubfolderNames | % {
            #We have to search for these using ServerRelativeUrls, but create them using LibraryRelativeUrls
            $libraryRelativePath = $_.Replace($pnpList.RootFolder.ServerRelativeUrl,"")
            if($libraryRelativePath.Substring(0,1) -eq "/"){$libraryRelativePath = $libraryRelativePath.Substring(1,$libraryRelativePath.Length-1)} #Trim any leading "/"
            if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Add-PnPDocumentSet -List $($pnpList.Id) [$($pnpList.Title)] -Name [$libraryRelativePath] -ContentType ""Document Set"""}
            $newFolderUrl = Add-PnPDocumentSet -List $pnpList.Id -Name $libraryRelativePath -ContentType "Document Set"
            }
        $newFolder = get-spoFolder -pnpList $pnpList -folderServerRelativeUrl $newFolderUrl.Replace("https://anthesisllc.sharepoint.com","") -adminCreds $spoCredentials -verboseLogging $verboseLogging
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
        $pnpTerm = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Identity $cleanTerm -ErrorAction Stop #Weirdly, Get-PnPTerm throws a non-terminating exception if the Term isn't found. We want an exception, so that catch{} returns $null value
        #$alreadyInStore = Get-PnPTaxonomyItem -TermPath "$termGroup|$termSet|$term"
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
function copy-spoFile($fromList,$from,$to,$spoCredentials){
    if($verboseLogging){Write-Host -ForegroundColor Magenta "copy-spoFile($fromList,$from,$to"}
    if($fromList.Substring(0,1) -ne "/"){$fromList = "/"+$fromList}
    if($(Split-Path $from -Leaf) -eq $(Split-Path $to -Leaf)){$to = $to.SubString(0,$($to.Length - $(split-path $to -leaf).Length) -1)} #Specififying a file name is broken for (presumably) Sites with large numbers of Libraries/Files
<#    $oldConnection = Get-PnPConnection
    if($oldConnection.Url -ne "https://anthesisllc.sharepoint.com$fromList"){
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Connected to wrong site - connecting to https://anthesisllc.sharepoint.com$fromList"}
        Connect-PnPOnline –Url $("https://anthesisllc.sharepoint.com$fromList") –Credentials $spoCredentials
        }#>
    if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Copy-PnPFile -SourceUrl $from -TargetUrl $to -Force"}
    Copy-PnPFile -SourceUrl $from -TargetUrl $to -force
    if($oldConnection){
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Reconnecting to $($oldConnection.Url)"}
        Connect-PnPOnline -Url $oldConnection.Url -Credentials $spoCredentials
        }
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
            if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Trying to retrieve Library by Client Name: Get-PnPList -Identity $($clientName)"}
            try{$thisClientLibrary = Get-PnPList -Identity $(sanitise-forSql $clientName)}
            catch{<#Meh.#>}
            }
        $thisClientLibrary
        }
    catch{log-error -myError $_ -myFriendlyMessage "Error retrieving Client Library in get-spoClientLibrary" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName}
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
        if($pnpFolder){
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
function get-spoKimbleClientListItems($spoCredentials, $verboseLogging){
    if($(Get-PnPConnection).Url -ne "https://anthesisllc.sharepoint.com/clients"){
        Connect-PnPOnline –Url $("https://anthesisllc.sharepoint.com/clients") –Credentials $spoCredentials
        }
    if($verboseLogging){Write-Host -ForegroundColor Magenta 'Get-PnPListItem -List "Kimble Clients" -PageSize 5000 -Fields "Title","GUID","KimbleId","ClientDescription","IsDirty","IsDeleted","Modified","LastModifiedDate","PreviousName","PreviousDescription","Id","LibraryGUID"'}
    $clientListItems = Get-PnPListItem -List "Kimble Clients" -PageSize 5000 -Fields "Title","GUID","KimbleId","ClientDescription","IsDirty","IsDeleted","Modified","LastModifiedDate","PreviousName","PreviousDescription","Id","LibraryGUID"
    $clientListItems.FieldValues | %{
        $thisClient = $_
        [array]$allSpoClients += New-Object psobject -Property $([ordered]@{"Id"=$thisClient["KimbleId"];"Name"=$thisClient["Title"];"GUID"=$thisClient["GUID"];"SPListItemID"=$thisClient["ID"];"IsDirty"=$thisClient["IsDirty"];"IsDeleted"=$thisClient["IsDeleted"];"LastModifiedDate"=$thisClient["LastModifiedDate"];"PreviousName"=$thisClient["PreviousName"];"ClientDescription"=$(sanitise-stripHtml $thisClient["ClientDescription"]);"PreviousDescription"=$thisClient["PreviousDescription"];"LibraryGUID"=$thisClient["LibraryGUID"]})
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
function new-spoClientLibrary($clientName, $clientDescription, $spoCredentials, $verboseLogging){
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
function update-spoTerm($termGroup,$termSet,$oldTerm,$newTerm,$kimbleId,$verboseLogging){
     if($verboseLogging){Write-Host -ForegroundColor Magenta "update-spoTerm($termGroup,$termSet,$oldTerm,$newTerm)"}
     $cleanOldTerm = $(sanitise-forTermStore $oldTerm)
     $cleanNewTerm = $(sanitise-forTermStore $newTerm)
    try{
        $pnpTermGroup = Get-PnPTermGroup $termGroup 
        $pnpTermSet = Get-PnPTermSet -TermGroup $pnpTermGroup -Identity $termSet
        $pnpOldTerm = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Identity $cleanOldTerm
        if(!$pnpOldTerm){Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Identity $oldTerm} #Try the dirty version if we can't find the clean version
        $pnpNewTerm = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Identity $cleanNewTerm
        #$alreadyInStore = Get-PnPTaxonomyItem -TermPath "$termGroup|$termSet|$term"
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