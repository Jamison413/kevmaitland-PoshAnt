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
function new-spoKimbleClientItem($kimbleClientObject, $pnpClientList, $fullLogPathAndName,$verboseLogging){
    #Create the new List item
    if($verboseLogging){Write-Host -ForegroundColor Magenta "new-spoKimbleClientItem($($kimbleClientObject.Name), $($pnpClientList.Title)"}
    #Check that PNP is connected to Clients Site
    #Check that the list is valid
    #Get the Content Type
    $contentType = $pnpClientList.ContentTypes | ? {$_.Name -eq "Item"}
    $updateValues = @{"Title"=$kimbleClientObject.Name;"KimbleId"=$kimbleClientObject.Id;"ClientDescription"=$(sanitise-stripHtml $kimbleClientObject.Description);"IsDirty"=$true;"IsDeleted"=$kimbleClientObject.IsDeleted}
    if($kimbleClientObject.LastModifiedDate){$updateValues.Add("LastModifiedDate",$(Get-Date $kimbleClientObject.LastModifiedDate -Format "MM/dd/yyyy hh:mm"))}
    try{
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "`tAdd-PnPListItem -List $($pnpClientList.Title) -ContentType $($contentType.Id.StringValue) -Values @{$(stringify-hashTable $updateValues)}"}
        $newItem = Add-PnPListItem -List $pnpClientList.Id -ContentType $($contentType.Id.StringValue) -Values $updateValues
        }
    catch{
        log-error -myError $_ -myFriendlyMessage "Error creating new [Kimble Clients] list item [$($kimbleClientObject.Name)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
        }
    if($newItem){if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "`tSUCCESS: Item [$($kimbleClientObject.Name)] created in [Kimble Clients]"}}
    else{Write-Host -ForegroundColor DarkRed "`tFAILED: Item NOT [$($kimbleClientObject.Name)] created in [Kimble Clients] :("}
    $newItem
    }
function new-spoKimbleProjectItem($kimbleProjectObject, $pnpProjectList, $fullLogPathAndName, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Magenta "new-spoKimbleProjectItem($($kimbleProjectObject.Name), $($pnpProjectList.Title)"}
    $contentType = $pnpProjectList.ContentTypes | ? {$_.Name -eq "Item"}
    $updateValues = @{"Title"=$kimbleProjectObject.Name;"KimbleId"=$kimbleProjectObject.Id;"KimbleClientId"=$kimbleProjectObject.KimbleOne__Account__c;"IsDirty"=$true;"IsDeleted"=$kimbleProjectObject.IsDeleted}
    if($kimbleProjectObject.LastModifiedDate){$updateValues.Add("LastModifiedDate",$(Get-Date $kimbleProjectObject.LastModifiedDate -Format "MM/dd/yyyy hh:mm"))}
    try{
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "`tAdd-PnPListItem -List $($pnpProjectList.Id) -ContentType $($contentType.Id.StringValue) -Values @{$(stringify-hashTable $updateValues)}"}
        $newItem = Add-PnPListItem -List $pnpProjectList.Id -ContentType $contentType.Id.StringValue -Values $updateValues
        }
    catch{
        log-error -myError $_ -myFriendlyMessage "Error creating new [Kimble Projects] list item [$($kimbleProjectObject.Name)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
        }
    if($newItem){if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "`tSUCCESS: Item [$($kimbleProjectObject.Name)] created in [Kimble Projects]"}}
    else{Write-Host -ForegroundColor DarkRed "`tFAILED: Item NOT [$($kimbleProjectObject.Name)] created in [Kimble Projects] :("}
    $newItem
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
function set-standardTeamSitePermissions($teamSiteAbsoluteUrl, $adminCredentials, $verboseLogging){
    #$teamSiteAbsoluteUrl = "https://anthesisllc.sharepoint.com/teams/Energy_Engineering_Team_All_365/"
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

        #Add Managers group to Site Coll Admins & Site Owners Group
        $guessedManagerGroupName = get-managersGroupNameFromTeamUrl -teamSiteUrl $teamSiteAbsoluteUrl
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "get-managersGroupNameFromTeamUrl -teamSiteUrl $teamSiteAbsoluteUrl = [$guessedManagerGroupName]"}
        if($guessedManagerGroupName){
            if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Getting OwnersGroup: Get-PnPGroup -AssociatedOwnerGroup"}
            $ownersSpoGroup = Get-PnPGroup -AssociatedOwnerGroup
            $managersGroup = Get-PnPUser | ? {$_.Email -eq $($guessedManagerGroupName+"@anthesisgroup.com")}
            if(!$managersGroup){$managersGroup = New-PnPUser -LoginName $($guessedManagerGroupName+"@anthesisgroup.com")}
            if($managersGroup){
                if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Add-PnPUserToGroup -EmailAddress $($managersGroup.Email) -Identity $($ownersSpoGroup.Title) -SendEmail:$false"}
                Add-PnPUserToGroup -EmailAddress $managersGroup.Email -Identity $ownersSpoGroup.Id -SendEmail:$false
                #if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Set-PnPTenantSite -Url $teamSiteAbsoluteUrl -Owners $($guessedManagerGroupName+"@anthesisgroup.com")"}
                Set-PnPTenantSite -Url $teamSiteAbsoluteUrl -Owners $managersGroup.LoginName
                if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Set-PnPTenantSite -Url $teamSiteAbsoluteUrl -Owners 'kimblebot@anthesisgroup.com'"}
                Set-PnPTenantSite -Url $teamSiteAbsoluteUrl -Owners "kimblebot@anthesisgroup.com"
                }
            else{Write-Warning "No Managers group could be guessed for [$teamSiteAbsoluteUrl] - it cannot be added to the Site Owners Group, nor as a Site Collection Admin"}
            }
        #Check the Site Collection Administrators
        $siteCollectionAdmins = Get-PnPSiteCollectionAdmin
        $o365Group = $siteCollectionAdmins | ? {$_.Email -eq $((Split-Path $teamSiteAbsoluteUrl -Leaf)+"@anthesisgroup.com")}
        if($o365Group){
            #Remove O365 Group from Site Collection Admins
            if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Remove-PnPSiteCollectionAdmin -Owners $((Split-Path $teamSiteAbsoluteUrl -Leaf)+"@anthesisgroup.com")"}
            Remove-PnPSiteCollectionAdmin -Owners $o365Group.LoginName
            }
        if($siteCollectionAdmins.Title -notcontains "Kimble Bot"){Write-Warning "KimbleBot is not a Site Collection Administrator"}
        if($siteCollectionAdmins.Email -notcontains $managersGroup.Email){
            if($managersGroup){Write-Warning "$($managersGroup.Title) was not added as a Site Collection Administrator"}
            }

        #Block all external sharing
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Blocking external Sharing: Set-PnPTenantSite -Url $teamSiteAbsoluteUrl -Sharing Disabled"}
        Set-PnPTenantSite -Url $teamSiteAbsoluteUrl -Sharing Disabled
        
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
        else{Remove-PnPUserFromGroup -LoginName (Get-PnPConnection).PSCredential.UserName -Identity $ownersSpoGroup.Id}

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
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "All finished"}
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
function update-spoKimbleClientItem($kimbleClientObject, $pnpClientList, $fullLogPathAndName,$verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Magenta "update-spoKimbleClientItem($($kimbleClientObject.Name), $($pnpClientList.Title)"}
    #Bodge the KimbeId value if it's not present
    if([string]::IsNullOrWhiteSpace($kimbleClientObject.KimbleId) -and $kimbleClientObject.Id.Length -eq 18){
        $kimbleClientObject | Add-Member -MemberType NoteProperty -Name KimbleId -Value $kimbleClientObject.Id
        }
    #Retrieve the existing item
    try{
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Get-PnPListItem -List $($pnpClientList.Title) -Query <View><Query><Where><Eq><FieldRef Name='KimbleId'/><Value Type='Text'>$($kimbleClientObject.KimbleId)</Value></Eq></Where></Query></View> -ErrorAction Stop"}
        $existingItem = Get-PnPListItem -List $pnpClientList -Query "<View><Query><Where><Eq><FieldRef Name='KimbleId'/><Value Type='Text'>$($kimbleClientObject.KimbleId)</Value></Eq></Where></Query></View>" -ErrorAction Stop
        }
    catch{
        log-error -myError $_ -myFriendlyMessage "Error retrieving [Kimble Clients] list item [$($kimbleClientObject.Name)] in update-spoKimbleClientItem()" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
        }

    #Update it
    if(!$existingItem){
        if($verboseLogging){Write-Host -ForegroundColor DarkRed "`tFAILED: Existing item [$($kimbleClientObject.Name)] in [Kimble Clients] not found"}
        }
    else{
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "`tSUCCESS: Existing item [$($existingItem.FieldValues.Title)] in [Kimble Clients] found"}
        $updateValues = @{"Title"=$kimbleClientObject.Name;"KimbleId"=$kimbleClientObject.Id;"ClientDescription"=$(sanitise-stripHtml $kimbleClientObject.Description);"IsDirty"=$true;"IsDeleted"=$kimbleClientObject.IsDeleted}
        if($kimbleClientObject.LastModifiedDate){$updateValues.Add("LastModifiedDate",$(Get-Date $kimbleClientObject.LastModifiedDate -Format "MM/dd/yyyy hh:mm"))}
        try{
            if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Set-PnPListItem -List $($pnpClientList.Id) -Identity $($existingItem.Id) -Values @{$(stringify-hashTable $updateValues)}"}
            $updatedItem = Set-PnPListItem -List $pnpClientList.Id -Identity $existingItem.Id -Values $updateValues -ErrorAction Stop
            }
        catch{
            log-error -myError $_ -myFriendlyMessage "Error updating [Kimble Clients] list item [$($existingItem.FieldValues.Title)] in update-spoKimbleClientItem()" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
            }
        if($updatedItem){if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "`tSUCCESS: Item [$($updatedItem.FieldValues.Title)] updated in [Kimble Clients]"}}
        else{Write-Host -ForegroundColor DarkRed "`tFAILED: Item [$($existingItem.FieldValues.Title)] NOT updated in [Kimble Clients], even though I found it to update :("}
        }
    #Return it
    $updatedItem
    }
function update-spoKimbleProjectItem($kimbleProjectObject, $pnpProjectList, $fullLogPathAndName, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Magenta "update-spoKimbleProjectItem($($kimbleProjectObject.Name), $($pnpProjectList.Title)"}
    #Bodge the KimbleId value if it's not present
    if([string]::IsNullOrWhiteSpace($kimbleProjectObject.KimbleId) -and $kimbleProjectObject.Id.Length -eq 18){
        $kimbleProjectObject | Add-Member -MemberType NoteProperty -Name KimbleId -Value $kimbleProjectObject.Id
        }
    #Retrieve the existing item
    try{
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Get-PnPListItem -List $($pnpProjectList.Title) -Query <View><Query><Where><Eq><FieldRef Name='KimbleId'/><Value Type='Text'>$($kimbleProjectObject.KimbleId)</Value></Eq></Where></Query></View> -ErrorAction Stop"}
        $existingItem = Get-PnPListItem -List $pnpProjectList -Query "<View><Query><Where><Eq><FieldRef Name='KimbleId'/><Value Type='Text'>$($kimbleProjectObject.KimbleId)</Value></Eq></Where></Query></View>" -ErrorAction Stop
        }
    catch{
        log-error -myError $_ -myFriendlyMessage "Error retrieving [Kimble Projects] list item [$($kimbleProjectObject.Name)] in update-spoKimbleProjectItem()" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
        }

    #Update it
    if(!$existingItem){
        if($verboseLogging){Write-Host -ForegroundColor DarkRed "`tFAILED: Existing item [$($kimbleProjectObject.Name)] in [Kimble Projects] not found"}
        }
    else{
        if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "`tSUCCESS: Existing item [$($existingItem.FieldValues.Title)] in [Kimble Clients] found"}
        $updateValues = @{"Title"=$kimbleProjectObject.Name;"KimbleId"=$kimbleProjectObject.Id;"KimbleClientId"=$kimbleProjectObject.KimbleOne__Account__c;"IsDirty"=$true;"IsDeleted"=$kimbleProjectObject.IsDeleted}
        if($kimbleProjectObject.LastModifiedDate){$updateValues.Add("LastModifiedDate",$(Get-Date $kimbleProjectObject.LastModifiedDate -Format "MM/dd/yyyy hh:mm"))}
        try{
            if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "Set-PnPListItem -List $($pnpProjectList.Id) -Identity $($existingItem.Id) -Values @{$(stringify-hashTable $updateValues)}"}
            $updatedItem = Set-PnPListItem -List $pnpProjectList.Id -Identity $existingItem.Id -Values $updateValues -ErrorAction Stop
            }
        catch{
            log-error -myError $_ -myFriendlyMessage "Error updating [Kimble Projects] list item [$($existingItem.FieldValues.Title)] in update-spoKimbleProjectItem()" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
            }
        if($updatedItem){if($verboseLogging){Write-Host -ForegroundColor DarkMagenta "`tSUCCESS: Item [$($updatedItem.FieldValues.Title)] updated in [Kimble Projects]"}}
        else{Write-Host -ForegroundColor DarkRed "`tFAILED: Item [$($existingItem.FieldValues.Title)] NOT updated in [Kimble Projects], even though I found it to update :("}
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