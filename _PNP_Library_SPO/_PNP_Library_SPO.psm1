$spoRoot = "https://anthesisllc.sharepoint.com"
$spoAdminRoot = "https://anthesisllc-admin.sharepoint.com"
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
function copy-spoFile(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [System.Uri]$sourceUrl
        ,[parameter(Mandatory = $true)]
            [System.Uri]$destinationSite
        ,[parameter(Mandatory = $false)]
            [System.Uri]$destinationFolder
        ,[parameter(Mandatory = $true)]
            [pscredential]$pnpCreds
        ,[parameter(Mandatory = $false)]
            [switch]$overwriteDestinationFile
        ,[parameter(Mandatory = $false)]
            [string]$renameFileAs
        )
    
    Write-Verbose "Downloading source file [$($sourceUrl.LocalPath)]"
    try{Get-PnPFile -Url $sourceUrl.LocalPath -Path "$env:TEMP" -Filename $([uri]::UnescapeDataString($(Split-Path $sourceUrl.AbsoluteUri -Leaf))) -AsFile -Force}
    catch{Write-Error "Error retrieving file [$($sourceUrl.LocalPath)] using Get-PnpFile in copy-spoFile";break}
    try{Connect-PnPOnline -Url $destinationSite.AbsoluteUri -Credentials $pnpCreds}
    catch{Write-Error "Error connecting to Site [$($destinationSite.AbsoluteUri)] using Connect-PnPOnline in copy-spoFile";break}
    if(test-pnpConnectionMatchesResource -resourceUrl $destinationSite.AbsoluteUri){
        try{
            if([string]::IsNullOrWhiteSpace($renameFileAs)){
                $file = Add-PnPFile -Path "$env:TEMP\$([uri]::UnescapeDataString($(Split-Path $sourceUrl.AbsoluteUri -Leaf)))" -Folder $destinationFolder -ErrorAction Stop #Added '$file = ' to avoid https://github.com/SharePoint/PnP-PowerShell/issues/722
                }
            else{$file = Add-PnPFile -Path "$env:TEMP\$([uri]::UnescapeDataString($(Split-Path $sourceUrl.AbsoluteUri -Leaf)))" -Folder $destinationFolder -ErrorAction Stop -NewFileName $renameFileAs} #Added '$file = ' to avoid https://github.com/SharePoint/PnP-PowerShell/issues/722
            }
        catch{Write-Error "Error uploading file [$([uri]::UnescapeDataString($(Split-Path $sourceUrl.AbsoluteUri -Leaf)))] to Site [$($destinationSite.AbsoluteUri)] using Add-PnPFile in copy-spoFile";break}
        }
    }
function copy-spoPage(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
        [ValidatePattern(".[SitePages].")]
        [System.Uri]$sourceUrl = "https://anthesisllc.sharepoint.com/sites/Resources-IT/SitePages/Candidate-Template-for-Global-Sites.aspx"

        ,[parameter(Mandatory = $true)]
        [System.Uri]$destinationSite

        ,[parameter(Mandatory = $true)]
        [pscredential]$pnpCreds

        ,[parameter(Mandatory = $false)]
        [bool]$overwriteDestinationFile = $false
        
        ,[parameter(Mandatory = $false)]
        [string]$renameFileAs
        )
    Write-Verbose "copy-spoPage($($sourceUrl),$($destinationSite))"
    $dirtyBodgeToGetSourceSite = $sourceUrl.Scheme+"://"+$sourceUrl.DnsSafeHost
    #$sourceUrl.Segments | %{ #break not supported in pipeline
    foreach ($segment in $sourceUrl.Segments ){
        if($segment -match "SitePages"){break}
        $dirtyBodgeToGetSourceSite += $segment
        }
    Write-Verbose "`$dirtyBodgeToGetSourceSite = $dirtyBodgeToGetSourceSite"
    
    $dirtyBodgeToGetDestinationSite = $destinationSite.Scheme+"://"+$destinationSite.DnsSafeHost
    foreach ($segment in $destinationSite.Segments){
        if($segment -match "SitePages"){break}
        $dirtyBodgeToGetDestinationSite += $segment
        }
    Write-Verbose "`$dirtyBodgeToGetDestinationSite = $dirtyBodgeToGetDestinationSite"

    try{
        if (test-pnpConnectionMatchesResource -resourceUrl $dirtyBodgeToGetSourceSite -connectIfDifferent $true -pnpCreds $pnpCreds){Write-Verbose "Already connected to source Site [$($dirtyBodgeToGetSourceSite)]"}
        try{
            Write-Verbose "Downloading source Page file [$($sourceUrl.LocalPath)]"
            Get-PnPFile -Url $sourceUrl.LocalPath -Path "$env:TEMP" -Filename $([uri]::UnescapeDataString($(Split-Path $sourceUrl.AbsoluteUri -Leaf))) -AsFile -Force
            try{
                Write-Verbose "Connecting to SPO Admin [https://anthesisllc-admin.sharepoint.com/] (same creds [$($pnpCreds.UserName)], but different permissions required)"
                Connect-SPOService -Url https://anthesisllc-admin.sharepoint.com/ -Credential $pnpCreds
                try{
                    Write-Verbose "Allowing upload of .aspx files to destination [$($destinationSite.AbsoluteUri.TrimEnd("/"))]"
                    Set-SPOSite -Identity $destinationSite.AbsoluteUri.TrimEnd("/") -DenyAddAndCustomizePages $false -ErrorAction Stop
                    try{
                        Write-Verbose "Uploading file to [$($destinationSite.AbsoluteUri+"/SitePages/"+$(Split-Path $sourceUrl.AbsoluteUri -Leaf))]"
                        Connect-PnPOnline -Url $destinationSite.AbsoluteUri -Credentials $pnpCreds
                        if([string]::IsNullOrWhiteSpace($renameFileAs)){
                            $file = Add-PnPFile -Path "$env:TEMP\$(Split-Path $sourceUrl.AbsoluteUri -Leaf)" -Folder "SitePages" -ErrorAction Stop #Added '$file = ' to avoid https://github.com/SharePoint/PnP-PowerShell/issues/722
                            }
                        else{$file = Add-PnPFile -Path "$env:TEMP\$(Split-Path $sourceUrl.AbsoluteUri -Leaf)" -Folder "SitePages" -ErrorAction Stop -NewFileName $renameFileAs} #Added '$file = ' to avoid https://github.com/SharePoint/PnP-PowerShell/issues/722
                        
                        try{
                            Write-Verbose "Disabling upload of .aspx files to destination [$($destinationSite.AbsoluteUri.TrimEnd("/"))]"
                            Set-SPOSite -Identity $destinationSite.AbsoluteUri.TrimEnd("/") -DenyAddAndCustomizePages $true -ErrorAction Stop
                            }
                        catch{
                            Write-Error "Failed to re-allow upload of .aspx files to Destination SitePages Lib [$($destinationSite.AbsoluteUri)]"
                            }
                        }
                    catch{
                        Write-Error "Failed to upload file to destination [$($destinationSite.AbsoluteUri+"/SitePages/"+$(Split-Path $sourceUrl.AbsoluteUri -Leaf))]"
                        }
                    }
                catch{
                    Write-Error "Could not enable upload of .aspx files to destination site [[$($destinationSite.AbsoluteUri)]]"
                    }
                }
            catch{
                Write-Error "Failed to connect to [https://anthesisllc-admin.sharepoint.com/]"
                }
            }
        catch{
             Write-Error "Failed to download source file [$($sourceUrl.LocalPath)]"
            }
        }
    catch{
        Write-Error "Could not connect to Source Site via PNP [$dirtyBodgeToGetSourceSite]"
        }
    
    }
function delete-versionHistory(){
     [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [System.Uri]$fileUrl = "https://anthesisllc.sharepoint.com/:p:/r/clients/Anthesis%20LLC/201908_Anthesis%20LLC_CE%20Centre%20CL%20Partners%20Chemical%20Recycling%20Screening%20(E005835)/Proposal/Archive/Anthesis%20Closed%20Loop%20RfP%20Addendum%20230919%20workstream%204.pptx?d=w3826c547c6db4e04bb6565b30b840395&csf=1&web=1&e=bg33ew"
        ,[parameter(Mandatory = $true)]
            [pscredential]$pnpCreds
        ,[parameter(Mandatory = $true)]
            [int]$numberOfPreviousVersionsToKeep = 100
        )

    Write-Verbose "delete-versionHistory($($fileUrl),$($numberOfPreviousVersionsToKeep))"
    $file = Get-PnPFile -Url $([uri]::UnescapeDataString($fileUrl.AbsolutePath.Replace("/:p:/r","")))
    $versions = $file.Versions
    $file.Context.Load($versions)
    $file.Context.ExecuteQuery()

    $versionsToDelete = $versions.Count - $numberOfPreviousVersionsToKeep
    Write-Verbose "`t[$($versions.Count)] versions exist, [$($numberOfPreviousVersionsToKeep)] to be retained .: deleting oldest [$($versionsToDelete)]"
    for($i=0; $i -lt $versionsToDelete; $i++){
        Write-Verbose "`tMarking version [$($versions[0].VersionLabel)] for deletion"
        $versions[0].DeleteObject() #$versions[0] is always the oldest version, even before we .ExecuteQuery()
        }
    
    try{
        $file.Context.ExecuteQuery()
        Write-Verbose "`t[$($versionsToDelete)] versions deleted successfully"
        }
    catch{$_}

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
                $thisDocumentLibrary = Get-PnPList -Identity $($docLibGuid) -Includes Description
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
    if([string]::IsNullOrWhiteSpace($pnpListItem) -and ![string]::IsNullOrWhiteSpace($kimbleEngagementCodeToLookFor)){
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
function set-standardSitePermissions(){
    [cmdletbinding(SupportsShouldProcess=$true)]
    param(
        [parameter(Mandatory = $true,ParameterSetName="UnifiedGroupObject")]
            [parameter(Mandatory = $true,ParameterSetName="UnifiedGroupId")]
            [PSCustomObject]$tokenResponse
        ,[parameter(Mandatory = $true,ParameterSetName="UnifiedGroupObject")]
            [PSCustomObject]$graphGroupExtended
        ,[parameter(Mandatory = $true,ParameterSetName="UnifiedGroupId")]
            [string]$groupId
        ,[parameter(Mandatory = $true,ParameterSetName="UnifiedGroupObject")]
            [parameter(Mandatory = $true,ParameterSetName="UnifiedGroupId")]
            [pscredential]$pnpCreds
        ,[parameter(Mandatory = $false,ParameterSetName="UnifiedGroupObject")]
            [parameter(Mandatory = $false,ParameterSetName="UnifiedGroupId")]
            [switch]$suppressEmailNotifications
        )
    Write-Verbose "set-standardSitePermissions([$($graphGroupExtended.Id)$unifiedGroupId])"

    #Get $graphGroupExtended, regardless of which parameters we've been given
    switch ($PsCmdlet.ParameterSetName){
        “UnifiedGroupId”  {
            Write-Verbose "We've been given a 365 Id, so we need the Group object"
            $graphGroupExtended = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -filterId $unifiedGroupId
            if(!$graphGroupExtended){
                Write-Error "Could not retrieve Unified Group from ID [$unifiedGroupId]"
                break
                }
            }
        }
    try{$pnpUnifiedGroupObject = Get-PnPUnifiedGroup -Identity $graphGroupExtended.id -ErrorAction Stop -WarningAction Stop}
    catch{#Connect to the root site if we're not connected to anything
        Write-Verbose "Connecting to Graph"
        Connect-PnPOnline -Url "https://anthesisllc-admin.sharepoint.com/" -AccessToken $tokenResponse.access_token
        $pnpUnifiedGroupObject = Get-PnPUnifiedGroup -Identity $graphGroupExtended.id
        }

    if([string]::IsNullOrWhiteSpace([uri]::UnescapeDataString($pnpUnifiedGroupObject.SiteUrl))){ #This is a more reliable test than the UnifiedGroup.SharePointSiteUrl property as it populates /much/ faster
        Write-Error "Could not retrieve 365 Group URL from Group [$($graphGroupExtended.DisplayName)][$($graphGroupExtended.id)]. Exiting without attempting to check/set permissions"
        break
        }
    else{Write-Verbose "SiteUrl for [$($graphGroupExtended.DisplayName)] is [$([uri]::UnescapeDataString($pnpUnifiedGroupObject.SiteUrl))] via Get-PnpUnifiedGroup"}

    #region Get connected to the Site
    try{
        $userWasAlreadyASiteAdmin = test-isUserSiteCollectionAdmin -pnpUnifiedGroupObject $pnpUnifiedGroupObject -accessToken $tokenResponse.access_token -pnpCreds $pnpCreds -addPermissionsIfMissing $true -ErrorAction Stop -Verbose:$VerbosePreference
        }
    catch{
        Write-Verbose "Error connecting to [$([uri]::UnescapeDataString($pnpUnifiedGroupObject.SiteUrl))] - cannot continue"
        $_
        break
        }
    #endregion

    Connect-PnPOnline -Url $([uri]::UnescapeDataString($pnpUnifiedGroupObject.SiteUrl)) -Credentials $pnpCreds -ErrorAction Stop -WarningAction Stop
    if(test-pnpConnectionMatchesResource -resourceUrl $([uri]::UnescapeDataString($pnpUnifiedGroupObject.SiteUrl))){
        if([string]::IsNullOrWhiteSpace((Get-PnPFeature -Scope Site -Identity "b50e3104-6812-424f-a011-cc90e6327318"))){
            Write-Verbose "Enabling the DocID service"
            Enable-PnPFeature -Identity "b50e3104-6812-424f-a011-cc90e6327318" -Scope Site -Verbose:$VerbosePreference
            }

        Write-Verbose "Now set the Classification-specific Sharing settings"
        #First, set the UnifiedGroup Guest access settings
        set-graphUnifiedGroupGuestSettings -tokenResponse $tokenResponse -graphUnifiedGroupExtended $graphGroupExtended -Verbose:$VerbosePreference
        #Then set the corresponding SharePoint Site sharing settings
        Write-Verbose "`[uri]::UnescapeDataString($pnpUnifiedGroupObject.SiteUrl) = [$([uri]::UnescapeDataString($pnpUnifiedGroupObject.SiteUrl))]"
        Write-Verbose "`$graphGroupExtended.anthesisgroup_UGSync.classification = [$($graphGroupExtended.anthesisgroup_UGSync.classification)]"
        Write-Verbose Get-PnPConnection
#        if([uri]::UnescapeDataString($pnpUnifiedGroupObject.SiteUrl) -notmatch "&" -and [uri]::UnescapeDataString($pnpUnifiedGroupObject.SiteUrl) -notmatch "%26"){ #There's a bug where URLs containing ampersands don't process correctly via a Scheduled Task. Must be something to do with encoding, but can't figure it out and have run out of time :'(
            switch($graphGroupExtended.anthesisgroup_UGSync.classification){
                    "External" {
                        Write-Verbose "Set-PnPSite -DisableSharingForNonOwners:$true -Sharing ExternalUserAndGuestSharing -Verbose:$VerbosePreference"
                        Set-PnPSite -DisableSharingForNonOwners:$true -Sharing ExternalUserAndGuestSharing -Verbose:$VerbosePreference #Allow external sharing #Unescape because Set-PnPSite doesn't like %26 in place of &
                        }
                    "Internal" {
                        Write-Verbose "Set-PnPSite -DisableSharingForNonOwners:$true -Sharing Disabled -Verbose:$VerbosePreference"
                        Set-PnPSite -DisableSharingForNonOwners:$true -Sharing Disabled -Verbose # :$VerbosePreference #Block all external sharing
                        }
                    "Confidential" {
                        Write-Verbose "Set-PnPSite -DisableSharingForNonOwners:$true -Sharing Disabled -Verbose:$VerbosePreference"
                        Set-PnPSite -DisableSharingForNonOwners:$true -Sharing Disabled -Verbose:$VerbosePreference #Block all external sharing
                        }
                    }
 #           }

        Write-Verbose "Now set Access Requests to go to default Owners Group (Set-PnPSite above seems to break this)"
        $pnpWeb = Get-PnPWeb -Includes RequestAccessEmail
        $pnpWeb.Context.Web.SetUseAccessRequestDefaultAndUpdate($true)
        $pnpWeb.Context.ExecuteQuery()

        #Reset any changes made to managed properties
        Write-Verbose "Reset any changes made to managed properties (pausing to allow AAD to catch up)"
        Start-Sleep -Seconds 5
        if($suppressEmailNotifications){reset-graphUnifiedGroupSettingsToOriginals -tokenResponse $tokenResponse -graphGroupExtended $graphGroupExtended -Verbose:$VerbosePreference -suppressEmailNotification}
        else{reset-graphUnifiedGroupSettingsToOriginals -tokenResponse $tokenResponse -graphGroupExtended $graphGroupExtended -Verbose:$VerbosePreference}

        Write-Verbose "Remove everything that isn't the 365 Group Owners object from Site Owners (it looks like adding the Data Managers AAD group has been deprecated to match the user-only membership behaviour of 365 Groups)"
        $spoOwnersGroup = Get-PnPGroup -AssociatedOwnerGroup
        [array]$unexpectedSiteOwners = $spoOwnersGroup.Users | ? {($_.LoginName -notmatch $unifiedGroupObject.ExternalDirectoryObjectId) -and ($_.LoginName -ne "SHAREPOINT\system")}
        if($unexpectedSiteOwners.Count -gt 0){
            #Remove Unexpected Site Owners
            $unexpectedSiteOwners | % {
                Write-Verbose "`tRemove-PnPUserFromGroup -LoginName $($_.LoginName) -Identity $($spoOwnersGroup.Id)"
                $dummy = Remove-PnPUserFromGroup -LoginName $_.LoginName -Identity $spoOwnersGroup.Id -Verbose:$VerbosePreference
                }
            }

        Write-Verbose "Remove everything that isn't the 365 Group Owners object from Site Collection Admins (except executing user otherwise we might saw off the branch we're sitting on) :)"
        $spoSiteCollectionAdmins = Get-PnPSiteCollectionAdmin
        [array]$unexpectedSiteCollectionAdmins = $spoSiteCollectionAdmins | ? {($_.LoginName -notmatch $unifiedGroupObject.ExternalDirectoryObjectId -and $_.LoginName -notmatch (Get-PnPConnection).PSCredential.UserName)}
        if($unexpectedSiteCollectionAdmins.Count -gt 0){
            #Remove Unexpected Site Admins
            $unexpectedSiteCollectionAdmins | % {
                Remove-PnPSiteCollectionAdmin -Owners $_.LoginName
                }
            }
        }
    else{Write-Warning "Cannot set security: Connected to incorrect Site [$((Get-PnPConnection).Url)]; expecting [$([uri]::UnescapeDataString($pnpUnifiedGroupObject.SiteUrl))]"}
    Write-Verbose "Finally, remove any owner/memberships we've temporarily granted ourselves"
    if(!$userWasAlreadyASiteAdmin){
        Remove-PnPSiteCollectionAdmin -Owners $($pnpCreds.UserName) -Verbose:$VerbosePreference
        }
    }
function test-isUserSiteCollectionAdmin(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true,ParameterSetName="pnpGroupObject")]
            [OfficeDevPnP.Core.Entities.UnifiedGroupEntity]$pnpUnifiedGroupObject
        ,[parameter(Mandatory = $true,ParameterSetName="UnifiedGroupId")]
            [string]$unifiedGroupId
        ,[parameter(Mandatory = $true,ParameterSetName="pnpSiteObject")]
            [Microsoft.SharePoint.Client.ClientObject]$pnpSiteObject
        ,[parameter(Mandatory = $true,ParameterSetName="SiteId")]
            [string]$pnpSiteId

        ,[parameter(Mandatory = $true,ParameterSetName="pnpGroupObject")]
        [parameter(Mandatory = $true,ParameterSetName="UnifiedGroupId")]
        [parameter(Mandatory = $true,ParameterSetName="pnpSiteObject")]
        [parameter(Mandatory = $true,ParameterSetName="SiteId")]
        [string]$accessToken = $true

        ,[parameter(Mandatory = $true,ParameterSetName="pnpGroupObject")]
        [parameter(Mandatory = $true,ParameterSetName="UnifiedGroupId")]
        [parameter(Mandatory = $true,ParameterSetName="pnpSiteObject")]
        [parameter(Mandatory = $true,ParameterSetName="SiteId")]
        [pscredential]$pnpCreds

        ,[parameter(Mandatory = $false,ParameterSetName="pnpGroupObject")]
        [parameter(Mandatory = $false,ParameterSetName="UnifiedGroupId")]
        [parameter(Mandatory = $true,ParameterSetName="pnpSiteObject")]
        [parameter(Mandatory = $true,ParameterSetName="SiteId")]
        [bool]$addPermissionsIfMissing = $false
        )

    Write-Verbose "test-isUserSiteCollectionAdmin [$($pnpUnifiedGroupObject.GroupId+$unifiedGroupId)]"
    #Get $unifiedGroupObject, regardless of which parameters we've been given
    switch ($PsCmdlet.ParameterSetName){
        “UnifiedGroupId”  {
            Write-Verbose "`ttest-isUserSiteCollectionAdmin | We've been given a 365 Id, so we need the PnPUnifiedGroup object"
            try{$pnpUnifiedGroupObject = Get-PnPUnifiedGroup -Identity $unifiedGroupId -ErrorAction Stop -WarningAction Stop}
            catch{#Connect to the root site if we're not connected to anything
                Connect-PnPOnline -Url "https://anthesisllc-admin.sharepoint.com/" -AccessToken $tokenResponse.access_token
                $pnpUnifiedGroupObject = Get-PnPUnifiedGroup -Identity $unifiedGroupId
                }
            if(!$pnpUnifiedGroupObject){
                Write-Error "`ttest-isUserSiteCollectionAdmin | Could not retrieve Unified Group from ID [$unifiedGroupId]"
                return
                }
            }
        {$_ -match "Group"} { #Catches both pnpGroupObject & UnifiedGroupId
             try{
                Write-Verbose "`ttest-isUserSiteCollectionAdmin | Checking to see if the executing user already has admin permissions for the Site"
                $pnpGroupAdmins = Get-PnPUnifiedGroupOwners -Identity $pnpUnifiedGroupObject.GroupId
                if($pnpGroupAdmins.UserPrincipalName -contains $pnpCreds.UserName){
                    $isAlreadyAnAdmin = $true
                    $isAlreadyAnAdmin
                    Write-Verbose "`ttest-isUserSiteCollectionAdmin | Yes - [$($pnpCreds.UserName)] is already an Owner, and therefore a Site Collection Admin"
                    return
                    }
                else{
                    Write-Warning "`ttest-isUserSiteCollectionAdmin | [$($pnpCreds.UserName)] is not an Owner, checking to see whether they are a Site Collection Admin ***This will break any existing Connect-PnPOnline sessions***"
                    #Bizarrely, there doesn;t seem to be a way of finding Site Collection Administrators without connecting to a Site within a Site Collection (which requires you to be a Site Collection Admin). You clearly /can/ do this because we can access this information via the Classic and Modern SharePoint 365 consoles (presumably via the -admin.sharepoint.com Site). Can't figure this out programatically though >:(
                    try{
                        Connect-PnPOnline -Url $([uri]::UnescapeDataString($pnpUnifiedGroupObject.SiteUrl)) -Credentials $pnpCreds -ErrorAction Stop -Verbose
                        $currentAdmins = Get-PnPSiteCollectionAdmin -ErrorAction Stop -Verbose
                        $isAlreadyAnAdmin = $true
                        $isAlreadyAnAdmin
                        Write-Verbose "`ttest-isUserSiteCollectionAdmin | Yes - [$($pnpCreds.UserName)] is already a Site Collection Admin, and this has been set explicitly"
                        return
                        }
                    catch{
                        #If the user is not a Site Collection Admin, Connect-PnPOnline will throw an error to here
                        Write-Verbose "`ttest-isUserSiteCollectionAdmin | No - [$($pnpCreds.UserName)] is not a Site Collection Admin"
                        $isAlreadyAnAdmin = $false
                        $isAlreadyAnAdmin
                        if($addPermissionsIfMissing){
                            Write-Verbose "`t`ttest-isUserSiteCollectionAdmin | Temporarily granting Site Collection Admin rights now for [$($pnpCreds.UserName)] to [$([uri]::UnescapeDataString($pnpUnifiedGroupObject.SiteUrl))]"
                            Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com" -Credentials $pnpCreds
                            Set-PnPTenantSite -Url $([uri]::UnescapeDataString($pnpUnifiedGroupObject.SiteUrl)) -Owners $pnpCreds.UserName -Verbose:$VerbosePreference
                            }
                        }
                    }
                }
            catch{
                Write-Error "Error connecting to [$([uri]::UnescapeDataString($pnpUnifiedGroupObject.SiteUrl))] - cannot continue"
                $_
                return
                }            
            }
        }



    }
function test-pnpConnectionMatchesResource(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
        [System.Uri]$resourceUrl = "https://anthesisllc.sharepoint.com"

        ,[parameter(Mandatory = $false)]
        [bool]$connectIfDifferent = $false

        ,[parameter(Mandatory = $false)]
        [pscredential]$pnpCreds
        )
    Write-Verbose "test-pnpConnectionMatchesResource($resourceUrl, $($pnpCreds.UserName)"
    try{
        Get-PnPConnection | Out-Null
        if((Get-PnPConnection).Url -eq $resourceUrl){
            Write-Verbose "Connect-PnPOnline connection matches [$resourceUrl]"
            return $true
            break #To avoid reconnecting and changing context later
            }
        else{
            Write-Verbose "Connect-PnPOnline connection [$([System.Uri](Get-PnPConnection).Url))] does not match [$resourceUrl]"
            $false
            }
        }
    catch{
        Write-Verbose "No Connect-PnPOnline connection available."
        }

    if($connectIfDifferent){
        Write-Verbose "Creating new Connect-PnpOnline to [$resourceUrl]"
        if($pnpCreds){
            try{Connect-PnPOnline -Url $resourceUrl -Credentials $pnpCreds}
            catch{Write-Error $_}
            }
        else{
            try{Connect-PnPOnline -Url $resourceUrl -CurrentCredentials}
            catch{Write-Error $_}
            }
        }
    }
function test-pnpConnectionMatchesResource_deprecated($resourceUrl, $verboseLogging){
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
    Write-Verbose "update-spoDocumentLibraryAndSubfoldersFromSqlKimbleListItem [$($sqlKimbleAccount.Name)] - looking for existing Library"
    log-action "update-spoDocumentLibraryAndSubfoldersFromSqlKimbleListItem [$($sqlKimbleAccount.Name)] - looking for existing Library" -logFile $fullLogPathAndName
    try{
        $duration = Measure-Command {
            #Try to get the Document Library by GUID (most accurate), then by PreviousName (next most likely), then by Name (least likely)
            Write-Verbose "`t`$existingLibrary = get-spoDocumentLibrary -docLibName [$($sqlKimbleAccount.PreviousName)] -docLibGuid [$($sqlKimbleAccount.DocumentLibraryGuid)]"
            $existingLibrary = get-spoDocumentLibrary -docLibName $sqlKimbleAccount.PreviousName -docLibGuid $sqlKimbleAccount.DocumentLibraryGuid
            if(!$existingLibrary){
                Write-Verbose "`t`$existingLibrary = get-spoDocumentLibrary -docLibName [$($sqlKimbleAccount.Name)]"
                $existingLibrary = get-spoDocumentLibrary -docLibName $sqlKimbleAccount.Name
                }
            }
        }
    catch{
        Write-Verbose "Error retrieving Document Library in update-spoDocumentLibraryAndSubfoldersFromSqlKimbleListItem [$($pnpListItem.Name)][$($pnpListItem.DocumentLibraryGuid)] `r`n$($Error[0].Exception.Response)`r`n$($Error[0].Exception.InnerException.Response)"
        Write-Error "Error retrieving Document Library in update-spoDocumentLibraryAndSubfoldersFromSqlKimbleListItem [$($pnpListItem.Name)][$($pnpListItem.DocumentLibraryGuid)] `r`n$($Error[0].Exception.Response)`r`n$($Error[0].Exception.InnerException.Response)"
        log-error -myError $_ -myFriendlyMessage "Error retrieving Document Library in update-spoDocumentLibraryAndSubfoldersFromSqlKimbleListItem [$($pnpListItem.Name)][$($pnpListItem.DocumentLibraryGuid)] $($Error[0].Exception.InnerException.Response)" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
        }

    if($existingLibrary){
        Write-Verbose "`tSUCCESS: [$($existingLibrary.RootFolder.ServerRelativeUrl)] found (GUID:[$($existingLibrary.Id.Guid)] [$($duration.TotalSeconds) seconds])"
        log-result -myMessage "SUCCESS: [$($existingLibrary.RootFolder.ServerRelativeUrl)] found (GUID:[$($existingLibrary.Id.Guid)] [$($duration.TotalSeconds) seconds])" -logFile $fullLogPathAndName
        Write-Verbose "Updating Document Library [$($existingLibrary.RootFolder.ServerRelativeUrl)]"
        log-action -myMessage "Updating Document Library [$($existingLibrary.RootFolder.ServerRelativeUrl)]" -logFile $fullLogPathAndName

        try{
            #Update the Library
            Write-Verbose "Updating Library [$($existingLibrary.RootFolder.ServerRelativeUrl)] Description:[$(sanitise-stripHtml -dirtyString $sqlKimbleAccount.Description)]"
            $duration = Measure-Command {
                $existingLibrary.Description = $(sanitise-stripHtml $sqlKimbleAccount.Description)
                $existingLibrary.Update()
                $existingLibrary.Context.ExecuteQuery()
                Write-Verbose "Updating Library [$($existingLibrary.RootFolder.ServerRelativeUrl)] Title:[$($sqlKimbleAccount.Name)]: `r`nSet-PnPList -Identity [$($existingLibrary.Id.Guid)] -Title [$($sqlKimbleAccount.Name)]"
                Set-PnPList -Identity $existingLibrary.Id -Title $sqlKimbleAccount.Name
                Write-Verbose "`$updatedLibrary = Get-PnPList -Identity [$($existingLibrary.Id)]"
                $updatedLibrary = get-spoDocumentLibrary -docLibGuid $existingLibrary.Id #The Id property is constant between $existingLibrary and $updatedLibrary 
                }
            #Check the update worked
            if($updatedLibrary.Title -eq $sqlKimbleAccount.Name -and $(sanitise-stripHtml $updatedLibrary.Description) -eq $(sanitise-stripHtml $sqlKimbleAccount.Description)){
                Write-Verbose "`tSUCCESS: Client Library [$($existingLibrary.RootFolder.ServerRelativeUrl)] updated successfully [$($duration.TotalSeconds) secs]"
                log-result -myMessage "SUCCESS: Client Library [$($existingLibrary.RootFolder.ServerRelativeUrl)] updated successfully [$($duration.TotalSeconds) secs]" -logFile $fullLogPathAndName
                $updatedLibrary
                }
            else{
                Write-Verbose "`tFAILED: Client Library not updated as expected!"
                if($updatedLibrary.Title -eq $sqlKimbleAccount.Name){
                    Write-Verbose "Client Name [$($updatedLibrary.Title)] matches correctly"
                    }
                else{
                    Write-Verbose "Original Client Name [$($sqlKimbleAccount.Name)] does not match [$($updatedLibrary.Title)]"
                    }
                if($(sanitise-stripHtml $updatedLibrary.Description) -eq $(sanitise-stripHtml $sqlKimbleAccount.Description)){
                    Write-Verbose "Client Description [$($updatedLibrary.Description)] matches correctly"
                    }
                else{
                    Write-Verbose "Original Client Description [$(sanitise-stripHtml $sqlKimbleAccount.Description)] does not match [$(sanitise-stripHtml $updatedLibrary.Description)]"
                    }
                }
            }
        catch{
            #Failed to update Client Library
            Write-Verbose "`tFAILED: Document Library [$($existingLibrary.Title)] was found, but not updated"
            log-result -myMessage "FAILED: Document Library [$($existingLibrary.Title)] was found, but not updated" -logFile $fullLogPathAndName
            Write-Verbose "Error updating Document Library [$($existingLibrary.Title)] - this is still marked as IsDirty=`$true :( [$($Error[0].Exception.Response)][$($Error[0].Exception.InnerException.Response)]"
            Write-Error "Error updating Document Library [$($existingLibrary.Title)] - this is still marked as IsDirty=`$true :( [$($Error[0].Exception.Response)][$($Error[0].Exception.InnerException.Response)]"
            log-error -myError $_ -myFriendlyMessage "Error updating Document Library [$($existingLibrary.Title)] - this is still marked as IsDirty=`$true :( [$($Error[0].Exception.InnerException.Response)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
            }
        }
    else{
        #Couldn't find the Library, so try creating a new one to paper over the cracks. #WCGW
        Write-Verbose "`tFAILED: Could not retrieve a Document Library for [$($sqlKimbleAccount.Name)] - sending it back for re-creation :/"
        log-result -myMessage "FAILED: Could not retrieve a Document Library for [$($sqlKimbleAccount.Name)] - sending it back for re-creation :/" -logFile $fullLogPathAndName
        Write-Verbose "Sending [$($sqlKimbleAccount.Name)] back for re-creation as it has mysteriously disappeared"
        log-action -myMessage "Sending [$($sqlKimbleAccount.Name)] back for re-creation as it has mysteriously disappeared" -logFile $fullLogPathAndName
        Write-Verbose "new-spoDocumentLibraryAndSubfoldersFromSqlKimbleListItem -sqlKimbleAccount $($sqlKimbleAccount.Name) -sqlDbConn $($sqlDbConn.DataSource) -arrayOfClientSubfolders @($($arrayOfSubfolders -join ",")) -recreateSubFolderOverride `$false"
        try{
            $duration = Measure-Command {$newLibrary = new-spoDocumentLibraryAndSubfoldersFromSqlKimbleListItem -sqlKimbleAccount $sqlKimbleAccount -sqlDbConn $sqlDbConn -arrayOfSubfolders $arrayOfSubfolders -recreateSubFolderOverride $false -adminCreds $adminCreds -fullLogPathAndName $fullLogPathAndName -errorLogPathAndName $errorLogPathAndName}
            if($newLibrary){Write-Verbose "`tSUCCESS: Weirdly unfindable Client Library [$($newLibrary.RootFolder.ServerRelativeUrl)] was recreated [$($duration.TotalSeconds) secs]" -logFile $fullLogPathAndName}
            else{
                Write-Verbose "`tFAILED: Someone left a sponge in the patient - I couldn't retrieve a Document Library for [$($sqlKimbleAccount.Name)] and I couldn't create a new one either..."
                log-result -myMessage "FAILED: Someone left a sponge in the patient - I couldn't retrieve a Document Library for [$($sqlKimbleAccount.Name)] and I couldn't create a new one either..." -logFile $fullLogPathAndName
                write-error "Someone left a sponge in the patient - I couldn't retrieve a Document Library for [$($sqlKimbleAccount.Name)] and I couldn't create a new one either..."
                log-error -myError $null -myFriendlyMessage "Someone left a sponge in the patient - I couldn't retrieve a Document Library for [$($sqlKimbleAccount.Name)] and I couldn't create a new one either..." -smtpServer "anthesisgroup-com.mail.protection.outlook.com" -mailTo "kevin.maitland@anthesisgroup.com" -mailFrom "$(split-path $PSCommandPath -Leaf)_netmon@sustain.co.uk"
                }
            }
        catch{
            Write-Verbose "Error: Borked update-spoDocumentLibraryAndSubfoldersFromSqlKimbleListItem [$($sqlKimbleAccount.Name)] [$($Error[0].Exception.Response)][$($Error[0].Exception.InnerException.Response)]"
            Write-Error "Error: Borked update-spoDocumentLibraryAndSubfoldersFromSqlKimbleListItem [$($sqlKimbleAccount.Name)] [$($Error[0].Exception.Response)][$($Error[0].Exception.InnerException.Response)]"
            log-error -myError $_ -myFriendlyMessage "Error: Borked update-spoDocumentLibraryAndSubfoldersFromSqlKimbleListItem [$($sqlKimbleAccount.Name)]" -smtpServer "anthesisgroup-com.mail.protection.outlook.com" -mailTo "kevin.maitland@anthesisgroup.com" -mailFrom "$(split-path $PSCommandPath -Leaf)_netmon@sustain.co.uk"
            }
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