$username = "kevin.maitland@anthesisgroup.com"
$password = Read-Host -Prompt "Password for $username" -AsSecureString

$webUrl = "https://anthesisllc.sharepoint.com/"

function import-CsomModules(){
    Import-Module 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.UserProfiles.dll' #CSOM for SPO User Profiles
    Import-Module 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll' #CSOM for SharePoint Online
    }
function sanitise-forSharePointFileName($dirtyString){ 
    $dirtyString.Replace("`"","").Replace("#","").Replace("%","").Replace("?","").Replace("<","").Replace(">","").Replace("\","").Replace("/","").Replace("...","").Replace("..","").Replace("'","`'")
    if($dirtyString.Substring(($dirtyString.Length-1),1) -eq "."){$dirtyString = $dirtyString.Substring(0,$dirtyString.Length-1)} #Trim trailing "."
    }
function sanitise-forSharePointUrl($dirtyString){ 
    $dirtyString = $dirtyString.Replace("`"","").Replace("#","").Replace("%","").Replace("?","").Replace("<","").Replace(">","").Replace("\","/").Replace("//","/").Replace(":","")
    $dirtyString = $dirtyString.Replace("$","`$").Replace("``$","`$").Replace("(","").Replace(")","").Replace("-","").Replace(".","").Replace("&","").Replace(",","").Replace("'","")
    $cleanString =""
    for($i= 0;$i -lt $dirtyString.Split("/").Count;$i++){ #Examine each virtual directory in the URL
        if($i -gt 0){$cleanString += "/"}
        if($dirtyString.Split("/")[$i].Length -gt 50){$tempString = $dirtyString.Split("/")[$i].SubString(0,49)} #Truncate long folder names to 50 characters
            else{$tempString = $dirtyString.Split("/")[$i]}
        if($tempString.Length -gt 0){
            if(@(".", " ") -contains $tempString.Substring(($tempString.Length-1),1)){$tempString = $tempString.Substring(0,$tempString.Length-1)} #Trim trailing "." and " ", even if this results in a truncation <50 characters
            }
        $cleanString += $tempString
        }
    $cleanString = $cleanString.Replace("//","/").Replace("https/","https://") #"//" is duplicated to catch trailing "/" that might now be duplicated. https is an exception that needs specific handling
    $cleanString
    }
function get-spoCredentials($sharePointAdminUsername, $sharePointAdminSecurePassword){
    $sharePointCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($sharePointAdminUsername, $sharePointAdminSecurePassword)
    $sharePointCredentials
    }
function new-csomContext($sharepointSite, $sharePointCredentials){
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($sharepointSite)
    $context.Credentials = $sharePointCredentials
    $context
    }
function copy-libraryItems($webUrl, $srcSite, $srcLibraryTitle, $srcLibraryName, $destLibraryTitle, $destLibraryName, $destSite, $spoCreds){
    #Title is what is displayed to the user, Name is the immutable InternalName property
    if(!$destSite){$destSite = $srcSite}
    $srcCtx = new-csomContext -sharepointSite ($webUrl+$srcSite) -sharePointCredentials $spoCreds
    $destCtx = new-csomContext -sharepointSite ($webUrl+$destSite) -sharePointCredentials $spoCreds

    $srcLibrary = $srcCtx.Web.Lists.GetByTitle($srcLibraryTitle)  
    $srcLibraryItems = $srcLibrary.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())  
    $srcCtx.Load($srcLibraryItems)  
    $srcCtx.Load($srcLibrary)  
    $srcCtx.ExecuteQuery()

    $destLibrary = $destCtx.Web.Lists.GetByTitle($destLibraryTitle)  
    $destCtx.Load($destLibrary)  
    $destCtx.ExecuteQuery()

    foreach ($doc in $srcLibraryItems){
        $destUrl = $null
        $webUrl+$destSite+$destLibraryName
        if($doc.FileSystemObjectType -eq "File"){
            $srcFile = $doc.File
            $srcCtx.Load($srcFile)
            $srcCtx.ExecuteQuery()
            $destRelativeUrl = $srcFile.ServerRelativeUrl.Replace($srcLibraryName,$destLibraryName).Replace($srcSite,$destSite)
            $srcFile.ServerRelativeUrl+" > "+$destRelativeUrl
            $srcFileData = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($srcCtx, $srcFile.ServerRelativeUrl)
            [Microsoft.SharePoint.Client.File]::SaveBinaryDirect($destCtx, $destRelativeUrl,$srcFileData.Stream,$true)
            #$srcCtx.ExecuteQuery()
            #$destCtx.ExecuteQuery()
            }

        if($doc.FileSystemObjectType -eq "Folder"){
            "FolderFound!"
            #Start-Sleep 10
            $srcFile = $doc.Folder
            $srcCtx.Load($srcFile)
            $srcCtx.ExecuteQuery()
            $tidySrcLibName = sanitise-forSharePointUrl $srcLibraryName
            $tidyDesLibName = sanitise-forSharePointUrl $destLibraryName
            $destRelativeUrl = $srcFile.ServerRelativeUrl.Replace($tidySrcLibName,$tidyDesLibName).Replace($srcSite,$destSite)
            $destFolders = $destCtx.Web.Folders.Add($destRelativeUrl)
            $destCtx.Load($destFolders)
            $destctx.ExecuteQuery()
            }
        }
    }
function add-userToSpoGroup($site, $spoCreds, $spoGroupName, $userOrGroupNameToAdd){
    $context = new-csomContext -sharepointSite $site -sharePointCredentials $spoCreds
    $groups = $context.Web.SiteGroups
    $context.Load($groups)
    $group = $groups.getByName($spoGroupName)
    $context.Load($group)

    $userOrGroup = $context.Web.EnsureUser($userOrGroupNameToAdd)
    $context.Load($userOrGroup)
    $addMe = $group.Users.AddUser($userOrGroup)
    $context.Load($addMe)
    $context.ExecuteQuery
    }

import-CsomModules
$spoCreds = get-spoCredentials -sharePointAdminUsername $username -sharePointAdminSecurePassword $password
$srcSite = "Anthesis Resources/"

$srcLibraryName = "Dummy"
$destSite = "Teams/Administration/"
$destLibraryName = "Shared Documents"
copy-libraryItems -webUrl $webUrl -srcSite $srcSite -srcLibraryName $srcLibraryName -destLibraryName $destLibraryName -destSite $destSite -spoCreds $spoCreds

$srcLibraryName = "Administration"
$destSite = "Teams/Administration/"
$destLibraryName = "Shared Documents"
copy-libraryItems -webUrl $webUrl -srcSite $srcSite -srcLibraryName $srcLibraryName -destLibraryName $destLibraryName -destSite $destSite -spoCreds $spoCreds

$srcLibraryName = "Policies & Training"
$destSite = "teams/hr/"
$destLibraryName = "Shared Documents"
copy-libraryItems -webUrl $webUrl -srcSite $srcSite -srcLibraryName $srcLibraryName -destLibraryName $destLibraryName -destSite $destSite -spoCreds $spoCreds

$srcLibraryName = "marketing"
$destSite = "teams/marketing/"
$destLibraryName = "Shared Documents"
copy-libraryItems -webUrl $webUrl -srcSite $srcSite -srcLibraryName $srcLibraryName -destLibraryName $destLibraryName -destSite $destSite -spoCreds $spoCreds

$srcSite = "Anthesis%20Projects/"
$srcLibraryName = "clients a-e"
$destSite = "clients/"
$destLibraryName = "Clients A-E"
copy-libraryItems -webUrl $webUrl -srcSite $srcSite -srcLibraryName $srcLibraryName -destLibraryName $destLibraryName -destSite $destSite -spoCreds $spoCreds
$srcLibraryName = "clients f-k"
$destSite = "clients/"
$destLibraryName = "Clients F-K"
copy-libraryItems -webUrl $webUrl -srcSite $srcSite -srcLibraryName $srcLibraryName -destLibraryName $destLibraryName -destSite $destSite -spoCreds $spoCreds
$srcLibraryName = "clients l-r"
$destSite = "clients/"
$destLibraryName = "Clients L-R"
copy-libraryItems -webUrl $webUrl -srcSite $srcSite -srcLibraryName $srcLibraryName -destLibraryName $destLibraryName -destSite $destSite -spoCreds $spoCreds
$srcLibraryName = "clients s-z"
$destSite = "clients/"
$destLibraryName = "Clients S-Z"
copy-libraryItems -webUrl $webUrl -srcSite $srcSite -srcLibraryName $srcLibraryName -destLibraryName $destLibraryName -destSite $destSite -spoCreds $spoCreds
$srcLibraryName = "Business Development"
$destSite = "clients/"
$destLibraryName = "Business Development"
copy-libraryItems -webUrl $webUrl -srcSite $srcSite -srcLibraryName $srcLibraryName -destLibraryName $destLibraryName -destSite $destSite -spoCreds $spoCreds

$srcSite = "Anthesis Resources/"
$srcLibraryTitle = "Policies & Training"
$srcLibraryName = "Policies & Training"
$destLibraryTitle = "Policies & Training"
$destSite = "teams/hr/"
$destLibraryName = "Policies & Training"
copy-libraryItems -webUrl $webUrl -srcSite $srcSite -srcLibraryName $srcLibraryName -srcLibraryTitle $srcLibraryTitle -destLibraryName $destLibraryName -destLibraryTitle $destLibraryTitle -destSite $destSite -spoCreds $spoCreds
