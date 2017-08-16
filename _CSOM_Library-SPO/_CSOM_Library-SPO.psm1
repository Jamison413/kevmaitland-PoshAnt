#Import-Module Microsoft.Online.Sharepoint.PowerShell
#Connect-SPOService -url 'https://anthesisllc-admin.sharepoint.com' -Credential $credential

[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Taxonomy") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.UserProfiles") | Out-Null
$webUrl = "https://anthesisllc.sharepoint.com" 

$loadInfo1
#region Functions
function add-memberToGroup($credentials, $webUrl, $siteCollection, $sitePath, $groupName, $memberToAdd){
    $ctx = new-csomContext -fullSitePath ($webUrl+$siteCollection+$sitePath) -sharePointCredentials $credentials
    $groups = $ctx.Web.SiteGroups
    $ctx.Load($groups)
    $group = $groups.GetByName($groupName)
    $ctx.Load($group)
    $userToAdd = $ctx.Web.EnsureUser($memberToAdd)
    $ctx.Load($userToAdd)
    $ctx.Load($group.Users.AddUser($userToAdd))
    $ctx.ExecuteQuery()
    }
function add-site($credentials, $webUrl, $siteCollection, $sitePath, $siteName, $siteUrlEndStub, $siteTemplate, $inheritPermissions, $inheritTopNav, $owner){
    #{8C3E419E-EADC-4032-A7CD-BC5778A30F9C}#Default External Sharing Site /sites/external
    #{7FD4CC3D-B615-4930-A041-3ADB8C6509EA}#Default Community Site /teams/communities
    $ctx = new-csomContext -fullSitePath ($webUrl+$siteCollection+$sitePath) -sharePointCredentials $credentials
    $webCreationInformation = New-Object Microsoft.SharePoint.Client.WebCreationInformation
    $webCreationInformation.Url = $siteUrlEndStub
    $webCreationInformation.Title = $siteName
    $webCreationInformation.WebTemplate = $siteTemplate
    $webCreationInformation.UseSamePermissionsAsParentSite = $inheritPermissions
    
    $newWeb = $ctx.Web.Webs.Add($webCreationInformation)
    $ctx.Load($newWeb)
    $nNav = $newWeb.Navigation
    $ctx.Load($nNav)
    $nNav.UseShared = $inheritTopNav
    $ctx.ExecuteQuery()
    #$newWeb.Navigation.UseShared = $inheritTopNav
    
    if($inheritPermissions -eq $false){
        #Create the standard groups
        $ownersGroup  = new-SPOGroup -title "$siteName Owners"  -description "Managers and Admins of $siteName" -spoSite $newWeb -ctx $ctx
        $membersGroup = new-SPOGroup -title "$siteName Members" -description "Contributors to $siteName" -spoSite $newWeb -ctx $ctx
        $visitorsGroup = new-SPOGroup -title "$siteName Visitors" -description "ReadOnly users of $siteName" -spoSite $newWeb -ctx $ctx
        
        #Get the standard Roles
        $roleDefBindFullControl = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($ctx)
        $roleDefBindFullControl.Add($newWeb.RoleDefinitions.GetByName("Full Control"))
        $roleDefBindEdit = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($ctx)
        $roleDefBindEdit.Add($newWeb.RoleDefinitions.GetByName("Edit"))
        $roleDefBindContribute = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($ctx)
        $roleDefBindContribute.Add($newWeb.RoleDefinitions.GetByName("Contribute"))
        $roleDefBindRead = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($ctx)
        $roleDefBindRead.Add($newWeb.RoleDefinitions.GetByName("Read"))

        #Assign the standard Roles to the standard Groups
        $ctx.Load($newWeb.RoleAssignments.Add($ownersGroup, $roleDefBindFullControl))
        $ctx.Load($newWeb.RoleAssignments.Add($membersGroup, $roleDefBindContribute))
        $ctx.Load($newWeb.RoleAssignments.Add($visitorsGroup, $roleDefBindRead))
        $ctx.ExecuteQuery()

        #Remove the current user from the Site
        #Get the current User
        $currentUser = $newWeb.CurrentUser
        $ctx.Load($currentUser)
        #Get the current RoleAssignments
        $members= @()
        $roleAssignments = $newWeb.RoleAssignments
        $ctx.Load($roleAssignments)
        $ctx.ExecuteQuery()

        #Look for the current user in the collection of RoleAssigments and delete it if we find it
        $roleAssignments.GetEnumerator() | % {
            $member = $newWeb.RoleAssignments.GetByPrincipalId($_.PrincipalId).Member
            $ctx.Load($member)
            $members += $member
            if($member.Title -eq $currentUser.Title){$newWeb.RoleAssignments.GetByPrincipalId($member.Id).DeleteObject()}
            }
        $ctx.ExecuteQuery()
 
        }

    #Set the RequestAccessEmail details
    if($owner -notmatch "@anthesisgroup.com"){$owner = $owner.Replace(" ",".") + "@anthesisgroup.com"}
    $newWeb.RequestAccessEmail = $owner
    $ctx.ExecuteQuery()

    #Brand the bugger
    $colorPaletteUrl = "/_catalogs/theme/15/AnthesisPalette_Orange.spcolor"
    $spFontUrl = "/_catalogs/theme/15/Anthesis_fontScheme_Montserrat_uploaded.spfont"
    apply-theme -credentials $credentials -webUrl $webUrl -siteCollection $siteCollection -site "$sitePath$siteUrlEndStub" -colorPaletteUrl $colorPaletteUrl -fontSchemeUrl $spFontUrl -backgroundImageUrl $null -shareGenerated $false
    }
function add-termToStore($credentials, $webUrl, $siteCollection, $pGroup,$pSet,$pTerm){
    #Sanitise the input:
    $pGroup = sanitise-forTermStore $pGroup
    $pSet = sanitise-forTermStore $pSet
    $pTerm = sanitise-forTermStore $pTerm
    $lcid = 1033 #Probably Language
    $ctx = new-csomContext -fullSitePath ($webUrl+$siteCollection) -sharePointCredentials $credentials
    Write-Host "Connected to DESTINATION SharePoint Online site: " $ctx.Url "" -ForegroundColor Green
    $taxonomySession = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($ctx)
    $taxonomySession.UpdateCache()
    $ctx.Load($taxonomySession)
    $ctx.ExecuteQuery()
    if (!$taxonomySession.ServerObjectIsNull){
        Write-Host "Destination Taxonomy session initiated: " $taxonomySession.Path.Identity "" -ForegroundColor Green
        $termStore = $taxonomySession.GetDefaultSiteCollectionTermStore()
        $ctx.Load($termStore)
        $ctx.ExecuteQuery()
        if ($termStore.IsOnline){
            Write-Host "...Default Term Store connected:" $termStore.Id "" -ForegroundColor Green
            # $termStoreId will be the SspId in the taxonomy column configs
            $continue = 1
            }
        
        #Create the TermGroup if necessary
        $ctx.Load($termStore.Groups)
        $ctx.ExecuteQuery()
        if(!(($termStore.Groups | select Name) -match $pGroup)){
            $newGroup = $termStore.CreateGroup($pGroup,(New-Guid))
            $ctx.Load($newGroup)
            $ctx.ExecuteQuery()
           }
        $termGroup = $termStore.Groups.GetByName($pGroup)
        $ctx.Load($termGroup)
        $ctx.ExecuteQuery()

        #Create the TermSet if necessary
        $ctx.Load($termGroup.TermSets)
        $ctx.ExecuteQuery()
        if(!(($termGroup.TermSets | select Name) -match $pSet)){
            $newSet = $termGroup.CreateTermSet($pSet,(New-Guid))
            $ctx.Load($newSet)
            $ctx.ExecuteQuery()
           }
        $termSet = $termGroup.TermSets.GetByName($pSet)
        $ctx.Load($termSet)
        $ctx.ExecuteQuery()

        #Create the Term if necessary
        $ctx.Load($termSet.Terms)
        $ctx.ExecuteQuery()
        if(!(($termSet.Terms | select Name) -match $pTerm)){
            $newTerm = $termSet.CreateTerm($pTerm,$lcid,(New-Guid))
            $ctx.Load($newTerm)
            $ctx.ExecuteQuery()
           }
        }
    }
function apply-theme($credentials, $webUrl, $siteCollection, $site, $colorPaletteUrl, $fontSchemeUrl, $backgroundImageUrl, $shareGenerated ){
    #$shareGenerated: true if the generated theme files should be placed in the root web, false to store them in this web.
    #Weirdly, it doesn't like $null as value
    #if($backgroundImageUrl -eq $null){$backgroundImageUrl = Out-Null}
    $backgroundImageUrl = Out-Null
    $ctx = new-csomContext -fullSitePath ($webUrl+$siteCollection+$site) -sharePointCredentials $credentials
    $web = $ctx.Web
    $ctx.Load($web)
    $web.ApplyTheme($siteCollection+$colorPaletteUrl, $siteCollection+$spFontUrl, $backgroundImageUrl, $shareGenerated)
    $web.Update()
    $ctx.ExecuteQuery()
    }
function copy-ListItems($credentials, $webUrl, $srcSite, $srcListName, $destListName, $destSite){
    if(!$destSite){$destSite = $srcSite}
    $srcCtx = new-csomContext -fullSitePath ($webUrl+$srcSite) -sharePointCredentials $credentials
    $destCtx = new-csomContext -fullSitePath ($webUrl+$destSite) -sharePointCredentials $credentials

    $srcList = $srcCtx.Web.Lists.GetByTitle($srcListName)  
    $destList = $destCtx.Web.Lists.GetByTitle($destListName)  
    $srcListItems = $srcList.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())  
    $fields = $srcList.Fields  
    $srcCtx.Load($srcListItems)  
    $srcCtx.Load($srcList)  
    $destCtx.Load($destList)  
    $srcCtx.Load($fields)  
    $srcCtx.ExecuteQuery()
    $destCtx.ExecuteQuery()
    
    foreach($item in $srcListItems){  
        Write-Host $item.ID  
        $listItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation  
        $destListItem = $destList.AddItem($listItemInfo)  
          
        foreach($field in $fields){  
            #Write-Host $field.InternalName " - " $field.ReadOnlyField   
            if((-Not ($field.ReadOnlyField)) -and (-Not ($field.Hidden)) -and ($field.InternalName -ne  "Attachments") -and ($field.InternalName -ne  "ContentType")){
                Write-Host $field.InternalName " - " $item[$field.InternalName]  
                $destListItem[$field.InternalName] = $item[$field.InternalName]  
                $destListItem.update()  
                }  
            }  
        }  
        $destCtx.ExecuteQuery()   
    }
function copy-allLibraryItems($credentials, $webUrl, $srcSite, $srcLibraryName, $destLibraryName, $destSite){
    if(!$destSite){$destSite = $srcSite}
    $srcCtx = new-csomContext -sharepointSite ($webUrl+$srcSite) -sharePointCredentials $credentials
    $destCtx = new-csomContext -sharepointSite ($webUrl+$destSite) -sharePointCredentials $credentials

    $srcLibrary = $srcCtx.Web.Lists.GetByTitle($srcLibraryName)  
    $srcLibraryItems = $srcLibrary.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())  
    $srcCtx.Load($srcLibraryItems)  
    $srcCtx.Load($srcLibrary)  
    $srcCtx.ExecuteQuery()

    $destLibrary = $destCtx.Web.Lists.GetByTitle($destLibraryName)  
    $destCtx.Load($destLibrary)  
    $destCtx.ExecuteQuery()

    foreach ($doc in $srcLibraryItems){
        $destUrl = $null
        $webUrl+$destSite+$destLibraryName
        if($doc.FileSystemObjectType -eq "File"){
            $srcFile = $doc.File
            $srcCtx.Load($srcFile)
            $srcCtx.ExecuteQuery()
            $destRelativeUrl = $srcFile.ServerRelativeUrl.Replace($srcSite,$destSite).Replace($srcLibraryName,$destLibraryName)
            $srcFile.ServerRelativeUrl+" > "+$destRelativeUrl
            $srcFileData = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($srcCtx, $srcFile.ServerRelativeUrl)
            [Microsoft.SharePoint.Client.File]::SaveBinaryDirect($destCtx, $destRelativeUrl,$srcFileData.Stream,$true)
            $srcCtx.ExecuteQuery()
            $destCtx.ExecuteQuery()
            }

        if($doc.FileSystemObjectType -eq "Folder"){
            "FolderFound!"
            }
        }
    }
function copy-libraryItems($credentials, $webUrl, $srcSite, $srcLibraryName, $destLibraryName, $destSite){
    if(!$destSite){$destSite = $srcSite}
    $srcCtx = new-csomContext -sharepointSite ($webUrl+$srcSite) -sharePointCredentials $credentials
    $destCtx = new-csomContext -sharepointSite ($webUrl+$destSite) -sharePointCredentials $credentials

    $srcLibrary = $srcCtx.Web.Lists.GetByTitle($srcLibraryName)  
    $srcLibraryItems = $srcLibrary.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())  
    $srcCtx.Load($srcLibraryItems)  
    $srcCtx.Load($srcLibrary)  
    $srcCtx.ExecuteQuery()

    $destLibrary = $destCtx.Web.Lists.GetByTitle($destLibraryName)  
    $destCtx.Load($destLibrary)  
    $destCtx.ExecuteQuery()

    foreach ($doc in $srcLibraryItems){
        $destUrl = $null
        $webUrl+$destSite+$destLibraryName
        if($doc.FileSystemObjectType -eq "File"){
            $srcFile = $doc.File
            $srcCtx.Load($srcFile)
            $srcCtx.ExecuteQuery()
            $destRelativeUrl = $srcFile.ServerRelativeUrl.Replace($srcSite,$destSite).Replace($srcLibraryName,$destLibraryName)
            $srcFile.ServerRelativeUrl+" > "+$destRelativeUrl
            $srcFileData = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($srcCtx, $srcFile.ServerRelativeUrl)
            [Microsoft.SharePoint.Client.File]::SaveBinaryDirect($destCtx, $destRelativeUrl,$srcFileData.Stream,$true)
            $srcCtx.ExecuteQuery()
            $destCtx.ExecuteQuery()
            }

        if($doc.FileSystemObjectType -eq "Folder"){
            "FolderFound!"
            }
        }
    }
function delete-termFromStore($credentials, $webUrl, $siteCollection, $pGroup,$pSet,$pTerm){
    #Sanitise the input:
    $pGroup = sanitise-forTermStore $pGroup
    $pSet = sanitise-forTermStore $pSet
    $pTerm = sanitise-forTermStore $pTerm
    $lcid = 1033 #Probably Language
    $ctx = new-csomContext -fullSitePath ($webUrl+$siteCollection) -sharePointCredentials $credentials
    Write-Host "Connected to DESTINATION SharePoint Online site: " $ctx.Url "" -ForegroundColor Green
    $taxonomySession = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($ctx)
    $taxonomySession.UpdateCache()
    $ctx.Load($taxonomySession)
    $ctx.ExecuteQuery()
    if (!$taxonomySession.ServerObjectIsNull){
        Write-Host "Destination Taxonomy session initiated: " $taxonomySession.Path.Identity "" -ForegroundColor Green
        $termStore = $taxonomySession.GetDefaultSiteCollectionTermStore()
        $ctx.Load($termStore)
        $termGroup = $termStore.Groups.GetByName($pGroup)
        $ctx.Load($termGroup)
        $termSet = $termGroup.TermSets.GetByName($pSet)
        $ctx.Load($termSet)
        $ctx.ExecuteQuery()        #Delete the Term if necessary
        $ctx.Load($termSet.Terms)
        $ctx.ExecuteQuery()
        if(($termSet.Terms.Name) -contains $pTerm){
            $term = $termSet.Terms.GetByName($pTerm)
            $ctx.Load($term)
            $term.DeleteObject()
            $ctx.ExecuteQuery()
           }
        }
    }
function get-webTempates($credentials, $webUrl, $siteCollection, $site){
    $ctx = new-csomContext -fullSitePath ($webUrl+$siteCollection+$site) -sharePointCredentials $credentials
    #New-Object Microsoft.SharePoint.Client.ClientContext($webUrl+$siteCollection+$site) 
    #$ctx.Credentials = $credentials
    $web = $ctx.Web
    $ctx.Load($web)
    $templateCollection = $web.GetAvailableWebTemplates("1033",$false)
    $ctx.Load($templateCollection)
    $siteTemplates = @()
    $siteTemplatesEnum = $templateCollection.GetEnumerator()
    $ctx.ExecuteQuery()
    while ($siteTemplatesEnum.MoveNext()) {
        $siteTemplates += $siteTemplatesEnum.Current.Name
        }
    $siteTemplates
    }
function new-csomContext ($fullSitePath, $sharePointCredentials){
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($fullSitePath) 
    $ctx.Credentials = $sharePointCredentials
    $ctx
    }
function new-SPOGroup($credentials, $title, $description, $spoSite, $ctx){
    $spoGroupCreationInfo=New-Object Microsoft.SharePoint.Client.GroupCreationInformation
    $spoGroupCreationInfo.Title=$title
    $spoGroupCreationInfo.Description=$description
    $newGroup = $spoSite.SiteGroups.Add($spoGroupCreationInfo)
    $ctx.ExecuteQuery()
    $newGroup
    }
function remove-memberFromGroup($credentials, $webUrl, $siteCollection, $sitePath, $groupName, $memberToRemove){
    $ctx = new-csomContext -fullSitePath ($webUrl+$siteCollection+$sitePath) -sharePointCredentials $credentials
    #Get the current RoleAssignments
    $roleAssignments = $ctx.Web.RoleAssignments
    $ctx.Load($roleAssignments)
    #Validate the user to be removed
    $userToRemove = $ctx.Web.EnsureUser($memberToRemove)
    $ctx.Load($userToRemove)
    $ctx.ExecuteQuery()
    #Look for the current user in the collection of RoleAssigments and delete it if we find it
    $roleAssignments.GetEnumerator() | % {
        $member = $ctx.Web.RoleAssignments.GetByPrincipalId($_.PrincipalId).Member
        $ctx.Load($member)
        if($member.Title -eq $userToRemove.Title){$newWeb.RoleAssignments.GetByPrincipalId($member.Id).DeleteObject()}
        }
    $ctx.ExecuteQuery()
 
    }
function rename-termInStore($credentials, $webUrl, $siteCollection, $pGroup,$pSet,$pOldTerm, $pNewTerm){
    #Sanitise the input:
    $pGroup = sanitise-forTermStore $pGroup
    $pSet = sanitise-forTermStore $pSet
    $pTerm = sanitise-forTermStore $pTerm
    $lcid = 1033 #Probably Language
    $ctx = new-csomContext -fullSitePath ($webUrl+$siteCollection) -sharePointCredentials $credentials
    Write-Host "Connected to DESTINATION SharePoint Online site: " $ctx.Url "" -ForegroundColor Green
    $taxonomySession = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($ctx)
    $taxonomySession.UpdateCache()
    $ctx.Load($taxonomySession)
    $ctx.ExecuteQuery()
    if (!$taxonomySession.ServerObjectIsNull){
        Write-Host "Destination Taxonomy session initiated: " $taxonomySession.Path.Identity "" -ForegroundColor Green
        $termStore = $taxonomySession.GetDefaultSiteCollectionTermStore()
        $ctx.Load($termStore)
        $termGroup = $termStore.Groups.GetByName($pGroup)
        $ctx.Load($termGroup)
        $termSet = $termGroup.TermSets.GetByName($pSet)
        $ctx.Load($termSet)
        $ctx.ExecuteQuery()        #Delete the Term if necessary
        $ctx.Load($termSet.Terms)
        $ctx.ExecuteQuery()
        if(($termSet.Terms.Name) -contains $pOldTerm){
            $term = $termSet.Terms.GetByName($pOldTerm)
            $ctx.Load($pOldTerm)
            $term.Name = $pNewTerm
            $term.Name = "BBC"
            $ctx.ExecuteQuery()
           }
        }
    }
function set-csomCredentials($username, $password){
    if ($username -eq $null -or $username -eq ""){$username = Read-Host -Prompt "Enter SharePoint Online username (blank for $($env:USERNAME)@anthesisgroup.com)"}
    if ($username -eq $null -or $username -eq ""){$username = "$($env:USERNAME)@anthesisgroup.com"}
    if ($password -eq $null -or $password -eq ""){$password = Read-Host -Prompt "Password for $username" -AsSecureString}
    New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $password)
    }
function set-navTopNodes($credentials, $webUrl, $siteCollection, $sitePath, $deleteAllBeforeAdding, $hashTableOfNodes){
    $ctx = new-csomContext -fullSitePath ($webUrl+$siteCollection+$sitePath) -sharePointCredentials $credentials
    $web = $ctx.Web 
    $ctx.Load($web) 
    $NavBar = $ctx.Web.Navigation.TopNavigationBar 
    if($deleteAllBeforeAdding){
        $ctx.Load($NavBar)
        $ctx.ExecuteQuery()
        while($NavBar.Count -gt 0){$NavBar.item(0).DeleteObject()}
        }
    foreach($key in $hashTableOfNodes.Keys){
        $NavigationNode = New-Object Microsoft.SharePoint.Client.NavigationNodeCreationInformation
        $NavigationNode.Title = $key
        $NavigationNode.Url = $hashTableOfNodes[$key]
        $NavigationNode.AsLastNode = $true           
        $NavigationNode.IsExternal = $true           
        $ctx.Load($NavBar.Add($NavigationNode)) 
        $ctx.ExecuteQuery()
        }    
    }
#endregion
