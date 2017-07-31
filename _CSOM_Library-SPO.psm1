$o365Admin = "kevin.maitland@anthesisgroup.com"
$o365AdminPassword = Read-Host -Prompt "Password for $o365Admin" -AsSecureString
$credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $o365Admin, $o365AdminPassword
Import-Module Microsoft.Online.Sharepoint.PowerShell
Connect-SPOService -url 'https://anthesisllc-admin.sharepoint.com' -Credential $credential

$loadInfo1 = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
$loadInfo2 = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
$loadInfo3 = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Taxonomy")
$webUrl = "https://anthesisllc.sharepoint.com" 

#region Functions
function copy-ListItems($webUrl, $srcSite, $srcListName, $destListName, $destSite){
    if(!$destSite){$destSite = $srcSite}
    $srcCtx = New-Object Microsoft.SharePoint.Client.ClientContext($global:webUrl+$srcSite) 
    $srcCtx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($o365Admin, $o365AdminPassword)
    $destCtx = New-Object Microsoft.SharePoint.Client.ClientContext($global:webUrl+$destSite) 
    $destCtx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($o365Admin, $o365AdminPassword)

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
function copy-allLibraryItems($webUrl, $srcSite, $srcLibraryName, $destLibraryName, $destSite, $spoCreds){
    if(!$destSite){$destSite = $srcSite}
    $srcCtx = new-csomContext -sharepointSite ($webUrl+$srcSite) -sharePointCredentials $spoCreds
    $destCtx = new-csomContext -sharepointSite ($webUrl+$destSite) -sharePointCredentials $spoCreds

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
function copy-libraryItems($webUrl, $srcSite, $srcLibraryName, $destLibraryName, $destSite, $spoCreds){
    if(!$destSite){$destSite = $srcSite}
    $srcCtx = new-csomContext -sharepointSite ($webUrl+$srcSite) -sharePointCredentials $spoCreds
    $destCtx = new-csomContext -sharepointSite ($webUrl+$destSite) -sharePointCredentials $spoCreds

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
function apply-theme($webUrl, $siteCollection, $site, $colorPaletteUrl, $fontSchemeUrl, $backgroundImageUrl, $shareGenerated ){
    #$shareGenerated: true if the generated theme files should be placed in the root web, false to store them in this web.
    #Weirdly, it doesn't like $null as value
    #if($backgroundImageUrl -eq $null){$backgroundImageUrl = Out-Null}
    $backgroundImageUrl = Out-Null
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($webUrl+$siteCollection+$site) 
    $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($o365Admin, $o365AdminPassword)
    $web = $ctx.Web
    $ctx.Load($web)
    $web.ApplyTheme($siteCollection+$colorPaletteUrl, $siteCollection+$spFontUrl, $backgroundImageUrl, $shareGenerated)
    $web.Update()
    $ctx.ExecuteQuery()
    }
function set-navTopNodes($webUrl,$siteCollectionOrSite,$deleteAllBeforeAdding,$hashTableOfNodes){
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($webUrl+$siteCollectionOrSite) 
    $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($o365Admin, $o365AdminPassword)
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
function add-termToStore($pGroup,$pSet,$pTerm){
    #Sanitise the input:
    $pGroup = sanitise-forTermStore $pGroup
    $pSet = sanitise-forTermStore $pSet
    $pTerm = sanitise-forTermStore $pTerm
    $lcid = 1033 #Probably Language
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($serverUrl) 
    $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($o365user, $o365Pass)
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
function delete-termFromStore($pGroup,$pSet,$pTerm){
    #Sanitise the input:
    $pGroup = sanitise-forTermStore $pGroup
    $pSet = sanitise-forTermStore $pSet
    $pTerm = sanitise-forTermStore $pTerm
    $lcid = 1033 #Probably Language
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($serverUrl) 
    $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($o365user, $o365Pass)
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
function rename-termInStore($pGroup,$pSet,$pOldTerm, $pNewTerm){
    #Sanitise the input:
    $pGroup = sanitise-forTermStore $pGroup
    $pSet = sanitise-forTermStore $pSet
    $pTerm = sanitise-forTermStore $pTerm
    $lcid = 1033 #Probably Language
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($serverUrl) 
    $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($o365user, $o365Pass)
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
function get-webTempates($webUrl, $siteCollection, $site){
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($webUrl+$siteCollection+$site) 
    $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($o365Admin, $o365AdminPassword)
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
function new-SPOGroup($title, $description, $spoSite, $ctx){
    $spoGroupCreationInfo=New-Object Microsoft.SharePoint.Client.GroupCreationInformation
    $spoGroupCreationInfo.Title=$title
    $spoGroupCreationInfo.Description=$description
    $newGroup = $spoSite.SiteGroups.Add($spoGroupCreationInfo)
    $ctx.ExecuteQuery()
    $newGroup
    }
function add-site($webUrl, $siteCollection, $sitePath, $siteName, $siteUrlEndStub, $siteTemplate, $inheritPermissions, $owner){
    #{8C3E419E-EADC-4032-A7CD-BC5778A30F9C}#Default External Sharing Site /sites/external
    #{7FD4CC3D-B615-4930-A041-3ADB8C6509EA}#Default Community Site /teams/communities
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($webUrl+$siteCollection+$site) 
    $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($o365Admin, $o365AdminPassword)
    $webCreationInformation = New-Object Microsoft.SharePoint.Client.WebCreationInformation
    $webCreationInformation.Url = $siteUrlEndStub
    $webCreationInformation.Title = $siteName
    $webCreationInformation.WebTemplate = $siteTemplate
    $webCreationInformation.UseSamePermissionsAsParentSite = $inheritPermissions
    
    $newWeb = $ctx.Web.Webs.Add($webCreationInformation)
    $ctx.Load($newWeb) 
    $ctx.ExecuteQuery()
    
    if($inheritPermissions -eq $false){
        #Create the standard groups
        $ownersGroupInfo  = new-SPOGroup -title "$siteName Owners"  -description "Managers and Admins of $siteName" -spoSite $newWeb -ctx $ctx
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



    #Brand the bugger
    $colorPaletteUrl = "/_catalogs/theme/15/AnthesisPalette_Orange.spcolor"
    $spFontUrl = "/_catalogs/theme/15/Anthesis_fontScheme_Montserrat_uploaded.spfont"
    apply-theme -webUrl $webUrl -siteCollection $siteCollection -site "$sitePath$siteUrlEndStub" -colorPaletteUrl $colorPaletteUrl -fontSchemeUrl $spFontUrl -backgroundImageUrl $null -shareGenerated $false
    }
function add-memberToGroup($webUrl, $siteCollection, $sitePath, $groupName, $memberToAdd){
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($webUrl+$siteCollection+$sitePath) 
    $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($o365Admin, $o365AdminPassword)
    $groups = $ctx.Web.SiteGroups
    $ctx.Load($groups)
    $group = $groups.GetByName($groupName)
    $ctx.Load($group)
    $userToAdd = $ctx.Web.EnsureUser($memberToAdd)
    $ctx.Load($userToAdd)
    $ctx.Load($group.Users.AddUser($userToAdd))
    $ctx.ExecuteQuery()
    }
function remove-memberFromGroup($webUrl, $siteCollection, $sitePath, $groupName, $memberToRemove){
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($webUrl+$siteCollection+$sitePath) 
    $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($o365Admin, $o365AdminPassword)
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
#endregion
#add-site -webUrl $webUrl -$siteCollection = "/teams/communities" -$sitePath = "/" -$siteName = "STEP" -$siteUrlEndStub = "step" -$siteTemplate = "{7FD4CC3D-B615-4930-A041-3ADB8C6509EA}#Default Community Site" -$inheritPermissions = $false -$owner = "Graeme Hadley"
#add-memberToGroup -webUrl $webUrl -siteCollection $siteCollection -$sitePath = "/step" -$groupName = "STEP Owners" -$memberToAdd = "STEP Team"
add-memberToGroup -webUrl $webUrl -siteCollection "/teams/communities" -sitePath "/step" -groupName "STEP Owners" -memberToAdd $owner
add-memberToGroup -webUrl $webUrl -siteCollection "/teams/communities" -sitePath "/step" -groupName "STEP Members" -memberToAdd "STEP Team"

#Brand Sites:
$siteCollection = "/sites/external" #Main Site Collections are "/", "/teams/communities", "/teams/sym" Full list available here: https://anthesisllc-admin.sharepoint.com/_layouts/15/online/SiteCollections.aspx
$sites = @("/unite") #@("/intsus","/anyOtherSite", "/anyOtherSite/SubSite")
foreach($site in $sites){
    $colorPaletteUrl = "/_catalogs/theme/15/AnthesisPalette_Orange.spcolor"
    $spFontUrl = "/_catalogs/theme/15/Anthesis_fontScheme_Montserrat_uploaded.spfont"
    apply-theme -webUrl $webUrl -siteCollection $siteCollection -site $site -colorPaletteUrl $colorPaletteUrl -fontSchemeUrl $spFontUrl -backgroundImageUrl $null -shareGenerated $false
    }

#Update TopNav Bar:
$sitesToUpdate= @("/teams/hr","/teams/finance","/teams/marketing","/teams/administration","/clients","/teams/confidential","/global", "/teams/all", "/teams/communities", "/subs")
foreach ($site in $sitesToUpdate){
    $nodesToAdd = [ordered]@{"Global"="https://anthesisllc.sharepoint.com/global";"Clients"= "https://anthesisllc.sharepoint.com/Anthesis Projects";"Resources"="https://anthesisllc.sharepoint.com/Anthesis Resources";"External"="https://anthesisllc.sharepoint.com/sites/external";"Kimble"="https://login.salesforce.com/";"Search"="https://anthesisllc.sharepoint.com/search";"Help"="https://anthesisllc.sharepoint.com/help"}
    $newNodesToAdd = [ordered]@{"Global"="https://anthesisllc.sharepoint.com/global";"Clients"= "https://anthesisllc.sharepoint.com/clients";"Resources"="https://anthesisllc.sharepoint.com/teams/all/Lists/Internal%20Teams";"External"="https://anthesisllc.sharepoint.com/sites/external";"Kimble"="https://login.salesforce.com/";"Search"="https://anthesisllc.sharepoint.com/search";"Help"="https://anthesisllc.sharepoint.com/help"}
    #set-navTopNodes -webUrl $webUrl -siteCollectionOrSite $site -deleteAllBeforeAdding $true -hashTableOfNodes $newNodesToAdd
    }







#copy-ListItems -webUrl $webUrl -srcSite "/teams/sym/" -srcListName "Sym Groups" -destSite "/teams/all/" -destListName "Sym Groups"
#copy-ListItems -webUrl $webUrl -srcSite "/teams/communities/" -srcListName "Communities" -destSite "/teams/all/" -destListName "Communities"
#copy-libraryItems -webUrl $webUrl -srcSite "/Anthesis Resources/" -srcLibraryName "Images" -destSite "/teams/all/" -destLibraryName "Images"


$extGroups  = $groups | ? {$_.Users -match "ext"} | select Title
$groups | ? {$_.users -eq "Admin Info"}
