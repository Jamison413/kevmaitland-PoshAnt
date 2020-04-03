

#[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client") 
#[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.ClientContext")
#[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
#[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Sharing") 
#[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Taxonomy") 
#[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.UserProfiles")
#Import-Module  "$env:USERPROFILE\SPO_CSOM\Microsoft.SharePointOnline.CSOM.16.1.8210.1200\lib\net45\Microsoft.SharePoint.Client.dll"
#Import-Module  "$env:USERPROFILE\SPO_CSOM\Microsoft.SharePointOnline.CSOM.16.1.8210.1200\lib\net45\Microsoft.SharePoint.Client.ClientContext.dll"
#Import-Module  "$env:USERPROFILE\SPO_CSOM\Microsoft.SharePointOnline.CSOM.16.1.8210.1200\lib\net45\Microsoft.SharePoint.Client.Runtime.dll"
#Import-Module  "$env:USERPROFILE\SPO_CSOM\Microsoft.SharePointOnline.CSOM.16.1.8210.1200\lib\net45\Microsoft.SharePoint.Client.SharePointOnlineCredentials.dll"
#Import-Module  "$env:USERPROFILE\SPO_CSOM\Microsoft.SharePointOnline.CSOM.16.1.8210.1200\lib\net45\Microsoft.SharePoint.Client.Sharing.dll"
#Import-Module  "$env:USERPROFILE\SPO_CSOM\Microsoft.SharePointOnline.CSOM.16.1.8210.1200\lib\net45\Microsoft.SharePoint.Client.Taxonomy.dll"
#Import-Module  "$env:USERPROFILE\SPO_CSOM\Microsoft.SharePointOnline.CSOM.16.1.8210.1200\lib\net45\Microsoft.SharePoint.Client.UserProfiles.dll"

#Import-Module _PS_Library_GeneralFunctionality
$webUrl = "https://anthesisllc.sharepoint.com" 

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
    <#
    #Get $newWeb again if we lose it
    $z_webs = $ctx.Web.Webs
    $ctx.Load($z_webs)
    $ctx.ExecuteQuery()
    $z_webs | %{
        $_.Title
        if($siteName -eq $_.Title){
            $newWeb = $_
            $ctx.Load($newWeb)
            $ctx.Load($newWeb.AllProperties)
            $ctx.Load($newWeb.AssociatedOwnerGroup)
            $ctx.Load($newWeb.AssociatedMemberGroup)
            $ctx.Load($newWeb.AssociatedVisitorGroup)
            $ctx.ExecuteQuery()
            }
        }
    #>

    $ctx = new-csomContext -fullSitePath ($webUrl+$siteCollection+$sitePath) -sharePointCredentials $credentials
    $webCreationInformation = New-Object Microsoft.SharePoint.Client.WebCreationInformation
    $webCreationInformation.Url = $siteUrlEndStub
    $webCreationInformation.Title = $siteName
    $webCreationInformation.WebTemplate = $siteTemplate
    $webCreationInformation.UseSamePermissionsAsParentSite = $inheritPermissions
    
    #Create the Site
    $newWeb = $ctx.Web.Webs.Add($webCreationInformation)
    $ctx.Load($newWeb)
    $nNav = $newWeb.Navigation
    $ctx.Load($nNav)
    $nNav.UseShared = $inheritTopNav
    $ctx.ExecuteQuery()

    #Add link to Parent's QuickLaunch (if the Subsites Node exists)
    $ql = $ctx.web.Navigation.QuickLaunch
    $ctx.Load($ql)
    $ctx.ExecuteQuery()
    $ql.GetEnumerator() | % {
        if("Subsites" -eq $_.Title){$subSitesId = $_.Id}
        }
    if($subSitesId){
        $subSitesNode = $ctx.Web.Navigation.GetNodeById($subSitesId)
        $ctx.Load($subSitesNode)
        $newNode = New-Object Microsoft.SharePoint.Client.NavigationNodeCreationInformation
        $newNode.Title = $siteName
        $newNode.Url = $siteCollection+$sitePath+$siteUrlEndStub
        $newNode.AsLastNode = $false
        $newNode.IsExternal = $true
        $ctx.Load($subSitesNode.Children.Add($newNode))
        $ctx.ExecuteQuery()
        }



    if(!$inheritPermissions){
        #Create the standard groups
        $siteNameForGroups = sanitise-forSharePointGroupName $siteName
        $ownersGroup  = new-SPOGroup -title "$siteNameForGroups Owners"  -description "Managers and Admins of $siteName" -spoSite $newWeb -ctx $ctx 
        $membersGroup = new-SPOGroup -title "$siteNameForGroups Members" -description "Contributors to $siteName" -spoSite $newWeb -ctx $ctx
        $visitorsGroup = new-SPOGroup -title "$siteNameForGroups Visitors" -description "ReadOnly users of $siteName" -spoSite $newWeb -ctx $ctx
        
        #Get the standard Roles
        $roleDefBindFullControl = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($ctx)
        $roleDefBindFullControl.Add($newWeb.RoleDefinitions.GetByName("Full Control"))
        $roleDefBindEdit = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($ctx)
        $roleDefBindEdit.Add($newWeb.RoleDefinitions.GetByName("Edit"))
        $roleDefBindContribute = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($ctx)
        $roleDefBindContribute.Add($newWeb.RoleDefinitions.GetByName("Contribute"))
        $roleDefBindRead = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($ctx)
        $roleDefBindRead.Add($newWeb.RoleDefinitions.GetByName("Read"))

        #Assign the standard Roles to the standard Groups, set them as the Default Groups and 
        $ctx.Load($newWeb.RoleAssignments.Add($ownersGroup, $roleDefBindFullControl))
        if($siteCollection -match "confidential"){$ctx.Load($newWeb.RoleAssignments.Add($membersGroup, $roleDefBindEdit))}
            else{$ctx.Load($newWeb.RoleAssignments.Add($membersGroup, $roleDefBindContribute))}
        $ctx.Load($newWeb.RoleAssignments.Add($visitorsGroup, $roleDefBindRead))
        $newWeb.AssociatedOwnerGroup = $ownersGroup
        $newWeb.AssociatedMemberGroup = $membersGroup
        $newWeb.AssociatedVisitorGroup = $visitorsGroup
        $newWeb.Update()
        $ctx.ExecuteQuery()

        #Remove the current user from the Site
        $currentUser = $newWeb.CurrentUser
        $ctx.Load($currentUser)
        $ctx.ExecuteQuery()
        remove-userFromSite -credentials $credentials -webUrl $webUrl -siteCollection $siteCollection -sitePath $sitePath$siteUrlEndStub -memberToRemove $currentUser.Title
        }

    #Set the RequestAccessEmail details and disable Members' ability to Share
    if($owner -notmatch "@anthesisgroup.com"){$owner = $owner.Replace(" ",".") + "@anthesisgroup.com"}
    $newWeb.RequestAccessEmail = $owner
    $newWeb.MembersCanShare = $false
    $newweb.Update()
    $ctx.ExecuteQuery()

    
    #Unfuckulate an inheritance bug on the Access Requests list
    if($siteCollection -match "external"){
        new-sharingInvite -credentials $credentials -webUrl $webUrl -siteCollection $siteCollection -sitePath $sitePath$siteUrlEndStub -userEmailAddressesArray "dummyexternaluser@spuriousdomain.com" -friendlyPermissionsLevel "Read" -sendEmail $false -customEmailMessageContent $null -additivePermission $true -allowExternalSharing $true
        set-listInheritance -credentials $credentials -webUrl $webUrl -siteCollection $siteCollection -sitePath $sitePath$siteUrlEndStub -listname "Access Requests" -enableInheritance $true
        delete-allItemsInList -credentials $credentials -webUrl $webUrl -siteCollection $siteCollection -sitePath $sitePath$siteUrlEndStub -listname "Access Requests" -areYouReallySure "YesIAmReallySure" # "Access Requests" 
        }

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
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Taxonomy") 

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
function combine-url($arrayOfStrings){
    $output = ""
    $arrayOfStrings | % {
        $output += $_.TrimStart("/").TrimEnd("/")+"/"
        }
    $output = $output.Substring(0,$output.Length-1)
    $output = $output.Replace("//","/").Replace("//","/").Replace("//","/")
    $output = $output.Replace("http:/","http://").Replace("https:/","https://")
    $output
    }
function convert-listItemToCustomObject($spoListItem, $spoTaxonomyData, $debugMe){
    $alwaysExclude = @("Activities","AttachmentFiles","Attachments","AuthorId","ComplianceAssetId","ContentType","ContentTypeId","EditorId","FieldValuesAsHtml","FieldValuesAsText","FieldValuesForEdit","File","FileSystemObjectType","FirstUniqueAncestorSecurableObject","Folder","GetDlpPolicyTip","ID-dup","OData__UIVersionString","ParentList","Properties","RoleAssignments","ServerRedirectedEmbedUrl","ServerRedirectedEmbedUri","Versions")
    #Ignoring the standard fields listed above, go through each remainig property and process it based on its Type
    $customObj = [psobject]::new()
    $spoListItem | Get-Member -MemberType NoteProperty | ?{$alwaysExclude -notcontains $_.Name} | % {
        $ourMember = $_
        #$ourMember = $spoListItem | Get-Member -MemberType NoteProperty | ?{$_.Name -eq "Id"}
        #$($ourMember.Name)
        #$spoListItem.$($ourMember.Name)
        #$spoListItem.$($ourMember.Name).GetType().Name
        try{
            switch ($spoListItem.$($ourMember.Name).GetType().Name){
                "String" {$customObj | Add-Member -Name $($ourMember.Name.Replace("_x0020_","_")) -MemberType NoteProperty -Value $(sanitise-stripHtml -dirtyString $($spoListItem.$($ourMember.Name)))}
                "Int32" {$customObj | Add-Member -Name $($ourMember.Name.Replace("_x0020_","_")) -MemberType NoteProperty -Value $spoListItem.$($ourMember.Name)}
                "Bool" {$customObj | Add-Member -Name $($ourMember.Name.Replace("_x0020_","_")) -MemberType NoteProperty -Value $spoListItem.$($ourMember.Name)}
                "Boolean" {$customObj | Add-Member -Name $($ourMember.Name.Replace("_x0020_","_")) -MemberType NoteProperty -Value $spoListItem.$($ourMember.Name)}
                "PSCustomObject" {
                    #Now for the complicated stuff
                    #Check for duff data first:
                    if($spoListItem.$($ourMember.Name).__deferred -ne $null){$customObj | Add-Member -Name $($ourMember.Name.Replace("_x0020_","_")) -MemberType NoteProperty -Value "Value was deferred. Check you expanded it correctly in the spoQuery"}
                    else{
                        #Is it a single-value User/Group?
                        if($spoListItem.$($ourMember.Name).Title -ne $null){if($debugMe){Write-Host -ForegroundColor DarkCyan "$($spoListItem.Title).$($ourMember.Name) is a single-value user/group field"}
                            #Process users and groups differently (get user UPN or group Title):
                            if($spoListItem.$($ourMember.Name).Name -match 'i:0\#\.f\|membership\|'){$customObj | Add-Member -Name $($ourMember.Name.Replace("_x0020_","_")) -MemberType NoteProperty -Value $spoListItem.$($ourMember.Name).Name.Replace("i:0#.f|membership|","")}
                            else{$customObj | Add-Member -Name $($ourMember.Name.Replace("_x0020_","_")) -MemberType NoteProperty -Value $spoListItem.$($ourMember.Name).Title -Force}
                            }
                        #Is it a single-value metadata field?
                        elseif($spoListItem.$($ourMember.Name).TermGuid -ne $null){if($debugMe){Write-Host -ForegroundColor DarkCyan "$($spoListItem.Title).$($ourMember.Name) is a single-value metadata field"}
                            $customObj | Add-Member -MemberType NoteProperty -Name $($ourMember.Name.Replace("_x0020_","_")) -Value $(($spoTaxonomyData | ?{$_.IdForTerm -eq $spoListItem.$($ourMember.Name).TermGuid})[0] | %{$_.Title}) #-Force
                            }
                        #Is it a multi-value field?
                        elseif($spoListItem.$($ourMember.Name).results.Count -gt 0){if($debugMe){Write-Host -ForegroundColor DarkCyan "$($spoListItem.Title).$($ourMember.Name) is a multi-value field"}
                            #Is it a multi-value User/Group?
                            if($spoListItem.$($ourMember.Name).results[0].Title -ne $null){if($debugMe){Write-Host -ForegroundColor DarkCyan "$($spoListItem.Title).$($ourMember.Name) is a multi-value user/group field"}
                                $userOrGroups = @()
                                foreach($userOrGroup in $spoListItem.$($ourMember.Name).results){$userOrGroups += $userOrGroup.Title}
                                $customObj | Add-Member -MemberType NoteProperty -Name $($ourMember.Name.Replace("_x0020_","_")) -Value $userOrGroups -Force
                                }
                            #Is it a multi-value metadata field?
                            if($spoListItem.$($ourMember.Name).results[0].TermGuid -ne $null){if($debugMe){Write-Host -ForegroundColor DarkCyan "$($spoListItem.Title).$($ourMember.Name) is a multi-value metadata field"}
                                $userOrGroups = @()
                                foreach($userOrGroup in $spoListItem.$($ourMember.Name).results){$userOrGroups += $userOrGroup.Label}
                                $customObj | Add-Member -MemberType NoteProperty -Name $($ourMember.Name.Replace("_x0020_","_")) -Value $userOrGroups -Force
                                }
                            }
                        #Otherwise, just add it
                        else{$customObj | Add-Member -MemberType NoteProperty -Name $($ourMember.Name.Replace("_x0020_","_")) -Value $spoListItem.$($ourMember.Name)}
                        }
                    }
                default {$customObj | Add-Member -Name $($ourMember.Name.Replace("_x0020_","_")) -MemberType NoteProperty -Value "Uh-oh, someone left a sponge in the patient :(" -Force} 
                }
            }
        #Not clever, but only $nulls should land here. Ha - "should".
        catch{
            write-host $_
            if ($($customObj | Get-Member -Name $($ourMember.Name.Replace("_x0020_","_"))) -eq $null){
                $customObj | Add-Member -Name $($ourMember.Name.Replace("_x0020_","_")) -MemberType NoteProperty -Value $null
                }
            }
        }
    $customObj
    }
function copy-allFilesAndFolders($credentials, $webUrl, $sourceCtx, $sourceSiteCollectionPath, $sourceSitePath, $sourceLibraryName, $sourceFolderPath, $destCtx, $destSiteCollectionPath, $destSitePath, $destLibraryName, $destFolderPath, [boolean]$overwrite){
    if(!$sourceCtx){$sourceCtx = new-csomContext -fullSitePath $($webUrl+$sourceSiteCollectionPath+$sourceSitePath) -sharePointCredentials $credentials}
    if(!$destCtx){$destCtx = new-csomContext -fullSitePath $($webUrl+$destSiteCollectionPath+$destSitePath) -sharePointCredentials $credentials}
    if(!$destSiteCollectionPath){$destSiteCollectionPath = $sourceSiteCollectionPath}
    if(!$destSitePath){$destSitePath = $sourceSitePath}
    if(!$destLibraryName){$destLibraryName = $sourceLibraryName}
    if(!$destFolderPath){$destFolderPath = $sourceFolderPath}

    $srcLib = $sourceCtx.Web.Lists.GetByTitle($sourceLibraryName)
    $srcQuery = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
    $srcQuery.FolderServerRelativeUrl = $(combine-url -arrayOfStrings @("/",$sourceSiteCollectionPath,$sourceSitePath,$sourceLibraryName,$sourceFolderPath))
    $srcQuery.FolderServerRelativeUrl = $srcQuery.FolderServerRelativeUrl.Replace("/Documents/","/Shared Documents/")
    Write-Host -ForegroundColor DarkYellow $srcQuery.FolderServerRelativeUrl
    $srcItems = $srcLib.GetItems($srcQuery)
    $sourceCtx.Load($srcItems)
    $sourceCtx.ExecuteQuery()


    #write-host -ForegroundColor DarkYellow $srcItems.Count " items found"
    foreach($item in $srcItems){
        switch ($item.FileSystemObjectType){
            ([Microsoft.SharePoint.Client.FileSystemObjectType]::File){
                $file = $item.File
                $sourceCtx.Load($file)
                $sourceCtx.ExecuteQuery() 
                Write-Host -ForegroundColor Yellow "Copying:" $file.ServerRelativeUrl
                #copy-file -credentials $credentials -webUrl $webUrl -sourceCtx $sourceCtx -sourceLibraryName $sourceLibraryName -sourceFolderPath $sourceFolderPath -sourceFileName $file.Name -destCtx $destCtx -destLibraryName $destLibraryName -destFolderPath $destFolderPath -destFileName $file.Name -overwrite $overwrite
                }
            ([Microsoft.SharePoint.Client.FileSystemObjectType]::Folder){
                $folder= $item.Folder
                $sourceCtx.Load($folder)
                $sourceCtx.ExecuteQuery()

                $destLib = $destCtx.Web.Lists.GetByTitle($destLibraryName)
                $destCtx.Load($destLib)
                $destCtx.ExecuteQuery()

                #Enable folder creation if necessary
                if([string]::IsNullOrEmpty($destFolderPath) -or "/" -eq $destFolderPath){
                    if(!$destLib.EnableFolderCreation){$destLib.EnableFolderCreation = $true;$destLib.Update();$destCtx.ExecuteQuery()}
                    }

                #Create the folder
                $creationinfo = [Microsoft.SharePoint.Client.ListItemCreationInformation]::new()
                $creationinfo.UnderlyingObjectType = [Microsoft.SharePoint.Client.FileSystemObjectType]::Folder
                $creationinfo.LeafName = $folder.Name
                $creationinfo.FolderUrl = $(combine-url -arrayOfStrings @("/",$destSiteCollectionPath, $destSitePath, $destLibraryName, $destFolderPath))
                $creationinfo.FolderUrl = $creationinfo.FolderUrl.Replace("/Documents/","/Shared Documents/")
                Write-Host -ForegroundColor Yellow "Creating:" $(combine-url @("/",$creationinfo.FolderUrl, $creationinfo.LeafName))

                $newFolder = $destLib.AddItem($creationinfo)
                $newFolder["Title"] = $folder.Name
                $newFolder.Update()
                $destCtx.ExecuteQuery()                
                #Recurse
                #Write-Host -ForegroundColor DarkYellow "copy-allFilesAndFolders -credentials $credentials -webUrl $webUrl -sourceCtx $sourceCtx -sourceSiteCollectionPath $sourceSiteCollectionPath -sourceSitePath $sourceSitePath -sourceLibraryName $sourceLibraryName -sourceFolderPath $(combine-url -arrayOfStrings @($sourceFolderPath, $folder.Name)) -destCtx $destCtx -destSiteCollectionPath $destSiteCollectionPath -destSitePath $destSitePath -destLibraryName $destLibraryName -destFolderPath $(combine-url -arrayOfStrings @($destFolderPath,$folder.Name)) -overwrite $overwrite"
                #copy-allFilesAndFolders -credentials $credentials -webUrl $webUrl -sourceCtx $sourceCtx -sourceSiteCollectionPath $sourceSiteCollectionPath -sourceSitePath $sourceSitePath -sourceLibraryName $sourceLibraryName -sourceFolderPath $(combine-url -arrayOfStrings @($sourceFolderPath, $folder.Name)) -destCtx $destCtx -destSiteCollectionPath $destSiteCollectionPath -destSitePath $destSitePath -destLibraryName $destLibraryName -destFolderPath $(combine-url -arrayOfStrings @($destFolderPath,$folder.Name)) -overwrite $overwrite
                }
            }
        }
    }
function copy-file($credentials, $webUrl, $sourceCtx, $sourceSiteCollectionPath, $sourceSitePath, $sourceLibraryName, $sourceFolderPath, $sourceFileName, $destCtx, $destSiteCollectionPath, $destSitePath, $destLibraryName, $destFolderPath, $destFileName, [boolean]$overwrite){
    if(!$sourceCtx){$sourceCtx = new-csomContext -fullSitePath $($webUrl+$sourceSiteCollectionPath+$sourceSitePath) -sharePointCredentials $credentials}
    if(!$destCtx){$destCtx = new-csomContext -fullSitePath $($webUrl+$destSiteCollectionPath+$destSitePath) -sharePointCredentials $credentials}
    if(!$destSiteCollectionPath){$destSiteCollectionPath = $sourceSiteCollectionPath}
    if(!$destSitePath){$destSitePath = $sourceSitePath}
    if(!$destLibraryName){$destLibraryName = $sourceLibraryName}
    if(!$destFolderPath){$destFolderPath = $sourceFolderPath}
    if(!$destFileName){$destFileName = $sourceFileName}

    $sanitisedSourceLibraryName = sanitise-LibraryNameForUrl -dirtyString $sourceLibraryName
    $sanitisedSourceFileName = sanitise-forSharePointFileName -dirtyString $sourceFileName
    $sanitisedSourceFullRelativePath = [uri]::EscapeUriString($sourceSiteCollectionPath+"/"+$sourceSitePath+"/"+$sanitisedSourceLibraryName+"/"+$sourceFolderPath+"/"+$sanitisedSourceFileName).Replace("//","/").Replace("//","/").Replace("//","/")
    $sanitisedSourceFullRelativePath = ($sourceSiteCollectionPath+"/"+$sourceSitePath+"/"+$sanitisedSourceLibraryName+"/"+$sourceFolderPath+"/"+$sanitisedSourceFileName).Replace("//","/").Replace("//","/").Replace("//","/")

    $sanitisedDestLibraryName = sanitise-LibraryNameForUrl -dirtyString $destLibraryName
    $sanitisedDestFileName = sanitise-forSharePointFileName -dirtyString $destFileName
    $sanitisedDestFullRelativePath = [uri]::EscapeUriString($destSiteCollectionPath+"/"+$destSitePath+"/"+$sanitisedDestLibraryName+"/"+$destFolderPath+"/"+$sanitisedDestLibraryName).Replace("//","/").Replace("//","/").Replace("//","/")
    $sanitisedDestFullRelativePath = ($destSiteCollectionPath+"/"+$destSitePath+"/"+$sanitisedDestLibraryName+"/"+$destFolderPath+"/"+$sanitisedDestFileName).Replace("//","/").Replace("//","/").Replace("//","/")

    Write-Host $sanitisedSourceFullRelativePath
    Write-Host $sanitisedDestFullRelativePath
    $fileObj = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($sourceCtx,$sanitisedSourceFullRelativePath)
    [Microsoft.SharePoint.Client.File]::SaveBinaryDirect($destCtx,$sanitisedDestFullRelativePath,$fileObj.Stream,$overwrite)
    $fileObj.Dispose()
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
function delete-allItemsInList($credentials, $webUrl, $siteCollection, $sitePath, $listname, $areYouReallySure){
    if("YesIAmReallySure" -eq $areYouReallySure){
        $ctx = new-csomContext -fullSitePath ($webUrl+$siteCollection+$sitePath) -sharePointCredentials $credentials
        $list = $ctx.Web.Lists.GetByTitle($listname)
        $ctx.Load($list)
        $items = $list.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())
        $ctx.Load($items)
        $ctx.ExecuteQuery()
        $items.GetEnumerator() | % {
            $_.deleteObject()
            }
        $ctx.ExecuteQuery()
        }
        else{"Sorry - please supply `"YesIAmReallySure`" as the string value for `$areYouReallySure to really delete stuff"}
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
function get-listID($credentials, $webUrl, $siteCollection, $sitePath, $listname){
    $ctx = new-csomContext -fullSitePath ($webUrl+$siteCollection+$sitePath) -sharePointCredentials $credentials
    $list = $ctx.Web.Lists.GetByTitle($listname)
    $ctx.Load($list)
    $ctx.ExecuteQuery()
    $list.Id
    }
function get-list($credentials, $webUrl, $siteCollection, $sitePath, $listname){
    $ctx = new-csomContext -fullSitePath ($webUrl+$siteCollection+$sitePath) -sharePointCredentials $credentials
    $list = $ctx.Web.Lists.GetByTitle($listname)
    $ctx.Load($list)
    $ctx.ExecuteQuery()
    $list
    }
function get-file($ctx, $credentials, $webUrl, $siteCollection, $sitePath, $listname, $folderPath, $fileName){
    if (!$ctx){$ctx = new-csomContext -fullSitePath ($webUrl+$siteCollection+$sitePath) -sharePointCredentials $credentials}
    #$file = $ctx.Web.GetFileByServerRelativePath([Microsoft.SharePoint.Client.ResourcePath]::FromDecodedUrl('/teams/communities/heathandsafetyteam/Shared%20Documents/RAs/Projects/Anthesis%20UK%20Project%20Risk%20Assessment.xlsx'))
    #$file = $ctx.Web.GetFileByServerRelativePath([Microsoft.SharePoint.Client.ResourcePath]::FromDecodedUrl('/teams/communities/heathandsafetyteam/Shared Documents/RAs/Projects/Anthesis UK Project Risk Assessment.xlsx'))
    Write-Host $('$ctx.Web.GetFileByServerRelativePath([Microsoft.SharePoint.Client.ResourcePath]::FromDecodedUrl("'+"$siteCollection$sitePath/$listname$folderPath/$fileName"+'))')
    $ctx.Load($file)
    $ctx.ExecuteQuery()
    $file
    }
function get-SPOGroup($ctx, $credentials, $webUrl, $siteCollection, $sitePath, $groupName){
    if ($ctx -eq $null){$ctx = new-csomContext -fullSitePath ($webUrl+$siteCollection+$sitePath) -sharePointCredentials $credentials}
    $group = $ctx.Web.SiteGroups.GetByName($groupName)
    $ctx.Load($group)
    $ctx.ExecuteQuery()

    #$groups = $ctx.Web.SiteGroups
    #$ctx.Load($groups)
    #$ctx.ExecuteQuery()
    #$groups.GetEnumerator() | % {
    #    if ($_.Title -eq $groupName){
    #        $group = $_ 
    #        $ctx.Load($group)
    #        $ctx.ExecuteQuery()
    #        }
    #    }
    $group
    }
function get-spoLocaleFromCountry($p3LetterCountryIsoCode){
    #$countryToLocaleHashTable = @{"Canada"="4105";"China"="2052";"Finland"="1035";"Germany"="1031";"Korea"="1042";"Spain"="1034";"Sri Lanka"="1097";"Philippines"="13321";"Sweden"="1053";"United Arab Emirates"="";"United Kingdom"="2057";"United States"="1033"}
    $countryToLocaleHashTable = @{"CAN"="4105";"CHN"="2052";"FIN"="1035";"DEU"="1031";"KOR"="1042";"ESP"="1034";"LKA"="1097";"PHL"="13321";"SWE"="1053";"ARE"="";"GBR"="2057";"USA"="1033"}
    $countryToLocaleHashTable[$p3LetterCountryIsoCode]
    }
function get-spoTimeZoneHashTable($credentials){
    $ctx = new-csomContext -fullSitePath "https://anthesisllc.sharepoint.com" -sharePointCredentials $credentials
    $tz = $ctx.Web.RegionalSettings.TimeZones
    $ctx.Load($tz) | Out-Null
    $tzEnum = $ctx.Web.RegionalSettings.TimeZones.GetEnumerator()
    $ctx.ExecuteQuery() | Out-Null
    $spoTimeZones = @{}
    while($tzEnum.MoveNext()){
        $spoTimeZones.Add($tzEnum.Current.Description, $tzEnum.Current.Id) | Out-Null
        }
    $spoTimeZones
    }
function get-termsInSet($credentials, $webUrl, $siteCollection, $pGroup,$pSet){
    #Sanitise the input:
    $group = sanitise-forTermStore $pGroup
    $set = sanitise-forTermStore $pSet
    $lcid = 1033 #Probably Language
    $ctx = new-csomContext -fullSitePath ($webUrl+$siteCollection) -sharePointCredentials $credentials
    $taxonomySession = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($ctx)
    $taxonomySession.UpdateCache()
    $ctx.Load($taxonomySession)
    $ctx.ExecuteQuery()
    if (!$taxonomySession.ServerObjectIsNull){
        Write-Host "Destination Taxonomy session initiated: " $taxonomySession.Path.Identity "" -ForegroundColor Green
        $termStore = $taxonomySession.GetDefaultSiteCollectionTermStore()
        $ctx.Load($termStore)
        $termGroup = $termStore.Groups.GetByName($group)
        $ctx.Load($termGroup)
        $termSet = $termGroup.TermSets.GetByName($set)
        $ctx.Load($termSet)
        $ctx.ExecuteQuery()        #Delete the Term if necessary
        $ctx.Load($termSet.Terms)
        $termsEnum = $termSet.Terms.GetEnumerator()
        $ctx.ExecuteQuery()
        while ($termsEnum.MoveNext()) {
            [array]$arrayOfTerms += $termsEnum.Current.Name
            }
        $arrayOfTerms
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
function new-sharingInvite($credentials, $webUrl, $siteCollection, $sitePath, $userEmailAddressesArray, $friendlyPermissionsLevel, $sendEmail, $customEmailMessageContent, $additivePermission, $allowExternalSharing){
    #Make a new context and load the Web
    $ctx = new-csomContext -fullSitePath ($webUrl+$siteCollection+$sitePath) -sharePointCredentials $credentials
    $web = $ctx.Web
    $ctx.Load($web)
    
    #Get the Role Definition
    switch ($friendlyPermissionsLevel){
        "Full Control" {$sharingRole = [Microsoft.SharePoint.Client.Sharing.Role]::Owner}
        "Edit" {$sharingRole = [Microsoft.SharePoint.Client.Sharing.Role]::Edit}
        "Contribute" {$sharingRole = [Microsoft.SharePoint.Client.Sharing.Role]::Edit}
        "Read" {$sharingRole = [Microsoft.SharePoint.Client.Sharing.Role]::View}
        "None" {$sharingRole = [Microsoft.SharePoint.Client.Sharing.Role]::None}
        default {"Uh-oh, someone left a sponge in the patient - `$friendlyPermissionsLevel value of $friendlyPermissionsLevel is invalid";break}
        }

    #Assign the Role Definition to each e-mail address in $userEmailAddressesArray
    [System.Reflection.Assembly]::LoadWithPartialName("System.Collections") | Out-Null
    $roleAssignments = New-Object "System.Collections.Generic.List[Microsoft.SharePoint.Client.Sharing.UserRoleAssignment]"
    $userEmailAddressesArray | %{
        #$ctx.Load($roleDef.Add($_, $roleDefBind))
        $roleAssignment = New-Object Microsoft.SharePoint.Client.Sharing.UserRoleAssignment
        $roleAssignment.UserId = $_
        $roleAssignment.Role = $sharingRole
        $roleAssignments.Add($roleAssignment)
        }
    
    [Microsoft.SharePoint.Client.Sharing.WebSharingManager]::UpdateWebSharingInformation($ctx, $ctx.Web, $roleAssignments, $sendEmail, $customEmailMessageContent, $additivePermission, $allowExternalSharing)
    $ctx.ExecuteQuery()
    }
function new-SPOGroup($ctx, $title, $description, $pGroupOwner, $spoSite){
    $spoGroupCreationInfo=New-Object Microsoft.SharePoint.Client.GroupCreationInformation
    $spoGroupCreationInfo.Title=$title
    $spoGroupCreationInfo.Description=$description
    $newGroup = $spoSite.SiteGroups.Add($spoGroupCreationInfo)
    if ($pGroupOwner -ne $null){
        if($pGroupOwner -eq $title){$newGroup.Owner = $newGroup} #If the new group is owned by itself, just do it now.
        else{
            try{
                $groupOwner = $ctx.Web.SiteGroups.GetByName($pGroupOwner) #Otherwise see if $pGroupOwner is a group
                $ctx.Load($groupOwner)
                $ctx.ExecuteQuery()
                }
            catch{
                try{
                    $groupOwner = $ctx.Web.EnsureUser($pGroupOwner)#Otherwise, see if $pGroupOwner is a User
                    $ctx.Load($groupOwner)
                    $ctx.ExecuteQuery()
                    }
                catch{Write-Host "$pGroupOwner is not a valid SPOGroup or User :(";$groupOwner = $null}
                }
            }
        }
    if($groupOwner -ne $null){
        $newGroup.Owner = $groupOwner
        $newGroup.Update()
        }
    $ctx.ExecuteQuery()
    $newGroup
    }
function remove-userFromSite($credentials, $webUrl, $siteCollection, $sitePath, $memberToRemove){
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
        $ctx.ExecuteQuery()
        #$member.Title
        if($member.Title -eq $userToRemove.Title){$ctx.Web.RoleAssignments.GetByPrincipalId($member.Id).DeleteObject();$found = $true}
        }
    if($found){
        $ctx.ExecuteQuery()
        "$memberToRemove removed"
        }
        else{"$memberToRemove not found"}
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
            #$term.Name = "BBC"
            $ctx.ExecuteQuery()
           }
        }
    }
function reuse-allTermsInTermStore($credentials,$webUrl,$siteCollection,$sourceTermGroup,$sourceTermSet,$destTermGroup,$destTermSet){
    $termsToReUse = get-termsInSet -credentials $credentials -webUrl $webUrl -siteCollection $siteCollection -pGroup $sourceTermGroup -pSet $sourceTermSet
    $termsToReUse | % {reuse-termInStore -credentials $credentials -webUrl $webUrl -siteCollection $siteCollection -pCurrentGroup $sourceTermGroup -pCurrentSet $sourceTermSet -pTerm $_ -pAdditionalGroup $destTermGroup -pAdditionalSet $destTermSet}
    }
function reuse-termInStore($credentials, $webUrl, $siteCollection, $pCurrentGroup,$pCurrentSet,$pTerm,$pAdditionalGroup,$pAdditionalSet){
    #Sanitise the input:
    $currentGroup = sanitise-forTermStore $pCurrentGroup
    $currentSet = sanitise-forTermStore $pCurrentSet
    $additionalGroup = sanitise-forTermStore $pAdditionalGroup
    $additionalSet = sanitise-forTermStore $pAdditionalSet
    $term = sanitise-forTermStore $pTerm
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
        $currentTermGroup = $termStore.Groups.GetByName($currentGroup)
        $ctx.Load($currentTermGroup)
        $currentTermSet = $currentTermGroup.TermSets.GetByName($currentSet)
        $ctx.Load($currentTermSet)
        $ctx.Load($currentTermSet.Terms)
        $targetTermGroup = $termStore.Groups.GetByName($additionalGroup)
        $ctx.Load($targetTermGroup)
        $targetTermSet = $targetTermGroup.TermSets.GetByName($additionalSet)
        $ctx.Load($targetTermSet)
        $ctx.ExecuteQuery()
        if(($currentTermSet.Terms.Name) -contains $term){
            $termToReuse = $currentTermSet.Terms.GetByName($term)
            $ctx.Load($termToReuse)
            $ctx.ExecuteQuery()
            Write-Host "Reusing term: [$($termToReuse.Name)] from [$($currentTermSet.Name)] in [$($targetTermSet.Name)]" -ForegroundColor Green
            $targetTermSet.ReuseTerm($termToReuse,$false) | Out-Null #This weirdly throws a "collection has not been initialized" error, but works fine if you ignore it
            $ctx.ExecuteQuery()
           }
        }
    }
function sanitise-forSharePointFileName($dirtyString){ 
    $dirtyString = $dirtyString.Trim()
    $dirtyString.Replace("`"","").Replace("#","").Replace("%","").Replace("?","").Replace("<","").Replace(">","").Replace("\","").Replace("/","").Replace("...","").Replace("..","").Replace("'","`'")
    if(@("."," ") -contains $dirtyString.Substring(($dirtyString.Length-1),1)){$dirtyString = $dirtyString.Substring(0,$dirtyString.Length-1)} #Trim trailing "."
    }
function sanitise-LibraryNameForUrl($dirtyString){
    $cleanerString = $dirtyString.Trim()
    $cleanerString = $dirtyString -creplace '[^a-zA-Z0-9 _/]+', ''
    $cleanerString
    }
function new-csomCredentials($username, $password){
    if ($username -eq $null -or $username -eq ""){$username = Read-Host -Prompt "Enter SharePoint Online username (blank for $($env:USERNAME)@anthesisgroup.com)"}
    if ($username -eq $null -or $username -eq ""){$username = "$($env:USERNAME)@anthesisgroup.com"}
    if ($password -eq $null -or $password -eq ""){$password = Read-Host -Prompt "Password for $username" -AsSecureString}
    New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $password)
    }
function set-listInheritance($credentials, $webUrl, $siteCollection, $sitePath, $listname, $enableInheritance){
    $ctx = new-csomContext -fullSitePath ($webUrl+$siteCollection+$sitePath) -sharePointCredentials $credentials
    $list = $ctx.Web.Lists.GetByTitle($listname)
    $ctx.Load($list)
    if($enableInheritance){$list.ResetRoleInheritance()}
        else{$list.BreakRoleInheritance($true, $true)} #1st Arg = Copy permissions from parent?; 2nd Arg = Remove any unquire permissions?
    $list.Update()
    $ctx.ExecuteQuery()
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
function set-SPOGroupAsDefault($credentials, $webUrl, $siteCollection, $sitePath, $groupName, $defaultForWhat){
    $ctx = new-csomContext -fullSitePath ($webUrl+$siteCollection+$sitePath) -sharePointCredentials $credentials
    $group = get-SPOGroup -ctx $ctx -webUrl $webUrl -siteCollection $siteCollection -sitePath $sitePath -groupName $groupName
    $web =$ctx.Web
    $ctx.Load($web)
    switch ($defaultForWhat){
        "Owners" {$web.AssociatedOwnerGroup = $group}
        "Editors" {$web.AssociatedMemberGroup = $group}
        "Members" {$web.AssociatedMemberGroup = $group}
        "Visitors" {$web.AssociatedVisitorGroup = $group}
        default {"Uh-oh, `$defaultForWhat needs to be a string containing `"Owners`", `"Editors`", `"Members`", or `"Visitors`""}
        }
    $web.Update()
    $ctx.ExecuteQuery()
    $ctx.Dispose()
    }
