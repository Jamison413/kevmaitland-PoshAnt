$365creds = set-MsolCredentials
$teamBotDetails = import-encryptedCsv -pathToEncryptedCsv "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\teambotdetails.txt"

$tokenResponse = get-graphTokenResponse -aadAppCreds $teamBotDetails
Connect-PnPOnline -AccessToken $tokenResponse.access_token
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/" -Credentials $365creds
$allPnpUnifiedGroups = Get-PnPUnifiedGroup

$allPnpSites.count
$allPnpUnifiedGroups |  % {
    $thisSite = $_
    $tokenResponse = get-graphTokenResponse -aadAppCreds $teamBotDetails
    Connect-PnPOnline -AccessToken $tokenResponse.access_token
    $userWasAlreadySiteAdmin = test-isUserSiteCollectionAdmin -pnpUnifiedGroupObject $thisSite -accessToken $tokenResponse.access_token -pnpCreds $365creds -addPermissionsIfMissing $true
    Connect-PnPOnline  $thisSite.SiteUrl -Credentials $365creds
    Write-Host -ForegroundColor Yellow "Site [$($thisSite.DisplayName)][$($thisSite.SiteUrl)]"
    $dummyFolderName = "DummyShareToDelete"

    $allDocLibsInThisSite = Get-PnPList -Includes Fields | ? {$_.TemplateFeatureId -eq "00bfea71-e717-4e80-aa17-d0c71b360101" -and @("_catalogs/hubsite","Form Templates","Site Assets","Site Pages","Style Library","TaxonomyHiddenList") -notcontains $_.Title}
    $allDocLibsInThisSite | % {
        $thisDocLib = $_
        $docLibName = $(Split-Path $thisDocLib.RootFolder.ServerRelativeUrl -Leaf)
        Write-Verbose "`tAdding Folder [$dummyFolderName] to [$docLibName]"
        Add-PnPFolder -Name $dummyFolderName -Folder $docLibName
        $dummyPnpFolderItem = Get-PnPFolderItem -FolderSiteRelativeUrl $docLibName -ItemType Folder -ItemName $dummyFolderName
        #Set-PnPListItemPermission -List $defaultDocLib.Id -Identity $dummyPnpFolderItem.ListItemAllFields -AddRole Contribute -User $($pnpCreds.UserName) This sets the permissions, but doesn't trigger a SharedWith event
        [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Sharing")  | Out-Null
        [System.Reflection.Assembly]::LoadWithPartialName("System.Collections") | Out-Null
        $roleAssignments = New-Object "System.Collections.Generic.List[Microsoft.SharePoint.Client.Sharing.UserRoleAssignment]"
        $roleAssignment = New-Object Microsoft.SharePoint.Client.Sharing.UserRoleAssignment
        $roleAssignment.UserId = $365creds.UserName
        $roleAssignment.Role = [Microsoft.SharePoint.Client.Sharing.Role]::Edit
        $roleAssignments.Add($roleAssignment)
        [Microsoft.SharePoint.Client.Sharing.DocumentSharingManager]::UpdateDocumentSharingInfo($dummyPnpFolderItem.Context,"https://anthesisllc.sharepoint.com"+$dummyPnpFolderItem.ServerRelativeUrl,$roleAssignments,$false,$true,$false,"",$false,$false)
        Write-Verbose "`tSharing Folder [$dummyFolderName] with [$($365creds.UserName)] via CSOM"
        $dummyPnpFolderItem.Context.ExecuteQuery() #This errors, but still adds the SharedWithUsers column to the Site
        Write-Verbose "`tRemoving Folder [$dummyFolderName]"
        Remove-PnPFolder -Name $dummyFolderName -Folder $docLibName -Force


        Write-Verbose "Setting default View in Documents Library"
        $defaulDocLibPnpView = $thisDocLib | Get-PnPView | ? {$_.DefaultView -eq $true}
        $defaulDocLibPnpView | Set-PnPView -Fields "DocIcon","LinkFilename","Modified","Editor","Created","Author","FileSizeDisplay","SharedWithUsers"
        }
    
    Remove-PnPSiteCollectionAdmin -Owners $365creds.UserName

    #if($userWasAlreadySiteAdmin){}
    #else{
    #    Remove-PnPSiteCollectionAdmin -Owners $365creds.UserName
    #    }
    }