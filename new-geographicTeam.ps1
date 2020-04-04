$365creds = set-MsolCredentials
connect-to365 -credential $365creds

$allAdminUGs= get-unifiedgroup -Filter "DisplayName -like 'Admin*'"  

$teamBotDetails = import-encryptedCsv -pathToEncryptedCsv "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\teambotdetails.txt"
$tokenResponse = get-graphTokenResponse -aadAppCreds $teamBotDetails
$tokenResponse = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponse -renewTokenExpiringInSeconds 30 -aadAppCreds $teamBotDetails


foreach($team in @("All Homeworkers (North America)","All Homeworkers (PHL)","All Madrid (ESP)","All Manchester (GBR)","All Manlleu (ESP)")){
    $tokenResponse = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponse -renewTokenExpiringInSeconds 300 -aadAppCreds $teamBotDetails

    $displayName = $team
    $areDataManagersLineManagers = $false
    $managedBy = "AAD"
    #$memberOf = ??
    $hideFromGal = $false
    $blockExternalMail = $true
    $accessType = "Private"
    $autoSubscribe = $true
    $groupClassification = "Internal"
    $alsoCreateTeam = $false #This doesn't work (race condition)
    $horriblyUnformattedStringOfManagers = "kevin.maitland@anthesisgroup.com"
    $horriblyUnformattedStringOfMembers = "
    "
    

    #region Get the Managers and Members in the right formats
    $managers = @()
    $originalManagers = convertTo-arrayOfEmailAddresses $horriblyUnformattedStringOfManagers | sort | select -Unique
    $managers = $originalManagers #So we can e-mail the right people at the end.
    $members = @()
    $members += convertTo-arrayOfEmailAddresses $horriblyUnformattedStringOfMembers | sort | select -Unique
    $members | % {
        $thisEmail = $_
        try{
            $dg = Get-DistributionGroup -Identity $thisEmail -ErrorAction Stop
            if($dg){
                $members += $(enumerate-nestedDistributionGroups -distributionGroupObject $dg -Verbose).WindowsLiveId
                $members = $members | ? {$_ -ne $thisEmail}
                }
            }
        catch{<# Anything that isn't an e-mail address for a Distribution Group will cause errors here, and we don't really care about them #>}
        }
    $members = $members | Sort-Object | select -Unique

    #See if we need to temporarily add the executing user as 
    if($managers -notcontains ($365creds.UserName)){
        $addExecutingUserAsTemporaryOwner = $true
        [array]$managers += ($365creds.UserName)
        }
    if($members -notcontains ($365creds.UserName)){
        $addExecutingUserAsTemporaryMember = $true
        [array]$members += ($365creds.UserName)
        }

    if($managedBy -eq "AAD"){$managers = "groupbot@anthesisgroup.com"} #Override the ownership of any aggregated / Parent Functional Teams as these are automated separately

    #endregion


    $newGroup = new-365Group -displayName $displayName -managerUpns $managers -teamMemberUpns $members -memberOf $memberOf -hideFromGal $hideFromGal -blockExternalMail $blockExternalMail -accessType $accessType -autoSubscribe $autoSubscribe -additionalEmailAddresses $additionalEmailAddresses -groupClassification $groupClassification -ownersAreRealManagers $areDataManagersLineManagers -membershipmanagedBy $managedBy -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference -tokenResponse $tokenResponse -alsoCreateTeam $alsoCreateTeam -pnpCreds $365creds
    Write-Verbose "Getting associated PnP and Graph objects for [$($newGroup.DisplayName)] - this is a faster way to do stuff than using the UnifiedGroup object"
    Connect-PnPOnline -AccessToken $tokenResponse.access_token
    $newPnpTeam = Get-PnPUnifiedGroup -Identity $newGroup.ExternalDirectoryObjectId
    $newGraphGroup = get-graphGroupFromUpn -tokenResponse $tokenResponse -groupUpn $newGroup.PrimarySmtpAddress
    $newGraphGroupDrive = get-graphDrives -tokenResponse $tokenResponse -teamUpn $newGroup.PrimarySmtpAddress -returnOnlyDefaultDocumentsLibrary

    #Remove GroupBot if required
    $newTeamOwners = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupUpn $newGroup.PrimarySmtpAddress -returnOnlyLicensedUsers -memberType Owners
    if($newTeamOwners.userPrincipalName -contains "groupbot@anthesisgroup.com"){
        if($newTeamOwners.userPrincipalName -notcontains $365creds.UserName){
            add-graphUsersToGroup -tokenResponse $tokenResponse -graphGroupId $newGraphGroup.id -memberType Owners -graphUserUpns $365creds.UserName
            }
        remove-graphUsersFromGroup -tokenResponse $tokenResponse -graphGroupId $newGraphGroup.id -memberType Owners -graphUserIds "00aa81e4-2e8f-4170-bc24-843b917fd7cf" #This works faster with Ids, so I've hard-coded GroupBot's Id
        remove-graphUsersFromGroup -tokenResponse $tokenResponse -graphGroupId $newGraphGroup.id -memberType Members -graphUserIds "00aa81e4-2e8f-4170-bc24-843b917fd7cf"
        }
    $newGraphTeam = new-graphTeam -tokenResponse $tokenResponse -groupId $newGraphGroup.id -allowMemberCreateUpdateChannels $true -allowMemberDeleteChannels $false #Create the Team if it doesn't already exist

    #region Do as much as we can in Graph because it's much quicker
    ########################################
    #Create, share and delete a dummy folder in this Site to trigger the SharedWith Site Column
    $dummyFolder = add-graphArrayOfFoldersToDrive -graphDriveId $newGraphGroupDrive.id -foldersAndSubfoldersArray "DummyFolder" -tokenResponse $tokenResponse -conflictResolution Replace
    grant-graphSharing -tokenResponse $tokenResponse -driveId $newGraphGroupDrive.id -itemId $dummyFolder.id -sharingRecipientsUpns @($365creds.UserName) -requireSignIn $true -sendInvitation $false -role Write -Verbose
    delete-graphDriveItem -tokenResponse $tokenResponse -graphDriveId $newGraphGroupDrive.id -graphDriveItemId $dummyFolder.id -eTag $dummyFolder.eTag 

    #Create corresponding Regional Folder in associated regional Administration Team site
    $currentRegion = get-3lettersInBrackets -stringMaybeContaining3LettersInBrackets $newGroup.DisplayName
    $associatedRegionalAdminTeam = $allAdminUGs | ? {$_.DisplayName -match $currentRegion}
    if([string]::IsNullOrWhiteSpace($associatedRegionalAdminTeam)){
        switch($newGroup.DisplayName){
            {$_ -match "(North America)" -or $_ -match "(CAN)" -or $_ -match "(USA)"}{
                $associatedRegionalAdminTeam = $allAdminUGs | ? {$_.DisplayName -match "North America"}
                }
            {$_ -match "(COL)" -or $_ -match "(AND)"}{
                $associatedRegionalAdminTeam = $allAdminUGs | ? {$_.DisplayName -match "ESP"}
                }
            }
        }
    if([string]::IsNullOrWhiteSpace($currentRegion)){
        Write-Error "Could not identify regional Administration Team. Cannot proceed with configuring the rest of the Site & Team." ;break
        }

    $associatedRegionalAdminDrive = get-graphDrives -tokenResponse $tokenResponse -teamUpn $associatedRegionalAdminTeam.PrimarySmtpAddress -returnOnlyDefaultDocumentsLibrary
    $newRegionalFolder = add-graphArrayOfFoldersToDrive -graphDriveId $associatedRegionalAdminDrive.id -foldersAndSubfoldersArray @($newGroup.DisplayName) -tokenResponse $tokenResponse -conflictResolution Fail

    #Set Edit permissions for this Team on new folder
    grant-graphSharing -tokenResponse $tokenResponse -driveId $associatedRegionalAdminDrive.id -itemId $newRegionalFolder.id -sharingRecipientsUpns @($newGroup.PrimarySmtpAddress) -requireSignIn $true -sendInvitation $false -role Write

    #Create text file explaining how this works / links to new folder from this Site
    $textFileTemplateContent = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/drives/b!AE2tHi4uHkKRdhUoe1wizoHfHdLv_DZOlObt1vtIejFDr6vvuqdFTaTWzb63-TzY/items/01V67YTVCXN7JPPNRJXBB3TPRS34DU3FA3/content" #This is the content of https://anthesisllc.sharepoint.com/teams/IT_Team_All_365/Shared%20Documents/File%20storage%20is%20disabled%20for%20this%20geographic%20team.txt
    $newTextFile = invoke-graphPut -tokenResponse $tokenResponse -graphQuery "/drives/$($newGraphGroupDrive.id)/items/root:/Try the General»Managed Files tab - file storage is disabled for geographic team $($newGroup.DisplayName).txt:/content" -binaryFileStream $textFileTemplateContent
    $anotherNewTextFile = invoke-graphPut -tokenResponse $tokenResponse -graphQuery "/drives/$($newGraphGroupDrive.id)/items/root:/General/Try the General»Managed Files tab - file storage is disabled for geographic team $($newGroup.DisplayName).txt:/content" -binaryFileStream $textFileTemplateContent
    $newHyperlinkContent = `
    "[InternetShortcut]
    URL=$($associatedRegionalAdminDrive.webUrl+"/"+[uri]::EscapeDataString($newGroup.DisplayName))
    "
    $newHyperlink = invoke-graphPut -tokenResponse $tokenResponse -graphQuery "/drives/$($newGraphGroupDrive.id)/items/root:/Link to storage manged by $($associatedRegionalAdminTeam.DisplayName).url:/content" -binaryFileStream $newHyperlinkContent
    #$anotherNewHyperlink = invoke-graphPut -tokenResponse $tokenResponse -graphQuery "/drives/$($newGraphGroupDrive.id)/items/root:/General/Link to storage manged by $($associatedRegionalAdminTeam.DisplayName).url:/content" -binaryFileStream $newHyperlinkContent #This doesn;t work so well in Teams, but we'll keep the first link as it's only really visible from SharePoint

    #Create new Tab in Teams linking to this location
    $newGraphTeamGeneralChannel = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/teams/$($newGraphGroup.id)/channels"
    $newGraphTeamGeneralChannelTabs = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/teams/$($newGraphGroup.id)/channels/$($newGraphTeamGeneralChannel.id)/tabs"
    $newGraphTeamGeneralChannelTabs | ? {$_.displayName -eq "Managed Files"} | % {
        invoke-graphDelete -tokenResponse $tokenResponse -graphQuery "/teams/$($newGraphGroup.id)/channels/$($newGraphTeamGeneralChannel.id)/tabs/$($_.id)" 
        }
    $tabConfiguration = @{
        "entityId"=$null
        "contentUrl"=$newRegionalFolder.webUrl
        "websiteUrl"=$newRegionalFolder.webUrl
        "removeUrl"=$null
        }
    $tabBody = @{
        "displayName"="Managed Files"
        "teamsApp@odata.bind"="https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.web"
        "configuration"=$tabConfiguration
        }
    $newGraphTeamGeneralChannelManagedFilesTab = invoke-graphPost -tokenResponse $tokenResponse -graphQuery "/teams/$($newGraphGroup.id)/channels/$($newGraphTeamGeneralChannel.id)/tabs" -graphBodyHashtable $tabBody


    #endregion

    #region We still have to do some stuff in PnP because it's not supported in Graph yet
    if($addExecutingUserAsTemporaryOwner){ #copy-spoPage requires Site Collection Admin rights
        $userWasAlreadySiteAdmin = test-isUserSiteCollectionAdmin -pnpUnifiedGroupObject $newPnpTeam -accessToken $tokenResponse.access_token -pnpCreds $365creds -addPermissionsIfMissing $true -Verbose
        }
    copy-spoPage -sourceUrl "https://anthesisllc.sharepoint.com/teams/IT_Team_All_365/SitePages/Candidate-template-for-regional-sites.aspx" -destinationSite $newPnpTeam.SiteUrl -pnpCreds $365creds -overwriteDestinationFile $true -renameFileAs "LandingPage.aspx" -Verbose | Out-Null

    test-pnpConnectionMatchesResource -resourceUrl $newPnpTeam.SiteUrl -pnpCreds $365creds -connectIfDifferent $true | Out-Null
    if((test-pnpConnectionMatchesResource -resourceUrl $newPnpTeam.SiteUrl) -eq $true){
        Write-Verbose "Setting Homepage"
        Set-PnPHomePage  -RootFolderRelativeUrl "SitePages/LandingPage.aspx" | Out-Null
        Write-Verbose "Disabling Comments"
        Set-PnPClientSidePage -Identity "LandingPage.aspx" -CommentsEnabled:$false
        Write-Verbose "ReTitling Homepage"
        $newHomepage = Get-PnPListItem -List "SitePages" -Query "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='Text'>LandingPage.aspx</Value></Eq></Where></Query></View>"
        Set-PnPListItem -Values @{"Title"=$newPnpTeam.DisplayName} -List "SitePages" -Identity $newHomepage.Id

        Write-Verbose "Setting default View in Documents Library"
        $thisDocLib = Get-PnPList -Identity "Documents" -Includes Fields, ParentWeb
        $defaulDocLibPnpView = $thisDocLib | Get-PnPView | ? {$_.DefaultView -eq $true}
        $defaulDocLibPnpView | Set-PnPView -Fields "DocIcon","LinkFilename","Modified","Editor","Created","Author","FileSizeDisplay","SharedWithUsers"

        #Set Site permissions to Members = ReadOnly
        $siteMembersGroup = Get-PnPGroup -AssociatedMemberGroup
        Set-PnPGroupPermissions -Identity $siteMembersGroup -AddRole Read -RemoveRole Edit
        #Set-PnPGroupPermissions -Identity $siteMembersGroup -AddRole Edit -RemoveRole Read


        }

    Add-PnPHubSiteAssociation -Site $newPnpTeam.SiteUrl -HubSite "https://anthesisllc.sharepoint.com/sites/TeamHub" | Out-Null
    #endregion

    Write-Verbose "Opening in browser - don't forget to edit the page to make the last few changes."
    Write-Host -f Yellow "TeamName:`t$($newGroup.DisplayName)"
    Write-Host -f Yellow "TeamLink:`t$($newGraphTeamGeneralChannel.webUrl)"
    Set-Clipboard -Value $newGraphTeamGeneralChannel.webUrl

    start-Process $newPnpTeam.SiteUrl

    Write-Verbose "set-standardSitePermissions [$($newGroup.DisplayName)]"
    set-standardSitePermissions -unifiedGroupObject $newGroup -tokenResponse $tokenResponse -pnpCreds $365creds -Verbose


    if($addExecutingUserAsTemporaryOwner){
        test-pnpConnectionMatchesResource -resourceUrl $newPnpTeam.SiteUrl -pnpCreds $365creds -connectIfDifferent $true | Out-Null
        Remove-UnifiedGroupLinks -Identity $newPnpTeam.GroupId -LinkType Owner -Links $($365creds.UserName) -Confirm:$false
        Remove-DistributionGroupMember -Identity $new365Group.CustomAttribute2 -Member $($365creds.UserName) -Confirm:$false -BypassSecurityGroupManagerCheck:$true
        Remove-PnPSiteCollectionAdmin -Owners $($365creds.UserName)
        }
    if($addExecutingUserAsTemporaryMember){
        Remove-UnifiedGroupLinks -Identity $newPnpTeam.GroupId -LinkType Member -Links $($365creds.UserName) -Confirm:$false
        }
    }