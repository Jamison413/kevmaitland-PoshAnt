﻿$365creds = set-MsolCredentials
connect-to365 -credential $365creds

$allAdminUGs= get-unifiedgroup -Filter "DisplayName -like 'Admin*'"  

$teamBotDetails = $(get-graphAppClientCredentials -appName TeamsBot)
$tokenResponseTeamsBot = get-graphTokenResponse -aadAppCreds $teamBotDetails -grant_type client_credentials 

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



foreach($team in @("Energy Practice Team (GBR)")){}
    $tokenResponseTeamsBot = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponseTeamsBot -renewTokenExpiringInSeconds 300 -aadAppCreds $teamBotDetails

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
    $horriblyUnformattedStringOfManagers = "andrew.ost@anthesisgroup.com"
    $horriblyUnformattedStringOfMembers = ""<# , Chaminda.Ranaweera@anthesisgroup.com  , Olga.Harrington@anthesisgroup.com , Hannah.Southan@anthesisgroup.com "#>
    
    

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

    #Run 188 - 193 twice so that it can see the Shared Mailbox is a member of the team or line 196 will fail :(
    $newGroup = new-365Group -displayName $displayName -managerUpns $managers -teamMemberUpns $members -memberOf $memberOf -hideFromGal $hideFromGal -blockExternalMail $blockExternalMail -accessType $accessType -autoSubscribe $autoSubscribe -additionalEmailAddresses $additionalEmailAddresses -groupClassification $groupClassification -ownersAreRealManagers $areDataManagersLineManagers -membershipmanagedBy $managedBy -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference -tokenResponse $tokenResponseTeamsBot -alsoCreateTeam $alsoCreateTeam -pnpCreds $365creds
    Write-Verbose "Getting associated PnP and Graph objects for [$($newGroup.DisplayName)] - this is a faster way to do stuff than using the UnifiedGroup object"
    Connect-PnPOnline -ClientId $teamBotDetails.ClientID -ClientSecret $teamBotDetails.Secret -Url "https://anthesisllc.sharepoint.com" #-RetryCount 2 -ReturnConnection
    $newPnpTeam = Get-PnPUnifiedGroup -Identity $newGroup.id
    $newGraphGroup = get-graphGroups -tokenResponse $tokenResponseTeamsBot -filterUpn $newGroup.mail -selectAllProperties
    $newGraphGroupDrive = get-graphDrives -tokenResponse $tokenResponseTeamsBot -groupGraphId $newGroup.id -returnOnlyDefaultDocumentsLibrary

    #Remove GroupBot if required
    $newTeamOwners = get-graphUsersFromGroup -tokenResponse $tokenResponseTeamsBot -groupUpn $newGroup.mail -returnOnlyLicensedUsers -memberType Owners
    if($newTeamOwners.userPrincipalName -contains "groupbot@anthesisgroup.com"){
        if($newTeamOwners.userPrincipalName -notcontains $365creds.UserName){
        $MyGraphUser = get-graphUsers -tokenResponse $tokenResponseTeamsBot -filterUpns $365creds.UserName
            add-graphUsersToGroup -tokenResponse $tokenResponseTeamsBot -graphGroupId $newGraphGroup.id -memberType Owners -graphUserIds $MyGraphUser.id
            }
        remove-graphUsersFromGroup -tokenResponse $tokenResponseTeamsBot -graphGroupId $newGraphGroup.id -memberType Owners -graphUserIds "00aa81e4-2e8f-4170-bc24-843b917fd7cf" #This works faster with Ids, so I've hard-coded GroupBot's Id
        #remove-graphUsersFromGroup -tokenResponse $tokenResponseTeamsBot -graphGroupId $newGraphGroup.id -memberType Members -graphUserIds "00aa81e4-2e8f-4170-bc24-843b917fd7cf"
        }
    $newGraphTeam = new-graphTeam -tokenResponse $tokenResponseTeamsBot -groupId $newGraphGroup.id -allowMemberCreateUpdateChannels $true -allowMemberDeleteChannels $false #Create the Team if it doesn't already exist

    #region Do as much as we can in Graph because it's much quicker
    ########################################
    #Create, share and delete a dummy folder in this Site to trigger the SharedWith Site Column
    $dummyFolder = add-graphArrayOfFoldersToDrive -graphDriveId $newGraphGroupDrive.id -foldersAndSubfoldersArray "DummyFolder" -tokenResponse $tokenResponseTeamsBot -conflictResolution Replace
    grant-graphSharing -tokenResponse $tokenResponseTeamsBot -driveId $newGraphGroupDrive.id -itemId $dummyFolder.id -sharingRecipientsUpns @($365creds.UserName) -requireSignIn $true -sendInvitation $false -role Write -Verbose
    delete-graphDriveItem -tokenResponse $tokenResponseTeamsBot -graphDriveId $newGraphGroupDrive.id -graphDriveItemId $dummyFolder.id -eTag $dummyFolder.eTag 

    #Create corresponding Regional Folder in associated regional Administration Team site
    <#$currentRegion = get-3lettersInBrackets -stringMaybeContaining3LettersInBrackets $newGroup.DisplayName
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
    if([string]::IsNullOrWhiteSpace($associatedRegionalAdminTeam)){
        Write-Error "Could not identify regional Administration Team. Cannot proceed with configuring the rest of the Site & Team." ;break
        }

    $associatedRegionalAdminDrive = get-graphDrives -tokenResponse $tokenResponseTeamsBot -teamUpn $associatedRegionalAdminTeam.PrimarySmtpAddress -returnOnlyDefaultDocumentsLibrary
    $newRegionalFolder = add-graphArrayOfFoldersToDrive -graphDriveId $associatedRegionalAdminDrive.id -foldersAndSubfoldersArray @($newGroup.DisplayName) -tokenResponse $tokenResponseTeamsBot -conflictResolution Fail

    #Set Edit permissions for this Team on new folder
    grant-graphSharing -tokenResponse $tokenResponseTeamsBot -driveId $associatedRegionalAdminDrive.id -itemId $newRegionalFolder.id -sharingRecipientsUpns @($newGroup.mail) -requireSignIn $true -sendInvitation $false -role Write
#>
    #Create text file explaining how this works / links to new folder from this Site
    $textFileTemplateContent = "This Hive Team has the file functionality disabled. You can still store data in the Practice Community though, or you can store data in your personal OneDrive area, share it with this Team, and post a link in a channel."

    $newTextFile = invoke-graphPut -tokenResponse $tokenResponseTeamsBot -graphQuery "/drives/$($newGraphGroupDrive.id)/items/root:/This Hive Team ($($newGroup.DisplayName)) has the file functionality disabled.txt:/content" -binaryFileStream $textFileTemplateContent
    $anotherNewTextFile = invoke-graphPut -tokenResponse $tokenResponseTeamsBot -graphQuery "/drives/$($newGraphGroupDrive.id)/items/root:/General/This Hive Team ($($newGroup.DisplayName)) has the file functionality disabled.txt:/content" -binaryFileStream $textFileTemplateContent
    <#$newHyperlinkContent = `
    "[InternetShortcut]
    URL=$($associatedRegionalAdminDrive.webUrl+"/"+[uri]::EscapeDataString($newGroup.DisplayName))
    "
    $newHyperlink = invoke-graphPut -tokenResponse $tokenResponseTeamsBot -graphQuery "/drives/$($newGraphGroupDrive.id)/items/root:/Link to storage manged by $($associatedRegionalAdminTeam.DisplayName).url:/content" -binaryFileStream $newHyperlinkContent#>
    #$anotherNewHyperlink = invoke-graphPut -tokenResponse $tokenResponseTeamsBot -graphQuery "/drives/$($newGraphGroupDrive.id)/items/root:/General/Link to storage manged by $($associatedRegionalAdminTeam.DisplayName).url:/content" -binaryFileStream $newHyperlinkContent #This doesn;t work so well in Teams, but we'll keep the first link as it's only really visible from SharePoint

    #Create new Tab in Teams linking to this location
    <#$newGraphTeamGeneralChannel = invoke-graphGet -tokenResponse $tokenResponseTeamsBot -graphQuery "/teams/$($newGraphGroup.id)/channels"
    $newGraphTeamGeneralChannelTabs = invoke-graphGet -tokenResponse $tokenResponseTeamsBot -graphQuery "/teams/$($newGraphGroup.id)/channels/$($newGraphTeamGeneralChannel.id)/tabs"
    $newGraphTeamGeneralChannelTabs | ? {$_.displayName -eq "Managed Files"} | % {
        invoke-graphDelete -tokenResponse $tokenResponseTeamsBot -graphQuery "/teams/$($newGraphGroup.id)/channels/$($newGraphTeamGeneralChannel.id)/tabs/$($_.id)" 
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
    $newGraphTeamGeneralChannelManagedFilesTab = invoke-graphPost -tokenResponse $tokenResponseTeamsBot -graphQuery "/teams/$($newGraphGroup.id)/channels/$($newGraphTeamGeneralChannel.id)/tabs" -graphBodyHashtable $tabBody#>


    #endregion

    #region We still have to do some stuff in PnP because it's not supported in Graph yet
    if($addExecutingUserAsTemporaryOwner){ #copy-spoPage requires Site Collection Admin rights
        $userWasAlreadySiteAdmin = test-isUserSiteCollectionAdmin -pnpUnifiedGroupObject $newPnpTeam -pnpAppCreds $teamBotDetails -pnpCreds $365creds -addPermissionsIfMissing $true -Verbose
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
    set-standardSitePermissions -tokenResponse $tokenResponseTeamsBot -pnpAppCreds $teamBotDetails -graphGroupExtended $newGraphGroup -pnpCreds $365creds -Verbose


    if($addExecutingUserAsTemporaryOwner){
        test-pnpConnectionMatchesResource -resourceUrl $newPnpTeam.SiteUrl -pnpCreds $365creds -connectIfDifferent $true | Out-Null
        Remove-UnifiedGroupLinks -Identity $newPnpTeam.GroupId -LinkType Owner -Links $($365creds.UserName) -Confirm:$false
        Remove-DistributionGroupMember -Identity $new365Group.CustomAttribute2 -Member $($365creds.UserName) -Confirm:$false -BypassSecurityGroupManagerCheck:$true
        Remove-PnPSiteCollectionAdmin -Owners $($365creds.UserName)
        }
    if($addExecutingUserAsTemporaryMember){
        Remove-UnifiedGroupLinks -Identity $newPnpTeam.GroupId -LinkType Member -Links $($365creds.UserName) -Confirm:$false
        }
  