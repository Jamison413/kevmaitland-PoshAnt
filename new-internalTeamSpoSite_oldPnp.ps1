﻿$365creds = set-MsolCredentials
connect-to365 -credential $365creds

#email out
$smtpBotDetails = get-graphAppClientCredentials -appName SmtpBot
$tokenResponseSmtp = get-graphTokenResponse -aadAppCreds $smtpBotDetails



$requests = @()
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/sites/TeamHub" -Credentials $365creds
$requests += Get-PnPListItem -List "Internal Team Site Requests" -Query "<View><Query><Where><Eq><FieldRef Name='Status'/><Value Type='String'>Awaiting creation</Value></Eq></Where></Query></View>"
if($requests){[array]$selectedRequests = $requests | select {$_.FieldValues.Title},{$_.FieldValues.Site_x0020_Type},{$_.FieldValues.DataManager.LookupValue},{$_.FieldValues.Members.LookupValue -join ", "},{$_.FieldValues.GUID.Guid} | Out-GridView -PassThru -Title "Highlight any requests to process and click OK"}


$areDataManagersLineManagers = $false
$managedBy = "365"
#$memberOf = ??
$hideFromGal = $false
$blockExternalMail = $true
$accessType = "Private"
$autoSubscribe = $true
$groupClassification = "Internal"
$alsoCreateTeam = $true

foreach($request in $selectedRequests){
    $request = $requests | ? {$_.FieldValues.GUID.Guid -eq $request.'$_.FieldValues.GUID.Guid'}
    if($request.FieldValues.Site_x0020_Type -eq "Functional"){$displayName = $($request.FieldValues.Title)}
    else{$displayName = "$($request.FieldValues.Title) $($request.FieldValues.Site_x0020_Type)"}

    $teamBotDetails = get-graphAppClientCredentials -appName TeamsBot
    $tokenResponse = get-graphTokenResponse -aadAppCreds $teamBotDetails

    #region Get the Managers and Members in the right formats
    $managers = @()
    $originalManagers = convertTo-arrayOfEmailAddresses $request.FieldValues.DataManager.Email | sort | select -Unique
    $managers = $originalManagers #So we can e-mail the right people at the end.
    $members = @()
    $members += convertTo-arrayOfEmailAddresses $request.FieldValues.Members.Email | sort | select -Unique
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
    if($managers -notcontains ((Get-PnPConnection).PSCredential.UserName)){
        $addExecutingUserAsTemporaryOwner = $true
        [array]$managers += ((Get-PnPConnection).PSCredential.UserName)
        }
    if($members -notcontains ((Get-PnPConnection).PSCredential.UserName)){
        $addExecutingUserAsTemporaryMember = $true
        [array]$members += ((Get-PnPConnection).PSCredential.UserName)
        }

    if($managedBy -eq "AAD"){$managers = "groupbot@anthesisgroup.com"} #Override the ownership of any aggregated / Parent Functional Teams as these are automated separately

    #endregion


    $newTeam = new-365Group -displayName $displayName -managerUpns $managers -teamMemberUpns $members -memberOf $memberOf -hideFromGal $hideFromGal -blockExternalMail $blockExternalMail -accessType $accessType -autoSubscribe $autoSubscribe -additionalEmailAddresses $additionalEmailAddresses -groupClassification $groupClassification -ownersAreRealManagers $areDataManagersLineManagers -membershipmanagedBy $managedBy -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference -tokenResponse $tokenResponse -alsoCreateTeam $alsoCreateTeam -pnpCreds $365creds
    Write-Verbose "Getting PnPUnifiedGroup [$displayName] - this is a faster way to get the SharePoint URL than using the UnifiedGroup object"
    #Connect-PnPOnline -AccessToken $tokenResponse.access_token
    Connect-PnPOnline -ClientId $teamBotDetails.ClientID -ClientSecret $teamBotDetails.Secret -Url "https://anthesisllc.sharepoint.com" -RetryCount 2 -ReturnConnection
    $newPnpTeam = Get-PnPUnifiedGroup -Identity $newTeam.id

    #Aggrivatingly, you can't manipulate Pages with Graph yet, and Add-PnpFile doesn;t support AccessTokens, so we need to go old-school:
    if($addExecutingUserAsTemporaryOwner){
        Connect-PnPOnline -ClientId $teamBotDetails.ClientID -ClientSecret $teamBotDetails.Secret -Url "https://anthesisllc.sharepoint.com"
        $executingUserAlreadySiteCollectionAdmin = test-isUserSiteCollectionAdmin -pnpUnifiedGroupObject $newPnpTeam -pnpAppCreds $teamBotDetails -pnpCreds $365creds -addPermissionsIfMissing $true
        }
    copy-spoPage -sourceUrl "https://anthesisllc.sharepoint.com/sites/Resources-IT/SitePages/Candiate-Template-for-Team-Site-Landing-Page.aspx" -destinationSite $newPnpTeam.SiteUrl -pnpCreds $365creds -overwriteDestinationFile $true -renameFileAs "LandingPage.aspx" -Verbose | Out-Null
    test-pnpConnectionMatchesResource -resourceUrl $newPnpTeam.SiteUrl -pnpCreds $365creds -connectIfDifferent $true | Out-Null
    if((test-pnpConnectionMatchesResource -resourceUrl $newPnpTeam.SiteUrl) -eq $true){
        Write-Verbose "Setting Homepage"
        Set-PnPHomePage  -RootFolderRelativeUrl "SitePages/LandingPage.aspx" | Out-Null
        Write-Verbose "ReTitling Homepage"
        $newHomepage = Get-PnPListItem -List "SitePages" -Query "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='Text'>LandingPage.aspx</Value></Eq></Where></Query></View>"
        Set-PnPListItem -Values @{"Title"=$newPnpTeam.DisplayName} -List "SitePages" -Identity $newHomepage.Id

        Write-Verbose "Create, Share and Delete a folder in the Documents Library to enable the SharedWithUsers metadata"
        $docLibName = "Shared Documents"
        $dummyFolderName = "DummyShareToDelete"
        Write-Verbose "`tAdding Folder [$dummyFolderName] to [$docLibName]"
        Add-PnPFolder -Name $dummyFolderName -Folder $docLibName
        $dummyPnpFolderItem = Get-PnPFolderItem -FolderSiteRelativeUrl $docLibName -ItemType Folder -ItemName $dummyFolderName
        [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Sharing")  | Out-Null
        [System.Reflection.Assembly]::LoadWithPartialName("System.Collections") | Out-Null
        $roleAssignments = New-Object "System.Collections.Generic.List[Microsoft.SharePoint.Client.Sharing.UserRoleAssignment]"
        $roleAssignment = New-Object Microsoft.SharePoint.Client.Sharing.UserRoleAssignment
        $roleAssignment.UserId = $365creds.UserName
        $roleAssignment.Role = [Microsoft.SharePoint.Client.Sharing.Role]::Edit
        $roleAssignments.Add($roleAssignment)
        [Microsoft.SharePoint.Client.Sharing.DocumentSharingManager]::UpdateDocumentSharingInfo($dummyPnpFolderItem.Context,"https://anthesisllc.sharepoint.com"+$dummyPnpFolderItem.ServerRelativeUrl,$roleAssignments,$false,$true,$false,"",$false,$false)
        Write-Verbose "`tSharing Folder [$dummyFolderName] with [$($365creds.UserName)] via CSOM"
        $dummyPnpFolderItem.Context.ExecuteQuery() 
        Write-Verbose "`tRemoving Folder [$dummyFolderName]"
        Remove-PnPFolder -Name $dummyFolderName -Folder $docLibName -Force


        Write-Verbose "Setting default View in Documents Library"
        $thisDocLib = Get-PnPList -Identity $docLibName -Includes Fields
        $defaulDocLibPnpView = $thisDocLib | Get-PnPView | ? {$_.DefaultView -eq $true}
        $defaulDocLibPnpView | Set-PnPView -Fields "DocIcon","LinkFilename","Modified","Editor","Created","Author","FileSizeDisplay","SharedWithUsers"



        <#--
        $ctx = (Get-PnPConnection).Context
        Write-Verbose "Create, Share and Delete a folder in the Documents Library to enable the SharedWithUsers metadata"
        $dummyFolderName = "DummyShareToDelete"
        $defaultDocLib = Get-PnPList -Identity "Documents"
        Write-Verbose "`tAdding Folder [$dummyFolderName]"
        Add-PnPFolder -Name $dummyFolderName -Folder "Shared Documents"
        $dummyPnpFolderItem = Get-PnPFolderItem -FolderSiteRelativeUrl "Shared Documents" -ItemType Folder -ItemName $dummyFolderName
        #Set-PnPListItemPermission -List $defaultDocLib.Id -Identity $dummyPnpFolderItem.ListItemAllFields -AddRole Contribute -User $($pnpCreds.UserName) This sets the permissions, but doesn't trigger a SharedWith event
        #[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Sharing")  | Out-Null
        #[System.Reflection.Assembly]::LoadWithPartialName("System.Collections") | Out-Null
        [System.Reflection.Assembly]::LoadFile("C:\Program Files\WindowsPowerShell\Modules\SharePointPnPPowerShellOnline\3.17.2001.2\Microsoft.SharePoint.Client.dll")
        $roleAssignments = New-Object "System.Collections.Generic.List[Microsoft.SharePoint.Client.Sharing.UserRoleAssignment]"
        $roleAssignment = New-Object Microsoft.SharePoint.Client.Sharing.UserRoleAssignment
        $roleAssignment.UserId = $pnpCreds.UserName
        $roleAssignment.Role = [Microsoft.SharePoint.Client.Sharing.Role]::Edit
        $roleAssignments.Add($roleAssignment)
        [Microsoft.SharePoint.Client.Sharing.DocumentSharingManager]::UpdateDocumentSharingInfo($dummyPnpFolderItem.Context,"https://anthesisllc.sharepoint.com"+$dummyPnpFolderItem.ServerRelativeUrl,$roleAssignments,$false,$true,$false,"",$false,$false)
        [Microsoft.SharePoint.Client.Sharing.DocumentSharingManager]::UpdateDocumentSharingInfo($dummyPnpFolderItem.Context.CastTo($_,[Microsoft.SharePoint.Client.ClientRuntimeContext]),"https://anthesisllc.sharepoint.com"+$dummyPnpFolderItem.ServerRelativeUrl,$roleAssignments,$false,$true,$false,"",$false,$false)
        Write-Verbose "`tSharing Folder [$dummyFolderName] with [$($pnpCreds.UserName)] via CSOM"
        $dummyPnpFolderItem.Context.ExecuteQuery() #This errors, but still adds the SharedWithUsers column to the Site
        Remove-PnPFolder -Name $dummyFolderName -Folder "Shared Documents" -Force


        Write-Verbose "Setting default View in Documents Library"
        $defaulDocLibPnpView = $defaultDocLib | Get-PnPView | ? {$_.DefaultView -eq $true}
        $defaulDocLibPnpView | Set-PnPView -Fields "DocIcon","LinkFilename","Modified","Editor","Created","Author","FileSizeDisplay","SharedWithUsers" --#>
        }

    Add-PnPHubSiteAssociation -Site $newPnpTeam.SiteUrl -HubSite "https://anthesisllc.sharepoint.com/sites/TeamHub" | Out-Null

    Write-Verbose "Opening in browser - no further steps, it's just to eyeball the Site and check it's worked."
    start-Process $newPnpTeam.SiteUrl


    Write-Host -f DarkYellow "`tset-standardSitePermissions [$($newTeam.DisplayName)]"
    try{
        Connect-PnPOnline -ClientId $teamBotDetails.ClientID -ClientSecret $teamBotDetails.Secret -Url "https://anthesisllc.sharepoint.com"
        set-standardSitePermissions -tokenResponse $tokenResponse -graphGroupExtended $newTeam -pnpAppCreds $teamBotDetails -pnpCreds $365creds -Verbose:$VerbosePreference -suppressEmailNotifications -ErrorAction Continue
        }
    catch{$_}


    if($addExecutingUserAsTemporaryOwner){
        Remove-UnifiedGroupLinks -Identity $newPnpTeam.GroupId -LinkType Owner -Links $((Get-PnPConnection).PSCredential.UserName) -Confirm:$false
        Remove-DistributionGroupMember -Identity $newTeam.anthesisgroup_UGSync.dataManagerGroupId -Member $((Get-PnPConnection).PSCredential.UserName) -Confirm:$false -BypassSecurityGroupManagerCheck:$true
        }
    if($addExecutingUserAsTemporaryMember){
        Remove-UnifiedGroupLinks -Identity $newPnpTeam.GroupId -LinkType Member -Links $((Get-PnPConnection).PSCredential.UserName) -Confirm:$false
        }


    Write-Verbose "Updating Team Request: Status = [Created]"
    Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/sites/TeamHub" -Credentials $365creds
    $dummy = Set-PnPListItem -List "34cad35a-4710-4fc9-bd53-ec35ae54574f" -Identity $request.Id -Values @{Status="Created"} #"Internal Team Site Requests" List 

    Write-Verbose "Preparing e-mail(s)"
    $originalManagers | % {
        $thisManager = $_
        $thisManagerFirstName = guess-nameFromString -ambiguousString $thisManager
        if(![string]::IsNullOrWhiteSpace($thisManagerFirstName)){$thisManagerFirstName = ($thisManagerFirstName.Split(" ")[0])}
        try{
            $body = "<HTML><BODY><p>Hi $thisManagerFirstName,</p>
                <p>Your new <a href=`"$($newPnpTeam.siteUrl)`">[$($newTeam.DisplayName)] Team Site</a> is available for you now. You are probably already 
                familiar with how these Sites work, but we have <a href=`"https://anthesisllc.sharepoint.com/sites/Resources-IT/SitePages/SharePoint-Training-Guides.aspx#data-managers-guides`">
                a good selection of guides for Data Mangers</a> available on the IT Resources Site, and a few of the most popular ones are below if
                you want to do anything fancier that simply sharing files:</p>

                <UL><LI><a href=`"https://anthesisllc.sharepoint.com/sites/Resources-IT/_layouts/15/DocIdRedir.aspx?ID=HXX7CE52TSD2-1759992947-62`">Changing
                the logo for your Site</a></LI>

                <LI><a href=`"https://anthesisllc.sharepoint.com/sites/Resources-IT/_layouts/15/DocIdRedir.aspx?ID=HXX7CE52TSD2-1759992947-196`">Creating/editing
                pages in SharePoint</a></LI>

                <LI><a href=`"https://anthesisllc.sharepoint.com/sites/Resources-IT/_layouts/15/DocIdRedir.aspx?ID=HXX7CE52TSD2-1759992947-105`">Creating
                links in SharePoint</a></LI>

                <LI><a href=`"https://anthesisllc.sharepoint.com/sites/Resources-IT/_layouts/15/DocIdRedir.aspx?ID=HXX7CE52TSD2-1759992947-42`">Adding
                icons to your link</a></LI></UL>

                <p>You and all the new members of your team will get another e-mail from 365 shortly telling you that the new team has been created, and you can find your way back to the file storage area in SharePoint either <a href=`"https://anthesisllc.sharepoint.com/sites/Resources-IT/_layouts/15/DocIdRedir.aspx?ID=HXX7CE52TSD2-1759992947-4`">via Outlook</a>, by <a href=`"https://anthesisllc.sharepoint.com/sites/Resources-IT/_layouts/15/DocIdRedir.aspx?ID=HXX7CE52TSD2-1759992947-32`">bookmarking the Site in Chrome</a>, or <a href=`"https://anthesisllc.sharepoint.com/sites/Resources-IT/_layouts/15/DocIdRedir.aspx?ID=HXX7CE52TSD2-1759992947-209`"><i>ridiculously</i> easily by searching in Chrome</a>.</p>

                <p>You won't be able to able to share data in this Site with external users or guests (if you want to do this, you'll need to take a look 
                at <a href=`"https://anthesisllc.sharepoint.com/sites/external/SitePages/External-Sharing-Sites.aspx`">External Sharing Sites</a>).</p>

                <p>Love,</p>

                <p>The Team Site Robot</p>
                </BODY></HTML>"
            #Send-MailMessage  -BodyAsHtml $body -Subject "[$($newTeam.DisplayName)] Team Site created" -to $thisManager -bcc $((Get-PnPConnection).PSCredential.UserName) -from "TeamSiteRobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Encoding UTF8
            send-graphMailMessage -tokenResponse $tokenResponseSmtp -fromUpn teamsiterobot@anthesisgroup.com -toAddresses $thisManager -subject "[$($newTeam.DisplayName)] Team Site created" -bodyHtml $body -bccAddresses $($365creds.UserName)

            Write-Verbose "E-mail sent"
            }
        catch{$_}
        }

    }