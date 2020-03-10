$365creds = set-MsolCredentials
connect-to365 -credential $365creds

$displayName = "Sustainable Chemistry Team (GBR)"
$areDataManagersLineManagers = $true
$managedBy = "365"
#$memberOf = ??
$hideFromGal = $false
$blockExternalMail = $true
$accessType = "Private"
$autoSubscribe = $true
$groupClassification = "Internal"
$alsoCreateTeam = $false
$horriblyUnformattedStringOfManagers = "groupbot@anthesisgroup.com, kevin.maitland@anthesisgroup.com"
$horriblyUnformattedStringOfMembers = "andrew.noone@anthesisgroup.com
ben.tuxworth@anthesisgroup.com
sarah.gilby@anthesisgroup.com
alex.mckay@anthesisgroup.com
"
    

$teamBotDetails = import-encryptedCsv -pathToEncryptedCsv "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\teambotdetails.txt"
$tokenResponse = get-graphTokenResponse -aadAppCreds $teamBotDetails

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
Connect-PnPOnline -AccessToken $tokenResponse.access_token
$newPnpTeam = Get-PnPUnifiedGroup -Identity $newTeam.ExternalDirectoryObjectId

#Aggrivatingly, you can't manipulate Pages with Graph yet, and Add-PnpFile doesn;t support AccessTokens, so we need to go old-school:
if($addExecutingUserAsTemporaryOwner){
    $addExecutingUserAsTemporarySiteCollectionAdmin = test-isUserSiteCollectionAdmin -pnpUnifiedGroupObject $newPnpTeam -accessToken $tokenResponse.access_token -pnpCreds $365creds -addPermissionsIfMissing $true
    }
copy-spoPage -sourceUrl "https://anthesisllc.sharepoint.com/sites/Resources-IT/SitePages/Candiate-Template-for-Team-Site-Landing-Page.aspx" -destinationSite $newPnpTeam.SiteUrl -pnpCreds $365creds -overwriteDestinationFile $true -renameFileAs "LandingPage.aspx" -Verbose | Out-Null
test-pnpConnectionMatchesResource -resourceUrl $newPnpTeam.SiteUrl -pnpCreds $pnpCreds -connectIfDifferent $true | Out-Null
if((test-pnpConnectionMatchesResource -resourceUrl $newPnpTeam.SiteUrl) -eq $true){
    Write-Verbose "Setting Homepage"
    Set-PnPHomePage  -RootFolderRelativeUrl "SitePages/LandingPage.aspx" | Out-Null
    }

Add-PnPHubSiteAssociation -Site $newPnpTeam.SiteUrl -HubSite "https://anthesisllc.sharepoint.com/sites/TeamHub" | Out-Null

Write-Verbose "Opening in browser - no further steps, it's just to eyeball the Site and check it's worked."
start-Process $newPnpTeam.SiteUrl

if($addExecutingUserAsTemporaryOwner){
    Remove-UnifiedGroupLinks -Identity $newPnpTeam.GroupId -LinkType Owner -Links $((Get-PnPConnection).PSCredential.UserName) -Confirm:$false
    Remove-DistributionGroupMember -Identity $new365Group.CustomAttribute2 -Member $((Get-PnPConnection).PSCredential.UserName) -Confirm:$false -BypassSecurityGroupManagerCheck:$true
    }
if($addExecutingUserAsTemporaryMember){
    Remove-UnifiedGroupLinks -Identity $newPnpTeam.GroupId -LinkType Member -Links $((Get-PnPConnection).PSCredential.UserName) -Confirm:$false
    }

Write-Verbose "set-standardSitePermissions [$($newTeam.DisplayName)]"
set-standardSitePermissions -unifiedGroupObject $newTeam -tokenResponse $tokenResponse -pnpCreds $365creds -Verbose


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
        Send-MailMessage  -BodyAsHtml $body -Subject "[$($newTeam.DisplayName)] Team Site created" -to $thisManager -bcc $((Get-PnPConnection).PSCredential.UserName) -from "TeamSiteRobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Encoding UTF8
        Write-Verbose "E-mail sent"
        }
    catch{$_}
    }