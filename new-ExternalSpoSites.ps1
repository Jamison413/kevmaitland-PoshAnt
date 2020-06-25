$365creds = set-MsolCredentials
connect-ToExo -credential $365creds

$teamBotDetails = import-encryptedCsv -pathToEncryptedCsv "$env:USERPROFILE\Desktop\teambotdetails.txt"
$tokenResponse = get-graphTokenResponse -aadAppCreds $teamBotDetails

$allClientSiteDocLibs = get-graphDrives -tokenResponse $tokenResponse -siteUrl "https://anthesisllc.sharepoint.com/clients" #-filterDriveName $($fullRequest.FieldValues.ClientName.Label) #$filters aren't currently supported on this endpoint :'(
$allSupplierSiteDocLibs = get-graphDrives -tokenResponse $tokenResponse -siteUrl "https://anthesisllc.sharepoint.com/subs" #-filterDriveName $($fullRequest.FieldValues.ClientName.Label) #$filters aren't currently supported on this endpoint :'(

#$toDo = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/sites/anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,faed84bc-70be-4e35-bfbf-cdab31aeeb99/Lists/06365ce6-07a5-41e9-b0aa-a90c9f1edc3f/items?expand=fields(select=ID,ClientName,Title,Site_x0020_Owners,Site_x0020_Members,Status,GuID)&filter=fields/Status eq 'Awaiting creation'"
#$toDo += invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/sites/anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,9fb8ecd6-c87d-485d-a488-26fd18c62303/Lists/0c68ca6f-06fe-449b-8cf1-c0dbe7fddd5c/items?expand=fields(select=ID,Subcontractor_x002f_Supplier_x00,ClientNameTitle,Site_x0020_Owners,Site_x0020_Members,Status,GuID)&filter=fields/Status eq 'Awaiting creation'"
#if($toDo){[array]$selectedRequests = $toDo | select {$_.Fields.Title},{$_.Fields.ClientName.Label},{$_.Fields.Site_x0020_Admin.LookupValue},{$_.Fields.Site_x0020_Owners.LookupValue -join ", "},{$_.Fields.Site_x0020_Members.LookupValue -join ", "},{$_.Fields.GUID} | Out-GridView -PassThru -Title "Highlight any requests to process and click OK"}

#clients
$requests = @()
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/clients" -Credentials $365creds
$requests += Get-PnPListItem -List "External Client Site Requests" -Query "<View><Query><Where><Eq><FieldRef Name='Status'/><Value Type='String'>Awaiting creation</Value></Eq></Where></Query></View>"
if($requests){[array]$selectedRequests = $requests | select {$_.FieldValues.Title},{$_.FieldValues.ClientName.Label},{$_.FieldValues.Site_x0020_Admin.LookupValue},{$_.FieldValues.Site_x0020_Owners.LookupValue -join ", "},{$_.FieldValues.Site_x0020_Members.LookupValue -join ", "},{$_.FieldValues.GUID.Guid} | Out-GridView -PassThru -Title "Highlight any requests to process and click OK"}
#subs
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/subs" -Credentials $365creds
$requests += Get-PnPListItem -List "External Subcontractor Site Requests" -Query "<View><Query><Where><Eq><FieldRef Name='Status'/><Value Type='String'>Awaiting creation</Value></Eq></Where></Query></View>"
if($requests){[array]$selectedRequests = $requests | select {$_.FieldValues.Title},{$_.FieldValues.Subcontractor_x002f_Supplier_x00.Label},{$_.FieldValues.Site_x0020_Admin.LookupValue},{$_.FieldValues.Site_x0020_Owners.LookupValue -join ", "},{$_.FieldValues.Site_x0020_Members.LookupValue -join ", "},{$_.FieldValues.GUID.Guid} | Out-GridView -PassThru -Title "Highlight any requests to process and click OK"}

    $hideFromGal = $false
    $blockExternalMail = $false
    $accessType = "Private"
    $autoSubscribe = $true
    $groupClassification = "External"



foreach ($currentRequest in $selectedRequests){
    $tokenResponse = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponse -renewTokenExpiringInSeconds 300 -aadAppCreds $teamBotDetails
    $fullRequest = $requests | ? {$_.FieldValues.GUID.Guid -eq $currentRequest.'$_.FieldValues.GUID.Guid'}
    $alsoCreateTeam = $($fullRequest.FieldValues.UpgradeToTeam)
    write-host -ForegroundColor Yellow "Creating External Sharing Site [$($fullRequest.FieldValues.Title)] for [$($fullRequest.FieldValues.ClientName.Label)]"
    $managers = convertTo-arrayOfEmailAddresses ($fullRequest.FieldValues.Site_x0020_Owners.Email +","+ $fullRequest.FieldValues.Site_x0020_Admin.Email) | sort | select -Unique
    $members = convertTo-arrayOfEmailAddresses ($managers + "," + $fullRequest.FieldValues.Site_x0020_Members.Email) | sort | select -Unique
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

    Write-Host -f DarkYellow "`tData Managers:`t$($managers -join ", ")"
    Write-Host -f DarkYellow "`tMembers:`t`t$($members -join ", ")"
    try{
        Write-Host -f DarkYellow "`tCreating Groups"
        $new365Group = new-365Group -displayName $("External - $($fullRequest.FieldValues.Title)".Trim(" ")) -managerUpns $managers -teamMemberUpns $members -memberOf $null -hideFromGal $hideFromGal -blockExternalMail $blockExternalMail -accessType Private -autoSubscribe $autoSubscribe -groupClassification $groupClassification -membershipManagedBy 365 -tokenResponse $tokenResponse -pnpCreds $365creds -ownersAreRealManagers $false -alsoCreateTeam $alsoCreateTeam -Verbose
        
        Write-Verbose "Getting PnPUnifiedGroup [$($new365Group.displayName)] - this is a faster way to get the SharePoint URL than using the UnifiedGroup object"
        Connect-PnPOnline -AccessToken $tokenResponse.access_token
        $newPnpTeam = Get-PnPUnifiedGroup -Identity $new365Group.id

        Write-Verbose "Adding Navigation to External Hub"
        Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/sites/external" -Credentials $365creds
        Add-PnPNavigationNode -Location QuickLaunch -Title $($fullRequest.FieldValues.Title) -Url $newPnpTeam.SiteUrl -First -External -Parent 2252 #2252 is the "Modern External Sites" NavNode

        #Add a link from the Clients/Suppliers Site folder to this Site for improved eyeball search
        Write-Host -f DarkYellow "`tGetting Client DocLibs - this might take a while!"
        Write-Verbose "Getting Client DocLibs - this might take a while!"
        switch($fullRequest.FieldValues.FileDirRef.Split("/")[1]){
            "clients" {
                $clientOrSupplierSiteDocLib = $allClientSiteDocLibs | ? {$_.Name -eq $($fullRequest.FieldValues.ClientName.Label)}
                }
            "subs"    {
                $clientOrSupplierSiteDocLib = $allSupplierSiteDocLibs | ? {$_.Name -eq $($fullRequest.FieldValues.ClientName.Label)}
                }
            }
        if($clientOrSupplierSiteDocLib){ #The Spaniards create their own Clients Managed MetaData records, so there might not be Clients DocLibs for every instance
            $newHyperlinkContent = `
            "[InternetShortcut]
            URL=$($newPnpTeam.SiteUrl)
            "
            Write-Host -f DarkYellow "`tCreating link from Clients/Suppliers Site to new External Sharing Site"
            $newHyperlink = invoke-graphPut -tokenResponse $tokenResponse -graphQuery "/drives/$($clientOrSupplierSiteDocLib.id)/items/root:/Link to $($new365Group.DisplayName).url:/content" -binaryFileStream $newHyperlinkContent
            }
        else{Write-Verbose "No Client/Supplier DocLib found."}
        

        #Add a Website tab in the General channel linking back to the Client Site 
        if($alsoCreateTeam){
            Write-Host -f DarkYellow "`tCreating Website Tab in General channel to  Clients/Suppliers Site "
            add-graphWebsiteTabToChannel -tokenResponse $tokenResponse -teamId $new365Group.id -channelName "General" -tabName "$($fullRequest.FieldValues.ClientName.Label) Client Data" -tabDestinationUrl $clientOrSupplierSiteDocLib.webUrl -Verbose
            }

        #Aggrivatingly, you can't manipulate Pages with Graph yet, and Add-PnpFile doesn;t support AccessTokens, so we need to go old-school:
        if($addExecutingUserAsTemporaryOwner){
            $executingUserAlreadySiteCollectionAdmin = test-isUserSiteCollectionAdmin -pnpUnifiedGroupObject $newPnpTeam -accessToken $tokenResponse.access_token -pnpCreds $365creds -addPermissionsIfMissing $true
            }
        Write-Host -f DarkYellow "`tCopying new homepage"
        copy-spoPage -sourceUrl "https://anthesisllc.sharepoint.com/sites/Resources-IT/SitePages/External-Site-Template-Candidate.aspx" -destinationSite $newPnpTeam.SiteUrl -pnpCreds $365creds -overwriteDestinationFile $true -renameFileAs "LandingPage.aspx" | Out-Null
        test-pnpConnectionMatchesResource -resourceUrl $newPnpTeam.SiteUrl -pnpCreds $pnpCreds -connectIfDifferent $true | Out-Null
        if((test-pnpConnectionMatchesResource -resourceUrl $newPnpTeam.SiteUrl) -eq $true){
            Write-Host -f DarkYellow "`tSetting Homepage"
            Set-PnPHomePage  -RootFolderRelativeUrl "SitePages/LandingPage.aspx" | Out-Null
            Write-Host -f DarkYellow "`tDisabling Comments & Retitling Page"
            Set-PnPClientSidePage -Identity "LandingPage.aspx" -Title $newPnpTeam.DisplayName -CommentsEnabled:$false | Out-Null

            Write-Host -f DarkYellow "`tSetting default View in Documents Library"
            $thisDocLib = Get-PnPList -Identity "Documents" -Includes Fields
            $defaulDocLibPnpView = $thisDocLib | Get-PnPView | ? {$_.DefaultView -eq $true}
            $defaulDocLibPnpView | Set-PnPView -Fields "DocIcon","LinkFilename","Modified","Editor","Created","Author","FileSizeDisplay","SharedWithUsers"
            }

        Write-Host -f DarkYellow "`tSetting Hub Site association"
        Add-PnPHubSiteAssociation -Site $newPnpTeam.SiteUrl -HubSite "https://anthesisllc.sharepoint.com/sites/ExternalHub" | Out-Null
        Write-Verbose "Opening in browser"
        Set-Clipboard -Value $fullRequest.FieldValues.Site_x0020_Admin.Email
        Write-Host -ForegroundColor Yellow "Site Admin is : [$($fullRequest.FieldValues.Site_x0020_Admin.LookupValue)] (copied to clipboard for you to paste into Site Admin webpart)"

        start-Process $newPnpTeam.SiteUrl

        Write-Host -f DarkYellow "`tset-standardSitePermissions [$($new365Group.DisplayName)]"
        set-standardSitePermissions -tokenResponse $tokenResponse -graphGroupExtended $new365Group -pnpCreds $365creds -Verbose:$VerbosePreference -suppressEmailNotifications

        if($addExecutingUserAsTemporaryOwner){
            Write-Host -f DarkYellow "`tRemoving temporary Admin role for [$($365creds.UserName)] from [$($new365Group.DisplayName)]"
            remove-graphUsersFromGroup -tokenResponse $tokenResponse -graphGroupId $new365Group.id -memberType Owners -graphUserUpns $365creds.UserName
            #Remove-UnifiedGroupLinks -Identity $newPnpTeam.GroupId -LinkType Owner -Links $((Get-PnPConnection).PSCredential.UserName) -Confirm:$false
            Remove-DistributionGroupMember -Identity $new365Group.anthesisgroup_UGSync.dataManagerGroupId -Member $365creds.UserName -Confirm:$false -BypassSecurityGroupManagerCheck:$true
            }
        if($addExecutingUserAsTemporaryMember){
            Write-Host -f DarkYellow "`tRemoving temporary Membership role for [$($365creds.UserName)] from [$($new365Group.DisplayName)]"
            remove-graphUsersFromGroup -tokenResponse $tokenResponse -graphGroupId $new365Group.id -memberType Members -graphUserUpns $365creds.UserName
            #Remove-UnifiedGroupLinks -Identity $newPnpTeam.GroupId -LinkType Member -Links $((Get-PnPConnection).PSCredential.UserName) -Confirm:$false
            }

        switch($fullRequest.FieldValues.FileDirRef.Split("/")[1]){
            "clients" {
                Write-Host -f DarkYellow "`tUpdating Client Request: Status = [Created],Url=[$($newPnpTeam.SiteUrl)]"
                test-pnpConnectionMatchesResource -resourceUrl "https://anthesisllc.sharepoint.com/clients" -connectIfDifferent $true -pnpCreds $365creds | Out-Null
                $dummy = Set-PnPListItem -List "External Client Site Requests" -Identity $fullRequest.Id -Values @{Status="Created";URL=$newPnpTeam.SiteUrl}
                $externalParty = $fullRequest.FieldValues.ClientName.Label
                $externalPartyType = "client"
                Write-Host -f DarkYellow "`tGetting Managed MetaData term"
                $termGroup = $(Get-PnPTermGroup "Kimble") 
                $termSet = $(Get-PnPTermSet -TermGroup $termGroup -Identity "Clients")
                $managedMetaDataTerm = Get-PnPTerm -Identity $fullRequest.FieldValues.ClientName.Label -TermSet $termSet -TermGroup $termGroup -Includes CustomProperties
                Write-Verbose "`tRetrieved: [$($managedMetaDataTerm.Name)]"
                }
            "subs"    {
                Write-Verbose "Updating Supplier Request: Status = [Created]"
                test-pnpConnectionMatchesResource -resourceUrl "https://anthesisllc.sharepoint.com/subs" -connectIfDifferent $true -pnpCreds $365creds | Out-Null
                $dummy = Set-PnPListItem -List "0c68ca6f-06fe-449b-8cf1-c0dbe7fddd5c" -Identity $fullRequest.Id -Values @{Status="Created"} #"External Subcontractor Site Requests" List 
                $externalParty = $fullRequest.FieldValues.Subcontractor_x002f_Supplier_x00.Label
                $externalPartyType = "subcontractor"
                }
            default   {}
            }
        
        Write-Verbose "Preparing e-mail"
        try{
            $body = "<HTML><BODY><p>Hi $($fullRequest.FieldValues.Site_x0020_Admin.LookupValue.Split(" ")[0]),</p>
                <p>Your new <a href=`"$($newPnpTeam.siteUrl)`">External
                Sharing Site</a> is available for you now. This is a new Modern-style External
                Sharing Site, which should be more familiar to work with than the
                older Classic-style Sites. We have also made some improvements to the way
                external users get access, which should make them significantly simpler to use
                (particularly where $externalPartyType`s don&#39;t use 365 themselves). There is <a 
                href=`"https://anthesisllc.sharepoint.com/:w:/r/sites/Resources-IT/Shared%20Documents/Guides/Guide%20to%20sharing%20Modern%20External%20Sites.docx?d=w00ab51f7f8d243ada762abef1a7d3a55&amp;csf=1&amp;e=LlJKZO&amp;web=1`">a
                new Sharing Guide available</a> and <a 
                href=`"https://anthesisllc.sharepoint.com/:w:/r/sites/Resources-IT/Shared%20Documents/Guides/Guide%20for%20External%20Users%20to%20access%20Anthesis%20External%20Sharing%20Sites.docx?d=w0e63dc2ec7b3483da8913a9124945e49&csf=1&e=0YKVyT`">
                an external version</a> you can send to $externalPartyType`s if they get stuck.</p>

                <p>There are also some additional guides to get you started if
                you want to do anything fancier than simply sharing files:</p>

                <UL><LI><a href=`"https://anthesisllc.sharepoint.com/sites/Resources-IT/_layouts/15/DocIdRedir.aspx?ID=HXX7CE52TSD2-1759992947-62`">Changing
                the logo for your Site</a></LI>

                <LI><a href=`"https://anthesisllc.sharepoint.com/sites/Resources-IT/_layouts/15/DocIdRedir.aspx?ID=HXX7CE52TSD2-1759992947-196`">Creating/editing
                pages in SharePoint</a></LI>

                <LI><a href=`"https://anthesisllc.sharepoint.com/sites/Resources-IT/_layouts/15/DocIdRedir.aspx?ID=HXX7CE52TSD2-1759992947-105`">Creating
                links in SharePoint</a></LI>

                <LI><a href=`"https://anthesisllc.sharepoint.com/sites/Resources-IT/_layouts/15/DocIdRedir.aspx?ID=HXX7CE52TSD2-1759992947-42`">Adding
                icons to your link</a></LI></UL>

                <p>Love,</p>

                <p>The External Sharing Site Robot</p>
                </BODY></HTML>"
            $cc = $(convertTo-arrayOfEmailAddresses ($fullRequest.FieldValues.Site_x0020_Owners.Email +","+ $fullRequest.FieldValues.Site_x0020_Members.Email) | sort | select -Unique)
            Write-Verbose "Sending e-mail"
            if($cc){
                Send-MailMessage  -BodyAsHtml $body -Subject "External Site for $externalParty created" -to $fullRequest.FieldValues.Site_x0020_Admin.Email -Cc $cc -bcc $($365creds.UserName) -from "ExternalSiteRobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Encoding UTF8
                Write-Verbose "E-mail sent"
                }
            else{
                Send-MailMessage  -BodyAsHtml $body -Subject "External Site for $externalParty created" -to $fullRequest.FieldValues.Site_x0020_Admin.Email -bcc $($365creds.UserName) -from "ExternalSiteRobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Encoding UTF8 -Credential $365creds
                Write-Verbose "E-mail sent"
                } #Send-MailMessage doesn't support Empty CC option
            }
        catch{$_}
        }
    catch{$_}
    }










<#
$clientSite = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/sites/anthesisllc.sharepoint.com:/subs" -Verbose -returnEntireResponse
$clientSiteLists = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/sites/anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,9fb8ecd6-c87d-485d-a488-26fd18c62303/Lists" -Verbose

$clientSiteLists| ? {$_.DisplayName -match "Request"}

06365ce6-07a5-41e9-b0aa-a90c9f1edc3f
$listItems = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/sites/anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,faed84bc-70be-4e35-bfbf-cdab31aeeb99/Lists/06365ce6-07a5-41e9-b0aa-a90c9f1edc3f/items" -Verbose
$columns = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/sites/anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,faed84bc-70be-4e35-bfbf-cdab31aeeb99/Lists/06365ce6-07a5-41e9-b0aa-a90c9f1edc3f/items/$($listItems[100].id)/?expand=fields"

#>