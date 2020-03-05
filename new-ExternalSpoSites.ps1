$365creds = set-MsolCredentials
connect-to365 -credential $365creds

$teamBotDetails = Import-Csv "$env:USERPROFILE\Desktop\teambotdetails.txt"
$resource = "https://graph.microsoft.com"
$tenantId = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.TenantId)
$clientId = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.ClientID)
$redirect = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.Redirect)
$secret   = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.Secret)

$ReqTokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    client_Id     = $clientID
    Client_Secret = $secret
    } 
$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody

$requests = @()
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/clients" -Credentials $365creds
$requests += Get-PnPListItem -List "External Client Site Requests" -Query "<View><Query><Where><Eq><FieldRef Name='Status'/><Value Type='String'>Awaiting creation</Value></Eq></Where></Query></View>"
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/subs" -Credentials $365creds
$requests += Get-PnPListItem -List "External Subcontractor Site Requests" -Query "<View><Query><Where><Eq><FieldRef Name='Status'/><Value Type='String'>Awaiting creation</Value></Eq></Where></Query></View>"
if($requests){[array]$selectedRequests = $requests | select {$_.FieldValues.Title},{$_.FieldValues.ClientName.Label},{$_.FieldValues.Site_x0020_Admin.LookupValue},{$_.FieldValues.Site_x0020_Owners.LookupValue -join ", "},{$_.FieldValues.Site_x0020_Members.LookupValue -join ", "},{$_.FieldValues.GUID.Guid} | Out-GridView -PassThru -Title "Highlight any requests to process and click OK"}

    $hideFromGal = $false
    $blockExternalMail = $false
    $accessType = "Private"
    $autoSubscribe = $true
    $groupClassification = "External"
    $alsoCreateTeam = $false


foreach ($currentRequest in $selectedRequests){

    $fullRequest = $requests | ? {$_.FieldValues.GUID.Guid -eq $currentRequest.'$_.FieldValues.GUID.Guid'}
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
    if($managers -notcontains ((Get-PnPConnection).PSCredential.UserName)){
        $addExecutingUserAsTemporaryOwner = $true
        [array]$managers += ((Get-PnPConnection).PSCredential.UserName)
        }
    if($members -notcontains ((Get-PnPConnection).PSCredential.UserName)){
        $addExecutingUserAsTemporaryMember = $true
        [array]$members += ((Get-PnPConnection).PSCredential.UserName)
        }

    
    try{
        $new365Group = new-365Group -displayName $("External - $($fullRequest.FieldValues.Title)".Trim(" ")) -managerUpns $managers -teamMemberUpns $members -memberOf $null -hideFromGal $hideFromGal -blockExternalMail $blockExternalMail -accessType Private -autoSubscribe $autoSubscribe -groupClassification $groupClassification -membershipManagedBy 365 -tokenResponse $tokenResponse -pnpCreds $365creds -ownersAreRealManagers $false -alsoCreateTeam $alsoCreateTeam -Verbose
        Write-Host -ForegroundColor Yellow "Site Admin is : [$($fullRequest.FieldValues.Site_x0020_Admin.LookupValue)]"
        
        Write-Verbose "Getting PnPUnifiedGroup [$($new365Group.displayName)] - this is a faster way to get the SharePoint URL than using the UnifiedGroup object"
        Connect-PnPOnline -AccessToken $tokenResponse.access_token
        $newPnpTeam = Get-PnPUnifiedGroup -Identity $new365Group.ExternalDirectoryObjectId

        Write-Verbose "Adding Navigation to External Hub"
        Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/sites/external" -Credentials $365creds
        Add-PnPNavigationNode -Location QuickLaunch -Title $($fullRequest.FieldValues.Title) -Url $newPnpTeam.SiteUrl -First -External -Parent 2252 #2252 is the "Modern External Sites" NavNode

        #Aggrivatingly, you can't manipulate Pages with Graph yet, and Add-PnpFile doesn;t support AccessTokens, so we need to go old-school:
        if($addExecutingUserAsTemporaryOwner){
            $addExecutingUserAsTemporarySiteCollectionAdmin = test-isUserSiteCollectionAdmin -pnpUnifiedGroupObject $newPnpTeam -accessToken $tokenResponse.access_token -pnpCreds $365creds -addPermissionsIfMissing $true
            }
        copy-spoPage -sourceUrl "https://anthesisllc.sharepoint.com/sites/Resources-IT/SitePages/External-Site-Template-Candidate.aspx" -destinationSite $newPnpTeam.SiteUrl -pnpCreds $365creds -overwriteDestinationFile $true -renameFileAs "LandingPage.aspx" | Out-Null
        test-pnpConnectionMatchesResource -resourceUrl $newPnpTeam.SiteUrl -pnpCreds $pnpCreds -connectIfDifferent $true | Out-Null
        if((test-pnpConnectionMatchesResource -resourceUrl $newPnpTeam.SiteUrl) -eq $true){
            Write-Verbose "Setting Homepage"
            Set-PnPHomePage  -RootFolderRelativeUrl "SitePages/LandingPage.aspx" | Out-Null
            }

        Write-Verbose "Setting Hub Site association"
        Add-PnPHubSiteAssociation -Site $newPnpTeam.SiteUrl -HubSite "https://anthesisllc.sharepoint.com/sites/ExternalHub" | Out-Null
        Write-Verbose "Opening in browser"
        start-Process $newPnpTeam.SiteUrl

        if($addExecutingUserAsTemporaryOwner){
            Remove-UnifiedGroupLinks -Identity $newPnpTeam.GroupId -LinkType Owner -Links $((Get-PnPConnection).PSCredential.UserName) -Confirm:$false
            Remove-DistributionGroupMember -Identity $new365Group.CustomAttribute2 -Member $((Get-PnPConnection).PSCredential.UserName) -Confirm:$false -BypassSecurityGroupManagerCheck:$true
            }
        if($addExecutingUserAsTemporaryMember){
            Remove-UnifiedGroupLinks -Identity $newPnpTeam.GroupId -LinkType Member -Links $((Get-PnPConnection).PSCredential.UserName) -Confirm:$false
            }
        <# Not required - this will be undone and redone by set-standardSitePermissions in a moment)
        if($addExecutingUserAsTemporarySiteCollectionAdmin){
            Remove-PnPSiteCollectionAdmin -Owners $($pnpCreds.UserName)
            }
        #>
        Write-Verbose "set-standardSitePermissions [$($new365Group.DisplayName)]"
        set-standardSitePermissions -unifiedGroupObject $new365Group -tokenResponse $tokenResponse -pnpCreds $365creds -Verbose

        switch($fullRequest.FieldValues.FileDirRef.Split("/")[1]){
            "clients" {
                Write-Verbose "Updating Client Request: Status = [Created],Url=[$($newPnpTeam.SiteUrl)]"
                test-pnpConnectionMatchesResource -resourceUrl "https://anthesisllc.sharepoint.com/clients" -connectIfDifferent $true -pnpCreds $365creds | Out-Null
                Set-PnPListItem -List "External Client Site Requests" -Identity $fullRequest.Id -Values @{Status="Created";URL=$newPnpTeam.SiteUrl}
                $externalParty = $fullRequest.FieldValues.ClientName.Label
                $externalPartyType = "client"
                Write-Verbose "Getting Managed MetaData term"
                $termGroup = $(Get-PnPTermGroup "Kimble") 
                $termSet = $(Get-PnPTermSet -TermGroup $termGroup -Identity "Clients")
                $managedMetaDataTerm = Get-PnPTerm -Identity $fullRequest.FieldValues.ClientName.Label -TermSet $termSet -TermGroup $termGroup -Includes CustomProperties
                Write-Verbose "`tRetrieved: [$($managedMetaDataTerm.Name)]"
                }
            "subs"    {
                Write-Verbose "Updating Supplier Request: Status = [Created]"
                test-pnpConnectionMatchesResource -resourceUrl "https://anthesisllc.sharepoint.com/subs" -connectIfDifferent $true -pnpCreds $365creds | Out-Null
                Set-PnPListItem -List "0c68ca6f-06fe-449b-8cf1-c0dbe7fddd5c" -Identity $fullRequest.Id -Values @{Status="Created"} #"External Subcontractor Site Requests" List 
                $externalParty = $fullRequest.FieldValues.Subcontractor_x002f_Supplier_x00.Label
                $externalPartyType = "subcontractor"
                }
            default   {}
            }
        
        <#if(![string]::IsNullOrWhiteSpace($managedMetaDataTerm.CustomProperties["DocLibId"])){
            Write-Verbose "Creating links between Client DocLib and new Site"
            Connect-PnPOnline -AccessToken $tokenResponse.access_token
            $clientDocLib = 
            }#>

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
                Send-MailMessage  -BodyAsHtml $body -Subject "External Site for $externalParty created" -to $fullRequest.FieldValues.Site_x0020_Admin.Email -Cc $cc -bcc $((Get-PnPConnection).PSCredential.UserName) -from "ExternalSiteRobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Encoding UTF8
                Write-Verbose "E-mail sent"
                }
            else{
                Send-MailMessage  -BodyAsHtml $body -Subject "External Site for $externalParty created" -to $fullRequest.FieldValues.Site_x0020_Admin.Email -bcc $((Get-PnPConnection).PSCredential.UserName) -from "ExternalSiteRobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Encoding UTF8
                Write-Verbose "E-mail sent"
                } #Send-MailMessage doesn't support Empty CC option
            }
        catch{$_}
        }
    catch{$_}
    }


