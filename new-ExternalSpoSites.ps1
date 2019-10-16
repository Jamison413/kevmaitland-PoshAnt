$365creds = set-MsolCredentials
connect-to365 -credential $365creds

$teamBotDetails = Import-Csv "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\teambotdetails.txt"
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

Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/clients" -Credentials $365creds
$requests = Get-PnPListItem -List "External Client Site Requests" -Query "<View><Query><Where><Eq><FieldRef Name='Status'/><Value Type='String'>Awaiting creation</Value></Eq></Where></Query></View>"
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/subs" -Credentials $365creds
$requests += Get-PnPListItem -List "External Subcontractor Site Requests" -Query "<View><Query><Where><Eq><FieldRef Name='Status'/><Value Type='String'>Awaiting creation</Value></Eq></Where></Query></View>"
if($requests){$selectedRequests = $requests | select {$_.FieldValues.Title},{$_.FieldValues.ClientName.Label},{$_.FieldValues.Site_x0020_Admin.LookupValue},{$_.FieldValues.Site_x0020_Owners.LookupValue -join ", "},{$_.FieldValues.Site_x0020_Members.LookupValue -join ", "},{$_.FieldValues.GUID.Guid} | Out-GridView -PassThru -Title "Highlight any requests to process and click OK"}

foreach ($currentRequest in $selectedRequests){
    $fullRequest = $requests | ? {$_.FieldValues.GUID.Guid -eq $currentRequest.'$_.FieldValues.GUID.Guid'}
    $managers = convertTo-arrayOfEmailAddresses ($fullRequest.FieldValues.Site_x0020_Owners.Email +","+ $fullRequest.FieldValues.Site_x0020_Admin.Email+","+ $((Get-PnPConnection).PSCredential.UserName)) | sort | select -Unique
    $members = convertTo-arrayOfEmailAddresses ($managers + $fullRequest.FieldValues.Site_x0020_Members.Email) | sort | select -Unique
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
    try{
        $result = new-externalGroup -displayName $("External - $($fullRequest.FieldValues.Title)").Trim(" ") -managerUpns $managers -teamMemberUpns $members -membershipManagedBy 365 -tokenResponse $tokenResponse -alsoCreateTeam $false -pnpCreds $365creds -Verbose -ErrorAction Stop
        #If there we no errors returned, assume it worked and notify finish the setup
        #Add a link to the new Site on the External Hub
        Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/sites/external" -Credentials $365creds
        Add-PnPNavigationNode -Location QuickLaunch -Title $($fullRequest.FieldValues.Title) -Url $result.SiteUrl -First -External -Parent 2252 #2252 is the "Modern External Sites" NavNode

        switch($fullRequest.FieldValues.FileDirRef.Split("/")[1]){
            "clients" {
                test-pnpConnectionMatchesResource -resourceUrl "https://anthesisllc.sharepoint.com/clients" -connectIfDifferent $true -pnpCreds $365creds | Out-Null
                Set-PnPListItem -List "External Client Site Requests" -Identity $fullRequest.Id -Values @{Status="Created"}
                $externalParty = $fullRequest.FieldValues.ClientName.Label
                $externalPartyType = "client"
                }
            "subs"    {
                test-pnpConnectionMatchesResource -resourceUrl "https://anthesisllc.sharepoint.com/subs" -connectIfDifferent $true -pnpCreds $365creds | Out-Null
                Set-PnPListItem -List "0c68ca6f-06fe-449b-8cf1-c0dbe7fddd5c" -Identity $fullRequest.Id -Values @{Status="Created"} #"External Subcontractor Site Requests" List 
                $externalParty = $fullRequest.FieldValues.Subcontractor_x002f_Supplier_x00.Label
                $externalPartyType = "subcontractor"
                }
            default   {}
            }

        $body = "<HTML><BODY><p>Hi $($fullRequest.FieldValues.Site_x0020_Admin.LookupValue.Split(" ")[0]),</p>
            <p>Your new <a href=`"$siteUrl`">External
            Sharing Site</a> is available for you now. This is a new Modern-style External
            Sharing Site, which should be more familiar to work with than the
            older Classic-style Sites. We have also made some improvements to the way
            external users get access, which should make them significantly simpler to use
            (particularly where $externalPartyType`s don&#39;t use 365 themselves). There is <a 
            href=`"https://anthesisllc.sharepoint.com/:w:/r/sites/Resources-IT/Shared%20Documents/Guides/Guide%20to%20sharing%20Modern%20External%20Sites.docx?d=w00ab51f7f8d243ada762abef1a7d3a55&amp;csf=1&amp;e=LlJKZO&amp;web=1`">a
            new Sharing Guide available</a> too - it&#39;s internal documentation, but it&#39;s not
            sensitive, so feel free to forward it on to your $externalPartyType`s if they get stuck.</p>

            <p>There are also some additional guides to get you started if
            you want to do anything fancier that simply sharing files:</p>

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
        Send-MailMessage  -BodyAsHtml $body -Subject "External Site for $externalParty created" -to $fullRequest.FieldValues.Site_x0020_Admin.Email -Cc $(convertTo-arrayOfEmailAddresses ($fullRequest.FieldValues.Site_x0020_Owners.Email +","+ $fullRequest.FieldValues.Site_x0020_Members.Email) | sort | select -Unique) -bcc $((Get-PnPConnection).PSCredential.UserName) -from "ExternalSiteRobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com"

        }
    catch{Write-Error $_}
    }


