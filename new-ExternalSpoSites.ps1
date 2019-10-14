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
if($requests){$selectedRequests = $requests | select {$_.FieldValues.Title},{$_.FieldValues.ClientName.Label},{$_.FieldValues.Site_x0020_Admin.LookupValue},{$_.FieldValues.Site_x0020_Owners.LookupValue -join ", "},{$_.FieldValues.Site_x0020_Members.LookupValue -join ", "},{$_.FieldValues.GUID.Guid} | Out-GridView -PassThru -Title "Highlight any requests to process and click OK"}
foreach ($currentRequest in $selectedRequests){
    $fullRequest = $requests | ? {$_.FieldValues.GUID.Guid -eq $currentRequest.'$_.FieldValues.GUID.Guid'}
    $managers = convertTo-arrayOfEmailAddresses ($fullRequest.FieldValues.Site_x0020_Owners.Email +","+ $fullRequest.FieldValues.Site_x0020_Admin.Email+","+ $((Get-PnPConnection).PSCredential.UserName)) | sort | select -Unique
    $members = convertTo-arrayOfEmailAddresses ($managers + $fullRequest.FieldValues.Site_x0020_Members.Email) | sort | select -Unique
    try{
        $result = new-externalGroup -displayName $fullRequest.FieldValues.Title -managerUpns $managers -teamMemberUpns $members -membershipManagedBy 365 -tokenResponse $tokenResponse -alsoCreateTeam $false -pnpCreds $365creds -Verbose -ErrorAction Stop
        #If there we no errors returned, assume it worked and notify iva e-mail
        $body = "<HTML><BODY><p>Hi $($fullRequest.FieldValues.Site_x0020_Admin.LookupValue.Split(" ")[0]),</p>
            <p>Your new <a href=`"$siteUrl`">External
            Sharing Site</a> is available for you now. This is a new Modern-style External
            Sharing Site, which should be more familiar to work with than the
            older Classic-style Sites. We have also made some improvements to the way
            external users get access, which should make them significantly simpler to use
            (particularly where clients don&#39;t use 365 themselves). There is <a 
            href=`"https://anthesisllc.sharepoint.com/:w:/r/sites/Resources-IT/Shared%20Documents/Guides/Guide%20to%20sharing%20Modern%20External%20Sites.docx?d=w00ab51f7f8d243ada762abef1a7d3a55&amp;csf=1&amp;e=LlJKZO&amp;web=1`">a
            new Sharing Guide available</a> too - it&#39;s internal documentation, but it&#39;s not
            sensitive, so feel free to forward it on to your clients if they get stuck.</p>

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
        Send-MailMessage  -BodyAsHtml $body -Subject "External Site for $($fullRequest.FieldValues.ClientName.Label) created" -to $fullRequest.FieldValues.Site_x0020_Admin.Email -Cc $(convertTo-arrayOfEmailAddresses ($fullRequest.FieldValues.Site_x0020_Owners.Email +","+ $fullRequest.FieldValues.Site_x0020_Members.Email) | sort | select -Unique) -bcc $((Get-PnPConnection).PSCredential.UserName) -from "ExternalSiteRobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com"

        }
    catch{}
    }

Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/subs" -Credentials $365creds
$requests += Get-PnPListItem -List "External Subcontractor Site Requests" -Query "<View><Query><Where><Eq><FieldRef Name='Status'/><Value Type='String'>Awaiting creation</Value></Eq></Where></Query></View>"

