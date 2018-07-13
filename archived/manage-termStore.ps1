Import-Module _CSOM_Library-SPO
$msolCredentials = set-MsolCredentials #Set these once as a PSCredential object and use that to build the CSOM SharePointOnlineCredentials object and set the creds for REST
$csomCredentials = new-csomCredentials -username $msolCredentials.UserName -password $msolCredentials.Password

$webUrl = "https://anthesisllc.sharepoint.com"


#Sync [Offices] and [Non-Geographic Workplaces] terms to [Primary Workplaces]
reuse-allTermsInTermStore -credentials $csomCredentials -webUrl $webUrl -siteCollection "/" -sourceTermGroup "Anthesis" -sourceTermSet "Offices" -destTermGroup "Anthesis" -destTermSet "Primary Workplaces"
reuse-allTermsInTermStore -credentials $csomCredentials -webUrl $webUrl -siteCollection "/" -sourceTermGroup "Anthesis" -sourceTermSet "Non-Geographic Workplaces" -destTermGroup "Anthesis" -destTermSet "Primary Workplaces"

#Sync [Offices] and [Geographic Preseneces] terms to [Secondary Workplaces]
reuse-allTermsInTermStore -credentials $csomCredentials -webUrl $webUrl -siteCollection "/" -sourceTermGroup "Anthesis" -sourceTermSet "Offices" -destTermGroup "Anthesis" -destTermSet "Secondary Workplaces"
reuse-allTermsInTermStore -credentials $csomCredentials -webUrl $webUrl -siteCollection "/" -sourceTermGroup "Anthesis" -sourceTermSet "Geographic Presences" -destTermGroup "Anthesis" -destTermSet "Secondary Workplaces"