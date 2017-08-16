Import-Module .\_CSOM_Library-SPO.psm1
Import-Module .\_PS_Library_MSOL.psm1

$anthesisMySite = 'https://anthesisllc-my.sharepoint.com/' # This needs to be the mySite where the userdata lives.
$anthesisAdminSite = 'https://anthesisllc-admin.sharepoint.com/' # This needs to be the "admin" site.
$csomCreds = set-csomCredentials
$ctx = new-csomContext -fullSitePath $anthesisMySite -sharePointCredentials $csomCreds

$anthesisUserProfileCollection = @()
$anthesisSharePointUsers = $ctx.Web.SiteUsers
$ctx.Load($anthesisSharePointUsers)
$ctx.ExecuteQuery()
#Create an Object [People Manager] to retrieve profile information
$anthesisPeople = New-Object Microsoft.SharePoint.Client.UserProfiles.PeopleManager($ctx)

ForEach($user in $anthesisSharePointUsers){
    $userprofile = $anthesisPeople.GetPropertiesFor($user.LoginName)
    $ctx.Load($userprofile)
    $ctx.ExecuteQuery() #Sorry :(
    $anthesisUserProfileCollection += $userprofile #Enumerate each user account and save it to the userProfileCollection array
    }

$kev = $anthesisUserProfileCollection | ?{$_.DisplayName -eq "Kev Maitland"}
$ellie = $anthesisUserProfileCollection | ?{$_.DisplayName -eq "Eleanor Penney"}


$adminCtx = new-csomContext -fullSitePath $anthesisAdminSite -sharePointCredentials $csomCreds
$adminAnthesisPeople = New-Object Microsoft.SharePoint.Client.UserProfiles.PeopleManager($adminCtx)
$adminAnthesisPeople.SetSingleValueProfileProperty($ellie.AccountName, "SPS-RegionalSettings-Initialized", $true)

$adminAnthesisPeople.SetSingleValueProfileProperty($ellie.AccountName, "SPS-MUILanguages", "en-GB")
$adminAnthesisPeople.SetSingleValueProfileProperty($ellie.AccountName, "SPS-RegionalSettings-FollowWeb", "false")
$adminAnthesisPeople.SetSingleValueProfileProperty($ellie.AccountName, "SPS-TimeZone", 2)
$adminAnthesisPeople.SetSingleValueProfileProperty($ellie.AccountName, "SPS-Locale", 2057)
$adminAnthesisPeople.SetSingleValueProfileProperty($ellie.AccountName, "SPS-CalendarType", 1)
$adminAnthesisPeople.SetSingleValueProfileProperty($ellie.AccountName, "SPS-AltCalendarType", 1)
$adminCtx.ExecuteQuery()







