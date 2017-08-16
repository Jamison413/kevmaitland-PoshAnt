Import-Module _CSOM_Library-SPO.psm1

#region Statics - don't change these.
$webUrl = "https://anthesisllc.sharepoint.com"
#$internalTeamTemplate = "{7FD4CC3D-B615-4930-A041-3ADB8C6509EA}#Default Community Site"
$internalTeamTemplate = "{32C80FAC-E19D-495E-B923-6216EE14A571}#AnthesisTeamSite_v1.1"
$externalSharingTemplate = "{8C3E419E-EADC-4032-A7CD-BC5778A30F9C}#Default External Sharing Site"
$colorPaletteUrl = "/_catalogs/theme/15/AnthesisPalette_Orange.spcolor"
$spFontUrl = "/_catalogs/theme/15/Anthesis_fontScheme_Montserrat_uploaded.spfont"
#endregion
#region Variables - change these to create new sites
#$siteCollection = "/teams/communities" 
$siteCollection = "/sites/external" 
$sitePath = "/" 
$siteName = "IKEA CAT18 - external site" 
$siteUrlEndStub = "fishwick-ikea" 
#$siteTemplate = $internalTeamTemplate
$siteTemplate = $externalSharingTemplate
$inheritPermissions = $false 
$inheritTopNav = $true
$owner = "Ellen Upton"
$precreatedSecurityGroupForMembers = ""
#endregion


#Build and customise a new Site
$csomCredentials = set-csomCredentials 
add-site -credentials $csomCredentials -webUrl $webUrl -siteCollection $siteCollection -sitePath $sitePath -siteName $siteName -siteUrlEndStub $siteUrlEndStub -siteTemplate $siteTemplate -inheritPermissions $inheritPermissions -inheritTopNav $inheritTopNav -owner $owner
add-memberToGroup -credentials $csomCredentials -webUrl $webUrl -siteCollection $siteCollection -sitePath $sitePath -groupName "$siteName Owners" -memberToAdd $owner
add-memberToGroup -credentials $csomCredentials -webUrl $webUrl -siteCollection $siteCollection -sitePath $sitePath -groupName "$siteName Members" -memberToAdd $precreatedSecurityGroupForMembers
apply-theme -credentials $csomCredentials -webUrl $webUrl -siteCollection $siteCollection -site $sitePath -colorPaletteUrl $colorPaletteUrl -fontSchemeUrl $spFontUrl -backgroundImageUrl $null -shareGenerated $false

get-webTempates -credentials $csomCredentials -webUrl $webUrl -siteCollection $siteCollection -site $sitePath

#Brand multiple Sites:
$siteCollection = "/sites/external" #Main Site Collections are "/", "/teams/communities", "/teams/sym" Full list available here: https://anthesisllc-admin.sharepoint.com/_layouts/15/online/SiteCollections.aspx
$sites = @("/unite") #@("/intsus","/anyOtherSite", "/anyOtherSite/SubSite")
foreach($site in $sites){
    apply-theme -webUrl $webUrl -siteCollection $siteCollection -site $site -colorPaletteUrl $colorPaletteUrl -fontSchemeUrl $spFontUrl -backgroundImageUrl $null -shareGenerated $false
    }

#Update TopNav Bar:
$sitesToUpdate= @("/teams/hr","/teams/finance","/teams/marketing","/teams/administration","/clients","/teams/confidential","/global", "/teams/all", "/teams/communities", "/subs")
foreach ($site in $sitesToUpdate){
    $nodesToAdd = [ordered]@{"Global"="https://anthesisllc.sharepoint.com/global";"Clients"= "https://anthesisllc.sharepoint.com/Anthesis Projects";"Resources"="https://anthesisllc.sharepoint.com/Anthesis Resources";"External"="https://anthesisllc.sharepoint.com/sites/external";"Kimble"="https://login.salesforce.com/";"Search"="https://anthesisllc.sharepoint.com/search";"Help"="https://anthesisllc.sharepoint.com/help"}
    $newNodesToAdd = [ordered]@{"Global"="https://anthesisllc.sharepoint.com/global";"Clients"= "https://anthesisllc.sharepoint.com/clients";"Resources"="https://anthesisllc.sharepoint.com/teams/all/Lists/Internal%20Teams";"External"="https://anthesisllc.sharepoint.com/sites/external";"Kimble"="https://login.salesforce.com/";"Search"="https://anthesisllc.sharepoint.com/search";"Help"="https://anthesisllc.sharepoint.com/help"}
    #set-navTopNodes -webUrl $webUrl -siteCollectionOrSite $site -deleteAllBeforeAdding $true -hashTableOfNodes $newNodesToAdd
    }







#copy-ListItems -webUrl $webUrl -srcSite "/teams/sym/" -srcListName "Sym Groups" -destSite "/teams/all/" -destListName "Sym Groups"
#copy-ListItems -webUrl $webUrl -srcSite "/teams/communities/" -srcListName "Communities" -destSite "/teams/all/" -destListName "Communities"
#copy-libraryItems -webUrl $webUrl -srcSite "/Anthesis Resources/" -srcLibraryName "Images" -destSite "/teams/all/" -destLibraryName "Images"


$extGroups  = $groups | ? {$_.Users -match "ext"} | select Title
$groups | ? {$_.users -eq "Admin Info"}
