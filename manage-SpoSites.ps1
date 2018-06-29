Import-Module _PS_Library_MSOL
Import-Module _CSOM_Library-SPO.psm1
Import-Module _REST_Library-SPO.psm1

$msolCredentials = set-MsolCredentials #Set these once as a PSCredential object and use that to build the CSOM SharePointOnlineCredentials object and set the creds for REST
$restCredentials = new-spoCred -username $msolCredentials.UserName -securePassword $msolCredentials.Password
$csomCredentials = new-csomCredentials -username $msolCredentials.UserName -password $msolCredentials.Password

$logFile = "C:\ScriptLogs\manage-SpoSites.log"
$errorLogFile = "C:\ScriptLogs\manage-SpoSites_Errors.log"
$smtpServer = "anthesisgroup-com.mail.protection.outlook.com"

#region Get the Admin to pick the request/s to process
#Get the Taxonomy Data for the Site Collection as there's Managed MetaData fields to retrieve 
$webUrl = "https://anthesisllc.sharepoint.com"
$taxonomyListName = "TaxonomyHiddenList"
$taxononmyData = get-itemsInList -serverUrl $webUrl -sitePath "/" -listName $taxonomyListName -suppressProgress $true -restCreds $restCredentials -verboseLogging $true -logFile $logFile

#Get the Client Site requests that have a status of "Awaiting creation"
$clientsSite = "/clients"
$clientSiteRequestListName = "External Client Site Requests"
#$oDataUnprocessedClientRequests = '$select=*'
$oDataUnprocessedClientRequests = '$select=ClientName,Id,Title,Site_x0020_AdminId,Site_x0020_Admin/Name,Site_x0020_Admin/Title'
$oDataUnprocessedClientRequests += ',Site_x0020_OwnersId,Site_x0020_Owners/Id,Site_x0020_Owners/Title'
$oDataUnprocessedClientRequests += ',Site_x0020_MembersId,Site_x0020_Members/Id,Site_x0020_Members/Title'
$oDataUnprocessedClientRequests += ',Site_x0020_VisitorsId,Site_x0020_Visitors/Id,Site_x0020_Visitors/Title'
$oDataUnprocessedClientRequests += '&$expand=Site_x0020_Admin/Id,Site_x0020_Owners/Id,Site_x0020_Members/Id,Site_x0020_Visitors/Id'    #,Site_x0020_Members,Site_x0020_Visitors
$oDataUnprocessedClientRequests += '&$filter=Status eq ''Awaiting creation'''
$unprocessedClientRequests = get-itemsInList -serverUrl $webUrl -sitePath $clientsSite -listName $clientSiteRequestListName -suppressProgress $false -oDataQuery $oDataUnprocessedClientRequests -restCreds $restCredentials
#Standardise the Requests:
foreach($request in $unprocessedClientRequests){
    $req = New-Object -TypeName PSObject
    $req | Add-Member -MemberType NoteProperty -Name RequestType -Value "Client"
    #These are read directly from the List:
    $req | Add-Member -MemberType NoteProperty -Name "SiteName" -Value $request.Title
    #These are taxonomy fields and the Ids need to cross-referenced with the TaxonomyHiddenList to get the labels
    $req | Add-Member -MemberType NoteProperty -Name "For" -Value $($taxononmyData | ?{$_.IdForTerm -eq $request.ClientName.TermGuid} | %{$_.Term})
    #These are People/Group fields and need expanding
    if($request.Site_x0020_Admin.__deferred -eq $null){
        $req | Add-Member -MemberType NoteProperty -Name "SiteAdminId" -Value $request.Site_x0020_Admin.Name
        $req | Add-Member -MemberType NoteProperty -Name "SiteAdminName" -Value $request.Site_x0020_Admin.Title
        }
    #These are Multi People/Group fields and need expanding and iterating through
    $owners = @{}
    if($request.Site_x0020_Owners.__deferred -eq $null){
        foreach($userOrGroup in $request.Site_x0020_Owners.results){[hashtable]$owners.Add($userOrGroup.id,$userOrGroup.Title)}
        }
    $members = @{}
    if($request.Site_x0020_Members.__deferred -eq $null){
        foreach($userOrGroup in $request.Site_x0020_Members.results){[hashtable]$members.Add($userOrGroup.id,$userOrGroup.Title)}
        }
    $visitors = @{}
    if($request.Site_x0020_Visitors.__deferred -eq $null){
        foreach($userOrGroup in $request.Site_x0020_Visitors.results){[hashtable]$visitors.Add($userOrGroup.id,$userOrGroup.Title)}
        }

    Add-Member -InputObject $req -MemberType NoteProperty -Name "Owners" -Value $owners
    Add-Member -InputObject $req -MemberType NoteProperty -Name "Members" -Value $members
    Add-Member -InputObject $req -MemberType NoteProperty -Name "Visitors" -Value $visitors
    Add-Member -InputObject $req -MemberType NoteProperty -Name "Id" -Value $request.id
    Add-Member -InputObject $req -MemberType NoteProperty -Name "listContentType" -Value $request.__metadata.type

    [array]$clientRequests += $req
    }

#Get the Supplier Site requests that have a status of "Awaiting creation"
$suppliersSite = "/subs"
$supplierSiteRequestListName = "External Subcontractor Site Requests"
$oDataUnprocessedSupplierRequests = '$select=Subcontractor_x002f_Supplier_x00,Title,Id,Site_x0020_AdminId,Site_x0020_Admin/Name,Site_x0020_Admin/Title'
$oDataUnprocessedSupplierRequests += ',Site_x0020_OwnersId,Site_x0020_Owners/Id,Site_x0020_Owners/Title'
$oDataUnprocessedSupplierRequests += ',Site_x0020_MembersId,Site_x0020_Members/Id,Site_x0020_Members/Title'
$oDataUnprocessedSupplierRequests += ',Site_x0020_VisitorsId,Site_x0020_Visitors/Id,Site_x0020_Visitors/Title'
$oDataUnprocessedSupplierRequests += '&$expand=Site_x0020_Admin/Id,Site_x0020_Owners/Id,Site_x0020_Members/Id,Site_x0020_Visitors/Id'    #,Site_x0020_Members,Site_x0020_Visitors
$oDataUnprocessedSupplierRequests += '&$filter=Status eq ''Awaiting creation'''
$unprocessedSupplierRequests = get-itemsInList -serverUrl $webUrl -sitePath $suppliersSite -listName $supplierSiteRequestListName -suppressProgress $false -oDataQuery $oDataUnprocessedSupplierRequests -restCreds $restCredentials
#Standardise the Requests:
foreach($request in $unprocessedSupplierRequests){
    $req = New-Object -TypeName PSObject
    $req | Add-Member -MemberType NoteProperty -Name RequestType -Value "Supplier"
    #These are read directly from the List:
    $req | Add-Member -MemberType NoteProperty -Name "SiteName" -Value $request.Title
    #These are taxonomy fields and the Ids need to cross-referenced with the TaxonomyHiddenList to get the labels
    $req | Add-Member -MemberType NoteProperty -Name "For" -Value $($taxononmyData | ?{$_.IdForTerm -eq $request.Subcontractor_x002f_Supplier_x00.TermGuid} | %{$_.Term})
    #These are People/Group fields and need expanding
    if($request.Site_x0020_Admin.__deferred -eq $null){
        $req | Add-Member -MemberType NoteProperty -Name "SiteAdminId" -Value $request.Site_x0020_Admin.Name
        $req | Add-Member -MemberType NoteProperty -Name "SiteAdminName" -Value $request.Site_x0020_Admin.Title
        }
    #These are Multi People/Group fields and need expanding and iterating through
    $owners = @{}
    if($request.Site_x0020_Owners.__deferred -eq $null){
        foreach($userOrGroup in $request.Site_x0020_Owners.results){[hashtable]$owners.Add($userOrGroup.id,$userOrGroup.Title)}
        }
    $members = @{}
    if($request.Site_x0020_Members.__deferred -eq $null){
        foreach($userOrGroup in $request.Site_x0020_Members.results){[hashtable]$members.Add($userOrGroup.id,$userOrGroup.Title)}
        }
    $visitors = @{}
    if($request.Site_x0020_Visitors.__deferred -eq $null){
        foreach($userOrGroup in $request.Site_x0020_Visitors.results){[hashtable]$visitors.Add($userOrGroup.id,$userOrGroup.Title)}
        }

    Add-Member -InputObject $req -MemberType NoteProperty -Name "Owners" -Value $owners
    Add-Member -InputObject $req -MemberType NoteProperty -Name "Members" -Value $members
    Add-Member -InputObject $req -MemberType NoteProperty -Name "Visitors" -Value $visitors
    Add-Member -InputObject $req -MemberType NoteProperty -Name "Id" -Value $request.Id
    Add-Member -InputObject $req -MemberType NoteProperty -Name "listContentType" -Value $request.__metadata.type

    [array]$supplierRequests += $req
    }

#Combine all the Requests into a signle list to present to the Admin
$allRequests = $clientRequests
$allRequests += $supplierRequests

#Present them and record any selections
$selectedRequests = $allRequests | Out-GridView -PassThru -Title "Highlight any requests to process and click OK"

#endregion


#Now process any requests authorised by the Admin
foreach ($currentRequest in $selectedRequests){
    switch ($currentRequest.RequestType){
        {$_ -in 'Client',"Supplier"} {
            #These define the specifics for External Client/Supplier Sites
            $sitePath = "/" 
            $siteName = $currentRequest.SiteName 
            $alphaNumericRegexPattern = '[^a-zA-Z0-9]'
            $siteUrlEndStub = $currentRequest.SiteName -replace $alphaNumericRegexPattern, ""
            $inheritPermissions = $false 
            $inheritTopNav = $false
            $siteTemplate = "{5C86D4B3-3D30-4C36-BDE1-6A1779799A45}#ExternalSiteTemplate"
            #$siteTemplate = "{8C3E419E-EADC-4032-A7CD-BC5778A30F9C}#Default External Sharing Site"
            $siteCollection = "/sites/external" 
            $colorPaletteUrl = "/_catalogs/theme/15/AnthesisPalette_Orange.spcolor"
            $spFontUrl = "/_catalogs/theme/15/Anthesis_fontScheme_Montserrat_uploaded.spfont"

            #Create the Site (branded automatically) then create and configure the membership groups
            add-site -credentials $csomCredentials -webUrl $webUrl -siteCollection $siteCollection -sitePath $sitePath -siteName $siteName -siteUrlEndStub $siteUrlEndStub -siteTemplate $siteTemplate -inheritPermissions $inheritPermissions -inheritTopNav $inheritTopNav -owner $currentRequest.SiteAdminName
            add-memberToGroup -credentials $csomCredentials -webUrl $webUrl -siteCollection $siteCollection -sitePath $sitePath -groupName "$siteName Owners" -memberToAdd $currentRequest.SiteAdminName
            $currentRequest.Owners.Keys | % {add-memberToGroup -credentials $csomCredentials -webUrl $webUrl -siteCollection $siteCollection -sitePath $sitePath -groupName "$siteName Owners" -memberToAdd $currentRequest.Owners[$_]}
            $currentRequest.Members.Keys | % {add-memberToGroup -credentials $csomCredentials -webUrl $webUrl -siteCollection $siteCollection -sitePath $sitePath -groupName "$siteName Members" -memberToAdd $currentRequest.Members[$_]}
            $currentRequest.Visitors.Keys | % {add-memberToGroup -credentials $csomCredentials -webUrl $webUrl -siteCollection $siteCollection -sitePath $sitePath -groupName "$siteName Visitors" -memberToAdd $currentRequest.Members[$_]}

            #Then update the Request to show that it's been processed
            #(there should be a check here to ensure no errors occurred during the Site creation)
            switch ($currentRequest.RequestType){
                'Client' {
                    $requestSite = $clientsSite
                    $requestListName = $clientSiteRequestListName
                    }
                'Supplier' {
                    $requestSite = $suppliersSite
                    $requestListName = $supplierSiteRequestListName
                    }
                }
            #Get the Id of the current Admin (no error checking required as we won't get this far if the Admin hasn't authenticated successfully)
            $ctx = new-csomContext -fullSitePath $($webUrl+$requestSite) -sharePointCredentials $csomCredentials
            $admin = $ctx.Web.EnsureUser($csomCredentials.UserName)
            $ctx.Load($admin)
            $ctx.ExecuteQuery()
            $ctx.Dispose()
            $digest = new-spoDigest -serverUrl $webUrl -sitePath $requestSite -restCreds $restCredentials 
            check-digestExpiry -serverUrl $webUrl -sitePath $requestSite -digest $digest -restCreds $restCredentials
            update-itemInList -serverUrl $webUrl -sitePath $requestSite -listName $requestListName -predeterminedItemType $currentRequest.listContentType -itemId $currentRequest.Id -hashTableOfItemData @{Status="Created";Site_x0020_Created_x0020_ById=$admin.Id} -restCreds $restCredentials -digest $digest
            }
        'Confidential' {}
        'Team' {}
        default {}
        }
    Write-host "$weburl/sites/external/$siteName".Replace(" ","")
    }
#endregion






<#
#region Statics - don't change these.
################
#    Confidential Site settings
$siteTemplate = "{4527360C-CD78-4F36-BE30-DFBE2EAF34E6}#ConfidentialSiteTemplate"
$siteCollection = "/teams/confidential" 
$colorPaletteUrl = "/_catalogs/theme/15/AnthesisPalette_Orange.spcolor"
$spFontUrl = "/_catalogs/theme/15/Anthesis_fontScheme_Montserrat_uploaded.spfont"
$owner = "kev maitland" #Override anything else to prevent numpties tinkering with the security settings
################
#    Internal Team Site settings
#$internalTeamTemplate = "{7FD4CC3D-B615-4930-A041-3ADB8C6509EA}#Default Community Site"
$siteTemplate = "{32C80FAC-E19D-495E-B923-6216EE14A571}#AnthesisTeamSite_v1.1"
$siteCollection = "/teams/communities" 
$colorPaletteUrl = "/_catalogs/theme/15/AnthesisPalette_Orange.spcolor"
$spFontUrl = "/_catalogs/theme/15/Anthesis_fontScheme_Montserrat_uploaded.spfont"
$sitePath = "/"
$siteName = "NA Growth Team"
$siteUrlEndStub = "nagrowthteam"
$inheritTopNav = $true
$inheritPermissions = $false
$precreatedSecurityGroupForMembers = "NAGrowthTeam"
$owner = "Rosanna.Collorafi"


#Build and customise a new Site
add-site -credentials $csomCredentials -webUrl $webUrl -siteCollection $siteCollection -sitePath $sitePath -siteName $siteName -siteUrlEndStub $siteUrlEndStub -siteTemplate $siteTemplate -inheritPermissions $inheritPermissions -inheritTopNav $inheritTopNav -owner $owner
add-memberToGroup -credentials $csomCredentials -webUrl $webUrl -siteCollection $siteCollection -sitePath $sitePath -groupName "$siteName Owners" -memberToAdd $owner
add-memberToGroup -credentials $csomCredentials -webUrl $webUrl -siteCollection $siteCollection -sitePath $sitePath -groupName "$siteName Members" -memberToAdd $precreatedSecurityGroupForMembers

#Rolled into add-site
#remove-userFromSite -credentials $csomCredentials -webUrl $webUrl -siteCollection $siteCollection -sitePath $sitePath$siteUrlEndStub -memberToRemove ("Kev Maitland")
#apply-theme -credentials $csomCredentials -webUrl $webUrl -siteCollection $siteCollection -site $sitePath$siteUrlEndStub -colorPaletteUrl $colorPaletteUrl -fontSchemeUrl $spFontUrl -backgroundImageUrl $null -shareGenerated $false
#apply-theme -credentials $csomCredentials -webUrl $webUrl -siteCollection "/teams/IT" -site "" -colorPaletteUrl $colorPaletteUrl -fontSchemeUrl $spFontUrl -backgroundImageUrl $null -shareGenerated $false

get-webTempates -credentials $csomCredentials -webUrl $webUrl -siteCollection $siteCollection -site $sitePath

#Brand multiple Sites:
$siteCollection = "/teams/IT" #Main Site Collections are "/", "/teams/communities", "/teams/sym" Full list available here: https://anthesisllc-admin.sharepoint.com/_layouts/15/online/SiteCollections.aspx
$sites = @("/unite") #@("/intsus","/anyOtherSite", "/anyOtherSite/SubSite")
foreach($site in $sites){
    apply-theme -credentials $csomCredentials -webUrl $webUrl -siteCollection $siteCollection -site $site -colorPaletteUrl $colorPaletteUrl -fontSchemeUrl $spFontUrl -backgroundImageUrl $null -shareGenerated $false
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

#>

