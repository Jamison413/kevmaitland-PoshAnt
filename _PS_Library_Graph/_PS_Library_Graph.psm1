function add-graphArrayOfFoldersToDrive(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true,ParameterSetName="DriveId")]
            [string]$graphDriveId 
        ,[parameter(Mandatory = $true,ParameterSetName="DriveObject")]
            [string]$graphDriveObject 
        ,[parameter(Mandatory = $true,ParameterSetName="DriveId")]
            [parameter(Mandatory = $true,ParameterSetName="DriveObject")]
            [array]$foldersAndSubfoldersArray
        ,[parameter(Mandatory = $true,ParameterSetName="DriveId")]
            [parameter(Mandatory = $true,ParameterSetName="DriveObject")]
            [psobject]$tokenResponse
        ,[parameter(Mandatory = $true,ParameterSetName="DriveId")]
            [parameter(Mandatory = $true,ParameterSetName="DriveObject")]
            [ValidateSet(“Fail”,”Rename”,”Replace”)]
            [string]$conflictResolution
        )

    switch ($PsCmdlet.ParameterSetName){
        'DriveObject' {$graphDriveId = $graphDriveObject.Id}
        }
    Write-Verbose "add-graphArrayOfFoldersToDrive [$($graphDriveId)]"    
    
    #Prep the folders array (in case the user has provided junk like $foldersAndSubfoldersArray = @("Test","test\test2\test3\test4","test","/test/TeSt2\","tEST #3","Test | #4")
    $expandedFoldersAndSubfoldersArray = ,@()
    $foldersAndSubfoldersArray | % {
        $thisFolder = $_.Replace("\","/").Trim("/")
        $expandingFolderPath = ""
        $thisFolder.Split("/") | % {
            $expandingFolderPath += "$(sanitise-forSharePointGroupName $_)/"
            $expandedFoldersAndSubfoldersArray += $expandingFolderPath.Trim("/")
            }
        }

    $driveItemsToReturn = ,@()
    #Iterate through our sanitised folder array and create the folders
    $expandedFoldersAndSubfoldersArray | Sort-Object -Unique | ? {![string]::IsNullOrWhiteSpace($_)} | % {
        $folderName = Split-Path $_ -Leaf
        if($folderName -eq $_){ #If it is _just_ a folder (i.e. not a subfolder), just create it
            try{
                $newFolder = add-graphFolderToDrive -graphDriveId $graphDriveId -folderName $folderName -tokenResponse $tokenResponse -conflictResolution $conflictResolution -Verbose:$VerbosePreference -ErrorAction Stop
                $driveItemsToReturn += $newFolder
                }
            catch{
                if($_.Exception.Message -eq "The remote server returned an error: (409) Conflict."){ #If the folder already existed, get and return it
                    $existingFolder = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/drives/$graphDriveId/root:/$folderName"
                    $driveItemsToReturn += $existingFolder
                    }
                else{Write-Error $_}
                }
            }
        else{ #If it _is_ a subfolder, we also need to supply the relative path (and invoke-graphGet doesn't like a $null value for -relativePathToFolder)
            try{
                $relativePath = Split-Path $_ -Parent
                $newFolder = add-graphFolderToDrive -graphDriveId $graphDriveId -folderName $folderName -tokenResponse $tokenResponse -conflictResolution $conflictResolution -Verbose:$VerbosePreference -ErrorAction Stop -relativePathToFolder $([uri]::EscapeDataString($relativePath))
                $driveItemsToReturn += $newFolder
                }
            catch{
                if($_.Exception.Message -eq "The remote server returned an error: (409) Conflict."){ #If the folder already existed, get and return it
                    $existingFolder = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/drives/$graphDriveId/root:/$relativePath/$folderName"
                    $driveItemsToReturn += $existingFolder
                    }
                else{Write-Error $_}
                }
            }
        }

    $driveItemsToReturn
    }
function add-graphFolderToDrive(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true,ParameterSetName="DriveId_Id")]
            [parameter(Mandatory = $true,ParameterSetName="DriveId_RelativePath")]
            [parameter(Mandatory = $true,ParameterSetName="DriveId_Neither")]
            [string]$graphDriveId 
        ,[parameter(Mandatory = $true,ParameterSetName="DriveObject_Id")]
            [parameter(Mandatory = $true,ParameterSetName="DriveObject_RelativePath")]
            [parameter(Mandatory = $true,ParameterSetName="DriveObject_Neither")]
            [string]$graphDriveObject 
        ,[parameter(Mandatory = $true,ParameterSetName="DriveId_Id")]
            [parameter(Mandatory = $true,ParameterSetName="DriveId_RelativePath")]
            [parameter(Mandatory = $true,ParameterSetName="DriveId_Neither")]
            [parameter(Mandatory = $true,ParameterSetName="DriveObject_Id")]
            [parameter(Mandatory = $true,ParameterSetName="DriveObject_RelativePath")]
            [parameter(Mandatory = $true,ParameterSetName="DriveObject_Neither")]
            [string]$folderName
        ,[parameter(Mandatory = $true,ParameterSetName="DriveId_Id")]
            [parameter(Mandatory = $true,ParameterSetName="DriveObject_Id")]
            [string]$parentItemId
        ,[parameter(Mandatory = $true,ParameterSetName="DriveId_RelativePath")]
            [parameter(Mandatory = $true,ParameterSetName="DriveObject_RelativePath")]
            [string]$relativePathToFolder
        ,[parameter(Mandatory = $true,ParameterSetName="DriveId_Id")]
            [parameter(Mandatory = $true,ParameterSetName="DriveId_RelativePath")]
            [parameter(Mandatory = $true,ParameterSetName="DriveId_Neither")]
            [parameter(Mandatory = $true,ParameterSetName="DriveObject_Id")]
            [parameter(Mandatory = $true,ParameterSetName="DriveObject_RelativePath")]
            [parameter(Mandatory = $true,ParameterSetName="DriveObject_Neither")]
            [psobject]$tokenResponse
        ,[parameter(Mandatory = $true,ParameterSetName="DriveId_Id")]
            [parameter(Mandatory = $true,ParameterSetName="DriveId_RelativePath")]
            [parameter(Mandatory = $true,ParameterSetName="DriveId_Neither")]
            [parameter(Mandatory = $true,ParameterSetName="DriveObject_Id")]
            [parameter(Mandatory = $true,ParameterSetName="DriveObject_RelativePath")]
            [parameter(Mandatory = $true,ParameterSetName="DriveObject_Neither")]
            [ValidateSet(“Fail”,”Rename”,”Replace”)]
            [string]$conflictResolution
        )
    switch ($PsCmdlet.ParameterSetName){
        {$_ -match 'DriveObject'} {$graphDriveId = $graphDriveObject.Id}
        {$_ -match 'RelativePath'} {
            $useRelativePath = $true
            $relativePathToFolder = $relativePathToFolder.Replace("\","/").Trim("/")
            }
        }

    if($parentItemId){Write-Verbose "add-graphFolderToDrive [$($graphDriveId)]\[$($parentItemId)]\[$($folderName)]"}
    else{Write-Verbose "add-graphFolderToDrive [$($graphDriveId)]\[$($folderName)] _[$($PsCmdlet.ParameterSetName)]_"}
    
    if(!$parentItemId){$parentItemId = "root"}

    $folderHash = @{
        "name"   = $folderName
        "folder" = @{}
        "@microsoft.graph.conflictBehavior" = "$($conflictResolution.ToLower())"
        }
    
    if($useRelativePath){$graphQuery = "/drives/$graphDriveId/root:/$relativePathToFolder`:/children".Replace("root:/:/","root:/")}
    else{$graphQuery = "/drives/$graphDriveId/items/$parentItemId/children".Replace("items/root","root")}
    Write-Verbose $graphQuery
    invoke-graphPost -tokenResponse $tokenResponse -graphQuery $graphQuery -graphBodyHashtable $folderHash -Verbose:$VerbosePreference
    }
function add-graphLicenseToUser(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true,ParameterSetName="Friendly")]
            [parameter(Mandatory = $true,ParameterSetName="Guid")]
            [parameter(Mandatory = $true,ParameterSetName="Guids")]
            [psobject]$tokenResponse
        ,[parameter(Mandatory = $true,ParameterSetName = "Friendly")]
            [parameter(Mandatory = $true,ParameterSetName="Guid")]
            [parameter(Mandatory = $true,ParameterSetName="Guids")]
            [string]$userIdOrUpn
        ,[parameter(Mandatory = $true,ParameterSetName = "Friendly")]
            [ValidateSet("K1","E1","E3","E5","EMS","AudioConferencing","DomesticCalling","InternationalCalling","Project","Visio")]
            [string]$licenseFriendlyName 
        ,[parameter(Mandatory = $true,ParameterSetName = "Guids")]
            [string]$licenseGuid
        ,[parameter(Mandatory = $true,ParameterSetName = "Guid")]
            [string[]]$disabledPlansGuids = @()
        ,[parameter(Mandatory = $true,ParameterSetName = "Guids")]
            [string[]]$licenseGuids
        )
    $licensesToRemove = @()
    if(<#Licenses contain K1/E1/E3/E5#>$false){
        #We have to remove any conflicting licenses at the same time
        #get user licesnses
        #build appropriate remove hash
        #$licensesToRemove = @("guidToRemove")
        }

    switch ($PsCmdlet.ParameterSetName){
        "Friendly" {
            [string[]]$licenseGuids = get-microsoftProductInfo -getType GUID -fromType FriendlyName -fromValue $licenseFriendlyName
            }
        "Guid" {
            [string[]]$licenseGuids = $licenseGuid
            }

        }

    #Iterate through the supplied/derived licenseGuids
    $licenseGuids | % {
        $thisLicenseDefinition = @{"skuId"=$_}
        $thisLicenseDefinition.Add("disabledPlans",$disabledPlansGuids) #This cannot proc if $PsCmdlet.ParameterSetName -eq "Guids", so we don't need to worry about which disabledPlans belong to which licenseGuid
        [array]$licenseArray += $thisLicenseDefinition
        }
    
    $graphBodyHashtable = @{
        "addLicenses"=$licenseArray
        "removeLicenses"=$licensesToRemove
        }

    invoke-graphPost -tokenResponse $tokenResponse -graphQuery "/users/$userIdOrUpn/assignLicense" -graphBodyHashtable $graphBodyHashtable -Verbose:$VerbosePreference
    }
function add-graphUsersToGroup(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true,ParameterSetName="UserIds")]
            [parameter(Mandatory = $true,ParameterSetName="UserUpns")]
            [psobject]$tokenResponse
        ,[parameter(Mandatory = $true,ParameterSetName = "UserIds")]
            [parameter(Mandatory = $true,ParameterSetName = "UserUpns")]
            [string]$graphGroupId
        ,[parameter(Mandatory = $true,ParameterSetName = "UserIds")]
            [parameter(Mandatory = $true,ParameterSetName = "UserUpns")]
            [ValidateSet("Members","Owners")]
            [string]$memberType 
        ,[parameter(Mandatory = $true,ParameterSetName = "UserUpns")]
            [string[]]$graphUserUpns
        ,[parameter(Mandatory = $true,ParameterSetName = "UserIds")]
            [string[]]$graphUserIds
        )
    
    switch ($PsCmdlet.ParameterSetName){
        "UserUpns" {$graphUserIds = $graphUserUpns} #UPNs work natively in place of Ids for this endpoint, so no need to handle them differrently after all! 
        }

    $graphUserIds | % {
        $bodyHash = @{"@odata.id"="https://graph.microsoft.com/v1.0/users/$_"}
        invoke-graphPost -tokenResponse $tokenResponse -graphQuery "/groups/$graphGroupId/$memberType/`$ref" -graphBodyHashtable $bodyHash -Verbose:$VerbosePreference
        }
    }
function add-graphWebsiteTabToChannel(){
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$teamId
        ,[parameter(Mandatory = $true)]
            [string]$channelName
        ,[parameter(Mandatory = $true)]
            [string]$tabName
        ,[parameter(Mandatory = $true)]
            [string]$tabDestinationUrl
        )

    Write-Verbose "add-graphWebsiteTabToChannel | Getting Channels"
    $newGraphTeamChannel = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/teams/$teamId/channels?`$filter displayName eq '$channelName'" | ? {$_.DisplayName -eq $channelName} #$filter doesn't currently work on this endpoint :/
    Write-Verbose "add-graphWebsiteTabToChannel | Getting Channels Tabs"
    $newGraphTeamGeneralChannelTabs = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/teams/$teamId/channels/$($newGraphTeamChannel.id)/tabs"
    $newGraphTeamGeneralChannelTabs | ? {$_.displayName -eq "$tabName"} | % {
        Write-Verbose "Removing old [$tabName] tab"
        invoke-graphDelete -tokenResponse $tokenResponse -graphQuery "/teams/$teamId/channels/$($newGraphTeamChannel.id)/tabs/$($_.id)" 
        }
    $tabConfiguration = @{
        "entityId"=$null
        "contentUrl"=$tabDestinationUrl
        "websiteUrl"=$tabDestinationUrl
        "removeUrl"=$null
        }
    $tabBody = @{
        "displayName"=$tabName
        "teamsApp@odata.bind"="https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.web"
        "configuration"=$tabConfiguration
        }

    Write-Verbose "add-graphWebsiteTabToChannel | Creating new [$tabName] Tab in [$channelName] Channel linking to [$tabDestinationUrl]"
    invoke-graphPost -tokenResponse $tokenResponse -graphQuery "/teams/$teamId/channels/$($newGraphTeamChannel.id)/tabs" -graphBodyHashtable $tabBody
    
    }
function delete-graphDriveItem(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse
        ,[parameter(Mandatory = $true)]
            [string]$graphDriveId 
        ,[parameter(Mandatory = $true)]
            [string]$graphDriveItemId
        ,[parameter(Mandatory = $false)]
            [string]$eTag
        )
    
    if($eTag){
        $deleteBody = @{"if-match"=$eTag}
        invoke-graphDelete -tokenResponse $tokenResponse -graphQuery "/drives/$graphDriveId/items/$graphDriveItemId" -graphBodyHashtable $deleteBody -Verbose:$VerbosePreference
        }
    else{
        invoke-graphDelete -tokenResponse $tokenResponse -graphQuery "/drives/$graphDriveId/items/$graphDriveItemId"  -Verbose:$VerbosePreference
        }
    }
function delete-graphListItem(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse
        ,[parameter(Mandatory = $true)]
            [string]$graphSiteId 
        ,[parameter(Mandatory = $true)]
            [string]$graphListId
        ,[parameter(Mandatory = $true)]
            [string]$graphItemId

        )
        #Need to expand to allow for ListName and SiteName as well as the Id's (to match other functions here)
        invoke-graphDelete -tokenResponse $tokenResponse -graphQuery "sites/$graphSiteId/lists/$graphListId/items/$graphItemId"  -Verbose:$VerbosePreference
        
}
function get-groupAdminRoleEmailAddresses(){
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        )

    $admins = @()
    get-graphAdministrativeRoleMembers -tokenResponse $tokenResponse -roleName 'User Account Administrator' | % {$admins += $_.userPrincipalName}
    get-graphAdministrativeRoleMembers -tokenResponse $tokenResponse -roleName 'Exchange Service Administrator' | % {$admins += $_.userPrincipalName}
    $admins | Sort-Object -Unique
    }
function get-graphAdministrativeRoleMembers(){
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [ValidateSet("Application Administrator","Authentication Administrator","Company Administrator","Conditional Access Administrator","Device Administrators","Device Managers","Directory Readers","Directory Writers","Exchange Service Administrator","Global Reader","Helpdesk Administrator","Intune Service Administrator","Lync Service Administrator","Service Support Administrator","SharePoint Service Administrator","Teams Service Administrator","User Account Administrator")]
            [string]$roleName 
        )
    $roleHash = @{ #This uses the DirectoryRole Id, not roleTemplateId
        "Application Administrator"="741ab5a1-5e1f-42be-8c6b-911e2db72b97"
        "Authentication Administrator"="117dfbf3-a802-4518-8cf7-3234f17b1fb8"
        "Company Administrator"="ceb5b002-5318-47fa-8e30-72d5fac100a4"
        "Conditional Access Administrator"="50383b78-5362-4206-809b-34f276985216"
        "Device Administrators"="8e9abe65-a35e-4053-b665-2b9d72738065"
        "Device Managers"="ca2c6c1c-a2a5-468d-b77c-ab6bd224a8cc"
        "Directory Readers"="1c5d15f2-c1b0-4dcb-b560-f0ae28e06d51"
        "Directory Writers"="97c440f3-d4b7-465e-baa5-99c7ebfe3f42"
        "Exchange Service Administrator"="41fe0192-9858-4741-a3cd-f593a78c2b1f"
        "Global Reader"="c6d7d5b2-4cc4-4e97-b79d-5217c8a87395"
        "Helpdesk Administrator"="bc35604d-8ef8-43aa-b7e3-3a3d4f2a3a3d"
        "Intune Service Administrator"="0990da52-1ee9-4539-b9cd-833f40c0d350"
        "Lync Service Administrator"="ca882d67-c3e0-49bf-be4a-a6c5aa255bf8"
        "Service Support Administrator"="2385910b-ed12-4102-b96c-c0009632ab44"
        "SharePoint Service Administrator"="e0de39cb-c1af-4d0d-b221-641ac49b4bec"
        "Teams Service Administrator"="50a0ea4d-b517-4405-8c56-00a55a8b165c"
        "User Account Administrator"="2c620b2d-2ac3-4663-bbd6-f09a8f861ebc"
        }
    <#--$roleHash = @{  
        "Application Administrator"="9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3"
        "Authentication Administrator"="c4e39bd9-1100-46d3-8c65-fb160da0071f"
        "Company Administrator"="62e90394-69f5-4237-9190-012177145e10"
        "Conditional Access Administrator"="b1be1c3e-b65d-4f19-8427-f6fa0d97feb9"
        "Device Administrators"="9f06204d-73c1-4d4c-880a-6edb90606fd8"
        "Device Managers"="2b499bcd-da44-4968-8aec-78e1674fa64d"
        "Directory Readers"="88d8e3e3-8f55-4a1e-953a-9b9898b8876b"
        "Directory Writers"="9360feb5-f418-4baa-8175-e2a00bac4301"
        "Exchange Service Administrator"="41fe0192-9858-4741-a3cd-f593a78c2b1f"
        "Global Reader"="f2ef992c-3afb-46b9-b7cf-a126ee74c451"
        "Helpdesk Administrator"="729827e3-9c14-49f7-bb1b-9608f156bbb8"
        "Intune Service Administrator"="3a2c62db-5318-420d-8d74-23affee5d9d5"
        "Lync Service Administrator"="75941009-915a-4869-abe7-691bff18279e"
        "Service Support Administrator"="f023fd81-a637-4b56-95fd-791ac0226033"
        "SharePoint Service Administrator"="f28a1f50-f6e7-4571-818b-6a12f2af6b6c"
        "Teams Service Administrator"="69091246-20e8-4a56-aa4d-066075b2a7a8"
        "User Account Administrator"=" fe930be7-5e62-47db-91af-98c3a49a38b1"        
        }--#>
    
    invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/directoryRoles/$($roleHash[$roleName])/members"

    }
function get-graphAuthCode() {
     [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [string]$clientID
        ,[parameter(Mandatory = $true)]
            [string]$redirectUri
        ,[parameter(Mandatory = $false)]
            [string]$scope
        )

    $clientIDEncoded = [System.Web.HttpUtility]::UrlEncode($clientID)
    $redirectUriEncoded =  [System.Web.HttpUtility]::UrlEncode($redirectUri)
    $resourceEncoded = [System.Web.HttpUtility]::UrlEncode("https://graph.microsoft.com")
    $scopeEncoded = [System.Web.HttpUtility]::UrlEncode($scope) #"https://outlook.office.com/user.readwrite.all" "https://outlook.office.com/Directory.AccessAsUser.All"

    Add-Type -AssemblyName System.Windows.Forms
    if($scope){$url = "https://login.microsoftonline.com/common/oauth2/authorize?response_type=code&redirect_uri=$redirectUriEncoded&client_id=$clientID&resource=$resourceEncoded&prompt=admin_consent&scope=$scopeEncoded"}
    else{$url = "https://login.microsoftonline.com/common/oauth2/authorize?response_type=code&redirect_uri=$redirectUriEncoded&client_id=$clientID&resource=$resourceEncoded&prompt=admin_consent"}
    $form = New-Object -TypeName System.Windows.Forms.Form -Property @{Width=440;Height=640}
    $web  = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width=420;Height=600;Url=($url -f ($Scope -join "%20")) }
    $docComp  = {
        $uri = $web.Url.AbsoluteUri        
        if ($uri -match "error=[^&]*|code=[^&]*") {$form.Close() }
        }
    $web.ScriptErrorsSuppressed = $true
    $web.Add_DocumentCompleted($docComp)
    $form.Controls.Add($web)
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() | Out-Null
    $queryOutput = [System.Web.HttpUtility]::ParseQueryString($web.Url.Query)
    $output = @{}
    foreach($key in $queryOutput.Keys){
        $output["$key"] = $queryOutput[$key]
        }
    $output
    }
function get-graphDrives(){
     [cmdletbinding()]
    param(
        [parameter(Mandatory = $true,ParameterSetName = "fromUrl")]
            [parameter(Mandatory = $true,ParameterSetName = "fromSiteId")]
            [parameter(Mandatory = $true,ParameterSetName = "fromUpn")]
            [parameter(Mandatory = $true,ParameterSetName = "fromGroupId")]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true,ParameterSetName = "fromUrl")]
            [string]$siteUrl
        ,[parameter(Mandatory = $true,ParameterSetName = "fromSiteId")]
            [string]$siteGraphId
        ,[parameter(Mandatory = $true,ParameterSetName = "fromGroupId")]
            [string]$groupGraphId
        ,[parameter(Mandatory = $true,ParameterSetName = "fromUpn")]
            [ValidatePattern("@")]
            [string]$teamUpn
        ,[parameter(Mandatory = $false,ParameterSetName = "fromUrl")]
            [parameter(Mandatory = $false,ParameterSetName = "fromSiteId")]
            [parameter(Mandatory = $false,ParameterSetName = "fromUpn")]
            [parameter(Mandatory = $false,ParameterSetName = "fromGroupId")]
            [switch]$returnOnlyDefaultDocumentsLibrary
        ,[parameter(Mandatory = $false,ParameterSetName = "fromUrl")]
            [parameter(Mandatory = $false,ParameterSetName = "fromSiteId")]
            [parameter(Mandatory = $false,ParameterSetName = "fromUpn")]
            [parameter(Mandatory = $false,ParameterSetName = "fromGroupId")]
            [string]$filterDriveName
        )
    
    if($returnOnlyDefaultDocumentsLibrary){$endpoint = "/drive"}
    else{$endpoint = "/drives"}

    switch ($PsCmdlet.ParameterSetName){ #Build the query based on the parameters supplied. Because we're dealing with several permutations of endpoints (/groups vs /sites & /drive vs /drives), this looks more complicated than it really is. 
        "fromUpn" { #If we've only got a UPN, we need to get the corresponding group Id
            Write-Verbose "get-graphDrives | Getting GroupId from UPN [$teamUpn]"
            $groupGraphId = (get-graphGroups -tokenResponse $tokenResponse -filterUpn $teamUpn).id
            }
         {@("fromUpn","fromGroupId") -contains $_} { #Now we can use the /groups endpoint with either the /drive or /drives endpoint (whichever we picked before the switch statement)
            Write-Verbose "get-graphDrives | Getting from $_"
            $query = "/groups/$groupGraphId$endpoint"
            }
        "fromUrl" { #If we're working with $siteUrl, we'll need to get $siteGraphId (which is more of a faff)
            Write-Verbose "get-graphDrives | Getting SiteId from URL [$siteUrl]"
            if([string]::IsNullOrWhiteSpace($siteGraphId)){ 
                if($siteUrl -match "anthesisllc.sharepoint.com"){$siteUrl = ($siteUrl -Split "anthesisllc.sharepoint.com")[1].Trim("/")} #Get the serverRelativeUrl
                $siteGraphId = (invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/sites/anthesisllc.sharepoint.com:/$siteUrl").id
                if([string]::IsNullOrWhiteSpace($siteGraphId)){ #Weirdly this doesn't seem to work, despite the same query being submitted to graph.
                    Write-Verbose "Weird, that should have worked. trying again"
                    $siteGraphId = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0//sites/anthesisllc.sharepoint.com:/$siteUrl" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method GET).id
                    }
                }
            }       
       {@("fromUrl","fromSiteId") -contains $_} {#Now we can use the /sites endpoint with either the /drive or /drives endpoint (whichever we picked before the switch statement)
            Write-Verbose "get-graphDrives | Getting from $_"
            $query = "/sites/$siteGraphId$endpoint"
            }
        }
    
    #Now build the refiner based on the other paramters supplied
    if($filterDriveName){
        $filter += " and name eq '$filterDriveName'"
        }
    if(![string]::IsNullOrWhiteSpace($filter)){
        if($filter.StartsWith(" and ")){$filter = $filter.Substring(5,$filter.Length-5)}
        $filter = "`$filter=$filter"
        }

    $refiner = "?"+$select
    if($filter){
        if($refiner.Length -gt 1){$refiner = $refiner+"&"} #If there is already another query option in the refiner, use the '&' symbol to concatenate the the strings
        $refiner = $refiner+$filter
        }    
    if($refiner.Length -gt 1){$query = $query+[uri]::EscapeDataString($refiner)}

    #Finally, submit the query and return the results
    $drives = invoke-graphGet -tokenResponse $tokenResponse -graphQuery $query
    $drives

    }
function get-graphGroups(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true,ParameterSetName = "ambiguous")]
            [parameter(Mandatory = $true,ParameterSetName = "explicit")]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true,ParameterSetName = "explicit")]
            [string]$filterId
        ,[parameter(Mandatory = $false,ParameterSetName = "ambiguous")]
            [ValidatePattern("@")]
            [string]$filterUpn
        ,[parameter(Mandatory = $false,ParameterSetName = "ambiguous")]
            [string]$filterDisplayName
        ,[parameter(Mandatory = $false,ParameterSetName = "ambiguous")]
            [string]$filterDisplayNameStartsWith
        ,[parameter(Mandatory = $false,ParameterSetName = "ambiguous")]
            [ValidateSet("Unified","Security","MailEnabledSecurity","Distribution")]
            [string]$filterGroupType
        ,[parameter(Mandatory = $false,ParameterSetName = "ambiguous")]
            [parameter(Mandatory = $false,ParameterSetName = "explicit")]
            [switch]$selectAllProperties = $false
        )
    if($selectAllProperties){ #We're only dealing with one select option (all or default)
        $select  = "`$select=displayName,id,description,mail,anthesisgroup_UGSync,deletedDateTime,classification,createdDateTime,creationOptions,groupTypes,isAssignableToRole,mailEnabled,mailNickname,onPremisesDomainName,onPremisesLastSyncDateTime,onPremisesNetBiosName,onPremisesSamAccountName,onPremisesSecurityIdentifier,onPremisesSyncEnabled,preferredDataLocation,proxyAddresses,renewedDateTime,resourceBehaviorOptions,resourceProvisioningOptions,securityEnabled,securityIdentifier,visibility,onPremisesProvisioningErrors"
        }

    switch ($PsCmdlet.ParameterSetName){
        “explicit”  {
            invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/groups/$filterId$select" -Verbose:$VerbosePreference
            return
            }
        }
    
    #Build up the filter & select conditions into a refiner we can append to the query string
    if($filterUpn){$filter += " and mail eq '$filterUpn'"}
    if($filterDisplayName){$filter += " and displayName eq '$([uri]::EscapeDataString($filterDisplayName))'"}
    if($filterDisplayNameStartsWith){$filter += " and startswith(displayName,'$([uri]::EscapeDataString($filterDisplayNameStartsWith))')"}
    switch($filterGroupType){
        "Unified" {$filter += " and groupTypes/any(a:a eq 'Unified')"}
        "Security" {$filter += " and mailEnabled eq 'false'"}
        #Graph doesn't support 'ne' or 'null' in lambda queries. Have to filter these client-side instead :/
        #"MailEnabledSecurity" {$filter += " and groupTypes/any(a:a ne 'Unified') and mailEnabled eq 'true' and securityEnabled eq 'true'"}
        "MailEnabledSecurity" {$filter += " and mailEnabled eq true and securityEnabled eq true"}
        #"Distribution" {$filter += " and groupTypes/any(a:a ne 'Unified') and mailEnabled eq 'true' and securityEnabled eq 'false'"}
        "Distribution" {$filter += " and mailEnabled eq true and securityEnabled eq false"}
        }

    if(![string]::IsNullOrWhiteSpace($filter)){
        if($filter.StartsWith(" and ")){$filter = $filter.Substring(5,$filter.Length-5)}
        $filter = "`$filter=$filter"
        }

    $refiner = "?"+$select
    if($filter){
        if($refiner.Length -gt 1){$refiner = $refiner+"&"} #If there is already another query option in the refiner, use the '&' symbol to concatenate the the strings
        $refiner = $refiner+$filter
        }
    
    Write-Verbose "`$filter = $filter"
    Write-Verbose "`$select = $select"
    Write-Verbose "`$refiner = $refiner"

    $results = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/groups$refiner" -Verbose:$VerbosePreference

    if($filterGroupType -eq "MailEnabledSecurity" -or $filterGroupType -eq "Distribution"){
        $results | ? {$_.groupTypes -notcontains "Unified"} | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name ExternalDirectoryObjectId -Value $_.id}
        $results 
        }
    else{$results}

    }
function get-graphGroupWithUGSyncExtensions(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $false)]
            [ValidatePattern("@")]
            [string]$filterUpn
        ,[parameter(Mandatory = $false)]
            [string]$filterId
        ,[parameter(Mandatory = $false)]
            [string]$filterDisplayName
        ,[parameter(Mandatory = $false)]
            [string]$filterDataManagersGroupId
        ,[parameter(Mandatory = $false)]
            [string]$filterMembersGroupId
        ,[parameter(Mandatory = $false)]
            [string]$filterCombinedGroupId
        ,[parameter(Mandatory = $false)]
            [string]$filterSharedMailboxId
        ,[parameter(Mandatory = $false)]
            [ValidateSet("365","AAD")]
            [string]$filterMasterMembership
        ,[parameter(Mandatory = $false)]
            [ValidateSet("Internal","External","Sym","Confidential")]
            [string]$filterClassifcation
        ,[parameter(Mandatory = $false)]
            [ValidateSet("Private","Public")]
            [string]$filterPrivacy
        ,[parameter(Mandatory = $false)]
            [switch]$selectAllProperties
        )
    #Add $filters for the various properties
    if($filterUpn){$additionalFilters += " and mail eq '$filterUpn'"}
    if($filterId){$additionalFilters += " and id eq '$filterId'"}
    if($filterDisplayName){$additionalFilters += " and displayName eq '$([uri]::EscapeDataString($filterDisplayName))'"}
    if($filterDataManagersGroupId){$additionalFilters += " and anthesisgroup_UGSync/dataManagerGroupId eq '$filterDataManagersGroupId'"}
    if($filterMembersGroupId){$additionalFilters += " and anthesisgroup_UGSync/memberGroupId eq '$filterMembersGroupId'"}
    if($filterCombinedGroupId){$additionalFilters += " and anthesisgroup_UGSync/combinedGroupId eq '$filterCombinedGroupId'"}
    if($filterSharedMailboxId){$additionalFilters += " and anthesisgroup_UGSync/sharedMailboxId eq '$filterSharedMailboxId'"}
    if($filterMasterMembership){$additionalFilters += " and anthesisgroup_UGSync/masterMembershipList eq '$filterMasterMembership'"}
    if($filterClassifcation){$additionalFilters += " and anthesisgroup_UGSync/classification eq '$filterClassifcation'"}
    if($filterPrivacy){$additionalFilters += " and anthesisgroup_UGSync/privacy eq '$filterPrivacy'"}

    if($selectAllProperties){$lotsOfProperties = ",deletedDateTime,classification,createdDateTime,creationOptions,groupTypes,isAssignableToRole,mailEnabled,mailNickname,onPremisesDomainName,onPremisesLastSyncDateTime,onPremisesNetBiosName,onPremisesSamAccountName,onPremisesSecurityIdentifier,onPremisesSyncEnabled,preferredDataLocation,proxyAddresses,renewedDateTime,resourceBehaviorOptions,resourceProvisioningOptions,securityEnabled,securityIdentifier,visibility,onPremisesProvisioningErrors"}
    invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/groups?`$filter=anthesisgroup_UGSync/extensionType eq 'UGSync'$additionalFilters&`$select=displayName,id,description,mail,anthesisgroup_UGSync$lotsOfProperties"
            
    }
function get-graphList(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true,ParameterSetName = "IdAndId")]
            [parameter(Mandatory = $true,ParameterSetName = "URLAndId")]
            [parameter(Mandatory = $true,ParameterSetName = "IdAndName")]
            [parameter(Mandatory = $true,ParameterSetName = "URLAndName")]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true,ParameterSetName = "IdAndId")]
            [parameter(Mandatory = $true,ParameterSetName = "IdAndName")]
            [string]$graphSiteId
        ,[parameter(Mandatory = $true,ParameterSetName = "URLAndId")]
            [parameter(Mandatory = $true,ParameterSetName = "URLAndName")]
            [string]$serverRelativeSiteUrl
        ,[parameter(Mandatory = $true,ParameterSetName = "URLAndId")]
            [parameter(Mandatory = $true,ParameterSetName = "IdAndId")]
            [string]$listId
        ,[parameter(Mandatory = $true,ParameterSetName = "URLAndName")]
            [parameter(Mandatory = $true,ParameterSetName = "IdAndName")]
            [string]$listName
        )

    switch ($PsCmdlet.ParameterSetName){
        {$_ -match "URL"} { #If we've got a URL to the Site, we'll need to get the Id
            Write-Verbose "get-graphList | Getting Site from URL [$serverRelativeSiteUrl]"
            $graphSiteId = $(get-graphSite -tokenResponse $tokenResponse -serverRelativeUrl $serverRelativeSiteUrl).Id
            }
        {$_ -match "AndName"} { #If we've got a URL to the Site, we'll need to get the Id
            $filter = "?`$filter= displayName eq '$listName'"
            Write-Verbose "get-graphList | Filter set to [$filter]"
            }
        {$_ -match "AndId"} { #If we've got a URL to the Site, we'll need to get the Id
            $listId = "/$listId"
            Write-Verbose "get-graphList | ListId [$listId]"
            }
        }

    invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/sites/$graphSiteId/lists$ListId$filter" -Verbose:$VerbosePreference

    }
function get-graphListItems(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true,ParameterSetName = "IdAndId")]
            [parameter(Mandatory = $true,ParameterSetName = "URLAndId")]
            [parameter(Mandatory = $true,ParameterSetName = "IdAndName")]
            [parameter(Mandatory = $true,ParameterSetName = "URLAndName")]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true,ParameterSetName = "IdAndId")]
            [parameter(Mandatory = $true,ParameterSetName = "IdAndName")]
            [string]$graphSiteId
        ,[parameter(Mandatory = $true,ParameterSetName = "URLAndId")]
            [parameter(Mandatory = $true,ParameterSetName = "URLAndName")]
            [string]$serverRelativeSiteUrl
        ,[parameter(Mandatory = $true,ParameterSetName = "URLAndId")]
            [parameter(Mandatory = $true,ParameterSetName = "IdAndId")]
            [string]$listId
        ,[parameter(Mandatory = $true,ParameterSetName = "URLAndName")]
            [parameter(Mandatory = $true,ParameterSetName = "IdAndName")]
            [string]$listName
        ,[parameter(Mandatory = $false,ParameterSetName = "IdAndId")]
            [parameter(Mandatory = $false,ParameterSetName = "URLAndId")]
            [parameter(Mandatory = $false,ParameterSetName = "IdAndName")]
            [parameter(Mandatory = $false,ParameterSetName = "URLAndName")]
            [switch]$expandAllFields
        ,[parameter(Mandatory = $false,ParameterSetName = "IdAndId")]
            [parameter(Mandatory = $false,ParameterSetName = "URLAndId")]
            [parameter(Mandatory = $false,ParameterSetName = "IdAndName")]
            [parameter(Mandatory = $false,ParameterSetName = "URLAndName")]
            [string]$filterId
        )

    switch ($PsCmdlet.ParameterSetName){
        {$_ -match "URL"} { #If we've got a URL to the Site, we'll need to get the Id
            Write-Verbose "get-graphListItems | Getting Site from URL [$serverRelativeSiteUrl]"
            $graphSiteId = $(get-graphSite -tokenResponse $tokenResponse -serverRelativeUrl $serverRelativeSiteUrl).Id
            }
        {$_ -match "AndName"} { #If we've got a URL to the Site, we'll need to get the Id
            $listId = $(get-graphList -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listName $listName).Id 
            Write-Verbose "get-graphListItems | getting ListId from name [$listName]"
            }
        }

    #Special case for Filter-by-Id as it expicitly requests a single result
    if($filterId){
        Write-Verbose "get-graphListItems | Requesting ListItem by Id [$filterId]"
        invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/sites/$graphSiteId/lists/$listId/items/$filterId" -Verbose:$VerbosePreference
        return
        }


    #Otherwise build up the filter & select conditions into a refiner we can append to the query string
    if($expandAllFields){$expand += " and fields"}

    $refiner = "?"+$select
    if(![string]::IsNullOrWhiteSpace($expand)){
        if($expand.StartsWith(" and ")){$expand = $expand.Substring(5,$expand.Length-5)}
        $expand = "`$expand=$expand"
        }
    if($expand){
        if($refiner.Length -gt 1){$refiner = $refiner+"&"} #If there is already another query option in the refiner, use the '&' symbol to concatenate the the strings
        $refiner = $refiner+$expand
        }

    invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/sites/$graphSiteId/lists/$listId/items$refiner" -Verbose:$VerbosePreference

    }
function get-graphSite(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true,ParameterSetName = "IdLonger")]
            [parameter(Mandatory = $true,ParameterSetName = "URLLonger")]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true,ParameterSetName = "IdLonger")]
            [string]$graphSiteId
        ,[parameter(Mandatory = $true,ParameterSetName = "URLLonger")]
            [string]$serverRelativeUrl
        )

    switch ($PsCmdlet.ParameterSetName){
        "URLLonger" { 
            $sanitisedServerRelativeUrl = $serverRelativeUrl.Replace("https://","").Replace("anthesisllc.sharepoint.com","").Replace(":","").Replace("//","/")
            if($sanitisedServerRelativeUrl.Substring(0,1) -ne "/"){$sanitisedServerRelativeUrl = "/" + $sanitisedServerRelativeUrl}
            Write-Verbose "get-graphSite | Getting Site from URL [$sanitisedServerRelativeUrl][$serverRelativeUrl]"
            $result = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/sites/anthesisllc.sharepoint.com:$sanitisedServerRelativeUrl" -Verbose:$VerbosePreference
            }
        "IdLonger" { #If we're working with $siteUrl, we'll need to get $siteGraphId (which is more of a faff)
            Write-Verbose "get-graphSite | Getting SiteId from URL [$siteUrl]"
            $result = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/sites/$graphSiteId"  -Verbose:$VerbosePreference
            }       
        }

    $result
    }
function get-graphteamsitedetails(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$siteurl
        )
    write-host "Siteurl: $($siteurl)"
    $sitename = ($siteurl -split ".com")[1].Trim("/")
    $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com:/$sitename" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Get
    $response
    write-host "The ID string for $($response.displayname) is: `
    $($response.id)" -ForegroundColor Yellow
    $siteid = $($response.id)
    $DocumentLibraryconfirmation = Read-Host "Would you like to see a list of Document Libraries for $($sitename)? (y/n)"
    If("y" -eq $DocumentLibraryconfirmation){
        $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$($siteid)/drives" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Get
        $response.value
        }
    <#
    .SYNOPSIS
    Find id's of Sharepoint sites and Document Libraries easily to save a few clicks
    #>
    }
function get-graphTokenResponse{
     [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [PSCustomObject]$aadAppCreds
        ,[parameter(Mandatory = $false)]
            [ValidateSet(“client_credentials”,”authorization_code”,"device_code")]
            [string]$grant_type = "client_credentials"
        ,[parameter(Mandatory = $false)]
            [string]$scope = "https://graph.microsoft.com/.default"
        )
    switch($grant_type){
        "authorization_code" {if(!$scope){$scope = "https://graph.microsoft.com/.default"}
            $authCode = get-graphAuthCode -clientID $aadAppCreds.ClientID -redirectUri $aadAppCreds.RedirectUri -scope $scope
            $ReqTokenBody = @{
                Grant_Type    = "authorization_code"
                Scope         = $scope
                client_Id     = $aadAppCreds.ClientID
                Client_Secret = $aadAppCreds.Secret
                redirect_uri  = $aadAppCreds.RedirectUri
                code          = $authCode
                #resource      = "https://graph.microsoft.com"
                }
            }
        "client_credentials" {
            $ReqTokenBody = @{
                Grant_Type    = "client_credentials"
                Scope         = $scope
                client_Id     = $aadAppCreds.ClientID
                Client_Secret = $aadAppCreds.Secret
                }
            }
        "device_code" {
            $tenant = "anthesisllc.onmicrosoft.com"
            $authUrl = "https://login.microsoftonline.com/$tenant"
            $postParams = @{
                resource = "https://graph.microsoft.com/"
                client_id = $aadAppCreds.ClientId
                }
            $response = Invoke-RestMethod -Method POST -Uri "$authurl/oauth2/devicecode" -Body $postParams
            $code = ($response.message -split "code " | Select-Object -Last 1) -split " to authenticate."
            Set-Clipboard -Value $code

            Add-Type -AssemblyName System.Windows.Forms
            $form = New-Object -TypeName System.Windows.Forms.Form -Property @{ Width = 440; Height = 640 }
            $web = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{ Width = 440; Height = 600; Url = "https://www.microsoft.com/devicelogin" }
            $web.Add_DocumentCompleted($DocComp)
            $web.DocumentText
            $form.Controls.Add($web)
            $form.Add_Shown({ $form.Activate() })
            $web.ScriptErrorsSuppressed = $true
            $form.AutoScaleMode = 'Dpi'
            $form.text = "Graph API Authentication"
            $form.ShowIcon = $False
            $form.AutoSizeMode = 'GrowAndShrink'
            $Form.StartPosition = 'CenterScreen'
            $form.ShowDialog() | Out-Null     

            $ReqTokenBody = @{
                grant_type    = "device_code"
                client_Id     = $aadAppCreds.ClientID
                code          = $response.device_code
                }

            }
        }

    $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$($aadAppCreds.TenantId)/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody
    $tokenResponse | Add-Member -MemberType NoteProperty -Name OriginalExpiryTime -Value $((Get-Date).AddSeconds($tokenResponse.expires_in))
    $tokenResponse
    }
function get-graphUsers(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $false)]
            #[ValidateSet("AD","AE","AF","AG","AI","AL","AM","AO","AQ","AR","AS","AT","AU","AW","AX","AZ","BA","BB","BD","BE","BF","BG","BH","BI","BJ","BL","BM","BN","BO","BQ","BR","BS","BT","BV","BW","BY","BZ","CA","CC","CD","CF","CG","CH","CI","CK","CL","CM","CN","CO","CR","CU","CV","CW","CX","CY","CZ","DE","DJ","DK","DM","DO","DZ","EC","EE","EG","EH","ER","ES","ET","FI","FJ","FK","FM","FO","FR","GA","GB","GD","GE","GF","GG","GH","GI","GL","GM","GN","GP","GQ","GR","GS","GT","GU","GW","GY","HK","HM","HN","HR","HT","HU","ID","IE","IL","IM","IN","IO","IQ","IR","IS","IT","JE","JM","JO","JP","KE","KG","KH","KI","KM","KN","KP","KR","KW","KY","KZ","LA","LB","LC","LI","LK","LR","LS","LT","LU","LV","LY","MA","MC","MD","ME","MF","MG","MH","MK","ML","MM","MN","MO","MP","MQ","MR","MS","MT","MU","MV","MW","MX","MY","MZ","NA","NC","NE","NF","NG","NI","NL","NO","NP","NR","NU","NZ","OM","PA","PE","PF","PG","PH","PK","PL","PM","PN","PR","PS","PT","PW","PY","QA","RE","RO","RS","RU","RW","SA","SB","SC","SD","SE","SG","SH","SI","SJ","SK","SL","SM","SN","SO","SR","SS","ST","SV","SX","SY","SZ","TC","TD","TF","TG","TH","TJ","TK","TL","TM","TN","TO","TR","TT","TV","TW","TZ","UA","UG","UM","US","UY","UZ","VA","VC","VE","VG","VI","VN","VU","WF","WS","YE","YT","ZA","ZM","ZW")]
            [ValidateSet("AD","AE","CH","CN","CO","DE","ES","FI","FR","GB","IE","IT","KR","LK","PH","SE","US")]
            [string]$filterUsageLocation
        ,[parameter(Mandatory = $false)]
            [ValidatePattern("@")]
            [string]$filterUpn
        ,[parameter(Mandatory = $false)]
            [hashtable]$filterCustomEq = @{}
        ,[parameter(Mandatory = $false)]
            [switch]$filterLicensedUsers = $false
        ,[parameter(Mandatory = $false)]
            [string[]]$selectCustomProperties
        ,[parameter(Mandatory = $false)]
            [switch]$selectAllProperties = $false
        )

    #We need the GroupId, so if we were only given the UPN, we need to find the Id from that.
    if($filterUsageLocation){
        $filter += " and usageLocation eq '$filterUsageLocation'"
        }
    if($filterUpn){
        $filter += " and userPrincipalName eq '$filterUpn'"
        }
    $filterCustomEq.Keys | % {
        $filter += " and $_ eq '$($filterCustomEq[$_])'"
        }

    if($filterLicensedUsers){
        $select = ",id,displayName,jobTitle,mail,userPrincipalName,usageLocation,assignedLicenses,companyName,country,department,anthesisgroup_employeeInfo"
        }
    if($selectAllProperties){
        $select = ",anthesisgroup_employeeInfo,accountEnabled,assignedLicenses,assignedPlans,businessPhones,city,companyName,country,createdDateTime,creationType,deletedDateTime,department,displayName,employeeId,faxNumber,givenName,id,identities,imAddresses,isResourceAccount,jobTitle,lastPasswordChangeDateTime,legalAgeGroupClassification,licenseAssignmentStates,mail,mailNickname,mobilePhone,officeLocation,onPremisesDistinguishedName,onPremisesDomainName,onPremisesExtensionAttributes,onPremisesImmutableId,onPremisesLastSyncDateTime,onPremisesProvisioningErrors,onPremisesSamAccountName,onPremisesSecurityIdentifier,onPremisesSyncEnabled,onPremisesUserPrincipalName,otherMails,passwordPolicies,passwordProfile,postalCode,preferredDataLocation,preferredLanguage,provisionedPlans,proxyAddresses,refreshTokensValidFromDateTime,showInAddressList,signInSessionsValidFromDateTime,state,streetAddress,surname,usageLocation,userPrincipalName,userType"
        } #Not Implemented yet: aboutMe, birthday, hireDate, interests, mailboxSettings, mySite,pastProjects, preferredName,responsibilities,schools, skills 
    $selectCustomProperties | % {
        $select += ",$_"
        }

    #Build the refiner based on the parameters supplied
    if(![string]::IsNullOrWhiteSpace($select)){
        if($select.StartsWith(",")){$select = $select.Substring(1,$select.Length-1)}
        $select = "`$select=$select"
        }
    if(![string]::IsNullOrWhiteSpace($filter)){
        if($filter.StartsWith(" and ")){$filter = $filter.Substring(5,$filter.Length-5)}
        $filter = "`$filter=$filter"
        }

    $refiner = "?"+$select
    if($filter){
        if($refiner.Length -gt 1){$refiner = $refiner+"&"} #If there is already another query option in the refiner, use the '&' symbol to concatenate the the strings
        $refiner = $refiner+$filter
        }

    Write-Verbose "Graph Query = [users$refiner]"
    try{
        $allUsers = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "users$refiner" -Verbose:$VerbosePreference
        }
    catch{
        Write-Error "Error retrieving Graph Users in get-graphUsers() using query [users$refiner]"
        Throw $_ #Terminate on this error
        }
    
    if($filterLicensedUsers){
        Write-Verbose "Returning all Licensed Users"
        $allUsers | ? {$_.assignedLicenses.Count -gt 0} | Sort-Object userPrincipalName -Unique
        }
    else{
        Write-Verbose "Returning all Users"
        $allUsers | Sort-Object userPrincipalName -Unique
        }
    }
function get-graphUsersFromGroup(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true,ParameterSetName = "groupId")]
            [parameter(Mandatory = $true,ParameterSetName = "groupUpn")]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true,ParameterSetName = "groupId")]
            [string]$groupId
        ,[parameter(Mandatory = $true,ParameterSetName = "groupUpn")]
            [ValidatePattern("@")]
            [string]$groupUpn
        ,[parameter(Mandatory = $true,ParameterSetName = "groupId")]
            [parameter(Mandatory = $true,ParameterSetName = "groupUpn")]
            [ValidateSet("Members","TransitiveMembers","Owners")]
            [string]$memberType 
        ,[parameter(Mandatory = $false,ParameterSetName = "groupId")]
            [parameter(Mandatory = $false,ParameterSetName = "groupUpn")]
            [switch]$returnOnlyUsers = $false
        ,[parameter(Mandatory = $false,ParameterSetName = "groupId")]
            [parameter(Mandatory = $false,ParameterSetName = "groupUpn")]
            [switch]$returnOnlyLicensedUsers = $false
        )

    #We need the GroupId, so if we were only given the UPN, we need to find the Id from that.
    switch ($PsCmdlet.ParameterSetName){
        “groupUpn”  {
            Write-Verbose "We've been given a GroupUPN, so we need the GroupId"
            $graphGroup = get-graphGroups -tokenResponse $tokenResponse -filterUpn $groupUpn -Verbose:$VerbosePreference
            if(!$graphGroup){
                Write-Error "Could not retrieve Graph Group using UPN [$groupUpn]. Check the UPN is valid and try again."
                break
                }
            $groupId = $graphGroup.id
            Write-Verbose "[$groupUpn] Id is [$groupId]"
            }
        }
    if($returnOnlyLicensedUsers){
        $refiner = "?`$select=id,displayName,jobTitle,mail,userPrincipalName,usageLocation,assignedLicenses,anthesisgroup_employeeInfo"
        $returnOnlyUsers = $true #Licensed Users are a subset of Users, so $returnOnlyUsers = $true is implied if $returnOnlyLicensedUsers = $true
        }
    Write-Verbose "Graph Query = [groups/$($groupId)/$($memberType+$refiner)]"
    try{
        $allMembers = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "groups/$groupId/$memberType$refiner" -Verbose:$VerbosePreference
        }
    catch{
        Write-Error "Error retrieving Graph Group $memberType in get-graphUsersFromGroup() using query [groups/$($groupId)/$($memberType+$refiner)]"
        Throw $_ #Terminate on this error
        }
    
    if($returnOnlyUsers){
        if($returnOnlyLicensedUsers){
            Write-Verbose "Returning all Licensed Users"
            $allLicensedUsers = $allMembers | ? {$_.'@odata.type' -eq "#microsoft.graph.user" -and $_.assignedLicenses.Count -gt 0} | Sort-Object userPrincipalName -Unique
            $allLicensedUsers
            }
        else{
            Write-Verbose "Returning all Users"
            $allUsers = $allMembers | ? {$_.'@odata.type' -eq "#microsoft.graph.user"} | Sort-Object userPrincipalName -Unique
            $allUsers
            }
        }
    else{
        Write-Verbose "Returning all Members"
        $allMembers
        }
    }
function get-graphUsersWithEmployeeInfoExtensions(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true,ParameterSetName = "ambiguous")]
            [parameter(Mandatory = $true,ParameterSetName = "explicitUpn")]
            [parameter(Mandatory = $true,ParameterSetName = "explicitId")]
            [parameter(Mandatory = $true,ParameterSetName = "explicitDisplayName")]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true,ParameterSetName = "ambiguous")]
            [switch]$filterNone
        ,[parameter(Mandatory = $true,ParameterSetName = "explicitUpn")]
            [ValidatePattern("@")]
            [string]$filterUpn
        ,[parameter(Mandatory = $true,ParameterSetName = "explicitId")]
            [string]$filterId
        ,[parameter(Mandatory = $true,ParameterSetName = "explicitDisplayName")]
            [string]$filterDisplayName
        ,[parameter(Mandatory = $false,ParameterSetName = "ambiguous")]
            [parameter(Mandatory = $false,ParameterSetName = "explicitUpn")]
            [parameter(Mandatory = $false,ParameterSetName = "explicitId")]
            [parameter(Mandatory = $false,ParameterSetName = "explicitDisplayName")]
            [ValidateSet("Employee","Subcontractor","Associate")]
            [string]$filterContractType
        ,[parameter(Mandatory = $false,ParameterSetName = "ambiguous")]
            [parameter(Mandatory = $false,ParameterSetName = "explicitUpn")]
            [parameter(Mandatory = $false,ParameterSetName = "explicitId")]
            [parameter(Mandatory = $false,ParameterSetName = "explicitDisplayName")]
            [switch]$selectAllProperties
        )
    #Add $filters for the various properties
    $customFilter = @{
        "anthesisgroup_employeeInfo/extensionType" = "employeeInfo"
        }
    if($filterContractType){
        $customFilter.Add("anthesisgroup_employeeInfo/contractType",$filterContractType)
        }

    switch ($PsCmdlet.ParameterSetName){
        "ambiguous"           {get-graphUsers -tokenResponse $tokenResponse -filterCustomEq $customFilter -selectCustomProperties @("anthesisgroup_employeeInfo") -selectAllProperties:$selectAllProperties -filterLicensedUsers -Verbose:$VerbosePreference}
        "explicitUpn"         {get-graphUsers -tokenResponse $tokenResponse -filterCustomEq $customFilter -selectCustomProperties @("anthesisgroup_employeeInfo") -selectAllProperties:$selectAllProperties -filterLicensedUsers -Verbose:$VerbosePreference -filterUpn $filterUpn}
        "explicitDisplayName" {
            $customFilter.Add("displayName",$filterDisplayName)
            get-graphUsers -tokenResponse $tokenResponse -filterCustomEq $customFilter -selectCustomProperties @("anthesisgroup_employeeInfo") -selectAllProperties:$selectAllProperties -filterLicensedUsers -Verbose:$VerbosePreference
            }
        "explicitId"          {
            $customFilter.Add("id",$filterId)
            get-graphUsers -tokenResponse $tokenResponse -filterCustomEq $customFilter -selectCustomProperties @("anthesisgroup_employeeInfo") -selectAllProperties:$selectAllProperties -filterLicensedUsers -Verbose:$VerbosePreference
            }
        }

    
            
    }
function grant-graphSharing(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true,ParameterSetName = "withInvitation")]
            [parameter(Mandatory = $true,ParameterSetName = "withoutInvitation")]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true,ParameterSetName = "withInvitation")]
            [parameter(Mandatory = $true,ParameterSetName = "withoutInvitation")]
            [string]$driveId
        ,[parameter(Mandatory = $true,ParameterSetName = "withInvitation")]
            [parameter(Mandatory = $true,ParameterSetName = "withoutInvitation")]
            [string]$itemId
        ,[parameter(Mandatory = $true,ParameterSetName = "withInvitation")]
            [parameter(Mandatory = $true,ParameterSetName = "withoutInvitation")]
            [array]$sharingRecipientsUpns
        ,[parameter(Mandatory = $true,ParameterSetName = "withInvitation")]
            [parameter(Mandatory = $true,ParameterSetName = "withoutInvitation")]
            [bool]$requireSignIn = $true
        ,[parameter(Mandatory = $true,ParameterSetName = "withInvitation")]
            [parameter(Mandatory = $true,ParameterSetName = "withoutInvitation")]
            [bool]$sendInvitation
        ,[parameter(Mandatory = $true,ParameterSetName = "withInvitation")]
            [parameter(Mandatory = $false,ParameterSetName = "withoutInvitation")]
            [string]$sharingMessage = $false
        ,[parameter(Mandatory = $true,ParameterSetName = "withInvitation")]
            [parameter(Mandatory = $true,ParameterSetName = "withoutInvitation")]
            [ValidateSet("Read","Write")]
            [string]$role
        )
    <#--$formattedRecipients = @{}
    $sharingRecipientsUpns | % {
        $formattedRecipients.Add("email",$_)
        }--#>
    $formattedRecipients = @()
    $sharingRecipientsUpns | % {
        $formattedRecipients += @{"email"=$_}
        }

    $graphParams =@{
        "requireSignIn"=$requireSignIn
        "sendInvitation"=$sendInvitation
        "roles"=@($role)
        "recipients"=$formattedRecipients
        "message"=$sharingMessage
        }
    invoke-graphPost -tokenResponse $tokenResponse -graphQuery "/drives/$driveId/items/$itemId/invite" -graphBodyHashtable $graphParams
    }
function invoke-graphDelete(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$graphQuery
        ,[parameter(Mandatory = $false)]
            [string]$graphBodyHashtable
        )
    $sanitisedGraphQuery = $graphQuery.Trim("/")
    Write-Verbose "https://graph.microsoft.com/v1.0/$sanitisedGraphQuery"

    if($graphBodyHashtable){
        $graphBodyJson = ConvertTo-Json -InputObject $graphBodyHashtable
        Write-Verbose $graphBodyJson
        $graphBodyJsonEncoded = [System.Text.Encoding]::UTF8.GetBytes($graphBodyJson)
        $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/$sanitisedGraphQuery" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method DELETE -Body $graphBodyJsonEncoded
        }
    else{
        $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/$sanitisedGraphQuery" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method DELETE
        }
    if($response.value){$response.value}
    else{$response}
    }
function invoke-graphGet(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$graphQuery
        ,[parameter(Mandatory = $false)]
            [switch]$firstPageOnly
        ,[parameter(Mandatory = $false)]
            [switch]$returnEntireResponse
        )
    $sanitisedGraphQuery = $graphQuery.Trim("/")
    do{
        Write-Verbose "https://graph.microsoft.com/v1.0/$sanitisedGraphQuery"
        $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/$sanitisedGraphQuery" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method GET -Verbose:$VerbosePreference
        if($response.value){
            $results += $response.value
            Write-Verbose "[$([int]$response.value.count)] results returned on this cycle, [$([int]$results.count)] in total"
            }
        elseif([string]::IsNullOrWhiteSpace($response) `
            -or ($response.'@odata.context' -and [string]::IsNullOrWhiteSpace($response.value) -and [string]::IsNullOrWhiteSpace($response.id))){ #If $response is $null, or if we get a response with a $null value
            Write-Verbose "[0] results returned on this cycle, [$([int]$results.count)] in total"
            }
        else{
            $results += $response
            Write-Verbose "[1] results returned on this cycle, [$([int]$results.count)] in total"
            }
        
        if($firstPageOnly){break}
        if(![string]::IsNullOrWhiteSpace($response.'@odata.nextLink')){$sanitisedGraphQuery = $response.'@odata.nextLink'.Replace("https://graph.microsoft.com/v1.0/","")}
        }
    #while($response.value.count -gt 0)
    while($response.'@odata.nextLink')
    if($returnEntireResponse){$response}
    else{$results}
    }
function invoke-graphPatch(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$graphQuery
        ,[parameter(Mandatory = $true)]
            [Hashtable]$graphBodyHashtable
        )

    $sanitisedGraphQuery = $graphQuery.Trim("/")
    Write-Verbose "https://graph.microsoft.com/v1.0/$sanitisedGraphQuery"
        
    $graphBodyJson = ConvertTo-Json -InputObject $graphBodyHashtable
    Write-Verbose $graphBodyJson
    $graphBodyJsonEncoded = [System.Text.Encoding]::UTF8.GetBytes($graphBodyJson)
    
    Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/$sanitisedGraphQuery" -Body $graphBodyJsonEncoded -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Patch
    }
function invoke-graphPost(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$graphQuery
        ,[parameter(Mandatory = $true)]
            [Hashtable]$graphBodyHashtable
        )

    $sanitisedGraphQuery = $graphQuery.Trim("/")
    Write-Verbose "https://graph.microsoft.com/v1.0/$sanitisedGraphQuery"
        
    $graphBodyJson = ConvertTo-Json -InputObject $graphBodyHashtable -Depth 10
    Write-Verbose $graphBodyJson
    $graphBodyJsonEncoded = [System.Text.Encoding]::UTF8.GetBytes($graphBodyJson)
    
    Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/$sanitisedGraphQuery" -Body $graphBodyJsonEncoded -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Post
    }
function invoke-graphPut(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true,ParameterSetName = "BinaryFileStream")]
            [parameter(Mandatory = $true,ParameterSetName = "NormalRequest")]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true,ParameterSetName = "BinaryFileStream")]
            [parameter(Mandatory = $true,ParameterSetName = "NormalRequest")]
            [string]$graphQuery
        ,[parameter(Mandatory = $true,ParameterSetName = "BinaryFileStream")]
            $binaryFileStream
        ,[parameter(Mandatory = $true,ParameterSetName = "NormalRequest")]
            [Hashtable]$graphBodyHashtable
        )

    $sanitisedGraphQuery = $graphQuery.Trim("/")
    Write-Verbose "https://graph.microsoft.com/v1.0/$sanitisedGraphQuery"
        
    if($binaryFileStream){
        $contentType = "text/plain"
        $bodyData = $binaryFileStream
        }
    elseif($graphBodyHashtable){
        $contentType = "application/json; charset=utf-8"
        $graphBodyJson = ConvertTo-Json -InputObject $graphBodyHashtable
        Write-Verbose $graphBodyJson
        $bodyData = [System.Text.Encoding]::UTF8.GetBytes($graphBodyJson)
        }

    Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/$sanitisedGraphQuery" -Body $bodyData -ContentType $contentType -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Put
    }
function new-graphListItem(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true,ParameterSetName = "IdAndId")]
            [parameter(Mandatory = $true,ParameterSetName = "URLAndId")]
            [parameter(Mandatory = $true,ParameterSetName = "IdAndName")]
            [parameter(Mandatory = $true,ParameterSetName = "URLAndName")]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true,ParameterSetName = "IdAndId")]
            [parameter(Mandatory = $true,ParameterSetName = "IdAndName")]
            [string]$graphSiteId
        ,[parameter(Mandatory = $true,ParameterSetName = "URLAndId")]
            [parameter(Mandatory = $true,ParameterSetName = "URLAndName")]
            [string]$serverRelativeSiteUrl
        ,[parameter(Mandatory = $true,ParameterSetName = "URLAndId")]
            [parameter(Mandatory = $true,ParameterSetName = "IdAndId")]
            [string]$listId
        ,[parameter(Mandatory = $true,ParameterSetName = "URLAndName")]
            [parameter(Mandatory = $true,ParameterSetName = "IdAndName")]
            [string]$listName
        ,[parameter(Mandatory = $true,ParameterSetName = "IdAndId")]
            [parameter(Mandatory = $true,ParameterSetName = "URLAndId")]
            [parameter(Mandatory = $true,ParameterSetName = "IdAndName")]
            [parameter(Mandatory = $true,ParameterSetName = "URLAndName")]
            [hashtable]$listItemFieldValuesHash
        )
    switch ($PsCmdlet.ParameterSetName){
        {$_ -match "URLAnd"} {
            Write-Verbose "new-graphListItem | Getting SiteId"
            $graphSiteId = $(get-graphSite -tokenResponse $tokenResponse -serverRelativeUrl $serverRelativeSiteUrl -Verbose:$VerbosePreference).id
            }
        {$_ -match "AndName"} {
            Write-Verbose "new-graphListItem | Getting ListId"
            $listId = $(get-graphList -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listName $listName -Verbose:$VerbosePreference).id
            }
        }
    $graphBodyHash = @{"fields"=$listItemFieldValuesHash}
    Write-Verbose "new-graphListItem | $(stringify-hashTable $listItemFieldValuesHash)"
    invoke-graphPost -tokenResponse $tokenResponse -graphQuery "/sites/$graphSiteId/lists/$listId/items" -graphBodyHashtable $graphBodyHash -Verbose:$VerbosePreference
    }
function new-graphTeam(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$groupId
        ,[parameter(Mandatory = $true)]
            [bool]$allowMemberCreateUpdateChannels
        ,[parameter(Mandatory = $true)]
            [bool]$allowMemberDeleteChannels
        ,[parameter(Mandatory = $false)]
            [bool]$allowGuestCreateUpdateChannels = $false
        ,[parameter(Mandatory = $false)]
            [bool]$allowGuestDeleteChannels = $false
        )
    try{$prexistingTeam = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/teams/$groupId"}
    catch{<#--Meh.--#>}
    if($prexistingTeam){
        Write-Verbose "Pre-existing Team found [$($prexistingTeam.DisplayName)][$($prexistingTeam.id)]"
        $prexistingTeam
        return #If the Team already exists, just return it
        }

    $memberSettings = @{
        "allowCreateUpdateChannels"=$allowMemberCreateUpdateChannels
        "allowDeleteChannels"=$allowMemberDeleteChannels
        }
    $guestSettings = @{
        "allowCreateUpdateChannels"=$allowGuestCreateUpdateChannels
        "allowDeleteChannels"=$allowGuestDeleteChannels
        }
    $newTeamBody = @{
        "memberSettings"=$memberSettings
        "guestSettings"=$guestSettings
        }
    try{$attempt = invoke-graphPut -tokenResponse $tokenResponse -graphQuery "groups/$groupId/team" -graphBodyHashtable $newTeamBody}
    catch{
        #https://docs.microsoft.com/en-us/graph/api/team-put-teams?view=graph-rest-1.0&tabs=http
        #If the group was created less than 15 minutes ago, it's possible for the Create team call to fail with a 404 error code due to replication delays. The recommended pattern is to retry the Create team call three times, with a 10 second delay between calls.
        Start-Sleep -Seconds 10
        try{$attempt = invoke-graphPut -tokenResponse $tokenResponse -graphQuery "groups/$groupId/team" -graphBodyHashtable $newTeamBody}
        catch{
            Start-Sleep -Seconds 10
            try{$attempt = invoke-graphPut -tokenResponse $tokenResponse -graphQuery "groups/$groupId/team" -graphBodyHashtable $newTeamBody}
            catch{
                Write-error "Failed to add Team component to Group [$groupId] after 3 attempts"
                return
                }
            }
        }
    $attempt
    }
function remove-graphUsersFromGroup(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true,ParameterSetName="UserIds")]
            [parameter(Mandatory = $true,ParameterSetName="UserUpns")]
            [psobject]$tokenResponse
        ,[parameter(Mandatory = $true,ParameterSetName = "UserIds")]
            [parameter(Mandatory = $true,ParameterSetName = "UserUpns")]
            [string]$graphGroupId
        ,[parameter(Mandatory = $true,ParameterSetName = "UserIds")]
            [parameter(Mandatory = $true,ParameterSetName = "UserUpns")]
            [ValidateSet("Members","Owners")]
            [string]$memberType 
        ,[parameter(Mandatory = $true,ParameterSetName = "UserUpns")]
            [string[]]$graphUserUpns
        ,[parameter(Mandatory = $true,ParameterSetName = "UserIds")]
            [string[]]$graphUserIds
        )
    
    switch ($PsCmdlet.ParameterSetName){
        "UserUpns" {
            $graphUserUpns | % {
                [array]$graphUserIds += $(get-graphUsers -tokenResponse $tokenResponse -filterUpn $_).id
                }
            
            } 
        }

    $graphUserIds | % {
        #$bodyHash = @{"@odata.id"="https://graph.microsoft.com/v1.0/users/$_"}
        invoke-graphDelete -tokenResponse $tokenResponse -graphQuery "/groups/$graphGroupId/$memberType/$_/`$ref" -Verbose:$VerbosePreference
        }
    }
function repair-graphGroupUGSyncSchemaExtensions(){
    param(
        [parameter(Mandatory = $true)]
            [PSCustomObject]$tokenResponse
        ,[parameter(Mandatory = $true)]
            [PSCustomObject]$graphGroup
        ,[Parameter(Mandatory=$true)]
            [ValidateSet ("Internal","Confidential","External","Sym")]
            [string]$groupClassifcation
        ,[Parameter(Mandatory=$true)]
            [ValidateSet ("AAD","365")]
            [string]$masterMembership
        ,[Parameter(Mandatory=$false)]
            [switch]$createGroupsIfMissing
        )

    
    $possibleSecurityGroupMatches = get-graphGroups -tokenResponse $tokenResponse -filterDisplayNameStartsWith $graphGroup.DisplayName -filterGroupType MailEnabledSecurity

    $dataManagerSG = @()
    $membersSG = @()
    $combinedSG = @()
    $smb = @()

    $dataManagerSG += $possibleSecurityGroupMatches | ? {$_.DisplayName -match "data managers"}
    $membersSG += $possibleSecurityGroupMatches | ? {$_.DisplayName -match "members"}
    $combinedSG += $possibleSecurityGroupMatches | ? {$_.DisplayName -eq $graphGroup.DisplayName -and $_.ExternalDirectoryObjectId -ne $graphGroup.id}
    $smb += Get-Mailbox -Filter "DisplayName -like `'*$graphGroup.DisplayName*`'"

    switch($groupClassifcation){
        "Internal" {$pubPriv = "Private"}
        "Confidential" {$pubPriv = "Private"}
        "External" {$pubPriv = "Private"}
        "Sym" {$pubPriv = "Public"}
        }


    if($dataManagerSG.Count -ne 1){
        Write-Warning "[$($dataManagerSG.Count)] Potential Data Manager groups identified [$($dataManagerSG.DisplayName -join ",")]. Cannot automatically resolve this problem."
        $bigProblem = $true
        }
    if($membersSG.Count -ne 1){
        Write-Warning "[$($membersSG.Count)] Potential Member groups identified [$($membersSG.DisplayName -join ",")]. Cannot automatically resolve this problem."
        $bigProblem = $true
        }
    if($combinedSG.Count -ne 1){
        Write-Warning "[$($combinedSG.Count)] Potential Combined groups identified [$($combinedSG.DisplayName -join ",")]. Cannot automatically resolve this problem."
        $bigProblem = $true
        }
    if($smb.Count -ne 1){
        Write-Warning "[$($smb.Count)] Potential Shared Mailboxes identified [$($smb.DisplayName -join ",")]. Cannot automatically resolve this problem."
        }

    if($bigProblem){
        if($createGroupsIfMissing){
            Write-Warning "Couldn't automatically identify the required groups to fix this. Will attempt to create missing groups"
            if(!$combinedSG){
                Write-Verbose "`tCreating Combined Security Group [$($graphGroup.DisplayName)]"
                try{
                    $combinedSg = new-mailEnabledSecurityGroup -dgDisplayName $graphGroup.DisplayName -membersUpns $null -hideFromGal $false -blockExternalMail $true -ownersUpns "ITTeamAll@anthesisgroup.com" -description "Mail-enabled Security Group for $($graphGroup.DisplayName)" -WhatIf:$WhatIfPreference
                    }
                catch{Write-Error $_}
                }
            if($combinedSG){#Dont try creating the subgroups if the Combined Group isn't available
                set-graphGroupUGSyncSchemaExtensions -tokenResponse $tokenResponse -groupId $graphGroup.id -combinedGroupId $combinedSg.ExternalDirectoryObjectId 
                if(!$dataManagerSG){ #Create a Managers SG if required
                    Write-Verbose "Creating Data Managers Security Group [$($graphGroup.DisplayName) - Data Managers Subgroup]"
                    try{$dataManagerSG = new-mailEnabledSecurityGroup -dgDisplayName "$($graphGroup.DisplayName) - Data Managers Subgroup" -fixedSuffix " - Data Managers Subgroup" -membersUpns $null -memberOf @($combinedSg.ExternalDirectoryObjectId,$combinedSG[0].ObjectId)-hideFromGal $false -blockExternalMail $true -ownersUpns "ITTeamAll@anthesisgroup.com" -description "Mail-enabled Security Group for $($graphGroup.DisplayName) Data Managers" -WhatIf:$WhatIfPreference -Verbose}
                    catch{Write-Error $_}
                    }
                if($dataManagerSG){set-graphGroupUGSyncSchemaExtensions -tokenResponse $tokenResponse -groupId $graphGroup.id -dataManagerGroupId $dataManagerSG.ExternalDirectoryObjectId}

                if(!$membersSg){ #And create a Members SG if required
                    Write-Verbose "Creating Members Security Group [$($graphGroup.DisplayName) - Members Subgroup]"
                    try{$membersSg = new-mailEnabledSecurityGroup -dgDisplayName "$($graphGroup.DisplayName) - Members Subgroup" -fixedSuffix " - Members Subgroup" -membersUpns $null -memberOf @($combinedSg.ExternalDirectoryObjectId,$combinedSG[0].ObjectId) -hideFromGal $false -blockExternalMail $true -ownersUpns "ITTeamAll@anthesisgroup.com" -description "Mail-enabled Security Group for mirroring membership of $($graphGroup.DisplayName) Unified Group" -WhatIf:$WhatIfPreference -Verbose}
                    catch{Write-Error $_}
                    }
                if($membersSG){set-graphGroupUGSyncSchemaExtensions -tokenResponse $tokenResponse -groupId $graphGroup.id -memberGroupId $membersSG.ExternalDirectoryObjectId}
                
                }
            break
            }
        else{
            Write-Warning "Couldn't automatically identify the required groups to fix this. Will attempt to set remaining CustomAttributes then exit"
            set-graphGroupUGSyncSchemaExtensions -tokenResponse $tokenResponse -groupId $graphGroup.id -masterMembershipList $masterMembership -classification $groupClassifcation -privacy $pubPriv
            Set-UnifiedGroup -Identity $unifiedGroup.ExternalDirectoryObjectId -CustomAttribute6 $masterMembership -CustomAttribute7 $groupType -CustomAttribute8 $pubPriv
            break
            }
        }

    if($smb.Count -eq 1){
        Write-Verbose "set-graphGroupUGSyncSchemaExtensions -tokenResponse $tokenResponse -groupId [$($graphGroup.id)] -sharedMailboxId [$($smb.id)] -dataManagerGroupId [$($dataManagerSG.id)] -memberGroupId [$($membersSG.id)] -combinedGroupId [$($combinedSg.id)] -masterMembershipList [$masterMembership] -classification [$groupClassifcation] -privacy [$pubPriv]"
        set-graphGroupUGSyncSchemaExtensions -tokenResponse $tokenResponse -groupId $graphGroup.id -sharedMailboxId $smb.id -dataManagerGroupId $dataManagerSG.id -memberGroupId $membersSG.id -combinedGroupId $combinedSg.id -masterMembershipList $masterMembership -classification $groupClassifcation -privacy $pubPriv
        }
    else{
        Write-Verbose "set-graphGroupUGSyncSchemaExtensions -tokenResponse $tokenResponse -groupId [$($graphGroup.id)]  -dataManagerGroupId [$($dataManagerSG.id)] -memberGroupId [$($membersSG.id)] -combinedGroupId [$($combinedSg.id)] -masterMembershipList [$masterMembership] -classification [$groupClassifcation] -privacy [$pubPriv]"
        set-graphGroupUGSyncSchemaExtensions -tokenResponse $tokenResponse -groupId $graphGroup.id -dataManagerGroupId $dataManagerSG.id -memberGroupId $membersSG.id -combinedGroupId $combinedSg.id -masterMembershipList $masterMembership -classification $groupClassifcation -privacy $pubPriv
        }
    }
function reset-graphUnifiedGroupSettingsToOriginals(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [PSCustomObject]$tokenResponse
        ,[parameter(Mandatory = $true)]
            [PSCustomObject]$graphGroupExtended
        ,[parameter(Mandatory = $false)]
            [string[]]$itAdminEmailAddresses
        ,[parameter(Mandatory = $false)]
            [switch]$suppressEmailNotification
        )
    #Compare current Unified Group settings against orginal settings and revert
    if([string]::IsNullOrWhiteSpace($graphGroupExtended.classification)){$graphGroupExtended = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -filterId $graphGroupExtended.id -selectAllProperties}
    [hashtable]$current = @{}
    [hashtable]$changes = @{}
    $combinedMesg = get-graphGroups -tokenResponse $tokenResponse -filterId $graphGroupExtended.anthesisgroup_UGSync.combinedGroupId
    if($combinedMesg.displayName -ne $graphGroupExtended.displayName){
        $current.Add("displayName",$graphGroupExtended.displayName)
        $changes.Add("displayName",$combinedMesg.displayName)
        }
    if($graphGroupExtended.anthesisgroup_UGSync.classification -ne $graphGroupExtended.classification){
        $current.Add("classification",$graphGroupExtended.classification)
        $changes.Add("classification",$graphGroupExtended.anthesisgroup_UGSync.classification)
        }
    if($graphGroupExtended.anthesisgroup_UGSync.privacy -ne $graphGroupExtended.visibility){
        $current.Add("visibility",$graphGroupExtended.visibility)
        $changes.Add("visibility",$graphGroupExtended.anthesisgroup_UGSync.privacy)
        }

    if($changes.Count -gt 0){
        Write-Warning "Unexpected changed found on UnifiedGroup [$($graphGroupExtended.displayName)][$($graphGroupExtended.id)]: $(stringify-hashTable $current)"
        if(!$suppressEmailNotification){
            if($itAdminEmailAddresses.Count -lt 1){$itAdminEmailAddresses = $(get-graphAdministrativeRoleMembers -tokenResponse $tokenResponse -roleName 'User Account Administrator').mail}
            $owners = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $graphGroupExtended.id -memberType Owners -returnOnlyUsers
            $groupOwnersFirstNames = $($($owners.givenName | Sort-Object givenName) -join ", ")
            $groupOwnersFirstNames = $groupOwnersFirstNames -replace "(.*),(.*)", "`$1 &`$2"
            $body = "<HTML><FONT FACE=`"Calibri`">Hello $groupOwnersFirstNames,`r`n`r`n<BR><BR>"
            $body += "Sorry, I found some changes to $($combinedMesg.displayName) and I'm rolling them back:`r`n`r`n<BR><BR><UL>"
            $changes.Keys | % {
                $body += "<LI>$_ reverted to [$($changes[$_])] from [$($current[$_])]</LI>"
                }
            $body += "</UL> Our Team names adhere to our <A HREF=`"https://anthesisllc.sharepoint.com/sites/Resources-IT/_layouts/15/DocIdRedir.aspx?ID=HXX7CE52TSD2-1759992947-11`">Naming Conventions</A> to ensure everyone in Anthesis is talking a common language, and we rely on Team Classification and Privacy/Visibilty settings to ensure robust and scalable access to data.`r`n`r`n<BR><BR>"
            $body += "If you think that these settings are wrong, you'll need to speak with one of the humans in the IT Team.`r`n`r`n<BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>The Helpful Teams Robot</FONT></HTML>"
            #Send-MailMessage -From groupbot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Team [$($combinedMesg.displayName)] settings rolled back" -BodyAsHtml $body -To $($owners.mail) -Cc $itAdminEmailAddresses -Encoding UTF8
            Send-MailMessage -From groupbot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Team [$($combinedMesg.displayName)] settings rolled back" -BodyAsHtml $body -To kevin.maitland@anthesisgroup.com  -Encoding UTF8
            }
        #Now, fix the settings:
        invoke-graphPatch -tokenResponse $tokenResponse -graphQuery "/groups/$($graphGroupExtended.id)" -graphBodyHashtable $changes -Verbose:$VerbosePreference
        #And check the Membership settings are correct too:
        set-graphUnifiedGroupGuestSettings -tokenResponse $tokenResponse -graphUnifiedGroupExtended $graphGroupExtended -Verbose:$VerbosePreference
        }
    }
function set-graphGroupSharedMailboxAccess(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true,ParameterSetName = "groupObject")]
            [parameter(Mandatory = $true,ParameterSetName = "groupUpn")]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true,ParameterSetName = "groupObject")]
            [psobject]$graphGroup
        ,[parameter(Mandatory = $true,ParameterSetName = "groupUpn")]
            [ValidatePattern("@")]
            [string]$groupUpn
        ,[parameter(Mandatory = $true,ParameterSetName = "groupObject")]
            [parameter(Mandatory = $true,ParameterSetName = "groupUpn")]
            [pscredential]$exoCreds
        ,[parameter(Mandatory = $false,ParameterSetName = "groupObject")]
            [parameter(Mandatory = $false,ParameterSetName = "groupUpn")]
            [switch]$reconcileFullAccessPermissions
        ,[parameter(Mandatory = $false,ParameterSetName = "groupObject")]
            [parameter(Mandatory = $false,ParameterSetName = "groupUpn")]
            [switch]$reconcileSendAsPermissions
        )

    if(!$reconcileFullAccessPermissions -and !$reconcileSendAsPermissions){
        Write-Warning "Neither `$reconcileFullAccessPermissions nor `$reconcileSendAsPermissions was set. Nothing to process."
        break
        }

    switch ($PsCmdlet.ParameterSetName){
        “groupUpn”  {
            Write-Verbose "We've been given a GroupUPN, so we need the Group object"
            $graphGroup = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -filterUpn $groupUpn -Verbose:$VerbosePreference
            if(!$graphGroup){
                Write-Error "Could not retrieve Graph Group using UPN [$groupUpn]. Check the UPN is valid and try again."
                break
                }
            }
        "groupObject" {
            if($graphGroup.psobject.Properties.Name -notcontains "CustomAttribute1"){
                Write-Verbose "We've been given a Group object, but it's missing the CustomAttributes"
                $graphGroup = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -filterUpn $groupUpn -Verbose:$VerbosePreference
                if(!$graphGroup){
                    Write-Error "Could not retrieve CustomAttributes of Graph Group [$($graphGroup.displayName)][$($graphGroup.id)]. Cannot identify linked Shared Mailbox without these."
                    break
                    }
                }
            }
        }

    if($graphGroup.groupTypes -notcontains "Unified"){
        Write-Error "Graph Group [$($graphGroup.displayName)][$($graphGroup.id)] is not a Unified Group, so has no Shared Mailbox associated with it"
        break
        }
    
    if([string]::IsNullOrWhiteSpace($graphGroup.CustomAttribute5)){
        Write-Error "Graph Group [$($graphGroup.displayName)][$($graphGroup.id)] has no associated Shared Mailbox (CustomAttribute5 is not set). Cannot set Members' permissions."
        break
        }
    
    $sharedMailbox = Get-Mailbox -Identity $graphGroup.CustomAttribute5
    if(!$sharedMailbox){
        Write-Error "Shared Mailbox with Id [$($graphGroup.CustomAttribute5)] for [$($graphGroup.displayName)][$($graphGroup.id)] cannot be retrieved. Check that the mailbox has not been deleted."
        break
        }

    #Get the list of users who *should* have access, and get it via the associated Members Subgroup so that we can get the transitive members
    $usersToSet = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $graphGroup.CustomAttribute3 -memberType TransitiveMembers -returnOnlyLicensedUsers -Verbose:$VerbosePreference
    
    if($reconcileFullAccessPermissions){
        Write-Verbose "Reconciling FullAccess permissions on Shared Mailbox [$($sharedMailbox.DisplayName)][$($sharedMailbox.ExternalDirectoryObjectId)]"
        #Get the current list of permissions
        $mailboxPermissions = Get-MailboxPermission -Identity $sharedMailbox.ExternalDirectoryObjectId | ? {$_.User -match "@"}

        #Compare what's there with what *should* be there
        if([string]::IsNullOrWhiteSpace($mailboxPermissions)){$mailboxPermissions = @()} #Prevent $null handling errors in Compare-Object
        $mailboxPermissions | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name "userPrincipalName" -Value $_.User}
        $comparison = Compare-Object -ReferenceObject $mailboxPermissions -DifferenceObject $usersToSet -Property userPrincipalName
        $comparison | ? {$_.SideIndicator -eq "=>"} | % {
            Write-Verbose "Adding FullAccess permission for [$($_.userPrincipalName)] to [$($sharedMailbox.DisplayName)][$($sharedMailbox.ExternalDirectoryObjectId)]"
            Add-MailboxPermission -Identity $sharedMailbox.ExternalDirectoryObjectId -AccessRights FullAccess -User $_.userPrincipalName 
            }
        $comparison | ? {$_.SideIndicator -eq "<="} | % {
            Write-Verbose "Removing FullAccess permission for [$($_.userPrincipalName)] from [$($sharedMailbox.DisplayName)][$($sharedMailbox.ExternalDirectoryObjectId)]"
            Remove-MailboxPermission -Identity $sharedMailbox.ExternalDirectoryObjectId -AccessRights FullAccess -User $_.userPrincipalName
            }
        }
    if($reconcileSendAsPermissions){
        Write-Verbose "Reconciling SendAs permissions on Shared Mailbox [$($sharedMailbox.DisplayName)][$($sharedMailbox.ExternalDirectoryObjectId)]"
        #Get the current list of permissions
        $recipientPermissions = Get-RecipientPermission -Identity $sharedMailbox.ExternalDirectoryObjectId | ? {$_.Trustee -match "@"}

        #Compare what's there with what *should* be there
        if([string]::IsNullOrWhiteSpace($recipientPermissions)){$recipientPermissions = @()} #Prevent $null handling errors in Compare-Object
        $recipientPermissions | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name "userPrincipalName" -Value $_.Trustee}
        $comparison = Compare-Object -ReferenceObject $recipientPermissions -DifferenceObject $usersToSet -Property userPrincipalName
        $comparison | ? {$_.SideIndicator -eq "=>"} | % {
            Write-Verbose "Adding SendAs permission for [$($_.userPrincipalName)] to [$($sharedMailbox.DisplayName)][$($sharedMailbox.ExternalDirectoryObjectId)]"
            Add-RecipientPermission -Identity $sharedMailbox.ExternalDirectoryObjectId -AccessRights SendAs -Trustee $_.userPrincipalName -Confirm:$false
            }
        $comparison | ? {$_.SideIndicator -eq "<="} | % {
            Write-Verbose "Removing FullAccess permission for [$($_.userPrincipalName)] from [$($sharedMailbox.DisplayName)][$($sharedMailbox.ExternalDirectoryObjectId)]"
            Remove-RecipientPermission -Identity $sharedMailbox.ExternalDirectoryObjectId -AccessRights FullAccess -User $_.userPrincipalName
            }

        #If Users have SendAs permissions, the SharedMailbox needs to be visible in the Global Address List. Otherwise, we should hide it.
        $recipientPermissions = Get-RecipientPermission -Identity $sharedMailbox.ExternalDirectoryObjectId | ? {$_.Trustee -match "@"}
        if([string]::IsNullOrWhiteSpace($recipientPermissions)){
            Write-Verbose "Hiding Shared Mailbox [$($sharedMailbox.DisplayName)][$($sharedMailbox.ExternalDirectoryObjectId)] from the Global Address List"
            Set-Mailbox -Identity $sharedMailbox.ExternalDirectoryObjectId -HiddenFromAddressListsEnabled:$true
            }
        else{
            Write-Verbose "Showing Shared Mailbox [$($sharedMailbox.DisplayName)][$($sharedMailbox.ExternalDirectoryObjectId)] in the Global Address List"
            Set-Mailbox -Identity $sharedMailbox.ExternalDirectoryObjectId -HiddenFromAddressListsEnabled:$false
            }
        }
    }
function set-graphGroupUGSyncSchemaExtensions(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$groupId        
        ,[parameter(Mandatory = $false)]
            [string]$dataManagerGroupId        
        ,[parameter(Mandatory = $false)]
            [string]$memberGroupId        
        ,[parameter(Mandatory = $false)]
            [string]$combinedGroupId        
        ,[parameter(Mandatory = $false)]
            [string]$sharedMailboxId        
        ,[parameter(Mandatory = $false)]
            [string]$masterMembershipList        
        ,[parameter(Mandatory = $false)]
            [string]$classification        
        ,[parameter(Mandatory = $false)]
            [string]$privacy        
        )
    $bodyHash = @{
        "anthesisgroup_UGSync" = @{
            "extensionType" = "UGSync"}
            }
    if($dataManagerGroupId){$bodyHash["anthesisgroup_UGSync"].Add("dataManagerGroupId",$dataManagerGroupId)}
    if($memberGroupId){$bodyHash["anthesisgroup_UGSync"].Add("memberGroupId",$memberGroupId)}
    if($combinedGroupId){$bodyHash["anthesisgroup_UGSync"].Add("combinedGroupId",$combinedGroupId)}
    if($sharedMailboxId){$bodyHash["anthesisgroup_UGSync"].Add("sharedMailboxId",$sharedMailboxId)}
    if($masterMembershipList){$bodyHash["anthesisgroup_UGSync"].Add("masterMembershipList",$masterMembershipList)}
    if($classification){$bodyHash["anthesisgroup_UGSync"].Add("classification",$classification)}
    if($privacy){$bodyHash["anthesisgroup_UGSync"].Add("privacy",$privacy)}

    invoke-graphPatch -tokenResponse $tokenResponse -graphQuery "/groups/$groupId" -graphBodyHashtable $bodyHash
    
    }
function set-graphUnifiedGroupGuestSettings(){
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true,ParameterSetName = "idOnly")]
            [parameter(Mandatory = $true,ParameterSetName = "groupObject")]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true,ParameterSetName = "idOnly")]
            [string]$groupId
        ,[parameter(Mandatory = $true,ParameterSetName = "groupObject")]
            [psobject]$graphUnifiedGroupExtended        
        ,[parameter(Mandatory = $false,ParameterSetName = "idOnly")]
            [parameter(Mandatory = $false,ParameterSetName = "groupObject")]
            [ValidateSet("Internal","External","Sym","Confidential")]
            [string]$classificationOverride
        )

    #If we're not given an override, look up what the Classification of the group should be
    if([string]::IsNullOrWhiteSpace($classificationOverride)){
        switch ($PsCmdlet.ParameterSetName){
            “idOnly”  {
                $graphUnifiedGroupExtended = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -filterId $groupId
                }
            }
        $classificationOverride = $graphUnifiedGroupExtended.anthesisgroup_UGSync.classification
        }

    if($classificationOverride -eq "External"){$allowToAddGuests = "True"} #Weird, but we can't use $true or $false
    else{$allowToAddGuests = "False"}

    $sharingSettings = [ordered]@{
        'name'='AllowToAddGuests'
        'value'=$allowToAddGuests
        }
    $sharingBody = [ordered]@{
        'displayName'='Group.Unified.Guest'
        'templateId' ='08d542b9-071f-4e16-94b0-74abb372e3d9'
        'values'     = @($sharingSettings)
            }

    $existingSettings = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/groups/$($graphUnifiedGroupExtended.id)/settings" -Verbose:$VerbosePreference
    if($existingSettings){
        $existingSettings.values | ? {$_.Name -eq "AllowToAddGuests"} | % { #"/groups/$($graphUnifiedGroupExtended.id)/settings" returns a weird object: the .values property is a 0+ array of [PSCustomObject]
            if($_.value -ne $allowToAddGuests){
                #If the wrong AllowToAddGuests settings are in place, fix them and notify IT
                invoke-graphPatch -tokenResponse $tokenResponse -graphQuery "/groups/$($graphUnifiedGroupExtended.id)/settings/$($existingSettings.id)" -graphBodyHashtable $sharingBody -Verbose:$VerbosePreference
                Write-Warning "AllowToAddGuests changed from [$($_.value)] to [$allowToAddGuests] for Unified Group [$($graphUnifiedGroupExtended.id)][$($graphUnifiedGroupExtended.DisplayName)]"
                Send-MailMessage -Subject "AllowToAddGuests changed from [$($_.value)] to [$sharingSettings] for Unified Group [$($graphUnifiedGroupExtended.id)][$($graphUnifiedGroupExtended.DisplayName)]" -to "kevin.maitland@anthesisgroup.com" -From securitybot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Encoding UTF8 -Priority High
                }
            else{Write-Verbose "AllowToAddGuests are correct for [$($graphUnifiedGroupExtended.id)][$($graphUnifiedGroupExtended.DisplayName)]"}
            
            }
        }
    else{#If there are no AllowToAddGuests settings, just create them
        invoke-graphPost -tokenResponse $tokenResponse -graphQuery "/groups/$($graphUnifiedGroupExtended.id)/settings" -graphBodyHashtable $sharingBody -Verbose:$VerbosePreference
        }

    }
function set-graphuser(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$userIdOrUpn
        ,[parameter(Mandatory = $false)]
            [hashtable]$userPropertyHash = @{}
        ,[parameter(Mandatory = $false)]
            [hashtable]$userEmployeeInfoExtensionHash
        )
      
    $validProperties = @("accountEnabled","assignedLicenses","assignedPlans","businessPhones","city","companyName","country","createdDateTime","creationType","deletedDateTime","department","displayName","employeeId","faxNumber","givenName","id","identities","imAddresses","isResourceAccount","jobTitle","lastPasswordChangeDateTime","legalAgeGroupClassification","licenseAssignmentStates","mail","mailNickname","mobilePhone","officeLocation","onPremisesDistinguishedName","onPremisesDomainName","onPremisesExtensionAttributes","onPremisesImmutableId","onPremisesLastSyncDateTime","onPremisesProvisioningErrors","onPremisesSamAccountName","onPremisesSecurityIdentifier","onPremisesSyncEnabled","onPremisesUserPrincipalName","otherMails","passwordPolicies","passwordProfile","postalCode","preferredDataLocation","preferredLanguage","provisionedPlans","proxyAddresses","refreshTokensValidFromDateTime","showInAddressList","signInSessionsValidFromDateTime","state","streetAddress","surname","usageLocation","userPrincipalName","userType")
    $dubiousProperties = @("aboutMe","birthday","hireDate","interests","mailboxSettings","mySite","pastProjects","preferredName","responsibilities","schools","skills")
    $validExtensionProperties = @("extensionType","businessUnit","employeeId","contractType")

    $duffProperties = @()
    $userPropertyHash.Keys | % { #Check the properties we're going to try and update the User with are valid:
        if($validProperties -notcontains $_ ){
            if($dubiousProperties -notcontains $_){
                $duffProperties += $_
                }
            else{Write-Warning "Property [$_] isn't fully supported and might cause problems"}
            }
        }

    if($userEmployeeInfoExtensionHash){
        $userEmployeeInfoExtensionHash.Keys | % { #Check the properties we're going to try and update the User with are valid:
            if($validExtensionProperties -notcontains $_){
                $duffProperties += "anthesisgroup_employeeInfo/$_"
                }
            }
        #Now add the Extension properties into the main hash in the correct format
        $userPropertyHash.Add("anthesisgroup_employeeInfo",$userEmployeeInfoExtensionHash)
        }

    if($duffProperties.Count -gt 0){
        Write-Error -Message "Property(s) [$($duffProperties -join ", ")] is invalild for Graph User object. Will not attempt to update."
        break
        }
    
    invoke-graphPatch -tokenResponse $tokenResponse -graphQuery "/users/$userIdOrUpn" -graphBodyHashtable $userPropertyHash -Verbose:$VerbosePreference
    }
function test-graphBearerAccessTokenStillValid(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true,ParameterSetName = "TestAndRenew")]
            [parameter(Mandatory = $true,ParameterSetName = "JustTest")]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $false,ParameterSetName = "TestAndRenew")]
            [int]$renewTokenExpiringInSeconds
        ,[parameter(Mandatory = $true,ParameterSetName = "TestAndRenew")]
            [PSCustomObject]$aadAppCreds
        )
    if($tokenResponse.OriginalExpiryTime -ge $(Get-Date).AddSeconds($renewTokenExpiringInSeconds)){$tokenResponse} #If the token  is still valid, just return it
    else{
        if($renewTokenExpiringInSeconds){
            get-graphTokenResponse -aadAppCreds $aadAppCreds -grant_type client_credentials #If it's expired (or will expire within the supplied limit), renew it
            }
        else{$false}#Otherwise return False
        }
    }
function update-mailboxCustomAttibutesToGraphSchemaExtensions(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [psobject]$mailbox
        ,[parameter(Mandatory = $true)]
            [ValidateSet("Employee","Subcontractor","Associate")]
            [string]$contractType
        )
    
    $bodyHash = @{
        "anthesisgroup_employeeInfo" = @{
            "extensionType" = "employeeInfo"
            "contractType" = "Employee"
            "employeeId" = $null
            "businessUnit" = $mailbox.CustomAttribute1
            }
        }
    invoke-graphPatch -tokenResponse $tokenResponse -graphQuery "/users/$($mailbox.ExternalDirectoryObjectId)" -graphBodyHashtable $bodyHash
    }
function update-unifiedGroupCustomAttibutesToGraphSchemaExtensions(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [psobject]$unifiedGroup        
        )
    
    $bodyHash = @{
        "anthesisgroup_UGSync" = @{
            "extensionType" = "UGSync"
            "dataManagerGroupId" = $unifiedGroup.CustomAttribute2
            "memberGroupId" = $unifiedGroup.CustomAttribute3
            "combinedGroupId" = $unifiedGroup.CustomAttribute4
            "sharedMailboxId" = $unifiedGroup.CustomAttribute5
            "masterMembershipList" = $unifiedGroup.CustomAttribute6
            "classification" = $unifiedGroup.CustomAttribute7
            "privacy" = $unifiedGroup.CustomAttribute8
            }
        }
    invoke-graphPatch -tokenResponse $tokenResponse -graphQuery "/groups/$($unifiedGroup.ExternalDirectoryObjectId)" -graphBodyHashtable $bodyHash
    }
