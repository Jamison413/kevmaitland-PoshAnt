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
        'DriveId'     {
            #Test that graphDriveId is valid
            try{$drive = get-graphDrives -tokenResponse $tokenResponse -driveId $graphDriveId}
            catch{
                if($_.Exception -match "(404)"){
                    Write-Error "Drive Object with Id [$($graphDriveId)] does not exist"
                    return
                    }
                }
            }
        'DriveObject' {$graphDriveId = $graphDriveObject.Id}
        }
    Write-Verbose "add-graphArrayOfFoldersToDrive [$($graphDriveId)]"    
    


    #Prep the folders array (in case the user has provided junk like $foldersAndSubfoldersArray = @("Test","test\test2\test3\test4","test","/test/TeSt2\","tEST #3","Test | #4")
    $expandedFoldersAndSubfoldersArray = ,@()
    $foldersAndSubfoldersArray | % {
        $thisFolder = $_.Replace("\","/").Trim("/")
        $expandingFolderPath = ""
        $thisFolder.Split("/") | % {
            $expandingFolderPath += "$(sanitise-forSharePointFolderName $_)/"
            $expandedFoldersAndSubfoldersArray += $expandingFolderPath.Trim("/")
            }
        }

    $driveItemsToReturn = @()
    #Iterate through our sanitised folder array and create the folders
    $expandedFoldersAndSubfoldersArray | Sort-Object -Unique | ? {![string]::IsNullOrWhiteSpace($_)} | % {
        $folderName = Split-Path $_ -Leaf
        if($folderName -eq $_){ #If it is _just_ a folder (i.e. not a subfolder), just create it
            try{
                $newFolder = add-graphFolderToDrive -graphDriveId $graphDriveId -folderName $folderName -tokenResponse $tokenResponse -conflictResolution $conflictResolution -ErrorAction Stop
                $driveItemsToReturn += $newFolder
                }
            catch{
                if($_.Exception -match "(409)" -or $_.InnerException -match "(409)"){ #If the folder already existed, get and return it
                    $existingFolder = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/drives/$graphDriveId/root:/$folderName"
                    $driveItemsToReturn += $existingFolder
                    }
                else{Write-Error $_}
                }
            }
        else{ #If it _is_ a subfolder, we also need to supply the relative path (and invoke-graphGet doesn't like a $null value for -relativePathToFolder)
            try{
                $relativePath = Split-Path $_ -Parent
                $newFolder = add-graphFolderToDrive -graphDriveId $graphDriveId -folderName $folderName -tokenResponse $tokenResponse -conflictResolution $conflictResolution -ErrorAction Stop -relativePathToFolder $([uri]::EscapeDataString($relativePath))
                $driveItemsToReturn += $newFolder
                }
            catch{
                if($_.Exception -match "(409)" -or $_.InnerException -match "(409)"){ #If the folder already existed, get and return it
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
    invoke-graphPost -tokenResponse $tokenResponse -graphQuery $graphQuery -graphBodyHashtable $folderHash
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
            [ValidateSet("Kiosk","E1","E3","E5","EMS","ATP","PowerBIFree","AudioConferencing","DomesticCalling","InternationalCalling","Project","Visio","M3","WinE3")]
            [string]$licenseFriendlyName 
        ,[parameter(Mandatory = $true,ParameterSetName = "Guid")]
            [string]$licenseGuid
        ,[parameter(Mandatory = $false,ParameterSetName = "Guid")]
            [parameter(Mandatory = $false,ParameterSetName = "Friendly")]
            [string[]]$disabledPlansGuids = @()
        ,[parameter(Mandatory = $true,ParameterSetName = "Guids")]
            [string[]]$licenseGuids
        )
    $specialLicenses = @("Kiosk","80b2d799-d2ba-4d2a-8842-fb0d0f3a4b82","E1","18181a46-0d4e-45cd-891e-60aabd171b4e","E3","6fd2c87f-b296-42f0-b197-1e91e994b900","E5","c7df2760-2c81-4ef7-b578-5b5392b571df","M3","05e9a617-0261-4cee-bb44-138d3ef5d965","M5","06ebc4ee-1bb5-47dd-8120-11324bc54e06")
    $licensesToRemove = @()
    if($specialLicenses -contains $licenseFriendlyName -or
        $specialLicenses -contains $licenseGuid 
        ){
        #We have to remove any conflicting licenses at the same time
        #get user licesnses
        $userRecord = get-graphUsers -tokenResponse $tokenResponse -filterUpns $userIdOrUpn -selectAllProperties
        #build appropriate remove hash
        $matchedLicenses = $(Compare-Object -ReferenceObject $userRecord.assignedLicenses.skuId -DifferenceObject $specialLicenses -IncludeEqual -ExcludeDifferent)
        @($matchedLicenses.InputObject | Select-Object) | % {
            $licensesToRemove += $_
            }
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
    @($licenseGuids | Select-Object) | % {
        $thisLicenseDefinition = @{"skuId"=$_}
        $thisLicenseDefinition.Add("disabledPlans",$disabledPlansGuids) #$disabledPlansGuids is $null if $PsCmdlet.ParameterSetName -eq "Guids", so we don't need to worry about which disabledPlans belong to which licenseGuid
        [array]$licenseArray += $thisLicenseDefinition
        }
    
    $graphBodyHashtable = @{
        "addLicenses"=$licenseArray
        "removeLicenses"=$licensesToRemove
        }

    invoke-graphPost -tokenResponse $tokenResponse -graphQuery "/users/$userIdOrUpn/assignLicense" -graphBodyHashtable $graphBodyHashtable 
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
        $bodyHash = @{"@odata.id"="https://graph.microsoft.com/v1.0/directoryObjects/$_"}
        invoke-graphPost -tokenResponse $tokenResponse -graphQuery "/groups/$graphGroupId/$memberType/`$ref" -graphBodyHashtable $bodyHash
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
        invoke-graphDelete -tokenResponse $tokenResponse -graphQuery "/drives/$graphDriveId/items/$graphDriveItemId" -graphBodyHashtable $deleteBody
        }
    else{
        invoke-graphDelete -tokenResponse $tokenResponse -graphQuery "/drives/$graphDriveId/items/$graphDriveItemId" 
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
        invoke-graphDelete -tokenResponse $tokenResponse -graphQuery "sites/$graphSiteId/lists/$graphListId/items/$graphItemId" 
}
function delete-graphCalendarEvent(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$userId
        ,[parameter(Mandatory = $true)]
            [string]$eventId
        )
    Write-Verbose "delete-graphCalendarEvent | $($eventId)"
    invoke-graphDelete -tokenResponse $tokenResponse -graphQuery "/users/$userId/calendar/events/$eventId" -Verbose:$VerbosePreference
}
function delete-graphRecurringCalendarEvent(){
    [cmdletbinding()]
        Param (
        [parameter(Mandatory = $true,ParameterSetName="UserId")]
            [parameter(Mandatory = $true,ParameterSetName="UserUpn")]
            [psobject]$tokenResponse
        ,[parameter(Mandatory = $true,ParameterSetName = "UserId")]
            [parameter(Mandatory = $true,ParameterSetName = "UserUpn")]
            [string]$seriesmasterId
        ,[parameter(Mandatory = $true,ParameterSetName = "UserUpn")]
            [string]$UserUpn
        ,[parameter(Mandatory = $true,ParameterSetName = "UserId")]
            [string]$UserId
        )
    Write-Verbose "Deleting recurring event ID: $($seriesmasterId)"
    invoke-graphDelete -tokenResponse $tokenResponse -graphQuery "/users/$userId/calendar/events/$eventId" -Verbose:$VerbosePreference
}
function delete-graphShiftUserShifts(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$teamId
        ,[parameter(Mandatory = $true)]
            [string]$MsAppActsAsUserId
        ,[parameter(Mandatory = $true)]
            [string]$shiftId
        )
    

    invoke-graphDelete -tokenResponse $tokenResponse -graphQuery "/teams/$teamId/schedule/Shifts/$shiftId" -additionalHeaders @{"MS-APP-ACTS-AS"=$MsAppActsAsUserId} -Verbose:$VerbosePreference
    
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
    $admins += "t0-kevin.maitland@anthesisgroup.com"
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
function get-graphAppClientCredentials{
     [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [ValidateSet("TeamsBot","SchemaBot","IntuneBot","SharePointBot","ShiftBot","ReportBot","SmtpBot")]
            [String]$appName
        )
    
    switch($appName){ #Figure out the name of the file
        "TeamsBot"  {$encryptedCredsFile = "teambotdetails.txt"}
        "SchemaBot" {$encryptedCredsFile = "schemabot.txt"}
        "IntuneBot" {$encryptedCredsFile = "intunebot.txt"}
        "SharePointBot" {$encryptedCredsFile = "spBotDetails.txt"}
        "ShiftBot" {$encryptedCredsFile = "shiftBotDetails.txt"}
        "ReportBot"{$encryptedCredsFile = "ReportBotDetails.txt"}
        "SmtpBot"{$encryptedCredsFile = "SmtpBot.txt"}
        }
    
    $placesToLook = @( #Figure out where to look
        "$env:USERPROFILE\Downloads\$encryptedCredsFile"
        ,"$env:USERPROFILE\Desktop\$encryptedCredsFile"
        ,"$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\$encryptedCredsFile"
        )

    for($i=0; $i -lt $placesToLook.Count; $i++){ #Look for the file in each location until we find it
        if(Test-Path $placesToLook[$i]){
            $pathToEncryptedCsv = $placesToLook[$i]
            break
            }
        }
    if([string]::IsNullOrWhiteSpace($pathToEncryptedCsv)){ #Break if we can't find it
        Write-Error "Encrypted Client_Crednetials file for [$appName] not found in any of these locations: $($placesToLook -join ", ")"
        break
        }
    else{ #Otherwise, import the file
        $clientCredentials = import-encryptedCsv -pathToEncryptedCsv $pathToEncryptedCsv
        }
    $clientCredentials
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
function get-graphCalendarEvent(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$userId
        ,[parameter(Mandatory = $false)]
            [string]$eventId
        ,[parameter(Mandatory = $false)]
            [string]$filterSubject
        )

    if($filterSubject){$filter += "`$filter=subject eq '$filterSubject'"}

    #Write-Verbose "get-graphCalendarEvent | $($eventId)"
    invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/users/$userId/calendar/events?$filter" -Verbose:$VerbosePreference
}
function get-graphDevices(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $false)]
            [ValidateSet("Android","iOS","Windows")]
            [string]$filterOperatingSystem
        ,[parameter(Mandatory = $false)]
            [string[]]$filterOwnerIds
        ,[parameter(Mandatory = $false)]
            [string[]]$filterDisplayNames
        ,[parameter(Mandatory = $false)]
            [hashtable]$filterCustomEq = @{}
        )

    #
    if($filterOperatingSystem){
        $filter += " and operatingsystem eq '$filterOperatingSystem'"
        }
    #These aren't supported by Graph API yet, so we have to filter client-side :(
    <#if($filterOwnerIds){
        $filter += " and userId in (`'$($filterOwnerIds -join "','")`')"
        }#>
    if($filterDisplayNames){
        $filter += " and displayName in (`'$($filterDisplayNames -join "','")`')"
        }
    $filterCustomEq.Keys | % {
        $filter += " and $_ eq '$($filterCustomEq[$_])'"
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
        if($refiner.Length -gt 1){$refiner = $refiner+"&"} #If there is already another parameter in the refiner, use the '&' symbol to concatenate the strings
        $refiner = $refiner+$filter
        }#>

    Write-Verbose "Graph Query = [/devices$refiner]"
    try{
        $allDevices = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/devices$refiner"
        }
    catch{
        Write-Error "Error retrieving Graph Devices in get-graphDevices() using query [/devices$refiner]"
        Throw $_ #Terminate on this error
        }
    
    if($filterOwnerIds){
        Write-Verbose "Returning all Devices owned by [$($filterOwnerIds -join ",")]"
        $allDevices | ? { #Extract the USER-GID guid and match it to the supplied OwnerIds
            $($($_.physicalIds | ? {$_ -match "USER-GID"}) | % {$($_ -split ":")[1]}) `
                -match `
                $($filterOwnerIds -join "|")
            } | Sort-Object displayName
        }
    elseif($filterDisplayNames2){
        Write-Verbose "Returning all Devices named [$($filterDeviceNames -join ",")]"
        $allDevices | ? {$filterDisplayNames -contains $_.displayName} | Sort-Object displayName
        }
    else{
        Write-Verbose "Returning all Devices"
        $allDevices | Sort-Object displayName
        }
    }
function get-graphDriveItems(){
     [cmdletbinding()]
    param(
        [parameter(Mandatory = $true,ParameterSetName = "root")]
            [parameter(Mandatory = $true,ParameterSetName = "itemId")]
            [parameter(Mandatory = $true,ParameterSetName = "path")]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true,ParameterSetName = "root")]
            [parameter(Mandatory = $true,ParameterSetName = "itemId")]
            [parameter(Mandatory = $true,ParameterSetName = "path")]
            [string]$driveGraphId
        ,[parameter(Mandatory = $true,ParameterSetName = "itemId")]
            [string]$itemGraphId = "root"
        ,[parameter(Mandatory = $true,ParameterSetName = "root")]
            [parameter(Mandatory = $true,ParameterSetName = "itemId")]
            [parameter(Mandatory = $true,ParameterSetName = "path")]
            [ValidateSet("Item","Children")]
            [string]$returnWhat
        ,[parameter(Mandatory = $true,ParameterSetName = "path")]
            [string]$folderPathRelativeToRoot
        ,[parameter(Mandatory = $false,ParameterSetName = "root")]
            [parameter(Mandatory = $false,ParameterSetName = "itemId")]
            [parameter(Mandatory = $false,ParameterSetName = "path")]
            [string]$filterNameRegex
        )
    
    if($returnWhat -eq "Children"){$getChildren = "/children"}

    switch ($PsCmdlet.ParameterSetName){ 
        "root" {
            $results = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/drives/$driveGraphId/root$getChildren" #-ErrorAction $ErrorActionPreference
            if($filterNameRegex){$results | ? {$_.name -match $filterNameRegex}}
            else{$results}
            return
            }
        "itemId" { 
            $results = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/drives/$driveGraphId/items/$itemGraphId$getChildren" #-ErrorAction $ErrorActionPreference
            if($filterNameRegex){$results | ? {$_.name -match $filterNameRegex}}
            else{$results}
            return
            }
        "path" { 
            $results = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/drives/$driveGraphId/root:/$folderPathRelativeToRoot`:$getChildren" #-ErrorAction $ErrorActionPreference
            if($filterNameRegex){$results | ? {$_.name -match $filterNameRegex}}
            else{$results}
            return
            }
        }
    
    }
function get-graphDrives(){
     [cmdletbinding()]
    param(
        [parameter(Mandatory = $true,ParameterSetName = "fromUrl")]
            [parameter(Mandatory = $true,ParameterSetName = "fromSiteId")]
            [parameter(Mandatory = $true,ParameterSetName = "fromUpn")]
            [parameter(Mandatory = $true,ParameterSetName = "fromGroupId")]
            [parameter(Mandatory = $true,ParameterSetName = "fromDriveId")]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true,ParameterSetName = "fromUrl")]
            [string]$siteUrl
        ,[parameter(Mandatory = $true,ParameterSetName = "fromSiteId")]
            [string]$siteGraphId
        ,[parameter(Mandatory = $false,ParameterSetName = "fromSiteId")]
            [string]$listGraphId
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
            [string]$filterDriveName_unsupported
        ,[parameter(Mandatory = $false,ParameterSetName = "fromSiteId")]
            [parameter(Mandatory = $true,ParameterSetName = "fromDriveId")]
            [string]$driveId
        )
    
    if($returnOnlyDefaultDocumentsLibrary){$endpoint = "/drive"}
    elseif($listGraphId){$endpoint = "/lists"}
    else{$endpoint = "/drives"}

    switch ($PsCmdlet.ParameterSetName){ #Build the query based on the parameters supplied. Because we're dealing with several permutations of endpoints (/groups vs /sites & /drive vs /drives), this looks more complicated than it really is. 
        "fromDriveId" {
            invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/drives/$driveId"
            return
            }
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
            if($filterDriveId){$query+="/$filterDriveId"}
            if($listGraphId){$query+="/$listGraphId/drive"}
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
        if($refiner.Length -gt 1){$refiner = $refiner+"&"} #If there is already another parameter in the refiner, use the '&' symbol to concatenate the strings
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
            invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/groups/$filterId$select"
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
        if($refiner.Length -gt 1){$refiner = $refiner+"&"} #If there is already another parameter in the refiner, use the '&' symbol to concatenate the strings
        $refiner = $refiner+$filter
        }
    
    Write-Verbose "`$filter = $filter"
    Write-Verbose "`$select = $select"
    Write-Verbose "`$refiner = $refiner"

    $results = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/groups$refiner"

    if($filterGroupType -eq "MailEnabledSecurity" -or $filterGroupType -eq "Distribution"){
        $results | ? {$_.groupTypes -notcontains "Unified"} | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name ExternalDirectoryObjectId -Value $_.id}
        $results | ? {$_.groupTypes -notcontains "Unified"}
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
function get-graphIntuneDevices(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $false)]
            [ValidateSet("Android","iOS","Windows")]
            [string]$filterOperatingSystem
        ,[parameter(Mandatory = $false)]
            [string[]]$filterOwnerIds
        ,[parameter(Mandatory = $false)]
            [string[]]$filterOwnerUPNs
        ,[parameter(Mandatory = $false)]
            [string[]]$filterDeviceNames
        ,[parameter(Mandatory = $false)]
            [hashtable]$filterCustomEq = @{}
        )

    #
    if($filterOperatingSystem){
        $filter += " and operatingsystem eq '$filterOperatingSystem'"
        }
    #These aren't supported by Graph API yet, so we have to filter client-side :(
    <#if($filterOwnerIds){
        $filter += " and userId in (`'$($filterOwnerIds -join "','")`')"
        }
    if($filterOwnerUPNs){
        $filter += " and userPrincipalName in (`'$($filterOwnerUPNs -join "','")`')"
        }
    if($filterDeviceNames){
        $filter += " and deviceName in (`'$($filterDeviceNames -join "','")`')"
        }
    $filterCustomEq.Keys | % {
        $filter += " and $_ eq '$($filterCustomEq[$_])'"
        }
    #>

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
        if($refiner.Length -gt 1){$refiner = $refiner+"&"} #If there is already another parameter in the refiner, use the '&' symbol to concatenate the strings
        $refiner = $refiner+$filter
        }#>

    Write-Verbose "Graph Query = [/deviceManagement/managedDevices$refiner]"
    try{
        $allIntuneDevices = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/deviceManagement/managedDevices$refiner"
        }
    catch{
        Write-Error "Error retrieving Graph Intune Devices in get-graphIntuneDevices() using query [/deviceManagement/managedDevices$refiner]"
        Throw $_ #Terminate on this error
        }
    
    if($filterOwnerIds){
        Write-Verbose "Returning all Intune Devices owned by [$($filterOwnerIds -join ",")]"
        $allIntuneDevices | ? {$filterOwnerIds -contains $_.userId} | Sort-Object userPrincipalName,deviceName
        }
    elseif($filterOwnerUPNs){
        Write-Verbose "Returning all Intune Devices owned by [$($filterOwnerUPNs -join ",")]"
        $allIntuneDevices | ? {$filterOwnerUPNs -contains $_.userPrincipalName} | Sort-Object userPrincipalName,deviceName
        }
    elseif($filterDeviceNames){
        Write-Verbose "Returning all Intune Devices named [$($filterDeviceNames -join ",")]"
        $allIntuneDevices | ? {$filterDeviceNames -contains $_.deviceName} | Sort-Object userPrincipalName,deviceName
        }
    else{
        Write-Verbose "Returning all Intune Devices"
        $allIntuneDevices | Sort-Object userPrincipalName,deviceName
        }
    }
function get-graphList(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true,ParameterSetName = "IdAndId")]
            [parameter(Mandatory = $true,ParameterSetName = "URLAndId")]
            [parameter(Mandatory = $true,ParameterSetName = "IdAndName")]
            [parameter(Mandatory = $true,ParameterSetName = "URLAndName")]
            [parameter(Mandatory = $true,ParameterSetName = "driveId")]
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
        ,[parameter(Mandatory = $true,ParameterSetName = "driveId")]
            [string]$graphDriveId
        )

    switch ($PsCmdlet.ParameterSetName){
        "driveId" {
            #$drive = get-graphDrives -tokenResponse $tokenResponse -driveId $graphDriveId
            invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/drives/$graphDriveId/list"
            #get-graphList -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listName $drive.name
            return
            }
        {$_ -match "URL"} { #If we've got a URL to the Site, we'll need to get the Id
            Write-Verbose "get-graphList | Getting Site from URL [$serverRelativeSiteUrl]"
            $graphSiteId = $(get-graphSite -tokenResponse $tokenResponse -serverRelativeUrl $serverRelativeSiteUrl).Id
            }
        {$_ -match "AndName"} { #If we've got a URL to the Site, we'll need to get the Id
            $filter = "?`$filter= displayName eq '$([uri]::EscapeDataString($listName))'"
            Write-Verbose "get-graphList | Filter set to [$filter]"
            }
        {$_ -match "AndId"} { #If we've got a URL to the Site, we'll need to get the Id
            $listId = "/$listId"
            Write-Verbose "get-graphList | ListId [$listId]"
            }
        }
    invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/sites/$graphSiteId/lists$ListId$filter"

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
        invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/sites/$graphSiteId/lists/$listId/items/$filterId"
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
        if($refiner.Length -gt 1){$refiner = $refiner+"&"} #If there is already another parameter in the refiner, use the '&' symbol to concatenate the strings
        $refiner = $refiner+$expand
        }

    invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/sites/$graphSiteId/lists/$listId/items$refiner"

    }
function get-graphMailboxSettings(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse
       ,[parameter(Mandatory = $true)]
            [string]$identity
         )
If(($identity -match "@anthesisgroup.com") -or ($identity.Length -eq 36)){
#Identity contains a upn or looks like a guid
$graphQuery = "users/$identity/mailboxSettings"
$response = invoke-graphGet -tokenResponse $tokenResponse -graphQuery $graphQuery -returnEntireResponse -Verbose
$response
}
Else{
Write-Error "Please provide a valid upn or guid"
}
}
function get-graphRoomCalendarEvents(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$roomemail
        ,[parameter(Mandatory = $true)]
            [string]$startDate
        ,[parameter(Mandatory = $true)]
            [string]$endDate
        ,[parameter(Mandatory = $true)]
            [string]$endTime
        ,[parameter(Mandatory = $true)]
            [string]$startTime

        )
    
    $startDate = get-date $startDate -Format "yyyy-MM-dd"
    $endDate = get-date $endDate -Format "yyyy-MM-dd" 
    $startTime = get-date $startTime -Format "HH:mm:ss"
    $endTime = get-date $endTime -Format "HH:mm:ss"

    #FormatDateTimeinfo
    $startDateTime = $startDate + "T" + $startTime
    $endDateTime = $endDate + "T" + $endTime

    #Need to use calendarView for all room events and specify a start/end datetime for query
    invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/users/$roomemail/calendarView/?endDateTime=$endDateTime&startDateTime=$startDateTime" -Verbose:$VerbosePreference
}
function get-graphSelectAllUserProperties(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $false)]
            [ValidateSet("beta","v1.0")]
            [string]$endpoint = "v1.0"
        )
    
    switch ($endpoint){
        'v1.0' {"anthesisgroup_employeeInfo,accountEnabled,assignedLicenses,assignedPlans,businessPhones,city,companyName,country,createdDateTime,creationType,deletedDateTime,department,displayName,employeeId,employeeHireDate,faxNumber,givenName,id,identities,imAddresses,isResourceAccount,jobTitle,lastPasswordChangeDateTime,legalAgeGroupClassification,licenseAssignmentStates,mail,mailNickname,mobilePhone,officeLocation,onPremisesDistinguishedName,onPremisesDomainName,onPremisesExtensionAttributes,onPremisesImmutableId,onPremisesLastSyncDateTime,onPremisesProvisioningErrors,onPremisesSamAccountName,onPremisesSecurityIdentifier,onPremisesSyncEnabled,onPremisesUserPrincipalName,otherMails,passwordPolicies,passwordProfile,postalCode,preferredDataLocation,preferredLanguage,provisionedPlans,proxyAddresses,refreshTokensValidFromDateTime,showInAddressList,signInSessionsValidFromDateTime,state,streetAddress,surname,usageLocation,userPrincipalName,userType"}
        'beta' {"anthesisgroup_employeeInfo,accountEnabled,assignedLicenses,assignedPlans,businessPhones,city,companyName,country,createdDateTime,creationType,deletedDateTime,department,displayName,employeeId,employeeHireDate,faxNumber,givenName,id,identities,imAddresses,isResourceAccount,jobTitle,lastPasswordChangeDateTime,legalAgeGroupClassification,licenseAssignmentStates,mail,mailNickname,mobilePhone,officeLocation,onPremisesDistinguishedName,onPremisesDomainName,onPremisesExtensionAttributes,onPremisesImmutableId,onPremisesLastSyncDateTime,onPremisesProvisioningErrors,onPremisesSamAccountName,onPremisesSecurityIdentifier,onPremisesSyncEnabled,onPremisesUserPrincipalName,otherMails,passwordPolicies,passwordProfile,postalCode,preferredDataLocation,preferredLanguage,provisionedPlans,proxyAddresses,refreshTokensValidFromDateTime,showInAddressList,signInSessionsValidFromDateTime,state,streetAddress,surname,usageLocation,userPrincipalName,userType,infoCatalogs,preferredDataLocation,signInActivity"}
        } #Not Implemented yet: aboutMe, birthday, hireDate, interests, mailboxSettings, mySite,pastProjects, preferredName,responsibilities,schools, skills 
    }
function get-graphShiftOpenShifts(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$teamId
        ,[parameter(Mandatory = $true)]
            [string]$MsAppActsAsUserId
#        ,[parameter(Mandatory = $false)]
#            [string[]]$filterIds
        ,[parameter(Mandatory = $false)]
            [string]$openShiftid
        )
    
    #if($filterIds){$filter += "`$filter=id in (`'$($filterIds -join "','")`')"}
    if($openShiftid){$filter = "/$openShiftid"}

    invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/teams/$teamId/schedule/openShifts$filter" -additionalHeaders @{"MS-APP-ACTS-AS"=$MsAppActsAsUserId}
    
    }
function get-graphShiftOpenShiftChangeRequests(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$teamId
        ,[parameter(Mandatory = $true)]
            [string]$MsAppActsAsUserId
        ,[parameter(Mandatory = $false)]
            [ValidateSet(“approved”,”pending”,"declined")]
            [string]$requestState
        )
    
    if($requestState){$filter += "`$filter = state eq '$requestState'"}

    invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/teams/$teamId/schedule/openShiftChangeRequests?$filter" -additionalHeaders @{"MS-APP-ACTS-AS"=$MsAppActsAsUserId} -Verbose
    
    }
function get-graphShiftOfferShiftRequests(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$teamId
        ,[parameter(Mandatory = $true)]
            [string]$MsAppActsAsUserId
        ,[parameter(Mandatory = $false)]
            [ValidateSet(“approved”,”pending”,"declined")]
            [string]$requestState
        )
    
    if($requestState){$filter += "`$filter = state eq '$requestState'"}

    invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/teams/$teamId/schedule/offerShiftRequests?$filter" -additionalHeaders @{"MS-APP-ACTS-AS"=$MsAppActsAsUserId} -Verbose
    
}
function get-graphShiftUserShifts(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$teamId
        ,[parameter(Mandatory = $true)]
            [string]$MsAppActsAsUserId
#        ,[parameter(Mandatory = $false)]
#            [string[]]$filterIds
        ,[parameter(Mandatory = $false)]
            [string]$filterId
        )
    

    invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/teams/$teamId/schedule/Shifts?$filter" -additionalHeaders @{"MS-APP-ACTS-AS"=$MsAppActsAsUserId} -Verbose:$VerbosePreference
    
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
            $result = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/sites/anthesisllc.sharepoint.com:$sanitisedServerRelativeUrl"
            }
        "IdLonger" { #If we're working with $siteUrl, we'll need to get $siteGraphId (which is more of a faff)
            Write-Verbose "get-graphSite | Getting SiteId from URL [$siteUrl]"
            $result = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/sites/$graphSiteId" 
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
function get-graphteamstructure(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $false)]
            [string]$filter365DisplayName
        ,[parameter(Mandatory = $false)]
            [string]$filter365UPN
        ,[parameter(Mandatory = $false)]
            [string]$filter365Id

        )

if(![string]::IsNullOrWhiteSpace($filter365DisplayName)){

    $365 = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -filterDisplayName $filter365DisplayName
    If($365){
        $managers = @(get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $365.anthesisgroup_UGSync.dataManagerGroupId -memberType Members)
        $members = @(get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $365.anthesisgroup_UGSync.memberGroupId -memberType Members)
        
        $result = New-Object psobject -Property @{
            managers = $managers.userPrincipalName
            members = $members.userPrincipalName
        }
        <#    
        Write-Host -f Cyan "$($team)"
        Write-Host -f White "Data Managers:"
        ForEach($manager in $managers){
        Write-Host -f DarkYellow "$($manager.userPrincipalName)"
        }
        Write-Host -f White "Members:"
        ForEach($member in $members){
        Write-Host -f DarkYellow "$($member.userPrincipalName)"
        }#>
        return $result
    }
    Else{
        Write-Host "$($filter365DisplayName): 365 group not found" -ForegroundColor Red
    }
}
if(![string]::IsNullOrWhiteSpace($filter365UPN)){
    $365 = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -filterUpn $filter365UPN
    If($365){
        $managers = @( get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $365.anthesisgroup_UGSync.dataManagerGroupId -memberType Members)
        $members = @(get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $365.anthesisgroup_UGSync.memberGroupId -memberType Members)
        
        $result = New-Object psobject -Property @{
            managers = $managers.userPrincipalName
            members = $members.userPrincipalName
        }
        <#   
        Write-Host -f Cyan "$($team)"
        Write-Host -f DarkYellow "Data Managers:"
        ForEach($manager in $managers){
        Write-Host -f White "$($manager.userPrincipalName)"
        }
        Write-Host -f DarkYellow "Members:"
        ForEach($member in $members){
        Write-Host -f White "$($member.userPrincipalName)"
        }
        #>
        return $result
    }
    Else{
        Write-Host "$($filter365UPN): group not found" -ForegroundColor Red
    }

}
if(![string]::IsNullOrWhiteSpace($filter365Id)){
    $365 = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -filterId $filter365Id
    If($365){
        $managers = @( get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $365.anthesisgroup_UGSync.dataManagerGroupId -memberType Members)
        $members = @(get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $365.anthesisgroup_UGSync.memberGroupId -memberType Members)
        
        $result = New-Object psobject -Property @{
            managers = $managers.userPrincipalName
            members = $members.userPrincipalName
        }
        <#   
        Write-Host -f Cyan "$($team)"
        Write-Host -f DarkYellow "Data Managers:"
        ForEach($manager in $managers){
        Write-Host -f White "$($manager.userPrincipalName)"
        }
        Write-Host -f DarkYellow "Members:"
        ForEach($member in $members){
        Write-Host -f White "$($member.userPrincipalName)"
        }
        #>
        return $result
    }
    Else{
        Write-Host "$($filter365UPN): group not found" -ForegroundColor Red
    }

}

}
function get-graphTeamsPrivateChannels(){
    [cmdletbinding()]
    param(
         [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $false)]
            [string]$serverRelativeUrl
        ,[parameter(Mandatory = $false)]
            [string]$graphUnifiedGroupId
        )

If($serverRelativeUrl){
$graphSite = get-graphSite -tokenResponse $tokenResponse -serverRelativeUrl $($serverRelativeUrl) -Verbose
}
If($graphSite){
$filter  = "`$filter=membershipType eq 'private'"
$unifiedgroup = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -filterDisplayName $($graphSite.displayName)
$PrivateChannels = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/teams/$($unifiedgroup.id)/channels?$($filter)" -useBetaEndPoint
$privatechannels
}

If($graphUnifiedGroupId){
$filter  = "`$filter=membershipType eq 'private'"
$PrivateChannels = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/teams/$($graphUnifiedGroupId)/channels?$($filter)" -Verbose -useBetaEndPoint
$privatechannels
}

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
            $authCode = get-graphAuthCode -clientID $aadAppCreds.ClientID -redirectUri $aadAppCreds.Redirect -scope $scope
            $ReqTokenBody = @{
                Grant_Type    = "authorization_code"
                Scope         = $scope
                client_Id     = $aadAppCreds.ClientID
                Client_Secret = $aadAppCreds.Secret
                redirect_uri  = $aadAppCreds.Redirect
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
function get-graphUserLineManager(){
    [cmdletbinding(DefaultParameterSetName = "noManagementLevels")]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse
        ,[parameter(Mandatory = $true)]
            [string]$userIdOrUpn
        ,[parameter(Mandatory = $true,ParameterSetName = "fixedManagementLevels")]
            [ValidateRange(1,1000)] #Graph API limit https://docs.microsoft.com/en-us/graph/api/user-list-manager?view=graph-rest-1.0&tabs=http
            [int]$returnManagementLevels
        ,[parameter(Mandatory = $true,ParameterSetName = "maxManagementLevels")]
            [switch]$returnAllManagementLevels
        ,[parameter(Mandatory = $false)]
            [switch]$selectAllProperties = $false
        )


    if($selectAllProperties){$select = "`$select=$(get-graphSelectAllUserProperties)"}
    switch ($PsCmdlet.ParameterSetName){
        'noManagementLevels'    {$refiner = "/manager?$select"}
        'fixedManagementLevels' {$refiner = "?`$expand=manager(`$levels=$returnManagementLevels;$select)"}
        'maxManagementLevels'   {$refiner = "?`$expand=manager(`$levels=max;$select)"}
        }
    try{
        invoke-graphGet -tokenResponse $tokenResponse -graphQuery "users/$userIdOrUpn$refiner"
        }
    catch{
        if($_.Exception -match "(404)"){return}#Error 404 means no Line Manager, so just ignore it
        else {
            write-host get-errorsummary $_
            }
        }

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
            [string[]]$filterUpns
        ,[parameter(Mandatory = $false)]
            [hashtable]$filterCustomEq = @{}
        ,[parameter(Mandatory = $false)]
            [ValidateSet($null,"Anthesis (UK) Ltd (GBR)","Anthesis Consulting Group (GBR)","Anthesis Consultoria Ambiental ltda (BRA)","Anthesis Energy UK Ltd (GBR)","Anthesis Enveco AB (SWE)","Anthesis Finland Oy (FIN)","Anthesis GmbH (DEU)","Anthesis Ireland Ltd (IRL)","Anthesis LLC (USA)","Anthesis Middle East (ARE)","Anthesis Philippines Inc. (PHL)","Anthesis Srl (ITA)","Caleb Management Services Ltd (GBR)","France (FRA)","Lavola 1981 SAU (ESP)","Lavola Andora SA (AND)","Lavola Columbia (COL)","The Goodbrand Works Ltd (GBR)")]
            [string]$filterBusinessUnit
        ,[parameter(Mandatory = $false)]
            [switch]$filterLicensedUsers = $false
        ,[parameter(Mandatory = $false)]
            [string[]]$selectCustomProperties
        ,[parameter(Mandatory = $false)]
            [switch]$selectAllProperties = $false
        ,[parameter(Mandatory = $false)]
            [switch]$includeLineManager = $false
        ,[parameter(Mandatory = $false)]
            [switch]$useBetaEndPoint = $false
        )

    #We need the GroupId, so if we were only given the UPN, we need to find the Id from that.
    if($filterUsageLocation){
        $filter += " and usageLocation eq '$filterUsageLocation'"
        }
    if($filterUpns){
        $filter += " and (userPrincipalName eq '$($filterUpns -join "' or userPrincipalName eq '")')"
        }
    if($filterBusinessUnit){
        $filter += " and anthesisgroup_employeeInfo/businessUnit eq '$filterBusinessUnit'"
        }
    $filterCustomEq.Keys | % {
        $filter += " and $_ eq '$($filterCustomEq[$_])'"
        }
    $filterCustomEq.Keys | % {
        $filter += " and $_ eq '$($filterCustomEq[$_])'"
        }

    if($filterLicensedUsers){
        $select = "id,displayName,jobTitle,mail,userPrincipalName,usageLocation,assignedLicenses,companyName,country,department,anthesisgroup_employeeInfo"
        }
    if($selectAllProperties){
        $includeLineManager = $true #Assume that Line Managers are part of "allProperties"
        if($useBetaEndPoint){$select = get-graphSelectAllUserProperties -endpoint beta}
        else{$select = get-graphSelectAllUserProperties}
        } #Not Implemented yet: aboutMe, birthday, hireDate, interests, mailboxSettings, mySite,pastProjects, preferredName,responsibilities,schools, skills 
    @($selectCustomProperties | Select-Object) | % {
        $select += ",$_"
        }
    
    if($includeLineManager){
        $expand += ",manager(`$levels=1)"
        }
    
    #Build the refiner based on the parameters supplied
    $refiner = "?"
    if($select){
        if($refiner.Length -gt 1){$refiner = $refiner+"&"} #If there is already another parameter in the refiner, use the '&' symbol to concatenate the strings
        $refiner = $refiner+"`$select=$($select -replace "^,",'')"#Add the select to the refiner, trimming off any leading "," (don't use .TrimStart() because it's bafflingly unpredictable)
        }
    if($filter){
        if($refiner.Length -gt 1){$refiner = $refiner+"&"} #If there is already another parameter in the refiner, use the '&' symbol to concatenate the strings
        $refiner = $refiner+"`$filter=$($filter -replace "^ and ",'')"#Add the filter to the refiner, trimming off any leading " and " (don't use .TrimStart() because it's bafflingly unpredictable)
        }
    if($expand){
        if($refiner.Length -gt 1){$refiner = $refiner+"&"} #If there is already another parameter in the refiner, use the '&' symbol to concatenate the strings
        $refiner = $refiner+"`$expand=$($expand -replace "^,",'')"#Add the expand to the refiner, trimming off any leading ","  (don't use .TrimStart() because it's bafflingly unpredictable)
        }

    Write-Verbose "Graph Query = [users$refiner]"
    try{
        $allUsers = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "users$refiner" -useBetaEndPoint:$useBetaEndPoint
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
        ,[parameter(Mandatory = $false)]
            [switch]$selectAllProperties = $false
        ,[parameter(Mandatory = $false)]
            [switch]$includeLineManager = $false
        )

    #We need the GroupId, so if we were only given the UPN, we need to find the Id from that.
    switch ($PsCmdlet.ParameterSetName){
        “groupUpn”  {
            Write-Verbose "We've been given a GroupUPN, so we need the GroupId"
            $graphGroup = get-graphGroups -tokenResponse $tokenResponse -filterUpn $groupUpn
            if(!$graphGroup){
                Write-Error "Could not retrieve Graph Group using UPN [$groupUpn]. Check the UPN is valid and try again."
                break
                }
            $groupId = $graphGroup.id
            Write-Verbose "[$groupUpn] Id is [$groupId]"
            }
        }
    if($returnOnlyLicensedUsers){
        $select = "id,displayName,jobTitle,mail,userPrincipalName,usageLocation,assignedLicenses,anthesisgroup_employeeInfo"
        $returnOnlyUsers = $true #Licensed Users are a subset of Users, so $returnOnlyUsers = $true is implied if $returnOnlyLicensedUsers = $true
        }
    if($selectAllProperties){
        $includeLineManager = $true #Assume that Line Managers are part of "allProperties"
        if($useBetaEndPoint){$select = get-graphSelectAllUserProperties -endpoint beta}
        else{$select = get-graphSelectAllUserProperties}
        } #Not Implemented yet: aboutMe, birthday, hireDate, interests, mailboxSettings, mySite,pastProjects, preferredName,responsibilities,schools, skills 

    #Build the refiner based on the parameters supplied
    $refiner = "?"
    if($select){
        if($refiner.Length -gt 1){$refiner = $refiner+"&"} #If there is already another parameter in the refiner, use the '&' symbol to concatenate the strings
        $refiner = $refiner+"`$select=$($select -replace "^,",'')"#Add the select to the refiner, trimming off any leading "," (don't use .TrimStart() because it's bafflingly unpredictable)
        }
    if($filter){
        if($refiner.Length -gt 1){$refiner = $refiner+"&"} #If there is already another parameter in the refiner, use the '&' symbol to concatenate the strings
        $refiner = $refiner+"`$filter=$($filter -replace "^ and ",'')" #Add the filter to the refiner, trimming off any leading " and " (don't use .TrimStart() because it's bafflingly unpredictable)
        }

    Write-Verbose "Graph Query = [groups/$($groupId)/$($memberType+$refiner)]"
    try{
        $allMembers = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "groups/$groupId/$memberType$refiner"
        }
    catch{
        Write-Error "Error retrieving Graph Group $memberType in get-graphUsersFromGroup() using query [groups/$($groupId)/$($memberType+$refiner)]"
        Throw $_ #Terminate on this error
        }

    if($includeLineManager){ #Relationships (like /owners) don't support $expand parameters, so we have to enumerate the Line Managers per-user
        $allMembers | ? {$_.'@odata.type' -eq "#microsoft.graph.user"} | % {
            try{$thisLineManager = $(get-graphUserLineManager -tokenResponse $tokenResponse -userIdOrUpn $_.userPrincipalName -selectAllProperties:$selectAllProperties)}
            catch{
                if($_.Exception -match "(404)"){
                    <#Do nothing - this means the user did not have a Line Manager assigned#>
                    write-warning "User [$($_.userPrincipalName)] has no Line Manager assigned"
                    }
                else{get-errorSummary -errorToSummarise $_}
                }
            Add-Member -InputObject $_ -MemberType NoteProperty -Name manager -Value $thisLineManager
            }
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
function get-graphUserGroupMembership(){
    [cmdletbinding()]
    param(
         [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$userUpn
        )
        invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/users/$($userUpn)/memberOf"  -Verbose:$VerbosePreference
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
        ,[parameter(Mandatory = $false)]
            [hashtable]$additionalHeaders
        )
    $sanitisedGraphQuery = $graphQuery.Trim("/")
    Write-Verbose "https://graph.microsoft.com/v1.0/$sanitisedGraphQuery"

    $headers = @{Authorization = "Bearer $($tokenResponse.access_token)"}
    if($additionalHeaders){
        $additionalHeaders.GetEnumerator() | %{
            $headers.Add($_.Key,$_.Value)
            }
        }

    if($graphBodyHashtable){
        $graphBodyJson = ConvertTo-Json -InputObject $graphBodyHashtable
        Write-Verbose $graphBodyJson
        $graphBodyJsonEncoded = [System.Text.Encoding]::UTF8.GetBytes($graphBodyJson)
        $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/$sanitisedGraphQuery" -ContentType "application/json; charset=utf-8" -Headers $headers -Method DELETE -Body $graphBodyJsonEncoded
        }
    else{
        $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/$sanitisedGraphQuery" -ContentType "application/json; charset=utf-8" -Headers $headers -Method DELETE
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
        ,[parameter(Mandatory = $false)]
            [hashtable]$additionalHeaders
        ,[parameter(Mandatory = $false)]
            [switch]$useBetaEndPoint = $false
        )
    $sanitisedGraphQuery = $graphQuery.Trim("/")
    #$sanitisedGraphQuery = [uri]::EscapeDataString($(sanitise-forSql $([uri]::UnescapeDataString($graphQuery).Trim("/"))))
    $headers = @{Authorization = "Bearer $($tokenResponse.access_token)"}
    if($additionalHeaders){
        $additionalHeaders.GetEnumerator() | %{
            $headers.Add($_.Key,$_.Value)
            }
        }
    #Write-Verbose $(stringify-hashTable -hashtable $headers -interlimiter "=" -delimiter ";")
    if($useBetaEndPoint){$endpoint = "beta"}
    else{$endpoint = "v1.0"}
    do{
        Write-Verbose "https://graph.microsoft.com/$endpoint/$sanitisedGraphQuery"
        $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/$endpoint/$sanitisedGraphQuery" -ContentType "application/json; charset=utf-8" -Headers $headers -Method GET  #-Verbose:$VerbosePreference -ErrorAction:$ErrorActionPreference
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
        if(![string]::IsNullOrWhiteSpace($response.'@odata.nextLink')){$sanitisedGraphQuery = $response.'@odata.nextLink'.Replace("https://graph.microsoft.com/$endpoint/","")}
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
        ,[parameter(Mandatory = $false)]
            [hashtable]$additionalHeaders
        )

    $sanitisedGraphQuery = $graphQuery.Trim("/")
    $headers = @{Authorization = "Bearer $($tokenResponse.access_token)"}
    if($additionalHeaders){
        $additionalHeaders.GetEnumerator() | %{
            $headers.Add($_.Key,$_.Value)
            }
        }
    Write-Verbose "https://graph.microsoft.com/v1.0/$sanitisedGraphQuery"
        
    $graphBodyJson = ConvertTo-Json -InputObject $graphBodyHashtable
    Write-Verbose $graphBodyJson
    $graphBodyJsonEncoded = [System.Text.Encoding]::UTF8.GetBytes($graphBodyJson)
    
    Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/$sanitisedGraphQuery" -Body $graphBodyJsonEncoded -ContentType "application/json; charset=utf-8" -Headers $headers -Method Patch
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
        ,[parameter(Mandatory = $false)]
            [hashtable]$additionalHeaders
        )

    $sanitisedGraphQuery = $graphQuery.Trim("/")
    $headers = @{Authorization = "Bearer $($tokenResponse.access_token)"}
    if($additionalHeaders){
        $additionalHeaders.GetEnumerator() | %{
            $headers.Add($_.Key,$_.Value)
            }
        }
    Write-Verbose "https://graph.microsoft.com/v1.0/$sanitisedGraphQuery"
        
    $graphBodyJson = ConvertTo-Json -InputObject $graphBodyHashtable -Depth 10
    Write-Verbose $graphBodyJson
    $graphBodyJsonEncoded = [System.Text.Encoding]::UTF8.GetBytes($graphBodyJson)
    
    Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/$sanitisedGraphQuery" -Body $graphBodyJsonEncoded -ContentType "application/json; charset=utf-8" -Headers $headers -Method Post
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
        ,[parameter(Mandatory = $false)]
            [hashtable]$additionalHeaders
        )

    $sanitisedGraphQuery = $graphQuery.Trim("/")
    $headers = @{Authorization = "Bearer $($tokenResponse.access_token)"}
    if($additionalHeaders){
        $additionalHeaders.GetEnumerator() | %{
            $headers.Add($_.Key,$_.Value)
            }
        }
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

    Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/$sanitisedGraphQuery" -Body $bodyData -ContentType $contentType -Headers $headers -Method Put
    }
function move-graphDriveItem(){
     [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$driveGraphIdSource
        ,[parameter(Mandatory = $true)]
            [string]$itemGraphIdSource
        ,[parameter(Mandatory = $false)]
            [string]$driveGraphIdDestination
        ,[parameter(Mandatory = $true)]
            [string]$parentItemGraphIdDestination
        ,[parameter(Mandatory = $false)]
            [string]$newItemName
        )

    $reqBodyParentRef = @{id=$parentItemGraphIdDestination} #All MOVE operations explicitly require a destination ID to move the item _to_ (no assumptions about Root destinations)
    if($driveGraphIdDestination){
        $reqBodyParentRef.Add("driveId",$driveGraphIdDestination) #If we're MOVE to a different Drive, add the ID here (otherwise MOVE assumed to be within the current Drive)
        }

    $reqBody = @{parentReference=$reqBodyParentRef}
    if($newItemName){
        $reqBody.Add("name",$newItemName) #If we're changing the Item during the MOVE, add that here
        }

    $query = "/drives/$driveGraphIdSource/items/$itemGraphIdSource"
    $movedItem = invoke-graphPatch -tokenResponse $tokenResponse -graphQuery $query -graphBodyHashtable $reqBody
    $movedItem

    <#https://docs.microsoft.com/en-us/graph/api/driveitem-move?view=graph-rest-1.0&tabs=http
    PATCH /me/drive/items/{item-id}
    Content-type: application/json

    {
      "parentReference": {
        "id": "{new-parent-folder-id}"
      },
      "name": "new-item-name.txt"
    }

    #https://docs.microsoft.com/en-us/graph/api/driveitem-copy?view=graph-rest-1.0
    POST /me/drive/items/{item-id}/copy
    Content-Type: application/json

    {
      "parentReference": {
        "driveId": "6F7D00BF-FC4D-4E62-9769-6AEA81F3A21B",
        "id": "DCD0D3AD-8989-4F23-A5A2-2C086050513F"
      },
      "name": "contoso plan (copy).txt"
    }    #>

    #Finally, submit the query and return the results
   
    }
function new-graphCalendarEvent(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$userId
        ,[parameter(Mandatory = $true)]
            [string]$subject
        ,[parameter(Mandatory = $true)]
            [datetime]$start
        ,[parameter(Mandatory = $true)]
            [ArgumentCompleter({
                param($Command, $Parameter, $WordToComplete, $CommandAst, $FakeBoundParams)
                $((Get-TimeZone -ListAvailable).Id | sort) #Use Get-TimeZone to populate the pseudoValidateSet
                })]
            [ValidateScript({$_ -in ((Get-TimeZone -ListAvailable).Id)})]
            [string]$startTimeZone
        ,[parameter(Mandatory = $true)]
            [datetime]$end
        ,[parameter(Mandatory = $true)]
            [ArgumentCompleter({
                param($Command, $Parameter, $WordToComplete, $CommandAst, $FakeBoundParams)
                $((Get-TimeZone -ListAvailable).Id | sort) #Use Get-TimeZone to populate the pseudoValidateSet
                })]
            [string]$endTimeZone
        ,[parameter(Mandatory = $false)]
            [string]$location
        ,[parameter(Mandatory = $false)]
            [string]$bodyHTML
        ,[parameter(Mandatory = $false)]
            [bool]$isTeamsMeeting
        ,[parameter(Mandatory = $false)]
            [int]$reminderMinutesBeforeStart
        ,[parameter(Mandatory = $false)]
            [ValidateSet ("free","tentative","busy","oof","workingElsewhere","unknown")]
            [string]$freeBusyStatus = "busy"
        ,[parameter(Mandatory = $false)]
            [string[]]$categories
        )

    $event = @{
        subject = $subject
        start   = @{
            dateTime = $(Get-Date $start -Format s)
            timeZone = $startTimeZone
            }
        end   = @{
            dateTime = $(Get-Date $end -Format s)
            timeZone = $endTimeZone
            }
        }
    if($location){
        $event.Add("location",@{displayName=$location})
        }
    if($bodyHTML){
        $body = @{}
        $body.Add("contentType","HTML")
        $body.Add("content",$bodyHTML)
        $event.Add("body",$body)
        }
    if($isTeamsMeeting){
        $event.Add("isOnlineMeeting",$true)
        $event.Add("onlineMeetingProvider","teamsForBusiness")
        }
    if($reminderMinutesBeforeStart){
        $event.Add("reminderMinutesBeforeStart",$reminderMinutesBeforeStart)
        $event.Add("isReminderOn",$true)
        }
    if($freeBusyStatus){
        $event.Add("showAs",$freeBusyStatus)
        }
    if($categories){
        $event.Add("categories",$categories)
        }

    Write-Verbose "new-graphCalendarEvent | $(stringify-hashTable $event -interlimiter "=" -delimiter "; ")"
    invoke-graphPost -tokenResponse $tokenResponse -graphQuery "/users/$userId/calendar/events" -graphBodyHashtable $event
    #https://docs.microsoft.com/en-us/graph/api/calendar-post-events?view=graph-rest-1.0&tabs=http
    }
function new-graphGroup(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$groupDisplayName
        ,[parameter(Mandatory = $false)]
            [string]$groupDescription
        ,[parameter(Mandatory = $true)]
            [ValidateSet ("Security")]
            [string]$groupType
        ,[parameter(Mandatory = $true)]
            [ValidateSet ("Assigned")]
            [string]$membershipType
        ,[parameter(Mandatory = $true)]
            [string[]]$groupOwners
        ,[parameter(Mandatory = $true)]
            [string[]]$groupMembers
        )    

    $bodyHash = @{"displayName"=$groupDisplayName;"description"=$groupDescription}
    switch($groupType){
        "Security" {
            $bodyHash.Add("mailEnabled",$false)
            $bodyHash.Add("mailNickname",$(guess-aliasFromDisplayName $groupDisplayName))
            $bodyHash.Add("securityEnabled",$true)
            }
        }

    #Get GUIDs for Owners & Members
    if($groupOwners){ #Split into GUIDs & (presumed) UPNs
        $ownersGuids = @()
        $groupOwners | Select-Object | % {
            if(test-isGuid $_){$ownersGuids += $_}
            elseif($_ -match "@"){$ownersUpns += $_}
            else{[array]$ownersUpns += "$_@anthesisgroup.com"}
            }
        if($ownersUpns){ #If we have UPNs, we need to convert them to GUIDs
            try{[array]$ownersGuids += $(get-graphUsers -tokenResponse $tokenResponse -filterUpns $ownersUpns).id} #Try doing it all in one query first (more efficient)
            catch{ #If it doesn't work, process the users individually so we can warn about the specific problems
                $ownersUpns | Select-Object | % {
                    $user = get-graphUsers -tokenResponse $tokenResponse -filterUpns $_
                    if($user){$ownersGuids += $user.id}
                    else{Write-Warning "Owner [$($_)] could not be resolved to an AAD user. Cannot add to group."}
                    }
                }
            }
        }
    if($groupMembers){ #Split into GUIDs & (presumed) UPNs
        $memberGuids = @()
        $groupMembers | Select-Object | % {
            if(test-isGuid $_){$memberGuids += $_}
            elseif($_ -match "@"){[array]$memberUpns += $_}
            else{[array]$memberUpns += "$_@anthesisgroup.com"}
            }
        if($memberUpns){ #If we have UPNs, we need to convert them to GUIDs
            try{[array]$memberGuids += $(get-graphUsers -tokenResponse $tokenResponse -filterUpns $memberUpns).id} #Try doing it all in one query first (more efficient)
            catch{ #If it doesn't work, process the users individually so we can warn about the specific problems
                $memberUpns | Select-Object | % {
                    $user = get-graphUsers -tokenResponse $tokenResponse -filterUpns $_
                    if($user){$memberGuids += $user.id}
                    else{Write-Warning "Member [$($_)] could not be resolved to an AAD user. Cannot add to group."}
                    }
                }
            }

        }

    #Max 20 combined Owners/Members can be supplied during Group creation: https://docs.microsoft.com/en-us/graph/api/group-post-groups?view=graph-rest-1.0&tabs=http#example-2-create-a-group-with-owners-and-members
    $ownersGuids | Select-Object | % {
        if($i -lt 20){
            [array]$ownersArray += "https://graph.microsoft.com/v1.0/directoryObjects/$_"
            $i++
            }
        }
    $memberGuids | Select-Object | % {
        if($i -lt 20){
            [array]$membersArray += "https://graph.microsoft.com/v1.0/directoryObjects/$_"
            $i++
            }
        }

    if($ownersArray){$bodyHash.Add("owners@odata.bind",$ownersArray)}
    if($membersArray){$bodyHash.Add("members@odata.bind",$membersArray)}

    $newGroup = invoke-graphPost -tokenResponse $tokenResponse -graphQuery "/groups" -graphBodyHashtable $bodyHash

    #Max 20 Owners/Members can be supplied during Group creation: https://docs.microsoft.com/en-us/graph/api/group-post-groups?view=graph-rest-1.0&tabs=http#example-2-create-a-group-with-owners-and-members
    if($memberGuids.Count + $ownersGuids.Count -gt 20){
        $fullOwners = add-graphUsersToGroup -tokenResponse $tokenResponse -graphGroupId $newGroup.id -memberType Owners -graphUserIds $ownersGuids
        $fullMembers = add-graphUsersToGroup -tokenResponse $tokenResponse -graphGroupId $newGroup.id -memberType Members -graphUserIds $memberGuids
        $newGroup = get-graphGroups -tokenResponse $tokenResponse -filterId $newGroup.id
        }

    $newGroup
    }
function new-graphList(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$siteGraphId
        ,[parameter(Mandatory = $true)]
            [string]$listDisplayName
        ,[parameter(Mandatory = $true)]
            [ValidateSet ("documentLibrary")]
            [string]$listType
        )    

    $bodyHash = @{"displayName"=$listDisplayName;"list"=@{"template"=$listType}}
    invoke-graphPost -tokenResponse $tokenResponse -graphQuery "/sites/$siteGraphId/lists" -graphBodyHashtable $bodyHash
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
            $graphSiteId = $(get-graphSite -tokenResponse $tokenResponse -serverRelativeUrl $serverRelativeSiteUrl).id
            }
        {$_ -match "AndName"} {
            Write-Verbose "new-graphListItem | Getting ListId"
            $listId = $(get-graphList -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listName $listName).id
            }
        }
    $graphBodyHash = @{"fields"=$listItemFieldValuesHash}
    Write-Verbose "new-graphListItem | $(stringify-hashTable $listItemFieldValuesHash)"
    invoke-graphPost -tokenResponse $tokenResponse -graphQuery "/sites/$graphSiteId/lists/$listId/items" -graphBodyHashtable $graphBodyHash
    }
function new-graphOpenShiftShared(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$teamId
        ,[parameter(Mandatory = $true)]
            [string]$MsAppActsAsUserId
        ,[parameter(Mandatory = $true)]
            [string]$schedulingGroupId
        ,[parameter(Mandatory = $true)]
            [string]$shiftName
        ,[parameter(Mandatory = $true)]
            [string]$shiftNotes
        ,[parameter(Mandatory = $true)]
            [int]$availableSlots
        ,[parameter(Mandatory = $true)]
            [datetime]$startDateTime
        ,[parameter(Mandatory = $true)]
            [datetime]$endDateTime
        ,[parameter(Mandatory = $false)]
            [ValidateSet ("White","Blue","Green","Purple","Pink","Yellow","Gray","DarkBlue","DarkGreen","DarkPurple","DarkPink","DarkYellow")]
            [string]$shiftColour
        )
    #Creates an already-shared OpenShift
    $shiftDetails=@{
        displayName=$shiftName
        notes=$shiftNotes
        startDateTime=$(Get-Date $startDateTime -Format o) #Format dates as ISO 8601
        endDateTime=$(Get-Date $endDateTime -Format o) #Format dates as ISO 8601
        theme=$shiftColour
        openSlotCount=$availableSlots
        }
    $newShift = @{
       schedulingGroupId=$schedulingGroupId
       sharedOpenShift=$shiftDetails
       }

    invoke-graphPost -tokenResponse $tokenResponse -graphQuery "/teams/$teamId/schedule/openShifts" -graphBodyHashtable $newShift -additionalHeaders @{"MS-APP-ACTS-AS"=$msAppActsAsUserId}
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
                [array]$graphUserIds += $(get-graphUsers -tokenResponse $tokenResponse -filterUpns $_).id
                }
            
            } 
        }

    $graphUserIds | % {
        #$bodyHash = @{"@odata.id"="https://graph.microsoft.com/v1.0/users/$_"}
        invoke-graphDelete -tokenResponse $tokenResponse -graphQuery "/groups/$graphGroupId/$memberType/$_/`$ref"
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
    Write-Verbose "`treset-graphUnifiedGroupSettingsToOriginals"    #Compare current Unified Group settings against orginal settings and revert
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
            Send-MailMessage -From groupbot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Team [$($combinedMesg.displayName)] settings rolled back" -BodyAsHtml $body -To $($owners.mail) -Cc $itAdminEmailAddresses -Encoding UTF8
            #Send-MailMessage -From groupbot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Team [$($combinedMesg.displayName)] settings rolled back" -BodyAsHtml $body -To kevin.maitland@anthesisgroup.com  -Encoding UTF8
            }
        #Now, fix the settings:
        invoke-graphPatch -tokenResponse $tokenResponse -graphQuery "/groups/$($graphGroupExtended.id)" -graphBodyHashtable $changes
        #And check the Membership settings are correct too:
        set-graphUnifiedGroupGuestSettings -tokenResponse $tokenResponse -graphUnifiedGroupExtended $graphGroupExtended
        }
    }
function send-graphMailMessage(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [ValidatePattern("@")]
            [string]$fromUpn
        ,[parameter(Mandatory = $true)]
            [ValidatePattern("@")]
            [string[]]$toAddresses
        ,[parameter(Mandatory = $false)]
            [ValidatePattern("@")]
            [string[]]$ccAddresses
        ,[parameter(Mandatory = $false)]
            [ValidatePattern("@")]
            [string[]]$bccAddresses
        ,[parameter(Mandatory = $true)]
            [string]$subject
        ,[parameter(Mandatory = $true,ParameterSetName = "text")]
            [string]$bodyText
        ,[parameter(Mandatory = $true,ParameterSetName = "HTML")]
            [string]$bodyHtml
        ,[parameter(Mandatory = $false)]
            [bool]$saveToSentItems = $true
        ,[parameter(Mandatory = $false)]
            [ValidateSet ("low","normal","high")]
            [string]$priority = "normal"
        )

    [array]$formattedToAddresses = $toAddresses | % {
        @{emailAddress=@{'address'=$_}}
        }
    [array]$formattedFromAddresses = $fromUpn | % {
        @{emailAddress=@{'address'=$_}}
        }
    $message = @{
        toRecipients = $formattedToAddresses
        subject = $subject
        importance=$priority
        #from = $formattedFromAddresses
        #sender = $formattedFromAddresses
        }

    if($ccAddresses){
        [array]$formattedCcAddresses = $ccAddresses | % {
            @{emailAddress=@{'address'=$_}}
            }
        $message.Add("ccRecipients",$formattedCcAddresses)
        }
    if($bccAddresses){
        [array]$formattedBccAddresses = $bccAddresses | % {
            @{emailAddress=@{'address'=$_}}
            }
        $message.Add("bccRecipients",$formattedBccAddresses)
        }
    if($bodyText){$message.Add("body",@{"contentType"="Text";"content"=$bodyText})}
    if($bodyHtml){$message.Add("body",@{"contentType"="HTML";"content"=$bodyHtml})}

    invoke-graphPost -tokenResponse $tokenResponse -graphQuery "/users/$fromUpn/sendMail" -graphBodyHashtable @{"message"=$message;"saveToSentItems"=$saveToSentItems}
    }
function set-graphDrive_unsupported(){
     [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$driveId
        ,[parameter(Mandatory = $true)]
            [hashtable]$drivePropertyHash = @{}
        )
    $validProperties = @("description","displayName","name")
    $duffProperties = @()
    $drivePropertyHash.Keys | % { #Check the properties we're going to try and update the Drive with are valid:
        if($validProperties -notcontains $_ ){
            $duffProperties += $_
            }
        }

    if($duffProperties.Count -gt 0){
        Write-Error -Message "Property(s) [$($duffProperties -join ", ")] is invalild for Graph Drive object. Will not attempt to update."
        return
        }

    invoke-graphPost -tokenResponse $tokenResponse -graphQuery "/drives/$driveId" -graphBodyHashtable $drivePropertyHash
    }
function set-graphDriveItem(){
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$driveId
        ,[parameter(Mandatory = $true)]
            [string]$driveItemId
        ,[parameter(Mandatory = $true)]
            [hashtable]$driveItemPropertyHash = @{}
        )
    $validProperties = @("name")
    $duffProperties = @()
    $driveItemPropertyHash.Keys | % { #Check the properties we're going to try and update the Drive with are valid:
        if($validProperties -notcontains $_ ){
            $duffProperties += $_
            }
        }

    if($duffProperties.Count -gt 0){
        Write-Error -Message "Property(s) [$($duffProperties -join ", ")] is invalild for Graph Drive object. Will not attempt to update."
        return
        }

    invoke-graphPatch -tokenResponse $tokenResponse -graphQuery "/drives/$driveId/items/$driveItemId" -graphBodyHashtable $driveItemPropertyHash
    
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
            $graphGroup = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -filterUpn $groupUpn
            if(!$graphGroup){
                Write-Error "Could not retrieve Graph Group using UPN [$groupUpn]. Check the UPN is valid and try again."
                break
                }
            }
        "groupObject" {
            if($graphGroup.psobject.Properties.Name -notcontains "CustomAttribute1"){
                Write-Verbose "We've been given a Group object, but it's missing the CustomAttributes"
                $graphGroup = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -filterUpn $groupUpn
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
    $usersToSet = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $graphGroup.CustomAttribute3 -memberType TransitiveMembers -returnOnlyLicensedUsers
    
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
function set-graphList(){
     [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$graphSiteId
        ,[parameter(Mandatory = $true)]
            [string]$graphListId
        ,[parameter(Mandatory = $true)]
            [hashtable]$listPropertyHash = @{}
        )
    $validProperties = @("description","displayName")
    $duffProperties = @()
    $listPropertyHash.Keys | % { #Check the properties we're going to try and update the Drive with are valid:
        if($validProperties -notcontains $_ ){
            $duffProperties += $_
            }
        }

    if($duffProperties.Count -gt 0){
        Write-Error -Message "Property(s) [$($duffProperties -join ", ")] is invalild for Graph List object. Will not attempt to update."
        return
        }

    invoke-graphPatch -tokenResponse $tokenResponse -graphQuery "/sites/$graphSiteId/lists/$graphListId" -graphBodyHashtable $listPropertyHash
    }
function set-graphMailboxSettings(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse
       ,[parameter(Mandatory = $true)]
            [string]$identity
       ,[parameter(Mandatory = $false)]
            [string]$timeZone
         )
If(($identity -match "@anthesisgroup.com") -or ($identity.Length -eq 36)){
#Identity contains a upn or looks like a guid
    If($timeZone){
        If((Get-TimeZone -ListAvailable | Select-Object -Property "Id") -match $timeZone){
        #timezone is available in Windows
        $graphQuery = "users/$identity/mailboxSettings"
        #We set the mailbox timezone and meeting hours timezone the same
        $graphBodyHashtable = [ordered]@{
        workingHours = @{
        timeZone=@{
        "name"="$($timeZone)";
        }      
        }
        "timeZone"= "$($timeZone)"
        }
        $response = invoke-graphPatch -tokenResponse $tokenResponse -graphQuery  $graphQuery -graphBodyHashtable $graphBodyHashtable -Verbose
        $response
        }
        Else{
        Write-Error "Please provide a valid timezone available in Windows"
        }
    }
}
Else{
Write-Error "Please provide a valid upn or guid"
}
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

    $existingSettings = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/groups/$($graphUnifiedGroupExtended.id)/settings"
    if($existingSettings){
        $existingSettings.values | ? {$_.Name -eq "AllowToAddGuests"} | % { #"/groups/$($graphUnifiedGroupExtended.id)/settings" returns a weird object: the .values property is a 0+ array of [PSCustomObject]
            if($_.value -ne $allowToAddGuests){
                #If the wrong AllowToAddGuests settings are in place, fix them and notify IT
                invoke-graphPatch -tokenResponse $tokenResponse -graphQuery "/groups/$($graphUnifiedGroupExtended.id)/settings/$($existingSettings.id)" -graphBodyHashtable $sharingBody
                Write-Warning "AllowToAddGuests changed from [$($_.value)] to [$allowToAddGuests] for Unified Group [$($graphUnifiedGroupExtended.id)][$($graphUnifiedGroupExtended.DisplayName)]"
                Send-MailMessage -Subject "AllowToAddGuests changed from [$($_.value)] to [$sharingSettings] for Unified Group [$($graphUnifiedGroupExtended.id)][$($graphUnifiedGroupExtended.DisplayName)]" -to "kevin.maitland@anthesisgroup.com" -From securitybot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Encoding UTF8 -Priority High
                }
            else{Write-Verbose "AllowToAddGuests are correct for [$($graphUnifiedGroupExtended.id)][$($graphUnifiedGroupExtended.DisplayName)]"}
            
            }
        }
    else{#If there are no AllowToAddGuests settings, just create them
        invoke-graphPost -tokenResponse $tokenResponse -graphQuery "/groups/$($graphUnifiedGroupExtended.id)/settings" -graphBodyHashtable $sharingBody
        }

    }
function set-graphUserManager(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
         [ValidatePattern("@")]
            [string]$userUPN
        ,[parameter(Mandatory = $true)]
         [ValidatePattern("@")]
            [string]$managerUPN
        )

$employeeid = get-graphUsers -tokenResponse $tokenResponse -filterUpns $($userUPN) | Select-Object -Property "id"
$managerid = get-graphUsers -tokenResponse $tokenResponse -filterUpns $($managerUPN) | Select-Object -Property "id"
If(($employeeid) -and ($managerid)){
$body = "{
  `"@odata.id`": `"https://graph.microsoft.com/v1.0/users/$($managerid.id)`"
}"
$graphQuery = "users/$($employeeid.id)" + '/manager/' + "`$ref"
Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/$graphQuery" -Body $body -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Put -Verbose
}
Else{
write-host "User or Manager ID missing" -ForegroundColor Red
}
}
function set-graphUser(){
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
      
    $validProperties = @("accountEnabled","assignedLicenses","assignedPlans","businessPhones","city","companyName","country","createdDateTime","creationType","deletedDateTime","department","displayName","employeeId","faxNumber","givenName","id","identities","imAddresses","isResourceAccount","jobTitle","lastPasswordChangeDateTime","legalAgeGroupClassification","licenseAssignmentStates","mail","mailNickname","manager","mobilePhone","officeLocation","onPremisesDistinguishedName","onPremisesDomainName","onPremisesExtensionAttributes","onPremisesImmutableId","onPremisesLastSyncDateTime","onPremisesProvisioningErrors","onPremisesSamAccountName","onPremisesSecurityIdentifier","onPremisesSyncEnabled","onPremisesUserPrincipalName","otherMails","passwordPolicies","passwordProfile","postalCode","preferredDataLocation","preferredLanguage","provisionedPlans","proxyAddresses","refreshTokensValidFromDateTime","showInAddressList","signInSessionsValidFromDateTime","state","streetAddress","surname","usageLocation","userPrincipalName","userType")
    $dubiousProperties = @("aboutMe","birthday","interests","mailboxSettings","hireDate","mySite","pastProjects","preferredName","responsibilities","schools","skills")
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
    
    invoke-graphPatch -tokenResponse $tokenResponse -graphQuery "/users/$userIdOrUpn" -graphBodyHashtable $userPropertyHash
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
    if($(get-date $tokenResponse.OriginalExpiryTime) -ge $(Get-Date).AddSeconds($renewTokenExpiringInSeconds)){$tokenResponse} #If the token  is still valid, just return it
    else{
        if($renewTokenExpiringInSeconds){
            get-graphTokenResponse -aadAppCreds $aadAppCreds -grant_type client_credentials -verbose #If it's expired (or will expire within the supplied limit), renew it
            }
        else{$false}#Otherwise return False
        }
    }
function update-graphGroupOfDevicesBasedOnOwners(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$userGroupId
        ,[parameter(Mandatory = $true)]
            [string]$devicesGroupId
        ,[parameter(Mandatory = $false)]
            [ValidateSet("Android","iOS","Windows")]
            [string]$deviceType
        )

    $usersInGroup = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $userGroupId -memberType TransitiveMembers -returnOnlyLicensedUsers
    if([string]::IsNullOrWhiteSpace($deviceType)){
        $devicesOwnedByUsers = get-graphDevices -tokenResponse $tokenResponse -filterOwnerIds $usersInGroup.id 
        }
    else{
        $devicesOwnedByUsers = get-graphDevices -tokenResponse $tokenResponse -filterOwnerIds $usersInGroup.id -filterOperatingSystem $deviceType
        }
    
    $devicesCurrentlyInGroup = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $devicesGroupId -memberType TransitiveMembers

    if([string]::IsNullOrWhiteSpace($usersInGroup.Id)){$devicesOwnedByUsers = @()}
    if([string]::IsNullOrWhiteSpace($devicesCurrentlyInGroup.Id)){$devicesCurrentlyInGroup = @()}
    $delta = Compare-Object -ReferenceObject $devicesCurrentlyInGroup -DifferenceObject $devicesOwnedByUsers -Property Id -PassThru

    $toAdd = $delta | ? {$_.SideIndicator -eq "=>"} 
    if($toAdd){
        add-graphUsersToGroup -tokenResponse $tokenResponse -graphGroupId $devicesGroupId -memberType Members -graphUserIds $toAdd.id
        }

    $toRemove = $delta | ? {$_.SideIndicator -eq "<="} 
    if($toRemove){
        remove-graphUsersFromGroup -tokenResponse $tokenResponse -graphGroupId $devicesGroupId -memberType Members -graphUserIds $toRemove.id
        }
    }
function update-graphListItem(){
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
            [string]$listitemId
        ,[parameter(Mandatory = $true)]
            [hashtable]$fieldHash = @{}


        )

    switch ($PsCmdlet.ParameterSetName){
        {$_ -match "URL"} { #If we've got a URL to the Site, we'll need to get the Id
            Write-Verbose "update-graphListItem | Getting Site from URL [$serverRelativeSiteUrl]"
            $graphSiteId = $(get-graphSite -tokenResponse $tokenResponse -serverRelativeUrl $serverRelativeSiteUrl).Id
            }
        {$_ -match "AndName"} { #If we've got a URL to the Site, we'll need to get the Id
            $listId = $(get-graphList -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listName $listName).Id 
            Write-Verbose "update-graphListItem | getting ListId from name [$listName]"
            }
        }
   
    $graphBodyHashtable = $fieldHash
    $reponse = invoke-graphPatch -tokenResponse $tokenResponse -graphQuery "/sites/$graphSiteId/lists/$listId/items/$listitemId/fields" -graphBodyHashtable $graphBodyHashtable -Verbose:$VerbosePreference
    $reponse
}
function update-graphOpenShiftShared(){
        [cmdletbinding()]
        param(
            [parameter(Mandatory = $true)]
                [psobject]$tokenResponse        
            ,[parameter(Mandatory = $true)]
                [string]$teamId
            ,[parameter(Mandatory = $true)]
                [string]$MsAppActsAsUserId
            ,[parameter(Mandatory = $false)]
                [string]$schedulingGroupId
            ,[parameter(Mandatory = $true)]
                [string]$openShiftId
            ,[parameter(Mandatory = $false)]
                [string]$shiftName
            ,[parameter(Mandatory = $false)]
                [string]$shiftNotes
            ,[parameter(Mandatory = $false)]
                [int]$availableSlots
            ,[parameter(Mandatory = $false)]
                [ValidateSet ("White","Blue","Green","Purple","Pink","Yellow","Gray","DarkBlue","DarkGreen","DarkPurple","DarkPink","DarkYellow")]
                [string]$shiftColour
            )

        #Get the Shift - we can't amend the start and end times of a shift as it RECREATES the shift, which means a new ID and subsequently all Shift requests are declined... 
        $currentShift = get-graphShiftOpenShifts -tokenResponse $tokenResponse -teamId $teamId -MsAppActsAsUserId $msAppActsAsUserId | Where-Object -Property "Id" -EQ $openShiftId
        if($currentShift){
        #Swap out anything that's missing with existing information from the Shift
        if([string]::IsNullOrWhiteSpace($shiftName)){$shiftName = $currentShift.sharedOpenShift.displayName}
        if([string]::IsNullOrWhiteSpace($shiftNotes)){$shiftNotes = $currentShift.sharedOpenShift.notes}
        if([string]::IsNullOrWhiteSpace($shiftColour)){$shiftColour = $currentShift.sharedOpenShift.theme}
        if(!$availableSlots){$availableSlots = $currentShift.sharedOpenShift.openSlotCount}

        #Create an already-shared OpenShift object to apply
        $shiftDetails=@{
            displayName=$shiftName
            notes=$shiftNotes
            startDateTime="$($currentShift.sharedOpenShift.startDateTime)"
            endDateTime="$($currentShift.sharedOpenShift.endDateTime)"
            theme=$shiftColour
            openSlotCount=$availableSlots
            }
        $notdraftshift=@{
        draftOpenShift="null"
        }
        $Shift = @{
           schedulingGroupId=$schedulingGroupId
           sharedOpenShift=$shiftDetails
           }
 
        invoke-graphPut -tokenResponse $tokenResponse -graphQuery "/teams/$teamId/schedule/openShifts/$openShiftId" -graphBodyHashtable $Shift -Verbose:$true -additionalHeaders @{"MS-APP-ACTS-AS"=$msAppActsAsUserId}
        }
        else{
        Write-Error -Exception "Couldn't find the Shift with the given teamId and openShiftId" -Message "Shift not found" 
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




