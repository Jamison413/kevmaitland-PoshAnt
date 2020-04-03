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
            [parameter(Mandatory = $true,ParameterSetName = "fromId")]
            [parameter(Mandatory = $true,ParameterSetName = "fromUpn")]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true,ParameterSetName = "fromUrl")]
            [string]$siteUrl
        ,[parameter(Mandatory = $true,ParameterSetName = "fromId")]
            [string]$siteGraphId
        ,[parameter(Mandatory = $true,ParameterSetName = "fromUpn")]
            [ValidatePattern("@")]
            [string]$teamUpn
        ,[parameter(Mandatory = $false,ParameterSetName = "fromUrl")]
            [parameter(Mandatory = $false,ParameterSetName = "fromId")]
            [parameter(Mandatory = $false,ParameterSetName = "fromUpn")]
            [switch]$returnOnlyDefaultDocumentsLibrary
        )
    
    switch ($PsCmdlet.ParameterSetName){ #We need $siteGraphId, so get it from any other parameter supplied
        "fromUpn" {
            Write-Verbose "get-graphDrives | Getting from Team UPN"
            $groupId = (get-graphGroupFromUpn -tokenResponse $tokenResponse -groupUpn $teamUpn).id
            $drives = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/groups/$groupId/drives"
            }
        {@("fromUrl","fromId") -contains $_} {
            Write-Verbose "get-graphDrives | Getting from $_"
            if([string]::IsNullOrWhiteSpace($siteGraphId)){
                if($siteUrl -match "anthesisllc.sharepoint.com"){$siteUrl = ($siteUrl -Split "anthesisllc.sharepoint.com")[1].Trim("/")} #Get the serverRelativeUrl
                $siteGraphId = (invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/sites/anthesisllc.sharepoint.com:/$siteUrl").id
                if([string]::IsNullOrWhiteSpace($siteGraphId)){ #Weirdly this doesn't seem to work, despite the same query being submitted to graph.
                    Write-Verbose "Weird, that should have worked. trying again"
                    $siteGraphId = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0//sites/anthesisllc.sharepoint.com:/$siteUrl" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method GET).id
                    }
                }
            $drives = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/sites/$($siteGraphId)/drives"
            }
        }

    if($returnOnlyDefaultDocumentsLibrary){
        $drives | Sort-Object -Property createdDateTime | Select-Object -First 1 #Select the oldest DocLib, regardless of Name
        }
    else{$drives}

    }
function get-graphGroupFromUpn(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true,ParameterSetName = "GraphOnly")]
            [parameter(Mandatory = $true,ParameterSetName = "Graph&Exchange")]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true,ParameterSetName = "GraphOnly")]
            [parameter(Mandatory = $true,ParameterSetName = "Graph&Exchange")]
            [ValidatePattern("@")]
            [string]$groupUpn
        ,[parameter(Mandatory = $true,ParameterSetName = "Graph&Exchange")]
            [switch]$returnCustomAttributes
        ,[parameter(Mandatory = $false,ParameterSetName = "Graph&Exchange")]
            [pscredential]$exoCreds
        )

    try{
        $graphGroup = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "groups/?`$filter=mail+eq+'$groupUpn'"
        }
    catch{
        Write-Error "Error retrieving Graph Group by UPN in get-graphUsersFromGroup()"
        Throw $_ #Terminate on this error
        }
    if(!$graphGroup){
        Write-Error "Could not retrieve Graph Group using UPN [$groupUpn]. Check the UPN is valid and try again."
        break
        }
    if($returnCustomAttributes){ #The CustomAttribute properties aren't exposed by the Graph API, so we need to revert to the EXO Cmdlets if they are required
        if($graphGroup.groupTypes -contains "Unified"){ #This will only work for a UnifiedGroup, so there's no point in trying with AAD/Exchange groups
            connect-ToExo -credential $exoCreds
            $ug = Get-UnifiedGroup -Identity $groupUpn
            $ug.psobject.Properties | ? {$_.Name -match "CustomAttribute"} | % {
                $graphGroup | Add-Member -MemberType NoteProperty -Name $_.Name -Value $_.Value
                }
            }
        else{Write-Warning "[$groupUpn] is not a Unified Group - cannot return CustomAttributes for it."}
        }

    $graphGroup
    }
function get-graphOwnersFromGroup(){
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
        )
    switch ($PsCmdlet.ParameterSetName){
        “groupUpn”  {
            Write-Verbose "We've been given a GroupUPN, so we need the GroupId"
            $graphGroup = get-graphGroupFromUpn -tokenResponse $tokenResponse -groupUpn $groupUpn -Verbose:$VerbosePreference
            if(!$graphGroup){
                Write-Error "Could not retrieve Graph Group using UPN [$groupUpn]. Check the UPN is valid and try again."
                break
                }
            $groupId = $graphGroup.id
            Write-Verbose "[$groupUpn] Id is [$groupId]"
            }
        }
    
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
                Scope         = "https://graph.microsoft.com/.default"
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
            [switch]$filterLicensedUsers = $false
        ,[parameter(Mandatory = $false)]
            [switch]$selectAllProperties = $false
        )

    #We need the GroupId, so if we were only given the UPN, we need to find the Id from that.
    if($filterUsageLocation){
        $filter += "and usageLocation eq '$filterUsageLocation'"
        }
    if($filterUpn){
        $filter += "and userPrincipalName eq '$filterUpn'"
        }
    if($returnOnlyLicensedUsers){
        $select = ",id,displayName,jobTitle,mail,userPrincipalName,usageLocation,assignedLicenses"
        }
    if($selectAllProperties){
        $select = ",id,id,displayName,givenName,surname,jobTitle,userPrincipalName,mail,mobilePhone,officeLocation,postalCode,usageLocation,preferredLanguage,assignedLicenses"
        }

    #Build the refiner based on the parameters supplied
    if(![string]::IsNullOrWhiteSpace($select) -and $select.StartsWith(",")){$select = $select.Substring(1,$select.Length-1)}
    if($select){$select = "`$select=$select"}
    if(![string]::IsNullOrWhiteSpace($filter) -and $filter.StartsWith("and ")){$filter = $filter.Substring(4,$filter.Length-4)}
    if($filter){$filter = "`$filter=$filter"}

    $refiner = $null
    if($filter -or $select){$refiner = "?"}
    if($select){
        if($refiner.Length -gt 1){$refiner = $refiner+"&"}
        $refiner = $refiner+$select
        }
    if($filter){
        if($refiner.Length -gt 1){$refiner = $refiner+"&"}
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
            $graphGroup = get-graphGroupFromUpn -tokenResponse $tokenResponse -groupUpn $groupUpn -Verbose:$VerbosePreference
            if(!$graphGroup){
                Write-Error "Could not retrieve Graph Group using UPN [$groupUpn]. Check the UPN is valid and try again."
                break
                }
            $groupId = $graphGroup.id
            Write-Verbose "[$groupUpn] Id is [$groupId]"
            }
        }
    if($returnOnlyLicensedUsers){
        $refiner = "?`$select=id,displayName,jobTitle,mail,userPrincipalName,usageLocation,assignedLicenses"
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
        $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/$sanitisedGraphQuery" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method GET
        if($response.value){
            $results += $response.value
            Write-Verbose "[$([int]$response.value.count)] results returned on this cycle, [$([int]$results.count)] in total"
            }
        elseif([string]::IsNullOrWhiteSpace($response)){
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
        
    $graphBodyJson = ConvertTo-Json -InputObject $graphBodyHashtable
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
                break
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
            $graphGroup = get-graphGroupFromUpn -tokenResponse $tokenResponse -groupUpn $groupUpn -Verbose:$VerbosePreference -returnCustomAttributes -exoCreds $exoCreds
            if(!$graphGroup){
                Write-Error "Could not retrieve Graph Group using UPN [$groupUpn]. Check the UPN is valid and try again."
                break
                }
            }
        "groupObject" {
            if($graphGroup.psobject.Properties.Name -notcontains "CustomAttribute1"){
                Write-Verbose "We've been given a Group object, but it's missing the CustomAttributes"
                $graphGroup = get-graphGroupFromUpn -tokenResponse $tokenResponse -groupUpn $groupUpn -Verbose:$VerbosePreference -returnCustomAttributes -exoCreds $exoCreds
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