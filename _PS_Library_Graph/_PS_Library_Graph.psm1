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
        ,[parameter(Mandatory = $false,ParameterSetName = "groupId")]
            [parameter(Mandatory = $false,ParameterSetName = "groupUpn")]
            [switch]$includeTransitiveMembers = $false
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
            $groupId = $graphGroup.id
            }
        }
    if($includeTransitiveMembers){$memberType = "transitiveMembers"}
    else{$memberType = "members"}
    if($returnOnlyLicensedUsers){
        $refiner = "?`$select=id,displayName,jobTitle,mail,userPrincipalName,assignedLicenses"
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
            $allUsers = $allMembers | ? {$_.'@odata.type' -eq "#microsoft.graph.user"} | Sort-Object userPrincipalName -Unique
            $allUsers
            }
        else{
            Write-Verbose "Returning all Users"
            $allLicensedUsers | ? {$_.'@odata.type' -eq "#microsoft.graph.user" -and $_.assignedLicenses.Count -gt 0} | Sort-Object userPrincipalName -Unique
            $allLicensedUsers
            }
        }
    else{
        Write-Verbose "Returning all Members"
        $allMembers
        }
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
        )
    $sanitisedGraphQuery = $graphQuery.Trim("/")
    do{
        Write-Verbose "https://graph.microsoft.com/v1.0/$sanitisedGraphQuery"
        $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/$sanitisedGraphQuery" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method GET
        Write-Verbose "[$($response.value.count)] results returned on this cycle, [$($results.count)] in total"
        $results += $response.value
        if($firstPageOnly){break}
        if(![string]::IsNullOrWhiteSpace($response.'@odata.nextLink')){$sanitisedGraphQuery = $response.'@odata.nextLink'.Replace("https://graph.microsoft.com/v1.0/","")}
        }
    #while($response.value.count -gt 0)
    while($response.'@odata.nextLink')
    $results
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

