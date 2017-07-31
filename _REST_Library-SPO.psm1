<#
.Synopsis
    Stores the credentials for Invoke-SPORestMethod.
.DESCRIPTION
    Stores the credentials for Invoke-SPORestMethod. This is done so that you
    don't have to provide your credentials on every call to Invoke-SPORestMethod.
.EXAMPLE
   Set-SPORestCredentials
.EXAMPLE
   Set-SPORestCredentials -Credential $cred
#>
#region SPO functions
function global:Set-SPORestCredentials {
    Param (
        [Parameter(ValueFromPipeline = $true)]
        [ValidateNotNull()]
        $Credential = (Get-Credential -Message "Enter your credentials for SharePoint Online:")
    )
    Begin {
        if ((Get-Module Microsoft.Online.SharePoint.PowerShell -ListAvailable) -eq $null) {
            throw "The Microsoft SharePoint Online PowerShell cmdlets have not been installed."
        }
        [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime") | Out-Null
    }
    Process {
        $global:spoCred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credential.UserName, $Credential.Password)
    }
} 
<#
.Synopsis
    Clears the SharePoint Online credentials stored in the global variable.
.DESCRIPTION
    Clears the SharePoint Online credentials stored in the global variable.
    You can also manually clear the variable by explicitly setting 
    $global:spoCred = $null.
.EXAMPLE
   Clear-SPORestCredentials
#>
function global:Clear-SPORestCredentials {
    $global:spoCred = $null
}
<#
.Synopsis
    Sends an HTTP or HTTPS request to a SharePoint Online REST-compliant web service.
.DESCRIPTION
    This function sends an HTTP or HTTPS request to a Representational State 
    Transfer (REST)-compliant ("RESTful") SharePoint Online web service.
    When connecting, if Set-SPORestCredentials is not called then you will be
    prompted for your credentials. Those credentials are stored in a global
    variable $global:spoCred so that it will be available on subsequent calls.
    Call Clear-SPORestCredentials to clear the variable.
.EXAMPLE
   Invoke-SPORestMethod -Url "https://contoso.sharepoint.com/_api/web"
.EXAMPLE
   Invoke-SPORestMethod -Url "https://contoso.sharepoint.com/_api/contextinfo" -Method "Post"
#>
function global:Invoke-SPORestMethod {
    [CmdletBinding()]
    [OutputType([int])]
    Param (
        # The REST endpoint URL to call.
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.Uri]$Url,

        # Specifies the method used for the web request. The default value is "Get".
        [Parameter(Mandatory = $false, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("Get", "Head", "Post", "Put", "Delete", "Trace", "Options", "Merge", "Patch")]
        [string]$Method = "Get",

        # Additional metadata that should be provided as part of the Body of the request.
        [Parameter(Mandatory = $false, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [object]$Metadata,

        # The "X-RequestDigest" header to set. This is most commonly used to provide the form digest variable. Use "(Invoke-SPORestMethod -Url "https://contoso.sharepoint.com/_api/contextinfo" -Method "Post").GetContextWebInformation.FormDigestValue" to get the Form Digest value.
        [Parameter(Mandatory = $false, Position = 3)]
        [ValidateNotNullOrEmpty()]
        [string]$RequestDigest,
        
        # The "If-Match" header to set. Provide this to make sure you are not overwritting an item that has changed since you retrieved it.
        [Parameter(Mandatory = $false, Position = 4)]
        [ValidateNotNullOrEmpty()]
        [string]$ETag, 
        
        # To work around the fact that many firewalls and other network intermediaries block HTTP verbs other than GET and POST, specify PUT, DELETE, or MERGE requests for -XHTTPMethod with a POST value for -Method.
        [Parameter(Mandatory = $false, Position = 5)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("Get", "Head", "Post", "Put", "Delete", "Trace", "Options", "Merge", "Patch")]
        [string]$XHTTPMethod,

        [Parameter(Mandatory = $false, Position = 6)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("Verbose", "MinimalMetadata", "NoMetadata")]
        [string]$JSONVerbosity = "Verbose",

        # If the returned data is a binary data object such as a file from a SharePoint site specify the output file name to save the data to.
        [Parameter(Mandatory = $false, Position = 7)]
        [ValidateNotNullOrEmpty()]
        [string]$OutFile
    )

    Begin {
        if ((Get-Module Microsoft.Online.SharePoint.PowerShell -ListAvailable) -eq $null) {
            throw "The Microsoft SharePoint Online PowerShell cmdlets have not been installed."
        }
        if ($global:spoCred -eq $null) {
            [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime") | Out-Null
            $cred = Get-Credential -Message "Enter your credentials for SharePoint Online:"
        } 

    }
    Process {
        $request = [System.Net.WebRequest]::Create($Url)
        $request.Credentials = $global:spoCred
        $odata = ";odata=$($JSONVerbosity.ToLower())"
        $request.Accept = "application/json$odata"
        $request.ContentType = "application/json;charset=UTF-8$odata"   
        $request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f")   
        $request.Method = $Method.ToUpper()

        if(![string]::IsNullOrEmpty($RequestDigest)) {
            $request.Headers.Add("X-RequestDigest", $RequestDigest)
        }
        if(![string]::IsNullOrEmpty($ETag)) {
            $request.Headers.Add("If-Match", $ETag)
        }
        if($XHTTPMethod -ne $null) {
            $request.Headers.Add("X-HTTP-Method", $XHTTPMethod.ToUpper())
        }
        if ($Metadata -is [string] -and ![string]::IsNullOrEmpty($Metadata)) {
            $body = [System.Text.Encoding]::UTF8.GetBytes($Metadata)
            $request.ContentLength = $body.Length
            $stream = $request.GetRequestStream()
            $stream.Write($body, 0, $body.Length)
        } elseif ($Metadata -is [byte[]] -and $Metadata.Count -gt 0) {
            $request.ContentLength = $Metadata.Length
            $stream = $request.GetRequestStream()
            $stream.Write($Metadata, 0, $Metadata.Length)
        } else {
            $request.ContentLength = 0
        }
 
        $response = $request.GetResponse()
        try {
            $streamReader = New-Object System.IO.StreamReader $response.GetResponseStream()
            try {
                # If the response is a file (a binary stream) then save the file our output as-is.
                if ($response.ContentType.Contains("application/octet-stream")) {
                    if (![string]::IsNullOrEmpty($OutFile)) {
                        $fs = [System.IO.File]::Create($OutFile)
                        try {
                            $streamReader.BaseStream.CopyTo($fs)
                        } finally {
                            $fs.Dispose()
                        }
                        return
                    }
                    $memStream = New-Object System.IO.MemoryStream
                    try {
                        $streamReader.BaseStream.CopyTo($memStream)
                        Write-Output $memStream.ToArray()
                    } finally {
                        $memStream.Dispose()
                    }
                    return
                }
                # We don't have a file so assume JSON data.
                $data = $streamReader.ReadToEnd()
                # In many cases we might get two ID properties with different casing.
                # While this is legal in C# and JSON it is not with PowerShell so the
                # duplicate ID value must be renamed before we convert to a PSCustomObject.
                if ($data.Contains("`"ID`":") -and $data.Contains("`"Id`":")) {
                    $data = $data.Replace("`"ID`":", "`"ID-dup`":");
                }

                $results = ConvertFrom-Json -InputObject $data
                $global:results2 = $results
                # The JSON verbosity setting changes the structure of the object returned.
                if ($JSONVerbosity -ne "verbose" -or $results.d -eq $null) {
                    Write-Output $results
                } else {
                    Write-Output $results.d 
                }
            } finally {
                $streamReader.Dispose()
            }
        } finally {
            $response.Dispose()
        }
    }
    End {
    }
} 
#endregion
#region Ant functions
function check-digestExpiry($serverUrl, $sitePath){
    $sitePath = format-path $sitePath
    if(($digestExpiryTime.AddSeconds(-30) -lt (Get-Date)) -or ($digest.GetContextWebInformation.WebFullUrl -ne $serverUrl+$sitePath)){get-newDigest $serverUrl $sitePath}
    }
function copy-fileInLibrary($sourceSitePath,$sourceLibraryAndFolderPath,$sourceFileName,$destinationSitePath,$destinationLibraryAndFolderPath,$destinationFileName,[boolean]$overwrite){
    check-digestExpiry -serverUrl $serverUrl -sitePath $sitePath  #this needs to be checked for all POST queries
    #$sourceFile = get-fileInLibrary -sitePath $sourceSitePath -libraryAndFolderPath $sourceLibraryAndFolderPath -fileName $sourceFileName
    if(!$destinationSitePath -and !$destinationLibraryAndFolderPath -and !$destinationFileName){$destinationSitePath = $sourceSitePath;$destinationLibraryAndFolderPath = $sourceLibraryAndFolderPath;$destinationFileName = $sourceFileName+"_copy"}
    if(!$destinationSitePath){$destinationSitePath = $sourceSitePath}
    if(!$destinationLibraryAndFolderPath){$destinationLibraryAndFolderPath = $sourceLibraryAndFolderPath}
    if(!$destinationFileName){$destinationFileName = $sourceFileName}
    $sourceSitePath = format-path (sanitise-forSharePointUrl $sourceSitePath)
    $sourceLibraryAndFolderPath = format-path (sanitise-forSharePointUrl $sourceLibraryAndFolderPath)
    $sourceFileName = sanitise-forSharePointFileName $sourceFileName
    $destinationSitePath = format-path (sanitise-forSharePointUrl $destinationSitePath)
    $destinationLibraryAndFolderPath = format-path (sanitise-forSharePointUrl $destinationLibraryAndFolderPath)
    $destinationFileName = sanitise-forSharePointFileName $destinationFileName

    $destinationUrl = $destinationSitePath+$destinationLibraryAndFolderPath+"/"+$destinationFileName
    $url = $serverUrl+$sourceSitePath+"/_api/web/GetFileByServerRelativeUrl('$sourceSitePath$sourceLibraryAndFolderPath/$sourceFileName')/CopyTo('$destinationUrl')"
    $url
    try{
        if($verboseLogging){log-action "Invoke-SPORestMethod -Url $url -Method `"POST`" -RequestDigest $($digest.GetContextWebInformation.FormDigestValue)"}
        Invoke-SPORestMethod -Url $url -Method "POST" -RequestDigest $digest.GetContextWebInformation.FormDigestValue
        if($verboseLogging){log-result "FILE COPIED: $destinationFileName"}
        }
    catch{
        if($verboseLogging){log-error -myError $Error -myFriendlyMessage "Failed to copy-FileInLibrary: Invoke-SPORestMethod -Url $url -Method `"POST`" -RequestDigest $($digest.GetContextWebInformation.FormDigestValue)" -doNotLogToEmail $true}
        $false
        }
    }
function delete-folderInLibrary($sitePath,$libraryName,$folderPathAndNameToBeDeleted){
    #This needs tidying up
    $sitePath = format-path (sanitise-forSharePointUrl $sitePath)
    $folderPathAndNameToBeDeleted = format-path $folderPathAndNameToBeDeleted
    $url = "$serverUrl$sitePath/_api/web/GetFolderByServerRelativeUrl('$sitePath$libraryName$folderPathAndNameToBeDeleted')"
    #$dummy = Invoke-SPORestMethod -Url $url 
    check-digestExpiry -serverUrl $serverUrl -sitePath $sitePath  #this needs to be checked for all POST queries
    try{
        if($verboseLogging){log-action "delete-folderInLibrary: Invoke-SPORestMethod -Url $url -Method `"POST`" -XHTTPMethod `"DELETE`" -RequestDigest $($digest.GetContextWebInformation.FormDigestValue) -ETag `"*`""}
        Invoke-SPORestMethod -Url $url -Method "POST" -XHTTPMethod "DELETE" -RequestDigest $digest.GetContextWebInformation.FormDigestValue -ETag "*"
        if($verboseLogging){log-result "FOLDER DELETED: $sitePath$libraryName$folderPathAndNameToBeDeleted"}
        }
    catch{
        if($verboseLogging){log-error -myError $_ -myFriendlyMessage "Failed to delete-folderInLibrary: Invoke-SPORestMethod -Url $url -Method `"POST`" -XHTTPMethod `"DELETE`" -RequestDigest $($digest.GetContextWebInformation.FormDigestValue) -ETag `"*`"" -doNotLogToEmail $true}
        $false
        }
    }
function format-path($dirtyPath){
    #All "path" variables should be prefixed with a "/", but not suffixed
    if($dirtyPath.Substring(0,1) -ne "/"){$dirtyPath = "/"+$dirtyPath}
    if($dirtyPath.Substring(($dirtyPath.Length-1),1) -eq "/"){$dirtyPath = $dirtyPath.Substring(0,$dirtyPath.Length-1)}
    $dirtyPath
    }
function get-fileInLibrary($sitePath, $libraryAndFolderPath, $fileName){
    #Sanitise parameters
    $sitePath = format-path $sitePath
    $libraryAndFolderPath = format-path $libraryAndFolderPath
    $fileName = format-path (sanitise-forSharePointFileName $fileName)
    $sanitisedPath = sanitise-forResourcePath $sitePath$libraryAndFolderPath$fileName
    #Build and execute REST statement
    $url = $serverUrl+$sitePath+"/_api/web/GetFileByServerRelativePath(decodedUrl='$sanitisedPath')"
    try{
        if($verboseLogging){log-action "get-fileInLibrary: Invoke-SPORestMethod -Url $url"}
        Invoke-SPORestMethod -Url $url
        if($verboseLogging){log-result "SUCCESS: File found in Library"}
        }
    catch{
        if($verboseLogging){log-error -myError $_ -myFriendlyMessage "Failed to get-fileInLibrary: Invoke-SPORestMethod -Url $url" -doNotLogToEmail $true}
        $false
        }
    }
function get-folderInLibrary($sitePath, $libraryName, $folderPathAndOrName){
    #Sanitise parameters
    $sitePath = format-path (sanitise-forSharePointUrl $sitePath)
    $libraryName = format-path (sanitise-forSharePointUrl $libraryName)
    $folderPathAndOrName = format-path ($folderPathAndOrName)
    #$libraryAndFolderPath = format-path (sanitise-forSharePointUrl  $libraryAndFolderPath)
    #$folderName = sanitise-forSharePointFileName $folderName
    $sanitisedPath = "decodedurl='"+(sanitise-forResourcePath $sitePath$libraryName$folderPathAndOrName)+"'"
    #Build and execute REST statement
    #$url = $serverUrl+$sitePath+"/_api/web/GetFolderByServerRelativeUrl('$sitePath$libraryAndFolderPath/$folderName"+"')"
    $url = $serverUrl+$sitePath+"/_api/web/GetFolderByServerRelativePath($sanitisedPath)"
    try{
        if($verboseLogging){log-action "get-folderInLibrary: Invoke-SPORestMethod -Url $url"}
        Invoke-SPORestMethod -Url $url
        if($verboseLogging){log-result "SUCCESS:`tFolder in Library found"}
        }
    catch{
        if($verboseLogging){log-error -myError $_ -myFriendlyMessage "Failed to get-folderInLibrary: Invoke-SPORestMethod -Url $url" -doNotLogToEmail $true}
        $false
        }
    }
function get-itemInListFromProperty($sitePath, $listName, $propertyName, $propertyValue){
    $sitePath = format-path (sanitise-forSharePointUrl $sitePath)
    $listName = sanitise-forSharePointUrl $listName
    $query = "?filter=$propertyName eq $propertyValue"
    #Build and execute REST statement
    $url = $serverUrl+$sitePath+"/_api/web/Lists/GetByTitle('$listName')/items"
    try{
        if($verboseLogging){log-action "get-itemInListFromProperty: Invoke-SPORestMethod -Url ($url$query)"}
        $item = Invoke-SPORestMethod -Url ($url+$query)
        if($item){
            if($verboseLogging){log-result "FOUND ITEM IN LIST FROM PROPERTY"}
            $item.results
            }
        else{
            if($verboseLogging){log-result -myFriendlyMessage "WARNING: get-itemInListFromProperty($sitePath, $listName, $propertyName, $propertyValue) returned no items"}
            $false
            }
        }
    catch{
        if($verboseLogging){log-error -myError $_ -myFriendlyMessage "get-itemInListFromProperty($sitePath, $listName, $propertyName, $propertyValue) failed to execute" -doNotLogToEmail $true}
        $false
        }
    }    
function get-itemsInList($sitePath, $listName, $oDataQuery, $suppressProgress){
    #Sanitise parameters
    $sitePath = format-path (sanitise-forSharePointUrl $sitePath)
    $listName = sanitise-forSharePointUrl $listName
    if($oDataQuery){if($oDataQuery.SubString(0,1) -ne "?"){$oDataQuery = "?$oDataQuery"}} #Prefix with ? if user hasn't done so already
    #Build the query
    $url = $serverUrl+$sitePath+"/_api/web/Lists/GetByTitle('$listName')/items$oDataQuery"
    #Run the query
    try{
        if($verboseLogging){log-action "get-itemsInList: Invoke-SPORestMethod -Url $url"}
        $partialItems = Invoke-SPORestMethod -Url $url
        if($partialItems){
            if($verboseLogging){log-result "SUCCESS: Initial $($partialItems.results.Count) items returned"}
            $queryResults = $partialItems.results
            }
        else{
            if($verboseLogging){log-result -myFriendlyMessage "WARNING: get-itemsInList($sitePath, $listName) returned no items"}
            $false
            }
        }
    catch{
        if($verboseLogging){log-error -myError $_ -myFriendlyMessage "get-itemsInList($sitePath, $listName) failed to execute" -doNotLogToEmail $true}
        $false
        }
    $i=$partialItems.results.Count
    #Check for additional results
    while($partialItems.__next){
        try{
            if($verboseLogging){log-action "get-itemsInList: Invoke-SPORestMethod -Url $($partialItems.__next)"}
            $partialItems = Invoke-SPORestMethod -Url $partialItems.__next
            if($partialItems){
                if($verboseLogging){log-result "SUCCESS: Subsequent $($partialItems.results.Count) items returned"}
                $queryResults += $partialItems.results
                }
            else{
                if($verboseLogging){log-result "WARNING: get-itemsInList($sitePath, $listName) returned no items"}
                $false
                }
            }
        catch{
            if($verboseLogging){log-error -myError $_ -myFriendlyMessage "get-itemsInList($sitePath, $listName) failed to execute"}
            $false
            }
        $i+=$partialItems.results.Count
        if(!$suppressProgress){Write-Host $i}
        }
    $queryResults
    }
function get-library($sitePath, $libraryName){
    #Sanitise parameters
    $sitePath = format-path (sanitise-forSharePointUrl $sitePath)
    $libraryName = format-path (sanitise-forSharePointUrl $libraryName) #The LibraryName cannot contain specific characters
    #Build and execute REST statement
    $url = $serverUrl+$sitePath+"/_api/web/GetFolderByServerRelativePath(decodedurl='$sitePath$libraryName')"
    try{
        if($verboseLogging){log-action "get-library: Invoke-SPORestMethod -Url $url"}
        Invoke-SPORestMethod -Url $url
        if($verboseLogging){log-result "SUCCESS: Library found"}
        }
    catch{
        if($verboseLogging){log-error -myError $_ -myFriendlyMessage "Failed to get-library($sitePath, $libraryName)" -doNotLogToEmail $true}
        $false
        }
    }
function get-list($sitePath, $listName){
    #Sanitise parameters
    $sitePath = format-path (sanitise-forSharePointUrl $sitePath)
    $listName = (sanitise-forSharePointUrl $listName).Replace("Lists/","")
    #Build and execute REST statement
    $url = $serverUrl+$sitePath+"/_api/web/Lists/GetByTitle('$listName')"
    try{
        if($verboseLogging){log-action "get-list: Invoke-SPORestMethod -Url $url"}
        Invoke-SPORestMethod -Url $url
        if($verboseLogging){log-result "SUCCESS: List found"}
        }
    catch{
        if($verboseLogging){log-error -myError $_ -myFriendlyMessage "Failed to get-list: Invoke-SPORestMethod -Url $url" -doNotLogToEmail $true}
        $false
        }
    }
function get-newDigest($serverUrl, $sitePath){
    $global:digest = (Invoke-SPORestMethod -Url "$serverUrl$sitePath/_api/contextinfo" -Method "POST")#.GetContextWebInformation.FormDigestValue
    $global:digestExpiryTime = (Get-Date).AddSeconds($global:digest.GetContextWebInformation.FormDigestTimeoutSeconds)
    }
function new-folderInLibrary($sitePath, $libraryName, $folderPathAndOrName){
    #$libraryName = $kimbleClientHashTable[$dirtyProject.KimbleClientId]
    #$libraryName = "Shared Documents"
    #$folderPathAndOrName = $dirtyProject.Title
    #$folderPathAndOrName = "/Håkon''s test @naughty folder Name!/Tes()t&2"
    #Sanitise parameters
    $sitePath = format-path (sanitise-forSharePointUrl $sitePath)
    $libraryName = format-path (sanitise-forSharePointUrl $libraryName)
    $folderPathAndOrName = format-path $folderPathAndOrName
    $sanitisedPath = "decodedurl='"+(sanitise-forResourcePath ($sitePath+$libraryName+$folderPathAndOrName))+"'"
    #Prepare security and log action
    check-digestExpiry -serverUrl $serverUrl -sitePath $sitePath  #this needs to be checked for all POST queries
    #Build and execute REST statement
    $url = $serverUrl+$sitePath+"/_api/web/folders/AddUsingPath($sanitisedPath)"
    $folderExists = (get-folderInLibrary -sitePath $sitePath -libraryName $libraryName -folderPathAndOrName $folderPathAndOrName)
    if($folderExists -eq $false){
        try{
            if($verboseLogging){log-action -myMessage "new-folderInLibrary: Invoke-SPORestMethod -Url $url -Method `"POST`" -RequestDigest $($digest.GetContextWebInformation.FormDigestValue)"}
            Invoke-SPORestMethod -Url $url -Method "POST" -RequestDigest $digest.GetContextWebInformation.FormDigestValue
            if($verboseLogging){log-result "SUCCESS: Created folder $sitePath$libraryName$folderPathAndOrName"}
            }
        catch{
            if($verboseLogging){log-error -myError $_ -myFriendlyMessage "Failed to create new folder: new-folderInLibrary($sitePath, $libraryAndFolderPath, $folderName)" -doNotLogToEmail $true}
            $false
            }
        }
    else{
        if($verboseLogging){log-result "WARNING: Folder already exists: $sitePath$libraryName$folderPathAndOrName"}
        $folderExists
        }
    }
function new-itemInList($sitePath,$listName,$predeterminedItemType,$hashTableOfItemData){
    #Error handling for no DataType?
    #Sanitise parameters
    $sitePath = format-path (sanitise-forSharePointUrl $sitePath)
    $listName = sanitise-forSharePointUrl(sanitise-forSharePointFileName ($listName.Replace("Lists/","")))
    #Prepare security and log action
    check-digestExpiry -serverUrl $serverUrl -sitePath $sitePath  #this needs to be checked for all POST queries
    log-action "new-itemInList($sitePath,$listName,$predeterminedItemType,$($hashTableOfItemData.Keys | %{"$_=$($hashTableOfItemData[$_]);"})"
    #Build and execute REST statement
    $url = $serverUrl+$sitePath+"/_api/web/Lists/GetByTitle('$listName')/items"
    foreach($key in $hashTableOfItemData.Keys){
        $formattedItemData += "`'$key`':`"$($hashTableOfItemData[$key])`", "
        }
    $formattedItemData = $formattedItemData.Substring(0,$formattedItemData.Length-2) #Trim off the final ","
    $metadata = "{ '__metadata': { 'type': '$predeterminedItemType' }, $formattedItemData}"
    try{
        if($verboseLogging){log-action "Invoke-SPORestMethod -Url $url -Method `"POST`" -Metadata $metadata -RequestDigest $($digest.GetContextWebInformation.FormDigestValue)"}
        Invoke-SPORestMethod -Url $url -Method "POST" -Metadata $metadata -RequestDigest $digest.GetContextWebInformation.FormDigestValue
        if($verboseLogging){log-result "SUCCESS: New item created"}
        }
    catch{
        if($verboseLogging){log-error $_ -myFriendlyMessage "Failed to create new list item: new-itemInList($sitePath,$listName,$predeterminedItemType,$($hashTableOfItemData.Keys | %{"$_=$($hashTableOfItemData[$_]);"})" -doNotLogToEmail $true}
        $false
        }
    }
function new-library($sitePath, $libraryName, $libraryDesc){
    #Sanitise parameters
    $sitePath = format-path (sanitise-forSharePointUrl $sitePath)
    $libraryName = sanitise-forSharePointFileName $libraryName
    #Prepare security and log action
    check-digestExpiry -serverUrl $serverUrl -sitePath $sitePath  #this needs to be checked for all POST queries
    #Build and execute REST statement
    $metadata = "{'__metadata':{'type':'SP.List'},'Description':`"$libraryDesc`",'BaseTemplate':101,'ContentTypesEnabled':true,'Title':`"$libraryName`",'AllowContentTypes':true}"
    $url = $serverUrl+$sitePath+"/_api/web/lists"
    $libraryExists = get-library -sitePath $sitePath -libraryName $libraryName
    if($libraryExists -eq $false){
        try{
            if($verboseLogging){log-action -myMessage "new-library: Invoke-SPORestMethod -Url $url -Method `"POST`" -Metadata $metadata -RequestDigest $($digest.GetContextWebInformation.FormDigestValue)"}
            Invoke-SPORestMethod -Url $url -Method "POST" -Metadata $metadata -RequestDigest $digest.GetContextWebInformation.FormDigestValue
            if($verboseLogging){log-result "SUCCESS: Library created: $sitePath/$libraryName"}
            }
        catch{
            if($verboseLogging){log-error -myError $_ -myFriendlyMessage "Failed to create new Library: new-library($sitePath, $libraryName, $libraryDesc)" -doNotLogToEmail $true}
            $false
            }
        }
    else{
        if($verboseLogging){log-result "WARNING: Library already exists: $sitePath/$libraryName"}
        $libraryExists
        }
    }
function update-list($sitePath, $listName,$hashTableOfUpdateData){
    #Sanitise parameters
    $sitePath = format-path (sanitise-forSharePointUrl $sitePath)
    $listName = sanitise-forSharePointUrl $listName
    #Prepare security and log action
    check-digestExpiry -serverUrl $serverUrl -sitePath $sitePath  #this needs to be checked for all POST queries
    #Build and execute REST statement
    $url = $serverUrl+$sitePath+"/_api/web/Lists/GetByTitle('$listName')"
    foreach($key in $hashTableOfUpdateData.Keys){
        $formattedItemData += "`'$key`':`'$($hashTableOfUpdateData[$key])`', "
        }
    $formattedItemData = $formattedItemData.Substring(0,$formattedItemData.Length-2) #Trim off the final ","
    #$metadata = "{ '__metadata': { 'type': '$predeterminedItemType' }, $formattedItemData}"
    $metadata = "{'__metadata':{'type':'SP.List'},$formattedItemData}"
    try{
        if($verboseLogging){log-action "update-list: Invoke-SPORestMethod -Url $url -Method `"POST`" -XHTTPMethod `"MERGE`" -Metadata $metadata -RequestDigest $($digest.GetContextWebInformation.FormDigestValue) -ETag `"*`""}
        Invoke-SPORestMethod -Url $url -Method "POST" -XHTTPMethod "MERGE" -Metadata $metadata -RequestDigest $digest.GetContextWebInformation.FormDigestValue -ETag "*"
        if($verboseLogging){log-result "SUCCESS: List updated: $formattedItemData"}
        }
    catch{
        if($verboseLogging){log-error $_ -myFriendlyMessage "Failed to update-list($sitePath, $listName,$($hashTableOfUpdateData.Keys | %{"$_=$($hashTableOfUpdateData[$_]);"})" -doNotLogToEmail $true}
        $false
        }
    }
function update-itemInList($serverUrl,$sitePath,$listName,$predeterminedItemType,$itemId,$hashTableOfItemData){
    #Sanitise parameters
    $sitePath = format-path (sanitise-forSharePointUrl $sitePath)
    $listName = sanitise-forSharePointUrl(sanitise-forSharePointFileName ($listName.Replace("Lists/","")))
    #Prepare security and log action
    check-digestExpiry -serverUrl $serverUrl -sitePath $sitePath
    #Build and execute REST statement
    $url = $serverUrl+$sitePath+"/_api/web/Lists/GetByTitle('$listName')/items($itemId)"
    foreach($key in $hashTableOfItemData.Keys){
        $formattedItemData += "`'$key`':`"$($hashTableOfItemData[$key])`", "
        }
    $formattedItemData = $formattedItemData.Substring(0,$formattedItemData.Length-2) #Trim off the final ","
    $metadata = "{ '__metadata': { 'type': '$predeterminedItemType' }, $formattedItemData}"
    try{
        if($verboseLogging){log-action "Invoke-SPORestMethod -Url $url -Method `"POST`" -XHTTPMethod `"MERGE`" -Metadata $metadata -RequestDigest $($digest.GetContextWebInformation.FormDigestValue) -ETag `"*`""}
        Invoke-SPORestMethod -Url $url -Method "POST" -XHTTPMethod "MERGE" -Metadata $metadata -RequestDigest $digest.GetContextWebInformation.FormDigestValue -ETag "*"
        if($verboseLogging){log-result "SUCCESS: Updated list item: $formattedItemData"}
        }
    catch{
        if($verboseLogging){log-error $_ -myFriendlyMessage "Failed to update item in List: update-itemInList($sitePath,$listName,$predeterminedItemType,$itemId,$($hashTableOfItemData.Keys | %{"$_=$($hashTableOfItemData[$_]);"})" -doNotLogToEmail $true}
        $false
        }
    }
function log-action($myMessage, $doNotLogToFile, $doNotLogToScreen){
    if($logActions){
        if(!$doNotLogToFile -or $logToFile){Add-Content -Value ((Get-Date -Format "yyyy-MM-dd HH:mm:ss")+"`t$myMessage") -Path $logfile}
        if(!$doNotLogToScreen -or $logToScreen){Write-Host -ForegroundColor Yellow $myMessage}
        }
    }
function log-error($myError, $myFriendlyMessage, $doNotLogToFile, $doNotLogToScreen, $doNotLogToEmail, $overrideErrorLogFile){
    if($logErrors){
        if($overrideErrorLogFile){$myErrorLogFile = $overrideErrorLogFile} else{$myErrorLogFile = $logFile}
        if(!$doNotLogToFile -or $logToFile){Add-Content -Value "`t`tERROR:`t$myFriendlyMessage" -Path $myErrorLogFile}
        if(!$doNotLogToFile -or $logToFile){Add-Content -Value "`t`t$($myError.Exception.Message)" -Path $myErrorLogFile}
        if(!$doNotLogToScreen -or $logToScreen){Write-Host -ForegroundColor Red $myFriendlyMessage}
        if(!$doNotLogToScreen -or $logToScreen){Write-Host -ForegroundColor Red $myError}
        if(!$doNotLogToEmail -or $logErrorsToEmail){Send-MailMessage -To $mailTo -From $mailFrom -Subject "Error in update-SpoClientsFolders - $myFriendlyMessage" -Body $myError -SmtpServer $smtpServer}
        }
    }
function log-result($myMessage, $doNotLogToFile, $doNotLogToScreen){
    if($logResults){
        if(!$doNotLogToFile -or $logToFile){Add-Content -Value ("`t$myMessage") -Path $logfile}
        if(!$doNotLogToScreen -or $logToScreen){Write-Host -ForegroundColor DarkYellow "`t$myMessage"}
        }
    }
function sanitise-forSharePointFileName($dirtyString){ 
    $dirtyString = $dirtyString.Trim()
    $dirtyString.Replace("`"","").Replace("#","").Replace("%","").Replace("?","").Replace("<","").Replace(">","").Replace("\","").Replace("/","").Replace("...","").Replace("..","").Replace("'","`'")
    if(@("."," ") -contains $dirtyString.Substring(($dirtyString.Length-1),1)){$dirtyString = $dirtyString.Substring(0,$dirtyString.Length-1)} #Trim trailing "."
    }
function sanitise-forSharePointUrl($dirtyString){ 
    $dirtyString = $dirtyString.Trim()
    $dirtyString = $dirtyString -creplace '[^a-zA-Z0-9 _/]+', ''
    #$dirtyString = $dirtyString.Replace("`"","").Replace("#","").Replace("%","").Replace("?","").Replace("<","").Replace(">","").Replace("\","/").Replace("//","/").Replace(":","")
    #$dirtyString = $dirtyString.Replace("$","`$").Replace("``$","`$").Replace("(","").Replace(")","").Replace("-","").Replace(".","").Replace("&","").Replace(",","").Replace("'","").Replace("!","")
    $cleanString =""
    for($i= 0;$i -lt $dirtyString.Split("/").Count;$i++){ #Examine each virtual directory in the URL
        if($i -gt 0){$cleanString += "/"}
        if($dirtyString.Split("/")[$i].Length -gt 50){$tempString = $dirtyString.Split("/")[$i].SubString(0,50)} #Truncate long folder names to 50 characters
            else{$tempString = $dirtyString.Split("/")[$i]}
        if($tempString.Length -gt 0){
            if(@(".", " ") -contains $tempString.Substring(($tempString.Length-1),1)){$tempString = $tempString.Substring(0,$tempString.Length-1)} #Trim trailing "." and " ", even if this results in a truncation <50 characters
            }
        $cleanString += $tempString
        }
    $cleanString = $cleanString.Replace("//","/").Replace("https/","https://") #"//" is duplicated to catch trailing "/" that might now be duplicated. https is an exception that needs specific handling
    $cleanString
    }
function sanitise-forResourcePath($dirtyString){
    if(@("."," ") -contains $dirtyString.Substring(($dirtyString.Length-1),1)){$dirtyString = $dirtyString.Substring(0,$dirtyString.Length-1)} #Trim trailing "."
    $dirtyString = $dirtyString.trim().replace("`'","`'`'")
    $dirtyString = $dirtyString.replace("#","").replace("%","") #As of 2017-05-26, these characters are not supported by SharePoint (even though https://msdn.microsoft.com/en-us/library/office/dn450841.aspx suggests they should be)
    #$dirtyString = $dirtyString -creplace "[^a-zA-Z0-9 _/()`'&-@!]+", '' #No need to strip non-standard characters
    #[uri]::EscapeUriString($dirtyString) #No need to encode the URL
    $dirtyString
    }
#endregion

get-help Invoke-WebRequest 