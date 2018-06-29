#region SPO functions
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
function Invoke-SPORestMethod {
    [CmdletBinding()]
    [OutputType([int])]
    Param (
        # The REST endpoint URL to call.
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.Uri]$Url,

        # The credentials used to authenticate.
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [Microsoft.SharePoint.Client.SharePointOnlineCredentials]$credentials,

        # Specifies the method used for the web request. The default value is "Get".
        [Parameter(Mandatory = $false, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("Get", "Head", "Post", "Put", "Delete", "Trace", "Options", "Merge", "Patch")]
        [string]$Method = "Get",

        # Additional metadata that should be provided as part of the Body of the request.
        [Parameter(Mandatory = $false, Position = 3)]
        [ValidateNotNullOrEmpty()]
        [object]$Metadata,

        # The "X-RequestDigest" header to set. This is most commonly used to provide the form digest variable. Use "(Invoke-SPORestMethod -Url "https://contoso.sharepoint.com/_api/contextinfo" -Method "Post").GetContextWebInformation.FormDigestValue" to get the Form Digest value.
        [Parameter(Mandatory = $false, Position = 4)]
        [ValidateNotNullOrEmpty()]
        [string]$RequestDigest,
        
        # The "If-Match" header to set. Provide this to make sure you are not overwritting an item that has changed since you retrieved it.
        [Parameter(Mandatory = $false, Position = 5)]
        [ValidateNotNullOrEmpty()]
        [string]$ETag, 
        
        # To work around the fact that many firewalls and other network intermediaries block HTTP verbs other than GET and POST, specify PUT, DELETE, or MERGE requests for -XHTTPMethod with a POST value for -Method.
        [Parameter(Mandatory = $false, Position = 6)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("Get", "Head", "Post", "Put", "Delete", "Trace", "Options", "Merge", "Patch")]
        [string]$XHTTPMethod,

        [Parameter(Mandatory = $false, Position = 7)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("Verbose", "MinimalMetadata", "NoMetadata")]
        [string]$JSONVerbosity = "Verbose",

        # If the returned data is a binary data object such as a file from a SharePoint site specify the output file name to save the data to.
        [Parameter(Mandatory = $false, Position = 8)]
        [ValidateNotNullOrEmpty()]
        [string]$OutFile,

        # Override the default timeout for long runnign queries
        [Parameter(Mandatory = $false, Position = 9)]
        [ValidateNotNullOrEmpty()]
        [long]$manualTimeOut
    )

    Begin {
        if ((Get-Module Microsoft.Online.SharePoint.PowerShell -ListAvailable) -eq $null) {
            throw "The Microsoft SharePoint Online PowerShell cmdlets have not been installed."
        }
        if ($credentials -eq $null) {
            [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime") | Out-Null
            $credentials = Get-Credential -Message "Enter your credentials for SharePoint Online:"
        } 

    }
    Process {
        $request = [System.Net.WebRequest]::Create($Url)
        $request.Credentials = $credentials
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
        if ($manualTimeOut -ne 0){$request.Timeout = $manualTimeOut}

        $global:dummy= $request
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
function add-attachmentToListItem($serverUrl,$sitePath,$listItem,$filePathAndName,$restCreds,$digest,$verboseLogging,$logFile){
    if($verboseLogging){Write-Host -ForegroundColor Yellow "add-attachmentToListItem -serverUrl $serverUrl -sitePath $sitePath -listItem $($listItem.__metadata.uri) -filePathAndName $filePathAndName -restCreds $restCreds"}
    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`$digest = $($digest.digest.GetContextWebInformation.FormDigestValue))"}
    $digest = check-digestExpiry -serverUrl $serverUrl -sitePath $sitePath -digest $digest -restCreds $restCreds -logFile $logFile -verboseLogging $verboseLogging
    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`$digest = $($digest.digest.GetContextWebInformation.FormDigestValue))"}
    $sanitisedFileName = [uri]::EscapeUriString($(Split-Path $filePathAndName -Leaf)) 
    $url = $listItem.__metadata.uri+"/AttachmentFiles/add(Filename=`'$sanitisedFileName`')"
    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`$url = $url"}
    [System.Net.WebRequest]$request = [System.Net.WebRequest]::CreateHttp($url)
    $request.Credentials = $restCreds
    $request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f")
    $request.Headers.Add("X-RequestDigest", $($digest.digest.GetContextWebInformation.FormDigestValue))
    $request.Method = "POST"

    $fileContent = [System.IO.File]::ReadAllBytes($filePathAndName)
    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`$fileContent.Length = $($fileContent.Length)"}
    $request.ContentLength = $fileContent.Length

    $requestStream = $request.GetRequestStream()
    $requestStream.Write($fileContent,0,$fileContent.Length)
    
    $response = $request.GetResponse()
    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`$response.StatusCode: $($response.StatusCode)"}
    $response.StatusCode
    $response.Dispose()
    }
function check-digestExpiry($serverUrl, $sitePath, $digest, $restCreds,$verboseLogging,$logFile){
    $sitePath = format-path $sitePath
    if($digest -eq $null){
        if($verboseLogging){write-host -ForegroundColor DarkYellow "Digest was `$null"} 
        new-spoDigest -serverUrl $serverUrl -sitePath $sitePath -restCreds $restCreds
        }
    elseif(($digest.expiryTime.AddSeconds(-30) -lt (Get-Date)) -or ($digest.digest.GetContextWebInformation.WebFullUrl -ne $serverUrl+$sitePath)){new-spoDigest -serverUrl $serverUrl -sitePath $sitePath -restCreds $restCreds}
    else{$digest}
    }
function copy-fileInLibrary($sourceSitePath,$sourceLibraryAndFolderPath,$sourceFileName,$destinationSitePath,$destinationLibraryAndFolderPath,$destinationFileName,[boolean]$overwrite, $restCreds, $digest,$verboseLogging,$logFile){
    $digest = check-digestExpiry -serverUrl $serverUrl -sitePath $sitePath -digest $digest -restCreds $restCreds #this needs to be checked for all POST queries
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
        if($verboseLogging){log-action "Invoke-SPORestMethod -Url $url -Method `"POST`" -RequestDigest $($digest.GetContextWebInformation.FormDigestValue)" -logFile $logFile}
        Invoke-SPORestMethod -Url $url -Method "POST" -RequestDigest $digest.digest.GetContextWebInformation.FormDigestValue -credentials $restCreds
        if($verboseLogging){log-result "FILE COPIED: $destinationFileName" -logFile $logFile}
        }
    catch{
        if($verboseLogging){log-error -myError $Error -myFriendlyMessage "Failed to copy-FileInLibrary: Invoke-SPORestMethod -Url $url -Method `"POST`" -RequestDigest $($digest.GetContextWebInformation.FormDigestValue)" -doNotLogToEmail $true -errorLogFile $logFile}
        $false
        }
    }
function delete-folderInLibrary($serverUrl, $sitePath,$libraryName,$folderPathAndNameToBeDeleted, $restCreds, $digest,$verboseLogging,$logFile){
    #This needs tidying up
    $sitePath = format-path (sanitise-forSharePointUrl $sitePath)
    $folderPathAndNameToBeDeleted = format-path $folderPathAndNameToBeDeleted
    $url = "$serverUrl$sitePath/_api/web/GetFolderByServerRelativeUrl('$sitePath$libraryName$folderPathAndNameToBeDeleted')"
    #$dummy = Invoke-SPORestMethod -Url $url 
    $digest = check-digestExpiry -serverUrl $serverUrl -sitePath $sitePath -digest $digest -restCreds $restCreds #this needs to be checked for all POST queries
    try{
        if($verboseLogging){log-action "delete-folderInLibrary: Invoke-SPORestMethod -Url $url -Method `"POST`" -XHTTPMethod `"DELETE`" -RequestDigest $($digest.GetContextWebInformation.FormDigestValue) -ETag `"*`"" -logFile $logFile}
        Invoke-SPORestMethod -Url $url -Method "POST" -XHTTPMethod "DELETE" -RequestDigest $digest.digest.GetContextWebInformation.FormDigestValue -ETag "*" -credentials $restCreds
        if($verboseLogging){log-result "FOLDER DELETED: $sitePath$libraryName$folderPathAndNameToBeDeleted" -logFile $logFile}
        }
    catch{
        if($verboseLogging){log-error -myError $_ -myFriendlyMessage "Failed to delete-folderInLibrary: Invoke-SPORestMethod -Url $url -Method `"POST`" -XHTTPMethod `"DELETE`" -RequestDigest $($digest.GetContextWebInformation.FormDigestValue) -ETag `"*`"" -doNotLogToEmail $true -errorLogFile $logFile}
        $false
        }
    }
function format-itemData($hashTableOfItemData){
    foreach($key in $hashTableOfItemData.Keys){
        if($hashTableOfItemData[$key] -eq $null){$formattedItemData += "`'$key`':`"`", "}
        elseif($hashTableOfItemData[$key].GetType().Name -eq "DateTime"){$formattedItemData += "`'$key`':`"$($hashTableOfItemData[$key])`", "} #If it's a DateTime, mark it up like a string
        elseif($hashTableOfItemData[$key].GetType().Name -eq "Boolean"){$formattedItemData += "`'$key`':`"$($hashTableOfItemData[$key])`", "} #If it's a Boolean, mark it up like a string
        elseif([regex]::Match($hashTableOfItemData[$key],"^[0-9\-\.]+$").Success){$formattedItemData += "`'$key`':$($hashTableOfItemData[$key]), "} #If it's a numeric value
        elseif($hashTableOfItemData[$key].Trim()[0] -eq "{"){[string]$formattedItemData += "`'$key`':$($hashTableOfItemData[$key]), "}#If it's a compound value
        else{$formattedItemData += "`'$key`':`"$($hashTableOfItemData[$key].Replace('"','\"'))`", "} #If it's anything else, treat it as a string
        }
    $formattedItemData = $formattedItemData.Substring(0,$formattedItemData.Length-2) #Trim off the final ","
    $formattedItemData
    }
function format-path($dirtyPath){
    #All "path" variables should be prefixed with a "/", but not suffixed
    if($dirtyPath.Substring(0,1) -ne "/"){$dirtyPath = "/"+$dirtyPath}
    if($dirtyPath.Substring(($dirtyPath.Length-1),1) -eq "/"){$dirtyPath = $dirtyPath.Substring(0,$dirtyPath.Length-1)}
    $dirtyPath
    }
function get-allLists($serverUrl, $sitePath,$restCreds,$verboseLogging,$logFile){
    #Sanitise parameters
    $sitePath = format-path (sanitise-forSharePointUrl $sitePath)
    $listName = (sanitise-forSharePointUrl $listName).Replace("Lists/","")
    #Build and execute REST statement
    $url = $serverUrl+$sitePath+"/_api/web/Lists/"
    try{
        if($verboseLogging){log-action "get-allLists: Invoke-SPORestMethod -Url $url" -logFile $logFile}
        Invoke-SPORestMethod -Url $url -credentials $restCreds -manualTimeOut 600000
        if($verboseLogging){log-result "SUCCESS: List found" -logFile $logFile}
        }
    catch{
        if($verboseLogging){$_;$url;log-error -myError $_ -myFriendlyMessage "Failed to get-allLists: Invoke-SPORestMethod -Url $url" -doNotLogToEmail $true -errorLogFile $logFile}
        $false
        }
    }
function get-fileInLibrary($serverUrl, $sitePath, $libraryAndFolderPath, $fileName, $restCreds,$verboseLogging,$logFile){
    #Sanitise parameters
    $sitePath = format-path $sitePath
    $libraryAndFolderPath = format-path $libraryAndFolderPath
    $fileName = format-path (sanitise-forSharePointFileName $fileName)
    $sanitisedPath = sanitise-forResourcePath $sitePath$libraryAndFolderPath$fileName
    #Build and execute REST statement
    $url = $serverUrl+$sitePath+"/_api/web/GetFileByServerRelativePath(decodedUrl='$sanitisedPath')"
    try{
        if($verboseLogging){log-action "get-fileInLibrary: Invoke-SPORestMethod -Url $url" -logFile $logFile}
        Invoke-SPORestMethod -Url $url -credentials $restCreds
        if($verboseLogging){log-result "SUCCESS: File found in Library" -logFile $logFile}
        }
    catch{
        if($verboseLogging){log-error -myError $_ -myFriendlyMessage "Failed to get-fileInLibrary: Invoke-SPORestMethod -Url $url" -doNotLogToEmail $true -errorLogFile $logFile}
        $false
        }
    }
function get-folderInLibrary($serverUrl, $sitePath, $libraryName, $folderPathAndOrName, $restCreds,$verboseLogging,$logFile){
    #Sanitise parameters
    $sitePath = format-path (sanitise-forSharePointUrl $sitePath)
    $libraryName = format-path (sanitise-forSharePointUrl $libraryName)
    $folderPathAndOrName = format-path ($folderPathAndOrName)
    #$libraryAndFolderPath = format-path (sanitise-forSharePointUrl  $libraryAndFolderPath)
    #$folderName = sanitise-forSharePointFileName $folderName
    $sanitisedPath = "decodedurl='"+(sanitise-forResourcePath $sitePath$libraryName$folderPathAndOrName)+"'"
    #Build and execute REST statement
    #$url = $serverUrl+$sitePath+"/_api/web/GetFolderByServerRelativeUrl('$sitePath$libraryAndFolderPath/$folderName"+"')"
    $url = $serverUrl+$sitePath+"/_api/web/GetFolderByServerRelativePath($sanitisedPath)"#/ListItemAllFields"
    try{
        if($verboseLogging){log-action "get-folderInLibrary: Invoke-SPORestMethod -Url $url" -logFile $logFile}
        Invoke-SPORestMethod -Url $url -credentials $restCreds
        if($verboseLogging){log-result "SUCCESS:`tFolder in Library found" -logFile $logFile}
        }
    catch{
        if($verboseLogging){log-error -myError $_ -myFriendlyMessage "Failed to get-folderInLibrary: Invoke-SPORestMethod -Url $url" -doNotLogToEmail $true -errorLogFile $logFile}
        $false
        }
    }
function get-itemInListFromProperty($serverUrl, $sitePath, $listName, $propertyName, $propertyValue, $restCreds,$verboseLogging,$logFile){
    $sitePath = format-path (sanitise-forSharePointUrl $sitePath)
    $listName = sanitise-forSharePointUrl $listName
    $query = "?filter=$propertyName eq $propertyValue"
    #Build and execute REST statement
    $url = $serverUrl+$sitePath+"/_api/web/Lists/GetByTitle('$listName')/items"
    try{
        if($verboseLogging){log-action "get-itemInListFromProperty: Invoke-SPORestMethod -Url ($url$query)" -logFile $logFile}
        $item = Invoke-SPORestMethod -Url ($url+$query) -credentials $restCreds
        if($item){
            if($verboseLogging){log-result "FOUND ITEM IN LIST FROM PROPERTY" -logFile $logFile}
            $item.results
            }
        else{
            if($verboseLogging){log-result -myFriendlyMessage "WARNING: get-itemInListFromProperty($sitePath, $listName, $propertyName, $propertyValue) returned no items" -logFile $logFile}
            $false
            }
        }
    catch{
        if($verboseLogging){log-error -myError $_ -myFriendlyMessage "get-itemInListFromProperty($sitePath, $listName, $propertyName, $propertyValue) failed to execute" -doNotLogToEmail $true -errorLogFile $logFile}
        $false
        }
    }    
function get-itemsInList($serverUrl, $sitePath, $listName, $oDataQuery, $suppressProgress, $restCreds,$verboseLogging,$logFile){
    #Sanitise parameters
    $sitePath = format-path (sanitise-forSharePointUrl $sitePath)
    $listName = sanitise-forSharePointUrl $listName
    if($oDataQuery){if($oDataQuery.SubString(0,1) -ne "?"){$oDataQuery = "?$oDataQuery"}} #Prefix with ? if user hasn't done so already
    #Build the query
    $url = $serverUrl+$sitePath+"/_api/web/Lists/GetByTitle('$listName')/items$oDataQuery"
    #Run the query
    try{
        if($verboseLogging){log-action "get-itemsInList: Invoke-SPORestMethod -Url $url" -logFile $logFile}
        $partialItems = Invoke-SPORestMethod -Url $url -credentials $restCreds
        if($partialItems){
            if($verboseLogging){log-result "SUCCESS: Initial $($partialItems.results.Count) items returned" -logFile $logFile}
            $queryResults = $partialItems.results
            }
        else{
            if($verboseLogging){log-result -myFriendlyMessage "WARNING: get-itemsInList($sitePath, $listName) returned no items" -logFile $logFile}
            $false
            }
        }
    catch{
        if($verboseLogging){log-error -myError $_ -myFriendlyMessage "get-itemsInList($sitePath, $listName) failed to execute" -doNotLogToEmail $true -errorLogFile $logFile}
        $false
        }
    $i=$partialItems.results.Count
    #Check for additional results
    while($partialItems.__next){
        try{
            if($verboseLogging){log-action "get-itemsInList: Invoke-SPORestMethod -Url $($partialItems.__next)" -logFile $logFile}
            $partialItems = Invoke-SPORestMethod -Url $partialItems.__next -credentials $restCreds
            if($partialItems){
                if($verboseLogging){log-result "SUCCESS: Subsequent $($partialItems.results.Count) items returned" -logFile $logFile}
                $queryResults += $partialItems.results
                }
            else{
                if($verboseLogging){log-result "WARNING: get-itemsInList($sitePath, $listName) returned no items" -logFile $logFile}
                $false
                }
            }
        catch{
            if($verboseLogging){log-error -myError $_ -myFriendlyMessage "get-itemsInList($sitePath, $listName) failed to execute" -errorLogFile $logFile}
            $false
            }
        $i+=$partialItems.results.Count
        if(!$suppressProgress){Write-Host $i}
        }
    $queryResults
    }
function get-library($serverUrl, $sitePath, $libraryName, $restCreds,$verboseLogging,$logFile){
    #Sanitise parameters
    $sitePath = format-path (sanitise-forSharePointUrl $sitePath)
    $libraryName = format-path (sanitise-forSharePointUrl $libraryName) #The LibraryName cannot contain specific characters
    #Build and execute REST statement
    $url = $serverUrl+$sitePath+"/_api/web/GetFolderByServerRelativePath(decodedurl='$sitePath$libraryName')"
    try{
        if($verboseLogging){log-action "get-library: Invoke-SPORestMethod -Url $url" -logFile $logFile}
        Invoke-SPORestMethod -Url $url -credentials $restCreds
        if($verboseLogging){log-result "SUCCESS: Library found" -logFile $logFile}
        }
    catch{
        if($verboseLogging){log-error -myError $_ -myFriendlyMessage "Failed to get-library($sitePath, $libraryName)" -doNotLogToEmail $true -errorLogFile $logFile}
        $false
        }
    }
function get-list($serverUrl, $sitePath, $listName, $restCreds,$verboseLogging,$logFile){
    #Sanitise parameters
    $sitePath = format-path (sanitise-forSharePointUrl $sitePath)
    $listName = (sanitise-forSharePointUrl $listName).Replace("Lists/","")
    #Build and execute REST statement
    $url = $serverUrl+$sitePath+"/_api/web/Lists/GetByTitle('$listName')"
    try{
        if($verboseLogging){log-action "get-list: Invoke-SPORestMethod -Url $url" -logFile $logFile}
        Invoke-SPORestMethod -Url $url -credentials $restCreds
        if($verboseLogging){log-result "SUCCESS: List found" -logFile $logFile}
        }
    catch{
        if($verboseLogging){$_;$url;log-error -myError $_ -myFriendlyMessage "Failed to get-list: Invoke-SPORestMethod -Url $url" -doNotLogToEmail $true -errorLogFile $logFile}
        $false
        }
    }
function get-propertyValueFromSpoMetadata([string]$__metadata, [string]$propertyName){
    $propertyValue = ($__metadata.Replace("@{","").Replace("}","") -split "; " | ?{$_.Substring(0,$propertyName.length) -imatch "$propertyName"})
    if(![string]::IsNullOrEmpty($propertyValue)){$propertyValue.Replace("$propertyName=","")}
    else{$false}
    }
function new-spoDigest($serverUrl, $sitePath, $restCreds,$verboseLogging,$logFile){
    #$digest = $(Invoke-SPORestMethod -Url "$serverUrl$sitePath/_api/contextinfo" -credentials $restCreds -Method "POST")
    $digest = New-Object psobject -Property @{"digest" = $(Invoke-SPORestMethod -Url "$serverUrl$sitePath/_api/contextinfo" -credentials $restCreds -Method "POST" -manualTimeOut 10000)}
    $digest | Add-Member -MemberType NoteProperty expiryTime -Value (Get-Date).AddSeconds($digest.digest.GetContextWebInformation.FormDigestTimeoutSeconds)
    $digest
    #$global:digest = (Invoke-SPORestMethod -Url "$serverUrl$sitePath/_api/contextinfo" -credentials $restCreds -Method "POST")#.GetContextWebInformation.FormDigestValue
    #$global:digestExpiryTime = (Get-Date).AddSeconds($global:digest.GetContextWebInformation.FormDigestTimeoutSeconds)
    }
function new-folderInLibrary($serverUrl, $sitePath, $libraryName, $folderPathAndOrName, $restCreds, $digest,$verboseLogging,$logFile){
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
    $digest = check-digestExpiry -serverUrl $serverUrl -sitePath $sitePath -digest $digest -restCreds $restCreds #this needs to be checked for all POST queries
    #Build and execute REST statement
    $url = $serverUrl+$sitePath+"/_api/web/folders/AddUsingPath($sanitisedPath)"
    $folderExists = (get-folderInLibrary -serverUrl $serverUrl -sitePath $sitePath -libraryName $libraryName -folderPathAndOrName $folderPathAndOrName -restCreds $restCreds -verboseLogging $verboseLogging -logFile $logFile)
    if($folderExists -eq $false){
        try{
            if($verboseLogging){log-action -myMessage "new-folderInLibrary: Invoke-SPORestMethod -Url $url -Method `"POST`" -RequestDigest $($digest.GetContextWebInformation.FormDigestValue)" -logFile $logFile}
            Invoke-SPORestMethod -Url $url -Method "POST" -RequestDigest $digest.digest.GetContextWebInformation.FormDigestValue -credentials $restCreds
            if($verboseLogging){log-result "SUCCESS: Created folder $sitePath$libraryName$folderPathAndOrName" -logFile $logFile}
            }
        catch{
            if($verboseLogging){log-error -myError $_ -myFriendlyMessage "Failed to create new folder: new-folderInLibrary($sitePath, $libraryAndFolderPath, $folderName)" -doNotLogToEmail $true -errorLogFile $logFile}
            $false
            }
        }
    else{
        if($verboseLogging){log-result "WARNING: Folder already exists: $sitePath$libraryName$folderPathAndOrName" -logFile $logFile}
        $folderExists
        }
    }
function new-itemInList($serverUrl, $sitePath,$listName,$predeterminedItemType,$hashTableOfItemData,$restCreds,$digest,$verboseLogging,$logFile){
    #Error handling for no DataType?
    #Sanitise parameters
    $sitePath = format-path (sanitise-forSharePointUrl $sitePath)
    $listName = sanitise-forSharePointUrl(sanitise-forSharePointFileName ($listName.Replace("Lists/","")))
    #Prepare security and log action
    $digest = check-digestExpiry -serverUrl $serverUrl -sitePath $sitePath -digest $digest -restCreds $restCreds #this needs to be checked for all POST queries
    if($verboseLogging){log-action "new-itemInList($sitePath,$listName,$predeterminedItemType,$($hashTableOfItemData.Keys | %{"$_=$($hashTableOfItemData[$_]);"})" -logFile $logFile}
    #Build and execute REST statement
    $url = $serverUrl+$sitePath+"/_api/web/Lists/GetByTitle('$listName')/items"
    $formattedItemData = format-itemData -hashTableOfItemData $hashTableOfItemData
    $metadata = "{ '__metadata': { 'type': '$predeterminedItemType' }, $formattedItemData}"
    try{
        if($verboseLogging){log-action "Invoke-SPORestMethod -Url $url -Method `"POST`" -Metadata $metadata -RequestDigest $($digest.digest.GetContextWebInformation.FormDigestValue)" -logFile $logFile}
        Invoke-SPORestMethod -Url $url -Method "POST" -Metadata $metadata -RequestDigest $digest.digest.GetContextWebInformation.FormDigestValue -credentials $restCreds
        if($verboseLogging){log-result "SUCCESS: New item created" -logFile $logFile}
        }
    catch{
        if($verboseLogging){log-error $_ -myFriendlyMessage "Failed to create new list item: new-itemInList($sitePath,$listName,$predeterminedItemType,$($hashTableOfItemData.Keys | %{"$_=$($hashTableOfItemData[$_]);"})" -doNotLogToEmail $true -errorLogFile $logFile}
        #See if it already exists, and if so, return that instead
        try{
            #Extract the Title property from $hashTableOfItemData and use that to try and retrieve the item
            if($verboseLogging){log-action -myMessage "Checking to see if the item has already been created" -logFile $logFile}
            $item = get-itemsInList -serverUrl $serverUrl -sitePath $sitePath -listName $listName -oDataQuery "?`$filter=Title eq `'$($hashTableOfItemData["Title"])`'" -restCreds $restCreds -verboseLogging $verboseLogging -logFile $logFile
            if($item){
                if($verboseLogging){log-result -myMessage "SUCCESS: Item already existed and has been returned" -logFile $logFile}
                return $item
                }
            else{
                write-host "did not find item $($hashTableOfItemData["Title"])"
                if($verboseLogging){log-result -myMessage "FAILURE: Item did not already exist either" -logFile $logFile}
                $false
                }
            }
        catch{
            if($verboseLogging){log-error -myError $_ -myFriendlyMessage "Error when attempting to retrieve List Item [$($hashTableOfItemData["Title"])] that might have already been created" -errorLogFile $logFile -doNotLogToEmail $true}
            $false
            }
        
        }
    }
function new-library($serverUrl, $sitePath, $libraryName, $libraryDesc, $digest, $restCreds){
    #Sanitise parameters
    $sitePath = format-path (sanitise-forSharePointUrl $sitePath)
    $libraryName = sanitise-forSharePointFileName $libraryName
    #Prepare security and log action
    $digest = check-digestExpiry -serverUrl $serverUrl -sitePath $sitePath -digest $digest -restCreds $restCreds #this needs to be checked for all POST queries
    #Build and execute REST statement
    $metadata = "{'__metadata':{'type':'SP.List'},'Description':`"$libraryDesc`",'BaseTemplate':101,'ContentTypesEnabled':true,'Title':`"$libraryName`",'AllowContentTypes':true}"
    $url = $serverUrl+$sitePath+"/_api/web/lists"
    $libraryExists = get-library -sitePath $sitePath -libraryName $libraryName -serverUrl $serverUrl -restCreds $restCreds
    if($libraryExists -eq $false){
        try{
            if($verboseLogging){log-action -myMessage "new-library: Invoke-SPORestMethod -Url $url -Method `"POST`" -Metadata $metadata -RequestDigest $($digest.GetContextWebInformation.FormDigestValue)" -logFile $logFile}
            Invoke-SPORestMethod -Url $url -Method "POST" -Metadata $metadata -RequestDigest $digest.digest.GetContextWebInformation.FormDigestValue -credentials $restCreds
            if($verboseLogging){log-result "SUCCESS: Library created: $sitePath/$libraryName" -logFile $logFile}
            }
        catch{
            if($verboseLogging){log-error -myError $_ -myFriendlyMessage "Failed to create new Library: new-library($sitePath, $libraryName, $libraryDesc)" -doNotLogToEmail $true -errorLogFile $logFile}
            $false
            }
        }
    else{
        if($verboseLogging){log-result "WARNING: Library already exists: $sitePath/$libraryName" -logFile $logFile}
        $libraryExists
        }
    }
function new-spoCred($username, $securePassword){
    New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userName, $securePassword)
    }
function update-list($serverUrl, $sitePath, $listName,$hashTableOfUpdateData, $restCreds, $digest,$verboseLogging,$logFile){
    #Sanitise parameters
    $sitePath = format-path (sanitise-forSharePointUrl $sitePath)
    $listName = sanitise-forSharePointUrl $listName
    #Prepare security and log action
    $digest = check-digestExpiry -serverUrl $serverUrl -sitePath $sitePath -digest $digest -restCreds $restCreds #this needs to be checked for all POST queries
    #Build and execute REST statement
    $url = $serverUrl+$sitePath+"/_api/web/Lists/GetByTitle('$listName')"
    $formattedItemData = format-itemData -hashTableOfItemData $hashTableOfUpdateData
    $metadata = "{'__metadata':{'type':'SP.List'},$formattedItemData}"
    try{
        if($verboseLogging){log-action "update-list: Invoke-SPORestMethod -Url $url -Method `"POST`" -XHTTPMethod `"MERGE`" -Metadata $metadata -RequestDigest $($digest.GetContextWebInformation.FormDigestValue) -ETag `"*`"" -logFile $logFile}
        Invoke-SPORestMethod -Url $url -Method "POST" -XHTTPMethod "MERGE" -Metadata $metadata -RequestDigest $digest.digest.GetContextWebInformation.FormDigestValue -ETag "*" -credentials $restCreds
        if($verboseLogging){log-result "SUCCESS: List updated: $formattedItemData" -logFile $logFile}
        }
    catch{
        if($verboseLogging){log-error $_ -myFriendlyMessage "Failed to update-list($sitePath, $listName,$($hashTableOfUpdateData.Keys | %{"$_=$($hashTableOfUpdateData[$_]);"})" -doNotLogToEmail $true -errorLogFile $logFile}
        $false
        }
    }
function update-itemInList($serverUrl,$sitePath,$listNameOrGuid,$predeterminedItemType,$itemId,$hashTableOfItemData,$restCreds,$digest,$verboseLogging,$logFile){
    #Sanitise parameters
    $sitePath = format-path (sanitise-forSharePointUrl $sitePath)
    
    #Prepare security and log action
    $digest = check-digestExpiry -serverUrl $serverUrl -sitePath $sitePath -digest $digest -restCreds $restCreds #this needs to be checked for all POST queries
    #Build and execute REST statement
    if ($listNameOrGuid.GetType().Name -eq  "Guid"){$url = $serverUrl+$sitePath+"/_api/web/Lists(guid'$listNameOrGuid')/items($itemId)"}
    else{
        $listName = sanitise-forSharePointUrl(sanitise-forSharePointFileName ($listNameOrGuid.Replace("Lists/","")))
        $url = $serverUrl+$sitePath+"/_api/web/Lists/GetByTitle('$listName')/items($itemId)"
        }
    $formattedItemData = format-itemData -hashTableOfItemData $hashTableOfItemData
    $metadata = "{ '__metadata': { 'type': '$predeterminedItemType' }, $formattedItemData}"
    try{
        if($verboseLogging){log-action "Invoke-SPORestMethod -Url $url -Method `"POST`" -XHTTPMethod `"MERGE`" -Metadata $metadata -RequestDigest $($digest.GetContextWebInformation.FormDigestValue) -ETag `"*`"" -logFile $logFile}
        Invoke-SPORestMethod -Url $url -Method "POST" -XHTTPMethod "MERGE" -Metadata $metadata -RequestDigest $digest.digest.GetContextWebInformation.FormDigestValue -ETag "*" -credentials $restCreds
        if($verboseLogging){log-result "SUCCESS: Updated list item: $formattedItemData" -logFile $logFile}
        }
    catch{
        if($verboseLogging){log-error $_ -myFriendlyMessage "Failed to update item in List: update-itemInList($sitePath,$listName,$predeterminedItemType,$itemId,$($hashTableOfItemData.Keys | %{"$_=$($hashTableOfItemData[$_]);"})" -doNotLogToEmail $true -errorLogFile $logFile}
        $false
        }
    }

