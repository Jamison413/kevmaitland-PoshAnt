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
        $request.ContentType = "application/json$odata"   
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
    $sitePath = format-path (sanitise-forSharePointUrl $sitePath)
    $libraryAndFolderPath = format-path (sanitise-forSharePointUrl $libraryAndFolderPath)
    $fileName = sanitise-forSharePointFileName $fileName
    #Build and execute REST statement
    $url = $serverUrl+$sitePath+"/_api/web/GetFileByServerRelativeUrl('$sitePath$libraryAndFolderPath/$fileName"+"')"
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
function get-folderInLibrary($sitePath, $libraryAndFolderPath, $folderName){
    #Sanitise parameters
    $sitePath = format-path (sanitise-forSharePointUrl $sitePath)
    $libraryAndFolderPath = format-path (sanitise-forSharePointUrl  $libraryAndFolderPath)
    $folderName = sanitise-forSharePointFileName $folderName
    #Build and execute REST statement
    $url = $serverUrl+$sitePath+"/_api/web/GetFolderByServerRelativeUrl('$sitePath$libraryAndFolderPath/$folderName"+"')"
    try{
        if($verboseLogging){log-action "get-folderInLibrary: Invoke-SPORestMethod -Url $url"}
        Invoke-SPORestMethod -Url $url
        if($verboseLogging){log-result "FOLDER IN LIBRARY FOUND"}
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
    $libraryName = sanitise-forSharePointFileName $libraryName
    #Build and execute REST statement
    $url = $serverUrl+$sitePath+"/_api/web/GetFolderByServerRelativeUrl('"+(sanitise-forSharePointUrl "$sitePath/$libraryName")+"')"
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
function new-folderInLibrary($sitePath, $libraryAndFolderPath, $folderName){
    #Sanitise parameters
    $sitePath = format-path (sanitise-forSharePointUrl $sitePath)
    $folderName = sanitise-forSharePointFileName $folderName
    $libraryAndFolderPath = format-path (sanitise-forSharePointUrl $libraryAndFolderPath)
    #Prepare security and log action
    check-digestExpiry -serverUrl $serverUrl -sitePath $sitePath  #this needs to be checked for all POST queries
    #Build and execute REST statement
    $metadata = "{'__metadata':{'type':'SP.Folder'},'ServerRelativeUrl':`"$sitePath$libraryAndFolderPath/$folderName`"}"
    $url = $serverUrl+$sitePath+"/_api/web/folders"
    if((get-folderInLibrary -sitePath $sitePath -libraryAndFolderPath $libraryAndFolderPath -folderName $folderName) -eq $false){
        try{
            if($verboseLogging){log-action -myMessage "new-folderInLibrary: Invoke-SPORestMethod -Url $url -Method `"POST`" -Metadata $metadata -RequestDigest $($digest.GetContextWebInformation.FormDigestValue)"}
            Invoke-SPORestMethod -Url $url -Method "POST" -Metadata $metadata -RequestDigest $digest.GetContextWebInformation.FormDigestValue
            if($verboseLogging){log-result "SUCCESS: Created folder $sitePath$libraryAndFolderPath/$folderName"}
            }
        catch{
            if($verboseLogging){log-error -myError $_ -myFriendlyMessage "Failed to create new folder: new-folderInLibrary($sitePath, $libraryAndFolderPath, $folderName)" -doNotLogToEmail $true}
            $false
            }
        }
    else{
        if($verboseLogging){log-result "WARNING: Folder already exists: $sitePath$libraryAndFolderPath/$folderName"}
        $false
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
    if((get-library -sitePath $sitePath -libraryName $libraryName) -eq $false){
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
        else{if($verboseLogging){log-result "WARNING: Library already exists: $sitePath/$libraryName"}}
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
    $dirtyString = $dirtyString.Replace("`"","").Replace("#","").Replace("%","").Replace("?","").Replace("<","").Replace(">","").Replace("\","/").Replace("//","/").Replace(":","")
    $dirtyString = $dirtyString.Replace("$","`$").Replace("``$","`$").Replace("(","").Replace(")","").Replace("-","").Replace(".","").Replace("&","").Replace(",","").Replace("'","").Replace("!","")
    $cleanString =""
    for($i= 0;$i -lt $dirtyString.Split("/").Count;$i++){ #Examine each virtual directory in the URL
        if($i -gt 0){$cleanString += "/"}
        if($dirtyString.Split("/")[$i].Length -gt 50){$tempString = $dirtyString.Split("/")[$i].SubString(0,50).Trim()} #Truncate long folder names to 50 characters
            else{$tempString = $dirtyString.Split("/")[$i]}
        if($tempString.Length -gt 0){
            if(@(".", " ") -contains $tempString.Substring(($tempString.Length-1),1)){$tempString = $tempString.Substring(0,$tempString.Length-1)} #Trim trailing "." and " ", even if this results in a truncation <50 characters
            }
        $cleanString += $tempString
        }
    $cleanString = $cleanString.Replace("//","/").Replace("https/","https://") #"//" is duplicated to catch trailing "/" that might now be duplicated. https is an exception that needs specific handling
    $cleanString
    }
#endregion
#region Kimble functions
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$callbackUri = "https://login.salesforce.com/services/oauth2/token"
#"https://test.salesforce.com/services/oauth2/token"
$grantType = "password"
$myInstance = "https://eu5.salesforce.com"
$queryUri = "$myInstance/services/data/v39.0/query/?q="
$querySuffixStub = " -H `"Authorization: Bearer "
$clientId = "3MVG9Rd3qC6oMalWu.nvQtpSk61bDN.lZwfX8gpDqVnnIVt9zRnzJlDlG59lMF4QFnj.mmycmnid3o94k6Y9j"
$clientSecret = "3010701969925148301"
$username = "kevin.maitland@anthesisgroup.com"
$password = "M0nkeyfucker"
$securityToken = "Ou4G5iks8m5axtp6iDldxUZq"
#$username = "system.admin@anthesisgroup.com.sandbox"
#$password = "SisethaN2017"
#$securityToken = "eOURcVPchk8Xogv2hlbV3NSV1"

#region Kimble functions
function Get-KimbleAuthorizationTokenWithUsernamePasswordFlowRequestBody ($client_id, $client_secret, $user_name, $pass_word, $security_token){
    Add-Type -AssemblyName System.Web
    $user_name = [System.Web.HttpUtility]::UrlEncode($user_name)
    $pass_word = [System.Web.HttpUtility]::UrlEncode($pass_word)
    $requestBody = "grant_type=$grantType"
    $requestBody += "&client_id=$client_id"
    $requestBody += "&client_secret=$client_secret"
    $requestBody += "&username=$user_name"
    $requestBody += "&password=$pass_word$security_token"
    $requestBody += "&Content-Type=application/x-www-form-urlencoded"
    $requestBody
    #Write-Host "Body:" $requestBody

    #Invoke-RestMethod -Method Post -Uri $callbackUri -Body $requestBody
    #try{Invoke-RestMethod -Method Post -Uri $callbackUri -Body $requestBody} catch {Failure}
    }
function Failure {
    $global:helpme = $body
    $global:helpmoref = $moref
    $global:result = $_.Exception.Response.GetResponseStream()
    $global:reader = New-Object System.IO.StreamReader($global:result)
    $global:responseBody = $global:reader.ReadToEnd();
    Write-Host -BackgroundColor:Black -ForegroundColor:Red "Status: A system exception was caught."
    Write-Host -BackgroundColor:Black -ForegroundColor:Red $global:responsebody
    Write-Host -BackgroundColor:Black -ForegroundColor:Red "The request body has been saved to `$global:helpme"
    break
    }
function Get-KimbleSoqlDataset($queryUri, $soqlQuery, $restHeaders){
    $soqlQuery = $soqlQuery.Replace(" ","+")
    $lastIndex = 0
    $nextIndex = 0
    do{
        $lastIndex = $nextIndex
        if($lastIndex -eq 0){
            #Write-Host -ForegroundColor Magenta $partialDataSet.totalSize
            $partialDataSet = Invoke-RestMethod -Uri $queryUri+$soqlQuery -Headers $restHeaders
            $fullDataSet = New-Object object[] $partialDataSet.totalSize
            }
            else{$partialDataSet = Invoke-RestMethod -Uri $myInstance$($partialDataSet.nextRecordsUrl) -Headers $restHeaders}
        try{[int]$nextIndex = $partialDataSet.nextRecordsUrl.Split("-")[1]}catch{$nextIndex = $partialDataSet.totalSize-1}
        $j=0
        for($i=$lastIndex;$i -le $nextIndex;$i++){
            $fullDataSet[$i] = $partialDataSet.records[$j]
            $j++
            if($i%100 -eq 0){Write-Host -ForegroundColor DarkMagenta $i $j}
            }
        Write-Host -ForegroundColor Yellow $partialDataSet.nextRecordsUrl
        }
    while($partialDataSet.nextRecordsUrl -ne $null)
    $fullDataSet
    }

#endregion


#endregion

##################################
#
#Get ready
#
##################################
$o365user = "kevin.maitland@anthesisgroup.com"
$o365Pass = ConvertTo-SecureString (Get-Content 'C:\New Text Document.txt') -AsPlainText -Force
$credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $o365user, $o365Pass
$logfile = "C:\users\administrator.sustainltd\Desktop\provisionSpoClients.log"
$logErrors = $true
$logMethodMain = $true
$logFunctionCalls = $true
Set-SPORestCredentials -Credential $credential

$oAuthReqBody = Get-KimbleAuthorizationTokenWithUsernamePasswordFlowRequestBody -client_id $clientId -client_secret $clientSecret -user_name $username -pass_word $password -security_token $securityToken
try{$kimbleAccessToken=Invoke-RestMethod -Method Post -Uri $callbackUri -Body $oAuthReqBody} catch {Failure}
$kimbleRestHeaders = @{Authorization = "Bearer " + $kimbleAccessToken.access_token}


##################################
#
#Do Stuff
#
##################################

#region Kimble Sync
#Get the last Lead modified in /lists/Kimble Leads to minimise the number of records to process
$serverUrl = "https://anthesisllc.sharepoint.com" 
$sitePath = "/clients"
$listName = "Kimble Leads"
get-newDigest -serverUrl $serverUrl -sitePath $sitePath
$kp = get-list -sitePath $sitePath -listName $listName

#Get the Kimble Leads that have been modifed since the last update
$cutoffDate = (Get-Date (Get-Date $kp.LastItemModifiedDate).AddHours(-1) -Format s) #Look one hour behind just in case there is ever a delay between polling Kimble and updating SharePoint
#$cutoffDate = (Get-Date (Get-Date $kp.LastItemModifiedDate).AddYears(-1) -Format s) #Bodge this once for the initial Sync
$soqlQuery = "SELECT Name,Id,KimbleOne__Account__c,LastModifiedDate,SystemModStamp,CreatedDate,IsDeleted FROM KimbleOne__SalesOpportunity__c WHERE LastModifiedDate > $cutoffDate`Z"

$kimbleModifiedLeads = Get-KimbleSoqlDataset -queryUri $queryUri -soqlQuery $soqlQuery -restHeaders $kimbleRestHeaders
$kimbleChangedLeads = $kimbleModifiedLeads | ?{$_.LastModifiedDate -lt $cutoffDate}
$kimbleNewLeads = $kimbleModifiedLeads | ?{$_.CreatedDate -ge $cutoffDate}
#Check any other Leads for changes
#At what point does it become more efficent to dump the whole [Kimble Leads] List from SP, rather than query individual items?
#SP pages results back in 100s, so when $spLead.Count/100 -gt $kimbleChangedLeads.Count, it will take fewer requests to query each $kimbleChangedLeads individually. This ought to happen most of the time (unless there is a batch update of Leads)

<# Use this is a batch update...
$spLeads = get-itemsInList -sitePath $sitePath -listName "Kimble Leads"
foreach($kimbleChangedLead in $kimbleChangedLeads){
    $spLead = $null
    $spLead = $spLeads | ?{$_.KimbleId -eq $kimbleChangedLead.Id}
    if($spLead){
        #Check whether spLead.Title = modLead.Name and update and mark IsDirty if necessary ;PreviousName=
        #if($spLead)
        }
    else{#Lead is missing from SP, so add it
        $kimbleNewLeads += $kimbleChangedLead
        }
    }
#>
#Otherwise, use this:
foreach($kimbleChangedLead in $kimbleChangedLeads){
    $kpListItem = get-itemsInList -sitePath $sitePath -listName "Kimble Leads" -oDataQuery "?&`$filter=KimbleId eq `'$($kimbleChangedLead.Id)`'"
    if($kpListItem){
        #Check whether the data has changed
        if($kpListItem.Title -ne $kimbleChangedLead.Name `
            -or $kpListItem.KimbleClientId -ne $kimbleChangedLead.KimbleOne__Account__c `
            -or $kpListItem.IsDeleted -ne $kimbleChangedLead.IsDeleted){
            #Update the entry in [Kimble Leads]
            $updateData = @{PreviousName=$kpListItem.LeadName;PreviousKimbleClientId=$kpListItem.KimbleClientId;Title=$kimbleChangedLead.Name;KimbleClientId=$kimbleChangedLead.KimbleOne__Account__c;IsDeleted=$kimbleChangedLead.IsDeleted;IsDirty=$true}
            try{update-itemInList -serverUrl $serverUrl -sitePath $sitePath -listName "Kimble Leads" -predeterminedItemType $kp.ListItemEntityTypeFullName -itemId $kpListItem.Id -hashTableOfItemData $updateData}
            catch{$false;log-error -myError $Error[0] -myFriendlyMessage "Failed to update [Kimble Leads].$($kimbleChangedLead.Id) with $updateData"}
            }
        }
    else{$kimbleNewLeads += $kimbleChangedLead} #The Library doesn't exist, so add it to the "to-be-created" array, even though we were expecting it to exist
    }


#Add the new Leads
foreach ($kimbleNewLead in $kimbleNewLeads){
#foreach ($kimbleNewLead in $kimbleNewLeads){
    $kimbleNewLeadData = @{KimbleId=$kimbleNewLead.Id;Title=$kimbleNewLead.Name;KimbleClientId=$kimbleNewLead.KimbleOne__Account__c;IsDeleted=$kimbleNewLead.IsDeleted;IsDirty=$true}
    try{new-itemInList -sitePath $sitePath -listName "Kimble Leads" -predeterminedItemType $kp.ListItemEntityTypeFullName -hashTableOfItemData $kimbleNewLeadData}
    catch{$false;log-error -myError $Error[0] -myFriendlyMessage "Failed to create new [Kimble Leads].$($kimbleNewLead.Id) with $kimbleNewLeadData"}
    }

#endregion



<##############################
#For building the initial Sync
###############################


$spLeads = get-itemsInList -sitePath $sitePath -listName "Kimble Leads" 
$remainingKimbleLeads = $kimbleModifiedLeads | ?{($spLeads | %{$_.KimbleId}) -notcontains $_.Id}

$remainingKimbleLeads = ,@();$j=0
foreach ($c in $kimbleModifiedLeads){
    $foundIt = $false
    foreach ($createdLead in $spLeads){
        if($c.Id -eq $createdLead.KimbleId){$foundIt= $true;break}
        }
    if(!$foundIt){$remainingKimbleLeads += $c}
    $j++
    if($j%100 -eq 0){$j}
    }

foreach ($kimbleNewLead in $remainingKimbleLeads){
#foreach ($kimbleNewLead in $kimbleNewLeads){
    $kimbleNewLeadData = @{KimbleId=$kimbleNewLead.Id;Title=$kimbleNewLead.Name;IsDeleted=$kimbleNewLead.IsDeleted;IsDirty=$true}
    switch ($kimbleNewLead.Description.Length){
        0 {break}
        {$_ -lt 255} {$kimbleNewLeadData.Add("LeadDescription","$($kimbleNewLead.Description)");break}
        default {$kimbleNewLeadData.Add("LeadDescription","$($kimbleNewLead.Description.Substring(0,254))")}
        }
    new-itemInList -sitePath $sitePath -listName "Kimble Leads" -predeterminedItemType $kp.ListItemEntityTypeFullName -hashTableOfItemData $kimbleNewLeadData
    }

$kimbleModifiedLeads.Count
$spLeads.Count
$remainingKimbleLeads.Count

#>