function combine-url($arrayOfStrings){
    $output = ""
    $arrayOfStrings | % {
        $output += $_.TrimStart("/").TrimEnd("/")+"/"
        }
    $output = $output.Substring(0,$output.Length-1)
    $output = $output.Replace("//","/").Replace("//","/").Replace("//","/")
    $output = $output.Replace("http:/","http://").Replace("https:/","https://")
    $output
    }
function convertTo-localisedSecureString($plainText){
    if ($(Get-Module).Name -notcontains "_PS_Library_Forms"){Import-Module _PS_Library_Forms}
    if (!$plainText){$plainText = form-captureText -formTitle "PlainText" -formText "Enter the plain text to be converted to a secure string" -sizeX 300 -sizeY 200}
    ConvertTo-SecureString $plainText -AsPlainText -Force | ConvertFrom-SecureString
    }
function log-action($myMessage, $logFile, $doNotLogToFile, $doNotLogToScreen){
    if(!$doNotLogToFile -or $logToFile){Add-Content -Value ((Get-Date -Format "yyyy-MM-dd HH:mm:ss")+"`tACTION:`t$myMessage") -Path $logFile}
    if(!$doNotLogToScreen -or $logToScreen){Write-Host -ForegroundColor Yellow $myMessage}
    }
function log-error($myError, $myFriendlyMessage, $fullLogFile, $errorLogFile, $doNotLogToFile, $doNotLogToScreen, $doNotLogToEmail, $smtpServer, $mailTo, $mailFrom){
    if(!$doNotLogToFile -or $logToFile){
        Add-Content -Value "`t`tERROR:`t$myFriendlyMessage" -Path $errorLogFile
        Add-Content -Value "`t`t$($myError.Exception.Message)" -Path $errorLogFile
        Add-Content -Value "`t`tERROR:`t$myFriendlyMessage" -Path $fullLogFile
        Add-Content -Value "`t`t$($myError.Exception.Message)" -Path $fullLogFile
        }
    if(!$doNotLogToScreen -or $logToScreen){
        Write-Host -ForegroundColor Red $myFriendlyMessage
        Write-Host -ForegroundColor Red $myError
        }
    if(!$doNotLogToEmail -or $logErrorsToEmail){
        Send-MailMessage -To $mailTo -From $mailFrom -Subject "Error in automated script - $myFriendlyMessage" -Body $myError -SmtpServer $smtpServer
        }
    }
function log-result($myMessage, $logFile, $doNotLogToFile, $doNotLogToScreen){
    if(!$doNotLogToFile -or $logToFile){Add-Content -Value ("`tRESULT:`t$myMessage") -Path $logfile}
    if(!$doNotLogToScreen -or $logToScreen){Write-Host -ForegroundColor DarkYellow "`t$myMessage"}
    }
function sanitise-forSharePointStandard($dirtyString){
    $dirtyString = $dirtyString.Trim()
    $dirtyString = $dirtyString.Replace(" "," ") #Weird instance where a space character is not a space character...
    if(@("."," ") -contains $dirtyString.Substring(($dirtyString.Length-1),1)){$dirtyString = $dirtyString.Substring(0,$dirtyString.Length-1)} #Trim trailing "."
    $dirtyString.Replace("`"","").Replace("#","").Replace("%","").Replace("?","").Replace("<","").Replace(">","").Replace("\","").Replace("...","").Replace("..","").Replace("'","`'").Replace("`t","").Replace("`r","").Replace("`n","").Replace("*","")
    }
function sanitise-LibraryNameForUrl($dirtyString){
    $cleanerString = $dirtyString.Trim()
    $cleanerString = $dirtyString -creplace '[^a-zA-Z0-9 _/]+', ''
    $cleanerString
    }
function sanitise-forSharePointListName($dirtyString){ 
    $cleanerString = sanitise-forSharePointStandard $dirtyString
    $cleanerString.Replace("/","")
    }
function sanitise-forSharePointFileName($dirtyString){ 
    $cleanerString = sanitise-forSharePointStandard $dirtyString
    $cleanerString.Replace("/","").Replace(":","")
    }
function sanitise-forSharePointFileName2($dirtyString){ 
    $dirtyString = $dirtyString.Trim()
    $dirtyString.Replace("`"","").Replace("#","").Replace("%","").Replace("?","").Replace("<","").Replace(">","").Replace("\","").Replace("/","").Replace("...","").Replace("..","").Replace("'","`'")
    if(@("."," ") -contains $dirtyString.Substring(($dirtyString.Length-1),1)){$dirtyString = $dirtyString.Substring(0,$dirtyString.Length-1)} #Trim trailing "."
    }
function sanitise-forSharePointFolderPath($dirtyString){ 
    $cleanerString = sanitise-forSharePointStandard $dirtyString
    $cleanerString.Replace(":","")
    }
function sanitise-forSharePointUrl($dirtyString){ 
    $dirtyString = $dirtyString.Trim()
    $dirtyString = $dirtyString.Replace(" "," ") #Weird instance where a space character is not a space character...
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
    if($dirtyString.Length -gt 0){
        if(@("."," ") -contains $dirtyString.Substring(($dirtyString.Length-1),1)){$dirtyString = $dirtyString.Substring(0,$dirtyString.Length-1)} #Trim trailing "."
        $dirtyString = $dirtyString.trim().replace("`'","`'`'")
        $dirtyString = $dirtyString.replace("#","").replace("%","") #As of 2017-05-26, these characters are not supported by SharePoint (even though https://msdn.microsoft.com/en-us/library/office/dn450841.aspx suggests they should be)
        #$dirtyString = $dirtyString -creplace "[^a-zA-Z0-9 _/()`'&-@!]+", '' #No need to strip non-standard characters
        #[uri]::EscapeUriString($dirtyString) #No need to encode the URL
        $dirtyString
        }
    }
function sanitise-stripHtml($dirtyString){
    $cleanString = $dirtyString -replace '<[^>]+>',''
    $cleanString = [System.Web.HttpUtility]::HtmlDecode($cleanString)# -replace '&amp;','&'
    $cleanString
    }
