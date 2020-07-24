$FTPRequest = [System.Net.FtpWebRequest]::Create("ftp://ftp.adobe.com/pub/adobe/reader/win/AcrobatDC/")
$FTPRequest.Method = [System.Net.WebRequestMethods+Ftp]::ListDirectoryDetails
$FTPResponse = $FTPRequest.GetResponse() 
$ResponseStream = $FTPResponse.GetResponseStream()
$StreamReader = New-Object System.IO.StreamReader $ResponseStream  
   
#Read each line of the stream and add it to an array list
$files = New-Object System.Collections.ArrayList
While ($file = $StreamReader.ReadLine()){
    [void] $files.add("$file")
    }
$latestVersion = $files | % {$_.Split(" ")[27]} | ? {$_.Length -gt 4} | sort -Descending | select -First 1

$FTPRequest = [System.Net.FtpWebRequest]::Create("ftp://ftp.adobe.com/pub/adobe/reader/win/AcrobatDC/$latestVersion")
$FTPRequest.Method = [System.Net.WebRequestMethods+Ftp]::ListDirectoryDetails
$FTPResponse = $FTPRequest.GetResponse() 
$ResponseStream = $FTPResponse.GetResponseStream()
$StreamReader = New-Object System.IO.StreamReader $ResponseStream
$files = New-Object System.Collections.ArrayList
While ($file = $StreamReader.ReadLine()){
    [void] $files.add("$file")
    }
