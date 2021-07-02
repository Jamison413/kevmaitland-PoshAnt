##############################################
#                                            #
#                                            #
#   GhostScript - Automated PDF Compression  #
#                                            #
#                                            #
##############################################

<#
Some extra useful notes to help

Ghostscript has a notion of 'output devices' which handle saving or displaying the results in a particular format. Ghostscript comes with a diverse variety of such devices supporting vector and raster file output, screen display, driving various printers and communicating with other applications. The command line option '-sDEVICE=device' selects which output device Ghostscript should use. The pdfwrite device outputs PDF.
To send the output to a file, use the -sOutputFile= switch

http://flukylogs.blogspot.com/2012/08/gs-ghostscript-cheat-sheet.html

We call gswin64c for GhostScript, all arguments are handled inline in PS (c for command-line, calling the non-c version will pop open a window despite using the -q switch, -q does not need to be declared).
Using -dBATCH -dNOPAUSE will disable interactive prompting, or it will only process one page (as it is waiting to be set to the next line in window)

Example: $result = $(& 'C:\Program Files\gs\gs9.27\bin\gswin64c.exe' -sDEVICE=pdfwrite -sOutputFile='C:\Users\Susmin-Emily\Desktop\pdftest\test.pdf' -dPDFSETTINGS=/ebook -dNOPAUSE -dBATCH -dSAFER 'C:\Users\Susmin-Emily\Desktop\pdftest\PowerApps Governance and Deployment Whitepaper.pdf')


**Please note, it looks as though Ghostscript cannot overwrite or replace the original files, so the below script will handle this part separately:
1. Get the target file and compress it
2. Save as a similar temporary name
3. Check if it made the temporary file - get the file, check the size is greater than 0 and check a result was retrieved from the compression task.
4. If successful based on step 3, then delete the original file and rename the temporary one to match. If unsuccessful skip this part and throw an error.
#>

#Set some basic logging
$Logname = "C:\Scripts" + "\Logs" + "\compress-pdf$(Get-Date -Format "yyMMdd").log"

#Set some nicer logging
$friendlyLogname = "C:\Scripts" + "\Logs" + "\friendlylog-compress-pdf$(Get-Date -Format "yyMMdd").log"
function friendlyLogWrite(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [string]$friendlyLogname
       ,[parameter(Mandatory = $true)]
            [string]$logstring
       ,[parameter(Mandatory = $true)]
            [validateset("WARNING","SUCCESS","ERROR","ERROR DETAILS","MESSAGE","END","START")]
            [String]$messagetype
        )
If($messagetype -eq "MESSAGE"){
Add-content $friendlyLogname -value $("**********************************************************************************************************************************************************")
Add-content $friendlyLogname -value $("$(get-date)" + " MESSAGE: " + "$($logstring)")
Add-content $friendlyLogname -value $("**********************************************************************************************************************************************************")
}

If($messagetype -eq "START"){
Add-content $friendlyLogname -value $("-------------------------------------------------------------------------------------------------------------------------------------------------------------")
Add-content $friendlyLogname -value $("$(get-date)" + " START: " + "$($logstring)")
Add-content $friendlyLogname -value $("-------------------------------------------------------------------------------------------------------------------------------------------------------------")
}


If($messagetype -eq "END"){
Add-content $friendlyLogname -value $("-------------------------------------------------------------------------------------------------------------------------------------------------------------")
Add-content $friendlyLogname -value $("$(get-date)" + " END: " + "$($logstring)")
Add-content $friendlyLogname -value $("-------------------------------------------------------------------------------------------------------------------------------------------------------------")
}


If($messagetype -eq "WARNING"){
Add-content $friendlyLogname -value $("$(get-date)" + "     WARNING: " + "$($logstring)")
}

If($messagetype -eq "SUCCESS"){
$content = 
Add-content $friendlyLogname -value $("$(get-date)" + "     SUCCESS: " + "$($logstring)")
}

If($messagetype -eq "ERROR"){
Add-content $friendlyLogname -value $("$(get-date)" + "     ERROR: " + "$($logstring)")
}

If($messagetype -eq "ERROR DETAILS"){
$value = $("$(get-date)" + "     ERROR DETAILS: " + "$($logstring)")
Add-content $friendlyLogname -value $value
}
}



$thisRunDateStamp = (get-date)
$lastRunDateStamp = $thisRunDateStamp.AddYears(-1).AddMonths(+6).AddDays(+10)

#Collect total saved kb for reporting
$savedkbsthisrun = @()

Start-Transcript -Path $Logname -Append
Write-Host "Script started:" (Get-date)
Write-host "**********************" -ForegroundColor White

#FIRST check the storage, Ghostscript breaks pdf's when it hasn't got enough space to run - stop completely if we get below 10gb on X
$driveSpace = Get-PSDrive -Name D
If(($driveSpace.Free/1gb) -lt 15){
Write-Host "Not enough space...stopping" -ForegroundColor Red
Write-Error -Exception "Not enough space" -Message "There is less than 10gb left on the D drive"
Exit
}
Else{
Write-Host "Ok, there is enough space"
}

#Collections
$pdftoolongpaths = @()
$pdfissues = @()
$pdfsuccesslist = @()


#Find all directories that aren't symbolic links
#$folders = Get-ChildItem -Directory "D:\Clients" -Recurse | Where-Object { $_.LinkType -ne "SymbolicLink" }
$folders = Get-ChildItem -Directory "C:\pdf tests" -Recurse | Where-Object { $_.LinkType -ne "SymbolicLink" }

foreach($folder in $folders){
    $thisFolder = $folder
    #Find all .pdf files in each folder
    $file = Get-ChildItem -Path $thisFolder.FullName -Filter *.pdf 
    If($thisFolder.LastWriteTime -gt $lastRunDateStamp){
    foreach($pdf in $file){
#We want to check the potential path size of any files going in as ECO use very long path names - this could break pdfs. So we take the path and we add 11 which is the longest addition we add to file names.
$maxpotentialpathlength = $pdf.FullName.Length + 11
If($maxpotentialpathlength -ge 256){
Write-Host "The file path is too long to create new files in this directory, skipping...[$($thisfilename)]" -ForegroundColor Red
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "[$($thisfilename)]: The file path is too long to create new files in this directory"
$pdftoolongpaths += $thisfilename
Break
}
Else{
Write-Host "The file path is an okay length to create new files in this directory, continuing..." -ForegroundColor Green
}
Write-host "**********************" -ForegroundColor White
#Set some useful variables
$thisFile = $pdf
$Originalfilesize = ($thisFile.Length/1024)
$thisfilename = $thisFile.FullName
$thistemporaryfilename = $thisFile.DirectoryName + "\" + $thisFile.BaseName + "temp.pdf"
Write-Host "Compressing PDF: $($thisfilename), saving temporarily as $($thistemporaryfilename). We will re-save the document as the original name if compression is successful." -ForegroundColor Yellow
#Compress the PDF to a new file in the same location
        $exe = "C:\Program Files\gs\gs9.27\bin\gswin64c.exe"
        $result = &$exe ""-sDEVICE=pdfwrite -sOutputFile="$thistemporaryfilename" -dPDFSETTINGS=/ebook -dNOPAUSE -dPDFWRDEBUG -dBATCH -dSAFER $thisfilename""        
#Run some checks
$Aoutputpath = $thisFile.DirectoryName + "\Aoutput.txt"
$Boutputpath = $thisFile.DirectoryName + "\Boutput.txt"
&$exe ""-sDEVICE=txtwrite -o $Aoutputpath -dNOPAUSE $thisfilename""        
&$exe ""-sDEVICE=txtwrite -o $Boutputpath -dNOPAUSE $thistemporaryfilename""   
$a = Get-Content $Aoutputpath
[string]$a = [string]$a -replace " ",""
$b = Get-Content $Boutputpath
[string]$b = [string]$b -replace " ",""
If($a.Length -eq $b.Length){
    #I work
    Write-Host "[$($thisfilename)] - .txt content matches" -ForegroundColor Green
    #Try to retrieve the new compressed PDF in the same location, if it's not there it wasn't successful.
    Start-Sleep -s 5
    $newfile = get-item -Path $thistemporaryfilename
    $newfilesize = ($newfile.Length/1024)
    $sizedifferencepercent = [math]::Round(100/($Originalfilesize/$newfilesize))
    $sizedifferenceactual = [math]::Round($Originalfilesize-$newfilesize)
    $savedkbsthisrun  += $sizedifferenceactual
    If(($newfile) -and ($newfilesize -gt 0) -and ($result)){
        Write-Host "Looks like Ghostscript successfully created a new compressed PDF for: $($thisfilename).`r`n" -ForegroundColor Yellow
        Write-Host "The new file size is approximately $($sizedifferencepercent)% of the original size, $($sizedifferenceactual)kb was saved!`r`n" -ForegroundColor Green
        Write-Host "We will attempt to delete the old file and rename the temporary file to match - effectively replacing the original file.`r`n" -ForegroundColor Yellow
        Remove-Item -Path $thisfilename
        Remove-Item -Path $Aoutputpath
        Remove-Item -Path $Boutputpath
        Rename-Item -Path $thistemporaryfilename -NewName $thisfilename
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "[$($thisfilename)]: File compressed"
        $pdfsuccesslist += $thisfilename
        }
        Else{
        Write-Host "Doesn't look like Ghostscript could create the new compressed PDF as we can't find it, so we won't delete the old one" -ForegroundColor Red
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "[$($thisfilename)]: File not compressed by Ghostscript"
        $pdfissues += $thisfilename
        }
}
Else{
#I failed because there is missing content
Write-Host "[$($thisfilename)] - .txt content does not match the output..." -ForegroundColor Red
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "[$($thisfilename)]: File not compressed by Ghostscript"
$pdfissues += $thisfilename
}
Write-host "**********************" -ForegroundColor White
}
}
Else{
Write-Host "$($thisFile.Name) skipped: $($thisFolder.LastWriteTime)" -ForegroundColor Red
}
}



#show the results
$totalsavedsizethisrun = ($savedkbsthisrun | Measure-object -Sum)
Write-Host "Total saved space: $($totalsavedsizethisrun.sum) kb"

friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype MESSAGE -logstring "Files compressed successfully"
ForEach($pdfsuccess in $pdfsuccesslist){friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "$($pdfsuccess)"}
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype MESSAGE -logstring "Files with issues - we have not deleted the original or output"
ForEach($pdfissue in $pdfissues){friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "$($pdfissue)"}
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype MESSAGE -logstring "Files with paths that are too long - we have not attempted to compress them and have skipped them"
ForEach($pdftoolongpath in $pdftoolongpaths){friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "$($pdftoolongpaths)"}


Stop-Transcript

#Set-Content -Path D:\pdfArchive.log -Value $(Get-Date $thisRunDateStamp -Format "yyyy-MM-dd hh:mm:ss") -Force





















