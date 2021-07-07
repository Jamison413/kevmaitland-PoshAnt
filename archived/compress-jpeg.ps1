##################################################
#                                                #
#                                                #
#   jpeg-recompress - Automated jpg Compression  #
#                                                #
#                                                #
##################################################



$lastRunDateStamp = Get-Date $(Get-Content D:\jpegArchive.log)
$thisRunDateStamp = Get-Date

#If you want to set a manual datestamp
#$lastRunDateStamp = $thisRunDateStamp.AddMonths(-4)


#Find all folders that aren't symbolic links
$folders = Get-ChildItem -Directory "D:\Clients" -Recurse | Where-Object { $_.LinkType -ne "SymbolicLink" }
foreach($folder in $folders){
    $thisFolder = $folder
    #find any jpg files
    $file = Get-ChildItem -Path $thisFolder.FullName -Filter *.jpg 
    #Check the last date stamp as we don't want to unnecessarily re-run through files
    If($thisFolder.LastWriteTime -gt $lastRunDateStamp){
    foreach($jpg in $file){
        $thisFile = $jpg
        Write-Host "$($thisFile.FullName)" -ForegroundColor Yellow
        #Process the jpg file
        $result = $(& C:\Users\administrator.SUSTAINLTD\Desktop\jpeg-archive\jpeg-recompress.exe $thisFile.FullName $thisFile.FullName)
        write-host "$result"
    }
    }
    Else{
    Write-Host "$($thisFile.Name) skipped: $($thisFolder.LastWriteTime)" -ForegroundColor Red
    }
    }
    
#Set date on a centralised datestamp file to help reduce re-processing at the first stage of the script.
Set-Content -Path D:\jpegArchive.log -Value $(Get-Date $thisRunDateStamp -Format "yyyy-MM-dd hh:mm:ss") -Force
