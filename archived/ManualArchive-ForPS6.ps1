##############################################
#                                            #
#                                            #
#              Manual Archive                #
#                                            #
#                                            #
##############################################







<#Sort the logging#>
$logName = "Manual Archive $(Get-Date -Format "yyMMdd").log"
$VerbosePreference = "Continue"
<#Find all Archive Folders that aren't Symbolic links#>
$Target = "D:\Clients"
#$Target = "D:\Clients\Npower Northern Ltd"
#$Target = "D:\Clients\Npower Northern Ltd\102159-NPower ECO3 Delivery 2019\Reports\Audits\Duplicates\August 2019\QBR7977001\RE-submitted Docs"
$standardArchiveLocationscollection = Get-ChildItem $Target "*Archive" -Recurse | Where-Object {$_.LinkType -ne "SymbolicLink"}
$standardArchiveLocations = $standardArchiveLocationscollection.FullName
$standardArchiveLocations = $standardArchiveLocations.Where({$_ -notmatch "EDF Energy"})
$standardArchiveLocations = $standardArchiveLocations.Where({$_ -notmatch "Ovo Energy"})
$standardArchiveLocations | Out-File 'C:\content.txt'

function archive-folder($folderToArchive){
    Invoke-Command -ScriptBlock {& "C:\Windows\system32\robocopy.exe" "$($folderToArchive.Replace('X:\','D:\'))" "$($folderToArchive.Replace('X:\','E:\X\').Replace('D:\','E:\X\').Replace('Y:\','E:\X\'))" "/E" "/MOVE" "/COPY:DATSO" "/DCOPY:DAT" "/XJD" "/LOG+:C:\Scripts\Logs\$logName" "/R:1" "/W:1" "/NP" "/B"} 
    #/E - Include all subfolders
    #/MOVE - Delete source after successful copy
    #/COPY:DATSO - Copy Data, Attributes, Timestamps, Security & Owner for files. Security is for auditing as the Archives Share only had Read permissions
    #/DCOPY:DAT - Copy Data, Attributes, Timestamps for directories. 
    #/XJD - Exclude Junctions for directories (we create SymLinks later to point to the archive location, and we don't want to follow any of these a second time if we're re-archiving a directory)
    #/LOG - Log what happens
    #/R - Retry 1 time. It'll either work or it won't - there's not much point in retrying this more than once.
    #/W - Wait 1 second. If the file's locked, it's unlikely to be unlocked in 30 seconds, so just retry faster in case it was a network issue.
    #/NP - Don't fill the log file with progress percentages for large files
    #/B - Use backup mode, just in case there are permission problems
    }

#Map a drive if it's not already there
#if((Test-Path -Path "X:\") -eq $false){Invoke-Command -ScriptBlock {& "net" "use X: \\sustainltd.local\data"}}
if((Test-Path -Path "X:\") -eq $false){New-PSDrive -PSProvider FileSystem -Root "\\sustainltd.local\data" -Name X }
if((Test-Path -Path "Y:\") -eq $false){New-PSDrive -PSProvider FileSystem -Root "E:\X" -Name Y }
#Try this in cmd (not as Admin) if it still deosn't work :)
#Subst Y: E:\X


$standardArchiveLocations | %{
    write-host -ForegroundColor Yellow "Archiving $_"
    archive-folder -folderToArchive $_
    #Recreate archived folder and create SymLink inside it
    New-Item -Path $_ -ItemType Directory
    New-Item -Path $($_+"\"+$(Split-Path $_ -Leaf)) -ItemType SymbolicLink -Value $($_.Replace('X:\','\\sustainltd.local\ArchivedData\Archives\X\'))
    }
