$matchThisFolderName = "Emails"
$fsrmExceptionGroups = "E-mail Files"
RecurseThis("D:\Clients")


function RecurseThis ([string]$currentFolderPath){
    #Write-Host "Calling Recurse function"
    if (-not (Test-Path $currentFolderPath)) 
        {
        Write-Error "$currentFolderPath is an invalid path."
        return $false
        }

    $currentFolder = Get-Item $currentFolderPath
    if ($matchThisFolderName -eq $currentFolder.Name){
        New-FsrmFileScreenException -Path $currentFolderPath -IncludeGroup $fsrmExceptionGroups -Description "Automagically generated via PowerShell"
        }
    
    foreach ($folder in $currentFolder.GetDirectories())
        {
        RecurseThis($folder.FullName)
        }
    }