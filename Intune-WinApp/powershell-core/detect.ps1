$thisApp = "powershell-core"
$filePathToTest = "${env:ProgramFiles}\PowerShell\7\pwsh.exe"
#$thisApp = "%%PLACEHOLDERAPPNAME%%"
#$filePathToTest = "%%PLACEHOLDERDETECTIONFILE%%"
$scheduledTaskToTest = "$env:ProgramData\CustomScripts\redo-choco$thisApp-scheduledTaskCreated.log"

if(Test-Path $filePathToTest){
    if(Test-Path $scheduledTaskToTest){
        Write-Host "Path [$($filePathToTest)] and Task [$($scheduledTaskToTest)] both exist!"
        }
    else{
        Throw "Path [$($scheduledTaskToTest)] not found - [$($thisApp)] is not managed!"
        }
    }
else{
    Throw "Path [$($filePathToTest)] not found - [$($thisApp)] is not installed"
    }