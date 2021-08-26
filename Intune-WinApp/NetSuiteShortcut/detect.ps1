$thisApp = "NetSuite shortcut"
$filePathToTest = "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\NetSuite (SSO).url"
$scheduledTaskToTest = "$env:ProgramData\CustomScripts\redo-choco$thisApp-scheduledTaskCreated.log"

if(Test-Path $filePathToTest){
    Write-Host "Path [$($filePathToTest)] exists!"
    }
else{
    Throw "Path [$($filePathToTest)] not found - [$($thisApp)] is not installed"
    }