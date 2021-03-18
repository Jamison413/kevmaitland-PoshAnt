$thisApp = "Google Chrome"
$filePathToTest = "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
#$thisApp = "%%PLACEHOLDERAPPNAME%%"
#$filePathToTest = "%%PLACEHOLDERDETECTIONFILE%%"
#$scheduledTaskToTest = "$env:ProgramData\CustomScripts\redo-choco$thisApp-scheduledTaskCreated.log"

if($(Test-Path $filePathToTest) -eq $false){
    Throw "Path [$($filePathToTest)] not found - [$($thisApp)] is not installed"
    }
else{
    Write-Host "Path [$($filePathToTest)] exists!"
    }