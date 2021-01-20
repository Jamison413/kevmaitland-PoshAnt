$thisApp = "splashtopstreamer"
$filePathToTest = "${env:ProgramFiles(x86)}\Splashtop\Splashtop Remote\Server\SRServer.exe"
#$thisApp = "%%PLACEHOLDERAPPNAME%%"
#$filePathToTest = "%%PLACEHOLDERDETECTIONFILE%%"

if(Test-Path $filePathToTest){
     Write-Host "Path [$($filePathToTest)] both exist!"
    }
else{
    Throw "Path [$($filePathToTest)] not found - [$($thisApp)] is not installed"
    }
