if((Test-Path "$env:ProgramFiles\Bluebeam Software\Bluebeam Revu\20\Revu\Revu.exe") -eq $true){Write-Host "BlueBeam Revu installed correctly"}
else{throw "BlueBeam did not install :'("}