if(Test-Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Printers\GBR-Bristol-5thFloor"){
    Write-Output "Printer [GBR-Bristol-5thFloor] installed"
    Exit 0
}
else{
    Write-Output "Printer [GBR-Bristol-5thFloor] failed to install - check [GBR-Bristol-5thFloor.log] in [%TEMP%]"
    Write-Error "Printer [GBR-Bristol-5thFloor] failed to install - check [GBR-Bristol-5thFloor.log] in [%TEMP%]"
    Exit 1
}
