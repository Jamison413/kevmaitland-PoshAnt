if(Test-Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Printers\GBR-Bristol-GroundFloor"){
    Write-Output "Printer [GBR-Bristol-GroundFloor] installed"
    Exit 0
}
else{
    Write-Output "Printer [GBR-Bristol-GroundFloor] failed to install - check [GBR-Bristol-GroundFloor.log] in [%TEMP%]"
    Write-Error "Printer [GBR-Bristol-GroundFloor] failed to install - check [GBR-Bristol-GroundFloor.log] in [%TEMP%]"
    Exit 1
}
