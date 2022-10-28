if(Test-Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Printers\ESP-Manlleu-Planta2"){
    Write-Output "Printer [ESP-Manlleu-Planta2] installed"
    Exit 0
}
else{
    Write-Output "Printer [ESP-Manlleu-Planta2] failed to install - check [ESP-Manlleu-Planta2.log] in [%TEMP%]"
    Write-Error "Printer [ESP-Manlleu-Planta2] failed to install - check [ESP-Manlleu-Planta2.log] in [%TEMP%]"
    Exit 1
}
