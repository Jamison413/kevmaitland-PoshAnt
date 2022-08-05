if(Test-Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Printers\GBR-London-KonicaC257i"){
    Write-Output "Printer [GBR-London-KonicaC257i] installed"
    Exit 0
}
else{
    Write-Output "Printer [GBR-London-KonicaC257i] failed to install - check [GBR-London-KonicaC257i.log] in [C:\Users\KEVMAI~1\AppData\Local\Temp]"
    Write-Error "Printer [GBR-London-KonicaC257i] failed to install - check [GBR-London-KonicaC257i.log] in [C:\Users\KEVMAI~1\AppData\Local\Temp]"
    Exit 1
}
