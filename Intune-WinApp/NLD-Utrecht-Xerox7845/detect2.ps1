if(Test-Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Printers\NLD-Utrecht-Xerox7845"){
    Write-Output "Printer [NLD-Utrecht-Xerox7845] installed"
	Exit 0
}
else{
    Write-Output "Printer [NLD-Utrecht-Xerox7845] failed to install - check [NLD-Utrecht-Xerox7845.log] in [%TEMP%]"
    Write-Error "Printer [NLD-Utrecht-Xerox7845] failed to install - check [NLD-Utrecht-Xerox7845.log] in [%TEMP%]"
	Exit -1
}
