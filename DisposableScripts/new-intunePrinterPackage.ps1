#Download driver from manufacturer's website

#Find and open relevant .INF file (use guide at https://msendpointmgr.com/2022/01/03/install-network-printers-intune-win32apps-powershell/)

#Locate files required for install, paste below:
$arrayOfFiles = convertTo-arrayOfStrings "KOAX5J_G.DLL
KOAX5J_X.DLL
KOAX5J_F.DLL
KOAX5J_C.DLL
KOAX5J_U.DLL
KOAX5J_S.DLL
KOAX5J_R.DLL
KOAX5J_J.DLL
KOAX5J_D.DLL
KOAX5J__ZH-TW.chm
KOAX5J_D.exe
KOAX5J_O.exe
KOBDrvAPIIF.dll
KOBDrvAPIIF32.dll
KOBDrvAPIW64.exe
"
$arrayOfFiles += "koax5j__.cat" #Manually add Catalogue file name
$arrayOfFiles += "KOAX5J__.inf" #Manually add Inf file name

#Set variables
$infFile = "KOAX5J__.inf"
$driverName = "KONICA MINOLTA Universal PCL"
$printerIP = "192.168.93.53"
$prettyPrinterName = "GBR-London-KonicaC257i"

$newDir = New-Item -Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\GitHub\PoshAnt\Intune-WinApp\$prettyPrinterName" -ItemType Directory
#$newDir = Get-Item -Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\GitHub\PoshAnt\Intune-WinApp\$prettyPrinterName"
$arrayOfFiles | ForEach-Object {
    $thisFile = $_
    switch($thisFile.SubString($thisFile.Length - 3,3)){
        "in_" {$newFileName = $thisFile.TrimEnd("_")+"f"}
        "dl_" {$newFileName = $thisFile.TrimEnd("_")+"l"}
        "ex_" {$newFileName = $thisFile.TrimEnd("_")+"e"}
        "ch_" {$newFileName = $thisFile.TrimEnd("_")+"m"}
        "ca_" {$newFileName = $thisFile.TrimEnd("_")+"b"}
    }
    Copy-Item "$env:USERPROFILE\Downloads\UPDPCL6Win_392120MU\driver\win_x64\$_" -Destination "$($newDir.FullName)\$newFileName" -Force
}
Copy-Item "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\GitHub\PoshAnt\Install-Printer.ps1"-Destination $newDir.FullName -Force
Copy-Item "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\GitHub\PoshAnt\Remove-Printer.ps1"-Destination $newDir.FullName -Force

#Generate install, uninstall & detect scripts
"powershell.exe -executionpolicy bypass -file Install-Printer.ps1 -PortName `"IP_$printerIP`" -PrinterIP `"$printerIP`" -PrinterName `"$prettyPrinterName`" -DriverName `"$driverName`" -INFFile `"$infFile`"" | Out-File -FilePath "$($newDir.FullName)\install.ps1"
"powershell.exe -executionpolicy bypass -file Remove-Printer.ps1 -PrinterName `"$prettyPrinterName`"" | Out-File -FilePath "$($newDir.FullName)\uninstall.ps1"
"if(Test-Path `"HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Printers\$prettyPrinterName`"){
    Write-Output `"Printer [$($prettyPrinterName)] installed`"
}
else{
    Throw `"Printer [$($prettyPrinterName)] failed to install - check [$prettyPrinterName.log] in [$env:TEMP]`"
}" | Out-File -FilePath "$($newDir.FullName)\detect.ps1"

#Build the Win32 package
Start-Process -NoNewWindow -FilePath "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\GitHub\PoshAnt\Intune-WinApp\IntuneWinAppUtil.exe" -ArgumentList "-c `"$($newDir.FullName)`"","-s install.ps1","-o `"$($newDir.FullName)`"","-q"

#region Create Security Group
$tokenResponseTeamBot = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName TeamsBot) -grant_type client_credentials
$newSG = new-graphGroup -tokenResponse $tokenResponseTeamBot -groupDisplayName "Printer - $prettyPrinterName" -groupDescription "Security Group for deploying Printer [$prettyPrinterName]" -groupType Security -membershipType Assigned
#endregion

<# This is essentially what the Install-Printer.ps1 script does:
pnputil.exe /add-driver $infFile
Add-PrinterPort -Name "IP_$printerIP" -PrinterHostAddress $printerIP
Add-PrinterDriver -DriverName $driverName
Add-Printer -Name $prettyPrinterName -DriverName $driverName -PortName "IP_$printerIP"
#>