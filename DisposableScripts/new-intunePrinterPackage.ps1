#Download driver from manufacturer's website

#Find and open relevant .INF file (use guide at https://msendpointmgr.com/2022/01/03/install-network-printers-intune-win32apps-powershell/)

#Locate files required for install, paste below:
$arrayOfFiles = convertTo-arrayOfStrings "x3up02T.dll
x3rpcl02T.dll
x3wfuv02T.dll
x3gui02T.dll
x3core02T.dll
x3util02T.dll
x3rnut02T.dll
x3txt02T.cab
x3coms02T.dll
x3jobt02T.exe
x3thpr02T.exe
x3ptpc02T.dll
x3fput02T.dll
xUNIVX.tag
x3UNIV02T.cab
x3JAR02T.cab 
x3fpb02T.exe
x2fpd02.dll
x3encr02T.dll 
x5print.dll
x5pp.dll
x5lrs.dll
x5lrsl.dll
api-ms-win-core-file-l1-2-0.dll
api-ms-win-core-file-l2-1-0.dll
api-ms-win-core-localization-l1-2-0.dll
api-ms-win-core-processthreads-l1-1-1.dll
api-ms-win-core-synch-l1-2-0.dll
api-ms-win-core-timezone-l1-1-0.dll
ucrtbase.dll
api-ms-win-crt-conio-l1-1-0.dll
api-ms-win-crt-convert-l1-1-0.dll
api-ms-win-crt-environment-l1-1-0.dll
api-ms-win-crt-filesystem-l1-1-0.dll
api-ms-win-crt-heap-l1-1-0.dll
api-ms-win-crt-locale-l1-1-0.dll
api-ms-win-crt-math-l1-1-0.dll
api-ms-win-crt-multibyte-l1-1-0.dll
api-ms-win-crt-private-l1-1-0.dll
api-ms-win-crt-process-l1-1-0.dll
api-ms-win-crt-runtime-l1-1-0.dll
api-ms-win-crt-stdio-l1-1-0.dll
api-ms-win-crt-string-l1-1-0.dll
api-ms-win-crt-time-l1-1-0.dll
api-ms-win-crt-utility-l1-1-0.dll
xUNIVX02T.gpd
xUNIVX02T.ini
xUNIVX02T.cfg
"
$arrayOfFiles += "x3UNIVx.cat" #Manually add Catalogue file name
$arrayOfFiles += "x3UNIVX.inf" #Manually add Inf file name

#Set variables
$infFile = "x3UNIVX.inf"
$driverName = "Xerox Global Print Driver PCL6"
$printerIP = "10.18.28.250"
$prettyPrinterName = "NLD-Utrecht-Xerox7845"

$newDir = New-Item -Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\GitHub\PoshAnt\Intune-WinApp\$prettyPrinterName" -ItemType Directory
#$newDir = Get-Item -Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\GitHub\PoshAnt\Intune-WinApp\$prettyPrinterName"
$arrayOfFiles | ForEach-Object {
    Copy-Item "$env:USERPROFILE\Downloads\UNIV_5.887.3.0_PCL6_x64\UNIV_5.887.3.0_PCL6_x64_Driver.inf\$_" -Destination $newDir.FullName -Force
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