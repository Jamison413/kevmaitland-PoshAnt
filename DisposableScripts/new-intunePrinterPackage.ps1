#Download driver from manufacturer's website

#Find and open relevant .INF file (use guide at https://msendpointmgr.com/2022/01/03/install-network-printers-intune-win32apps-powershell/)

#Locate files required for install, paste below:
$arrayOfFiles = convertTo-arrayOfStrings "mfricr64.dl_
oemsetup.dsc
oemsetup.inf
rd05rd64.dl_
Readme.html
rica5x.cat
rica5Xcb.dl_
rica5Xcd.dl_
rica5Xcd.psz
rica5Xcf.cfz
rica5Xch.chm
rica5Xci.dl_
rica5Xcj.dl_
rica5Xcl.ini
rica5Xct.dl_
rica5Xcz.dlz
rica5Xgl.dl_
rica5Xgr.dl_
rica5Xlm.dl_
rica5Xtc.ex_
rica5Xtf.ex_
rica5Xtl.ex_
rica5Xtt.ex_
rica5Xug.dl_
rica5Xug.miz
rica5Xui.dl_
rica5Xui.irj
rica5Xui.rcf
rica5Xui.rdj
rica5Xur.dl_
ricdb64.dl_
rica5Xui.dll,rica5Xui.dl_
rica5Xui.irj
rica5Xui.rdj
rica5Xui.rcf
rica5Xug.dll,rica5Xug.dl_
rica5Xug.miz
rica5Xur.dll,rica5Xur.dl_
rica5Xgr.dll,rica5Xgr.dl_
rica5Xgl.dll,rica5Xgl.dl_
rica5Xci.dll,rica5Xci.dl_
rica5Xcd.dll,rica5Xcd.dl_
rica5Xcd.psz
rica5Xcf.cfz
rica5Xcl.ini
rica5Xch.chm
rica5Xcz.dlz
rica5Xcj.dll,rica5Xcj.dl_
rica5Xct.dll,rica5Xct.dl_
rica5Xcb.dll,rica5Xcb.dl_
rica5Xtl.exe,rica5Xtl.ex_
rica5Xtc.exe,rica5Xtc.ex_
rica5Xtt.exe,rica5Xtt.ex_
rica5Xtf.exe,rica5Xtf.ex_
"
#$arrayOfFiles += "koax5j__.cat" #Manually add Catalogue file name
#$arrayOfFiles += "KOAX5J__.inf" #Manually add Inf file name

#Set variables
$infFile = "oemsetup.inf"
$driverName = "RICOH MP C3004ex PCL 6"
$printerIP = "192.168.0.244"
$prettyPrinterName = "ESP-Manlleu-Planta2"

$newDir = New-Item -Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\GitHub\PoshAnt\Intune-WinApp\$prettyPrinterName" -ItemType Directory
#$newDir = Get-Item -Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\GitHub\PoshAnt\Intune-WinApp\$prettyPrinterName"
$arrayOfFiles | ForEach-Object {
    $thisFile = $_
    switch($thisFile.SubString($thisFile.Length - 3,3)){
<#        "in_" {$newFileName = $thisFile.TrimEnd("_")+"f"}
        "dl_" {$newFileName = $thisFile.TrimEnd("_")+"l"}
        "ex_" {$newFileName = $thisFile.TrimEnd("_")+"e"}
        "ch_" {$newFileName = $thisFile.TrimEnd("_")+"m"}
        "ca_" {$newFileName = $thisFile.TrimEnd("_")+"b"}
#>        default {$newFileName = $thisFile}
    }
    Copy-Item "$env:USERPROFILE\Downloads\z97499L16\disk1\$thisFile" -Destination "$($newDir.FullName)\$newFileName" -Force
}
Copy-Item "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\GitHub\PoshAnt\Install-Printer.ps1"-Destination $newDir.FullName -Force
Copy-Item "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\GitHub\PoshAnt\Remove-Printer.ps1"-Destination $newDir.FullName -Force

#Generate install, uninstall & detect scripts
"powershell.exe -executionpolicy bypass -file Install-Printer.ps1 -PortName `"IP_$printerIP`" -PrinterIP `"$printerIP`" -PrinterName `"$prettyPrinterName`" -DriverName `"$driverName`" -INFFile `"$infFile`"" | Out-File -FilePath "$($newDir.FullName)\install.ps1"
"powershell.exe -executionpolicy bypass -file Remove-Printer.ps1 -PrinterName `"$prettyPrinterName`"" | Out-File -FilePath "$($newDir.FullName)\uninstall.ps1"
"if(Test-Path `"HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Printers\$prettyPrinterName`"){
    Write-Output `"Printer [$($prettyPrinterName)] installed`"
    Exit 0
}
else{
    Write-Output `"Printer [$($prettyPrinterName)] failed to install - check [$prettyPrinterName.log] in [%TEMP%]`"
    Write-Error `"Printer [$($prettyPrinterName)] failed to install - check [$prettyPrinterName.log] in [%TEMP%]`"
    Exit 1
}" | Out-File -FilePath "$($newDir.FullName)\detect.ps1"
#Open detect.ps1 in Notepad++ and change the encoding to UTF8


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