#Download driver from manufacturer's website

#Find and open relevant .INF file (use guide at https://msendpointmgr.com/2022/01/03/install-network-printers-intune-win32apps-powershell/)

#Locate files required for install, paste below:
$arrayOfFiles = convertTo-arrayOfStrings "cnp60m.cat
CNP60MA64.INF
gppcl6.cab
"
#$arrayOfFiles += "koax5j__.cat" #Manually add Catalogue file name
#$arrayOfFiles += "KOAX5J__.inf" #Manually add Inf file name

#Set variables
$infFile = "CNP60MA64.inf"
$driverName = "Canon Generic Plus PCL6"
$printerIP = "94.190.240.217"
$prettyPrinterName = "GBR-Bristol-5thFloor"

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
    Copy-Item "$env:USERPROFILE\Downloads\GPlus_PCL6_Driver_V260_32_64_00\x64\Driver\$_" -Destination "$($newDir.FullName)\$newFileName" -Force
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