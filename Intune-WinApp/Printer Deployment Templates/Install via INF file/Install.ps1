#Install.ps1

##############################################
#  London Universal Printer Driver Install   #
##############################################
$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$drivername = "KONICA MINOLTA Universal PCL" #driver name decalred in .inf file
$portName = "TCPIP:GBR-London-BizHub"
$PortAddress = "" #IP address of printer in network

###################
#      Staging    #
###################
C:\Windows\SysNative\pnputil.exe /add-driver "$($PSScriptRoot)\win_x64\KOAX8J__.inf" /install

#######################
#     Installing      #
#######################

Add-PrinterDriver -Name $drivername

##########################################################
#     check if the port already exist, else install      #
##########################################################
$checkPortExists = Get-Printerport -Name $portname -ErrorAction SilentlyContinue
if (-not $checkPortExists) 
{
Add-PrinterPort -name $portName -PrinterHostAddress $PortAddress
}

########################################
#     Check if Driver Exists install   #
########################################
$printDriverExists = Get-PrinterDriver -name $DriverName -ErrorAction SilentlyContinue

if ($printDriverExists)
{
Add-Printer -Name "GBR-London (BizHub)" -PortName $portName -DriverName $DriverName
}
else
{
Write-Warning "Printer driver not installed"
write-host $($error)
}


