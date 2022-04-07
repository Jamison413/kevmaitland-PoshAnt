###Notes for setup (example London printer)

1. Find the driver files online for the printer, download and extract
2. Find and open the INF file in text editor (e.g. Notepad), make a note of the .INF file name
3. Find the Driver name in the .INF file e.g. example from below "KONICA MINOLTA Universal PCL"

"
[Manufacturer]
%KM%=KONICA MINOLTA, NTamd64, NTamd64.6.0

[KONICA MINOLTA.NTamd64]
"KONICA MINOLTA Universal PCL" = KOAX8J_G.DLL.NTamd64, USBPRINT\KONICA_MINOLTAUniverFE33, KONICA_MINOLTAUniverFE33
"KONICA MINOLTA Universal PCL" = KOAX8J_G.DLL.NTamd64, USBPRINT\KONICA_MINOLTAC286iSA714, KONICA_MINOLTAC286iSA714
"KONICA MINOLTA Universal PCL" = KOAX8J_G.DLL.NTamd64, USBPRINT\KONICA_MINOLTAC287iS6BD5, KONICA_MINOLTAC287iS6BD5
"KONICA MINOLTA Universal PCL" = KOAX8J_G.DLL.NTamd64, USBPRINT\KONICA_MINOLTA4750iSB86C, KONICA_MINOLTA4750iSB86C
"KONICA MINOLTA Universal PCL" = KOAX8J_G.DLL.NTamd64, USBPRINT\KONICA_MINOLTA4700iCBE2, KONICA_MINOLTA4700iCBE2
"KONICA MINOLTA Universal PCL" = KOAX8J_G.DLL.NTamd64, USBPRINT\KONICA_MINOLTAC750iC104, KONICA_MINOLTAC750iC104
"KONICA MINOLTA Universal PCL" = KOAX8J_G.DLL.NTamd64, USBPRINT\KONICA_MINOLTA750iDD12, KONICA_MINOLTA750iDD12
"



4. Find the .cat file and make a note of the file name e.g. koax8j_.cat
5. Find the .cab file and make a note of the file name e.g. KOAX8J_.cab
6. Create a new directory with the driver folder (with all the files in the folder, including .INF, .cab, .cat and all the .dll files) copied, create install.ps1 and uninstall.ps1
7. Copy the example install.ps1 code into your install.ps1 new file, configure printer name, port and driver details
8. Copy the example Uninstall.ps1 code into your Uninstall.ps1 new file, configure printer name and port details
9. Run the IntuneWinAppUtil.exe tool:

	- point the source folder to the new directory with both the driver folder, install and uninstall files in
	- point the setup file as the install.ps1 file
	- point the output folder to the new directory with both the driver folder, install and uninstall files in - this is where the Install.intunewin file we'll upload to Intune will sit.
10. Find the output Install.intunewin and go to Intune > Apps > Windows > Add App
11. Select Windows app (Win32) and click select
12. Fill out win app details in Intune (I've omitted the basic stuff):

Install command
powershell.exe -ExecutionPolicy Bypass .\Install.ps1

Uninstall command
powershell.exe -ExecutionPolicy Bypass .\Uninstall.ps1

Install behavior
System

Device restart behavior
No specific action

For detection rules:

- select Registry type

Key path: 
Computer\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Print\Printers\[printer display name]
e.g. "Computer\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Print\Printers\GBR-London (BizHub)"

Value Name:
Name

Detection method:
String comparison

Operator:
Equals

Value:

Printer display name, e.g. GBR-London (BizHub)

Leave everything else as is for detection rule

13. Save and test deploy

