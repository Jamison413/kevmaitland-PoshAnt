#Remove Junk Apps
Get-AppxPackage | ? {$_.Publisher -notmatch "Microsoft Corporation"} | Remove-AppxPackage 
Get-AppxPackage | ? {$_.Name -cmatch "Xbox"} | Remove-AppxPackage 
Get-AppxPackage | ? {$_.Name -cmatch "Zune"} | Remove-AppxPackage 
Get-AppxPackage | ? {$_.Name -cmatch "Bing"} | Remove-AppxPackage 
Get-AppxPackage | ? {$_.Name -cmatch "SkypeApp"} | Remove-AppxPackage 
Get-AppxPackage | ? {$_.Name -cmatch "OneNote"} | Remove-AppxPackage 
Get-AppxPackage | ? {$_.Name -cmatch "Solitaire"} | Remove-AppxPackage 
Get-AppxPackage -AllUsers | ? {$_.Publisher -notmatch "Microsoft Corporation"} | Remove-AppxPackage -AllUsers
Get-AppxPackage -AllUsers | ? {$_.Name -cmatch "Xbox"} | Remove-AppxPackage -AllUsers
Get-AppxPackage -AllUsers | ? {$_.Name -cmatch "Zune"} | Remove-AppxPackage -AllUsers
Get-AppxPackage -AllUsers | ? {$_.Name -cmatch "Bing"} | Remove-AppxPackage -AllUsers
Get-AppxPackage -AllUsers | ? {$_.Name -cmatch "SkypeApp"} | Remove-AppxPackage -AllUsers
Get-AppxPackage -AllUsers | ? {$_.Name -cmatch "OneNote"} | Remove-AppxPackage -AllUsers

#Block Consumer Junk
New-Item -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows -Name CloudContent
New-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent -Name DisableWindowsConsumerFeatures -PropertyType $([Microsoft.Win32.RegistryValueKind]::DWord) -Value 1
New-Item -Path HKLM:\SOFTWARE\Policies\Microsoft -Name WindowsStore
New-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\WindowsStore -Name RemoveWindowsStore -PropertyType $([Microsoft.Win32.RegistryValueKind]::DWord) -Value 1
New-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\WindowsStore -Name DisableStoreApps -PropertyType $([Microsoft.Win32.RegistryValueKind]::DWord) -Value 1
New-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\WindowsStore -Name DisableOSUpgrade -PropertyType $([Microsoft.Win32.RegistryValueKind]::DWord) -Value 1


#DeProvision junk
$dismPackages = dism /online /Get-ProvisionedAppxPackages | Select-String PackageName
$dismPackages | ?{$_.Line -match "xbox"} |  %{& DISM /Online /Remove-ProvisionedAppxPackage /PackageName:$($_.Line.Replace("PackageName : ",''))}
$dismPackages | ?{$_.Line -match "zune"} |  %{& DISM /Online /Remove-ProvisionedAppxPackage /PackageName:$($_.Line.Replace("PackageName : ",''))}
$dismPackages | ?{$_.Line -match "skypeapp"} |  %{& DISM /Online /Remove-ProvisionedAppxPackage /PackageName:$($_.Line.Replace("PackageName : ",''))}
$dismPackages | ?{$_.Line -match "onenote"} |  %{& DISM /Online /Remove-ProvisionedAppxPackage /PackageName:$($_.Line.Replace("PackageName : ",''))}


<#Remove OneDrive#>
.\taskkill.exe /im:onedrive.exe
if(Test-Path $env:SystemRoot\System32\OneDriveSetup.exe){& $env:SystemRoot\System32\OneDriveSetup.exe /uninstall}
if(Test-Path $env:SystemRoot\SysWOW64\OneDriveSetup.exe){& $env:SystemRoot\SysWOW64\OneDriveSetup.exe /uninstall}
Remove-Item $env:USERPROFILE\OneDrive -Force -Recurse
Remove-Item $env:LOCALAPPDATA\Microsoft\OneDrive -Force -Recurse
Remove-Item "$env:ProgramData\Microsoft OneDrive" -Force -Recurse
New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT
Remove-Item -Path 'HKCR:\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}' -Force -Recurse
Remove-Item -Path 'HKCR:\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}' -Force -Recurse

#Unpin as many apps from the Start Menu as possible
$apps = New-Object -ComObject Shell.Application
$apps.NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | % {
    $_.Verbs() | ?{$_.Name.Replace('&','') -match 'From "Start" UnPin|Unpin from Start'} | % {$_.DoIt()}
    }

#Show file extensions
New-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced -Name HideFileExt -PropertyType $([Microsoft.Win32.RegistryValueKind]::DWord) -Value 0

