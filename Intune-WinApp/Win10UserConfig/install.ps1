#region Set Remote Desktop Gateway preferences
#Build out HKCU:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\ path if required
if((Test-Path 'HKCU:\SOFTWARE\Policies\Microsoft\Windows NT') -eq $false){
    New-Item –Path "HKCU:\SOFTWARE\Policies\Microsoft\" –Name "Windows NT"
    }
if((Test-Path 'HKCU:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services') -eq $false){
    New-Item –Path "HKCU:\SOFTWARE\Policies\Microsoft\Windows NT\" –Name "Terminal Services"
    }
if((Test-Path 'HKCU:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services') -eq $false){
    New-Item –Path "HKCU:\SOFTWARE\Policies\Microsoft\Windows NT\" –Name "Terminal Services"
    }
#Set UseProxy = 1 [Administrative Templates (Users) | Windows Components | Remote Desktop Services | RD Gateway | Enable connection through RD Gateway]
if((Get-ItemProperty 'HKCU:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\' -Name UseProxy -ErrorAction SilentlyContinue) -eq $null){
    New-ItemProperty –Path "HKCU:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\" –Name "UseProxy" -Value 1 -PropertyType DWORD
    }
else{
    Set-ItemProperty –Path "HKCU:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\" –Name "UseProxy" -Value 1 
    }
#Set AllowExplicitUseProxy = 1 [Administrative Templates (Users) | Windows Components | Remote Desktop Services | RD Gateway | Enable connection through RD Gateway | Allow users to change this setting]
if((Get-ItemProperty 'HKCU:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\' -Name AllowExplicitUseProxy -ErrorAction SilentlyContinue) -eq $null){
    New-ItemProperty –Path "HKCU:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\" –Name "AllowExplicitUseProxy" -Value 1 -PropertyType DWORD
    }
else{
    Set-ItemProperty –Path "HKCU:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\" –Name "AllowExplicitUseProxy" -Value 1 
    }
#Set ProxyName = "remote.sustain.co.uk" [Administrative Templates (Users) | Windows Components | Remote Desktop Services | RD Gateway | Set RD Gateway server address]
if((Get-ItemProperty 'HKCU:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\' -Name ProxyName -ErrorAction SilentlyContinue) -eq $null){
    New-ItemProperty –Path "HKCU:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\" –Name "ProxyName" -Value "remote.sustain.co.uk" -PropertyType String
    }
else{
    Set-ItemProperty –Path "HKCU:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\" –Name "ProxyName" -Value "remote.sustain.co.uk" 
    }
#Set AllowExplicitProxyName = 1 [Administrative Templates (Users) | Windows Components | Remote Desktop Services | RD Gateway | Set RD Gateway server address | Allow users to change this setting]
if((Get-ItemProperty 'HKCU:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\' -Name AllowExplicitProxyName -ErrorAction SilentlyContinue) -eq $null){
    New-ItemProperty –Path "HKCU:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\" –Name "AllowExplicitProxyName" -Value 1 -PropertyType DWORD
    }
else{
    Set-ItemProperty –Path "HKCU:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\" –Name "AllowExplicitProxyName" -Value 1 
    }
#endregion

#region Set Explorer preferences
#Disable Group By for Downloads folder
    #Remove default Bags for Downloads folder (885A186E-A440-4ADA-812B-DB871B942259)
    $Bags = 'HKCU:\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\Bags'
    $DLID = '{885A186E-A440-4ADA-812B-DB871B942259}'
    (Get-ChildItem $bags -recurse | ? PSChildName -like $DLID ) | Remove-Item
    
    #Build out HKCU:\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{885A186E-A440-4ADA-812B-DB871B942259}\ path if required
    if((Test-Path 'HKCU:\SOFTWARE\Microsoft\Windows\Shell\Bags') -eq $false){
        New-Item –Path "HKCU:\SOFTWARE\Microsoft\Windows\Shell\" –Name "Bags"
        }
    if((Test-Path 'HKCU:\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders') -eq $false){
        New-Item –Path "HKCU:\SOFTWARE\Microsoft\Windows\Shell\Bags\" –Name "AllFolders"
        }
    if((Test-Path 'HKCU:\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell') -eq $false){
        New-Item –Path "HKCU:\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\" –Name "Shell"
        }
    if((Test-Path 'HKCU:\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{885A186E-A440-4ADA-812B-DB871B942259}') -eq $false){
        New-Item –Path "HKCU:\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\" –Name "{885A186E-A440-4ADA-812B-DB871B942259}"
        }
    #Set Mode = 4 [No GPO equivalent]
    if((Get-ItemProperty 'HKCU:\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{885A186E-A440-4ADA-812B-DB871B942259}\' -Name Mode -ErrorAction SilentlyContinue) -eq $null){
        New-ItemProperty –Path "HKCU:\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{885A186E-A440-4ADA-812B-DB871B942259}\" –Name "Mode" -Value 4 -PropertyType DWORD
        }
    else{
        Set-ItemProperty –Path "HKCU:\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{885A186E-A440-4ADA-812B-DB871B942259}\" –Name "Mode" -Value 4 -Confirm:$false
        }
    #Set GroupView = 0 [No GPO equivalent]
    if((Get-ItemProperty 'HKCU:\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{885A186E-A440-4ADA-812B-DB871B942259}\' -Name GroupView -ErrorAction SilentlyContinue) -eq $null){
        New-ItemProperty –Path "HKCU:\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{885A186E-A440-4ADA-812B-DB871B942259}\" –Name "GroupView" -Value 0 -PropertyType DWORD
        }
    else{
        Set-ItemProperty –Path "HKCU:\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{885A186E-A440-4ADA-812B-DB871B942259}\" –Name "GroupView" -Value 0 -Confirm:$false
        }

#Set default view settings
    #Enable Hidden Items
    Set-ItemProperty –Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\" –Name "Hidden" -Value 1 -Confirm:$false
    #Unhide File Extensions
    Set-ItemProperty –Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\" –Name "HideFileExt" -Value 0 -Confirm:$false
    #Hide Cortana button
    Set-ItemProperty –Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\" –Name "ShowCortanaButton" -Value 0 -Confirm:$false
    #Hide Task View button
    Set-ItemProperty –Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\" –Name "ShowTaskViewButton" -Value 0 -Confirm:$false
    #Hide StoreAppsOnTaskbar button
    Set-ItemProperty –Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\" –Name "StoreAppsOnTaskbar" -Value 0 -Confirm:$false


#Restart Explorer
    gps explorer | spps


#endregion

#region Deploy Desktop Tools
    #Deploy Splashtop
    if((Test-Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop") -eq $false){
        Copy-Item -Path ".\Anthesis SplashTop Remote Support.exe" -Destination "$env:USERPROFILE\Desktop" -Force
        Remove-Item -Path "$env:USERPROFILE\Desktop\Anthesis_SplashTop_Support.exe" -Force -Confirm:$false -ErrorAction SilentlyContinue
        }
    else{
        Copy-Item -Path ".\Anthesis SplashTop Remote Support.exe" -Destination "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop" -Force
        Remove-Item -Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\Anthesis_SplashTop_Support.exe" -Force -Confirm:$false -ErrorAction SilentlyContinue
        }

    #Deploy WUMT
    if((Test-Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents") -eq $false){
        if((Test-Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\WUMT") -eq $false){
            New-Item -Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\WUMT" -ItemType Directory
            }
        Copy-Item -Path ".\wumt_x64.exe" -Destination "$env:USERPROFILE\Documents\WUMT" -Force -Confirm:$false
        }
    else{
        if((Test-Path "$env:USERPROFILE\Documents\WUMT") -eq $false){
            New-Item -Path "$env:USERPROFILE\Documents\WUMT" -ItemType Directory
            }
        Copy-Item -Path ".\wumt_x64.exe" -Destination "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents" -Force
        }

#endregion

