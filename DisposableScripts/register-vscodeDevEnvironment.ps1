#As User
$adminAccountLocalProfile = "t1-$env:USERNAME"
if($(Test-Path "$env:SystemDrive\Users\$adminAccountLocalProfile") -eq $false){
    New-Item -Path "$env:SystemDrive\Users\" -Name $adminAccountLocalProfile -ItemType Directory
    }
if($(Test-Path "$env:SystemDrive\Users\$adminAccountLocalProfile\Documents") -eq $false){
    New-Item -Path "$env:SystemDrive\Users\$adminAccountLocalProfile" -Name "Documents" -ItemType Directory
    }
if($(Test-Path "$env:SystemDrive\Users\$adminAccountLocalProfile\Documents\PowerShell") -eq $false){
    New-Item -Path "$env:SystemDrive\Users\$adminAccountLocalProfile\Documents" -Name "PowerShell" -ItemType Directory
    }
if($(Test-Path "$env:SystemDrive\Users\$adminAccountLocalProfile\Documents\PowerShell\Modules") -eq $false){
    if($(Test-Path "$env:SystemDrive\Users\$env:USERNAME\Documents\WindowsPowerShell\Modules") -eq $false){ #If the old 5.1 
        New-Item -Path "$env:SystemDrive\Users\$adminAccountLocalProfile\Documents\PowerShell" -Name "Modules" -ItemType Directory
        }
    else{
        New-Item -Name Modules -Path $env:UserProfile\OneDrive – Anthesis LLC\Documents\PowerShell\ -Value $env:UserProfile\OneDrive – Anthesis LLC\Documents\WindowsPowerShell\Modules\ -ItemType SymbolicLink
        }
    }


<# Add T1 account as local admin temporarily
Run this as local admin
net localgroup administrators /add azuread\t1-$(whoami /upn)

#>

#As T1
$userAccountLocalProfile = $env:USERNAME.Trim("T1-").Trim("T2-").Trim("T0-")
$adminAccountLocalProfile = $env:USERNAME

<#
Runas /user:azuread\$(whoami /upn) powershell.exe
Start-Process Powershell -Verb runAs
#>

if($(Test-Path "$env:SystemDrive\Users\$adminAccountLocalProfile") -eq $false){
    New-Item -Path "$env:SystemDrive\Users\" -Name $adminAccountLocalProfile -ItemType Directory
    }
if($(Test-Path "$env:SystemDrive\Users\$adminAccountLocalProfile\Documents") -eq $false){
    New-Item -Path "$env:SystemDrive\Users\$adminAccountLocalProfile" -Name "Documents" -ItemType Directory
    }
if($(Test-Path "$env:SystemDrive\Users\$adminAccountLocalProfile\Documents\PowerShell") -eq $false){
    New-Item -Path "$env:SystemDrive\Users\$adminAccountLocalProfile\Documents" -Name "PowerShell" -ItemType Directory
    }
if($(Test-Path "$env:SystemDrive\Users\$adminAccountLocalProfile\Documents\PowerShell\Modules") -eq $false){
    if($(Test-Path "$env:SystemDrive\Users\$userAccountLocalProfile\Documents\WindowsPowerShell\Modules") -eq $false){ #If the old 5.1 Modules are missing (without OneDrive redirection)
        if($(Test-Path "$env:SystemDrive\Users\$userAccountLocalProfile\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules") -eq $false){ #And if the old 5.1 Modules are present (with OneDrive redirection)
            New-Item -Path "$env:SystemDrive\Users\$adminAccountLocalProfile\Documents\PowerShell" -Name "Modules" -ItemType Directory #Just create a new empty Modules folder
            }
        else{
            New-Item -Name Modules -Path "$env:SystemDrive\Users\$adminAccountLocalProfile\Documents\PowerShell\" -Value "$env:SystemDrive\Users\$userAccountLocalProfile\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules" -ItemType SymbolicLink #SymbolicLink to the redirected Modules folder
            }
        }
    else{
        New-Item -Name Modules -Path "$env:SystemDrive\Users\$adminAccountLocalProfile\Documents\PowerShell\" -Value "$env:SystemDrive\Users\$userAccountLocalProfile\Documents\WindowsPowerShell\Modules" -ItemType SymbolicLink #SymbolicLink to the *un*redirected Modules folder
        }
    }

if($(Test-Path "$env:SystemDrive\Users\$adminAccountLocalProfile\Documents\WindowsPowerShell") -eq $false){
    New-Item -Path "$env:SystemDrive\Users\$adminAccountLocalProfile\Documents" -Name "WindowsPowerShell" -ItemType Directory
    }
if($(Test-Path "$env:SystemDrive\Users\$adminAccountLocalProfile\Documents\WindowsPowerShell\Modules") -eq $false){
    if($(Test-Path "$env:SystemDrive\Users\$userAccountLocalProfile\Documents\WindowsPowerShell\Modules") -eq $false){ #If the old 5.1 Modules are missing (without OneDrive redirection)
        if($(Test-Path "$env:SystemDrive\Users\$userAccountLocalProfile\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules") -eq $false){ #And if the old 5.1 Modules are present (with OneDrive redirection)
            New-Item -Path "$env:SystemDrive\Users\$adminAccountLocalProfile\Documents\WindowsPowerShell" -Name "Modules" -ItemType Directory #Just create a new empty Modules folder
            }
        else{
            New-Item -Name Modules -Path "$env:SystemDrive\Users\$adminAccountLocalProfile\Documents\WindowsPowerShell\" -Value "$env:SystemDrive\Users\$userAccountLocalProfile\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules" -ItemType SymbolicLink #SymbolicLink to the redirected Modules folder
            }
        }
    else{
        New-Item -Name Modules -Path "$env:SystemDrive\Users\$adminAccountLocalProfile\Documents\WindowsPowerShell\" -Value "$env:SystemDrive\Users\$userAccountLocalProfile\Documents\WindowsPowerShell\Modules" -ItemType SymbolicLink #SymbolicLink to the *un*redirected Modules folder
        }
    }


new-shortcut -name "T1 VS Code" -path "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop" -target "$env:programfiles\Microsoft VS Code\Code.exe" -runasUser "azuread\$adminAccountLocalProfile@anthesisgroup.com"


<#From within VS Code (once it's running as T1)
code --install-extension AwesomeAutomationTeam.azureautomation
code --install-extension ms-azuretools.vscode-azurefunctions --force
code --install-extension ms-azuretools.vscode-azureresourcegroups
code --install-extension ms-vscode.azure-account
code --install-extension ms-vscode.powershell

Automation Extension default settings stored separately
#>

