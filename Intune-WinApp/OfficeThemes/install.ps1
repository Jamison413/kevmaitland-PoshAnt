if((Test-Path -Path $env:APPDATA\Microsoft\Templates\Normal.dotm.old) -eq $false){
    Rename-Item -Path $env:APPDATA\Microsoft\Templates\Normal.dotm -NewName Normal.dotm.old -Force
    if((Test-Path -Path $env:APPDATA\Microsoft\Templates\Normal.dotm.old) -eq $true -or (Test-Path -Path $env:APPDATA\Microsoft\Templates\Normal.dotm) -eq $false){
        Copy-Item -Path .\Normal.dotm  -Destination $env:APPDATA\Microsoft\Templates\ -Force
        }
    }

if((Test-Path -Path $env:APPDATA\Microsoft\Templates\NormalEmail.dotm.old) -eq $false){
    Rename-Item -Path $env:APPDATA\Microsoft\Templates\NormalEmail.dotm -NewName NormalEmail.dotm.old -Force
    if((Test-Path -Path $env:APPDATA\Microsoft\Templates\NormalEmail.dotm.old) -eq $true -or (Test-Path -Path $env:APPDATA\Microsoft\Templates\NormalEmail.dotm) -eq $false){
        Copy-Item -Path .\NormalEmail.dotm  -Destination $env:APPDATA\Microsoft\Templates\ -Force
        }
    }

if((Test-Path -Path $env:APPDATA\Microsoft\Templates\Blank.potx.old) -eq $false){
    Rename-Item -Path $env:APPDATA\Microsoft\Templates\Blank.potx -NewName Blank.potx.old -Force
    if((Test-Path -Path $env:APPDATA\Microsoft\Templates\Blank.potx.old) -eq $true -or (Test-Path -Path $env:APPDATA\Microsoft\Templates\Blank.potx) -eq $false){
        Copy-Item -Path .\Blank.potx  -Destination $env:APPDATA\Microsoft\Templates\ -Force
        }
    }

Copy-Item -Path ".\Document Themes\" -Destination "$env:APPDATA\Microsoft\Templates" -Force -Recurse
