if((Test-Path -Path $env:APPDATA\Microsoft\Templates\Normal.dotm.old) -eq $true){
    Rename-Item -Path $env:APPDATA\Microsoft\Templates\Normal.dotm -NewName Normal.dotm.removed -Force
    if((Test-Path -Path $env:APPDATA\Microsoft\Templates\Normal.dotm.removed) -eq $true){
        Rename-Item -Path  $env:APPDATA\Microsoft\Templates\Normal.dotm.old  -NewName Normal.dotm -Force
        }
    }

if((Test-Path -Path $env:APPDATA\Microsoft\Templates\NormalEmail.dotm.old) -eq $true){
    Rename-Item -Path $env:APPDATA\Microsoft\Templates\NormalEmail.dotm -NewName NormalEmail.dotm.removed -Force
    if((Test-Path -Path $env:APPDATA\Microsoft\Templates\NormalEmail.dotm.removed) -eq $true){
        Rename-Item -Path  $env:APPDATA\Microsoft\Templates\NormalEmail.dotm.old  -NewName NormalEmail.dotm -Force
        }
    }

if((Test-Path -Path $env:APPDATA\Microsoft\Templates\Blank.potx.old) -eq $true){
    Rename-Item -Path $env:APPDATA\Microsoft\Templates\Blank.potx -NewName Blank.potx.removed -Force
    if((Test-Path -Path $env:APPDATA\Microsoft\Templates\Blank.potx.removed) -eq $true){
        Rename-Item -Path  $env:APPDATA\Microsoft\Templates\Blank.potx.old  -NewName Blank.potx -Force
        }
    }

if((Test-Path "$env:APPDATA\Microsoft\Templates\Document Themes\Anthesis.thmx") -eq $true){
    Remove-Item -Path "$env:APPDATA\Microsoft\Templates\Document Themes\Anthesis.thmx" -Force
    }
if((Test-Path "$env:APPDATA\Microsoft\Templates\Document Themes\Theme Colors\Anthesis_Official.xml") -eq $true){
    Remove-Item -Path "$env:APPDATA\Microsoft\Templates\Document Themes\Theme Colors\Anthesis_Official.xml" -Force
    }

