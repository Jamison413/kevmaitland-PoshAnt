$fileName = "Anthesis_SplashTopSupportTool_Windows.exe"


if(Test-Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop"){
    if(Test-Path -Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\$($fileName)"){
        Write-Host "Path [$("$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\$($fileName)")] exists!"
        }
    else{
        Throw "Path [$("$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\$($fileName)")] not found - [$($thisApp)] is not installed"
        }    
    }
else{
    if(Test-Path -Path "$env:USERPROFILE\Desktop\$($fileName)"){
        Write-Host "Path [$("$env:USERPROFILE\Desktop\$($fileName)")] exists!"
        }
    else{
        Throw "Path [$("$env:USERPROFILE\Desktop\$($fileName)")] not found - [$($fileName)] is not installed"
        }    
    }
