$file = Get-Item .\Anthesis_SplashTopSupportTool_Windows.exe

if(Test-Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop"){
    Remove-Item -Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\$($file.Name)" -Force
    }
else{
    Remove-Item -Path "$env:USERPROFILE\Desktop\$($file.Name)" -Force
    }