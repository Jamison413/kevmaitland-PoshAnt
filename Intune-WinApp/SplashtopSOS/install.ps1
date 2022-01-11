$file = Get-Item .\Anthesis_SplashTopSupportTool_Windows.exe 
if(Test-Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop"){
    Copy-Item -Path $file.PSPath -Destination "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop" -Force
    }
else{
    Copy-Item -Path $file.PSPath -Destination "$env:USERPROFILE\Desktop"  -Force
    }