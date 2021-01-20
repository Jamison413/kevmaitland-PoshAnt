$file = Get-Item .\Splashtop_Streamer_Windows_DEPLOY_INSTALLER_v3.4.2.2_HJ5TPTZAATPR.msi
$MSIArguments = @(
    "/i"
    ('"{0}"' -f $file.fullname)
    "/qn"
    "/norestart"
    )
Start-Process "msiexec.exe" -ArgumentList $MSIArguments -Wait -NoNewWindow 