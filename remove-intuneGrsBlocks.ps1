Get-ChildItem HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps\ | ? {@("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps\Reporting","HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps\00000000-0000-0000-0000-000000000000" -notcontains $_.Name )} | % { 
    Remove-Item $_.Name.Replace("HKEY_LOCAL_MACHINE","HKLM:") -Confirm:$false -Force 
    } 
 