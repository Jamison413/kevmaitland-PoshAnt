Get-ChildItem HKLM:\SOFTWARE\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps\ | ? {@("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Enrollments\Context","HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Enrollments\Ownership","HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Enrollments\Status","HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Enrollments\ValidNodePaths" -notcontains $_.Name )} | % { 
    Remove-Item $_.Name.Replace("HKEY_LOCAL_MACHINE","HKLM:") -Confirm:$false -Force 
    } 
 