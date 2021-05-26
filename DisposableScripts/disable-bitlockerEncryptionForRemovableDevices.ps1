if(Test-Path HKLM:\SYSTEM\CurrentControlSet\Policies\Microsoft\FVE){
    try{
        $RDVDenyWriteAccess = Get-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Policies\Microsoft\FVE -Name "RDVDenyWriteAccess" -ErrorAction Stop
        if($RDVDenyWriteAccess.RDVDenyWriteAccess -eq 1){Set-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Policies\Microsoft\FVE -Name "RDVDenyWriteAccess" -Value 0}
        }
    catch{
        New-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Policies\Microsoft\FVE -Name "RDVDenyWriteAccess" -Value 0
        }
    }
else{
    New-Item -Path HKLM:\SYSTEM\CurrentControlSet\Policies\Microsoft -Name FVE
    New-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Policies\Microsoft\FVE -Name "RDVDenyWriteAccess" -Value 0
    }