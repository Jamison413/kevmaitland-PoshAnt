﻿#anthesisgroup-com.mail.protection.outlook.com
Import-Module _PS_Library_MSOL.psm1
connect-ToExo

$mbs = Get-Mailbox | ? {$_.RecipientTypeDetails -eq "UserMailbox"}
$mbs | %{
    $mbInfoObj = [psobject]::new()
    $mbInfoObj | Add-Member -MemberType NoteProperty -Name Mailbox -Value $_
    $mbInfoObj | Add-Member -MemberType NoteProperty -Name Clutter -Value $(Get-Clutter -Identity $_.Id)
    [array]$mbInfo += $mbInfoObj
    }


for ($i=0;$i -lt $mbInfo.Count; $i++){
    if($mbInfo[$i].Clutter.IsEnabled -eq $true){[array]$clutterOn += $mbInfo[$i]}
        else{[array]$clutterOff += $mbInfo[$i]}
    }

$mbInfo.Count
$mbs.Count
$mbInfo | ?{$_.Clutter -eq $null} | % {$_.Clutter = $(Get-Clutter -Identity $_.Mailbox.Id)}

$mbInfo.Mailbox -eq "Kevin.Maitland"

Export-Csv -Path $env:USERPROFILE\desktop\ClutterOff.csv
Export-Csv -InputObject $clutterOn -Path $env:USERPROFILE\desktop\ClutterOn.csv

$clutterOff | % {
    $dummy = $_
    Select $dummy.Mailbox, $dummy.Clutter
    }

$mbInfo | % {
    [array]$allTheThings += New-Object psobject -Property $([ordered]@{"MailBox"=$_.Mailbox;"Clutter"=$_.Clutter.IsEnabled})
    }

$allTheThings | Export-Csv -Path $env:USERPROFILE\desktop\Clutter\ClutterStatus_$((get-date -Format "yyyy_MM_dd")).csv -NoTypeInformation -Encoding UTF8