#anthesisgroup-com.mail.protection.outlook.com
Import-Module _PS_Library_MSOL.psm1
connect-ToExo

$mbs = Get-Mailbox | ? {$_.RecipientTypeDetails -eq "UserMailbox"}
$mbs | %{
    $mbInfoObj = [psobject]::new()
    $mbInfoObj | Add-Member -MemberType NoteProperty -Name Mailbox -Value $_
    $mbInfoObj | Add-Member -MemberType NoteProperty -Name Clutter -Value $(Get-Clutter -Identity $_.Id)
    [array]$mbInfo += $mbInfoObj
    }

$mbInfo = $sminfo
for ($i=0;$i -lt $mbInfo.Count; $i++){
    if($mbInfo[$i].Clutter -eq $true){[array]$clutterOn += $mbIfo[$i]}
        else{[array]$cluetteroff += $mbIfo[$i]}
    }

