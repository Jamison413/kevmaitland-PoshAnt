connect-ToMsol -credential $msolCredentials
connect-ToExo -credential $msolCredentials

$GroupsTohide = Get-UnifiedGroup | where {($_.DisplayName -notmatch "External - " -and $_.HiddenFromAddresslistsEnabled -eq $False)} 
$GroupsTohide | Set-UnifiedGroup -HiddenFromAddressListsEnabled $True