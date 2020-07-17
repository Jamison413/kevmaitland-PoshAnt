#Dump S4B / Teams config
connect-ToS4b $365creds
$s4bConfigLog = "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\S4B_Config_PostInstall_Working.log"
$s4bCmdlets = Get-Command | ? {$_.Name -match "Get-CS" -and $_.Name -ne "Get-CsOnlinePowerShellAccessToken"}

"Skype For Business Config Log $(Get-Date -Format s)" | Out-File $s4bConfigLog

$s4bCmdlets | % {
    "`t$($_.Name)" | Out-File $s4bConfigLog -Append
    Invoke-Expression $_.Name | Out-File $s4bConfigLog -Append
    }


