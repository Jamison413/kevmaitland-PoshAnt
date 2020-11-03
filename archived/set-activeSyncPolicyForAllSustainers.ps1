Import-Module _PS_Library_MSOL
$gaCreds = set-MsolCredentials 
connect-ToExo -credential $gaCreds

Get-Mailbox -ResultSize unlimited -filter {CustomAttribute1 -eq "Sustain"} | Set-CASMailbox -ActiveSyncMailboxPolicy "Sustain"
