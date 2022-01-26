$logFileLocation = "C:\ScriptLogs\"
if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)){
    $fullLogPathAndName = $logFileLocation+"alert-passwordExpiry365_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = $logFileLocation+"alert-passwordExpiry365_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
else{
    $fullLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
if($PSCommandPath){
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_Transcript_$(Get-Date -Format "yyMMdd").log"
    Start-Transcript $transcriptLogName -Append
    }

Import-Module _PS_Library_MSOL
Import-Module _PS_Library_GeneralFunctionality

$groupAdmin = "groupbot@anthesisgroup.com"
#convertTo-localisedSecureString ""
$groupAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Downloads\GroupBot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $groupAdmin, $groupAdminPass
connect-ToMsol -credential $adminCreds
connect-ToExo -credential $adminCreds

$allUsers = @()
$AllUsers = Get-MsolUser -All -EnabledFilter EnabledOnly | select ObjectID, UserPrincipalName, FirstName, LastName, StrongAuthenticationRequirements, StsRefreshTokensValidFrom, StrongPasswordRequired, LastPasswordChangeTimestamp | Where-Object {($_.UserPrincipalName -notlike "*#EXT#*")}

$UserInboxRules = @()
$UserDelegates = @()

foreach ($User in $allUsers)
{
    Write-Host "Checking inbox rules and delegates for user: " $User.UserPrincipalName;
    $UserInboxRules += Get-InboxRule -Mailbox $User.UserPrincipalname | Select Name, Description, Enabled, Priority, ForwardTo, ForwardAsAttachmentTo, RedirectTo, DeleteMessage | Where-Object {($_.ForwardTo -ne $null) -or ($_.ForwardAsAttachmentTo -ne $null) -or ($_.RedirectsTo -ne $null)}
    $UserDelegates += Get-MailboxPermission -Identity $User.UserPrincipalName | Where-Object {($_.IsInherited -ne "True") -and ($_.User -notlike "*SELF*")}
}

$SMTPForwarding = Get-Mailbox -ResultSize Unlimited | select DisplayName,ForwardingAddress,ForwardingSMTPAddress,DeliverToMailboxandForward | where {$_.ForwardingSMTPAddress -ne $null}

$UserInboxRules | Export-Csv C:\Users\kevinm\Desktop\AuditLogs\MailForwardingRulesToExternalDomains_$(Get-Date -Format "yyMMdd").csv
$UserDelegates | Export-Csv C:\Users\kevinm\Desktop\AuditLogs\MailboxDelegatePermissions_$(Get-Date -Format "yyMMdd").csv
$SMTPForwarding | Export-Csv C:\Users\kevinm\Desktop\AuditLogs\Mailboxsmtpforwarding_$(Get-Date -Format "yyMMdd").csv


$SMTPForwarding | ? {$_.ForwardingSmtpAddress -notmatch "anthesisgroup.com"}
$UserDelegates[0]