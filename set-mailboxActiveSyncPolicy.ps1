param(
    # Specifies which  $mailbox.CustomAttribute1 we are applying the policy to
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    #[ValidateSet()]
    [string]$businessUnit

    )

$logFileLocation = "C:\ScriptLogs\"
if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)){
    $fullLogPathAndName = $logFileLocation+"set-mailboxActiveSyncPolicy.ps1_$businessUnit`_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = $logFileLocation+"set-mailboxActiveSyncPolicy.ps1_$businessUnit`_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
else{
    $fullLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_$businessUnit`_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_$businessUnit`_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
if($PSCommandPath){
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_$businessUnit`_Transcript_$(Get-Date -Format "yyMMdd").log"
    Start-Transcript $transcriptLogName -Append
    }

Import-Module _PS_Library_MSOL
Import-Module _PS_Library_GeneralFunctionality

$groupAdmin = "groupbot@anthesisgroup.com"
#convertTo-localisedSecureString ""
$groupAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Downloads\GroupBot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $groupAdmin, $groupAdminPass

connect-ToMsol -credential $adminCreds

$groupAdmin = "groupbot@anthesisgroup.com"
$groupAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Downloads\GroupBot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $groupAdmin, $groupAdminPass
connect-ToExo -credential $adminCreds


$mailboxesToProcess = Get-Mailbox | ? {$_.CustomAttribute1 -eq $businessUnit}
$mailboxesToProcess | Set-CASMailbox -ActiveSyncMailboxPolicy "Interim security policy"
