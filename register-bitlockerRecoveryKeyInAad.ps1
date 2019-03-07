$logFileLocation = "$env:USERPROFILE\"
$transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_Transcript_$(Get-Date -Format "yyMMdd").log"
if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)){
    $fullLogPathAndName = $logFileLocation+"register-bitlockerRecoveryKeyInAad_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = $logFileLocation+"register-bitlockerRecoveryKeyInAad_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
else{
    $fullLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
Start-Transcript $transcriptLogName -Append

$bitlockerVolume = Get-BitLockerVolume | ? {$_.ProtectionStatus -eq "on"} 
$bitlockerVolume | % {
    rv latestRecoveryKey -ErrorAction SilentlyContinue
    $_.KeyProtector | %{if($_.KeyProtectorType -eq "RecoveryPassword"){$latestRecoveryKey = $_}}
    if($latestRecoveryKey){
        BackupToAAD-BitLockerKeyProtector -MountPoint $_.MountPoint -KeyProtectorId $latestRecoveryKey.KeyProtectorId
        }
    }
Stop-Transcript