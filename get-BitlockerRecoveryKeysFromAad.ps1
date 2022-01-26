
$logFileLocation = "C:\ScriptLogs\"
$scriptName = "get-bitlockerRecoveryKeysFromAad"
$fullLogPathAndName = $logFileLocation+$scriptName+".ps1_FullLog_$(Get-Date -Format "yyMMdd").log"
$errorLogPathAndName = $logFileLocation+$scriptName+".ps1_ErrorLog_$(Get-Date -Format "yyMMdd").log"
if($PSCommandPath){
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_$whatToSync`_Transcript_$(Get-Date -Format "yyMMdd").log"
    Start-Transcript $transcriptLogName -Append
    }

Import-Module _PS_Library_GeneralFunctionality

$aadUser = "groupbot@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
$aadUserPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Downloads\GroupBot.txt) 
$aadCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $aadUser, $aadUserPass

    #$userDevices = Get-AzureADUser -SearchString $SearchString | Get-AzureADUserRegisteredDevice -All:$true
    $allDevices = Get-AzureADDevice -All:$true

    $bitLockerKeys = @()

    foreach ($device in $allDevices) {
        $url = "https://main.iam.ad.ext.azure.com/api/Device/$($device.objectId)"
        $deviceRecord = Invoke-RestMethod -Uri $url -Headers $header -Method Get
        if ($deviceRecord.bitlockerKey.count -ge 1) {
            $bitLockerKeys += [PSCustomObject]@{
                Device      = $deviceRecord.displayName
                DriveType   = $deviceRecord.bitLockerKey.driveType
                KeyId       = $deviceRecord.bitLockerKey.keyIdentifier
                RecoveryKey = $deviceRecord.bitLockerKey.recoveryKey
            }
        }
    }

$bitlockerRecoveryKeys = Get-AzureADBitLockerKeysForAllDevices -aadCreds $aadCreds -Verbose
$bitlockerRecoveryKeys | Export-Csv -Path "C:\users\kevinm\Desktop\AuditLogs\BitlockerKeys_$(Get-Date -Format "yyMMdd").csv" -NoTypeInformation
