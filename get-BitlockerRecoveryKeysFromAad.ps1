
$logFileLocation = "C:\ScriptLogs\"
$scriptName = "get-bitlockerRecoveryKeysFromAad"
$fullLogPathAndName = $logFileLocation+$scriptName+".ps1_FullLog_$(Get-Date -Format "yyMMdd").log"
$errorLogPathAndName = $logFileLocation+$scriptName+".ps1_ErrorLog_$(Get-Date -Format "yyMMdd").log"
if($PSCommandPath){
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_$whatToSync`_Transcript_$(Get-Date -Format "yyMMdd").log"
    Start-Transcript $transcriptLogName -Append
    }

Import-Module _PS_Library_GeneralFunctionality

#$aadUser = "groupbot@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
#$aadUserPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\GroupBot.txt) 
#$aadCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $aadUser, $aadUserPass
$aadCreds = set-MsolCredentials
Connect-AzureRmAccount -Credential $aadCreds

$bitlockerRecoveryKeys = Get-AzureADBitLockerKeysForAllDevices -aadCreds $aadCreds -Verbose
$bitlockerRecoveryKeys | export-csv $env:USERPROFILE\Desktop\BitlockerKeys_$(Get-Date -Format "yyMMdd").csv -NoTypeInformation