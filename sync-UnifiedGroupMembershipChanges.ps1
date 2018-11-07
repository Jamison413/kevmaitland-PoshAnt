$logFileLocation = "C:\ScriptLogs\"
if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)){
    $fullLogPathAndName = $logFileLocation+"sync-UnifiedGroupMembershipChanges_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = $logFileLocation+"sync-UnifiedGroupMembershipChanges_ErrorLog_$(Get-Date -Format "yyMMdd").log"
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
Import-Module _PS_Library_Groups
Import-Module _PS_Library_GeneralFunctionality

$groupAdmin = "groupbot@anthesisgroup.com"
#convertTo-localisedSecureString ""
$groupAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\GroupBot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $groupAdmin, $groupAdminPass

connect-toAAD -credential $adminCreds

$groupAdmin = "groupbot@anthesisgroup.com"
$groupAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\GroupBot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $groupAdmin, $groupAdminPass
connect-ToExo -credential $adminCreds

$all365Groups = Get-UnifiedGroup
$toExclude = @("Sym - Supply Chain","Apparel Team (All)","All North America","Business Development Team (GBR)","Pre Sales Team (All)","Teams Testing Team", "Finance Team (North America)","Finance Team (North America)")
$365GroupsToProcess = $all365Groups | ? {$toExclude -notcontains $($_.DisplayName) -and $_.DisplayName -notmatch "Confidential"}


$365GroupsToProcess | % {
    $365Group = $_
    sync-365GroupMembersToMirroredSecurityGroup -unifiedGroupObject $365Group -reallyDoIt $true -dontSendEmailReport $false -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
    sync-managersTo365GroupOwners -unifiedGroupObject $365Group -reallyDoIt $true -dontSendEmailReport $false -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
    }



Stop-Transcript

<#
$365Group = Get-UnifiedGroup "Software Team (PHI)"
Remove-Module _PS_Library_Groups
Import-Module _PS_Library_Groups
#>