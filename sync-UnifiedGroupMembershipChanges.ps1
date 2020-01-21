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

connect-ToMsol -credential $adminCreds
connect-toAAD -credential $adminCreds
connect-ToExo -credential $adminCreds

$all365Groups = Get-UnifiedGroup
$toExclude = @("Sym - Supply Chain","Apparel Team (All)","All North America","Business Development Team (GBR)","Pre Sales Team (All)","Teams Testing Team", "Finance Team (North America)","Finance Team (North America)")
$365GroupsToProcess = $all365Groups | ? {$toExclude -notcontains $($_.DisplayName) -and $_.DisplayName -notmatch "Confidential"}

$adminEmailAddresses = get-groupAdminRoleEmailAddresses

$startProcessing = $false
$365GroupsToProcess | % {
    $365Group = $_
    $365Group.DisplayName
    if($365Group.DisplayName -match "External - Assistència tècnica per al programa d'Acords voluntaris per la reducció d'emissions de GEH"){$startProcessing = $true}
    if($startProcessing -eq $false){return}
    try{sync-groupMemberships -UnifiedGroup $365Group -syncWhat Members -sourceGroup $365Group.CustomAttribute6 -adminEmailAddresses $adminEmailAddresses -enumerateSubgroups $true -Verbose }
    catch{$_;break}
    try{sync-groupMemberships -UnifiedGroup $365Group -syncWhat Owners -sourceGroup AAD -adminEmailAddresses $adminEmailAddresses -enumerateSubgroups $true -Verbose}
    catch{$_;break}
    }



Stop-Transcript
