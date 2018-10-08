$logFileLocation = "C:\ScriptLogs\"
if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)){
    $fullLogPathAndName = $logFileLocation+"set-initialTeamMembership_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = $logFileLocation+"set-initialTeamMembership_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
else{
    $fullLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
if($PSCommandPath){
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_Transcript_$(Get-Date -Format "yyMMdd").log"
    Start-Transcript $transcriptLogName -Append
    }

Import-Module _PS_Library_GeneralFunctionality
Import-Module SharePointPnPPowerShellOnline
Import-Module _PNP_Library_SPO

$groupAdmin = "groupbot@anthesisgroup.com"
#convertTo-localisedSecureString ""
$groupAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\GroupBot.txt) 
$exoCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $groupAdmin, $groupAdminPass

$sharePointAdmin = "kimblebot@anthesisgroup.com"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\KimbleBot.txt) 
$sharePointCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
#$sharePointCreds = set-MsolCredentials

connect-ToExo -credential $exoCreds
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com" -Credentials $sharePointCreds

#$membersToAdd = convertTo-arrayOfEmailAddresses "Andy Marsh <Andy.Marsh@anthesisgroup.com>; Anne O’Brien <Anne.OBrien@anthesisgroup.com>; Beth Simpson <Beth.Simpson@anthesisgroup.com>; Claudia Amos <Claudia.Amos@anthesisgroup.com>; Debbie Hitchen <debbie.hitchen@anthesisgroup.com>; Ellen Struthers <Ellen.Struthers@anthesisgroup.com>; Emma Hampsey <Emma.Hampsey@anthesisgroup.com>; Hannah Dick <Hannah.Dick@anthesisgroup.com>; Julian Parfitt <Julian.Parfitt@anthesisgroup.com>; Mark Sayers <Mark.Sayers@anthesisgroup.com>; Nick Cuomo <Nick.Cuomo@anthesisgroup.com>; Peter Scholes <Peter.Scholes@anthesisgroup.com>; Richard Peagam <richard.peagam@anthesisgroup.com>; Simone Aplin <Simone.Aplin@anthesisgroup.com>; Stephanie Egee <Stephanie.Egee@anthesisgroup.com>"
#$ownersToAdd = convertTo-arrayOfEmailAddresses " Richard Peagam <richard.peagam@anthesisgroup.com>"

$membersToAdd = Get-DistributionGroupMember "Senior Management Team (Energy Division)"
$groupStub = "Senior Management Team (GBR)"


$sg = Get-DistributionGroup $($groupStub.Replace(" ","").Replace("(","").Replace(")","")+"@anthesisgroup.com")
$mirror = Get-DistributionGroup $($groupStub.Replace(" ","").Replace("(","").Replace(")","")+"-365Mirror@anthesisgroup.com")
$managers = Get-DistributionGroup $($groupStub.Replace(" ","").Replace("(","").Replace(")","")+"-Managers@anthesisgroup.com")
$365 = Get-UnifiedGroup $($groupStub.Replace(" ","_").Replace("(","").Replace(")","")+"_365@anthesisgroup.com")

Add-UnifiedGroupLinks -Identity $365.Id -LinkType Member -Links $($membersToAdd.ExternalDirectoryObjectId)
#Add-UnifiedGroupLinks -Identity $365.Id -LinkType Owner -Links $ownersToAdd

$membersToAdd | % {
    Add-DistributionGroupMember -Identity $sg.ExternalDirectoryObjectId -Member $_.ExternalDirectoryObjectId -BypassSecurityGroupManagerCheck
    Add-DistributionGroupMember -Identity $mirror.ExternalDirectoryObjectId -Member $_.ExternalDirectoryObjectId -BypassSecurityGroupManagerCheck
    }
#$ownersToAdd | % {
#    Add-DistributionGroupMember -Identity $managers.ExternalDirectoryObjectId -Member $_ -BypassSecurityGroupManagerCheck
#    }

Stop-Transcript