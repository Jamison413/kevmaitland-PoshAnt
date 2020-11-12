$teamBotDetails = get-graphAppClientCredentials -appName TeamsBot
$tokenResponseTeams = get-graphTokenResponse -aadAppCreds $teamBotDetails

$intuneBotDetails = get-graphAppClientCredentials -appName IntuneBot
$tokenResponseIntune = get-graphTokenResponse -aadAppCreds $intuneBotDetails

$exoCreds = set-MsolCredentials -username t0-kevin.maitland@anthesisgroup.com
connect-ToExo -credential $exoCreds

$members = convertTo-arrayOfEmailAddresses "
Gregor.Pecnik@anthesisgroup.com
conflictminerals@anthesisgroup.com
wdpec@anthesisgroup.com
ProductCompliance@anthesisgroup.com
micron.pec@anthesisgroup.com
cupid@anthesisgroup.com
L3.Pec@anthesisgroup.com
crownprince@anthesisgroup.com
summit@anthesisgroup.com
WMSI@anthesisgroup.com
VarexConflictMinerals@anthesisgroup.com
ruiz@anthesisgroup.com
Varex.PEC@anthesisgroup.com
pmi@anthesisgroup.com
Eaton.Pec@anthesisgroup.com
cayre@anthesisgroup.com
VarianConflictMinerals@anthesisgroup.com
cray@anthesisgroup.com
almar@anthesisgroup.com
kcc.pec@anthesisgroup.com
AvayaConflictMinerals@anthesisgroup.com
burkert.pec@anthesisgroup.com
Yara.pec@anthesisgroup.com
Microsoft.ECM@anthesisgroup.com
dexcom.pec@anthesisgroup.com
akm.pec@anthesisgroup.com
akmstaging.pec@anthesisgroup.com
wackerneuson.pec@anthesisgroup.com
wackerneusonstaging.pec@anthesisgroup.com
"
new-mailEnabledSecurityGroup -dgDisplayName "Mailbox Access - ACS Mailer App" -description "Grants access to members' mailbxoes for the ACS Mailer App" -hideFromGal $true -membersUpns $members -blockExternalMail $true
New-ApplicationAccessPolicy -AppId b32a2aee-b0bd-46c8-83c8-88b726b180f7 -PolicyScopeGroupId Mailbox_Access_-_ACS_Mailer_App@anthesisgroup.com -AccessRight RestrictAccess -Description "Restrict [ACS Mailer] App to members of distribution group [Mailbox Access - ACS Mailer App]." 
Test-ApplicationAccessPolicy -Identity kevin.maitland@anthesisgroup.com -AppId b32a2aee-b0bd-46c8-83c8-88b726b180f7
Test-ApplicationAccessPolicy -Identity wackerneusonstaging.pec@anthesisgroup.com -AppId b32a2aee-b0bd-46c8-83c8-88b726b180f7

$smbx = Get-Mailbox -Filter {recipienttypedetails -eq "SharedMailbox"}
$results = $smbx | Get-MailboxPermission | select identity,user,accessrights  | where { ($_.User -like 'acsmailboxaccess@anthesisgroup.com')   }
$results | % {Add-DistributionGroupMember -Identity Mailbox_Access_-_ACS_Mailer_App@anthesisgroup.com -Member $_.Identity -Confirm:$false}
$results | select identity

$membersAre = get-graphUsersFromGroup -tokenResponse $tokenResponseTeams -groupUpn Mailbox_Access_-_ACS_Mailer_App@anthesisgroup.com -memberType Members
$membersAre.mail | sort