Import-Module _PS_Library_MSOL.psm1
Import-Module _PS_Library_GeneralFunctionality
connect-ToExo

$members = convertTo-arrayOfStrings -blockOfText ""
$members = convertTo-arrayOfEmailAddresses -blockOfText ""
$members = @("Jessica.onyshko@anthesisgroup.com","Paul.Ashford@anthesisgroup.com","Drew.ONeil@Target.com","Don.Asleson@Target.com")
$members = @("kevin.maitland")

$memberOf = @("All Europe")
$owners = @("kevin.maitland")
$name = "All (COL)"
$hideFromGal = $false
$blockExternalMail = $true
$public365Site = $true
$autoSubscribe = $true

function new-mailEnabledDistributionGroup($displayName, $members, $memberOf, $hideFromGal, $blockExternalMail, $owners){
    New-DistributionGroup -Name $displayName -Type Security -Members $members -PrimarySmtpAddress $($name.Replace(" ","").Replace("(","").Replace(")","")+"@anthesisgroup.com")
    Set-DistributionGroup $displayName -HiddenFromAddressListsEnabled $hideFromGal -RequireSenderAuthenticationEnabled $blockExternalMail -ManagedBy $owners
    }
function new-365Group($displayName, $members, $memberOf, $hideFromGal, $blockExternalMail, $owners, $isPublic, $autoSubscribe){
    $shortName = $Name.Replace(" (All)","")
    new-mailEnabledDistributionGroup -displayName $("Managers - $shortName") -members $owners -memberOf "Managers - All" -hideFromGal $true -blockExternalMail $true -owners "IT Team"
    $newManagerGroup = Get-DistributionGroup -Identity $("Managers - $shortName")
    if($isPublic){$accessType = "Public"}else{$accessType = "Private"}
    $alias = $name.replace(" ","").Replace("(","_").Replace(")","")
    New-UnifiedGroup -RequireSenderAuthenticationEnabled $blockExternalMail -AutoSubscribeNewMembers:$autoSubscribe -AlwaysSubscribeMembersToCalendarEvents:$autoSubscribe -DisplayName $Name -Members $members -AccessType $accessType -Name $alias -Alias $alias -Owner $($owners -join ",")
    }



new-mailEnabledDistributionGroup -displayName $name -members $members -memberOf $memberOf -hideFromGal $hideFromGal -blockExternalMail $blockExternalMail -owners $owners
new-365Group -displayName $name -members $members -memberOf $memberOf -hideFromGal $hideFromGal -blockExternalMail $blockExternalMail -owners $owners -isPublic $public365Site -autoSubscribe $autoSubscribe

#Add-MailboxPermission -AccessRights fullaccess -Identity nigel.arnott -User mary.short -AutoMapping $true
#$members | %{Add-DistributionGroupMember -Identity "iONA Capital Team" -Member $_}
#$members | % {Add-DistributionGroupMember -Identity "Clients - Chinook UM Team" -Member $_}
Set-InboxRule -Mailbox NA-CareersAutoreply -
$oldRules = Get-InboxRule -Mailbox NA-CareersAutoreply
New-InboxRule -Mailbox NA-CareersAutoreply -Name "Forward to NA Careers DG" -ForwardTo "na-careers@anthesisgroup.com"
$newrules = Get-InboxRule -Mailbox NA-CareersAutoreply

$newrules

$rules[0] | fl

$guineapigs = Get-DistributionGroupMember "Guineapigs Spam Control Group"
$guineapigs += Get-DistributionGroupMember "Sustain"



$guinepigs = @("Rosanna Collorafi","Emma Armstrong","Curtis Harnanan","Matt Wood","Mary Short","Chris Jennings","Georgie Edwards","Sion Fenwick","Stuart Miller","Pete Best","James Carberry","Amy MacGrain","Tobias Parker","Nigel Arnott","Henrietta Bird","Matthew Gitsham","James Walker","Josep Porta","Lorna Kelly","Margaret Davis","Duncan Faulkes","Wai Cheung","Stuart Gray","Laurie Eldridge","Huw Blackwell","Rebecca Hughes")
$guinepigs | % {
    Add-DistributionGroupMember "Guineapigs Spam Experimental Group" -Member $($_.Replace(" ",".")+"@anthesisgroup.com")
    }


    $dummy = Get-DistributionGroup -Identity "Managers - All"


@("ali.mahdavi","katie.swain","simon.white","laura.sponti","sion.fenwick","ben.buffery","laura.pugh","tilly.shaw","catherine.green") | % {
    $user = $_
    $u = Get-User -Identity $user@anthesisgroup.com
    Get-DistributionGroup -Filter "Members -eq '$($u.DistinguishedName)'" | % {
        Remove-DistributionGroupMember -Identity $_.Id -Member $user@anthesisgroup.com -Confirm:$false
        }
    Set-Mailbox $user -HiddenFromAddressListsEnabled $true -InactiveMailbox
    }

