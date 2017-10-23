Import-Module _PS_Library_MSOL.psm1
connect-ToExo

$members = @("Mary.Short@anthesisgroup.com","Ian.Forrester@anthesisgroup.com","debbie.hitchen@anthesisgroup.com","Fiona.Place@anthesisgroup.com","Andrew.Noone@anthesisgroup.com","Jono.Adams@anthesisgroup.com","Helen.Kean@anthesisgroup.com","brad.blundell@anthesisgroup.com")
$memberOf = @()
$name = "Sustainability Senior Management Team"
$hideFromGal = $false
$blockExternalMail = $true


function new-mailEnabledDistributionGroup($displayName, $members, $memberOf, $hideFromGal, $blockExternalMail){
    New-DistributionGroup -Name $name -Type Security -Members $members -PrimarySmtpAddress $($name.Replace(" ","")+"@anthesisgroup.com")
    Set-DistributionGroup $name -HiddenFromAddressListsEnabled $hideFromGal -RequireSenderAuthenticationEnabled $blockExternalMail
    }


new-mailEnabledDistributionGroup -displayName $name -members $members -memberOf $memberOf -hideFromGal $hideFromGal -blockExternalMail $blockExternalMail

Add-MailboxPermission -AccessRights fullaccess -Identity nigel.arnott -User mary.short -AutoMapping $true

$members | %{Add-DistributionGroupMember -Identity "iONA Capital Team" -Member $_}

$members | % {Add-DistributionGroupMember -Identity "Guineapigs Sustain Spam" -Member $_}