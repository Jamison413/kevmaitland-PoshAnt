Import-Module _PS_Library_MSOL.psm1
connect-ToExo

$members = @("Polly Stebbings","Georgie Edwards","Graeme Hadley")
$memberOf = @("")
$name = "PULSE Team"
$hideFromGal = $false
$blockExternalMail = $true


function new-mailEnabledDistributionGroup($displayName, $members, $memberOf, $hideFromGal, $blockExternalMail){
    New-DistributionGroup -Name $name -Type Security -Members $members -PrimarySmtpAddress $($name.Replace(" ","")+"@anthesisgroup.com")
    Set-DistributionGroup $name -HiddenFromAddressListsEnabled $hideFromGal -RequireSenderAuthenticationEnabled $blockExternalMail
    }

Add-MailboxPermission -AccessRights fullaccess -Identity nigel.arnott -User mary.short -AutoMapping $true

$members | %{Add-DistributionGroupMember -Identity "iONA Capital Team" -Member $_}

