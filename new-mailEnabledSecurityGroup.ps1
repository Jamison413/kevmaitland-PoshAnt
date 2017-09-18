Import-Module _PS_Library_MSOL.psm1
connect-ToExo


$members = @("Claudia Amos","Peter Scholes","Brad Blundell", "Thomas Milne", "Hannah Dick","Michael Kirk-Smith")
$memberOf = @("")
$name = "Helios Team"
$hideFromGal = $false
$blockExternalMail = $true
New-DistributionGroup -Name $name -Type Security -Members $members -PrimarySmtpAddress $($name.Replace(" ","")+"@anthesisgroup.com")
Set-DistributionGroup $name -HiddenFromAddressListsEnabled $hideFromGal -RequireSenderAuthenticationEnabled $blockExternalMail


Add-MailboxPermission -AccessRights fullaccess -Identity nigel.arnott -User mary.short -AutoMapping $true

$members | %{Add-DistributionGroupMember -Identity "iONA Capital Team" -Member $_}

