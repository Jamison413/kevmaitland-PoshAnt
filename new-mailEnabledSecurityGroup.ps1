Import-Module _PS_Library_MSOL.psm1
connect-ToExo


$members = @("Amy.Dartington","Georgie.Edwards","Stuart.Gray","Sion.Fenwick","Tom.Willis")
$memberOf = @("")
$name = "Smart Islans Energy Team"
$hideFromGal = $false
$blockExternalMail = $false
New-DistributionGroup -Name $name -Type Security -Members $members -PrimarySmtpAddress $($name.Replace(" ","")+"@anthesisgroup.com")
Set-DistributionGroup $name -HiddenFromAddressListsEnabled $hideFromGal -RequireSenderAuthenticationEnabled $blockExternalMail


Add-MailboxPermission -AccessRights fullaccess -Identity nigel.arnott -User mary.short -AutoMapping $true