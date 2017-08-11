Import-Module .\_PS_Library_MSOL.psm1
connect-ToExo


$members = @("Alan Spray","Andrew Noone", "Carl van Tonder", "Craig Simmons", "Erika Bata", "Helen Kean", "Helen Tyrrell", "Ian Forrester", "Laura Thompson", "Lorna Kelly", "Lucy Boreham", "Lucy Richardson", "Maggie Weglinski", "Mary Short", "Rosanna Collorafi", "Sophie Sapienza", "Stuart McLachlan", "Tecla Castella", "Tobias Parker")
$memberOf = @("")
$name = "Recruitment Team"
$hideFromGal = $false
$blockExternalMail = $true
New-DistributionGroup -Name $name -Type Security -Members $members -PrimarySmtpAddress $($name.Replace(" ","")+"@anthesisgroup.com")
Set-DistributionGroup $name -HiddenFromAddressListsEnabled $hideFromGal -RequireSenderAuthenticationEnabled $blockExternalMail


