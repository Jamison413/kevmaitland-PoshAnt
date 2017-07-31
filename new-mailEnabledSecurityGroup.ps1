$credential = get-credential -Credential kevin.maitland@anthesisgroup.com
Import-Module MSOnline
Connect-MsolService -Credential $credential

$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
Import-PSSession $ExchangeSession

$members = @("Graeme.Hadley","Polly.Stebbings","Yvonne.Ngo","Hannah.Dick","Pearl.Nemeth","Harriet.Bell","Rosie.Sibley","Chloe.McCloskey","lucy.welch","Georgie.Edwards","Beth.Simpson","Ellen.Upton","Sophie.Taylor")
$memberOf = @()
$name = "STEP Team"
$hideFromGal = $false
$blockExternalMail = $true
New-DistributionGroup -Name $name -Type Security -Members $members -PrimarySmtpAddress $($name.Replace(" ","")+"@anthesisgroup.com")
Set-DistributionGroup $name -HiddenFromAddressListsEnabled $hideFromGal -RequireSenderAuthenticationEnabled $blockExternalMail


