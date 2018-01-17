Import-Module _PS_Library_MSOL.psm1
Import-Module _PS_Library_GeneralFunctionality
connect-ToExo

$members = convertTo-arrayOfStrings -blockOfText "Stuart Mclachlan
Ian Forrester
Enda.Colfer@anthesisgroup.com
Alan matthews
Chris Keller
Malcolm Paul
Jason Urry
Paul Crewe
"

$members = convertTo-arrayOfEmailAddresses -blockOfText "Kev Maitland
 
kevin.maitland@anthesisgroup.com
 
 
G
 
Georgie Edwards
 
Georgie.Edwards@anthesisgroup.com
 
 
S
 
Sion Fenwick
 
Sion.Fenwick@anthesisgroup.com
 
 
P
 
Pete Best
 
Pete.Best@anthesisgroup.com
 
 
J
 
James Walker
 
james.walker@anthesisgroup.com
 
 
M
 
Margaret Davis
 
Margaret.Davis@anthesisgroup.com
 
 
L
 
Louise Trimby
 
Louise.Trimby@anthesisgroup.com
 
 
M
 
Mark Griffin
 
Mark.Griffin@anthesisgroup.com
 
 
S
 
Stuart Gray"

$members = @("Carl.van.Tonder","istvan.denes","Craig.Simmons","Ben.Diallo","Stephen.Ellis","Greg.Stencel")
$memberOf = @()
$name = "Products & Technology Team2"
$hideFromGal = $false
$blockExternalMail = $true


function new-mailEnabledDistributionGroup($displayName, $members, $memberOf, $hideFromGal, $blockExternalMail){
    New-DistributionGroup -Name $name -Type Security -Members $members -PrimarySmtpAddress $($name.Replace(" ","")+"@anthesisgroup.com")
    Set-DistributionGroup $name -HiddenFromAddressListsEnabled $hideFromGal -RequireSenderAuthenticationEnabled $blockExternalMail
    }


    new-mailEnabledDistributionGroup -displayName $name -members $members -memberOf $memberOf -hideFromGal $hideFromGal -blockExternalMail $blockExternalMail

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