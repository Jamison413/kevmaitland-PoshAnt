Import-Module _PS_Library_MSOL.psm1
Import-Module _PS_Library_GeneralFunctionality
Import-Module _PS_Library_Groups
connect-ToExo

$displayName = "Yara - PEC"
$owner = "michael.malate@anthesisgroup.com"
$arrayOfFullAccessMembers = @("Wolfgang.Wick@anthesisgroup.com","sharleen.rivera@anthesisgroup.com","irene.oliquino@anthesisgroup.com","gerber.manalo@anthesisgroup.com","Zarel.Adame@anthesisgroup.com","michael.malate@anthesisgroup.com")
$grantSendAsToo = $true
$hideFromGal = $true


function new-sharedMailbox($displayName, $owner, $arrayOfFullAccessMembers, $hideFromGal, $grantSendAsToo){
    $exchangeAlias = $(guess-aliasFromDisplayName -displayName $displayName)
    New-Mailbox -Shared -ModeratedBy $owner -DisplayName $displayName -Name $displayName -Alias $exchangeAlias | Set-Mailbox -HiddenFromAddressListsEnabled $hideFromGal -RequireSenderAuthenticationEnabled $false
    $arrayOfFullAccessMembers  | %{
        Add-MailboxPermission -AccessRights "FullAccess" -User $_ -AutoMapping $true -Identity $exchangeAlias
        if ($grantSendAsToo){Add-RecipientPermission -Identity $exchangeAlias -Trustee $_ -AccessRights SendAs -Confirm:$false}
        }
    }

new-sharedMailbox -displayName $displayName -arrayOfFullAccessMembers $arrayOfFullAccessMembers -hideFromGal $hideFromGal -owner $owner -grantSendAsToo $grantSendAsToo

#Add-MailboxPermission -AccessRights fullaccess -Identity nigel.arnott -User mary.short -AutoMapping $true
#$members | %{Add-DistributionGroupMember -Identity "iONA Capital Team" -Member $_}
#$members | % {Add-DistributionGroupMember -Identity "Clients - Chinook UM Team" -Member $_}

<#
Set OoO replies

$oOo = '<p dir="ltr">We acknowledge receipt of your application for a position at Sustain and sincerely appreciate your interest in our company.</p>
<p dir="ltr">We will screen all applicants and select candidates whose qualifications seem to meet our needs. We will carefully consider your application during the initial screening and will contact you if you are selected to continue in the recruitment process. We wish you every success.</p>
<p dir="ltr">Many thanks</p>
'
Set-MailboxAutoReplyConfiguration -Identity $displayName.Replace(" ",".") -InternalMessage $oOo -ExternalMessage $oOo -AutoReplyState Enabled
#>

<#
Move e-mail addresses
$oldPf = Get-MailPublicFolder -Identity "\1.Public Folders\3.Accounts\Sustain\Management\Recruitment" 
$oldPf.EmailAddresses | % {if($_ -notmatch "@AnthesisLLC.onmicrosoft.com"){[array]$emailAddressesToMove += $_.Replace("smtp:","")}}
$newSm = Get-Mailbox -Identity $displayName.Replace(" ",".")
$emailAddressesToMove | %{
    $oldPf.EmailAddresses.Remove("smtp:$_")
    Set-MailPublicFolder -Identity $oldPf.Identity -EmailAddresses $oldPf.EmailAddresses
    Set-Mailbox $newSm.Identity -EmailAddresses @{add="$_"}
    }
(Get-Mailbox -Identity $displayName.Replace(" ",".")).EmailAddresses | fl



#>