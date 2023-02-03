Import-Module _PS_Library_MSOL.psm1
Import-Module _PS_Library_GeneralFunctionality
Import-Module _PS_Library_Groups

set-MsolCredentials
connect-msolService
connect-ToExo


$displayName = "displayName"
$primaryEmail = "name1.name2@anthesisgroup.com"
$owner = "firstname.lastname@anthesisgroup.com"
$arrayOfFullAccessMembers = convertTo-arrayOfEmailAddresses "firstname.lastname@anthesisgroup.com; firstname.lastname@anthesisgroup.com"
$additionalEmailAddresses = convertTo-arrayOfEmailAddresses ""
$allEmailAddresses = convertTo-arrayOfEmailAddresses "$primaryEmail , $additionalEmailAddresses"
$grantSendAsToo = $true
$hideFromGal = $false


function new-sharedMailbox($displayName, $owner, $arrayOfFullAccessMembers, $hideFromGal, $grantSendAsToo){
    $exchangeAlias = $(guess-aliasFromDisplayName -displayName $displayName)
    New-Mailbox -Shared -ModeratedBy $owner -DisplayName $displayName -Name $displayName -Alias $exchangeAlias -PrimarySmtpAddress $primaryEmail | Set-Mailbox -HiddenFromAddressListsEnabled $hideFromGal -RequireSenderAuthenticationEnabled $false -MessageCopyForSendOnBehalfEnabled $true -MessageCopyForSentAsEnabled $true -EmailAddresses $allEmailAddresses
    $arrayOfFullAccessMembers  | %{
        Add-MailboxPermission -AccessRights "FullAccess" -User $_ -AutoMapping $true -Identity $exchangeAlias
        if ($grantSendAsToo){Add-RecipientPermission -Identity $exchangeAlias -Trustee $_ -AccessRights SendAs -Confirm:$false}
        }
    Set-User -Identity $exchangeAlias -AuthenticationPolicy "Block Basic Auth"
    }

new-sharedMailbox -displayName $displayName -arrayOfFullAccessMembers $arrayOfFullAccessMembers -hideFromGal $hideFromGal -owner $owner -grantSendAsToo $grantSendAsToo


#Block sign-in for new shared mailbox
$newsharedmailbox = get-user -Identity $primaryEmail
If($newsharedmailbox){
Set-MsolUser -UserPrincipalName $newsharedmailbox.UserPrincipalName -BlockCredential $true
}
Else{
Write-Host "Couldn't find new shared mailbox $($primaryEmail), this could be a time delay - try again in a few minutes" -ForegroundColor Red
}




<#Set current shared mailbox authentication policy
$newsharedmailbox = get-user -Identity $primaryEmail
If($newsharedmailbox){
Set-User -Identity $newsharedmailbox.UserPrincipalName -AuthenticationPolicy "Allow IMAP"
}
Else{
Write-Host "Couldn't find new shared mailbox $($primaryEmail), this could be a time delay - try again in a few minutes" -ForegroundColor Red
}
#>


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

