Connect-ExchangeOnline -UserPrincipalName "exchangeAdmin@lavola.com"

#region Shared Mailbox Prep
$arrayOfCngSharedMailboxAddresses = convertTo-arrayOfEmailAddresses "accounts@climateneutralgroup.co.za
administratie@climateneutralgroup.com
administratieechtgoed@climateneutralgroup.com
AVG@climateneutralgroup.com
certificatengreenseat@climateneutralgroup.com
certification@climateneutralgroup.com
Communicatie@climateneutralgroup.com
communication@climateneutralgroup.com
consultancy@climateneutralgroup.com
website@climateneutralgroup.co.za
csc@climateneutralgroup.com
fcc@climateneutralgroup.com
ferry@echtgoed.nl
footprintdata@climateneutralgroup.com
greendreams@climateneutralgroup.com
Infodocdata@climateneutralgroup.com
Info@climateneutralgroup.com
Info@climateneutralgroup.co.za
info@greenseat.com
inspiratie@climateneutralgroup.com
internship@climateneutralgroup.com
Marketing@climateneutralgroup.co.za
nacalculatie.voetafdruk@climateneutralgroup.com
Nieuwsbrief@climateneutralgroup.com
officemanagement@climateneutralgroup.com
pers@climateneutralgroup.com
planning.simapro@climateneutralgroup.com
Recruitment@climateneutralgroup.com
Recruitment@climateneutralgroup.co.za
Sollicitatie@climateneutralgroup.com
"
$allMailboxes = Get-EXOMailbox -ResultSize Unlimited
$smbxToMigrate = $allMailboxes | Where-Object {$_.PrimarySmtpAddress -in $arrayOfCngSharedMailboxAddresses} #See later for $allMailboxes

$smbxToMigrate | ForEach-Object {
    $thisMailbox = $_
    Write-Output "[$($thisMailbox.PrimarySmtpAddress)]"
    $thisMailboxStats = Get-EXOMailboxStatistics -Identity  $thisMailbox.Identity
    $thisMailbox | Add-Member -MemberType NoteProperty -Name TotalItemSize -Value $thisMailboxStats.TotalItemSize -Force
    $thisMailboxPermissions = Get-EXOMailboxPermission -Identity $thisMailbox.ExternalDirectoryObjectId
    $thisMailboxDelegates = $thisMailboxPermissions | Where-Object {$_.User -notmatch "NT AUTHORITY" -and $_.User -ne "kev.maitland@climateneutralgroup.com" -and $_.AccessRights -contains "FullAccess"} #Strip out junk
    $thisMailbox | Add-Member -MemberType NoteProperty -Name Delegates -Value $thisMailboxDelegates.User -Force
    $thisMailboxSendAs = $thisMailboxPermissions | Where-Object {$_.User -notmatch "NT AUTHORITY" -and $_.User -ne "kev.maitland@climateneutralgroup.com" -and $_.AccessRights -contains "SendAs"} #Strip out junk
    $thisMailbox | Add-Member -MemberType NoteProperty -Name SendAs -Value $thisMailboxSendAs.User -Force
    $thisMailboxOoO = Get-MailboxAutoReplyConfiguration -Identity  $thisMailbox.Identity
    $thisMailbox | Add-Member -MemberType NoteProperty -Name OoO -Value $thisMailboxOoO -Force
    $anthesisAddress = "lavola-$($thisMailbox.PrimarySmtpAddress.Split("@")[0])@anthesisgroup.com"
    $anthesisDisplayName = "$($thisMailbox.DisplayName) - Lavola Shared Mailbox"

    $thisMailbox | Add-Member -MemberType NoteProperty -Name anthesisAddress -Value $anthesisAddress -Force
    $thisMailbox | Add-Member -MemberType NoteProperty -Name anthesisDisplayName -Value $anthesisDisplayName -Force
    $thisMailbox | Add-Member -MemberType NoteProperty -Name nonRoutableUpn -Value "$($thisMailbox.UserPrincipalName.Split("@")[0])@lavola.onmicrosoft.com" -Force
}


