

[array]$newChumps = convertTo-arrayOfEmailAddresses "contact@ceooffice.info"

$blackListTheseChumpsRuleName = "Blacklist these chumps"
$blackListRepliesToTheseChumpsRuleName = "Blacklist replies to these chumps"

$blackListTheseChumpsRule = Get-TransportRule | Where-Object {$_.Identity -contains $blackListTheseChumpsRuleName}
$blackListRepliesToTheseChumpsRule = Get-TransportRule | Where-Object {$_.Identity -contains $blackListRepliesToTheseChumpsRuleName}

$newChumps | % {
    $blackListTheseChumpsRule.From.Add($_)
    $blackListRepliesToTheseChumpsRule.SentTo.Add($_)
    }
$blackListTheseChumpsRule | Set-TransportRule -From $blackListTheseChumpsRule.From
$blackListTheseChumpsRuleName | Set-TransportRule -SentTo $blackListRepliesToTheseChumpsRule.SentTo
