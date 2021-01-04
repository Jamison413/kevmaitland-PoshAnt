$exoCreds = set-MsolCredentials
connect-ToExo -credential $exoCreds

[array]$newChumps = convertTo-arrayOfEmailAddresses "zoom.notfication@mnioose.com"

$blackListTheseChumpsRuleName = "Blacklist these chumps"
$blackListRepliesToTheseChumpsRuleName = "Blacklist replies to these chumps"

$blackListTheseChumpsRule = Get-TransportRule | Where-Object {$_.Identity -contains $blackListTheseChumpsRuleName}
$blackListRepliesToTheseChumpsRule = Get-TransportRule | Where-Object {$_.Identity -contains $blackListRepliesToTheseChumpsRuleName}

$newChumps | % {
    $blackListTheseChumpsRule.From.Add($_)
    $blackListRepliesToTheseChumpsRule.SentTo.Add($_)
    }
$blackListTheseChumpsRule | Set-TransportRule -From $blackListTheseChumpsRule.From
$blackListRepliesToTheseChumpsRule | Set-TransportRule -SentTo $blackListRepliesToTheseChumpsRule.SentTo
