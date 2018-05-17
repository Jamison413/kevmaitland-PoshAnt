$users = convertTo-arrayOfStrings "mark.sayers"
$users | % {
    set-mailbox $_ -CustomAttribute2 "CanScrape"
    }