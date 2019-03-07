#new-MailTraceDetailed

$daysToLookBack = 1
if(!$toDate){$toDate = $(Get-Date).AddDays(1)}
$fromDate = $toDate.AddDays(-($daysToLookBack+1))
$toAddress = "sophie.taylor@anthesisgroup.com"
$trace = Get-MessageTrace -EndDate $toDate -StartDate $fromDate -RecipientAddress $toAddress
$details = $trace | Get-MessageTraceDetail

for ($i = 0; $i -lt $trace.Count; $i++){
    $trace[$i] | fl
    $details | ? {$_.MessageId -eq $trace[$i].MessageId} | fl
    }