function get-allToAddressXHours($recipientAddress,$hoursAgo){
    $dateEnd = get-date
    $dateStart = $dateEnd.AddHours(-$hoursAgo)
    Get-MessageTrace -StartDate $dateStart -EndDate $dateEnd -RecipientAddress $recipientAddress
    }

$duffers = Get-DistributionGroup -Filter "DisplayName -like '∂_*'"

$duffers | %{
    $_.EmailAddresses | % {
        $emailAddress = $_
        #if(matchContains -term $emailAddress -arrayOfStrings @("footprintreporter.com", "@anthesisgroup.com", "@bestfootforward.com", "@lrsconsultancy.com", "@sustain.co.uk", "@pcrrg.uk")){$emailAddressesToCheck += $_.Replace("smtp:","")}
        if(($_ -match "@anthesisgroup.com") -or ($_ -match "@bestfootforward.com") -or ($_ -match "@lrsconsultancy.com") -or ($_ -match "@sustain.co.uk") -or ($_ -match "@pcrrg.uk") -or ($_ -match "footprintreporter.com")){
            [array]$emailAddressesToCheck+= $_.Replace("smtp:","").Replace("SMTP:","")
            }
        
        }
    }

$emailAddressesToCheck 

$trace = get-allToAddressXHours -recipientAddress $emailAddressesToCheck -hoursAgo 720
$trace | Export-Csv -Path C:\Users\kevinm\Desktop\AuditLogs\duffDGMailTraces_2018_05_25.csv
