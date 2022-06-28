





connect-ToMsol -credential $msolCredentials
connect-ToExo -credential $msolCredentials
connect-toAAD -credential $msolCredentials





Get-MsolUser -All | Where-Object {($_.Licenses).AccountSkuId -match "EXCHANGEDESKLESS" -or ($_.Licenses).AccountSkuId -match  "ENTERPRISEPACK" -or ($_.Licenses).AccountSkuId -match "STANDARDPACK" -or ($_.Licenses).AccountSkuId -match "ENTERPRISEPREMIUM" -and 
    $_.UserPrincipalName -notmatch "conflictminerals" -and $_.UserPrincipalName -notmatch "acsmailboxaccess" -and $_.UserPrincipalName -notmatch "SoS.test" -and $_.UserPrincipalName -notmatch "t0-" -and 
        $_.UserPrincipalName -notmatch "t1-" -and $_.UserPrincipalName -notmatch "UKcareers" -and $_.UserPrincipalName -notmatch "Varex.PEC" -and $_.UserPrincipalName -notmatch "Barry.Holt" -and 
        $_.UserPrincipalName -notmatch "AnthesisUKFinance" -and $_.UserPrincipalName -notmatch "Microsoft.ECM" -and $_.UserPrincipalName -notmatch "ACSSupport" -and 
        $_.UserPrincipalName -notmatch "VarexConflictMinerals"}  | Select DisplayName, UsageLocation | Export-Csv -path C:\Users\Andrew.Ost\Desktop\CSVs\AnthesisGlobalPresence2.csv
