$tokenResponseSharePointBot = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName ReportBot) -grant_type client_credentials
$tokenResponseReportBot = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName ReportBot) -grant_type client_credentials

$storageReport = invoke-graphGet -tokenResponse $authResult -graphQuery "/reports/getSharePointSiteUsageStorage(period='D7')"
GET /reports/getSharePointSiteUsageDetail(period='{period_value}')
GET /reports/getSharePointSiteUsageDetail(date={date_value})
$usageReport = invoke-graphGet -tokenResponse $authResult -graphQuery "/reports/getSharePointSiteUsageDetail(period='$((Get-Date -f u).Split(" ")[0])')"
$usageReport = invoke-graphGet -tokenResponse $authResult -graphQuery "/reports/getSharePointSiteUsageDetail(period='D7')"
$usageReport | % {Add-Content $_ -Path $("$env:USERPROFILE\Downloads\UsageReport_$((Get-Date -f u).Split(" ")[0]).csv")}
$usageReport = invoke-graphGet -tokenResponse $authResult -graphQuery "/reports/getSharePointSiteUsageDetail(date=2021-09-27)"

$report  = invoke-graphGet -tokenResponse $tokenResponseReportBot -graphQuery "/reports/getOffice365ActiveUserDetail(period='D180')" 
$report2 = invoke-graphGet -tokenResponse $authResult -graphQuery "/reports/getOffice365ActiveUserDetail(period='D180')" 

$authResult | Add-Member -MemberType NoteProperty -Name access_token -Value $authResult.AccessToken

get-graphSite -tokenResponse $tokenResponseSharePointBot -


$fileReport = invoke-graphGet -tokenResponse $authResult -graphQuery "/reports/getSharePointSiteUsageFileCounts(period='D7')"

