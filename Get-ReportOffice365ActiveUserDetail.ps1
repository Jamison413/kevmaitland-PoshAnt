
    
$reportbotDetails = get-graphAppClientCredentials -appName ReportBot
$tokenResponseReportBot = get-graphTokenResponse -aadAppCreds $reportbotDetails -grant_type client_credentials
$report = invoke-graphGet -tokenResponse $tokenResponseReportBot -graphQuery "/reports/getOffice365ActiveUserDetail(period='D180')" 
$arrayReport = ($report -split '\r?\n').Trim()
$arrayReport  | out-file $env:TEMP\test.csv -Force utf8


$csvReport = Import-Csv -Path $env:TEMP\test.csv 
$csvReport | Export-Csv -Path $env:USERPROFILE\Downloads\test4.csv -NoTypeInformation -Force -Encoding UTF8



