#anthesisgroup-com.mail.protection.outlook.com
Import-Module _PS_Library_MSOL.psm1
connect-ToExo

function get-allMailLastXHours($hoursAgo){
    $dateEnd = get-date
    $dateStart = $dateEnd.AddHours(-$hoursAgo)
    Get-MessageTrace -StartDate $dateStart -EndDate $dateEnd 
    }
function get-allFromAddressXHours($senderAddress,$hoursAgo){
    $dateEnd = get-date
    $dateStart = $dateEnd.AddHours(-$hoursAgo)
    Get-MessageTrace -StartDate $dateStart -EndDate $dateEnd -SenderAddress $senderAddress
    }
function get-allToAddressXHours($recipientAddress,$hoursAgo){
    $dateEnd = get-date
    $dateStart = $dateEnd.AddHours(-$hoursAgo)
    Get-MessageTrace -StartDate $dateStart -EndDate $dateEnd -RecipientAddress $recipientAddress
    }
function format-MailTracePrettily($traceBlob){
    #$trace | Select-Object Received, SenderAddress, RecipientAddress, Subject, Status, ToIP, FromIP, Size, MessageID, MessageTraceID | export-csv -NoTypeInformation -Path C:\Users\kevin.maitland\Desktop\mailTrace.log -Encoding UTF8
    $trace | Out-GridView 
    }

    
$trace = get-allToAddressXHours -recipientAddress "uk.software@anthesisgroup.com" -hoursAgo 720
$trace = get-allToAddressXHours -recipientAddress "US-Anthesis@anthesisgroup.com" -hoursAgo 48
$trace = get-allFromAddressXHours -senderAddress "graeme.hadley@anthesisgroup.com" -hoursAgo 384


format-MailTracePrettily $trace
$trace | Export-Csv -Path "$env:USERPROFILE\Desktop\MailTrace_$(Get-Date -Format yyyy-MM-dd).csv" -NoTypeInformation


Import-Module LyncOnlineConnector
Get-OrganizationConfig | Format-Table -Auto Name,OAuth*
Get-CsOAuthConfiguration

Set-OrganizationConfig -OAuth2ClientProfileEnabled $true
Set-CsOAuthConfiguration -ClientAdalAuthOverride Allowed

$users.windowsliveid