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
    Get-MessageTrace -StartDate $dateStart -EndDate $dateEnd -SenderAddress $senderAddress -PageSize 5000
    }
function get-allToAddressXHours($recipientAddress,$hoursAgo){
    $dateEnd = get-date
    $dateStart = $dateEnd.AddHours(-$hoursAgo)
    Get-MessageTrace -StartDate $dateStart -EndDate $dateEnd -RecipientAddress $recipientAddress
    }
function format-MailTracePrettily($traceBlob){
   # $trace | Select-Object Received, SenderAddress, RecipientAddress, Subject, Status, ToIP, FromIP, Size, MessageID, MessageTraceID | export-csv -NoTypeInformation -Path C:\Users\kevin.maitland\Desktop\mailTrace.log -Encoding UTF8
    $trace | Out-GridView 
    }

$emailAddressForTrace = "Alan.Matthews@anthesisgroup.com"
    
$trace = get-allToAddressXHours -recipientAddress $emailAddressForTrace -hoursAgo 720
$trace = get-allToAddressXHours -recipientAddress $emailAddressForTrace -hoursAgo 5
$trace = get-allFromAddressXHours -senderAddress $emailAddressForTrace -hoursAgo 30


format-MailTracePrettily $trace
$trace | Export-Csv -Path "$env:USERPROFILE\Desktop\AuditLogs\MailTrace_$emailAddressForTrace$(Get-Date -Format yyyy-MM-dd).csv" -NoTypeInformation


<#
Import-Module LyncOnlineConnector
Get-OrganizationConfig | Format-Table -Auto Name,OAuth*
Get-CsOAuthConfiguration

Set-OrganizationConfig -OAuth2ClientProfileEnabled $true
Set-CsOAuthConfiguration -ClientAdalAuthOverride Allowed

$users.windowsliveid

$OoO = "Thanks for your email. I am currently out of office on paternity leave until the 2 January 2018, and will be in touch shortly upon return. For project related matters, please contact Margaret Davis on 0117 403 2663."
#>

<#$OoO = '<HTML>
<body>
<div>
<p style="font-family: Calibri,Verdana,Arial; color: black;">Thank you for your email. </p>
<p style="font-family: Calibri,Verdana,Arial; color: black;">The Finance Department will process any invoices and respond to any enquires within 2-3 working days.</p>
<p style="font-family: Calibri,Verdana,Arial; color: black;">Our company name has changed to Anthesis Energy UK Ltd and our email addresses have changed as well. Please update your records.</p>
<p style="font-family: Calibri,Verdana,Arial; color: black;">Invoices to be sent to energyinvoices@anthesisgroup.com</p>
<p style="font-family: Calibri,Verdana,Arial; color: black;">Remittances advices to energyremittances@anthesisgroup.com</p>
<p style="font-family: Calibri,Verdana,Arial; color: black;">Enquiries and statements to energyfinance@anthesisgroup.com</p>
<p style="font-family: Calibri,Verdana,Arial; color: black;">If you have any queries then please contact kath.addison-scott@anthesisgroup.com or greg.francis@anthesisgroup.com</p>
<p style="font-family: Calibri,Verdana,Arial; color: black;">Kind Regards,</p>
<p style="font-family: Calibri,Verdana,Arial; color: black;">Anthesis Energy UK''s AutoReply Robot</p>
<p></p>
</div>
</body>
</HTML>'
#>
