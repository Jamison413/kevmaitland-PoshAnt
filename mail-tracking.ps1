#anthesisgroup-com.mail.protection.outlook.com
$credential = get-credential -Credential kevin.maitland@anthesisgroup.com
Import-Module MSOnline
Connect-MsolService -Credential $credential

$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
Import-PSSession $ExchangeSession


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
    $traceBlob | Select-Object Received, SenderAddress, RecipientAddress, Subject, Status, ToIP, FromIP, Size, MessageID, MessageTraceID | Out-GridView 
    }

    
$trace = get-allFromAddressXHours -senderAddress "xerox9303main@sustain.co.uk" -hoursAgo 10
$trace = get-allToAddressXHours -recipientAddress "bidstenders@sustain.co.uk" -hoursAgo 10