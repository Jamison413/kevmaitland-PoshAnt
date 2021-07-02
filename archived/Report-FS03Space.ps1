$Logname = "C:\Scripts" + "\Logs" + "\report-FS03Space $(Get-Date -Format "yyMMdd").log"
Start-Transcript -Path $Logname -Append


$disk = Get-WmiObject Win32_LogicalDisk -ComputerName Fs03 -Filter "DeviceID='D:'" |
Select-Object Size,FreeSpace

($disk.Size / 1GB)
$freespace = [Math]::Round($disk.FreeSpace / 1GB)

$body = "<HTML><BODY><p>Hi IT Team </p>

<p>FS03 is running out of space again :(</p>
<p>Free space: $($freespace)gb left</p>

<p>Love,</p>

<p>The Unfriendly 'FS03 is outta space' Robot</p>
</BODY></HTML>"



If($freespace -lt 8){
Send-MailMessage  -BodyAsHtml $body -Subject "FS03 is almost out of space" -to "emily.pressey@anthesisgroup.com" -from "FS03Robot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Encoding UTF8 
}


Stop-Transcript

