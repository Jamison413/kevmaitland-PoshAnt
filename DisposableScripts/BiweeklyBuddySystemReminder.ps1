#Set some logging
$Logname = "C:\Scripts" + "\Logs" + "\BuddySystem $(Get-Date -Format "yyMMdd").log" #Check this location before live
Start-Transcript -Path $Logname -Append
Write-Host "Script started:" (Get-date)
Write-host "**********************" -ForegroundColor White

#Connect to site
Import-Module _PS_Library_GeneralFunctionality
Import-Module _PNP_Library_SPO

$smtpaddress = "groupbot@anthesisgroup.com"

$smtpBotDetails = get-graphAppClientCredentials -appName SmtpBot
$tokenResponseSmtp = get-graphTokenResponse -aadAppCreds $smtpBotDetails


$sharePointAdmin = "kimblebot@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content "$env:USERPROFILE\Downloads\KimbleBot.txt") 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass

#Get all the current users in the waiting list
$connect = Connect-PnPOnline -url "https://anthesisllc.sharepoint.com/teams/IT_Team_All_365/" -Credentials $adminCreds
$allwaiting = Get-PnPListItem -List "Buddy System Waiting List"

ForEach($person in $allwaiting){
$a = Add-PnPListItem -List "Buddy System Repeat Sign Up" -Values @{"Yourname" = $($person.FieldValues.Yourname.LookupValue); "Yourcommunity_x0028_ifapplicable" = $($person.FieldValues.Yourcommunity_x0028_ifapplicable); "Youtimezone" = $($person.FieldValues.Youtimezone); "Yourcountry" = $($person.FieldValues.Yourcountry)} 
$b = Remove-PnPListItem -List "Buddy System Waiting List" -Identity $($person.Id) -Force
}


$allwaitingCheck = Get-PnPListItem -List "Buddy System Waiting List"
$allRepeatSignUpCheck = Get-PnPListItem -List "Buddy System Repeat Sign Up"

$subject = "Anthesis Buddy System Reminder Email"
$body3 = "<HTML><FONT FACE=`"Calibri`">I should have processed reminders `r`n`r`n<BR><BR>"
$body3 += "Counts - WaitingList (should be 0): $(($allwaitingCheck | Measure-Object).count) `r`n`r`n<BR><BR>"
$body3 += "Counts - RepeatSignUpList (should be <0): $(($allRepeatSignUpCheck | Measure-Object).count) `r`n`r`n<BR><BR>"
$body3 += "$($allwaiting.FieldValues.Yourname.LookupValue)<BR><BR>"
send-graphMailMessage -tokenResponse $tokenResponseSmtp -fromUpn buddy.system@anthesisgroup.com -toAddresses "emily.pressey@anthesisgroup.com" -subject $subject -bodyHtml $body3 -priority high



Write-host "**********************" -ForegroundColor White
Write-Host "Script finished:" (Get-date)

Stop-Transcript



