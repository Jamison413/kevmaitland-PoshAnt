#Get-ADUser -Filter * -Properties * | ? {$_.PasswordExpired -eq $true} 
Start-Transcript "$($MyInvocation.MyCommand.Definition).log" -Append


Import-Module ActiveDirectory
$smtpServer = "anthesisgroup-com.mail.protection.outlook.com"

#Get-ADUser  -filter * | select Name
Get-ADUser -Filter * -Properties * -SearchBase "OU=Users,OU=Sustain,DC=Sustainltd,DC=local" | select name,PasswordLastSet,PasswordNeverExpires,PasswordExpired 
#Get-ADUser -Filter * -Properties * -SearchBase "OU=Users,OU=Sustain,DC=Sustainltd,DC=local" | ? {$_.PasswordLastSet -lt (Get-Date).AddDays(-83) -and $_.PasswordNeverExpires -eq $false -and $_.PasswordExpired -eq $false}

foreach ($user in Get-ADUser -Filter * -Properties * -SearchBase "OU=Users,OU=Sustain,DC=Sustainltd,DC=local" | ? {$_.PasswordLastSet -lt (Get-Date).AddDays(-83) -and $_.PasswordNeverExpires -eq $false -and $_.PasswordExpired -eq $false}){
    "E-mailing "+$user.Name
    Send-MailMessage -From "itnn@sustain.co.uk" -To $user.EmailAddress -SmtpServer $smtpServer -Subject "Password expiry warning" -Body "Hello $($user.GivenName),
    
I just wanted to let you know that your `"$($user.Company.Replace(" Limited",""""))`" password will expire on $($user.PasswordLastSet.AddDays(90)). 
    
If you are in the office, you can change your password by:
`t1) Signing  into your laptop 
`t2) Pressing Ctrl-Alt-Del 
`t3) Choosing `"Change a password`"

If you are out of the office, you can change your password by:
`t1) Signing into the Remote Desktop Server (http://remote.sustain.co.uk)
`t2) Pressing Ctrl-Alt-End  (not Del)
`t3) Choosing `"Change a password`"

I'd strongly recommend that you change you Anthesis password too, to keep them in sync:
`t1) Go to https://portal.office.com and log in
`t2) Click on the Cog (top-right), then Password
`t3) Enter your old password again, followed by your new password twice then click Submit

If you pick up your e-mail on your phone, you will automatically be prompted to to update the password on there once you've updated you Anthesis one. 

Love,

The PasswordReminderRobot"
    }

Stop-Transcript
#If you use the ANTHESIS WiFi for anything other than your laptop, you will need to connect to it again using your username ($(($user.DistinguishedName -Split ',DC=')[1])\$($user.SamAccountName)) and your new password