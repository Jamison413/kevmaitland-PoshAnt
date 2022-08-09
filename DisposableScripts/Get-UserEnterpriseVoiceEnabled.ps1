Import-Module -name MicrosoftTeams
Update-Module -name MicrosoftTeams 
Connect-MicrosoftTeams

#Get all Teams users with enterprise voice enabled and spit them out to the screen so you can eyeball them
Get-CsOnlineUser | Where-Object "EnterpriseVoiceEnabled" -eq "$True" | select displayname, enterprisevoiceenabled, hostedvoicemail, dialplan | Out-GridView

#Enable Enterprise voice for the targeted user
Set-CsPhoneNumberAssignment -Identity "lucy.dornan@anthesisgroup.com" -EnterpriseVoiceEnabled $True