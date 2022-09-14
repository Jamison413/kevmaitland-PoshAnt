Import-Module -name MicrosoftTeams
#Update-Module -name MicrosoftTeams 
Connect-MicrosoftTeams

#Get all Teams users with enterprise voice enabled and spit them out to the screen so you can eyeball them
Get-CsOnlineUser | Where-Object "EnterpriseVoiceEnabled" -eq "$True" | select displayname, enterprisevoiceenabled, Country | Export-Csv -Path C:\Users\AndrewOst\Documents\PowerCsv\AllEntVoiceUsers.csv -NoTypeInformation

#Enable Enterprise voice for the targeted user
Set-CsPhoneNumberAssignment -Identity "" -EnterpriseVoiceEnabled $True

 