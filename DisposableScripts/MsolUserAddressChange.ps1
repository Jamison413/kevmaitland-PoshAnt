#Import your modules
Import-Module -Name AzureAD
Import-Module -Name MSOnline
Import-Module -Name ExchangeOnlineManagement
#Import-Module -Name PnP.PowerShell

#Connectify with your credentials
connect-ToMsol -interactive
connect-ToExo -interactive
connect-toAAD -interactive

#Find all office based users for specific office
$UsersToFix = Get-msoluser -All | where-object {$_.islicensed -eq "True" -and $_.City -eq "London, GBR" -and $_.Office -eq "London, GBR" } 
$UsersToFix | Set-MsolUser -StreetAddress "Floor 1, Fitzroy House, 355 Euston Road" -City "London, GBR" -PostalCode "NW1 3AL" -Office "London, GBR" -Country "United Kingdom" 

#Find all Home based users for a specific office
$HomeUsersToFix = Get-MsolUser -All | Where-Object {$_.islicensed -eq "True" -and $_.City -eq "Home worker" -or $_.City -eq "Homeworker" -and $_.Office -eq "London, GBR" }
$HomeUsersToFix | Set-MsolUser -StreetAddress "Floor 1, Fitzroy House, 355 Euston Road" -City "Home worker" -PostalCode "NW1 3AL" -Office "London, GBR" -Country "United Kingdom" 


