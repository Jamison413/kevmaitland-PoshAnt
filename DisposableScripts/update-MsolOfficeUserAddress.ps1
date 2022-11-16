#Import your modules
Import-Module -Name MSOnline

#Connectify with your credentials
connect-ToMsol -interactive

<#Current Bristol office address
Floor 4, Runway East Bristol Bridge, 1 Victoria Street BS1 6AA

Current London office address
Floor 1, Fitzroy House, 355 Euston Road, NW1 3AL

#>

#Find all office based users for specific office
$UsersToFix = Get-msoluser -All | where-object {$_.islicensed -eq "True" -and $_.City -eq "Bristol, GBR" -and $_.Office -eq "Bristol, GBR"} 

#Set the office address for users returned
$UsersToFix | Set-MsolUser -StreetAddress "Floor 4, Runway East Bristol Bridge, 1 Victoria Street" -City "Bristol, GBR" -PostalCode "BS1 6AA" -Office "Bristol, GBR" -Country "United Kingdom" 

#Find all Home workers for a specific office
$HomeworkersToFix = Get-MsolUser -All | Where-Object {$_.islicensed -eq "True" -and $_.City -eq "Home worker" -or $_.City -eq "Homeworker" -and $_.Office -eq "Bristol, GBR"}

#Set the office address for users returned
$HomeworkersToFix | Set-MsolUser -StreetAddress "Floor 4, Runway East Bristol Bridge, 1 Victoria Street" -City "Home worker" -PostalCode "BS1 6AA" -Office "Bristol, GBR" -Country "United Kingdom" 
