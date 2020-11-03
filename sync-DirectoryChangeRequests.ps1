#NOTE
#This is the second half of the Anthesis Directory Sync Scripts. This just handles change requests via the People Services and Administration Teams, which writes back to 365 and to the Anthesis Directory and Reporting Lists.


$friendlyLogname = "C:\ScriptLogs" + "\friendlylogsync-PeopleDirectory $(Get-Date -Format "yyMMdd").log"
function friendlyLogWrite(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [string]$friendlyLogname
       ,[parameter(Mandatory = $true)]
            [string]$logstring
       ,[parameter(Mandatory = $true)]
            [validateset("WARNING","SUCCESS","ERROR","ERROR DETAILS","MESSAGE","Change in 365","Change by Request","END","START")]
            [String]$messagetype
        )
If($messagetype -eq "MESSAGE"){
Add-content $friendlyLogname -value $("*************************************************************************************************************************************************************")
Add-content $friendlyLogname -value $("$(get-date)" + " MESSAGE: " + "$($logstring)")
Add-content $friendlyLogname -value $("*************************************************************************************************************************************************************")
}

If($messagetype -eq "START"){
Add-content $friendlyLogname -value $("-------------------------------------------------------------------------------------------------------------------------------------------------------------")
Add-content $friendlyLogname -value $("$(get-date)" + " START: " + "$($logstring)")
Add-content $friendlyLogname -value $("-------------------------------------------------------------------------------------------------------------------------------------------------------------")
}


If($messagetype -eq "END"){
Add-content $friendlyLogname -value $("-------------------------------------------------------------------------------------------------------------------------------------------------------------")
Add-content $friendlyLogname -value $("$(get-date)" + " END: " + "$($logstring)")
Add-content $friendlyLogname -value $("-------------------------------------------------------------------------------------------------------------------------------------------------------------")
}


If($messagetype -eq "WARNING"){
Add-content $friendlyLogname -value $("$(get-date)" + "     WARNING: " + "$($logstring)")
}

If($messagetype -eq "SUCCESS"){
$content = 
Add-content $friendlyLogname -value $("$(get-date)" + "     SUCCESS: " + "$($logstring)")
}

If($messagetype -eq "ERROR"){
Add-content $friendlyLogname -value $("$(get-date)" + "     ERROR: " + "$($logstring)")
}

If($messagetype -eq "ERROR DETAILS"){
$value = $("$(get-date)" + "     ERROR DETAILS: " + "$($logstring)")
Add-content $friendlyLogname -value $value
}

If($messagetype -eq "Change by Request"){
$value = $("<>" + " Change by Request: " + "$($logstring)" + "<>")
Add-content $friendlyLogname -value $value
}

If($messagetype -eq "Change in 365"){
$value = $("<>" + " Change in 365: " + "$($logstring)" + "<>")
Add-content $friendlyLogname -value $value
}
}
$TeamsLog = @()

$Logname = "C:\ScriptLogs" + "\sync-PeopleDirectory $(Get-Date -Format "yyMMdd").log"
Start-Transcript -Path $Logname -Append
Write-Host "Script started:" (Get-date)

Import-Module _PS_Library_Graph.psm1
Import-Module _PNP_Library_SPO.psm1
Import-Module _CSOM_Library-SPO.psm1
Import-Module _PS_Library_UserManagement.psm1


$groupAdmin = "groupbot@anthesisgroup.com"
$groupAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\GroupBot.txt) 
$exoCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $groupAdmin, $groupAdminPass
connect-ToExo -credential $exoCreds


#Conn - CSOM for SharepointUserID
$sharePointAdmin = "kimblebot@anthesisgroup.com"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\kimblebot.txt) 
$spoCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
$peopleservicessite = "https://anthesisllc.sharepoint.com/teams/People_Services_Team_All_365/"
$conn = Connect-PnPOnline -Url $peopleservicessite -Credentials $spoCreds
$ctx = Get-PnPContext


#Conn - Graph for overall user profile
$teamBotDetails = Import-Csv "$env:USERPROFILE\Desktop\teambotdetails.txt"
$resource = "https://graph.microsoft.com"
$tenantId = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.TenantId)
$clientId = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.ClientID)
$redirect = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.Redirect)
$secret   = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.Secret)

$ReqTokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    client_Id     = $clientID
    Client_Secret = $secret
    } 
$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody

friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype START -logstring "Starting run for sync-DirectoryChangeRequests"

#Set  Sharepoint list id's
$graphSiteId = "anthesisllc.sharepoint.com,cd82f435-8404-4c16-9ef5-c1e357ac5b96,2373d950-6dea-4ed5-9224-dea4c41c7da3"
$directoryListId = "009bb573-f305-402d-9b21-e6f597473256"
$changeListId = "8d633223-396f-4c45-974e-35ae8871d886"
$reportinglinesListId = "42dca4b4-170c-4caf-bcfe-62e00cb62819"


#Get all licensed graph users
$usersarray = get-graphUsers -tokenResponse $tokenResponse -filterLicensedUsers:$true -selectAllProperties:$true -Verbose
$allgraphusers = remove-mailboxesandbots -usersarray $usersarray
#Get all current Anthesians in the list
$allanthesians = get-graphListItems -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $directoryListId -expandAllFields
$allanthesianGUIDS = $allanthesians | select -ExpandProperty "fields"
#Get all current Live Reporting Lines List
$allPOPreports = get-graphListItems -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $reportinglinesListId -expandAllFields

#Set Teams Messaging holder
$TeamsReport = @()

<#------------------------------------------------------------------------------------Process Changes by Request--------------------------------------------------------------------------------------------#>

#Get change requests from Sharepoint List
$allchanges = get-graphListItems -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $changeListId -expandAllFields
$livechanges = $allchanges | Select-Object -ExpandProperty Fields | Where-Object -Property "Status" -EQ "Awaiting"
If($livechanges){
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype MESSAGE -logstring "We've found some change requests to process [$(($livechanges | Measure-Object).Count)]"
}
Else{
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype MESSAGE -logstring "We've found no change requests to process"
}
ForEach($change in $livechanges){



#Get the graphuser, their entry in the staff list and the POP Reporting Lines list first
$graphuser = get-graphUsers -tokenResponse $tokenResponse -filterUpns $change.AnthesianEmail -selectAllProperties:$true
$thisanthesian = $allanthesians  | Where-Object {$_.fields.Email -eq "$($change.AnthesianEmail)"}
$thisreport = $allPOPreports | Where-Object {$_.fields.Email -eq "$($change.AnthesianEmail)"}



#Check they aren't a 365 admin, we can't set some properties on them due to security
$isadmin = get-graphAdministrativeRoleMembers -tokenResponse $tokenResponse -roleName 'User Account Administrator' | Where-Object -Property "userPrincipalName" -EQ $graphuser.userPrincipalName
$isglobaladmin = get-graphAdministrativeRoleMembers -tokenResponse $tokenResponse -roleName 'Company Administrator' | Where-Object -Property "userPrincipalName" -EQ $graphuser.userPrincipalName
If((!$isglobaladmin) -and (!$isadmin)){
Write-Host "Not an admin/global admin, continuing..." -ForegroundColor Yellow
}
Else{
Write-Host "Is an admin/global admin, breaking..." -ForegroundColor Red
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype WARNING -logstring "[$($graphuser.userPrincipalName)] is a either an admin or global admin, we cannot amend their details"
$changetable = $change | Format-Table JobTitle,CellPhone,Office_x0020_Number,Community,Office,City,ManagerEmail,Business_x0020_Unit
$subject = "Woops! We tried to update 365 details for either an IT Admin"
$body = "<HTML><FONT FACE=`"Calibri`">Hello there,`r`n`r`n<BR><BR>"
$body += "You're receiving this email as someone has tried update your details in 365. This won't be possible as you are a 365 admin and will need to complete these changes yourself due to security.`r`n`r`n<BR><BR>"
$body += "Job title: $($change.JobTitle)`r`n<BR><BR>"
$body += "Mobile number: $($change.CellPhone)`r`n<BR><BR>"
$body += "Office number: $($change.Office_x0020_Number)`r`n<BR><BR>"
$body += "City: $($change.City)`r`n<BR><BR>"
$body += "Office: $($change.Office)`r`n<BR><BR>"
$body += "Manager: $($change.ManagerEmail)`r`n<BR><BR>"
$body += "Business Unit: $($change.Business_x0020_Unit)`r`n`r`n<BR><BR><BR><BR>"
$body += "Love,`r`n`r`n<BR><BR>"
$body += "The People Services Robot"
Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
update-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $changeListId -listitemId "$($change.id)" -fieldHash @{"Status" = "This Anthesian is in the IT Team - we'll let them know the changes"} -Verbose
Break
}

If($($change.JobTitle) -ne $($graphuser.jobTitle)){
    $error.clear()
    write-host "Job title has been changed...amending 365" -ForegroundColor White
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'Change by Request' -logstring "Updating the job title for [$($graphuser.userPrincipalName)]: $($graphuser.jobtitle) to $($change.JobTitle)"
    set-graphuser -tokenResponse $tokenResponse -userIdOrUpn  $graphuser.id -userPropertyHash @{"jobTitle" = $($change.JobTitle)} -Verbose
    $jobtitlecheck = get-graphUsers -tokenResponse $tokenResponse -filterUpns $graphuser.userPrincipalName
    #Check graph user was changed 
    If($jobtitlecheck.jobtitle -eq $($change.JobTitle)){
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(365) Updated the job title for [$($graphuser.userPrincipalName)]: $($graphuser.jobtitle) to $($jobtitlecheck.jobTitle)"
        #Update Directory List (reporting list only has name of employee and manager fields, so we won't update this one)
        $jobtitledirectoryupdate = update-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $directoryListId -listitemId "$($thisanthesian.id)" -fieldHash @{"JobTitle" = "$($change.JobTitle)"} -Verbose
        If($jobtitledirectoryupdate.JobTitle -eq $change.JobTitle){
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(Directory) Updated the job title for [$($graphuser.userPrincipalName)]: $($graphuser.jobtitle) to $($change.JobTitle)"
        }
        Else{
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(Directory) Something went wrong updating the job title for [$($graphuser.userPrincipalName)] in the Directory"
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
        $TeamsReport += @{"$(get-date) (Change Request > Directory List) ERROR - Job Title not changed to $($change.JobTitle)" = "[$($graphuser.userPrincipalName)]"}
        }
    }
    Else{
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(365) Something went wrong updating the job title for [$($graphuser.userPrincipalName)] in 365"
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
    $TeamsReport += @{"$(get-date) (Change Request > 365/Directory List) ERROR - Job Title not changed to $($change.JobTitle)" = "[$($graphuser.userPrincipalName)]"}
    }
}
If($($change.CellPhone) -ne $($graphuser.mobilePhone) -and ("no number" -ne $change.CellPhone)){
    $error.clear()
    write-host "Mobile number has been changed...amending 365" -ForegroundColor White
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'Change by Request' -logstring "Updating the mobile number for [$($graphuser.userPrincipalName)]: $($graphuser.mobilePhone) to $($change.CellPhone)"
    set-graphuser -tokenResponse $tokenResponse -userIdOrUpn  $graphuser.id -userPropertyHash @{"mobilePhone" = $($change.CellPhone)}
    $mobilecheck = get-graphUsers -tokenResponse $tokenResponse -filterUpns $graphuser.userPrincipalName
    #Check graph user was changed 
    If($mobilecheck.mobilePhone -eq $($change.CellPhone)){
       #Update Directory List (reporting list only has name of employee and manager fields, so we won't update this one)
       friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(365) Updated the mobile number for [$($graphuser.userPrincipalName)]: $($graphuser.mobilePhone) to $($change.JobTitle)"
       $mobiledirectoryupdate = update-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $directoryListId -listitemId "$($thisanthesian.id)" -fieldHash @{"CellPhone" = "$($change.CellPhone)"} -Verbose
       If($mobiledirectoryupdate.CellPhone -eq $change.CellPhone){
       friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(Directory) Updated the mobile number for [$($graphuser.userPrincipalName)]: $($graphuser.mobilePhone) to $($change.CellPhone)"
       }
       Else{
       friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(Directory) Something went wrong updating the mobile number for [$($graphuser.userPrincipalName)] in the Directory"
       friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
       $TeamsReport += @{"$(get-date) (Change Request > Directory List) ERROR - Mobile number not changed to $($change.CellPhone)" = "[$($graphuser.userPrincipalName)]"}
       }
    }
    Else{
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(365) Something went wrong updating the mobile number for [$($graphuser.userPrincipalName)] in 365"
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
    $TeamsReport += @{"$(get-date) (Change Request > 365/Directory List) ERROR -  Mobile number not changed to $($change.CellPhone)" = "[$($graphuser.userPrincipalName)]"}
    }
}
If($($change.Office_x0020_Number) -ne $($graphuser.businessPhones) -and ("no number" -ne $change.Office_x0020_Number)){
    $error.clear()
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'Change by Request' -logstring "Updating the office number for [$($graphuser.userPrincipalName)]: $($graphuser.businessPhones) to $($change.Office_x0020_Number)"
    write-host "Office number has been changed...amending 365" -ForegroundColor White
    $businessnumberhash = @{businessPhones=@(“$($change.Office_x0020_Number)”)}
    set-graphuser -tokenResponse $tokenResponse -userIdOrUpn  $graphuser.id -userPropertyHash $businessnumberhash -Verbose
    $officenumbercheck = get-graphUsers -tokenResponse $tokenResponse -filterUpns $graphuser.userPrincipalName
    #Check graph user was changed
    If($officenumbercheck.businessPhones -eq $change.Office_x0020_Number){
       #Update Directory List (reporting list only has name of employee and manager fields, so we won't update this one)
       friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(365) Updated the office number for [$($graphuser.userPrincipalName)]: $($graphuser.businessPhones) to $($change.Office_x0020_Number)"
       $officenumberdirectoryupdate = update-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $directoryListId -listitemId "$($thisanthesian.id)" -fieldHash @{"Office_x0020_Phone" = "$($change.Office_x0020_Number)"} -Verbose
       If($officenumberdirectoryupdate.Office_x0020_Phone -eq $change.Office_x0020_Number){
       friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(Directory) Updated the office number for [$($graphuser.userPrincipalName)]: $($graphuser.businessPhones) to $($change.Office_x0020_Number)"
       }
       Else{
       friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(Directory) Something went wrong updating the office number for [$($graphuser.userPrincipalName)] in the Directory"
       friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
       $TeamsReport += @{"$(get-date) (Change Request > Directory List) ERROR - Office number not changed to $change.Office_x0020_Number" = "[$($graphuser.userPrincipalName)]"}
       }
    }
    Else{
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(365) Something went wrong updating the office number for [$($graphuser.userPrincipalName)] in 365"
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
    $TeamsReport += @{"$(get-date) (Change Request > 365/Directory List) ERROR -  Office number not changed to $change.Office_x0020_Number" = "[$($graphuser.userPrincipalName)]"}
    }
}
If($($change.Community) -ne $($graphuser.department) -and "Select one" -ne ($change.Community) -and "Not Applicable" -ne ($change.Community)){
    $error.clear()
    write-host "Community has been changed...amending 365" -ForegroundColor White
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'Change by Request' -logstring "Updating the community for [$($graphuser.userPrincipalName)]: $($graphuser.department) to $($change.Community)"
    set-graphuser -tokenResponse $tokenResponse -userIdOrUpn  $graphuser.id -userPropertyHash @{"department" = $($change.Community)}
    $communitycheck = get-graphUsers -tokenResponse $tokenResponse -filterUpns $graphuser.userPrincipalName -selectAllProperties:$true
    #Check graph user was changed
    If($communitycheck.department -eq $change.Community){
       #Update Directory List (reporting list only has name of employee and manager fields, so we won't update this one)
       friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(365) Updated the community for [$($graphuser.userPrincipalName)]: $($graphuser.department) to $($change.Community)"
       $communitydirectoryupdate = update-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $directoryListId -listitemId "$($thisanthesian.id)" -fieldHash @{"Community" = "$($change.Community)"} -Verbose
       If($communitydirectoryupdate.Community -eq $change.Community){
       friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(Directory) Updated the community for [$($graphuser.userPrincipalName)]: $($graphuser.department) to $($change.Community)"
       }
       Else{
       friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(Directory) Something went wrong updating the community for [$($graphuser.userPrincipalName)] in the Directory"
       friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
       $TeamsReport += @{"$(get-date) (Change Request > Directory List) ERROR - Community not changed to $($change.Community)" = "[$($graphuser.userPrincipalName)]"}
        }
    }
    Else{
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(365) Something went wrong updating the community for [$($graphuser.userPrincipalName)] in 365"
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
    $TeamsReport += @{"$(get-date) (Change Request > 365/Directory List) ERROR - Community not changed to $($change.Community)" = "[$($graphuser.userPrincipalName)]"}
    }
}
If($($change.Office) -ne $($graphuser.officeLocation) -and "Select one" -ne ($change.Office)){
$error.clear()
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'Change by Request' -logstring "Updating the office for [$($graphuser.userPrincipalName)]: $($graphuser.officeLocation) to $($change.Office)"
write-host "Office has been changed...amending 365" -ForegroundColor White
Connect-PnPOnline "https://anthesisllc-admin.sharepoint.com/" -Credentials $spoCreds
$officeterm = Get-PnPTerm -TermSet "Offices" -TermGroup "Anthesis" -Includes CustomProperties | Where-Object -Property "Name" -EQ $change.Office
If($officeterm){
    #Set graphuser office
    set-graphuser -tokenResponse $tokenResponse -userIdOrUpn $graphuser.userPrincipalName -userPropertyHash  @{"officeLocation" = $($officeterm.Name)}
    set-graphuser -tokenResponse $tokenResponse -userIdOrUpn $graphuser.userPrincipalName -userPropertyHash  @{"streetAddress" = $($officeterm.CustomProperties.'Street Address')}
    set-graphuser -tokenResponse $tokenResponse -userIdOrUpn $graphuser.userPrincipalName -userPropertyHash  @{"postalCode" = $($officeterm.CustomProperties.'Postal Code')}
    set-graphuser -tokenResponse $tokenResponse -userIdOrUpn $graphuser.userPrincipalName -userPropertyHash  @{"country" = $($officeterm.CustomProperties.'Country')}
    #Set SPO timezone
    Set-SPOTimezone -upn $graphuser.userPrincipalName -office $officeterm.Name
    #Set EXO timezone
    $mailboxupdate = set-graphMailboxSettings -tokenResponse $tokenResponse -identity $graphuser.userPrincipalName -timeZone "$($officeterm.CustomProperties.Timezone)"
    
    #Check timezones after waiting 10 seconds for changes to sync up and run checks
    Start-Sleep -Seconds 10
        $timezonesync = @()
        #Mailbox timezone check
        If(($mailboxupdate.timeZone -eq $officeterm.CustomProperties.Timezone) -and ($($mailboxupdate.workingHours.timeZone).name -eq $officeterm.CustomProperties.Timezone)){
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(365) Mailbox timezone matches the Office Location for [$($graphuser.userPrincipalName)]: $($mailboxupdate.timeZone)"
        $timezonesync += 0
        }
        Else{
        $currenttimezone = get-graphMailboxSettings -tokenResponse $tokenResponse -identity $graphuser.userPrincipalName
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(365) Something went wrong updating the Exchange Online Mailbox timezone for [$($graphuser.userPrincipalName)]"
        $TeamsReport += @{"$(get-date) (Change Request > 365) ERROR - [Part of Office Location change]  Mailbox timezone not changed to $($officeterm.CustomProperties.Timezone). Current mailbox timezone is: $($currenttimezone.timeZone)" = "[$($graphuser.userPrincipalName)]"}
        $timezonesync += 1
        }
        #sharepoint timezone check
        $spoTimezonecheck =  Get-PnPUserProfileProperty -Account $($graphuser.userPrincipalName)
        #PNP pulls back a version of the timezone name from the SPO-userobject not used elsewhere so we have to process the UTC as we can't use the ID
        $allsptimezones = Get-PnPTimeZoneId
        #lazy match for first parenthesis (in all timezone strings) - and then proceed through the 'battenburg problem'
        $formattedtimezone = ($spoTimezonecheck.UserProfileProperties.'SPS-TimeZone')    
        $regex = [regex]"\(.+?\)"
        $utcvalue = [regex]::match($formattedtimezone, $regex).Groups[0]
        $utcvalue = [string]($utcvalue.Value)
        $utcvalue = $utcvalue.Split('\(')[1]
        $utcvalue = $utcvalue.Split('\)')[0]
        #Get the timezone we tried to set
        $attemptedtimezone = $allsptimezones | Where-Object -Property "Id" -EQ $($officeterm.CustomProperties.'Sharepoint Timezone ID')
        #Compare both on the UTC value - closest estimation we can get to see if it's correct to what we tried to change it to
        $timezonecheck = Compare-Object -ReferenceObject $utcvalue -DifferenceObject $attemptedtimezone.Identifier -IncludeEqual
        If($timezonecheck.SideIndicator -eq "=="){
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(365) Sharepoint timezone matches the Office Location for [$($graphuser.userPrincipalName)]: $($attemptedtimezone.Identifier)"
        $timezonesync += 0            
        }
        Else{
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(365) Updating the Sharepoint timezone for [$($graphuser.userPrincipalName)] has gone wrong, the UTC value does not match what we tried to set it to ($($attemptedtimezone.Identifier))"
        $TeamsReport += @{"$(get-date) (Change Request > 365) ERROR - [Part of Office Location change]  Sharepoint timezone not changed to $($officeterm.CustomProperties.Timezone). Current mailbox timezone is: $($currentexoTimezone.Id)" = "[$($graphuser.userPrincipalName)]"}
        $timezonesync += 1
        }                 
    If($timezonesync -eq 0){
    #If all returns okay, update the Directory with the friendly utc timezone 
    $exoTimezone = get-graphMailboxSettings -tokenResponse $tokenResponse -identity "$($graphuser.userPrincipalName)" -Verbose
    #Philippine's uses several timezone names
    If(($exoTimezone.timeZone -eq "Singapore Standard Time") -or ($exoTimezone.timeZone -eq "Taipei Standard Time") -or ($exoTimezone.timeZone -eq "China Standard Time")){
    $exoTimezone = New-Object -TypeName psobject @{
    "DisplayName" = "(UTC+08:00) $($graphuser.country) Standard Time"
    }
    }
    Else{
    $exoTimezone = Get-TimeZone $exoTimezone.timeZone
    }    
    $officedirectoryupdate = update-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $directoryListId -listitemId "$($thisanthesian.id)" -fieldHash @{"Office" = "$($change.Office)";"Timezone" = "$($exoTimezone.DisplayName)";"Country" = "$($officeterm.CustomProperties.Country)"} -Verbose
        If($officedirectoryupdate.office -eq $change.Office){
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "Updated the office for [$($graphuser.userPrincipalName)] in the Directory: $($graphuser.officeLocation) to $($change.Office)"
        }
        Else{
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(Directory) Something went wrong updating the Office for [$($graphuser.userPrincipalName)]: $($graphuser.officeLocation) to $($change.Office)"
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
        $TeamsReport += @{"$(get-date) (Change Request > Directory List) ERROR - Office not changed to $($change.Office)" = "[$($graphuser.userPrincipalName)]"}
        }
    }
}
Else{
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(365) Something went wrong updating the office for [$($graphuser.userPrincipalName)], we couldn't retrieve the Term for the office from Sharepoint: $($change.Office)"
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
$TeamsReport += @{"$(get-date) (Change Request > 365/Directory List) ERROR - Office not changed to $($change.Office), Term couldn't be retrieved from the Term Store" = "[$($graphuser.userPrincipalName)]"}
}
}
$graphQuery = "/users/$($graphuser.id)/manager"
$graphmanager = invoke-graphGet -tokenResponse $tokenResponse -graphQuery $graphQuery -ErrorAction SilentlyContinue
$graphmanager = $graphmanager.userPrincipalName
If($($change.ManagerEmail) -ne ($($graphmanager) -or !$graphmanager) -and "groupbot@anthesisgroup.com" -ne ($change.ManagerEmail)){
    $error.clear()
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'Change by Request' -logstring "Updating the manager for [$($graphuser.userPrincipalName)]: $($graphmanager.userPrincipalName) to $($change.ManagerEmail)"
    Write-Host "Manager has been changed from $($graphManager) to $($change.ManagerEmail)...amending" -ForegroundColor Yellow
    set-graphuserManager -tokenResponse $tokenResponse -userUPN $($graphuser.userPrincipalName) -managerUPN $($change.ManagerEmail) -Verbose
    #Check it was set correctly
    $graphmanagercheck = invoke-graphGet -tokenResponse $tokenResponse -graphQuery $graphQuery -ErrorAction SilentlyContinue
    If($($change.ManagerEmail) -eq $($graphmanagercheck.userPrincipalName)){
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(365) Updated the manager for [$($graphuser.userPrincipalName)]: $($graphmanager.userPrincipalName) to $($change.ManagerEmail)"
    #If it was set correctly, update the Directory and Reporting list
    $spoManager = $allanthesians  | Where-Object {$_.fields.Email -eq "$($change.ManagerEmail)"}
    $managerdirectoryupdate = update-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $directoryListId -listitemId "$($thisanthesian.id)" -fieldHash @{"ManagerLookupId" = "$($spomanager.fields.AnthesianLookupId)";"ManagerEmail" = $($spomanager.fields.Email)} -Verbose
        If($managerdirectoryupdate.ManagerEmail -eq $change.ManagerEmail){
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(Directory) Updated the manager for [$($graphuser.userPrincipalName)]: $($graphmanager.userPrincipalName) to $($change.ManagerEmail)"
        #If the Directory update, go ahead and update the reporting list, just to keep them nice and in sync
        $managerreportingupdate = update-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $reportinglinesListId -listitemId "$($thisreport.id)" -fieldHash @{"ManagerLookupId" = "$($spomanager.fields.AnthesianLookupId)";"ManagerEmail" = $($spomanager.fields.Email)} -Verbose
            If($managerreportingupdate.ManagerEmail -eq $change.ManagerEmail){
            friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(Reporting) Updated the manager for [$($graphuser.userPrincipalName)]: $($graphmanager.userPrincipalName) to $($change.ManagerEmail)"
            }
            Else{
            friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(Reporting) Something went wrong updating the manager for [$($graphuser.userPrincipalName)] in the Reporting List: $($graphmanager) to $($change.ManagerEmail)"
            friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
            $TeamsReport += @{"$(get-date) (Change Request > Reporting List) ERROR - Manager not changed to $($change.ManagerEmail)" = "[$($graphuser.userPrincipalName)]"}
            }
        }
        Else{
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(Directory) Something went wrong updating the manager for [$($graphuser.userPrincipalName)] in the Directory: $($graphmanager) to $($change.ManagerEmail)"
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
        $TeamsReport += @{"$(get-date) (Change Request > Directory List) ERROR - Manager not changed to $($change.ManagerEmail)" = "[$($graphuser.userPrincipalName)]"}
        }   
    }
    Else{
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(365) Something went wrong updating the manager for [$($graphuser.userPrincipalName)]: $($graphmanager) to $($change.ManagerEmail)"
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
    $TeamsReport += @{"$(get-date) (Change Request > 365/Directory List/Reporting List) ERROR - Manager not changed to $($change.ManagerEmail)" = "[$($graphuser.userPrincipalName)]"}
    }
}
If($($change.City) -ne $($graphuser.city) -and ("Select one" -ne $change.City)){
    $error.clear()
    write-host "City has been changed...amending 365" -ForegroundColor White
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'Change by Request' -logstring "Updating the city for [$($graphuser.userPrincipalName)]: $($graphuser.city) to $($change.City)"
    set-graphuser -tokenResponse $tokenResponse -userIdOrUpn  $graphuser.id -userPropertyHash @{"city" = $($change.City)}
    $citycheck = get-graphUsers -tokenResponse $tokenResponse -filterUpns $graphuser.userPrincipalName -selectAllProperties:$true
    #Check graph user was changed
    If($citycheck.city -eq $change.City){
        #Update Directory List (reporting list only has name of employee and manager fields, so we won't update this one)
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(365) Updated the city for [$($graphuser.userPrincipalName)]: $($graphuser.city) to $($change.City)"
        $citydirectoryupdate = update-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $directoryListId -listitemId "$($thisanthesian.id)" -fieldHash @{"City" = "$($change.City)"} -Verbose
        If($citydirectoryupdate.City -eq $change.City){
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(Directory) Updated the city for [$($graphuser.userPrincipalName)]: $($graphuser.city) to $($change.City)"
        }
        Else{
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(Directory) Something went wrong updating the city for [$($graphuser.userPrincipalName)] in the Directory"
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
        $TeamsReport += @{"$(get-date) (Change Request > Directory List) ERROR - City not changed to $($change.City)" = "[$($graphuser.userPrincipalName)]"}
        }
    }
    Else{
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(365) Something went wrong updating the city for [$($graphuser.userPrincipalName)]"
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
    $TeamsReport += @{"$(get-date) (Change Request > 365/Directory List) ERROR - City not changed to $($change.City)" = "[$($graphuser.userPrincipalName)]"}
    }
}
If($($change.Business_x0020_Unit) -ne $($graphuser.anthesisgroup_employeeInfo.businessUnit) -and ("Select one" -ne $change.Business_x0020_Unit)){
    $error.clear()
    write-host "Business Unit has been changed...amending 365" -ForegroundColor White
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'Change by Request' -logstring "Updating the Business Unit for [$($graphuser.userPrincipalName)]: $($graphuser.anthesisgroup_employeeInfo.businessUnit) to $($change.Business_x0020_Unit)"
    set-graphuser -tokenResponse $tokenResponse -userIdOrUpn $graphuser.id -userEmployeeInfoExtensionHash @{"businessUnit" = $($change.Business_x0020_Unit)}
    $businessunitcheck = get-graphUsers -tokenResponse $tokenResponse -filterUpns $graphuser.userPrincipalName -selectAllProperties:$true
    #Check graph user was changed
    If($businessunitcheck.anthesisgroup_employeeInfo.businessUnit -eq $change.Business_x0020_Unit){
        #Update Directory List (reporting list only has name of employee and manager fields, so we won't update this one)
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(365) Updated the business unit for [$($graphuser.userPrincipalName)]: $($graphuser.anthesisgroup_employeeInfo.businessUnit) to $($change.Business_x0020_Unit)"
        $businessunitdirectoryupdate = update-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $directoryListId -listitemId "$($thisanthesian.id)" -fieldHash @{"BusinessUnit" = "$($change.Business_x0020_Unit)"} -Verbose
        If($businessunitdirectoryupdate.BusinessUnit -eq $change.Business_x0020_Unit){
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(Directory) Updated the business unit for [$($graphuser.userPrincipalName)]: $($graphuser.anthesisgroup_employeeInfo.businessUnit) to $($change.Business_x0020_Unit)"
        }
        Else{
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(Directory) Something went wrong updating the business unit for [$($graphuser.userPrincipalName)] in the Directory"
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
        $TeamsReport += @{"$(get-date) (Change Request > Directory List) ERROR - Business Unit not changed to $($change.Business_x0020_Unit)" = "[$($graphuser.userPrincipalName)]"}
        }
    }
    Else{
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(365) Something went wrong updating the business unit for [$($graphuser.userPrincipalName)] in 365"
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
    $TeamsReport += @{"$(get-date) (Change Request > 365/Directory List) ERROR - Business Unit not changed to $($change.Business_x0020_Unit)" = "[$($graphuser.userPrincipalName)]"}
    }
}

#If there are ANY errors, don't update the change request so we have another chance to spot issues in the chain
If(!$TeamsReport){update-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $changeListId -listitemId "$($change.id)" -fieldHash @{"Status" = "Complete"} -Verbose}
Else{update-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $changeListId -listitemId "$($change.id)" -fieldHash @{"Status" = "Issue during update - IT have been notified"} -Verbose}

}

<#------------------------------------------------------------------------------------------Teams Report-----------------------------------------------------------------------------------------------------#>

If($TeamsReport){
    $report = @()
    $report += "***************Errors found in 365/Directory/Reporting List Sync***************" + "<br><br>"
    $report += "*******************************Change Request Side*****************************" + "<br><br>"
        ForEach($t in $TeamsReport){
        $report += "$($t.Keys)" + " - " + "$($t.Values)" + "<br><br>"
}

$report = $report | out-string

Send-MailMessage -To "cb1d8222.anthesisgroup.com@amer.teams.ms" -From "PeopleServicesRobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "test" -BodyAsHtml $report -Encoding UTF8 -Credential $exocreds
}


#Finish run
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype END -logstring "End of run for sync-DirectoryChangeRequests"
