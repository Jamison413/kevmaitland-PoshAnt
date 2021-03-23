#NOTE
#This is the first half of the Anthesis Directory Sync Scripts. This handles most of the heavy listing around adding, removing and amending entries in the Anthesis Direcotry and Reporting Lists via changes in 365.

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

$Admin = "kimblebot@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
$AdminPass = ConvertTo-SecureString (Get-Content "$env:USERPROFILE\Desktop\KimbleBot.txt") 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin, $AdminPass

$exoCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin, $AdminPass
connect-ToExo -credential $exoCreds


#Conn - CSOM for SharepointUserID
<#
$sharePointAdmin = "kimblebot@anthesisgroup.com"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\kimblebot.txt) 
$spoCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
#>
#$sharePointAdmin = "emily.pressey@anthesisgroup.com"
#$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\Emily.txt) 
$sharePointAdmin = "kimblebot@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content "$env:USERPROFILE\Desktop\KimbleBot.txt") 


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

friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype START -logstring "Starting run for sync-Directory365Changes"

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

<#-----------------------------------------------------------------------------Add/Remove members from overall list-----------------------------------------------------------------------------#>

<#Add new members#>
$newanthesians = Compare-Object -ReferenceObject $allgraphusers.Id -DifferenceObject $allanthesianGUIDS.UserGUID | where-object -Property "SideIndicator" -EQ "<="
If($newanthesians){
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype MESSAGE -logstring "We've found some Anthesians to add to the Directory [$(($newanthesians| Measure-Object).Count)]"
}
Else{
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype MESSAGE -logstring "We've found no new Anthesians to add to the Directory"
}
$fullUsers = @()
#Process any new Anthesians
ForEach($newanthesian in $newanthesians){

$graphuser = $allgraphusers | Where-Object -Property "id" -EQ $newanthesian.InputObject
Write-Host "New Anthesian: Adding $($graphuser.userPrincipalName) to Anthesis People Directory" -ForegroundColor Yellow

#Get info across objects
Write-Host "Processing: $($graphuser.userPrincipalName)" -ForegroundColor White
#SPO - for Sharepoint lookup ID to input data into people picker columns
$spomanager = $null
$context = $ctx
#Fetch the users in Site Collection
$sharepointUsers = $context.Web.SiteUsers
$context.Load($sharepointUsers)
$context.ExecuteQuery()
$spoUser = $context.web.EnsureUser("i:0#.f|membership|$($graphuser.userPrincipalName)")
$context.Load($spoUser)
$context.ExecuteQuery()

#Graph and SPO manager
$graphQuery = "/users/$($graphuser.id)/manager"
$graphmanager = invoke-graphGet -tokenResponse $tokenResponse -graphQuery $graphQuery -ErrorAction SilentlyContinue
    If($graphmanager){
        $context = $ctx
        $sharepointUsers = $context.Web.SiteUsers
        $context.Load($sharepointUsers)
        $context.ExecuteQuery()
        $spomanager = $context.web.EnsureUser("i:0#.f|membership|$($graphmanager.userPrincipalName)")
        $context.Load($spomanager)
        $context.ExecuteQuery()
        }
    Else{
    Write-Host "WARNING: No line manager for $($graphuser.userPrincipalName)" -ForegroundColor White
    }

#Exchange timezone
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


#sharepoint timezone
$spoTimezone =  Get-PnPUserProfileProperty -Account $($graphuser.userPrincipalName)

$antUser = New-Object psobject -Property @{
        graphuser = $graphuser
        spouser = $spoUser
        linemanager = $spomanager
        exoTimezone = $exoTimezone
        spoTimezone = $spoTimezone
        teamslink = "https://teams.microsoft.com/l/chat/0/0?users=" + "$($graphuser.userPrincipalName)"
}
Write-Host "Adding: $($graphuser.displayName)" -ForegroundColor Yellow
$fullUsers += $antUser
}
#Add new entry to the People Directory first
ForEach($user in $fullUsers){
Write-Host "Adding $($user.graphuser.userPrincipalName) to the Directory" -ForegroundColor Yellow
If($user.linemanager.Id){
#Line manager
$body = "{
  `"fields`": {
    `"AnthesianLookupId`": `"$($user.spouser.Id)`",
    `"JobTitle`": `"$($user.graphuser.jobTitle)`",
    `"Community`": `"$($user.graphuser.department)`",
    `"Office_x0020_Phone`": `"$($user.graphuser.businessPhones)`",
    `"CellPhone`": `"$($user.graphuser.mobilePhone)`",
    `"City`": `"$($user.graphuser.city)`",
    `"Office`": `"$($user.graphuser.officeLocation)`",
    `"Country`": `"$($user.graphuser.country)`",
    `"Timezone`": `"$($user.exotimezone.DisplayName)`",
    `"ManagerLookupId`": `"$($user.linemanager.Id)`",
    `"ManagerEmail`": `"$($user.linemanager.Email)`",
    `"BusinessUnit`": `"$($user.graphuser.anthesisgroup_employeeInfo.businessUnit)`",
    `"Email`": `"$($user.graphuser.userPrincipalName)`",
    `"TeamsLink`": `"$($user.teamslink)`",
    `"UserGUID`": `"$($user.graphuser.id)`",
    `"plaintextname`": `"$($user.graphuser.displayName)`",
    `"Contract`": `"$($user.graphuser.anthesisgroup_employeeInfo.contractType)`",
    `"TenureDays`": `"0`"

  }
}"
}
Else{
#No line manager
$body = "{
  `"fields`": {
    `"AnthesianLookupId`": `"$($user.spouser.Id)`",
    `"JobTitle`": `"$($user.graphuser.jobTitle)`",
    `"Community`": `"$($user.graphuser.department)`",
    `"Office_x0020_Phone`": `"$($user.graphuser.businessPhones)`",
    `"CellPhone`": `"$($user.graphuser.mobilePhone)`",
    `"City`": `"$($user.graphuser.city)`",
    `"Office`": `"$($user.graphuser.officeLocation)`",
    `"Country`": `"$($user.graphuser.country)`",
    `"Timezone`": `"$($user.exotimezone.DisplayName)`",
    `"BusinessUnit`": `"$($user.graphuser.anthesisgroup_employeeInfo.businessUnit)`",
    `"Email`": `"$($user.graphuser.userPrincipalName)`",
    `"TeamsLink`": `"$($user.teamslink)`",
    `"UserGUID`": `"$($user.graphuser.id)`",
    `"plaintextname`": `"$($user.graphuser.displayName)`",
    `"Contract`": `"$($user.graphuser.anthesisgroup_employeeInfo.contractType)`",
    `"TenureDays`": `"0`"
  }
}"
}
$graphQuery = "https://graph.microsoft.com/v1.0/sites/$($graphSiteId)/lists/$($directoryListId)/items"
$body = [System.Text.Encoding]::UTF8.GetBytes($body)
$directoryresponse = Invoke-RestMethod -Uri "$graphQuery" -Body $body -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Post -verbose
If($directoryresponse.fields.Email -eq $user.graphuser.userPrincipalName){
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "user added to Directory: [$($user.graphuser.userPrincipalName)]"
}
Else{
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "user not added to Directory: [$($user.graphuser.userPrincipalName)]"
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($error[0])"
$TeamsReport += @{"$(get-date) (365 > Directory) ERROR - New User not added" = "[$($user.graphuser.userPrincipalName)]"}
}

#Add new entry to the POP Reporting List second
Write-Host "Adding $($user.graphuser.userPrincipalName) to the Reporting list" -ForegroundColor Yellow
If($user.linemanager.Id){
#Line manager
$body = "{
  `"fields`": {
    `"AnthesianLookupId`": `"$($user.spouser.Id)`",
    `"ManagerLookupId`": `"$($user.linemanager.Id)`",
    `"ManagerEmail`": `"$($user.linemanager.Email)`",
    `"Email`": `"$($user.graphuser.userPrincipalName)`",
    `"plaintextname`": `"$($user.graphuser.displayName)`",
    `"UserGUID`": `"$($user.graphuser.id)`"
  }
}"
}
Else{
#No line manager
$body = "{
  `"fields`": {
    `"AnthesianLookupId`": `"$($user.spouser.Id)`",
    `"Email`": `"$($user.graphuser.userPrincipalName)`",
    `"plaintextname`": `"$($user.graphuser.displayName)`",
    `"plaintextname`": `"$($user.graphuser.displayName)`",
    `"UserGUID`": `"$($user.graphuser.id)`",
  }
}"
}
$graphQuery = "https://graph.microsoft.com/v1.0/sites/$($graphSiteId)/lists/$($reportinglinesListId)/items"
$body = [System.Text.Encoding]::UTF8.GetBytes($body)
$reportresponse = Invoke-RestMethod -Uri "$graphQuery" -Body $body -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Post -verbose
If($reportresponse.fields.Email -eq $user.graphuser.userPrincipalName){
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "user added to Reporting List: [$($user.graphuser.userPrincipalName)]"
}
Else{
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "user not added to Reporting List: [$($user.graphuser.userPrincipalName)]"
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($error[0])"
$TeamsReport += @{"$(get-date) (365 > Reporting List) ERROR - New User not added" = "[$($user.graphuser.userPrincipalName)]"}
}
}


<#Remove members#>
$removedanthesians = Compare-Object -ReferenceObject $allgraphusers.Id -DifferenceObject $allanthesianGUIDS.UserGUID | where-object -Property "SideIndicator" -EQ "=>"
If($removedanthesians){
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype MESSAGE -logstring "We've found some Anthesians to remove from the Directory [$(($removedanthesians | Measure-Object).Count)]"
}
Else{
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype MESSAGE -logstring "We've found no Anthesians to remove from the Directory"
}
ForEach($removedanthesian in $removedanthesians){
    $anthesiantoremove = $allanthesianGUIDS | Where-Object -Property "UserGUID" -EQ "$($removedanthesian.InputObject)"#needs testing
    $thisreporttoremove = $allPOPreports | Where-Object {$_.fields.UserGUID -eq "$($removedanthesian.InputObject)"}
    Write-Host "Removed Anthesian: Removing $($anthesiantoremove.Email) from Anthesis People Directory" -ForegroundColor Yellow
    $graphItemId = $anthesiantoremove.id
    delete-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -graphListId $directoryListId -graphItemId $graphItemId
    $graphItemId = $thisreporttoremove.id
    delete-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -graphListId $reportinglinesListId -graphItemId $graphItemId

    #Check if it was deleted, we can't capture any details of the response for the graph api weirdly. We try to query the list and if it was removed we should recieve an error code saying it can't find the item
    $error.clear()
    $graphItemId = $anthesiantoremove.id
    get-graphListItems -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $directoryListId -filterId $graphItemId
    $removeitemcheck = If($($error[0]) -match '"code": "itemNotFound"'){1}
    $error.clear()
    $graphItemId = $thisreporttoremove.id
    get-graphListItems -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $reportinglinesListId -filterId $graphItemId
    If($($error[0]) -match '"code": "itemNotFound"'){$removeitemcheck += 2} #########///////////this section doesn't work because null is also -notcontains############
    $error.clear()
        If($removeitemcheck -eq 3){
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "Removed from Lists: [$($anthesiantoremove.Email)]"
        }
        If($removeitemcheck -eq 2){
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(Directory) Not Removed From Directory: [$($anthesiantoremove.Email)]"
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "Item Not Found: [$($anthesiantoremove.Email)]"
        $TeamsReport += @{"$(get-date) (365 > Directory List) ERROR - User not removed from Directory" = "$($anthesiantoremove.Email)"}
        }
        If($removeitemcheck -eq 1){
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "Not Removed From Reporting List: [$($anthesiantoremove.Email)]"
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "Item Not Found: [$($anthesiantoremove.Email)]"
        $TeamsReport += @{"$(get-date) (365 > Reporting List) ERROR - User not removed from Reporting List" = "$($anthesiantoremove.Email)"}
        }
}

If($TeamsReport){
    $report = @()
    $report += "***************Errors found in 365/Directory/Reporting List Sync***************" + "<br><br>"
    $report += "*******************************365 add/remove users****************************************" + "<br><br>"
        ForEach($t in $TeamsReport){
        $report += "$($t.Keys)" + " - " + "$($t.Values)" + "<br><br>"
}

$report = $report | out-string

Send-MailMessage -To "cb1d8222.anthesisgroup.com@amer.teams.ms" -From "PeopleServicesRobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "test" -BodyAsHtml $report -Encoding UTF8 -Credential $exocreds
}



<#------------------------------------------------------------------------------------Process Changes by 365 amend--------------------------------------------------------------------------------------------#>

ForEach($graphuser in $allgraphusers){
$TeamsReport = @()
#Find the list entries for the staff list and POP Reporting list

#Directory List
$thisanthesian = $allanthesians | Where-Object {$_.fields.Email -eq "$($graphuser.userPrincipalName)"}
#Reporting List
$thisreport = $allPOPreports | Where-Object {$_.fields.Email -eq "$($graphuser.userPrincipalName)"}

If($($thisanthesian.fields.JobTitle) -ne $($graphuser.jobTitle)){
    $error.clear()
    Write-Host "Job title has been changed in 365 for $($graphuser.userPrincipalName)...amending lists" -ForegroundColor Yellow
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'Change in 365' -logstring "user job title changed in 365: [$($graphuser.userPrincipalName)]"
    $jobtitledirectoryupdate = update-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $directoryListId -listitemId "$($thisanthesian.id)" -fieldHash @{"JobTitle" = "$($graphuser.JobTitle)"} -Verbose
    If($jobtitledirectoryupdate.JobTitle -eq $graphuser.jobTitle){
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(Directory) Updated job title for [$($graphuser.userPrincipalName)]: $($thisanthesian.fields.JobTitle) to $($graphuser.JobTitle)"
    }
    Else{
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(Directory) Something went wrong updating the job title for [$($graphuser.userPrincipalName)] in the Directory"
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
    $TeamsReport += @{"$(get-date) (Change in 365 > Directory List) ERROR - Job Title not changed to $($graphuser.jobTitle)" = "[$($graphuser.userPrincipalName)]"}
    }
}
If($($thisanthesian.fields.cellphone) -ne $($graphuser.mobilePhone)){
    $error.clear()
    Write-Host "Mobile number has been changed in 365 for $($graphuser.userPrincipalName)...amending lists" -ForegroundColor Yellow
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'Change in 365' -logstring "user mobile number changed in 365: [$($graphuser.userPrincipalName)]"
    $mobilenumberdirectoryupdate = update-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $directoryListId -listitemId "$($thisanthesian.id)" -fieldHash @{"CellPhone" = "$($graphuser.mobilePhone)"} -Verbose
    If($mobilenumberdirectoryupdate.CellPhone -eq $graphuser.mobilePhone){
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(Directory) Updated mobile number for [$($graphuser.userPrincipalName)]: $($thisanthesian.fields.CellPhone) to $($graphuser.mobilePhone)"
    }
    Else{
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(Directory) Something went wrong updating the mobile number for [$($graphuser.userPrincipalName)] in the Directory"
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
    $TeamsReport += @{"$(get-date) (Change in 365 > Directory List) ERROR - Mobile Number not changed to $($graphuser.mobilePhone)" = "[$($graphuser.userPrincipalName)]"}
    }
}
If($($thisanthesian.fields.Office_x0020_Phone) -ne $($graphuser.businessPhones)){
    $error.clear()
    Write-Host "Office number has been changed in 365 for $($graphuser.userPrincipalName)...amending lists" -ForegroundColor Yellow
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'Change in 365' -logstring "user office number changed in 365: [$($graphuser.userPrincipalName)]"
    $officenumberdirectoryupdate = update-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $directoryListId -listitemId "$($thisanthesian.id)" -fieldHash @{"Office_x0020_Phone" = "$($graphuser.businessPhones)"} -Verbose
    If($officenumberdirectoryupdate.Office_x0020_Phone -eq $graphuser.businessPhones){
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(Directory) Updated office number for [$($graphuser.userPrincipalName)]: $($thisanthesian.fields.Office_x0020_Phone) to $($graphuser.businessPhones)"
    }
    Else{
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(Directory) Something went wrong updating the office number for [$($graphuser.userPrincipalName)] in the Directory"
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
    $TeamsReport += @{"$(get-date) (Change in 365 > Directory List) ERROR - Office Number not changed to $($graphuser.businessPhones)" = "[$($graphuser.userPrincipalName)]"}
    }  
}
If($($thisanthesian.fields.Community) -ne $($graphuser.department)){
    $error.clear()
    Write-Host "Community has been changed in 365 for $($graphuser.userPrincipalName)...amending lists" -ForegroundColor Yellow
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'Change in 365' -logstring "user community changed in 365: [$($graphuser.userPrincipalName)]"
    $communitydirectoryupdate = update-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $directoryListId -listitemId "$($thisanthesian.id)" -fieldHash @{"Community" = "$($graphuser.department)"} -Verbose
    If($communitydirectoryupdate.Community -eq $graphuser.department){
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(Directory) Updated community for [$($graphuser.userPrincipalName)]: $($thisanthesian.fields.Community) to $($graphuser.department)"
    }
    Else{
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(Directory) Something went wrong updating the community for [$($graphuser.userPrincipalName)] in the Directory"
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
    $TeamsReport += @{"$(get-date) (Change in 365 > Directory List) ERROR - Community not changed to $($graphuser.department)" = "[$($graphuser.userPrincipalName)]"}
    }
}
If($($thisanthesian.fields.Office) -ne $($graphuser.officeLocation)){
$error.clear()
Write-Host "Office has been changed in 365 for $($graphuser.userPrincipalName)...amending lists" -ForegroundColor Yellow
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'Change in 365' -logstring "Updating the office for [$($graphuser.userPrincipalName)]: [$($graphuser.officeLocation)]"
Connect-PnPOnline "https://anthesisllc-admin.sharepoint.com/" -Credentials $spoCreds
$officeterm = Get-PnPTerm -TermSet "Offices" -TermGroup "Anthesis" -Includes CustomProperties | Where-Object -Property "Name" -EQ $graphuser.officeLocation
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
    $timezonesync = @()
    Start-Sleep -Seconds 10
        #Mailbox timezone check
        If(($mailboxupdate.timeZone -eq $officeterm.CustomProperties.Timezone) -and ($($mailboxupdate.workingHours.timeZone).name -eq $officeterm.CustomProperties.Timezone)){
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(365) Mailbox timezone matches the Office Location for [$($graphuser.userPrincipalName)]: $($mailboxupdate.timeZone)"
        $timezonesync += 0
        }
        Else{
        $currenttimezone = get-graphMailboxSettings -tokenResponse $tokenResponse -identity "emily.pressey@anthesisgroup.com"
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(365) Something went wrong updating the Exchange Online Mailbox timezone for [$($graphuser.userPrincipalName)]"
        $TeamsReport += @{"$(get-date) (365) ERROR - [Part of Office Location change]  Mailbox timezone not changed to $($officeterm.CustomProperties.Timezone). Current mailbox timezone is: $($currenttimezone.timeZone)" = "[$($graphuser.userPrincipalName)]"}
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
        #Only report back if there is an error
        If($timezonecheck.SideIndicator -eq "=="){
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(365) Sharepoint timezone matches the Office Location for [$($graphuser.userPrincipalName)]: $($attemptedtimezone.Identifier)"
        $timezonesync += 0            
        }
        Else{
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(365) Updating the Sharepoint timezone for [$($graphuser.userPrincipalName)] has gone wrong, the UTC value does not match what we tried to set it to ($($attemptedtimezone.Identifier))"
        $TeamsReport += @{"$(get-date) (365) ERROR - [Part of Office Location change]  Sharepoint timezone not changed to $($officeterm.CustomProperties.Timezone). Current mailbox timezone is: $($currentexoTimezone.Id)" = "[$($graphuser.userPrincipalName)]"}
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
    $officedirectoryupdate = update-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $directoryListId -listitemId "$($thisanthesian.id)" -fieldHash @{"Office" = "$($graphuser.officeLocation)";"Timezone" = "$($exoTimezone.DisplayName)";"Country" = "$($officeterm.CustomProperties.Country)"} -Verbose    
            If($officedirectoryupdate.Office -eq  $($graphuser.officeLocation)){
            friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(Directory) Updated Office for [$($graphuser.userPrincipalName)]: $($graphuser.officeLocation) to $($thisanthesian.fields.Office)"
            }
            Else{
            friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(Directory) Something went wrong updating the Office for [$($graphuser.userPrincipalName)]: $($graphuser.officeLocation)"
            friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
            $TeamsReport += @{"$(get-date) (365 > Directory List) ERROR - Office not changed to $($change.Office)" = "[$($graphuser.userPrincipalName)]"}
            }
    }
}
Else{
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(365) Something went wrong updating the Office for [$($graphuser.userPrincipalName)], we couldn't retrieve the Term for the office from Sharepoint: $($graphuser.officeLocation)"
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
$TeamsReport += @{"$(get-date) (365/Directory List) ERROR - Office not changed to  $($graphuser.officeLocation), Term couldn't be retrieved from the Term Store" = "[$($graphuser.userPrincipalName)]"}
}
}

$error.clear()
$graphQuery = "/users/$($graphuser.id)/manager"
$graphmanager = ""
$graphmanager = invoke-graphGet -tokenResponse $tokenResponse -graphQuery $graphQuery -ErrorAction SilentlyContinue
#We only want to amend the directory and reporting list if there is a) a manager present in Azure AD AND b) it is different to the entries in both lists
If($graphmanager){
        $spomanager = $allanthesians  | Where-Object {$_.fields.Email -eq "$($graphmanager.userPrincipalName)"}
        If($($thisanthesian.fields.ManagerEmail) -ne $($graphmanager.userPrincipalName)){
        Write-Host "Manager has been changed in 365 for $($graphuser.userPrincipalName)...amending lists" -ForegroundColor Yellow
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'Change in 365' -logstring "user manager changed in 365: [$($graphuser.userPrincipalName)]"
        #If it was set correctly, update the Directory and Reporting list
        $spoManager = $allanthesians  | Where-Object {$_.fields.Email -eq "$($graphmanager.userPrincipalName)"}
        $managerdirectoryupdate = update-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $directoryListId -listitemId "$($thisanthesian.id)" -fieldHash @{"ManagerLookupId" = "$($spomanager.fields.AnthesianLookupId)";"ManagerEmail" = $($spomanager.fields.Email)} -Verbose
        If($managerdirectoryupdate.ManagerEmail -eq $graphmanager.userPrincipalName){
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(Directory) Updated manager for [$($graphuser.userPrincipalName)]: $($thisanthesian.fields.ManagerEmail) to $($graphmanager.userPrincipalName)"
        #If the Directory update, go ahead and update the reporting list, just to keep them nice and in sync
        $managerreportingupdate = update-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $reportinglinesListId -listitemId "$($thisreport.id)" -fieldHash @{"ManagerLookupId" = "$($spomanager.fields.AnthesianLookupId)";"ManagerEmail" = $($spomanager.fields.Email)} -Verbose
            If($managerreportingupdate.ManagerEmail -eq $graphmanager.userPrincipalName){
            friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(Reporting) Updated the manager for [$($graphuser.userPrincipalName)]: $($graphmanager.userPrincipalName) to $($graphmanager.userPrincipalName)"
            }
            Else{
            friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(Reporting) Something went wrong updating the manager for [$($graphuser.userPrincipalName)] in the Reporting List: $($graphmanager) to $($spomanager.fields.Email)"
            friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
            $TeamsReport += @{"$(get-date) (Change in 365 > Reporting List) ERROR - Manager not changed to $($thisanthesian.fields.ManagerEmail)" = "[$($graphuser.userPrincipalName)]"}
            }
            }
        Else{
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(Directory) Something went wrong updating the manager for [$($graphuser.userPrincipalName)] in the Directory: $($graphmanager.userPrincipalName) to $($spomanager.fields.Email)"
        friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
        $TeamsReport += @{"$(get-date) (Change in 365 > Directory List) ERROR - Manager not changed to $($thisanthesian.fields.ManagerEmail)" = "[$($graphuser.userPrincipalName)]"}
        }
    }
}
If($($thisanthesian.fields.City) -ne $($graphuser.city)){
    $error.clear()
    Write-Host "City has been changed in 365 for $($graphuser.userPrincipalName)...amending lists" -ForegroundColor Yellow
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'Change in 365' -logstring "user city changed in 365: [$($graphuser.userPrincipalName)]"
    $citydirectoryupdate = update-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $directoryListId -listitemId "$($thisanthesian.id)" -fieldHash @{"City" = "$($graphuser.city)"} -Verbose
    If($citydirectoryupdate.City -eq $graphuser.city){
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(Directory) Updated city for [$($graphuser.userPrincipalName)]: $($thisanthesian.fields.City) to $($graphuser.city)"
    }
    Else{
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(Directory) Something went wrong updating the city for [$($graphuser.userPrincipalName)] in the Directory"
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
    $TeamsReport += @{"$(get-date) (Change in 365 > Directory List) ERROR - City not changed to $($graphuser.city)" = "[$($graphuser.userPrincipalName)]"}
    }
}
If($($thisanthesian.fields.BusinessUnit) -ne $($graphuser.anthesisgroup_employeeInfo.businessUnit)){
    $error.clear()
    Write-Host "Business Unit has been changed in 365 for $($graphuser.userPrincipalName)...amending lists" -ForegroundColor Yellow
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'Change in 365' -logstring "user business unit changed in 365: [$($graphuser.userPrincipalName)]"
    $businessunitdirectoryupdate = update-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $directoryListId -listitemId "$($thisanthesian.id)" -fieldHash @{"BusinessUnit" = "$($graphuser.anthesisgroup_employeeInfo.businessUnit)"} -Verbose
    If($businessunitdirectoryupdate.BusinessUnit -eq $graphuser.anthesisgroup_employeeInfo.businessUnit){
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype SUCCESS -logstring "(Directory) Updated business unit for [$($graphuser.userPrincipalName)]: $($thisanthesian.fields.BusinessUnit) to $($graphuser.anthesisgroup_employeeInfo.businessUnit)"
    }
    Else{
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "(Directory) Something went wrong updating the business unit for [$($graphuser.userPrincipalName)] in the Directory"
    friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype 'ERROR DETAILS' -logstring "$($Error[0].Exception.Message)"
    $TeamsReport += @{"$(get-date) (Change in 365 > Directory List) ERROR - Business Unit not changed to $($graphuser.anthesisgroup_employeeInfo.businessUnit)" = "[$($graphuser.userPrincipalName)]"}
    }
}
}



<#------------------------------------------------------------------------------------Something for Dupe Checking--------------------------------------------------------------------------------------------#>

#Get all current Anthesians in the list
$allanthesians = get-graphListItems -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $directoryListId -expandAllFields
$allanthesiansDetails = $allanthesians | select -ExpandProperty "fields"

$dupecheck = $allanthesiansDetails
$allanthesiansDetails = $allanthesiansDetails | sort email -Unique
$dupestoremove = Compare-Object -ReferenceObject $dupecheck.email -DifferenceObject $allanthesiansDetails.email
$dupeuniqueupns = $dupestoremove

If($dupestoremove){
    ForEach($dupe in $dupeuniqueupns){
    
        $dupecount = $dupecheck | Where-Object -property "email" -EQ  "$($dupe.InputObject)"
        $totaltoremove = ($dupecount.count) - 1
        If($dupecount.Count -gt 1){
        Write-Host "Removing $($totaltoremove) dupes for $($dupe.InputObject)" -ForegroundColor Yellow
        $removalIDs = $($dupecount | Select-Object -First $($totaltoremove)) | select -Property "ID"
            foreach($removalID in $removalIDs){
            delete-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -graphListId $directoryListId -graphItemId $removalID.id
        }
}
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "DUPLICATE REMOVED FROM DIRECTORY LIST: $($dupe.InputObject)"
$TeamsReport += @{"DUPLICATE REMOVED FROM DIRECTORY LIST:" = $($dupe.InputObject)}
}
}
Else{
Write-Host "No dupes found in Directory List" -ForegroundColor Yellow
}



$dupestoremove = @()
#Get all current Live Reporting Lines List
$allPOPreports = get-graphListItems -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $reportinglinesListId -expandAllFields
$allPOPreportsDetails = $allPOPreports | select -ExpandProperty "fields"

$dupecheck = $allPOPreportsDetails
$allPOPreportsDetails = $allPOPreportsDetails | sort email -Unique
$dupestoremove = Compare-Object -ReferenceObject $dupecheck.email -DifferenceObject $allPOPreportsDetails.email
$dupeuniqueupns = $dupestoremove 

If($dupestoremove){
    ForEach($dupe in $dupeuniqueupns){
    
        $dupecount = $dupecheck | Where-Object -property "email" -EQ  "$($dupe.InputObject)"
        $totaltoremove = ($dupecount.count) - 1
        If($dupecount.Count -gt 1){
        Write-Host "Removing $($totaltoremove) dupes for $($dupe.InputObject)" -ForegroundColor Yellow
        $removalIDs = $($dupecount | Select-Object -First $($totaltoremove)) | select -Property "ID"
            foreach($removalID in $removalIDs){
            delete-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -graphListId $reportinglinesListId -graphItemId $removalID.id
        }
}
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype ERROR -logstring "DUPLICATE REMOVED FROM REPORTING LIST: $($dupe.InputObject)"
$TeamsReport += @{"DUPLICATE REMOVED FROM REPORTING LIST:" = $($dupe.InputObject)}
}
}
Else{
Write-Host "No dupes found in POP Reporting List" -ForegroundColor Yellow
}


<#------------------------------------------------------------------------------------------Sync checking between the two lists/Cleaning Up POP Reporting List for Mismatches-----------------------------------------------------------------------------------------------------#>

#If graph is interrupted pulling graph users back, the lists can become misaligned.
$unsyncedanthesians = Compare-Object -ReferenceObject $allanthesiansDetails.UserGUID -DifferenceObject $allPOPreportsDetails.UserGUID | where-object -Property "SideIndicator" -EQ "<="
If($unsyncedanthesians){
    ForEach($unsyncedanthesian in $unsyncedanthesians){
    $idtoremove = $allanthesiansDetails | Where-Object -Property "UserGUID" -EQ "$($unsyncedanthesian.InputObject)" | Select-Object -Property "Id"
    #Remove each entry in the main Directory
    delete-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -graphListId $directoryListId -graphItemId $idtoremove.id

    #Kick the script off again via scheduled task so they will get created again without having to duplicate code
    Start-ScheduledTask -TaskName "Ant - Sync Directory 365 Changes"
    }
}


#Sometimes the system gets out of sync due to licensing (I think...not sure but the below will fix it), where deactivated users are in the POP list
$missingpops = Compare-Object -ReferenceObject $allanthesiansDetails.UserGUID -DifferenceObject $allPOPreportsDetails.UserGUID | where-object -Property "SideIndicator" -EQ "=>"
ForEach($missingpop in $missingpops){
    $popidtoremove = $allPOPreportsDetails | Where-Object -Property "UserGUID" -EQ "$($missingpop.InputObject)" | Select-Object -Property "Id"
    #Remove each entry in the pop list
    delete-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -graphListId $reportinglinesListId -graphItemId $popidtoremove.id
}


<#------------------------------------------------------------------------------------------Calculate Hire date (if empty) and Tenure-----------------------------------------------------------------------------------------------------#>

#Update Hire date on People Services list item

#This is a real pain - we can't update HireDate at the moment on the graph object, so we're searching the two new starter request lists for an entry. I've used pnp for my sanity here just to differentiate processing from the two lists we're processing with Graph 
$allanthesians = get-graphListItems -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $directoryListId -expandAllFields
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/teams/hr" -UseWebLogin #-Credentials $msolCredentials
$oldrequests = Get-PnPListItem -List "New User Requests"

Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/teams/People_Services_Team_All_365" -UseWebLogin #-Credentials $msolCredentials
$newrequests = Get-PnPListItem -List "New Starter Details"


ForEach($anthesian in $allanthesians){
    If($anthesian.fields.HireDate -eq $null){
    
    #Try to find them in the old list
    $thisFoundUser = ""
    $thisFoundUser = $oldrequests | ? {($(remove-diacritics $($_.FieldValues.Title.Trim().Replace(" ",".")+"@anthesisgroup.com"))) -eq $anthesian.fields.Email}
        If(($thisFoundUser | Measure-Object).count -eq 1){
        write-host "Updating hireDate for $($anthesian.Fields.Email)" -foregroundcolor Cyan
        update-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $directoryListId -listitemId $anthesian.id -fieldHash @{"HireDate" = $($thisFoundUser.FieldValues.StartDate | get-date -format "o")} -Verbose
        }
        Else{
        Write-Host "Not found in old user request list: $($anthesian.Fields.Email)" -ForegroundColor Red
        }

    #if not in old list, connect and pull all requests from the new list and try to find them
        If($thisFoundUser -eq $null){
        $thisFoundUser = $newrequests | ? {($(remove-diacritics $($_.FieldValues.Employee_x0020_Preferred_x0020_N.Trim().Replace(" ",".")+"@anthesisgroup.com"))) -eq $anthesian.fields.Email}
        If(($thisFoundUser | Measure-Object).count -eq 1){
        write-host "Updating hireDate for $($anthesian.Fields.Email)" -foregroundcolor Cyan
        update-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $directoryListId -listitemId $anthesian.id -fieldHash @{"HireDate" = $($thisFoundUser.FieldValues.StartDate | get-date -format "o")} -Verbose
        }
        Else{
        Write-Host "Not found in new user request list: $($anthesian.Fields.Email)" -ForegroundColor Red
        }

    }

    #If we can't find them at all - not much we can do - it will need to be updated manually
}
}

#Calculate tenure in days and update People Directory list item
ForEach($anthesian in $allanthesians){
If($anthesian.fields.HireDate){
$TenureinDays = New-TimeSpan -Start $anthesian.fields.HireDate -End $(get-date) | Select-Object -Property "Days"
$TenureinDays = $TenureinDays.Days + 1
[int]$currentListItemTenure = $anthesian.fields.TenureDays
    If($currentListItemTenure -ne $TenureinDays.Days){
    write-host "Updating tenure for $($anthesian.Fields.Email): $($TenureinDays.Days)" -foregroundcolor Cyan
    update-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $directoryListId -listitemId $anthesian.id -fieldHash @{"TenureDays" = $($TenureinDays).days} -Verbose
    }
}
}


<#------------------------------------------------------------------------------------------Teams Report for 365 amends-----------------------------------------------------------------------------------------------------#>

If($TeamsReport){
    $report = @()
    $report += "***************Errors found in 365/Directory/Reporting List Sync***************" + "<br><br>"
    $report += "*******************************365 Amends****************************************" + "<br><br>"
        ForEach($t in $TeamsReport){
        $report += "$($t.Keys)" + " - " + "$($t.Values)" + "<br><br>"
}

$report = $report | out-string

Send-MailMessage -To "cb1d8222.anthesisgroup.com@amer.teams.ms" -From "PeopleServicesRobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "test" -BodyAsHtml $report -Encoding UTF8 -Credential $exocreds
}



#Finish run
friendlyLogWrite -friendlyLogname $friendlyLogname -messagetype END -logstring "End of run for sync-Directory365Changes"




