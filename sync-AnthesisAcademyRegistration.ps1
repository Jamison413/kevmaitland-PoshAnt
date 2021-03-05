####################
#                  #
# Anthesis Academy #
#                  #
####################

$Logname = "C:\ScriptLogs" + "\sync-AnthesisAcademyRegistration$(Get-Date -Format "yyMMdd").log" #Check this location before live
Start-Transcript -Path $Logname -Append
Write-Host "Script started:" (Get-date)

Import-Module _PNP_Library_SPO
Remove-Module PnP.PowerShell
Import-Module SharePointPnPPowerShellOnline
Remove-Module SharePointPnPPowerShellOnline
import-Module PnP.PowerShell



$Admin = "kimblebot@anthesisgroup.com"
$AdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\kimblebot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin, $AdminPass

$exoCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin, $AdminPass
connect-ToExo -credential $exoCreds
Connect-AzureAD -credential $adminCreds
connect-toAAD -credential $adminCreds

#Connect to the People Services site
$pnpconnect = Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/teams/People_Services_Team_All_365/" -Credentials $adminCreds

$registrantProcessingList = "Anthesis Academy: Registrant Processing List"
$masterModuleList = "Anthesis Academy: Master Module List" 
$modulecompletelist = "Anthesis Academy: Module Completion Record"
$ITTraininglist = "Anthesis Academy: IT Training"

#Generate new module codes
$allmodules = Get-PnPListItem -List $masterModuleList
ForEach($moduleitem in $allmodules){

  If(!($moduleitem.FieldValues.ModuleCode)){
    
    #Generate code
    Write-Host "Generating new module code for $($moduleitem.FieldValues.ModuleName)" -ForegroundColor Yellow
    $generatemodulecode = "$($moduleitem.Id)" + "_" + (($moduleitem.FieldValues.Created_x0020_Date).Split("T")[0]) + "_" + (($moduleitem.FieldValues.ModuleName).Replace(" ",""))

    #If no module code, check no duplicates and add one, add to complete list also
    $completedmodules = Get-PnPListItem -List $modulecompletelist
    If(($completedmodules.where({$_.modulecode -eq $generatemodulecode}))){
    write-host "Something has gone wrong and there are duplicate Module Codes" -ForegroundColor Red
        $report = @()
        $report += "***************Errors found in Anthesis Academy Sync: Duplicate Module Codes***************" + "<br><br>"
        $report += "Errors found on this Module: $($moduleitem.fieldvalues.ModuleName). This will cause issues in Powerapps and needs to be manually resolved." + "<br><br>"       
        $report = $report | out-string
Send-MailMessage -To "8ed81bd4.anthesisgroup.com@amer.teams.ms" -From "PeopleServicesRobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Anthesis Academy Sync: Error" -BodyAsHtml $report -Encoding UTF8 -Credential $exocreds

    }
    Else{
    $11 = Set-PnPListItem -List $masterModuleList -Identity $moduleitem.Id -Values @{"ModuleCode" = $($generatemodulecode)}
    $12 = Add-PnPListItem -List $modulecompletelist -Values @{"ModuleName" = $moduleitem.FieldValues.ModuleName; "ModuleCode" = $generatemodulecode}
    }

  }
}

#There is a race condition from the Flow as it will overwrite a Powershell change to a field if it gets approved after (we can't cancel approval flows...), so we process Closed Modules first if they've reached their max registrant count > then registrants for closed modules > then anyone waiting to be processed

#close modules that have hit their max registrant count
$allmodules = Get-PnPListItem -List $masterModuleList
ForEach($module in $allmodules){

#If they've hit their max registrant count
$isMaxCount = $module.FieldValues.MaxRegistrantAmount -eq $module.FieldValues.RegistrantCount
If(($isMaxCount -eq "True") -and ($module.FieldValues.Status -eq "Sign Up Live")){
#close the module
Write-Host "$($module.FieldValues.ModuleName): Max registrant count reached...Closing sign up" -ForegroundColor Cyan
$closemodule = Set-PnPListItem -list $masterModuleList -Identity $module.Id -Values @{"Status" = "Closed"} 
}

}

#Clean up old events in IT Training (as we have no team management of this in SPO, this will save some time)
$ITTrainingEvents = Get-PnPListItem -List $ITTraininglist
ForEach($currentevent in $ITTrainingEvents){

$isitaftertheevent = New-TimeSpan -Start (get-date) -End $currentevent.FieldValues.Start_x0020_Time
    If($isitaftertheevent.Minutes -lt -5){
    Write-Host "$($currentevent.FieldValues.Training_x0020_Session): $($currentevent.FieldValues.Start_x0020_Time) - looks like this event is old and has passed, we'll remove it" -ForegroundColor Yellow
    $isitaftertheevent.Minutes
    $removepastevent = Remove-PnPListItem -List $ITTraininglist -Identity $currentevent.Id -Recycle -Force
    }

}





#Sync Anthesis Academy Registrants


#Just a note: on the Registrant processing list there are two 'trigger' columns, Processed (Powershell) and FlowProcessed (Flow).

#Clean out registrants for closed modules (might have been waiting for approval)
$allmodules = Get-PnPListItem -List $masterModuleList #re-get the modules we might have processed above so its up to date
$allregistrants = Get-PnPListItem -List $registrantProcessingList
$allnonwaitingregistrants = $allregistrants.Where({($_.FieldValues.FlowProcessed -eq "Waiting for Approval") -or ($_.FieldValues.FlowProcessed -eq "Approved - Waiting to be Processed as Registrant")})
ForEach($nonwaitingregistrant in $allnonwaitingregistrants){

#Find the corresponding module for the registrant
$thisRegistrantModule = $allmodules.where({$_.FieldValues.ModuleCode -eq $nonwaitingregistrant.FieldValues.ModuleCode})
    #If the Module is closed, reject it
    If($thisRegistrantModule.FieldValues.Status -eq "Closed"){
    Write-Host "Rejecting $($nonwaitingregistrant.FieldValues.RegistrantName.Email) for Module $($thisRegistrantModule.FieldValues.ModuleName), Module code $($thisRegistrantModule.FieldValues.ModuleCode): Module $($thisRegistrantModule.FieldValues.Status)" -ForegroundColor Yellow
    $rejectRegistrant = Set-PnPListItem -List $registrantProcessingList -Identity $nonwaitingregistrant.Id -Values @{"Processed" = "Module Closed - Cannot Process"; "FlowProcessed" = "Module Closed - Cannot Process"}

    #For those who were waiting approval or got caught waiting for processing and were addedd too late
    $body = "<HTML><BODY><p>Hi $($nonwaitingregistrant.FieldValues.RegistrantName.LookupValue),</p>
    <p>Unfortunately, we couldn't sign you up for the Anthesis Academy module <b>$($thisRegistrantModule.fieldvalues.ModuleName)</b> as we didn't receive Line Manager approval before sign up close.</p>
    <p>The module may be re-run in the future, please keep visiting the <a href='https://anthesisllc.sharepoint.com/sites/Resources-HR/SitePages/Anthesis-Academy.aspx?source=https%3a//anthesisllc.sharepoint.com/sites/Resources-HR/SitePages/Forms/ByAuthor.aspx'>Anthesis Academy page</a> to see newly listed Modules.<br><br></p>
    <p></p>
    <p>The Anthesis Academy</p>
    </BODY></HTML>"
     Send-MailMessage  -BodyAsHtml $body -Subject "Anthesis Academy: Could not finalise sign up to $($thisRegistrantModule.FieldValues.ModuleName)" -to $($newregistrant.FieldValues.RegistrantName.Email) -from "AnthesisAcademy@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Encoding UTF8    
     #Send-MailMessage  -BodyAsHtml $body -Subject "Anthesis Academy: Could not finalise sign up to $($thisRegistrantModule.FieldValues.ModuleName)" -to "emily.pressey@anthesisgroup.com" -from "AnthesisAcademy@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Encoding UTF8    
    }

}


$allnewregistrants = Get-PnPListItem -List $registrantProcessingList  -Query "<View><Query><Where><Eq><FieldRef Name='FlowProcessed'/><Value Type='Text'>Approved - Waiting to be Processed as Registrant</Value></Eq></Where></Query></View>"
#$allnewregistrants = $allnewregistrants.where({$_.FieldValues.Processed -ne "Module Full - Cannot Process"})

ForEach($newregistrant in $allnewregistrants){

#Doublecheck we aren't processing a non-approved registrant (looking at the Flow column)
If(($newregistrant.FieldValues.FlowProcessed -eq "Approval Denied") -or ($newregistrant.FieldValues.FlowProcessed -eq "Waiting for Approval")){
Write-Host "We shouldn't be processing this registrant, they are unapproved by line manager: $($newregistrant.FieldValues.RegistrantName.Email)" -ForegroundColor Red
        $report = @()
        $report += "***************Errors found in Anthesis Academy Sync: Powershell is trying to process an Unapproved Registrant***************" + "<br><br>"
        $report += "Weird - it's $($newregistrant.FieldValues.RegistrantName.Email). ID $($newregistrant.Id). This shouldn't be happening!" + "<br><br>"       
        $report = $report | out-string
Send-MailMessage -To "8ed81bd4.anthesisgroup.com@amer.teams.ms" -From "PeopleServicesRobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Anthesis Academy Sync: Error" -BodyAsHtml $report -Encoding UTF8 -Credential $exocreds
}


###Update the list item in the Anthesis Academy: Master Module List by looking up the Module Code###

#Get the Module list each time we iterate to get the most up to date registrant list and count
$alllivemodules = Get-PnPListItem -List $masterModuleList -Query "<View><Query><Where><Eq><FieldRef Name='Status'/><Value Type='Text'>Sign Up Live</Value></Eq></Where></Query></View>"

#Get the module (check we only brought back 1) and check it's live and as a last ditch attempt check that the Max registration count hasn't been reached - if so something is wrong in PowerApps
$thismodule = $alllivemodules.where({$_.FieldValues.ModuleCode -eq $newregistrant.FieldValues.ModuleCode})

If(($thismodule | Measure-Object).Count -eq 1){

#Get the current registrant list
$currentregistrants = @($thismodule.fieldvalues.RegistrantList.Email)

    #Check for count - greater than
If($currentregistrants.Count -gt $thismodule.fieldvalues.MaxRegistrantAmount){
Write-Host "Something has gone very wrong, too many people have signed up to this module ($($thismodule.fieldvalues.modulename)). Messaging Emily.)" -ForegroundColor Red
        $report = @()
        $report += "***************Errors found in Anthesis Academy Sync: Too Many People Have Signed Up For a Module***************" + "<br><br>"
        $report += "Errors found on this Module: $($thismodule.fieldvalues.ModuleName). The number of registered people has exceeded the maximum number of allowed registrants." + "<br><br>"       
        $report = $report | out-string
Send-MailMessage -To "8ed81bd4.anthesisgroup.com@amer.teams.ms" -From "PeopleServicesRobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Anthesis Academy Sync: Error" -BodyAsHtml $report -Encoding UTF8 -Credential $exocreds
}
    #Check for count - equal to
If($currentregistrants.Count -eq $thismodule.fieldvalues.MaxRegistrantAmount){
Write-Host "Something has gone wrong in powerapps, we shouldn't be processing new registrations for this module ($($thismodule.fieldvalues.modulename)) as the max amount has been reached. Messaging Emily." -ForegroundColor Red
        $report = @()
        $report += "***************Errors found in Anthesis Academy Sync: We have reached the Maximum Sign Up Count For a Module***************" + "<br><br>"
        $report += "Errors found on this Module: $($thismodule.fieldvalues.ModuleName). We shouldn't be processing any more people as they shouldn't have had the option to sign up. This might have been a timing issue (unlikely though)." + "<br><br>"
        $report = $report | out-string
    Send-MailMessage -To "8ed81bd4.anthesisgroup.com@amer.teams.ms" -From "PeopleServicesRobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Anthesis Academy Sync: Error" -BodyAsHtml $report -Encoding UTF8 -Credential $exocreds
}


#Add the new registrant if they aren't already there - update the module list item
If($currentregistrants -notcontains $newregistrant.FieldValues.RegistrantName.Email){

#############
#If there are no current registrants (this is a pnp bug, some sort of array issue), just add 1, else add to the list and up the count
#############
If(!$currentregistrants){
    $moduleUpdate = Set-PnPListItem -List $masterModuleList -Identity $thismodule.Id -Values @{"RegistrantList" = $newregistrant.FieldValues.RegistrantName.Email; "RegistrantCount" = "1"}
}
Else{
    $currentregistrants += $newregistrant.FieldValues.RegistrantName.Email
    #Overwrite the current list and update the count
    $registrantcount = ($currentregistrants | Measure-Object).Count
    $moduleUpdate = Set-PnPListItem -List $masterModuleList -Identity $thismodule.Id -Values @{"RegistrantList" = $currentregistrants; "RegistrantCount" = $registrantcount}
}        
    #Check it worked
    
    #Registrant count matches number of registrants
    $thenewregistrantlist = Get-PnPListItem -List $masterModuleList -Id $thismodule.Id
    $registrants = $thenewregistrantlist.FieldValues.RegistrantList
    
    #New Registrant in Module Registrant list
    If(($moduleUpdate.FieldValues.RegistrantList.Email -contains $newregistrant.FieldValues.RegistrantName.Email) -and (($registrants | Measure-Object).Count -eq $thenewregistrantlist.FieldValues.RegistrantCount)){
    write-host "Success! $($newregistrant.FieldValues.RegistrantName.Email) now registered for $($thismodule.fieldvalues.ModuleName)" -ForegroundColor Yellow
    $27 = Set-PnPListItem -List $registrantProcessingList -Identity $newregistrant.Id -Values @{"Processed" = "Approved - Processed as Registrant"}
    
    #Send the Registrant an email confirming (still can't send user messages via Graph, only channel messages :(....)
                $body = "<HTML><BODY><p>Hi $($newregistrant.FieldValues.RegistrantName.Email),</p>
                <p>We just wanted to let you know that you have been successfully signed up for the Anthesis Academy Module <b>$($thismodule.fieldvalues.ModuleName)</b></p>
                <p>You don't need to do anything else - keep an eye on your Teams and Inbox for next steps from the Module Leader ($($thismodule.fieldvalues.ModuleLeader.LookupValue))</b><br><br></p>
                <p></p>
                <p>The Anthesis Academy</p>
                </BODY></HTML>"
                Send-MailMessage  -BodyAsHtml $body -Subject "You've Signed Up to an Anthesis Academy Module!" -to $($newregistrant.FieldValues.RegistrantName.Email) -from "AnthesisAcademy@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Encoding UTF8    
    }
    Else{
    Write-Host "Something went wrong registering $($newregistrant.FieldValues.RegistrantName.Email) to module: $($thismodule.fieldvalues.ModuleName). Messaging Emily." -ForegroundColor Red
        $report = @()
        $report += "***************Errors found in Anthesis Academy Sync: Something Went Wrong Processing a Registrant to a Module***************" + "<br><br>"
        $report += "Something went wrong registering $($newregistrant.FieldValues.RegistrantName.Email) to module: $($thismodule.fieldvalues.ModuleName)." + "<br><br>"
        $report = $report | out-string
    Send-MailMessage -To "8ed81bd4.anthesisgroup.com@amer.teams.ms" -From "PeopleServicesRobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Anthesis Academy Sync: Error" -BodyAsHtml $report -Encoding UTF8 -Credential $exocreds
    }
}
Else{
    write-host "$($newregistrant.FieldValues.RegistrantName.Email): Already signed up to the '$($thismodule.fieldvalues.ModuleName)' module. Not added." -ForegroundColor Red
}
}
Else{
Write-Host "Error: $($thismodule) Too many modules were found, we couldn't find the one needed! There are likely to be duplicate Module Codes in the list" -ForegroundColor Red
}
}

#Update each module with emails of Registrant
$allmodules = Get-PnPListItem -List $masterModuleList
ForEach($module in $allmodules){

#Get emails
$registrantEmails = $module.FieldValues.RegistrantList.Email
If($registrantEmails){
$emailArray = convertTo-arrayOfEmailAddresses $registrantEmails
    $formattedemails = @()
    ForEach($email in $emailArray){
    $emailtoadd = $email + ";"
    $formattedemails += $emailtoadd
    }
    $formated
$updateModuleEmailList = Set-PnPListItem -List $masterModuleList -Identity $module.Id -Values @{"Registrant_x0020_Emails" = "$($formattedemails)"}
}
}

