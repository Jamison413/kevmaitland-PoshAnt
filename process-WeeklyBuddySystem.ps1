#Set some logging
$Logname = "C:\Scripts" + "\Logs" + "\BuddySystem $(Get-Date -Format "yyMMdd").log" #Check this location before live
Start-Transcript -Path $Logname -Append
Write-Host "Script started:" (Get-date)
Write-host "**********************" -ForegroundColor White

#Connect to site
Import-Module _PS_Library_GeneralFunctionality
Import-Module _PNP_Library_SPO

$sharePointAdmin = "kimblebot@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content "$env:USERPROFILE\Desktop\KimbleBot.txt") 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass




<#---Process New Assignees---#>


#Process
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#



#Pull data for processing for this week (needs a cutoff date)
Connect-PnPOnline -url "https://anthesisllc.sharepoint.com/teams/IT_Team_All_365/" -Credentials $adminCreds

#Hardcode Adriana Quintero
Add-PnPListItem -List "Buddy System" -Values @{"Yourname" = "Adriana.Quintero@anthesisgroup.com"; "Community" = "Not applicable"; "Youtimezone" = "UTC+01:00"; "Yourcountry" = "France"}


#Process historical matches
$allhistoricassignees = Get-PnPListItem -List "Buddy System Historical Matches"
$HistoricIDList = $allhistoricassignees.FieldValues.MatchID

#Process assignees for the week
$allassignees = Get-PnPListItem -List "Buddy System"
#Process into pscustom object
$processedassignees = @()
ForEach($unprocessedassignee in $allassignees){

$processedassignees += New-Object psobject -Property @{

    "SharepointID" = $($unprocessedassignee.FieldValues.ID)
    "email" = $($unprocessedassignee.FieldValues.Yourname.Email);
    "community" = $($unprocessedassignee.FieldValues.Community);
    "timezone" = $($unprocessedassignee.FieldValues.Youtimezone);
    "userID" = $($unprocessedassignee.FieldValues.Yourname.LookupId);
    "name" = $($unprocessedassignee.FieldValues.Yourname.LookupValue);
    "country" = $($unprocessedassignee.FieldValues.Yourcountry);
}
}
#Get rid of dupes from the Buddy System SP List - or they will be moved over with everyone at script end into the waiting list
$dupecheck = $processedassignees
$processedassignees = $processedassignees | sort email -Unique
$dupestoremove = Compare-Object -ReferenceObject $dupecheck.email -DifferenceObject $processedassignees.email
$dupeuniqueupns = $dupestoremove

If($dupestoremove){
    ForEach($dupe in $dupeuniqueupns){
        $dupecount = $dupecheck | Where-Object -property "email" -EQ  "$($dupe.InputObject)"
        $totaltoremove = ($dupecount.count) - 1
        If($dupecount.Count -gt 1){
        Write-Host "Removing $($totaltoremove) dupes for $($dupe.InputObject)" -ForegroundColor Yellow
        $removalIDs = $($dupecount | Select-Object -First $($totaltoremove)) | select -Property "SharepointID"
            foreach($removalID in $removalIDs){
            Remove-PnPListItem -List "Buddy System" -Identity $removalID.SharepointID -Force
        }       
}
}
}
Else{
Write-Host "No dupes this week" -ForegroundColor Yellow
}
$count = $processedassignees | Measure-Object

#If there is an odd number - add and email Emily  ####this needs amending to loop round to top again
If((($count.count % 2) -eq "1") -or ($count.count -eq "1")){
write-host "We have an uneven number of assigneees this week, sending Emily a message..." -ForegroundColor Yellow
Exit
}
Else{
Write-Host "It looks like we've got an even number of assignees this week, let's continue..." -ForegroundColor Green
}

$thisweekslist = $processedassignees

#We create a multi-dimensional array to hold pairs of contacts by iterating through using two counters, one that starts from 0 and one that starts from half way through the array
$pairedArray = @($false)*[math]::Ceiling($processedassignees.length / 2)
#Set the second counter in the middle of the array to start, this needs to be reset every loop
$j = [math]::floor($processedassignees.length / 2)

for($r = 0; $r -lt 1000; $r++){
[System.Collections.ArrayList]$matchIDs = @()
#[System.Collections.ArrayList]$matchIDsold = @()
$j = [math]::floor($processedassignees.length / 2)
#We want to run this once and then check the pairings again historical records 
write-host "Run $($r + 1)..." -ForegroundColor Yellow
    for ($i = 0; $i -lt [math]::floor($processedassignees.length / 2); $i++){
        $pmatchID1 = "$($processedassignees[$i].userID)" + "$($processedassignees[$j].userID)"
        $pmatchID2 = "$($processedassignees[$j].userID)" + "$($processedassignees[$i].userID)"
        #$pmatchID1old = "$($processedassignees[$i].userID)" + "$($processedassignees[$j].userID)"
        #$pmatchID2old = "$($processedassignees[$j].userID)" + "$($processedassignees[$i].userID)"
        $pairedArray[$i] = @($processedassignees[$i],$processedassignees[$j])
        $matchIDs += $pmatchID1
        $matchIDs += $pmatchID2
        #$matchIDsold += $pmatchID1old
        #$matchIDsold += $pmatchID2old
        $j++
    }
#Check output
write-host "Here is our output for run $($r + 1):"
$matchIDs

#Then check if it matches any historical records - if so, do it again, if not break
[System.Collections.ArrayList]$badmatches = @()
#[System.Collections.ArrayList]$badmatchesold = @()
ForEach($ID in $matchIDs){
        Write-Host "$ID" -ForegroundColor White
        #Good
        If($HistoricIDList -notcontains $ID){$badmatches.Add(0)}
        #Bad
        Else{$badmatches.Add(1)}
        }
<#ForEach($ID in $matchIDsold){
        Write-Host "$ID" -ForegroundColor White
        #Good
        If($HistoricIDList -notcontains $ID){$badmatchesold.Add(0)}
        #Bad
        Else{$badmatchesold.Add(1)}
        }
#>
#If we found any bad matches from the last loop, continue and loop around again, if everything looks okay break out of the loop and end the block        
If($badmatches -contains "1"){
write-host "Looks like we've got historical duplicates, running again..." -ForegroundColor Red
#Do it again and overwrite current $pairedArray pairs, randomise array order of $processedassignees
$processedassignees = $processedassignees | Sort-Object {Get-Random}
Continue
}
Else{
write-host "Looks like our pairings are not like historical matches" -ForegroundColor Yellow
#Break the loop 
Break
    }
}

#If we can't find a match, email Emily and exit the script
If(($r = 1000) -and ($badmatches -contains 1)){
write-host "It looks like it wasn't mathematically possible to find  unique parinings this week :( Emailing Emily..." -ForegroundColor Red
$subject = "Buddy System: It looks like it wasn't mathemetically possible to find  unique pairings this week :("
$body = "Emails out were cancelled, womp, womp..."
Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "buddy.system@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 -Credential $adminCreds 
#Exit
}
Else{
#Add historical records
write-host "It looks like we have a set of unique pairings this week :) Continuing..." -ForegroundColor Green

}

#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#


#Checking
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
ForEach($pair in $pairedArray){

write-host "$($pair[0].email) and $($pair[1].email)" -ForegroundColor Green

}



#Finalise
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

#Send the emails :)
Write-host "Let's finish up and send out the emails!" -ForegroundColor Green
$summaryemailmatches = @()
ForEach($pair in $pairedArray){

#Generate a new Teams chat link - this will open an existing one if it is there already or start a new one (opens in browser but if Teams is installed user just has to press "Open in Teams" to open the chat in the Teams client. I cannot for the life of me render the links as href properly, Outlook refuses to process them properly, they will have to stay as ugly URL links for now (this might be a solution: https://social.technet.microsoft.com/Forums/ie/en-US/1703e410-e53f-45a7-9aa5-09b50df072c6/adding-a-href-hyperlink-to-htmlconverttohtml-via-powershell?forum=sharepointgeneralprevious)
$teamschatlink = "https://teams.microsoft.com/l/chat/0/0?users=" + "$($pair[0].email)" + "," + "$($pair[1].email)"
$wellbeingpageslink = "https://anthesisllc.sharepoint.com/sites/WellbeingResourcesSite"
$ITemail = "ITTeamAll@anthesisgroup.com"

#Construct the first email
[string]$community = $pair[1].community
$friendlyname0 = $($pair[0].name.Trim().Split(" ")[0].Trim())
$subject = "You have been matched with a new Anthesis Buddy!"
$body0 = "<HTML><FONT FACE=`"Calibri`">Hi $($friendlyname0),`r`n`r`n<BR><BR>"
$body0 += "You’re receiving this email as you signed up with the Anthesis Buddy system and you have a new match!`r`n`r`n<BR><BR>"
$body0 += "You have been matched with <font color='f36e21'><b>$($pair[1].name)</b></font color> in <font color='f36e21'><b>$($pair[1].country)</b></font color>.`r`n<BR>"
If(($community -ne "Not applicable")){$body0 += "<BR>They are in the $($community) Service Area"}
$body0 += "`r`n`r`n<BR><BR>"
$body0 += "<font color='f36e21'><b>What do I do next?</b></font color>`r`n`r`n<BR><BR>"
$body0 += "Over the next week, get in touch with your Buddy to get to know them more!`r`n`r`n<BR><BR>"
$body0 += "<font color='f36e21'><b>How do I contact my Anthesis Buddy?</b></font color>`r`n`r`n<BR><BR>"
$body0 += "You can make good use of Microsoft Team’s excellent chat and video functionality - jump into a chat with them now by clicking this link: `r`n<BR>$($teamschatlink)`r`n`r`n<BR><BR>"
$body0 += "Your Buddy is in timezone $($pair[1].timezone) so please be aware of time differences when getting in touch!`r`n`r`n<BR><BR>"
$body0 += "<font color='f36e21'><b>What happens at the end of the week?</b></font color>`r`n`r`n<BR><BR>"
$body0 += "You can still chat to your Buddy after the week has ended. You can also sign up again for another new Buddy on Thursdays by replying to the Anthesis Buddy System Re-Sign Up email (keep an eye on your inbox for this!). If you do not want to receive a new Buddy in the next week, either press the reject button or ignore the email and we will register your choice.`r`n`r`n<BR><BR>"
$body0 += "Don’t forget to check out the Anthesis Wellbeing pages on Sharepoint ($($wellbeingpageslink)) to find advice on how to look after Mind, Body, and Spirit, #AnthesisWFH posts and more!`r`n`r`n<BR><BR>"
$body0 += "With love,`r`n<BR>"
$body0 += "The Buddy System Robot <3`r`n`r`n<BR><BR>"
$body0 += "(Ps I’m managed by the IT Team, if I have broken or if you have any questions, please get in touch via $($ITemail))"


#Construct the second email
[string]$community = $pair[0].community
$friendlyname1 = $($pair[1].name.Trim().Split(" ")[0].Trim())
$subject = "You have been matched with a new Anthesis Buddy!"
$body1 = "<HTML><FONT FACE=`"Calibri`">Hi $($friendlyname1),`r`n`r`n<BR><BR>"
$body1 += "You’re receiving this email as you signed up with the Anthesis Buddy system and you have a new match!`r`n`r`n<BR><BR>"
$body1 += "You have been matched with <font color='f36e21'><b>$($pair[0].name)</b></font color> in <font color='f36e21'><b>$($pair[0].country)</b></font color>.`r`n<BR>"
If(($community -ne "Not applicable")){$body1 += "<BR>They are in the $($community) Service Area"}
$body1 += "`r`n`r`n<BR><BR>"
$body1 += "<font color='f36e21'><b>What do I do next?</b></font color>`r`n`r`n<BR><BR>"
$body1 += "Over the next week, get in touch with your Buddy to get to know them more!`r`n`r`n<BR><BR>"
$body1 += "<font color='f36e21'><b>How do I contact my Anthesis Buddy?</b></font color>`r`n`r`n<BR><BR>"
$body1 += "You can make good use of Microsoft Team’s excellent chat and video functionality - jump into a chat with them now by clicking this link: `r`n<BR>$($teamschatlink)`r`n`r`n<BR><BR>"
$body1 += "Your Buddy is in timezone $($pair[0].timezone) so please be aware of time differences when getting in touch!`r`n`r`n<BR><BR>"
$body1 += "<font color='f36e21'><b>What happens at the end of the week?</b></font color>`r`n`r`n<BR><BR>"
$body1 += "You can still chat to your Buddy after the week has ended. You can also sign up again for another new Buddy on Thursdays by replying to the Anthesis Buddy System Re-Sign Up email (keep an eye on your inbox for this!). If you do not want to receive a new Buddy in the next week, either press the reject button or ignore the email and we will register your choice.`r`n`r`n<BR><BR>"
$body1 += "Don’t forget to check out the Anthesis Wellbeing pages on Sharepoint ($($wellbeingpageslink)) to find advice on how to look after Mind, Body, and Spirit, #AnthesisWFH posts and more!`r`n`r`n<BR><BR>"
$body1 += "With love,`r`n<BR>"
$body1 += "The Buddy System Robot <3`r`n`r`n<BR><BR>"
$body1 += "(Ps I’m managed by the IT Team, if I have broken or if you have any questions, please get in touch via $($ITemail))"

#Send the emails and add to summary email
write-host "Sending emails to $($pair[0].email) and $($pair[1].email)" -ForegroundColor Green
$summaryemailmatches += "$($pair[0].email) and $($pair[1].email) `r`n <BR>"
Try{
Send-MailMessage -To "$($pair[0].email)" -From "buddy.system@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body0 -Encoding UTF8 -Credential $adminCreds
Send-MailMessage -To "$($pair[1].email)" -From "buddy.system@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body1 -Encoding UTF8 -Credential $adminCreds
}
Catch{
#send-mailmessage does not return anything if it fails oddly enough, so have captured the Powershell error to test instead
[string]$test = $Error[0].CategoryInfo.Activity
$emailcheck = 0
}
}

#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#


#Reporting and finalising
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#


#If emails are successful we can finalise all the processing

#update the intemediary Waiting List ready for Thursday - we need to clear out the main list to allow new submissions in the meantime
ForEach($person in $thisweekslist){
Add-PnPListItem -List "Buddy System Waiting List" -Values @{"Yourname" = $($person.name); "Yourcommunity_x0028_ifapplicable" = $($person.community); "Youtimezone" = $($person.timezone); "Yourcountry" = $($person.country)} 
Remove-PnPListItem -List "Buddy System" -Identity $($person.SharepointID) -Force
}

#Update the historical list of matches
write-host "It looks like we have a set of unique pairings this week :) Updating our historical record of pairings..." -ForegroundColor Green
ForEach($validID in $matchIDs){
Add-PnPListItem -List "Buddy System Historical Matches" -Values @{"MatchID" = $($validID)}
}


$buddyapplink = "https://apps.powerapps.com/play/dce65d94-2361-4d05-8bf3-8d4ad4362ffd?tenantId=271df584-ab64-437f-85b6-80ff9bef6c9f"
$leftoversignups = Get-PnPListItem -List "Buddy System Repeat Sign Up"
$summaryleftoverpeople = @()
ForEach($leftoversignup in $leftoversignups){

$subject = "Notification: Anthesis Buddy System Participation"
$body2 = "<HTML><FONT FACE=`"Calibri`">Hi $($leftoversignup.FieldValues.Yourname.LookupValue),`r`n`r`n<BR><BR>"
$body2 += "You’re receiving this email to confirm that you have been removed from the Anthesis Buddy system`r`n`r`n<BR><BR>"
$body2 += "<b>What do I need to do now?</b>`r`n`r`n<BR><BR>"
$body2 += "You do not need to do anything else. You can sign up again at any point to receive a Buddy on Fridays by going to the Buddy App. If you have missed the Re-Sign Up email or Teams message (sent on Thursdays) and want to sign up again for next week, you will need to sign back up via the Buddy App form here: $($buddyapplink).`r`n<BR><BR>"
$body2 += "With love,`r`n<BR>"
$body2 += "The Buddy System Robot <3`r`n`r`n<BR><BR>"
$body2 += "(Ps I’m managed by the IT Team, if I have broken or if you have any questions, please get in touch via $($ITemail))"

Write-Host "Emailing $($leftoversignup.FieldValues.Yourname.Email) to let them know they have been removed" -ForegroundColor Yellow
Send-MailMessage -To "$($leftoversignup.FieldValues.Yourname.Email)" -From "buddy.system@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body2 -Encoding UTF8 -Credential $adminCreds

$summaryleftoverpeople += "$($leftoversignup.FieldValues.Yourname.Email)<BR>"

Remove-PnPListItem -List "Buddy System Repeat Sign Up" -Identity $($leftoversignup.Id) -Force

}

#Add countries to summary email
$summaryemailcountries = @()
ForEach($processedperson in $processedassignees){
$summaryemailcountries += "$($processedperson.country) `r`n <BR>"
}


#Send summary email
$subject = "Anthesis Buddy System Summary Email"
$body3 = "<HTML><FONT FACE=`"Calibri`">Here are this week's numbers: $($processedassignees.count)`r`n`r`n<BR><BR>"
$body3 += "<HTML><FONT FACE=`"Calibri`">$($leftoversignups.count) either missed the re-sign up prompt or did not want to sign back up this week:`r`n`r`n<BR><BR>"
$body3 += "<HTML><FONT FACE=`"Calibri`">$summaryleftoverpeople`r`n`r`n<BR><BR>"

$body3 += "<HTML><FONT FACE=`"Calibri`">Here are this week's matches:`r`n`r`n<BR><BR>"
$body3 += "<HTML><FONT FACE=`"Calibri`">$summaryemailmatches`r`n`r`n<BR><BR>"

$body3 += "<HTML><FONT FACE=`"Calibri`">Here are this week's locations:`r`n`r`n<BR><BR>"
$body3 += "<HTML><FONT FACE=`"Calibri`">$summaryemailcountries`r`n`r`n<BR><BR>"

Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "buddy.system@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body3 -Encoding UTF8 -Credential $adminCreds
Send-MailMessage -To "paul.crewe@anthesisgroup.com" -From "buddy.system@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body3 -Encoding UTF8 -Credential $adminCreds

#end here
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#





<#---Process Re-Sign Up List---#>

#Get all the current users in the waiting list
$connect = Connect-PnPOnline -url "https://anthesisllc.sharepoint.com/teams/IT_Team_All_365/" -Credentials $adminCreds
$allwaiting = Get-PnPListItem -List "Buddy System Waiting List"

ForEach($person in $allwaiting){
$a = Add-PnPListItem -List "Buddy System Repeat Sign Up" -Values @{"Yourname" = $($person.FieldValues.Yourname.LookupValue); "Yourcommunity_x0028_ifapplicable" = $($person.FieldValues.Yourcommunity_x0028_ifapplicable); "Youtimezone" = $($person.FieldValues.Youtimezone); "Yourcountry" = $($person.FieldValues.Yourcountry)} 
$b = Remove-PnPListItem -List "Buddy System Waiting List" -Identity $($person.Id) -Force
}

Write-host "**********************" -ForegroundColor White
Write-Host "Script finished:" (Get-date)

Stop-Transcript








