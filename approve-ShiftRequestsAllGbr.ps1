#$VerbosePreference = 2
$logFileLocation = "C:\ScriptLogs\"
$logFileName = "approve-shiftRequestsAllGbr"
$fullLogPathAndName = $logFileLocation+$logFileName+"_$whatToSync`_FullLog_$(Get-Date -Format "yyMMdd").log"
$errorLogPathAndName = $logFileLocation+$logFileName+"_$whatToSync`_ErrorLog_$(Get-Date -Format "yyMMdd").log"
if($PSCommandPath){
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_$whatToSync`_Transcript_$(Get-Date -Format "yyMMdd").log"
    Start-Transcript $transcriptLogName -Append
    }


<#$Admin = "emily.pressey@anthesisgroup.com"
$AdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\Emily.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin, $AdminPass
$exoCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin, $AdminPass
connect-ToExo -credential $exoCreds
#>

$TeamsReport = @()

$teamBotDetails = get-graphAppClientCredentials -appName TeamsBot
$tokenResponseTeams = get-graphTokenResponse -aadAppCreds $teamBotDetails
$smtpBotDetails = get-graphAppClientCredentials -appName SmtpBot
$tokenResponseSmtp = get-graphTokenResponse -aadAppCreds $smtpBotDetails
$shiftBotDetails = get-graphAppClientCredentials -appName ShiftBot
$tokenResponseShiftBot = get-graphTokenResponse -aadAppCreds $shiftBotDetails


$allgraphusers = get-graphUsers -tokenResponse $tokenResponseTeams -filterLicensedUsers

#$teamId = "2bea0e44-9491-4c30-9e8f-7620ccacac73" #Teams Testing Team
$teamId = "549dd0d0-251f-4c23-893e-9d0c31c2dc13" #All (GBR)
#$msAppActsAsUserId = "36bc6f20-feed-422d-b2f2-7758e9708604" #Kev Maitland
$groupBotGuid = "00aa81e4-2e8f-4170-bc24-843b917fd7cf"
$msAppActsAsUserId = "00aa81e4-2e8f-4170-bc24-843b917fd7cf" #GroupBot

#Write-Verbose "Access Token: [$($tokenResponseShiftBot.access_token)]"

Write-Host "Approving OpenShiftChangeRequests" -ForegroundColor Cyan
#Approve OpenShiftChangeRequests - these are normal requests to book a slot
#$allShifts = get-graphShiftOpenShifts -tokenResponse $tokenResponseShiftBot -teamId $teamId -MsAppActsAsUserId $msAppActsAsUserId #Swap this shit out when openShifts supports filtering on id
$pendingRequests = get-graphShiftOpenShiftChangeRequests -tokenResponse $tokenResponseshiftBot -teamId $teamId -MsAppActsAsUserId $groupBotGuid -requestState pending -Verbose:$VerbosePreference
if($pendingRequests){
    Write-Host "We've found pending requests to process [$(($pendingRequests | Measure-Object).count)]" -ForegroundColor Cyan
    Write-Verbose ($pendingRequests | Out-String)
    $pendingRequests | % {
        $thisRequest = $_
        Write-Host "Processing OpenShiftChangeRequest for $($allgraphusers.where({$_.Id -eq $thisRequest.senderUserId}).userPrincipalName) for $($thisRequest.openShiftId)" -ForegroundColor Cyan
        invoke-graphPost -tokenResponse $tokenResponseShiftBot -graphQuery "/teams/$teamId/schedule/openShiftChangeRequests/$($thisRequest.id)/approve" -additionalHeaders @{"MS-APP-ACTS-AS"=$groupBotGuid} -graphBodyHashtable @{message="Approve-ulated"}
        $shift = get-graphShiftOpenShifts -tokenResponse $tokenResponseShiftBot -teamId $teamId -MsAppActsAsUserId $groupBotGuid -openShiftid $thisRequest.openShiftId 
        #$shift = $allShifts | ? {$_.id -eq $thisRequest.openShiftId}
        new-graphCalendarEvent -tokenResponse $tokenResponseShiftBot -userId $thisRequest.senderUserId -subject "$($shift.sharedOpenShift.displayName) desk reservation" -start $shift.sharedOpenShift.startDateTime -startTimeZone 'GMT Standard Time' -end $shift.sharedOpenShift.endDateTime -endTimeZone 'GMT Standard Time' -location $($shift.sharedOpenShift.displayName.Split(" ")[0]) -bodyHTML $shift.sharedOpenShift.notes -reminderMinutesBeforeStart $(20*60) -freeBusyStatus free -categories @("BookedByShiftBot",$shift.id) -Verbose:$VerbosePreference
        }
    }


Write-Host "Approving ShiftOfferRequests" -ForegroundColor Cyan
#Approve OfferShiftRequests with Group Bot - this is to hand back a slot (someone has cancelled their booking)
$allShifts = get-graphShiftOpenShifts -tokenResponse $tokenResponseShiftBot -teamId $teamId -MsAppActsAsUserId $msAppActsAsUserId
$offerPendingRequests = get-graphShiftofferShiftRequests -tokenResponse $tokenResponseshiftBot -teamId $teamId -MsAppActsAsUserId $msAppActsAsUserId -requestState pending  -Verbose:$VerbosePreference
#Find only the Shift offers to groupbot using the hard-coded ID (we're unlikely to change it)
$offerPendingRequests = $offerPendingRequests | Where-Object -Property "recipientUserId" -eq "00aa81e4-2e8f-4170-bc24-843b917fd7cf"
if($offerPendingRequests){
    Write-Host "We've found pending offer requests to process [$(($offerPendingRequests | Measure-Object).count)]" -ForegroundColor Cyan
    Write-Verbose ($offerPendingRequests | Out-String)
    $offerPendingRequests | % {
        $thisOfferRequest = $_
        Write-Host "Processing ShiftOfferRequest for [$($allgraphusers.where({$_.Id -eq $thisOfferRequest.senderUserId}).userPrincipalName)] for [$($thisOfferRequest.senderShiftId)]" -ForegroundColor Cyan
        #First get the existing Shift AND OpenShift that has been offered
        $existingUserShift = invoke-graphGet -tokenResponse $tokenResponseShiftBot -graphQuery "/teams/$teamId/schedule/shifts/$($thisOfferRequest.senderShiftId)" -additionalHeaders @{"MS-APP-ACTS-AS"=$MsAppActsAsUserId} -Verbose:$VerbosePreference
        $thisOpenShift = $allShifts.Where({($_.sharedOpenShift.displayName -eq $existingUserShift.sharedShift.displayName) -and ($_.sharedOpenShift.startDateTime -eq $existingUserShift.sharedShift.startDateTime) -and ($_.sharedOpenShift.endDateTime -eq $existingUserShift.sharedShift.endDateTime)})
        #$thisOpenShift = get-graphShiftOpenShifts -tokenResponse $tokenResponseShiftBot -teamId $teamId -MsAppActsAsUserId $groupBotGuid -openShiftid $existingUserShift.id 
        
        If(($thisOpenShift | Measure-Object).Count -eq 1){
            Write-Host "we've found the open shift - processing the approval, updating the open shift slot count and deleting the calendar entry for the user" -ForegroundColor Cyan
            #Process the offer approval, and update the OPENSHFT slot count to reflect the change (or it stays as one less slot available - the offer shift is just to give to another user, not back to the OPENSHFT)
            invoke-graphPost -tokenResponse $tokenResponseShiftBot -graphQuery "/teams/$teamId/schedule/offerShiftRequests/$($thisOfferRequest.id)/approve" -additionalHeaders @{"MS-APP-ACTS-AS"=$msAppActsAsUserId} -graphBodyHashtable @{message="Approve-ulated"}
            #$openshift = get-graphShiftOpenShifts -tokenResponse $tokenResponseShiftBot -teamId $teamId -MsAppActsAsUserId $msAppActsAsUserId -filterId $thisevent.categories[1] -Verbose:$VerbosePreference #for when we can filter on OPENSHFT id
            update-graphOpenShiftShared -tokenResponse $tokenResponseShiftBot -schedulingGroupId $thisOpenShift.schedulingGroupId -openShiftId $thisOpenShift.Id -teamId $teamId -availableSlots ($thisOpenShift.sharedOpenShift.openSlotCount + 1) -MsAppActsAsUserId $msAppActsAsUserId -Verbose:$VerbosePreference
        
            #Try to find the calendar event by searching the user's calendar for an event the name of the shift that is handed back - this is a real pain as we can't look up the OPENSHFT from the SHFT record in the offer request....so we're searching by subject and start/end time :(
            $formattedshiftdate = ($existingUserShift.sharedShift.startDateTime).Split("T")[0]
            $events = get-graphCalendarEvent -tokenResponse $tokenResponseShiftBot -userId $existingUserShift.userId -filterSubject $($existingUserShift.sharedShift.displayName + " desk reservation") -Verbose:$VerbosePreference
            #Also being lazy, working with datetime in Exchange is often a pain - match on the start time, should be accurate enough after filtering on event subject (as we divide by office locations + morning and afternoon reservations)
            $thisevent = $events | Where-Object {(($_.start.dateTime).split("T")[0]) -eq $formattedshiftdate} -ErrorAction SilentlyContinue
            #If we've only managed to grab one event - delete the calendar event in the user's calendar, 
                If(($thisevent | Measure-Object).count -eq 1){
                    delete-graphCalendarEvent -tokenResponse $tokenResponseShiftBot -userId $existingUserShift.userId -eventId $thisevent.id -Verbose:$VerbosePreference
                }
                Else{
                    Write-Host "Woops, we couldn't find the EXACT calendar event to delete from the Shift :(" -ForegroundColor Red
                    get-graphUsers -tokenResponse $tokenResponseShiftBot $tokenResponseShiftBot -filterCustomEq
                    $TeamsReport += @{"(Shift Offer - Calendar Event Deletion) ERROR" = "Error finding calendar event for $($allgraphusers.where({$_.Id -eq $thisOfferRequest.senderUserId}).userPrincipalName) for event $($existingUserShift.sharedShift.displayName) on $($existingUserShift.sharedShift.startDateTime). <BR>They could have deleted it."}
                }
        }
        Else{
        Write-Host "We couldn't find the open shift - we can't do anything without this info and so nothing has been changed" -ForegroundColor Red
        $TeamsReport += @{"(Shift Offer - Approval Processing) ERROR" = "Error finding Open Shift for $($allgraphusers.where({$_.Id -eq $thisOfferRequest.senderUserId}).userPrincipalName) for shift $($existingUserShift.sharedShift.displayName) on $($existingUserShift.sharedShift.startDateTime). <BR>Nothing was changed and this will need looking into."}
        }
    }
}

#Clear out groupbot Shifts or the OPENSHFT will still show 1 extra as assigned
$groupbotshifts = get-graphShiftUserShifts -tokenResponse $tokenResponseShiftBot -teamId $teamId -MsAppActsAsUserId $msAppActsAsUserId | Where-Object -Property "userId" -EQ "00aa81e4-2e8f-4170-bc24-843b917fd7cf"
ForEach($groupbotshift in $groupbotshifts){delete-graphShiftUserShifts -tokenResponse $tokenResponseShiftBot -teamId $teamId -shiftId $groupbotshift.id -MsAppActsAsUserId $msAppActsAsUserId -Verbose:$VerbosePreference}



If($TeamsReport){

    $report = @()
    $report += "***************Errors found in Shift Approval***************" + "<br><br>"
        ForEach($t in $TeamsReport){
        $report += "$($t.Keys)" + " - " + "$($t.Values)" + "<br><br>"
}
$report = $report | out-string

#Send-MailMessage -To "c6167716.anthesisgroup.com@amer.teams.ms" -From "ShiftsRobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Shifts Approval Error" -BodyAsHtml $report -Encoding UTF8 -Credential $exocreds
send-graphMailMessage -tokenResponse $tokenResponseSmtp -fromUpn "groupbot@anthesisgroup.com" -toAddresses "c6167716.anthesisgroup.com@amer.teams.ms" -subject "Shifts Approval Error" -bodyHtml $report
}

Stop-Transcript