#$VerbosePreference = 2
$logFileLocation = "C:\ScriptLogs\"
$logFileName = "approve-shiftRequestsAllGbr"
$fullLogPathAndName = $logFileLocation+$logFileName+"_$whatToSync`_FullLog_$(Get-Date -Format "yyMMdd").log"
$errorLogPathAndName = $logFileLocation+$logFileName+"_$whatToSync`_ErrorLog_$(Get-Date -Format "yyMMdd").log"
if($PSCommandPath){
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_$whatToSync`_Transcript_$(Get-Date -Format "yyMMdd").log"
    Start-Transcript $transcriptLogName -Append
    }


#$teamId = "2bea0e44-9491-4c30-9e8f-7620ccacac73" #Teams Testing Team
$teamId = "549dd0d0-251f-4c23-893e-9d0c31c2dc13" #All (GBR)
#$msAppActsAsUserId = "36bc6f20-feed-422d-b2f2-7758e9708604" #Kev Maitland
$msAppActsAsUserId = "00aa81e4-2e8f-4170-bc24-843b917fd7cf" #GroupBot

$shiftBotDetails = get-graphAppClientCredentials -appName ShiftBot
$tokenResponseShiftBot = get-graphTokenResponse -grant_type client_credentials -aadAppCreds $shiftBotDetails
#Write-Verbose "Access Token: [$($tokenResponseShiftBot.access_token)]"

Write-Host "Approving OpenShiftChangeRequests" -ForegroundColor Cyan
#Approve OpenShiftChangeRequests - these are normal requests to book a slot
$allShifts = get-graphShiftOpenShifts -tokenResponse $tokenResponseShiftBot -teamId $teamId -MsAppActsAsUserId $msAppActsAsUserId #Swap this shit out when openShifts supports filtering on id
$pendingRequests = get-graphShiftOpenShiftChangeRequests -tokenResponse $tokenResponseshiftBot -teamId $teamId -MsAppActsAsUserId $msAppActsAsUserId -requestState pending -Verbose:$VerbosePreference
if($pendingRequests){
    Write-Verbose ($pendingRequests | Out-String)
    $pendingRequests | % {
        $thisRequest = $_
        invoke-graphPost -tokenResponse $tokenResponseShiftBot -graphQuery "/teams/$teamId/schedule/openShiftChangeRequests/$($thisRequest.id)/approve" -additionalHeaders @{"MS-APP-ACTS-AS"=$msAppActsAsUserId} -graphBodyHashtable @{message="Approve-ulated"}
        #$shift = get-graphShiftOpenShifts -tokenResponse $tokenResponseShiftBot -teamId $teamId -MsAppActsAsUserId $msAppActsAsUserId -filterId $pendingRequests[0].openShiftId  #FFS - openShifts doesn't support filtering on id yet (2020-08-19) >:(
        $shift = $allShifts | ? {$_.id -eq $thisRequest.openShiftId}
        new-graphCalendarEvent -tokenResponse $tokenResponseShiftBot -userId $thisRequest.senderUserId -subject "$($shift.sharedOpenShift.displayName) desk reservation" -start $shift.sharedOpenShift.startDateTime -startTimeZone 'GMT Standard Time' -end $shift.sharedOpenShift.endDateTime -endTimeZone 'GMT Standard Time' -location $($shift.sharedOpenShift.displayName.Split(" ")[0]) -bodyHTML $shift.sharedOpenShift.notes -reminderMinutesBeforeStart $(20*60) -freeBusyStatus free -categories @("BookedByShiftBot",$shift.id) -Verbose:$VerbosePreference
        }
    }

Write-Host "Approving ShiftOfferRequests" -ForegroundColor Cyan
#Approve OfferShiftRequests with Group Bot - this is to hand back a slot (someone has cancelled their booking)
$allShifts = get-graphShiftOpenShifts -tokenResponse $tokenResponseShiftBot -teamId $teamId -MsAppActsAsUserId $msAppActsAsUserId
$offerpendingRequests = get-graphShiftofferShiftRequests -tokenResponse $tokenResponseshiftBot -teamId $teamId -MsAppActsAsUserId $msAppActsAsUserId -requestState pending  -Verbose:$VerbosePreference
#Find only the Shift offers to groupbot using the hard-coded ID (we're unlikely to change it)
$offerpendingRequests = $offerpendingRequests | Where-Object -Property "recipientUserId" -eq "00aa81e4-2e8f-4170-bc24-843b917fd7cf"
if($offerpendingRequests){
    Write-Verbose ($offerpendingRequests | Out-String)
    $offerpendingRequests | % {
        $thisOfferRequest = $_
        #First get the existing Shift that has been offered
        $existingusershift = invoke-graphGet -tokenResponse $tokenResponseShiftBot -graphQuery "/teams/$teamId/schedule/shifts/$($thisOfferRequest.senderShiftId)" -additionalHeaders @{"MS-APP-ACTS-AS"=$MsAppActsAsUserId} -Verbose:$VerbosePreference
        $formattedshiftdate = ($existingusershift.sharedShift.startDateTime).Split("T")[0]
        #Find the calendar event by searching the user's calendar for an event the name of the shift that is handed back - this is a real pain as we can't look up the OPENSHFT from the SHFT record in the offer request....so we're searching by subject and start/end time :(
        $events = get-graphCalendarEvent -tokenResponse $tokenResponseShiftBot -userId $existingusershift.userId -filterSubject $($existingusershift.sharedShift.displayName + " desk reservation") -Verbose:$VerbosePreference
        #Also being lazy, working with datetime in Exchange is often a pain - match on the start time, should be accurate enough after filtering on event subject (as we divide by office locations + morning and afternoon reservations)
        $thisevent = $events | Where-Object {(($_.start.dateTime).split("T")[0]) -eq $formattedshiftdate}
        #If we've only managed to grab one event - delete the calendar event in the user's calendar, process the offer approval, and update the OPENSHFT slot count to reflect the change (or it stays as one less slot available - the offer shift is just to give to anotehr user, not back to the OPENSHFT)
        If(($thisevent | Measure-Object).count -eq 1){
            delete-graphCalendarEvent -tokenResponse $tokenResponseShiftBot -userId $existingusershift.userId -eventId $thisevent.id -Verbose:$VerbosePreference
            invoke-graphPost -tokenResponse $tokenResponseShiftBot -graphQuery "/teams/$teamId/schedule/offerShiftRequests/$($thisOfferRequest.id)/approve" -additionalHeaders @{"MS-APP-ACTS-AS"=$msAppActsAsUserId} -graphBodyHashtable @{message="Approve-ulated"}
            #$openshift = get-graphShiftOpenShifts -tokenResponse $tokenResponseShiftBot -teamId $teamId -MsAppActsAsUserId $msAppActsAsUserId -filterId $thisevent.categories[1] -Verbose:$VerbosePreference #for when we can filter on OPENSHFT id
            $thisOpenShift = get-graphShiftOpenShifts -tokenResponse $tokenResponseShiftBot -teamId $teamId -MsAppActsAsUserId $msAppActsAsUserId | Where-Object -Property "Id" -EQ $thisevent.categories[1]
            Update-graphOpenShiftShared -tokenResponse $tokenResponseShiftBot -schedulingGroupId $thisOpenShift.schedulingGroupId -openShiftId $thisOpenShift.Id -teamId $teamId -availableSlots ($thisOpenShift.sharedOpenShift.openSlotCount + 1) -MsAppActsAsUserId $msAppActsAsUserId -Verbose:$VerbosePreference
        }
        Else{
            Write-Host "Woops, we couldn't find the EXACT calendar event to delete from the Shift :( We can't do anything without this information - nothing was changed" -ForegroundColor Red
        }
    }
}

#Clear out groupbot Shifts or the OPENSHFT will still show 1 extra as assigned
$groupbotshifts = get-graphShiftUserShifts -tokenResponse $tokenResponseShiftBot -teamId $teamId -MsAppActsAsUserId $msAppActsAsUserId | Where-Object -Property "userId" -EQ "00aa81e4-2e8f-4170-bc24-843b917fd7cf"
ForEach($groupbotshift in $groupbotshifts){delete-graphShiftUserShifts -tokenResponse $tokenResponseShiftBot -teamId $teamId -shiftId $groupbotshift.id -MsAppActsAsUserId $msAppActsAsUserId -Verbose:$VerbosePreference}


Stop-Transcript