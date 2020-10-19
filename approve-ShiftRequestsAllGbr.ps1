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

$allShifts = get-graphShiftOpenShifts -tokenResponse $tokenResponseShiftBot -teamId $teamId -MsAppActsAsUserId $msAppActsAsUserId #Swap this shit out when openShifts supports filtering on id
$pendingRequests = get-graphShiftOpenShiftChangeRequests -tokenResponse $tokenResponseshiftBot -teamId $teamId -MsAppActsAsUserId "36bc6f20-feed-422d-b2f2-7758e9708604" -requestState pending -Verbose:$VerbosePreference
if($pendingRequests){
    Write-Verbose $pendingRequests
    $pendingRequests | % {
        $thisRequest = $_
        invoke-graphPost -tokenResponse $tokenResponseShiftBot -graphQuery "/teams/$teamId/schedule/openShiftChangeRequests/$($thisRequest.id)/approve" -additionalHeaders @{"MS-APP-ACTS-AS"=$msAppActsAsUserId} -graphBodyHashtable @{message="Approve-ulated"}
        #$shift = get-graphShiftOpenShifts -tokenResponse $tokenResponseShiftBot -teamId $teamId -MsAppActsAsUserId $msAppActsAsUserId -filterId $pendingRequests[0].openShiftId  #FFS - openShifts doesn't support filtering on id yet (2020-08-19) >:(
        $shift = $allShifts | ? {$_.id -eq $thisRequest.openShiftId}
        new-graphCalendarEvent -tokenResponse $tokenResponseShiftBot -userId $thisRequest.senderUserId -subject "$($shift.sharedOpenShift.displayName) desk reservation" -start $shift.sharedOpenShift.startDateTime -startTimeZone 'GMT Standard Time' -end $shift.sharedOpenShift.endDateTime -endTimeZone 'GMT Standard Time' -location $($shift.sharedOpenShift.displayName.Split(" ")[0]) -bodyHTML $shift.sharedOpenShift.notes -reminderMinutesBeforeStart $(20*60) -freeBusyStatus free -categories @("BookedByShiftBot",$shift.id) -Verbose:$VerbosePreference
        }
    }



Stop-Transcript