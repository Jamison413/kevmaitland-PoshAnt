function new-dayShiftHash(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [datetime]$date 
        )

    $shifts = @(
        [ordered]@{
            shiftName  = "Morning"
            shiftStart = "$(get-date $date -Format "yyyy-MM-dd")T08:00:00Z"
            shiftEnd   = "$(get-date $date -Format "yyyy-MM-dd")T12:00:00Z"
            }
        ,[ordered]@{
            shiftName  = "Afternoon"
            shiftStart = "$(get-date $date -Format "yyyy-MM-dd")T12:00:00Z"
            shiftEnd   = "$(get-date $date -Format "yyyy-MM-dd")T16:00:00Z"
            }
        )
    $shifts
    }
function new-weekShiftHash(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [datetime]$date
        ,[parameter(Mandatory = $false)]
            [bool]$startFromPreviousMonday = $true
        ,[parameter(Mandatory = $false)]
            [bool]$suppressMonday = $false
        ,[parameter(Mandatory = $false)]
            [bool]$suppressTuesday = $false
        ,[parameter(Mandatory = $false)]
            [bool]$suppressWednesday = $false
        ,[parameter(Mandatory = $false)]
            [bool]$suppressThursday = $false
        ,[parameter(Mandatory = $false)]
            [bool]$suppressFriday = $false
        ,[parameter(Mandatory = $false)]
            [bool]$suppressSaturday = $false
        ,[parameter(Mandatory = $false)]
            [bool]$suppressSunday = $false
        )

    if($startFromPreviousMonday){
        $delta = $date.DayOfWeek.value__ - 1
        $date = $date.AddDays(-$delta).Date
        }

    $dayMap = ,@()*7
    $dayMap[0] = $suppressSunday
    $dayMap[1] = $suppressMonday
    $dayMap[2] = $suppressTuesday
    $dayMap[3] = $suppressWednesday
    $dayMap[4] = $suppressThursday
    $dayMap[5] = $suppressFriday
    $dayMap[6] = $suppressSaturday

    for($i=0;$i -lt 7;$i++){
        if($dayMap[$($date.AddDays($i).DayOfWeek.value__)] -eq $false){ #Weird double-negative, but it'll make more sense to people calling the function if they actively have to suppress days fo the week
            [array]$weekOfShifts += new-dayShiftHash -date $date.AddDays($i)
            }
        }
    $weekOfShifts
    }

$offices = @()
$offices += [ordered]@{
    OfficeName="GBR-Oxford"
    OfficeColour="Yellow"
    OfficeWeekendColour="darkYellow"
    OfficeDesks=8
    ShiftNotes="Remember to clock in using the Wheelhouse App (available on the Anthesis and/or personal App Store), and to use hand sanitiser when entering/exiting the office please"
    }
$offices += [ordered]@{
    OfficeName="GBR-Bristol"
    OfficeColour="Green"
    OfficeWeekendColour="darkGreen"
    OfficeDesks=18
    ShiftNotes="Remember to use hand sanitiser when entering/exiting the office please"
    }
$offices += [ordered]@{
    OfficeName="GBR-London"
    OfficeColour="Blue"
    OfficeWeekendColour="darkBlue"
    OfficeDesks=16
    ShiftNotes="Remember to use hand sanitiser when entering/exiting the office please"
    }
$offices += [ordered]@{
    OfficeName="GBR-Manchester"
    OfficeColour="Pink"
    OfficeWeekendColour="darkPink"
    OfficeDesks=6
    ShiftNotes="Remember to use hand sanitiser when entering/exiting the office please"
    }

#$teamId = "2bea0e44-9491-4c30-9e8f-7620ccacac73" #Teams Testing Team
$teamId = "549dd0d0-251f-4c23-893e-9d0c31c2dc13" #All (GBR)
$msAppActsAsUserId = "00aa81e4-2e8f-4170-bc24-843b917fd7cf" #GroupBot

$shiftBotDetails = get-graphAppClientCredentials -appName ShiftBot
$tokenResponseShiftBot = get-graphTokenResponse -grant_type client_credentials -aadAppCreds $shiftBotDetails
#$tokenResponseShiftBot = get-graphTokenResponse -grant_type device_code -aadAppCreds $shiftBotDetails
$teamBotDetails = get-graphAppClientCredentials -appName TeamsBot
$tokenResponseTeamBot = get-graphTokenResponse -aadAppCreds $teamBotDetails


$openShifts = invoke-graphGet -tokenResponse $tokenResponseshiftBot -graphQuery "/teams/$teamId/schedule/openshifts" -additionalHeaders @{"MS-APP-ACTS-AS"=$msAppActsAsUserId}
[datetime]$lastScheduledDay = $openShifts | Group-Object schedulingGroupId | % {$_.Group.sharedOpenShift.endDateTime} | Sort-Object -Descending | select -Index 0
[datetime]$nextMonday = $lastScheduledDay.AddDays(-$($lastScheduledDay.DayOfWeek.value__ - 1)+7) 

if($lastScheduledDay.DayOfWeek.value__ -eq 0){$nextMonday = $nextMonday.AddDays(-7)} #Special case for Sundays being part of the wrong week
$nextWeekOfShifts = new-weekShiftHash -date $nextMonday -suppressMonday:$true -suppressFriday:$true -suppressSaturday:$true -suppressSunday:$true

#Get the SchedulingGroupId values for the offices
$officeSchedulingGroups = invoke-graphGet -tokenResponse $tokenResponseshiftBot -graphQuery "/teams/$teamId/schedule/schedulingGroups" -additionalHeaders @{"MS-APP-ACTS-AS"=$msAppActsAsUserId}
$teamMembers = get-graphUsersFromGroup -tokenResponse $tokenResponseTeamBot -groupId $teamId -memberType TransitiveMembers -returnOnlyLicensedUsers 

#Process the new Shifts
$offices | % {
    $thisOffice = $_
    $thisSchedulingGroup = $officeSchedulingGroups | ? {$_.displayName -match $thisOffice["OfficeName"]}
    if([string]::IsNullOrWhiteSpace($thisSchedulingGroup)){
        Write-Warning "Couldn't match an Id for Office [$($thisOffice["OfficeName"])]"
        }
    else{
        Write-Verbose "Adding id [$($thisSchedulingGroup.id)] to office [$($thisOffice["OfficeName"])]"
        $_.Add("id",$thisSchedulingGroup.id) #Add the SchedulingGroupId to our Offices hashtable
        
        $updatedHash = @{ #Keep the SchedulingGroups up-to-date
            displayName="$($thisOffice["OfficeName"]) (Max $($thisOffice["OfficeDesks"]))" #This will update the SchedulingGroup Name if the number of available desks changes
            isActive = $thisSchedulingGroup.isActive
            userIds = @($teamMembers.id) #This will automaticlly add all Team Members to each SchedulingGroup
            }
        invoke-graphPut -tokenResponse $tokenResponseshiftBot -graphQuery "/teams/$teamId/schedule/schedulingGroups/$($thisSchedulingGroup.id)" -graphBodyHashtable $updatedHash -additionalHeaders @{"MS-APP-ACTS-AS"=$msAppActsAsUserId} -Verbose:$VerbosePreference #PATCH doesn't work on schedulingGroups yet :'( But PUT works!
             
        $nextWeekOfShifts | % { #Add 1 week's worth of shifts 
            $thisShift = $_
            if(@(6,0) -contains $(Get-Date $thisShift["shiftStart"]).DayOfWeek.value__){ #If the shift is at a weekend, use a different colour
                Write-Verbose "`tCreating weekend shift for [$($thisOffice["OfficeName"])] on [$($(Get-Date $thisShift["shiftStart"]).DayOfWeek)][$(Get-Date $thisShift["shiftStart"] -Format g)]"
                new-graphOpenShiftShared -tokenResponse $tokenResponseshiftBot -teamId $teamId -schedulingGroupId $thisSchedulingGroup.id -shiftName $("$($thisOffice["OfficeName"]) $($thisShift["shiftName"])") -shiftNotes $thisOffice["shiftNotes"] -availableSlots $thisOffice["OfficeDesks"] -startDateTime $thisShift["shiftStart"] -endDateTime $thisShift["shiftEnd"] -shiftColour $thisOffice["OfficeWeekendColour"] -MsAppActsAsUserId $msAppActsAsUserId -Verbose:$VerbosePreference
                }
            else{
                Write-Verbose "`tCreating weekday shift for [$($thisOffice["OfficeName"])] on [$($(Get-Date $thisShift["shiftStart"]).DayOfWeek)][$(Get-Date $thisShift["shiftStart"] -Format g)]"
                new-graphOpenShiftShared -tokenResponse $tokenResponseshiftBot -teamId $teamId -schedulingGroupId $thisSchedulingGroup.id -shiftName $("$($thisOffice["OfficeName"]) $($thisShift["shiftName"])") -shiftNotes $thisOffice["shiftNotes"] -availableSlots $thisOffice["OfficeDesks"] -startDateTime $thisShift["shiftStart"] -endDateTime $thisShift["shiftEnd"] -shiftColour $thisOffice["OfficeColour"] -MsAppActsAsUserId $msAppActsAsUserId -Verbose:$VerbosePreference
                }
            }
        }
    }
