﻿function new-dayShiftHash(){
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
        )

    if($startFromPreviousMonday){
        $delta = $date.DayOfWeek.value__ - 1
        $date = $date.AddDays(-$delta).Date
        }

    for($i=0;$i -lt 7;$i++){
        [array]$weekOfShifts += new-dayShiftHash -date $date.AddDays($i)
        }
    $weekOfShifts
    }

$offices = @()
$offices += [ordered]@{
    OfficeName="GBR-Oxford"
    OfficeColour="Yellow"
    OfficeWeekendColour="darkYellow"
    OfficeDesks=8
    }
$offices += [ordered]@{
    OfficeName="GBR-Bristol"
    OfficeColour="Green"
    OfficeWeekendColour="darkGreen"
    OfficeDesks=18
    }
$offices += [ordered]@{
    OfficeName="GBR-London"
    OfficeColour="Blue"
    OfficeWeekendColour="darkBlue"
    OfficeDesks=16
    }
$offices += [ordered]@{
    OfficeName="GBR-Manchester"
    OfficeColour="Pink"
    OfficeWeekendColour="darkPink"
    OfficeDesks=4
    }

$teamId = "2bea0e44-9491-4c30-9e8f-7620ccacac73" #Teams Testing Team
#$teamId = "549dd0d0-251f-4c23-893e-9d0c31c2dc13" #All (GBR)
$standardShiftNotes = "Remember to sign in/out with the Blip! App, and use hand sanitiser when entering/exiting the office please"

$shiftBotDetails = get-graphAppClientCredentials -appName ShiftBot
$tokenResponseShiftBot = get-graphTokenResponse -grant_type client_credentials -aadAppCreds $shiftBotDetails
#$tokenResponseShiftBot = get-graphTokenResponse -grant_type device_code -aadAppCreds $shiftBotDetails
$teamBotDetails = import-encryptedCsv -pathToEncryptedCsv "$env:USERPROFILE\Desktop\teambotdetails.txt"
$tokenResponseTeamBot = get-graphTokenResponse -aadAppCreds $teamBotDetails

#Get the last listed OpenShift, then generate standard shifts for the following week 
$openShifts = invoke-graphGet -tokenResponse $tokenResponseshiftBot -graphQuery "/teams/$teamId/schedule/openshifts" -additionalHeaders @{"MS-APP-ACTS-AS"=$msAppActsAsUserId}
$lastOpenShifts = $openShifts | Group-Object schedulingGroupId | % {$_.Group | Sort-Object sharedOpenShift.endDateTime | Select-Object -Last 1}
[datetime]$lastScheduledDay = $($lastOpenShifts | Sort-Object {$_.sharedOpenShift.endDateTime} | Select-Object -Last 1).sharedOpenShift.endDateTime
[datetime]$nextMonday = $lastScheduledDay.AddDays(-$($lastScheduledDay.DayOfWeek.value__ - 1)+7) 
if($lastScheduledDay.DayOfWeek.value__ -eq 0){$nextMonday = $nextMonday.AddDays(-7)} #Special case for Sundays being part of the wrong week
$nextWeekOfShifts = new-weekShiftHash -date $nextMonday

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
        invoke-graphPut -tokenResponse $tokenResponseshiftBot -graphQuery "/teams/$teamId/schedule/schedulingGroups/$($thisSchedulingGroup.id)" -graphBodyHashtable $updatedHash -additionalHeaders @{"MS-APP-ACTS-AS"="36bc6f20-feed-422d-b2f2-7758e9708604"} #PATCH doesn't work on schedulingGroups yet :'( But PUT works!
             
        $nextWeekOfShifts | % { #Add 1 week's worth of shifts 
            $thisShift = $_
            if(@(6,0) -contains $(Get-Date $thisShift["shiftStart"]).DayOfWeek.value__){ #If the shift is at a weekend, use a different colour
                Write-Verbose "`tCreating weekend shift for [$($thisOffice["OfficeName"])] on [$($(Get-Date $thisShift["shiftStart"]).DayOfWeek)][$(Get-Date $thisShift["shiftStart"] -Format g)]"
                new-graphOpenShiftShared -tokenResponse $tokenResponseshiftBot -teamId $teamId -schedulingGroupId $thisSchedulingGroup.id -shiftName $("$($thisOffice["OfficeName"]) $($thisShift["shiftName"])") -shiftNotes $standardShiftNotes -availableSlots $thisOffice["OfficeDesks"] -startDateTime $thisShift["shiftStart"] -endDateTime $thisShift["shiftEnd"] -shiftColour $thisOffice["OfficeWeekendColour"] -MsAppActsAsUserId $msAppActsAsUserId -Verbose:$VerbosePreference
                }
            else{
                Write-Verbose "`tCreating weekday shift for [$($thisOffice["OfficeName"])] on [$($(Get-Date $thisShift["shiftStart"]).DayOfWeek)][$(Get-Date $thisShift["shiftStart"] -Format g)]"
                new-graphOpenShiftShared -tokenResponse $tokenResponseshiftBot -teamId $teamId -schedulingGroupId $thisSchedulingGroup.id -shiftName $("$($thisOffice["OfficeName"]) $($thisShift["shiftName"])") -shiftNotes $standardShiftNotes -availableSlots $thisOffice["OfficeDesks"] -startDateTime $thisShift["shiftStart"] -endDateTime $thisShift["shiftEnd"] -shiftColour $thisOffice["OfficeColour"] -MsAppActsAsUserId $msAppActsAsUserId -Verbose:$VerbosePreference
                }
            }
        }
    }