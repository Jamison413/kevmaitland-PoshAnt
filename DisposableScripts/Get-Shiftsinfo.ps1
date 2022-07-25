#Connect to Shifts
$teamBotDetails = get-graphAppClientCredentials -appName TeamsBot
$tokenResponseTeams = get-graphTokenResponse -aadAppCreds $teamBotDetails
$shiftBotDetails = get-graphAppClientCredentials -appName ShiftBot
$tokenResponseShiftBot = get-graphTokenResponse -aadAppCreds $shiftBotDetails

$teamId = "549dd0d0-251f-4c23-893e-9d0c31c2dc13" #All (GBR)
$groupBotGuid = "00aa81e4-2e8f-4170-bc24-843b917fd7cf"
$msAppActsAsUserId = "00aa81e4-2e8f-4170-bc24-843b917fd7cf" #GroupBot

#Get all Open Shifts for all offices for all time
$AllShiftsLikeEver = get-graphShiftOpenShifts -tokenResponse $tokenResponseShiftBot -teamId $teamId -MsAppActsAsUserId $msAppActsAsUserId -Verbose


#Filter for London for the week of 27th June-1st July 2022
$TargetAMshifts = $AllShiftsLikeEver.Where({($_.sharedOpenShift.displayName -eq "GBR-London Morning") -and ($_.sharedOpenShift.startDateTime -ge "2022-06-27T08:00:00Z")})
$TargetPMShifts = $AllShiftsLikeEver.Where({($_.sharedOpenShift.displayName -eq "GBR-London Afternoon") -and ($_.sharedOpenShift.startDateTime -ge "2022-06-27T08:00:00Z")})


#Get all Graph users
$users = get-graphUsers -tokenResponse $tokenResponseTeams -filterLicensedUsers -filterUsageLocation GB -Verbose
#Get all user shifts
$AllShiftusers = get-graphShiftUserShifts -tokenResponse $tokenResponseShiftBot -teamId $teamId -MsAppActsAsUserId $msAppActsAsUserId -Verbose

#Filter for London for the week of 27th June-1st July 2022 (this time for users)
$LonduserAMShift = $AllShiftusers.Where({($_.sharedShift.displayName -eq "GBR-London Morning") -and ($_.sharedShift.startDateTime -ge "2022-06-27T08:00:00Z")})
$LonduserPMShift = $AllShiftusers.Where({($_.sharedShift.displayName -eq "GBR-London Afternoon") -and ($_.sharedShift.startDateTime -ge "2022-06-27T08:00:00Z")})

$MdataToExport = @()
#Iterate through each London Shift, update each object with friendly user email/upn which you find from the ID
foreach($Shift in $LonduserAMShift){
    #this quieries users variable ID property to look for matching user ID on shift object
    $useremail = $users | Where-Object -Property "id" -eq $shift.userId
    #Adding upn to shift object so we can see who they are in human terms
    Add-Member -InputObject $Shift -MemberType NoteProperty -Name "UPN" -Value $useremail.userPrincipalName -Force
    $minimalShiftInformation = [PSCustomObject]@{
        ShiftName     = $shift.sharedShift.displayName
        upn = $userEmail.userprincipalname

        Time = $shift.sharedShift.StartDateTime
    }
    $MdataToExport  += $minimalShiftInformation  
} #export csv and send to whomever
$MdataToExport | Export-Csv -Path C:\Users\Andrew.Ost\Desktop\CSVs\LondonMorningShifts.csv -NoTypeInformation

$AdataToExport = @()
foreach($Shift in $LonduserPMShift){
    #this quieries users variable ID property to look for matching user ID on shift object
    $useremail = $users | Where-Object -Property "id" -eq $shift.userId
    #Adding upn to shift object so we can see who they are in human terms
    Add-Member -InputObject $Shift -MemberType NoteProperty -Name "UPN" -Value $useremail.userPrincipalName -Force
    $minimalShiftInformation = [PSCustomObject]@{
        ShiftName     = $shift.sharedShift.displayName
        upn = $userEmail.userprincipalname

        Time = $shift.sharedShift.StartDateTime
    }
    $AdataToExport  += $minimalShiftInformation  
} #export csv and send to whomever
$AdataToExport | Export-Csv -Path C:\Users\Andrew.Ost\Desktop\CSVs\LondonAfternoonShifts.csv -NoTypeInformation



#reading 

#Example of printing result to screen for sanity check (everything in result brought back)
$Targetopenshifts.sharedOpenShift.startDateTime
#Example of the above but targeting the first response for question
$Targetopenshifts[0].sharedOpenShift.startDateTime

####OpenShifts (Office)
https://docs.microsoft.com/en-us/graph/api/resources/openshift?view=graph-rest-1.0

####Shifts (User)
https://docs.microsoft.com/en-us/graph/api/resources/shift?view=graph-rest-1.0

#add-member example
$object = get-graphUsers -tokenResponse $tokenResponseTeams -filterUpns "emily.pressey@anthesisgroup.com" -Verbose
Add-Member -InputObject $object -MemberType NoteProperty -Name "Curry" -Value "Korma (vegan)" -Force