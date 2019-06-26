#Needs to be run as GoldMineLink (for elevated mailbox permissions)
#
# Script to read calendar events from FocalPoint Views and populate calendar entries in Outlook
#
# v1.1 Kev Maitland 7/5/15
# Revised SQL view to flatten multiple contiguous all-day entries into single multi-day events
# v1.2 Kev Maitland 18/1/17
# Adapted to send data to Anthesis' O365 platform
$logFileLocation = "C:\ScriptLogs\"
$transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_Transcript_$(Get-Date -Format "yyMMdd").log"
if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)){
    $fullLogPathAndName = $logFileLocation+"sync-focalPointToO365_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = $logFileLocation+"sync-focalPointToO365_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
else{
    $fullLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
$debugLog = "$env:USERPROFILE\Desktop\debugdump.log"
Start-Transcript $transcriptLogName -Append
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


$EWSServicePath = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'
#$EWSServicePath = '\\EX02\C$\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'
Import-Module $EWSServicePath

$sqlServer = "SQL01"
$sqlDbName = "SUSTAIN_LIVE"
$calendarSyncViewOutstanding = "SUS_VW_CUSTOM_CALENDAR_SYNC_CombinedSchedulerEntries_Outstanding_KM"
$calendarSyncViewAll = "SUS_VW_CUSTOM_CALENDAR_SYNC_CombinedSchedulerEntries_KM"
$calendarProcessingTable = "TS_CUSTOM_CALENDAR_SYNC"
$calendarProcessingTableId = "TS_CALENDAR_SYNC_PRI"
$calendarProcessingTableError = "Error"
$calendarProcessingTableRetry = "Retry"
$focalPointCategoryName = "From FocalPoint"
$projectsView = "SUS_VW_Projects_KM"
$bookingsView = "SUS_VW_Bookings_KM"
$ExchVer = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016
$ewsUrl = "https://outlook.office365.com/EWS/Exchange.asmx"
$maximumNumberOfSyncRetryAttempts = 13
$upnSMA = "sustainmailboxaccess@anthesisgroup.com"
#$passSMA = ConvertTo-SecureString -String '' -AsPlainText -Force | ConvertFrom-SecureString
$passSMA =  ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\SustainMailboxAccess.txt) 

#$daysToLookBehind = 30
#$daysToLookAhead = 30
#$from = (Get-Date -Format yyyy-MM-dd (get-date).AddDays(-$daysToLookBehind))
#$to = (Get-Date -Format yyyy-MM-dd (get-date).AddDays($daysToLookAhead))

#region functions
function Execute-SQLQueryOnSQLDB([string]$query, [string]$queryType) { 
  # NonQuery - Insert/Update/Delete query where no return data is required
    $sql = New-Object System.Data.SqlClient.SqlCommand
    $sql.Connection = $connDB
    $sql.CommandText = $query
    switch ($queryType){
        "NonQuery" {$sql.ExecuteNonQuery()}
        "Scalar" {$sql.ExecuteScalar()}
        "Reader" {    
            $oReader = $sql.ExecuteReader()
            $results = @()
            while ($oReader.Read()){
                $result = New-Object PSObject
                for ($i = 0; $oReader.FieldCount -gt $i; $i++){
                        $columnName = ($query.Replace(",","") -split '\s+')[$i+1]
                        if (1 -lt $columnName.Split(".").Length){$columnName = $columnName.Split(".")[1]} #Trim off any table names
                        $result | Add-Member NoteProperty $columnName $oReader[$i]
                        }
                 $results += $result
                }
            $oReader.Close()
            return $results
            }
        }
    }
function CatchNull([String]$x) {
   if ($x) { $x } else { -1 }
}
function CatchNull2($x,$returnIfNull) {
   if ($x -eq $null -or $x -eq [System.DBNull]::Value){$returnIfNull}
   else {$x}
}

function WriteOutputLogToSQL([double]$id, [string]$errorMessage, $previousNumberOfRetrys){
    if ("" -eq $entry.Retry) {Execute-SQLQueryOnSQLDB -query "INSERT INTO $calendarProcessingTable ($calendarProcessingTableId, $calendarProcessingTableError, $calendarProcessingTableRetry) VALUES ($id, `'$(CatchNull $errorMessage)`', $($(CatchNull $previousNumberOfRetrys)+1))" -queryType "NonQuery"}
    else {Execute-SQLQueryOnSQLDB -query "UPDATE $calendarProcessingTable SET $calendarProcessingTableError=`'$(CatchNull $errorMessage)`', $calendarProcessingTableRetry=$($(CatchNull $previousNumberOfRetrys)+1) WHERE $calendarProcessingTableId=$id" -queryType "NonQuery"}
    }
function SanitiseThatString([string]$dirtyString){
    $dirtyString.Replace("'", "''")
    }
function GetResourceBookingProbability([float]$eventId){
    Execute-SQLQueryOnSQLDB -query "SELECT [BGK_USER_CHAR1] AS Probability FROM [$bookingsView] INNER JOIN [$calendarSyncViewAll] ON [$calendarSyncViewAll].[TSCS_KEY] = [$bookingsView].[BKG_PRIMARY] WHERE [$calendarSyncViewAll].[TSCS_KEY_NAME] LIKE 'FP_Booking_ID' AND [$calendarSyncViewAll].[TSCS_KEY] = $eventId" -queryType "Scalar"
    }
function GetResourceBookingDuration([float]$eventId){
    Execute-SQLQueryOnSQLDB -query "SELECT [BGK_USER_NUM1] AS Duration FROM [$bookingsView] INNER JOIN [$calendarSyncViewAll] ON [$calendarSyncViewAll].[TSCS_KEY] = [$bookingsView].[BKG_PRIMARY] WHERE [$calendarSyncViewAll].[TSCS_KEY_NAME] LIKE 'FP_Booking_ID' AND [$calendarSyncViewAll].[TSCS_KEY] = $eventId" -queryType "Scalar"
    }
function GetProjectCodeFromProjectName([string]$projName){
    Execute-SQLQueryOnSQLDB -query "SELECT CH_CODE FROM $projectsView WHERE CH_NAME LIKE `'$(SanitiseThatString $projName)`'" -queryType "Scalar"
    }
#endregion

$connDB = New-Object System.Data.SqlClient.SqlConnection
$connDB.ConnectionString = "Server = $sqlServer; Database = $sqlDbName; Integrated Security = True" #This relies on the current user having the appropriate Login/Role Membership ont he DB
$connDB.Open()

$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($exchver)
$service.Credentials = New-Object System.Net.NetworkCredential($upnSMA,$passSMA)
$service.Url = $ewsUrl
#$service.AutodiscoverUrl($upnSMA,{$true})
#$service.Url = $ewsUrl #Set this again becasue something in the Sustain environment borks it.


$entriesToSync = Execute-SQLQueryOnSQLDB -query "SELECT 
      TSCS_PRIMARY
      ,TS_CALENDAR_SYNC_PRI
      ,TSCS_TYPE
      ,TSCS_STATUS
      ,TSCS_PREV_MAILBOX
      ,TSCS_NEW_MAILBOX
      ,TSCS_PREV_START
      ,TSCS_NEW_START
      ,TSCS_PREV_END
      ,TSCS_NEW_END
      ,TSCS_SUBJECT
      ,TSCS_LOCATION
      ,TSCS_ALLDAY
      ,TSCS_REMINDER
      ,TSCS_COLOUR
      ,TSCS_KEY_NAME
      ,TSCS_KEY
      ,TSCS_USERID
      ,TSCS_TIMESTAMP
      ,TSCS_ERROR
      ,TSCS_RETRY 
      ,Error
      ,Retry FROM $calendarSyncViewOutstanding" -queryType "Reader"
#      ,Retry FROM $calendarSyncViewOutstanding WHERE ((TSCS_PREV_START >= '$from' OR TSCS_NEW_START >= '$from') AND (TSCS_PREV_START <= '$to' OR TSCS_NEW_START <= '$to'))" -queryType "Reader"
foreach ($entry in $entriesToSync){
    $spongesInThePatient = ""
    if ($maximumNumberOfSyncRetryAttempts -gt $(CatchNull $entry.Retry) -and (((CatchNull2 $entry.TSCS_NEW_START -returnIfNull "1900-01-01") -gt "2018-01-01") -or ((CatchNull2 $entry.TSCS_PREV_START -returnIfNull "1900-01-01") -gt "2018-01-01"))){
        $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $(if("" -ne $entry.TSCS_PREV_MAILBOX){$entry.TSCS_PREV_MAILBOX}else{$entry.TSCS_NEW_MAILBOX}).replace("sustain.co.uk","anthesisgroup.com"))

        switch ($entry.TSCS_TYPE)
            {
            'N' {#New appointment
                Write-Host -ForegroundColor Yellow "Creating NEW $($entry.TSCS_SUBJECT) appointment $($entry.TSCS_PRIMARY) for $($entry.TSCS_NEW_MAILBOX) on $($entry.TSCS_NEW_START)"
                $appointment = New-Object Microsoft.Exchange.WebServices.Data.Appointment -ArgumentList $service
                $appointment.Start = $entry.TSCS_NEW_START
                $appointment.End = $entry.TSCS_NEW_END
                $appointment.Categories.Add($focalPointCategoryName)
                switch ($entry.TSCS_KEY_NAME){
                    "FP_Absence_ID" {
                        $appointment.Categories.Add("Holiday")
                        $appointment.Subject = $entry.TSCS_SUBJECT
                        $appointment.IsAllDayEvent = $true
                        $appointment.IsReminderSet = $false
                        }
                    "FP_Booking_ID" {
                        $projCode = GetProjectCodeFromProjectName -projName $entry.TSCS_SUBJECT
                        $probability = GetResourceBookingProbability -eventId $entry.TSCS_KEY
                        [string]$reservationDuration = GetResourceBookingDuration -eventId $entry.TSCS_KEY
                        if ($reservationDuration -eq "1"){$reservationDuration = $reservationDuration + " day"} else {$reservationDuration = $reservationDuration + " days"}
                        $duration = $($($entry.TSCS_NEW_END-$entry.TSCS_NEW_START).TotalMinutes) / 60
                        $appointment.Subject = "$reservationDuration, $projCode $($entry.TSCS_SUBJECT) [$($entry.TSCS_KEY)]"
                        $appointment.IsAllDayEvent = $true
                        $appointment.Categories.Add("Resourcing")
                        $appointment.Body = $entry.TSCS_TEXT
                        $appointment.LegacyFreeBusyStatus = [Microsoft.Exchange.WebServices.Data.LegacyFreeBusyStatus]::Tentative
                        $appointment.ReminderDueBy = $(get-date($entry.TSCS_NEW_START.ToShortDateString())).AddHours(8) #Set alarm for 8AM on the day
                        $appointment.IsReminderSet = $false
                        #$appointment.ReminderMinutesBeforeStart =
                        if (100 -eq $probability){$appointment.Categories.Add("Important")} #Make entry show as Red if confirmed booking
                        else {$appointment.Categories.Add("Admin")} #Make entry show as Orange if tentative booking
                        }
                    default {
                        $appointment.Subject = "Unknown Appointment Type from FocalPoint"
                        $appointment.Categories.Add("From FocalPoint")
                        $appointment.Body = $entry.TSCS_SUBJECT
                        $appointment.IsAllDayEvent = $true
                        }
                    }
                $spongesInThePatient = $appointment.Save([Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToAllAndSaveCopy)
                WriteOutputLogToSQL -id $entry.TSCS_PRIMARY -errorMessage $spongesInThePatient -previousNumberOfRetrys $entry.Retry
                }
            'E' {#Edit existing appointment
                $spongesInThePatient = "Potential appointments found, but none with the Category `"$focalPointCategoryName`" and beginning $($($entry.TSCS_SUBJECT -split " ")[0])"
                Write-Host -ForegroundColor Yellow "Searching for appointment to EDIT for $($entry.TSCS_NEW_MAILBOX) on $($entry.TSCS_NEW_START)"
                $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$entry.TSCS_NEW_MAILBOX)
                $CalendarFolder = [Microsoft.Exchange.WebServices.Data.CalendarFolder]::Bind($service,$folderid)
                $cvCalendarview = new-object Microsoft.Exchange.WebServices.Data.CalendarView($entry.TSCS_PREV_START,$entry.TSCS_PREV_START,2000)
                $cvCalendarview.PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
                $frCalendarResult = $CalendarFolder.FindAppointments($cvCalendarview)
                if($null -eq $frCalendarResult){$spongesInThePatient = "No potential appointments found that match the criteria Start:$($entry.TSCS_PREV_START)"}
                switch ($entry.TSCS_KEY_NAME){
                    'FP_Absence_ID' {
                        $matchFound = $false
                        foreach($possibleMatch in $frCalendarResult.Items){
                            if(($possibleMatch.Categories -contains $focalPointCategoryName) -and ($($possibleMatch.Subject -split " ")[0] -eq $($entry.TSCS_SUBJECT +" dummyTextToEnsureMoreThanWordIsFound" -split " ")[0])){
                                Write-Host -ForegroundColor DarkYellow "Updating EXISTING ABSENCE $($entry.TSCS_SUBJECT) appointment $($entry.TSCS_PRIMARY) for $($entry.TSCS_NEW_MAILBOX) on $($entry.TSCS_PREV_START)"
                                $possibleMatch.Subject = $entry.TSCS_SUBJECT
                                $possibleMatch.Start = $entry.TSCS_NEW_START
                                $possibleMatch.End = $entry.TSCS_NEW_END
                                $spongesInThePatient = $possibleMatch.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
                                $matchFound = $true
                                }
                            }
                        if(!$matchFound){
                            Write-Host -ForegroundColor DarkYellow "No EXISTING ABSENCE $($entry.TSCS_SUBJECT) appointment found for $($entry.TSCS_NEW_MAILBOX) on $($entry.TSCS_PREV_START) - Creating new entry"
                            $appointment = New-Object Microsoft.Exchange.WebServices.Data.Appointment -ArgumentList $service
                            $appointment.Start = $entry.TSCS_NEW_START
                            $appointment.End = $entry.TSCS_NEW_END
                            $appointment.Categories.Add($focalPointCategoryName)
                            $appointment.Categories.Add("Holiday")
                            $appointment.Subject = $entry.TSCS_SUBJECT
                            $appointment.IsAllDayEvent = $true
                            $appointment.IsReminderSet = $false
                            $spongesInThePatient = $appointment.Save([Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToAllAndSaveCopy)
                            }
                        }
                    'FP_Booking_ID'{
                        foreach($possibleMatch in $frCalendarResult.Items){
                            $projCode = GetProjectCodeFromProjectName -projName $entry.TSCS_SUBJECT
                            #if(($possibleMatch.Categories -contains $focalPointCategoryName) -and ($($possibleMatch.Subject -split "h,")[0] -eq $prevDuration) -and ($($possibleMatch.Subject -split "%, ")[1] -eq $projCode)){
                            if($($possibleMatch.Subject -split " ")[$($possibleMatch.Subject -split " ").Count -1] -eq "[$($entry.TSCS_KEY)]"){ #Look for the last word in the Subject and compare it to the $entry.TSCS_KEY value
                                Write-Host -ForegroundColor DarkYellow "Updating EXISTING BOOKING $($entry.TSCS_SUBJECT) appointment $($entry.TSCS_PRIMARY) for $($entry.TSCS_NEW_MAILBOX) on $($entry.TSCS_PREV_START)"
                                $newDuration = $($($entry.TSCS_NEW_END-$entry.TSCS_NEW_START).TotalMinutes) / 60
                                $probability = GetResourceBookingProbability -eventId $entry.TSCS_KEY
                                [string]$reservationDuration = GetResourceBookingDuration -eventId $entry.TSCS_KEY
                                if ($reservationDuration -eq "1"){$reservationDuration = $reservationDuration + " day"} else {$reservationDuration = $reservationDuration + " days"}
                                $possibleMatch.Subject = "$reservationDuration, $projCode $($entry.TSCS_SUBJECT) [$($entry.TSCS_KEY)]"
                                $possibleMatch.Start = $entry.TSCS_NEW_START
                                $possibleMatch.End = $entry.TSCS_NEW_END
                                $possibleMatch.LegacyFreeBusyStatus = [Microsoft.Exchange.WebServices.Data.LegacyFreeBusyStatus]::Tentative
                                $possibleMatch.ReminderDueBy = $(get-date($entry.TSCS_NEW_START.ToShortDateString())).AddHours(8) #Set alarm for 8AM on the day
                                $possibleMatch.IsReminderSet = $true
                                if (100 -eq $probability){
                                    if ($possibleMatch.Categories -contains "Admin"){$possibleMatch.Categories.Remove("Admin")}
                                    if ($possibleMatch.Categories -notcontains "Important"){$possibleMatch.Categories.Add("Important")}
                                    }
                                else {
                                    if ($possibleMatch.Categories -contains "Important"){$possibleMatch.Categories.Remove("Important")}
                                    if ($possibleMatch.Categories -notcontains "Admin"){$possibleMatch.Categories.Add("Admin")}
                                    }
                                $spongesInThePatient = $possibleMatch.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
                                }
                            }
                        }
                    default {}
                    }

                WriteOutputLogToSQL -id $entry.TSCS_PRIMARY -errorMessage $spongesInThePatient -previousNumberOfRetrys $entry.Retry
                }
            'D' {#Delete existing appointment}
                $spongesInThePatient = "Potential appointments found, but none with the Category `"$focalPointCategoryName`""
                Write-Host -ForegroundColor Yellow "Searching for appointment to DELETE for $($entry.TSCS_PREV_MAILBOX) on $($entry.TSCS_PREV_START)"
                $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$entry.TSCS_PREV_MAILBOX)
                $CalendarFolder = [Microsoft.Exchange.WebServices.Data.CalendarFolder]::Bind($service,$folderid)
                $cvCalendarview = new-object Microsoft.Exchange.WebServices.Data.CalendarView($entry.TSCS_PREV_START,$entry.TSCS_PREV_START,2000)
                $cvCalendarview.PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
                $frCalendarResult = $CalendarFolder.FindAppointments($cvCalendarview)
                if($null -eq $frCalendarResult){$spongesInThePatient = "No potential appointments found that match the criteria Start:$($entry.TSCS_PREV_START)"}

                switch ($entry.TSCS_KEY_NAME){
                    'FP_Absence_ID' {
                        foreach($possibleMatch in $frCalendarResult.Items){
                            if(($possibleMatch.Categories -contains $focalPointCategoryName) -and ($($possibleMatch.Subject -split " ")[0] -eq $($entry.TSCS_SUBJECT +" dummyTextToEnsureMoreThanWordIsFound" -split " ")[0])){
                                 Write-Host -ForegroundColor DarkYellow "DELETING appointment $($entry.TSCS_PRIMARY) for $($entry.TSCS_PREV_MAILBOX) on $($entry.TSCS_PREV_START)"
                                $spongesInThePatient = $possibleMatch.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)
                                }
                            }
                        }
                    'FP_Booking_ID'{
                        foreach($possibleMatch in $frCalendarResult.Items){
                            $projCode = GetProjectCodeFromProjectName -projName $entry.TSCS_SUBJECT
                            if($($possibleMatch.Subject -split " ")[$($possibleMatch.Subject -split " ").Count -1] -eq "[$($entry.TSCS_KEY)]"){ #Look for the last word in the Subject and compare it to the $entry.TSCS_KEY value
                                 Write-Host -ForegroundColor DarkYellow "DELETING appointment $($entry.TSCS_PRIMARY) for $($entry.TSCS_PREV_MAILBOX) on $($entry.TSCS_PREV_START)"
                                $spongesInThePatient = $possibleMatch.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)
                                }
                            }
                        }
                    default {}
                    }

                WriteOutputLogToSQL -id $entry.TSCS_PRIMARY -errorMessage $spongesInThePatient -previousNumberOfRetrys $entry.Retry
                }
        default{
            #Find some automated method of punching the AccessGroup's developers in the face over the internet
            }
        }

    }
}

$connDB.Close()

Stop-Transcript