$teamBotDetails = import-encryptedCsv -pathToEncryptedCsv "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\teambotdetails.txt"
$resource = "https://graph.microsoft.com"

$ReqTokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    client_Id     = $teamBotDetails.ClientID
    Client_Secret = $teamBotDetails.Secret
    } 
$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody

$reportingPeriod = 7

$activityUserDetailReport = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(period='D$reportingPeriod')" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method GET
$tempFile = $("$env:TEMP\$([guid]::NewGuid().Guid).csv")
New-Item -Path $tempFile -ItemType File -Value $activityUserDetailReport 
$audReport = Import-Csv -Path $tempFile 
Remove-Item $tempFile

$activityCountsReport = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityCounts(period='D$reportingPeriod')" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method GET
$tempFile = $("$env:TEMP\$([guid]::NewGuid().Guid).csv")
New-Item -Path $tempFile -ItemType File -Value $activityCountsReport 
$acReport = Import-Csv -Path $tempFile 
Remove-Item $tempFile

$activityUserCountsReport = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserCounts(period='D$reportingPeriod')" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method GET
$tempFile = $("$env:TEMP\$([guid]::NewGuid().Guid).csv")
New-Item -Path $tempFile -ItemType File -Value $activityUserCountsReport 
$aucReport = Import-Csv -Path $tempFile 
Remove-Item $tempFile



$teamsUsersGroupId = "ec90dbd2-a1fe-4a43-9116-2e1553f6c43f"
$itTeamAllGroupId = "78616081-85b8-422d-b392-4070657d6cb9"
#$itTeamSite = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$itTeamAllGroupId/sites/root" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method GET
$itTeamSiteId = "anthesisllc.sharepoint.com,1ead4d00-2e2e-421e-9176-15287b5c22ce,d21ddf81-fcef-4e36-94e6-edd6fb487a31"
#$itTeamSiteLists = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$($itTeamSite.id)/lists" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method GET
$itTeamSiteMetricsListId = "42b2d4eb-6c47-4fc9-b1ae-a0cbd5310e9c"

$pilotGroups = @{
    TeamsPilotTeam_All = "64a86b83-e871-40f0-a5bd-6704f1e23a4e"
    All_Swe = "752529da-8a6f-48cb-86a7-95884600be1d"
    AdministrationTeam_Gbr = "19b1f332-1972-4031-8739-46905052b2d2"
    AdministrationTeam_NorthAmerica = "efc4361b-d410-4af8-939b-f957e9472858"
    AdministrationTeam_Fin = "1317f31c-cba1-48bb-862b-44700724453c"
    AdministrationTeam_Ita = "c015793a-6372-4639-8acc-e28d5c83bf2f"
    AnalystsTeam_Gbr = "c5d71ff2-a3f4-4089-98d8-8cebcf24c476"
    AnalystsTeam_Phl = "619f1830-d25b-46b4-a909-5b9fcbef1eff"
    EnergyTeam_Swe = "813f8c05-9b42-4237-a35a-3da8aab4b487"
    FinanceTeam_Fin = "f62253fe-b9d3-427f-9673-310dde6a4857"
    HrTeam_Fin = "6310cb12-9202-471b-bbfc-bd24c48a922a"
    ItTeam_Esp = "49f7523a-d453-447d-a895-4ca760754c9b"
    ItTeam_Gbr = "8d35013e-584a-40e6-84dd-3b5fe07802f2"
    SoftwareTeam_Gbr = "cec757d4-9870-4a7a-8618-58a1b0077da9"
    SoftwareTeam_Phl = "f1aeed5e-9471-4d0b-afee-086400410f65"
    StepTeam_All = "2fffdb8b-e332-4f90-a382-75c2f5d8a6cd"
    TransactionCorportaeServicesTeam_Fin = "dcda509e-87d5-473a-8b2b-a2f88d5d3763"
    WasteAndResourceSustainabilityTeam_Gbr = "32a59db2-6749-4f75-bcb8-a10dfc86a8a6"
    WGCollaborationImprovement_All = "80610abc-5290-4a90-b4af-23de5d0211f1"
    }
 <#
$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$teamsPilotUsersGroupId/transitiveMembers" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method GET
$teamsPilotObjects = $response.value
while (![string]::IsNullOrWhiteSpace($response.'@odata.nextLink')){
    $response = Invoke-RestMethod -Uri $response.'@odata.nextLink' -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method GET
    $teamsPilotObjects += $teamsUsers = $response.value
    }
$teamsPilotUsers = $teamsPilotObjects | ? {$_.'@odata.type' -eq '#microsoft.graph.user' -and $_.displayName -notmatch "Shared Mailbox"}

$teamsPilotUsers | % {
    $_ | Add-Member -MemberType NoteProperty -Name 'User Principal Name' -Value $_.userPrincipalName
    }
$comparison = Compare-Object -ReferenceObject $audReport -DifferenceObject $teamsPilotUsers -Property 'User Principal Name' -IncludeEqual -ExcludeDifferent -PassThru

$comparison | %{
    $thisPilotUser = $_
    $thisPostQuery = "{
          'fields': {
            'Title': '$($thisPilotUser.'User Principal Name')'
            ,'SnapshotDate': '$(get-date $thisPilotUser.'ï»¿Report Refresh Date' -Format yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z')'
            ,'ReportingPeriodInDays': $reportingPeriod
            ,'LastActivityDate': '$(if([string]::IsNullOrWhiteSpace($thisPilotUser.'Last Activity Date')){'1999-12-31T00:00:00Z'}else{$(get-date $thisPilotUser.'Last Activity Date' -Format yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z')})'
            ,'isDeleted': $($thisPilotUser.'Is Deleted'.ToLower())
            ,'TeamChatMessageCount': $($thisPilotUser.'Team Chat Message Count')
            ,'PrivateChatMessageCount': $($thisPilotUser.'Private Chat Message Count')
            ,'CallCount': $($thisPilotUser.'Call Count')
            ,'MeetingCount': $($thisPilotUser.'Meeting Count')
            ,'HasOtherAction': $(if($thisPilotUser.'Has Other Action' -eq "Yes"){"true"}else{"false"})
          }
        }"

    #$dummy = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$itTeamSiteId/lists/$itTeamSiteMetricsListId/items?expand=fields"  -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method GET
    try{$newReportRecord = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$itTeamSiteId/lists/$itTeamSiteMetricsListId/items" -Body $thisPostQuery -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Post -ErrorAction Stop}
    catch {$_;break}
    }
#>

$surveyRespondees = convertTo-arrayOfEmailAddresses "Maria.Hammar@anthesisgroup.com
Diana.Namoco@anthesisgroup.com
Richard.Santiago@anthesisgroup.com
Jonathan.Pont@anthesisgroup.com
Sara.Telahoun@anthesisgroup.com
irene.oliquino@anthesisgroup.com"

$pstnTesters = convertTo-arrayOfEmailAddresses "Elle.Wright@anthesisgroup.com
Elle.Wright@anthesisgroup.com
Elle.Wright@anthesisgroup.com
Jamie.Warmington@anthesisgroup.com
Jamie.Warmington@anthesisgroup.com
Jamie.Warmington@anthesisgroup.com
Charlie.Oliver@anthesisgroup.com
Hanna.Westling@anthesisgroup.com
Hanna.Westling@anthesisgroup.com
Agneta.Persson@anthesisgroup.com
Simone.Aplin@anthesisgroup.com
Mikko.Vuorela@anthesisgroup.com
James.MacPherson@anthesisgroup.com
Samantha.Mullender@anthesisgroup.com
Albert.Masnou@anthesisgroup.com
Ellen.Upton@anthesisgroup.com
Ian.Forrester@anthesisgroup.com
Agneta.Persson@anthesisgroup.com
Tecla.Castella@anthesisgroup.com
" 
$pstnTesters  = $pstnTesters | Sort-Object -Unique

$audReport | % {
    $inSurveyRespondees = Compare-Object -ReferenceObject $_.'User Principal Name' -DifferenceObject $surveyRespondees -IncludeEqual -ExcludeDifferent
    if($inSurveyRespondees){$surveyResponse = $true}else{$surveyResponse = $false}
    $_ | Add-Member -MemberType NoteProperty -Name 'RespondedToSurvey' -Value $surveyResponse
    $inPstnTesters = Compare-Object -ReferenceObject $_.'User Principal Name' -DifferenceObject $pstnTesters -IncludeEqual -ExcludeDifferent
    if($inPstnTesters){$testedPstn = $true}else{$testedPstn = $false}
    $_ | Add-Member -MemberType NoteProperty -Name 'TestedPstn' -Value $testedPstn
    }

$pilotGroupMembers = @{}
$pilotGroupMemberStats = @{}
$pilotGroups.Keys | % {
    $thisGroupId = $pilotGroups[$_]
    $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$thisGroupId/transitiveMembers" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method GET
    $theseObjects = $response.value
    while (![string]::IsNullOrWhiteSpace($response.'@odata.nextLink')){
        $response = Invoke-RestMethod -Uri $response.'@odata.nextLink' -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method GET
        $theseObjects += $response.value
        }
    $theseObjects = $theseObjects | ? {$_.'@odata.type' -eq '#microsoft.graph.user' -and $_.displayName -notmatch "Shared Mailbox"}
    $theseObjects | % {
        $_ | Add-Member -MemberType NoteProperty -Name 'User Principal Name' -Value $_.userPrincipalName -Force
        }
    $pilotGroupMembers.Add($_,$theseObjects)
    $pilotGroupMemberStats.Add($_,$(Compare-Object -ReferenceObject $audReport -DifferenceObject $($pilotGroupMembers[$_]) -Property 'User Principal Name' -IncludeEqual -ExcludeDifferent -PassThru))
    Write-Host -f Yellow $_
    $pilotGroupMemberStats[$_] | select "User Principal Name","Last Activity Date","Team Chat Message Count","Private Chat Message Count","Call Count","Meeting Count","RespondedToSurvey" | ft
    }



#Upload to SharePoint
$comparison = Compare-Object -ReferenceObject $audReport -DifferenceObject $pilotGroupMembers["TeamsPilotTeam_All"] -Property 'User Principal Name' -IncludeEqual -ExcludeDifferent -PassThru

$comparison | %{
    $thisPilotUser = $_
    $thisPostQuery = "{
          'fields': {
            'Title': '$($thisPilotUser.'User Principal Name')'
            ,'SnapshotDate': '$(get-date $thisPilotUser.'ï»¿Report Refresh Date' -Format yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z')'
            ,'ReportingPeriodInDays': $reportingPeriod
            ,'LastActivityDate': '$(if([string]::IsNullOrWhiteSpace($thisPilotUser.'Last Activity Date')){'1999-12-31T00:00:00Z'}else{$(get-date $thisPilotUser.'Last Activity Date' -Format yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z')})'
            ,'isDeleted': $($thisPilotUser.'Is Deleted'.ToLower())
            ,'TeamChatMessageCount': $($thisPilotUser.'Team Chat Message Count')
            ,'PrivateChatMessageCount': $($thisPilotUser.'Private Chat Message Count')
            ,'CallCount': $($thisPilotUser.'Call Count')
            ,'MeetingCount': $($thisPilotUser.'Meeting Count')
            ,'HasOtherAction': $(if($thisPilotUser.'Has Other Action' -eq "Yes"){"true"}else{"false"})
          }
        }"

    #$dummy = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$itTeamSiteId/lists/$itTeamSiteMetricsListId/items?expand=fields"  -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method GET
    try{$newReportRecord = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$itTeamSiteId/lists/$itTeamSiteMetricsListId/items" -Body $thisPostQuery -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Post -ErrorAction Stop}
    catch {$_;break}
    }

$pilotGroupMemberStats.Keys | % {
    $thisGroup = $_
    $pilotGroupMemberStats[$thisGroup] | % {
         $_ | Add-Member -MemberType NoteProperty -Name 'Team' -Value $($thisGroup)
         $_ | Add-Member -MemberType NoteProperty -Name 'Region' -Value $($thisGroup.Split("_")[1]) 
        }
    }


$pilotGroupStats = @()
$pilotGroups.Keys | Sort-Object | % {
    $i = 0; $totalIMs =0; $totalPosts=0;$totalCalls=0;$totalAttendees=0
    $pilotGroupMemberStats[$_] | % {
        $totalIMs = $totalIMs + $_.'Private Chat Message Count'
        $totalPosts = $totalPosts + $_.'Team Chat Message Count'
        $totalCalls = $totalCalls + $_.'Call Count'
        $totalAttendees = $totalAttendees +$_.'Meeting Count'
        $i++
        }
    Write-Host -f DarkYellow "MeanIMs=$($totalIMs/$i)`tMeanPosts=$($totalPosts/$i)`tMeanCalls=$($totalCalls/$i)`tMeanAttendees=$($totalAttendees/$i)`tMeanTeamsocity=$(($totalIMs+$totalPosts*2+$totalCalls*5+$totalAttendees*10)/$i)"
    Write-Host -f Yellow $_
    $pilotGroupStatsObject = New-Object psobject -Property @{
        "TeamName" = $_
        "MeanIMs" = $($totalIMs/$i)
        "MeanPosts" = $($totalPosts/$i)
        "MeanCalls" = $($totalCalls/$i)
        "MeanAttendees" = $($totalAttendees/$i)
        "MeanTeamsocity" =  $(($totalIMs+$totalPosts*2+$totalCalls*5+$totalAttendees*10)/$i)
        }
    $pilotGroupStats += $pilotGroupStatsObject
    $pilotGroupMemberStats[$_] | Sort-Object "User Principal Name" | select "User Principal Name","Last Activity Date","Team Chat Message Count","Private Chat Message Count","Call Count","Meeting Count","RespondedToSurvey","TestedPstn" | ft
    }

$pilotGroupStats | Sort-Object MeanTeamsocity -Descending | Select TeamName, MeanIMs, MeanPosts, MeanCalls, MeanAttendees, MeanTeamsocity | ft

