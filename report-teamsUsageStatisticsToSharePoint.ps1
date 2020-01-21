$teamBotDetails = Import-Csv "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\teambotdetails.txt"
$resource = "https://graph.microsoft.com"
$tenantId = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.TenantId)
$clientId = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.ClientID)
$redirect = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.Redirect)
$secret   = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.Secret)

$ReqTokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    client_Id     = $clientID
    Client_Secret = $secret
    } 
$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody

$reportingPeriod = 7

$activityUserDetailReport = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(period='D$reportingPeriod')" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method GET
$tempFile = $("$env:TEMP\$([guid]::NewGuid().Guid).csv")
New-Item -Path $tempFile -ItemType File -Value $activityUserDetailReport 
$audReport = Import-Csv -Path $tempFile 
Remove-Item $tempFile

$activityCountsReport = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityCounts(period='D30')" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method GET
$tempFile = $("$env:TEMP\$([guid]::NewGuid().Guid).csv")
New-Item -Path $tempFile -ItemType File -Value $activityCountsReport 
$acReport = Import-Csv -Path $tempFile 
Remove-Item $tempFile

$activityUserCountsReport = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserCounts(period='D30')" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method GET
$tempFile = $("$env:TEMP\$([guid]::NewGuid().Guid).csv")
New-Item -Path $tempFile -ItemType File -Value $activityUserCountsReport 
$aucReport = Import-Csv -Path $tempFile 
Remove-Item $tempFile



$teamsUsersGroupId = "ec90dbd2-a1fe-4a43-9116-2e1553f6c43f"
$teamsPilotUsersGroupId = "64a86b83-e871-40f0-a5bd-6704f1e23a4e"
$itTeamAllGroupId = "78616081-85b8-422d-b392-4070657d6cb9"
#$itTeamSite = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$itTeamAllGroupId/sites/root" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method GET
$itTeamSiteId = "anthesisllc.sharepoint.com,1ead4d00-2e2e-421e-9176-15287b5c22ce,d21ddf81-fcef-4e36-94e6-edd6fb487a31"
#$itTeamSiteLists = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$($itTeamSite.id)/lists" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method GET
$itTeamSiteMetricsListId = "42b2d4eb-6c47-4fc9-b1ae-a0cbd5310e9c"
Get-UnifiedGroup "IT Team (All)" | fl

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


