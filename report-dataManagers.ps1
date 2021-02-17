#----- To do -----#

# automated expired manager removal


#Try{
#We need Groupbot to manage Mail-Enabled Distribution Group membership (still unavailable via Graph [https://microsoftgraph.uservoice.com/forums/920506-microsoft-graph-feature-requests/suggestions/39551191-add-an-endpoint-to-allow-managing-mail-enabled-sec])
$groupAdmin = "groupbot@anthesisgroup.com"
$groupAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\GroupBot.txt) 
$exoCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $groupAdmin, $groupAdminPass
connect-ToExo -credential $exoCreds

#We need Groupbot to access the User Training Records stored in SharePoint via PnP (Graph doesn't have good enough ListItem functionality yet [https://microsoftgraph.uservoice.com/forums/920506-microsoft-graph-feature-requests/suggestions/40175989-standardise-the-returned-data-for-single-and-multi])
$sharePointAdmin = "kimblebot@anthesisgroup.com"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\KimbleBot.txt) 
$sharePointCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass

$teamBotDetails = get-graphAppClientCredentials -appName TeamsBot
$tokenResponse = get-graphTokenResponse -aadAppCreds $teamBotDetails


#Reporting
$ITReporting = "d95c1342-cdfb-4a16-9881-4c2d6154e5a1"
#$reportsyncdataManagers = update-graphListItem -tokenResponse $tokenResponse -graphSiteId "1ead4d00-2e2e-421e-9176-15287b5c22ce,d21ddf81-fcef-4e36-94e6-edd6fb487a31" -listName $ITReporting -listitemId 4 -fieldHash @{"Status" = "Ok";"Notes" = ""}
#$lastUnauthorisedManagerReport = get-graphListItems -tokenResponse $tokenResponse -graphSiteId "1ead4d00-2e2e-421e-9176-15287b5c22ce,d21ddf81-fcef-4e36-94e6-edd6fb487a31" -listId $UnauthorisedManagersList -expandAllFields
#}
#Catch{
#Write-Error "Data Manager's might not have been processed - we found an error in the initial connections"
#$reportsyncdataManagers = update-graphListItem -tokenResponse $tokenResponse -graphSiteId "1ead4d00-2e2e-421e-9176-15287b5c22ce,d21ddf81-fcef-4e36-94e6-edd6fb487a31" -listId $ITReporting -listitemId 4 -fieldHash @{"Status" = "Error - Connection Issue";"Notes" = $([string]$error.ErrorDetails.Message)}
#}



###################################################################
#                                                                 #
#    Update subgroups to add to [Data Managers - Current (All)    #
#                                                                 #
###################################################################

#Make sure we've got all Data Manager Subgroups added into [Data Managers - Current (All)]
$allIndividualDataManagerGroupIds = $(get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse).anthesisgroup_UGSync.dataManagerGroupId
$allIndividualDataManagerGroupIds = $allIndividualDataManagerGroupIds | Sort-Object -Unique | ? {![string]::IsNullOrWhiteSpace($_)}
$allIndividualDataManagerGroups = @($false)*$allIndividualDataManagerGroupIds.Count #Initialise an array of the correct length
for($i=0; $i -lt $allIndividualDataManagerGroupIds.Count;$i++){
    $allIndividualDataManagerGroups[$i] = New-Object -TypeName PSCustomObject -Property @{"id"=$allIndividualDataManagerGroupIds[$i]}
    }
$allDataManagersGroup = get-graphGroups -tokenResponse $tokenResponse -filterUpn datamanagers-current@anthesisgroup.com
$allDataManagerSubGroups = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupUpn datamanagers-current@anthesisgroup.com -memberType Members | ? {$_.'@odata.type' -eq "#microsoft.graph.group"}
$mismatchedDataManagerGroups = Compare-Object -ReferenceObject $allIndividualDataManagerGroups -DifferenceObject $allDataManagerSubGroups -Property id -PassThru
$mismatchedDataManagerGroups | ? {$_.SideIndicator -eq "<="} | % {
    $groupToAdd = get-graphGroups -tokenResponse $tokenResponse -filterId $_.Id
    Write-Host -f Yellow "Adding [$($groupToAdd.DisplayName)] to [Data Managers - Current (All)]"
    Add-DistributionGroupMember -Identity $allDataManagersGroup.id -Member $_.id -Confirm:$false -BypassSecurityGroupManagerCheck:$true
    $groupChangesWereMade = $true
    }
if($groupChangesWereMade){#Refresh $allDataManagerSubGroups if it's changed
    $allDataManagerSubGroups = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupUpn datamanagers-current@anthesisgroup.com -memberType Members | ? {$_.'@odata.type' -eq "#microsoft.graph.group"}
    }

#Get the users and training dates for all authorised Data Managers
#$trainingRecords = get-graphListItems -tokenResponse $tokenResponse -serverRelativeSiteUrl "/sites/Resources-HR" -Verbose -ListName "User Training Records" -expandAllFields #This requires two additional calls to get the IDs for the Site & List
#$trainingRecords = get-graphListItems -tokenResponse $tokenResponse -graphSiteId "anthesisllc.sharepoint.com,8658f988-7c7d-4a35-a4db-8baea3b54ca5,5786d001-5418-4f96-88fc-9e4e9e5922d8" -Verbose -listId "ca4d708a-57e9-46b6-8b9f-a2d82bb94d24" -expandAllFields #This is more efficient as it requires fewer API calls to get the Ids
#$dataManagerTrainingRecords = $trainingRecords | ? {$_.fields.Training_x0020_session.Label -eq "Data Manager"}
#Graph doesn't expose SharePoint Users correctly, so it's much simpler to use PnP where PeoplePickers and Managed MetaData is used

#$validTrainedDataManagers > The "real" list of users who have valid Data Manager training
#$allAuthorisedDataManagers > Who had valid Data Manager training last time this script was run
#$allCurrentDataManagers > All users who are currently Data Managers, regardless of whether they are trained
#$newAuthorisedDataManagers = $validTrainedDataManagers - $allAuthorisedDataManagers
#$unauthorisedDataManagers = $allCurrentDataManagers - $validTrainedDataManagers ***
#$authorisedButUnassignedDataManagers = $validTrainedDataManagers -  $allCurrentDataManagers

###################################################################
#                                                                 #
#                     Report Data Managers                        #
#                                                                 #
###################################################################


Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/sites/Resources-HR" -Credentials $sharePointCreds
$dataManagerTrainingRecords = Get-PnPListItem -List "User Training Records" -Query "<View><Query><Where><Eq><FieldRef Name='Training_x0020_session' Label='True'/><Value Type='String'>Data Manager</Value></Eq></Where></Query></View>" #Get the Data Manager Training records
$dataManagerTrainingRecords | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name mail -Value $_.FieldValues.User.Email} #Add this property so we can compare-object with Graph Users later
$dataManagerTrainingRecords = $dataManagerTrainingRecords | Sort-Object {$_.FieldValues.User.Email}, {$_.FieldValues.Date_x0020_of_x0020_training} -Descending #Sort them by User, then by Training Date
$mostRecentTrainingRecords = @($false) * $($dataManagerTrainingRecords.FieldValues.User.Email | sort -Unique).Count #Build an array to hold the most recent training event for each user
$j=0 #Iterate through all the training events and copy the most recent one to $mostRecentTrainingRecords
for($i=0;$i -lt $dataManagerTrainingRecords.Count;$i++){
    if($dataManagerTrainingRecords[$i].FieldValues.User.Email -ne $lastEmail){
        $mostRecentTrainingRecords[$j] = $dataManagerTrainingRecords[$i]
        $j++
        }
    $lastEmail = $dataManagerTrainingRecords[$i].FieldValues.User.Email
    }
$validTrainedDataManagers = $mostRecentTrainingRecords | ? {$_.FieldValues.Date_x0020_of_x0020_training -ge $(Get-Date).AddYears(-1)}
$expiringSoonDataManagers = $mostRecentTrainingRecords | ? {$_.FieldValues.Date_x0020_of_x0020_training -ge $(Get-Date).AddYears(-1) -and $_.FieldValues.Date_x0020_of_x0020_training -lt $(Get-Date).AddMonths(-10)}
$c = $mostRecentTrainingRecords | ? {$_.FieldValues.Date_x0020_of_x0020_training -lt $(Get-Date).AddYears(-1)}

#Get the members of the relevant AAD groups
$authorisedDataManagerGroup = get-graphGroups -tokenResponse $tokenResponse -filterUpn datamanagers@anthesisgroup.com
$currentDataManagerGroup =  get-graphGroups -tokenResponse $tokenResponse -filterUpn datamanagers-current@anthesisgroup.com
$allCurrentDataManagers = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $currentDataManagerGroup.id -memberType TransitiveMembers -returnOnlyUsers
$allAuthorisedDataManagers = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $authorisedDataManagerGroup.id -memberType Members -returnOnlyUsers




###################################################################
#                                                                 #
#                  Find Who Needs to Do What                      #
#                                                                 #
###################################################################

#Compare who is currently in the [Data Managers - Authorised (All)] group with who *should* be in there, and make any changes

#Remove anyone who's training has lapsed from Data Managers (Authorised) - we'll remove them from the individual teams later once we've got the full list of unauthorised Data Managers

# Valid Trained Managers are found from the training record list - anyone that has had training within the last year
# All Authorised Data Managers are from the [Data Managers - Authorised (All)] group

if($validTrainedDataManagers -eq $null){$validTrainedDataManagers = @()}
if($allAuthorisedDataManagers -eq $null){$allAuthorisedDataManagers = @()}
#Compare the two and find anyone who is in the security group but hasn't got a record for training in the last year
$mismatchedAuthorisedDataManagers = Compare-Object -ReferenceObject $validTrainedDataManagers -DifferenceObject $allAuthorisedDataManagers -Property "mail" -IncludeEqual -PassThru

#Add anyone who has had training recently, but not in the  Data Managers (Authorised) group
$newAuthorisedDataManagers = $mismatchedAuthorisedDataManagers | ? {$_.SideIndicator -eq "<="} 
$newAuthorisedDataManagers | % { #Add anyone new
    Write-Verbose "Adding [$( $_.mail)] to [$($authorisedDataManagerGroup.displayName)]"
    Add-DistributionGroupMember -Identity $authorisedDataManagerGroup.id -Member $_.mail -Confirm:$false -BypassSecurityGroupManagerCheck:$true 
    #$userChangesWereMade = $true
    }

$expiredDataManagers = $mismatchedAuthorisedDataManagers | ? {$_.SideIndicator -eq "=>"}
$expiredDataManagers | % { 
    Write-Verbose "Removing [$( $_.mail)] from [$($authorisedDataManagerGroup.displayName)]"
    Write-host "Removing [$( $_.mail)] from [$($authorisedDataManagerGroup.displayName)]"
    Remove-DistributionGroupMember -Identity $authorisedDataManagerGroup.id -Member $_.mail -Confirm:$false -BypassSecurityGroupManagerCheck:$true 
    #$userChangesWereMade = $true
    }


#if($userChangesWereMade){#Refresh $allAuthorisedDataManagers if it's changed  #We don;t have to wait for this to refresh if we just re-use $validTrainedDataManagers
#    $oldCount = $allAuthorisedDataManagers.Count
#    do{
#        Write-Host -ForegroundColor Yellow "Waiting for EXO to sync to AAD"
#        Start-Sleep -Seconds 60
#        $allAuthorisedDataManagers = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $authorisedDataManagerGroup.id -memberType Members -returnOnlyUsers
#        }
#    while($allAuthorisedDataManagers.Count -eq $oldCount)
#    }

$tokenResponse = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponse -renewTokenExpiringInSeconds 600 -aadAppCreds $teamBotDetails

#Create a list of who-owns-what and what-is-owned-by-who
$whoOwnsWhatHash = @{}
$whatisOwnedByWhoHash = @{}
$allDataManagerSubGroups | % {
    $tokenResponse = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponse -renewTokenExpiringInSeconds 300 -aadAppCreds $teamBotDetails #Just in case this ever takes a really long time to complete
    $thisGroup = $_
    #what-is-owned-by-who
    $whatisOwnedByWhoHash.Add($thisGroup.id,@())
    #Who-owns-what
    $theseDataMangers = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $thisGroup.id -memberType Members -returnOnlyUsers 
    $theseDataMangers | % {
        if(!$whoOwnsWhatHash.ContainsKey($_.mail)){$whoOwnsWhatHash.Add($_.mail,@())}
        $whoOwnsWhatHash[$_.mail]+= ,@($thisGroup.displayName,$thisGroup.id) #The , between += and @() prevents the array from becoming unrolled as it is added
        #what-is-owned-by-who part2
        $whatisOwnedByWhoHash[$thisGroup.id] += ,$($_.displayName,$_.mail) #The , between += and @() prevents the array from becoming unrolled as it is added
        }
    }

#As a sanity-check, $whoOwnsWhatHash.Keys should now match $allCurrentDataManagers (they are both derived from enumerating the members of all Data Mananger Subgroups)
#$whoOwnsWhatHash.Keys | % {
#    [array]$whoOwnsWhatTest += New-Object psobject -Property @{mail=$_}
#    }
#Compare-Object -ReferenceObject $whoOwnsWhatTest -DifferenceObject $allCurrentDataManagers -Property mail #This should output nothing if everything is correct




#Find unauthorised data managers

#Find who hasn't completed Data Manager training in the past year, but is still currently a Data Manager
$mismatchedDataManagers = Compare-Object -ReferenceObject $allCurrentDataManagers -DifferenceObject $validTrainedDataManagers -Property mail -PassThru -IncludeEqual

$unauthorisedDataManagers = $mismatchedDataManagers  | ? {$_.SideIndicator -eq "<="}
#get rid of internal IT accounts
$unauthorisedDataManagers = $unauthorisedDataManagers.Where({$_.mail -ne "groupbot@anthesisgroup.com"})
$unauthorisedDataManagers = $unauthorisedDataManagers.Where({($_.mail -ne "t0-kevin.maitland@anthesisgroup.com") -and ($_.mail -ne "t1-emily.pressey@anthesisgroup.com") -and ($_.mail -ne "t1-andrew.ost@anthesisgroup.com")})

$authorisedButUnassignedDataManagers = $mismatchedDataManagers | ? {$_.SideIndicator -eq "=>"}


###################################################################
#                                                                 #
#           Tell Them What Will Happen and Process                #
#                                                                 #
###################################################################

### New
#Welcome new Data Managers $newauthorisedDataManagers.FieldValues.User.LookupValue | sort 
$newauthorisedDataManagers | % {
    $thisUser = $_
    $welcomeBodyTrunk =  "<HTML><FONT FACE=`"Calibri`">Hello $($thisUser.FieldValues.User.LookupValue.Split(" ")[0]),<BR><BR>`r`n`r`n"
    $welcomeBodyTrunk += "I'm sure you remember from your recent Data Manager training session on $(Get-Date $thisUser.FieldValues.Date_x0020_of_x0020_training -Format "dd MMMM yyyy") that we tell our clients that we train all Data Managers annually (to help us comply with their demands about how we manage their data). We will remind you 2 months in advance to ensure that you have time to join a refresher course.<BR><BR>`r`n`r`n"
    $welcomeBodyTrunk += "Just as a quick reminder, <A HREF='https://anthesisllc.sharepoint.com/:p:/r/teams/Working_Group_-_Collaboration_Improvement_365/Training%20Slides%20Shared/Data%20Management%20Training.pptx?d=wc75ff51e0236403b905979e34d5a1bdd&csf=1&e=FrfOxz'>the slides from the training session</A> are available for you to revist, and we have <A HREF='https://anthesisllc.sharepoint.com/sites/Resources-IT/SitePages/SharePoint-Training-Guides.aspx#data-managers-guides'>a series of Data Manager self-help guides</A> available too.<BR><BR>`r`n`r`n"
    $welcomeBodyTrunk += "If you have any more questions about Data Management, you can always ask the <A HREF='mailto:itteamall@anthesisgroup.com'>IT Team</A>,<BR><BR>`r`n`r`n"
    $welcomeBodyTrunk += "Love,`r`n`r`n<BR><BR>The Data Manager Robot</FONT></HTML>"
    #Send-MailMessage -From groupbot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Resources for new Data Managers" -BodyAsHtml $welcomeBodyTrunk -To kevin.maitland@anthesisgroup.com  -Encoding UTF8
    Send-MailMessage -From groupbot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Resources for new Data Managers" -BodyAsHtml $welcomeBodyTrunk -To $thisUser.FieldValues.User.Email -Encoding UTF8
    }

#Notify authorised Data Managers who don't own any sites that they can own sites, and to get in touch with the IT Team.
$authorisedButUnassignedDataManagers | ? {$newAuthorisedDataManagers.mail -notcontains $_.mail} | % { #Don't double-message new recruits
    $thisUser = $_
    if($thisUser.FieldValues.Date_x0020_of_x0020_training -ge $(Get-Date).AddMonths(-1)){ #Stop spamming users who trained > 1 month ago
        $nudgeBodyTrunk =  "<HTML><FONT FACE=`"Calibri`">Hello $($thisUser.FieldValues.User.LookupValue.Split(" ")[0]),<BR><BR>`r`n`r`n"
        $nudgeBodyTrunk += "It looks like you attended Data Manager training on $(Get-Date $thisUser.FieldValues.Date_x0020_of_x0020_training -Format "dd MMMM yyyy") but you haven't been allocated as a Data Manager of any teams yet.<BR><BR>`r`n`r`n"
        $nudgeBodyTrunk += "If this is an oversight, please let the <A HREF='mailto:itteamall@anthesisgroup.com'>IT Team</A> know which teams you should be managing, and they will fix this.<BR><BR>`r`n`r`n"
        $nudgeBodyTrunk += "Love,`r`n`r`n<BR><BR>The Data Manager Robot</FONT></HTML>"
        #Send-MailMessage -From groupbot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Resources for new Data Managers" -BodyAsHtml $welcomeBodyTrunk -To kevin.maitland@anthesisgroup.com  -Encoding UTF8
        Send-MailMessage -From groupbot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Data Managers status" -BodyAsHtml $nudgeBodyTrunk -To $thisUser.FieldValues.User.Email -Encoding UTF8
        }
    }


### Expiring
#Warn users that they will need to book onto a refresher course in the next two months $expiringSoonDataManagers.FieldValues.User.LookupValue | sort 
$expiringSoonDataManagers | % {

    $thisUser = $_
    
    #2 months remaining message (for the first time they'll be classed as an "expiring soon" Data Manager in this script - there won't be an entry on the most recent training record for them for a reminder email)
    If(!($thisUser.FieldValues.LastReminderEmailSent)){
    $warningBodyTrunk =  "<HTML><FONT FACE=`"Calibri`">Hello $($thisUser.FieldValues.User.LookupValue.Split(" ")[0]),<BR><BR>`r`n`r`n"
    $warningBodyTrunk += "To help us comply with the demands that our clients make about how we manage their data, we tell them that we train all Data Managers annually. Your last recorded Data Manager training session was on $(Get-Date $thisUser.FieldValues.Date_x0020_of_x0020_training -Format "dd MMMM yyyy"), so it's about time to <A HREF='https://anthesisllc.sharepoint.com/sites/ResourcesHub/SitePages/Upcoming-Training-Events.aspx'>book onto a refresher Data Manager training session</A>.<BR><BR>`r`n`r`n"
    $warningBodyTrunk += "A lot has changed in the past year. You can sign up for any session that is convenient for you and we'll go through some of the improvements that we've introduced, which will help you to work <I>even more</I> efficiently. <BR><BR>`r`n`r`n"
    $warningBodyTrunk += "If there aren't any suitable sessions available for you, please contact the <A HREF='mailto:itteamall@anthesisgroup.com'>IT Team</A> and they will arrange more. <BR><BR>`r`n`r`n"
    $warningBodyTrunk += "If you don't renew your training, you will automatically be changed to a Member of the following teams, and you won't be able to manage them until you join another <A HREF='https://anthesisllc.sharepoint.com/sites/ResourcesHub/SitePages/Upcoming-Training-Events.aspx'>Data Manager training session</A>:<BR><BR>`r`n<UL>"
    $whoOwnsWhatHash[$thisUser.FieldValues.User.Email] | Sort-Object {$_[0]} | % {
        $warningBodyTrunk += "`r`n`t<LI>$($_[0].Replace(" - Data Managers Subgroup",''))</LI>" #Then sublist each Team they are a Data Manager of
        }
    $warningBodyTrunk += "</UL><BR><BR>`r`n`r`nLove,`r`n`r`n<BR><BR>The Data Manager Robot</FONT></HTML>"
    #Send-MailMessage -From groupbot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Renew your Data Manager training before $(Get-Date (Get-Date $thisUser.FieldValues.Date_x0020_of_x0020_training.AddYears(1)) -f "yyyy-MM-dd")" -BodyAsHtml $warningBodyTrunk -To kevin.maitland@anthesisgroup.com  -Encoding UTF8
    Send-MailMessage -From groupbot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Renew your Data Manager training before $(Get-Date (Get-Date $thisUser.FieldValues.Date_x0020_of_x0020_training.AddYears(1)) -f "yyyy-MM-dd")" -BodyAsHtml $warningBodyTrunk -To $thisUser.FieldValues.User.Email  -Encoding UTF8
    #Send-MailMessage -From groupbot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Renew your Data Manager training before $(Get-Date (Get-Date $thisUser.FieldValues.Date_x0020_of_x0020_training.AddYears(1)) -f "yyyy-MM-dd")" -BodyAsHtml $warningBodyTrunk -To "emily.pressey@anthesisgroup.com"  -Encoding UTF8
    Set-PnPListItem -List "User Training Records" -Identity $thisUser.Id -Values @{"LastReminderEmailSent" = "$(get-date)"}
    }

    #1 month and 2 weeks remaining message
    $daysRemaining = (New-TimeSpan -start (get-date) -End ($thisUser.FieldValues.Date_x0020_of_x0020_training).AddYears(1)).Days
    If(($daysRemaining -eq 30) -or ($daysRemaining -eq 14)){
        $warningBodyTrunk =  "<HTML><FONT FACE=`"Calibri`">Hello $($thisUser.FieldValues.User.LookupValue.Split(" ")[0]),<BR><BR>`r`n`r`n"
    $warningBodyTrunk += "To help us comply with the demands that our clients make about how we manage their data, we tell them that we train all Data Managers annually. Your last recorded Data Manager training session was on $(Get-Date $thisUser.FieldValues.Date_x0020_of_x0020_training -Format "dd MMMM yyyy"), so it's about time to <A HREF='https://anthesisllc.sharepoint.com/sites/ResourcesHub/SitePages/Upcoming-Training-Events.aspx'>book onto a refresher Data Manager training session</A>.<BR><BR>`r`n`r`n"
    $warningBodyTrunk += "A lot has changed in the past year. You can sign up for any session that is convenient for you and we'll go through some of the improvements that we've introduced, which will help you to work <I>even more</I> efficiently. <BR><BR>`r`n`r`n"
    $warningBodyTrunk += "If there aren't any suitable sessions available for you, please contact the <A HREF='mailto:itteamall@anthesisgroup.com'>IT Team</A> and they will arrange more. <BR><BR>`r`n`r`n"
    $warningBodyTrunk += "If you don't renew your training, you will automatically be changed to a Member of the following teams, and you won't be able to manage them until you join another <A HREF='https://anthesisllc.sharepoint.com/sites/ResourcesHub/SitePages/Upcoming-Training-Events.aspx'>Data Manager training session</A>:<BR><BR>`r`n<UL>"
    $whoOwnsWhatHash[$thisUser.FieldValues.User.Email] | Sort-Object {$_[0]} | % {
        $warningBodyTrunk += "`r`n`t<LI>$($_[0].Replace(" - Data Managers Subgroup",''))</LI>" #Then sublist each Team they are a Data Manager of
        }
    $warningBodyTrunk += "</UL><BR><BR>`r`n`r`nLove,`r`n`r`n<BR><BR>The Data Manager Robot</FONT></HTML>"
    #Send-MailMessage -From groupbot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Renew your Data Manager training before $(Get-Date (Get-Date $thisUser.FieldValues.Date_x0020_of_x0020_training.AddYears(1)) -f "yyyy-MM-dd")" -BodyAsHtml $warningBodyTrunk -To kevin.maitland@anthesisgroup.com  -Encoding UTF8
    Send-MailMessage -From groupbot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Renew your Data Manager training before $(Get-Date (Get-Date $thisUser.FieldValues.Date_x0020_of_x0020_training.AddYears(1)) -f "yyyy-MM-dd")" -BodyAsHtml $warningBodyTrunk -To $thisUser.FieldValues.User.Email  -Encoding UTF8
    #Send-MailMessage -From groupbot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Renew your Data Manager training before $(Get-Date (Get-Date $thisUser.FieldValues.Date_x0020_of_x0020_training.AddYears(1)) -f "yyyy-MM-dd")" -BodyAsHtml $warningBodyTrunk -To "emily.pressey@anthesisgroup.com"  -Encoding UTF8
    Set-PnPListItem -List "User Training Records" -Identity $thisUser.Id -Values @{"LastReminderEmailSent" = "$(get-date)"}        
    }
        
    }


### Unauthorised
#Warn users that they have/will be automatically removed from Data Manager groups     $unauthorisedDataManagers.displayName | sort 
$unauthorisedDataManagers | ? {$_.mail -ne "groupbot@anthesisgroup.com"} | % {
    $thisUser = $_
    Write-Host "Sending email to $($thisUser.userPrincipalName) to let them know they've been booted" -ForegroundColor Yellow
    $removedBodyTrunk =  "<HTML><FONT FACE=`"Calibri`">Hello $($thisUser.displayName),<BR><BR>`r`n`r`n"
    $removedBodyTrunk += "To help us comply with the demands that our clients make about how we manage their data, we tell them that we train all Data Managers annually. We don't have a record of you attending a training session in the past year, so it's time to <A HREF='https://anthesisllc.sharepoint.com/sites/ResourcesHub/SitePages/Upcoming-Training-Events.aspx'>book onto a refresher Data Manager training session</A>.<BR><BR>`r`n`r`n"
    $removedBodyTrunk += "A lot has changed in the past year. You can sign up for any session that is convenient for you and we'll go through some of the improvements that we've introduced, which will help you to work <I>even more</I> efficiently. <BR><BR>`r`n`r`n"
    $removedBodyTrunk += "If there aren't any suitable sessions available for you, please contact the <A HREF='mailto:itteamall@anthesisgroup.com'>IT Team</A> and they will arrange more. <BR><BR>`r`n`r`n"
    $removedBodyTrunk += "You have been automatically changed to a Member of the following teams, and you won't be able to manage membership of them until you join another <A HREF='https://anthesisllc.sharepoint.com/sites/ResourcesHub/SitePages/Upcoming-Training-Events.aspx'>Data Manager training session</A>. <br><br><b>You will still have <u>edit</u> access to these sites to change any documents and any existing client/internal access will remain - you will be unable to add or remove both internal or external members to the site until you renew your training.</b><BR><BR>`r`n<UL>"
    $whoOwnsWhatHash[$thisUser.mail] | Sort-Object {$_[0]} | % {
        $removedBodyTrunk += "`r`n`t<LI>$($_[0].Replace(" - Data Managers Subgroup",''))</LI>" #Then sublist each Team they are a Data Manager of
        }
    $removedBodyTrunk += "</UL>Love,`r`n`r`n<BR><BR>The Data Manager Robot</FONT></HTML>"
    #Send-MailMessage -From groupbot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Renew your Data Manager training before 2020-07-01" -BodyAsHtml $removedBodyTrunk -To kevin.maitland@anthesisgroup.com  -Encoding UTF8;break
    Send-MailMessage -From groupbot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Renew your Data Manager training to be reinstated as a Data Manager" -BodyAsHtml $removedBodyTrunk -To $thisUser.userPrincipalName  -Encoding UTF8 -Priority High
    #Send-MailMessage -From groupbot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Renew your Data Manager training to be reinstated as a Data Manager" -BodyAsHtml $removedBodyTrunk -To "emily.pressey@anthesisgroup.com"  -Encoding UTF8 -Priority High
    #Now remove them from each of their groups (this will remove them from their "- Data Manager" groups, and the sync-UnifiedGroupMembership will then demote them on the next cycle). If they are the last Data Manager, they will be replaced with GroupBot
    if($thisUser.mail -ne "groupbot@anthesisgroup.com"){
        $whoOwnsWhatHash[$($thisUser.mail)] | % {
            Write-Verbose "Removing [$($thisUser.mail)] from [$($_[0])]"
            remove-DataManagerFromGroup -dataManagerGroupId $_[1] -upnToRemove $thisUser.mail -ErrorAction Stop
            }
        }
    
    }
    
        
   



###################################################################
#                                                                 #
#                           Reporting                             #
#                                                                 #
###################################################################

#Unauthorised managers are managed via SP list to track reminders and removal dates
<#
#Tell powerapps I'm working
$reportsyncdataManagers = update-graphListItem -tokenResponse $tokenResponse -graphSiteId "1ead4d00-2e2e-421e-9176-15287b5c22ce,d21ddf81-fcef-4e36-94e6-edd6fb487a31" -listName $ITReporting -listitemId 4 -fieldHash @{"Status" = "Ok";"Notes" = ""}

#Add some reporting info to IT Reporting SPO list for Authorised Data Managers
$allAuthorisedDataManagers = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $authorisedDataManagerGroup.id -memberType Members -returnOnlyUsers
$postArray = @()
forEach ($manager in $allAuthorisedDataManagers){
$usageLocation = get-graphUsers -tokenResponse $tokenResponse -filterUpns $manager.userPrincipalName -selectCustomProperties "usageLocation"
$hash = @(
"($($usageLocation.usageLocation)) " + "$($manager.userPrincipalName)"
)
$postArray += $hash
}
$postArray = $postArray | Sort-Object -Descending
[string]$postString = $postArray | out-string
$reportunauthorisedDataManagers = update-graphListItem -tokenResponse $tokenResponse -graphSiteId "1ead4d00-2e2e-421e-9176-15287b5c22ce,d21ddf81-fcef-4e36-94e6-edd6fb487a31" -listName $ITReporting -listitemId 3 -fieldHash @{"Notes" = $postString; "Objectcount" = $(($postArray | Measure-Object).Count)}


#Add some reporting info to IT Reporting SPO list for expiring soon Data Managers
$postArray = @()
forEach ($manager in $expiringSoonDataManagers){
$date = ($manager.FieldValues.Date_x0020_of_x0020_training | get-date -Format "dd/MM/yyyy")
$date = [datetime]::Parse(“$date”)
$manager | Add-Member -MemberType NoteProperty -Name "realDate" -Value $date  -Force
$hash = @(
"$(($manager.FieldValues.Date_x0020_of_x0020_training | get-date -Format "dd/MM/yyyy"))" + " - ($($manager.FieldValues.Location)) " + "$($manager.FieldValues.User.Email)"
)
$postArray += $hash
}
$postArray = $postArray | Sort-Object -Property "realDate"
[string]$postString = $postArray | out-string
$reportexpiringsoonDataManagers = update-graphListItem -tokenResponse $tokenResponse -graphSiteId "1ead4d00-2e2e-421e-9176-15287b5c22ce,d21ddf81-fcef-4e36-94e6-edd6fb487a31" -listName $ITReporting -listitemId 2 -fieldHash @{"Notes" = $postString; "Objectcount" = ($postArray | Measure-Object).Count}
#>




<#

Catch{
Write-Error "Data Manager's might not have been processed - we found an error"
$error
$reportsyncdataManagers = update-graphListItem -tokenResponse $tokenResponse -graphSiteId "1ead4d00-2e2e-421e-9176-15287b5c22ce,d21ddf81-fcef-4e36-94e6-edd6fb487a31" -listName $ITReporting -listitemId 4 -fieldHash @{"Status" = "Error - Unknown";"Notes" = $([string]$error[0].ErrorDetails.Message)}
}
#>







#region Overview report
$overviewBodyTrunk =  "<HTML><FONT FACE=`"Calibri`">Hello User/Exchange 365 Admins,<BR><BR>`r`n`r`n"
$overviewBodyTrunk += "This report combines information in <A HREF='https://anthesisllc.sharepoint.com/sites/Resources-HR/Lists/User%20Training%20Records/AllItems.aspx?viewpath=%2Fsites%2FResources-HR%2FLists%2FUser%20Training%20Records%2FAllItems.aspx'>User Training Records</A>, membership in [Data Managers - Authorised (All)] and membership in the individual [XYZ Team - Data Manager Subgroup] groups to ensure that all Data Managers have received training within the past 12 months. This will allow us to embed best practices within the business, meet client security requirements more easily and prepare for 3rd party accreditation (like ISO27001). <A HREF='https://github.com/kevmaitland/PoshAnt/blob/master/report-dataManagers.ps1'>GitCode</A><BR><BR>`r`n"
$overviewBodyTrunk += "<A HREF='https://anthesisllc.sharepoint.com/sites/Resources-HR/Lists/User%20Training%20Records/AllItems.aspx?viewpath=%2Fsites%2FResources-HR%2FLists%2FUser%20Training%20Records%2FAllItems.aspx'>User Training Records</A> must be completed for all people who attend a <A HREF='https://anthesisllc.sharepoint.com/sites/ResourcesHub/SitePages/Upcoming-Training-Events.aspx'>Data Manager training session</A>. The least-unreliable approach is for the person leading the training session to add these records on behalf of the attendees.<BR><BR>`r`n"
$overviewBodyTrunk += "Weekly nofitications are sent to users whose training will expire in the next 2 months prompting them to join a <A HREF='https://anthesisllc.sharepoint.com/sites/ResourcesHub/SitePages/Upcoming-Training-Events.aspx'>Data Manager training session</A>.<BR><BR>`r`n"
$overviewBodyTrunk += "Beginning 2020-07-01, users who have no valid <A HREF='https://anthesisllc.sharepoint.com/sites/Resources-HR/Lists/User%20Training%20Records/AllItems.aspx?viewpath=%2Fsites%2FResources-HR%2FLists%2FUser%20Training%20Records%2FAllItems.aspx'>Data Manager training record</A> will be automatically removed from all Data Manager groups (and replaced with GroupBot if they were the last Data Manager).<BR><BR>`r`n"
$overviewBodyTrunk += "The following users have recently had training records created and have been re/added into [Data Managers - Authorised (All)]:<BR><BR>`r`n<UL>"
$newauthorisedDataManagers | Sort-Object {$_.FieldValues.User.Email} | % {
    $overviewBodyTrunk += "<LI>$($_.FieldValues.User.Email)</LI>`r`n"
    }
$overviewBodyTrunk +=  "</UL>`r`n`r`n<BR><BR>The following users' training will expire in the next 2 months:<BR><BR>`r`n<UL>"
$expiringSoonDataManagers | Sort-Object {$_.FieldValues.User.Email} | % {
    $overviewBodyTrunk += "<LI>$($_.FieldValues.User.Email)</LI>`r`n"
    }
$overviewBodyTrunk +=  "</UL>`r`n`r`n<BR><BR>The following users have not renewed their training and have been removed from [Data Managers - Authorised (All)]:<BR><BR>`r`n<UL>"
$expiredDataManagers | Sort-Object {$_.FieldValues.User.Email} | % {
    $overviewBodyTrunk += "<LI>$($_.FieldValues.User.Email)</LI>`r`n"
    }
$overviewBodyTrunk +=  "</UL>`r`n`r`n<BR><BR>The following users have been removed as Data Managers (they are currently members of a [XYZ Team - Data Manager Subgroup], but have no valid training record):<BR><BR>`r`n<UL>"
$unauthorisedDataManagers | Sort-Object {$_.mail} | % {
    $thisManager = $_
    $overviewBodyTrunk += "`r`n<LI><B>$($thisManager.mail)</B><UL>" #List the Managers alphabetically
    $whoOwnsWhatHash[$thisManager.mail] | Sort-Object {$_[0]} | % {
        $overviewBodyTrunk += "`r`n`t<LI>$($_[0].Replace(" - Data Managers Subgroup",''))</LI>" #Then sublist each Team they are a Data Manager of
        }
    $overviewBodyTrunk += "</UL>" 
    }
$overviewBodyTrunk +=  "</UL>`r`n`r`n<BR><BR><B>All</B> Data Managers and the groups they manage (unauthorised Data Managers are in <I>italics</I>):<BR><BR>`r`n<UL>"
$allCurrentDataManagers | Sort-Object {$_.mail} | % {
    $thisManager = $_
    if($unauthorisedDataManagers.mail -contains $thisManager.mail){
        $overviewBodyTrunk += "`r`n<LI><B><I>$($thisManager.mail)</I></B><UL>" #List the Managers alphabetically (markup Unauthorised Data Managers in italics too)
        }
    else{
        $overviewBodyTrunk += "`r`n<LI><B>$($thisManager.mail)</B><UL>" #List the Managers alphabetically 
        }
    $whoOwnsWhatHash[$thisManager.mail] | Sort-Object {$_[0]} | % {
        $overviewBodyTrunk += "`r`n`t<LI>$($_[0].Replace(" - Data Managers Subgroup",''))</LI>" #Then sublist each Team they are a Data Manager of
        }
    $overviewBodyTrunk += "</UL>" 
    }
$overviewBodyTrunk +=  "</UL>`r`n`r`n<BR><BR>All Groups and who they are managed by:<BR><BR>`r`n<UL>"
$allDataManagerSubGroups | Sort-Object {$_.displayName} | % {
    $thisGroup = $_
    $overviewBodyTrunk += "`r`n<LI><B>$($thisGroup.displayName.Replace(" - Data Managers Subgroup",''))</B><UL>" #List the Groups alphabetically
    $whatisOwnedByWhoHash[$thisGroup.id] | Sort-Object {$_[0]} | % {
        $overviewBodyTrunk += "`r`n`t<LI>$($_[0])`t:`t$($_[1])</LI>" #Then sublist each Data Manager
        }
    $overviewBodyTrunk += "</UL>" 
    }
$overviewBodyTrunk += "Love,`r`n`r`n<BR><BR>The Data Manager Robot</FONT></HTML>"

$groupAndExchangeAdmins = get-graphAdministrativeRoleMembers -tokenResponse $tokenResponse -roleName 'Exchange Service Administrator' 
$groupAndExchangeAdmins += get-graphAdministrativeRoleMembers -tokenResponse $tokenResponse -roleName 'User Account Administrator' 
$groupAndExchangeAdmins = $groupAndExchangeAdmins | Sort-Object userPrincipalName -Unique

#Send-MailMessage -From groupbot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Data Manager Summary $(Get-Date -f "yyyy-MM-dd")" -BodyAsHtml $overviewBodyTrunk -To kevin.maitland@anthesisgroup.com  -Encoding UTF8
Send-MailMessage -From groupbot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Data Manager Summary $(Get-Date -f "yyyy-MM-dd")" -BodyAsHtml $overviewBodyTrunk -To "emily.pressey@anthesisgroup.com" -Encoding UTF8
#endregion



