$logFileLocation = "C:\ScriptLogs\"
$transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_Transcript_$(Get-Date -Format "yyMMdd").log"
if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)){
    $fullLogPathAndName = $logFileLocation+"PeopleServices_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = $logFileLocation+"PeopleServices_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
else{
    $fullLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
Start-Transcript $transcriptLogName -Append

<#Connect to everything and load modules#>
Import-Module _PNP_Library_SPO
#create a funciton to capture script lines to help troubleshoot
Function Get-CurrentLine {
    $Myinvocation.ScriptlineNumber
}

$sharePointAdmin = "kimblebot@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content "$env:USERPROFILE\Desktop\KimbleBot.txt") 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass


#Set Variables to connect to Sharepoint People Services Site and some other list variables
$SiteURL = "https://anthesisllc.sharepoint.com/teams/People_Services_Team_All_365"

<#--------------Connect to Sharepoint--------------#>
Connect-PnPOnline -Url $SiteURL -Credentials $adminCreds
$context = Get-PnPContext


<#--------------Get all lists--------------#>
$RecruitmentArea = Get-PnPList -Identity "Recruitment Area"
$NewStarterList = Get-PnPList -Identity "New Starter Details"
$items = Get-PnPListItem -List "Recruitment Area"


<#--------------Get all the Lists from the Site, find the live ones, "Live Candidate Tracker" will be in the description"--------------#>

$FullListQuery = Get-PnPList 
$LiveCandidateTrackers = @()
ForEach($List in $FullListQuery){
If($List.Description -match "Live Candidate Tracker"){

        $RoleId = ($($List.Description) -split ':')[1]

        $LiveCandidateTrackers += New-Object psobject -Property @{
        'Title' = $List.Title;
        'Guid' = $List.Id;
        'Description' = $List.Description;
        'RoleID' = $RoleId;
        
        }
     }
}

  


<#--------------Process each list--------------#>

#Iterate through each list and check for any actions against candidates need processing - is the date modified more recent than the Last Modified Date?
$Folderstocreate = @()
ForEach($LiveTracker in $LiveCandidateTrackers){

    $Itemstoprocess = Get-PnPListItem -List $LiveTracker.Guid  
    
    foreach($Item in $Itemstoprocess){
           
     #First, check for blanks in the trigger columns: Decision 1 and Final decision Last Entry fields. Errors will occur if thses fields are blank for the compare-object sections below.

        #check for blanks in the Decision 1 Last Entry column, just in case - the default value is "Decision Pending" on item creation
        If(!$Item.FieldValues.D1LE){
        write-host "Looks like Decision 1 Last Entry is blank, filling it in" -ForegroundColor Yellow
        Set-PnPListItem -List $List -Identity $item.ID -Values @{"D1LE" = "$($Item.FieldValues.Decision_x0020_1)"}
        Continue
            }
        #Check for blanks in the Final Decision Last Entry column  - the default value is "Decision Pending" on item creation
        If(!$Item.FieldValues.FDLE){
        write-host "Looks like Final Decision Last Entry is blank, filling it in" -ForegroundColor Yellow
        Set-PnPListItem -List $List -Identity $item.ID -Values @{"FDLE" = "$($Item.FieldValues.Final_x0020_Decision)"}
        Continue
            }

      #Second, check for Declined Candidates so we can skip the iteration and then compare the key columns for changes for non-declined candidates, if there is a change, we can process it below.

       #Check for Declined candidates - if Decision 1 is set to Decline, set the Last Entry fields and Final Decision  fields also as Decline and continue onto the next iteration. Same concept for Final Decision, but this won't change Interview 1 decision fields, just the Last entry field.
        If(("Decline" -eq $Item.FieldValues.Decision_x0020_1) -and ("Decline" -ne $Item.FieldValues.D1LE)){
           Write-Host "Looks like $($Item.FieldValues.Candidate_x0020_Name) has been declined after Interview 1 since we last ran through - setting other fields to decline and moving on" -ForegroundColor Yellow
           Set-PnPListItem -List $LiveTracker.Guid -Identity $Item.ID -Values @{'D1LE' = "$($Item.FieldValues.Decision_x0020_1)"}
           Set-PnPListItem -List $LiveTracker.Guid -Identity $Item.ID -Values @{'Final_x0020_Decision' = "$($Item.FieldValues.Decision_x0020_1)"}
           Set-PnPListItem -List $LiveTracker.Guid -Identity $Item.ID -Values @{'FDLE' = "$($Item.FieldValues.Decision_x0020_1)"}
           Continue
            }
        If(("Decline" -eq $Item.FieldValues.Final_x0020_Decision) -and ("Decline" -ne $Item.FieldValues.FDLE)){
           Write-Host "Looks like $($Item.FieldValues.Candidate_x0020_Name) has been declined after final interview since we last ran through - setting other fields to decline and moving on" -ForegroundColor Yellow
           Set-PnPListItem -List $LiveTracker.Guid -Identity $Item.ID -Values @{'FDLE' = "$($Item.FieldValues.Final_x0020_Decision)"}
           Continue
            }

      #Third, the Decision Columns - only process them if they are different as it indicates a change.

        $InterView1Decision = (Compare-Object -ReferenceObject $Item.FieldValues.Decision_x0020_1 -DifferenceObject $Item.FieldValues.D1LE)
        $FinalDecision = (Compare-Object -ReferenceObject $Item.FieldValues.Final_x0020_Decision -DifferenceObject $Item.FieldValues.FDLE)
        If($InterView1Decision){
        Write-host "$($Item.FieldValues.Candidate_x0020_Name): Something has changed on the Interview 1 Decision Field! Let's maybe do something about it!" -ForegroundColor Yellow
        }
        If($FinalDecision){
        Write-host "$($Item.FieldValues.Candidate_x0020_Name): Something has changed on the Final Decision Field! Let's maybe do something about it!" -ForegroundColor Yellow
        }
        
      #Fourth, check which part has changed and action based on input. We have included an -and statement to just include those that have changed since the last run or it will keep sending out emails
        
        #Check if Interview 1 needs processing and it matches "Move to Next Stage" so we can let People Services know
        If(($InterView1Decision) -and ($Item.FieldValues.Decision_x0020_1 -match "Move to Next Stage")) {

        write-host "Interview 1 Decision for $($Item.FieldValues.Candidate_x0020_Name) has changed from $($Item.FieldValues.D1LE) to $($Item.FieldValues.Decision_x0020_1)" -ForegroundColor Yellow
        $currentline = (Get-CurrentLine)
        Try{
            #Send email to People services letting them know to schedule a second interview
            $subject = "Recruitment Update: A Candidate is Ready to Move to Second Interview"
            $body = "<HTML><FONT FACE=`"Calibri`">Hello People Services Team,`r`n`r`n<BR><BR>"
            $body += "The Candidate $($Item.FieldValues.Candidate_x0020_Name) for role $($LiveTracker.Title) has been moved to the next stage.`r`n`r`n<BR><BR>"
            $body += "Please schedule an interview with the candidate and fill in the details of the date and type of interview in the candidate tracker.`r`n`r`n<BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>"
            $body += "The People Services Robot"
            Send-MailMessage -To "nina.cairns@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            #Set the Decision 1 Last Entry column (D1LE) to the new Entry, this will stop it from re-processing - we don't want people getting multiple emails
            Set-PnPListItem -List $LiveTracker.Guid -Identity $Item.ID -Values @{'D1LE' = "$($Item.FieldValues.Decision_x0020_1)"}
            }
        Catch{
            $ErrorMessage = $_.Exception.Message
            #Send email to People services letting them know to schedule a second interview
            $subject = "Recruitment Update: something has gone wrong sending an email to People Services for the next interview stage for Candidate $($Item.FieldValues.Candidate_x0020_Name)."
            $body = "<HTML><FONT FACE=`"Calibri`">Hello People Services Team,`r`n`r`n<BR><BR>"
            $body += "Should probably check it out - I'm breaking at line $($currentline). The error is: $($ErrorMessage)`r`n`r`n<BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>"
            $body += "The People Services Robot"
            Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            }
       }


       #Fifth, Check if Final Decision needs processing and it matches "Make Offer" so we can let People Services know
        If(($FinalDecision) -and ($Item.FieldValues.Final_x0020_Decision -match "Make Offer")) {

        write-host "FinalDecision for $($Item.FieldValues.Candidate_x0020_Name) has changed from $($Item.FieldValues.FDLE) to $($Item.FieldValues.Final_x0020_Decision)" -ForegroundColor Yellow
        $currentline = (Get-CurrentLine)
            try{
            #Send email to People services letting them know to make an offer to this Candidate
            $subject = "Recruitment Update: A Candidate is Ready to Receive an Offer"
            $body = "<HTML><FONT FACE=`"Calibri`">Hello People Services Team,`r`n`r`n<BR><BR>"
            $body += "The Candidate $($Item.FieldValues.Candidate_x0020_Name) for role $($LiveTracker.Title) is ready to recieve an offer.`r`n`r`n<BR><BR>"
            $body += "Please send an offer to the candidate. The candidate tracker 'Offer Outcome' column has automatically been set to 'Pending'. Please set this to either 'Accepted' or 'Rejected' based on the Candidates response and enter the proposed starting date as soon as possible <b>as the last action on the Candidate Tracker, this will label the hiring process as complete</b>. This will inform our internal teams of an upcoming starter to ensure things like IT hardware is in-stock.`r`n`r`n<BR><BR>"
            $body += "When a proposed date is entered, a template entry will be added to the New Starter Form (you will receive an email with a link to this). This will use information we already know about the role and candidate, but will not be complete. Please fill this entry out as soon as possible so internal teams can set them up ready for their first day. `r`n`r`n<BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>"
            $body += "The People Services Robot"
            Send-MailMessage -To "nina.cairns@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 
            
            #Set item 'Offer Outcome' to 'Pending', which People Services will change on Candidate response.
            Set-PnPListItem -List $LiveTracker.Guid -Identity $Item.ID -Values @{"Offer_x0020_Outcome" = "Pending"}
            Set-PnPListItem -List $LiveTracker.Guid -Identity $Item.ID -Values @{'FDLE' = "$($Item.FieldValues.Final_x0020_Decision)"}

            #Add the Candidate to a list so we can create their Employee folder later
            $Folderstocreate += New-Object psobject -Property @{"Candidate Name" = $Item.FieldValues.Candidate_x0020_Name}
            }
            Catch{
            $ErrorMessage = $_.Exception.Message
            $subject = "Recruitment Update: Something has gone wrong with the Final Decision email for Candidate $($Item.FieldValues.Candidate_x0020_Name)"
            $body = "<HTML><FONT FACE=`"Calibri`">Hello People Services Team,`r`n`r`n<BR><BR>"
            $body += "Should probably check it out - I'm breaking at the script block starting with line $($currentline). The error is $($ErrorMessage).`r`n`r`n<BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>"
            $body += "The People Services Robot"
            Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 
            }

        }

       #Sixth, check if there is a Proposed Start Date, this suggests some firm date has been set and the Candidate is likely to start then, create a new template entry with what we know about the Candidate already in the 'New Starter Details' List. Set the Candidate Tracker as "Complete"
        
        #Check for Start Date, if so create a new template entry in the New starter Details List.
        If($Item.FieldValues.Proposed_x0020_Start_x0020_Date){
        Write-host "Looks like the Hiring Process is complete for $($LiveTracker.Title). Let's create a template entry in the 'New Starter Details' Form based on what we know already" -ForegroundColor Yellow
        $currentline = (Get-CurrentLine)


        Try{
            #Create some SP-friendly variables
            [datetime]$Friendlystartdate = $($Item.FieldValues.Proposed_x0020_Start_x0020_Date)
            $RecruitmentAreaItem = Get-PnPListItem -List $RecruitmentArea -Id $($LiveTracker.RoleId)

            #Start Pre-populating the New Starter Details Form
                Write-host "creating New Starter entry for $($Item.FieldValues.Candidate_x0020_Name)..." -ForegroundColor Yellow
                $newstarterenrty = Add-PnPListItem -List $NewStarterList -Values @{
                "Employee_x0020_Preferred_x0020_N" = "$($Item.FieldValues.Candidate_x0020_Name)"; 
                "StartDate" = "$Friendlystartdate"; 
                "JobTitle" = "$($RecruitmentAreaItem.FieldValues.Role_x0020_Name)";
                "Line_x0020_Manager" = "$($RecruitmentAreaItem.FieldValues.Hiring_x0020_Manager.LookupValue)";
                "Primary_x0020_Team0" = "$($RecruitmentAreaItem.FieldValues.Primary_x0020_Team0.TermGuid)";
                "Community0" = "$($RecruitmentAreaItem.FieldValues.Community0.TermGuid)";
                "Business_x0020_Unit0" = "$($RecruitmentAreaItem.FieldValues.Business_x0020_Unit0.TermGuid)";
                }
            }
            Catch{
            $ErrorMessage = $_.Exception.Message
            Write-Host "Failure! ): Not sure what happened but a template entry could not be added to the New Starters From: $($Item.FieldValues.Candidate_x0020_Name)"
            $subject = "Employee New Starter Template Creation: Woops, something went wrong..."
            $body = "<HTML><FONT FACE=`"Calibri`">Hello IT Team,`r`n`r`n<BR><BR>"
            $body += "Something went wrong when trying to create template in the New Starters Form for <b>$($Item.FieldValues.Candidate_x0020_Name)</b>. Should probably take a look and see what's gone wrong - I'm breaking at the script block starting with line $($currentline). Error message is: $($ErrorMessage)`r`n`r`n<BR><BR>"
            $body += "<b>Timestamp: </b>$(get-date)`r`n`r`n<BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>"
            $body += "The People Services Robot"
            Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 
            }

            
            #Check a new item was actually created in the New Starters List as a secondary check - sometimes PnP doesn't tell us if it is unsuccessful, if so, close it off so it doesn't re-run. Set Candidate tracker description as complete
            
            If($newstarterenrty){
            write-host "Success! A new template entry was made in the New Starters Form for $($Item.FieldValues.Candidate_x0020_Name). Closing the Candidate tracker..." -ForegroundColor Yellow
            Set-PnPList -Identity "ID$($ListTitle)" -Description "Closed Candidate Tracker - RoleID:$($Role.'ID')" #Set the description to omit this in our processing script, which searches for "Live Candidate Tracker" in the List Description
            Set-PnPListItem -List $RecruitmentArea -Identity $LiveTracker.RoleId -Values @{"Role_x0020_Hire_x0020_Status" = "Complete"}

            #Send a confirmation email to People Services
            $subject = "Recruitment Update: A Candidate is set to start and a Template Entry has been added to the New Starter Details List"
            $body = "<HTML><FONT FACE=`"Calibri`">Hello People Services Team,`r`n`r`n<BR><BR>"
            $body += "The Candidate $($Item.FieldValues.Candidate_x0020_Name) for role $($LiveTracker.Title) now has a provisional start date!`r`n`r`n<BR><BR>"
            $body += "A new template entry has been created in the New Starter Details List, ready to be finished. Please finish this entry in good time before the start date so that Internal Teams can be ready for them to start.`r`n`r`n<BR><BR>"
            $body += "You can see the New Starter Details List here: https://anthesisllc.sharepoint.com/teams/People_Services_Team_All_365/Lists/New%20Starter%20Details/AllItems.aspx `r`n`r`n<BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>"
            $body += "The People Services Robot"
            Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 
            Send-MailMessage -To "nina.cairns@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            }
            Else{
            Write-Host "Failure! ): Not sure what happened but a template entry could not be added to the New Starters Form: $($Item.FieldValues.Candidate_x0020_Name)"
            $subject = "Employee New Starter Template Creation: Woops, something went wrong..."
            $body = "<HTML><FONT FACE=`"Calibri`">Hello IT Team,`r`n`r`n<BR><BR>"
            $body += "Something went wrong when trying to create template in the New Starters Form for <b>$($Item.FieldValues.Candidate_x0020_Name)</b>. Should probably take a look and see what's gone wrong - I'm breaking at the script block starting with line $($currentline)`r`n`r`n<BR><BR>"
            $body += "<b>Timestamp: </b>$(get-date)`r`n`r`n<BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>"
            $body += "The People Services Robot"
            Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 
            }

        }
    }
 }


<#--------------Connect to the confidential HR team site with Graph--------------#>  #Kimblebot is currently not allowed to connect to this site


#Get salted credentials and get an Accesstoken
$teamBotDetails = Import-Csv "$env:USERPROFILE\Desktop\teambotdetails.txt"
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


#Connnect to sharepoint via Graph
Connect-PnPOnline -AccessToken $tokenResponse.access_token


<#--------Create the Employee Folder Structure--------#>


ForEach($folder in $Folderstocreate){

<#--------create the initial Parent folder--------#>
$body = "{
    `"name`": `"$($folder.'Candidate Name')`",
    `"folder`": { },
    `"@microsoft.graph.conflictBehavior`": `"rename`"
}"
$body = [System.Text.Encoding]::UTF8.GetBytes($body)
$CandidateNameResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,e43ccfa7-1258-4a83-a6a9-483577275b99,d21ddf81-fcef-4e36-94e6-edd6fb487a31/drives/b!p8885FgSg0qmqUg1dydbmYHfHdLv_DZOlObt1vtIejFDr6vvuqdFTaTWzb63-TzY/items/01LLWAYUILOIXGORD4QBFYI6MMKVPW4HZI/children" -Body $body -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Post


<#--------create the subfolders within the parent folder created above--------#>
#Subfolder 1.Onboarding
$body = "{
    `"name`": `"1. Onboarding`",
    `"folder`": { },
    `"@microsoft.graph.conflictBehavior`": `"rename`"
}"
$body = [System.Text.Encoding]::UTF8.GetBytes($body)
$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,e43ccfa7-1258-4a83-a6a9-483577275b99,d21ddf81-fcef-4e36-94e6-edd6fb487a31/drives/b!p8885FgSg0qmqUg1dydbmYHfHdLv_DZOlObt1vtIejFDr6vvuqdFTaTWzb63-TzY/items/$($CandidateNameResponse.Id)/children" -Body $body -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Post

#Place new Onboarding folder ID into variable to use next
$OnboardingfolderID = $response.id
$ParentFolderID = $CandidateNameResponse.id
$PSConfidentialDriveID = "b!p8885FgSg0qmqUg1dydbmYHfHdLv_DZOlObt1vtIejFDr6vvuqdFTaTWzb63-TzY"

#Create New Starter Checklist template in the Onboarding folder we created above
$body = "{
    `"parentReference`": {
        `"driveId`": `"b!AE2tHi4uHkKRdhUoe1wizoHfHdLv_DZOlObt1vtIejFDr6vvuqdFTaTWzb63-TzY`",
        `"id`": `"01V67YTVHO2Y3JHJUM35EZTMY3LNCLRVNR`"
         },
    `"name`": `"New Starter Checklist.xlsx`",
    `"@microsoft.graph.conflictBehavior`": `"rename`"
}"
$body = $body.Replace("b!AE2tHi4uHkKRdhUoe1wizoHfHdLv_DZOlObt1vtIejFDr6vvuqdFTaTWzb63-TzY","$($PSConfidentialDriveID)")#Replace JSON parent folder
$body = $body.Replace("01V67YTVHO2Y3JHJUM35EZTMY3LNCLRVNR","$($OnboardingfolderID)")#Replace JSON Onboarding subfolder
$body = [System.Text.Encoding]::UTF8.GetBytes($body)
$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,1ead4d00-2e2e-421e-9176-15287b5c22ce,d21ddf81-fcef-4e36-94e6-edd6fb487a31/drives/b!AE2tHi4uHkKRdhUoe1wizoHfHdLv_DZOlObt1vtIejFDr6vvuqdFTaTWzb63-TzY/items/01V67YTVFEUVDEF5VSBRFLKMCSRJZZOCK6/copy" -Body $body -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Post


#Subfolder 2. Lifecycle
$body = "{
    `"name`": `"2. Lifecycle`",
    `"folder`": { },
    `"@microsoft.graph.conflictBehavior`": `"rename`"
}"
$body = [System.Text.Encoding]::UTF8.GetBytes($body)
$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,e43ccfa7-1258-4a83-a6a9-483577275b99,d21ddf81-fcef-4e36-94e6-edd6fb487a31/drives/b!p8885FgSg0qmqUg1dydbmYHfHdLv_DZOlObt1vtIejFDr6vvuqdFTaTWzb63-TzY/items/$($CandidateNameResponse.Id)/children" -Body $body -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Post


#Check the last subfolder was created, this won't create if the parent folder creation was not successful. Send an email if there are any problems.
#Subfolder 3. Offboarding
Try{
$body = "{
    `"name`": `"3. Offboarding`",
    `"folder`": { },
    `"@microsoft.graph.conflictBehavior`": `"rename`"
}"
$body = [System.Text.Encoding]::UTF8.GetBytes($body)
$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,e43ccfa7-1258-4a83-a6a9-483577275b99,d21ddf81-fcef-4e36-94e6-edd6fb487a31/drives/b!p8885FgSg0qmqUg1dydbmYHfHdLv_DZOlObt1vtIejFDr6vvuqdFTaTWzb63-TzY/items/$($CandidateNameResponse.Id)/children" -Body $body -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Post
}
catch{

            $subject = "Employee Folder Creation: Woops, something went wrong..."
            $body = "<HTML><FONT FACE=`"Calibri`">Hello IT Team,`r`n`r`n<BR><BR>"
            $body += "Something went wrong when trying to create an employee folder for <b>$($folder.'Candidate Name')</b>. Should probably take a look and see what's gone wrong.`r`n`r`n<BR><BR>"
            $body += "<b>Timestamp: </b>$(get-date)`r`n`r`n<BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>"
            $body += "The People Services Robot"

            Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 

}

}


#Other working out stuff I don't want to delete just yet - please excuse the mess!

#$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,e43ccfa7-1258-4a83-a6a9-483577275b99,d21ddf81-fcef-4e36-94e6-edd6fb487a31/drives/b!p8885FgSg0qmqUg1dydbmYHfHdLv_DZOlObt1vtIejFDr6vvuqdFTaTWzb63-TzY/items/01LLWAYUNWN7IAV3SML5EYVWC6XIAWAJPF/children" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Get
#$response.value.Name

#$response.value = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,e43ccfa7-1258-4a83-a6a9-483577275b99,d21ddf81-fcef-4e36-94e6-edd6fb487a31/drives/efabaf43-a7ba-4d45-a4d6-cdbeb7f93cd8/items/" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Get
#$response.value.Name

#$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,e43ccfa7-1258-4a83-a6a9-483577275b99,d21ddf81-fcef-4e36-94e6-edd6fb487a31/drives/efabaf43-a7ba-4d45-a4d6-cdbeb7f93cd8/items/" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Get
#$response.value.Name


#Shared Documents = efabaf43-a7ba-4d45-a4d6-cdbeb7f93cd8
    #AUK Employee Folders = 01LLWAYUOPYQEQ4RYWDJGZ2MPQ6QENEXEN
        #Current Employees = 01LLWAYUILOIXGORD4QBFYI6MMKVPW4HZI
            #1. Staff folder template = 01LLWAYUNWN7IAV3SML5EYVWC6XIAWAJPF

#$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,e43ccfa7-1258-4a83-a6a9-483577275b99,d21ddf81-fcef-4e36-94e6-edd6fb487a31/drives/b!p8885FgSg0qmqUg1dydbmYHfHdLv_DZOlObt1vtIejFDr6vvuqdFTaTWzb63-TzY/items/01LLWAYUOPYQEQ4RYWDJGZ2MPQ6QENEXEN/children/" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Get
#$response.value.Name


    
    #$Onboardingfoldername = "Shared Documents" + "\" + $folder.'Candidate Name' + "1. Onboarding"
    #Copy-PnPFile -SourceUrl "https://anthesisllc.sharepoint.com/:x:/r/sites/Confidential_Human_Resources_HR_Team_GBR_365/_layouts/15/Doc.aspx?sourcedoc=%7BAFD940AB-DED8-4C2B-BD2F-4AE144B72460%7D&file=New%20Starter%20Checklist.xlsx&action=default&mobileredirect=true" -TargetUrl $Onboardingfoldername


  
#$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com:/teams/IT_Team_All_365" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Get
#write-host "The ID for $($response.displayname) is $($response.id)"

#$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,1ead4d00-2e2e-421e-9176-15287b5c22ce,d21ddf81-fcef-4e36-94e6-edd6fb487a31/drives" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Get

#$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,1ead4d00-2e2e-421e-9176-15287b5c22ce,d21ddf81-fcef-4e36-94e6-edd6fb487a31/drives" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Get
#$response.value

#Shared Documents doclib
#b!AE2tHi4uHkKRdhUoe1wizoHfHdLv_DZOlObt1vtIejFDr6vvuqdFTaTWzb63-TzY

#$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,1ead4d00-2e2e-421e-9176-15287b5c22ce,d21ddf81-fcef-4e36-94e6-edd6fb487a31/drives/b!AE2tHi4uHkKRdhUoe1wizoHfHdLv_DZOlObt1vtIejFDr6vvuqdFTaTWzb63-TzY/root/children" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Get
#$response.value

#$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,1ead4d00-2e2e-421e-9176-15287b5c22ce,d21ddf81-fcef-4e36-94e6-edd6fb487a31/drives/b!AE2tHi4uHkKRdhUoe1wizoHfHdLv_DZOlObt1vtIejFDr6vvuqdFTaTWzb63-TzY/items/01V67YTVHO2Y3JHJUM35EZTMY3LNCLRVNR/children" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Get
#$response.value




#This points directly to the file
#$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,1ead4d00-2e2e-421e-9176-15287b5c22ce,d21ddf81-fcef-4e36-94e6-edd6fb487a31/drives/b!AE2tHi4uHkKRdhUoe1wizoHfHdLv_DZOlObt1vtIejFDr6vvuqdFTaTWzb63-TzY/items/01V67YTVFEUVDEF5VSBRFLKMCSRJZZOCK6" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Get
#$response.value


#file ID
#01V67YTVFEUVDEF5VSBRFLKMCSRJZZOCK6

#Folder ID
#01V67YTVHO2Y3JHJUM35EZTMY3LNCLRVNR
   

<#
$body = "{
    `"name`": `"1. Onboarding2`",
    `"folder`": { },
    `"@microsoft.graph.conflictBehavior`": `"rename`"
}"
$body = [System.Text.Encoding]::UTF8.GetBytes($body)
#$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,e43ccfa7-1258-4a83-a6a9-483577275b99,d21ddf81-fcef-4e36-94e6-edd6fb487a31/drives/b!p8885FgSg0qmqUg1dydbmYHfHdLv_DZOlObt1vtIejFDr6vvuqdFTaTWzb63-TzY/items/$($CandidateNameResponse.Id)/children" -Body $body -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Post
$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,1ead4d00-2e2e-421e-9176-15287b5c22ce,d21ddf81-fcef-4e36-94e6-edd6fb487a31/drives/b!AE2tHi4uHkKRdhUoe1wizoHfHdLv_DZOlObt1vtIejFDr6vvuqdFTaTWzb63-TzY/items/01V67YTVHO2Y3JHJUM35EZTMY3LNCLRVNR/children" -Body $body -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Post





$body = "{
    `"parentReference`": {
        `"driveId`": `"b!AE2tHi4uHkKRdhUoe1wizoHfHdLv_DZOlObt1vtIejFDr6vvuqdFTaTWzb63-TzY`",
        `"id`": `"01V67YTVHO2Y3JHJUM35EZTMY3LNCLRVNR`"
         },
    `"name`": `"New Starter Checklist1.xlsx`",
    `"@microsoft.graph.conflictBehavior`": `"rename`"
}"
$body = [System.Text.Encoding]::UTF8.GetBytes($body)
$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,1ead4d00-2e2e-421e-9176-15287b5c22ce,d21ddf81-fcef-4e36-94e6-edd6fb487a31/drives/b!AE2tHi4uHkKRdhUoe1wizoHfHdLv_DZOlObt1vtIejFDr6vvuqdFTaTWzb63-TzY/items/01V67YTVFEUVDEF5VSBRFLKMCSRJZZOCK6/copy" -Body $body -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Post

$body = "{
    `"parentReference`": {
        `"id`": `"01V67YTVHO2Y3JHJUM35EZTMY3LNCLRVNR`",
         },
    `"name`": `"New Starter Checklist`",
    `"@microsoft.graph.conflictBehavior`": `"rename`"
}"
$body = [System.Text.Encoding]::UTF8.GetBytes($body)
$response.value = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,1ead4d00-2e2e-421e-9176-15287b5c22ce,d21ddf81-fcef-4e36-94e6-edd6fb487a31/drives/b!AE2tHi4uHkKRdhUoe1wizoHfHdLv_DZOlObt1vtIejFDr6vvuqdFTaTWzb63-TzY/items/01V67YTVHO2Y3JHJUM35EZTMY3LNCLRVNR/children/01V67YTVFEUVDEF5VSBRFLKMCSRJZZOCK6" -Body $body -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Post
#using Post causes "message": "Either 'folder' or 'file' must be provided, but not both.": https://stackoverflow.com/questions/35631616/move-item-doesnt-work-in-ms-graph-api-for-onedrive
#Documentation is using Patch method instead, tried this and results in "The parameter parentReference does not exist in method getByIdThenPath."
#Someone worked around this by using 
#>

