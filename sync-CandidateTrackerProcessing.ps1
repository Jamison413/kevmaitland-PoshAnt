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

$sharePointAdmin = "kimblebot@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\KimbleBot.txt") 
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
ForEach($LiveTracker in $LiveCandidateTrackers[6]){

    $Itemstoprocess = Get-PnPListItem -List $LiveTracker.Guid  
    
    foreach($Item in $Itemstoprocess){

        #I don't work, don't believe me - just compare the Decision Columns below instead...keeping this here in case I get fixed
        #$LastModifiedDate = $Item.FieldValues.Last_x0020_Modified_x0020_Date
        #$ModifiedDate = $Item.FieldValues.Modified
        #write-host "The last modified date of this item is older the the current Modified date, something has changed! Comparing the old entries to the new entries"

    #First, check for Declined Candidates so we can skip the iteration and then compare the key columns for changes for non-declined candidates, if there is a change, we can process it below.


        #Check for Declined candidates - if Decision 1 is set to Decline, set the Last Entry fields and Final Decision  fields also as Decline and continue onto the next iteration. Same concept for Final Decision, but this won't change Interview 1 decision fields, just the Last entry field.
        If(("Decline" -eq $Item.FieldValues.Decision_x0020_1) -and ("Decline" -ne $Item.FieldValues.D1LE)){
           Write-Host "Looks like this Candidate has been declined after Interview 1 since we last ran through - setting other fields to decline and moving on"
           Set-PnPListItem -List $LiveTracker.Guid -Identity $Item.ID -Values @{'D1LE' = "$($Item.FieldValues.Decision_x0020_1)"}
           Set-PnPListItem -List $LiveTracker.Guid -Identity $Item.ID -Values @{'FDLE' = "$($Item.FieldValues.Decision_x0020_1)"}
           Continue
        }

        If(("Decline" -eq $Item.FieldValues.Final_x0020_Decision) -and ("Decline" -ne $Item.FieldValues.FDLE)){
           Write-Host "Looks like this Candidate has been declined after final interview since we last ran through - setting other fields to decline and moving on"
           Set-PnPListItem -List $LiveTracker.Guid -Identity $Item.ID -Values @{'FDLE' = "$($Item.FieldValues.Final_x0020_Decision)"}
           Continue
        }

        #Compare the Decision Columns - only process them if they are different as it indicates a change.
        $InterView1Decision = (Compare-Object -ReferenceObject $Item.FieldValues.Decision_x0020_1 -DifferenceObject $Item.FieldValues.D1LE)
        $FinalDecision = (Compare-Object -ReferenceObject $Item.FieldValues.Final_x0020_Decision -DifferenceObject $Item.FieldValues.FDLE)



        If($InterView1Decision){
        Write-host "$($Item.FieldValues.Candidate_x0020_Name): Something has changed on the Interview 1 Decision Field! Let's maybe do something about it!" -ForegroundColor Yellow
        }
        If($FinalDecision){
        Write-host "$($Item.FieldValues.Candidate_x0020_Name): Something has changed on the Final Decision Field! Let's maybe do something about it!" -ForegroundColor Yellow
        }
        
    

    #Second, check which part has changed and action based on input. We have included an -and statement to just include those that have changed since the last run or it will keep sending out emails
        
        #Check if Interview 1 needs processing and it matches "Move to Next Stage" so we can let People Services know
        If(($InterView1Decision) -and ($Item.FieldValues.Decision_x0020_1 -match "Move to Next Stage")) {

        write-host "Interview 1 Decision has changed from $($Item.FieldValues.D1LE) to $($Item.FieldValues.Decision_x0020_1)"
            
            #Send email to People services letting them know to schedule a second interview
            $subject = "Recruitment Update: A Candidate is Ready to Move to Second Interview"
            $body = "<HTML><FONT FACE=`"Calibri`">Hello People Services Team,`r`n`r`n<BR><BR>"
            $body += "The Candidate $($Item.FieldValues.Candidate_x0020_Name) for role $($LiveTracker.Title) has been moved to the next stage.`r`n`r`n<BR><BR>"
            $body += "Please schedule an interview with the candidate and fill in the details of the date and type of interview in the candidate tracker.`r`n`r`n<BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>"
            $body += "The People Services Robot"

            Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8

            
            #Set the Decision 1 Last Entry column (D1LE) to the new Entry, this will stop it from re-processing - we don't want people getting multiple emails
            Set-PnPListItem -List $LiveTracker.Guid -Identity $Item.ID -Values @{'D1LE' = "$($Item.FieldValues.Decision_x0020_1)"}
            #Set-PnPListItem -List $LiveTracker.Guid -Identity $Item.ID -Values @{"Last_x0020_Modified_x0020_Date" = "$Item.FieldValues.Modified"}
            }

       }


       #Check if Final Decision needs processing and it matches "Make Offer" so we can let People Services know
        If(($FinalDecision) -and ($Item.FieldValues.Final_x0020_Decision -match "Make Offer")) {

        write-host "FinalDecision has changed from $($Item.FieldValues.FDLE) to $($Item.FieldValues.Final_x0020_Decision)"
            
            #Send email to People services letting them know to make an offer to this Candidate
            $subject = "Recruitment Update: A Candidate is Ready to Receive an Offer"
            $body = "<HTML><FONT FACE=`"Calibri`">Hello People Services Team,`r`n`r`n<BR><BR>"
            $body += "The Candidate $($Item.FieldValues.Candidate_x0020_Name) for role $($LiveTracker.Title) is ready to recieve an offer.`r`n`r`n<BR><BR>"
            $body += "Please send an offer to the candidate. The candidate tracker 'Offer Outcome' column has automatically been set to 'Pending'. Please set this to either 'Accepted' or 'Rejected' based on the Candidates response and enter the proposed starting date as soon as possible <b>as the last action on the Candidate Tracker, this will label the hiring process as complete</b>. This will inform our internal teams of an upcoming starter to ensure things like IT hardware is in-stock.`r`n`r`n<BR><BR>"
            $body += "When a proposed date is entered, a template entry will be added to the New Starter Form (you will receive an email with a link to this). This will use information we already know about the role and candidate, but will not be complete. Please fill this entry out as soon as possible so internal teams can set them up ready for their first day. `r`n`r`n<BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>"
            $body += "The People Services Robot"

            Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 

    
            #Set item 'Offer Outcome' to 'Pending', which People Services will change on Candidate response.
            Set-PnPListItem -List $LiveTracker.Guid -Identity $Item.ID -Values @{"Offer_x0020_Outcome" = "Pending"}
            Set-PnPListItem -List $LiveTracker.Guid -Identity $Item.ID -Values @{'FDLE' = "$($Item.FieldValues.Final_x0020_Decision)"}
            #Set-PnPListItem -List $LiveTracker.Guid -Identity $Item.ID -Values @{"$Item.FieldValues.Last_x0020_Modified_x0020_Date" = "$Item.FieldValues.Modified"}

            #Add the Candidate to a list so we can create their Employee folder later
            $Folderstocreate += New-Object psobject -Property @{"Candidate Name" = $Item.FieldValues.Candidate_x0020_Name}

        }


    #Third, check if there is a Proposed Start Date, this suggests some firm date has been set and the Candidate is likely to start then, create a new template entry with what we know about the Candidate already in the 'New Starter Details' List. Set the Candidate Tracker as "Complete"

        If(($Item.FieldValues.Proposed_x0020_Start_x0020_Date) -and ("0" -eq $Item.FieldValues.IsDirty)){
        Write-host "Looks like the Hiring Process is complete. Let's set this Candidate Tracker to 'Complete' and put a placeholder in the 'New Starter Details' Form based on what we know already" -ForegroundColor Yellow
            Set-PnPListItem -List $RecruitmentArea -Identity $LiveTracker.RoleId -Values @{"Role_x0020_Hire_x0020_Status" = "Complete"}
            $RecruitmentAreaItem = Get-PnPListItem -List $RecruitmentArea -Id $LiveTracker.RoleId

        #Start Pre-populating the New Starter Details Form
            Add-PnPListItem -List $NewStarterList -Values @{
            "Employee_x0020_Preferred_x0020_N" = $Item.FieldValues.Candidate_x0020_Name; 
            "StartDate" = $Item.FieldValues.Proposed_x0020_Start_x0020_Date;  
            "JobTitle" = $RecruitmentAreaItem.FieldValues.Role_x0020_Name;
            "Line_x0020_Manager" = $RecruitmentAreaItem.FieldValues.Hiring_x0020_Manager.LookupValue;
            "Primary_x0020_Team" = $RecruitmentAreaItem.FieldValues.Primary_x0020_Team0.Label;
            "Community0" = $RecruitmentAreaItem.FieldValues.Community0.Label;
            "Business_x0020_Unit0" = $RecruitmentAreaItem.FieldValues.Business_x0020_Unit0.Label;
            }
        
        #Send a confirmation email to People Services
            $subject = "Recruitment Update: A Candidate is set to start and a Template Entry has been added to the New Starter Details List"
            $body = "<HTML><FONT FACE=`"Calibri`">Hello People Services Team,`r`n`r`n<BR><BR>"
            $body += "The Candidate $($Item.FieldValues.Candidate_x0020_Name) for role $($LiveTracker.Title) now has a provisional start date!`r`n`r`n<BR><BR>"
            $body += "A new template entry has been created in the New Starter Details List, ready to be finished. Please finish this entry in good time before the start date so that Internal Teams can be ready for them to start.`r`n`r`n<BR><BR>"
            $body += "You can see the New Starter Details List here: https://anthesisllc.sharepoint.com/teams/People_Services_Team_All_365/Lists/New%20Starter%20Details/AllItems.aspx `r`n`r`n<BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>"
            $body += "The People Services Robot"

            Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 

            #Set is Dirty to 2 so it does not re-process (used 2 as a final action, 0 is the default for the column for other lists
            Set-PnPListItem -List $LiveTracker.Guid -Identity $Item.ID -Values @{'IsDirty' = "2"}

        }
    }



<#--------------Connect to the confidential HR team site with new pnp-context--------------#>  #Kimblebot is currently not allowed to connect to this site

#Set Variables to connect to Sharepoint confidential HR site 
#$SiteURL = "https://anthesisllc.sharepoint.com/sites/Confidential_Human_Resources_HR_Team_GBR_365/"



#Connect to Sharepoint
#Connect-PnPOnline -Url $SiteURL -Credentials $adminCreds
#$context = Get-PnPContext

<#--------------Process each new Employee folder--------------#>


<#ForEach ($folder in $Folderstocreate){

write-host "Creating employee folders on confidential HR site" -ForegroundColor Yellow

    Add-PnPFolder -Name $folder.'Candidate Name' -Folder "Shared Documents"
    $parentfolder = "Shared Documents" + "\" + $folder.'Candidate Name'
    Add-PnPFolder -Name "1. Onboarding" -Folder $parentfolder
    Add-PnPFolder -Name "2. Lifecycle" -Folder $parentfolder
    Add-PnPFolder -Name "3. Offboarding" -Folder $parentfolder
    
    $Onboardingfoldername = "Shared Documents" + "\" + $folder.'Candidate Name' + "1. Onboarding"
    Copy-PnPFile -SourceUrl "https://anthesisllc.sharepoint.com/:x:/r/sites/Confidential_Human_Resources_HR_Team_GBR_365/_layouts/15/Doc.aspx?sourcedoc=%7BAFD940AB-DED8-4C2B-BD2F-4AE144B72460%7D&file=New%20Starter%20Checklist.xlsx&action=default&mobileredirect=true" -TargetUrl $Onboardingfoldername
}

#>