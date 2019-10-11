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
$RecruitmentArea = Get-PnPList -Identity "Recruitment Area"
$NewStarterList = Get-PnPList -Identity "New Starter Details"

$items = Get-PnPListItem -List "Recruitment Area"



#Connect to Sharepoint - Groupbot? Couldn't work this bit out
Connect-PnPOnline -Url $SiteURL -Credentials $adminCreds
$context = Get-PnPContext


<#Get the master list of lists, have gone a weird way around this - can't figure out how to filter the list object directly so have used arrays#>
#Get all the Lists from the Site, find the live ones, "Live Candidate Tracker" will be in the description"


<#Testing

$list.FieldValues.Interview_x0020_1_x003A__x0020_N0 = Get-pnplistitem -List "ID23  Analyst 403"


#>

$FullListQuery = Get-PnPList
$LiveCandidateTrackers = @()
ForEach($List in $FullListQuery){
If($List.Description -match "Live Candidate Tracker"){

        $RoleId = $($List.Description) -split ':'

        $LiveCandidateTrackers += New-Object psobject -Property @{
        'Title' = $List.Title;
        'Guid' = $List.Id;
        'Description' = $List.Description;
        'RoleID' = $RoleId;
        
        }
     }
}

$thing.FieldValues.Final_x0020_Decision

#Iterate through each list and check for any actions against candidates need processing - is the date modified more recent than the Last Modified Date?
$Folderstocreate = @()
ForEach($LiveTracker in $LiveCandidateTrackers[6]){

    $Items = Get-PnPListItem -List $LiveTracker.Guid  
    
    foreach($thing in $Items){

        #First, check if there are any that have been modified recently - Ha! I don't work, don't believe me - just compare the Decision Columns...
    #$LastModifiedDate = $Item.FieldValues.Last_x0020_Modified_x0020_Date
    #$ModifiedDate = $Item.FieldValues.Modified
        #write-host "The last modified date of this item is older the the current Modified date, something has changed! Comparing the old entries to the new entries"


        #Compare the Decision Columns
        $InterView1Decision = (Compare-Object -ReferenceObject $thing.FieldValues.Decision_x0020_1 -DifferenceObject $thing.FieldValues.D1LE)
        $FinalDecision = (Compare-Object -ReferenceObject $thing.FieldValues.Final_x0020_Decision -DifferenceObject $thing.FieldValues.FDLE)

        If($InterView1Decision){
        Write-host "$($thing.FieldValues.Candidate_x0020_Name): Something has changed on the Interview 1 Decision Field! Let's maybe do something about it!" -ForegroundColor Yellow
        }
        If($FinalDecision){
        Write-host "$($thing.FieldValues.Candidate_x0020_Name): Something has changed on the Final Decision Field! Let's maybe do something about it!" -ForegroundColor Yellow
        }
        
    

    #Second, check which part has changed and action based on input. We have included an -and statement to just include those that have changed since the last run or it will keep sending out emails
        If(($InterView1Decision) -and ($thing.FieldValues.Decision_x0020_1 -match "Move to Next Stage")) {

        write-host "Interview 1 Decision has changed from $($thing.FieldValues.D1LE) to $($thing.FieldValues.Decision_x0020_1)"
            
            <#Send email to People services letting them know to schedule a second interview#>
            $subject = "Recruitment Update: A Candidate is Ready to Move to Second Interview"
            $body = "<HTML><FONT FACE=`"Calibri`">Hello People Services Team,`r`n`r`n<BR><BR>"
            $body += "The Candidate $($thing.FieldValues.Candidate_x0020_Name) for role $($LiveTracker.Title) has been moved to the next stage.`r`n`r`n<BR><BR>"
            $body += "Please schedue an interview with the candidate and fill in the details of the date and type of interview in the candidate tracker.`r`n`r`n<BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>"
            $body += "The People Services Robot"


            Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8

            
            <#Set the 'Next Steps_LastEntry' column to the new Entry, this will stop it from re-processing - we don't want people getting multiple emails, and the Last Modified Date column to the Modified date entry for the same purpose.#>
            Set-PnPListItem -List $LiveTracker.Guid -Identity $thing.ID -Values @{'D1LE' = "$($thing.FieldValues.Decision_x0020_1)"}
            #Set-PnPListItem -List $LiveTracker.Guid -Identity $Item.ID -Values @{"Last_x0020_Modified_x0020_Date" = "$Item.FieldValues.Modified"}
            }

       }
        If(($FinalDecision) -and ($thing.FieldValues.Final_x0020_Decision -match "Make Offer")) {

        write-host "FinalDecision has changed from $($thing.FieldValues.FDLE) to $($thing.FieldValues.Final_x0020_Decision)"
            
            <#Send email to People services letting them know to schedule a second interview#>
            $subject = "Recruitment Update: A Candidate is Ready to Receive an Offer"
            $body = "<HTML><FONT FACE=`"Calibri`">Hello People Services Team,`r`n`r`n<BR><BR>"
            $body += "The Candidate $($thing.FieldValues.Candidate_x0020_Name) for role $($LiveTracker.Title) has been moved to the next stage and are ready to recieve an offer.`r`n`r`n<BR><BR>"
            $body += "Please send an offer to the candidate. The candidate tracker 'Offer Outcome' column has automatically been set to 'Pending'. Please set this to either 'Accepted' or 'Rejected' based on the Candidates response. This will inform our internal teams of an upcoming starter (you will still need to fill out the new starter form, this will ensure things like IT hardware is in-stock).`r`n`r`n<BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>"
            $body += "The People Services Robot"

            Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 

    
            <#Set item 'Offer Outcome' to 'Pending', which People Services will change on Candidate response.#>
            Set-PnPListItem -List $LiveTracker.Guid -Identity $thing.ID -Values @{"Offer_x0020_Outcome" = "Pending"}
            #Set-PnPListItem -List $LiveTracker.Guid -Identity $Item.ID -Values @{"$Item.FieldValues.Last_x0020_Modified_x0020_Date" = "$Item.FieldValues.Modified"}


            $Folderstocreate += New-Object psobject -Property @{"Candidate Name" = $thing.FieldValues.Candidate_x0020_Name}

        }

        If($thing.FieldValues.Proposed_x0020_Start_x0020_Date){
        Write-host "Looks like the Hiring Process is complete. Let's set this Candidate Tracker to 'Complete' and put a placeholder in the 'New Starter Details' Form based on what we know already" -ForegroundColor Yellow
            Set-PnPListItem -List $RecruitmentArea -Identity $LiveTracker.RoleId -Values @{"Role_x0020_Hire_x0020_Status" = "Complete"}
            $RecruitmentAreaItem = Get-PnPListItem -List $RecruitmentArea -Id $LiveTracker.RoleId

        #Start Pre-populating the New Starter Details Form

            Add-PnPListItem -List $NewStarterList -Values @{
            "Employee_x0020_Preferred_x0020_N" = $thing.FieldValues.Candidate_x0020_Name; 
            "StartDate" = $thing.FieldValues.Proposed_x0020_Start_x0020_Date;  
            "JobTitle" = $RecruitmentAreaItem.FieldValues.Role_x0020_Name;
            "Line_x0020_Manager" = $RecruitmentAreaItem.FieldValues.Hiring_x0020_Manager;
            "Primary_x0020_Team" = $RecruitmentAreaItem.FieldValues.Primary_x0020_Team;
            "Community0" = $RecruitmentAreaItem.FieldValues.Community0 ;
            "Business_x0020_Unit0" = $RecruitmentAreaItem.FieldValues.Business_x0020_Unit;
            }
        
        #Send a confirmation email to People Services
            $subject = "Recruitment Update: A Candidate is set to start and a Template Entry has been added to the New Starter Details List"
            $body = "<HTML><FONT FACE=`"Calibri`">Hello People Services Team,`r`n`r`n<BR><BR>"
            $body += "The Candidate $($thing.FieldValues.Candidate_x0020_Name) for role $($LiveTracker.Title) now has a provisional start date!`r`n`r`n<BR><BR>"
            $body += "A new template entry has been created in the New Starter Details List, ready to be finished. Please finish this entry in good time before the start date so that Internal Teams can be ready for them to start.`r`n`r`n<BR><BR>"
            $body += "You can see the New Starter Details List here: https://anthesisllc.sharepoint.com/teams/People_Services_Team_All_365/Lists/New%20Starter%20Details/AllItems.aspx `r`n`r`n<BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>"
            $body += "The People Services Robot"

            Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 
        }
    }



<#Connect to the confidential HR team site with new pnp-context#>

#Set Variables to connect to Sharepoint confidential HR site 
$SiteURL = "https://anthesisllc.sharepoint.com/teams/Confidential_Human_Resources_HR_Team_GBR_365"

#Connect to Sharepoint - Groupbot? Couldn't work this bit out
Connect-PnPOnline -Url $SiteURL -Credentials $adminCreds
$context = Get-PnPContext


ForEach ($folder in $Folderstocreate){

write-host "Creating employee folders on confidential HR site" -ForegroundColor Yellow

    Add-PnPFolder -Name $folder.'Candidate Name' -Folder "Shared Documents"
    $parentfolder = "Shared Documents" + "\" + $folder.'Candidate Name'
    Add-PnPFolder -Name "1. Onboarding" -Folder $parentfolder
    Add-PnPFolder -Name "2. Lifecycle" -Folder $parentfolder
    Add-PnPFolder -Name "3. Offboarding" -Folder $parentfolder
    
    $Onboardingfoldername = "Shared Documents" + "\" + $folder.'Candidate Name' + "1. Onboarding"
    Copy-PnPFile -SourceUrl "https://anthesisllc.sharepoint.com/:x:/r/sites/Confidential_Human_Resources_HR_Team_GBR_365/_layouts/15/Doc.aspx?sourcedoc=%7BAFD940AB-DED8-4C2B-BD2F-4AE144B72460%7D&file=New%20Starter%20Checklist.xlsx&action=default&mobileredirect=true" -TargetUrl $Onboardingfoldername
}





<#------------------Future Development Potential------------------


#We could set the hire date on the SPO User profile - this is an existing attribute on the SPO profile Service. We can make more of these quite easily 

#Set Variables
$SiteURL = "https://anthesisllc.sharepoint.com/"
$UserAccount = "emily.pressey@anthesisgroup.com"
 
#Connect to PNP Online
Connect-PnPOnline -Url $SiteURL
 
#Get all properties of a User Profile
$UserProfile = Get-PnPUserProfileProperty -Account $UserAccount
$UserProfile.UserProfileProperties

#Update properties
Set-PnPUserProfileProperty -Account $UserAccount -PropertyName "Department" -Value "Operations - IT"

#>
    
    


