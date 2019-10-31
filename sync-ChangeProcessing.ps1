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
$credential = Import-CliXml -Path 'C:\Users\Emily.Pressey\Desktop\JScred'


<#--------------Connect to Jira--------------#>

Set-JiraConfigServer 'https://anthesisit.atlassian.net'
New-JiraSession -Credential $credential


#######################################################################################
#                                                                                     #
#                               New Starters List Processing                          #
#                                                                                     #
#######################################################################################

                                                                    <#----------Sequential Evevnts----------#>

# - Candidate Tracker has proposal Date set on item, creates template entry in New Starters List OR someone manually adds New Starter form scratch with no Candidate Tracker
# - Microsoft Flow creates new Calendar entry in the Starters, Changers, Leavers Calendar
# - IT, Admin and People Services recieves an email
# - If the start date changes (found by comparison columns), then Powershell sets the FlowTrigger column to 'Change' to set off the Calendar Management Flow.

#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------#


#Set Variables to connect to Sharepoint - People Services (All) and New Starter Details List
$SiteURL = "https://anthesisllc.sharepoint.com/teams/People_Services_Team_All_365"
$List = "New Starter Details"

Connect-PnPOnline -Credentials $adminCreds -Url $SiteURL
$context = Get-PnPContext

#Get all the items
$AllNewStartersitems = Get-PnPListItem -List $List

<#--------------New Starter Intitial Notification Email---------------#>

#If PowersehllTrigger is set to "1"
$NewStarterInformation = @()
ForEach($Item in $AllNewStartersitems){

$NewStarterInformation = @()

If("1" -eq $Item.FieldValues.PowershellTrigger){
            
            
            write-host "An item has been added, and needs processing! Let's send an email to IT and People Services" -ForegroundColor Yellow

            #Get the information for the New Starter, convert it to an HTML table, create a friendly link to the item and send an email

            $NewStarterInformation += New-Object psobject -Property @{

            "Employee Preferred Name" = $($Item.FieldValues.Employee_x0020_Preferred_x0020_N); 
            "Start Date" = $($Item.FieldValues.StartDate);  
            "Job Title" = $($Item.FieldValues.JobTitle);
            "Line Manager" = $($Item.FieldValues.Hiring_x0020_Manager.LookupValue);
            "Primary Team" = $($Item.FieldValues.Primary_x0020_Team0.Label);
            "Community" = $($Item.FieldValues.Community0.Label);
            "Business Unit" = $($Item.FieldValues.Business_x0020_Unit0.Label);
            "Starting Office" = $($Item.FieldValues.Starting_x0020_Office0.Label);
            }

            $NewStarterHTML = $NewStarterInformation | ConvertTo-Html -Property "Employee Preferred Name","Start Date","Job Title","Line Manager","Primary Team","Community","Business Unit","Starting Office" -Head "<style>table, th, td {border: 1px solid;border-collapse: collapse ;padding: 5px;text-align: left;}</style>"

            $htmlfriendlytitle = $List -replace " ",'%20'
            $StarterItemLink = $SiteURL + "/Lists" + "/$($htmlfriendlytitle)" +  "/DispForm.aspx?" + "ID=$($Item.FieldValues.ID)"
            #[datetime]$date = $($Item.FieldValues.StartDate)
            #$date = $($Item.FieldValues.StartDate) -split "/"

            #Send an email to People Services and IT to notify of the change and to make the change 

            $subject = "New Starters Update: A New Starter has been Added!"
            $body = "<HTML><FONT FACE=`"Calibri`">Hello People Services & IT Teams,`r`n`r`n<BR><BR>"
            $body += "You're receiving this email as someone has added a New Starter to the New Starters List; a new entry will be added to the New Starters, Changers and Leavers Shared Calendar. Here is some information about them:`r`n`r`n<BR><BR>"
            $body += "$($NewStarterHTML)`r`n`r`n<BR><BR>"
            $body += "You can see more information about the New Starter here: $($StarterItemLink)`r`n`r`n<BR><BR><BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>"
            $body += "The People Services Robot"

            Send-MailMessage -To "IT_Team_GBR_365@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "nina.cairns@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "elle.wright@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "wai.cheung@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "greg.francis@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 
            Set-PnPListItem -List $List -Identity $Item.ID -Values @{"PowershellTrigger" = "0"}

            #We will also try to make a Jira task for the IT Team
            If("1" -eq $Item.FieldValues.JiraTaskCreated){
            #Set the fields for Jira Ticket
                    $fields = @{
                    #Workplace field
                    customfield_10045 = @{
                    value = "Bristol, GBR"
                    id = "10110"
                    }
                    #IT Team Responsible
                    customfield_10048 = @{
                    value = "Bristol"
                    id = "10129"
                    }
                    #customfield_10010 = @{
                    #id = "46"
                    #issueTypeId = "10002"
                    #}  
                    }
                    #Create Jira ticket, if created set JiraTaskCreated column to "0" so it doesn't re-create (finding exisitng Jira Tickets is a pain so have opted for this route instead).
                    Write-host "Creating Jira Ticket for New Starter Request: $($Item.FieldValues.Employee_x0020_Preferred_x0020_N)" -ForegroundColor Yellow
                    $newissue = New-JiraIssue -Project ITC -IssueType 'Service Request' -Summary "New Starter Request: $($Item.FieldValues.Employee_x0020_Preferred_x0020_N)" -Description $($StarterItemLink) -Fields $fields

            If($newissue){
            Set-PnPListItem -List $List -Identity $Item.ID -Values @{"JiraTaskCreated" = "0"}
                         }
            Else{
            Write-Host "Woops, something went wrong whilst creating a Jira Ticket for New Starter Request: $($Item.FieldValues.Employee_x0020_Preferred_x0020_N)"
            }
            }
            

}
Else{
write-host "Looks like there are no new starters" -ForegroundColor Yellow
}
}


<#--------------Start Date Change Processing---------------#>

 #Iterate through each item and see if anything has changed by comparing the Start Date and Last_Start Date columns.
    ForEach($item in $AllNewStartersitems){
    
   #I don't work at the moment
    #$LastModifiedDate = $Item.FieldValues.Last_x0020_Modified_x0020_Date
    #$ModifiedDate = $Item.FieldValues.Modified
        #If($ModifiedDate -gt $LastModifiedDate){
        #Compare the live and last entry columns
        #write-host "The last modified date of this item is older the the current Modified date, something has changed! Comparing the old entries to the new entries" -ForegroundColor Yellow
           
        #Format the relevant fields - Sharepoint gets confused with DateTime
        [datetime]$startdateformat = $($Item.FieldValues.StartDate)
        
        #If there is no Last Start Date, then set the Last Start Date to the same as the current Start Date and then skip over this iteration onto the next element.
        If(!$Item.FieldValues.Last_StartDate){
            Set-PnPListItem -List $List -Identity $item.ID -Values @{"Last_StartDate" = "$startdateformat"}
            Continue
            }
        Else{
        #If there is a Last Start Date then compare the two and see if it is different, because this implies that the Start Date has changed.
            [string]$Startdate = $Item.FieldValues.StartDate
            [string]$Last_StartDate = $Item.FieldValues.Last_StartDate
            $Startdatecomparison = (Compare-Object -ReferenceObject $Startdate -DifferenceObject $Last_StartDate)
            }        
        #Check if there is a difference, if there $startdate variable is null, there is no change, if there is something in there, then looks like there must be a change. Set the Last start Date to the Current Start Date and amend the FlowTrigger to set of the Calendar Management Flow.
        If($Startdatecomparison){
        Write-host "There has been a change to the Start Date: '$($Item.FieldValues.Employee_x0020_Preferred_x0020_N)'" -ForegroundColor Yellow

        #Send email letting people know
                    $subject = "New Starters Update: The Start Date for $($Item.FieldValues.Employee_x0020_Preferred_x0020_N) has been changed!"
            $body = "<HTML><FONT FACE=`"Calibri`">Hello People Services & IT Teams,`r`n`r`n<BR><BR>"
            $body += "You're receiving this email as someone has changed the Start Date for a New Starter. We'll try to update the calendar entry in the New Starters and Leavers shared calendar.`r`n`r`n<BR><BR>"
            $body += "$($Item.FieldValues.Employee_x0020_Preferred_x0020_N): From $($Item.FieldValues.Last_StartDate) to $($Item.FieldValues.StartDate).`r`n`r`n<BR><BR>"
            $body += "You can see more information about the New Starter here: $($StarterItemLink)`r`n`r`n<BR><BR><BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>"
            $body += "The People Services Robot"

            Send-MailMessage -To "IT_Team_GBR_365@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "nina.cairns@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "elle.wright@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "wai.cheung@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "greg.francis@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 
            
            
            Set-PnPListItem -List $List -Identity $Item.ID -Values @{"PowershellTrigger" = "0"}
            Set-PnPListItem -List $List -Identity $item.ID -Values @{"Last_StartDate" = "$startdateformat"}
            Set-PnPListItem -List $List -Identity $item.ID -Values @{"FlowTrigger" = "Change"}
         }

}


#######################################################################################
#                                                                                     #
#                              Leavers List Processing                                #
#                                                                                     #
#######################################################################################

                                                                    <#----------Sequential Evevnts----------#>


# - Microsoft Flow creates new Calendar entry in the Starters, Changers, Leavers Calendar
# - IT, Admin and People Services recieves an email
# - If the leave date changes (found by comparison columns), then Powershell sets the FlowTrigger column to 'Change' to set off the Calendar Management Flow.

#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------#


#Set Variables to connect to Sharepoint - People Services (All) and Notify Internal Teams of a Leaver
$SiteURL = "https://anthesisllc.sharepoint.com/teams/People_Services_Team_All_365"
$List = "Notify Internal Teams of a Leaver"

#Connect to Sharepoint - Groupbot? Couldn't work this bit out
Connect-PnPOnline -Credentials $adminCreds -Url $SiteURL
$context = Get-PnPContext

#Get all the items
$AllLeaversitems = Get-PnPListItem -List $List  

<#--------------New Leaver Intitial Notification Email---------------#>
$NewLeaverInformation = @()
ForEach($Item in $AllLeaversitems){

    If("1" -eq $item.FieldValues.PowershellTrigger){
            
            
            write-host "An item has been added, and needs processing: '$($Item.FieldValues.Employee_x0020_Name.LookupValue)'. Let's send an email to IT and People Services" -ForegroundColor Yellow

            #Get the information for the New Starter, convert it to an HTML table, create a friendly link to the item and send an email
            $NewLeaverInformation += New-Object psobject -Property @{

            "Employee Name" = $($Item.FieldValues.Employee_x0020_Name.LookupValue); 
            "Notes" = $($Item.FieldValues.Notes1);
            "Proposed Leaving Date" = $($Item.FieldValues.Proposed_x0020_Leaving_x0020_Dat);
            }

            $NewLeaverHTML = $NewLeaverInformation | ConvertTo-Html -Property "Employee Name","Notes","Proposed Leaving Date" -Head "<style>table, th, td {border: 1px solid;border-collapse: collapse ;padding: 5px;text-align: left;}</style>"

            $htmlfriendlytitle = $List -replace " ",'%20'
            $LeaverItemLink = $SiteURL + "/Lists" + "/$($htmlfriendlytitle)" +  "/DispForm.aspx?" + "ID=$($Item.FieldValues.ID)"

            #Send an email to People Services and IT to notify of the change and to make the change 
            $subject = "Leavers Update: A New Leaver has been Added"
            $body = "<HTML><FONT FACE=`"Calibri`">Hello People Services & IT Teams,`r`n`r`n<BR><BR>"
            $body += "You're receiving this email as someone has added a New Leaver to the Leavers List; a new entry will be added to the New Starters, Changers and Leavers Shared Calendar. Here is some information about them:`r`n`r`n<BR><BR>"
            $body += "$($NewLeaverHTML)`r`n`r`n<BR><BR>"
            $body += "They will need to be de-provisioned on the leaving date.`r`n`r`n<BR><BR>"
            $body += "You can see more information about the New Leaver here: $($LeaverItemLink)`r`n`r`n<BR><BR><BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>"
            $body += "The People Services Robot"


            Send-MailMessage -To "IT_Team_GBR_365@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "nina.cairns@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "elle.wright@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "wai.cheung@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "greg.francis@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 
            Set-PnPListItem -List $List -Identity $item.ID -Values @{"PowershellTrigger" = "0"}
            }

            If("1" -eq $Item.FieldValues.JiraTaskCreated){
            #Set the fields for Jira Ticket
                    $fields = @{
                    #Workplace field
                    customfield_10045 = @{
                    value = "Bristol, GBR"
                    id = "10110"
                    }
                    #IT Team Responsible
                    customfield_10048 = @{
                    value = "Bristol"
                    id = "10129"
                    }
                    }
                    #Create Jira ticket, if created set JiraTaskCreated column to "0" so it doesn't re-create (finding exisitng Jira Tickets is a pain so have opted for this route instead).
                    Write-host "Creating Jira Ticket for Leaver Request: $($Item.FieldValues.Employee_x0020_Name.LookupValue)" -ForegroundColor Yellow
                    $newissue = New-JiraIssue -Project ITC -IssueType 'Service Request' -Summary "Leaver Request: $($Item.FieldValues.Employee_x0020_Name.LookupValue)" -Description "Proposed Leaving Date: $($Item.FieldValues.Proposed_x0020_Leaving_x0020_Dat)` $($LeaverItemLink)" -Fields $fields

            If($newissue){
            Write-Host "Success! Jira ticket created"
            Set-PnPListItem -List $List -Identity $Item.ID -Values @{"JiraTaskCreated" = "0"}
                         }
            Else{
            Write-Host "Woops, something went wrong whilst creating a Jira Ticket for Leaver Request: $($Item.FieldValues.Employee_x0020_Name.LookupValue)"
            }
            }


Else{
write-host "Looks like there are no new leavers" -ForegroundColor Yellow
}

}


<#--------------Leave Date Change Processing---------------#>

#Iterate through each item and see if anything has changed by comparing the Leaving Date and Last_Leaving Date columns.
    ForEach($item in $AllLeaversitems){
    
    #$LastModifiedDate = $Item.FieldValues.Last_x0020_Modified_x0020_Date
    #$ModifiedDate = $Item.FieldValues.Modified
        #If($ModifiedDate -gt $LastModifiedDate){
        #Compare the live and last entry columns
        #write-host "The last modified date of this item is older the the current Modified date, something has changed! Comparing the old entries to the new entries" -ForegroundColor Yellow
            #$Leaverdate = (Compare-Object -ReferenceObject $Item.FieldValues.Proposed_x0020_Leaving_x0020_Dat -DifferenceObject $Item.Last_LeavingDate)
        #}
        #Else{
        #Write-Host "Looks like nothing has been modified!"
        #}

         #Format the relevant fields - Sharepoint gets confused with DateTime
        [datetime]$Leavedateformat = $($Item.FieldValues.Proposed_x0020_Leaving_x0020_Dat)

        If(!$Item.FieldValues.Last_LeavingDate){
            write-host "Looks like there was no Last Leaving Date recording for '$($Item.FieldValues.Employee_x0020_Name.LookupValue)', will record one now" -ForegroundColor Yellow
            Set-PnPListItem -List $List -Identity $item.ID -Values @{"Last_LeavingDate" = "$Leavedateformat"}
            Continue
            }
        Else{
        #If there is a Last Leaving Date then compare the two and see if it is different, because this implies that the Leaving Date has changed.
            [string]$Leavingdate = $Item.FieldValues.Proposed_x0020_Leaving_x0020_Dat
            [string]$Last_LeavingDate = $Item.FieldValues.Last_LeavingDate
            $Leavingdate = (Compare-Object -ReferenceObject $Leavingdate -DifferenceObject $Last_LeavingDate)
            }        

        #If the leaving date is different, set the last leaving date as current leaving date for next run and amend the FlowTrigger to set of the Calendar Management Flow.
        If($Leavingdate){
            Write-host "There has been a change to the End date" -ForegroundColor Yellow
            Set-PnPListItem -List $List -Identity $item.ID -Values @{"Last_LeavingDate" = "$Leavedateformat"}
            Set-PnPListItem -List $List -Identity $item.ID -Values @{"FlowTrigger" = "Change"}
        }

}



#######################################################################################
#                                                                                     #
#                       Changers List Processing                                      #
#                                                                                     #
#######################################################################################

                                                                    <#----------Sequential Evevnts----------#>


# - Microsoft Flow creates new Calendar entry in the Starters, Changers, Leavers Calendar
# - IT, Admin and People Services recieves an email

#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------#


#Set Variables to connect to Sharepoint - People Services (All) and Notify Internal Teams of a Leaver
$SiteURL = "https://anthesisllc.sharepoint.com/teams/People_Services_Team_All_365"
$List = "Request Change for Employee"

#Connect to Sharepoint
Connect-PnPOnline -Credentials $adminCreds -Url $SiteURL
$context = Get-PnPContext


#Get all the items
$AllChangersitems = Get-PnPListItem -List $List

#Check for items that need processing
ForEach($item in $AllChangersitems){

    $htmlfriendlytitle = $List -replace " ",'%20'
    $ChangeRequestLink = $SiteURL + "/Lists" + "/$($htmlfriendlytitle)" +  "/DispForm.aspx?" + "ID=$($item.FieldValues.ID)"

    If("1" -eq $item.FieldValues.IsDirty){
            #Send an email to People Services and IT to notify of the change and to make the change 
            write-host "An item has been added, and needs processing! Let's send an email to IT and People Services" -ForegroundColor Yellow
            $subject = "Changers Update: A Request has Been Made to Change an Employee!"
            $body = "<HTML><FONT FACE=`"Calibri`">Hello People Services & IT Teams,`r`n`r`n<BR><BR>"
            $body += "You're receiving this email as someone has requested a change to be made to an Employee (these are usually changes in licensing or access requirements).`r`n`r`n<BR><BR>"
            $body += "<b>Here is a description of the change:</b>.`r`n`r`n<BR><BR>"
            $body += "$($item.FieldValues.Change_x0020_Description)`r`n`r`n<BR><BR><BR><BR>"
            $body += "Please make a change to the item in the 'Request Change for Employee' List on the People Services (All) Site when the change is applied.`r`n`r`n<BR><BR>"
            $body += "You can see more information about this request here: $($ChangeRequestLink)`r`n`r`n<BR><BR><BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>"
            $body += "The People Services Robot"

            Send-MailMessage -To "IT_Team_GBR_365@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "nina.cairns@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "elle.wright@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "wai.cheung@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "greg.francis@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 
            Set-PnPListItem -List $List -Identity $item.ID -Values @{"IsDirty" = "0"}
            }
    

    }




#######################################################################################
#                                                                                     #
#                       Maternity/Paternity List Processing                           #
#                                                                                     #
#######################################################################################

#Set Variables to connect to Sharepoint - People Services (All) and Notify Internal Teams of a Leaver
$SiteURL = "https://anthesisllc.sharepoint.com/teams/People_Services_Team_All_365"
$List = "Notify of Maternity / Paternity Leave"

#Connect to Sharepoint
Connect-PnPOnline -Credentials $adminCreds -Url $SiteURL
$context = Get-PnPContext


#Get all the items
$AllMatPatitems = Get-PnPListItem -List $List

<#
        #Format the relevant fields - Sharepoint gets confused with DateTime
        [datetime]$startdateformat = $($Item.FieldValues.StartDate)
        
        #If there is no Last Start Date, then set the Last Start Date to the same as the current Start Date and then skip over this iteration onto the next element.
        If(!$Item.FieldValues.Last_StartDate){
            Set-PnPListItem -List $List -Identity $item.ID -Values @{"Last_StartDate" = "$startdateformat"}
            Continue
            }
        Else{
        #If there is a Last Start Date then compare the two and see if it is different, because this implies that the Start Date has changed.
            [string]$Startdate = $Item.FieldValues.StartDate
            [string]$Last_StartDate = $Item.FieldValues.Last_StartDate
            $Startdatecomparison = (Compare-Object -ReferenceObject $Startdate -DifferenceObject $Last_StartDate)
            }        
        #Check if there is a difference, if there $startdate variable is null, there is no change, if there is something in there, then looks like there must be a change. Set the Last start Date to the Current Start Date and amend the FlowTrigger to set of the Calendar Management Flow.
        If($Startdatecomparison){
        Write-host "There has been a change to the Start Date: '$($Item.FieldValues.Employee_x0020_Preferred_x0020_N)'" -ForegroundColor Yellow

        #Send email letting people know
                    $subject = "New Starters Update: The Start Date for $($Item.FieldValues.Employee_x0020_Preferred_x0020_N) has been changed!"
            $body = "<HTML><FONT FACE=`"Calibri`">Hello People Services & IT Teams,`r`n`r`n<BR><BR>"
            $body += "You're receiving this email as someone has changed the Start Date for a New Starter. We'll try to update the calendar entry in the New Starters and Leavers shared calendar.`r`n`r`n<BR><BR>"
            $body += "$($Item.FieldValues.Employee_x0020_Preferred_x0020_N): From $($Item.FieldValues.Last_StartDate) to $($Item.FieldValues.StartDate).`r`n`r`n<BR><BR>"
            $body += "You can see more information about the New Starter here: $($StarterItemLink)`r`n`r`n<BR><BR><BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>"
            $body += "The People Services Robot"

            Send-MailMessage -To "IT_Team_GBR_365@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "nina.cairns@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "elle.wright@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "wai.cheung@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "greg.francis@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
            Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 
            
            
            Set-PnPListItem -List $List -Identity $Item.ID -Values @{"PowershellTrigger" = "0"}
            Set-PnPListItem -List $List -Identity $item.ID -Values @{"Last_StartDate" = "$startdateformat"}
            Set-PnPListItem -List $List -Identity $item.ID -Values @{"FlowTrigger" = "Change"}
         }


         #>