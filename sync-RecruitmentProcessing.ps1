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

<#------------------- Connect to everything and load modules -------------------#>

Import-Module _PNP_Library_SPO

#Set Variables to connect to Sharepoint - this will be the HR Site soon
$SiteURL = "https://anthesisllc.sharepoint.com/teams/People_Services_Team_All_365"

$sharePointAdmin = "kimblebot@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content "$env:USERPROFILE\Desktop\KimbleBot.txt") 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass

#Connect to Sharepoint
Connect-PnPOnline -Url $SiteURL -Credentials $adminCreds
$context = Get-PnPContext

<#------------------- Find what needs processing -------------------#>

#Query all items in the "Recruitment Area" List, find all ones where IsDirty -eq "1" - Microsoft Flow will label these as "1" after approval is accepted.
$RequestItems = Get-PnPListItem -List "Recruitment Area" -Query "<View Scope='Processing'><Query><Where><Eq><FieldRef Name='IsDirty'/><Value Type='Text'>1</Value></Eq></Where></Query></View>"
 

#Build processing Array from all the IsDirty items found above - grabs some useful information to help us create the Candidate Trackers
$ItemstoProcess = @()
ForEach ($Item in $RequestItems){

    $ItemstoProcess += New-Object psobject -Property @{

    'ID' = $Item.FieldValues.ID;
    'Role Name' = $Item.FieldValues.Role_x0020_Name;
    'IsDirty' = $Item.FieldValues.IsDirty;
    'GUID' = $Item.FieldValues.GUID;
    'Hiring Manager' = $Item.FieldValues.Hiring_x0020_Manager.LookupValue;
    'Hiring Manager Email' = $Item.FieldValues.Hiring_x0020_Manager.Email
        }
}



<#------------------- Process Each Approved Request -------------------#>

#Pass details for each item above and create a new list for each role to track candidates. These are templated lists from the Candidate Tracker V2 Content type.

ForEach ($Role in $ItemstoProcess){

#Create the list title and the time for logging
$ListTitle = "$($Role.'ID')" + "  " + "$($Role.'Role Name')"
$datetime = (Get-date)

    #Create the List
    write-host "Creating new Candidate Tracker List for Role: ID$($Role.'ID') $($Role.'Role Name')" -ForegroundColor Yellow
    New-PnPList -Title "ID$($ListTitle)"  -Template GenericList
    Set-PnPList -Identity "ID$($ListTitle)" -Description "Live Candidate Tracker - RoleID:$($Role.'ID')" #Set the description to find this in our processing script, which searches for "Live Candidate Tracker" in the List Description, add RoleID to tie to Candidate Tracker

    #Add the Candidate Tracker V2 Content Type    
    write-host "Adding Content Types and Views to list (also removing default views)'" -ForegroundColor Yellow
    Add-PnPContentTypeToList -List "ID$($ListTitle)" -ContentType "Candidate Tracker V2" -DefaultContentType

    #Remove the "All Items" default Content type or it will appear on forms. Create some useful views - "Candidates" is the default opening view.     
    Remove-PnPContentTypeFromList -List "ID$($ListTitle)" -ContentType "Item"
    Add-PnPView -List "ID$($ListTitle)" -Title "Candidates" -Fields "ID","Candidate Name","Interview 1: Date","Interview 1: Type","Interview 1: Feedback", "Decision 1","Interview 2: Type","Interview 2: Feedback","Final Decision","Offer Outcome","Proposed Start Date"
    Add-PnPView -List "ID$($ListTitle)" -Title "People Services" -Fields "ID","Candidate Name","Interview 1: Date","Interview 1: Type","Interview 1: Feedback", "Decision 1","D1LE","Interview 2: Type","Interview 2: Feedback","Final Decision","FDLE","Offer Outcome","Proposed Start Date","Recruiter","Candidate Source"
    Add-PnPView -List "ID$($ListTitle)" -Title "IT" -Fields "ID","Candidate Name","Interview 1: Date","Interview 1: Type","Interview 1: Feedback", "Decision 1","D1LE","Interview 2: Type","Interview 2: Feedback","Final Decision","FDLE","Offer Outcome","Proposed Start Date","Recruiter","Candidate Source","IsDirty"
    Remove-PnPView -List "ID$($ListTitle)" -Identity "All Items" -Force
    Set-PnPView -List "ID$($ListTitle)" -Identity "Candidates" -Values @{DefaultView=$True}


    #Apply unique permisisons, add People Services, add IT, and add the Hiring Manager.
    write-host "Applying Hiring Manager Permissions" -ForegroundColor Yellow
    Set-PnPList -Identity "ID$($ListTitle)" -BreakRoleInheritance -ClearSubscopes
    Set-PnPListPermission -Identity "ID$($ListTitle)" -User $Role.'Hiring Manager Email' -AddRole "Contribute" #Hiring Manager Permissions
    Set-PnPListPermission -Identity "ID$($ListTitle)" -User nina.cairns@anthesisgroup.com -AddRole "Contribute" #People Services Permissions
    Set-PnPListPermission -Identity "ID$($ListTitle)" -User emily.pressey@anthesisgroup.com -AddRole "Full Control" #IT Permissions
    Set-PnPListPermission -Identity "ID$($ListTitle)" -User kevin.maitland@anthesisgroup.com -AddRole "Full Control" #IT Permissions

    #Wrap up the process by adding a link to the Candidate Tracker in it's item in the Recruitment Area, set the role to "Live" and label IsDirty as "0" to stop it from reprocessing
    write-host "Setting item as processed in Recruitment Area: $($Role.'Role Name'). Setting link to Role Candidate Tracker." -ForegroundColor Yellow
    $CandidateListPathway = "$($SiteURL)" + "/Lists/" + "ID$($ListTitle)" + "/Candidates.aspx"
    $fullurl = [uri]::EscapeUriString($CandidateListPathway)
    Set-PnPListItem -List "Recruitment Area" -Identity $Role.ID -Values @{"IsDirty" = "0"}
    Set-PnPListItem -List "Recruitment Area" -Identity $Role.ID -Values @{"Candidate_x0020_Tracker" = "$($fullurl), ID$($ListTitle) Candidate Tracker"}
    Set-PnPListItem -List "Recruitment Area" -Identity $Role.ID -Values @{"Role_x0020_Hire_x0020_Status" = "Live"}
    $link = "<a href=$($fullurl)>ID$($ListTitle)</a>"


    #Check for success and send an email if successful
    $currentlist = Get-PnPList -Identity "ID$($ListTitle)"
    If($currentlist){
    write-host "It looks like the correct list was made:" "ID$($ListTitle)" -ForegroundColor Yellow
    
    #Send a success email
            $subject = "Recruitment Update: A Candidate Tracker has been made for role " + "ID$($ListTitle)"
            $body = "<HTML><FONT FACE=`"Calibri`">Hello People Services and IT Team,`r`n`r`n<BR><BR>"
            $body += "This email is just to let you know that it looks like a Candidate Tracker has been successfully created: $link`r`n`r`n<BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>"
            $body += "The People Services Robot<BR><BR><BR><BR>"
            $body += "*Please note, this is an automated email. If you notice any issues, please get in touch with the IT Team"
            Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
    }

    Else{
    Write-Host "Woops, looks like something has gone wrong" -ForegroundColor Yellow

    #Send a failure email if not successful to IT
            $subject = "Failure: Recruitment Processing - a Candidate Tracker has *not* been made for role " + "ID$($ListTitle)"
            $body = "<HTML><FONT FACE=`"Calibri`">Hello IT Team,`r`n`r`n<BR><BR>"
            $body += "This email is just to let you know that it looks like a Candidate Tracker <b>has not been successfully created:</b> " + "ID$($ListTitle)`r`n`r`n<BR><BR>"
            $body += "Timestamp: " + "<b>$datetime</b>`r`n`r`n<BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>"
            $body += "The People Services Robot<BR><BR><BR><BR>"
            $body += "*Please note, this is an automated email. If you notice any issues, please get in touch with the IT Team"
            
            Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
    }
    
}





   









