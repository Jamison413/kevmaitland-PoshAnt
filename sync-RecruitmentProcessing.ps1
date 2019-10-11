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

#Set Variables to connect to Sharepoint - this will be the HR Site soon
$SiteURL = "https://anthesisllc.sharepoint.com/teams/People_Services_Team_All_365"

$sharePointAdmin = "kimblebot@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\KimbleBot.txt") 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass


#Connect to Sharepoint - Groupbot? Couldn't work this bit out
Connect-PnPOnline -Url $SiteURL -Credentials $adminCreds
$context = Get-PnPContext



#Get the Processing List
$RequestItems = Get-PnPListItem -List "Recruitment Area" -Query "<View Scope='Processing'><Query><Where><Eq><FieldRef Name='IsDirty'/><Value Type='Text'>1</Value></Eq></Where></Query></View>"
 

#Build processing Array - grabs some useful information
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


#Pass details for each object above and create new list for each role to track candidates

ForEach ($Role in $ItemstoProcess){

$ListTitle = "$($Role.'ID')" + "  " + "$($Role.'Role Name')"

write-host "Creating new Candidate Tracker List for Role: ID$($Role.'ID') $($Role.'Role Name')" -ForegroundColor Yellow
    New-PnPList -Title "ID$($ListTitle)"  -Template GenericList
    Set-PnPList -Identity "ID$($ListTitle)" -Description "Live Candidate Tracker - RoleID:$($Role.'ID')" #Set the description to find this in our processing script, which searches for "Live Candidate Tracker" in the List Description, add RoleID to tie to Candidate Tracker

write-host "Adding Content Types and Views to list (also removing default views)'" -ForegroundColor Yellow
    Add-PnPContentTypeToList -List "ID$($ListTitle)" -ContentType "Candidate Tracker V2" -DefaultContentType
    
    Remove-PnPContentTypeFromList -List "ID$($ListTitle)" -ContentType "Item"
    Add-PnPView -List "ID$($ListTitle)" -Title "Candidates" -Fields "ID","Candidate Name","Interview 1: Date","Interview 1: Type","Interview 1: Feedback", "Decision 1","D1LE","Interview 2: Type","Interview 2: Feedback","Final Decision","FDLE","Offer Outcome","Proposed Start Date"
    #Add-PnPView -List "ID$($ListTitle)" -Title "People Services" -Fields "ID","Candidate Name","Recruiter","Candidate Source","Interview 1: Date","Interview 1: Type","Interview 1: Feedback", "Interview 1: Next Steps","Interview 2: Date","Interview 2: Type","Interview 2: Feedback","Final Decision","Offer Outcome","Proposed Start Date" #need to figure out how to restrict this view
    
    #Add-PnPView -List "ID$($ListTitle)" -Title "Processing" -Fields "ID","Candidate Name","Recruiter","Candidate Source","Interview 1: Date","Interview 1: Type","Interview 1: Feedback", "Interview 1: Next Steps","Interview 1: Next Steps_LastEntry", "Interview 1: Email  Processed","Interview 2: Date","Interview 2: Type","Interview 2: Feedback","Final Decision","Final Decision_LastEntry","Offer Outcome","Proposed Start Date", "Last Modified Date", "Modified" #need to figure out how to restrict this view
    Remove-PnPView -List "ID$($ListTitle)" -Identity "All Items" -Force
    Set-PnPView -List "ID$($ListTitle)" -Identity "Candidates" -Values @{DefaultView=$True}



write-host "Applying Hiring Manager Permissions" -ForegroundColor Yellow
    Set-PnPList -Identity "ID$($ListTitle)" -BreakRoleInheritance -ClearSubscopes
    Set-PnPListPermission -Identity "ID$($ListTitle)" -User $Role.'Hiring Manager Email' -AddRole "Contribute" #Hiring Manager Permissions
    Set-PnPListPermission -Identity "ID$($ListTitle)" -User nina.cairns@anthesisgroup.com -AddRole "Contribute" #People Services Permissions
    Set-PnPListPermission -Identity "ID$($ListTitle)" -User emily.pressey@anthesisgroup.com -AddRole "Full Control" #IT Permissions
    Set-PnPListPermission -Identity "ID$($ListTitle)" -User kevin.maitland@anthesisgroup.com -AddRole "Full Control" #IT Permissions


write-host "Setting item as processed in Recruitment Area: $($Role.'Role Name'). Setting link to Role Candidate Tracker." -ForegroundColor Yellow
    $CandidateListPathway = "$($SiteURL)" + "/Lists/" + "ID$($ListTitle)" + "/Candidates.aspx"
    $fullurl = [uri]::EscapeUriString($CandidateListPathway)
    Set-PnPListItem -List "Recruitment Area" -Identity $Role.ID -Values @{"IsDirty" = "0"}
    Set-PnPListItem -List "Recruitment Area" -Identity $Role.ID -Values @{"Candidate_x0020_Tracker" = "$($fullurl), ID$($ListTitle) Candidate Tracker"}
    Set-PnPListItem -List "Recruitment Area" -Identity $Role.ID -Values @{"Role_x0020_Hire_x0020_Status" = "Live"}
    
}


write-host "I currently have no error handling, so I don't know if I haven't worked! It would be worth checking the Recruitment Area and resulting List in the Site Contents" -ForegroundColor Red

   









