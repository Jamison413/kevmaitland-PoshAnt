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

<#----------------Connect to everything and load modules----------------#>

Import-Module _PNP_Library_SPO

#Set Variables to connect to Sharepoint

$SiteURL = "https://anthesisllc.sharepoint.com/teams/People_Services_Team_All_365"
$sharePointAdmin = "kimblebot@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content "$env:USERPROFILE\Desktop\KimbleBot.txt") 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass



#Connect to Sharepoint People Services Team (All)
Connect-PnPOnline -Url $SiteURL -Credentials $adminCreds
$context = Get-PnPContext

 




###############################################################################                                      
#                                                                             #
#                              New Starters                                   #
#                                                                             #
###############################################################################


#First, get all the items in the New Starters List
$List = "New Starter Details"
$FullItemQuery = Get-PnPListItem -List $List

$htmlfriendlytitle = $List -replace " ",'%20'


#Second, iterate through them and check for any upcoming/starting today
$upcomingNewStarters = @()
ForEach($NewStarter in $FullItemQuery){

        
      #check for null or it throws errors and blank roles
      If($NewStarter.FieldValues.StartDate){


        [datetime]$startdate = $NewStarter.FieldValues.StartDate #start date wil always be set to 23:00 by Sharepoint, hopefully will not cause issues?
        $todaysdate = Get-Date
    
            If(($startdate -gt $todaysdate) -or ($StartDate -eq $todaysdate)){

                $NewStarterLink = $SiteURL + "/Lists" + "/$($htmlfriendlytitle)" +  "/DispForm.aspx?" + "ID=$($NewStarter.FieldValues.ID)"

    
                 $upcomingNewStarters += New-Object psobject -Property @{
                'New Starter Name' = $NewStarter.FieldValues.Employee_x0020_Preferred_x0020_N
                'Job Title' = $NewStarter.FieldValues.JobTitle;
                'Start Date' = $NewStarter.FieldValues.StartDate;
                'Starting Office' = $NewStarter.FieldValues.Starting_x0020_Office0.Label;
                'Primary Office' = $NewStarter.FieldValues.Main_x0020_Office0.Label;
                'Line Manager' = $NewStarter.FieldValues.Line_x0020_Manager.LookupValue;
                'IT Setup Notes' = $NewStarter.FieldValues.IT_x0020_Setup_x0020_Notes;
                'People Services Setup Notes' = $NewStarter.FieldValues.People_x0020_Services_x0020_Setu;
                'Link' = $NewStarterLink
                }
        }
    }
}

#Convert it to an HTML table
$NewStartersHTML = $upcomingNewStarters  | ConvertTo-Html -Property "New Starter Name","Job Title","Start Date","Starting Office","Primary Office","Line Manager","IT Setup Notes","People Services Setup Notes","Link" -Head "<style>table, th, td {border: 1px solid;border-collapse: collapse ;padding: 5px;text-align: left;}</style>"
If(!$upcomingNewStarters){$NewStartersHTML = "Looks like there are no upcoming Leavers!"}



###############################################################################                                      
#                                                                             #
#                               Leavers                                       #
#                                                                             #
###############################################################################

#Get the full list of leavers
$List = "Notify Internal Teams of a Leaver"
$AllLeavers = Get-PnPListItem -List $List
$htmlfriendlytitle = $List -replace " ",'%20'


#Iterate through each leaver and figure out whether the leavving date it within the previous 10 days, or grater than the current date (to include reminders of people that have recently left).
$LiveLeavers = @()
ForEach ($Leaver in $AllLeavers){


      #check for null or it throws errors and blank roles
      If($NewStarter.FieldValues.StartDate){

            [datetime]$Leaversdate = $Leaver.FieldValues.Proposed_x0020_Leaving_x0020_Dat
            $thresholddate = (Get-Date) - ($timespan = New-TimeSpan -days 40)

                    If($Leaversdate -gt $thresholddate){

                    $LeaverLink = $SiteURL + "/Lists" + "/$($htmlfriendlytitle)" +  "/DispForm.aspx?" + "ID=$($Leaver.FieldValues.ID)"

                    $LiveLeavers += New-Object psobject -Property @{
                    'Employee Name' = $($Leaver.FieldValues.Employee_x0020_Name.Lookupvalue)
                    'Notes' = $($Leaver.FieldValues.Notes1)
                    'Proposed Leaving Date' = $($Leaver.FieldValues.Proposed_x0020_Leaving_x0020_Dat)
                    'Link' = $LeaverLink
                    }

            }
        }
}


#Convert it to an HTML table
$LeaversHTML = $LiveLeavers  | ConvertTo-Html -Property "Employee Name","Notes","Proposed Leaving Date","Link" #-Head "<style>table, th, td {border: 1px solid;border-collapse: collapse ;padding: 5px;text-align: left;}</style>"
If(!$LiveLeavers){$LeaversHTML = "Looks like there are no upcoming Leavers!"}



###############################################################################                                      
#                                                                             #
#                               Maternity Leave                               #
#                                                                             #
###############################################################################

#Get the full list of leavers
$List = "Notify of Maternity and Paternity Leave"
$AllMP = Get-PnPListItem -List $List
$htmlfriendlytitle1 = $List -replace " ",'%20'
$htmlfriendlytitle = $htmlfriendlytitle1 -replace "and",''

#Iterate through each leaver and figure out whether the leavving date it within the previous 10 days, or grater than the current date (to include reminders of people that have recently left).
$LiveMP = @()
ForEach ($MP in $AllMP){


      #check for null or it throws errors and blank roles
      If($MP.FieldValues.Proposed_x0020_Leaving_x0020_Dat){

            $MPstartdate = $MP.FieldValues.Proposed_x0020_Leaving_x0020_Dat  
            $MPreturndate = $MP.Proposed_x0020_Return_x0020_Date
            $todaysdate = (Get-Date)

                    If(($todaysdate -gt $MPstartdate) -and ($todaysdate -gt $MPreturndate)){

                    $MPLink = $SiteURL + "/Lists" + "/$($htmlfriendlytitle)" +  "/DispForm.aspx?" + "ID=$($MP.FieldValues.ID)"

                    $LiveMP += New-Object psobject -Property @{
                    'Employee Name' = $($MP.FieldValues.Employee_x0020_Name.Lookupvalue)
                    'Start Date' = $($MP.FieldValues.Proposed_x0020_Leaving_x0020_Dat)
                    'End Date' = $($MP.FieldValues.Proposed_x0020_Return_x0020_Date)
                    'Notes' = $($MP.FieldValues.Notes1)
                    'Status' = $($MP.FieldValues.Maternity_x0020_Paternity_x0020_)
                    'Link' = $MPLink
                    }

            }
        }
}


#Convert it to an HTML table
$MPHTML = $LiveMP | ConvertTo-Html -Property "Employee Name","Start Date","End Date","Notes","Status","Link" #-Head "<style>table, th, td {border: 1px solid;border-collapse: collapse ;padding: 5px;text-align: left;}</style>"
If(!$LiveMP){$MPHTML = "Looks like there are no upcoming Leavers!"}




##################################################################################################################################################                                     
#                                                                                                                                                #
#                                                         --Reporting Email--                                                                    #
#                                                                                                                                                #
##################################################################################################################################################



#Put it all in an email and send!
$subject = "Current People Services Portal Report"
            $body = "<HTML><FONT FACE=`"Calibri`">Hello People Services, Administration & IT Teams!`r`n`r`n<BR><BR>"
            $body += "`r`n`r`n<BR><BR>"
            $body += "This is a current report from the entire People Services Portal. If something is amiss, please make any changes in the relevant areas and this will reflect in the next report.`r`n`r`n<BR><BR>"
            $body += "<b>Here is a list of recent or upcoming New Starters on the People Services Site:</b>`r`n`r`n<BR><BR>"
            $body += "$NewStartersHTML`r`n`r`n<BR><BR><BR><BR>"
            $body += "<b>Here is a list of recent or upcoming Leavers on the People Services Site:</b>`r`n`r`n<BR><BR>"
            $body += "$LeaversHTML`r`n`r`n<BR><BR><BR><BR>"
            $body += "<b>Here is a list of live Maternity and Paternity leave on the People Services Site:</b>`r`n`r`n<BR><BR>"
            $body += "$MPHTML`r`n`r`n<BR><BR><BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>"
            $body += "The People Services Robot"
            $body += "`r`n`r`n<BR><BR>"
            $body += "`r`n`r`n<BR><BR>"
            $body += "If you have any issues accessing the links or information in this email, or have any feedback, please get in touch with IT.`r`n`r`n<BR><BR>"
            Write-Information $body

Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
Send-MailMessage -To "andrew.ost@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
Send-MailMessage -To "kevin.maitland@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
Send-MailMessage -To "nina.cairns@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
Send-MailMessage -To "elle.wright@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
Send-MailMessage -To "wai.cheung@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
Send-MailMessage -To "greg.francis@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8

