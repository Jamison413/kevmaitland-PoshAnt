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
$sharePointAdminPass = ConvertTo-SecureString (Get-Content "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\KimbleBot.txt") 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass



#Connect to Sharepoint People Services Team (All)
Connect-PnPOnline -Url $SiteURL -Credentials $adminCreds
$context = Get-PnPContext

 


###############################################################################                                      
#                                                                             #
#                               Live Roles                                    #
#                                                                             #
###############################################################################


#Get the Processing List, all items with 'Live'

$FullListQuery = Get-PnPList

#Iterate through and process into a nice usable array, with lovely HTML friendly URL format

$LiveCandidateTrackers = @()
ForEach($List in $FullListQuery){
If($List.Description -match "Live Candidate Tracker"){

        $RoleId = ($($List.Description) -split ':')[1]
        $htmlfriendlytitle = $List.Title -replace " ",'%20'

        $LiveCandidateTrackers += New-Object psobject -Property @{
        'Title' = $List.Title;
        #'Guid' = $List.Id;
        #'Description' = $List.Description;
        'RoleID' = $RoleId;
        'Candidate Tracker Link' = $SiteURL + "/Lists" + "/$($htmlfriendlytitle)"
        
        }
     }
}

#Convert it into an HTML table
$LiveRolesHTML = $LiveCandidateTrackers | ConvertTo-Html -Property "Title","Candidate Tracker Link" -Head "<style>table, th, td {border: 1px solid;border-collapse: collapse ;padding: 5px;text-align: left;}</style>"


###############################################################################                                      
#                                                                             #
#                               Offers                                        #
#                                                                             #
###############################################################################


#First, get all Live Trackers

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

#Then check each tracker for Items, add them to a big array with all the details

$LiveOffers = @()
ForEach($LiveTracker in $LiveCandidateTrackers){

    $Items = Get-PnPListItem -List $LiveTracker.Guid  
    
        foreach($Candidate in $Items){

        $FinalDecision = $Candidate.FieldValues.Final_x0020_Decision
        $StartDate = $Candidate.FieldValues.Proposed_x0020_Start_x0020_Date

        #If 'Make Offer' and no start date, this indicates Offer is still pending
            If(("Make Offer" -eq $FinalDecision) -and ($null -eq $StartDate)){

                $htmlfriendlytitle = $LiveTracker.Title -replace " ",'%20'
        
                $LiveOffers += New-Object psobject -Property @{
                        
                        'Title' = $LiveTracker.Title;
                        'Candidate Name' = $($Candidate.FieldValues.Candidate_x0020_Name)
                        'Candidate Tracker Link' = $SiteURL + "/Lists" + "/$($htmlfriendlytitle)"
                                    }

        
            }

      }

}

$OffersHTML = $LiveOffers  | ConvertTo-Html -Property "Title","Candidate Name","Candidate Tracker Link" -Head "<style>table, th, td {border: 1px solid;border-collapse: collapse ;padding: 5px;text-align: left;}</style>"


###############################################################################                                      
#                                                                             #
#                              New Starters    #needs testing                 #
#                                                                             #
###############################################################################


#First, get all the items in the New Starters List

$FullItemQuery = Get-PnPListItem -List "New Starter Details"


#Second, iterate through them and check for any upcoming/starting today
$upcomingNewStarters = @()
ForEach($NewStarter in $FullItemQuery){


    

    [datetime]$startdate = $NewStarter.FieldValues.StartDate #start date wil always be set to 23:00 by Sharepoint, hopefully will not cause issues?
    $todaysdate = Get-Date

    If(($startdate -lt $todaysdate) -or ($StartDate -eq $todaysdate)){
    
                 $upcomingNewStarters += New-Object psobject -Property @{
                'New Starter Name' = $NewStarter.FieldValues.Employee_x0020_Preferred_x0020_N
                'Job Title' = $NewStarter.FieldValues.JobTitle;
                'Start Date' = $NewStarter.FieldValues.StartDate;
                'Starting Office' = $NewStarter.FieldValues.Starting_x0020_Office.Label0;
                'Primary Office' = $NewStarter.FieldValues.Main_x0020_Office0.Label;
                'Line Manager' = $NewStarter.FieldValues.Line_x0020_Manager.Label;
            }
    }
}




$NewStartersHTML = $upcomingNewStarters  | ConvertTo-Html -Property "'New Starter Name'","Job Title","Start Date","Starting Office","Primary Office","Line Manager" -Head "<style>table, th, td {border: 1px solid;border-collapse: collapse ;padding: 5px;text-align: left;}</style>"





















#Send an email!
$subject = "Current List of Live Offers"
            $body = "<HTML><FONT FACE=`"Calibri`">Hello People Services Team,`r`n`r`n<BR><BR>"
            $body += "<b>Here is a list of live Roles from Live Candidate Trackers on the People Services All Site.</b>`r`n`r`n<BR><BR>"
            $body += "$LiveRolesHTML`r`n`r`n<BR><BR><BR><BR>"
            $body += "<b>Here is a list of Live Offers from the Live Candidate Trackers on the People Services Site.</b>`r`n`r`n<BR><BR>"
            $body += "$OffersHTML`r`n`r`n<BR><BR><BR><BR>"
            $body += "<b>Here is a list of New Starters on the People Services Site.</b>`r`n`r`n<BR><BR>"
            $body += "$NewStartersHTML`r`n`r`n<BR><BR><BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>"
            $body += "The People Services Robot"
            
            Write-Information $body

Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
