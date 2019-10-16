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




#######################################################################################
#                                                                                     #
#                       New Starters List Processing                                  #
#                                                                                     #
#######################################################################################

#Set Variables to connect to Sharepoint - People Services (All) and New Starter Details List
$SiteURL = "https://anthesisllc.sharepoint.com/teams/People_Services_Team_All_365"
$List = "New Starter Details"

Connect-PnPOnline -Credentials $adminCreds -Url $SiteURL
$context = Get-PnPContext


#Get all the items
$AllNewStartersitems = Get-PnPListItem -List $List
    
    #First, check if there are any that have been modified recently
    ForEach($item in $AllNewStartersitems){

   #I don't work at the moment
    #$LastModifiedDate = $Item.FieldValues.Last_x0020_Modified_x0020_Date
    #$ModifiedDate = $Item.FieldValues.Modified
        #If($ModifiedDate -gt $LastModifiedDate){
        #Compare the live and last entry columns
        #write-host "The last modified date of this item is older the the current Modified date, something has changed! Comparing the old entries to the new entries" -ForegroundColor Yellow
           

        #If there is no Last Start Date, then set the Last Start Date to the same as the current Start Date and then skip over this iteration onto the next element.
        If(!$Item.FieldValues.Last_StartDate){
            Set-PnPListItem -List $List -Identity $item.ID -Values @{"$Item.FieldValues.Last_StartDate" = "$Item.FieldValues.StartDate"}
            Continue
            }
        Else{
        #If there is a Last Start Date then compare the two and see if it is different, because this implies that the Start Date has changed.
            $Startdate = (Compare-Object -ReferenceObject $Item.FieldValues.StartDate -DifferenceObject $Item.FieldValues.Last_StartDate)
            }        
        #Check if there is a difference, if there $startdate variable is null, there is no change, if there is something in there, then looks like there must be a change. Set the Last start Date to the Current Start Date and amend the FlowTrigger to set of the Calendar Management Flow.
        If($Startdate){
        Write-host "There has been a change to the Start Date" -ForegroundColor Yellow
        Set-PnPListItem -List $List -Identity $item.ID -Values @{"$Item.FieldValues.Last_StartDate" = "$Item.FieldValues.StartDate"}
        Set-PnPListItem -List $List -Identity $item.ID -Values @{"$Item.FieldValues.FlowTrigger" = "Change"}
        }

   }


#######################################################################################
#                                                                                     #
#                       Leavers List Processing                                       #
#                                                                                     #
#######################################################################################

#Set Variables to connect to Sharepoint - People Services (All) and Notify Internal Teams of a Leaver
$SiteURL = "https://anthesisllc.sharepoint.com/teams/People_Services_Team_All_365"
$List = "Notify Internal Teams of a Leaver"

#Connect to Sharepoint - Groupbot? Couldn't work this bit out
Connect-PnPOnline -Credentials $adminCreds -Url $SiteURL
$context = Get-PnPContext


#Get all the items
$AllLeaversitems = Get-PnPListItem -List $List
    
#First, check if there are any that have been modified recently
    ForEach($item in $AllLeaversitems){

    $LastModifiedDate = $Item.FieldValues.Last_x0020_Modified_x0020_Date
    $ModifiedDate = $Item.FieldValues.Modified
        If($ModifiedDate -gt $LastModifiedDate){
        #Compare the live and last entry columns
        write-host "The last modified date of this item is older the the current Modified date, something has changed! Comparing the old entries to the new entries" -ForegroundColor Yellow
            $Leaverdate = (Compare-Object -ReferenceObject $Item.FieldValues.Proposed_x0020_Leaving_x0020_Dat -DifferenceObject $Item.Last_LeavingDate)

        }
        Else{
        Write-Host "Looks like nothing has been modified!"
        }


        If($Leaverdate){
        Write-host "There has been a change to the End date" -ForegroundColor Yellow
        Set-PnPListItem -List $List -Identity $item.ID -Values @{"$Item.FieldValues.Last_LeavingDate" = "$Item.FieldValues.Proposed_x0020_Leaving_x0020_Dat"}
        Set-PnPListItem -List $List -Identity $item.ID -Values @{"$Item.FieldValues.FlowTrigger" = "Change"}

        }

        }



#######################################################################################
#                                                                                     #
#                       Changers List Processing                                      #
#                                                                                     #
#######################################################################################

#Set Variables to connect to Sharepoint - People Services (All) and Notify Internal Teams of a Leaver
$SiteURL = "https://anthesisllc.sharepoint.com/teams/People_Services_Team_All_365"
$List = "Request Change for Employee"

#Connect to Sharepoint - Groupbot? Couldn't work this bit out
Connect-PnPOnline -Credentials $adminCreds -Url $SiteURL
$context = Get-PnPContext


#Get all the items
$AllChangersitems = Get-PnPListItem -List $List

#Check for items that need processing
ForEach($item in $AllChangersitems){
    
    
    If($item.FieldValues.IsDirty = "1"){
            write-host "An item has been added, and needs processing! Let's send an email to IT and People Services" -ForegroundColor Yellow
            $subject = "Changers Update: A Request has Been Made to Change an Employee!"
            $body = "<HTML><FONT FACE=`"Calibri`">Hello People Services & IT Teams,`r`n`r`n<BR><BR>"
            $body += "You're receiving this email as someone has requested a change to be made to an Employee (these are usually changes in licensing or access requirements).`r`n`r`n<BR><BR>"
            $body += "<b>Here is a description of the change:<\b>.`r`n`r`n<BR><BR>"
            $body += "$item.FieldValues.Change_x0020_Description`r`n`r`n<BR><BR><BR><BR>"
            $body += "<b>Please make a change to the item in the 'Request Change for Employee' List on the People Services (All) Site when the change is applied. <\b>`r`n`r`n<BR><BR>"
            $body += "<b>You can see more information about this request here: https://anthesisllc.sharepoint.com/teams/People_Services_Team_All_365/Lists/Request%20Change%20for%20Employee/AllItems.aspx<\b>`r`n`r`n<BR><BR><BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>"
            $body += "The People Services Robot"

            Send-MailMessage -To "<#People services Team address for region#>" -From "thehelpfulpeopleservicesrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 
            Set-PnPListItem -List $List -Identity $item.ID -Values @{"$Item.FieldValues.IsDirty" = "0"}
            }
    

    }

}


#######################################################################################
#                                                                                     #
#                       Maternity/Paternity List Processing                           #
#                                                                                     #
#######################################################################################

