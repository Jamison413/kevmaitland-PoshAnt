#Script to check the \ICT\CheckIn Public Folder for e-mails (from TextMarketer), then
#match the phone number to an employee, create an appointment in their Outlook calendar
#using the message body and timestamp, then e-mail TextMarketer back (who will SMS the 
#original sender)
#
# Needs to authenticate with o365 as SustainMailboxAccess to enable impersonation via EWS
#
# Kev Maitland 15/1/15
#
# Edited Kev Maitland 30/04/15 - added function FindSmsInEmail to handle a change in formatting from TextMagic
# Edited Kev Maitland 01/02/17 - revised for Office 365 and changed Get-User to Get-ADUser to avoid hitting o365 Exchange
Start-Transcript "$($MyInvocation.MyCommand.Definition).log" #-Append

$EWSServicePath = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'
Import-Module $EWSServicePath
Import-Module -Name ActiveDirectory

#Set some variables
$ExchVer = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016
$ewsUrl = "https://outlook.office365.com/EWS/Exchange.asmx"
$upnExtension = "anthesisgroup.com"
$smtpServer = "anthesisgroup-com.mail.protection.outlook.com"
$smsServiceDomain = "@textmagic.com"
$checkInPFPath   = "\1.Public Folders\1.Admin\IT\Check-In"
$processedPFPath = "\1.Public Folders\1.Admin\IT\Check-In\Processed"
$failedPFPath    = "\1.Public Folders\1.Admin\IT\Check-In\Failed"
$internationalDialCode = "00"
$localCountryCode = "44"
$sharePointServer = "SP01"
$sharePointPSSesssionConfigName = "SharePointFarmScripts"
$logFile = "C:\ScriptLogs\checkInTool.log"
$errorLogFile = "C:\ScriptLogs\checkInTool_error.log"
$upnSMA = "sustainmailboxaccess@anthesisgroup.com"
#$passSMA = ConvertTo-SecureString -String '' -AsPlainText -Force | ConvertFrom-SecureString
$passSMA =  ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\SustainMailboxAccess.txt) 
$upnLdap = "ldapqueries@sustainltd.local"
$passLdap = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\ldapqueries.txt) 
$credsLdap = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $upnLdap, $passLdap
$upnGMLink = "goldminelink@sustainltd.local"
$passGMLink = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\GMLink.txt) 
$credsGMLink = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $upnGMLink, $passGMLink

#Connect to Exchange using EWS
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($exchver)
$service.Credentials = New-Object System.Net.NetworkCredential($upnSMA,$passSMA)
$service.Url = $ewsUrl

#region functions
function FolderIdFromPath{  
    param ($FolderPath = "$( throw 'Folder Path is a mandatory Parameter' )")  
    process{  
        ## Find and Bind to Folder based on Path    
        #Define the path to search should be seperated with \    
        #Bind to the MSGFolder Root    
        $folderId = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot)     
        $tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderId)    
        #Split the Search path into an array    
        $fldArray = $FolderPath.Split("\")  
         #Loop through the Split Array and do a Search for each level of folder  
        for ($lint = 1; $lint -lt $fldArray.Length; $lint++) {  
            #Perform search based on the displayname of each folder level  
            $fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1)  
            $SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$fldArray[$lint])  
            $findFolderResults = $service.FindFolders($tfTargetFolder.Id,$SfSearchFilter,$fvFolderView)  
            if ($findFolderResults.TotalCount -gt 0){  
                foreach($folder in $findFolderResults.Folders){  
                    $tfTargetFolder = $folder                 
                }  
            }  
            else{  
                "Error Folder Not Found"   
                $tfTargetFolder = $null   
                break   
            }      
        }   
        if($tfTargetFolder -ne $null){ 
            return $tfTargetFolder.Id.UniqueId.ToString() 
        } 
    } 
} 
function GetStandardisedUKMobileNumber([string]$dirtyNumber){
    $dirtyNumber = $dirtyNumber.Replace(" ","") #Remove any spaces in the number
    if ($dirtyNumber -match "(0)"){$dirtyNumber = $dirtyNumber.Replace("(0)","")} #Drop any optional zero-code and hope it still works
    if (!($dirtyNumber.StartsWith($internationalDialCode))){#If it's not in standard international format...
        if ($dirtyNumber.StartsWith("0")){$dirtyNumber = $internationalDialCode+$localCountryCode+$dirtyNumber.Substring(1)} #If the number begins with a zero, swap it for the local country code
        if ($dirtyNumber.StartsWith("+")){$dirtyNumber = $internationalDialCode+$dirtyNumber.Substring(1)} #If the number begins with a +, swap it for 00
        if ($dirtyNumber.StartsWith($localCountryCode)){$dirtyNumber = $internationalDialCode+$dirtyNumber} #If it's a valid local country code, but missing the 00, add it now
        }
    if ($dirtyNumber.Length -eq 14){$dirtyNumber}
    else {"Problem: $dirtyNumber"}
    #else {"There was a sponge in the patient with the mobile number"}
    }
function get-managerSamAccountNameFromCn($cNString, $adCreds){
    (Get-ADUser $cNString -Credential $adCreds).SamAccountName
    }
function FindSmsInEmail([string]$emailBody){
    $i = -1
    do { #TextMagic-specific way to find the useful part of the e-mail.
        $i++
        $tempBody = $(($emailBody -split "`n")[$i] -split "<")[0] #Take the first section from every line of the HTML (this is where the "text" lands when we split by "<")
        }
    While ((($emailBody -split "`n")[$i].Substring(0,1) -match "\W") `
        -and ($i+1 -lt ($emailBody -split "`n").count)) #Until the first letter of the line contains a letter or number, or we run out of lines
    $tempBody
    }
function LogMessage([string]$logMessage){
    Add-Content -Value "$(Get-Date -Format G): $logMessage" -Path $logFile
    }
function LogError([string]$errorMessage){
    Add-Content -Value "$(Get-Date -Format G): $errorMessage" -Path $logFile
    Add-Content -Value "$(Get-Date -Format G): $errorMessage" -Path $errorLogFile
    Send-MailMessage -To "itnn@sustain.co.uk" -From scriptrobot@sustain.co.uk -SmtpServer $smtpServer -Subject "Error in $($MyInvocation.ScriptName) on $env:COMPUTERNAME" -Body $errorMessage
    }        
#endregion

$folderId = new-object Microsoft.Exchange.WebServices.Data.FolderId(FolderIdFromPath -FolderPath $checkInPFPath)
$checkInPF = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderId)    
$folderId = new-object Microsoft.Exchange.WebServices.Data.FolderId(FolderIdFromPath -FolderPath $processedPFPath)
$processedPF = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderId)    
$folderId = new-object Microsoft.Exchange.WebServices.Data.FolderId(FolderIdFromPath -FolderPath $failedPFPath)
$failedPF = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderId)    
$itemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(10)  
$propertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  

#Get an array of objects from SharePoint with the users' mobile phone numbers:
    try {$userData = Invoke-Command -ComputerName $sharePointServer -ConfigurationName $sharePointPSSesssionConfigName -ScriptBlock {param($actionToPerform, $findProperty, $findPropertyCurrentValue) \\SP01\Scripts\MySites_Maintenance_UserProfileData.ps1 @PSBoundParameters} -ArgumentList "GetAll", "CellPhone", "dummyValue" -ErrorAction Stop -Credential $credsGMLink}
    catch {
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        LogError -errorMessage "ERROR: Failed to read MobilePhone info from MySites `n Item: $FailedItem`n Message: $ErrorMessage"
        }

    #Build an up-to-date contacts list from AD to get Company mobile phone numbers too
    $usersList = Get-ADUser -Filter * -SearchBase "OU=Users,OU=Sustain,DC=Sustainltd,DC=local" -Properties @("SAMAccountName","DisplayName","GivenName","SurName","Title","Company","Department","mail","OfficePhone","MobilePhone","Manager") -Credential $credsLdap
    #Generalise the data from AD so that we can use it with the data from SharePoint
    foreach ($user in $usersList){
        $generalisedUser = New-Object Object
        if($user.MobilePhone){$generalisedUser | Add-Member NoteProperty CellPhone ($user.MobilePhone).Replace(" ","")}
            else{$generalisedUser | Add-Member NoteProperty CellPhone ""}
        $generalisedUser | Add-Member NoteProperty AccountName "SUSTAINLTD\$($user.SAMAccountName)"
        $generalisedUser | Add-Member NoteProperty WhatIsMyName $user.GivenName
        $generalisedUser | Add-Member NoteProperty Manager "SUSTAINLTD\$(get-managerSamAccountNameFromCn -cNString $user.Manager -adCreds $credsLdap)"
        $userData += $generalisedUser
        }


$foundItems = $null  
do{  
    $foundItems = $service.FindItems($checkInPF.Id,$itemView)  
    [Void]$service.LoadPropertiesForItems($foundItems,$propertySet)
    foreach($email in $foundItems.Items){
        LogMessage -logMessage "INFO: E-mail received from $($email.Sender)"
        $texterNumber = GetStandardisedUKMobileNumber $(($email.Subject -split " ")[($email.Subject -split " ").Count -1]).Replace("(", "").Replace(")", "")

        $texterName = $null
        foreach ($user in $userData){
            if ($(GetStandardisedUKMobileNumber $user.CellPhone) -eq $texterNumber){
                LogMessage -logMessage "INFO: SMS identified as being from $(($user.AccountName -split "\\")[1])"
                #CreateAppointment -service $service -user $user.AccountName -email $email
                try {$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $("$(($user.AccountName -split "\\")[1])@$upnExtension")) -ErrorAction Stop}
                catch {
                    $ErrorMessage = $_.Exception.Message
                    $FailedItem = $_.Exception.ItemName
                    LogError -errorMessage "ERROR: Failed to impersonate $($user.AccountName) in EWS to create appointment`n Item: $FailedItem`n Message: $ErrorMessage"
                    }
                $texterName = $user.WhatIsMyName
                $appointment = New-Object Microsoft.Exchange.WebServices.Data.Appointment -ArgumentList $service
                $appointment.Start = $email.DateTimeReceived
                $appointment.End = $email.DateTimeReceived.AddMinutes(15)
                $appointment.Subject = "Check-in via SMS"
                $appointment.Body = FindSmsInEmail -emailBody $email.Body.ToString()
                $appointment.IsAllDayEvent = $false
                $appointment.ReminderDueBy = $email.DateTimeReceived
                $appointment.IsReminderSet = $true
                $appointment.Categories.Add("Check-In")
                $appointment.Categories.Add("Important")
                try {$appointment.Save([Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToAllAndSaveCopy)
                } 
                catch {
                    $ErrorMessage = $_.Exception.Message
                    $FailedItem = $_.Exception.ItemName
                    LogError -errorMessage "ERROR: Failed to save newly created appointment for $($user.AccountName) in EWS `nItem: $FailedItem`n Message: $ErrorMessage"
                    break
                    }
                try {
                    $email.Move($processedPF.Id)
                    #Confirm receipt and update via email-to-SMS service at TextMarketer if the message is moved (to prevent responding to the same message repeatedly)
                    $mycreds = New-Object System.Management.Automation.PSCredential ($upnSMA, $passSMA)
                    Send-MailMessage -To "$($texterNumber.Substring(2))$smsServiceDomain" -From "checkinrobot@sustain.co.uk" -SmtpServer $smtpServer -Subject "Thanks $texterName - I've updated your Outlook calendar with the contents of your text message. Love, The Sustain Check-In Robot" -Body " " -Credential $mycreds
                    Send-MailMessage -To ($user.Manager.Replace("SUSTAINLTD\","").Replace("sustainltd\","")+"@$upnExtension") -From "checkinrobot@anthesisgroup.com" -SmtpServer $smtpServer -Subject "$texterName has checked in via the Lone Worker Check-In tool." -Body "$(FindSmsInEmail -emailBody $email.Body.ToString()) `r`n$(get-date $email.DateTimeReceived -Format "dd/MM/yyyy HH:mm:ss")`r`n`r`nLove, `r`n`r`nThe Sustain Check-In Robot (on behalf of $texterName)" -Credential $mycreds
                    Send-MailMessage -To ("wai.cheung@$upnExtension") -From "checkinrobot@anthesisgroup.com" -SmtpServer $smtpServer -Subject "$texterName has checked in via the Lone Worker Check-In tool." -Body "$(FindSmsInEmail -emailBody $email.Body.ToString()) `r`n$(get-date $email.DateTimeReceived -Format "dd/MM/yyyy HH:mm:ss")`r`n`r`nLove, `r`n`r`nThe Sustain Check-In Robot (on behalf of $texterName)" -Credential $mycreds
                    }
                catch {
                    $ErrorMessage = $_.Exception.Message
                    $FailedItem = $_.Exception.ItemName
                    LogError -errorMessage "ERROR: Failed to move processed Check-In for $($user.AccountName) via EWS `n Item: $FailedItem`n Message: $ErrorMessage"
                    break
                    }
                break #break here to avoid multiple matches where the user has listed their company phone in SharePoint
                }
            }
    if ($texterName -eq $null) {
        LogError -errorMessage "WARNING: SMS received from an unknown phone number: $texterNumber"
        Send-MailMessage -To "officemanagementteam@sustain.co.uk" -From checkinrobot@sustain.co.uk -SmtpServer $smtpServer -Subject "Check-in text received from unknown number: $texterNumber" -Body "Hello Office Management Team,`n`nI've received a Check-In text, but I don't know who it's from (so I can't update their calendar) :(`n`nIt's from: $texterNumber`nAnd it reads: $((($email.Body.ToString() -split "`n")[28] -split "<")[0])`n`nLove,`n`nThe Lone Worker Check-In Tool Robot" -Credential $mycreds
        Send-MailMessage -To "$($texterNumber.Substring(2))$smsServiceDomain" -From "checkinrobot@sustain.co.uk" -SmtpServer $smtpServer -Subject "I'm sorry, but I couldn't find your mobile number in any MySustain profile. Please call 0117 4032 700 and check in verbally. Love, The Sustain Check-In Robot" -Body " " -Credential $mycreds
        try {$email.Move($failedPF.Id)}
        catch {
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            LogError -errorMessage "ERROR: Failed to move failed Check-In for $($user.AccountName) via EWS `n Item: $FailedItem`n Message: $ErrorMessage"
            }
        }
    }  
    $itemView.Offset += $foundItems.Items.Count  
}while($foundItems.MoreAvailable -eq $true) 




Stop-Transcript