# Script to check the UKCareers Mailbox and create SharePoint list items for pre-configured Job Roles
#
#
# Needs to authenticate with o365 as KimbleBot to enable delegate access via EWS (not impersonation)
#
# Kev Maitland 17/1/18
#
$logFileLocation = "C:\ScriptLogs\"
$transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_Transcript_$(Get-Date -Format "yyMMdd").log"
if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)){
    $fullLogPathAndName = $logFileLocation+"sync-kimbleProjectsToSpo_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = $logFileLocation+"sync-kimbleProjectsToSpo_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
else{
    $fullLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
Start-Transcript $transcriptLogName

$EWSServicePath = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'
Import-Module $EWSServicePath
Import-Module _PS_Library_GeneralFunctionality
Import-Module _CSOM_Library-SPO.psm1
Import-Module _REST_Library-SPO.psm1


#region functions
function get-allEwsItems($exchangeService, $folderId, $searchFilter){
    #Example $FolderId = [Microsoft.Exchange.WebServices.Data.FolderId]::new([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $ukCareersEmailAddress)
    $bind = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService, $folderId)
    #$ukCareersItems = $service.FindItems($bind.Id,$searchFilter,[Microsoft.Exchange.WebServices.Data.ItemView]::new(100))

    $itemsOffset = [Microsoft.Exchange.WebServices.Data.ItemView]::new(100)
    do {
        $foundItems = $exchangeService.FindItems($bind.Id,$null,$itemsOffset)
        $itemsOffset.Offset = $foundItems.NextPageOffset
        $allItems += $foundItems.Items
        Write-Host -ForegroundColor DarkYellow "`t`t$($allItems.count)/$($foundItems.TotalCount) retrieved"
        }
    while ($foundItems.MoreAvailable -eq $true) 
    $allItems
    }
function get-allEwsFolders($exchangeService, $folderId){
    #Example FolderId = [Microsoft.Exchange.WebServices.Data.FolderId]::new([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox
    $folderOffset = [Microsoft.Exchange.WebServices.Data.FolderView]::new(100)
    do {
        $foundFolders = $exchangeService.FindFolders($folderId, $folderOffset)
        $folderOffset.Offset = $foundFolders.NextPageOffset
        $allFolders += $foundFolders.Folders
        }
    while ($foundFolders.MoreAvailable -eq $true) 
    $allFolders
    }

function FolderIdFromPath{  
    param ($FolderPath = "$( throw 'Folder Path is a mandatory Parameter' )"
        , $exchangeService = "$( throw 'exchangeService is a mandatory Parameter' )"
        , $smtpAddress)  
    process{  
        ## Find and Bind to Folder based on Path    
        #Define the path to search should be seperated with \    
        #Bind to the MSGFolder Root    
        $folderId = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)     
        $tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService,$folderId,$smtpAddress)    
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

function LogMessage([string]$logMessage){
    Add-Content -Value "$(Get-Date -Format G): $logMessage" -Path $logFile
    }
function LogError([string]$errorMessage){
    Add-Content -Value "$(Get-Date -Format G): $errorMessage" -Path $logFile
    Add-Content -Value "$(Get-Date -Format G): $errorMessage" -Path $errorLogFile
    Send-MailMessage -To "itnn@sustain.co.uk" -From scriptrobot@sustain.co.uk -SmtpServer $smtpServer -Subject "Error in $($MyInvocation.ScriptName) on $env:COMPUTERNAME" -Body $errorMessage
    }        
#endregion

#Set some variables
$ExchVer = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016
$ewsUrl = "https://outlook.office365.com/EWS/Exchange.asmx"
$upnExtension = "anthesisgroup.com"
$ukCareersEmailAddress = "ukcareers@$upnExtension"
$sendReportToAddress = @("amanda.cox@anthesisgroup.com","helen.tyrrell@anthesisgroup.com","wai.cheung@anthesisgroup.com","lorna.kelly@anthesisgroup.com")
#$sendReportToAddress = "kevin.maitland@anthesisgroup.com"

$smtpServer = "anthesisgroup-com.mail.protection.outlook.com"
$closedJobMailFolderName = "z_closed"
$onHoldJobMailFolderName = "z_on-hold"
$duffMailFolderName = "_UnableToAutomate"
$logFile = "C:\ScriptLogs\process-ukCareersEmail.log"
$errorLogFile = "C:\ScriptLogs\process-ukCareersEmail_error.log"
$verboseLogging = $true
$upnSMA = "kimblebot@anthesisgroup.com"
#$passSMA = ConvertTo-SecureString -String '' -AsPlainText -Force | ConvertFrom-SecureString
$passSMA =  ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\KimbleBot.txt) 
log-action -myMessage "Transcript saved to $($MyInvocation.MyCommand.Definition).log" -logFile $logFile

#Connect to Exchange using EWS
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($exchver)
$service.Credentials = New-Object System.Net.NetworkCredential($upnSMA,$passSMA)
$service.Url = $ewsUrl
#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $ukCareersEmailAddress) -ErrorAction Stop

$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $upnSMA, $passSMA
$restCreds = new-spoCred -Credential -username $adminCreds.UserName -securePassword $adminCreds.Password
$csomCreds = new-csomCredentials -username $adminCreds.UserName -password $adminCreds.Password

$webUrl = "https://anthesisllc.sharepoint.com"
$recruitmentSiteCollection = "/teams/communities"
$recruitmentSite = "/Recruitment"
$ukJobsListName = "UK Roles"
$ukCandidatsListName = "UK Role Candidates"
$taxonomyListName = "TaxonomyHiddenList"
$taxonomyData = get-itemsInList -serverUrl $webUrl -sitePath $recruitmentSiteCollection -listName $taxonomyListName -suppressProgress $false -restCreds $restCreds -logFile $logFile -verboseLogging $verboseLogging

$ukJobs = get-itemsInList -serverUrl $webUrl -sitePath $($recruitmentSiteCollection+$recruitmentSite) -listName $ukJobsListName -restCreds $restCreds -logFile $logFile -verboseLogging $verboseLogging -oDataQuery "?`$filter=MailFolderArchived eq 0" #Get all Jobs where we haven't archived the Mail Folder (as they're done and dusted)
$ukJobsObjects =@()
$ukJobs | % {[array]$ukJobsObjects += convert-listItemToCustomObject -spoListItem $_ -spoTaxonomyData $taxonomyData} #Convert the List Items to PS Objects as some of the attributes are difficult to work with (e.g. MetaData or People/Group)
$ukJobsObjects | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name DisplayName -Value $_.UniqueJobID} #Add the UniqueJobId property again as DisplayName so we can use compare-object efficiently later
$ukCandidateList = get-list -serverUrl $webUrl -sitePath $($recruitmentSiteCollection+$recruitmentSite) -listName $ukCandidatsListName -restCreds $restCreds -verboseLogging $verboseLogging -logFile $logFile

#Reconcile the current Job Roles with the E-mail folders
$inboxFolders = get-allEwsFolders -exchangeService $service -folderId $([Microsoft.Exchange.WebServices.Data.FolderId]::new([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $ukCareersEmailAddress)) 
if($verboseLogging){Write-Host -ForegroundColor Yellow "Reconciling e-mail folders: $($inboxFolders.Count) Inbox folders and $($ukJobsObjects.Count) Job Roles found"}
$onHoldFolderId = $inboxFolders | ? {$_.DisplayName -eq $onHoldJobMailFolderName} | % {$_.Id}
$closedFolderId = $inboxFolders | ? {$_.DisplayName -eq $closedJobMailFolderName} | % {$_.Id}
$dufferFolderId = $inboxFolders | ? {$_.DisplayName -eq $duffMailFolderName} | % {$_.Id}
$newFoldersToReconcile = Compare-Object $inboxFolders -DifferenceObject $ukJobsObjects -Property "DisplayName" -PassThru -ErrorAction Continue
#If the Role is open, but there's no folder...
if($verboseLogging){if(($newFoldersToReconcile | ?{$_.SideIndicator -eq "=>" -and $_.Recruitment_Status -match "Open"} | % {})-eq $null){Write-Host -ForegroundColor DarkYellow "No new e-mail folders need creating in Inbox"}}
$newFoldersToReconcile | ?{$_.SideIndicator -eq "=>" -and $_.Recruitment_Status -match "Open"} | % { 
    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "Creating e-mail folder $($_.UniqueJobId) in Inbox"}
    $newJobMailFolder = [Microsoft.Exchange.WebServices.Data.Folder]::new($service)
    $newJobMailFolder.DisplayName = $_.UniqueJobID
    $newJobMailFolder.Save([Microsoft.Exchange.WebServices.Data.FolderId]::new([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $ukCareersEmailAddress))
    }
#If the Role is On-hold, but there is a folder...
$ukJobsObjectsOnHold = $ukJobsObjects | ?{$_.Recruitment_Status -match "hold"}
if($ukJobsObjectsOnHold){ #Compare-Object will throw a wobbly if we send a $null -DifferenceObject
    $onHoldFoldersToReconcile = Compare-Object $inboxFolders -DifferenceObject $ukJobsObjectsOnHold -Property "DisplayName" -PassThru -ExcludeDifferent -IncludeEqual -ErrorAction Continue
    if($onHoldFoldersToReconcile){#Process any folders
        $onHoldFoldersToReconcile | %{
            if($verboseLogging){Write-Host -ForegroundColor DarkYellow "Moving folder $($_.DisplayName) to $onHoldJobMailFolderName"}
            $folderToMoveBind = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $_.Id)
            $folderToMoveBind.Move($onHoldFolderId)
            }
        }
    else{if($verboseLogging){Write-Host -ForegroundColor DarkYellow "No e-mail folders need moving to $onHoldJobMailFolderName"}}
    }
else{if($verboseLogging){Write-Host -ForegroundColor DarkYellow "No Job Roles are On-Hold"}}
#If the Role is Closed, but there is a folder...
$ukJobsObjectsClosed = $ukJobsObjects | ?{$_.Recruitment_Status -match "Closed"}
if($ukJobsObjectsClosed){ #Compare-Object will throw a wobbly if we send a $null -DifferenceObject
    $closedFoldersToReconcile = Compare-Object $inboxFolders -DifferenceObject $ukJobsObjectsClosed -Property "DisplayName" -PassThru -ExcludeDifferent -IncludeEqual -ErrorAction Continue
    if($closedFoldersToReconcile){#Process any folders
        $closedFoldersToReconcile | %{
            #Move the folder and mark the JobRole ListItem as MailFolderArchived
            try{
                #First move the folder
                if($verboseLogging){Write-Host -ForegroundColor DarkYellow "Moving folder $($_.DisplayName) to $closedJobMailFolderName"}
                $folderToMoveBind = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $_.Id)
                $folderToMoveBind.Move($closedFolderId)
                #Second, update the ListItem
                $folderToMoveObject = $_
                try{
                    $jobRecordToUpdate = $ukJobsObjectsClosed | ?{$_.UniqueJobID -eq $folderToMoveObject.DisplayName}
                    $recruitmentDigest = check-digestExpiry -serverUrl $webUrl -sitePath $($recruitmentSiteCollection+$recruitmentSite) -restCreds $restCreds -logFile $logFile -digest $recruitmentDigest
                    update-itemInList -serverUrl $webUrl -sitePath $($recruitmentSiteCollection+$recruitmentSite) -listNameOrGuid $ukJobsListName -predeterminedItemType $jobRecordToUpdate.__metadata.type -itemId $jobRecordToUpdate.Id -hashTableOfItemData @{"MailFolderArchived"=$true} -restCreds $restCreds -digest $recruitmentDigest -logFile $logFile -verboseLogging $true
                    }
                catch{[array]$spongeInThePatient += "Could not update JobRole $($jobRecordToUpdate.DisplayName) to set MailFolderArchived = `$true"}
                }
            catch{[array]$spongeInThePatient += "Could not Move mail folder $($_.DisplayName) to $closedJobMailFolderName"}
            }
        }
    else{if($verboseLogging){Write-Host -ForegroundColor DarkYellow "No e-mail folders need moving to $closedJobMailFolderName"}}
    }
else{if($verboseLogging){Write-Host -ForegroundColor DarkYellow "No Job Roles are Closed (that we haven't already archived)"}}

#Get this again now that we might have created some new folders
if($newFoldersToReconcile | ?{$_.SideIndicator -eq "=>" -and $_.Recruitment_Status -match "Open"}){
    $inboxFolders = get-allEwsFolders -exchangeService $service -folderId $([Microsoft.Exchange.WebServices.Data.FolderId]::new([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $ukCareersEmailAddress)) 
    }

#Get any e-mails that haven't been processed (i.e. are still in the root Inbox)
$ukCareersItems = get-allEwsItems -exchangeService $service -folderId $([Microsoft.Exchange.WebServices.Data.FolderId]::new([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $ukCareersEmailAddress)) -searchFilter $null
if($verboseLogging){Write-Host -ForegroundColor Yellow "Found $($ukCareersItems.Count) e-mails that need processing"}
$ukCareersItems | %{
    #Read the Subject to see if it contains a jobID
    $email = $_
    if($verboseLogging){Write-Host -ForegroundColor Yellow "Found e-mail from $($email.Sender.Address) with Subject $($email.Subject)"}
    $jobCodes = @()
    $email.Subject.Split(" ") | % {
        if($ukJobsObjects.UniqueJobID -icontains $_){[array]$jobCodes += $_} #We need -icontains, not -imatch to avoid matching WS1 to WS12
        }
    if($jobCodes.Count -eq 1){#If it contains exactly 1 job code, create a List Item
        if($verboseLogging){Write-Host -ForegroundColor DarkYellow "Exactly 1 JobCode found [$($jobCodes[0])]"}
        #$dummyCandidate = get-itemsInList -serverUrl $webUrl -sitePath $($recruitmentSiteCollection+$recruitmentSite) -listName "UK Role Candidates" -restCreds $restCreds -logFile $logFile -oDataQuery "&`$filter=Id eq 18" -verboseLogging $true #Quickly find the field names in SPO
        $recruitmentDigest = check-digestExpiry -serverUrl $webUrl -sitePath $($recruitmentSiteCollection+$recruitmentSite) -restCreds $restCreds -logFile $logFile -digest $recruitmentDigest
        $candidateHash=@{}
        $candidateHash.Add("Application_x0020_received",$email.DateTimeReceived)
        $candidateHash.Add("E_x002d_mail_x0020_address",$email.Sender.Address)
        $candidateHash.Add("Title", $email.Sender.Name)
        #Try to work out the first/last name
        if($email.Sender.Name.Split(" ").Count -eq 2){ #Check that there are exactly two words separated by a space
            if($email.Sender.Name.Split(" ")[0] -notmatch ","){ #If the first word doesn't contain a comma, assume the format "FirstName LastName"
                $candidateHash.Add("First_x0020_Name", $email.Sender.Name.Split(" ")[0])
                $candidateHash.Add("Last_x0020_Name", $email.Sender.Name.Split(" ")[1])
                }
            else{# If the first word /does/ contain a comma, assume the format "LastName, FirstName"
                $candidateHash.Add("First_x0020_Name", $email.Sender.Name.Split(" ")[1])
                $candidateHash.Add("Last_x0020_Name", $email.Sender.Name.Split(" ")[0])
                }
            }
        #Get the required information to build a metadata value for the JobID
        $jobTerm = $taxonomyData | ?{$_.Term -contains $jobCodes[0]}
        <#This is what the real object looks like, but we only need the string values to build the JSON REST command
        $jobIDMetaDataValue = [Microsoft.SharePoint.Client.Taxonomy.TaxonomyFieldValue]::new() 
        $jobIDMetaDataValue.Label = $jobTerm.Term
        $jobIDMetaDataValue.TermGuid = $jobTerm.IdForTerm
        $jobIDMetaDataValue.WssId = -1#>
        $jobIDMetaDataHash = @{"Label"=$jobTerm.Term;"TermGuid"=$jobTerm.IdForTerm;"WssId"=-1} #No idea what WssId refers to, but -1 works fine.
        $formattedJobIdData = format-itemData $jobIDMetaDataHash
        if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`$formattedJobIdData = $formattedJobIdData"}
        $candidateHash.Add("RoleID","{'__metadata':{'type':'SP.Taxonomy.TaxonomyFieldValue'},$formattedJobIdData}")
        if($verboseLogging){Write-Host -ForegroundColor DarkYellow "Trying to create Candidate List Item"}
        $newCandidateRecord = new-itemInList -serverUrl $webUrl -sitePath $($recruitmentSiteCollection+$recruitmentSite) -listName $ukCandidatsListName -predeterminedItemType $ukCandidateList.ListItemEntityTypeFullName -hashTableOfItemData $candidateHash -restCreds $restCreds -digest $recruitmentDigest -logFile $logFile -verboseLogging $verboseLogging
        if(!$newCandidateRecord){
            if($verboseLogging){Write-Host -ForegroundColor DarkYellow "New Candidate ListItem did not create for $($email.Sender.Address) :("}
            [array]$spongeInThePatient += "Something went wrong when creating the Candidate Record for $($email.Sender.Address) in SharePoint"
            }
        else{[array]$allNewCandidateRecords += $newCandidateRecord}

        #If the e-mail has attachments, add them to the newly-created List Item
        if($email.HasAttachments){
            if($verboseLogging){Write-Host -ForegroundColor DarkYellow "E-mail has attachments"}
            $emailWithAttachments = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service, $email.Id, [Microsoft.Exchange.WebServices.Data.PropertySet]::new([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.ItemSchema]::Attachments))
            $emailWithAttachments.Attachments | % {
                $attachment = $_
                if($attachment.GetType().Name -eq "FileAttachment" -and ($attachment.Name -match ".doc" -or $attachment.Name -match ".pdf")){
                    $tempFilePathAndName = $env:TEMP+"\"+$attachment.Name.Replace("'","''")
                    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "Attempting to save attachment as $tempFilePathAndName"}
                    $attachment.Load($tempFilePathAndName)
                    $webResponseCode = add-attachmentToListItem -serverUrl $webUrl -sitePath $($recruitmentSiteCollection+$recruitmentSite) -listItem $newCandidateRecord -filePathAndName $tempFilePathAndName -restCreds $restCreds -digest $recruitmentDigest -logFile $logFile -verboseLogging $true
                    if ($webResponseCode -eq "OK"){Write-Host "High-five!"}
                    else{[array]$spongeInThePatient += "There were valid attachments from $($email.Sender.Address), but something went wrong and I couldn't attach them $($attachment.Name) in SharePoint :("}
                    #$webResponseCode.Dispose()
                    Remove-Item -Path $tempFilePathAndName -Force #Delete the local copy
                    }
                else{[array]$spongeInThePatient += "There were attachments that weren't Word or PDF files ($($attachment.Name)) from $($email.Sender.Address)"}
                }
            }

        #If it's worked, move the e-mail to the appropriate subfolder
        if($newCandidateRecord){
            $jobFolder = $null
            $jobFolder = $inboxFolders | ?{$_.DisplayName -eq $jobCodes[0]}
            if($jobFolder){
                $email.Move($jobFolder.Id)
                }
            else{
                [array]$spongeInThePatient += "Job $($jobCodes[0]) is not Open! I've created the Candidate record anyway, but I'm moving the e-mail from $($email.Sender.Address) to $duffMailFolderName"
                $email.Move($dufferFolderId)
                }
            }
        }
    #If it doesn't have exactly one Job Code, move it to the "_UnableToAutomate" subfolder & leave it for a human to process
    else{
        if($verboseLogging){Write-Host -ForegroundColor DarkYellow "Could not work out the JobCode for e-mail from $($email.Sender.Address)"}
        $email.Move($dufferFolderId)
        $duffCount++
        }
    }
        
#Write a nice report if we've done anything
if($ukCareersItems){
    $reportBody = "<HTML><BODY><P>Hello Recruitment Human,</P>
    <P>I've been working away on your behalf and I've:</P><UL>"
    $reportBody += "<LI>Looked at $($ukCareersItems.Count) e-mails</LI>"
    if($duffCount){$reportBody += "<LI>But I couldn't process $duffCount automatically, so you'll need to take a look at them (in the $duffMailFolderName folder)</LI>"}
    if($allNewCandidateRecords){
        $reportBody += "<LI>And I've created new <A HREF=`"$webUrl$recruitmentSiteCollection$recruitmentSite/Lists/Role%20Candidates/`">Candidate records</A> for:<UL>"
        $allNewCandidateRecords | %{
            $newRecordObject = convert-listItemToCustomObject -spoListItem $_ -spoTaxonomyData $taxonomyData
            $reportBody += "<LI>$($newRecordObject.RoleID)`t$($newRecordObject.First_Name) $($newRecordObject.Last_Name)`t$($newRecordObject.E_x002d_mail_address)</LI>"
            }
        $reportBody += "</UL>"
        }
    if($newFoldersToReconcile | ?{$_.SideIndicator -eq "=>" -and $_.Recruitment_Status -match "Open"}){
        $reportBody += "<LI>And I've created new E-mail folders these new Job Roles:<UL>"
        $newFoldersToReconcile | ?{$_.SideIndicator -eq "=>" -and $_.Recruitment_Status -match "Open"} | %{
            $newRecordObject = convert-listItemToCustomObject -spoListItem $_ -spoTaxonomyData $taxonomyData
            $reportBody += "<LI>$($newRecordObject.UniqueJobID) - $($newRecordObject.Title)</LI>"
            }
        $reportBody += "</UL>"
        }
    if($onHoldFoldersToReconcile){
        $reportBody += "<LI>And I've moved these E-mail folders to the $onHoldJobMailFolderName folder:<UL>"
        $onHoldFoldersToReconcile | % {
            $reportBody += "<LI>$($_.DisplayName)</LI>"
            }
        $reportBody += "</UL>"
        }
    if($closedFoldersToReconcile){
        $reportBody += "<LI>And I've moved these E-mail folders to the $closedJobMailFolderName folder:<UL>"
        $closedFoldersToReconcile | % {
            $reportBody += "<LI>$($_.DisplayName)</LI>"
            }
        $reportBody += "</UL>"
        }
    if($spongeInThePatient){
        $reportBody += "I also encountered these problems:<UL>"
        $spongeInThePatient | %{
            $reportBody += "<LI>$($_)</LI>"
            }
        $reportBody += "</UL>"
        }
    $reportBody += "</UL>"
    $reportBody += "<P></P><P>Love,</P><P>The Recruitment Robot</P></BODY></HTML>"
    Send-MailMessage -To $sendReportToAddress -Subject "Recruitment automation update" -BodyAsHtml $reportBody -From "therecruitmentrobot@anthesisgroup.com" -SmtpServer $smtpServer
    }
Stop-Transcript


<#
$toRestore = @("AAMkAGU3MzI0NjllLTQxMzAtNGJlMy1iYmFkLWI2NDliNTQ2M2UwYQAuAAAAAADlFHvuHN6MTIgoCwE/HEs2AQBBwLL8RoSKQpoh0zPoc4a5AAKm3/ZkAAA=",
"AAMkAGU3MzI0NjllLTQxMzAtNGJlMy1iYmFkLWI2NDliNTQ2M2UwYQAuAAAAAADlFHvuHN6MTIgoCwE/HEs2AQBBwLL8RoSKQpoh0zPoc4a5AAL/KEcuAAA=",
"AAMkAGU3MzI0NjllLTQxMzAtNGJlMy1iYmFkLWI2NDliNTQ2M2UwYQAuAAAAAADlFHvuHN6MTIgoCwE/HEs2AQBBwLL8RoSKQpoh0zPoc4a5AAI8jQi3AAA=",
"AAMkAGU3MzI0NjllLTQxMzAtNGJlMy1iYmFkLWI2NDliNTQ2M2UwYQAuAAAAAADlFHvuHN6MTIgoCwE/HEs2AQBBwLL8RoSKQpoh0zPoc4a5AAMWM5tYAAA=",
"AAMkAGU3MzI0NjllLTQxMzAtNGJlMy1iYmFkLWI2NDliNTQ2M2UwYQAuAAAAAADlFHvuHN6MTIgoCwE/HEs2AQBBwLL8RoSKQpoh0zPoc4a5AAMVyrQ7AAA=",
"AAMkAGU3MzI0NjllLTQxMzAtNGJlMy1iYmFkLWI2NDliNTQ2M2UwYQAuAAAAAADlFHvuHN6MTIgoCwE/HEs2AQBBwLL8RoSKQpoh0zPoc4a5AALIrgRBAAA=",
"AAMkAGU3MzI0NjllLTQxMzAtNGJlMy1iYmFkLWI2NDliNTQ2M2UwYQAuAAAAAADlFHvuHN6MTIgoCwE/HEs2AQBBwLL8RoSKQpoh0zPoc4a5AAKu4J6QAAA=",
"AAMkAGU3MzI0NjllLTQxMzAtNGJlMy1iYmFkLWI2NDliNTQ2M2UwYQAuAAAAAADlFHvuHN6MTIgoCwE/HEs2AQBBwLL8RoSKQpoh0zPoc4a5AAKra0PeAAA=",
"AAMkAGU3MzI0NjllLTQxMzAtNGJlMy1iYmFkLWI2NDliNTQ2M2UwYQAuAAAAAADlFHvuHN6MTIgoCwE/HEs2AQBBwLL8RoSKQpoh0zPoc4a5AAKra0PfAAA=",
"AAMkAGU3MzI0NjllLTQxMzAtNGJlMy1iYmFkLWI2NDliNTQ2M2UwYQAuAAAAAADlFHvuHN6MTIgoCwE/HEs2AQBBwLL8RoSKQpoh0zPoc4a5AAI8jQiSAAA=",
"AAMkAGU3MzI0NjllLTQxMzAtNGJlMy1iYmFkLWI2NDliNTQ2M2UwYQAuAAAAAADlFHvuHN6MTIgoCwE/HEs2AQBBwLL8RoSKQpoh0zPoc4a5AALeeEBeAAA=",
"AAMkAGU3MzI0NjllLTQxMzAtNGJlMy1iYmFkLWI2NDliNTQ2M2UwYQAuAAAAAADlFHvuHN6MTIgoCwE/HEs2AQBBwLL8RoSKQpoh0zPoc4a5AAMTucacAAA=",
"AAMkAGU3MzI0NjllLTQxMzAtNGJlMy1iYmFkLWI2NDliNTQ2M2UwYQAuAAAAAADlFHvuHN6MTIgoCwE/HEs2AQBBwLL8RoSKQpoh0zPoc4a5AAMVyrQ8AAA=",
"AAMkAGU3MzI0NjllLTQxMzAtNGJlMy1iYmFkLWI2NDliNTQ2M2UwYQAuAAAAAADlFHvuHN6MTIgoCwE/HEs2AQBBwLL8RoSKQpoh0zPoc4a5AALaVOQLAAA=")

$allClosedFolders = get-allEwsFolders -exchangeService $service -folderId $onHoldFolderId
foreach ($folder in $toRestore){
    $restoreId = $null
    $restoreId = $allClosedFolders | ?{$_.Id.UniqueId -contains $folder} | % {$_.Id}
    if($restoreId){
        $folderToMoveBind = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $restoreId)
        $folderToMoveBind.Move($inboxFolderId)    
        }
    else {write-host -ForegroundColor DarkBlue "$folder failed"}
    }

$toRestore = @("AAMkAGU3MzI0NjllLTQxMzAtNGJlMy1iYmFkLWI2NDliNTQ2M2UwYQAuAAAAAADlFHvuHN6MTIgoCwE/HEs2AQBBwLL8RoSKQpoh0zPoc4a5AALaVOQMAAA=", "AAMkAGU3MzI0NjllLTQxMzAtNGJlMy1iYmFkLWI2NDliNTQ2M2UwYQAuAAAAAADlFHvuHN6MTIgoCwE/HEs2AQBBwLL8RoSKQpoh0zPoc4a5AAMWM5tZAAA=")
#>