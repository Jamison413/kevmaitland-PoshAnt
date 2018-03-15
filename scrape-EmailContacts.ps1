# Script to scrape Contactsfrom users' mailboxes
#
# Needs to authenticate with o365 as SustainMailboxAccess to enable impersonation access via EWS (not delegate, so will work for all Sustain users only)
#
# Kev Maitland 09/03/18
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


#region functions
function get-allEwsItems($exchangeService, $folderId, $searchFilter){
    #Example $FolderId = [Microsoft.Exchange.WebServices.Data.FolderId]::new([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $mailboxEmailAddress)
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
    #Example $FolderId = [Microsoft.Exchange.WebServices.Data.FolderId]::new([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
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
$mailboxEmailAddress = "kevin.maitland@$upnExtension"
#$sendReportToAddress = "kevin.maitland@anthesisgroup.com"

$smtpServer = "anthesisgroup-com.mail.protection.outlook.com"
$logFile = "C:\ScriptLogs\process-ukCareersEmail.log"
$errorLogFile = "C:\ScriptLogs\process-ukCareersEmail_error.log"
$verboseLogging = $true
$upnSMA = "SustainMailboxAccess@anthesisgroup.com"
#$passSMA = ConvertTo-SecureString -String '' -AsPlainText -Force | ConvertFrom-SecureString
$passSMA =  ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\SustainMailboxAccess.txt) 
log-action -myMessage "Transcript saved to $($MyInvocation.MyCommand.Definition).log" -logFile $logFile
$excludedCompanies = @("*Sustain*","*Anthesis*")
$excludedDomains = @("*sustain.co.uk","*anthesisgroup.com")
$excludedPhoneNumbers = @("*117403*","*1934864*")


#Connect to Exchange using EWS
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($exchver)
$service.Credentials = New-Object System.Net.NetworkCredential($upnSMA,$passSMA)
$service.Url = $ewsUrl

$listOfFolders = $(Get-ChildItem "\\sustainltd.local\data\Personal").Name

foreach($user in $edited){
    $mailboxEmailAddress = $user+"@"+$upnExtension
    Write-Host -ForegroundColor Yellow "Processing $mailboxEmailAddress"
    $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $mailboxEmailAddress) -ErrorAction Stop

    if($contacts){rv contacts}
    $contacts = get-allEwsItems -exchangeService $service -folderId $([Microsoft.Exchange.WebServices.Data.FolderId]::new([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts, $mailboxEmailAddress)) 
    Write-Host -ForegroundColor DarkYellow "`t$($contacts.count) contacts found (total)"

    if($goodContacts){rv goodContacts}
    $contacts | %{
        if($containsData){$containsData = $false}
        if(!(matchContains -term $_.CompanyName -arrayOfStrings $excludedCompanies)){
            $newContact = New-Object psobject -Property @{"displayName"=$null;"firstName" = $null;"lastName"=$null;"email1"=$null;"email2"=$null;"email3"=$null;"businessPhone"=$null;"mobile"=$null;"company"=$null;"jobTitle"=$null;"scrapedFrom"=$user+"@"+$upnExtension}
            #displayName
            if([string]::IsNullOrWhiteSpace($_.DisplayName)){#Try to make a DisplayName from First & Last Names
                if(!([string]::IsNullOrWhiteSpace($_.GivenName) -and [string]::IsNullOrWhiteSpace($_.Surname))){$newContact.displayName = $($_.GivenName + " " + $_.Surname).Trim()}
                }
            else{$newContact.displayName = $_.DisplayName.Trim()}
            #firstName
            if([string]::IsNullOrWhiteSpace($_.GivenName)){
                if(!([string]::IsNullOrWhiteSpace($_.DisplayName))){$newContact.firstName = $($_.DisplayName.Split(" ")[0]).Trim()}
                }
            else{$newContact.firstName = $_.GivenName.Trim()}
            #lastName
            if([string]::IsNullOrWhiteSpace($_.Surname)){
                if(!([string]::IsNullOrWhiteSpace($_.DisplayName))){$newContact.lastName = $($_.DisplayName.Split(" ")[$_.DisplayName.Split(" ").Count -1]).Trim()}
                }
            else{$newContact.lastName = $_.Surname.Trim()}
            #companyName
            if([string]::IsNullOrWhiteSpace($_.CompanyName)){}
                else{$newContact.company = $_.CompanyName.Trim()}
            #jobTitle
            if([string]::IsNullOrWhiteSpace($_.JobTitle)){}
                else{$newContact.jobTitle = $_.JobTitle.Trim()}
            #emails
            if($_.EmailAddresses){
                if(!(matchContains $_.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address -arrayOfStrings $excludedDomains) -and $_.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address -match "@"){
                    $newContact.email1 = $_.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address
                    $containsData = $true
                    }
                if(!(matchContains $_.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2].Address -arrayOfStrings $excludedDomains) -and $_.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2].Address -match "@"){
                    if([string]::IsNullOrWhiteSpace($newContact.email1)){$newContact.email1 = $_.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2].Address}
                    else{$newContact.email2 = $_.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2].Address}
                    $containsData = $true
                    }
                if(!(matchContains $_.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress3].Address -arrayOfStrings $excludedDomains) -and $_.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress3].Address -match "@"){
                    if([string]::IsNullOrWhiteSpace($newContact.email1)){$newContact.email1 = $_.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress3].Address}
                    elseif([string]::IsNullOrWhiteSpace($newContact.email2)){$newContact.email2 = $_.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress3].Address}
                    else{$newContact.email3 = $_.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress3].Address}
                    $containsData = $true
                    }
                }
            #phones
            if($_.PhoneNumbers){
                if(!([string]::IsNullOrWhiteSpace($_.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone]))){
                    if(!(matchContains -term $($_.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone].Replace(" ","")) -arrayOfStrings $excludedPhoneNumbers)){
                        $newContact.businessPhone = $_.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone].Replace(" ","")
                        $containsData = $true
                        }
                    }
                if(!([string]::IsNullOrWhiteSpace($_.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone]))){
                    if(!(matchContains -term $($_.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone].Replace(" ","")) -arrayOfStrings $excludedPhoneNumbers)){
                        $newContact.mobile = $_.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone].Replace(" ","")
                        $containsData = $true
                        }
                    }
                }
            }
        if($containsData -and ($newContact.company -or $newContact.email1 -or $newContact.businessPhone -or $newContact.mobile) -and (!([string]::IsNullOrWhiteSpace($newContact.displayName)))){[array]$goodContacts += $newContact}
        }
    Write-Host -ForegroundColor DarkYellow "`t$($goodContacts.Count) contacts processed"
    $outputfile = "\\sustainltd.local\data\Personal\$user\Contacts2.csv"
    #$goodContacts | % {$_ | Export-Csv -Path "\\sustainltd.local\data\Personal\$user\Contacts3.csv" -NoTypeInformation -Append}
    #if(!(Test-Path $outputfile)){
    $testMe = Get-Item $outputfile
    #if($testMe.LastWriteTime -lt "2018-03-13 10:15" -and $(get-acl $testMe).Owner -eq "SUSTAINLTD\kevin.maitland"){
        $goodContacts | select displayName,firstName,lastName,jobTitle,company,businessPhone,mobile,email1,email2,email3,scrapedFrom |  Export-Csv -Path $outputfile -NoTypeInformation -Encoding UTF8
        #$goodContacts[0] | select displayName,mobile |  Export-Csv -Path "\\sustainltd.local\data\Personal\$user\Contacts2.csv" -NoTypeInformation -Append
        #}
    #else{Write-Host -ForegroundColor Magenta "$user has edited Contacts"; [array]$manualUsers}
    }

Stop-Transcript
   

foreach($user in $listOfFolders){
    $mailboxEmailAddress = $user+"@"+$upnExtension
    $outputfile = "\\sustainltd.local\data\Personal\$user\Contacts.csv"
    $testMe = Get-Item $outputfile
    if($testMe){
        if($(get-acl $testMe).Owner -eq "SUSTAINLTD\kevin.maitland"){
            if($testMe.LastWriteTime -gt "2018-03-13 11:50" ){
                [array]$rebuildMe += $user
                }
            else{[array]$edited += $user}
            }
        else{[array]$manaulExport += $user}
        }
    else{[array]$missing += $user}
    }

