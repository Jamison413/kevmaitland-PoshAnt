# Script to scrape Contact History from users' mailboxes
#
## Needs to authenticate with o365 as KimbleBot to enable delegate access via EWS (not impersonation)
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
function scrape-recursively($exchangeService, $mailboxEmailAddress,$ewsFolder,$mailIsOutbound,$contactHistoryHash){
    Write-Host -ForegroundColor Yellow "Scraping $($ewsFolder.TotalCount) messages in $($ewsFolder.DisplayName) from $mailboxEmailAddress"
    if($mailIsOutbound){Write-Host -ForegroundColor DarkYellow "`tThese are Outbound messages, so they are much slower to process :("}
    get-mostRecentMailFromEachContact -exchangeService $exchangeService -mailboxEmailAddress $mailboxEmailAddress -folderId $ewsFolder.Id -mailIsOutbound $mailIsOutbound -contactHistoryHash $contactHistoryHash

    $subFolders = get-allEwsFolders -exchangeService $service -folderId $ewsFolder.Id
    Write-Host -ForegroundColor Yellow "`t$($subFolders.Count) subfolders found!"
    $subFolders | %{scrape-recursively -exchangeService $exchangeService -mailboxEmailAddress $mailboxEmailAddress -ewsFolder $_ -mailIsOutbound $mailIsOutbound -contactHistoryHash $contactHistoryHash}
    }
function get-mostRecentMailFromEachContact($exchangeService,$mailboxEmailAddress,$folderId,$mailIsOutbound,$contactHistoryHash){
    $itemsInFolder = get-allEwsItems -exchangeService $exchangeService -folderId $folderId
    $itemsInFolder | %{
        $unboundEmail = $_
        if($mailIsOutbound){
            $boundEmail = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($exchangeService, $unboundEmail.Id, [Microsoft.Exchange.WebServices.Data.PropertySet]::new([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties))
            $boundEmail.ToRecipients | %{
                if(!(matchContains -term $_.Address -arrayOfStrings $excludedDomains)){do-contactHistoryStuff -myAddress $mailboxEmailAddress -theirAddress $_.Address -mailIsOutbound $true -dateStamp $boundEmail.DateTimeSent -subject $boundEmail.Subject.Replace(",","") -contactHistoryHash $contactHistoryHash -theirName $((guess-nameFromString $_.Name).Replace(",",""))}
                }
            $boundEmail.CcRecipients | %{
                if(!(matchContains -term $_.Address -arrayOfStrings $excludedDomains)){do-contactHistoryStuff -myAddress $mailboxEmailAddress -theirAddress $_.Address -mailIsOutbound $true -dateStamp $boundEmail.DateTimeSent -subject $boundEmail.Subject.Replace(",","") -contactHistoryHash $contactHistoryHash -theirName $((guess-nameFromString $_.Name).Replace(",",""))}
                }
            $boundEmail.BccRecipients | %{
                if(!(matchContains -term $_.Address -arrayOfStrings $excludedDomains)){do-contactHistoryStuff -myAddress $mailboxEmailAddress -theirAddress $_.Address -mailIsOutbound $true -dateStamp $boundEmail.DateTimeSent -subject $boundEmail.Subject.Replace(",","") -contactHistoryHash $contactHistoryHash -theirName $((guess-nameFromString $_.Name).Replace(",",""))}
                }
            }
        else{
            if(!([string]::IsNullOrEmpty($unboundEmail.From.Address))){
                if(!(matchContains -term $unboundEmail.From.Address -arrayOfStrings $excludedDomains) -and $unboundEmail.From.Address -match "@"){
                    do-contactHistoryStuff -myAddress $mailboxEmailAddress -theirAddress $unboundEmail.From.Address -mailIsOutbound $false -dateStamp $unboundEmail.DateTimeReceived -subject $unboundEmail.Subject.Replace(",","") -contactHistoryHash $contactHistoryHash -theirName $((guess-nameFromString $unboundEmail.From.Name).Replace(",",""))
                    }
                }
            }
        }
    }
function do-contactHistoryStuff($myAddress,$theirAddress,$theirName,$mailIsOutbound,$dateStamp,$subject,$contactHistoryHash){
    if($contactHistoryHash.Keys -notcontains $theirAddress){
        #add it
        $detailsHash = [ordered]@{"mailbox"=$myAddress;"from"=$theirAddress;"to"=$myAddress;"directionOutbound"=$mailIsOutbound;"inboundDate"=$null;"outboundDate"=$null;"theirDomain"=$($theirAddress.Split("@")[1]);"inboundMessageCount"=0;"outboundMessageCount"=0;"guessedName"=$theirName;"lastSubject"=$subject}
        if($mailIsOutbound){ #reverse the to/from
            $detailsHash["to"] = $theirAddress
            $detailsHash["from"] = $myAddress
            $detailsHash["outboundDate"] = $dateStamp
            $detailsHash["outboundMessageCount"] = 1
            }
        else{
            $detailsHash["inboundDate"] = $dateStamp
            $detailsHash["inboundMessageCount"] = 1
            }
        $details = New-Object psobject -Property $detailsHash
        $contactHistoryHash.Add($theirAddress,$details)
        }
    elseif($mailIsOutbound){
        if($contactHistoryHash[$theirAddress].inboundDate -lt $dateStamp){
            #overwrite it
            $contactHistoryHash[$theirAddress].inboundDate = $dateStamp
            $contactHistoryHash[$theirAddress].directionOutbound = $mailIsOutbound
            $contactHistoryHash[$theirAddress].lastSubject = $subject
            $contactHistoryHash[$theirAddress].inboundMessageCount = $contactHistoryHash[$theirAddress].inboundMessageCount + 1
            }
        else{
            #Just increment inboundMessageCount
            $contactHistoryHash[$theirAddress].inboundMessageCount = $contactHistoryHash[$theirAddress].inboundMessageCount + 1
            }
        }
    else{
        if($contactHistoryHash[$theirAddress].outboundDate -lt $dateStamp){
            #overwrite it
            $contactHistoryHash[$theirAddress].outboundDate = $dateStamp
            $contactHistoryHash[$theirAddress].directionOutbound = $mailIsOutbound
            $contactHistoryHash[$theirAddress].lastSubject = $subject
            $contactHistoryHash[$theirAddress].outboundMessageCount = $contactHistoryHash[$theirAddress].outboundMessageCount + 1
            }
        else{
            #Just increment outboundMessageCount
            $contactHistoryHash[$theirAddress].outboundMessageCount = $contactHistoryHash[$theirAddress].outboundMessageCount + 1
            }
        }
    }
#endregion

#Set some variables
$ExchVer = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016
$ewsUrl = "https://outlook.office365.com/EWS/Exchange.asmx"
$upnExtension = "anthesisgroup.com"
$smtpServer = "anthesisgroup-com.mail.protection.outlook.com"
$excludedDomains = @("*sustain.co.uk","*anthesisgroup.com","*AnthesisLLC.onmicrosoft.com","*gmail.com","*hotmail.com","*hotmail.co.uk","*yahoo.com","*yahoo.co.uk","*twitter.com","*twitter.co.uk","*linkedin.com","*outlook.com","*bestfootforward.com","*bestfootforward.co.ukget-m","*calebgroup.net")
$logFile = "C:\ScriptLogs\scrape-EmailContactHistory.log"
$errorLogFile = "C:\ScriptLogs\scrape-EmailContactHistory_error.log"
#$verboseLogging = $true
$upnSMA = "SustainMailboxAccess@anthesisgroup.com"
#$passSMA = ConvertTo-SecureString -String '' -AsPlainText -Force | ConvertFrom-SecureString
$passSMA =  ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\SustainMailboxAccess.txt) 
log-action -myMessage "Transcript saved to $($MyInvocation.MyCommand.Definition).log" -logFile $logFile

#Connect to Exchange using EWS
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($exchver)
$service.Credentials = New-Object System.Net.NetworkCredential($upnSMA,$passSMA)
$service.Url = $ewsUrl

$listOfMailboxesToScrape = @("tim.clare","jono.adams","brad.blundell","ian.forrester","craig.simmons")
foreach($user in $listOfMailboxesToScrape){
    $mailboxEmailAddress = "$user@$upnExtension"
    $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $mailboxEmailAddress) -ErrorAction Stop

    $contactHistoryHash = [ordered]@{}
    $inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,[Microsoft.Exchange.WebServices.Data.PropertySet]::new([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties))
    $sentItems = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems,[Microsoft.Exchange.WebServices.Data.PropertySet]::new([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties))

    scrape-recursively -exchangeService $service -mailboxEmailAddress $mailboxEmailAddress -ewsFolder $inbox -mailIsOutbound $false -contactHistoryHash $contactHistoryHash
    scrape-recursively -exchangeService $service -mailboxEmailAddress $mailboxEmailAddress -ewsFolder $sentItems -mailIsOutbound $true -contactHistoryHash $contactHistoryHash

    $contactHistoryHash.Keys | %{
        $contactHistoryHash[$_] |  Export-Csv -Path "$env:USERPROFILE\Desktop\Scrape_$($mailboxEmailAddress)_Initial.csv" -NoTypeInformation -Append
        }
    }
Stop-Transcript
