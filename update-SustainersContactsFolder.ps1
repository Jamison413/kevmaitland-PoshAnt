param([string]$userEmailAddress,[string]$contactsFolderName)
Write-Host -ForegroundColor Yellow $userEmailAddress
#$userEmailAddress = 'thomas.milne@anthesisgroup.com'
#
#Script to create/update Sustain Contacts in a user's mailbox containing current staff info
#Originally based on Copy-OrgContactsToUserMailboxContacts.ps1, but updated for Exchange 2013
#and improved functionality
#
#Needs to connect to o365 as SustainMailboxAccess security context to enable impersonation
#
#Kev Maitland
#27/01/15
#
#Updated 01/02/17 Kev Maitland - Redesigned for Office 365


$rlbDoorCode = "#7456"
$sustainMobileNumbers= @("07824505548","07824505549","07824505550","07824505551","07887984649","07887986964","07887986470","07703821918","07703821925","07703821926","07703821930","07703821932","07824505552")

$ewsServicePath = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'
Import-Module $EWSServicePath
Import-Module -Name ActiveDirectory
$ExchVer = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016
$ewsUrl = "https://outlook.office365.com/EWS/Exchange.asmx"
$upnExtension = "anthesisgroup.com"
$upnSMA = "sustainmailboxaccess@anthesisgroup.com"
$passSMA =  ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\SustainMailboxAccess.txt) 

#$sustainFlowerLogoUrl = "http`://intranet.sustain.co.uk/PublishingImages/sustain_flower_mini.jpg"
$sustainFlowerLogoUrl = "C:\Scripts\sustain_flower_mini.jpg"
$contactsFolderName = 'Sustainers'
$contactMapping=@{
    "FirstName" = "GivenName"
    "LastName" = "Surname"
    "Company" = "CompanyName"
    "Department" = "Department"
    "Title" = "JobTitle"
    "mail" = "Email:EmailAddress1"
    "2ndEmailAddress" = "Email:EmailAddress2"
    "3rdEmailAddress" = "Email:EmailAddress3"
    "OfficePhone" = "Phone:BusinessPhone"
    "MobilePhone" = "Phone:MobilePhone"
    "BusinessPhone" = "Phone:BusinessPhone"
    "OtherTelephone" = "Phone:OtherTelephone"
    "Pager" = "Phone:Pager"
    }

function SetContactDetails([Microsoft.Exchange.WebServices.Data.Contact]$exchangeContact, [System.Object]$newContactDetails){
    # This uses the Contact Mapping above to save coding each and every field, one by one. Instead we look for a mapping and perform an action on
    # what maps across. As some methods need more "code" a fake multi-dimensional array (seperated by :'s) is used where needed.
    foreach ($key in $contactMapping.Keys){
        if ($newContactDetails.$key){ # Only do something if the key exists
            if ($contactMapping[$key] -like "*:*"){ # Will this call a more complicated mapping?
                $mappingArray = $contactMapping[$key].Split(":") # Make an array using the : to split items.
                switch ($mappingArray[0]){
                    "Email" {$exchangeContact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::($mappingArray[1])] = $newContactDetails.$key.ToString()}
                    "Phone" {$exchangeContact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::($mappingArray[1])] = $newContactDetails.$key}
                    }                
                } 
            else {$exchangeContact.($contactMapping[$key]) = $newContactDetails.$key}
            }    
        }
    $exchangeContact
    }

#Connect to Exchange using EWS
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($exchver)
$service.Credentials = New-Object System.Net.NetworkCredential($upnSMA,$passSMA)
$service.Url = $ewsUrl

# Open the main Contacts folder
$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $userEmailAddress)
$mainContactsFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts)
$mainContactsFolder.Load()

#Enumerate the Contacts Folder
$existingExchangeContacts = @()
$existingExchangeContacts += $service.FindItems($mainContactsFolder.Id,[Microsoft.Exchange.WebServices.Data.ItemView](500))  #Needs updating when more than 200 users

#Build an up-to-date contacts list
$usersList = Get-ADUser -Filter * -SearchBase "OU=Users,OU=Sustain,DC=Sustainltd,DC=local" -Properties @("SAMAccountName","DisplayName","GivenName","SurName","Title","Company","Department","mail","OfficePhone","MobilePhone")

#Add any non-standard contacts
$userCheckIn = New-Object Object
$userCheckIn | Add-Member NoteProperty DisplayName "Sustain Check-In for Lone Workers"
$userCheckIn | Add-Member NoteProperty MobilePhone "07520615285"
$userCheckIn | Add-Member NoteProperty Department "ICT"
$userCheckIn | Add-Member NoteProperty Company "Sustain Limited"
$userCheckIn | Add-Member NoteProperty mail ("checkinrobot@sustain.co.uk")
$usersList += $userCheckIn

$userSwitchboard = New-Object Object
$userSwitchboard | Add-Member NoteProperty DisplayName "Sustain Switchboard"
#$userSwitchboard | Add-Member NoteProperty FirstName "Sustain"
#$userSwitchboard | Add-Member NoteProperty LastName "Switchboard"
$userSwitchboard | Add-Member NoteProperty BusinessPhone "01174032700"
$userSwitchboard | Add-Member NoteProperty 2ndEmailAddress "$rlbDoorCode@sustain.co.uk"
$userSwitchboard | Add-Member NoteProperty 3rdEmailAddress "C2579Z_C2347Z@sustain.co.uk"
$userSwitchboard | Add-Member NoteProperty Company "Sustain Limited"
$userSwitchboard | Add-Member NoteProperty mail ("info@sustain.co.uk")
$usersList += $userSwitchboard


foreach ($contactItem in $usersList){
    $contactIsNew = $true
    $i = 0
    do { #Go through the list of Contacts until the name matches our current $contactItem, or we run out of Contacts
        if ($contactItem.DisplayName -eq $($existingExchangeContacts[$i]).DisplayName -and $contactItem.Company -eq "Sustain Limited"){ #If we find a match, update it 
            $updatedExchangeContact = SetContactDetails -exchangeContact $($existingExchangeContacts[$i]) -newContactDetails $contactItem
            #if ($updatedExchangeContact.DisplayName -eq "Sustain Check-In for Lone Workers" -or $updatedExchangeContact.DisplayName -eq "Sustain Switchboard"){$updatedExchangeContact.SetContactPicture($sustainFlowerLogoUrl)}
            $updatedExchangeContact.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
            #Write-Host -ForegroundColor DarkYellow "Updated $($updatedExchangeContact.DisplayName) with DDI $($updatedExchangeContact.PhoneNumbers.Item("BusinessPhone"))"
            $contactIsNew = $false
            }
        #if ($updatedExchangeContact.DisplayName -eq "Sustain Switchboard"){$updatedExchangeContact.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)} #Uncomment this line if you ever need to delete the Sustain Switchboard contact
        $i++
        }
    while ($true -eq $contactIsNew -and $existingExchangeContacts.Count -gt $i)
    
    if ($contactIsNew){ #If we didn't find a match in the latest Sustain list of Contacts, create a new Contact

        $newExchangeContact = New-Object Microsoft.Exchange.WebServices.Data.Contact($service)
	    if (!$contactItem.FirstName -and !$contactItem.LastName){$newExchangeContact.NickName = $contactItem.DisplayName} #Do some weird jiggery-pokery to handle incomplete Contacts
        else {$newExchangeContact.NickName = ($contactItem.FirstName + " " + $contactItem.LastName).Trim()}
        $newExchangeContact.DisplayName = $newExchangeContact.NickName
        $newExchangeContact.FileAs = $newExchangeContact.NickName
        $newExchangeContact = SetContactDetails -exchangeContact $newExchangeContact -newContactDetails $contactItem
        if ($newExchangeContact.DisplayName -eq "Sustain Check-In for Lone Workers" -or $newExchangeContact.DisplayName -eq "Sustain Switchboard"){$newExchangeContact.SetContactPicture($sustainFlowerLogoUrl)}
        Write-Host -ForegroundColor DarkYellow "New contact: $($newExchangeContact.DisplayName)"
        $newExchangeContact.Save($mainContactsFolder.Id)
        }
    }

#Then go through the existing contacts that are assigned to "Sustain Limited" and compare them to our list of current employees.
foreach ($existingContact in $existingExchangeContacts | ? {$_.CompanyName -eq "Sustain Limited"}){
    $isCurrentlyEmployed = $false
    $i = 0
    do {
        if ($existingContact.DisplayName -eq $usersList[$i].DisplayName){$isCurrentlyEmployed = $true}
        $i++
        }
    while ($isCurrentlyEmployed -eq $false -and $i -lt $usersList.Count)
    if (!$isCurrentlyEmployed){ #If the Contact is an ex-Sustainer
        Write-Host -ForegroundColor DarkYellow "$($existingContact.DisplayName) is an Ex-Sustainer"
        if ($existingContact.PhoneNumbers["BusinessPhone"]){
            if ($existingContact.PhoneNumbers["BusinessPhone"].Replace(" ","") -match "1934864"){$existingContact.PhoneNumbers["BusinessPhone"] = "."} #And they have a Sustain DDI, set it to "." (this cannot be nulled easily from here)
            if ($existingContact.PhoneNumbers["BusinessPhone"].Replace(" ","") -match "1174032"){$existingContact.PhoneNumbers["BusinessPhone"] = "."} #And they have a Sustain DDI, set it to "." (this cannot be nulled easily from here)
            }
        if ($existingContact.PhoneNumbers["MobilePhone"]){
            switch ($existingContact.PhoneNumbers["MobilePhone"].Replace(" ","")){ #And if they have a Sustain mobile, set it to "." (this cannot be nulled easily from here)
                "07824505548" {$existingContact.PhoneNumbers["MobilePhone"] = "."}
                "07824505549" {$existingContact.PhoneNumbers["MobilePhone"] = "."}
                "07824505550" {$existingContact.PhoneNumbers["MobilePhone"] = "."}
                "07824505551" {$existingContact.PhoneNumbers["MobilePhone"] = "."}
                "07887984649" {$existingContact.PhoneNumbers["MobilePhone"] = "."}
                "07887986964" {$existingContact.PhoneNumbers["MobilePhone"] = "."}
                "07887986470" {$existingContact.PhoneNumbers["MobilePhone"] = "."}
                "07703821918" {$existingContact.PhoneNumbers["MobilePhone"] = "."}
                "07703821925" {$existingContact.PhoneNumbers["MobilePhone"] = "."}
                "07703821926" {$existingContact.PhoneNumbers["MobilePhone"] = "."}
                "07703821930" {$existingContact.PhoneNumbers["MobilePhone"] = "."}
                "07703821932" {$existingContact.PhoneNumbers["MobilePhone"] = "."}
                "07824505552" {$existingContact.PhoneNumbers["MobilePhone"] = "."}
                default {}
                }
            }
        }
    if ($existingContact.IsDirty){$existingContact.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)}
    }


#Get-GlobalAddressList | Update-GlobalAddressList
#Get-OfflineAddressBook | Update-OfflineAddressBook