function internationalisePhoneNumber([string]$dirtyNumber, [string]$internationalCountryCode){
    if ($dirtyNumber.StartsWith("0")){$dirtyNumber = $internationalCountryCode+$dirtyNumber.Substring(1)} #If the number begins with a zero, swap it for the local country code
    $dirtyNumber
    }

$credential = get-credential -Credential kevin.maitland@anthesisgroup.com
Import-Module MSOnline
Connect-MsolService -Credential $credential

Import-Module -Name ActiveDirectory
$adUsers = Get-ADUser -SearchBase "OU=Users,OU=Sustain,DC=Sustainltd,DC=local" -Filter * -Properties ipphone, mobile, pager, manager, title, homepage, company, department, officephone
#$adUser = $adUsers | ?{$_.Name -eq "Kevin Maitland"}

##########################
# Get the SharePoint profiles
##########################
# Download and install this: http://www.microsoft.com/en-us/download/details.aspx?id=42038
Import-Module 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.UserProfiles.dll'
Import-Module 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll'
$mySiteUrl = 'https://anthesisllc-my.sharepoint.com/' # This needs to be the mySite where the userdata lives.
$adminsite = 'https://anthesisllc-admin.sharepoint.com/' # This needs to be the "admin" site.
$userProfileCollection = @()
$admin = 'kevin.maitland@anthesisgroup.com'
$password = Read-Host 'Enter Password' -AsSecureString
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($admin , $password)
#Get the Client Context and Bind the Site Collection for the MySites first to get teh full list of users
$context = New-Object Microsoft.SharePoint.Client.ClientContext($mySiteUrl)
$context.Credentials = $credentials
#Fetch the users in Site Collection
$sharepointUsers = $context.Web.SiteUsers
$context.Load($sharepointUsers)
$context.ExecuteQuery()
#Create an Object [People Manager] to retrieve profile information
$people = New-Object Microsoft.SharePoint.Client.UserProfiles.PeopleManager($context)

ForEach($user in $sharepointUsers){
    $userprofile = $people.GetPropertiesFor($user.LoginName)
    $context.Load($userprofile)
    $context.ExecuteQuery()
    $userProfileCollection += $userprofile #Enumerate each user account and save it to the userProfileCollection array
    }
#Then connect to the Admin Site Collection so that you can actually make the changes
$context = New-Object Microsoft.SharePoint.Client.ClientContext($mysite)
$context.Credentials = $credentials
$people = New-Object Microsoft.SharePoint.Client.UserProfiles.PeopleManager($context)


##########################
# Get the Mailboxes profiles
##########################
$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
Import-PSSession $ExchangeSession
$mailUsers = Get-MailUser



foreach($adUser in $adUsers){
    $newUserUpn = "$($adUser.GivenName).$($adUser.Surname)@anthesisgroup.com" 
    #Remove the exisitng corresponding MailUser object
    $mailUsers | ?{$_.Name -eq $adUser.Name} | Remove-MailUser -WhatIf
    #Create the new user object (unlicensed)
    New-MsolUser `
        -UserPrincipalName $newUserUpn `
        -FirstName $adUser.GivenName `
        -LastName $adUser.Surname `
        -DisplayName "$($adUser.GivenName) $($adUser.Surname)" `
        -Title $adUser.Title `
        -Department "SPARKE (Energy)" `
        -Office "Bristol, UK" `
        -PhoneNumber $adUser.OfficePhone `
        -MobilePhone $adUser.mobile `
        -StreetAddress "42-46 Baldwin Street" `
        -City "Bristol" `
        -PostalCode "BS1 1PN" `
        -Country "UK" `
        -UsageLocation "GB" `
        -StrongPasswordRequired $true `
        -Password "Welcome123" `
        -ForceChangePassword $true
    sleep 10 #Wait for the mailbox to be provisioned then add Mailbox access for Calendar Scheduling tools
    Add-MailboxPermission -Identity $newUserUpn -AccessRights FullAccess -user SustainMailboxAccess
    Add-MailboxPermission -Identity $newUserUpn -AccessRights SendAs -User SustainMailboxAccess -InheritanceType all
    $sharepointProfile = $sharepointUsers[0] | 
    }



foreach ($duffer in $duffers){
    $people.SetSingleValueProfileProperty($duffer, "HideFromSearch", $true)
    $context.ExecuteQuery()
    }
    Add-MailboxPermission -Identity $NewADUserName -AccessRights FullAccess -User $adUser.Manager -InheritanceType all | Out-Null
