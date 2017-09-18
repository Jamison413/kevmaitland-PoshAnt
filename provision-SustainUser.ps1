Import-Module -Name ActiveDirectory
Import-Module .\_PS_Library_MSOL.psm1
Import-Module .\_REST_Library-SPO.psm1

$userSAM = "Ali.Midhani"
$userFirstName = "Ali"
$userSurname = "Midhani"
$userManagerSAM = "Duncan.Faulkes"
$userCommunity = "SPARK"
$userDepartment = "Sustain"
$userJobTitle = "Associate"
$plaintextPassword = "Welcome123"
$licenses = @("E1")
$timeZone = "GMT Standard Time"
$countryLocale = "2057"


$logFile = "C:\Scripts\Logs\provision-User.log"
$errorLogFile = "C:\Scripts\Logs\provision-User_Errors.log"
$smtpServer = "anthesisgroup-com.mail.protection.outlook.com"
$msolAdmin = "kevin.maitland@anthesisgroup.com"

$creds = set-MsolCredentials
connect-ToMsol -credential $creds
connect-ToExo -credential $creds
Set-SPORestCredentials -Credential $creds


$sharePointServerUrl = "https://anthesisllc.sharepoint.com"
$hrSite = "/teams/hr"
$taxonomyListName = "TaxonomyHiddenList"
$taxononmyData = get-itemsInList -serverUrl $sharePointServerUrl -sitePath $hrSite -listName $taxonomyListName -suppressProgress $false 

$newUserListName = "New User Requests"
#$oDataUnprocessedUsers = '$filter=Current_x0020_Status eq ''1 - Waiting for IT Team to set up accounts'''
#$oDataUnprocessedUsers = '$select=*&$filter=Current_x0020_Status eq ''1 - Waiting for IT Team to set up accounts''&$expand=Line_x0020_Manager/Id'
$oDataUnprocessedUsers = '$select=*,Line_x0020_Manager/Name,Line_x0020_Manager/Title,Prinicpal_x0020_Community_x0020_/Name,Prinicpal_x0020_Community_x0020_/Title&$filter=Current_x0020_Status eq ''1 - Waiting for IT Team to set up accounts''&$expand=Line_x0020_Manager/Id,Prinicpal_x0020_Community_x0020_/Id'
$unprocessedStarters = get-itemsInList -serverUrl $sharePointServerUrl -sitePath $hrSite -listName $newUserListName -suppressProgress $false -oDataQuery $oDataUnprocessedUsers 
#$unprocessedStarters | %{
foreach($unprocessedStarter in $unprocessedStarters){
    $startingUser = New-Object -TypeName PSObject
    #These are read directly from the List:
    $startingUser | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $unprocessedStarter.Title
    if($unprocessedStarter.Employee_x0020_Legal_x0020_Name.__deferred -eq $null){
        $startingUser | Add-Member -MemberType NoteProperty -Name "LegalName" -Value $unprocessedStarter.Employee_x0020_Legal_x0020_Name
        }
        else {$startingUser | Add-Member -MemberType NoteProperty -Name "LegalName" -Value $unprocessedStarter.Title}
    $startingUser | Add-Member -MemberType NoteProperty -Name "JobTitle" -Value $unprocessedStarter.Job_x0020_title
    #These are taxonomy fields and the Ids need to cross-referenced with the TaxonomyHiddenList to get the labels
    $startingUser | Add-Member -MemberType NoteProperty -Name "Region" -Value $($taxononmyData | ?{$_.IdForTerm -eq $unprocessedStarter.Region.TermGuid} | %{$_.Term})
    $startingUser | Add-Member -MemberType NoteProperty -Name "NearestOffice" -Value $($taxononmyData | ?{$_.IdForTerm -eq $unprocessedStarters[0].Nearest_x0020_Office.TermGuid} | %{$_.Term})
    $startingUser | Add-Member -MemberType NoteProperty -Name "Community" -Value $($taxononmyData | ?{$_.IdForTerm -eq $unprocessedStarters[0].Community.TermGuid} | %{$_.Term})
    $startingUser | Add-Member -MemberType NoteProperty -Name "Company" -Value $($taxononmyData | ?{$_.IdForTerm -eq $unprocessedStarters[0].Finance_x0020_Cost_x0020_Attribu.TermGuid} | %{$_.Term})
    #These are People/Group fields and need expanding
    if($unprocessedStarter.Line_x0020_Manager.__deferred -eq $null){
        $startingUser | Add-Member -MemberType NoteProperty -Name "LineManager" -Value $unprocessedStarter.Line_x0020_Manager.Name.Replace("i:0#.f|membership|","")
        }
        else{$startingUser | Add-Member -MemberType NoteProperty -Name "LineManager" -Value ""}
    if($unprocessedStarter.Prinicpal_x0020_Community_x0020_.__deferred -eq $null){
        $startingUser | Add-Member -MemberType NoteProperty -Name "CommunityManager" -Value $unprocessedStarter.Prinicpal_x0020_Community_x0020_.Name.Replace("i:0#.f|membership|","")
        }
        else{$startingUser | Add-Member -MemberType NoteProperty -Name "CommunityManager" -Value ""}

    $unprocessedStartersFormatted += $startingUser
    }

$selectedStartersrs = $unprocessedStartersFormatted | Out-GridView -PassThru



#region functions
function create-ADUser([string]$userSAM, [string]$userFirstName, [string]$userSurname, [string]$userManagerSAM, [string]$userDepartment, [string]$userJobTitle, $plaintextPassword){
    New-ADUser `
        -AccountPassword (ConvertTo-SecureString $plaintextPassword -AsPlainText -force) `
        -CannotChangePassword $False `
        -ChangePasswordAtLogon $False `
        -Company "Sustain Limited" `
        -Department $userDepartment `
        -DisplayName "$userFirstName $userSurname"`
        -Enabled $true `
        -Fax "SustainLtd" `
        -GivenName $userFirstName `
        -HomePage "www.sustain.co.uk" `
        -Manager $(Get-ADUser $userManagerSAM) `
        -Name "$userFirstName $userSurname"`
        -OfficePhone "0117 403 2XXX" `
        -Path 'OU=Users,OU=Sustain,DC=Sustainltd,DC=local' `
        -SAMAccountName $userSAM `
        -Surname $userSurname `
        -Title $userJobTitle `
        -UserPrincipalName "$userSAM@sustain.co.uk" `
        -EmailAddress "$userSAM@anthesisgroup.com"
        -OtherAttributes @{'ipPhone'="XXX";'pager'="0117 403 2700"} | Out-Null 
    }
function create-msolUser($userSAM){
    #create the Mailbox rather than the MSOLUser, which will effectively create an unlicensed E1 user
    New-Mailbox -Name $userSAM.Replace("."," ") -Password (ConvertTo-SecureString -AsPlainText $plaintextPassword -Force) -MicrosoftOnlineServicesID $userSAM@anthesisgroup.com
    }
function license-msolUser($userSAM, $licenseType){
    switch ($licenseType){
        "E1" {$licenseToAssign = Get-MsolAccountSku | ?{$_.AccountSkuId -eq "AnthesisLLC:STANDARDPACK"}}
        "E3" {$licenseToAssign = Get-MsolAccountSku | ?{$_.AccountSkuId -eq "AnthesisLLC:ENTERPRISEPACK"}}
        "VISIO" {$licenseToAssign = Get-MsolAccountSku | ?{$_.AccountSkuId -eq "AnthesisLLC:VISIOCLIENT"}}
        "PROJECT" {$licenseToAssign = Get-MsolAccountSku | ?{$_.AccountSkuId -eq "AnthesisLLC:PROJECTPROFESSIONAL"}}
        }
    Set-MsolUserLicense -UserPrincipalName $userSAM@anthesisgroup.com -AddLicenses $licenseToAssign.AccountSkuId
    }
function update-MsolUser([string]$userSAM, [string]$userFirstName, [string]$userSurname, [string]$userDisplayName, [string]$userManagerSAM, [string]$userDepartment, [string]$userJobTitle, [string]$userPhoneExtension){
    #$msolUser = New-MsolUser `
    Set-MsolUser -UserPrincipalName "$userSAM@anthesisgroup.com" `
        -FirstName $userFirstName `
        -LastName $userSurname `
        -DisplayName $userDisplayName `
        -Title $userJobTitle `
        -Department "Sustain" `
        -Office "Bristol, GBR" `
        -PhoneNumber "+44 117 430 2$userPhoneExtension" `
        -StreetAddress "42-46 Baldwin Street" `
        -City "Bristol, GBR" `
        -PostalCode "BS1 1PN" `
        -Country "United Kingdom" `
        -UsageLocation "GB" `
        -StrongPasswordRequired $true 
        #-Password "Welcome123" `
        #-ForceChangePassword $true
    }
function update-msolMailbox($userSAM,$userFirstName,$userSurname,$userDisplayName){
    Get-Mailbox $userSAM@anthesisgroup.com | Set-Mailbox  -CustomAttribute1 "Sustain" -Alias $userSAM -DisplayName $userDisplayName -Name "$userFirstName $userSurname" -Office "Bristol, UK" -EmailAddresses @{add="$userSAM@sustain.co.uk"}
    Get-Mailbox $userSAM@anthesisgroup.com | Set-CASMailbox -ActiveSyncMailboxPolicy "Sustain"
    Set-User -Identity $userSAM@anthesisgroup.com -Company "Sustain"
    Set-MailboxRegionalConfiguration -Identity $userSAM@anthesisgroup.com  -TimeZone $timeZone
    }
function update-msolSharePointProfileFromAnotherProfile($sourceSpProfile,$destSpProfile,$destContext,$destPeopleManager){
    if($sourceSpProfile.UserProfileProperties["AboutMe"] -ne $null){$destPeopleManager.SetSingleValueProfileProperty($destSpProfile.AccountName, "AboutMe", $sourceSpProfile.UserProfileProperties["AboutMe"])}
    if($sourceSpProfile.UserProfileProperties["SPS-Birthday"] -ne $null){$destPeopleManager.SetSingleValueProfileProperty($destSpProfile.AccountName, "SPS-Birthday", $sourceSpProfile.UserProfileProperties["SPS-Birthday"])}
    if($sourceSpProfile.UserProfileProperties["Bio"] -ne $null){$destPeopleManager.SetSingleValueProfileProperty($destSpProfile.AccountName, "Bio", $sourceSpProfile.UserProfileProperties["Bio"])}
    
    if($sourceSpProfile.UserProfileProperties["SPS-PastProjects"] -ne $null){$destPeopleManager.SetMultiValuedProfileProperty($destSpProfile.AccountName, "SPS-PastProjects", $sourceSpProfile.UserProfileProperties["SPS-PastProjects"].Split("|"))} 
    if($sourceSpProfile.UserProfileProperties["SPS-Skills"] -ne $null){$destPeopleManager.SetMultiValuedProfileProperty($destSpProfile.AccountName, "SPS-Skills", $sourceSpProfile.UserProfileProperties["SPS-Skills"].Split("|"))} 
    if($sourceSpProfile.UserProfileProperties["SPS-School"] -ne $null){$destPeopleManager.SetMultiValuedProfileProperty($destSpProfile.AccountName, "SPS-School", $sourceSpProfile.UserProfileProperties["SPS-School"].Split("|"))} 
    if($sourceSpProfile.UserProfileProperties["SPS-Interests"] -ne $null){$destPeopleManager.SetMultiValuedProfileProperty($destSpProfile.AccountName, "SPS-Interests", $sourceSpProfile.UserProfileProperties["SPS-Interests"].Split("|"))} 
    if($sourceSpProfile.UserProfileProperties["Qualifications"] -ne $null){$destPeopleManager.SetMultiValuedProfileProperty($destSpProfile.AccountName, "Qualifications", $sourceSpProfile.UserProfileProperties["Qualifications"].Split("|"))} 
    $destContext.ExecuteQuery()
    }
function update-SharePointInitialConfig([string]$userSAM, $anthesisAdminSite, $csomCreds, $timeZone, $countryLocale){
    $adminContext = new-csomContext -fullSitePath $anthesisAdminSite -sharePointCredentials $csomCreds
    $spoUsers = New-Object Microsoft.SharePoint.Client.UserProfiles.PeopleManager($adminContext)
    $spoUsers.SetSingleValueProfileProperty("$userSAM@anthesisgroup.com", "SPS-RegionalSettings-Initialized", $true)
    $spoUsers.SetSingleValueProfileProperty("$userSAM@anthesisgroup.com", "SPS-RegionalSettings-FollowWeb", $false)
    #Getting the TimeZoneID is a massive PITA:
    if($timeZones -eq $null){$timeZones = get-timeZones}
    $tz = $timeZones | ?{$_.PSChildName -eq $timeZone} #Look that up in the registry list
    if($spoTimeZones -eq $null){$spoTimeZones = get-spoTimeZoneHashTable}
    $tzID = $spoTimeZones[$tz.Display.replace("+00:00","")] #Then match a different property of the registry object to the SPO object
    if($tzID.Length -gt 0){$spoUsers.SetSingleValueProfileProperty("$userSAM@anthesisgroup.com", "SPS-TimeZone", $tzID)}
    if($countryToLocaleHashTable[$antUser.Country].length -gt 0){$spoUsers.SetSingleValueProfileProperty("$userSAM@anthesisgroup.com", "SPS-Locale", $countryLocale)}
    $spoUsers.SetSingleValueProfileProperty("$userSAM@anthesisgroup.com", "SPS-MUILanguages", "en-GB")
    $spoUsers.SetSingleValueProfileProperty("$userSAM@anthesisgroup.com", "SPS-CalendarType", 1)
    $spoUsers.SetSingleValueProfileProperty("$userSAM@anthesisgroup.com", "SPS-AltCalendarType", 1)
    $adminContext.ExecuteQuery()
    }
function create-personalFolder([string]$userSAM){
    $dirRoot = "X:\Personal"
    
    #Create the user's Personal Folder and give them Modify rights
    $personalFolder = New-Item -Path "$dirRoot\$userSAM" -ItemType Directory
    $acl = Get-Acl $personalFolder
    $perm = "Modify"
    $permInherit = "ContainerInherit, ObjectInherit" #This folder, files & subfolders - see http://powershell.nicoh.me/powershell-1/files-and-folders/set-folders-acl-owner-and-ntfs-rights 
    $permProp = "None" #This folder, files & subfolders - see http://powershell.nicoh.me/powershell-1/files-and-folders/set-folders-acl-owner-and-ntfs-rights 
    $ace = New-Object System.Security.AccessControl.FileSystemAccessRule($userSAM, $perm, $permInherit, $permProp, "Allow")
    $acl.AddAccessRule($ace)
    Set-Acl -Path $personalFolder -AclObject $acl

    #Create the user's Secure folder, break the inheritance permissions
    $secureFolder = New-Item -Path "$dirRoot\$userSAM\Secure" -ItemType Directory
    $acl = Get-Acl $secureFolder
    $acl.SetAccessRuleProtection($true,$true)  #Note that SetAccessRuleProtection takes two boolean arguments; the first turns inheritance on ($False) or off ($True) and the second determines whether the previously inherited permissions are retained ($True) or removed ($False)
    Set-Acl -Path $secureFolder -AclObject $acl
    
    #Now remove all permissions that are not the user's or backup-related
    foreach ($ace in $acl.Access){
        if (!($ace.IdentityReference -eq "SUSTAINLTD\Backup Process Account - do not block permissions" -or $ace.IdentityReference -eq "SUSTAINLTD\$userSAM")){
            icacls $secureFolder /remove `"$($ace.IdentityReference)`" | Out-Null
            }
        }
    }
function set-mailboxPermissions([string]$userSAM, [string]$userManagerSAM){
    Add-MailboxPermission -Identity $userSAM -AccessRights FullAccess -user SustainMailboxAccess@anthesisgroup.com | Out-Null
    #Add-MailboxPermission -Identity $userSAM -AccessRights SendAs -User SustainMailboxAccess@anthesisgroup.com -InheritanceType all | Out-Null
    Add-MailboxPermission -Identity $userSAM -AccessRights FullAccess -User $userManagerSAM -InheritanceType all | Out-Null
    Add-MailboxFolderPermission "$($userSAM):\Calendar" -User "View all Sustain calendars" -AccessRights "Reviewer" | Out-Null
    Add-MailboxFolderPermission "$($userSAM):\Calendar" -User "Edit all Sustain calendars" -AccessRights "PublishingEditor" | Out-Null
    }
function log-Message([string]$logMessage, $colour){
    Write-Host -Object $logMessage -ForegroundColor $colour 
    Add-Content -Value "$(Get-Date -Format G): $logMessage" -Path $logFile
    }
function log-Error([string]$errorMessage){
    Write-Host -f Red $errorMessage
    Add-Content -Value "$(Get-Date -Format G): $errorMessage" -Path $logFile
    Add-Content -Value "$(Get-Date -Format G): $errorMessage" -Path $errorLogFile
    Send-MailMessage -To "itnn@sustain.co.uk" -From scriptrobot@sustain.co.uk -SmtpServer $smtpServer -Subject "Error in $MyInvocation.ScriptName on $env:COMPUTERNAME" -Body $errorMessage
    }
#endregion

function provision-user($userSAM, $userFirstName, $userSurname, $userManagerSAM, $userDepartment, $userJobTitle, $plaintextPassword){
    try{
        log-Message "Creating AD account for $userSAM" -colour "Yellow"
        create-ADUser -userSAM $userSAM -userFirstName $userFirstName -userSurname $userSurname -userManagerSAM $userManagerSAM -userDepartment $userDepartment -userJobTitle $userJobTitle -plaintextPassword $plaintextPassword
        log-Message "Account created" -colour "DarkYellow"
        }
    catch{
        log-Error "Failed to create AD account"
        log-Error $Error
        }
    try{
        log-Message "Creating personal folder for $userSAM" -colour "Yellow"
        create-personalFolder -userSAM $userSAM
        log-Message "Folder created" -colour "DarkYellow"
        }
    catch{
        log-Error "Failed to create personal folder"
        log-Error $Error
        }
    try{
        log-Message "Creating MSOL account for $userSAM" -colour "Yellow"
        create-msolUser -userSAM $userSAM
        log-Message "Account created" -colour "DarkYellow"
        }
    catch{
        log-Error "Failed to create MSOL account"
        log-Error $Error
        }
    Start-Sleep -Seconds 5 #Give EXO a chance to catch up
    try{
        log-Message "Setting mailbox permissions for $userSAM" -colour "Yellow"
        set-mailboxPermissions -userSAM $userSAM -userManagerSAM $userManagerSAM
        log-Message "Mailbox permissions set" -colour "DarkYellow"
        }
    catch{
        log-Error "Failed to set mailbox permissions"
        log-Error $Error
        }
    try{
        log-Message "Updating mailbox for $userSAM" -colour "Yellow"
        update-msolMailbox -userSAM $userSAM -userFirstName $userFirstName -userSurname $userSurname -userDisplayName "$userFirstName $userSurname"
        log-Message "Mailbox updated" -colour "DarkYellow"
        }
    catch{
        log-Error "Failed to update mailbox"
        log-Error $Error
        }
    }
function update-msolUserFromAd($userSAM){
    $adU = Get-ADUser $userSAM -Properties DisplayName,Title,Department,Office,ipPhone,Manager
    $userManagerSAM = (Get-ADUser $adu.Manager).SamAccountName
    try{
        log-Message "Updating MSOL account for $userSAM" -colour "Yellow"
        update-MsolUser -userSAM $userSAM -userFirstName $adU.GivenName -userSurname $adU.Surname -userManagerSAM $userManagerSAM -userDepartment "SPARK" -userJobTitle $adU.Title -userDisplayName $adU.DisplayName -userPhoneExtension $adU.ipPhone
        log-Message "Account updated" -colour "DarkYellow"
        }
    catch{
        log-Error "Failed to update MSOL account"
        log-Error $Error
        }
    try{
        log-Message "Updating MSOL mailbox for $userSAM" -colour "Yellow"
        update-msolMailbox -userSAM $userSAM -userFirstName $adU.GivenName -userSurname $adU.Surname -userDisplayName $adU.DisplayName
        log-Message "Mailbox updated" -colour "DarkYellow"
        }
    catch{
        log-Error "Failed to update MSOL mailbox"
        log-Error $Error
        }
    }






provision-user -userSAM $userSAM -userFirstName $userFirstName -userSurname $userSurname -userManagerSAM $userManagerSAM -userDepartment $userDepartment -userJobTitle $userJobTitle -plaintextPassword $plaintextPassword
#Now assign the user a phone number via http://shoretel/shorewaredirector and set their ipPhone and telephoneNumber AD attributes
start-sleep -Seconds 10
update-msolUserFromAd -userSAM $userSAM
#foreach($license in $licenses){license-msolUser -userSAM $userSAM -licenseType $license}

