Import-Module -Name ActiveDirectory
Import-Module _PS_Library_MSOL.psm1
Import-Module _PS_Library_GeneralFunctionality
Import-Module _REST_Library-SPO.psm1
Import-Module _CSOM_Library-SPO

<#
$userSAM = "Ali.Midhani"
$userFirstName = "Ali"
$userSurname = "Midhani"
$userManagerSAM = "Duncan.Faulkes"
$userCommunity = "SPARK"
$userDepartment = "Sustain"
$userJobTitle = "Associate"
$plaintextPassword = ""
$licenses = @("E1")
$timeZone = "GMT Standard Time"
$countryLocale = "2057"
#>

$logFile = "C:\Scripts\Logs\provision-User.log"
$errorLogFile = "C:\Scripts\Logs\provision-User_Errors.log"
$smtpServer = "anthesisgroup-com.mail.protection.outlook.com"

$adCredentials = Get-Credential -Message "Enter local AD Administrator credentials to create a new user in AD" -UserName "$env:USERDOMAIN\username"
$msolCredentials = set-MsolCredentials #Set these once as a PSCredential object and use that to build the CSOM SharePointOnlineCredentials object and set the creds for REST
$restCredentials = new-spoCred -username $msolCredentials.UserName -securePassword $msolCredentials.Password
$csomCredentials = new-csomCredentials -username $msolCredentials.UserName -password $msolCredentials.Password
connect-ToMsol -credential $msolCredentials
connect-ToExo -credential $msolCredentials

$sharePointServerUrl = "https://anthesisllc.sharepoint.com"
$hrSite = "/teams/hr"
$taxonomyListName = "TaxonomyHiddenList"
$taxononmyData = get-itemsInList -serverUrl $sharePointServerUrl -sitePath $hrSite -listName $taxonomyListName -suppressProgress $false -restCreds $restCredentials -logFile $logFile

$newUserListName = "New User Requests"
#$oDataUnprocessedUsers = '$select=*,Line_x0020_Manager/Name,Line_x0020_Manager/Title,Prinicpal_x0020_Community_x0020_/Name,Prinicpal_x0020_Community_x0020_/Title,Primary_x0020_Team/Name,Primary_x0020_Team/Title,Additional_x0020_Teams/Id,Additional_x0020_Teams/Title&$filter=Current_x0020_Status eq ''1 - Waiting for IT Team to set up accounts''&$expand=Line_x0020_Manager/Id,Prinicpal_x0020_Community_x0020_/Id,Primary_x0020_Team/Id,Additional_x0020_Teams/Id'
$oDataUnprocessedUsers = "`$select=*"
$oDataUnprocessedUsers += ",Line_x0020_Manager/Name,Line_x0020_Manager/Title"
$oDataUnprocessedUsers += ",Prinicpal_x0020_Community_x0020_/Name,Prinicpal_x0020_Community_x0020_/Title"
$oDataUnprocessedUsers += ",Primary_x0020_Team/Name,Primary_x0020_Team/Title"
$oDataUnprocessedUsers += ",Additional_x0020_Teams/Id,Additional_x0020_Teams/Title"
$oDataUnprocessedUsers += "&`$filter=Current_x0020_Status eq '1 - Waiting for IT Team to set up accounts'"
$oDataUnprocessedUsers += "&`$expand=Line_x0020_Manager/Id,Prinicpal_x0020_Community_x0020_/Id,Primary_x0020_Team/Id,Additional_x0020_Teams/Id"
$unprocessedStarters = get-itemsInList -serverUrl $sharePointServerUrl -sitePath $hrSite -listName $newUserListName -suppressProgress $false -oDataQuery $oDataUnprocessedUsers -restCreds $restCredentials -logFile $logFile
#$unprocessedStarters | %{
if($null -ne $unprocessedStartersFormatted){rv unprocessedStartersFormatted}
$unprocessedStarters | %{[array]$unprocessedStartersFormatted += $(convert-listItemToCustomObject -spoListItem $_ -spoTaxonomyData $taxononmyData)}
$selectedStartersrs = $unprocessedStartersFormatted | Out-GridView -PassThru



#region functions
function create-ADUser($pUPN, $pFirstName, $pSurname, $pDisplayName, $pManagerSAM, $pPrimaryTeam, $pSecondaryTeams, $pJobTitle, $plaintextPassword, $pBusinessUnit, $adCredentials){
    #Set Domain-specific variables
    switch ($pBusinessUnit) {
        "Sustain Limited" {$upnSuffix = "@anthesisgroup.com"; $twitterAccount = "SustainLtd"; $DDI = "0117 403 2XXX"; $receptionDDI = "0117 403 2700";$ouDn = "OU=Users,OU=Sustain,DC=Sustainltd,DC=local"; $website = "www.sustain.co.uk"}
        "Anthesis (UK) Limited" {$upnSuffix = "@bf.local"; $twitterAccount = "anthesis_group"; $DDI = ""; $receptionDDI = "";$ouDn = "???,DC=Bf,DC=local"; $website = "www.anthesisgroup.com"}
        "Anthesis Consulting Group Ltd" {}
        "Anthesis LLC" {}
        default {}
        }
    #Create a new AD User account
    New-ADUser `
        -AccountPassword (ConvertTo-SecureString $plaintextPassword -AsPlainText -force) `
        -CannotChangePassword $False `
        -ChangePasswordAtLogon $False `
        -Company $pBusinessUnit `
        -Department $pPrimaryTeam `
        -DisplayName $pDisplayName `
        -EmailAddress $pUPN `
        -Enabled $true `
        -Fax $twitterAccount `
        -GivenName $pFirstName `
        -HomePage $website `
        -Manager $(Get-ADUser $pManagerSAM) `
        -Name "$pFirstName $pSurname"`
        -OfficePhone $DDI `
        -Path $ouDn `
        -SAMAccountName $pUPN `
        -Surname $pSurname `
        -Title $pJobTitle `
        -UserPrincipalName "$pUPN" `
        -EmailAddress "$pUPN$upnSuffix" `
        -OtherAttributes @{'ipPhone'="XXX";'pager'=$receptionDDI} `
        -Credential $adCredentials
    $newAdUserAccount = Get-ADUser -filter {UserPrincipalName -like $pUPN} -Credential $adCredentials 
    Get-ADGroup -Filter {name -like $pPrimaryTeam} | %{Add-ADGroupMember -Identity $_ -Members $newAdUserAccount -Credential $adCredentials}
    $pSecondaryTeams | %{Get-ADGroup -Filter {name -like $_}} | %{Add-ADGroupMember -Identity $_ -Members $newAdUserAccount -Credential $adCredentials}
    }
function create-msolUser($pUPN){
    #create the Mailbox rather than the MSOLUser, which will effectively create an unlicensed E1 user
    New-Mailbox -Name $pUPN.Replace("."," ") -Password (ConvertTo-SecureString -AsPlainText $plaintextPassword -Force) -MicrosoftOnlineServicesID $pUPN@anthesisgroup.com
    }
function license-msolUser($pUPN, $licenseType){
    switch ($licenseType){
        "E1" {$licenseToAssign = Get-MsolAccountSku | ?{$_.AccountSkuId -eq "AnthesisLLC:STANDARDPACK"}}
        "E3" {$licenseToAssign = Get-MsolAccountSku | ?{$_.AccountSkuId -eq "AnthesisLLC:ENTERPRISEPACK"}}
        "VISIO" {$licenseToAssign = Get-MsolAccountSku | ?{$_.AccountSkuId -eq "AnthesisLLC:VISIOCLIENT"}}
        "PROJECT" {$licenseToAssign = Get-MsolAccountSku | ?{$_.AccountSkuId -eq "AnthesisLLC:PROJECTPROFESSIONAL"}}
        "EMS" {$licenseToAssign = Get-MsolAccountSku | ?{$_.AccountSkuId -eq "AnthesisLLC:EMS"}}
        }
    Set-MsolUserLicense -UserPrincipalName $pUPN@anthesisgroup.com -AddLicenses $licenseToAssign.AccountSkuId
    }
function update-MsolUser($pUPN, $pFirstName, $pSurname, $pDisplayName, $pManagerSAM, $pPrimaryTeam, $pSecondaryTeams, $pPrimaryOffice, $pSecondaryOffice, $pCountry, $pJobTitle, $pDDI, $pMobile){
    switch($primaryOffice){
        "Home worker" {$streetAddress = $null;$postalCode=$null;$country=$pCountry;$usageLocation=$(get-2letterIsoCodeFromCountryName $pCountry)}
        "Bristol, GBR" {$streetAddress = "Royal London Buildings, 42-46 Baldwin Street";$postalCode="BS1 1PN";$country="United Kingdom";$usageLocation="GB"}
        "London, GBR" {$streetAddress = "Unit 12.2.1, The Leathermarket, 11-13 Weston Street";$postalCode="SE1 3ER";$country="United Kingdom";$usageLocation="GB"}
        "Oxford, GBR" {$streetAddress = "9 Newtec Place, Magdalen Road";$postalCode="OX4 1RE";$country="United Kingdom";$usageLocation="GB"}
        "Macclesfield, GBR" {$streetAddress = "Riverside Suite 1, Sunderland House, Sunderland St";$postalCode="SK11 6LF";$country="United Kingdom";$usageLocation="GB"}
        "Manchester, GBR" {$streetAddress = "40 King Street";$postalCode="M2 6BA";$country="United Kingdom";$usageLocation="GB"}
        "Manila, PHI" {}
        "Boulder, CO, USA" {$streetAddress = "1877 Broadway #100";$postalCode="80302";$country="United States";$usageLocation="US"}
        "Emeryville, CA, USA" {$streetAddress = "1900 Powell Street, Ste 600";$postalCode="94608";$country="United States";$usageLocation="US"}
        }
    #$msolUser = New-MsolUser `
    Set-MsolUser -UserPrincipalName "$pUPN@anthesisgroup.com" `
        -FirstName $pFirstName `
        -LastName $pSurname `
        -DisplayName $pDisplayName `
        -Title $pJobTitle `
        -Department $pPrimaryTeam `
        -Office $primaryOffice `
        -PhoneNumber $pDDI `
        -StreetAddress $streetAddress `
        -City $secondaryOffice `
        -PostalCode $postalCode `
        -Country $country `
        -UsageLocation $usageLocation `
        -StrongPasswordRequired $true 
        #-Password "Welcome123" `
        #-ForceChangePassword $true
    Add-DistributionGroupMember -Identity $pPrimaryTeam -Member $pUPN
    $pSecondaryTeams | % {Add-DistributionGroupMember -Identity $_ -Member $pUPN}
    }
function update-msolMailbox($pUPN,$pFirstName,$pSurname,$pDisplayName,$pBusinessUnit,$pTimeZone){
    Get-Mailbox $pUPN@anthesisgroup.com | Set-Mailbox  -CustomAttribute1 $pBusinessUnit -Alias $pUPN -DisplayName $pDisplayName -Name "$pFirstName $pSurname" -AuditEnabled $true
    if ($pBusinessUnit -match "Sustain"){Get-Mailbox $pUPN@anthesisgroup.com | Set-Mailbox -EmailAddresses @{add="$pUPN@sustain.co.uk"}}
    Get-Mailbox $pUPN@anthesisgroup.com | Set-CASMailbox -ActiveSyncMailboxPolicy "Sustain"
    Set-User -Identity $pUPN@anthesisgroup.com -Company $pBusinessUnit
    Set-MailboxRegionalConfiguration -Identity $pUPN@anthesisgroup.com -TimeZone $pTimeZone
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
function update-sharePointInitialConfig($pUPN, $anthesisAdminSite, $csomCreds, $timeZone, $p3LetterCountryIsoCode){
    $countryLocale = get-spoLocaleFromCountry -p3LetterCountryIsoCode $p3LetterCountryIsoCode
    $languageCode = guess-languageCodeFromCountry -p3LetterCountryIsoCode $p3LetterCountryIsoCode
    $adminContext = new-csomContext -fullSitePath $anthesisAdminSite -sharePointCredentials $csomCreds
    $spoUsers = New-Object Microsoft.SharePoint.Client.UserProfiles.PeopleManager($adminContext)
    $spoUsers.SetSingleValueProfileProperty("$pUPN@anthesisgroup.com", "SPS-RegionalSettings-Initialized", $true)
    $spoUsers.SetSingleValueProfileProperty("$pUPN@anthesisgroup.com", "SPS-RegionalSettings-FollowWeb", $false)
    #Getting the TimeZoneID is a massive PITA:
    if($timeZones -eq $null){$timeZones = get-timeZones}
    $tz = $timeZones | ?{$_.PSChildName -eq $timeZone} #Look that up in the registry list
    if($spoTimeZones -eq $null){$spoTimeZones = get-spoTimeZoneHashTable -credentials $csomCredentials}
    $tzID = $spoTimeZones[$tz.Display.replace("+00:00","")] #Then match a different property of the registry object to the SPO object
    if($tzID.Length -gt 0){$spoUsers.SetSingleValueProfileProperty("$pUPN@anthesisgroup.com", "SPS-TimeZone", $tzID)}
    if($countryLocale.length -gt 0){$spoUsers.SetSingleValueProfileProperty("$pUPN@anthesisgroup.com", "SPS-Locale", $countryLocale)}
    if($languageCode.length -gt 0){$spoUsers.SetSingleValueProfileProperty("$pUPN@anthesisgroup.com", "SPS-MUILanguages", $languageCode)}
    $spoUsers.SetSingleValueProfileProperty("$pUPN@anthesisgroup.com", "SPS-CalendarType", 1)
    $spoUsers.SetSingleValueProfileProperty("$pUPN@anthesisgroup.com", "SPS-AltCalendarType", 1)
    $adminContext.ExecuteQuery()
    }
function create-personalFolder($pUPN){
    $dirRoot = "X:\Personal"
    
    #Create the user's Personal Folder and give them Modify rights
    $personalFolder = New-Item -Path "$dirRoot\$pUPN" -ItemType Directory
    $acl = Get-Acl $personalFolder
    $perm = "Modify"
    $permInherit = "ContainerInherit, ObjectInherit" #This folder, files & subfolders - see http://powershell.nicoh.me/powershell-1/files-and-folders/set-folders-acl-owner-and-ntfs-rights 
    $permProp = "None" #This folder, files & subfolders - see http://powershell.nicoh.me/powershell-1/files-and-folders/set-folders-acl-owner-and-ntfs-rights 
    $ace = New-Object System.Security.AccessControl.FileSystemAccessRule($pUPN, $perm, $permInherit, $permProp, "Allow")
    $acl.AddAccessRule($ace)
    Set-Acl -Path $personalFolder -AclObject $acl

    #Create the user's Secure folder, break the inheritance permissions
    $secureFolder = New-Item -Path "$dirRoot\$pUPN\Secure" -ItemType Directory
    $acl = Get-Acl $secureFolder
    $acl.SetAccessRuleProtection($true,$true)  #Note that SetAccessRuleProtection takes two boolean arguments; the first turns inheritance on ($False) or off ($True) and the second determines whether the previously inherited permissions are retained ($True) or removed ($False)
    Set-Acl -Path $secureFolder -AclObject $acl
    
    #Now remove all permissions that are not the user's or backup-related
    foreach ($ace in $acl.Access){
        if (!($ace.IdentityReference -eq "SUSTAINLTD\Backup Process Account - do not block permissions" -or $ace.IdentityReference -eq "SUSTAINLTD\$pUPN")){
            icacls $secureFolder /remove `"$($ace.IdentityReference)`" | Out-Null
            }
        }
    }
function set-mailboxPermissions($pUPN,$pManagerSAM,$pBusinessUnit){
    Add-MailboxPermission -Identity $pUPN -AccessRights FullAccess -User $pManagerSAM -InheritanceType all -AutoMapping $false
    if($pBusinessUnit -match "Sustain"){
        Add-MailboxPermission -Identity $pUPN -AccessRights FullAccess -user SustainMailboxAccess@anthesisgroup.com
        #Add-MailboxPermission -Identity $pUPN -AccessRights SendAs -User SustainMailboxAccess@anthesisgroup.com -InheritanceType all
        Add-MailboxFolderPermission "$($pUPN):\Calendar" -User "View all Sustain calendars" -AccessRights "Reviewer"
        Add-MailboxFolderPermission "$($pUPN):\Calendar" -User "Edit all Sustain calendars" -AccessRights "PublishingEditor"
        }
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

#region meta-functions
function provision-user($userUPN, $userFirstName, $userSurname, $userManagerSAM, $userCommunity, $userPrimaryTeam, $userSecondaryTeams, $userBusinessUnit, $userJobTitle, $plaintextPassword, $adCredentials){
    try{
        log-Message "Creating AD account for $userUPN" -colour "Yellow"
        create-ADUser -pUPN $userUPN -pFirstName $userFirstName -pSurname $userSurname -pDisplayName $userDisplayName -pManagerSAM $userManagerSAM -pDepartment $userDepartment -pJobTitle $userJobTitle -plaintextPassword $plaintextPassword -pPrimaryTeam $userPrimaryTeam -pSecondaryTeams $userSecondaryTeams -pBusinessUnit $userBusinessUnit -adCredentials $adCredentials
        log-Message "Account created" -colour "DarkYellow"
        }
    catch{
        log-Error "Failed to create AD account"
        log-Error $Error
        }
    try{
        log-Message "Creating MSOL account for $userUPN" -colour "Yellow"
        create-msolUser -userUPN $userUPN
        log-Message "Account created" -colour "DarkYellow"
        }
    catch{
        log-Error "Failed to create MSOL account"
        log-Error $Error
        }
    Start-Sleep -Seconds 5 #Give EXO a chance to catch up
    try{
        log-Message "Setting mailbox permissions for $userUPN" -colour "Yellow"
        set-mailboxPermissions -pUPN $userUPN -pManagerSAM $userManagerSAM -pBusinessUnit $userBusinessUnit
        log-Message "Mailbox permissions set" -colour "DarkYellow"
        }
    catch{
        log-Error "Failed to set mailbox permissions"
        log-Error $Error
        }
    try{
        log-Message "Updating mailbox for $userUPN" -colour "Yellow"
        update-msolMailbox -pUPN $userUPN -pFirstName $userFirstName -pSurname $userSurname -pDisplayName $userDisplayName -pBusinessUnit $userBusinessUnit -pTimeZone $userTimeZone
        log-Message "Mailbox updated" -colour "DarkYellow"
        }
    catch{
        log-Error "Failed to update mailbox"
        log-Error $Error
        }
    if($userBusinessUnit -match "Sustain"){
        try{
            log-Message "Creating personal folder for $userUPN" -colour "Yellow"
            create-personalFolder -pUPN $userUPN
            log-Message "Folder created" -colour "DarkYellow"
            }
        catch{
            log-Error "Failed to create personal folder"
            log-Error $Error
            }
        }
    }
function update-msolUserFromAd($userUPN){
    $adU = Get-ADUser -filter {UserPrincipalName -like $userUPN} -Properties DisplayName,Title,Department,Office,ipPhone,Manager,*
    $userManagerSAM = (Get-ADUser $adu.Manager).SamAccountName
    $DDI = format-internationalPhoneNumber -pDirtyNumber $adu.OfficePhone -p3letterIsoCountryCode $(get-3letterIsoCodeFromCountryName -pCountryName $adu.Country)
    $mobile = format-internationalPhoneNumber -pDirtyNumber $adu.MobilePhone -p3letterIsoCountryCode $(get-3letterIsoCodeFromCountryName -pCountryName $adu.Country)


    try{
        log-Message "Updating MSOL account for $userUPN" -colour "Yellow"
        update-MsolUser -pUPN $userUPN -pFirstName $adU.GivenName -pSurname $adU.Surname -pManagerSAM $userManagerSAM -pDDI $DDI -pMobile $mobile  -pJobTitle $adU.Title -pDisplayName $adU.DisplayName -pPhoneExtension $adU.ipPhone
        log-Message "Account updated" -colour "DarkYellow"
        }
    catch{
        log-Error "Failed to update MSOL account"
        log-Error $Error
        }
    try{
        log-Message "Updating MSOL mailbox for $userUPN" -colour "Yellow"
        ######This is as far as you got!!
        update-msolMailbox -userUPN $userUPN -userFirstName $adU.GivenName -userSurname $adU.Surname -userDisplayName $adU.DisplayName
        log-Message "Mailbox updated" -colour "DarkYellow"
        }
    catch{
        log-Error "Failed to update MSOL mailbox"
        log-Error $Error
        }
    }
#endregion





provision-user -userUPN $userUPN -userFirstName $userFirstName -userSurname $userSurname -userManagerSAM $userManagerSAM -userDepartment $userDepartment -userJobTitle $userJobTitle -plaintextPassword $plaintextPassword
#Now assign the user a phone number via http://shoretel/shorewaredirector and set their ipPhone and telephoneNumber AD attributes
start-sleep -Seconds 10
update-msolUserFromAd -userUPN $userUPN
#foreach($license in $licenses){license-msolUser -userUPN $userUPN -licenseType $license}

