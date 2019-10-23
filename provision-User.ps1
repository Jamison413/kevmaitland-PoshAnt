Import-Module -Name ActiveDirectory

<#
$userSAM = "Dummy.User"
$userFirstName = "Dummy"
$userSurname = "User"
$userManagerSAM = "Kevin.Maitland"
$userCommunity = "Roots"
$userDepartment = "Sustain"
$userJobTitle = "Associate"
$plaintextPassword = ""
$licenses = @("E1")
$timeZone = "GMT Standard Time"
$countryLocale = "2057"
#>

$logFile = "C:\ScriptLogs\provision-User.log"
$errorLogFile = "C:\ScriptLogs\provision-User_Errors.log"
$smtpServer = "anthesisgroup-com.mail.protection.outlook.com"

$msolCredentials = set-MsolCredentials #Set these once as a PSCredential object and use that to build the CSOM SharePointOnlineCredentials object and set the creds for REST
$restCredentials = new-spoCred -username $msolCredentials.UserName -securePassword $msolCredentials.Password
$csomCredentials = new-csomCredentials -username $msolCredentials.UserName -password $msolCredentials.Password
connect-ToMsol -credential $msolCredentials
connect-ToExo -credential $msolCredentials
connect-toAAD -credential $msolCredentials
#connect-ToSpo -credential $msolCredentials

$adCredentials = Get-Credential -Message "Enter local AD Administrator credentials to create a new user in AD" -UserName "$env:USERDOMAIN\username"

#Get the New User Requests that have not been marked as processed
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/teams/hr" -Credentials $msolCredentials
$requests = (Get-PnPListItem -List "New User Requests" -Query "<View><Query><Where><Eq><FieldRef Name='Current_x0020_Status'/><Value Type='String'>1 - Waiting for IT Team to set up accounts</Value></Eq></Where></Query></View>") |  % {Add-Member -InputObject $_ -MemberType NoteProperty -Name Guid -Value $_.FieldValues.GUID.Guid;$_}
if($requests){#Display a subset of Properties to help the user identify the correct account(s)
    $selectedRequests = $requests | Sort-Object -Property {$_.FieldValues.Start_x0020_Date} -Descending | select {$_.FieldValues.Title},{$_.FieldValues.Start_x0020_Date},{$_.FieldValues.Job_x0020_title},{$_.FieldValues.Primary_x0020_Workplace.Label},{$_.FieldValues.Line_x0020_Manager.LookupValue},{$_.FieldValues.Primary_x0020_Team.LookupValue},{$_.FieldValues.GUID.Guid} | Out-GridView -PassThru -Title "Highlight any requests to process and click OK" | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name "Guid" -Value $_.'$_.FieldValues.GUID.Guid';$_}
    #Then return the original requests as these contain the full details
    [array]$selectedRequests = Compare-Object -ReferenceObject $requests -DifferenceObject $selectedRequests -Property Guid -IncludeEqual -ExcludeDifferent -PassThru
    }


#region functions
function create-ADUser($pUPN, $pFirstName, $pSurname, $pDisplayName, $pManagerSAM, $pPrimaryTeam, $pSecondaryTeams, $pJobTitle, $plaintextPassword, $pBusinessUnit, $adCredentials, $pPrimaryOffice){
    #Set Domain-specific variables
    <#
    $pUPN = $userUPN; $pFirstName = $userFirstName; $pSurname = $userSurname;$pDisplayName=$userDisplayName;$pManagerSAM=$userManagerSAM;$pPrimaryTeam=$userPrimaryTeam;$pSecondaryTeams=$userSecondaryTeams;$pJobTitle=$userJobTitle;$pBusinessUnit=$userBusinessUnit;$pPrimaryOffice=$userPrimaryOffice
    #> 
    switch ($pBusinessUnit) {
        "Anthesis Energy UK Ltd (GBR)" {$upnSuffix = "@anthesisgroup.com"; $twitterAccount = "anthesis_group"; $DDI = "0117 403 2XXX"; $receptionDDI = "0117 403 2700";$ouDn = "OU=Users,OU=Sustain,DC=Sustainltd,DC=local"; $website = "www.anthesisgroup.com"}
        "Anthesis (UK) Ltd (GBR)"  {write-host -ForegroundColor Magenta "AUK, but creating Sustain account"; $upnSuffix = "@anthesisgroup.com"; $twitterAccount = "anthesis_group"; $DDI = "0117 403 2XXX"; $receptionDDI = "0117 403 2700";$ouDn = "OU=Users,OU=Sustain,DC=Sustainltd,DC=local"; $website = "www.anthesisgroup.com"}
        #"Anthesis (UK) Limited (GBR)" {$upnSuffix = "@bf.local"; $twitterAccount = "anthesis_group"; $DDI = ""; $receptionDDI = "";$ouDn = "???,DC=Bf,DC=local"; $website = "www.anthesisgroup.com"}
        "Anthesis Consulting Group Ltd (GBR)" {}
        "Anthesis LLC" {}
        default {Write-Host -ForegroundColor DarkRed "Warning: Could not not identify Business Unit [$pBusinessUnit]"}
        }
    #Create a new AD User account
    write-host -ForegroundColor DarkYellow "UPN:`t$pUPN"
    write-host -ForegroundColor DarkYellow "GivenName:`t$pFirstName"
    write-host -ForegroundColor DarkYellow "Surname:`t$pSurname"
    write-host -ForegroundColor DarkYellow "Company:`t$pBusinessUnit"
    write-host -ForegroundColor DarkYellow "Department:`t$pPrimaryTeam"
    write-host -ForegroundColor DarkYellow "DisplayName:`t$pDisplayName"
    write-host -ForegroundColor DarkYellow "Fax:`t$twitterAccount"
    write-host -ForegroundColor DarkYellow "HomePage:`t$website"
    write-host -ForegroundColor DarkYellow "Manager:`t$(Get-ADUser -filter {SamAccountName -like $pManagerSAM})"
    write-host -ForegroundColor DarkYellow "Name:`t$pFirstName $pSurname"
    write-host -ForegroundColor DarkYellow "Office:`t$pPrimaryOffice"
    write-host -ForegroundColor DarkYellow "OfficePhone:`t$DDI"
    write-host -ForegroundColor DarkYellow "Title:`t$pJobTitle"
    write-host -ForegroundColor DarkYellow "EmailAddress:`t$($pUPN.Split("@")[0])$upnSuffix"
    write-host -ForegroundColor DarkYellow "pager:`t$receptionDDI"
    write-host -ForegroundColor DarkYellow "Path:`t$ouDn"
    write-host -ForegroundColor DarkYellow "SAMAccountName:`t$($pUPN.Split("@")[0])"
    New-ADUser `
        -AccountPassword (ConvertTo-SecureString $plaintextPassword -AsPlainText -force) `
        -CannotChangePassword $False `
        -ChangePasswordAtLogon $False `
        -Company $pBusinessUnit `
        -Department $pPrimaryTeam `
        -DisplayName $pDisplayName `
        -Enabled $true `
        -Fax $twitterAccount `
        -GivenName $pFirstName `
        -HomePage $website `
        -Manager $(Get-ADUser -filter {SamAccountName -like $pManagerSAM}) `
        -Name "$pFirstName $pSurname"`
        -Office $pPrimaryOffice `
        -OfficePhone $DDI `
        -Path $ouDn `
        -SAMAccountName $($pUPN.Split("@")[0]) `
        -Surname $pSurname `
        -Title $pJobTitle `
        -UserPrincipalName "$($pUPN.Split("@")[0])$upnSuffix" `
        -EmailAddress "$($pUPN.Split("@")[0])$upnSuffix" `
        -OtherAttributes @{'ipPhone'="XXX";'pager'=$receptionDDI} `
        -Credential $adCredentials
    $newAdUserAccount = Get-ADUser -filter {UserPrincipalName -like $pUPN} -Credential $adCredentials 
    $primaryTeam = Get-ADGroup -Filter {name -like $pPrimaryTeam} 
    if($primaryTeam){
        Write-Host "Adding [$($newAdUserAccount.Name)] to [$($primaryTeam.Name)]"
        Add-ADGroupMember -Identity $primaryTeam.ObjectGUID -Members $newAdUserAccount -Credential $adCredentials
        }
    if($pSecondaryTeams){
        $pSecondaryTeams | %{Get-ADGroup -Filter {name -like $_}} | %{Add-ADGroupMember -Identity $_ -Members $newAdUserAccount -Credential $adCredentials}
        }
    }
function create-msolUser($pUPN,$pPlaintextPassword){
    #create the Mailbox rather than the MSOLUser, which will effectively create an unlicensed E1 user
    New-Mailbox -Name $pUPN.Replace("."," ").Split("@")[0] -Password (ConvertTo-SecureString -AsPlainText $pPlaintextPassword -Force) -MicrosoftOnlineServicesID $pUPN
    }
function license-msolUser($pUPN, $licenseType, $usageLocation){
    switch ($licenseType){
        "E1" {
            $licenseToAssign = Get-MsolAccountSku | ?{$_.AccountSkuId -eq "AnthesisLLC:STANDARDPACK"}
            if((Get-MsolUser -UserPrincipalName $pUPN).Licenses.AccountSkuId -contains "AnthesisLLC:ENTERPRISEPACK"){$licenseToRemove = Get-MsolAccountSku | ?{$_.AccountSkuId -eq "AnthesisLLC:ENTERPRISEPACK"}}
            }
        "E3" {
            $licenseToAssign = Get-MsolAccountSku | ?{$_.AccountSkuId -eq "AnthesisLLC:ENTERPRISEPACK"}
            if((Get-MsolUser -UserPrincipalName $pUPN).Licenses.AccountSkuId -contains "AnthesisLLC:STANDARDPACK"){$licenseToRemove = Get-MsolAccountSku | ?{$_.AccountSkuId -eq "AnthesisLLC:STANDARDPACK"}}
            }
        "VISIO" {$licenseToAssign = Get-MsolAccountSku | ?{$_.AccountSkuId -eq "AnthesisLLC:VISIOCLIENT"}}
        "PROJECT" {$licenseToAssign = Get-MsolAccountSku | ?{$_.AccountSkuId -eq "AnthesisLLC:PROJECTPROFESSIONAL"}}
        "EMS" {$licenseToAssign = Get-MsolAccountSku | ?{$_.AccountSkuId -eq "AnthesisLLC:EMS"}}
        }
    Write-Host -ForegroundColor Yellow "Set-MsolUserLicense -UserPrincipalName $pUPN -AddLicenses $($licenseToAssign.AccountSkuId) -RemoveLicenses $($licenseToRemove.AccountSkuId)"
    Set-MsolUserLicense -UserPrincipalName $pUPN -AddLicenses $licenseToAssign.AccountSkuId -RemoveLicenses $licenseToRemove.AccountSkuId
    }
function update-MsolUser($pUPN, $pFirstName, $pSurname, $pDisplayName, $pPrimaryTeam, $pSecondaryTeams, $pPrimaryOffice, $pSecondaryOffice, $pCountry, $pJobTitle, $pDDI, $pMobile){
    #$pUPN = $userUPN; $pFirstName = $userFirstName; $pSurname = $userSurname;$pDisplayName=$userDisplayName;$pPrimaryTeam=$userPrimaryTeam;$pSecondaryTeams=$userSecondaryTeams;$pPrimaryOffice=$userPrimaryOffice;$pSecondaryOffice=$userSecondaryOffice;$pJobTitle=$userJobTitle;$pDDI=$userDDI;$pMobile=$userMobile
    $currentUser = Get-MsolUser -UserPrincipalName $pUPN
    if([string]::IsNullOrEmpty($pFirstName)){$firstName = $currentUser.FirstName}else{$firstname = $pFirstName}
    if([string]::IsNullOrEmpty($firstName)){$firstName = $currentUser.DisplayName.Split(" ")[0]}
    if([string]::IsNullOrEmpty($pSurname)){$surName = $currentUser.LastName}else{$surName = $pSurname}
    if([string]::IsNullOrEmpty($surname)){$surname = $currentUser.DisplayName.Split(" ")[$currentUser.DisplayName.Split(" ").Count-1]}
    if([string]::IsNullOrEmpty($pDisplayName)){$displayName = $currentUser.DisplayName}else{$displayName = $pDisplayName}
    if([string]::IsNullOrEmpty($pPrimaryOffice)){$primaryOffice = $currentUser.Office}else{$primaryOffice = $pPrimaryOffice}
    if([string]::IsNullOrEmpty($pSecondaryOffice)){$secondaryOffice = $primaryOffice}else{$secondaryOffice = $pSecondaryOffice}
    if([string]::IsNullOrEmpty($pJobTitle)){$jobTitle = $currentUser.Title}else{$jobTitle = $pJobTitle}
    if([string]::IsNullOrEmpty($pDDI)){$ddi = $currentUser.PhoneNumber}else{$ddi = $pDDI}
    if([string]::IsNullOrEmpty($pMobile)){$mobile = $currentUser.MobilePhone}else{$mobile = $pMobile}
    
    Write-Host -ForegroundColor DarkYellow "`tPrimaryOffice: $pPrimaryOffice"
    switch($pPrimaryOffice){
        "Home worker" {$streetAddress = $null;$postalCode=$null;$country=$pCountry;$usageLocation=$(get-2letterIsoCodeFromCountryName $pCountry;$group = "All Homeworkers")}
        "Bristol, GBR" {$streetAddress = "Royal London Buildings, 42-46 Baldwin Street";$postalCode="BS1 1PN";$country="United Kingdom";$usageLocation="GB";$group = "All Bristol (GBR)"}
        "London, GBR" {$streetAddress = "Unit 12.2.1, The Leathermarket, 11-13 Weston Street";$postalCode="SE1 3ER";$country="United Kingdom";$usageLocation="GB";$group = "All London (GBR)"}
        "Oxford, GBR" {$streetAddress = "9 Newtec Place, Magdalen Road";$postalCode="OX4 1RE";$country="United Kingdom";$usageLocation="GB";$group = "All Oxford (GBR)"}
        "Macclesfield, GBR" {$streetAddress = "Riverside Suite 1, Sunderland House, Sunderland St";$postalCode="SK11 6LF";$country="United Kingdom";$usageLocation="GB";$group = "All Macclesfield (GBR)"}
        "Manchester, GBR" {$streetAddress = "40 King Street";$postalCode="M2 6BA";$country="United Kingdom";$usageLocation="GB";$group = "All Manchester (GBR)"}
        "Dubai, ARE" {$streetAddress = "1605 The Metropolis Building, Burj Khalifa St";$postalCode="PO Box 392563";$country="United Arab Emirates";$usageLocation="AE";$group = "All (ARE)"}
        "Manila, PHI" {$streetAddress = "10F Unit C & D, Strata 100 Condominium, F. Ortigas Jr. Road, Ortigas Center Brgy. San Antonio";$postalCode="1605";$country="Philippines";$usageLocation="PH";$group = "All (PHI)"}
        "Frankfurt, DEU" {$streetAddress = "Münchener Str. 7";$postalCode="60329";$country="Germany";$usageLocation="DE";$group = "All (DEU)"}
        "Nuremberg, DEU" {$streetAddress = "Sulzbacher Str. 70";$postalCode="90489";$country="Germany";$usageLocation="DE";$group = "All (DEU)"}
        "Boulder, CO, USA" {$streetAddress = "1877 Broadway #100";$postalCode="80302";$country="United States";$usageLocation="US";$group = "All (North America)"}
        "Emeryville, CA, USA" {$streetAddress = "1900 Powell Street, Ste 600";$postalCode="94608";$country="United States";$usageLocation="US";$group = "All (North America)"}
        "Stockholm, SWE" {$streetAddress = "Barnhusgatan 4";$postalCode="SE-111 23";$country="Sweden";$usageLocation="SE";$group = "All (SWE)"}
        default {$streetAddress = $currentUser.StreetAddress;$postalCode=$currentUser.PostalCode;$country=$currentUser.Country;$usageLocation=$currentUser.UsageLocation}
        }
    Write-Host -ForegroundColor DarkYellow "`tUsername:`t`t`t$($pUPN.Split("@")[0])@anthesisgroup.com"
    Write-Host -ForegroundColor DarkYellow "`tfirstName:`t`t`t$firstName"
    Write-Host -ForegroundColor DarkYellow "`tsurname:`t`t`t$surname"
    Write-Host -ForegroundColor DarkYellow "`tdisplayName:`t`t$displayName"
    Write-Host -ForegroundColor DarkYellow "`tjobTitle:`t`t`t$jobTitle"
    Write-Host -ForegroundColor DarkYellow "`tprimaryTeam:`t`t$primaryTeam"
    Write-Host -ForegroundColor DarkYellow "`tddi:`t`t`t`t$ddi"
    Write-Host -ForegroundColor DarkYellow "`tStreetAddress:`t`t$streetAddress"
    Write-Host -ForegroundColor DarkYellow "`tsecondaryOffice:`t$secondaryOffice"
    Write-Host -ForegroundColor DarkYellow "`tpostalCode:`t`t`t$postalCode"
    Write-Host -ForegroundColor DarkYellow "`tusageLocation:`t`t$usageLocation"
    #$msolUser = New-MsolUser `
    Set-MsolUser -UserPrincipalName "$($pUPN.Split("@")[0])@anthesisgroup.com" `
        -FirstName $firstName `
        -LastName $surname `
        -DisplayName $displayName `
        -Title $jobTitle `
        -Department $primaryTeam `
        -Office $primaryOffice `
        -PhoneNumber $ddi `
        -StreetAddress $streetAddress `
        -City $secondaryOffice `
        -PostalCode $postalCode `
        -Country $country `
        -UsageLocation $usageLocation `
        -StrongPasswordRequired $true 
        #-Password "Welcome123" `
        #-ForceChangePassword $true
    if($pPrimaryTeam -ne $null){Add-DistributionGroupMember -Identity $pPrimaryTeam -Member $pUPN -BypassSecurityGroupManagerCheck}
    if($pSecondaryTeams -ne $null){$pSecondaryTeams | % {Add-DistributionGroupMember -Identity $_ -Member $pUPN -BypassSecurityGroupManagerCheck}}
    if($group -ne $null){Add-DistributionGroupMember -Identity $group -Member $pUPN -BypassSecurityGroupManagerCheck}
    Add-DistributionGroupMember -Identity "b264f337-ef04-432e-a139-3574331a4d18" -Member $pUPN -BypassSecurityGroupManagerCheck #"MDM - BYOD Users"
    }
function update-msolMailbox($pUPN,$pFirstName,$pSurname,$pDisplayName,$pBusinessUnit,$pTimeZone){
    #$pUPN = $userUPN; $pFirstName = $userFirstName; $pSurname = $userSurname;$pDisplayName=$userDisplayName;$pBusinessUnit=$userBusinessUnit,$pTimeZone=$userTimeZone
    Get-Mailbox $pUPN | Set-Mailbox  -CustomAttribute1 $pBusinessUnit -Alias $($pUPN.Split("@")[0]) -DisplayName $pDisplayName -Name "$pFirstName $pSurname" -AuditEnabled $true -AuditLogAgeLimit 180 -AuditAdmin Update, MoveToDeletedItems, SoftDelete, HardDelete, SendAs, SendOnBehalf, Create, UpdateFolderPermission -AuditDelegate Update, SoftDelete, HardDelete, SendAs, Create, UpdateFolderPermissions, MoveToDeletedItems, SendOnBehalf -AuditOwner UpdateFolderPermission, MailboxLogin, Create, SoftDelete, HardDelete, Update, MoveToDeletedItems 
    if ($pBusinessUnit -match "Germany"){Get-Mailbox $pUPN | Set-Mailbox -LitigationHoldEnabled $true -RetentionComment "Ligation Hold (DEU)" -RetentionUrl "https://anthesisllc.sharepoint.com/sites/Resources-IT/SitePages/ALessSinisterExplaination.aspx"}
    if ($pBusinessUnit -match "UK"){Get-Mailbox $pUPN | Set-Mailbox -LitigationHoldEnabled $true -RetentionComment "Ligation Hold (GBR - Energy)" -RetentionUrl "https://anthesisllc.sharepoint.com/sites/Resources-IT/SitePages/ALessSinisterExplaination.aspx"}
    #Get-Mailbox $pUPN | Set-CASMailbox -ActiveSyncMailboxPolicy "Sustain"
    #Get-Mailbox $pUPN | Set-Clutter -Enable $true
    Set-User -Identity $pUPN -Company $pBusinessUnit -Manager $pLineManager
    Set-MailboxRegionalConfiguration -Identity $pUPN -TimeZone $(convertTo-exTimeZoneValue $pTimeZone)
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
function update-sharePointInitialConfig($pUPN, $csomCreds, $pTimeZone, $p3LetterCountryIsoCode){
    $spoLoginPrefix = "i:0#.f|membership|"
    $countryLocale = get-spoLocaleFromCountry -p3LetterCountryIsoCode $p3LetterCountryIsoCode
    $languageCode = guess-languageCodeFromCountry -p3LetterCountryIsoCode $p3LetterCountryIsoCode

    #Getting the TimeZoneID is a massive PITA:
    $timeZones = get-timeZones
    $tz = $timeZones | ?{$_.PSChildName -eq $pTimeZone} #Look that up in the registry list
    if($spoTimeZones -eq $null){$spoTimeZones = get-spoTimeZoneHashTable -credentials $csomCredentials}
    $tzID = $spoTimeZones[$tz.Display.replace("+00:00","")] #Then match a different property of the registry object to the SPO object

    $adminContext = new-csomContext -fullSitePath "https://anthesisllc-admin.sharepoint.com/" -sharePointCredentials $csomCreds
    $spoUsers = New-Object Microsoft.SharePoint.Client.UserProfiles.PeopleManager($adminContext)
    $spoUsers.SetSingleValueProfileProperty($spoLoginPrefix+$pUPN, "SPS-RegionalSettings-Initialized", $true)
    $spoUsers.SetSingleValueProfileProperty($spoLoginPrefix+$pUPN, "SPS-RegionalSettings-FollowWeb", $false)
    if($tzID.Length -gt 0){$spoUsers.SetSingleValueProfileProperty($spoLoginPrefix+$pUPN, "SPS-TimeZone", $tzID)}
    if($countryLocale.length -gt 0){$spoUsers.SetSingleValueProfileProperty($spoLoginPrefix+$pUPN, "SPS-Locale", $countryLocale)}
    if($languageCode.length -gt 0){$spoUsers.SetSingleValueProfileProperty($spoLoginPrefix+$pUPN, "SPS-MUILanguages", $languageCode)}
    $spoUsers.SetSingleValueProfileProperty($spoLoginPrefix+$pUPN, "SPS-CalendarType", 1)
    $spoUsers.SetSingleValueProfileProperty($spoLoginPrefix+$pUPN, "SPS-AltCalendarType", 1)
    $adminContext.ExecuteQuery()
    }
function create-personalFolder($pUPN){
    $dirRoot = "X:\Personal"
    $username = $pUPN.split("@")[0]
    #Create the user's Personal Folder and give them Modify rights
    $personalFolder = New-Item -Path "$dirRoot\$username" -ItemType Directory -Credential $adCredentials
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
    Add-MailboxFolderPermission "$($pUPN):\Calendar" -User "All$(get-3lettersInBrackets -stringMaybeContaining3LettersInBrackets $pBusinessUnit)@anthesisgroup.com" -AccessRights "LimitedDetails"
    if($pBusinessUnit -match "Anthesis Energy UK Ltd (GBR)"){
        Add-MailboxPermission -Identity $pUPN -AccessRights FullAccess -user SustainMailboxAccess@anthesisgroup.com
        #Add-MailboxPermission -Identity $pUPN -AccessRights SendAs -User SustainMailboxAccess@anthesisgroup.com -InheritanceType all
        #Add-MailboxFolderPermission "$($pUPN):\Calendar" -User "View all Sustain calendars" -AccessRights "Reviewer"
        #Add-MailboxFolderPermission "$($pUPN):\Calendar" -User "Edit all Sustain calendars" -AccessRights "PublishingEditor"
        }
    }
function update-newUserRequest($listItem, $digest, $restCredentials, $logFile){
    $digest = check-digestExpiry -serverUrl $sharePointServerUrl -sitePath $hrSite -digest $digest -restCreds $restCredentials -logFile $logFile
    [guid]$listGuid = [regex]::Match($listItem.__metadata.uri,"\'([^)]+)\'").Groups[1].Value
    update-itemInList -serverUrl $sharePointServerUrl -sitePath $hrSite -listNameOrGuid $listGuid -predeterminedItemType $listItem.__metadata.type -itemId $listItem.Id -hashTableOfItemData @{"Current_x0020_Status"="2 - Waiting for HR to populate user data";"Previous_x0020_Status"=$listItem.Current_Status} -restCreds $restCredentials -digest $digest -logFile $logFile
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
function provision-365user($userUPN, $userFirstName, $userSurname, $userDisplayName, $userManagerSAM, $userCommunity, $userPrimaryTeam, $userSecondaryTeams, $userPrimaryOffice, $userSecondaryOffice, $userBusinessUnit, $userJobTitle, $plaintextPassword, $restCredentials, $newUserListItem, $userTimeZone, $user365License){
    $lastError = $Error[0]
    if ([string]::IsNullOrEmpty($userDisplayName)){$userDisplayName = "$userFirstName $userSurname"}
    try{
        log-Message "Creating MSOL account for $userUPN" -colour "Yellow"
        create-MsolUser -pUPN $userUPN -pPlaintextPassword $plaintextPassword
        log-Message "Account created" -colour "DarkYellow"
        }
    catch{
        log-Error "Failed to create MSOL account"
        log-Error $Error
        }
    Start-Sleep -Seconds 20 #Let MSOL & EXO Syncronise
    try{
        log-Message "Updating MSOL account for $userUPN" -colour "Yellow"
        update-MsolUser -pUPN $userUPN -pPrimaryOffice $userPrimaryOffice -pSecondaryOffice $userSecondaryOffice -pPrimaryTeam $userPrimaryTeam -pSecondaryTeams $userSecondaryTeams -pJobTitle $userJobTitle #-pSurname $userSurname -pDisplayName $userDisplayName
        log-Message "Account updated" -colour "DarkYellow"
        }
    catch{
        log-Error "Failed to update MSOL account"
        log-Error $Error
        }
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
    try{
        log-Message "Licensing $userUPN with $user365License" -colour "Yellow"
        license-msolUser -pUPN $userUPN -licenseType $user365License
        log-Message "Mailbox updated" -colour "DarkYellow"
        }
    catch{
        log-Error "Failed to update mailbox"
        log-Error $Error
        }
    try{
        log-Message "Setting SharePoint Timezone" -colour "Yellow"
        $isoCountryCode = $(get-3letterIsoCodeFromCountryName -pCountryName (get-3lettersInBrackets -stringMaybeContaining3LettersInBrackets $userPrimaryOffice))
        if([string]::IsNullOrWhiteSpace($isoCountryCode)){$isoCountryCode = get-3letterIsoCodeFromCountryName -pCountryName (get-trailing3LettersIfTheyLookLikeAnIsoCountryCode -ambiguousString $userPrimaryOffice)}
        if(![string]::IsNullOrWhiteSpace($isoCountryCode)){update-sharePointInitialConfig -pUPN $userUPN -csomCreds $csomCredentials -pTimeZone $userTimeZone -p3LetterCountryIsoCode $isoCountryCode}
        log-Message "SharePoint Timezone updated" -colour "DarkYellow"
        }
    catch{
        log-Error "Failed to update SharePoint Timezone"
        log-Error $Error
        }

    if ($lastError -eq $Error[0]){
        update-newUserRequest -listItem $newUserListItem -digest $digest -restCredentials $restCredentials -logFile $logFile
        }
    }
function provision-SustainADUser($userUPN, $userFirstName, $userSurname, $userDisplayName, $userManagerSAM, $userCommunity, $userPrimaryTeam, $userSecondaryTeams, $userBusinessUnit, $userJobTitle, $userPrimaryOffice, $plaintextPassword, $adCredentials){
    try{
        log-Message "Creating AD account for $userUPN" -colour "Yellow"
        create-ADUser -pUPN $userUPN -pFirstName $userFirstName -pSurname $userSurname -pDisplayName $userDisplayName -pManagerSAM $($userManagerSAM.Split("@")[0]) -pDepartment $userPrimaryTeam -pJobTitle $userJobTitle -plaintextPassword $plaintextPassword -pPrimaryTeam $userPrimaryTeam -pSecondaryTeams $userSecondaryTeams -pBusinessUnit $userBusinessUnit -adCredentials $adCredentials -pPrimaryOffice $userPrimaryOffice
        log-Message "Account created" -colour "DarkYellow"
        }
    catch{
        log-Error "Failed to create AD account"
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
        update-MsolUser -pUPN $userUPN -pFirstName $adU.GivenName -pSurname $adU.Surname -pManagerSAM $userManagerSAM -pDDI $DDI -pMobile $mobile  -pJobTitle $adU.Title -pDisplayName $adU.DisplayName -pPhoneExtension $adU.ipPhone -pPrimaryOffice $adu.physicalDeliveryOfficeName
        log-Message "Account updated" -colour "DarkYellow"
        }
    catch{
        log-Error "Failed to update MSOL account"
        log-Error $Error
        }
    try{
        log-Message "Updating MSOL mailbox for $userUPN" -colour "Yellow"
        update-msolMailbox -pUPN $userUPN -pFirstName $adU.GivenName -pSurname $adU.Surname -pDisplayName $adU.DisplayName -pBusinessUnit $adU.Company -pTimeZone "GMT Standard Time"
        log-Message "Mailbox updated" -colour "DarkYellow"
        }
    catch{
        log-Error "Failed to update MSOL mailbox"
        log-Error $Error
        }
    }
#endregion




$selectedRequests | % {
    $thisUser = $_
    provision-365user -userUPN $(remove-diacritics $($thisUser.FieldValues.Title.Trim().Replace(" ",".")+"@anthesisgroup.com")) `
        -userFirstName $thisUser.FieldValues.Title.Trim().Split(" ")[0].Trim() `
        -userSurname $($thisUser.FieldValues.Title.Trim().Split(" ")[$thisUser.FieldValues.Title.Trim().Split(" ").Count-1]).Trim() `
        -userDisplayName $($thisUser.FieldValues.Title).Trim() `
        -userManagerSAM $($thisUser.FieldValues.Line_x0020_Manager.Email).Replace("@anthesisgroup.com","") `
        -userPrimaryOffice $thisUser.FieldValues.Primary_x0020_Workplace.Label `
        -userCommunity $null `
        -userPrimaryTeam $thisUser.FieldValues.Primary_x0020_Team.Email `
        -userSecondaryTeams $thisUser.FieldValues.Additional_x0020_Teams `
        -userBusinessUnit $thisUser.FieldValues.Finance_x0020_Cost_x0020_Attribu.Label `
        -userJobTitle $thisUser.FieldValues.Job_x0020_title `
        -plaintextPassword "Anthesis123" `
        -adCredentials $adCredentials `
        -restCredentials $restCredentials `
        -newUserListItem $_ `
        -userTimeZone $thisUser.FieldValues.TimeZone `
        -user365License $thisUser.FieldValues.Office_x0020_365_x0020_license `
        -userSecondaryOffice $thisUser.FieldValues.Nearest_x0020_Office
    }
$selectedStarters  | % {
    provision-SustainADUser -userUPN $($_.Title.Trim().Replace(" ",".")+"@anthesisgroup.com") `
        -userFirstName $_.Title.Split(" ")[0] `
        -userSurname $($_.Title.Split(" ")[$_.Title.Split(" ").Count-1]) `
        -userDisplayName $($_.Title) `
        -userManagerSAM $_.Line_Manager `
        -userCommunity $null `
        -userPrimaryTeam $_.Primary_Team `
        -userSecondaryTeams $_.Additional_Teams `
        -userBusinessUnit $_.Finance_Cost_Attribu `
        -userJobTitle $_.Job_title `
        -plaintextPassword "Anthesis123" `
        -adCredentials $adCredentials `
        -restCredentials $restCredentials `
        -newUserListItem $_ `
        -userTimeZone $_.TimeZone `
        -user365License $_.Office_365_license `
        -userPrimaryOffice $_.Primary_Workplace
    }

$selectedStarters | % {
    update-msolUserFromAd -userUPN $($_.Title.Trim().Replace(" ",".")+"@anthesisgroup.com")
    }

#<# Bodge this stuff
    $userUPN = remove-diacritics $($thisUser.FieldValues.Title.Trim().Replace(" ",".")+"@anthesisgroup.com") 
    $userFirstName = $thisUser.FieldValues.Title.Split(" ")[0].Trim()
    $userSurname = $($thisUser.FieldValues.Title.Split(" ")[$thisUser.FieldValues.Title.Split(" ").Count-1]).Trim()
    $userDisplayName = $($thisUser.FieldValues.Title).Trim()
    $userManagerSAM = $($thisUser.FieldValues.Line_x0020_Manager.Email).Replace("@anthesisgroup.com","")
    $userCommunity = $null 
    $userPrimaryTeam = $thisUser.FieldValues.Primary_x0020_Team.Email
    $userSecondaryTeams = $thisUser.FieldValues.Additional_x0020_Teams.Email
    $userPrimaryOffice = $thisUser.FieldValues.Primary_x0020_Workplace.Label
    $userSecondaryOffice = $thisUser.FieldValues.Nearest_x0020_Office.Label
    $userBusinessUnit = $thisUser.FieldValues.Finance_x0020_Cost_x0020_Attribu.Label
    $userJobTitle = $thisUser.FieldValues.Job_x0020_title
    $plaintextPassword = "Welcome123" 
    $adCredentials = $adCredentials 
    $restCredentials = $restCredentials 
    $newUserListItem = $_ 
    $userTimeZone = $thisUser.FieldValues.TimeZone
    $user365License = $thisUser.FieldValues.Office_x0020_365_x0020_license
    $userDDI="0117 403 2XXX"
#>

