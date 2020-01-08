

#Done and tested :)
function create-ADUser{
    [cmdletbinding()]
Param (
    [parameter(Mandatory = $true,ParameterSetName="UPN")]
        [String]$upn
   ,[parameter(Mandatory = $true,ParameterSetName="UPN")]
        [String]$FirstName
    ,[parameter(Mandatory = $true,ParameterSetName="UPN")]
        [String]$Surname
    ,[parameter(Mandatory = $true,ParameterSetName="UPN")]
        [String]$DisplayName
    ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [String]$ManagerSAM
    ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [String]$PrimaryTeam
    ,[parameter(Mandatory = $true,ParameterSetName="UPN")]
        [String]$JobTitle
    ,[parameter(Mandatory = $true,ParameterSetName="UPN")]
        [String]$plaintextPassword
    ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [String]$BusinessUnit
    ,[parameter(Mandatory = $true,ParameterSetName="UPN")]
        [System.Management.Automation.PSCredential]$adCredentials
    ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
    [ValidateSet("Andorra, AND", "Barcelona, ESP", "Bogota, COL","Boulder, CO, USA","Bristol, GBR","Dubai, ARE","Emeryville, CA, USA","Frankfurt, DEU","Helsinki, FIN","London, GBR","Macclesfield, GBR","Madrid, ESP","Manchester, GBR","Manila, PHL","Manlleu, ESP","Nuremberg, DEU","Oxford, GBR","Rome, ITA","Stockholm, SWE","Tormarton, GBR")]
        [String[]]$PrimaryOffice
    ,[parameter(Mandatory = $true,ParameterSetName="UPN")]
        [String]$TwitterAccount
    ,[parameter(Mandatory = $true,ParameterSetName="UPN")]
        [String]$website
    ,[parameter(Mandatory = $true,ParameterSetName="UPN")]
        [String]$DDI
    ,[parameter(Mandatory = $true,ParameterSetName="UPN")]
        [String]$ouDn
    ,[parameter(Mandatory = $true,ParameterSetName="UPN")]
        [String]$upnSuffix
    ,[parameter(Mandatory = $true,ParameterSetName="UPN")]
        [String]$receptionDDI
    )

#Get BU details

switch ($BusinessUnit) {
    "Anthesis Energy UK Ltd (GBR)" {$upnSuffix = "@anthesisgroup.com"; $twitterAccount = "anthesis_group"; $DDI = "0117 403 2XXX"; $receptionDDI = "0117 403 2700";$ouDn = "OU=Users,OU=Sustain,DC=Sustainltd,DC=local"; $website = "www.anthesisgroup.com"}
    "Anthesis (UK) Ltd (GBR)"  {write-host -ForegroundColor Magenta "AUK, but creating Sustain account"; $upnSuffix = "@anthesisgroup.com"; $twitterAccount = "anthesis_group"; $DDI = "0117 403 2XXX"; $receptionDDI = "0117 403 2700";$ouDn = "OU=Users,OU=Sustain,DC=Sustainltd,DC=local"; $website = "www.anthesisgroup.com"}
    "Anthesis Consulting Group Ltd (GBR)" {}
    "Anthesis LLC" {}
    default {Write-Host -ForegroundColor DarkRed "Warning: Could not not identify Business Unit [$BusinessUnit]"}
    }
write-host "Business Unit is $($BusinessUnit)" -ForegroundColor White
#Create AD account

if(![string]::IsNullOrWhiteSpace($upn)){New-ADUser -Name $upn.Replace("."," ").Split("@")[0] `
    -AccountPassword (ConvertTo-SecureString $plaintextPassword -AsPlainText -force) `
    -CannotChangePassword $False `
    -ChangePasswordAtLogon $False `
    -Company $BusinessUnit `
    -Department $DisplayName `
    -Enabled $True `
    -Fax $twitterAccount `
    -GivenName $FirstName `
    -HomePage $website `
    -Manager $(Get-ADUser -Filter {SamAccountName -like $ManagerSAM}) `
    -OfficePhone $DDI `
    -Path $ouDn `
    -SAMAccountName $($upn.Split("@")[0]) `
    -Surname $Surname `
    -Title $JobTitle `
    -UserPrincipalName "$($upn.Split("@")[0])$upnSuffix" `
    -EmailAddress "$($upn.Split("@")[0])$upnSuffix" `
    -OtherAttributes @{'ipPhone'="XXX";'pager'=$receptionDDI} `
    -Credential $adCredentials
    
#Check the account was created and add the new account to a group if there is a PrimaryTeam specified.

$newAdUserAccount = Get-ADUser -filter {UserPrincipalName -like $upn} -Credential $adCredentials 
$primaryTeam = Get-ADGroup -Filter {name -like $PrimaryTeam} 
if($primaryTeam){
        Write-Host "Adding [$($newAdUserAccount.Name)] to [$($primaryTeam.Name)]"
        Add-ADGroupMember -Identity $primaryTeam.ObjectGUID -Members $newAdUserAccount -Credential $adCredentials
}
}
<#
.SYNOPSIS
Creates AD user object

.EXAMPLE
  create-ADUser -upn $upn -ManagerSAM $ManagerSAM -PrimaryTeam $PrimaryTeam -plaintextPassword $plaintextPassword -adCredentials $adCredentials -PrimaryOffice $PrimaryOffice -DDI $DDI -ouDn $ouDn -website $website -receptionDDI $receptionDDI -Fax $twitterAccount -JobTitle $JobTitle -upnSuffix $upnSuffix
#>

}


#done! :)
function create-msolUser{
        [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true,ParameterSetName="UPN")]
            [String]$upn
       ,[parameter(Mandatory = $true,ParameterSetName="UPN")]
            [String]$Plaintextpassword
            )

        Try{
        #create the Mailbox rather than the MSOLUser, which will effectively create an unlicensed E1 user
        if(![string]::IsNullOrWhiteSpace($upn)){New-Mailbox -Name $upn.Replace("."," ").Split("@")[0] -Password (ConvertTo-SecureString -AsPlainText $PlaintextPassword -Force) -MicrosoftOnlineServicesID $upn}
        }
        Catch{
        Write-Error "Failed to create new Msol user [$($upn)] in create-msoluser"
        Write-Error $_
        }

<#
.SYNOPSIS
Creates Msol User object by first creating a new mailbox, which will create an unlicensed E1 user.

.EXAMPLE
create-msolUser -upn "jo.bloggs@anthesisgroup.com" -Plaintextpassword $PlaintextPassword
#>
}


#done! :) Main licensing tested - issues with having $lO and remove license on same line as add license
function license-msolUser{
        [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true,ParameterSetName="UPN")]
            [String]$upn
       ,[parameter(Mandatory = $true,ParameterSetName="UPN")]
            [String]$licenseType
       ,[parameter(Mandatory = $true,ParameterSetName="UPN")]
            [String]$usageLocation
            )

#Core 365 licensing
    Try{
    switch ($licenseType){
        "E1" {
            $licenseToAssign = Get-MsolAccountSku | ?{$_.AccountSkuId -eq "AnthesisLLC:STANDARDPACK"}
            if((Get-MsolUser -UserPrincipalName $upn).Licenses.AccountSkuId -contains "AnthesisLLC:ENTERPRISEPACK"){$licenseToRemove = Get-MsolAccountSku | ?{$_.AccountSkuId -eq "AnthesisLLC:ENTERPRISEPACK"}}
            }
        "E3" {
            $licenseToAssign = Get-MsolAccountSku | ?{$_.AccountSkuId -eq "AnthesisLLC:ENTERPRISEPACK"}
            if((Get-MsolUser -UserPrincipalName $upn).Licenses.AccountSkuId -contains "AnthesisLLC:STANDARDPACK"){$licenseToRemove = Get-MsolAccountSku | ?{$_.AccountSkuId -eq "AnthesisLLC:STANDARDPACK"}}
            }
        }
        Write-Host -ForegroundColor Yellow "Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses $($licenseToAssign.AccountSkuId) -RemoveLicenses $($licenseToRemove.AccountSkuId)"
        $LO = New-MsolLicenseOptions -AccountSkuId "AnthesisLLC:ENTERPRISEPACK" -DisabledPlans "YAMMER_ENTERPRISE" #restrict Yammer
        Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses $($licenseToAssign.AccountSkuId)
        Set-MsolUserLicense -UserPrincipalName $upn -RemoveLicenses $($licenseToRemove.AccountSkuId) -LicenseOptions $LO          
        }
        Catch{
        Write-Error "Failed to license new Msol user [$($upn)] in license-msoluser (Core 365 Licensing)"
        Write-Error $_
        }
#Optional licensing    
    Try{
        If("GB" -eq $usageLocation){
        $licenseToAssign = Get-MsolAccountSku | ?{$_.AccountSkuId -eq "AnthesisLLC:EMS"}
        Set-MsolUserLicense -UserPrincipalName $UPN -AddLicenses $licenseToAssign.AccountSkuId
        }
        Else{
        write-host "I Shouldn't have an EMS license (yet) - Usage Location is $($usageLocation)"
        }
        }
        Catch{
        Write-Error "Failed to license new Msol user [$($upn)] in license-msoluser (EMS)"
        Write-Error $_
        }
<#
.SYNOPSIS
Licenses Msol user object with given licenses. The function breaks down the two types of licensning into 'Core Licensing' for E1 and E3 and 'Optional Licesnning' for licenses such as EMS which might not apply to the whole business. 

.EXAMPLE
license-msolUser -upn "jo.bloggs@anthesisgroup.com" -licenseType "E1" -usageLocation GB"
#>

}


#I've split this one into two functions for re-use - details vs groups as there may be more scope down the line with groups?
#Needs testing
function update-msoluserdetails{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true,ParameterSetName="UPN")]
        [PSObject]$upn
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [string]$firstName
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [string]$lastName
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [string]$displayName
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [string]$primaryTeam
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [ValidateSet("Andorra, AND", "Barcelona, ESP", "Bogota, COL","Boulder, CO, USA","Bristol, GBR","Dubai, ARE","Emeryville, CA, USA","Frankfurt, DEU","Helsinki, FIN","London, GBR","Macclesfield, GBR","Madrid, ESP","Manchester, GBR","Manila, PHL","Manlleu, ESP","Nuremberg, DEU","Oxford, GBR","Rome, ITA","Stockholm, SWE","Tormarton, GBR")]
        [string[]]$primaryOffice
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [string]$country
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [string]$jobTitle
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [string]$DDI
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [string]$mobile
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [string]$businessUnit
    )

#If key details aren't null, set them on the Msol user object
Try{
if(![string]::IsNullOrWhiteSpace($firstName)){Set-MsolUser -UserPrincipal $upn -FirstName $firstName}
if(![string]::IsNullOrWhiteSpace($lastName)){Set-MsolUser -UserPrincipal $upn -LastName $lastName}
if(![string]::IsNullOrWhiteSpace($displayName)){Set-MsolUser -UserPrincipal $upn -DisplayName $displayName}
if(![string]::IsNullOrWhiteSpace($country)){Set-MsolUser -UserPrincipal $upn -Country $country}
if(![string]::IsNullOrWhiteSpace($jobTitle)){Set-MsolUser -UserPrincipal $upn -title $jobTitle}
if(![string]::IsNullOrWhiteSpace($primaryOffice)){Set-MsolUser -UserPrincipal $upn -City $primaryOffice}
}
catch{
    Write-Error "Failed to update msoluser object [$($upn)] in update-msoluser"
    Write-Error $_
}
<#
.SYNOPSIS
Updates Msol User object with correct details, such as first name, last name, etc.
#>
}



#Needs testing
function update-msolusercoregroups{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true,ParameterSetName="UPN")]
        [PSObject]$upn
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [ValidateSet("Andorra, AND", "Barcelona, ESP", "Bogota, COL","Boulder, CO, USA","Bristol, GBR","Dubai, ARE","Emeryville, CA, USA","Frankfurt, DEU","Helsinki, FIN","London, GBR","Macclesfield, GBR","Madrid, ESP","Manchester, GBR","Manila, PHL","Manlleu, ESP","Nuremberg, DEU","Oxford, GBR","Rome, ITA","Stockholm, SWE","Tormarton, GBR")]
        [string[]]$primaryOffice
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [string]$businessUnit
    )

#If key details aren't null, let's add the msoluser to the correct regional group based on Office location (from Term Store)
Try{
    if(![string]::IsNullOrWhiteSpace($PrimaryOffice)){
    $OfficeTerm = Get-PnPTerm -Identity $PrimaryOffice -TermGroup "Anthesis" -TermSet "Offices" -Includes CustomProperties
    $365RegionalGroup = Get-DistributionGroup -Identity $OfficeTerm.CustomProperties.'365 Regional Group' | select-object -Property guid
    Add-DistributionGroupMember -Identity $365RegionalGroup -Member $upn -BypassSecurityGroupManagerCheck 
    }
    }
    catch{
        Write-Error "Failed to update msoluser group membership for regional group [$($upn)] in update-msoluser"
        Write-Error $_
    }
    #If they are in one of the GBR business units, add them to the MDM BYOD group
    try {
    if(![string]::IsNullOrWhiteSpace($businessUnit) -and (("Anthesis Energy UK Ltd (GBR)" -eq $businessUnit) -or ("Anthesis (UK) Ltd (GBR)" -eq $businessUnit) -or ("Anthesis Consulting Group Ltd (GBR)" -eq $businessUnit))){
    Add-DistributionGroupMember -Identity "b264f337-ef04-432e-a139-3574331a4d18" -Member $upn -BypassSecurityGroupManagerCheck
    }
    }
    catch {
        Write-Error "Failed to update msoluser group membership for MDM BYOD group [$($upn)] in update-msoluser"
        Write-Error $_
    }
<#
.SYNOPSIS
Updates Msol User object with correct core groups, such as regional and MDM groups.
#>
}





<#Kev's example#>
function update-msolMailboxViaUpn{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true,ParameterSetName="Mailbox")]
            [PSObject]$mailbox
        ,[parameter(Mandatory = $true,ParameterSetName="UPN")]
            [ValidatePattern("@anthesisgroup.com")]   # Need to amend ther functions with this function
            [string]$upn
        ,[parameter(Mandatory = $false,ParameterSetName="Mailbox")]
            [parameter(Mandatory = $false,ParameterSetName="UPN")]
            [string]$displayName
        ,[parameter(Mandatory = $false,ParameterSetName="Mailbox")]
            [parameter(Mandatory = $false,ParameterSetName="UPN")]
            [string]$businessUnit
        ,[parameter(Mandatory = $false,ParameterSetName="Mailbox")]
            [parameter(Mandatory = $false,ParameterSetName="UPN")]
            [string]$timeZone
         ,[parameter(Mandatory = $false,ParameterSetName="Mailbox")]
            [parameter(Mandatory = $false,ParameterSetName="UPN")]
            [ValidatePattern("@anthesisgroup.com")]
            [string]$lineManager
            ,[parameter(Mandatory = $false,ParameterSetName="Mailbox")]
        [parameter(Mandatory = $false,ParameterSetName="UPN")]
        [ValidateSet("Andorra, AND", "Barcelona, ESP", "Bogota, COL","Boulder, CO, USA","Bristol, GBR","Dubai, ARE","Emeryville, CA, USA","Frankfurt, DEU","Helsinki, FIN","London, GBR","Macclesfield, GBR","Madrid, ESP","Manchester, GBR","Manila, PHL","Manlleu, ESP","Nuremberg, DEU","Oxford, GBR","Rome, ITA","Stockholm, SWE","Tormarton, GBR")]
            [string[]]$Office
        )

    switch ($PsCmdlet.ParameterSetName){
        “Mailbox”  {$upn = $mailbox.UserPrincipalName}
        }

    Write-Verbose "update-msolMailbox($($upn),)"
    try{
        if(![string]::IsNullOrWhiteSpace($displayName)){Set-Mailbox -Identity $upn -DisplayName $displayName}
        if(![string]::IsNullOrWhiteSpace($businessUnit)){Set-Mailbox -Identity $upn -CustomAttribute1 $businessUnit}
        if(![string]::IsNullOrWhiteSpace($lineManager)){Set-User -Identity $upn -Manager $lineManager}
        }
    catch{
        Write-Error "Failed to set DisplayName or CustomAttribute1 on mailbox [$($upn)] in update-msolMailbox"
        Write-Error $_
        }
    if(![string]::IsNullOrWhiteSpace($Office)){
        #Get the correct timezone from term store by Office
        $OfficeTerm = Get-PnPTerm -Identity $($Office) -TermGroup "Anthesis" -TermSet "Offices" -Includes CustomProperties
        $timeZone = $($OfficeTerm.CustomProperties.Timezone)
        if(![string]::IsNullOrWhiteSpace($timeZone)){Set-MailboxRegionalConfiguration -Identity $upn -TimeZone $timeZone}
        }
    Set-Mailbox -Identity $upn -Alias $($upn.Split("@")[0]) -AuditEnabled $true -AuditLogAgeLimit 180 -AuditAdmin Update, MoveToDeletedItems, SoftDelete, HardDelete, SendAs, SendOnBehalf, Create, UpdateFolderPermission -AuditDelegate Update, SoftDelete, HardDelete, SendAs, Create, UpdateFolderPermissions, MoveToDeletedItems, SendOnBehalf -AuditOwner UpdateFolderPermission, MailboxLogin, Create, SoftDelete, HardDelete, Update, MoveToDeletedItems 
    }
    

<#Kev's example/#>


#Needs testing
function update-sharePointConfig{
        [cmdletbinding()]
    Param (
         [parameter(Mandatory = $true,ParameterSetName="UPN")]
            [String]$upn
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
            [String]$timezoneID
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
            [String]$countryLocale
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
            [String]$languageCode
            )
        

        #Main provision user script will need to figure out the p3letter country code from the Main Office field, which is term store based. we can assume this is where (or at least close to where) they will be working.

       if(![string]::IsNullOrWhiteSpace($upn)){
        
        Try{
            Write-Host "Setting Sharepoint initial config" -ForegroundColor Yellow
            Set-PnPUserProfileProperty -Account $UPN -PropertyName 'SPS-RegionalSettings-FollowWeb' -Value "False"
            Set-PnPUserProfileProperty -Account $UPN -PropertyName 'SPS-RegionalSettings-Initialized' -Value "True"
            Set-PnPUserProfileProperty -Account $UPN -PropertyName 'SPS-TimeZone' -Value $($timezoneID)
            Set-PnPUserProfileProperty -Account $UPN -PropertyName 'SPS-Locale' -Value $($countryLocale)
            Set-PnPUserProfileProperty -Account $UPN -PropertyName 'SPS-MUILanguages' -Value $($languageCode)
            Set-PnPUserProfileProperty -Account $UPN -PropertyName 'SPS-CalendarType' -Value "1"
            Set-PnPUserProfileProperty -Account $UPN -PropertyName 'SPS-AltCalendarType' -Value "1"
            }
       Catch{
            Write-Error "Failed to update SPO user [$($upn)] in update-sharePointInitialConfig"
            Write-Error $_
       }
       }

}






<#Still to do
The three letter bu string is a question, as is the manager access?
#>
function set-mailboxPermissions{
        [cmdletbinding()]
    Param (
         [parameter(Mandatory = $true,ParameterSetName="UPN")]
            [String]$upn
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
            [String]$ManagerSAM
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
            [String]$BusinessUnit
            )

      if(![string]::IsNullOrWhiteSpace($upn)){

      Try{
        Add-MailboxPermission -Identity $UPN -AccessRights FullAccess -User $ManagerSAM -InheritanceType all -AutoMapping $false
        Add-MailboxFolderPermission "$($UPN):\Calendar" -User "All$(get-3lettersInBrackets -stringMaybeContaining3LettersInBrackets $BusinessUnit)@anthesisgroup.com" -AccessRights "LimitedDetails"
         }
      Catch{
        Write-Error "Failed to update SPO user [$($upn)] in update-sharePointInitialConfig"
        Write-Error $_
      }
}
}











