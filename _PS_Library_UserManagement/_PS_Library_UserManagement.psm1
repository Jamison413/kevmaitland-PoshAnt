function create-ADUser{
    [cmdletbinding()]
Param (
    [parameter(Mandatory = $true,ParameterSetName="UPN")]
        [String]$upn
   ,[parameter(Mandatory = $true,ParameterSetName="UPN")]
        [String]$firstname
    ,[parameter(Mandatory = $true,ParameterSetName="UPN")]
        [String]$surname
    ,[parameter(Mandatory = $true,ParameterSetName="UPN")]
        [String]$displayname
    ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [String]$managerSAM
    ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [String]$primaryteam
    ,[parameter(Mandatory = $true,ParameterSetName="UPN")]
        [String]$jobtitle
    ,[parameter(Mandatory = $true,ParameterSetName="UPN")]
        [String]$plaintextpassword
    ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [String]$businessunit
    ,[parameter(Mandatory = $true,ParameterSetName="UPN")]
        [System.Management.Automation.PSCredential]$adCredentials
    ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
    [ValidateSet("Andorra, AND", "Barcelona, ESP", "Bogota, COL","Boulder, CO, USA","Bristol, GBR","Dubai, ARE","Emeryville, CA, USA","Frankfurt, DEU","Helsinki, FIN","London, GBR","Macclesfield, GBR","Madrid, ESP","Manchester, GBR","Manila, PHL","Manlleu, ESP","Nuremberg, DEU","Oxford, GBR","Rome, ITA","Stockholm, SWE","Tormarton, GBR")]
        [String[]]$office
    ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
    [String]$allpermanentstaffadgroupprompt
    ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
    [String]$SAMaccountname
    )

#Get BU details
switch ($businessunit) {
    "Anthesis Energy UK Ltd (GBR)" {$upnsuffix = "@anthesisgroup.com"; $twitteraccount = "anthesis_group"; $DDI = "0117 403 2XXX"; $receptionDDI = "0117 403 2700";$ouDn = "OU=Users,OU=Sustain,DC=Sustainltd,DC=local"; $website = "www.anthesisgroup.com"}
    "Anthesis (UK) Ltd (GBR)"  {$upnsuffix = "@anthesisgroup.com"; $twitteraccount = "anthesis_group"; $DDI = "0117 403 2XXX"; $receptionDDI = "0117 403 2700";$ouDn = "OU=Users,OU=Sustain,DC=Sustainltd,DC=local"; $website = "www.anthesisgroup.com"}
    "Anthesis Consulting Group Ltd (GBR)" {}
    "Anthesis LLC" {}
    default {Write-Host -ForegroundColor DarkRed "Warning: Could not not identify Business Unit [$businessunit]"}
    }
write-host "Business Unit is $($businessunit)" -ForegroundColor White

#Create AD account
write-host "*******************This is what we will try to set:*******************" -ForegroundColor White
write-host "ad upn: $($upn)" -ForegroundColor Yellow
write-host "firstname: $($firstname)" -ForegroundColor Yellow
write-host "lastname: $($lastname)" -ForegroundColor Yellow
write-host "displayname: $($displayname)" -ForegroundColor Yellow
write-host "jobtitle: $($jobtitle)" -ForegroundColor Yellow
write-host "businessunit: $($businessunit)" -ForegroundColor Yellow
write-host "department: $($primaryteam)" -ForegroundColor Yellow
write-host "managerSAM: $($managerSAM)" -ForegroundColor Yellow
write-host "**********************************************************************" -ForegroundColor White

if(![string]::IsNullOrWhiteSpace($upn)){New-ADUser -Name $upn.Replace("."," ").Split("@")[0] `
    -AccountPassword (ConvertTo-SecureString $plaintextpassword -AsPlainText -force) `
    -CannotChangePassword $False `
    -ChangePasswordAtLogon $False `
    -Company $businessunit `
    -DisplayName $displayname `
    -Department  $primaryteam `
    -Enabled $True `
    -Fax $twitteraccount `
    -GivenName $firstname `
    -HomePage $website `
    -OfficePhone $DDI `
    -Path $ouDn `
    -SAMAccountName $SAMAccountName `
    -Surname $surname `
    -Title $jobtitle `
    -UserPrincipalName "$($upn.Split("@")[0])$upnsuffix" `
    -EmailAddress "$($upn.Split("@")[0])$upnsuffix" `
    -OtherAttributes @{'ipPhone'="XXX";'pager'=$receptionDDI} `
    -Credential $adCredentials
    
#Check the account was created and add the new account to a group if there is a primaryteam specified.
$newAdUserAccount = Get-ADUser -filter {SamAccountName -eq $SAMaccountname} -Credential $adCredentials 
Write-Host "Looks like the AD account for $($upn) was created successfully!" -ForegroundColor Green

<#--------Add to the primary team--------#>
#$primaryteam = Get-ADGroup -Filter {name -like $primaryteam} 
#if($primaryteam){
 #       Write-Host "Adding [$($newAdUserAccount.Name)] to [$($primaryteam.Name)]"
 #       Add-ADGroupMember -Identity $primaryteam.ObjectGUID -Members $newAdUserAccount -Credential $adCredentials
#}

<#--------Set Manager field (if the Manager has an AD account)--------#>
$manager = (Get-ADUser -Filter {SamAccountName -eq $managerSAM} -Credential $adcredentials)
If($manager){
    Set-ADUser -Identity $SAMaccountname -Manager $managerSAM -Credential $adcredentials
}
Else{
write-host "The Line Manager doesn't appear to have an account on our domain, skipping..." -ForegroundColor White
}

Write-Host "Adding to the relevant AD groups..." -ForegroundColor Yellow
<#--------Add to relevant AD groups--------#>
#Prompt to add to all permanent staff
If("y" -eq $allpermanentstaffadgroupprompt){
Add-ADGroupMember -Identity "All Permanent Staff" -members $SAMaccountname -Credential $adCredentials
}
Else{
write-host "Okay, we won't add $($upn) to the All Permanent Staff group" -ForegroundColor White
}
#If London based, add to Taper Building AD group
If(("London, GBR" -eq $office) -and ("y" -eq $allpermanentstaffadgroupprompt)){
    Write-Host "We'll add the new starter to the Taper Building AD group...." -ForegroundColor White
    Add-ADGroupMember -Identity "The Taper Building" -members $SAMaccountname -Credential $adCredentials
    Write-Host "...and the AlwaysOn VPN Users group"
    Add-ADGroupMember -Identity "AlwaysOn VPN Users" -members $SAMaccountname -Credential $adCredentials
    }
    Else{
    }
#Add to AlwaysOn VPN AD Group if Bristol based
If(("Bristol, GBR" -eq $office) -and ("y" -eq $allpermanentstaffadgroupprompt)){
    Write-Host "We'll add them to the AlwaysOn VPN Users group"
    Add-ADGroupMember -Identity "AlwaysOn VPN Users" -members $SAMaccountname -Credential $adCredentials
    }
    Else{
    }          
}
}
<#
.SYNOPSIS
Creates AD user object

.EXAMPLE
  create-ADUser -upn $upn -managerSAM $managerSAM -primaryteam $primaryteam -plaintextpassword $plaintextpassword -adCredentials $adCredentials -office $office -DDI $DDI -ouDn $ouDn -website $website -receptionDDI $receptionDDI -Fax $twitteraccount -jobtitle $jobtitle -upnsuffix $upnsuffix
#>
function create-msolUser{
        [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true,ParameterSetName="UPN")]
            [String]$upn
       ,[parameter(Mandatory = $true,ParameterSetName="UPN")]
            [String]$plaintextpassword
            )

Try{
        #create the Mailbox rather than the MSOLUser, which will effectively create an unlicensed E1 user
        if(![string]::IsNullOrWhiteSpace($upn)){New-Mailbox -Name $upn.Replace("."," ").Split("@")[0] -Password (ConvertTo-SecureString -AsPlainText $plaintextpassword -Force) -MicrosoftOnlineServicesID $upn}
        }
        Catch{
        Write-Error "Failed to create new Msol user [$($upn)] in create-msoluser"
        Write-Error $_
        }
    }
<#
.SYNOPSIS
Creates Msol User object by first creating a new mailbox, which will create an unlicensed E1 user.

.EXAMPLE
create-msolUser -upn "jo.bloggs@anthesisgroup.com" -plaintextpassword $plaintextpassword
#>
function license-msolUser{
        [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true,ParameterSetName="UPN")]
            [String]$upn
       ,[parameter(Mandatory = $true,ParameterSetName="UPN")]
            [String]$licensetype
       #,[parameter(Mandatory = $true,ParameterSetName="UPN")]
            #[String]$usagelocation
            )

#Core 365 licensing
    Try{
    switch ($licensetype){
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
        write-host "Adding licenses: $($licenseToAssign.AccountSkuId)" -ForegroundColor Yellow
        Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses $($licenseToAssign.AccountSkuId)
        write-host "Removing licenses: Yammer" -ForegroundColor Yellow
        Set-MsolUserLicense -UserPrincipalName $upn -RemoveLicenses $($licenseToRemove.AccountSkuId) -LicenseOptions $LO     
        }
        Catch{
        Write-Error "Failed to license new Msol user [$($upn)] in license-msoluser (Core 365 Licensing)"
        Write-Error $_
        }
#Optional licensing    
    Try{
        If("GB" -eq $usagelocation){
        $licenseToAssign = Get-MsolAccountSku | ?{$_.AccountSkuId -eq "AnthesisLLC:EMS"}
        write-host "****************Adding EMS license****************"
        Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses $licenseToAssign.AccountSkuId
        }
        Else{
        write-host "I Shouldn't have an EMS license (yet) - Usage Location is $($usagelocation)"
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
license-msolUser -upn "jo.bloggs@anthesisgroup.com" -licensetype "E1" -usagelocation GB"
#>

}
function update-msoluserdetails{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true,ParameterSetName="UPN")]
        [PSObject]$upn
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [string]$firstname
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [string]$lastname
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [string]$displayname
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [string]$primaryteam
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [string]$country
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [string]$streetaddress
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [string]$office
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [string]$city
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [string]$usagelocation
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [string]$postcode
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [string]$jobtitle
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [string]$DDI
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [string]$mobile
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [string]$businessunit
    )

#If key details aren't null, set them on the Msol user object
Try{
    write-host "Setting Firstname user: $($upn): $($firstname)" -ForegroundColor Yellow
    if(![string]::IsNullOrWhiteSpace($firstname)){Set-MsolUser -UserPrincipal $upn -firstname $firstname}
}
Catch{
    Write-Error "Failed to update msoluser object firstname [$($upn)] in update-msoluser"
} 
Try{
    write-host "Setting lastname user: $($upn): $($lastname)" -ForegroundColor Yellow
    if(![string]::IsNullOrWhiteSpace($lastname)){Set-MsolUser -UserPrincipal $upn -lastname $lastname}
}
Catch{
    Write-Error "Failed to update msoluser object lastname [$($upn)] in update-msoluser"
} 
Try{
    write-host "Setting displayname user: $($upn): $($displayname)" -ForegroundColor Yellow
    if(![string]::IsNullOrWhiteSpace($displayname)){Set-MsolUser -UserPrincipal $upn -displayname $displayname}
}
Catch{
    Write-Error "Failed to update msoluser object displayname [$($upn)] in update-msoluser"
} 
Try{
    write-host "Setting country user: $($upn): $($country)" -ForegroundColor Yellow
    if(![string]::IsNullOrWhiteSpace($country)){Set-MsolUser -UserPrincipal $upn -Country $country}
}
Catch{
    Write-Error "Failed to update msoluser object country [$($upn)] in update-msoluser"
} 
Try{
    write-host "Setting jobtitle user: $($upn): $($jobtitle)" -ForegroundColor Yellow
    if(![string]::IsNullOrWhiteSpace($jobtitle)){Set-MsolUser -UserPrincipal $upn -title $jobtitle}
}
Catch{
    Write-Error "Failed to update msoluser object jobtitle [$($upn)] in update-msoluser"
} 
Try{
    write-host "Setting city user: $($upn): $($city)" -ForegroundColor Yellow
    if(![string]::IsNullOrWhiteSpace($city)){Set-MsolUser -UserPrincipal $upn -City $city}
}
Catch{
    Write-Error "Failed to update msoluser object city [$($upn)] in update-msoluser"
} 
Try{
    write-host "Setting usagelocation user: $($upn): $($usagelocation)" -ForegroundColor Yellow
    if(![string]::IsNullOrWhiteSpace($city)){Set-MsolUser -UserPrincipal $upn -UsageLocation $usagelocation}
}
Catch{
    Write-Error "Failed to update msoluser object city [$($upn)] in update-msoluser"
} 
Try{
    write-host "Setting streetaddress user: $($upn): $($streetaddress)" -ForegroundColor Yellow
    if(![string]::IsNullOrWhiteSpace($streetaddress)){Set-MsolUser -UserPrincipal $upn -StreetAddress $streetaddress}
}
Catch{
    Write-Error "Failed to update msoluser object streetaddress [$($upn)] in update-msoluser"
} 
Try{
    write-host "Setting postcode user: $($upn): $($postcode)" -ForegroundColor Yellow
    if(![string]::IsNullOrWhiteSpace($postcode)){Set-MsolUser -UserPrincipal $upn -PostalCode $postcode}
}
Catch{
    Write-Error "Failed to update msoluser object postcode [$($upn)] in update-msoluser"
} 
Try{
    write-host "Setting office user: $($upn): $($office)" -ForegroundColor Yellow
    if(![string]::IsNullOrWhiteSpace($office)){Set-MsolUser -UserPrincipal $upn -Office $office}
}
Catch{
    Write-Error "Failed to update msoluser object office [$($upn)] in update-msoluser"
} 
}
<#
.SYNOPSIS
Updates Msol User object with correct details, such as first name, last name, etc.
#>
function update-msolusercoregroups{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true,ParameterSetName="UPN")]
        [PSObject]$upn
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [ValidateSet("Andorra, AND", "Barcelona, ESP", "Bogota, COL","Boulder, CO, USA","Bristol, GBR","Dubai, ARE","Emeryville, CA, USA","Frankfurt, DEU","Helsinki, FIN","London, GBR","Macclesfield, GBR","Madrid, ESP","Manchester, GBR","Manila, PHL","Manlleu, ESP","Nuremberg, DEU","Oxford, GBR","Rome, ITA","Stockholm, SWE","Tormarton, GBR")]
        [string[]]$office
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [string]$businessunit
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
        [string]$regionalgroup
    )

#If key details aren't null, let's add the msoluser to the correct regional group based on Office location (from Term Store)
Try{
    if(![string]::IsNullOrWhiteSpace($office)){
        write-host "Adding to regionalgroup: $($office)"
    Add-DistributionGroupMember -Identity $regionalgroup -Member $upn -BypassSecurityGroupManagerCheck 
    }
    }
    catch{
        Write-Error "Failed to update msoluser group membership for regional group [$($upn)] in update-msoluser"
        Write-Error $_
    }
    #If they are in one of the GBR business units, add them to the MDM BYOD group
    try {
    if(![string]::IsNullOrWhiteSpace($businessunit) -and (("Anthesis Energy UK Ltd (GBR)" -eq $businessunit) -or ("Anthesis (UK) Ltd (GBR)" -eq $businessunit) -or ("Anthesis Consulting Group Ltd (GBR)" -eq $businessunit))){
    write-host "Adding to MDM BYOD Group"
        Add-DistributionGroupMember -Identity "b264f337-ef04-432e-a139-3574331a4d18" -Member $upn -BypassSecurityGroupManagerCheck
    }
    }
    catch {
        Write-Error "Failed to update msoluser group membership for MDM BYOD group [$($upn)] in update-msoluser"
        Write-Error $_
    }
    try {
        if(![string]::IsNullOrWhiteSpace($businessunit) -and (("Anthesis Energy UK Ltd (GBR)" -eq $businessunit) -or ("Anthesis (UK) Ltd (GBR)" -eq $businessunit) -or ("Anthesis Consulting Group Ltd (GBR)" -eq $businessunit))){
        write-host "Adding to All Sharepoint Users"
        Add-DistributionGroupMember -Identity "f30dfb2c-88d4-4e4d-8953-3d68f0d92a9e" -Member $upn -BypassSecurityGroupManagerCheck
        }
        }
        catch {
            Write-Error "Failed to update msoluser group membership for All Sharepoint Users [$($upn)] in update-msoluser"
            Write-Error $_
        }
    
<#
.SYNOPSIS
Updates Msol User object with correct core groups, such as regional and MDM groups.
#>
}
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
            [string]$displayname
        ,[parameter(Mandatory = $false,ParameterSetName="Mailbox")]
            [parameter(Mandatory = $false,ParameterSetName="UPN")]
            [string]$businessunit
        ,[parameter(Mandatory = $false,ParameterSetName="Mailbox")]
            [parameter(Mandatory = $false,ParameterSetName="UPN")]
            [string]$timezone
         ,[parameter(Mandatory = $false,ParameterSetName="Mailbox")]
            [parameter(Mandatory = $false,ParameterSetName="UPN")]
            [ValidatePattern("@anthesisgroup.com")]
            [string]$linemanager
            ,[parameter(Mandatory = $false,ParameterSetName="Mailbox")]
        [parameter(Mandatory = $false,ParameterSetName="UPN")]
        [ValidateSet("Andorra, AND", "Barcelona, ESP", "Bogota, COL","Boulder, CO, USA","Bristol, GBR","Dubai, ARE","Emeryville, CA, USA","Frankfurt, DEU","Helsinki, FIN","London, GBR","Macclesfield, GBR","Madrid, ESP","Manchester, GBR","Manila, PHL","Manlleu, ESP","Nuremberg, DEU","Oxford, GBR","Rome, ITA","Stockholm, SWE","Tormarton, GBR")]
            [string[]]$office
        )

    switch ($PsCmdlet.ParameterSetName){
        “Mailbox”  {$upn = $mailbox.UserPrincipalName}
        }

    Write-Verbose "update-msolMailbox($($upn),)"
    try{
        if(![string]::IsNullOrWhiteSpace($displayname)){Set-Mailbox -Identity $upn -displayname $displayname}
        if(![string]::IsNullOrWhiteSpace($businessunit)){Set-Mailbox -Identity $upn -CustomAttribute1 $businessunit}
        if(![string]::IsNullOrWhiteSpace($linemanager)){Set-User -Identity $upn -Manager $linemanager}
        if(![string]::IsNullOrWhiteSpace($timezone)){Set-MailboxRegionalConfiguration -Identity $upn -timezone $timezone}
        }
    catch{
        Write-Error "Failed to set displayname or CustomAttribute1 on mailbox [$($upn)] in update-msolMailbox"
        Write-Error $_
        }
    try{
    Set-Mailbox -Identity $upn -Alias $($upn.Split("@")[0]) -AuditEnabled $true -AuditLogAgeLimit 180 -AuditAdmin Update, MoveToDeletedItems, SoftDelete, HardDelete, SendAs, SendOnBehalf, Create, UpdateFolderPermission -AuditDelegate Update, SoftDelete, HardDelete, SendAs, Create, UpdateFolderPermissions, MoveToDeletedItems, SendOnBehalf -AuditOwner UpdateFolderPermission, MailboxLogin, Create, SoftDelete, HardDelete, Update, MoveToDeletedItems 
    }
    catch{
        Write-Error "Failed to set audit info [$($upn)] in update-msolMailbox"
        Write-Error $_
    }
}    
function update-sharePointConfig{
    [cmdletbinding()]
Param (
     [parameter(Mandatory = $true,ParameterSetName="upn")]
        [String]$upn
    ,[parameter(Mandatory = $false,ParameterSetName="upn")]
        [String]$office
        )
    
   if(![string]::IsNullOrWhiteSpace($upn)){
        #Check if there is an SPO profile with this upn
        $profilename = ("i:0#.f|membership|" + "$($upn)").Trim()
        Write-Host "$($profilename)" -ForegroundColor Yellow
        $SPOUserProfile = Get-PnPUser -Identity $profilename
        If($SPOUserProfile){write-host "Success! SPOUserProfile retrieved for $($upn)"}
        Else{write-host "Failure! SPOUserProfile could not be retrieved for $($upn)"
        break}
        #Then see if the profile follows default profile settings via "follow-web"
        $SPOUserProfileProperties = Get-PnPUserProfileProperty -Account $upn
        If("True" -eq $($SPOUserProfileProperties.UserProfileProperties.'SPS-RegionalSettings-FollowWeb')){write-host "It looks like they are using the default settings!"}
        Else{
        write-host "It looks like they have unique settings already, meaning they are following their current offices timezone and locale settings" -ForegroundColor Yellow
        If("True" -eq ([Environment]::UserInteractive)){
            $syncconfiguration = Read-host "Do you want to re-sync their Sharepoint configuration? (y/n)"
                If("y" -eq $syncconfiguration){write-host "Okay! Attempting to re-sync Sharepoint configuration"}
                Else {break}
            }
            Else{
                Write-Host "You look like a robot, hello! We'll continue syncing Sharepoint Config" -ForegroundColor Yellow
            }
        }
        #If we've gotten this far, the profile has default settings or we are interactively telling the script to update
        If(("True" -eq $($SPOUserProfileProperties.UserProfileProperties.'SPS-RegionalSettings-FollowWeb') -and ("True" -ne ([Environment]::UserInteractive)))){
            Write-Host "It looks like they are using the default settings and you are a robot, continuing...." -ForegroundColor Yellow
        }
        Else{
            Write-Host "It looks like they have unique settings already, meaning they are following their current offices timezone and locale settings (unless this has changed in 365)...stopping" -ForegroundColor Yellow
            break
        }
        #If office is missing, get the MSOL User object and office from there, Get secondary geographic information for office from term store
        if([string]::IsNullOrWhiteSpace($office)){
        write-host "It looks like an office wasn't provided, so we'll try to retrieve it from 365" -ForegroundColor Yellow 
       $office = Get-MsolUser -UserPrincipalName $upn | select-object -Property "Office"
       $termtofind = ($office.split("=").Replace("}",""))[1]
        Write-Host "I'm in $($termtofind) according to 365"
        $officeterm = Get-PnPTerm -Identity $($termtofind) -TermGroup "Anthesis" -TermSet "offices" -Includes CustomProperties
        Write-Host "Here is the term I tried to get: $($officeterm.Name)" -ForegroundColor Yellow
        #Set variables
        $timezoneID = $($officeterm.CustomProperties.'Sharepoint Timezone ID')
        $countrylocale = $($officeterm.CustomProperties.'Locale')
        $languagecode = $($officeterm.CustomProperties.'Language Code')
        }
        Else{
        Write-host "It looks like an office was provided! Retrieving term from term store"
        $officeterm = Get-PnPTerm -Identity $($office) -TermGroup "Anthesis" -TermSet "offices" -Includes CustomProperties
        $timezoneID = $($officeterm.CustomProperties.'Sharepoint Timezone ID')
        $countrylocale = $($officeterm.CustomProperties.'Locale')
        $languagecode = $($officeterm.CustomProperties.'Language Code')
        Write-host "Warning: We'll update the users timezone to match the provided office, however this will be out of sync with the office record for the user in 365 if the provided office is different - you'll need to change it to match if required " -ForegroundColor Red
        }
        Write-Host "Setting Sharepoint configuration for $($upn) in $($termtofind)" -ForegroundColor Yellow
        Try{
        Set-PnPUserProfileProperty -Account $upn -PropertyName 'SPS-RegionalSettings-FollowWeb' -Value "False"
        Set-PnPUserProfileProperty -Account $upn -PropertyName 'SPS-RegionalSettings-Initialized' -Value "True"
        }
        Catch{
            Write-Error "Failed to update SPO user [$($upn)] in update-sharePointConfig: RegionalSettings"
            Write-Error $_
        }
        if(![string]::IsNullOrWhiteSpace($timezoneID)){
        Try{
        Set-PnPUserProfileProperty -Account $upn -PropertyName 'SPS-timezone' -Value $($timezoneID)
        }
        Catch{
            Write-Error "Failed to update SPO user [$($upn)] in update-sharePointConfig: TimezoneID"
            Write-Error $_
        }
        }
        if(![string]::IsNullOrWhiteSpace($countrylocale)){
        Try{
        if(![string]::IsNullOrWhiteSpace($countrylocale)){Set-PnPUserProfileProperty -Account $upn -PropertyName 'SPS-Locale' -Value $($countrylocale)}
        }
        Catch{
        Write-Error "Failed to update SPO user [$($upn)] in update-sharePointConfig: Country Locale"
        Write-Error $_
        }
        }
        if(![string]::IsNullOrWhiteSpace($languagecode)){
        Try{
        if(![string]::IsNullOrWhiteSpace($languagecode)){Set-PnPUserProfileProperty -Account $upn -PropertyName 'SPS-MUILanguages' -Value $($languagecode)}
        }
        Catch{
        Write-Error "Failed to update SPO user [$($upn)] in update-sharePointConfig: languagecode"
        Write-Error $_
        }
        }
        Try{
        Set-PnPUserProfileProperty -Account $upn -PropertyName 'SPS-CalendarType' -Value "1"
        Set-PnPUserProfileProperty -Account $upn -PropertyName 'SPS-AltCalendarType' -Value "1"
        }
        Catch{
            Write-Error "Failed to update SPO user [$($upn)] in update-sharePointConfig: CalendarType"
            Write-Error $_
        }

}
}
<#
.SYNOPSIS
Updates SPO User profile with correct details according to 365 msol office details or provided office. Must be connected via Kimblebot for automation and to the main "https://anthesisllc.sharepoint.com/" site (NOT the admin site)
#>
<#Still to do
The three letter bu string is a question, as is the manager access?
#>
function set-mailboxPermissions{
        [cmdletbinding()]
    Param (
         [parameter(Mandatory = $true,ParameterSetName="UPN")]
            [String]$upn
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
            [String]$managerSAM
        ,[parameter(Mandatory = $false,ParameterSetName="UPN")]
            [String]$businessunit
            )

      if(![string]::IsNullOrWhiteSpace($upn)){

      Try{
        Add-MailboxPermission -Identity $UPN -AccessRights FullAccess -User $managerSAM -InheritanceType all -AutoMapping $false
        Add-MailboxFolderPermission "$($UPN):\Calendar" -User "All$(get-3lettersInBrackets -stringMaybeContaining3LettersInBrackets $businessunit)@anthesisgroup.com" -AccessRights "LimitedDetails"
         }
      Catch{
        Write-Error "Failed to update SPO user [$($upn)] in update-sharePointInitialConfig"
        Write-Error $_
      }
}
}
function remove-mailboxesandbots{
        [cmdletbinding()]
    Param (
         [parameter(Mandatory = $true)]
            [array]$usersarray
            )

#Remove licensed mailbox accounts and bots

$usersarray = $usersarray | Where-Object -Property "userPrincipalName" -NE "conflictminerals@anthesisgroup.com"
$usersarray = $usersarray | Where-Object -Property "userPrincipalName" -NE "VarexConflictMinerals@anthesisgroup.com"
$usersarray = $usersarray | Where-Object -Property "userPrincipalName" -NE "ACSSupport@anthesisgroup.com"
$usersarray = $usersarray | Where-Object -Property "userPrincipalName" -NE "acsmailboxaccess@anthesisgroup.com"
$usersarray = $usersarray | Where-Object -Property "userPrincipalName" -NE "Microsoft.ECM@anthesisgroup.com"
$usersarray = $usersarray | Where-Object -Property "userPrincipalName" -NE "qwest_ga@anthesisgroup.com"
$usersarray = $usersarray | Where-Object -Property "userPrincipalName" -NE "info@umr-gmbh.com"
$usersarray = $usersarray | Where-Object -Property "userPrincipalName" -NE "Anthesis Energy UK Mailbox Robot"
$usersarray = $usersarray | Where-Object -Property "userPrincipalName" -NE "Varex.PEC@anthesisgroup.com"
$usersarray = $usersarray | Where-Object -Property "userPrincipalName" -NE "UKcareers@anthesisgroup.com"
$usersarray = $usersarray | Where-Object -Property "userPrincipalName" -NE "Diana.Correal@anthesisgroup.com"
$usersarray = $usersarray | Where-Object -Property "userPrincipalName" -NE "groupbot@anthesisgroup.com"
$usersarray = $usersarray | Where-Object -Property "userPrincipalName" -NE "SustainMailboxAccess@anthesisgroup.com"
$usersarray = $usersarray | Where-Object -Property "userPrincipalName" -NE "barry.holt@anthesisgroup.com"
$usersarray = $usersarray | Where-Object -Property "userPrincipalName" -NE "AnthesisUKFinance@anthesisgroup.com"



$usersarray
}
function set-SPOTimezone{
    [cmdletbinding()]
Param (
     [parameter(Mandatory = $true,ParameterSetName="upn")]
        [String]$upn
    ,[parameter(Mandatory = $true,ParameterSetName="upn")]
        [String]$office
        )
    
   if(![string]::IsNullOrWhiteSpace($upn) -and ![string]::IsNullOrWhiteSpace($office)){
        $officeterm = Get-PnPTerm -Identity $($office) -TermGroup "Anthesis" -TermSet "offices" -Includes CustomProperties
        Write-Host "$($officeterm.Name)"
        Write-Host "$($upn)"
        $timezoneID = $($officeterm.CustomProperties.'Sharepoint Timezone ID')
        $countrylocale = $($officeterm.CustomProperties.'Locale')
        $languagecode = $($officeterm.CustomProperties.'Language Code')
        Try{
        Set-PnPUserProfileProperty -Account $upn -PropertyName 'SPS-RegionalSettings-FollowWeb' -Value "False"
        Set-PnPUserProfileProperty -Account $upn -PropertyName 'SPS-RegionalSettings-Initialized' -Value "True"
        }
        Catch{
            Write-Error "Failed to update SPO user [$($upn)] in update-sharePointConfig: RegionalSettings"
            Write-Error $_
        }
        if(![string]::IsNullOrWhiteSpace($timezoneID)){
        Try{
        Set-PnPUserProfileProperty -Account $upn -PropertyName 'SPS-timezone' -Value $($timezoneID)
        }
        Catch{
            Write-Error "Failed to update SPO user [$($upn)] in update-sharePointConfig: TimezoneID"
            Write-Error $_
        }
        }
        if(![string]::IsNullOrWhiteSpace($countrylocale)){
        Try{
        if(![string]::IsNullOrWhiteSpace($countrylocale)){Set-PnPUserProfileProperty -Account $upn -PropertyName 'SPS-Locale' -Value $($countrylocale)}
        }
        Catch{
        Write-Error "Failed to update SPO user [$($upn)] in update-sharePointConfig: Country Locale"
        Write-Error $_
        }
        }
        if(![string]::IsNullOrWhiteSpace($languagecode)){
        Try{
        if(![string]::IsNullOrWhiteSpace($languagecode)){Set-PnPUserProfileProperty -Account $upn -PropertyName 'SPS-MUILanguages' -Value $($languagecode)}
        }
        Catch{
        Write-Error "Failed to update SPO user [$($upn)] in update-sharePointConfig: languagecode"
        Write-Error $_
        }
        }
        Try{
        Set-PnPUserProfileProperty -Account $upn -PropertyName 'SPS-CalendarType' -Value "1"
        Set-PnPUserProfileProperty -Account $upn -PropertyName 'SPS-AltCalendarType' -Value "1"
        }
        Catch{
            Write-Error "Failed to update SPO user [$($upn)] in update-sharePointConfig: CalendarType"
            Write-Error $_
        }

}
}



