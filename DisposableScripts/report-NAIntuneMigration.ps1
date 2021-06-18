#Get Users from AAD
$tokenResponseTeamsBot = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName TeamsBot) -grant_type client_credentials
#$allUsers = get-graphUsersWithEmployeeInfoExtensions -tokenResponse $tokenResponseTeamsBot -filterNone -selectAllProperties 
#$llcUsers = get-graphUsersWithEmployeeInfoExtensions -tokenResponse $tokenResponseTeamsBot -filterBusinessUnit "Anthesis LLC (USA)" -filterNone -selectAllProperties
#$phlUsers = get-graphUsersWithEmployeeInfoExtensions -tokenResponse $tokenResponseTeamsBot -filterBusinessUnit "Anthesis Philippines Inc. (PHL)" -filterNone -selectAllProperties
#$chnUsers = $allUsers | ? {$_.anthesisgroup_employeeInfo.businessUnit -eq "China"} 
$allUsers = get-graphUsers -tokenResponse $tokenResponseTeamsBot -filterLicensedUsers -selectAllProperties
$llcUsers = get-graphUsers -tokenResponse $tokenResponseTeamsBot -filterCustomEq  @{"anthesisgroup_employeeInfo/businessUnit" = "Anthesis LLC (USA)"} -filterLicensedUsers -selectAllProperties
$phlUsers = get-graphUsers -tokenResponse $tokenResponseTeamsBot -filterCustomEq  @{"anthesisgroup_employeeInfo/businessUnit" = "Anthesis Philippines Inc. (PHL)"} -filterLicensedUsers -selectAllProperties
$chnUsers = $allUsers | ? {$_.anthesisgroup_employeeInfo.businessUnit -eq "China"} 

$naUsers = $llcUsers
$naUsers += $phlUsers
$naUsers += $chnUsers

#Get all Intune, ATP & AAD devices
$tokenResponseIntuneBot = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName IntuneBot) -grant_type client_credentials 
$intuneDevices = get-graphIntuneDevices -tokenResponse $tokenResponseIntuneBot
$intuneDevices | % {
    Add-Member -InputObject $_ -MemberType NoteProperty -Name "deviceId" -Value $_.azureADDeviceId -Force
    }
$tokenResponseIntuneBotAtp = get-atpTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName IntuneBot) -grant_type client_credentials 
$atpDevices = get-atpMachines -tokenResponse $tokenResponseIntuneBotAtp
$atpDevices | % {
    Add-Member -InputObject $_ -MemberType NoteProperty -Name "deviceId" -Value $_.aadDeviceId -Force
    }
$allAadDevices = get-graphDevices -tokenResponse $tokenResponseTeamsBot

$allAadDevices | % {
    $thisDevice = $_
    $thisUserId = $thisDevice.physicalIds | ? {$_ -match "USER-GID"} | % {$($_ -split ":")[1]}
    $thisUser = $allUsers | ? {$_.id -eq $thisUserId}
    Add-Member -InputObject $_ -MemberType NoteProperty -Name "CurrentOwnerName" -Value $thisUser.displayName -Force
    Add-Member -InputObject $_ -MemberType NoteProperty -Name "CurrentOwnerEmail" -Value $thisUser.mail -Force
    Add-Member -InputObject $_ -MemberType NoteProperty -Name "CurrentOwnerStatus" -Value $thisUser.anthesisgroup_employeeInfo.contractType -Force
    Add-Member -InputObject $_ -MemberType NoteProperty -Name "BusinessUnit" -Value $thisUser.anthesisgroup_employeeInfo.businessUnit -Force
    Add-Member -InputObject $_ -MemberType NoteProperty -Name "HWID" -Value $($_.physicalIds | ? {$_ -match "\[HWID\]"}) -Force
    }

$dedupedAadDevices = $allAadDevices | Group-Object {$_.HWID} | % {$_.Group | Sort-Object approximateLastSignInDateTime | Select-Object -Last 1} #DeDupe and keep only the most recent instance of each Device
$dedupedAadDevices = $dedupedAadDevices | Group-Object {$_.CurrentOwnerEmail} | % {$_.Group | Sort-Object approximateLastSignInDateTime | Select-Object -Last 1} #DeDupe and keep only the most recent Device for each user
$dedupedAadDevices | % {
    $thisIntuneDevice = Compare-Object -ReferenceObject $intuneDevices -DifferenceObject $_ -Property deviceId -ExcludeDifferent -IncludeEqual -PassThru
    Add-Member -InputObject $_ -MemberType NoteProperty -Name EnrolledInIntune -Value $(if([string]::IsNullOrWhiteSpace($thisIntuneDevice.id)){$false}else{$true}) -Force
    $thisAtpDevice = Compare-Object -ReferenceObject $atpDevices -DifferenceObject $_ -Property deviceId -ExcludeDifferent -IncludeEqual -PassThru
    Add-Member -InputObject $_ -MemberType NoteProperty -Name EnrolledInMde -Value $(if([string]::IsNullOrWhiteSpace($thisAtpDevice.id)){$false}else{$true}) -Force
    Add-Member -InputObject $_ -MemberType NoteProperty -Name OsVersionFromMde -Value $(if([string]::IsNullOrWhiteSpace($thisAtpDevice.id)){}else{$thisAtpDevice.version}) -Force
    }
$dedupedAadDevicesWin = $dedupedAadDevices | ? {$_.operatingSystem -eq "Windows"}

$now = (Get-Date -f s).Replace(":",".")

$naUsers | sort-object {$_.anthesisgroup_employeeInfo.businessUnit},displayName | % {
    $thisUser = $_
    $thisUsersPC = $dedupedAadDevicesWin | ? {$_.CurrentOwnerEmail -eq $thisUser.mail}
    $thisUsefulObject = New-Object psobject -Property $([ordered]@{
        UserName = $thisUser.displayName
        UserEmail = $thisUser.mail
        UserBusinessUnit = $thisUser.anthesisgroup_employeeInfo.businessUnit
        UserLocation = $thisUser.usageLocation
        UserEmploymentContract = $thisUser.anthesisgroup_employeeInfo.contractType
        MachineName = $thisUsersPC.displayName
        MachineOs = $thisUsersPC.operatingSystem
        MachineOsVersion = $thisUsersPC.operatingSystemVersion
        MachineJoinType = $thisUsersPC.trustType
        MachineProfileType = $thisUsersPC.RegisteredDevice
        EnrolledInIntune = $thisUsersPC.EnrolledInIntune
        EnrolledInMde = $thisUsersPC.EnrolledInMde
        OsVersionFromMde = $thisUsersPC.OsVersionFromMde
        })
    $thisUsefulObject | Export-Csv -Path "C:\Users\KevMaitland\Downloads\NA-IntuneAdoptionReport_$now.csv" -NoTypeInformation -Append
    }


<#
$serviceAccounts = convertTo-arrayOfEmailAddresses "acsmailboxaccess@anthesisgroup.com
ACSSupport@anthesisgroup.com
conflictminerals@anthesisgroup.com
Microsoft.ECM@anthesisgroup.com
Varex.PEC@anthesisgroup.com"

$allUsers | ? {$serviceAccounts -contains $_.mail} | %  {
    set-graphUser -tokenResponse $tokenResponseTeamsBot -userIdOrUpn $_.id -userEmployeeInfoExtensionHash @{contractType="ServiceAccount"}
    }

$allUsers | ? {$serviceAccounts -contains $_.mail} | %  {
    set-graphUser -tokenResponse $tokenResponseTeamsBot -userIdOrUpn $_.id -userEmployeeInfoExtensionHash @{businessUnit="Anthesis Philippines Inc. (PHL)"}
    }

$subbies = convertTo-arrayOfEmailAddresses "DeAnn.Sarver@anthesisgroup.com
    Deby.Stabler@anthesisgroup.com
    qwest_ga@anthesisgroup.com
    Therese.Karkowski@anthesisgroup.com
    "
$allUsers | ? {$subbies -contains $_.mail} | %  {
    set-graphUser -tokenResponse $tokenResponseTeamsBot -userIdOrUpn skye.lei@anthesisgroup.com -userEmployeeInfoExtensionHash @{contractType="Employee"}
    set-graphUser -tokenResponse $tokenResponseTeamsBot -userIdOrUpn Alexa.Cotton@anthesisgroup.com -userEmployeeInfoExtensionHash @{contractType="Employee"}
    }
    


#>