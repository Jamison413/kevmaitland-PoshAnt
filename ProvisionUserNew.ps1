
##########################################################################################################
#                                                                                                        #
#                                             Provision User                                             #
#                                                                                                        #
##########################################################################################################




<#---------------------------------------------Script Notes---------------------------------------------#>

#This script handles both 365 and AD user creation
#It will be able to pull new starter information from both the new and old lists (see below for the correct sections)
#Most of the cmdlets this script uses resides in the _PS_Library_UserManagement.psm1 module
#This script CANNOT be run in PS6 (pscore) if you are trying to create an AD user, the ActiveDirectory module is not ported over yet fully.


<#------------------------------------------------------------------------------------------------------#>


#######################
#                     #
#        Setup        #
#                     #
#######################

<#--------Import Modules--------#>

Import-Module -Name ActiveDirectory #Not compatible with pscore
Import-Module -Name _PS_Library_UserManagement.psm1 #This has a compatibilty issue with core - something to dig into, shuld work in ISE?
Import-Module -Name _PS_Library_GeneralFunctionality.psm1
Import-Module -Name _PS_Library_Graph


<#--------Logging--------#>

$logFile = "C:\ScriptLogs\provision-User.log"
$errorLogFile = "C:\ScriptLogs\provision-User_Errors.log"
$smtpServer = "anthesisgroup-com.mail.protection.outlook.com"


<#--------Service Connections--------#>

$msolCredentials = set-MsolCredentials #Set these once as a PSCredential object and use that to build the CSOM SharePointOnlineCredentials object and set the creds for REST
$restCredentials = new-spoCred -username $msolCredentials.UserName -securePassword $msolCredentials.Password
$csomCredentials = new-csomCredentials -username $msolCredentials.UserName -password $msolCredentials.Password
connect-ToMsol -credential $msolCredentials
connect-ToExo -credential $msolCredentials
connect-toAAD -credential $msolCredentials
Connect-MsolService -credential $msolCredentials

$teamBotDetails = import-encryptedCsv -pathToEncryptedCsv "$env:USERPROFILE\Desktop\teambotdetails.txt"
$tokenResponse = get-graphTokenResponse -aadAppCreds $teamBotDetails

$adCredentials = Get-Credential -Message "Enter local AD Administrator credentials to create a new user in AD" -UserName "$env:USERDOMAIN\username"

<#--------Available License Check--------#>

#Just in case you want to save yourself a few more clicks, this will show you currently available licensing

get-available365licensecount -licensetype "all"



<#--------Create Meta-Functions--------#>

Function provision-365user {

     Param($upn)

        try{
        write-host "Creating MSOL account for $upn" -ForegroundColor Yellow
        create-MsolUser -upn $upn -plaintextpassword $plaintextpassword
        write-host "Account created" -ForegroundColor Yellow
        }
        catch{
        Write-host "Failed to create MSOL account" -ForegroundColor Red
        log-Error "Failed to create MSOL account"
        log-Error $Error
        }
Start-Sleep -Seconds 60 #Let MSOL & EXO Syncronise
$msoluser = Get-MsolUser -UserPrincipalName $upn #check it can be retrieved
If($msoluser){
        try{
        update-msoluserdetails -upn $upn -firstname $firstname -lastname $lastname -displayname $displayname -primaryteam $primaryteam -country $country -jobtitle $jobtitle -DDI $DDI -mobile $mobile -businessunit $businessunit -city $city -streetaddress $streetaddress -office $office -postcode $postcode -usagelocation $usagelocation
        }
        catch{
        Write-host "Failed to update MSOL account details" -ForegroundColor Red
        log-Error "Failed to update MSOL account details"
        log-Error $Error
        }
        try{
        update-msolusercoregroups -upn $upn -office $office -businessunit $businessunit -regionalgroup $regionalgroup #Need to add office option to figure this out just from the office location?
        }
        catch{
        Write-host "Failed to update MSOL account core groups" -ForegroundColor Red
        log-Error "Failed to update MSOL account core groups"
        log-Error $Error
        }
        try{
        update-msolMailboxViaUpn -upn $upn -displayname $displayname -businessunit $businessunit -timezone $timezone -linemanager $linemanager -office $office #this seemed to work - potential problem with connect-toexo in remote session
        }
        catch{
        Write-host "Failed to update MSOL account mailbox details" -ForegroundColor Red
        log-Error "Failed to update MSOL account mailbox details"
        log-Error $Error
        }
        try{
        set-mailboxPermissions -upn $upn -managerSAM $managerSAM -businessunit $businessunit
        }
        catch{
        Write-host "Failed to update MSOL account mailbox premissions" -ForegroundColor Red
        log-Error "Failed to update MSOL account mailbox premissions"
        log-Error $Error
        }
        try{
        license-msolUser -upn $upn -licensetype $licensetype
        }
        catch{
        Write-host "Failed to update MSOL account licensing" -ForegroundColor Red
        log-Error "Failed to update MSOL account licensing"
        log-Error $Error
        }
}
Else{
write-host "*****************Failed to retrieve msol user account in time: $($upn)*****************" -ForegroundColor red
}
}




#################################
#                               #
#      Runthrough - New List    #
#                               #
#################################




<#--------Retrieve Requests from Sharepoint--------#>

#Get the New User Requests that have not been marked as processed
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/teams/People_Services_Team_All_365/" -UseWebLogin #-Credentials $msolCredentials
$requests = (Get-PnPListItem -List "New Starter Details" -Query "<View><Query><Where><Eq><FieldRef Name='New_x0020_Starter_x0020_Setup_x0'/><Value Type='String'>Waiting</Value></Eq></Where></Query></View>") |  % {Add-Member -InputObject $_ -MemberType NoteProperty -Name Guid -Value $_.FieldValues.GUID.Guid;$_}
if($requests){#Display a subset of Properties to help the user identify the correct account(s)
    $selectedRequests = $requests | Sort-Object -Property {$_.FieldValues.StartDate} -Descending | select {$_.FieldValues.Employee_x0020_Preferred_x0020_N},{$_.FieldValues.jobtitle},{$_.FieldValues.StartDate},{$_.FieldValues.Main_x0020_office0.Label},{$_.FieldValues.Line_x0020_Manager.LookupValue},{$_.FieldValues.Licensing},{$_.FieldValues.Primary_x0020_Team0.Label},{$_.FieldValues.GUID.Guid} | Out-GridView -PassThru -Title "Highlight any requests to process and click OK" | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name "Guid" -Value $_.'$_.FieldValues.GUID.Guid';$_}
    #Then return the original requests as these contain the full details
    [array]$selectedRequests = Compare-Object -ReferenceObject $requests -DifferenceObject $selectedRequests -Property Guid -IncludeEqual -ExcludeDifferent -PassThru
    }



ForEach($thisUser in $selectedRequests){

#Get secondary geographic data from the term store
$officeterm = Get-PnPTerm -Identity $($thisUser.FieldValues.Main_x0020_Office0.Label) -TermGroup "Anthesis" -TermSet "offices" -Includes CustomProperties
$regionalgroup = (Get-DistributionGroup -Identity $officeterm.CustomProperties.'365 Regional Group').guid
$country = $officeTerm.CustomProperties.Country
   
#365 user account: Create the 365 user
write-host "Creating MSOL account for $($upn = (remove-diacritics $($thisUser.FieldValues.Employee_x0020_Preferred_x0020_N.Trim().Replace(" ",".")+"@anthesisgroup.com"))) first, which will create the unliscensed 365 E1 user"    
    provision-365user -upn ($upn = (remove-diacritics $($thisUser.FieldValues.Employee_x0020_Preferred_x0020_N.Trim().Replace(" ",".")+"@anthesisgroup.com"))) `
    -plaintextpassword ($plaintextpassword = "Anthesis123") `
    -firstname ($firstname = "$($thisUser.FieldValues.Employee_x0020_Preferred_x0020_N.Trim().Split(" ")[0].Trim())") `
    -lastname = ($lastname = "$(($thisUser.FieldValues.Employee_x0020_Preferred_x0020_N.Trim().Split(" ")[$thisUser.FieldValues.Employee_x0020_Preferred_x0020_N.Trim().Split(" ").Count-1]).Trim())") `
    -displayname = ($displayname = "$(($thisUser.FieldValues.Employee_x0020_Preferred_x0020_N).Trim())") `
    -primaryteam = ($primaryteam = "$(($thisUser.FieldValues.Primary_x0020_Team0.Label).Trim())") `
    -regionalgroup = $regionalgroup `
    -office = ($office = "$(($thisUser.FieldValues.Main_x0020_Office0.Label).Trim())") `
    -streetaddress = ($streetaddress = ($officeTerm.CustomProperties.'Street Address')) `
    -country = ($country = ($officeTerm.CustomProperties.Country)) `
    -city = ($city = "$(($thisUser.FieldValues.Main_x0020_Office0.Label).Trim())") `
    -postcode = ($postcode = ($officeTerm.CustomProperties.'Postal Code')) `
    -linemanager = ($linemanager = ($thisUser.FieldValues.Line_x0020_Manager.Email)) `
    -businessunit = ($businessunit = ($thisUser.FieldValues.Business_x0020_Unit0.Label)) `
    -jobtitle = ($jobtitle = ($thisUser.FieldValues.JobTitle)) `
    -plaintextpassword = "Anthesis123" `
    -adCredentials = $adCredentials `
    -restCredentials = $restCredentials `
    -licensetype = ($licensetype = ($thisUser.FieldValues.Licensing.Split(" ")[1].Trim())) `
    -usagelocation = ($usagelocation = ($officeTerm.CustomProperties.'Usage Location')) `
    -timezone = ($timezone = ($officeterm.CustomProperties.'Timezone')) `

#update employee extension info with graph (just business unit)
set-graphuser -tokenResponse $tokenResponse -userIdOrUpn $upn -userEmployeeInfoExtensionHash @{"businessUnit" = $($businessunit)}

#Update phone numbers with graph (whole thing needs re-writing like this - fastest way to make amends at the moment)
$businessnumberhash = @{businessPhones=@("$(($thisUser.FieldValues.WorkPhone).Trim())")}
set-graphuser -tokenResponse $tokenResponse -userIdOrUpn $upn -userPropertyHash $businessnumberhash
set-graphuser -tokenResponse $tokenResponse -userIdOrUpn $upn -userPropertyHash @{"mobilePhone" = "$(($thisUser.FieldValues.CellPhone).Trim())"}

    
#AD user account: If user will be based in Bristol or London office, offer to create an AD user account
If((![string]::IsNullOrWhiteSpace($upn)) -and (("Bristol, GBR" -eq $office) -or ("London, GBR" -eq $office))){
write-host "It looks like this user will either be based in the Bristol or London offices." -ForegroundColor Yellow
$confirmation = Read-Host "Create an AD account? (y/n)"
if ($confirmation -eq 'y') {
Write-Host "Okay, let's create an AD account for $($upn)..." -ForegroundColor Yellow
}
$allpermanentstaffadgroupprompt = Read-Host "Do we also want to add the New Starter to the All Permanant Staff AD Group? (y/n)"
    try{
write-host "Creating AD account for $(remove-diacritics $($thisUser.FieldValues.Employee_x0020_Preferred_x0020_N.Trim().Replace(" ",".")+"@anthesisgroup.com"))"    
    create-ADUser -upn $upn `
    -firstname $firstname `
    -surname ($surname = $lastname) `
    -displayname $displayname `
    -managerSAM ($managerSAM =  ($thisUser.FieldValues.Line_x0020_Manager.Email.split("@")[0])) `
    -primaryteam $primaryteam `
    -jobtitle $jobtitle `
    -plaintextpassword $plaintextpassword `
    -businessunit $businessunit `
    -adCredentials $adCredentials `
    -office $office `
    -allpermanentstaffadgroupprompt $allpermanentstaffadgroupprompt `
    -SAMaccountname ($SAMaccountname = $($upn.Split("@")[0]))
    }
Catch{
    Write-host "Failed to create AD account" -ForegroundColor Red
    log-Error "Failed to create AD account"
    log-Error $Error
    }
}
Else{
write-host "Okay, we will stop here." -ForegroundColor White
}
}






#################################
#                               #
#      Runthrough - Old List    #
#                               #
#################################





<#--------Retrieve Requests from Sharepoint--------#>

#Get the New User Requests
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/teams/hr" -UseWebLogin #-Credentials $msolCredentials
$requests = (Get-PnPListItem -List "New User Requests" -Query "<View><Query><Where><Eq><FieldRef Name='Current_x0020_Status'/><Value Type='String'>1 - Waiting for IT Team to set up accounts</Value></Eq></Where></Query></View>") |  % {Add-Member -InputObject $_ -MemberType NoteProperty -Name Guid -Value $_.FieldValues.GUID.Guid;$_}
if($requests){#Display a subset of Properties to help the user identify the correct account(s)
    $selectedRequests = $requests | Sort-Object -Property {$_.FieldValues.Start_x0020_Date} -Descending | select {$_.FieldValues.Title},{$_.FieldValues.Start_x0020_Date},{$_.FieldValues.Job_x0020_title},{$_.FieldValues.Primary_x0020_Workplace.Label},{$_.FieldValues.Line_x0020_Manager.LookupValue},{$_.FieldValues.Primary_x0020_Team.LookupValue},{$_.FieldValues.GUID.Guid} | Out-GridView -PassThru -Title "Highlight any requests to process and click OK" | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name "Guid" -Value $_.'$_.FieldValues.GUID.Guid';$_}
    #Then return the original requests as these contain the full details
    [array]$selectedRequests = Compare-Object -ReferenceObject $requests -DifferenceObject $selectedRequests -Property Guid -IncludeEqual -ExcludeDifferent -PassThru
    }



ForEach($thisUser in $selectedRequests){

#Get secondary geographic data from the term store
$officeterm = Get-PnPTerm -Identity $($thisUser.fieldvalues.Primary_x0020_Workplace.Label) -TermGroup "Anthesis" -TermSet "offices" -Includes CustomProperties
$regionalgroup = (Get-DistributionGroup -Identity $officeterm.CustomProperties.'365 Regional Group').guid
$country = $officeTerm.CustomProperties.Country

   
#365 user account: Create the 365 user
write-host "Creating MSOL account for $($upn = (remove-diacritics $($thisUser.FieldValues.Title.Trim().Replace(" ",".")+"@anthesisgroup.com"))) first, which will create the unliscensed 365 E1 user"    
    provision-365user -upn ($upn = (remove-diacritics $($thisUser.FieldValues.Title.Trim().Replace(" ",".")+"@anthesisgroup.com"))) `
    -plaintextpassword ($plaintextpassword = "Anthesis123") `
    -firstname ($firstname = "$($thisUser.FieldValues.Title.Trim().Split(" ")[0].Trim())") `
    -lastname = ($lastname = "$(($thisUser.FieldValues.Title.Trim().Split(" ")[$thisUser.FieldValues.Employee_x0020_Preferred_x0020_N.Trim().Split(" ").Count-1]).Trim())") `
    -displayname = ($displayname = "$(($thisUser.FieldValues.Title).Trim())") `
    -primaryteam = ($primaryteam = "$(($thisUser.FieldValues.Primary_x0020_Team.Label).Trim())") `
    -regionalgroup = $regionalgroup `
    -office = ($office = "$(($thisUser.FieldValues.Primary_x0020_Workplace.Label).Trim())") `
    -streetaddress = ($streetaddress = ($officeTerm.CustomProperties.'Street Address')) `
    -country = ($country = ($officeTerm.CustomProperties.Country)) `
    -city = ($city = "$(($thisUser.FieldValues.Primary_x0020_Workplace.Label).Trim())") `
    -postcode = ($postcode = ($officeTerm.CustomProperties.'Postal Code')) `
    -linemanager = ($linemanager = ($thisUser.FieldValues.Line_x0020_Manager.Email)) `
    -businessunit = ($businessunit = ($thisUser.FieldValues.Finance_x0020_Cost_x0020_Attribu.Label)) `
    -jobtitle = ($jobtitle = ($thisUser.FieldValues.Job_x0020_title)) `
    -plaintextpassword = "Anthesis123" `
    -adCredentials = $adCredentials `
    -restCredentials = $restCredentials `
    -licensetype = ($licensetype = ($thisUser.FieldValues.Office_x0020_365_x0020_license.Split(" ").Trim())) `
    -usagelocation = ($usagelocation = ($officeTerm.CustomProperties.'Usage Location')) `
    -timezone = ($timezone = ($officeterm.CustomProperties.'Timezone')) 

#update employee extension info with graph (just business unit)
set-graphuser -tokenResponse $tokenResponse -userIdOrUpn $upn -userEmployeeInfoExtensionHash @{"businessUnit" = $($businessunit)}

#Update phone numbers with graph (whole thing needs re-writing like this - fastest way to make amends at the moment)
$businessnumberhash = @{businessPhones=@("$(($thisUser.FieldValues.Landline_x0020_phone_x0020_numbe).Trim())")}
set-graphuser -tokenResponse $tokenResponse -userIdOrUpn $upn -userPropertyHash $businessnumberhash
set-graphuser -tokenResponse $tokenResponse -userIdOrUpn $upn -userPropertyHash @{"mobilePhone" = "$(($thisUser.FieldValues.Mobile_x002f_Cell_x0020_phone_x0).Trim())"}

    
#AD user account: If user will be based in Bristol or London office, offer to create an AD user account
If((![string]::IsNullOrWhiteSpace($upn)) -and (("Bristol, GBR" -eq $office) -or ("London, GBR" -eq $office))){
write-host "It looks like this user will either be based in the Bristol or London offices." -ForegroundColor Yellow
$confirmation = Read-Host "Create an AD account? (y/n)"
if ($confirmation -eq 'y') {
Write-Host "Okay, let's create an AD account for $($upn)..." -ForegroundColor Yellow
}
$allpermanentstaffadgroupprompt = Read-Host "Do we also want to add the New Starter to the All Permanant Staff AD Group? (y/n)"
    try{
write-host "Creating AD account for $(remove-diacritics $($thisUser.FieldValues.Employee_x0020_Preferred_x0020_N.Trim().Replace(" ",".")+"@anthesisgroup.com"))"    
    create-ADUser -upn $upn `
    -firstname $firstname `
    -surname ($surname = $lastname) `
    -displayname $displayname `
    -managerSAM ($managerSAM =  ($thisUser.FieldValues.Line_x0020_Manager.Email.split("@")[0])) `
    -primaryteam $primaryteam `
    -jobtitle $jobtitle `
    -plaintextpassword $plaintextpassword `
    -businessunit $businessunit `
    -adCredentials $adCredentials `
    -office $office `
    -allpermanentstaffadgroupprompt $allpermanentstaffadgroupprompt `
    -SAMaccountname ($SAMaccountname = $($upn.Split("@")[0]))
    }
Catch{
    Write-host "Failed to create AD account" -ForegroundColor Red
    log-Error "Failed to create AD account"
    log-Error $Error
    }
}
Else{
write-host "Okay, we will stop here." -ForegroundColor White
}
}















