
##########################################################################################################
#                                                                                                        #
#                                             Provision User                                             #
#                                                                                                        #
##########################################################################################################

<#---------------------------------------------Script Notes---------------------------------------------#>

#This script handles JUST 365 user creation - if you need an AD account, please create manually or use ad-newuser
#It will be able to pull new starter information from both the new and old lists (see below for the correct sections)
#Most of the cmdlets this script uses resides in the _PS_Library_UserManagement.psm1 module

<#------------------------------------------------------------------------------------------------------#>


#######################
#                     #
#        Setup        #
#                     #
#######################

<#--------Import Modules--------#>

Import-Module -Name _PS_Library_UserManagement.psm1
Import-Module -Name _PS_Library_GeneralFunctionality.psm1
Import-Module -Name _PS_Library_Graph.psm1


<#--------Logging--------#>

$logFile = "C:\ScriptLogs\provision-User.log"
$errorLogFile = "C:\ScriptLogs\provision-User_Errors.log"
$smtpServer = "anthesisgroup-com.mail.protection.outlook.com"


<#--------Service Connections--------#>

#365 services

connect-ToMsol 
connect-ToExo 
connect-toAAD 
Connect-MsolService 



#Graph - with userBot
$userBotDetails = get-graphAppClientCredentials -appName UserBot
$tokenResponse = get-graphTokenResponse -aadAppCreds $userBotDetails


#email out
$smtpBotDetails = get-graphAppClientCredentials -appName SmtpBot
$tokenResponseSmtp = get-graphTokenResponse -aadAppCreds $smtpBotDetails





<#--------Available License Check--------#>

#Just in case you want to save yourself a few more clicks, this will show you currently available licensing

get-available365licensecount -licensetype "all"

<#--------Current Users Check--------#>

$allCurrentUsers = get-graphUsers -tokenResponse $tokenResponse -filterLicensedUsers

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
        Write-host "Failed to update MSOL account mailbox permissions" -ForegroundColor Red
        log-Error "Failed to update MSOL account mailbox permissions"
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
        try{
        Set-User -Identity $upn -AuthenticationPolicy "Block Basic Auth"
        }
        catch{
        Write-host "Failed to update authentication policy to Block Basic Auth" -ForegroundColor Red
        log-Error "Failed to update authentication policy to Block Basic Auth"
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
$requests = $requests | Where-Object {(($_.FieldValues.StartDate | get-date -format s) -gt ((get-date).AddDays(-7) | get-date -Format s))}

if($requests){#Display a subset of Properties to help the user identify the correct account(s)
    $selectedRequests = $requests | Sort-Object -Property {$_.FieldValues.StartDate} -Descending | select {$_.FieldValues.Employee_x0020_Preferred_x0020_N},{$_.FieldValues.jobtitle},{$_.FieldValues.StartDate},{$_.FieldValues.Main_x0020_office0.Label},{$_.FieldValues.Line_x0020_Manager.LookupValue},{$_.FieldValues.Licensing},{$_.FieldValues.Primary_x0020_Team0.Label},{$_.FieldValues.GUID.Guid},{$_.FieldValues.GraphUserGUID} | Out-GridView -PassThru -Title "Highlight any requests to process and click OK" | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name "Guid" -Value $_.'$_.FieldValues.GUID.Guid';$_}
    #Then return the original requests as these contain the full details
    [array]$selectedRequests = Compare-Object -ReferenceObject $requests -DifferenceObject $selectedRequests -Property Guid -IncludeEqual -ExcludeDifferent -PassThru
    }


#Building
ForEach($thisUser in $selectedRequests){



#We don't to create user accounts that are more than 14 days in advance, this creates a security risk - if more than 14 days in advance pop box open and warn script runner with choice box to either continue or break (end loop).
$timespan = New-TimeSpan -Start ($thisUser.FieldValues.StartDate | get-date -format s) -End ((get-date).AddDays(15) | get-date -Format s)
If($timespan.days -gt 0){
 write-host "c1"
}
Else{

    Add-Type -AssemblyName PresentationCore,PresentationFramework
    $ButtonType = [System.Windows.MessageBoxButton]::YesNoCancel
    $MessageIcon = [System.Windows.MessageBoxImage]::Error
    $MessageBody = "The user you are trying to create is more than 14 days out from their start date. Creating users more than two weeks in advance can create a security risk, unless there are extenuating circumstances please check to make sure they are less than two weeks out from starting."
    $MessageTitle = "Confirm Deletion"
    $Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
    Write-Host "$Result"
        If($Result -eq "Yes"){
             write-host "continued"
        }
        Else{
            write-host "stopped"
            Break
            
        }
}




#Before we start, check the contract type
write-host "Before we start, what is the contract type?"
write-host "A: Employee"
write-host "B: Subcontractor"
$selection = Read-Host "Type A or B"
Switch($selection){
"A" {$contracttype = "Employee"}
"B" {$contracttype = "Subcontractor"}
}


#Get secondary geographic data from the term store
$officeterm = Get-PnPTerm -Identity $($thisUser.FieldValues.Main_x0020_Office0.Label) -TermGroup "Anthesis" -TermSet "Primary Workplaces" -Includes CustomProperties
$country = $officeTerm.CustomProperties.Country

#Get credential info from term store
$userProvisionCredentials =  Get-PnPTerm -Identity "UserProvision" -TermGroup "Anthesis" -TermSet "IT" -Includes CustomProperties -ErrorAction Stop

 
#365 user account: Create the 365 user
write-host "Creating MSOL account for $($upn = (remove-diacritics $($thisUser.FieldValues.Employee_x0020_Preferred_x0020_N.Trim().Replace(" ",".")+"@anthesisgroup.com"))) first, which will create the unliscensed 365 E1 user"    
    provision-365user -upn ($upn = (remove-diacritics $($thisUser.FieldValues.Employee_x0020_Preferred_x0020_N.Trim().Replace(" ",".")+"@anthesisgroup.com"))) `
    -plaintextpassword ($plaintextpassword = "$($userProvisionCredentials[0].CustomProperties.Values)") `
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
    -managerSAM = ($managerSAM = (($thisUser.FieldValues.Line_x0020_Manager.Email).split("@")[0]).replace("."," ")) `
    -businessunit = ($businessunit = ($thisUser.FieldValues.Business_x0020_Unit0.Label)) `
    -jobtitle = ($jobtitle = ($thisUser.FieldValues.JobTitle)) `
    -adCredentials = $adCredentials `
    -restCredentials = $restCredentials `
    -licensetype = ($licensetype = ($thisUser.FieldValues.Licensing.Split(" ")[1].Trim())) `
    -usagelocation = ($usagelocation = ($officeTerm.CustomProperties.'Usage Location')) `
    -timezone = ($timezone = ($officeterm.CustomProperties.'Timezone')) `



If($selection -ne "B"){
    #Add to a regional group - this needs rewriting into a function, bodging for now
    #$thisoffice = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -filterDisplayName "$($officeterm.CustomProperties.'365 Regional Group')" -Verbose
    $thisoffice = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -filterId "$($officeterm.CustomProperties.'356 Group GUID')" -Verbose
    $regionalmembersgroup = get-graphGroups -tokenResponse $tokenResponse -filterId "$($thisoffice.anthesisgroup_UGSync.memberGroupId)"
    If(($regionalmembersgroup | Measure-Object).Count -eq 1){
        add-DistributionGroupMember -Identity $regionalmembersgroup.mail -Member $upn -Confirm:$false -BypassSecurityGroupManagerCheck
        $graphuser = get-graphUsers -tokenResponse $tokenResponse -filterUpns $upn
        add-graphUsersToGroup -tokenResponse $tokenResponse -graphGroupId $thisoffice.id -memberType Members -graphUserUpns $graphuser.id -Verbose
    }
    Else{
    Write-Host "More than 1 group found for regional group. They haven't been added" -ForegroundColor Red
    Write-Error "More than 1 group found for regional group. They haven't been added"
}

    #Add to MDM groups - this is for Intune enrollment
    $BYOD = Read-Host "Add to MDM - BYOD user group? (y/n)"
    If ($BYOD -eq 'y') {
    Add-DistributionGroupMember -Identity "MDM-BYOD-MobileDeviceUsers@anthesisgroup.com" -Member $upn -Confirm:$false -BypassSecurityGroupManagerCheck
    }
    $COBO = Read-Host "Add to MDM - COBO user group (are they in GBR/NA/PHL/CHN/SWE/BRA/FRA/IRE/NLD/ZAF)? (y/n)"
    If ($COBO -eq 'y') {
    Add-DistributionGroupMember -Identity "MDM-CorporateMobileDeviceUsers@anthesisgroup.com" -Member $upn -Confirm:$false -BypassSecurityGroupManagerCheck
    add-graphLicenseToUser -tokenResponse $tokenResponse -userIdOrUpn $upn -licenseFriendlyName EMS_E3
    add-graphLicenseToUser -tokenResponse $tokenResponse -userIdOrUpn $upn -licenseFriendlyName MDE
}
}
Else{
Write-Host "Subcontractor - not adding to regional groups" -ForegroundColor White
}


#update employee extension info with graph
set-graphuser -tokenResponse $tokenResponse -userIdOrUpn $upn -userEmployeeInfoExtensionHash @{"businessUnit" = $($businessunit)}
set-graphuser -tokenResponse $tokenResponse -userIdOrUpn $upn -userEmployeeInfoExtensionHash @{"contractType" = $($contracttype)}
set-graphuser -tokenResponse $tokenResponse -userIdOrUpn $upn -userEmployeeInfoExtensionHash @{"extensionType" = "employeeInfo"}

#Set hire date from start date
#set-graphuser -tokenResponse $tokenResponse -userIdOrUpn $upn -userPropertyHash @{"employeeHireDate" = $($thisUser.FieldValues.StartDate)} -not available yet?

#Update phone numbers with graph (whole thing needs re-writing like this - fastest way to make amends at the moment)
if($thisUser.FieldValues.WorkPhone){
$businessnumberhash = @{businessPhones=@("$(($thisUser.FieldValues.WorkPhone).Trim())")}
set-graphuser -tokenResponse $tokenResponse -userIdOrUpn $upn -userPropertyHash $businessnumberhash
}
if($thisUser.FieldValues.CellPhone){
set-graphuser -tokenResponse $tokenResponse -userIdOrUpn $upn -userPropertyHash @{"mobilePhone" = "$(($thisUser.FieldValues.CellPhone).Trim())"}
}


#Return user to check what was set
sleep -Seconds 10
$thisProvisionedUser = ""
$thisProvisionedUser = get-graphUsers -tokenResponse $tokenResponse -filterUpns $upn -selectAllProperties -Verbose
If(($thisProvisionedUser | Measure-Object).count -eq 1){

Write-Host "Graph user object found - updating SPO list item"
Set-PnPListItem -List "New Starter Details" -Identity $thisUser.Id -Values @{"GraphUserGUID" = $thisProvisionedUser.id}
If($thisProvisionedUser.assignedPlans){
Write-Host "User appears to be licensed, emailing"
#Send email
 $body = "<HTML><BODY><p>Hi $($thisUser.FieldValues.Author.LookupValue),</p>
                <p>We have created a Microsoft account for $($thisUser.FieldValues.Employee_x0020_Preferred_x0020_N):</p>
                <p><b>username:</b> $($upn)<br>
                <b>password:</b> $($userProvisionCredentials[0].CustomProperties.Values)</p>
                <p>We have cc'd in the Line Manager added in the request</p> 
                <p>Kind regards,<br>
                The IT Team</p>
                "
send-graphMailMessage -tokenResponse $tokenResponseSmtp -fromUpn "Shared_Mailbox_-_IT_Team_GBR@anthesisgroup.com" -toAddresses $thisUser.FieldValues.Author.Email -subject "New User Requests - $($thisUser.FieldValues.Employee_x0020_Preferred_x0020_N)" -bodyHtml $body -ccAddresses $($thisUser.FieldValues.Line_x0020_Manager.Email) -bccAddresses "IT_Team_GBR@anthesisgroup.com" -Verbose
}
Else{
Write-Host "User does not appear to be licensed - you can buy and assign licenses and re-run lines 424-432 to send an automated email"
}
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
$requests = $requests | Where-Object {(($_.FieldValues.Start_x0020_Date | get-date -format s) -gt ((get-date).AddDays(-7) | get-date -Format s))}

if($requests){#Display a subset of Properties to help the user identify the correct account(s)
    $selectedRequests = $requests | Sort-Object -Property {$_.FieldValues.Start_x0020_Date} -Descending | select {$_.FieldValues.Title},{$_.FieldValues.Start_x0020_Date},{$_.FieldValues.Job_x0020_title},{$_.FieldValues.Primary_x0020_Workplace.Label},{$_.FieldValues.Line_x0020_Manager.LookupValue},{$_.FieldValues.Primary_x0020_Team.LookupValue},{$_.FieldValues.GUID.Guid},{$_.FieldValues.GraphUserGUID} | Out-GridView -PassThru -Title "Highlight any requests to process and click OK" | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name "Guid" -Value $_.'$_.FieldValues.GUID.Guid';$_}
    #Then return the original requests as these contain the full details
    [array]$selectedRequests = Compare-Object -ReferenceObject $requests -DifferenceObject $selectedRequests -Property Guid -IncludeEqual -ExcludeDifferent -PassThru
    }


ForEach($thisUser in $selectedRequests){

If($thisUser.FieldValues.GraphUserGUID){
Write-Host "It looks like this user has already been created"
break
}

#We don't to create user accounts that are more than 14 days in advance, this creates a security risk - if more than 14 days in advance pop box open and warn script runner with choice box to either continue or break (end loop).
$timespan = New-TimeSpan -Start ($thisUser.FieldValues.Start_x0020_Date | get-date -format s) -End ((get-date).AddDays(15) | get-date -Format s)
If($timespan.days -gt 0){
}
Else{

    Add-Type -AssemblyName PresentationCore,PresentationFramework
    $ButtonType = [System.Windows.MessageBoxButton]::YesNoCancel
    $MessageIcon = [System.Windows.MessageBoxImage]::Error
    $MessageBody = "The user you are trying to create is more than 14 days out from their start date. Creating users more than two weeks in advance can create a security risk, unless there are extenuating circumstances please check to make sure they are less than two weeks out from starting."
    $MessageTitle = "Confirm Deletion"
    $Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
    Write-Host "$Result"
        If($Result -eq "Yes"){
        }
        Else{
            Break
        }
}


$thisUser.FieldValues.Title
$thisUser.FieldValues.Job_x0020_title

#Before we start, check the contract type
write-host "Before we start, what is the contract type?"
write-host "A: Employee"
write-host "B: Subcontractor"
$selection = Read-Host "Type A or B"
Switch($selection){
"A" {$contracttype = "Employee"}
"B" {$contracttype = "Subcontractor"}
}

<#switch($thisUser.FieldValues.Is_x0020_a_x0020_Subcontractor_x){
    {$_ -match "Yes"}  {$contracttype = "Subcontractor"}
    {$_ -match "No"}   {$contracttype = "Employee"}
    default {write-host -f Red "[$($thisUser.FieldValues.Title)] does not have a valid Employment Status set. Will not create account.";break}
    }#>

#Use Secondary Office location if homeworker - primary for anything else
If($thisUser.fieldvalues.Primary_x0020_Workplace.Label -eq "Home worker"){
    #If there is actually data in the secondary workplace field - we can't make it mandatory
    If($thisUser.fieldvalues.Nearest_x0020_Office.Label){
    $officeterm = Get-PnPTerm -Identity $($thisUser.fieldvalues.Nearest_x0020_Office.Label) -TermGroup "Anthesis" -TermSet "offices" -Includes CustomProperties
    $country = $officeTerm.CustomProperties.Country
    $regionalgroup = $officeterm
    }
}
Else{
#Get secondary geographic data from the term store
$officeterm = Get-PnPTerm -Identity $($thisUser.fieldvalues.Primary_x0020_Workplace.Label) -TermGroup "Anthesis" -TermSet "offices" -Includes CustomProperties
$country = $officeTerm.CustomProperties.Country
$regionalgroup = $officeterm
}

#Get credential info from term store
$userProvisionCredentials =  Get-PnPTerm -Identity "UserProvision" -TermGroup "Anthesis" -TermSet "IT" -Includes CustomProperties -ErrorAction Stop

#365 user account: Create the 365 user
write-host "Creating MSOL account for $($upn = (remove-diacritics $($thisUser.FieldValues.Title.Trim().Replace(" ",".")+"@anthesisgroup.com"))) first, which will create the unliscensed 365 E1 user"    
    provision-365user -upn ($upn = (remove-diacritics $($thisUser.FieldValues.Title.Trim().Replace(" ",".")+"@anthesisgroup.com"))) `
    -plaintextpassword ($plaintextpassword = "$($userProvisionCredentials[0].CustomProperties.Values)") `
    -firstname ($firstname = "$($thisUser.FieldValues.Title.Trim().Split(" ")[0].Trim())") `
    -lastname = ($lastname = "$($thisUser.FieldValues.Title.Trim().Split(" ")[1].Trim())") `
    -displayname = ($displayname = "$(($thisUser.FieldValues.Title).Trim())") `
    -primaryteam = ($primaryteam = "$(($thisUser.FieldValues.Primary_x0020_Team.Label).Trim())") `
    -regionalgroup = $regionalgroup `
    -office = ($office = "$(($thisUser.FieldValues.Primary_x0020_Workplace.Label).Trim())") `
    -streetaddress = ($streetaddress = ($officeTerm.CustomProperties.'Street Address')) `
    -country = ($country = ($officeTerm.CustomProperties.Country)) `
    -city = ($city = "$(($thisUser.FieldValues.Primary_x0020_Workplace.Label).Trim())") `
    -postcode = ($postcode = ($officeTerm.CustomProperties.'Postal Code')) `
    -linemanager = ($linemanager = ($thisUser.FieldValues.Line_x0020_Manager.Email)) `
    -managerSAM = ($managerSAM = (($thisUser.FieldValues.Line_x0020_Manager.Email).split("@")[0]).replace("."," ")) `
    -businessunit = ($businessunit = ($thisUser.FieldValues.Finance_x0020_Cost_x0020_Attribu.Label)) `
    -jobtitle = ($jobtitle = ($thisUser.FieldValues.Job_x0020_title)) `
    -adCredentials = $adCredentials `
    -restCredentials = $restCredentials `
    -licensetype = ($licensetype = ($thisUser.FieldValues.Office_x0020_365_x0020_license.Split(" ").Trim())) `
    -usagelocation = ($usagelocation = ($officeTerm.CustomProperties.'Usage Location')) `
    -timezone = ($timezone = ($officeterm.CustomProperties.'Timezone')) 

    $graphUser = ""

    while(!$graphUser){$graphuser = get-graphUsers -tokenResponse $tokenResponse -filterUpns $upn -Verbose}
    
    If(($graphuser | Measure-Object).count -eq 1){

    Write-Host "Some requests will not be labelled correctly as a subcontractor, for Spain they have a lot of users that need an account, but should not have acess to internal Teams, `
and they have their own group called All Educators and Mediators (ESP). `
If they have a job title 'Educators', have a kiosk license, and if Primary Workplace is in Spain, it's likely that it is one of these users (a Subcontractor)." -ForegroundColor Yellow

    $addToAllEducatorsGroup = read-host "Do you want to add this user to the All Educators and Mediators (ESP) group? Answer Y or N"
    Switch($addToAllEducatorsGroup){
        "Y" {
            add-graphUsersToGroup -tokenResponse $tokenResponse -graphGroupId "b4a43a31-c282-4859-a038-76960efa2bd9" -memberType members -graphUserIds $graphuser.id -Verbose
            $contracttype = "Subcontractor"
                   }
        "N" {write-host "Not adding to the group"}
        default{write-host "Error: Please write Y or N" -ForegroundColor Red}
    }
    }

#Add to a regional group - this needs rewriting into a function, bodging for now
If($contracttype -ne "Subcontractor"){
    #$thisoffice = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -filterDisplayName "$($officeterm.CustomProperties.'365 Regional Group')" -Verbose
    $thisoffice = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -filterId "$($officeterm.CustomProperties.'356 Group GUID')" -Verbose
    $regionalmembersgroup = get-graphGroups -tokenResponse $tokenResponse -filterId "$($thisoffice.anthesisgroup_UGSync.memberGroupId)"
    If(($regionalmembersgroup | Measure-Object).Count -eq 1){
    add-DistributionGroupMember -Identity $regionalmembersgroup.mail -Member $upn -Confirm:$false -BypassSecurityGroupManagerCheck
    add-graphUsersToGroup -tokenResponse $tokenResponse -graphGroupId $thisoffice.id -memberType Members -graphUserUpns $graphuser.id -Verbose
    }
    Else{
    Write-Host "More than 1 group found for regional group. They haven't been added" -ForegroundColor Red
    Write-Error "More than 1 group found for regional group. They haven't been added"
    }
    

    #Add to MDM groups - this is for Intune enrollment
    $BYOD = Read-Host "Add to MDM - BYOD user group? (y/n)"
    If ($BYOD -eq 'y') {
    Add-DistributionGroupMember -Identity "MDM-BYOD-MobileDeviceUsers@anthesisgroup.com" -Member $upn -Confirm:$false -BypassSecurityGroupManagerCheck
    }
    $COBO = Read-Host "Add to MDM - COBO user group (are they in GBR/NA/PHL/CHN/SWE/BRA/FRA/IRE/NLD/ZAF)? (y/n)"
    If ($COBO -eq 'y') {
    Add-DistributionGroupMember -Identity "MDM-CorporateMobileDeviceUsers@anthesisgroup.com" -Member $upn -Confirm:$false -BypassSecurityGroupManagerCheck
    add-graphLicenseToUser -tokenResponse $tokenResponse -userIdOrUpn $upn -licenseFriendlyName EMS_E3
    add-graphLicenseToUser -tokenResponse $tokenResponse -userIdOrUpn $upn -licenseFriendlyName MDE
    }    
}

Else{
Write-Host "Subcontractor - not adding to regional groups" -ForegroundColor White
}
   

#update employee extension info with graph
set-graphuser -tokenResponse $tokenResponse -userIdOrUpn $upn -userEmployeeInfoExtensionHash @{"businessUnit" = $($businessunit)}
set-graphuser -tokenResponse $tokenResponse -userIdOrUpn $upn -userEmployeeInfoExtensionHash @{"contractType" = $($contracttype)}
set-graphuser -tokenResponse $tokenResponse -userIdOrUpn $upn -userEmployeeInfoExtensionHash @{"extensionType" = "employeeInfo"}

#Set hire date from start date
#set-graphuser -tokenResponse $tokenResponse -userIdOrUpn $upn -userPropertyHash @{"employeeHireDate" = $($thisUser.FieldValues.Start_x0020_Date)} -not available yet?

#Set hire date from start date
#set-graphuser -tokenResponse $tokenResponse -userIdOrUpn $upn -userPropertyHash @{"employeeHireDate" = $($thisUser.FieldValues.Start_x0020_Date)} -not available yet?

#Update phone numbers with graph (whole thing needs re-writing like this - fastest way to make amends at the moment)
If($thisUser.FieldValues.Landline_x0020_phone_x0020_numbe){
$businessnumberhash = @{businessPhones=@("$(($thisUser.FieldValues.Landline_x0020_phone_x0020_numbe).Trim())")}
set-graphuser -tokenResponse $tokenResponse -userIdOrUpn $upn -userPropertyHash $businessnumberhash
}
If($thisUser.FieldValues.Mobile_x002f_Cell_x0020_phone_x0){
set-graphuser -tokenResponse $tokenResponse -userIdOrUpn $upn -userPropertyHash @{"mobilePhone" = "$(($thisUser.FieldValues.Mobile_x002f_Cell_x0020_phone_x0).Trim())"}
}

#Return user to check what was set
sleep -Seconds 10
$thisProvisionedUser = ""
$thisProvisionedUser = get-graphUsers -tokenResponse $tokenResponse -filterUpns $upn -selectAllProperties -Verbose
If(($thisProvisionedUser | Measure-Object).count -eq 1){

Write-Host "Graph user object found - updating SPO list item"
Set-PnPListItem -List "New User Requests" -Identity $thisUser.Id -Values @{"GraphUserGUID" = $thisProvisionedUser.id}

If($thisProvisionedUser.assignedPlans){
Write-Host "User appears to be licensed, emailing"
#Send email
 $body = "<HTML><BODY><p>Hi $($thisUser.FieldValues.Author.LookupValue),</p>
                <p>We have created a Microsoft account for $($thisUser.FieldValues.Employee_x0020_Legal_x0020_Name):</p>
                <p><b>username:</b> $($upn)<br>
                <b>password:</b> $($userProvisionCredentials[0].CustomProperties.Values)</p>
                <p>We have cc'd in the Line Manager added in the request</p> 
                <p>Kind regards,<br>
                The IT Team</p>
                "
send-graphMailMessage -tokenResponse $tokenResponseSmtp -fromUpn "Shared_Mailbox_-_IT_Team_GBR@anthesisgroup.com" -toAddresses $thisUser.FieldValues.Author.Email -subject "New User Requests - $($thisUser.FieldValues.Employee_x0020_Legal_x0020_Name)" -bodyHtml $body -ccAddresses $($thisUser.FieldValues.Line_x0020_Manager.Email) -bccAddresses "IT_Team_GBR@anthesisgroup.com" -Verbose
}
Else{
Write-Host "User does not appear to be licensed - you can buy and assign licenses and re-run lines 436-448 to send an automated email"
}
}
}








