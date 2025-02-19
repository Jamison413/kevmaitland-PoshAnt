﻿#Param function to place UPN string into - who do you want to deactivate? Place their name in the prompt in the terminal
param(
    [CmdletBinding()]
    [parameter(Mandatory = $true)]
    [ValidateNotNull()]
    [ValidatePattern(".[@].")]
    [string]$upnsString
    )

#Import other code we need - modules contain useful functions to make tasks consistent    
Import-Module _PS_Library_GeneralFunctionality
Import-Module _PS_Library_Databases.psm1
Import-Module _PS_Library_MSOL.psm1
Import-Module _REST_Library-SPO.psm1

#Some functions are just needed for this script, so we put them here and load them
#region functions
function add-emailAddressesToPublicFolder($publicFolder, $emailAddressArray){
    $tempPF = $publicFolder #Bodge to get [Microsoft.Exchange.Data.ProxyAddressCollection] without Library
    foreach ($externalEmailAddress in $emailAddressArray){$tempPF.EmailAddresses += $externalEmailAddress}
    $publicFolder | Set-MailPublicFolder -EmailAddresses $tempPF.EmailAddresses -EmailAddressPolicyEnabled $false
    }
function remove-emailAddressesToPublicFolder($publicFolder, $emailAddressArray){
    $tempPF = $publicFolder #Bodge to get [Microsoft.Exchange.Data.ProxyAddressCollection] without Library
    $tempPF.EmailAddresses = $publicFolder.EmailAddresses | ?{$emailAddressArray -notcontains $_}
    $publicFolder | Set-MailPublicFolder -EmailAddresses $tempPF.EmailAddresses -EmailAddressPolicyEnabled $false
    }
function take-ownership([String]$folderPath, $newOwner){
	if($newOwner.split("\").Count -lt 2){$newOwner = "SUSTAINLTD\$newOwner"}
    elseif($newOwner.Split("\")[0] -notmatch "SUSTAINLTD"){$newOwner = "SUSTAINLTD\"+($newOwner.Split("\")[1])}
        
    iex -Command "takeown.exe /A /F $($folderPath)"
	$CurrentACL = Get-Acl $folderPath
	write-host ...Adding NT Authority\SYSTEM to $folderPath -Fore Yellow
	$SystemACLPermission = "NT AUTHORITY\SYSTEM","FullControl","ContainerInherit,ObjectInherit","None","Allow"
	$SystemAccessRule = new-object System.Security.AccessControl.FileSystemAccessRule $SystemACLPermission
	$CurrentACL.AddAccessRule($SystemAccessRule)
	write-host ...Adding AdminGroup to $folderPath -Fore Yellow
	$AdminACLPermission = "$newOwner","FullControl","ContainerInherit,ObjectInherit","None","Allow"
	$SystemAccessRule = new-object System.Security.AccessControl.FileSystemAccessRule $AdminACLPermission
	$CurrentACL.AddAccessRule($SystemAccessRule)
	Set-Acl -Path $folderPath -AclObject $CurrentACL
    }
function get-msolUpn($userSAM){
    try{$oUser = Get-MsolUser -SearchString $userSAM.Replace("."," ")
        if($oUser.Count -eq 1){$oUser.UserPrincipalName}
        else{
            $allOUsers = Get-MsolUser
            for($i=0;$i -lt $userSAM.Split(".")[0].Length;$i++){
                $allOUsers | ?{($_.DisplayName -match $userSAM.Split(".")[1]) -and ($_.DisplayName -match $userSAM.Split(".")[0].Substring(0,$i))}
                if($oUser.Count -eq 1){$oUser.UserPrincipalName;break}
                }
            return $false #If we can't find the individual after searching, just return $false
            }
        }
    catch{$Error[0]}
    }
function reassign-emailAddresses($userSAM,$exportAdmin,$reassignTo,$guid){
    $externalEmailAddresses = ,@()
    if (!$guid){
        $guid = New-Guid
        Set-MsolUserPrincipalName -UserPrincipalName "$userSAM@anthesisgroup.com" -NewUserPrincipalName "$guid@anthesisgroup.com"
        Set-Mailbox $userSAM -WindowsEmailAddress $guid@anthesisgroup.com 
        }
    #New-MailboxSearch -Name $userSAM -Description "Hold for Export $userSAM" -InPlaceHoldEnabled $true -SourceMailboxes (Get-Mailbox $userSAM@anthesisgroup.com).Id
    Add-MailboxPermission -Identity $userSAM -AccessRights FullAccess -user $exportAdmin -AutoMapping $true

    Write-Host -ForegroundColor Yellow "$userSAM ready for manual export by $exportAdmin"
    foreach ($externalEmailAddress in ((Get-MsolUser -UserPrincipalName $guid@anthesisgroup.com).ProxyAddresses | ? {(($_ -match "@sustain.co.uk") -or ($_ -match "@anthesisgroup.com")) -and ($_ -notmatch $guid)})){    #Get any external e-mail addresses associated with the User's mailbox 
        Write-Host -ForegroundColor Yellow "Removing e-mail address" $externalEmailAddress
        $externalEmailAddresses += $externalEmailAddress
        Set-Mailbox $userSAM -EmailAddresses @{remove=$externalEmailAddress}
        }
    Start-Sleep -Seconds 3 #Give EXO a chance to process the changes
    if($reassignTo){
        foreach($externalEmailAddress in $externalEmailAddresses){Set-Mailbox $reassignTo -EmailAddresses @{add=$externalEmailAddress}}
        }
    else{
        $pf = Get-MailPublicFolder "\1.Public Folders\1.Admin\IT\AutoReplyBot"
        add-emailAddressesToPublicFolder -publicFolder $pf -emailAddressArray $externalEmailAddresses
        }
    #return $guid
	}
function remove-msolLicenses($userSAM){
    $oUser = Get-MsolUser -UserPrincipalName $userSAM@anthesisgroup.com
    if($oUser.Licenses.Count -eq 0){Write-Host -ForegroundColor DarkYellow "$userSAM had no licenses to remove"}
    foreach($license in $oUser.Licenses){Set-MsolUserLicense -UserPrincipalName $userSAM@anthesisgroup.com -RemoveLicenses $license.AccountSkuId}
    }
function reset-msolPassword($userSAM, $plaintextPassord){
    Set-MsolUserPassword -UserPrincipalName $userSAM@anthesisgroup.com -NewPassword $plaintextPassord -ForceChangePassword $true
    }
function reset-adPassword($userSAM, $plaintextPassord){
    Set-ADAccountPassword -Identity $userSAM -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "$plaintextPassord" -Force)
    }
function delete-msolAccount($userSAM){
    Remove-MsolUser -UserPrincipalName $userSAM@anthesisgroup.com -force
    }
function delete-adAccount($userSAM){
    Get-ADUser -Identity $userSAM | Remove-ADUser
    }
function move-PersonalFolder($userSAM){
    $personalFolderPath = "\\sustainltd.local\data\Personal\$userSAM"
    $archivedFolderPath = "\\hv06\Exsustainus\Personal Folders\$userSAM"
    foreach($folder in (gci $personalFolderPath -Recurse -Directory)){take-ownership -Folder $folder.FullName -newOwner "$env:USERDNSDOMAIN\$env:USERNAME"}
    Move-Item -Path $personalFolderPath -Destination $archivedFolderPath -Force
    }
function disable-ArenaAccount($userSAM){
    $sql = "UPDATE TS_USERS
            SET TSU_ENABLED = 0
            WHERE TSU_EMAIL_ADDRESS LIKE '$userSAM%'"
    Execute-SQLQueryOnSQLDB -query $sql -queryType "NonQuery" -sqlServerConnection $sqlConnection
    }
function delete-goldMineAccount($userSAM){Write-Host -ForegroundColor DarkCyan "You need to manually delete the GoldMine user $userSAM"}
function redirect-phone($userSAM){Write-Host -ForegroundColor DarkCyan "You need to manually delete the ShoreTel user $userSAM"}
function export-mailbox($userSAM){Write-Host -ForegroundColor DarkCyan "You need to manually export the mailbox for user $userSAM"}
function deprovision-user($userSAM, $plaintextPassword, $exportAdmin, $reassignEmailAddressesTo){
    try{reset-adPassword -userSAM $userSAM -plaintextPassord $plaintextPassword;Write-Host -ForegroundColor DarkYellow "AD password reset"}
    catch{Write-Host -ForegroundColor Red "Failed to reset AD password"}
    try{reset-msolPassword -userSAM $userSAM -plaintextPassord $plaintextPassword;Write-Host -ForegroundColor DarkYellow "MSOL password reset"}
    catch{Write-Host -ForegroundColor Red "Failed to reset MSOL password"}
    try{move-PersonalFolder -userSAM $userSAM;Write-Host -ForegroundColor DarkYellow "Personal Folder moved"}
    catch{Write-Host -ForegroundColor Red "Failed to move Personal Folder"}
    try{disable-ArenaAccount -userSAM $userSAM;Write-Host -ForegroundColor DarkYellow "ARENA account disabled"}
    catch{Write-Host -ForegroundColor Red "Failed to disable ARENA account"}
    delete-goldMineAccount -userSAM $userSAM
    redirect-phone -userSAM $userSAM
    export-mailbox -userSAM $userSAM
    try{reassign-emailAddresses -userSAM $userSAM -exportAdmin $exportAdmin -reassignTo $reassignEmailAddressesTo;Write-Host -ForegroundColor DarkYellow "E-mail addresses reassigned"}
    catch{Write-Host -ForegroundColor Red "Failed to reassign e-mail addresses"} #This changes the UPN for msol User
	}
function delete-userAccounts($userSAM){
    if((Read-Host -Prompt "Are you sure that you want to delete the user accounts for $userSAM`? And have you exported their mailbox?`r`nType YES to proceed") -eq "YES"){
        $newUserSAM = (get-msolUpn $userSAM).Replace("@anthesisgroup.com","")
        try{delete-adAccount -userSAM $userSAM Write-Host -ForegroundColor DarkYellow "AD Account deleted!"}
        catch{Write-Host -ForegroundColor Red "Failed to delete AD Account!";$Error[0]}
        try{delete-msolAccount -userSAM $newUserSAM Write-Host -ForegroundColor DarkYellow "MSOL Account deleted!"}
        catch{Write-Host -ForegroundColor Red "Failed to delete MSOL Account!";$Error[0]}
        }
    }
#endregion

#Set logging so we have more details if something goes wrong - we print the output of the terminal to the log.txt file
$logFileLocation = "C:\ScriptLogs\"
$scriptName = "deprovision-user"
$fullLogPathAndName = $logFileLocation+$scriptName+".ps1_FullLog_$(Get-Date -Format "yyMMdd").log"
$errorLogPathAndName = $logFileLocation+$scriptName+".ps1_ErrorLog_$(Get-Date -Format "yyMMdd").log"
if($PSCommandPath){
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_$whatToSync`_Transcript_$(Get-Date -Format "yyMMdd").log"
    Start-Transcript $transcriptLogName -Append
    }


#/Ignore if running manually
$userAdmin = "groupbot@anthesisgroup.com"
#convertTo-localisedSecureString "IntuneAdminPasswordHere"
$userAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Downloads\GroupBot.txt) 
#$adminCreds = set-MsolCredentials -username $intuneAdmin
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $userAdmin, $userAdminPass
#/

#Create a secure credential object
#$adminCreds = get-credential
#Connect to services by passing in secure credential object we set above
connect-ToMsol -Interactive
connect-toAAD -Interactive
connect-ToExo -Interactive

#Create an array of email address so we can iterate through them
$upnsToDeactivate = convertTo-arrayOfEmailAddresses $upnsString

#region deprovision - now go through each email address in the array and deactivate them
foreach($user in $upnsToDeactivate){

    if($user){
        $userExoObject = Get-User -Identity $user
        $userAadObject = Get-AzureADUser -SearchString $user.Replace("@anthesisgroup.com","")
        $userMsolObject = Get-MsolUser -UserPrincipalName $user
        if($userExoObject.DistinguishedName -ne $null){
            write-host "Disabling $($userExoObject.DisplayName)"
            Set-MsolUser -UserPrincipalName $userExoObject.UserPrincipalName -BlockCredential $true
            Set-MsolUserPassword -UserPrincipalName $userExoObject.UserPrincipalName -NewPassword "TTFN123!" -ForceChangePassword $true
            Get-DistributionGroup -Filter "Members -eq '$($userExoObject.DistinguishedName)'" | % {
                Remove-DistributionGroupMember -Identity $_.Id -Member $userExoObject.UserPrincipalName -Confirm:$false -BypassSecurityGroupManagerCheck:$true
                }
            Set-Mailbox $userExoObject.UserPrincipalName -HiddenFromAddressListsEnabled $true -Type Shared
            if($userExoObject.DisplayName.Substring(0,1) -ne "Ω"){Set-MsolUser -UserPrincipalName $userExoObject.UserPrincipalName -DisplayName $("Ω_"+$userExoObject.DisplayName)}
            Set-MsolUserLicense -UserPrincipalName $($userExoObject.UserPrincipalName) -RemoveLicenses $userMsolObject.Licenses.accountskuid
            Revoke-AzureADUserAllRefreshToken -ObjectId $userAadObject.ObjectId
            #Initiate Retire on Intune devices
            #
            # ?
            #get line reports and save to clipboard to send in email (we can't define who this needs to be sent to per region as this is an ad-hoc request process at the moment)
            $userLineReports = Get-AzureADUserDirectReport -ObjectId $userAadObject.ObjectId -All $true
            Write-Host "User's line reports are (we've copied to your clipboard):" -ForegroundColor Yellow
            $userLineReports.userPrincipalName
            Set-Clipboard -Value $userLineReports.userPrincipalName
            }
        }
    }
#-InactiveMailbox 


Disconnect-ExchangeOnline -Confirm:$false



<#

$usersToReassign = @{}

$sharePointServerUrl = "https://anthesisllc.sharepoint.com"
$hrSite = "/teams/hr"
$leavingUserListName = "Leaving User Requests"
$oDataUnprocessedUsers = '$filter=Status ne ''Completed'''
$oDataUnprocessedUsers = '$select=User_x0020_to_x0020_deprovision/Name,User_x0020_to_x0020_deprovision/Title,Last_x0020_day_x0020_of_x0020_em,Mailbox_x0020_action,E_x002d_mail_x0020_address_x0020,Reassign_x0020_e_x002d_mail_x002/Name,Reassign_x0020_e_x002d_mail_x002/Title,Additional_x0020_e_x002d_mail_x0,Title&$expand=User_x0020_to_x0020_deprovision/Id,Reassign_x0020_e_x002d_mail_x002/Id'
#$unprocessedLeavers = get-itemsInList -serverUrl $sharePointServerUrl -sitePath $hrSite -listName $leavingUserListName -oDataQuery $oDataUnprocessedUsers -suppressProgress $false
$unprocessedLeavers | %{
    $leavingUser = New-Object -TypeName PSObject
    $leavingUser | Add-Member -MemberType NoteProperty -Name "LeavingUserId" -Value $_.User_x0020_to_x0020_deprovision.Name.Replace("i:0#.f|membership|","")
    $leavingUser | Add-Member -MemberType NoteProperty -Name "LeavingUserName" -Value $_.User_x0020_to_x0020_deprovision.Title
    $leavingUser | Add-Member -MemberType NoteProperty -Name "LeavingDate" -Value $_.Last_x0020_day_x0020_of_x0020_em
    $leavingUser | Add-Member -MemberType NoteProperty -Name "MailboxAction" -Value $_.Mailbox_x0020_action
    $leavingUser | Add-Member -MemberType NoteProperty -Name "UpnAction" -Value $_.E_x002d_mail_x0020_address_x0020
    if($_.Reassign_x0020_e_x002d_mail_x002.__deferred -eq $null){
        $leavingUser | Add-Member -MemberType NoteProperty -Name "RedirectToId" -Value $_.Reassign_x0020_e_x002d_mail_x002.Name.Replace("i:0#.f|membership|","")
        $leavingUser | Add-Member -MemberType NoteProperty -Name "RedirectToName" -Value $_.Reassign_x0020_e_x002d_mail_x002.Title
        }
        else{
            $leavingUser | Add-Member -MemberType NoteProperty -Name "RedirectToId" -Value ""
            $leavingUser | Add-Member -MemberType NoteProperty -Name "RedirectToName" -Value ""
            }
    $leavingUser | Add-Member -MemberType NoteProperty -Name "AliasAction" -Value $_.Additional_x0020_e_x002d_mail_x0
    $leavingUser | Add-Member -MemberType NoteProperty -Name "AdditionalDetails" -Value $_.Title
    $unprocessedLeaversFormatted += $leavingUser
    }

$selectedLeavers = $unprocessedLeaversFormatted | Out-GridView -PassThru
$usersToDeprovision = $selectedLeavers | ?{$_.UpnAction -ne "Reassign to another user"}
$selectedLeavers | ?{$_.UpnAction -eq "Reassign to another user"} | % {$usersToReassign.Add($_.LeavingUserId.Split("@")[0],$_.RedirectToId.Split("@")[0])}

$sqlConnection = connect-toSqlServer -SQLServer "sql.sustain.co.uk" -SQLDBName "SUSTAIN_LIVE" #This is required to disable ARENA accounts






#@("ali.mahdavi","katie.swain","simon.white","laura.sponti","sion.fenwick","ben.buffery","laura.pugh","tilly.shaw","catherine.green") | % {
$binMe | %{
    $user = $_
    $u = Get-User -Identity $user@anthesisgroup.com
    Get-DistributionGroup -Filter "Members -eq '$($u.DistinguishedName)'" | % {
        Remove-DistributionGroupMember -Identity $_.Id -Member $user@anthesisgroup.com -Confirm:$false
        }
    Set-Mailbox $user -HiddenFromAddressListsEnabled $true -InactiveMailbox
    }


foreach ($userSAM in $usersToDeprovision){
    Write-Host -ForegroundColor Yellow "Deprovisioning $userSAM"
    deprovision-user -userSAM $userSAM -plaintextPassword $plaintextPassword -exportAdmin $exportAdmin -reassignEmailAddressesTo $null
    }
foreach($userSAM in $usersToReassign.Keys){
    Write-Host -ForegroundColor Yellow "Deprovisioning $userSAM"
    deprovision-user -userSAM $userSAM -plaintextPassword $plaintextPassword -exportAdmin $exportAdmin -reassignEmailAddressesTo $usersToReassign[$userSAM] 
    }
#endregion
#region Unlicense
foreach ($userSAM in $usersToDeprovision){
    Write-Host -ForegroundColor Yellow "Unlicensing $userSAM"
    $newUserSAM = (get-msolUpn -userSAM $userSAM).Replace("@anthesisgroup.com","")
    try{remove-msolLicenses -userSAM $newUserSAM;Write-Host -ForegroundColor DarkYellow "MSOL License removal completed successfully"}
    catch{Write-Host -ForegroundColor Red "Failed to remove MSOL Licenses"}
    }
foreach($userSAM in $usersToReassign.Keys){
    Write-Host -ForegroundColor Yellow "Unlicensing $userSAM"
    $newUserSAM = (get-msolUpn -userSAM $userSAM).Replace("@anthesisgroup.com","")
    try{remove-msolLicenses -userSAM $newUserSAM;Write-Host -ForegroundColor DarkYellow "MSOL License removal completed successfully"}
    catch{Write-Host -ForegroundColor Red "Failed to remove MSOL Licenses"}
    }
#endregion
#region Delete
foreach ($userSAM in $usersToDeprovision){
    Write-Host -ForegroundColor Yellow "Deleting $userSAM"
    try{delete-userAccounts -userSAM $userSAM;Write-Host -ForegroundColor DarkYellow "Deleting $userSAM"}
    catch{Write-Host -ForegroundColor Red "Failed to delete user account for $userSAM";$error[0]}
    }
foreach($userSAM in $usersToReassign.Keys){
    Write-Host -ForegroundColor Yellow "Deleting $userSAM"
    try{delete-userAccounts -userSAM $userSAM;Write-Host -ForegroundColor DarkYellow "Deleting $userSAM"}
    catch{Write-Host -ForegroundColor Red "Failed to delete user account for $userSAM"}
    }
#endregion

$sqlConnection.Close()

#Remove-MailboxPermission -Identity $userSAM -AccessRights FullAccess -user $exportAdmin 
#>