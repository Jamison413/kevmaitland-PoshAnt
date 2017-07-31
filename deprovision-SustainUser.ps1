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
function connect-toSqlServer($SQLServer,$SQLDBName){
    #SQL Server connection string
    $connDB = New-Object System.Data.SqlClient.SqlConnection
    $connDB.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True" #This relies on the current user having the appropriate Login/Role Membership ont he DB
    $connDB.Open()
    $connDB
    }
function Execute-SQLQueryOnSQLDB([string]$query, [string]$queryType, $sqlServerConnection) { 
  # NonQuery - Insert/Update/Delete query where no return data is required
    $sql = New-Object System.Data.SqlClient.SqlCommand
    $sql.Connection = $sqlServerConnection
    $sql.CommandText = $query
    switch ($queryType){
        "NonQuery" {$sql.ExecuteNonQuery()}
        "Scalar" {$sql.ExecuteScalar()}
        "Reader" {    
            $oReader = $sql.ExecuteReader()
            $results = @()
            while ($oReader.Read()){
                $result = New-Object PSObject
                for ($i = 0; $oReader.FieldCount -gt $i; $i++){
                        $columnName = ($query.Replace(",","") -split '\s+')[$i+1]
                        if (1 -lt $columnName.Split(".").Length){$columnName = $columnName.Split(".")[1]} #Trim off any table names
                        $result | Add-Member NoteProperty $columnName $oReader[$i]
                        }
                 $results += $result
                }
            $oReader.Close()
            return $results
            }
        }
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
    }#endregion

$credential = get-credential -Credential kevin.maitland@anthesisgroup.com
Import-Module MSOnline
Connect-MsolService -Credential $credential
$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
Import-PSSession $ExchangeSession

$exportAdmin = "kevin.maitland@anthesisgroup.com"
$usersToDeprovision = @('Colette.Ford') 
$usersToReassign = @{'Richard.Hopkins'='Tilly.Shaw'}
    $plaintextPassword = "Ttfn123!"


    $sqlConnection = connect-toSqlServer -SQLServer "sql.sustain.co.uk" -SQLDBName "SUSTAIN_LIVE" #This is required to disable ARENA accounts
    #region deprovision
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