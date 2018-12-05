$GAUsername = "kevin.maitland@anthesisgroup.com"
$password = Read-Host -Prompt "Password for $GAUsername" -AsSecureString
$outputFile = "$env:USERPROFILE\Desktop\MsolAdminRoles_$(Get-Date -Format "yyyy-MM-dd").csv"

$credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $GAUsername, $password
Import-Module MSOnline
Connect-MsolService -Credential $credential

$roles = Get-MsolRole
$roleDataHash = @{}
$roleUsers = @()
foreach ($role in $roles){
    $roleDataHash.Add($role.Name,@())
    Get-MsolRoleMember -RoleObjectId $role.ObjectId | %{
        $roleDataHash[$role.Name] += @($_.DisplayName+",".Trim())
        $roleUsers += $_.EmailAddress
        }
    }
#    "$($role.Name)`t$($role.Description)" | Add-Content -Path $outputFile
#    select Displayname,EmailAddress,IsLicensed,RoleMemeberType,ValidationStatus | Add-Content -Path $outputFile 
#    "" | Add-Content -Path $outputFile

$roleDataHash.Keys | %{$_+","+$roleDataHash[$_] | Add-Content -Path $outputFile}

$roleUsers = $roleUsers | select -Unique | sort DisplayName
$msolUsers = Get-MsolUser -All | ?{$roleUsers -contains $_.UserPrincipalName}
"" | Add-Content -Path $outputFile
"DisplayName,UPN,IsLicensed,MFAState,MFADefault" | Add-Content -Path $outputFile
$msolUsers | %{$_.DisplayName+","+$_.UserPrincipalName+","+$_.IsLicensed+","+$_.LastPasswordChangeTimestamp,$_.StrongAuthenticationRequirements.State+","+$($msolUsers[0].StrongAuthenticationMethods | ? {$_.IsDefault -eq $true}).MethodType | Add-Content -Path $outputFile} 
