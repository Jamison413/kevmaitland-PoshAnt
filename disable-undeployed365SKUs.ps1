﻿$365creds = set-MsolCredentials
connect-to365 -credential $365creds


$teamBotDetails = import-encryptedCsv -pathToEncryptedCsv "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\teambotdetails.txt"
$tokenResponse = get-graphTokenResponse -aadAppCreds $teamBotDetails

$teamsPilotGroup = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "groups/?`$filter=mail+eq+'teamspilot@anthesisgroup.com'"
$teamsPilotUsers = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "groups/$($teamsPilotGroup.id)/transitiveMembers?`$select=id,displayName,jobTitle,mail,userPrincipalName,assignedLicenses"
$teamsPilotUPNs = $teamsPilotUsers | ? {$_.'@odata.type' -eq "#microsoft.graph.user" -and $_.assignedLicenses.Count -gt 0} | select userPrincipalName -Unique | Sort-Object userPrincipalName | % {$_.userPrincipalName}

$allTeamsGroup = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "groups/?`$filter=mail+eq+'teamsusers@anthesisgroup.com'"
$allTeamsUsers = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "groups/$($allTeamsGroup.id)/transitiveMembers?`$select=id,displayName,jobTitle,mail,userPrincipalName,assignedLicenses"
$allTeamsUPNs = $allTeamsUsers | ? {$_.'@odata.type' -eq "#microsoft.graph.user" -and $_.assignedLicenses.Count -gt 0} | select userPrincipalName -Unique | Sort-Object userPrincipalName | % {$_.userPrincipalName}
Write-host -fore Yellow "[$($teamsPilotUPNs.Count)] Audio Conferencing licenses required"

#$dg = Get-DistributionGroup -Identity teamspilot@anthesisgroup.com
#$teamsPilotUsers = $(enumerate-nestedDistributionGroups -distributionGroupObject $dg -Verbose).WindowsLiveId

#$dg = Get-DistributionGroup -Identity teamsusers@anthesisgroup.com
#$allTeamsUsers = $(enumerate-nestedDistributionGroups -distributionGroupObject $dg -Verbose).WindowsLiveId


$teamsPilotUsersDesiredState = [ordered]@{"TEAMS1"="Success";"YAMMER_ENTERPRISE"="Disabled";"AnthesisLLC:MCOMEETADV"="Add"} 
$teamsUsersDesiredState = [ordered]@{"TEAMS1"="Success";"YAMMER_ENTERPRISE"="Disabled"}
$mostUsersDesiredState = [ordered]@{"TEAMS1"="Success";"YAMMER_ENTERPRISE"="Disabled"}


#Get All Licensed Users
$users = Get-MsolUser -All | Where-Object {$_.isLicensed -eq $true}
#$allUsers = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "users?`$select=id,displayName,jobTitle,mail,userPrincipalName,assignedLicenses"
#$licensedUsers = $allUsers | ? {$_.assignedLicenses.Count -gt 0}

foreach ($user in $users){
    #Set the apprporiate DesiredState
    if($teamsPilotUPNs -contains $user.UserPrincipalName){$desiredState = $teamsPilotUsersDesiredState}
    elseif($allTeamsUPNs -contains $user.UserPrincipalName){$desiredState = $teamsUsersDesiredState}
    else{$desiredState = $mostUsersDesiredState}

    #Add/remove any licenses before we check individual services
    $licensesToAdd = @()
    $licensesToAdd += $($desiredState.GetEnumerator() | ? {$_.Name -match "AnthesisLLC:" -and $_.Value -eq "Add"}).Name | ? {$_ -ne $null}
    $licensesToAdd = Compare-Object -ReferenceObject $licensesToAdd -DifferenceObject $user.Licenses.AccountSkuId -PassThru | ? {$_.SideIndicator -eq "<="} #Prevent re-adding licenses unnecessarily
    #$licensesToAdd = Compare-Object -ReferenceObject $licensesToAdd -DifferenceObject $user.assignedLicenses.GetEnumerator().skuId -PassThru | ? {$_.SideIndicator -eq "<="} #Prevent re-adding licenses unnecessarily
    $licensesToRemove = @()
    $licensesToRemove += $($desiredState.GetEnumerator() | ? {$_.Name -match "AnthesisLLC:" -and $_.Value -eq "Remove"}).Name | ? {$_ -ne $null}
    $licensesToRemove = Compare-Object -ReferenceObject $licensesToRemove -DifferenceObject  $user.Licenses.AccountSkuId -PassThru -IncludeEqual | ? {$_.SideIndicator -eq "=="} #Prevent attempt to remove license that have already been removed
    
    if($licensesToAdd.Count -gt 0 -or $licenseToRemove.Count -gt 0){
        Write-Host -ForegroundColor Yellow "Set-MsolUserLicense -UserPrincipalName $($user.UserPrincipalName) -AddLicenses [$($licensesToAdd -join ", ")] -RemoveLicenses [$($licensesToRemove -join ", ")]"
        Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -AddLicenses $licensesToAdd -RemoveLicenses $licenseToRemove
        #Refresh the $user object if we've changed the licenses
        $user = Get-MsolUser -UserPrincipalName $user.UserPrincipalName
        }

    #Check/set the status for individual services within each license
	foreach($license in $user.Licenses){
        #Note any services that are disabled (so we don't accidentally re-enable them later if we)
        $alreadyDisabled = @()
        $alreadyDisabled += $($license.ServiceStatus | ?{$_.ProvisioningStatus -eq "Disabled"}).ServicePlan.ServiceName | ? {$_ -ne $null}

        #Figure out if anthing is not in the desired state
        $toDisable = @()
        $toEnable = @()
        $desiredState.Keys | % {
            $thisService = $_
            $thisServiceCurrentStatus = $($license.ServiceStatus | ?{$_.ServicePlan.ServiceName -eq $thisService}).ProvisioningStatus
            if($thisServiceCurrentStatus -ne $null -and $thisServiceCurrentStatus -ne $desiredState[$thisService]){
                switch($desiredState[$_]){
                    "Disabled" {$toDisable += $thisService}
                    "Success"  {$toEnable += $thisService}
                    }
                }
            }

        #If anything needs changing, change it.
        if($toDisable.Count -gt 0 -or $toEnable.Count -gt 0){
            #Figure out the final list of Services that should be disabled by adding anything that was already disabled (but isn't in the enable list) to the array of services that we've specifically identifying for disabling
            Compare-Object -ReferenceObject $toEnable  -DifferenceObject $alreadyDisabled | ? {$_.SideIndicator -eq "=>"} | % {$toDisable += $_.InputObject}
            $correctlyConfiguredSku = New-MsolLicenseOptions -AccountSkuId $license.AccountSkuid -DisabledPlans $toDisable
            Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -LicenseOptions $correctlyConfiguredSku
            Write-host "Set-MsolUserLicense -UserPrincipalName $($user.UserPrincipalName) -LicenseOptions [$($correctlyConfiguredSku.AccountSkuId.SkuPartNumber)] : [$($correctlyConfiguredSku.DisabledServicePlans -join ", ")]"
            }

        }

    
    }

        
