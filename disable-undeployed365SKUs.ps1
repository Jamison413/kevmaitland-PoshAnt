
$arrayOfUserstoIgnore = @("kevin.maitland","kirsten.doddy","elle.smith","sophie.taylor","emily.pressey")
$arrayOfUnusedServices = @("TEAMS1","YAMMER_ENTERPRISE") 

#Get All Licensed Users
$users = Get-MsolUser -All | Where-Object {$_.isLicensed -eq $true}

foreach ($user in $users){
    if(!($arrayOfUserstoIgnore -contains $user.UserPrincipalName.Replace("@anthesisgroup.com",""))){
	    foreach($license in $user.Licenses){
            $toDisable = @()
            [array]$alreadyDisabled = $($license.ServiceStatus | ?{$_.ProvisioningStatus -eq "Disabled"}).ServicePlan.ServiceName

            $arrayOfUnusedServices | % {
                $thisLicense = $license.ServiceStatus | ?{$_.ServicePlan.ServiceName -eq $_}
                switch ($thisLicense.ProvisioningStatus) {
                    "Success" {
                        $toDisable += "$_"
                        Write-Host -ForegroundColor Yellow "Disabling $_ for $($user.DisplayName)"
                        break
                        }
                    {![string]::IsNullOrWhiteSpace($_)} {
                        Write-Host -ForegroundColor DarkYellow "$_ already disabled for $($user.DisplayName)"
                        break;
                        }
                    #default {Write-Host -ForegroundColor DarkYellow "TEAMS1 not foundfor $($user.DisplayName)"}
                    }
                }
            $teamsLicense = $license.ServiceStatus | ?{$_.ServicePlan.ServiceName -eq "TEAMS1"}


            $yammerLicense = $license.ServiceStatus | ?{$_.ServicePlan.ServiceName -eq "YAMMER_ENTERPRISE"}
            switch ($yammerLicense.ProvisioningStatus) {
                "Success" {
                    $toDisable += "YAMMER_ENTERPRISE"
                    Write-Host -ForegroundColor Yellow "Disabling YAMMER_ENTERPRISE for $($user.DisplayName)"
                    break
                    }
                {![string]::IsNullOrWhiteSpace($_)} {
                    Write-Host -ForegroundColor DarkYellow "YAMMER_ENTERPRISE already disabled for $($user.DisplayName)"
                    break;
                    }
                #default {Write-Host -ForegroundColor DarkYellow "YAMMER_ENTERPRISE not foundfor $($user.DisplayName)"}
                }

            #Now disable anything that needs disabling
            if ($toDisable.Count -gt 0){
                [array]$disableThese = $toDisable
                if ($alreadyDisabled.Count -gt 0) {$alreadyDisabled | % {$disableThese += $_}}
                Write-Host -ForegroundColor Magenta "New-MsolLicenseOptions -AccountSkuId $($license.AccountSkuid) -DisabledPlans $($disableThese -join ",")"
                $NewSkU = New-MsolLicenseOptions -AccountSkuId $license.AccountSkuid -DisabledPlans $disableThese
                Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -LicenseOptions $NewSkU
                }

            }
        }
    }