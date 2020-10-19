$tokenResponseTeamsBot = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName TeamsBot) -grant_type client_credentials
$ukUsers = get-graphUsersWithEmployeeInfoExtensions -tokenResponse $tokenResponseTeamsBot -filterBusinessUnit 'Anthesis (UK) Ltd (GBR)' -filterNone
$ukUsers += get-graphUsersWithEmployeeInfoExtensions -tokenResponse $tokenResponseTeamsBot -filterBusinessUnit 'Anthesis Energy UK Ltd (GBR)' -filterNone
#$ukAadDevices = get-graphDevices -tokenResponse $tokenResponseTeamsBot -filterOwnerIds $ukUsers.id
#$ukAadDevices = $ukAadDevices | Group-Object displayName | % {$_.Group | Sort-Object approximateLastSignInDateTime | Select-Object -Last 1} #DeDupe and keep only the most recent

$tokenResponseIntuneBot = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName IntuneBot) -grant_type client_credentials 
$intuneDevices = get-graphIntuneDevices -tokenResponse $tokenResponseIntuneBot

$tokenResponseSharePointBot = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName SharePointBot) -grant_type client_credentials
$itTeamAllSite = get-graphSite -tokenResponse $tokenResponseSharePointBot -serverRelativeUrl "/teams/IT_Team_All_365"
$assetRegister = get-graphList -tokenResponse $tokenResponseSharePointBot -graphSiteId $itTeamAllSite.id -listName "Anthesis IT Asset Register"
$assetRegisterItems = get-graphListItems -tokenResponse $tokenResponseSharePointBot -graphSiteId $itTeamAllSite.id -listId $assetRegister.id -expandAllFields #$assetRegisterItems.fields.AssetStatus | select -Unique
$assetRegisterComputers = $assetRegisterItems | ? {$_.fields.ContentType -eq "Computers"}
$assetRegisterPhones = $assetRegisterItems | ? {$_.fields.ContentType -eq "Mobiles"}

$allAadDevices = get-graphDevices -tokenResponse $tokenResponseTeamsBot
$allAadDevices | % { #Gather Intune & Asset data for AAD devices
    $thisAadDevice = $_
    if($correspondingIntuneDevice){rv correspondingIntuneDevice}
    if($correspondingAsset){rv correspondingAsset}
    Write-Host "Processing [$($thisAadDevice.displayName)][$($thisAadDevice.deviceId)]"
    $correspondingIntuneDevice = $intuneDevices | ? {$_.azureADDeviceId -eq $thisAadDevice.deviceId}
    if($correspondingIntuneDevice){
        Write-Host "`tAdding Intune information to [$($thisAadDevice.displayName)][$($thisAadDevice.deviceId)]"
        $intuneHash = @{}
        Get-Member -InputObject $correspondingIntuneDevice -MemberType Properties | % {
            $intuneHash.Add($_.Name, $correspondingIntuneDevice.$($_.Name))
            }
        $_ | Add-Member -MemberType NoteProperty -Name intune -Value $intuneHash -Force

        #Then try matching the Asset using the serial number
        $correspondingAsset = $assetRegisterItems | ? {$_.fields.ManufacturerSerialNumber -eq $correspondingIntuneDevice.serialNumber}
        if(!$correspondingAsset){
            $correspondingAsset = $assetRegisterComputers | ? {$_.fields.IT_x0020_Product_x0020_Code -eq $correspondingIntuneDevice.serialNumber}
            if(!$correspondingAsset){
                $correspondingAsset = $assetRegisterComputers | ? {![string]::IsNullOrWhiteSpace($_.fields.MACAddresses)} | ? {$_.fields.MACAddresses.Replace(":","").Replace("-","") -match $correspondingIntuneDevice.wiFiMacAddress}
                if(!$correspondingAsset){
                    $correspondingAsset = $assetRegisterComputers | ? {$_.fields.ComputerName -eq $thisAadDevice.displayName}
                    if(!$correspondingAsset){
                        $correspondingAsset = $assetRegisterPhones | ? {$_.fields.IMEI -eq $correspondingIntuneDevice.imei}
                        if(!$correspondingAsset){}
                        else{Write-Host "`t`tAsset matched by IMEI"}
                        }
                    else{Write-Host "`t`tAsset matched by ComputerName"}
                    }
                else{Write-Host "`t`tAsset matched by MACAddresses"}
                }
            else{Write-Host "`t`tAsset matched by IT_x0020_Product_x0020_Code"}
            }
        else{Write-Host "`t`tAsset matched by ManufacturerSerialNumber"}

        }
    else{
        Write-Warning "No Intune device found for [$($thisAadDevice.displayName)][$($thisAadDevice.deviceId)]"
        [array]$notInIntune += $thisAadDevice
        }

    if($correspondingAsset){
        Write-Host "`tAdding Asset information to [$($thisAadDevice.displayName)][$($thisAadDevice.deviceId)]"
        $assetHash = @{}
        Get-Member -InputObject $correspondingAsset.fields -MemberType Properties | % {
            $assetHash.Add($_.Name, $correspondingAsset.fields.$($_.Name))
            }
        $_ | Add-Member -MemberType NoteProperty -Name asset -Value $assetHash -Force
        }

    }
$ukAadDevices = $allAadDevices | ? {$_.intune.userId -match $($ukUsers.id -join "|")}
$ukAadDevices = $ukAadDevices | Group-Object {$_.intune.serialNumber} | % {$_.Group | Sort-Object approximateLastSignInDateTime | Select-Object -Last 1} #DeDupe and keep only the most recent

#Find any assets that aren;t in AAD yet.
#invoke-graphGet -tokenResponse $tokenResponseSharePointBot -graphQuery '/sites/anthesisllc.sharepoint.com,1ead4d00-2e2e-421e-9176-15287b5c22ce,d21ddf81-fcef-4e36-94e6-edd6fb487a31/lists/79768c18-98dd-4fef-b0cb-fc0ceef84a77/items/127?$expand=fields(select=Author)' -Verbose -useBetaEndpoint
$extraAssets = @()
$assetRegisterComputersUseful = $assetRegisterComputers | ? {$_.fields.AssetStatus -notmatch "Dispose" -and $_.fields.AssetStatus -notmatch "Retire"}
$assetRegisterComputersUseful | % {
    $thisAsset = $_
    if($matchedDevice){rv matchedDevice}
    if(![string]::IsNullOrWhiteSpace($thisAsset.fields.ManufacturerSerialNumber)){
        $matchedDevice = $ukAadDevices | ? {$_.asset.ManufacturerSerialNumber -eq $thisAsset.fields.ManufacturerSerialNumber}
        
        if(!$matchedDevice){
            $extraAssets += $_
            write-host "`t[$($thisAsset.fields.ComputerName)] matched to [$($matchedDevice.displayName)]" -f Yellow
            }
        else{write-host "[$($thisAsset.fields.ComputerName)] matched to [$($matchedDevice.displayName)]"}
        }
    }
$assetRegisterPhonesUseful = $assetRegisterPhones | ? {$_.fields.AssetStatus -notmatch "Dispose" -and $_.fields.AssetStatus -notmatch "Retire" -and $_.fields.AssetStatus -notmatch "Broke"}
$assetRegisterPhonesUseful | % {
    $thisAsset = $_
    if($matchedDevice){rv matchedDevice}
    if(![string]::IsNullOrWhiteSpace($thisAsset.fields.ManufacturerSerialNumber)){
        $matchedDevice = $ukAadDevices | ? {$_.asset.IMEI -eq $thisAsset.fields.IMEI}
        
        if(!$matchedDevice){
            $extraAssets += $_
            write-host "`t[$($thisAsset.fields.ComputerName)] matched to [$($matchedDevice.displayName)]" -f Yellow
            }
        else{write-host "[$($thisAsset.fields.ComputerName)] matched to [$($matchedDevice.displayName)]"}
        }
    }

#Make the final objects for exporting
$prettyObjects = @()
$ukAadDevices | % {
    $thisAadDevice = $_
    $userId = $thisAadDevice.physicalIds | ? {$_ -match "USER-GID"} | % {$($_ -split ":")[1]}
    if($userId){$user = $ukUsers | ? {$_.id -eq $userId}}
    else{$user = $ukUsers | ? {$_.userPrincipalName -eq $thisAadDevice.intune.userPrincipalName}}
    
    $prettyObject = New-Object -TypeName psobject -Property $([ordered]@{
        AssetTag = $thisAadDevice.asset.AnthesisSerialNumber
        AssetType = $thisAadDevice.asset.ContentType
        Ownership = $thisAadDevice.intune.managedDeviceOwnerType
        AssetBusinessUnit = $thisAadDevice.asset.Business_x0020_Unit.Label
        AssetSupplier = $thisAadDevice.asset.AssetSupplier
        AssetCost = $thisAadDevice.asset.AssetPriceAtPurchase
        AssetPO = $thisAadDevice.asset.AssetPO
        AssetPurchaseDate = $(get-date $thisAadDevice.asset.InvoiceDate -f U)
        AssetStatus = $thisAadDevice.asset.AssetStatus
        DeviceNameAAD=$thisAadDevice.displayName
        DeviceNameAsset = $thisAadDevice.asset.ComputerName
        DeviceManufacturer = $thisAadDevice.manufacturer
        DeviceModel = $thisAadDevice.model
        DeviceSerialNumber = $thisAadDevice.intune.serialNumber
        DeviceWifiMacAddress = $thisAadDevice.intune.wiFiMacAddress
        OperatingSystem = $thisAadDevice.operatingSystem
        OperatingSystemVersion = $thisAadDevice.operatingSystemVersion
        DeviceCompliance = $thisAadDevice.intune.complianceState
        DeviceEncryption = $thisAadDevice.intune.isEncrypted
        DeviceEnrollmentType = $thisAadDevice.intune.deviceEnrollmentType
        DeviceTrustType = $thisAadDevice.trustType
        DeviceProfileType = $thisAadDevice.profileType
        DevicePhoneNumber = $thisAadDevice.intune.phoneNumber
        DeviceNetworkCarrier = $thisAadDevice.intune.subscriberCarrier
        DeviceNetworkImei = $thisAadDevice.intune.imei
        LastUser = $user.userPrincipalName
        LastUserBusinessUnit = $user.anthesisgroup_employeeInfo.businessUnit
        LastSignInDateTime = $(get-date $thisAadDevice.approximateLastSignInDateTime -f U)
        InAAD = $true
        })
    if(![string]::IsNullOrWhiteSpace($thisAadDevice.intune)){$prettyObject | Add-Member -MemberType NoteProperty -Name InIntune -Value $true}
    else{$prettyObject | Add-Member -MemberType NoteProperty -Name InIntune -Value $false}
    if(![string]::IsNullOrWhiteSpace($thisAadDevice.asset)){$prettyObject | Add-Member -MemberType NoteProperty -Name InAsset -Value $true}
    else{$prettyObject | Add-Member -MemberType NoteProperty -Name InAsset -Value $false}
    if($thisAadDevice.physicalIds -match "ZTDID"){$prettyObject | Add-Member -MemberType NoteProperty -Name InAutoPilot -Value $true}
    else{$prettyObject | Add-Member -MemberType NoteProperty -Name InAutoPilot -Value $false}

    [array]$prettyObjects += $prettyObject
    }
$extraAssets | % {
    $thisAsset = $_
    $userId = $thisAsset.physicalIds | ? {$_ -match "USER-GID"} | % {$($_ -split ":")[1]}
    $user = $ukUsers | ? {$_.id -eq $userId}
    $prettyObject = New-Object -TypeName psobject -Property $([ordered]@{
        AssetTag = $thisAsset.fields.AnthesisSerialNumber
        AssetType = $thisAsset.fields.ContentType
        Ownership = $thisAsset.intune.managedDeviceOwnerType
        AssetBusinessUnit = $thisAsset.fields.Business_x0020_Unit.Label
        AssetSupplier = $thisAsset.fields.AssetSupplier
        AssetCost = $thisAsset.fields.AssetPriceAtPurchase
        AssetPO = $thisAsset.fields.AssetPO
        AssetPurchaseDate = $(get-date $thisAsset.fields.InvoiceDate -f U)
        AssetStatus = $thisAsset.fields.AssetStatus
        DeviceNameAAD=$thisAsset.displayName
        DeviceNameAsset = $thisAsset.fields.ComputerName
        DeviceManufacturer = $thisAsset.fields.Manufacturer
        DeviceModel = $thisAsset.model
        DeviceSerialNumber = $thisAsset.intune.serialNumber
        DeviceWifiMacAddress = $thisAsset.intune.wiFiMacAddress
        OperatingSystem = $thisAsset.operatingSystem
        OperatingSystemVersion = $thisAsset.operatingSystemVersion
        DeviceCompliance = $thisAsset.intune.complianceState
        DeviceEncryption = $thisAsset.intune.isEncrypted
        DeviceEnrollmentType = $thisAsset.intune.deviceEnrollmentType
        DeviceTrustType = $thisAsset.trustType
        DeviceProfileType = $thisAsset.profileType
        DevicePhoneNumber = "N/A"
        DeviceNetworkCarrier = "N/A"
        DeviceNetworkImei = $thisAsset.fields.IMEI
        LastUser = $user.userPrincipalName
        LastUserBusinessUnit = $user.anthesisgroup_employeeInfo.businessUnit
        LastSignInDateTime = $thisAsset.approximateLastSignInDateTime
        InAsset = $true
        InAAD = $false
        InIntune = $false
        InAutoPilot = "Unknown"
        })
    [array]$prettyObjects += $prettyObject
    }

$prettyObjects | ? {$_.InIntune -eq $true -or $_.InAsset -eq $true -or $_.InAutoPilot -eq $true} | Export-Csv  -Path 'C:\Users\KevMaitland\OneDrive - Anthesis LLC\Desktop\Hardware2.csv' -NoTypeInformation 


$prettyObjects = @()
$ukAadDevices | ? {$_.intune.managedDeviceOwnerType -ne "personal" -and $_.intune.managedDeviceOwnerType -ne "unknown"} | Group-Object Manufacturer, Model, OperatingSystem, OperatingSystemVersion | % {
    $thisGroup = $_
    $prettyObject = New-Object -TypeName psobject -Property $([ordered]@{
        Manufacturer = $thisGroup.Name.Split(",")[0]
        Model = $thisGroup.Name.Split(",")[1]
        OperatingSystem = $thisGroup.Name.Split(",")[2]
        OperatingSystemVersion = $thisGroup.Name.Split(",")[3]
        Count = $thisGroup.Count
        })
    $prettyObjects += $prettyObject
    }

$prettyObjects | Sort-Object Manufacturer, Model, OperatingSystem, OperatingSystemVersion |  Export-Csv -Path 'C:\Users\KevMaitland\OneDrive - Anthesis LLC\Desktop\Hardware_Simplified.csv' -NoTypeInformation 




$prettyObjects = @()
$ukAadDevices | ? {$_.Ownership -ne "personal"} | % {
    $thisAadDevice = $_
    $userId = $thisAadDevice.physicalIds | ? {$_ -match "USER-GID"} | % {$($_ -split ":")[1]}
    if($userId){$user = $ukUsers | ? {$_.id -eq $userId}}
    else{$user = $ukUsers | ? {$_.userPrincipalName -eq $thisAadDevice.intune.userPrincipalName}}
    
    $prettyObject = New-Object -TypeName psobject -Property $([ordered]@{
        AssetTag = $thisAadDevice.asset.AnthesisSerialNumber
        AssetType = $thisAadDevice.asset.ContentType
        Ownership = $thisAadDevice.intune.managedDeviceOwnerType
        AssetBusinessUnit = $thisAadDevice.asset.Business_x0020_Unit.Label
        AssetSupplier = $thisAadDevice.asset.AssetSupplier
        AssetCost = $thisAadDevice.asset.AssetPriceAtPurchase
        AssetPO = $thisAadDevice.asset.AssetPO
        AssetPurchaseDate = $(get-date $thisAadDevice.asset.InvoiceDate -f U)
        AssetStatus = $thisAadDevice.asset.AssetStatus
        DeviceNameAAD=$thisAadDevice.displayName
        DeviceNameAsset = $thisAadDevice.asset.ComputerName
        DeviceManufacturer = $thisAadDevice.manufacturer
        DeviceModel = $thisAadDevice.model
        DeviceSerialNumber = $thisAadDevice.intune.serialNumber
        DeviceWifiMacAddress = $thisAadDevice.intune.wiFiMacAddress
        OperatingSystem = $thisAadDevice.operatingSystem
        OperatingSystemVersion = $thisAadDevice.operatingSystemVersion
        DeviceCompliance = $thisAadDevice.intune.complianceState
        DeviceEncryption = $thisAadDevice.intune.isEncrypted
        DeviceEnrollmentType = $thisAadDevice.intune.deviceEnrollmentType
        DeviceTrustType = $thisAadDevice.trustType
        DeviceProfileType = $thisAadDevice.profileType
        DevicePhoneNumber = $thisAadDevice.intune.phoneNumber
        DeviceNetworkCarrier = $thisAadDevice.intune.subscriberCarrier
        DeviceNetworkImei = $thisAadDevice.intune.imei
        LastUser = $user.userPrincipalName
        LastUserBusinessUnit = $user.anthesisgroup_employeeInfo.businessUnit
        LastSignInDateTime = $(get-date $thisAadDevice.approximateLastSignInDateTime -f U)
        InAAD = $true
        })
    if(![string]::IsNullOrWhiteSpace($thisAadDevice.intune)){$prettyObject | Add-Member -MemberType NoteProperty -Name InIntune -Value $true}
    else{$prettyObject | Add-Member -MemberType NoteProperty -Name InIntune -Value $false}
    if(![string]::IsNullOrWhiteSpace($thisAadDevice.asset)){$prettyObject | Add-Member -MemberType NoteProperty -Name InAsset -Value $true}
    else{$prettyObject | Add-Member -MemberType NoteProperty -Name InAsset -Value $false}
    if($thisAadDevice.physicalIds -match "ZTDID"){$prettyObject | Add-Member -MemberType NoteProperty -Name InAutoPilot -Value $true}
    else{$prettyObject | Add-Member -MemberType NoteProperty -Name InAutoPilot -Value $false}

    [array]$prettyObjects += $prettyObject
    }
$extraAssets | % {
    $thisAsset = $_
    $userId = $thisAsset.physicalIds | ? {$_ -match "USER-GID"} | % {$($_ -split ":")[1]}
    $user = $ukUsers | ? {$_.id -eq $userId}
    $prettyObject = New-Object -TypeName psobject -Property $([ordered]@{
        AssetTag = $thisAsset.fields.AnthesisSerialNumber
        AssetType = $thisAsset.fields.ContentType
        Ownership = $thisAsset.intune.managedDeviceOwnerType
        AssetBusinessUnit = $thisAsset.fields.Business_x0020_Unit.Label
        AssetSupplier = $thisAsset.fields.AssetSupplier
        AssetCost = $thisAsset.fields.AssetPriceAtPurchase
        AssetPO = $thisAsset.fields.AssetPO
        AssetPurchaseDate = $(get-date $thisAsset.fields.InvoiceDate -f U)
        AssetStatus = $thisAsset.fields.AssetStatus
        DeviceNameAAD=$thisAsset.displayName
        DeviceNameAsset = $thisAsset.fields.ComputerName
        DeviceManufacturer = $thisAsset.fields.Manufacturer
        DeviceModel = $thisAsset.model
        DeviceSerialNumber = $thisAsset.intune.serialNumber
        DeviceWifiMacAddress = $thisAsset.intune.wiFiMacAddress
        OperatingSystem = $thisAsset.operatingSystem
        OperatingSystemVersion = $thisAsset.operatingSystemVersion
        DeviceCompliance = $thisAsset.intune.complianceState
        DeviceEncryption = $thisAsset.intune.isEncrypted
        DeviceEnrollmentType = $thisAsset.intune.deviceEnrollmentType
        DeviceTrustType = $thisAsset.trustType
        DeviceProfileType = $thisAsset.profileType
        DevicePhoneNumber = "N/A"
        DeviceNetworkCarrier = "N/A"
        DeviceNetworkImei = $thisAsset.fields.IMEI
        LastUser = $user.userPrincipalName
        LastUserBusinessUnit = $user.anthesisgroup_employeeInfo.businessUnit
        LastSignInDateTime = $thisAsset.approximateLastSignInDateTime
        InAsset = $true
        InAAD = $false
        InIntune = $false
        InAutoPilot = "Unknown"
        })
    [array]$prettyObjects += $prettyObject
    }

$prettyObjects | ? {$_.InIntune -eq $true -or $_.InAsset -eq $true -or $_.InAutoPilot -eq $true} | Export-Csv  -Path 'C:\Users\KevMaitland\OneDrive - Anthesis LLC\Desktop\Hardware2.csv' -NoTypeInformation 




$tokenResponseIntuneBotAtp = get-atpTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName IntuneBot) -grant_type client_credentials 
$atpSoftware = invoke-atpGet -tokenResponse $tokenResponseIntuneBotAtp -atpQuery "/software" -Verbose
$atpSoftware | % {
    if($_.distribution -eq $null){
        $distribution = invoke-atpGet -tokenResponse $tokenResponseIntuneBotAtp -atpQuery "/software/$($_.id)/distributions" -Verbose
        $_ | Add-Member -MemberType NoteProperty -Name "distribution" -Value $distribution -Force
        }
    }


$prettyObjects = @()
$atpSoftware | Sort-Object Vendor,Name | % {
    $thisSoftware = $_
    @($thisSoftware.distribution | Select-Object) | Sort-Object Version -Descending | % {
        $prettyObject = New-Object -TypeName psobject -Property $([ordered]@{
            Vendor =   $thisSoftware.vendor
            Software = $thisSoftware.name
            Version = $_.version
            Installations = $_.installations
            Vulnerabilities = $_.vulnerabilities
            })
        $prettyObjects += $prettyObject     
        }
    }
$prettyObjects | Export-Csv  -Path 'C:\Users\KevMaitland\OneDrive - Anthesis LLC\Desktop\Software.csv' -NoTypeInformation 

$prettyObjects = @()
$atpSoftware | Sort-Object Vendor,Name | % {
    $thisSoftware = $_
    $installations = 0
    @($thisSoftware.distribution | Select-Object) | Sort-Object Version -Descending | % {
        $installations = $installations + $_.installations
        }
    $prettyObject = New-Object -TypeName psobject -Property $([ordered]@{
        Vendor =   $thisSoftware.vendor
        Software = $thisSoftware.name
        Installations = $installations
        })
    $prettyObjects += $prettyObject     
    }
$prettyObjects | Export-Csv  -Path 'C:\Users\KevMaitland\OneDrive - Anthesis LLC\Desktop\Software_Simplified.csv' -NoTypeInformation -Force
