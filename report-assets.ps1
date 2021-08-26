if($PSCommandPath){
    $VerbosePreference = 0
    $logFileLocation = "C:\ScriptLogs\"
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))`_Transcript_$(Get-Date -Format "yyyy-MM-dd").log"
    Start-Transcript $transcriptLogName -Append
    }

#region Get records to reconcile
#Get UK Users from AAD
$tokenResponseTeamsBot = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName TeamsBot) -grant_type client_credentials
$ukUsers = get-graphUsers -tokenResponse $tokenResponseTeamsBot -filterBusinessUnit 'Anthesis (UK) Ltd (GBR)' -selectAllProperties
$ukUsers += get-graphUsers -tokenResponse $tokenResponseTeamsBot -filterBusinessUnit 'Anthesis Energy UK Ltd (GBR)' -selectAllProperties 

#Get Asset records from SharePoint
$tokenResponseSharePointBot = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName SharePointBot) -grant_type client_credentials
$itTeamAllSite = get-graphSite -tokenResponse $tokenResponseSharePointBot -serverRelativeUrl "/teams/IT_Team_All_365"
$assetRegister = get-graphList -tokenResponse $tokenResponseSharePointBot -graphSiteId $itTeamAllSite.id -listName "Anthesis IT Asset Register"
$assetRegisterItems = get-graphListItems -tokenResponse $tokenResponseSharePointBot -graphSiteId $itTeamAllSite.id -listId $assetRegister.id -expandAllFields #$assetRegisterItems.fields.AssetStatus | select -Unique
$assetRegisterComputers = $assetRegisterItems | ? {$_.fields.ContentType -eq "Computers"}
$assetRegisterPhones = $assetRegisterItems | ? {$_.fields.ContentType -eq "Mobiles"}

#Get all Intune, ATP & AAD devices
$tokenResponseIntuneBot = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName IntuneBot) -grant_type client_credentials 
$intuneDevices = get-graphIntuneDevices -tokenResponse $tokenResponseIntuneBot
$tokenResponseIntuneBotAtp = get-atpTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName IntuneBot) -grant_type client_credentials 
$atpDevices = get-atpMachines -tokenResponse $tokenResponseIntuneBotAtp
$allAadDevices = get-graphDevices -tokenResponse $tokenResponseTeamsBot -includeOwners

#Get encryption state report - we can't pull this by 
$deviceEncryptionStates = get-DeviceEncryptionStates -tokenResponse $tokenResponseIntuneBot -Verbose

#endregion

#region Find stuff
#Match records and add Intune/ATP/Asset data onto the AAD device record (assuming that all devices exist in AAD, which isn;t 100% accurate for ATP)

#Iteration steps: #Get Aad device as main interation object -> Try to find the Atp device -> try to find the Intune device -> try to find the asset register device -> #~Add all the info onto the Aad object as properties~#

$allAadDevices | % { 
    $thisAadDevice = $_
    ##Clear any existing variables ready to go for the next run
    if($correspondingAtpDevice){rv correspondingAtpDevice}
    if($correspondingIntuneDevice){rv correspondingIntuneDevice}
    if($correspondingAsset){rv correspondingAsset}

    ##Grab atp device by filtering all atp machines for this device's aad DEVICE id - atp object has aad record on it
    Write-Host "Processing [$($thisAadDevice.displayName)][$($thisAadDevice.deviceId)]"
    $correspondingAtpDevice = $atpDevices | ? {$_.aadDeviceId -eq $thisAadDevice.deviceId} | Sort-Object firstSeen -Descending | Select-Object -First 1

    ##If we find a corresponding atp device that lives in aad, add the atp object info into a hash table and add add it to the $thisAadDevice object as a propery/element (whatever, it's on there somewhere and its query-able)
    if($correspondingAtpDevice){
        Write-Host "`tAdding ATP information to [$($thisAadDevice.displayName)][$($thisAadDevice.deviceId)]"
        $atpHash = [ordered]@{}
        Get-Member -InputObject $correspondingAtpDevice -MemberType Properties | % {
            $atpHash.Add($_.Name, $correspondingAtpDevice.$($_.Name))
            }
        $thisAadDevice | Add-Member -MemberType NoteProperty -Name atp -Value $atpHash -Force
        }

    ##Do the above for Intune as well to see if we can find an Intune device, using the Aad device id
    $correspondingIntuneDevice = $intuneDevices | ? {$_.azureADDeviceId -eq $thisAadDevice.deviceId} | Sort-Object enrolledDateTime -Descending | Select-Object -First 1
    if($correspondingIntuneDevice){
        Write-Host "`tAdding Intune information to [$($thisAadDevice.displayName)][$($thisAadDevice.deviceId)]"
        $intuneHash = @{}
        Get-Member -InputObject $correspondingIntuneDevice -MemberType Properties | % {
            $intuneHash.Add($_.Name, $correspondingIntuneDevice.$($_.Name))
            }
        $_ | Add-Member -MemberType NoteProperty -Name intune -Value $intuneHash -Force

    ##Do the above for advanced encryption state information stored *somewhere* in Intune and not on the direct Intune device object, using the Aad device id
     $correspondingEncryptionDevice = $deviceEncryptionStates | ? {$_.Id -eq $thisAadDevice.intune.id}
     if($correspondingEncryptionDevice){
        Write-Host "`tAdding Encryption information to [$($thisAadDevice.displayName)][$($thisAadDevice.deviceId)]"
        $encryptionHash = @{}
        Get-Member -InputObject $correspondingEncryptionDevice -MemberType Properties | % {
            $encryptionHash.Add($_.Name, $correspondingEncryptionDevice.$($_.Name))
            }
        $_ | Add-Member -MemberType NoteProperty -Name encryptiondata -Value $encryptionHash -Force
        }

    ##Then try matching the Asset using the manufacturer serial number - this lives on the Intune object which we found using the Aad device id
        $correspondingAsset = $assetRegisterItems | ? {$_.fields.ManufacturerSerialNumber -eq $correspondingIntuneDevice.serialNumber}
        ##If we can't find it by serial number, try product code against the Intune serial number
        if(!$correspondingAsset){
            $correspondingAsset = $assetRegisterComputers | ? {$_.fields.IT_x0020_Product_x0020_Code -eq $correspondingIntuneDevice.serialNumber}
            ##If we can't find it by product tag, try matching with MAC addresses (also lives on the Intune object)
            if(!$correspondingAsset){
                $correspondingAsset = $assetRegisterComputers | ? {![string]::IsNullOrWhiteSpace($_.fields.MACAddresses)} | ? {$_.fields.MACAddresses.Replace(":","").Replace("-","") -match $correspondingIntuneDevice.wiFiMacAddress}
                ##If we STILL can't find it, try using the computer name against the Aad display name as a last ditch attempt
                if(!$correspondingAsset){
                    $correspondingAsset = $assetRegisterComputers | ? {$_.fields.ComputerName -eq $thisAadDevice.displayName}
                    ##If it's a mobile, we can match using the IMEI if its in the asset register
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
        ##If we can't find it in Intune, add it to a running list in $notInIntune - we won't check the asset register
        Write-Warning "No Intune device found for [$($thisAadDevice.displayName)][$($thisAadDevice.deviceId)]"
        [array]$notInIntune += $thisAadDevice
        }
    #If we find the corresponding asset, add the asset info into a hash table and add to the Aad device object
    if($correspondingAsset){
        Write-Host "`tAdding Asset information to [$($thisAadDevice.displayName)][$($thisAadDevice.deviceId)]"
        $assetHash = @{}
        Get-Member -InputObject $correspondingAsset.fields -MemberType Properties | % {
            $assetHash.Add($_.Name, $correspondingAsset.fields.$($_.Name))
            }
        $_ | Add-Member -MemberType NoteProperty -Name asset -Value $assetHash -Force
        }

    }

#Filter & de-dupe the objects that belong to UK users
$ukAadDevices = $allAadDevices | ? {$_.intune.userId -match $($ukUsers.id -join "|")}
$ukAadDevices = $ukAadDevices | Group-Object {$_.intune.serialNumber} | % {$_.Group | Sort-Object approximateLastSignInDateTime | Select-Object -Last 1} #DeDupe and keep only the most recent

<#
#Find any assets that are missing from AAD (for gap analysis)
$extraAssets = @()
$assetRegisterComputersUseful = $assetRegisterComputers | ? {$_.fields.AssetStatus -notmatch "Dispose" -and $_.fields.AssetStatus -notmatch "Retire"}
$assetRegisterComputersUseful | % {
    $thisAsset = $_
    if($matchedDevice){rv matchedDevice}
    if(![string]::IsNullOrWhiteSpace($thisAsset.fields.ManufacturerSerialNumber)){
        $matchedDevice = $ukAadDevices | ? {$_.asset.ManufacturerSerialNumber -eq $thisAsset.fields.ManufacturerSerialNumber}
        if(!$matchedDevice){
            $extraAssets += $_
            write-host "`t[$($thisAsset.fields.ComputerName)] not matched" -f Yellow
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
    #>
#endregion

#region Report stuff
#region Update the Computers Asset Register with data fron AAD, Intune & ATP
$ukAadDevices | ? {$_.asset.ContentType -eq "Computers"} | % {

    $thisComputer = $_
    $thisUserId = $thisComputer.physicalIds | ? {$_ -match "USER-GID"} | % {$($_ -split ":")[1]}
    $thisUser = $ukUsers | ? {$_.id -eq $thisUserId}
    
  
    $updateHash = [ordered]@{
        Computer_AadDeviceName = $thisComputer.displayName
        Computer_AadLastUser = $thisUser.userPrincipalName
        Computer_AadLastUserSignin = $thisComputer.approximateLastSignInDateTime
        Computer_AadManufacturer = $thisComputer.manufacturer
        Computer_AadModel = $thisComputer.model
        Computer_AadOsBuildNumber = $thisComputer.operatingSystemVersion
        Computer_AadProfileType = $thisComputer.profileType
        Computer_AadTrustType = $thisComputer.trustType
        Computer_AtpExposureLevel = $thisComputer.atp.exposureLevel
        Computer_AtpFirstSeen = $thisComputer.atp.firstSeen
        Computer_AtpIsAadJoined = $thisComputer.atp.isAadJoined
        Computer_AtpLastExternalIP = $thisComputer.atp.lastExternalIpAddress
        Computer_AtpLastInternalIP = $thisComputer.atp.lastIpAddress
        Computer_AtpLastSeen = $thisComputer.atp.lastSeen
        Computer_AtpOs = $thisComputer.atp.osPlatform
        Computer_AtpOsProcessor = $thisComputer.atp.osProcessor
        Computer_AtpOsVersion = $thisComputer.atp.version
        Computer_IntuneComplianceState = $thisComputer.intune.complianceState
        Computer_IntuneEnrollmentType = $thisComputer.intune.deviceEnrollmentType
        Computer_IntuneIsEncrypted = $thisComputer.intune.isEncrypted
        Computer_IntuneLastUser = $thisComputer.intune.userPrincipalName
        Computer_IntuneLastUserBusinessU = $thisUser.anthesisgroup_employeeInfo.businessUnit
        Computer_IntuneSerialNumber = $thisComputer.intune.serialNumber
        Computer_IntuneWiFiMacAddress = $thisComputer.intune.wiFiMacAddress
        Computer_AadID = $thisComputer.id
        Computer_AtpID = $thisComputer.atp.id
        Computer_IntuneID = $thisComputer.Intune.id
        Computer_IntuneEncryptionState = $thisComputer.encryptiondata.encryptionState 
        Computer_IntuneTpmPresent = if($thisComputer.encryptiondata.tpmSpecificationVersion){"Yes"}` else{"No"}
        Computer_IntuneAdvancedBitLocker = $thisComputer.encryptiondata.advancedBitLockerStates
        Computer_EncryptionPolicyDetails = [string]$thisComputer.encryptiondata.policyDetails.policyName
        PresentInAad = $true
        }
    if(![string]::IsNullOrWhiteSpace($thisComputer.atp)){$updateHash.Add("PresentInAtp",$true)}
    else{$updateHash.Add("PresentInAtp",$false)}
    if(![string]::IsNullOrWhiteSpace($thisComputer.intune)){$updateHash.Add("PresentInIntune",$true)}
    else{$updateHash.Add("PresentInIntune",$false)}
    if($thisComputer.physicalIds -match "ZTDID"){$updateHash.Add("PresentInAutopilot",$true)}
    else{$updateHash.Add("PresentInAutopilot",$false)}
    $null = update-graphListItem -tokenResponse $tokenResponseSharePointBot -graphSiteId $itTeamAllSite.id -listId $assetRegister.id -listitemId $thisComputer.asset.id -fieldHash $updateHash #-Verbose
    Write-Host -f Yellow "Updated [$($thisComputer.displayName)]"
    #Tag atp device with the asset register status if not already there to help reduce admin time :)

    #get any current status options from It Asset Register AssetStatus column
    $possibleAssetTags = $assetRegisterItems.fields.AssetStatus | select -Unique

    If(($thisComputer.atp) -and ($thisComputer.asset) -and ![string]::IsNullOrWhiteSpace($thisComputer.atp.machineTags) -and ($thisComputer.atp.machineTags -notcontains $thisComputer.asset.AssetStatus)){

    #Remove any old status
    $thisTag = Compare-Object -ReferenceObject $thisComputer.atp.machineTags -differenceobject $possibleAssetTags -IncludeEqual | Where-Object -Property "SideIndicator" -EQ "==" 
    ForEach($tag in $thisTag){
    remove-atpDeviceTag -tokenResponse $tokenResponseIntuneBotAtp -deviceid $thisComputer.atp.id -tagstring $tag.InputObject
    }    
    #Add new status
    add-atpDeviceTag -tokenResponse $tokenResponseIntuneBotAtp -deviceid $thisComputer.atp.id -tagstring $thisComputer.asset.AssetStatus
    }
    

    #Update status with maternity/paternity leave, leavers - so we have some sort of update
    #-get lists
    #-find current users from lists
    #-update it asset register
    #-update atp?    
    
    }
        




#endregion


<#
#region Report detailed hardware CSV
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
        FirstSeenInATP = $thisAadDevice.atp.firstSeen
        LastSeenInATP = $thisAadDevice.atp.lastSeen
        AtpOsVersion = $thisAadDevice.atp.version
        AtpOsBuild = $thisAadDevice.atp.osBuild
        AtpHealthStatus = $thisAadDevice.atp.healthStatus
        AtpExposureLevel = $thisAadDevice.atp.exposureLevel
        AtpLastInternalIp = $thisAadDevice.atp.lastIpAddress
        AtpLastExternalIp = $thisAadDevice.atp.lastExternalIpAddress
        InAAD = $true
        })
    if(![string]::IsNullOrWhiteSpace($thisAadDevice.atp)){$prettyObject | Add-Member -MemberType NoteProperty -Name InATP -Value $true}
    else{$prettyObject | Add-Member -MemberType NoteProperty -Name InATP -Value $false}
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
#Output to file
$prettyObjects | ? {$_.InIntune -eq $true -or $_.InAsset -eq $true -or $_.InAutoPilot -eq $true} | Export-Csv  -Path 'C:\Users\KevMaitland\OneDrive - Anthesis LLC\Desktop\Hardware5.csv' -NoTypeInformation 
#endregion
#region Report simplified hardware CSV
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
#endregion
#region Report detailed Software from ATP
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

#endregion
#region Report simplified hardware from ATP
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
#endregion
#endregion

#>

Stop-Transcript


