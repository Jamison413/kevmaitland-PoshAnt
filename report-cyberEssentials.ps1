$tokenResponseTeamsBot = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName TeamsBot) -grant_type client_credentials
$ukUsers = get-graphUsers -tokenResponse $tokenResponseTeamsBot -filterBusinessUnit 'Anthesis (UK) Ltd (GBR)' -selectAllProperties -filterLicensedUsers
$ukUsers += get-graphUsers -tokenResponse $tokenResponseTeamsBot -filterBusinessUnit 'Anthesis Energy UK Ltd (GBR)' -selectAllProperties  -filterLicensedUsers
$ukUsers += get-graphUsers -tokenResponse $tokenResponseTeamsBot -filterBusinessUnit 'Anthesis Consulting Group (GBR)' -selectAllProperties  -filterLicensedUsers

#Get all Intune, ATP & AAD devices
$tokenResponseIntuneBot = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName IntuneBot) -grant_type client_credentials 
$intuneDevices = get-graphIntuneDevices -tokenResponse $tokenResponseIntuneBot
$tokenResponseIntuneBotAtp = get-atpTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName IntuneBot) -grant_type client_credentials 
$atpDevices = get-atpMachines -tokenResponse $tokenResponseIntuneBotAtp
$allAadDevices = get-graphDevices -tokenResponse $tokenResponseTeamsBot -includeOwners

$allAadDevices | ForEach-Object {Add-Member -InputObject $_ -MemberType NoteProperty -Name OwnerId -Value $_.registeredOwners[0].id  -Force}
$ukUsers | ForEach-Object {Add-Member -InputObject $_ -MemberType NoteProperty -Name OwnerId -Value $_.id -Force}
$thisGeoDevices = $allAadDevices | ? {@($ukUsers.OwnerId) -Contains $_.ownerId -or [string]::IsNullOrWhiteSpace($_.ownerId)}
#$test = Compare-Object -ReferenceObject @($allAadDevices | Select-Object) -DifferenceObject @($thisGeoUsers | Select-Object) -Property OwnerId -IncludeEqual -ExcludeDifferent -PassThru 

$deviceEncryptionStates = get-DeviceEncryptionStates -tokenResponse $tokenResponseIntuneBot -Verbose
$thisGeoDevices | foreach-object { 
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
        $thisAadDevice | Add-Member -MemberType NoteProperty -Name intune -Value $intuneHash -Force

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



$prettyDevices = @($null) * $thisGeoDevices.Count
$i=0
$thisGeoDevices | % {
    $thisDevice = $_
    $thisDeviceObject = New-Object -TypeName PSCustomObject -Property ([ordered]@{
        Owner = $thisDevice.registeredOwners.userPrincipalName
        DeviceName = $thisDevice.displayName
        Manufacturer = $null
        Model = $null
        Serial = $null
        OSType = $null
        OSVersion = $null
        EnrollmentType = $thisDevice.enrollmentType
        TrustType = $thisDevice.trustType
        Ownership = $thisDevice.deviceOwnership
        LastSeenAAD = $thisDevice.approximateLastSignInDateTime
        LastSeenIntune = $null
        LastSeenMde = $null
        })
    if([string]::IsNullOrWhiteSpace($thisDevice.intune)){
        $thisDeviceObject.Manufacturer = $thisDevice.manufacturer
        $thisDeviceObject.Model = $thisDevice.model
        $thisDeviceObject.Serial = $null
        $thisDeviceObject.OSType = $thisDevice.operatingSystem
        $thisDeviceObject.LastSeenIntune = $null
        }
    else{
        $thisDeviceObject.Manufacturer = $thisDevice.intune.manufacturer
        $thisDeviceObject.Model = $thisDevice.intune.model
        $thisDeviceObject.Serial = $thisDevice.intune.serialNumber
        $thisDeviceObject.OSType = $thisDevice.intune.operatingSystem
        $thisDeviceObject.LastSeenIntune = $thisDevice.intune.lastSyncDateTime
        }
    if([string]::IsNullOrWhiteSpace($thisDevice.atp)){
        $thisDeviceObject.OSVersion = $(
            switch -Regex ($thisDevice.operatingSystemVersion) {
                '^10.0.19043'{"21H1"}
                '^10.0.19042'{"20H2"}
                '^10.0.19041'{"2004"}
                '^10.0.18363'{"1909"}
                '^10.0.18362'{"1903"}
                '^10.0.17763'{"1809"}
                '^10.0.17134'{"1803"}
                '^10.0.16299'{"1709"}
                '^10.0.15063'{"1703"}
                '^10.0.14393'{"1607"}
                '^10.0.10586'{"1511"}
                '^10.0.10240'{"1507"}
                default {
                    try{[int]$thisDevice.operatingSystemVersion}
                    catch{$thisDevice.operatingSystemVersion.Split(".")[0]}
                    }
                }
            )
        $thisDeviceObject.LastSeenMde = $null
        }
    else{
        $thisDeviceObject.OSVersion = $thisDevice.atp.version
        $thisDeviceObject.LastSeenMde = $thisDevice.atp.lastSeen
        }
    $prettyDevices[$i] = $thisDeviceObject
    $i++
    }

$prettyIntuneDevices = $prettyDevices | ?{![string]::IsNullOrWhiteSpace($_.LastSeenIntune)} | Group-Object {$_.Serial} | % {$_.Group | Sort-Object LastSeenIntune | Select-Object -Last 1} #DeDupe and keep only the most recent
$prettyNonIntuneDevices = $prettyDevices | ?{[string]::IsNullOrWhiteSpace($_.LastSeenIntune)} | Group-Object {$_.DeviceName} | % {$_.Group | Sort-Object LastSeenAAD | Select-Object -Last 1} #DeDupe and keep only the most recent
$prettyDedupededDevices = $prettyIntuneDevices + $prettyNonIntuneDevices

#$prettyDedupededDevices = $prettyDedupededDevices | Group-Object {$_.DeviceName} | % {$_.Group | Sort-Object LastSeenAAD | Select-Object -Last 1} #DeDupe and keep only the most recent
$prettyDedupededAndPrunedDevices = $prettyDedupededDevices | Where-Object {[string]::IsNullOrWhiteSpace($_.LastSeenAAD) -or (Get-Date ($_.LastSeenAAD)) -gt (Get-Date).AddMonths(-3)} 
$usersWithNoHardware = Compare-Object -ReferenceObject $($ukUsers.userPrincipalName | Select-Object -Unique) -DifferenceObject $($prettyDedupededAndPrunedDevices.Owner | Select-Object -Unique) | ? {$_.SideIndicator -eq "<="}
$usersWithNoHardware | % {$prettyDedupededAndPrunedDevices += New-Object -TypeName PSCustomObject -Property ([ordered]@{
        Owner = $_.InputObject
        DeviceName = $null
        Manufacturer = $null
        Model = $null
        Serial = $null
        OSType = $null
        OSVersion = $null
        EnrollmentType = $null
        TrustType = $null
        Ownership = $null
        LastSeenAAD = $null
        LastSeenIntune = $null
        LastSeenMde = $null
        })}

$prettyDedupededAndPrunedDevices | Select-Object | Sort-Object Owner,Ownership,OSType,OSVersion | export-csv -Path $env:USERPROFILE\Downloads\CyberEssentialsDump.csv -NoTypeInformation -Force

Write-Host -ForegroundColor Yellow "COBO Windows"
$coboWindows = $prettyDedupededAndPrunedDevices | Select-Object | Where-Object {$_.Ownership -eq "Company" -and @("Windows") -contains $_.OSType} | Sort-Object OSType, OSVersion | Group-Object OSType, OSVersion | Select-Object Count, Name
$coboWindows | %{
    #Write-Host "$($_.Count)x $($_.Name -replace "^(?=(?:[^,]*,){2})([^,]*),", '$1')"
    Write-Host "$($_.Count)x $((Get-Culture).TextInfo.ToTitleCase($($_.Name -replace "^(?=(?:[^,]*,){2})([^,]*),", '$1')).Replace("Windows","Windows 10"))"
    }
Write-Host -ForegroundColor Yellow "BYOD Windows"
$byodWindows = $prettyDedupededAndPrunedDevices | Select-Object | Where-Object {$_.Ownership -ne "Company" -and @("Windows") -contains $_.OSType} | Sort-Object OSType, OSVersion | Group-Object OSType, OSVersion | Select-Object Count, Name
$byodWindows | %{
    #Write-Host "$($_.Count)x $($_.Name -replace "^(?=(?:[^,]*,){2})([^,]*),", '$1')"
    Write-Host "$($_.Count)x $((Get-Culture).TextInfo.ToTitleCase($($_.Name -replace "^(?=(?:[^,]*,){2})([^,]*),", '$1')).Replace("Windows","Windows 10"))"
    }

Write-Host -ForegroundColor Yellow "COBO Mobile"
$coboMobile = $prettyDedupededAndPrunedDevices | Select-Object | Where-Object {$_.Ownership -eq "Company" -and @("Android","AndroidEnterprise","iOS","IPhone") -contains $_.OSType} | Sort-Object Manufacturer, Model, OSVersion | Group-Object Manufacturer, Model, OSVersion | Select-Object Count, Name
$coboMobile | %{
    #Write-Host "$($_.Count)x $($_.Name -replace "^(?=(?:[^,]*,){2})([^,]*),", '$1')"
    Write-Host "$($_.Count)x $((Get-Culture).TextInfo.ToTitleCase($($_.Name -replace "^(?=(?:[^,]*,){2})([^,]*),", '$1')))"
    }

Write-Host -ForegroundColor Yellow "BYOD Mobile"
$byodMobile = $prettyDedupededAndPrunedDevices | Select-Object | Where-Object {$_.Ownership -eq "Personal" -and @("Android","AndroidEnterprise","iOS","IPhone") -contains $_.OSType} | Sort-Object Manufacturer, Model, OSVersion | Group-Object Manufacturer, Model, OSVersion #| Select-Object Count, Name
$byodMobile | %{
    #Write-Host "$($_.Count)x $($_.Name -replace "^(?=(?:[^,]*,){2})([^,]*),", '$1')"
    Write-Host "$($_.Count)x $((Get-Culture).TextInfo.ToTitleCase($($_.Name -replace "^(?=(?:[^,]*,){2})([^,]*),", '$1')).Replace("Ipad","iPad").Replace("Iphone","iPhone"))"
    }
