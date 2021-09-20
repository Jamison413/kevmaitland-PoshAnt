
function add-userToMdmByodDistributionGroup(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [string]$upn 
        ,[parameter(Mandatory = $false)]
        [PSObject]$fullLogFile 
        ,[parameter(Mandatory = $false)]
        [PSObject]$errorLogFile 
        ,[parameter(Mandatory = $false)]
        [PSObject]$mdmByodDistributionGroup 
        )

    if([string]::IsNullOrWhiteSpace($mdmByodDistributionGroup)){
        if(![string]::IsNullOrWhiteSpace($fullLogFile) -and ![string]::IsNullOrWhiteSpace($errorLogFile)){
            $mdmByodDistributionGroup = get-mdmByodDistributionGroup -fullLogFile $fullLogFile -errorLogFile $errorLogFile
            }
        else{$mdmByodDistributionGroup = get-mdmByodDistributionGroup}
        }

    try{#Add to "MDM - BYOD Mobile Device Users"
        Write-Verbose "Adding [$upn] to [$($mdmByodDistributionGroup.DisplayName)]" 
        Add-DistributionGroupMember -Identity $mdmByodDistributionGroup.ExternalDirectoryObjectId -Member $upn -ErrorAction Stop #Add to "MDM - BYOD Mobile Device Users"
        Write-Verbose "[$upn] successfully added to [$($mdmByodDistributionGroup.DisplayName)]"
        }
    catch{
        if(![string]::IsNullOrWhiteSpace($fullLogFile) -and ![string]::IsNullOrWhiteSpace($errorLogFile)){
            if($_.Exception.HResult -eq -2146233087){Write-Verbose "[$($upn)] already a member of [$($mdmByodDistributionGroup.DisplayName)]"}
            else{log-error -myError $_ -myFriendlyMessage "Error adding [$($upn)] to [$($mdmByodDistributionGroup.DisplayName)] in add-userToByodMdmGroup" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName}
            }
        else{Write-Error $_;$_}
        }
    }
function disable-legacyMailboxProtocols(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [string]$upn 
        ,[parameter(Mandatory = $false)]
        [PSObject]$fullLogFile 
        ,[parameter(Mandatory = $false)]
        [PSObject]$errorLogFile 
        )

    try{# Disable legacy mailbox protocols
        Write-Verbose "Disabling legacy mailbox protocols for [$upn]"
        Set-CASMailbox -Identity $upn -ImapEnabled $false -ActiveSyncEnabled $false -PopEnabled $false -OWAforDevicesEnabled $false -ActiveSyncMailboxPolicy "Default" -ErrorAction Stop #Disable legacy mailbox protocols to avoid MFA bypass -MAPIEnabled $false
        Write-Verbose "Legacy Protocols successfully disabled for [$upn]"
        }
    catch{
        if(![string]::IsNullOrWhiteSpace($fullLogFile) -and ![string]::IsNullOrWhiteSpace($errorLogFile)){
            log-error -myError $_ -myFriendlyMessage "Error disabling legacy mailbox protocols for [$upn] in disable-legacyMailboxProtocols" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
            }
        else{Write-Error $_;$_}
        }
    }
function get-DeviceEncryptionStates(){
     [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [pscustomobject]$tokenResponse
        )
        invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/deviceManagement/managedDeviceEncryptionStates" -useBetaEndPoint
}
function get-mdmByodDistributionGroup(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $false)]
        [PSObject]$fullLogFile 
        ,[parameter(Mandatory = $false)]
        [PSObject]$errorLogFile 
        )

    try{
        $mdmByodDistributionGroup = Get-DistributionGroup -Identity b264f337-ef04-432e-a139-3574331a4d18 #"MDM - BYOD Mobile Device Users"
        Write-Verbose "[$($mdmByodDistributionGroup.DisplayName)] retrieved"
        }
    catch{
        if(![string]::IsNullOrWhiteSpace($fullLogFile) -and ![string]::IsNullOrWhiteSpace($errorLogFile)){
            log-error -myError $_ -myFriendlyMessage "Error Retrieving `"MDM - BYOD Mobile Device Users`" Distribution Group in get-mdmByodDistributionGroup" -fullLogFile $fullLogFile -errorLogFile $errorLogFile -doNotLogToEmail $true
            }
        else{Write-Error $_;$_}
        }
    $mdmByodDistributionGroup
    }
function get-mdmPolicyDeviceConfig(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse   
        ,[parameter(Mandatory = $true,ParameterSetName = "explicit")]
            [string]$filterId
        ,[parameter(Mandatory = $false,ParameterSetName = "ambiguous")]
            [string]$filterDisplayNameEquals
        ,[parameter(Mandatory = $false,ParameterSetName = "ambiguous")]
            [string]$filterDisplayNameStartsWith
        ,[parameter(Mandatory = $false)]
            [switch]$selectAllProperties = $false
        )     

    if($selectAllProperties){ #We're only dealing with one select option (all or default)
        $select  = "`$select=id,lastModifiedDateTime,createdDateTime,description,displayName,version,enterpriseCloudPrintDiscoveryEndPoint,enterpriseCloudPrintOAuthAuthority,enterpriseCloudPrintOAuthClientIdentifier,enterpriseCloudPrintResourceIdentifier,enterpriseCloudPrintDiscoveryMaxLimit,enterpriseCloudPrintMopriaDiscoveryResourceIdentifier,searchBlockDiacritics,searchDisableAutoLanguageDetection,searchDisableIndexingEncryptedItems,searchEnableRemoteQueries,searchDisableIndexerBackoff,searchDisableIndexingRemovableDrive,searchEnableAutomaticIndexSizeManangement,diagnosticsDataSubmissionMode,oneDriveDisableFileSync,smartScreenEnableAppInstallControl,personalizationDesktopImageUrl,personalizationLockScreenImageUrl,bluetoothAllowedServices,bluetoothBlockAdvertising,bluetoothBlockDiscoverableMode,bluetoothBlockPrePairing,edgeBlockAutofill,edgeBlocked,edgeCookiePolicy,edgeBlockDeveloperTools,edgeBlockSendingDoNotTrackHeader,edgeBlockExtensions,edgeBlockInPrivateBrowsing,edgeBlockJavaScript,edgeBlockPasswordManager,edgeBlockAddressBarDropdown,edgeBlockCompatibilityList,edgeClearBrowsingDataOnExit,edgeAllowStartPagesModification,edgeDisableFirstRunPage,edgeBlockLiveTileDataCollection,edgeSyncFavoritesWithInternetExplorer,cellularBlockDataWhenRoaming,cellularBlockVpn,cellularBlockVpnWhenRoaming,defenderBlockEndUserAccess,defenderDaysBeforeDeletingQuarantinedMalware,defenderDetectedMalwareActions,defenderSystemScanSchedule,defenderFilesAndFoldersToExclude,defenderFileExtensionsToExclude,defenderScanMaxCpu,defenderMonitorFileActivity,defenderProcessesToExclude,defenderPromptForSampleSubmission,defenderRequireBehaviorMonitoring,defenderRequireCloudProtection,defenderRequireNetworkInspectionSystem,defenderRequireRealTimeMonitoring,defenderScanArchiveFiles,defenderScanDownloads,defenderScanNetworkFiles,defenderScanIncomingMail,defenderScanMappedNetworkDrivesDuringFullScan,defenderScanRemovableDrivesDuringFullScan,defenderScanScriptsLoadedInInternetExplorer,defenderSignatureUpdateIntervalInHours,defenderScanType,defenderScheduledScanTime,defenderScheduledQuickScanTime,defenderCloudBlockLevel,lockScreenAllowTimeoutConfiguration,lockScreenBlockActionCenterNotifications,lockScreenBlockCortana,lockScreenBlockToastNotifications,lockScreenTimeoutInSeconds,passwordBlockSimple,passwordExpirationDays,passwordMinimumLength,passwordMinutesOfInactivityBeforeScreenTimeout,passwordMinimumCharacterSetCount,passwordPreviousPasswordBlockCount,passwordRequired,passwordRequireWhenResumeFromIdleState,passwordRequiredType,passwordSignInFailureCountBeforeFactoryReset,privacyAdvertisingId,privacyAutoAcceptPairingAndConsentPrompts,privacyBlockInputPersonalization,startBlockUnpinningAppsFromTaskbar,startMenuAppListVisibility,startMenuHideChangeAccountSettings,startMenuHideFrequentlyUsedApps,startMenuHideHibernate,startMenuHideLock,startMenuHidePowerButton,startMenuHideRecentJumpLists,startMenuHideRecentlyAddedApps,startMenuHideRestartOptions,startMenuHideShutDown,startMenuHideSignOut,startMenuHideSleep,startMenuHideSwitchAccount,startMenuHideUserTile,startMenuLayoutEdgeAssetsXml,startMenuLayoutXml,startMenuMode,startMenuPinnedFolderDocuments,startMenuPinnedFolderDownloads,startMenuPinnedFolderFileExplorer,startMenuPinnedFolderHomeGroup,startMenuPinnedFolderMusic,startMenuPinnedFolderNetwork,startMenuPinnedFolderPersonalFolder,startMenuPinnedFolderPictures,startMenuPinnedFolderSettings,startMenuPinnedFolderVideos,settingsBlockSettingsApp,settingsBlockSystemPage,settingsBlockDevicesPage,settingsBlockNetworkInternetPage,settingsBlockPersonalizationPage,settingsBlockAccountsPage,settingsBlockTimeLanguagePage,settingsBlockEaseOfAccessPage,settingsBlockPrivacyPage,settingsBlockUpdateSecurityPage,settingsBlockAppsPage,settingsBlockGamingPage,windowsSpotlightBlockConsumerSpecificFeatures,windowsSpotlightBlocked,windowsSpotlightBlockOnActionCenter,windowsSpotlightBlockTailoredExperiences,windowsSpotlightBlockThirdPartyNotifications,windowsSpotlightBlockWelcomeExperience,windowsSpotlightBlockWindowsTips,windowsSpotlightConfigureOnLockScreen,networkProxyApplySettingsDeviceWide,networkProxyDisableAutoDetect,networkProxyAutomaticConfigurationUrl,networkProxyServer,accountsBlockAddingNonMicrosoftAccountEmail,antiTheftModeBlocked,bluetoothBlocked,cameraBlocked,connectedDevicesServiceBlocked,certificatesBlockManualRootCertificateInstallation,copyPasteBlocked,cortanaBlocked,deviceManagementBlockFactoryResetOnMobile,deviceManagementBlockManualUnenroll,safeSearchFilter,edgeBlockPopups,edgeBlockSearchSuggestions,edgeBlockSendingIntranetTrafficToInternetExplorer,edgeSendIntranetTrafficToInternetExplorer,edgeRequireSmartScreen,edgeEnterpriseModeSiteListLocation,edgeFirstRunUrl,edgeSearchEngine,edgeHomepageUrls,edgeBlockAccessToAboutFlags,smartScreenBlockPromptOverride,smartScreenBlockPromptOverrideForFiles,webRtcBlockLocalhostIpAddress,internetSharingBlocked,settingsBlockAddProvisioningPackage,settingsBlockRemoveProvisioningPackage,settingsBlockChangeSystemTime,settingsBlockEditDeviceName,settingsBlockChangeRegion,settingsBlockChangeLanguage,settingsBlockChangePowerSleep,locationServicesBlocked,microsoftAccountBlocked,microsoftAccountBlockSettingsSync,nfcBlocked,resetProtectionModeBlocked,screenCaptureBlocked,storageBlockRemovableStorage,storageRequireMobileDeviceEncryption,usbBlocked,voiceRecordingBlocked,wiFiBlockAutomaticConnectHotspots,wiFiBlocked,wiFiBlockManualConfiguration,wiFiScanInterval,wirelessDisplayBlockProjectionToThisDevice,wirelessDisplayBlockUserInputFromReceiver,wirelessDisplayRequirePinForPairing,windowsStoreBlocked,appsAllowTrustedAppsSideloading,windowsStoreBlockAutoUpdate,developerUnlockSetting,sharedUserAppDataAllowed,appsBlockWindowsStoreOriginatedApps,windowsStoreEnablePrivateStoreOnly,storageRestrictAppDataToSystemVolume,storageRestrictAppInstallToSystemVolume,gameDvrBlocked,experienceBlockDeviceDiscovery,experienceBlockErrorDialogWhenNoSIM,experienceBlockTaskSwitcher,logonBlockFastUserSwitching,tenantLockdownRequireNetworkDuringOutOfBoxExperience"
        }

    switch ($PsCmdlet.ParameterSetName){
        “explicit”  {
            invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/deviceManagement/deviceConfigurations/$filterId$select"
            return
            }
        }
    
    if($filterDisplayNameEquals){$filter += " and displayName eq '$([uri]::EscapeDataString($filterDisplayNameEquals))'"}
    #if($filterDisplayNameEquals){$filter += " and displayName eq '$(($filterDisplayNameEquals))'"}
    if($filterDisplayNameStartsWith){$filter += " and startswith(displayName,'$([uri]::EscapeDataString($filterDisplayNameStartsWith))')"}
    if(![string]::IsNullOrWhiteSpace($filter)){
        if($filter.StartsWith(" and ")){$filter = $filter.Substring(5,$filter.Length-5)}
        $filter = "`$filter=$filter"
        }

    $refiner = "?"+$select
    if($filter){
        if($refiner.Length -gt 1){$refiner = $refiner+"&"} #If there is already another query option in the refiner, use the '&' symbol to concatenate the the strings
        $refiner = $refiner+$filter
        }

    $results = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/deviceManagement/deviceConfigurations/$refiner"
    $results
    }
function get-mdmPolicyDeviceConfigAssignment(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse   
        ,[parameter(Mandatory = $true)]
            [string]$configurationId
        ,[parameter(Mandatory = $false)]
            [string]$assignmentId
        )     

    $results = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/deviceManagement/deviceConfigurations/$configurationId/assignments/$assignmentId"
    $results
    }
function new-mdmLocalAdminPolicy(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponseTeams
        ,[parameter(Mandatory = $true)]
            [psobject]$tokenResponseIntune
        ,[parameter(Mandatory = $true)]
            [ValidatePattern(".[@].")]
            [string]$userUPN
        ,[parameter(Mandatory = $true,ParameterSetName="DeviceName")]
            [string]$deviceName
        ,[parameter(Mandatory = $true,ParameterSetName="FindCurrentDevice")]
            [Switch]$currentWindowsDevice
        ,[parameter(Mandatory = $false)]
            [Switch]$removeOtherMembers
        ,[parameter(Mandatory = $false)]
            [Switch]$overrideOtherPolicies
        ,[parameter(Mandatory = $false)]
            [Switch]$excludeDeviceAdmins
        ,[parameter(Mandatory = $false)]
            [Switch]$excludeGlobalAdmins
        ,[parameter(Mandatory = $false)]
            [Switch]$includeMe
        )

    $user = get-graphUsers -tokenResponse $tokenResponseTeams -filterUpns $userUPN
    if(!$user){
        Write-Error "[$($userUPN)] could not be resolved to a user in AAD. Cannot create policy."
        return
        }
    
    switch ($PsCmdlet.ParameterSetName){
        "DeviceName"        {
            $device = get-graphDevices -tokenResponse $tokenResponseTeams -filterDisplayNames $deviceName
            if(!$device){
                Write-Error "[$($deviceName)] could not be resolved to a device in AAD. Cannot create policy."
                return
                }
            }
        "FindCurrentDevice" {
            $device = get-graphDevices -tokenResponse $tokenResponseTeams -filterOwnerIds $user.id -filterOperatingSystem Windows | Sort-Object approximateLastSignInDateTime | Select-Object -Last 1
            if(!$device){
                Write-Error "Could not identify [$userUPN]'s most recent Windows device in AAD. Cannot create policy."
                return
                }
            }
        }

    #Test whether AAD group exists
    $namingConvention = "MDM - LocalAdmin - $userUPN"
    $existingGroup = get-graphGroups -tokenResponse $tokenResponseTeams -filterDisplayName $namingConvention
    if($existingGroup -and $removeOtherMembers){
        $existingMembers = get-graphUsersFromGroup -tokenResponse $tokenResponseTeams -groupId $existingGroup.id -memberType Members
        $existingMembers = $existingMembers | ? {$_.userPrincipalName -ne $userUPN}
        if($existingMembers.Count -gt 0){
            remove-graphUsersFromGroup -tokenResponse $tokenResponseTeams -graphGroupId $existingGroup.id -memberType Members -graphUserIds $existingMembers.id
            }
        }
    if($existingGroup){$mdmGroup = $existingGroup}
    else{
        $mdmGroup = new-graphGroup -tokenResponse $tokenResponseTeams -groupDisplayName $namingConvention -groupDescription "Used to assign $userUPN as Local Admin" -groupType Security -membershipType Assigned -groupOwners "00aa81e4-2e8f-4170-bc24-843b917fd7cf" -groupMembers $userUPN
        }

    #Define the new Policy
    if($includeMe){ #This will ad the AAD Device Administrators Role/Group to the Local Admin group
        $additionalAdmins = "
    <member name = `"$(Invoke-Expression "whoami")`""
        }
    if(!$excludeDeviceAdmins){ #This will ad the AAD Device Administrators Role/Group to the Local Admin group
        $additionalAdmins += "
    <member name = `"S-1-12-1-2392505957-1079223134-2636866998-1702916978`""
        }
    if(!$excludeGlobalAdmins){ #This will ad the Global Administrators Role/Group to the Local Admin group
        $additionalAdmins += "
    <member name = `"S-1-12-1-1468013570-1207587608-3581030542-2751513082`""
        }
    $omaSettingsLocalAdmin = @{
        '@odata.type' = "#microsoft.graph.omaSettingString"
        displayName = $namingConvention
        description = "Used to assign $userUPN as Local Admin"
        omaUri = "./Vendor/MSFT/Policy/Config/RestrictedGroups/ConfigureGroupMembership"
        value = "<groupmembership>
<accessgroup desc = `"Administrators`">
    <member name = `"Administrator`" />$deviceAdmins
    <member name = `"AzureAD\$userUPN`" />
</accessgroup>
</groupmembership>"
        }
    $bodyHash = @{
        displayName = $namingConvention
        description = "Used to assign $userUPN as Local Admin"
        omaSettings = @($omaSettingsLocalAdmin)
        }

    #Test whether policy exists
    $existingPolicy = get-mdmPolicyDeviceConfig -tokenResponse $tokenResponseIntune -filterDisplayNameEquals $namingConvention
    if($existingPolicy.Count -gt 1){ #Filter not yet supported for this endpoint
        $existingPolicy = $existingPolicy | ? {$_.displayName -eq $namingConvention}
        }
    if($existingPolicy){
        $mdmPolicy = invoke-graphPatch -tokenResponse $tokenResponseIntune -graphQuery "/deviceManagement/deviceConfigurations/$($existingPolicy.id)" -graphBodyHashtable $bodyHash
        }
    else{
        $bodyHash.Add('@odata.type',"#microsoft.graph.windows10CustomConfiguration") #We only define the objectType if we're creating a new policy
        $mdmPolicy = invoke-graphPost -tokenResponse $tokenResponseIntune -graphQuery "/deviceManagement/deviceConfigurations" -graphBodyHashtable $bodyHash
        }
    
    #Assign Policy to Group
    if($mdmGroup -and $mdmPolicy){
        $target = @{
            '@odata.type'="#microsoft.graph.groupAssignmentTarget"
            groupId = $mdmGroup.id
            }
        $bodyHash = @{
             "@odata.type" = "#microsoft.graph.deviceConfigurationAssignment"
             target = $target
            }
        $assignment = invoke-graphPost -tokenResponse $tokenResponseIntune -graphQuery "/deviceManagement/deviceConfigurations/$($mdmPolicy.id)/assignments" -graphBodyHashtable $bodyHash
        }

    if($overrideOtherPolicies){ #Make sure that this device is not included in any other policies that manage Local Admins (otherwise they will conflict, and not work)
        $otherMdmLocalAdminPolicies = get-mdmPolicyDeviceConfig -tokenResponse $tokenResponseIntune | ? {$_.omaSettings -match "/Vendor/MSFT/Policy/Config/RestrictedGroups/ConfigureGroupMembership"}
        $otherMdmLocalAdminPolicies = $otherMdmLocalAdminPolicies | ? {$_.id -ne $mdmPolicy.id} #Exclude the policy we've just created
        $otherMdmLocalAdminPolicies | Select-Object | % { #Add the Assignments for each Policy as a property of the Policy
            Add-Member -InputObject $_ -MemberType NoteProperty -Name Assignments -Value $(get-mdmPolicyDeviceConfigAssignment -tokenResponse $tokenResponseIntune -configurationId $_.id) -Force
            $_.Assignments | Select-Object | % { #Check the members of each group in each assignment to see whether $deviceId is already affected by another POlicy
            $thisAssignment = $_
                $otherMembers = get-graphUsersFromGroup -tokenResponse $tokenResponseTeams -groupId $thisAssignment.target.groupId -memberType TransitiveMembers
                if($otherMembers.id -match $device.id -and $thisAssignment.target.groupId -ne $mdmGroup.id){ #If we find another Policy affecting $device, try to remove $device from the AAD Group in the Assignment
                    $thisGroup = get-graphGroups -tokenResponse $tokenResponseTeams -filterId $thisAssignment.target.groupId
                    try{
                        Write-Warning "Device [$($device.displayName)][$($device.id)] is a member of Group [$($thisGroup.displayName)][$($thisAssignment.target.groupId)], which would conflict with new Policy [$($mdmPolicy.displayName)][$($mdmPolicy.id)].`nRemoving from old group"
                        remove-graphUsersFromGroup -tokenResponse $tokenResponseTeams -graphGroupId $thisAssignment.target.groupId -memberType Members -graphUserIds $device.id
                        }
                    catch{
                        if((ConvertFrom-Json $Error[0].ErrorDetails.Message).error.code -eq "Request_ResourceNotFound"){
                            Write-Warning "Group [$($thisGroup.displayName)][$($thisAssignment.target.groupId)] does not contain [$($device.displayName)][$($device.id)] as a direct member - they must be a member of a SubGroup, so a human will need to solve this."
                            }
                        }
                    }
                }
            }
        }
    
    $mdmPolicy
    }


