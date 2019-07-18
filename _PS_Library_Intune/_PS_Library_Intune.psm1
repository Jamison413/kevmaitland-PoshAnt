
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

