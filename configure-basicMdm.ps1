param(
    # Specifies whether we are working with Android, Apple or Windows devices/users.
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Android", "Apple","Windows")]
    [string]$deviceType

    # Specifies whether we add the user to the Corporate group or not.
    ,[Parameter(Mandatory = $false, Position = 1)]
    [ValidateSet($true,$false)]
    [bool]$addToCorporateGroup
    )
    <#
    .Synopsis
        Sets group memberships and service statuses for users with Corporate-owned Mobile Devices
    .DESCRIPTION
        Moves users from the corporate enrollment group for their device type to "MDM Coprporate Device Users" and disables deprecated connection methods (e.g. ActiveSync, IMAP, etc.)
    .EXAMPLE
       configure-basicMdm -deviceType "Android" -addToCorporateGroup $true
    #>

$logFileLocation = "C:\ScriptLogs\"
$logFileName = "configure-basicMdm"
$fullLogPathAndName = $logFileLocation+$logFileName+"_$deviceType`_FullLog_$(Get-Date -Format "yyMMdd").log"
$errorLogPathAndName = $logFileLocation+$logFileName+"_$deviceType`_ErrorLog_$(Get-Date -Format "yyMMdd").log"
$smtpServer = "anthesisgroup-com.mail.protection.outlook.com"
$mailFrom = "$logFileName`_$env:COMPUTERNAME@anthesisgroup.com"
$mailTo = "kevin.maitland@anthesisgroup.com"
if($PSCommandPath){
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_$objectType`_Transcript_$(Get-Date -Format "yyMMdd").log"
    Start-Transcript $transcriptLogName -Append
    }

Import-Module _PS_Library_GeneralFunctionality
Import-Module _PS_Library_MSOL

$intuneAdmin = "kevin.maitland@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
$intuneAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\Kev.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $intuneAdmin, $intuneAdminPass
Connect-AzureAD -Credential $adminCreds
connect-ToExo -credential $adminCreds

switch ($deviceType){
    "Android" {
        log-action "Processing Android users" -logFile $fullLogPathAndName
        $mdmEnrollmentGroup = Get-AzureADGroup -ObjectId 2a88b3e5-9ce1-467b-87cd-43cdf8d6b8eb # -SearchString "MDM - Corporate User Enrollment Group (remove members once enrolled)"
        log-result "[$($mdmEnrollmentGroup.DisplayName)] retrieved" -logFile $fullLogPathAndName
        $mdmCorporateGroup = Get-DistributionGroup -Identity 22ee9443-1004-43c3-8e5d-88e9445a95d5 #"MDM - Corporate Mobile Device Users"
        log-result "[$($mdmCorporateGroup.DisplayName)] retrieved" -logFile $fullLogPathAndName
        $mdmByodGroup = Get-DistributionGroup -Identity b264f337-ef04-432e-a139-3574331a4d18 #"MDM - BYOD Mobile Device Users"
        log-result "[$($mdmByodGroup.DisplayName)] retrieved" -logFile $fullLogPathAndName

        $usersToProcess = Get-AzureADGroupMember -ObjectId $mdmEnrollmentGroup.ObjectId
        log-result "[$($usersToProcess.Count)] users retrieved to process" -logFile $fullLogPathAndName
        $usersToProcess | % {
            $thisUser = $_
            log-action "Processing [$($thisUser.DisplayName)]" -logFile $fullLogPathAndName
            try{
                if($addToCorporateGroup){
                    try{ #Add to "MDM - Corporate Mobile Device Users"
                        log-action "Adding to [$($mdmCorporateGroup.DisplayName)]" -logFile $fullLogPathAndName
                        Add-DistributionGroupMember -Identity $mdmCorporateGroup.ExternalDirectoryObjectId -Member $thisUser.ObjectId -ErrorAction Stop
                        log-result "[$($thisUser.DisplayName)] successfully added to [$($mdmCorporateGroup.DisplayName)]" -logFile $fullLogPathAndName
                        } 
                    catch{
                        if($_.Exception.HResult -eq -2146233087){log-result "[$($thisUser.DisplayName)] already a member of [$($mdmCorporateGroup.DisplayName)]" -logFile $fullLogPathAndName}
                        else{log-error -myError $_ -myFriendlyMessage "Error adding [$($thisUser.DisplayName)] to [$($mdmCorporateGroup.DisplayName)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName}
                        }
                    }
                else{log-action "`$addToCorporateGroup = `$false - not adding to [$($mdmCorporateGroup.DisplayName)]" -logFile $fullLogPathAndName}
                try{#Add to "MDM - BYOD Mobile Device Users"
                    log-action "Adding to [$($mdmByodGroup.DisplayName)]" -logFile $fullLogPathAndName
                    Add-DistributionGroupMember -Identity $mdmByodGroup.ExternalDirectoryObjectId -Member $thisUser.ObjectId -ErrorAction Stop #Add to "MDM - BYOD Mobile Device Users"
                    log-result "[$($thisUser.DisplayName)] successfully added to [$($mdmByodGroup.DisplayName)]" -logFile $fullLogPathAndName
                    }
                catch{
                    if($_.Exception.HResult -eq -2146233087){log-result "[$($thisUser.DisplayName)] already a member of [$($mdmByodGroup.DisplayName)]" -logFile $fullLogPathAndName}
                    else{log-error -myError $_ -myFriendlyMessage "Error adding [$($thisUser.DisplayName)] to [$($mdmByodGroup.DisplayName)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName}
                    }

                try{# Disable legacy mailbox protocols
                    log-action "Disabling legacy mailbox protocols for [$($thisUser.DisplayName)]" -logFile $fullLogPathAndName
                    $thisMailbox = Get-Mailbox -Identity $thisUser.UserPrincipalName
                    $thisMailbox | Set-CASMailbox -ImapEnabled $false -ActiveSyncEnabled $false -PopEnabled $false -OWAforDevicesEnabled $false -ErrorAction Stop #Disable legacy mailbox protocols to avoid MFA bypass -MAPIEnabled $false
                    log-result "[$($thisUser.DisplayName)] successfully added to [$($mdmByodGroup.DisplayName)]" -logFile $fullLogPathAndName
                    }
                catch{log-error -myError $_ -myFriendlyMessage "Error disabling legacy mailbox protocols for [$($thisUser.DisplayName)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName}

                try{
                    log-action "Removing [$($thisUser.DisplayName)] from [$($mdmEnrollmentGroup.DisplayName)]" -logFile $fullLogPathAndName
                    Remove-AzureADGroupMember -ObjectId $mdmEnrollmentGroup.ObjectId -MemberId $thisUser.ObjectId
                    }
                catch{log-error -myError $_ -myFriendlyMessage "Error removing [$($thisUser.DisplayName)] to [$($mdmEnrollmentGroup.DisplayName)]" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName}
                }
            catch{
                log-error -myError $_ -myFriendlyMessage "Something went wrong processing [$($thisUser.DisplayName)] in [$logFileName]. Check the error log at [$errorLogPathAndName] on [$env:COMPUTERNAME] for full details" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -smtpServer $smtpServer -mailTo $mailTo -mailFrom $mailFrom
                }
            }
        }
    "Apple" {
        
        }
    "Windows" {
        
        }
    }