[cmdletbinding()]
param(
    )
    <#
    .Synopsis
        Removes any other group memberships from users who are a member of [All COntractors]
    .DESCRIPTION
        Enumerates members of [All Contractors], then removes them from all other groups (to prevent granting excessive permissions)
    .EXAMPLE
       reset-contractorGroupMembership
    #>

$logFileLocation = "C:\ScriptLogs\"
$logFileName = "reset-contractorGroupMembership"
$fullLogPathAndName = $logFileLocation+$logFileName+"_FullLog_$(Get-Date -Format "yyMMdd").log"
$errorLogPathAndName = $logFileLocation+$logFileName+"_ErrorLog_$(Get-Date -Format "yyMMdd").log"
$smtpServer = "anthesisgroup-com.mail.protection.outlook.com"
$mailFrom = "$logFileName`_$env:COMPUTERNAME@anthesisgroup.com"
$mailTo = "kevin.maitland@anthesisgroup.com"
if($PSCommandPath){
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_$objectType`_Transcript_$(Get-Date -Format "yyMMdd").log"
    Start-Transcript $transcriptLogName -Append
    }

Import-Module _PS_Library_GeneralFunctionality
Import-Module _PS_Library_MSOL

$365admin = "groupbot@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
$365AdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Downloads\GroupBot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $365admin, $365AdminPass
Connect-AzureAD -Credential $adminCreds
connect-ToExo -credential $adminCreds

$groupMembershipExceptions = @()
$subbieExceptions = @()
$subcontractorMesg = Get-DistributionGroup cff78155-8974-44f3-b0b3-f9d2b3b53a23 #"All Subcontractors"
$mdmByodUsers = Get-DistributionGroup b264f337-ef04-432e-a139-3574331a4d18 #"MDM - BYOD Mobile Device Users"
$groupMembershipExceptions += $subcontractorMesg.ExternalDirectoryObjectId
$groupMembershipExceptions += $mdmByodUsers.ExternalDirectoryObjectId
#$subbieExceptions += "36bc6f20-feed-422d-b2f2-7758e9708604" # $(Get-User "kevin.maitland").ExternalDirectoryObjectId

$subbies = Get-DistributionGroupMember $subcontractorMesg.ExternalDirectoryObjectId | ? {$subbieExceptions -notcontains $_.ExternalDirectoryObjectId}
log-action "[$(if([string]::IsNullOrWhiteSpace($subbies.Count)){1}else{$subbies.Count})] Subbies found (excluding [$(if([string]::IsNullOrWhiteSpace($subbieExceptions.Count)){1}else{$subbieExceptions.Count})] exceptions) - will remove them from all groups except [$($groupMembershipExceptions -join ",")]" -logFile $fullLogPathAndName
$subbies | % {
    $thisUser = $_
    $thisUsersMemberships = Get-AzureADUserMembership -ObjectId $thisUser.ExternalDirectoryObjectId
    $thisUsersMemberships | ? {$groupMembershipExceptions -notcontains $_.ObjectId -and $_.ObjectType -eq "Group"} | % {
        $thisGroup = $_
        if($thisGroup.MailEnabled -eq $false){ #Security groups
            try{
                log-action -myMessage "Removing [$($thisUser.DisplayName)] from Security Group [$($thisGroup.DisplayName)]" -logFile $fullLogPathAndName
                Send-MailMessage -Subject "Removing Contractor [$($thisUser.DisplayName)] from Security Group [$($thisGroup.DisplayName)]" -From $mailFrom -To $mailTo -SmtpServer $smtpServer
                Remove-AzureADGroupMember -ObjectId $thisGroup.ObjectId -MemberId $thisUser.ExternalDirectoryObjectId
                log-result -myMessage "Success!" -logFile $fullLogPathAndName
                }
            catch{
                log-error -myError $_ -myFriendlyMessage "Error removing [$($thisUser.DisplayName)] from Security Group [$($thisGroup.DisplayName)] in $logFileName" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
                }
            }
        elseif($thisGroup.ProxyAddresses -match "SPO:"){ #All 365 groups created after ~ 2018-01
            try{
                log-action -myMessage "Removing [$($thisUser.DisplayName)] from 365 Group [$($thisGroup.DisplayName)]" -logFile $fullLogPathAndName
                Send-MailMessage -Subject "Removing Contractor [$($thisUser.DisplayName)] from 365 Group [$($thisGroup.DisplayName)]" -From $mailFrom -To $mailTo -SmtpServer $smtpServer
                Remove-UnifiedGroupLinks -Identity $thisGroup.ObjectId -LinkType Member -Links $thisUser.ExternalDirectoryObjectId  
                log-result -myMessage "Success!" -logFile $fullLogPathAndName
                }
            catch{
                log-error -myError $_ -myFriendlyMessage "Error removing [$($thisUser.DisplayName)] from 365 Group [$($thisGroup.DisplayName)] in $logFileName" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
                }
            }
        else{ #Mail-enabled security groups & Distribution groups
            try{
                log-action -myMessage "Removing [$($thisUser.DisplayName)] from Distribution Group [$($thisGroup.DisplayName)]" -logFile $fullLogPathAndName
                Send-MailMessage -Subject "Removing Contractor [$($thisUser.DisplayName)] from Distribution Group [$($thisGroup.DisplayName)]" -From $mailFrom -To $mailTo -SmtpServer $smtpServer
                Remove-DistributionGroupMember -Identity $thisGroup.ObjectId -Member $thisUser.ExternalDirectoryObjectId -BypassSecurityGroupManagerCheck
                log-result -myMessage "Success!" -logFile $fullLogPathAndName
                }
            catch{
                log-error -myError $_ -myFriendlyMessage "Error removing [$($thisUser.DisplayName)] from Distribution Group [$($thisGroup.DisplayName)] in $logFileName" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName
                }
            }
        }
    }
