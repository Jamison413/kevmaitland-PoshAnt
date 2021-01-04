param(
    [CmdletBinding()]
    [parameter(Mandatory = $true)]
    [ValidateNotNull()]
    [ValidatePattern(".[@].")]
    [string]$upnsString
    ,[parameter(Mandatory = $false)]
    [pscredential]$authenticationAdminCreds
    )

    Clara.Borras@anthesisgroup.com ; Daniela.Romero@anthesisgroup.com  ; Gabriela.Peixoto@anthesisgroup.com  ; Marina.Clara@anthesisgroup.com  ; Neus.Cardona@anthesisgroup.com ; Noelia.Lopez@anthesisgroup.com 

Import-Module _PS_Library_GeneralFunctionality
Import-Module _PS_Library_MSOL

#Start logging
$logFileLocation = "C:\ScriptLogs\"
$transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_Transcript_$(Get-Date -Format "yyMMdd").log"
Start-Transcript $transcriptLogName -Append
$fullLogPathAndName = $logFileLocation+"enableMFA_FullLog_$(Get-Date -Format "yyMMdd").log"
$errorLogPathAndName = $logFileLocation+"enableMFA_ErrorLog_$(Get-Date -Format "yyMMdd").log"

#Connect as Authenication Administrator
if($authenticationAdminCreds -eq $null){
    $Admin = "kevin.maitland@anthesisgroup.com"
    $AdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\Kev.txt) 
    $adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin, $AdminPass
    }
else{
    $adminCreds = $authenticationAdminCreds
    }
connect-ToMsol -Credential $adminCreds


#########################################
#                                       #
#             Enable MFA                #
#                                       #
#########################################

#Create an empty StrongAuthenticationRequirement object
$emptyAuthObject = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
$emptyAuthObject.RelyingParty = "*"
$emptyAuthObject.State = "Enabled"
$emptyAuthObject.RememberDevicesNotIssuedBefore = (Get-Date)

#Get the GUID for the SSPR Group
#$ssprGroup = Get-MsolGroup -SearchString "SSPR Testers"
[guid]$ssprGroupObjectId = "fee80bd5-6e2f-4888-a51c-9581cf64eb18" #This is the GUID for the SSPR Testers Group


#Figure out who to run this for
$upnsToEnable = convertTo-arrayOfEmailAddresses $upnsString


$upnsToEnable | % {
    $thisUser = Get-MsolUser -UserPrincipalName $_
    Write-Verbose "MFA is currently set to [$($thisUser.StrongAuthenticationRequirements.State)] for $_"
    if([string]::IsNullOrWhiteSpace($thisUser.StrongAuthenticationRequirements)){
        Write-Verbose "Enabling MFA for $_"
        Set-MsolUser -UserPrincipalName $thisUser.UserPrincipalName -StrongAuthenticationRequirements $emptyAuthObject
        }
    else{Write-Verbose "MFA already [$($thisUser.StrongAuthenticationRequirements.State)] for $_"}
    Add-MsolGroupMember -GroupObjectId $ssprGroupObjectId -GroupMemberType User -GroupMemberObjectId $thisUser.ObjectId
    }
Stop-Transcript



#########################################
#                                       #
#             Disable MFA               #
#                                       #
#########################################


#As long as the StrongAuthenticationRequirements is empty, it is disabled, click through form "More information needed" on login and press "Skip Setup" in the bottom right corner of the MFA screen.


$upnsToDisable = convertTo-arrayOfEmailAddresses $upnsString

$upnsToDisable | % {
    $thisUser = Get-MsolUser -UserPrincipalName $upnsToDisable
    Set-MsolUser -UserPrincipalName $thisUser.UserPrincipalName -StrongAuthenticationRequirements @()
    $thisUser = Get-MsolUser -UserPrincipalName $upnsToDisable
    $thisUser | Format-List | select -Property StrongAuthenticationRequirements #if blank it is successful

}
Stop-Transcript



$e = Get-MsolUser -UserPrincipalName "Cora.Philpott@anthesisgroup.com"
$e.StrongAuthenticationUserRequirements


<#
$disabledUsers | % {
    $thisUser = $_
    Write-Host -ForegroundColor DarkYellow "MFA is currently set to [$($thisUser.StrongAuthenticationRequirements.State)] for [$($thisUser.DisplayName)]"
    if([string]::IsNullOrWhiteSpace($thisUser.StrongAuthenticationRequirements)){
        Write-Host -ForegroundColor Yellow "Enabling MFA for [$($thisUser.DisplayName)]"
        Set-MsolUser -UserPrincipalName $thisUser.UserPrincipalName -StrongAuthenticationRequirements $auth
        }
    else{Write-Host -ForegroundColor DarkYellow "MFA already $($thisUser.StrongAuthenticationRequirements.State) for [$($thisUser.DisplayName)]"}
    
    }
  
<#
Get-MsolUser -UserPrincipalName ben.lynch@anthesisgroup.com | fl
$allUsers = Get-MsolUser -all 
$allUsers | ? {$_.StrongAuthenticationRequirements -ne $null -and $_.StrongAuthenticationUserDetails -eq $null}
$allUsers | ? {$_.DisplayName -match "Greg Francis"} | fl
$allUsers[0]


$users | % {
    Add-DistributionGroupMember -Identity GuineapigsSpamExperimentalGroup@anthesisgroup.com -Member "Rebecca Hughes"
    }






$allUsers = Get-MsolUser -all 
$msolg = Get-MsolGroup -All 
$msolg | ? {$_.DisplayName -notmatch "∂" }| % {
    $thisGroup = $_
    $members = Get-MsolGroupMember -GroupObjectId $thisGroup.ObjectId
    $members | ? {$_.GroupMemberType -eq "User"} | %{
        $thisMember = $_
        if($($allUsers | ? {$_.userprincipalname -eq $thisMember.EmailAddress}).IsLicensed){
            $detailObject = New-Object psobject -Property @{
                "DisplayName" = $thisMember.DisplayName;
                "Email" = $thisMember.EmailAddress;
                "Group" = $thisGroup.DisplayName
                "GroupType" = $thisGroup.GroupType
                "MfaStatus" = $($allUsers | ? {$_.userprincipalname -eq $thisMember.EmailAddress}).StrongAuthenticationRequirements.State
                "MfaOptions" =  $($allUsers | ? {$_.userprincipalname -eq $thisMember.EmailAddress}).StrongAuthenticationUserDetails
                }
            [array]$mfaDetails += $detailObject
            }
        }
    }

$mfaDetails | Export-Csv  $env:USERPROFILE\Desktop\MfaStatus_$(Get-Date -Format "yyMMdd").csv -NoTypeInformation


get-help Set-MailboxAutoReplyConfiguration -Detailed
Set-MailboxAutoReplyConfiguration -Identity Shared_Mailbox_Bodge_-_Finance_Team_GBR_-_Energy -AutoReplyState enabled -ExternalMessage "Thank you for your email.
The Finance Department will process any invoices and respond to any enquires within 2-3 working days.
Our company name has changed to Anthesis Energy UK Ltd and our email addresses have changed as well. Please update your records.
Invoices to be sent to energyinvoices@anthesisgroup.com
Remittances advices to energyremittances@anthesisgroup.com
Enquiries and statements to energyfinance@anthesisgroup.com
If you have any queries then please contact kath.addison-scott@anthesisgroup.com or greg.francis@anthesisgroup.com

Kind Regards,
Anthesis  Energy UK's AutoReply Robot"




$gb = Get-MsolUser -all | ?{($_.Country -eq "United Kingdom" -or $_.UsageLocation -eq "GB") -and $_.IsLicensed -eq $true}


$gb | %{
    Write-Host $_.DisplayName`t $_.UserPrincipalName`t$_.Country`t$_.UsageLocation`t $_.StrongAuthenticationRequirements[0].State`t $($_.StrongAuthenticationMethods | ?{$_.IsDefault}).MethodType
    }

Get-MsolUser -all | ? {$_.IsLicensed -eq $true} | %{
    Write-Host $_.DisplayName`t$($_.UserPrincipalName)`t$($_.Country)`t$($_.UsageLocation)`t $_.StrongAuthenticationRequirements[0].State`t $($_.StrongAuthenticationMethods | ?{$_.IsDefault}).MethodType
    }
#>


