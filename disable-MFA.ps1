param(
    [CmdletBinding()]
    [parameter(Mandatory = $true)]
    [ValidateNotNull()]
    [ValidatePattern(".[@].")]
    [string]$upnsString
    ,[parameter(Mandatory = $false)]
    [pscredential]$authenticationAdminCreds
    )

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

#Figure out who to run this for
$upnsToDisable = convertTo-arrayOfEmailAddresses $upnsString


$upnsToDisable | % {
    $thisUser = Get-MsolUser -UserPrincipalName $_
    Write-Verbose "MFA is currently set to [$($thisUser.StrongAuthenticationRequirements.State)] for $_"

$emptyAuthObject = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
$emptyAuthObject.RelyingParty = $thisUser.StrongAuthenticationRequirements[0].RelyingParty
$emptyAuthObject.State = "Disabled"
$emptyAuthObject.RememberDevicesNotIssuedBefore = $thisUser.StrongAuthenticationRequirements[0].RememberDevicesNotIssuedBefore
Write-Verbose $emptyAuthObject
#$emptyAuthObject.ExtensionData = $thisUser.StrongAuthenticationRequirements.ExtensionData
    if([string]::IsNullOrWhiteSpace($thisUser.StrongAuthenticationRequirements)){
        Write-Verbose "MFA already [$($thisUser.StrongAuthenticationRequirements.State)] for $_"
        }
    else{
        Write-Verbose "Disabling MFA for $_"
        Set-MsolUser -UserPrincipalName $_ -StrongAuthenticationRequirements $emptyAuthObject
        }
    }

Stop-Transcript

<#Disable MFA for specific user#>


