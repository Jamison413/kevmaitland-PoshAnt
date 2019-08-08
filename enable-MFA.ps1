param(
    [CmdletBinding()]
    [parameter(Mandatory = $true)]
    [ValidateNotNull()]
    [ValidatePattern(".[@].")]
    [string]$upnsString
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
$Admin = "kevin.maitland@anthesisgroup.com"
$AdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\Kev.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin, $AdminPass
connect-ToMsol -Credential $adminCreds

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
