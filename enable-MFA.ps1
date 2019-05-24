$logFileLocation = "C:\ScriptLogs\"
$transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_Transcript_$(Get-Date -Format "yyMMdd").log"
if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)){
    $fullLogPathAndName = $logFileLocation+"enableMFA_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = $logFileLocation+"enableMFA_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
else{
    $fullLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
Start-Transcript $transcriptLogName -Append

Import-Module _PS_Library_GeneralFunctionality
Import-Module _PS_Library_MSOL

$Admin = "kevin.maitland@anthesisgroup.com"
$AdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\Kev.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin, $AdminPass

connect-ToMsol -credential $adminCreds

$auth = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
$auth.RelyingParty = "*"
$auth.State = "Enabled"
$auth.RememberDevicesNotIssuedBefore = (Get-Date)

#$users = convertTo-arrayOfEmailAddresses "Matt Whitehead <Matt.Whitehead@anthesisgroup.com>; Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>; Mark Sayers <Mark.Sayers@anthesisgroup.com>; Amy Dartington <Amy.Dartington@anthesisgroup.com>; Debra Haylings <debra.haylings@anthesisgroup.com>; Michael Kirk-Smith <Michael.Kirk-Smith@anthesisgroup.com>"
#$users = convertTo-arrayOfEmailAddresses "Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>; Debra Haylings <debra.haylings@anthesisgroup.com>; Michael Kirk-Smith <Michael.Kirk-Smith@anthesisgroup.com>"
#$users = convertTo-arrayOfEmailAddresses "Alex McKay <alex.mckay@anthesisgroup.com>; Fiona Place <Fiona.Place@anthesisgroup.com>; Andrew Noone <Andrew.Noone@anthesisgroup.com>; Chris Morris <Chris.Morris@anthesisgroup.com>; Harriet Bell <Harriet.Bell@anthesisgroup.com>; Jennifer Wilson <jennifer.wilson@anthesisgroup.com>; Graeme Hadley <Graeme.Hadley@anthesisgroup.com>; Ben Tuxworth <Ben.Tuxworth@anthesisgroup.com>; Claire Richards <Claire.Richards@anthesisgroup.com>; James MacPherson <James.MacPherson@anthesisgroup.com>"
#$users = convertTo-arrayOfEmailAddresses "Chris Stanley <Chris.Stanley@anthesisgroup.com>; Chris Turner <Chris.Turner@anthesisgroup.com>; Helen Kean <Helen.Kean@anthesisgroup.com>; Ian Forrester <Ian.Forrester@anthesisgroup.com>; Jessica Onyshko <jessica.onyshko@anthesisgroup.com>; Karen Cooksey <Karen.Cooksey@anthesisgroup.com>; Paul Ashford <Paul.Ashford@anthesisgroup.com>; Paul Dornan <Paul.Dornan@anthesisgroup.com>; Pearl Németh <Pearl.Nemeth@anthesisgroup.com>; Terry Wood <Terry.Wood@anthesisgroup.com>"
#$users = convertTo-arrayOfEmailAddresses "Alan Spray <Alan.Spray@anthesisgroup.com>; Alec Burslem <Alec.Burslem@anthesisgroup.com>; Chloe McCloskey <Chloe.McCloskey@anthesisgroup.com>; Claire Stentiford <Claire.Stentiford@anthesisgroup.com>; Eleanor Penney <Eleanor.Penney@anthesisgroup.com>; Matt Fishwick <Matt.Fishwick@anthesisgroup.com>; Michael Kirk-Smith <Michael.Kirk-Smith@anthesisgroup.com>; Sophie Sapienza <Sophie.Sapienza@anthesisgroup.com>; Tecla Castella <Tecla.Castella@anthesisgroup.com>"
#$users = convertTo-arrayOfEmailAddresses "Mark Hawker <Mark.Hawker@anthesisgroup.com>; Heather Ball <Heather.Ball@anthesisgroup.com>; Jaime Dingle <Jaime.Dingle@anthesisgroup.com>; Tharaka Naga <Tharaka.Naga@anthesisgroup.com>; Ashwini Arul <Ashwini.Arul@anthesisgroup.com>; Alan Dow <Alan.Dow@anthesisgroup.com>; Matt Rooney <Matt.Rooney@anthesisgroup.com>; Sarah Gilby <Sarah.Gilby@anthesisgroup.com>; Tim Clare <Tim.Clare@anthesisgroup.com>"
#$users = convertTo-arrayOfEmailAddresses "Bethany Munyard <Bethany.Munyard@anthesisgroup.com>; Ellen Upton <Ellen.Upton@anthesisgroup.com>; Jono Adams <Jono.Adams@anthesisgroup.com>; Polly Stebbings <Polly.Stebbings@anthesisgroup.com>; Alan Matthews <Alan.Matthews@anthesisgroup.com>; Dee Moloney <Dee.Moloney@anthesisgroup.com>; Enda Colfer <Enda.Colfer@anthesisgroup.com>; Ian Bailey <Ian.Bailey@anthesisgroup.com>; Paul Crewe <Paul.Crewe@anthesisgroup.com>; Anne O’Brien <Anne.OBrien@anthesisgroup.com>; Beth Simpson <Beth.Simpson@anthesisgroup.com>; Julian Parfitt <Julian.Parfitt@anthesisgroup.com>; Nick Cuomo <Nick.Cuomo@anthesisgroup.com>; Peter Scholes <Peter.Scholes@anthesisgroup.com>; Simone Aplin <Simone.Aplin@anthesisgroup.com>; Stephanie Egee <Stephanie.Egee@anthesisgroup.com>"
$users = convertTo-arrayOfEmailAddresses "Pernilla.Holgersson@anthesisgroup.com
Maria.Hammar@anthesisgroup.com
"
$ssprGroup = Get-MsolGroup -SearchString "SSPR Testers"

$users | % {
    $thisUser = Get-MsolUser -UserPrincipalName $_
    Write-Host -ForegroundColor DarkYellow "MFA is currently set to $($thisUser.StrongAuthenticationRequirements.State) for $_"
    if([string]::IsNullOrWhiteSpace($thisUser.StrongAuthenticationRequirements)){
        Write-Host -ForegroundColor Yellow "Enabling MFA for $_"
        Set-MsolUser -UserPrincipalName $thisUser.UserPrincipalName -StrongAuthenticationRequirements $auth
        }
    else{Write-Host -ForegroundColor DarkYellow "MFA already $($thisUser.StrongAuthenticationRequirements.State) for $_"}
    Add-MsolGroupMember -GroupObjectId $ssprGroup.ObjectId -GroupMemberType User -GroupMemberObjectId $thisUser.ObjectId
    }

Stop-Transcript
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


