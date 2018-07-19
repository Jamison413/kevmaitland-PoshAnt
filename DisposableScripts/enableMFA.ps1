
$auth = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
$auth.RelyingParty = "*"
$auth.State = "Enabled"
$auth.RememberDevicesNotIssuedBefore = (Get-Date)

$users = convertTo-arrayOfEmailAddresses "Matt Whitehead <Matt.Whitehead@anthesisgroup.com>; Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>; Mark Sayers <Mark.Sayers@anthesisgroup.com>; Amy Dartington <Amy.Dartington@anthesisgroup.com>; Debra Haylings <debra.haylings@anthesisgroup.com>; Michael Kirk-Smith <Michael.Kirk-Smith@anthesisgroup.com>"
$users = convertTo-arrayOfEmailAddresses "Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>; Debra Haylings <debra.haylings@anthesisgroup.com>; Michael Kirk-Smith <Michael.Kirk-Smith@anthesisgroup.com>"
$users = convertTo-arrayOfEmailAddresses "Alex McKay <alex.mckay@anthesisgroup.com>; Fiona Place <Fiona.Place@anthesisgroup.com>; Andrew Noone <Andrew.Noone@anthesisgroup.com>; Chris Morris <Chris.Morris@anthesisgroup.com>; Harriet Bell <Harriet.Bell@anthesisgroup.com>; Jennifer Wilson <jennifer.wilson@anthesisgroup.com>; Graeme Hadley <Graeme.Hadley@anthesisgroup.com>; Ben Tuxworth <Ben.Tuxworth@anthesisgroup.com>; Claire Richards <Claire.Richards@anthesisgroup.com>; James MacPherson <James.MacPherson@anthesisgroup.com>"
$users = convertTo-arrayOfEmailAddresses "Chris Stanley <Chris.Stanley@anthesisgroup.com>; Chris Turner <Chris.Turner@anthesisgroup.com>; Helen Kean <Helen.Kean@anthesisgroup.com>; Ian Forrester <Ian.Forrester@anthesisgroup.com>; Jessica Onyshko <jessica.onyshko@anthesisgroup.com>; Karen Cooksey <Karen.Cooksey@anthesisgroup.com>; Paul Ashford <Paul.Ashford@anthesisgroup.com>; Paul Dornan <Paul.Dornan@anthesisgroup.com>; Pearl Németh <Pearl.Nemeth@anthesisgroup.com>; Terry Wood <Terry.Wood@anthesisgroup.com>"
$users = convertTo-arrayOfEmailAddresses "Alan Spray <Alan.Spray@anthesisgroup.com>; Alec Burslem <Alec.Burslem@anthesisgroup.com>; Chloe McCloskey <Chloe.McCloskey@anthesisgroup.com>; Claire Stentiford <Claire.Stentiford@anthesisgroup.com>; Eleanor Penney <Eleanor.Penney@anthesisgroup.com>; Matt Fishwick <Matt.Fishwick@anthesisgroup.com>; Michael Kirk-Smith <Michael.Kirk-Smith@anthesisgroup.com>; Sophie Sapienza <Sophie.Sapienza@anthesisgroup.com>; Tecla Castella <Tecla.Castella@anthesisgroup.com>"

$users | % {
    $thisUser = Get-MsolUser -UserPrincipalName $_
    Write-Host -ForegroundColor DarkYellow "MFA is currently set to $($thisUser.StrongAuthenticationRequirements.State) for $_"
    if([string]::IsNullOrWhiteSpace($thisUser.StrongAuthenticationRequirements)){
        Write-Host -ForegroundColor Yellow "Enabling MFA for $_"
        Set-MsolUser -UserPrincipalName $_ -StrongAuthenticationRequirements $auth
        }
    else{Write-Host -ForegroundColor DarkYellow "MFA already $($thisUser.StrongAuthenticationRequirements.State) for $_"}
    
    }
    

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