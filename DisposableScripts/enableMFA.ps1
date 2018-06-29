
$auth = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
$auth.RelyingParty = "*"
$auth.State = "Enforced"
$auth.RememberDevicesNotIssuedBefore = (Get-Date)

$users = convertTo-arrayOfEmailAddresses "Alex Matthews <Alex.Matthews@anthesisgroup.com>; Ben Lynch <Ben.Lynch@anthesisgroup.com>; Chris Jennings <Chris.Jennings@anthesisgroup.com>; Duncan Faulkes <Duncan.Faulkes@anthesisgroup.com>; Gavin Way <Gavin.Way@anthesisgroup.com>; Huw Blackwell <Huw.Blackwell@anthesisgroup.com>; Josep Porta <josep.porta@anthesisgroup.com>; Laurie Eldridge <Laurie.Eldridge@anthesisgroup.com>; Matt Landick <Matt.Landick@anthesisgroup.com>; Matthew Gitsham <Matthew.Gitsham@anthesisgroup.com>; Pete Best <Pete.Best@anthesisgroup.com>; Stuart Miller <Stuart.Miller@anthesisgroup.com>; Thomas Milne <Thomas.Milne@anthesisgroup.com>"
$users | % {
    Set-MsolUser -UserPrincipalName $_ -StrongAuthenticationRequirements $auth
    }


Get-MsolUser -UserPrincipalName ben.lynch@anthesisgroup.com | fl
Get-MsolUser -all | ? {$_.StrongAuthenticationRequirements -ne $null -and $_.StrongAuthenticationUserDetails -eq $null}



$users | % {
    Add-DistributionGroupMember -Identity GuineapigsSpamExperimentalGroup@anthesisgroup.com -Member "Rebecca Hughes"
    }

$users = convertTo-arrayOfEmailAddresses "Tom.Mitchell@anthesisgroup.com       Tom Mitchell       True      
Thomas.Milne@anthesisgroup.com       Thomas Milne       True      
James.Walker@anthesisgroup.com       James Walker       True      
Fuchsia.Wildgoose@anthesisgroup.com  Fuchsia Wildgoose  True      
Harry.Shepherd@anthesisgroup.com     Harry Shepherd     True      
Nigel.Arnott@anthesisgroup.com       Nigel Arnott       True      
Kath.Addison-Scott@anthesisgroup.com Kath Addison-Scott True      
Alex.Peers@anthesisgroup.com         Alex Peers         True      
Wai.Cheung@anthesisgroup.com         Wai Cheung         True      
James.Carberry@anthesisgroup.com     James Carberry     True      
Kirsty.Smart@anthesisgroup.com       Kirsty Smart       True      
Alex.Matthews@anthesisgroup.com      Alex Matthews      True      
Amy.MacGrain@anthesisgroup.com"


$km = get-mailbox "kevin.maitland"
$cj = get-mailbox Chris.jennings
$km | fl