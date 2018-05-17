
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

$users = convertTo-arrayOfStrings "Alex Peers
Amy MacGrain
Chris Jennings
Curtis Harnanan
Duncan Faulkes
Harry Shepherd
Jack Dodd Sachdev
James Carberry
Jennifer Clipsham
Josep Porta
Kath Addison-Scott
Kev Maitland
Lorna Kelly
Margaret Davis
Mary Short
Matt Whitehead
Matthew Gitsham
Nigel Arnott
Pearl Gyongi
Rebecca Hughes
Stuart Gray
Stuart Miller
Thomas Milne
Wai Cheung"