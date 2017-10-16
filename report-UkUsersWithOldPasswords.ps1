Import-Module _PS_Library_MSOL
$creds = set-MsolCredentials
connect-ToMsol -credential $creds
connect-ToExo -credential $creds

$alluk = Get-DistributionGroupMember "AllUKexcludingSustain"
$allOxford = Get-DistributionGroupMember "AllOxford"
$allLondon = Get-DistributionGroupMember "AllLondon"
$msolUsers = Get-MsolUser

foreach ($user in $allOxford){
    if (($msolUsers | ?{$_.UserPrincipalName -eq $($user.Name+"@anthesisgroup.com")}).LastPasswordChangeTimestamp -lt $(get-date("1 September 2017 00:00:00"))){[array]$toDoOxford += $user}
    }

rv toDoLondon
foreach ($user in $allLondon){
    if (($msolUsers | ?{$_.UserPrincipalName -eq $($user.Name+"@anthesisgroup.com")}).LastPasswordChangeTimestamp -lt $(get-date("1 September 2017 00:00:00"))){
        [array]$toDoLondon += $user
        Write-Host -ForegroundColor Yellow ($user.Name+"@anthesisgroup.com`t$(($msolUsers | ?{$_.UserPrincipalName -eq $($user.Name+"@anthesisgroup.com")}).LastPasswordChangeTimestamp)" )
        }
    }

$allElse = $alluk | ?{($allOxford.Name -notcontains $_.Name) -and ($allLondon.Name -notcontains $_.Name)}
rv toDoElse
foreach ($user in $allElse){
    if (($msolUsers | ?{$_.UserPrincipalName -eq $($user.Name+"@anthesisgroup.com")}).LastPasswordChangeTimestamp -lt $(get-date("1 September 2017 00:00:00"))){
        [array]$toDoElse += $user
        Write-Host -ForegroundColor Yellow ($user.Name+"@anthesisgroup.com`t$(($msolUsers | ?{$_.UserPrincipalName -eq $($user.Name+"@anthesisgroup.com")}).LastPasswordChangeTimestamp)" )
        }
    }

Write-Host -ForegroundColor Magenta "Oxford:"
Write-Host -ForegroundColor Magenta $toDoOxford
Write-Host -ForegroundColor Cyan "London:"
Write-Host -ForegroundColor Cyan $toDoLondon
Write-Host -ForegroundColor Yellow "AllElse:"
Write-Host -ForegroundColor Yellow $toDoElse

$toDoOxford | select WindowsLiveID
$toDoLondon | select WindowsLiveID, LastPasswordChangeTimestamp
$toDoElse | select WindowsLiveID

$toDoOxford[0] | fl

$users = @("Graeme.Hadley@anthesisgroup.com","Pravin.Selvarajah@anthesisgroup.com","Dee.Moloney@anthesisgroup.com","richard.peagam@anthesisgroup.com","Laura.Thompson@anthesisgroup.com","Hannah.Dick@anthesisgroup.com","Ellen.Struthers@anthesisgroup.com","debbie.hitchen@anthesisgroup.com","Beth.Simpson@anthesisgroup.com","Andy.Marsh@anthesisgroup.com","Ellen.Upton@anthesisgroup.com","Jono.Adams@anthesisgroup.com","Thomas.Milne@anthesisgroup.com","Ben.Diallo@anthesisgroup.com","Laurie.Eldridge@anthesisgroup.com")
$users = @("brad.blundell@anthesisgroup.com","Paul.Dornan@anthesisgroup.com","Sara.Angrill.Toledo@anthesisgroup.com","Pearl.Nemeth@anthesisgroup.com","Phil.Harrison@anthesisgroup.com","Andrew.Noone@anthesisgroup.com","Peter.Scholes@anthesisgroup.com","Claudia.Amos@anthesisgroup.com","Sophie.Martin@anthesisgroup.com","Nadeem.Butt@anthesisgroup.com","Matt.Rooney@anthesisgroup.com")
rv toDoLondon
foreach ($user in $users){
     if (($msolUsers | ?{$_.UserPrincipalName -eq $user}).LastPasswordChangeTimestamp -lt $(get-date("1 September 2017 00:00:00"))){
        [array]$toDoLondon += $user
        Write-Host -ForegroundColor Yellow ($user + $(($msolUsers | ?{$_.UserPrincipalName -eq $user}).LastPasswordChangeTimestamp))
        }
    
    }

foreach ($slacker in $toDoLondon){
    Write-Host -ForegroundColor Yellow "Resetting: $slacker"
    Set-MsolUserPassword -UserPrincipalName $slacker -NewPassword "SorryAboutThis!" -ForceChangePassword $true
    }