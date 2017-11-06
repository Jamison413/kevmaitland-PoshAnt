Import-Module _PS_Library_MSOL
$msolCreds = set-MsolCredentials
connect-ToMsol -credential $msolCreds
connect-ToExo -credential $msolCreds

$alluk = Get-DistributionGroupMember "AllUKexcludingSustain"
$allOxford = Get-DistributionGroupMember "AllOxford"
$allLondon = Get-DistributionGroupMember "AllLondon"
$msolUsers = Get-MsolUser -All

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

rv toDoElse
$allElse = $alluk | ?{($allOxford.Name -notcontains $_.Name) -and ($allLondon.Name -notcontains $_.Name)}
foreach ($user in $allElse){
    if (($msolUsers | ?{$_.UserPrincipalName -eq $($user.Name+"@anthesisgroup.com")}).LastPasswordChangeTimestamp -lt $(get-date("1 September 2017 00:00:00"))){[array]$toDoElse += $user}
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
$slackers= @("andrea.smerek@anthesisgroup.com","sherwood.li@anthesisgroup.com","Jae.Ryu@anthesisgroup.com","Adam.Wheeler@anthesisgroup.com","Chris.Jones@anthesisgroup.com","DeAnn.Sarver@anthesisgroup.com","jeff.gibbons@anthesisgroup.com","jessica.onyshko@anthesisgroup.com","Jill.Stoneberg@anthesisgroup.com","john.heckman@anthesisgroup.com","marca.hagenstad@anthesisgroup.com","Arul.Subra@anthesisgroup.com")
$users = @("czech@anthesisgroup.com","Finance.Support@anthesisgroup.com","France@anthesisgroup.com","George.Davey@anthesisgroup.com","Italy@anthesisgroup.com","Markus.Kamila@anthesisgroup.com","Paul.Crewe@anthesisgroup.com","Spain@anthesisgroup.com","Tharaka.Naga@anthesisgroup.com")
$users = @("czech@anthesisgroup.com","andrew.hennig@anthesisgroup.com","AnthesisUKFinance@anthesisgroup.com","Kevin.Lewis@anthesisgroup.com","Mike.Hoggan@anthesisgroup.com","Sinead.Fenton@anthesisgroup.com","UKcareers@anthesisgroup.com")

rv toDoLondon
foreach ($user in $users){
     if (($msolUsers | ?{$_.UserPrincipalName -eq $user}).LastPasswordChangeTimestamp -lt $(get-date("27 September 2017 00:00:00"))){
        [array]$toDoLondon += $user
        Write-Host -ForegroundColor Yellow ($user + $(($msolUsers | ?{$_.UserPrincipalName -eq $user}).LastPasswordChangeTimestamp))
        }
    }

foreach ($slacker in $toDoLondon){
    Write-Host -ForegroundColor Yellow "Resetting: $slacker"
    Set-MsolUserPassword -UserPrincipalName $slacker -NewPassword "SorryAboutThis!" -ForceChangePassword $true
    }


foreach($user in $($msolUsers|?{$_.IsLicensed -eq $true})){
    Write-Host -ForegroundColor Yellow $user.DisplayName";"$user.userprincipalname";"$user.Office";"$user.Country";"$user.LastPasswordChangeTimestamp # | select DisplayName,Office,LastPasswordChangeTimestamp
    }

$msolUsers | ? {$_.IsLicensed -eq $true} | select DisplayName, userPrincipalName, Office, Country, LastPasswordChangeTimestamp | Export-Csv -Path $env:USERPROFILE\Desktop\AntUsersPasswords9.csv -NoClobber -NoTypeInformation -Encoding UTF8
$msolUsers[10].LastPasswordChangeTimestamp

$huw = Get-MsolUser -UserPrincipalName huw.blackwell@anthesisgroup.com | fl
$huw.StrongAuthenticationRequirements.State




$users2 = Get-msoluser -All 
$users2 | select DisplayName,@{N='Email';E={$_.UserPrincipalName}},@{N='StrongAuthenticationRequirements';E={($_.StrongAuthenticationRequirements.State)}} | Export-Csv -NoTypeInformation $env:USERPROFILE\desktop\dummy.csv

Get-MsolUser -All | ? {$_.IsLicensed -eq $true -and $_.LastPasswordChangeTimestamp -lt "2017-09-27"} | select DisplayName, userPrincipalName, Office, Country, LastPasswordChangeTimestamp | ft
$slackers = Get-MsolUser -All | ? {$_.IsLicensed -eq $true -and $_.LastPasswordChangeTimestamp -lt "2017-09-27"} 
foreach ($slacker in $slackers){
    Write-Host -ForegroundColor Yellow "Resetting: $($slacker.DisplayName)"
    Set-MsolUserPassword -UserPrincipalName $slacker.UserPrincipalName -NewPassword "SorryAboutThis!" -ForceChangePassword $true
    }
