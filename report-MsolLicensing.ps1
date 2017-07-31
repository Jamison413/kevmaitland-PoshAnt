$oUsers = Get-MsolUser -All | Sort-Object -Property DisplayName
foreach($oUser in $oUsers) {
    #Write-Host -ForegroundColor Yellow $oUser.DisplayName
    for($i=0;$i -lt $oUser.Licenses.Count;$i++){
        #$oUser.Licenses[$i].ServiceStatus
        Write-Host -ForegroundColor DarkYellow $oUser.DisplayName"`t"$oUser.Licenses[$i].AccountSkuId
        }
    }


$Sustainers =  Get-Mailbox -Filter {CustomAttribute1 -eq "Sustain" -and RecipientTypeDetails -eq "UserMailbox"}
$allUsers = Get-MsolUser
$e1Users = ,@()
#$user = $allUsers | ?{$_.DisplayName -eq "Tobias Parker"}
foreach ($user in $allUsers){
    if($($user.Licenses | %{$_.AccountSkuId -contains "AnthesisLLC:STANDARDPACK"}) -eq $true){$e1Users += $user}
    }
$e1Sustainers = ,@()
foreach($e1 in $e1Users){
    if($($Sustainers | %{$_.MicrosoftOnlineServicesID}) -contains $e1.UserPrincipalName){
        $e1Sustainers += $e1
        }
    }
$Sustainers.Count 
$e1Sustainers.Count

$mismatch = ,@()
$mismatch += $Sustainers | ?{($e1Sustainers | %{$_.UserPrincipalName}) -notcontains $_.MicrosoftOnlineServicesID}
$mismatch += $e1Sustainers | ?{($Sustainers | %{$_.MicrosoftOnlineServicesID}) -notcontains $_.UserPrincipalName}

foreach($user in $mismatch){
    $msolUser = Get-MsolUser -UserPrincipalName $user.MicrosoftOnlineServicesID
    $msolUser.DisplayName
    $msolUser.licenses
    }

$Sustainers | Sort-Object Name | select Alias, MicrosoftOnlineServicesID
$e1Sustainers | Sort-Object DisplayName | select DisplayName, UserPrincipalName
$allUsers| Sort-Object DisplayName | select DisplayName, UserPrincipalName