Import-Module _PS_Library_MSOL.psm1
Import-Module _PS_Library_GeneralFunctionality
connect-ToExo

function set-primaryEmailAddress($user,$address){
    try{
        $mb = Get-Mailbox $user
        }
    catch{$_}
    if($mb){
        if ($mb.EmailAddresses -icontains "smtp:$address"){
            $mb.EmailAddresses | %{
                if($_ -match $address){[array]$newAddressList += "SMTP:$address"}
                else{[array]$newAddressList += $_.ToLower()}
                }
            }
        else{$problem = "E-mail address $address not associated with User $user"}
        }
    else{$problem = "User Mailbox $user not found"}
    if(!$problem -and $newAddressList.Count -eq $mb.EmailAddresses.Count){
        try{
            Set-Mailbox -Identity $user -EmailAddresses $newAddressList
            Write-Host "$address set as primary on $user"
            $true
            }
        catch{$_}
        }
    elseif(!$problem -and $newAddressList.Count -ne $mb.EmailAddresses.Count){
        $problem = "There's somethign weird about the e-mail addresses assigned to $user and I can't automate this reliably. Do it manually instead."
        $false
        }
    else{$false}
    Write-Host $problem
    }

set-primaryEmailAddress -user $newSm.Identity -address "careers@sustain.co.uk"