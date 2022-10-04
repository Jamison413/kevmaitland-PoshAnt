#user upn
param(
    [CmdletBinding()]
    [parameter(Mandatory = $true)]
    [ValidateNotNull()]
    [ValidatePattern(".[@].")]
    [string]$upnString
    )

#connect
$userBotDetails = get-graphAppClientCredentials -appName UserBot
$tokenResponse = get-graphTokenResponse -aadAppCreds $userBotDetails
connect-ToExo


#get Graph user
$user = get-graphUsers -tokenResponse $tokenResponse -filterUpns $upnString -Verbose

#MDM - COBO Users
If(($user | Measure-Object).Count -eq 1){
 Add-DistributionGroupMember -Identity "MDM - COBO Users" -Member $user.userPrincipalName -Confirm:$false  
}
Else{
write-host "Too many users found"
}

#MDM - BYOD Users
If(($user | Measure-Object).Count -eq 1){
 Add-DistributionGroupMember -Identity "MDM - BYOD Users" -Member $user.userPrincipalName -Confirm:$false  
}
Else{
write-host "Too many users found"
}


