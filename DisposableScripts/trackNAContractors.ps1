$tokenTeams = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName TeamsBot) -grant_type client_credentials

$naContractors = convertTo-arrayOfEmailAddresses "Ben.Dukes@anthesisgroup.com
Chris.Hazen@anthesisgroup.com
Deby.Stabler@anthesisgroup.com
Leslie.Macdougall@anthesisgroup.com
Matt.Dion@anthesisgroup.com
Susan.Mazzarella@anthesisgroup.com
Tanmoy.Mondal@anthesisgroup.com
Mariko.Thorbecke@anthesisgroup.com
DeAnn.Sarver@anthesisgroup.com
Harmony.Eberhardt@anthesisgroup.com
john.heckman@anthesisgroup.com
John.Hennessey@anthesisgroup.com
Matt.Hannafin@anthesisgroup.com
Meri.Mullins@anthesisgroup.com
Nico.van.Exel@anthesisgroup.com
Sophia.Traweek@anthesisgroup.com"

#Only person without @anthesisgroup.com account is jstittmann@gmail.com


$naContractorAccounts = @($null) * $naContractors.Count
for ($i=0;$i -lt $naContractors.Count;$i++){
    $naContractorAccounts[$i] = get-graphUsers -tokenResponse $tokenTeams -filterUpns $naContractors[$i] -selectAllProperties -useBetaEndPoint -selectCustomProperties "signInActivity"
    $naContractorAccounts[$i].licenseAssignmentStates | % {
        $_ | Add-Member -MemberType NoteProperty -Name FriendlyName -Value $(get-microsoftProductInfo -getType FriendlyName -fromType GUID -fromValue $_.skuId) -Force
    }
    $thisContractorGroups = invoke-graphGet -tokenResponse $tokenTeams -graphQuery "/users/$($naContractorAccounts[$i].id)/transitiveMemberOf"
    $naContractorAccounts[$i] | Add-Member -MemberType NoteProperty -Name Groups -Value $thisContractorGroups -Force
}

$naContractorAccounts | select displayName, userPrincipalName, @{N="LastSignIn"; E={$_.signInActivity.lastSignInDateTime}}, accountEnabled, @{N="LineManager"; E={$_.manager.userPrincipalName}}, @{N="Licenses"; E={@($_.licenseAssignmentStates.FriendlyName)}}, @{N="Groups"; E={$_.Groups.displayName}}, @{N="BusinessUnit"; E={$_.anthesisgroup_employeeInfo.businessUnit}}, @{N="contractTypeAccordingToIT"; E={$_.anthesisgroup_employeeInfo.contractType}} | Export-Csv -Path $env:USERPROFILE\Downloads\NAContractors_$(get-date -f FileDateTimeUniversal).csv -NoTypeInformation -Force




get-graphUsers -tokenResponse $tokenTeams -filterUpns $naContractors[0] -selectAllProperties -useBetaEndPoint -selectCustomProperties "signInActivity"