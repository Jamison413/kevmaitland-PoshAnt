#https://learn.microsoft.com/en-us/azure/synapse-analytics/security/how-to-set-up-access-control
$workspace = "NetSuiteMI_Prod"
$synapseSgs = @(
     "AzSynapse - $workspace - Admins"
    ,"AzSynapse - $workspace - SqlAdmins"
    ,"AzSynapse - $workspace - Contibutors"
    ,"AzSynapse - $workspace - ComputeOperators"
    ,"AzSynapse - $workspace - CredentialUsers"
    ,"AzSynapse - $workspace - DataReaders"
    )

$groupsToken = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName TeamAndCommunityBot)

$synapseSgs | ForEach-Object {
    new-graphGroup -tokenResponse $groupsToken -groupDisplayName $_ -groupDescription $_ -groupType Security -membershipType Assigned -groupOwners $(whoami -upn)
}

