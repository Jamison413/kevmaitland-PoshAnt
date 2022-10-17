$usersToLicense = convertTo-arrayOfEmailAddresses "verdiana.russo@anthesisgroup.com
Esperanca.Garcia@anthesisgroup.com
ruben.ruiz@anthesisgroup.com
mar.vives@anthesisgroup.com
meritxell.pastor@anthesisgroup.com
marina.clara@anthesisgroup.com
laura.toset@anthesisgroup.com
eleanor.penney@anthesisgroup.com
ida.fiorillo@anthesisgroup.com
toni.mansilla@anthesisgroup.com
estel.homs@anthesisgroup.com
sergi.vila@anthesisgroup.com
daniela.romero@anthesisgroup.com
natalia.reyes@anthesisgroup.com
camilo.alvarez@anthesisgroup.com
ignacio.sojo@anthesisgroup.com
tatiana.garcia@anthesisgroup.com
pilar.martin@anthesisgroup.com
yolanda.fulgueiras@anthesisgroup.com
joel.diez@anthesisgroup.com
marc.anguera@anthesisgroup.com
anna.mas@anthesisgroup.com
francesc.romero@anthesisgroup.com
merce.sorts@anthesisgroup.com
ricard.saborit@anthesisgroup.com
eulalia.miralles@anthesisgroup.com
xavier.codina@anthesisgroup.com
alba.bonas@anthesisgroup.com
rosa.rovira@anthesisgroup.com
nuria.asensio@anthesisgroup.com
miquel.rubio@anthesisgroup.com
montse.goma@anthesisgroup.com
roger.cardellach@anthesisgroup.com
irantzu.aured@anthesisgroup.com
alberta.gil@anthesisgroup.com
cristina.puig@anthesisgroup.com
marta.porra@anthesisgroup.com
Xavi.Benito@anthesisgroup.com
nerea.coca@anthesisgroup.com
barbara.montes@anthesisgroup.com
toni.soler@anthesisgroup.com
cristina.bayes@anthesisgroup.com
ester.padros@anthesisgroup.com
sandra.bordallo@anthesisgroup.com
Ester.Castillejo@anthesisgroup.com
pere.pous@anthesisgroup.com
nuria.sole@anthesisgroup.com
nuria.castells@anthesisgroup.com
laia.puig@anthesisgroup.com
nerea.lopez@anthesisgroup.com
Joana.Soares@anthesisgroup.com"

$teamsToken = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName TeamsBot)
$migrationGroup = get-graphGroups -tokenResponse $teamsToken -filterId 5bef0536-2c0e-4976-a2a5-089dceca978d


$usersToLicense | ForEach-Object {
    $thisUser = $_
    $thisGraphUser = get-graphUsers -tokenResponse $teamsToken -filterUpns $thisUser
    add-graphUsersToGroup -tokenResponse $teamsToken -graphGroupId $migrationGroup.id -memberType members -graphUserIds $thisGraphUser.id
    add-graphLicenseToUser -tokenResponse $teamsToken -userIdOrUpn $thisUser -licenseFriendlyName EMS_E3
    }

$usersInGroup = get-graphUsersFromGroup -tokenResponse $teamsToken -groupId $migrationGroup.id -memberType Members -selectAllProperties
$usersMissingFromGroup = Compare-Object -ReferenceObject $usersToLicense -DifferenceObject $usersInGroup.userPrincipalName
$usersMissingFromGroup

$emsLicense = get-microsoftProductInfo -getType GUID -fromType FriendlyName -fromValue EMS_E3
$unlicenedUsers = $usersInGroup | Where-Object {$_.assignedLicenses.skuId -notcontains $emsLicense}
$licenedUsers = $usersInGroup | Where-Object {$_.assignedLicenses.skuId -contains $emsLicense}

$unlicenedUsers.Count
$licenedUsers.Count

