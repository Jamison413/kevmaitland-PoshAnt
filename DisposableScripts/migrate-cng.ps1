$tokenTeams = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName TeamsBot) -grant_type client_credentials
Connect-ExchangeOnline -UserPrincipalName kev.maitland@climateneutralgroup.com
Connect-ExchangeOnline -UserPrincipalName t0-kevin.maitland@anthesisgroup.com
$365creds = set-MsolCredentials


$johnLeppers = get-graphUsers -tokenResponse $tokenTeams -filterUpns "john.leppers@anthesisgroup.com" -selectAllProperties -includeLineManager

$johnLeppers.manager.userPrincipalName

$allUsers = get-graphUsers -tokenResponse $tokenTeams -selectAllProperties 
$allCngUsers = $allUsers | Where-Object {$_.manager.userPrincipalName -eq "John.Leppers@anthesisgroup.com" -or $_.userPrincipalName -in @("John.Leppers@anthesisgroup.com","Rene.Toet@anthesisgroup.com","Marieke.Godding@anthesisgroup.com","Franz.Rentel@anthesisgroup.com","Arjen.Struijk@anthesisgroup.com")}
$allCngUsers = get-graphUsers -tokenResponse $tokenTeams -selectAllProperties 

$nldGroup = get-graphGroups -tokenResponse $tokenTeams -filterDisplayName "All Utrecht (NLD)" -filterGroupType Unified
$gbrGroup = get-graphGroups -tokenResponse $tokenTeams -filterDisplayName "All Homeworkers (GBR)" -filterGroupType Unified

$allCngUsers.count
$allCngUsers | foreach-object {
    $thisUser = $_
    write-host "Licensing [$($thisUser.userPrincipalName)] with E3"
    $dummy = add-graphLicenseToUser -tokenResponse $tokenTeams -userIdOrUpn $thisUser.id -licenseFriendlyName Office_E3 -disabledPlansGuids "7547a3fe-08ee-4ccb-b430-5077c5041653" 
    write-host "Licensing [$($thisUser.userPrincipalName)] with MDE"
    $dummy = add-graphLicenseToUser -tokenResponse $tokenTeams -userIdOrUpn $thisUser.id -licenseFriendlyName MDE
    write-host "Licensing [$($thisUser.userPrincipalName)] with EMS"
    $dummy = add-graphLicenseToUser -tokenResponse $tokenTeams -userIdOrUpn $thisUser.id -licenseFriendlyName EMS_E3
    write-host "Licensing [$($thisUser.userPrincipalName)] with Audio"
    $dummy = add-graphLicenseToUser -tokenResponse $tokenTeams -userIdOrUpn $thisUser.id -licenseFriendlyName TeamsAudioConferencingSelect
    set-graphUser -tokenResponse $tokenTeams -userIdOrUpn $thisUser.id -userPropertyHash @{officeLocation="Utrecht, NLD"; usageLocation="NL";city="Utrecht, NLD"} -userEmployeeInfoExtensionHash @{businessUnit="Climate Neutral Group (NLD)"}
    add-graphUsersToGroup -tokenResponse $tokenTeams -graphGroupId $nldGroup.id -memberType members -graphUserIds $thisUser.id
}

$allCngUsers | foreach-object {
    $thisUser = $_
    write-host "Processing [$($thisUser.userPrincipalName)] group membership"
    remove-graphUsersFromGroup -tokenResponse $tokenTeams -graphGroupId $gbrGroup.id -memberType Members -graphUserIds $thisUser.id
    add-graphUsersToGroup -tokenResponse $tokenTeams -graphGroupId $nldGroup.id -memberType members -graphUserIds $thisUser.id
}


$nldMembers = get-graphUsersFromGroup -tokenResponse $tokenTeams -groupId $nldGroup.id -memberType TransitiveMembers -selectAllProperties | Where-Object {$_.displayName -notmatch "kev maitland"}
$nldMembers.userPrincipalName

$allCngUsers = $nldMembers | Where-Object 

$activePasswords = $nldGroup | Where-Object {$_.passwordProfile.forceChangePasswordNextSignIn -eq $false}
$unChangedPasswords = $nldMembers | Where-Object {$_.lastPasswordChangeDateTime -le $_.createdDateTime}


#region Sites & Teams
$arrayOfTeamNames = convertTo-arrayOfStrings "Administration Team (NLD)
Administration Team (ZAF)
Human Resources (HR) Team (NLD)
Human Resources (HR) Team (ZAF)
Marketing Team (NLD)
Marketing Team (ZAF)
Climate Neutral Group (CNG) Interim Team
Finance Team (NLD)
Sales (Carbon Advisory and Partnerships) Team (NLD)
Sales (Internal) Team (NLD)
Consulting Team (NLD)
Certification Team (NLD)
Project Portfolio Team (NLD)
Project Development Team (NLD)
Management Team (NLD)
Agri-Carbon Team (ZAF)
New Project Team (ZAF)
Board Team (Climate Neutral Group)
Management Team (ZAF)
All (BEL)
All (NLD)
All (ZAF)
All Brussels (BEL)
All Homeworkers (BEL)
All Homeworkers (NLD)
All Utrecht (NLD)
All Cape Town (ZAF)
All Johannesburg (ZAF)
All Homeworkers (ZAF)
External - CB Ecocert
External - CB Preferred by Nature
"
$cngTeams = @($null) * $arrayOfTeamNames.Count
for ($i=0; $i -lt $arrayOfTeamNames.Count; $i++){
    Write-Output "Getting Group [$($arrayOfTeamNames[$i])]"
    $cngTeams[$i] = get-graphGroups -tokenResponse $tokenTeams -filterDisplayName $arrayOfTeamNames[$i] -filterGroupType Unified
    Write-Output "`tGetting Site"
    $thisSite = get-graphSite -tokenResponse $tokenTeams -groupId $cngTeams[$i].id
    $cngTeams[$i] | Add-Member -MemberType NoteProperty -Name SiteUrl -Value $thisSite.webUrl -Force
    $cngTeams[$i] | Add-Member -MemberType NoteProperty -Name SiteId -Value $thisSite.id -Force
}
$marketingResourcesSite = get-graphSite -tokenResponse $tokenTeams -serverRelativeUrl "sites/Resources-Marketing"
$cngTeams += [PSCustomObject]@{
    SiteUrl = $marketingResourcesSite.webUrl
    SiteId = $marketingResourcesSite.id
}
#endregion


#region Page Templates #Still relies on legacy PnP code :'(
for ($i=0; $i -lt $cngTeams.Count; $i++){
    copy-spoPage -sourceUrl "https://anthesisllc.sharepoint.com/sites/Resources-IT/SitePages/Candidate-Page-for-Regional-Sites.aspx" -destinationSite $cngTeams[$i].SiteUrl -pnpCreds $365creds -overwriteDestinationFile $true -renameFileAs "TemplatePage.aspx" -Verbose 
}
#endregion


#Region DocLibPrep
$docLibPrep = import-csv -Delimiter "," -Path $env:USERPROFILE\Downloads\DataMigrationDestinationDocLibPrep.csv
$docLibPrep | ForEach-Object {
    $thisDocLibPrep = $_
    Write-Output "Prepping [$($thisDocLibPrep.DestinationUrl)]"
    Write-Output "`tRetrieving DocLibs"
    $thisSiteDocLibs = get-graphDrives -tokenResponse $tokenTeams -siteUrl $thisDocLibPrep.'Destination SiteUrl' #Get the Drives/DocLibs from the Site if it has changed
    if([string]::IsNullOrWhiteSpace($thisDocLibPrep.PreDocLib)){
        $realDocLib = [uri]::EscapeDataString($thisDocLibPrep.PrettyDestinationDocLib) #If there is no PreDocLib value, use PrettyDestinationDocLib
        Write-Output "`tNo PreDocLib found, using PrettyDestinationDocLib [$($thisDocLibPrep.PrettyDestinationDocLib)]"                
    }
    else{
        $realDocLib = [uri]::EscapeDataString($thisDocLibPrep.PreDocLib)#Else use PreDocLib
        Write-Output "`tPreDocLib [$($thisDocLibPrep.PreDocLib)] found - using this"                
    } 
    $realDocLib = $realDocLib.TrimEnd("/").TrimEnd("%2F")
    if($realDocLib -notin @(split-path $thisSiteDocLibs.webUrl -Leaf)){ #If the DocLib hasn't been created yet...
        $thisTeam = $cngTeams | Where-Object {$_.SiteUrl -eq $thisDocLibPrep.'Destination SiteUrl'} #Find the apprporiate $cngTeam (we need the SiteId)
        Write-Output "`tMatched to Team [$($thisTeam.displayName)]"
        $newDocLib = new-graphList -tokenResponse $tokenTeams -siteGraphId $thisTeam.SiteId -listType documentLibrary -listDisplayName $([uri]::UnescapeDataString($realDocLib)) #Unescape this at the last moment as Graph we escape it again with weird results
        Write-Output "`t`tDocLib [$($newDocLib.displayName)][$($newDocLib.id)] created"
    }
    else{
        Write-Output "`t`tDocLib [$($realDocLib)] already found in destination"
    }
}
#endregion

#region FolderPrep
$docLibPrep = import-csv -Delimiter "," -Path $env:USERPROFILE\Downloads\DataMigrationDestinationDocLibPrep.csv
$docLibPrep | ForEach-Object {
    $thisDocLibPrep = $_
    if([string]::IsNullOrWhiteSpace($thisDocLibPrep.PreDocLib)){
        $realDocLib = [uri]::EscapeDataString($thisDocLibPrep.PrettyDestinationDocLib.TrimEnd("/")) #If there is no PreDocLib value, use PrettyDestinationDocLib
        Write-Output "`tNo PreDocLib found, using PrettyDestinationDocLib [$($thisDocLibPrep.PrettyDestinationDocLib)]"                
    }
    else{
        $realDocLib = [uri]::EscapeDataString($thisDocLibPrep.PreDocLib.TrimEnd("/"))#Else use PreDocLib
        Write-Output "`tPreDocLib [$($thisDocLibPrep.PreDocLib)] found - using this"                
    } 
    $subfolders = [uri]::UnescapeDataString($thisDocLibPrep.DestinationUrl.Replace($thisDocLibPrep.'Destination SiteUrl'+"/"+$realDocLib,"").TrimEnd("/"))
    if([string]::IsNullOrWhiteSpace($subfolders)){Write-Output "No subfolders to create!"; return} #No subfolders to create!

    $thisDestinationTeam = $cngTeams | Where-Object {$_.SiteUrl -eq $thisDocLibPrep.'Destination SiteUrl'}
    if($thisDestinationTeam){
        $thisDocLib = get-graphDrives -tokenResponse $tokenTeams -siteGraphId $thisDestinationTeam.SiteId | Where-Object {$_.webUrl -match "$realDocLib$"}
        if ([string]::IsNullOrWhiteSpace($thisDocLib)) {{Write-Error "No DocLib matched to name [$($realDocLib)] in [$($thisDestinationTeam.displayName)]";return}}
        elseif(-not $($thisDocLib.Count -gt 1)){
            Write-Output "Creating folder(s) [$($subfolders)] in [$($thisDocLib.webUrl)]"
            [array]$newFolders = add-graphArrayOfFoldersToDrive -graphDriveId $thisDocLib.id -foldersAndSubfoldersArray $subfolders -tokenResponse $tokenTeams -conflictResolution Fail
            Write-Output "`tFolder(s) [$($newFolders.webUrl -join "`r`n`t`t")] created!"
        }
        else {Write-Error "Too manys DocLibs matched to name [$($realDocLib)] in [$($thisDestinationTeam.displayName)]";return}
    
    }
    else{Write-Error "No Team matched to destination [$($thisLine.DestinationUrl)]";return}
}

#endRegion


#region User Prep
$arrayOfUsersToMigrate = convertTo-arrayOfEmailAddresses "irem.gurdal@climateneutralgroup.com
Olav.Provily@climateneutralgroup.com
Derk.deHaan@climateneutralgroup.com
Jos.Cozijnsen@climateneutralgroup.com
rogier.vanveenendaal@climateneutralgroup.com
Janne.Kuhn@climateneutralgroup.com
Gray.Maguire@climateneutralgroup.com
Franz.Rentel@climateneutralgroup.com
Derek.Groot@climateneutralgroup.com
Silvana.Claassen@climateneutralgroup.com
Melissa.Baird@climateneutralgroup.com
Rene.Toet@climateneutralgroup.com
Giacomo.diLallo@climateneutralgroup.com
Marjan.Verbeek@climateneutralgroup.com
Alicia.Kok@climateneutralgroup.com
Arjen.Struijk@climateneutralgroup.com
Mandy.Ngada@climateneutralgroup.com
Nonkululeko.Hadebe@climateneutralgroup.com
Siviwe.Malongweni@climateneutralgroup.com
Jana.Hofmann@climateneutralgroup.com
Nathan.Jansen@climateneutralgroup.com
Marieke.Andringa@climateneutralgroup.com
Emma.Cuijpers@climateneutralgroup.com
Omar.Hamouda@climateneutralgroup.com
Willem.Melis@climateneutralgroup.com
Paul.Zuiderbeek@climateneutralgroup.com
Jack.Everling@climateneutralgroup.com
Rianne.Kluitmans@climateneutralgroup.com
Elenoor.vanEs@climateneutralgroup.com
Esther.Snijder@climateneutralgroup.com
Michiel.Tijmensen@climateneutralgroup.com
Andrew.Lancefield@climateneutralgroup.com
Russell.Holmes@climateneutralgroup.com
Geert.Eenhoorn@climateneutralgroup.com
jorie.vanrooijen@climateneutralgroup.com
katerina.kadlecova@climateneutralgroup.com
Anton.Kool@climateneutralgroup.com
Grant.Little@climateneutralgroup.com
Jouke.Roelfzema@climateneutralgroup.com
Marieke.Godding@climateneutralgroup.com
Ciska.Uijlenbroek@climateneutralgroup.com
Rosa.Esnard@climateneutralgroup.com
Lorna.Tasker@climateneutralgroup.com
daan.vandekamp@climateneutralgroup.com
Eline.Weerts@climateneutralgroup.com
Maurits.Wesseling@climateneutralgroup.com
Mark.HuisintVeld@climateneutralgroup.com
Sandra.Slotboom@climateneutralgroup.com
Liezl.Julius@climateneutralgroup.com
AnneWil.Broersma@climateneutralgroup.com
Gabrielle.Smith@climateneutralgroup.com
Marloes.vanLuijk@climateneutralgroup.com
Bas.Ooteman@climateneutralgroup.com
Sanne.Dallinga@climateneutralgroup.com
Mandy.Momberg@climateneutralgroup.com
Egbert.Koetsier@climateneutralgroup.com
john.leppers@climateneutralgroup.com
Wopke.Geurts@climateneutralgroup.com
Ellen.brouwer@climateneutralgroup.com
"
$arrayOfUsersToMigrate = convertTo-arrayOfEmailAddresses "Alicia.Kok@climateneutralgroup.onmicrosoft.com
Andrew.Lancefield@climateneutralgroup.onmicrosoft.com
AnneWil.Broersma@climateneutralgroup.onmicrosoft.com
anton.kool@climateneutralgroup.onmicrosoft.com
Arjen.Struijk@climateneutralgroup.onmicrosoft.com
bas.ooteman@climateneutralgroup.onmicrosoft.com
Ciska.Uijlenbroek@climateneutralgroup.onmicrosoft.com
daan.vandekamp@climateneutralgroup.onmicrosoft.com
derek.groot@climateneutralgroup.onmicrosoft.com
derk.dehaan@climateneutralgroup.onmicrosoft.com
Egbert.Koetsier@climateneutralgroup.onmicrosoft.com
elenoor.vanes@climateneutralgroup.onmicrosoft.com
eline.weerts@climateneutralgroup.onmicrosoft.com
Ellen.brouwer@climateneutralgroup.onmicrosoft.com
emma.cuijpers@climateneutralgroup.onmicrosoft.com
esther.snijder@climateneutralgroup.onmicrosoft.com
Franz.Rentel@climateneutralgroup.onmicrosoft.com
gabrielle.smith@climateneutralgroup.onmicrosoft.com
geert.eenhoorn@climateneutralgroup.onmicrosoft.com
Giacomo.diLallo@climateneutralgroup.onmicrosoft.com
grant.little@climateneutralgroup.onmicrosoft.com
gray.maguire@climateneutralgroup.onmicrosoft.com
irem.gurdal@climateneutralgroup.onmicrosoft.com
jack.everling@climateneutralgroup.onmicrosoft.com
janne.kuhn@climateneutralgroup.onmicrosoft.com
john.leppers@climateneutralgroup.onmicrosoft.com
jorie.vanrooijen@climateneutralgroup.onmicrosoft.com
jos.cozijnsen@climateneutralgroup.onmicrosoft.com
jouke.roelfzema@climateneutralgroup.onmicrosoft.com
katerina.kadlecova@climateneutralgroup.onmicrosoft.com
liezl.julius@climateneutralgroup.onmicrosoft.com
lorna.tasker@climateneutralgroup.onmicrosoft.com
mandy.momberg@climateneutralgroup.onmicrosoft.com
mandy.ngada@climateneutralgroup.onmicrosoft.com
marieke.andringa@climateneutralgroup.onmicrosoft.com
Marieke.Godding@climateneutralgroup.onmicrosoft.com
marjan.verbeek@climateneutralgroup.onmicrosoft.com
Mark.HuisintVeld@climateneutralgroup.onmicrosoft.com
Marloes.vanLuijk@climateneutralgroup.onmicrosoft.com
maurits.wesseling@climateneutralgroup.onmicrosoft.com
melissa.baird@climateneutralgroup.onmicrosoft.com
michiel.tijmensen@climateneutralgroup.onmicrosoft.com
nathan.jansen@climateneutralgroup.onmicrosoft.com
Nonkululeko.Hadebe@climateneutralgroup.onmicrosoft.com
Olav.Provily@climateneutralgroup.onmicrosoft.com
omar.hamouda@climateneutralgroup.onmicrosoft.com
paul.zuiderbeek@climateneutralgroup.onmicrosoft.com
Rene.Toet@climateneutralgroup.onmicrosoft.com
rianne.kluitmans@climateneutralgroup.onmicrosoft.com
rogier.vanveenendaal@climateneutralgroup.onmicrosoft.com
rosa.esnard@climateneutralgroup.onmicrosoft.com
Russell.Holmes@climateneutralgroup.onmicrosoft.com
Sandra.Slotboom@climateneutralgroup.onmicrosoft.com
Sanne.Dallinga@climateneutralgroup.onmicrosoft.com
Silvana.Claassen@climateneutralgroup.onmicrosoft.com
siviwe.malongweni@climateneutralgroup.onmicrosoft.com
Willem.Melis@climateneutralgroup.onmicrosoft.com
Wopke.Geurts@climateneutralgroup.onmicrosoft.com
"
#endregion

#region Shared Mailbox Prep
$arrayOfCngSharedMailboxAddresses = convertTo-arrayOfEmailAddresses "accounts@climateneutralgroup.co.za
administratie@climateneutralgroup.com
administratieechtgoed@climateneutralgroup.com
AVG@climateneutralgroup.com
certificatengreenseat@climateneutralgroup.com
certification@climateneutralgroup.com
Communicatie@climateneutralgroup.com
communication@climateneutralgroup.com
consultancy@climateneutralgroup.com
website@climateneutralgroup.co.za
csc@climateneutralgroup.com
fcc@climateneutralgroup.com
ferry@echtgoed.nl
footprintdata@climateneutralgroup.com
greendreams@climateneutralgroup.com
Infodocdata@climateneutralgroup.com
Info@climateneutralgroup.com
Info@climateneutralgroup.co.za
info@greenseat.com
inspiratie@climateneutralgroup.com
internship@climateneutralgroup.com
Marketing@climateneutralgroup.co.za
nacalculatie.voetafdruk@climateneutralgroup.com
Nieuwsbrief@climateneutralgroup.com
officemanagement@climateneutralgroup.com
pers@climateneutralgroup.com
planning.simapro@climateneutralgroup.com
Recruitment@climateneutralgroup.com
Recruitment@climateneutralgroup.co.za
Sollicitatie@climateneutralgroup.com
"
$smbxToMigrate = $allMailboxes | Where-Object {$_.PrimarySmtpAddress -in $arrayOfCngSharedMailboxAddresses} #See later for $allMailboxes

Connect-ExchangeOnline -UserPrincipalName kev.maitland@climateneutralgroup.onmicrosoft.com

$allMailboxes = Get-EXOMailbox -ResultSize Unlimited
$allMailboxes | ForEach-Object {
    $thisMailbox = $_
    Write-Output "[$($thisMailbox.PrimarySmtpAddress)]"
    $thisMailboxStats = Get-EXOMailboxStatistics -Identity  $thisMailbox.Identity
    $thisMailbox | Add-Member -MemberType NoteProperty -Name TotalItemSize -Value $thisMailboxStats.TotalItemSize -Force
    $thisMailboxPermissions = Get-EXOMailboxPermission -Identity $thisMailbox.ExternalDirectoryObjectId
    $thisMailboxDelegates = $thisMailboxPermissions | Where-Object {$_.User -notmatch "NT AUTHORITY" -and $_.User -ne "kev.maitland@climateneutralgroup.com" -and $_.AccessRights -contains "FullAccess"} #Strip out junk
    $thisMailbox | Add-Member -MemberType NoteProperty -Name Delegates -Value $thisMailboxDelegates.User -Force
    $thisMailboxSendAs = $thisMailboxPermissions | Where-Object {$_.User -notmatch "NT AUTHORITY" -and $_.User -ne "kev.maitland@climateneutralgroup.com" -and $_.AccessRights -contains "SendAs"} #Strip out junk
    $thisMailbox | Add-Member -MemberType NoteProperty -Name SendAs -Value $thisMailboxSendAs.User -Force
    $thisMailboxOoO = Get-MailboxAutoReplyConfiguration -Identity  $thisMailbox.Identity
    $thisMailbox | Add-Member -MemberType NoteProperty -Name OoO -Value $thisMailboxOoO -Force
    if($thisMailbox.PrimarySmtpAddress -in $arrayOfCngSharedMailboxAddresses){
        if ($thisMailbox.PrimarySmtpAddress -match "greenseat") {$anthesisAddress = "cng-greenseat-$($thisMailbox.PrimarySmtpAddress.Split("@")[0])@anthesisgroup.com"}
        elseif ($thisMailbox.PrimarySmtpAddress -match "co.za") {$anthesisAddress = "cng-za-$($thisMailbox.PrimarySmtpAddress.Split("@")[0])@anthesisgroup.com"}
        else {$anthesisAddress = "cng-$($thisMailbox.PrimarySmtpAddress.Split("@")[0])@anthesisgroup.com"}
        $anthesisDisplayName = "$($thisMailbox.DisplayName) - CNG Shared Mailbox"
    }
    else{
        $anthesisAddress = "$($thisMailbox.UserPrincipalName.Split("@")[0])@anthesisgroup.com"
        $anthesisDisplayName = $thisMailbox.DisplayName
    }
    $thisMailbox | Add-Member -MemberType NoteProperty -Name anthesisAddress -Value $anthesisAddress -Force
    $thisMailbox | Add-Member -MemberType NoteProperty -Name anthesisDisplayName -Value $anthesisDisplayName -Force
    $thisMailbox | Add-Member -MemberType NoteProperty -Name nonRoutableUpn -Value "$($thisMailbox.UserPrincipalName.Split("@")[0])@climateneutralgroup.onmicrosoft.com" -Force
}

$usersToMigrate = $allMailboxes | Where-Object {$_.PrimarySmtpAddress -in $arrayOfUsersToMigrate} #See later for $allMailboxes
$smbxToMigrate | ForEach-Object {
    if ($thisMailbox.PrimarySmtpAddress -match "greenseat") {$anthesisAddress = "cng-greenseat-$($thisMailbox.PrimarySmtpAddress.Split("@")[0])@anthesisgroup.com"}
    elseif ($thisMailbox.PrimarySmtpAddress -match "co.za") {$anthesisAddress = "cng-za-$($thisMailbox.PrimarySmtpAddress.Split("@")[0])@anthesisgroup.com"}
    else {$anthesisAddress = "cng-$($thisMailbox.PrimarySmtpAddress.Split("@")[0])@anthesisgroup.com"}
    $thisMailbox | Add-Member -MemberType NoteProperty -Name anthesisAddress -Value $anthesisAddress -Force
}

Connect-ExchangeOnline -UserPrincipalName kev.maitland@climateneutralgroup.com



#Make SMBX based on CNG SMBX ($smbxToMigrate)
Connect-ExchangeOnline -UserPrincipalName t0-kevin.maitland@anthesisgroup.com
$owner = "t0-kevin.maitland@anthesisgroup.com" 
$smbxToMigrate | ForEach-Object {
    $thisMailbox = $_
    Write-Output "[$($thisMailbox.PrimarySmtpAddress)]"
    $exchangeAlias = $(guess-aliasFromDisplayName -displayName $thisMailbox.anthesisDisplayName)
    New-Mailbox -Shared -ModeratedBy $owner -DisplayName $thisMailbox.anthesisDisplayName -Name $thisMailbox.anthesisDisplayName -Alias $exchangeAlias -PrimarySmtpAddress $thisMailbox.anthesisAddress | Set-Mailbox -HiddenFromAddressListsEnabled $false -RequireSenderAuthenticationEnabled $false -MessageCopyForSendOnBehalfEnabled $true -MessageCopyForSentAsEnabled $true #-EmailAddresses $allEmailAddresses
    Set-User -Identity $exchangeAlias -AuthenticationPolicy "Block Basic Auth"
}

$smbxToMigrate | Select-Object PrimarySmtpAddress, anthesisAddress | Export-csv  -path   "$logFolder\$(get-date -f FileDateTimeUniversal)_SmbxToMigrate.csv" -NoTypeInformation

#Set SMBX Delegates & access
$smbxToMigrate | ForEach-Object {
    $thisMailbox = $_
    $thisMailbox.Delegates | ? {$_ -notmatch "kev.maitland"} | ForEach-Object {
        Write-Output "Adding [$($_)] as delegate for [$($thisMailbox.anthesisAddress)]"
        Add-MailboxPermission -AccessRights "FullAccess" -User $_.Replace("climateneutralgroup.onmicrosoft.com","anthesisgroup.com") -AutoMapping $true -Identity $thisMailbox.anthesisAddress
    }
    $thisMailbox.SendAs | select-object | ForEach-Object {
        Add-RecipientPermission -Identity $thisMailbox.anthesisAddress -Trustee $_.Replace("climateneutralgroup.onmicrosoft.com","anthesisgroup.com") -AccessRights SendAs -Confirm:$false
    }
}

#endregion




#region migrationday
$logFolder = "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\CNG-MigrationDayLogs"
#region Dump Users, SMBX, Groups, Resources
$allMailboxes | Convert-OutputForCSV | Export-Csv -Path "$logFolder\All_MBX_$(get-date -f FileDateTimeUniversal).csv" -NoTypeInformation

$dg = Get-DistributionGroup -ResultSize Unlimited
$dg | ForEach-Object {
    $thisDg = $_
    $theseMembers = Get-DistributionGroupMember -Identity $thisDg.Id
    $thisDg | Add-Member -MemberType NoteProperty -Name WindowsLiveID -Value $theseMembers.WindowsLiveID -Force
}
$dg | Export-Csv -Path "$logFolder\DistributionGroups_$(get-date -f FileDateTimeUniversal).csv" -NoTypeInformation

#Get 365 Teams
#endregion


#region remove routable domains
$cngCreds = Get-Credential -UserName "kev.maitland@climateneutralgroup.com" -Message "Creds"
$cngCreds = Get-Credential -UserName "kev.maitland@climateneutralgroup.onmicrosoft.com" -Message "Creds"
Connect-MsolService -Credential $cngCreds
connect-toAAD -credential $cngCreds
Connect-ExchangeOnline -credential $cngCreds

$domainsToMigrate = convertTo-arrayOfStrings "climateneutralgroup.com
climateneutralgroup.co.za
greenseat.nl
klimaatneutraal.nl
greenseat.com
echtgoed.nl
"
$domainsToMigrate | ForEach-Object{
    Get-MsolUser -DomainName $_ | Convert-OutputForCSV | Export-Csv -Path "$logFolder\Routable_$_ $(get-date -f FileDateTimeUniversal).csv"
}
Get-MsolUser -DomainName climateneutralgroup.com | Convert-OutputForCSV | Export-Csv -Path "$logFolder\Routable_climateneutralgroup.com_$(get-date -f FileDateTimeUniversal).csv"


#Switch UPN to climateNeutralGroup.onmicrosoft.com & remove routable domains
$usersToMigrate | ForEach-Object {
    $thisMailbox = $_
    Write-Output "Processing [$($thisMailbox.UserPrincipalName)]"
    Write-Output "`tSetting UPN to [$($thisMailbox.nonRoutableUpn)]"
    Set-AzureADUser -ObjectId $thisMailbox.ExternalDirectoryObjectId -UserPrincipalName $thisMailbox.nonRoutableUpn
}
$usersToMigrate | ForEach-Object {
    $thisMailbox = $_
    $thisMailbox.EmailAddresses | ForEach-Object {
        $thisEmail = $_
        if($thisEmail -imatch "smtp:" -and ($thisEmail -match "@climateneutralgroup.com")){
            Write-Output "`t`tRemoving routable email address [$($thisEmail)]"
            Set-Mailbox $thisMailbox.Identity -EmailAddresses @{remove=$thisEmail}        
        }
    }
}
# -or $thisEmail -match "@climateneutralgroup.co.za
#I *think* Exchange will automatically switch the SIP: address to match the UPN without further intervention (I don;t think you even *need* to remove the routable addresses for this to happen) 

$allMailboxes | ForEach-Object {
    $thisMailbox = $_
    Write-Output "Processing [$($thisMailbox.UserPrincipalName)]"
    Write-Output "`tSetting UPN to [$($thisMailbox.nonRoutableUpn)]"
    Set-AzureADUser -ObjectId $thisMailbox.ExternalDirectoryObjectId -UserPrincipalName $thisMailbox.nonRoutableUpn
}
$usersToMigrate | ForEach-Object {
    $thisMailbox = $_
    $thisMailbox.EmailAddresses | ForEach-Object {
        $thisEmail = $_
        if($thisEmail -imatch "smtp:" -and ($thisEmail -match "@climateneutralgroup.com")){
            Write-Output "`t`tRemoving routable email address [$($thisEmail)]"
            Set-Mailbox $thisMailbox.Identity -EmailAddresses @{remove=$thisEmail}        
        }
    }
}

$ugs = Get-UnifiedGroup -ResultSize Unlimited


$allMailboxes | ForEach-Object {
    $thisMailbox = $_
    Write-Output "Processing [$($thisMailbox.UserPrincipalName)]"
    Set-Mailbox $thisMailbox.anthesisAddress -EmailAddresses @{add=$thisMailbox.UserPrincipalName}        
}
#endregion





#endregion






$allgraphUsers = get-graphUsers -tokenResponse $tokenTeams -selectAllProperties

$graphUsers = $allgraphUsers | Where-Object {$_.userPrincipalName -in $allMailboxes.anthesisAddress} #get-graphUsers -tokenResponse $tokenTeams -filterUpns $allMailboxes.anthesisAddress -selectAllProperties
$graphUsers  | ForEach-Object {
    $thisGraphUser = $_
    [array]$emailsToAdd = "smpt:$($thisGraphUser.userPrincipalName.Replace("anthesisgroup","climateneutralgroup"))"
    $emailsToAdd += $thisGraphUser.otherMails
    $thisGraphUser = set-graphUser -tokenResponse $tokenTeams -userIdOrUpn $thisGraphUser.id -userPropertyHash @{proxyAddresses=$emailsToAdd}
}



#Region SharePoint ReadOnly
$AdminCenterURL="https://climateneutralgroup-admin.sharepoint.com/"
Connect-SPOService -Url $AdminCenterURL -Credential $cngCreds
$cngSites = get-spoSite 
$cngSites | Convert-OutputForCSV | Export-Csv -Path "$logFolder\All_SPOSites_$(get-date -f FileDateTimeUniversal).csv" -NoTypeInformation
$cngSites | ForEach-Object{
    Write-Output "Processing [$($_.Title)]"
    Set-SPOSite -Identity $_.Url -LockState ReadOnly
}

$cngPersonalSites = get-spoSite -IncludePersonalSite:$true
$cngPersonalSites = $cngPersonalSites | Where-Object {$_.Url -match "/personal/"}
$cngPersonalSites | Convert-OutputForCSV | Export-Csv -Path "$logFolder\All_SPOPersonalSites_$(get-date -f FileDateTimeUniversal).csv" -NoTypeInformation
$cngPersonalSites | ForEach-Object{
    Write-Output "Processing [$($_.Title)]"
    Set-SPOSite -Identity $_.Url -LockState ReadOnly
}


#Read more: https://www.sharepointdiary.com/2019/02/set-sharepoint-online-site-to-read-only-using-powershell.html#ixzz7YXUhQqEt
#endregion



$usersToMigrate[0].OoO
$usersToMigrate | ForEach-Object { 
    $thisMailbox = $_
    Write-Output "Setting OoO for [$($thisMailbox.DisplayName)]"
    Set-MailboxAutoReplyConfiguration -Identity $thisMailbox.anthesisAddress `
        -AutoDeclineFutureRequestsWhenOOF $thisMailbox.OoO.AutoDeclineFutureRequestsWhenOOF `
        -AutoReplyState $thisMailbox.OoO.AutoReplyState `
        -CreateOOFEvent $thisMailbox.OoO.CreateOOFEvent `
        -DeclineAllEventsForScheduledOOF $thisMailbox.OoO.DeclineAllEventsForScheduledOOF `
        -DeclineEventsForScheduledOOF  $thisMailbox.OoO.DeclineEventsForScheduledOOF `
        -DeclineMeetingMessage  $thisMailbox.OoO.DeclineMeetingMessage `
        -EndTime $thisMailbox.OoO.EndTime `
        -EventsToDeleteIDs $thisMailbox.OoO.EventsToDeleteIDs `
        -ExternalAudience $thisMailbox.OoO.ExternalAudience `
        -ExternalMessage $thisMailbox.OoO.ExternalMessage `
        -InternalMessage $thisMailbox.OoO.InternalMessage `
        -StartTime $thisMailbox.OoO.StartTime `
        -OOFEventSubject $thisMailbox.OoO.OOFEventSubject
}


$arrayOfTeams = convertTo-arrayOfStrings "Administration Team (NLD)
Administration Team (ZAF)
Human Resources (HR) Team (NLD)
Human Resources (HR) Team (ZAF)
Marketing Team (NLD)
Marketing Team (ZAF)
Climate Neutral Group (CNG) Interim Team
Finance Team (NLD)
Sales (Carbon Advisory and Partnerships) Team (NLD)
Sales (Internal) Team (NLD)
Consulting Team (NLD)
Certification Team (NLD)
Project Portfolio Team (NLD)
Project Development Team (NLD)
Management Team (NLD)
Agri-Carbon Team (ZAF)
New Project Team (ZAF)
Board Team (Climate Neutral Group)
Management Team (ZAF)"


$thisTeam = get-graphGroups -tokenResponse $tokenTeams -filterDisplayName "Climate Neutral Group (CNG) Interim Team" -filterGroupType Unified
$thesePeople = $(convertTo-arrayOfStrings "Stuart McLachlan
Luc Albert
Jason Urry
")
$theseUsers = $usersToMigrate | Where-Object {$_.DisplayName -in $thesePeople}
write-output "ThesePeople.Count = [$($thesePeople.Count)]"
write-output "theseUsers.Count = [$($theseUsers.Count)]"
if ($thesePeople.Count -eq $theseUsers.Count){
    $theseUsers | ForEach-Object {
        #$thisUser = get-graphUsers -tokenResponse $tokenTeams -filterUpns $_.anthesisAddress
        $thisUser = $_
        Write-Output "Adding [$($thisUser.UserPrincipalName)] to [$($thisteam.Displayname)]"
        #Add-DistributionGroupMember -Identity $thisTeam.mail -Member $_.UserPrincipalName -BypassSecurityGroupManagerCheck:$true
        add-graphUsersToGroup -tokenResponse $tokenTeams -graphGroupId $thisteam.id -memberType members -graphUserIds $thisUser.ID
    }
}



#Get migration status
$allCngUsers = get-graphUsers -tokenResponse $tokenTeams -selectAllProperties -useBetaEndPoint -selectCustomProperties "signInActivity" -filterBusinessUnit "Climate Neutral Group (NLD)"
$allCngUsers | % {
    $thisUser = $_
    Write-Output "Getting devices for [$($thisUser.displayName)]"
    $thisUsersDevices = invoke-graphGet -tokenResponse $tokenTeams -graphQuery "/users/$($thisUser.id)/ownedDevices"
    $thisUser | Add-Member -MemberType NoteProperty -Name Devices -Value $thisUsersDevices -Force
}

$notMigrated = $allCngUsers | ? {$($_.Devices.operatingSystem -contains "Windows") -eq $false}
$migrated = $allCngUsers | ? {$($_.Devices.operatingSystem -contains "Windows") -eq $true}
$migrated.displayName | sort
$notMigrated.displayName | sort


$outOfTune = $allCngUsers | ? {$($_.Devices.operatingSystem -contains "IPhone" -or $_.Devices.operatingSystem -contains "AndroidForWork") -eq $false}
$intune = $allCngUsers | ? {$($_.Devices.operatingSystem -contains "IPhone" -or $_.Devices.operatingSystem -contains "AndroidForWork") -eq $true}

$intune.displayName | sort
$outoftune.displayName | sort

$migrated.count
$notMigrated.count
$arjen = $allCngUsers | ? {$_.displayname -match "Arjen"}
$willem = $allCngUsers | ? {$_.displayname -match "Willem"}

$subs = $allCngUsers | ? {$_.displayname -match "Anton Kool" -or $_.displayname -match "Andrew Lancefield"}
$subs | % {
    set-graphUser -tokenResponse $tokenTeams -userIdOrUpn $_.id -userEmployeeInfoExtensionHash @{contractType="Subcontractor"}
}



#Assign CNG address as primary
Connect-ExchangeOnline -UserPrincipalName t0-kevin.maitland@anthesisgroup.com
$allCngUsers | % {
    $thisUser = $_
    $newPrimaryMail = "$($thisUser.mail.Split("@")[0])@climateneutralgroup.com"
    $updatedAddresses = @()

    $thisUser.proxyAddresses | %{
        $address = $_
        $prefix = $address.Split(":")[0] 
        $mail = $address.Split(":")[1] 
  
        if ($mail.ToLower() -eq $newPrimaryMail.ToLower()) {$address = "SMTP:" + $mail} 
        else {$address = $prefix.ToLower() + ":" + $mail} 
       
        $updatedAddresses += $address
    }
    Write-Output "Setting [$($thisUser.DisplayName)] addresses to [`r`n`t$($updatedAddresses -join "`r`n`t")`r`n`t]"
    set-mailbox -Identity $thisUser.userPrincipalName -EmailAddresses $updatedAddresses
}

#Set default calendar free/busy visibility
$allCngUsers | % {
    $thisUser = $_
    Write-Output "Setting [$($thisUser.DisplayName)]"
    Add-MailboxFolderPermission "$($thisUser.userPrincipalName):\Calendar" -User "AllNLD@anthesisgroup.com" -AccessRights "LimitedDetails"
    Add-MailboxFolderPermission "$($thisUser.userPrincipalName):\Calendar" -User "AllZAF@anthesisgroup.com" -AccessRights "LimitedDetails"
    Add-MailboxFolderPermission "$($thisUser.userPrincipalName):\Calendar" -User "AllBEL@anthesisgroup.com" -AccessRights "LimitedDetails"
    #Some mailboxes are in Dutch
    Add-MailboxFolderPermission "$($thisUser.userPrincipalName):\Agenda" -User "AllNLD@anthesisgroup.com" -AccessRights "LimitedDetails"
    Add-MailboxFolderPermission "$($thisUser.userPrincipalName):\Agenda" -User "AllZAF@anthesisgroup.com" -AccessRights "LimitedDetails"
    Add-MailboxFolderPermission "$($thisUser.userPrincipalName):\Agenda" -User "AllBEL@anthesisgroup.com" -AccessRights "LimitedDetails"
}


#Set ApplicatinAccessPolicy for website@climateneutralgroup.co.za mailer
New-ApplicationAccessPolicy -AppId 138da3d8-52e8-49c8-9426-6d36c483383e -PolicyScopeGroupId AAP-climateneutralgroup.co.zamailer@anthesisgroup.com -AccessRight RestrictAccess -Description "Restrict this app to members of mail-enabled security group [AAP - climateneutralgroup.co.za mailer]."


#Get InboxRules for Resources
$rules = Get-InboxRule -Mailbox lease.auto@climateneutralgroup.onmicrosoft.com




#Get ZAF Teams and Members
$allTeams = get-graphGroups -tokenResponse $tokenTeams -filterGroupType Unified -selectAllProperties
$zafTeams = $allTeams | Where-Object {
    $_.displayName -match "(ZAF)"
}
$zafTeams | % {
    $thisTeam = $_
    $thisTeamMembers = get-graphUsersFromGroup -tokenResponse $tokenTeams -groupId $_.id -memberType TransitiveMembers -returnOnlyUsers
    $thisTeam | Add-Member NoteProperty -Name Members -Value $thisTeamMembers.userPrincipalName
    $thisTeamOwners = get-graphUsersFromGroup -tokenResponse $tokenTeams -groupId $_.id -memberType Owners -returnOnlyUsers
    $thisTeam | Add-Member NoteProperty -Name Owners -Value $thisTeamOwners.userPrincipalName
}

$zafTeams | Convert-OutputForCSV | Export-Csv -Path "$env:userprofile\Downloads\ZafTeams_$(get-date -f FileDateTimeUniversal).csv" -NoTypeInformation