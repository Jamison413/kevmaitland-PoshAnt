$tokenTeams = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName TeamsBot) -grant_type client_credentials
$tokenSharePoint = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName SharePointBot) -grant_type client_credentials
$365creds = set-MsolCredentials
connect-to365 -credential $365creds

$listofTeams = convertTo-arrayOfStrings "Growth and Impact Team (North America)
Marketing Team (North America)
Emerging Talent Collective Team (North America)
North America Leadeship Team (North America)
Climate and Net Zero Leadership Team (North America)
Carbon Markets Team (North America)
Climate Risk and Task Force on Climate-Related Financial Disclosures (TCFD) Team (North America)
Renewable Energy Team (North America)
Science Based Target (SBT) and Net Zero Team (North America)
Greenhouse Gas (GHG) Accounting Team (North America)
Water Stewardship Team (North America)
Communications and Reporting Team (North America)
Investment Strategy and Teams Team (North America)
Performance Data and Metrics Team (North America)
Strategy Setting Team (North America)
Social Impact Team (North America)
Life Cycle Assessment Team (North America)
Sustainable Packaging Team (North America)
Waste Team (North America)
Circular Business Models (North America)
Supply Chain and Operations Leadership Team (North America)
Supplier Engagement Team (North America)
"
$listofTeams += "Sustainable Products, Packaging, and Circularity Leadership Team (North America)"
$listofTeams += "Environmental, Social and Governance (ESG) and Sustainability Strategy Leadership Team (North America)"

$listofTeams = convertTo-arrayOfStrings "Product Sustainability Team (North America)
Information Solutions Team (North America)
Ventures Team (North America)
Operations Leadership Team (North America)
Administration Team (North America)
Finance Team (North America)
Human Resources (HR) Team (North America)
"


$listofTeams | % {
    Add-PnPListItem -List "Internal Team Site Requests" -Values @{Title=$_;Site_x0020_Type="Functional"}
    }

$communities = convertTo-arrayOfStrings "Solutions Community (North America)
Climate and Net Zero Community (North America)
Supply Chain and Operations Community (North America)
"
$communities += "Environmental, Social and Governance (ESG) and Sustainability Strategy Community (North America)"
$communities += "Sustainable Products, Packaging, and Circularity Community (North America)"


$communities | % {
    Add-PnPListItem -List "Internal Team Site Requests" -Values @{Title=$_;Site_x0020_Type="Aggregated Functional"}
    }




$teamsInCOmmunities=@(
    @{Team="Climate and Net Zero Leadership Team (North America)";Community="Climate and Net Zero Community (North America)"}
    @{Team="Carbon Markets Team (North America)";Community="Climate and Net Zero Community (North America)"}
    @{Team="Climate Risk and Task Force on Climate-Related Financial Disclosures (TCFD) Team (North America)";Community="Climate and Net Zero Community (North America)"}
    @{Team="Renewable Energy Team (North America)";Community="Climate and Net Zero Community (North America)"}
    @{Team="Science Based Target (SBT) and Net Zero Team (North America)";Community="Climate and Net Zero Community (North America)"}
    @{Team="Greenhouse Gas (GHG) Accounting Team (North America)";Community="Climate and Net Zero Community (North America)"}
    @{Team="Water Stewardship Team (North America)";Community="Climate and Net Zero Community (North America)"}
    @{Team="Environmental, Social and Governance (ESG) and Sustainability Strategy Leadership Team (North America)";Community="Environmental, Social and Governance (ESG) and Sustainability Strategy Community (North America)"}
    @{Team="Communications and Reporting Team (North America)";Community="Environmental, Social and Governance (ESG) and Sustainability Strategy Community (North America)"}
    @{Team="Investment Strategy and Teams Team (North America)";Community="Environmental, Social and Governance (ESG) and Sustainability Strategy Community (North America)"}
    @{Team="Performance Data and Metrics Team (North America)";Community="Environmental, Social and Governance (ESG) and Sustainability Strategy Community (North America)"}
    @{Team="Strategy Setting Team (North America)";Community="Environmental, Social and Governance (ESG) and Sustainability Strategy Community (North America)"}
    @{Team="Social Impact Team (North America)";Community="Environmental, Social and Governance (ESG) and Sustainability Strategy Community (North America)"}
    @{Team="Sustainable Products, Packaging, and Circularity Leadership Team (North America)";Community="Sustainable Products, Packaging, and Circularity Community (North America)"}
    @{Team="Life Cycle Assessment Team (North America)";Community="Sustainable Products, Packaging, and Circularity Community (North America)"}
    @{Team="Sustainable Packaging Team (North America)";Community="Sustainable Products, Packaging, and Circularity Community (North America)"}
    @{Team="Waste Team (North America)";Community="Sustainable Products, Packaging, and Circularity Community (North America)"}
    @{Team="Circular Business Models (North America)";Community="Sustainable Products, Packaging, and Circularity Community (North America)"}
    @{Team="Supply Chain and Operations Leadership Team (North America)";Community="Supply Chain and Operations Community (North America)"}
    @{Team="Supplier Engagement Team (North America)";Community="Supply Chain and Operations Community (North America)"}
    @{Team="Supply Chain and Operations Community (North America)";Community="Solutions Community (North America)"}
    @{Team="Climate and Net Zero Community (North America)";Community="Solutions Community (North America)"}
    @{Team="Environmental, Social and Governance (ESG) and Sustainability Strategy Community (North America)";Community="Solutions Community (North America)"}
    @{Team="Sustainable Products, Packaging, and Circularity Community (North America)";Community="Solutions Community (North America)"}
    )
    
$teamsInCOmmunities=@(
    @{Team="Product Sustainability (North America)";Community="Sustainable Products, Packaging, and Circularity Community (North America)"}
    @{Team="Information Solutions Team (North America)";Community="Supply Chain and Operations Community (North America)"}
    @{Team="Ventures Team (North America)";Community="Solutions Community (North America)"}
    )



$teamsInCOmmunities | ForEach-Object {
    $thisPair = $_
    $thisTeam = get-graphGroups -tokenResponse $tokenTeams -filterDisplayNameStartsWith $thisPair["Team"] -filterGroupType Unified -selectAllProperties
    $thisCommunity = get-graphGroups -tokenResponse $tokenTeams -filterDisplayNameStartsWith $thisPair["Community"] -filterGroupType Unified -selectAllProperties
    if($thisCommunity.displayName -match "aggregat"){
        set-graphGroup -tokenResponse $tokenTeams -groupId $thisCommunity.id -groupPropertyHash @{displayName=$($thisCommunity.displayName.Replace(" Aggregated Functional",""))}
        $combined = get-graphGroups -tokenResponse $tokenTeams -filterId $thisCommunity.anthesisgroup_UGSync.combinedGroupId
        Set-DistributionGroup -Identity $combined.id -DisplayName $combined.displayName.Replace(" Aggregated Functional","")
        $datamanagers = get-graphGroups -tokenResponse $tokenTeams -filterId $thisCommunity.anthesisgroup_UGSync.dataManagerGroupId
        Set-DistributionGroup -Identity $datamanagers.id -DisplayName $datamanagers.displayName.Replace(" Aggregated Functional","")
        $members = get-graphGroups -tokenResponse $tokenTeams -filterId $thisCommunity.anthesisgroup_UGSync.combinedGroupId
        Set-DistributionGroup -Identity $members.id -DisplayName $members.displayName.Replace(" Aggregated Functional","")
        }
    Write-host "Adding [$($thisTeam.displayName)] to [$($thisCommunity.displayName)]"
    Add-DistributionGroupMember -Identity $thisCommunity.anthesisgroup_UGSync.memberGroupId -Member $thisTeam.anthesisgroup_UGSync.memberGroupId -Confirm:$false -BypassSecurityGroupManagerCheck
    }

$oldFolderNames = @()
#$oldFolderNames += "0 ADMIN"
#$oldFolderNames += "1 TEMPLATES"
#$oldFolderNames += "2 PRESENTATIONS"
#$oldFolderNames += "3 CASE STUDIES"
#$oldFolderNames += "4 MEDIA"
#$oldFolderNames += "5 REFERENCE"
$oldFolderNames += "2 PRESENTATIONS (Shared)"
$oldFolderNames += "3 CASE STUDIES (Shared)"
$oldFolderNames += "4 MEDIA (Shared)"
$oldFolderNames += "5 REFERENCE (Shared)"
$oldFolderNames += "6 MEDIA (Shared)"
$oldFolderNames += "7 REFERENCE (Shared)"


$solutionFoldersToCreate = @()
$solutionFoldersToCreate += "0 ADMIN"
$solutionFoldersToCreate += "1 TEMPLATES (Shared)"
$solutionFoldersToCreate += "2 CASE STUDIES (Shared)"
$solutionFoldersToCreate += "3 MEDIA (Shared)"
$solutionFoldersToCreate += "4 REFERENCE (Shared)"
$solutionFoldersToCreate += "5 PRESENTATIONS (Shared)"

$teamFoldersToCreate = $solutionFoldersToCreate
$teamFoldersToCreate += "6 MEDIA (Shared)"
$teamFoldersToCreate += "7 TEAM BIOS (Shared)"
$teamFoldersToCreate += "8 PROPOSALS (Shared)"

$listofTeams | % {
    $thisTeam = $_
    $thisTeamTeam = get-graphGroups -tokenResponse $tokenTeams -filterDisplayNameStartsWith $thisTeam -filterGroupType Unified
    #Temporarily add me
    #Add-DistributionGroupMember -Identity $thisTeamTeam.anthesisgroup_UGSync.dataManagerGroupId -BypassSecurityGroupManagerCheck:$true -Member t0-kevin.maitland@anthesisgroup.com
    add-graphUsersToGroup -tokenResponse $tokenTeams -graphGroupId $thisTeamTeam.id -memberType owners -graphUserIds 135feab0-fb9d-4ac1-a7a8-c40b66c75ddc #-graphUserUpns t0-kevin.maitland@anthesisgroup.com 

    if([string]::IsNullOrWhiteSpace($thisTeamTeam.id)){write-host "Team [$thisTeam] not retrieved";return}
    else{Write-Host -f Yellow "Team [$($thisTeamTeam.displayName)] retrieved"}
    $thisTeamDrive = get-graphDrives -tokenResponse $tokenTeams -groupGraphId $thisTeamTeam.id -returnOnlyDefaultDocumentsLibrary
    if([string]::IsNullOrWhiteSpace($thisTeamDrive.id)){write-host "Team Drive for [$($thisTeamTeam.displayName)] not retrieved";return}
    #Remove any duff Channels
    $currentChannels = get-graphTeamChannels -tokenResponse $tokenTeams -teamId $thisTeamTeam.id -channelType Both
    $currentChannels | ? {$oldFolderNames -contains $_.displayName} | % {
        $thisChannel = $_
        write-host "`tDeleting Channel [$($thisChannel.displayName)]"
        delete-graphTeamChannel -tokenResponse $tokenTeams -teamId $thisTeamTeam.id -channelId $thisChannel.id
        }
    #Remove any duff folders
    $currentFolders = get-graphDriveItems -tokenResponse $tokenTeams -driveGraphId $thisTeamDrive.id -returnWhat Children
    $currentFolders | ? {$_.folder -ne $null -and ($oldFolderNames -contains $_.name -or $_.name -eq "Team Files"-or $_.name -eq "Community Files")} | % {
        $thisFolder = $_
        write-host "`tDeleting Folder [$($thisFolder.name)]"
        #grant-graphSharing -tokenResponse $tokenTeams -driveId $thisTeamDrive.id -itemId $thisFolder.id -sharingRecipientsUpns kevin.maitland@anthesisgroup.com -requireSignIn $true -sendInvitation $false -role Write
        try   {delete-graphDriveItem -tokenResponse $tokenSharePoint -graphDriveId $thisTeamDrive.id -graphDriveItemId $thisFolder.id}
        catch {
            try   {delete-graphDriveItem -tokenResponse $tokenTeams -graphDriveId $thisTeamDrive.id -graphDriveItemId $thisFolder.id}
            catch {[array]$duffSites += $thisTeamTeam}
            }
        }
    #deploy the final folder names
    $resultFolders = add-graphArrayOfFoldersToDrive -tokenResponse $tokenTeams -graphDriveId $thisTeamDrive.id -foldersAndSubfoldersArray $teamFoldersToCreate -conflictResolution Fail
    #deploy channels with matching names
    $teamFoldersToCreate | % {
        $thisChannel = $_
        new-graphTeamChannel -tokenResponse $tokenTeams -teamId $thisTeamTeam.id -membershipType standard -channelName $thisChannel -isFavourite $true

        }

    #Share folders
    $resultFolders | ? {$_.name -match "(Shared)"} | % {
        $thisFolder = $_
        #grant-graphSharing -tokenResponse $tokenTeams -driveId $thisTeamDrive.id -itemId $thisFolder.id -sharingRecipientsUpns Solutions_Community_North_America_Aggregated_Functional_365@anthesisgroup.com -sendInvitation $false -role Read -requireSignIn $true
        }

    #create _Archive subfolders
    $resultFolders | % {
        $thisResultFolder= $_ 
        $archiveSubfolder = add-graphArrayOfFoldersToDrive -tokenResponse $tokenTeams -graphDriveId $thisTeamDrive.id -foldersAndSubfoldersArray "$($thisResultFolder.Name)\_ARCHIVE" -conflictResolution Fail | ? {$_.name -eq "_ARCHIVE"}
        $archiveSubfolderPermissions = get-graphDriveItemPermissions -tokenResponse $tokenTeams -driveGraphId $thisTeamDrive.id -itemGraphId $archiveSubfolder.id
        #Hide _archive subfolders
        $communityPermissions = $archiveSubfolderPermissions | ? {$_.grantedToV2.group.email -eq "Solutions_Community_North_America_Aggregated_Functional_365@anthesisgroup.com"}
        $communityPermissions | % {
            delete-graphDriveItemPermission -tokenResponse $tokenTeams -graphDriveId $thisTeamDrive.id -graphDriveItemId $archiveSubfolder.id -graphDriveItemPermissionId $_.id
            }
        }


    }

get-graphUsers -tokenResponse $tokenTeams -filterUpns @("t0-kevin.maitland@anthesisgroup.com","groupbot@anthesisgroup.com")

$communities | % {
    $thisTeam = $_
    $thisTeamTeam = get-graphGroups -tokenResponse $tokenTeams -filterDisplayNameStartsWith $thisTeam -filterGroupType Unified
    add-graphUsersToGroup -tokenResponse $tokenTeams -graphGroupId $thisTeamTeam.id -memberType owners -graphUserIds 135feab0-fb9d-4ac1-a7a8-c40b66c75ddc
    remove-graphUsersFromGroup -tokenResponse $tokenTeams -graphGroupId $thisTeamTeam.id -memberType owners -graphUserIds 00aa81e4-2e8f-4170-bc24-843b917fd7cf
    }

$communities | % {
    $thisCommunity = $_
    $thisCommunityTeam = get-graphGroups -tokenResponse $tokenTeams -filterDisplayNameStartsWith $thisCommunity -filterGroupType Unified
    if([string]::IsNullOrWhiteSpace($thisCommunityTeam.id)){write-host "Team [$thisTeam] not retrieved";return}
    else{Write-Host -f Yellow "Team [$($thisCommunityTeam.displayName)] retrieved"}
    $thisCommunityDrive = get-graphDrives -tokenResponse $tokenTeams -groupGraphId $thisCommunityTeam.id -returnOnlyDefaultDocumentsLibrary
    if([string]::IsNullOrWhiteSpace($thisCommunityDrive.id)){write-host "Team Drive for [$($thisCommunityTeam.displayName)] not retrieved";return}
    #Remove any duff Channels
    $currentChannels = get-graphTeamChannels -tokenResponse $tokenTeams -teamId $thisCommunityTeam.id -channelType Both
    $currentChannels | ? {$oldFolderNames -contains $_.displayName} | % {
        $thisChannel = $_
        write-host "`tDeleting Channel [$($thisChannel.displayName)]"
        delete-graphTeamChannel -tokenResponse $tokenTeams -teamId $thisCommunityTeam.id -channelId $thisChannel.id
        }
    #Remove any duff folders
    $currentFolders = get-graphDriveItems -tokenResponse $tokenTeams -driveGraphId $thisCommunityDrive.id -returnWhat Children
    $currentFolders | ? {$_.folder -ne $null -and ($oldFolderNames -contains $_.name -or $_.name -eq "Team Files"-or $_.name -eq "Community Files")} | % {
        $thisFolder = $_
        write-host "`tDeleting Folder [$($thisFolder.name)]"
        #grant-graphSharing -tokenResponse $tokenTeams -driveId $thisCommunityDrive.id -itemId $thisFolder.id -sharingRecipientsUpns kevin.maitland@anthesisgroup.com -requireSignIn $true -sendInvitation $false -role Write
        try   {delete-graphDriveItem -tokenResponse $tokenSharePoint -graphDriveId $thisCommunityDrive.id -graphDriveItemId $thisFolder.id}
        catch {
            try   {delete-graphDriveItem -tokenResponse $tokenTeams -graphDriveId $thisCommunityDrive.id -graphDriveItemId $thisFolder.id}
            catch {[array]$duffSites += $thisCommunityTeam}
            }
        }
    #deploy the final folder names
    $resultFolders = add-graphArrayOfFoldersToDrive -tokenResponse $tokenTeams -graphDriveId $thisCommunityDrive.id -foldersAndSubfoldersArray $solutionFoldersToCreate -conflictResolution Fail
    #deploy channels with matching names
    $solutionFoldersToCreate | % {
        $thisChannel = $_
        new-graphTeamChannel -tokenResponse $tokenTeams -teamId $thisCommunityTeam.id -membershipType standard -channelName $thisChannel -isFavourite $true

        }

    #Share folders with super-community
    $resultFolders | ? {$_.name -match "(Shared)"} | % {
        $thisFolder = $_
        grant-graphSharing -tokenResponse $tokenTeams -driveId $thisCommunityDrive.id -itemId $thisFolder.id -sharingRecipientsUpns Solutions_Community_North_America_Aggregated_Functional_365@anthesisgroup.com -sendInvitation $false -role Read -requireSignIn $true
        }

    #create _Archive subfolders
    $resultFolders | % {
        $thisResultFolder= $_ 
        $archiveSubfolder = add-graphArrayOfFoldersToDrive -tokenResponse $tokenTeams -graphDriveId $thisCommunityDrive.id -foldersAndSubfoldersArray "$($thisResultFolder.Name)\_ARCHIVE" -conflictResolution Fail | ? {$_.name -eq "_ARCHIVE"}
        $archiveSubfolderPermissions = get-graphDriveItemPermissions -tokenResponse $tokenTeams -driveGraphId $thisCommunityDrive.id -itemGraphId $archiveSubfolder.id
        #Hide _archive subfolders
        if($thisCommunityTeam.mail -ne "Solutions_Community_North_America_Aggregated_Functional_365@anthesisgroup.com"){
            $communityPermissions = $archiveSubfolderPermissions | ? {$_.grantedToV2.group.email -eq "Solutions_Community_North_America_Aggregated_Functional_365@anthesisgroup.com"}
            $communityPermissions | % {
                delete-graphDriveItemPermission -tokenResponse $tokenTeams -graphDriveId $thisCommunityDrive.id -graphDriveItemId $archiveSubfolder.id -graphDriveItemPermissionId $_.id
                }
            }    
        }
    }


$teamsInCOmmunities | % {
    $thisPair = $_
    $thisTeam = get-graphGroups -tokenResponse $tokenTeams -filterDisplayNameStartsWith $thisPair["Team"] -filterGroupType Unified
    if($thisCommunity.displayName -ne $thisPair["Community"]){$thisCommunity = get-graphGroups -tokenResponse $tokenTeams -filterDisplayNameStartsWith $thisPair["Community"] -filterGroupType Unified}
    if([string]::IsNullOrWhiteSpace($thisTeam.id)){write-host "Team [$($thisPair["Team"])] not retrieved";return}
    else{Write-Host -f Yellow "Team [$($thisTeam.displayName)] retrieved"}
    if([string]::IsNullOrWhiteSpace($thisCommunity.id)){write-host "Team [$($thisPair["Community"])] not retrieved";return}
    else{Write-Host -f Yellow "Team [$($thisCommunity.displayName)] retrieved"}
    $thisTeamDrive = get-graphDrives -tokenResponse $tokenTeams -groupGraphId $thisTeam.id -returnOnlyDefaultDocumentsLibrary
    if([string]::IsNullOrWhiteSpace($thisTeamDrive.id)){write-host "Team Drive for [$($thisTeam.displayName)] not retrieved";return}
    $thisCommunityDrive = get-graphDrives -tokenResponse $tokenTeams -groupGraphId $thisCommunity.id -returnOnlyDefaultDocumentsLibrary
    if([string]::IsNullOrWhiteSpace($thisCommunity.id)){write-host "Team Drive for [$($thisCommunity.displayName)] not retrieved";return}

    #Create Channels
    $solutionFoldersToCreate | % {
        $thisChannel = $_
        new-graphTeamChannel -tokenResponse $tokenTeams -teamId $thisTeam.id -membershipType standard -channelName $thisChannel -isFavourite $true
        }
    #Create Folders
    $thisTeamDriveItems = add-graphArrayOfFoldersToDrive -tokenResponse $tokenTeams -graphDriveId $thisTeamDrive.id -foldersAndSubfoldersArray $solutionFoldersToCreate -conflictResolution Fail | ? {$solutionFoldersToCreate -contains $_.name} #Create any Channel folders that haven't been automatically provisioned
    #Share "shared" folders with super-community
    $thisTeamDriveItems | ? {$_.name -match "(Shared)"} | % { #Share the stanadrd folders in the Team with the Community
        $thisTeamDriveItem = $_
        grant-graphSharing -tokenResponse $tokenTeams -driveId $thisTeamDrive.id -itemId $thisTeamDriveItem.id -sharingRecipientsUpns Solutions_Community_North_America_Aggregated_Functional_365@anthesisgroup.com -role Read -sendInvitation $false -requireSignIn $true
        }

    #create _Archive subfolders
    $thisTeamDriveItems | % {
        $thisResultFolder= $_ 
        $archiveSubfolder = add-graphArrayOfFoldersToDrive -tokenResponse $tokenTeams -graphDriveId $thisCommunityDrive.id -foldersAndSubfoldersArray "$($thisResultFolder.Name)\_ARCHIVE" -conflictResolution Fail | ? {$_.name -eq "_ARCHIVE"}
        $archiveSubfolderPermissions = get-graphDriveItemPermissions -tokenResponse $tokenTeams -driveGraphId $thisCommunityDrive.id -itemGraphId $archiveSubfolder.id
        #Hide _archive subfolders from Community
        if($thisCommunityTeam.mail -ne "Solutions_Community_North_America_Aggregated_Functional_365@anthesisgroup.com"){
            $communityPermissions = $archiveSubfolderPermissions | ? {$_.grantedToV2.group.email -eq "Solutions_Community_North_America_Aggregated_Functional_365@anthesisgroup.com"}
            $communityPermissions | % {
                delete-graphDriveItemPermission -tokenResponse $tokenTeams -graphDriveId $thisCommunityDrive.id -graphDriveItemId $archiveSubfolder.id -graphDriveItemPermissionId $_.id
                }
            }    
        }

    #Add Community Files tab to Team
    $thisGeneralChannel = get-graphTeamChannels -tokenResponse $tokenTeams -teamId $thisTeam.id -channelType Both | ? {$_.displayName -eq "General"}
    $theseGeneralChannelTabs = get-graphTeamChannelTabs -tokenResponse $tokenTeams -teamId $thisTeam.id -channelId $thisGeneralChannel.id
    if($theseGeneralChannelTabs.displayName -notcontains "Community Files"){
        add-graphWebsiteTabToChannel -tokenResponse $tokenTeams -teamId $thisTeam.id -channelName "General" -tabName "Community Files" -tabDestinationUrl $thisCommunityDrive.webUrl
        }


    #now do stuff in the Community
    $teamFolderInCommunity = get-graphDriveItems -tokenResponse $tokenTeams -driveGraphId $thisCommunityDrive.id -returnWhat Children | ? {$_.name -eq "$($thisTeam.displayName) files"} #Get or create the Team-in-Community folder
    if($teamFolderInCommunity.id -eq $null){$teamFolderInCommunity = add-graphFolderToDrive -tokenResponse $tokenTeams -graphDriveId $thisCommunityDrive.id -folderName "$($thisTeam.displayName) files" -conflictResolution Fail}
    $thisTeamDriveItems | ? {$_.name -match "(Shared)"}  | % {
        $thisTeamDriveItems = $_
        $newHyperlinkContent = `
"[InternetShortcut]
URL=$($thisTeamDriveItems.webUrl)
"
        $newHyperlink = invoke-graphPut -tokenResponse $tokenTeams -graphQuery "/drives/$($thisCommunityDrive.id)/items/$($teamFolderInCommunity.id):/$([uri]::EscapeUriString($thisTeamDriveItems.name)).url:/content" -binaryFileStream $newHyperlinkContent
        }



    }



#Set GroupBot as Data Manager
$listofTeams | % {
    $thisTeam = $_
    $thisTeamTeam = get-graphGroups -tokenResponse $tokenTeams -filterDisplayNameStartsWith $thisTeam -filterGroupType Unified -selectAllProperties
    #Add-DistributionGroupMember -Identity $thisTeamTeam.anthesisgroup_UGSync.dataManagerGroupId -BypassSecurityGroupManagerCheck:$true -Member groupbot@anthesisgroup.com
    Add-DistributionGroupMember -Identity $thisTeamTeam.anthesisgroup_UGSync.dataManagerGroupId -BypassSecurityGroupManagerCheck:$true -Member t0-kevin.maitland@anthesisgroup.com
    #Remove-DistributionGroupMember -Identity $thisTeamTeam.anthesisgroup_UGSync.dataManagerGroupId -BypassSecurityGroupManagerCheck:$true -Member t0-kevin.maitland@anthesisgroup.com
    }
$communities | % {
    $thisTeam = $_
    $thisTeamTeam = get-graphGroups -tokenResponse $tokenTeams -filterDisplayNameStartsWith $thisTeam -filterGroupType Unified -selectAllProperties
    Add-DistributionGroupMember -Identity $thisTeamTeam.anthesisgroup_UGSync.dataManagerGroupId -BypassSecurityGroupManagerCheck:$true -Member groupbot@anthesisgroup.com
    }


$renameTeams = @(
    @("Sustainable Chemistry Team (North America)","Product Sustainability (North America)"),
    @("Operational Leadership Team (North America)","Operations Leadership Team (North America)")
    )




$dataManagers = @(
	@("Administration Team (North America)","maggie.weglinski@anthesisgroup.com"),
	@("Carbon Markets Team (North America)","Miquel.Rubio@anthesisgroup.com"),
	@("Circular Business Models (North America)","Dawn.ManciniMoyer@anthesisgroup.com"),
	@("Climate and Net Zero Community (North America)","Stephen.Russell@anthesisgroup.com"),
	@("Climate and Net Zero Leadership Team (North America)","Stephen.Russell@anthesisgroup.com"),
	@("Climate Risk and Task Force on Climate-Related Financial Disclosures (TCFD) Team (North America)","Peter.van.Dijk@anthesisgroup.com"),
	@("Communications and Reporting Team (North America)","Amanda.Pinyan@anthesisgroup.com"),
	@("Emerging Talent Collective Team (North America)","Manisha.Paralikar@anthesisgroup.com"),
	@("Environmental, Social and Governance (ESG) and Sustainability Strategy Community (North America)","Jennifer.Clipsham@anthesisgroup.com"),
	@("Environmental, Social and Governance (ESG) and Sustainability Strategy Leadership Team (North America)","Jennifer.Clipsham@anthesisgroup.com"),
	@("Finance Team (North America)","Kelvin.Cabaldo@anthesisgroup.com"),
	@("Finance Team (North America)","maggie.weglinski@anthesisgroup.com"),
	@("Greenhouse Gas (GHG) Accounting Team (North America)","Sophia.Traweek@anthesisgroup.com"),
	#@("Growth and Impact Leadership Team (North America)","Chantelle.Ludski@anthesisgroup.com"),
	#@("Growth and Impact Leadership Team (North America)","John.Heckman@anthesisgroup.com"),
	@("Growth and Impact Team (North America)","John.Heckman@anthesisgroup.com"),
	@("Human Resources (HR) Team (North America)","maggie.weglinski@anthesisgroup.com"),
	@("Information Solutions Team (North America)","Jason.Gooden@anthesisgroup.com"),
	@("Investment Strategy and Teams Team (North America)","Barrett.Lawson@anthesisgroup.com"),
	@("Life Cycle Assessment Team (North America)","Caroline.Gaudreault@anthesisgroup.com"),
	@("Marketing Team (North America)","Jackie.Fleming@anthesisgroup.com"),
	@("North America Leadeship Team (North America)","Chantelle.Ludski@anthesisgroup.com"),
	@("Operations Leadership Team (North America)","maggie.weglinski@anthesisgroup.com"),
	@("Performance Data and Metrics Team (North America)","Jon.Taylor@anthesisgroup.com"),
	@("Product Sustainability Team (North America)","Lemis.Tarajano.Noya@anthesisgroup.com"),
	@("Renewable Energy Team (North America)","Stephen.Russell@anthesisgroup.com"),
	@("Science Based Target (SBT) and Net Zero Team (North America)","Curtis.Harnanan@anthesisgroup.com"),
	@("Social Impact Team (North America)","Emma.Armstrong@anthesisgroup.com"),
	@("Social Impact Team (North America)","Jason.Pearson@anthesisgroup.com"),
	@("Solutions Community (North America)","maggie.weglinski@anthesisgroup.com"),
	@("Strategy Setting Team (North America)","Alison.Murphy@anthesisgroup.com"),
	@("Strategy Setting Team (North America)","Jason.Pearson@anthesisgroup.com"),
	@("Supplier Engagement Team (North America)","Elena.Kocherovsky@anthesisgroup.com"),
	@("Supply Chain and Operations Community (North America)","Hannah.Aneiros@anthesisgroup.com"),
	@("Supply Chain and Operations Leadership Team (North America)","Jason.Gooden@anthesisgroup.com"),
	@("Sustainable Packaging Team (North America)","Carolyn.Klindt@anthesisgroup.com"),
	@("Sustainable Products, Packaging, and Circularity Community (North America)","Dawn.ManciniMoyer@anthesisgroup.com"),
	@("Sustainable Products, Packaging, and Circularity Leadership Team (North America)","Dawn.ManciniMoyer@anthesisgroup.com"),
	@("Ventures Team (North America)","Gabriel.Vanloozen@anthesisgroup.com"),
	@("Waste Team (North America)","Dawn.ManciniMoyer@anthesisgroup.com"),
	@("Water Stewardship Team (North America)","Stephen.Russell@anthesisgroup.com")
    )
$dataManagers | % {
    $thisDataManagerAssignment = $_
    $thisTeam = $thisDataManagerAssignment[0]
    $thisDataManager = $thisDataManagerAssignment[1]
    $thisGraphTeam = get-graphGroups -tokenResponse $tokenTeams -filterDisplayName $thisTeam -filterGroupType Unified -selectAllProperties
    if([string]::IsNullOrWhiteSpace($thisGraphTeam)){Write-Host -f Red "[$($thisTeam)] could not be retrieved"}
    else{Write-Host -f Yellow "[$($thisTeam)] was retrieved"}
    Add-DistributionGroupMember -Identity $thisGraphTeam.anthesisgroup_UGSync.dataManagerGroupId -BypassSecurityGroupManagerCheck:$true -Member $thisDataManager
    remove-DistributionGroupMember -Identity $thisGraphTeam.anthesisgroup_UGSync.dataManagerGroupId -BypassSecurityGroupManagerCheck:$true -Member t0-kevin.maitland@anthesisgroup.com
    }

$newGeoTeams = @(
	"All Washington DC (USA)",
	"All New York (USA)",
	"All Seattle (USA)",
	"All Guelph (CAN)",
	"All Montreal (CAN)",
	"All Vancouver (CAN)"
    )


#Add members to relevant Teams
$memberships = Import-Csv $env:USERPROFILE\Downloads\NA2_Members.csv
$teams = @{}
$($memberships[0].psobject.Properties).Name | ? {$_ -ne "Email"} | % {
    $graphTeam = get-graphGroups -tokenResponse $tokenTeams -filterDisplayName $_ -filterGroupType Unified
    if([string]::IsNullOrWhiteSpace($graphTeam.id)){write-host -f darkred "[$_] could not be retrieved"}
    $teams.Add($_,$graphTeam.id)
    }

$duffUsers = @()
$memberships | % {
    $thisUser = $_#memberships[0]
    $thisGraphUser = get-graphUsers -tokenResponse $tokenTeams -filterUpns $($thisUser.Email).Trim()
    if([string]::IsNullOrWhiteSpace($thisGraphUser.id)){write-host -f Red "[$($thisUser.Email)] could not be retrieved";$duffUsers += $thisUser.Email;return}
    for($i=1;$i -lt $($thisUser.psobject.Properties).Count; $i++){
        if(![string]::IsNullOrWhiteSpace($($thisUser.psobject.Properties)[$i].Value)){
            Write-Host -f Yellow "Adding [$($thisUser.Email)] to [$($($thisUser.psobject.Properties)[$i].Name)][$($teams[$($($thisUser.psobject.Properties)[$i].Name)])]"
            add-graphUsersToGroup -tokenResponse $tokenTeams -graphGroupId $teams[$($($thisUser.psobject.Properties)[$i].Name)] -memberType members -graphUserIds $thisGraphUser.id
            }
        }
    }


#Finally, remove t0-kevin.maitland and groupbot (if possible)
$teams.Values | % {
    $teamId = $_
    $thisTeam = get-graphGroups -tokenResponse $tokenTeams -filterId $teamId -selectAllProperties
    write-host -f Yellow "Processing [$($thisTeam.displayName)]"
    $dataManagers = get-graphUsersFromGroup -tokenResponse $tokenTeams -groupId $thisTeam.anthesisgroup_UGSync.dataManagerGroupId -memberType Members
    write-host -f Yellow "`t[$($dataManagers.Count)] Data Managers retrieved: `r`n`t`t$($dataManagers.userPrincipalName -join "`r`n`t`t")"
    if($dataManagers.userPrincipalName -contains "t0-kevin.maitland@anthesisgroup.com"){
        try{
            #remove-graphUsersFromGroup -tokenResponse $tokenTeams -graphGroupId $thisTeam.anthesisgroup_UGSync.dataManagerGroupId -memberType Members -graphUserIds 135feab0-fb9d-4ac1-a7a8-c40b66c75ddc
            write-host -f Magenta "`tRemoving [t0-kevin.maitland@anthesisgroup.com]"
            Remove-DistributionGroupMember -Identity $thisTeam.anthesisgroup_UGSync.dataManagerGroupId -BypassSecurityGroupManagerCheck:$true -Confirm:$false -Member t0-kevin.maitland@anthesisgroup.com
            }
        catch{get-errorSummary $_}
        }
    if($dataManagers.userPrincipalName -contains "groupbot@anthesisgroup.com"){
        if($dataManagers.Count -ge 3 -or ($dataManagers.Count -eq 2 -and $dataManagers.userPrincipalName -notcontains "t0-kevin.maitland@anthesisgroup.com")){
            try{
                #remove-graphUsersFromGroup -tokenResponse $tokenTeams -graphGroupId $thisTeam.anthesisgroup_UGSync.dataManagerGroupId -memberType Members -graphUserIds 00aa81e4-2e8f-4170-bc24-843b917fd7cf
                write-host -f Magenta "`tRemoving [groupbot@anthesisgroup.com]"
                Remove-DistributionGroupMember -Identity $thisTeam.anthesisgroup_UGSync.dataManagerGroupId -BypassSecurityGroupManagerCheck:$true -Confirm:$false -Member groupbot@anthesisgroup.com
                }
            catch{get-errorSummary $_}
            }
        }
    } 



$dataManagers | % {
    #$thisDMPair = $_
    $_[1]
    } | Sort-Object -Unique
$dataManagersObjects = @()
$dataManagers | % {
     #$thisDM = New-Object psobject -ArgumentList @{userPrincipalName=$_[1]}
    $dataManagersObjects += [PSCustomObject]@{userPrincipalName=$_[1]}
    }

$allUsers = get-graphUsers -tokenResponse $tokenTeams -selectAllProperties

$trainedDataManagers = Import-Csv $env:USERPROFILE\Downloads\DataManagerTrainingDates.csv -Encoding UTF7
$trainedDataManagers = $trainedDataManagers | Group-Object {$_.User} | % {$_.Group | Sort-Object {Get-Date $_.'Date of training'} -Descending | Select-Object -First 1}
$trainedDataManagers = $trainedDataManagers | ? {$_.User -notmatch '\?'}
$trainedDataManagers | % {
    $thisDMRecord = $_
    $thisDM = $allUsers | ? {$_.DisplayName -eq $thisDMRecord.User} 
    if([string]::IsNullOrWhiteSpace($thisDM.id)){
        Write-Host -f red "Could not identify [$($thisDMRecord.User)]"
        }
    else{
        $thisDMRecord | Add-Member -MemberType NoteProperty -Name userPrincipalName -Value $thisDM.userPrincipalName -Force
        }
    }
$trainedDataManagers = $trainedDataManagers | ? {![string]::IsNullOrWhiteSpace($_.userPrincipalName)}

$dataManagersObjects | % {
    $thisNADM = $_ 
    $thisRecord = Compare-Object -ReferenceObject $trainedDataManagers -DifferenceObject $thisNADM -IncludeEqual -PassThru -Property userPrincipalName -ExcludeDifferent
    $thisNADM | Add-Member -MemberType NoteProperty -Name LastTraining -Value $thisRecord.'Date of training' -Force
    $thisNADM | Add-Member -MemberType NoteProperty -Name LastReminder -Value $thisRecord.'Last Reminder Email Sent' -Force
    }

$dataManagersObjects | Sort-Object userPrincipalName -Unique







$thingsInCommunities = @(
    @{Team="Climate and Net Zero Leadership Team (North America)";Community="Climate and Net Zero Community (North America)"}
    ,@{Team="Carbon Markets Team (North America)";Community="Climate and Net Zero Community (North America)"}
    ,@{Team="Climate Risk and Task Force on Climate-Related Financial Disclosures (TCFD) Team (North America)";Community="Climate and Net Zero Community (North America)"}
    ,@{Team="Renewable Energy Team (North America)";Community="Climate and Net Zero Community (North America)"}
    ,@{Team="Science Based Target (SBT) and Net Zero Team (North America)";Community="Climate and Net Zero Community (North America)"}
    ,@{Team="Greenhouse Gas (GHG) Accounting Team (North America)";Community="Climate and Net Zero Community (North America)"}
    ,@{Team="Water Stewardship Team (North America)";Community="Climate and Net Zero Community (North America)"}

    ,@{Team="Environmental, Social and Governance (ESG) and Sustainability Strategy Leadership Team (North America)";Community="Environmental, Social and Governance (ESG) and Sustainability Strategy Community (North America)"}
    ,@{Team="Communications and Reporting Team (North America)";Community="Environmental, Social and Governance (ESG) and Sustainability Strategy Community (North America)"}
    ,@{Team="Investment Strategy and Teams Team (North America)";Community="Environmental, Social and Governance (ESG) and Sustainability Strategy Community (North America)"}
    ,@{Team="Performance Data and Metrics Team (North America)";Community="Environmental, Social and Governance (ESG) and Sustainability Strategy Community (North America)"}
    ,@{Team="Strategy Setting Team (North America)";Community="Environmental, Social and Governance (ESG) and Sustainability Strategy Community (North America)"}
    ,@{Team="Social Impact Team (North America)";Community="Environmental, Social and Governance (ESG) and Sustainability Strategy Community (North America)"}

    ,@{Team="Sustainable Products, Packaging, and Circularity Leadership Team (North America)";Community="Sustainable Products, Packaging, and Circularity Community (North America)"}
    ,@{Team="Life Cycle Assessment Team (North America)";Community="Sustainable Products, Packaging, and Circularity Community (North America)"}
    ,@{Team="Sustainable Packaging Team (North America)";Community="Sustainable Products, Packaging, and Circularity Community (North America)"}
    ,@{Team="Waste Team (North America)";Community="Sustainable Products, Packaging, and Circularity Community (North America)"}
    ,@{Team="Circular Business Models (North America)";Community="Sustainable Products, Packaging, and Circularity Community (North America)"}
    ,@{Team="Product Sustainability Team (North America)";Community="Sustainable Products, Packaging, and Circularity Community (North America)"}

    ,@{Team="Supply Chain and Operations Leadership Team (North America)";Community="Supply Chain and Operations Community (North America)"}
    ,@{Team="Supplier Engagement Team (North America)";Community="Supply Chain and Operations Community (North America)"}
    ,@{Team="Information Solutions Team (North America)";Community="Supply Chain and Operations Community (North America)"}

    ,@{Team="Supply Chain and Operations Community (North America)";Community="Solutions Community (North America)"}
    ,@{Team="Climate and Net Zero Community (North America)";Community="Solutions Community (North America)"}
    ,@{Team="Environmental, Social and Governance (ESG) and Sustainability Strategy Community (North America)";Community="Solutions Community (North America)"}
    ,@{Team="Sustainable Products, Packaging, and Circularity Community (North America)";Community="Solutions Community (North America)"}
    ,@{Team="Ventures Team (North America)";Community="Solutions Community (North America)"}
    )
###########################################################################
#Build Navigation
###########################################################################
$thingsInCommunities | ForEach-Object {
    $thisPair = $_
    $allTeamNames += [array]$thisPair["Team"]
    $allCommunityNames += [array]$thisPair["Community"]
    }

$allTeamNames = $allTeamNames | Select-Object -Unique
$allTeamNames | ForEach-Object {
    $allTeams += [array]$(get-graphGroups -tokenResponse $tokenTeams -filterDisplayName $_ -filterGroupType Unified)
    } 
$allTeams | ForEach-Object {
    Write-Host "Team [$($_.displayName)]"
    $thisTeamDefaultDrive = $(get-graphDrives -tokenResponse $tokenTeams -groupGraphId $_.id -returnOnlyDefaultDocumentsLibrary)
    $_ | Add-Member -MemberType NoteProperty -Name DefaultDriveId -Value $thisTeamDefaultDrive.id -Force
    $_ | Add-Member -MemberType NoteProperty -Name DefaultDriveUrl -Value $thisTeamDefaultDrive.webUrl -Force
    }

$allCommunityNames = $allCommunityNames | Select-Object -Unique
$allCommunityNames | ForEach-Object {
    $allCommunities += [array]$(get-graphGroups -tokenResponse $tokenTeams -filterDisplayName $_ -filterGroupType Unified)
    }
$allCommunities | ForEach-Object {
    $thisCommunityDefaultDrive = $(get-graphDrives -tokenResponse $tokenTeams -groupGraphId $_.id -returnOnlyDefaultDocumentsLibrary)
    $_ | Add-Member -MemberType NoteProperty -Name DefaultDriveId -Value $thisCommunityDefaultDrive.id -Force
    $_ | Add-Member -MemberType NoteProperty -Name DefaultDriveUrl -Value $thisCommunityDefaultDrive.webUrl -Force
    }
$superCommunity = $allCommunities | Where-Object {$_.displayName -eq "Solutions Community (North America)"}

$thingsInCommunities | ForEach-Object {
    $thisPair = $_
    $thisTeam = $allTeams | ? {$_.displayName -eq $thisPair["Team"]}
    $thisCommunity = $allCommunities | ? {$_.displayName -eq $thisPair["Community"]}

    #1 Create SharePoint tabs pointing to the Team, Community & Super Community files
    $thisTeamGeneralTab = get-graphTeamChannels -tokenResponse $tokenTeams -teamId $thisTeam.id -channelType Both | Where-Object {$_.displayName -eq "General"}
    if($thisTeam.displayName -match "Team"){
        $teamFilesTab = new-graphTeamChannelTab -tokenResponse $tokenTeams -teamId $thisTeam.id -channelId $thisTeamGeneralTab.id -tabType SharePoint -tabName "Team Files" -tabDestinationUrl $thisTeam.DefaultDriveUrl
        $communityFilesTab = new-graphTeamChannelTab -tokenResponse $tokenTeams -teamId $thisTeam.id -channelId $thisTeamGeneralTab.id -tabType SharePoint -tabName "Community Files" -tabDestinationUrl $thisCommunity.DefaultDriveUrl
        }
    else{
        $communityFilesTab = new-graphTeamChannelTab -tokenResponse $tokenTeams -teamId $thisTeam.id -channelId $thisTeamGeneralTab.id -tabType SharePoint -tabName "Community Files" -tabDestinationUrl $thisTeam.DefaultDriveUrl #If we're looking at a Community, we don't want to call the Tab "Team Files"
        }
    $superCommunityFilesTab = new-graphTeamChannelTab -tokenResponse $tokenTeams -teamId $thisTeam.id -channelId $thisTeamGeneralTab.id -tabType SharePoint -tabName "Super Community Files" -tabDestinationUrl $superCommunity.DefaultDriveUrl

    #2 create links from the Community back to the Team
    $teamFolderInCommunity = get-graphDriveItems -tokenResponse $tokenTeams -driveGraphId $thisCommunity.DefaultDriveId -returnWhat Children | ? {$_.name -eq "$($thisTeam.displayName) files"} #Get or create the Team-in-Community folder
#    if($teamFolderInCommunity.id -eq $null){$teamFolderInCommunity = add-graphFolderToDrive -tokenResponse $tokenTeams -graphDriveId $thisCommunity.DefaultDriveId -folderName "$($thisTeam.displayName) files" -conflictResolution Fail}
    if($teamFolderInCommunity.id -ne $null){
        delete-graphDriveItem -tokenResponse $tokenTeams -graphDriveId $thisCommunity.DefaultDriveId -graphDriveItemId $teamFolderInCommunity.id
        }
    $teamFolderInCommunity = add-graphFolderToDrive -tokenResponse $tokenTeams -graphDriveId $thisCommunity.DefaultDriveId -folderName "$($thisTeam.displayName) files" -conflictResolution Fail

    $thisTeamDriveItems = add-graphArrayOfFoldersToDrive -tokenResponse $tokenTeams -graphDriveId $thisTeam.DefaultDriveId -foldersAndSubfoldersArray $solutionFoldersToCreate -conflictResolution Fail | ? {$solutionFoldersToCreate -contains $_.name} #Create any Channel folders that haven't been automatically provisioned
    $thisTeamDriveItems = get-graphDriveItems -tokenResponse $tokenTeams -driveGraphId $thisTeam.DefaultDriveId -returnWhat Children
    $thisTeamDriveItems | ? {$_.name -match "(Shared)"}  | % {
        $thisTeamDriveItems = $_
        $newHyperlinkContent = `
"[InternetShortcut]
URL=$($thisTeamDriveItems.webUrl)
"
        $newHyperlink = invoke-graphPut -tokenResponse $tokenTeams -graphQuery "/drives/$($thisCommunity.DefaultDriveId)/items/$($teamFolderInCommunity.id):/$([uri]::EscapeUriString($thisTeamDriveItems.name)).url:/content" -binaryFileStream $newHyperlinkContent
        }

    #3 Remove any retired Channels
    $thisTeamChannels = get-graphTeamChannels -tokenResponse $tokenTeams -teamId $thisTeam.id -channelType Public
    $thisTeamChannels | ? {$teamFoldersToCreate -contains $_.displayName} | % {
        $thisChannel = $_
        write-host "`tDeleting Channel [$($thisChannel.displayName)] from Team [$($thisTeam.displayName)]"
        delete-graphTeamChannel -tokenResponse $tokenTeams -teamId $thisTeam.id -channelId $thisChannel.id
        }

    }



$dummyTeam = get-graphGroups -tokenResponse $tokenTeams -filterUpn Solutions_Community_North_America_Aggregated_Functional_365@anthesisgroup.com -selectAllProperties
$dummyTeamDrive = get-graphDrives -tokenResponse $tokenTeams -groupGraphId $dummyTeam.id -returnOnlyDefaultDocumentsLibrary
$dummyChannels = get-graphTeamChannels -tokenResponse $tokenTeams -teamId $dummyTeam.id -channelType Both
$dummyChannelTabs = get-graphTeamChannelTabs -tokenResponse $tokenTeams -teamId $dummyTeam.id -channelId 19:ldRR5kJB6rfd7XaZ05NJN-vNDiJo0RBC01-Ga8tRogc1@thread.tacv2



new-graphTeamChannelTab -tokenResponse $tokenTeams -teamId $thisCommunity.id -channelId 19:2e9b0a2b45d44bb5ac49dd232255c4c5@thread.skype -tabType Website -tabName "TestWeb" -tabDestinationUrl $dummyTeamDrive.webUrl
new-graphTeamChannelTab -tokenResponse $tokenTeams -teamId $dummy.id -channelId 19:2e9b0a2b45d44bb5ac49dd232255c4c5@thread.skype -tabType DocumentLibrary -tabName "TestDocLib" -tabDestinationUrl $dummyTeamDrive.webUrl
new-graphTeamChannelTab -tokenResponse $tokenTeams -teamId $dummyTeam.id -channelId 19:ldRR5kJB6rfd7XaZ05NJN-vNDiJo0RBC01-Ga8tRogc1@thread.tacv2 -tabType SharePoint -tabName "Super Community Files" -tabDestinationUrl $dummyTeamDrive.webUrl



https://anthesisllc.sharepoint.com/sites/Resources-IT/_layouts/15/teamslogon.aspx?spfx=true&dest=https%3A%2F%2Fanthesisllc.sharepoint.com%2Fsites%2FResources-IT%2F_layouts%2F15%2Flistallitems.aspx%3Fapp%3DteamsPage%26listUrl%3D%2Fsites%2FResources-IT%2FShared%20Documents

[uri]::UnescapeDataString("https%3A%2F%2Fanthesisllc.sharepoint.com%2Fsites%2FResources-IT%2F_layouts%2F15%2Flistallitems.aspx%3Fapp%3DteamsPage%26listUrl%3D%2Fsites%2FResources-IT%2FShared%20Documents")