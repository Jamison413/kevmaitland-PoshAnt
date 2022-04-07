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
    Sustainable Products, Packaging, and Circularity Community (North America) Aggregated Function
    Sustainable Products, Packaging, and Circularity Community (North America)
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
$oldFolderNames += "0 ADMIN"
$oldFolderNames += "1 TEMPLATES"
$oldFolderNames += "2 PRESENTATIONS"
$oldFolderNames += "3 CASE STUDIES"
$oldFolderNames += "4 MEDIA"
$oldFolderNames += "5 REFERENCE"

$solutionFoldersToCreate = @()
$solutionFoldersToCreate += "0 ADMIN"
$solutionFoldersToCreate += "1 TEMPLATES (Shared)"
$solutionFoldersToCreate += "2 PRESENTATIONS (Shared)"
$solutionFoldersToCreate += "3 CASE STUDIES (Shared)"
$solutionFoldersToCreate += "4 MEDIA (Shared)"
$solutionFoldersToCreate += "5 REFERENCE (Shared)"

$teamFoldersToCreate = $solutionFoldersToCreate
$teamFoldersToCreate += "6 MEDIA (Shared)"
$teamFoldersToCreate += "7 REFERENCE (Shared)"

$listofTeams | % {
    $thisTeam = $_
    $thisTeamTeam = get-graphGroups -tokenResponse $tokenTeams -filterDisplayNameStartsWith $thisTeam -filterGroupType Unified
    #Temporarily add me
    Add-DistributionGroupMember -Identity $thisTeamTeam.anthesisgroup_UGSync.dataManagerGroupId -BypassSecurityGroupManagerCheck:$true -Member t0-kevin.maitland@anthesisgroup.com

    if([string]::IsNullOrWhiteSpace($thisTeamTeam.id)){write-host "Team [$thisTeam] not retrieved";return}
    else{Write-Host -f Yellow "Team [$($thisTeamTeam.displayName)] retrieved"}
    $thisTeamDrive = get-graphDrives -tokenResponse $tokenTeams -groupGraphId $thisTeamTeam.id
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
        grant-graphSharing -tokenResponse $tokenTeams -driveId $thisTeamDrive.id -itemId $thisFolder.id -sharingRecipientsUpns Solutions_Community_North_America_Aggregated_Functional_365@anthesisgroup.com -sendInvitation $false -role Read -requireSignIn $true
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
    $thisCommunityDrive = get-graphDrives -tokenResponse $tokenTeams -groupGraphId $thisCommunityTeam.id
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
    $thisTeamDrive = get-graphDrives -tokenResponse $tokenTeams -groupGraphId $thisTeam.id
    if([string]::IsNullOrWhiteSpace($thisTeamDrive.id)){write-host "Team Drive for [$($thisTeam.displayName)] not retrieved";return}
    $thisCommunityDrive = get-graphDrives -tokenResponse $tokenTeams -groupGraphId $thisCommunity.id
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