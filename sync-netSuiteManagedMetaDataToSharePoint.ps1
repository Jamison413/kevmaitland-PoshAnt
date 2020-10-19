if($PSCommandPath){
    $InformationPreference = 2
    $VerbosePreference = 0
    $logFileLocation = "C:\ScriptLogs\"
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))`_Transcript_$(Get-Date -Format "yyyy-MM-dd").log"
    Start-Transcript $transcriptLogName -Append
    }

function get-clientDrives(){
    #Get the Drives from Graph to compare against
    $sharePointBotDetails = get-graphAppClientCredentials -appName SharePointBot
    $tokenResponseSharePointBot = get-graphTokenResponse -aadAppCreds $sharePointBotDetails
    $clientSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,faed84bc-70be-4e35-bfbf-cdab31aeeb99"
    #$supplierSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,9fb8ecd6-c87d-485d-a488-26fd18c62303"
    #$devSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,8ba7475f-dad0-4d16-bdf5-4f8787838809"
    $allClientDrives = get-graphDrives -tokenResponse $tokenResponseSharePointBot -siteGraphId $clientSiteId
    $allClientDrives | % {
        Add-Member -InputObject $_ -MemberType NoteProperty -Name Name2 -Value $(tidy-name $_.name) -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name DriveId -Value $($_.id) -Force
        }
    $allClientDrives
    }
function get-clientTerm(){
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
            [Microsoft.SharePoint.Client.Taxonomy.TermSetItem[]]$arrayOfAllClientTerms
        ,[parameter(Mandatory = $true)]
            [string]$netSuiteId
        ,[parameter(Mandatory = $true)]
            [ValidateSet(“ReturnAll”,”Oldest”)]
            [string]$duplicateBehaviour
        )
    
    $clientTerm = $arrayOfAllClientTerms | ? {$_.CustomProperties.NetSuiteId -eq $netSuiteId}
    if($clientTerm.Count -gt 1){
        switch ($duplicateBehaviour){
            "ReturnAll" {$clientTerm;return} #Just return the duplicates and exit early
            "Oldest"    {
                $clientTerm = $clientTerm | Sort-Object CreatedDate | Select-Object -First 1 #Select the oldest ClientTerm
                }
            }
        }
    $clientTerm
    }
function get-oppTerm(){
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
            [Microsoft.SharePoint.Client.Taxonomy.TermSetItem[]]$arrayOfAllOppTerms
        ,[parameter(Mandatory = $true)]
            [string]$netSuiteProjectId
        ,[parameter(Mandatory = $true)]
            [ValidateSet(“ReturnAll”,”Oldest”)]
            [string]$duplicateBehaviour
        )
    
    $oppTerm = $arrayOfAllOppTerms | ? {$_.CustomProperties.NetSuiteProjectId -eq $netSuiteProjectId}
    if($oppTerm.Count -gt 1){
        switch ($duplicateBehaviour){
            "ReturnAll" {$oppTerm;return} #Just return the duplicates and exit early
            "Oldest"    {
                $oppTerm = $oppTerm | Sort-Object CreatedDate | Select-Object -First 1 #Select the oldest ClientTerm
                }
            }
        }
    $oppTerm
    }
function merge-folders(){
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [psobject]$sourceDriveItem
        ,[parameter(Mandatory = $true)]
            [Microsoft.SharePoint.Client.Taxonomy.TermSetItem]$oppProjTerm
        ,[parameter(Mandatory = $true)]
            [Microsoft.SharePoint.Client.Taxonomy.TermSetItem]$sourceClientTerm
        ,[parameter(Mandatory = $false)]
            [Microsoft.SharePoint.Client.Taxonomy.TermSetItem]$destinationClientTerm
        ,[parameter(Mandatory = $false)]
            [switch]$updateOppProjTerm
        )
    
    if([string]::IsNullOrWhiteSpace($destinationClientTerm)){$destinationClientTerm = $sourceClientTerm} #If we're merging folders within the same Drive, use the same DriveID

    try{
        $destinationDriveItem = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $destinationClientTerm.CustomProperties.GraphDriveId -returnWhat Children | ? {$_.name -eq $oppProjTerm.Name}
        }
    catch{
        Write-Error "Error getting destination Folder after collision updating driveItem [$($sourceDriveItem.Name)][$($sourceDriveItem.Id)][$($sourceDriveItem.webUrl)] name to [$($oppProjTerm.Name)] | Retrying with -Verbose"
        get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $destinationClientTerm.CustomProperties.GraphDriveId -returnWhat Children -Verbose
        return
        }

    #Pick which Folder to keep. Prefer nonempty folders over empty, then correctly-named over incorrectly-named. Never delete nonempty Folders - e-mail a human to figure it out.
    if($sourceDriveItem.size -eq 0){
        $keptDriveItem = $destinationDriveItem
        $deleteThisDriveItem = $sourceDriveItem
        $deleteThisDriveItemFromDriveId = $sourceClientTerm.CustomProperties.GraphDriveId
        $deletedFolderFriendlyName = "Source"
        }
    elseif($destinationDriveItem.fileSystemInfo.createdDateTime -eq $destinationDriveItem.fileSystemInfo.lastModifiedDateTime){
        $keptDriveItem = $sourceDriveItem
        $deleteThisDriveItem = $destinationDriveItem
        $deleteThisDriveItemFromDriveId = $destinationClientTerm.CustomProperties.GraphDriveId
        $deletedFolderFriendlyName = "Destination"
        }

    if($deleteThisDriveItem){
        Write-Warning "Deleting *EMPTY* $deletedFolderFriendlyName Folder [$($deleteThisDriveItem.Name)][$($deleteThisDriveItem.id)][$($deleteThisDriveItem.webUrl)] instead of merging - a better version already exists"
        try{
            $muteMe = delete-graphDriveItem -tokenResponse $tokenResponseSharePointBot -graphDriveId $deleteThisDriveItemFromDriveId -graphDriveItemId $deleteThisDriveItem.id 
            }
        catch{
            Write-Error "Error deleting driveItem [$($clientTerm.CustomProperties.GraphDriveId)][$($driveFolder.id)] | Retrying with -Verbose"
            delete-graphDriveItem -tokenResponse $tokenResponseSharePointBot -graphDriveId $deleteThisDriveItemFromDriveId -graphDriveItemId $deleteThisDriveItem.id  -Verbose
            }

        if($updateOppProjTerm){
            if($oppProjTerm.CustomProperties.DriveItemId -ne $keptDriveItem.id -and ![string]::IsNullOrWhiteSpace($keptDriveItem.id)){#If we've changed the DriveItemId and it's not $null, update the Term now
                $oppProjTerm.SetCustomProperty("DriveItemId",$keptDriveItem.id)
                $oppProjTerm.Context.ExecuteQuery()
                }
            }

        $keptDriveItem #Return $keptDriveItem so we know which object still exists


        }
    else{
        #This scenario is too copmlicated to handle automatically - just e-mail some humans to fix it
        Write-Warning "`t`tCannot perform a simple merge on [$($sourceDriveItem.Name)][$($sourceDriveItem.id)][$($sourceDriveItem.webUrl)] into [$($destinationDriveItem.Name)][$($destinationDriveItem.id)][$($destinationDriveItem.webUrl)]. This is too complicated for me - I'm going to e-mail the SharePoint Admins"
        $tokenResponseTeamsBot = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName TeamsBot) -grant_type client_credentials
        Send-MailMessage -To $(get-graphAdministrativeRoleMembers -tokenResponse $tokenResponseTeamsBot -roleName 'SharePoint Service Administrator') -From netsuitebot@anthesisgroup.com -Subject "[NetSuite/SharePoint integration] Opp/Proj folders cannot be moved automatically" -BodyAsHtml "<BODY>Hi SharePoint Admins,<BR><BR>These Opp/Proj Folders: <BR>&emsp;[$($sourceClientTerm.Name)][$($sourceClientTerm.CustomProperties.GraphDriveId)]<BR>&emsp;[$($oppProjTerm.Name)][$($oppProjTerm.CustomProperties.DriveItemId)][$($sourceDriveItem.webUrl)]<BR><BR>couldn't be automatically moved here:<BR>&emsp;[$($destinationClientTerm.Name)][$($destinationClientTerm.CustomProperties.GraphDriveId)][$($destinationDriveItem.webUrl)]<BR><BR>They need a human to manually merge them - could one of you give me a hand please? :)<BR><BR>Love,<BR><BR>The NetSuiteSyncBot" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com"        
        }
        
    }
function process-folder(){
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [Microsoft.SharePoint.Client.Taxonomy.TermSetItem]$oppProjTerm
        ,[parameter(Mandatory = $true)]
            [Microsoft.SharePoint.Client.Taxonomy.TermSetItem[]]$arrayOfAllClientTerms
        ,[parameter(Mandatory = $true)]
            [Microsoft.SharePoint.Client.Taxonomy.TermSetItem[]]$arrayOfAllOppTerms
        ,[parameter(Mandatory = $true)]
            [String[]]$arrayOfLeadProjSubFolders
        
        )

    #Get Folder
        #Update Folder
        #Create Folder
    switch($oppProjTerm.Name){
        {$_.StartsWith("O-")} {$thisIsA = "Opp"}
        {$_.StartsWith("P-")} {$thisIsA = "Proj"}
        default               {$thisIsA = "Unknown"}
        }
    
    $flagForReprocessing = $true #By default, we wanbt to keep reprocessing Opps/Proj Terms that fail

    #Get the Client Term and do some preliminary error-checking
    $clientTerm = get-clientTerm -arrayOfAllClientTerms $arrayOfAllClientTerms -netSuiteId $oppProjTerm.CustomProperties.NetSuiteClientId -duplicateBehaviour Oldest
    if([string]::IsNullOrWhiteSpace($clientTerm)){
        Write-Error "Error: Client Term cannot be found for $thisIsA [$($oppProjTerm.Name)][$($oppProjTerm.Id)]. Folders cannot be processed for this $thisIsA"
        return #Exit early - there's nothing else to do and return nothing to show it failed
        }
    if([string]::IsNullOrWhiteSpace($clientTerm.CustomProperties.GraphDriveId)){
        Write-Error "Error: Client Term [$($clientTerm.Name)][$($clientTerm.Id)] has no DriveId for $thisIsA [$($oppProjTerm.Name)][$($oppProjTerm.Id)]. Folders cannot be processed for this $thisIsA"
        return #Exit early - there's nothing else to do and return nothing to show it failed
        }
    
    #region Get Folder
    if([string]::IsNullOrWhiteSpace($oppProjTerm.CustomProperties.DriveItemId)){
        try{
            $driveFolder = search-folder -tokenResponse $tokenResponseSharePointBot -oppProjTerm $oppProjTerm -clientTerm $clientTerm -arrayOfAllClientTerms $arrayOfAllClientTerms -updateTermIfFound -moveFolderIfFound
            }
        catch{
            Write-Error "Error searching for $($thisIsA)Folder (no ID available) [$($oppProjTerm.Name)][$($oppProjTerm.Id)] by Name  | Retrying with -Verbose"
            search-folder -tokenResponse $tokenResponseSharePointBot -oppProjTerm $oppProjTerm -clientTerm $clientTerm -arrayOfAllClientTerms $arrayOfAllClientTerms -updateTermIfFound -moveFolderIfFound -Verbose
            }
        }
    else{
        try{
            $driveFolder = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $clientTerm.CustomProperties.GraphDriveId -itemGraphId $oppProjTerm.CustomProperties.DriveItemId -returnWhat Item -ErrorAction stop
            }
        catch{
            if($_.Exception -match "(404)"){ #If the folder is missing, look through any previous client drives the Opp/Proj was associated with
                Write-Host "`t$($thisIsA)Folder for $($thisIsA) [$($oppProjTerm.Name)][$($oppProjTerm.Id)] is missing from Client [$($clientTerm.Name)][$($clientTerm.CustomProperties.GraphDriveId)]."
                try{
                    $driveFolder = search-folder -tokenResponse $tokenResponseSharePointBot -oppProjTerm $oppProjTerm -clientTerm $clientTerm -arrayOfAllClientTerms $arrayOfAllClientTerms -updateTermIfFound -moveFolderIfFound
                    }
                catch{
                    Write-Error "Error searching for $($thisIsA)Folder [$($oppProjTerm.CustomProperties.DriveItemId)] [$($oppProjTerm.Name)][$($oppProjTerm.Id)] by Name | Retrying with -Verbose"
                    search-folder -tokenResponse $tokenResponseSharePointBot -oppProjTerm $oppProjTerm -clientTerm $clientTerm -arrayOfAllClientTerms $arrayOfAllClientTerms -updateTermIfFound -moveFolderIfFound -Verbose
                    }
                }
            else{
                Write-Error "Error retrieving $($thisIsA)Folder for $($thisIsA) [$($oppProjTerm.Name)][$($oppProjTerm.id)][$($oppProjTerm.CustomProperties.DriveItemId)] for for Client [$($clientTerm.Name)][$($clientTerm.CustomProperties.GraphDriveId)] |  Retrying with Verbose"
                get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $clientTerm.CustomProperties.GraphDriveId -itemGraphId $oppProjTerm.CustomProperties.DriveItemId -returnWhat Item -Verbose
                }
            }
        }
    #endregion

    #region Update Folder
        #Update Name (if changed)
        #Reassign to new Client (if changed)
        #Notify if neither
    if($driveFolder){
        #Notify if neither - do this first, otherwise the Names/ClientIds will always match (because we'll have fixed them) :)
        if(($oppProjTerm.Name -eq $driveFolder.name) -and ($driveFolder.parentReference.driveId -eq $clientTerm.CustomProperties.GraphDriveId)){
            Write-Host "`tNothing significant has changed for $($thisIsA) Term [$($oppProjTerm.Name)][$($oppProjTerm.id)][$($oppProjTerm.CustomProperties.DriveItemId)]"
            $flagForReprocessing = $false #If this worked, mark the Term as clean
            }
        #region Update Name (if changed)
        if($oppProjTerm.Name -ne $driveFolder.name){#Fix the name if it has changed
            Write-Host "`t$($thisIsA)Folder Name [$($driveFolder.name)] is out-of-date - updating $($thisIsA)Folder name to [$($oppProjTerm.Name)]"
            try{
                $updatedFolder = set-graphDriveItem -tokenResponse $tokenResponseSharePointBot -driveId $clientTerm.CustomProperties.GraphDriveId -driveItemId $oppProjTerm.CustomProperties.DriveItemId -driveItemPropertyHash @{name=$oppProjTerm.Name} -ErrorAction Stop
                if($updatedFolder){
                    $driveFolder = $updatedFolder
                    $flagForReprocessing = $false #If this worked, mark the Term as clean
                    }
                else{Write-Host "`t`test-graphDriveItem didn't return the updated Folder, but didn't produce an error either :/"}
                }
            catch{
                if($_.Exception -match "(409)"){ #Folder already exists
                    Write-Host "`t`tA different $($thisIsA)Folder with the name [$($oppProjTerm.Name)] already exists. Attempting simple Merge."
                    try{
                        $updatedFolder = merge-folders -tokenResponse $tokenResponseSharePointBot -sourceDriveItem $driveFolder -oppProjTerm $oppProjTerm -sourceClientTerm $clientTerm -updateOppProjTerm
                        if($updatedFolder){
                            $driveFolder = $updatedFolder
                            $flagForReprocessing = $false #If this worked, mark the Term as clean
                            }
                        else{Write-Host "`t`tmerge-folders didn't return the updated Folder, but didn't produce an error either :/"}
                        }
                    catch{
                        Write-Error "Error merging Folders follwing a Name collision while updating Name for $($thisIsA) [$($oppProjTerm.Name)][$($oppProjTerm.id)][$($oppProjTerm.CustomProperties.DriveItemId)] for Client [$($clientTerm.Name)][$($clientTerm.CustomProperties.GraphDriveId)]. driveItem being renamed was [$($driveFolder.Name)][$($driveFolder.id)][$($driveFolder.webUrl)] | Retrying with -Verbose"
                        merge-folders -tokenResponse $tokenResponseSharePointBot -sourceDriveItem $driveFolder -oppProjTerm $oppProjTerm -sourceClientTerm $clientTerm -updateOppProjTerm -Verbose
                        }
                    }
                else{
                    Write-Error "Error updating Name of $($thisIsA)Folder [$($driveFolder.name)][$($driveFolder.webUrl)] for $thisIsA [$($oppProjTerm.Name)][$($oppProjTerm.CustomProperties.DriveItemId)] for Client [$($clientTerm.Name)][$($clientTerm.CustomProperties.GraphDriveId)] | Retrying with Verbose"
                    set-graphDriveItem -tokenResponse $tokenResponseSharePointBot -driveId $clientTerm.CustomProperties.GraphDriveId -driveItemId $oppProjTerm.CustomProperties.DriveItemId -driveItemPropertyHash @{name=$oppProjTerm.Name} -Verbose
                    Return #Exit early - we don't want *more* folders being created if we've already got name collisions.
                    }
                }
            }
            #endregion
        #region Reassign to new Client (if changed)
        if($driveFolder.parentReference.driveId -ne $clientTerm.CustomProperties.GraphDriveId){ #If the Opp has been reassigned to another client, move the OppFolders to the new location
            Write-Host "`t$($thisIsA) [$($oppProjTerm.Name)] has been reassigned from ClientDrive [$($driveFolder.parentReference.driveId)] to [$($clientTerm.Name)][$($clientTerm.CustomProperties.GraphDriveId)] - attempting to move $($thisIsA)Folders"
            try{
                $clientDriveRoot = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $clientTerm.CustomProperties.GraphDriveId -returnWhat Item #the Graph API behind move-driveItem explicitly requires the ID for the Root folder (we can't just use /root:)
                }
            catch{
                Write-Error "Error retrieving Root Folder for Client Drive [$($clientTerm.Name)][$($clientTerm.CustomProperties.GraphDriveId)] | Retrying with Verbose"
                get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $clientTerm.CustomProperties.GraphDriveId -returnWhat Item -Verbose
                }

            if($clientDriveRoot){ #If we've got enough information to attempt a move-driveItem
                try{
                    $movedFolder = move-driveItem -tokenResponse $tokenResponseSharePointBot -driveGraphIdSource $driveFolder.parentReference.driveId -itemGraphIdSource $driveFolder.id -driveGraphIdDestination $clientTerm.CustomProperties.GraphDriveId -parentItemGraphIdDestination $clientDriveRoot.id -newItemName $oppProjTerm.Name
                    if($movedFolder){
                        $driveFolder = $movedFolder
                        $flagForReprocessing = $false #If this worked, mark the Term as clean
                        }
                    else{Write-Host "`t`tmove-driveItem didn't return the new Folder, but didn't produce an error either "}
                    }
                catch{
                    if($_.Exception -match "(409)"){ #Folder already exists
                        Write-Host "`t`tA different $($thisIsA)Folder with the name [$($oppProjTerm.Name)] already exists in the new Client Drive [$($clientTerm.Name)][$($clientTerm.Id)][$($clientDriveRoot.webUrl)]. Attempting simple Merge."
                        $previousClientTerm = get-clientTerm -arrayOfAllClientTerms $arrayOfAllClientTerms -netSuiteId $($arrayOfAllClientTerms | ? {$_.CustomProperties.GraphDriveId -eq $driveFolder.parentReference.driveId}).CustomProperties.NetSuiteId -duplicateBehaviour Oldest
                        try{
                            $movedFolder = merge-folders -tokenResponse $tokenResponseSharePointBot -sourceDriveItem $driveFolder -oppProjTerm $oppProjTerm -sourceClientTerm $previousClientTerm -destinationClientTerm $clientTerm -updateOppProjTerm
                            if($movedFolder){
                                $driveFolder = $movedFolder
                                $flagForReprocessing = $false #If this worked, mark the Term as clean
                                }
                            else{Write-Host "`t`tmerge-folder didn't return the new Folder, but didn't produce an error either "}
                            }
                        catch{
                            Write-Error "Error merging Folders follwing a Name collision while reassigning $thisIsA [$($oppProjTerm.Name)][$($oppProjTerm.id)][$($oppProjTerm.CustomProperties.DriveItemId)] to a new Client [$($clientTerm.Name)][$($clientTerm.CustomProperties.GraphDriveId)] from [$($previousClientTerm.Name)][$($previousClientTerm.CustomProperties.GraphDriveId)]. driveItem being moved was [$($driveFolder.Name)][$($driveFolder.id)][$($driveFolder.webUrl)] | Retrying with -Verbose"
                            merge-folders -tokenResponse $tokenResponseSharePointBot -sourceDriveItem $driveFolder -oppProjTerm $oppProjTerm -sourceClientTerm $previousClientTerm -destinationClientTerm $clientTerm -updateOppProjTerm -Verbose
                            }
                        }
                    else{
                        Write-Error "Error moving $($thisIsA)Folders [$($oppProjTerm.Name)][$($driveFolder.id)] to from old Client Drive [$($previousClientTerm.Name)][$($previousClientTerm.CustomProperties.GraphDriveId)] to new Client Drive [$($clientTerm.Name)][$($clientTerm.CustomProperties.GraphDriveId)] | Retrying with Verbose"
                        move-driveItem -tokenResponse $tokenResponseSharePointBot -driveGraphIdSource $driveFolder.parentReference.driveId -itemGraphIdSource $driveFolder.id -driveGraphIdDestination $clientTerm.CustomProperties.GraphDriveId -parentItemGraphIdDestination $clientDriveRoot.id -newItemName $oppProjTerm.Name -Verbose
                        }
                    }
                if($oppProjTerm.CustomProperties.DriveItemId -ne $movedFolder.id -and ![string]::IsNullOrWhiteSpace($movedFolder.id)){#If we've changed the DriveItemId, update the Term now
                    $oppProjTerm.SetCustomProperty("DriveItemId",$movedFolder.id)
                    $oppProjTerm.Context.ExecuteQuery()
                    }
                }
            }
            #endregion
        }
    #endregion

    #region Create Folder
    else{
        #Extra step for Projects as we might need to rename the corresponding OppFolder
        if($thisIsA -eq "Proj"){
            Write-Host "`tSearching for corresponding OppFolder for Project [$($oppProjTerm.Name)][$($oppProjTerm.Id)]"
            $correspondingOppTerm = get-oppTerm -arrayOfAllOppTerms $arrayOfAllOppTerms -netSuiteProjectId $oppProjTerm.CustomProperties.NetSuiteProjId -duplicateBehaviour Oldest
            try{
                $oppFolder = search-folder -tokenResponse $tokenResponseSharePointBot -oppProjTerm $correspondingOppTerm -clientTerm $clientTerm -arrayOfAllClientTerms $arrayOfAllClientTerms #-updateTermIfFound -moveFolderIfFound
                }
            catch{
                Write-Error "Error searching for corresponding OppFolder for Proj [$($oppProjTerm.Name)][$($oppProjTerm.Id)] | Retrying with -Verbose"
                search-folder -tokenResponse $tokenResponseSharePointBot -oppProjTerm $correspondingOppTerm -clientTerm $clientTerm -arrayOfAllClientTerms $arrayOfAllClientTerms -Verbose 
                }
            if($oppFolder){
                try{
                    Write-Host "`t`tUpdating DriveItemId to [$($correspondingOppTerm.CustomProperties.DriveItemId)]"
                    $oppProjTerm.SetCustomProperty("DriveItemId",$oppFolder.id)
                    $oppProjTerm.Context.ExecuteQuery()
                    $flagForReprocessing = $true #Even if this worked, mark the Term as dirty so the Folder name gets updated on the next iteration
                    }
                catch{
                    Write-Error "Error updating DriveItemId for Proj [$($oppProjTerm.Name)][$($oppProjTerm.Id)] to the same value [$($correspondingOppTerm.CustomProperties.DriveItemId)] as the Opp [$($correspondingOppTerm.Name)][$($correspondingOppTerm.Id)]"
                    }
                Return #If we found a corresponding OppFolder, we don't want to create a new ProjFolder, so jump out of the current iteration of the loop
                }
            }

        #If we didn't find the Opp Folder anywhere, we want to create new folders in the current Client Drive
        Write-Host "`tCreating new $($thisIsA)Folders for [$($oppProjTerm.Name)]"
        [array]$customisedFolderList = $oppProjTerm.Name
        $customisedFolderList += $arrayOfLeadProjSubFolders | % {"$($oppProjTerm.Name)\$_"}
        try{
            $newFolders = add-graphArrayOfFoldersToDrive -graphDriveId $clientTerm.CustomProperties.GraphDriveId -foldersAndSubfoldersArray $customisedFolderList -tokenResponse $tokenResponseSharePointBot -conflictResolution Fail #-ErrorAction SilentlyContinue
            if($newFolders){
                $oppProjTerm.SetCustomProperty("DriveItemId",$newFolders[1].id)
                $oppProjTerm.Context.ExecuteQuery()
                $flagForReprocessing = $false #If this worked, mark the Term as clean
                }
            else{Write-Host "`t`tadd-graphArrayOfFoldersToDrive didn't return the new Folders, but didn't produce an error either :/"}
            }
        catch{
            Write-Error "Error creating $($thisIsA)Folders for [$($oppProjTerm.Name)][$($oppProjTerm.id)] for Client [$($clientTerm.Name)][$($clientTerm.CustomProperties.GraphDriveId)] | Retrying with Verbose"
            add-graphArrayOfFoldersToDrive -graphDriveId $clientTerm.CustomProperties.GraphDriveId -foldersAndSubfoldersArray $customisedFolderList -tokenResponse $tokenResponseSharePointBot -conflictResolution Fail -Verbose
            }
        }
    #endregion

    try{
        if($flagForReprocessing){Write-Host "`tSomething didn't work with [$($oppProjTerm.Name)][$($oppProjTerm.id)] for Client [$($clientTerm.Name)][$($clientTerm.CustomProperties.GraphDriveId)] - Flagging for Reprocessing"}
        else{Write-Host "`tSUCCESS: Setting `$flagForReprocessing = $flagForReprocessing"}
        $oppProjTerm.SetCustomProperty("flagForReprocessing",$flagForReprocessing)
        $oppProjTerm.Context.ExecuteQuery()
        }
    catch{
        Write-Error "Error setting `$flagForReprocessing = $flagForReprocessing :("
        }

    }
function search-folder(){
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [Microsoft.SharePoint.Client.Taxonomy.TermSetItem]$oppProjTerm
        ,[parameter(Mandatory = $false)]
            [Microsoft.SharePoint.Client.Taxonomy.TermSetItem]$clientTerm
        ,[parameter(Mandatory = $true)]
            [Microsoft.SharePoint.Client.Taxonomy.TermSetItem[]]$arrayOfAllClientTerms
        ,[parameter(Mandatory = $false)]
            [Switch]$updateTermIfFound
        ,[parameter(Mandatory = $false)]
            [Switch]$moveFolderIfFound
        )
    switch($oppProjTerm.Name){
        {$_.StartsWith("O-")} {$thisIsA = "Opp"}
        {$_.StartsWith("P-")} {$thisIsA = "Proj"}
        default               {$thisIsA = "Unknown"}
        }

    if([string]::IsNullOrWhiteSpace($clientTerm)){
        $clientTerm = get-clientTerm -arrayOfAllClientTerms $arrayOfAllClientTerms -netSuiteId $oppProjTerm.CustomProperties.NetSuiteClientId -duplicateBehaviour Oldest
        }
    
    Write-Host "`tLooking for folder with name [$($oppProjTerm.Name)] for $thisIsA [$($oppProjTerm.Name)][$($oppProjTerm.Id)] from Client [$($clientTerm.Name)][$($clientTerm.CustomProperties.GraphDriveId)]."
    try{#First check the Drive for an item with the same Name and re-use that if possible
        $driveFolder = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $clientTerm.CustomProperties.GraphDriveId -returnWhat Children -ErrorAction stop | ? {$(tidy-name $_.name) -eq $(tidy-name $oppProjTerm.Name)}
        if($driveFolder.Count -gt 1){#If there are multiple matches, use the oldest one
            Write-Warning "`Multiple $($thisIsA)Folder matches for [$($oppProjTerm.Name)][$($oppProjTerm.Id)] - selecting oldest"
            $driveFolder = $driveFolder | Sort-Object createdDateTime | Select-Object -First 1
            } 
        if($driveFolder){
            Write-Host "`tDifferent $($thisIsA)Folder with same name found"
            if($updateTermIfFound){
                Write-Host "`t`tUpdating $($thisIsA)Term.CustomProperties.DriveItemId from [$($oppProjTerm.CustomProperties.DriveItemId)] to [$($driveFolder.id)]"
                try{
                    $oppProjTerm.SetCustomProperty("DriveItemId",$driveFolder.id)
                    $oppProjTerm.Context.ExecuteQuery()
                    }
                catch{
                    Write-Error "Error Updating $($thisIsA)Term.CustomProperties.DriveItemId for $thisIsA [$($oppProjTerm.Name)][$($oppProjTerm.Id)] from Client [$($clientTerm.Name)][$($clientTerm.CustomProperties.GraphDriveId)] in search-folder()"
                    }
                }
            }
        }
    catch{
        if($_.Exception -match "(404)"){ #If the folder is missing, look through any previous client drives the Opp was associated with
            $newClientTerm = $clientTerm #Get this now in case we need to move the OppFolder to the current Client later
            if(![string]::IsNullOrWhiteSpace($oppProjTerm.CustomProperties.NetSuiteClientId_previous)){ #If the Opp has previously been reassigned...
                while(![string]::IsNullOrWhiteSpace($($oppProjTerm.CustomProperties."NetSuiteClientId_previous$($i+1)"))){$i++} #Check NetSuiteClientId_previous$i and increment $i until we find a null/empty property (this will give us the highest number for NetSuiteClientId_previous)
                do{#Work from $i back down to zero through the list of NetSuiteClientId_previous$i as higher values of $i are more likely to be correct
                    $clientTerm = get-clientTerm -arrayOfAllClientTerms $arrayOfAllClientTerms -netSuiteId $oppProjTerm.CustomProperties."NetSuiteClientId_previous$i" -duplicateBehaviour Oldest
                    if([string]::IsNullOrWhiteSpace($clientTerm.CustomProperties.GraphDriveId)){
                        Write-Error "Client [$($clientTerm.Name)][$($oppProjTerm.CustomProperties."NetSuiteClientId_previous$i")] does not have a GraphDriveId - $($thisIsA)s cannot be processed for this client."
                        Return
                        }
                    try{
                        $driveFolder = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $clientTerm.CustomProperties.GraphDriveId -itemGraphId $oppProjTerm.CustomProperties.DriveItemId -returnWhat Item -ErrorAction SilentlyContinue
                        }
                    catch{
                        if($_.Exception -match "(404)"){#There's a good chance we'll get 404 errors here as we're not expecting the OppFolder to be in the wrong place often. -ErrorAction SilentlyContinue doesn;t suppress the errors though, so we just catch and drop
                            }
                        else{#If something weird has gone wrong, we want to know about it though.
                            Write-Error "Error getting $($thisIsA)Folder for [$($oppProjTerm.Name)][$($oppProjTerm.Id)] for previous client [$($clientTerm.Name)][$($clientTerm.Id)] in search-folder() | Retrying with -Verbose"
                            get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $clientTerm.CustomProperties.GraphDriveId -itemGraphId $oppProjTerm.CustomProperties.DriveItemId -returnWhat Item -Verbose
                            }
                        } 
                    $i-- 
                    }
                while($driveFolder -eq $null -and $i -gt 0)
                }               
            else{
                #Client has never been reassigned
                }
            }
        else{#If something weird has gone wrong, we want to know about it though.
            Write-Error "Error getting $($thisIsA)Folder for [$($oppProjTerm.Name)][$($oppProjTerm.Id)] for primary client [$($clientTerm.Name)][$($clientTerm.Id)] in search-folder()"
            $_
            }
        } 

    $driveFolder #If we've got one, return the $driveFolder Object
    }
function tidy-name($string){
    $string.Replace("&","").Replace("＆","").Replace("  "," ")
    }

$listOfClientFolders = @("_NetSuite automatically creates Opportunity & Project folders","Background","Non-specific BusDev")
$listOfLeadProjSubFolders = @("Admin & contracts", "Analysis","Data & refs","Meetings","Proposal","Reports","Summary (marketing) - end of project")
$now = Get-Date

$sharePointBotDetails = get-graphAppClientCredentials -appName SharePointBot 
$tokenResponseSharePointBot = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName SharePointBot )
$clientSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,faed84bc-70be-4e35-bfbf-cdab31aeeb99"
#$supplierSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,9fb8ecd6-c87d-485d-a488-26fd18c62303"
#$devSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,8ba7475f-dad0-4d16-bdf5-4f8787838809"

$sharePointAdmin = "kimblebot@anthesisgroup.com"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content "$env:USERPROFILE\Desktop\KimbleBot.txt") 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com" -Credentials $adminCreds

Write-Host "sync-netsuiteManagedMetaDataToSharePoint started at [$(Get-Date -Format s)]"
$fullSyncTime = Measure-Command {
    [datetime]$lastSpoSyncRun = $(Get-PnPTerm -TermGroup "Anthesis" -TermSet "IT" -Identity "LastModified" -Includes CustomProperties).CustomProperties.LastSpoSyncRun
    #region Clients
    $pnpTermGroup = "Kimble"
    $pnpTermSet = "Clients"
    $allClientTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes CustomProperties | ? {$_.IsDeprecated -eq $false}

    #Filter these client-side (CSOM, eh?) to get only the changes since this script last completed successfully
    [array]$clientTermsToCheck = $allClientTerms | ? {($_.LastModifiedDate -gt $lastSpoSyncRun -or $_.CustomProperties.flagForReprocessing -eq $true) -and ![string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteId)}
    Write-Host "Processing [$($clientTermsToCheck.Count)] Clients"
    if($clientTermsToCheck){
        $clientTermsToCheck | % {
            Add-Member -InputObject $_ -MemberType NoteProperty -Name Name2 -Value $($_.Name).Replace("&","").Replace("＆","").Replace("  "," ") -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name DriveId -Value $($_.CustomProperties.GraphDriveId) -Force
            }
        if(!$allClientDrives){$allClientDrives = get-clientDrives} #Only load the Client Drives via graph if there is work to do and we haven't already got them

        #############################
        #Create new Prospects/Clients
        #############################
        $missingFromSpo = Compare-Object -ReferenceObject @($clientTermsToCheck | Sort-Object DriveId | Select-Object) -DifferenceObject @($allClientDrives | Sort-Object DriveId | Select-Object) -Property DriveId -PassThru -IncludeEqual | ? {$_.SideIndicator -eq "<="}
        Write-Host "Creating [$($missingFromSpo.Count)] new Client DocLibs"
        $missingFromSpo | Select-Object | % {
            $thisNewClient = $_
            Write-Host "Creating new GraphList [$($thisNewClient.Name)]"
            try{
                $newGraphList = new-graphList -tokenResponse $tokenResponseSharePointBot -siteGraphId $clientSiteId -listDisplayName $thisNewClient.Name -listType documentLibrary
                } #Graph doesn't support creating Drives, so we need to create a List
            catch{
                if($_.Exception -match "(409)"){ #Folder already exists
                    Write-Warning "`tClient DocLib for [$($thisNewClient.Name)] already exists!"
                    }
                else{
                    Write-Error "Error creating new DocLib for Client [$($thisNewClient.Name)] - retrying with Verbose"
                    $oldVerbosePreference = $VerbosePreference
                    $VerbosePreference = 2
                    new-graphList -tokenResponse $tokenResponseSharePointBot -siteGraphId $clientSiteId -listDisplayName $thisNewClient.Name -listType documentLibrary -Verbose
                    $VerbosePreference = $oldVerbosePreference
                    }
                }
            if(!$newGraphList){$newGraphList = get-graphList -tokenResponse $tokenResponseSharePointBot -graphSiteId $clientSiteId -listName $thisNewClient.Name}
            if(!$newGraphList){
                Write-Error "Could not retrieve Graph List for Client [$($thisNewClient.Name)] - not checking for Drive"
                continue
                }
            else{
                Write-Host "`tGetting new GraphDrive from GraphList [$($newGraphList.name)][$($newGraphList.id)]"
                $newGraphListDrive = get-graphDrives -tokenResponse $tokenResponseSharePointBot -siteGraphId $clientSiteId -listGraphId $newGraphList.id #Then get the new Drive object form the List
                Write-Host "`tCreating standard Client folders in Drive [$($newGraphListDrive.name)][$($newGraphListDrive.id)]"
                $newFolders = add-graphArrayOfFoldersToDrive -graphDriveId $newGraphListDrive.id -foldersAndSubfoldersArray $listOfClientFolders -tokenResponse $tokenResponseSharePointBot -conflictResolution Fail
                Write-Host "`tUpdating Term [$($thisNewClient.Name)] with CustomProperties @{DocLibId=$($newGraphList.id);GraphDriveId=$($newGraphListDrive.id)}"
                $thisNewClient.SetCustomProperty("DocLibId",$newGraphList.id)
                $thisNewClient.SetCustomProperty("GraphDriveId",$newGraphListDrive.id)
                }
            try{
                Write-Verbose "`tTrying to update Term [$($thisNewClient.Name)] with CustomProperties @{DocLibId=$($newGraphList.id);GraphDriveId=$($newGraphListDrive.id)}"
                $thisNewClient.Context.ExecuteQuery()
                $thisNewClient.SetCustomProperty("flagForReprocessing",$false) #If the previous ExecuteQuery() worked, deflag the Term so it doesn;t get processed next time
                $thisNewClient.Context.ExecuteQuery()
                }
            catch{
                Write-Error "Error updating Term [$($thisNewClient.Name)] with CustomProperties @{DocLibId=$($newGraphList.id);GraphDriveId=$($newGraphListDrive.id)} in sync-netsuiteManagedMetaDataToSharePoint()"
                [array]$flagForReprocessing += $thisNewClient
                }
            }


        #############################
        #Update any Clients Drives that have changed their names in NetSuite
        #############################
        $matchedGraphId = Compare-Object -ReferenceObject $clientTermsToCheck -DifferenceObject $allClientDrives -Property DriveId -IncludeEqual -ExcludeDifferent -PassThru #We find out which $clientTermsToCheck records already have valid GraphDriveId values
        $matchedGraphIdReversed = Compare-Object -ReferenceObject $allClientDrives -DifferenceObject @($matchedGraphId | Select-Object) -Property DriveId -IncludeEqual -ExcludeDifferent -PassThru #We then use $matchedGraphId to filter only the Drive objects with corresponding $clientTermsToCheck records
        $deltaName = Compare-Object -ReferenceObject @($matchedGraphId | Select-Object) -DifferenceObject @($matchedGraphIdReversed | Select-Object) -Property DriveId,Name2 -PassThru -IncludeEqual #We compare the two equal sets on both DriveId and Name2 to see which pairs have mismatched Name values
        $clientsWithChangedNames = $deltaName | ? {$_.SideIndicator -eq "<="} #Anything on this side has a different Name2 in NetSuite
        Write-Host "Updating [$($clientsWithChangedNames.Count)] Client name changes"
        $clientsWithChangedNames | Select-Object | % {
            $thisUpdatedClient = $_
            Write-Host "Company name [$($thisUpdatedClient.Name)][$($thisUpdatedClient.id)] seems to have changed. Investigating further."
            if([string]::IsNullOrWhiteSpace($thisUpdatedClient.CustomProperties.DocLibId)){ #If it's missing it's DocLibID, try to fix it
                Write-Host "[$($thisUpdatedClient.Name)][$($thisUpdatedClient.id)] is missing its .CustomProperties.DocLibId value - attempting repair"
                $graphList = get-graphList -tokenResponse $tokenResponseSharePointBot -graphSiteId $clientSiteId -graphDriveId $thisUpdatedClient.CustomProperties.GraphDriveId
                if($graphList){
                    $thisUpdatedClient.SetCustomProperty("DocLibId",$graphList.id)
                    try{
                        Write-Verbose "`tTrying: [$($thisUpdatedClient.Name)].SetCustomProperty(`"DocLibId`",$($graphList.id))"
                        $thisUpdatedClient.Context.ExecuteQuery()
                        $thisUpdatedClient.SetCustomProperty("flagForReprocessing",$false) #If the previous ExecuteQuery() worked, deflag the Term so it doesn;t get processed next time
                        $thisUpdatedClient.Context.ExecuteQuery()
                        $thisUpdatedClient = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Identity $thisUpdatedClient.Id -Includes CustomProperties
                        }
                    catch{
                        Write-Error "Error setting [$($thisUpdatedClient.Name)].SetCustomProperty(`"DocLibId`",$($graphList.id)) on Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisUpdatedClient.Name)] in sync-netsuiteManagedMetaDataToSharePoint()"
                        }
                    }
                else{
                    Write-Warning "Could not retrive GraphList for [$($thisUpdatedClient.Name)][$($thisUpdatedClient.id)][$($thisUpdatedClient.CustomProperties.GraphDriveId)] - cannot repair Term. Something weird is going on here."
                    }
                }
            if($thisUpdatedClient.CustomProperties.DocLibId){
                Write-Verbose "[$($thisUpdatedClient.Name)][$($thisUpdatedClient.id)] has .CustomProperties.DocLibId value [$($thisUpdatedClient.CustomProperties.DocLibId)] - attempting to rename List to match Term name"
                try{
                    Write-Verbose "`tTrying: set-graphList -graphSiteId $clientSiteId -graphListId $($thisUpdatedClient.CustomProperties.DocLibId) -listPropertyHash @{displayName=$($thisUpdatedClient.Name)}"
                    $updatedGraphList = set-graphList -tokenResponse $tokenResponseSharePointBot -graphSiteId $clientSiteId -graphListId $thisUpdatedClient.CustomProperties.DocLibId -listPropertyHash @{displayName=$thisUpdatedClient.Name}
                    }
                catch{
                    Write-Error "Error setting List [$($clientSiteId)][$($thisUpdatedClient.CustomProperties.DocLibId)] DisplayName to [$($thisUpdatedClient.Name)] in sync-netsuiteManagedMetaDataToSharePoint() | Retrying with Verbose"
                    $oldVerbosePreference = $VerbosePreference
                    $VerbosePreference = 2
                    set-graphList -tokenResponse $tokenResponseSharePointBot -graphSiteId $clientSiteId -graphListId $thisUpdatedClient.CustomProperties.DocLibId -listPropertyHash @{displayName=$thisUpdatedClient.Name} -Verbose
                    $VerbosePreference = $oldVerbosePreference
                    }
        
                if($updatedGraphList.displayName -ne $thisUpdatedClient.Name){#If this didn;t work, it might be because of a DisplayName collision, but it'll need investigating by a human for now as no errors are returned for us to handle
                    Write-Warning "Failed to update List [$($updatedGraphList.displayName)][$($updatedGraphList.CustomProperties.DocLibId)][$($updatedGraphList.id)] DisplayName to [$($thisUpdatedClient.Name)] in sync-netsuiteManagedMetaDataToSharePoint()"
                    [array]$problemChilds += ,@($thisUpdatedClient,$updatedGraphList,"Failed to update List DisplayName via graph") 
                    }
                else{
                    Write-Verbose "`tSuccess! List renamed to [$($thisUpdatedClient.Name)]"
                    }
                }
            else{
                Write-Error "Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisUpdatedClient.Name)] does not have a .CustomProperties.DocLibId value, and one could not be determined - cannot update DisplayName."
                [array]$problemChilds += ,@($thisUpdatedClient,$updatedGraphList,"Failed to update List DisplayName via graph") 
                }
            }
        }

    #endregion

    #region Opportunities
    if($missingFromSpo){ #Refresh $allClientTerms if we've created new Clients
        $pnpTermGroup = "Kimble"
        $pnpTermSet = "Clients"
        $allClientTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes CustomProperties | ? {$_.IsDeprecated -eq $false}
        }
    $pnpTermGroup = "Kimble"
    $pnpTermSet = "Opportunities"
    $allOppTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes CustomProperties | ? {$_.IsDeprecated -eq $false}
    #Filter these client-side (CSOM, eh?) to get only the changes since this script last completed successfully
    [array]$oppTermsToCheck = $allOppTerms | ? {($_.LastModifiedDate -gt $lastSpoSyncRun -or $_.CustomProperties.flagForReprocessing -eq $true) -and ![string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteOppId) -and [string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteProjId)}
    
    Write-Host "Processing [$($oppTermsToCheck.Count)] Opportunities"
    if($oppTermsToCheck){
        if(!$allClientDrives){$allClientDrives = get-clientDrives} #Only load the Client Drives via graph if there is work to do and we haven't already got them

        @($oppTermsToCheck | Select-Object) | % {
            $tokenResponseSharePointBot = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponseSharePointBot -aadAppCreds $sharePointBotDetails
            $thisOppTerm = $_
            Write-Host "Processing Opp Term [$($thisOppTerm.Name)][$($thisOppTerm.id)][$($thisOppTerm.CustomProperties.DriveItemId)] for NetSuiteClientId [$($thisOppTerm.CustomProperties.NetSuiteClientId)]"
            process-folder -tokenResponse $tokenResponseSharePointBot -oppProjTerm $thisOppTerm -arrayOfAllClientTerms $allClientTerms -arrayOfLeadProjSubFolders $listOfLeadProjSubFolders -arrayOfAllOppTerms $allOppTerms
            } 
        }


    #endregion

    #region Projects
    $pnpTermGroup = "Kimble"
    $pnpTermSet = "Projects"
    $allProjTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes CustomProperties | ? {$_.IsDeprecated -eq $false}
    #Filter these client-side (CSOM, eh?) to get only the changes since this script last completed successfully
    [array]$projTermsToCheck = $allProjTerms | ? {($_.LastModifiedDate -gt $lastSpoSyncRun -or $_.CustomProperties.flagForReprocessing -eq $true) -and ![string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteProjId) -and ![string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteClientId)}
    Write-Host "Processing [$($projTermsToCheck.Count)] Projects"
    if($projTermsToCheck){
        if(!$allClientDrives){$allClientDrives = get-clientDrives} #Only load the Client Drives via graph if there is work to do and we haven't already got them

        @($projTermsToCheck | Select-Object) | % {
            $tokenResponseSharePointBot = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponseSharePointBot -aadAppCreds $sharePointBotDetails -renewTokenExpiringInSeconds 30
            $thisProjTerm = $_
            Write-Host "Processing Proj Term [$($thisProjTerm.Name)][$($thisProjTerm.id)][$($thisProjTerm.CustomProperties.DriveItemId)] for NetSuiteClientId [$($thisProjTerm.CustomProperties.NetSuiteClientId)]"
            process-folder -tokenResponse $tokenResponseSharePointBot -oppProjTerm $thisProjTerm -arrayOfAllClientTerms $allClientTerms -arrayOfLeadProjSubFolders $listOfClientFolders -arrayOfAllOppTerms $allOppTerms
            }

        #############################
        #Process Projects en-masse as don't want to query every DriveItem in every Drive in advance
        #############################
        #Check for DriveItemId
            #If DriveItemId, update folder
            #If no DriveItemId, Check for Opp
                #If Opp, update DriveItemId and add to flagForReprocessing (to include in next run)
                #If no Opp, Create new folders
        [array]$flagForReprocessing = @()
        $projTermsToCheck | % {
            $thisProjTerm = $_
            Write-Host "Processing Project Term [$($thisProjTerm.Name)][$($thisProjTerm.id)][$($thisProjTerm.CustomProperties.DriveItemId)] for NetSuiteClientId [$($thisProjTerm.CustomProperties.NetSuiteClientId)]"

            if($thisProjTerm.CustomProperties.DriveItemId){        #If DriveItemId, update folder
            try{
                $thisClientTerm = $allClientTerms | ? {$_.CustomProperties.NetSuiteId -eq $thisProjTerm.CustomProperties.NetSuiteClientId}
                Write-Host "`tProject Term [$($thisProjTerm.Name)].CustomProperties.DriveItemId is [$($thisProjTerm.CustomProperties.DriveItemId)] - setting Folder name to [$($thisProjTerm.Name)]"
                try{
                    $updatedFolder = set-graphDriveItem -tokenResponse $tokenResponseSharePointBot -driveId $thisClientTerm.CustomProperties.GraphDriveId -driveItemId $thisProjTerm.CustomProperties.DriveItemId -driveItemPropertyHash @{name=$thisProjTerm.Name}
                    }
                catch{
                    if($_.Exception -match "(409) Conflict"){
                        Write-Warning "`tPotential duplicate Project folder found for [$($thisClientTerm.Name)][$($thisProjTerm.Name)]"
                        $duplicateNetProjectFolder = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $thisClientTerm.CustomProperties.GraphDriveId | ? {$_.name -eq $thisProjTerm.Name}
                        if($duplicateNetProjectFolder.size -eq 0 -and $duplicateNetProjectFolder.id -ne $thisProjTerm.CustomProperties.DriveItemId){
                            Write-Warning "`tDeleting empty duplicate project folder"
                            delete-graphDriveItem -tokenResponse $tokenResponseSharePointBot -graphDriveId $thisClientTerm.CustomProperties.GraphDriveId -graphDriveItemId $duplicateNetProjectFolder.id -eTag $duplicateNetProjectFolder.eTag -Verbose
                            $updatedFolder = set-graphDriveItem -tokenResponse $tokenResponseSharePointBot -driveId $thisClientTerm.CustomProperties.GraphDriveId -driveItemId $thisProjTerm.CustomProperties.DriveItemId -driveItemPropertyHash @{name=$thisProjTerm.Name}
                            }
                        }
                    if($_.exception -eq "Cannot bind argument to parameter 'graphDriveId' because it is an empty string."){
                        [array]$noDriveForYou += $thisProjTerm
                        #Check whether this is an InterCompany ("pretend") client, and ignore it if it is.
                        }
                    }
                }
            catch{
                $_
                #There's a whole lot that could wrong here ($thisClientTerm.DriveId could be $null, the Drive could be missing or the DriveItem could be missing. Log the error for further investigation
                }
            }
                                                                                                                                            else{#If no DriveItemId, Check for Opp
            Write-Host "`tProject Term [$($thisProjTerm.Name)].CustomProperties.DriveItemId is missing - checking for Opportunity with NetSuiteProjectId -eq [$($thisProjTerm.CustomProperties.NetSuiteProjId)]"
            $thisOppTerm = $allOppTerms | ? {$_.CustomProperties.NetSuiteProjectId -eq $thisProjTerm.Id}
            if(![string]::IsNullOrWhiteSpace($thisOppTerm.CustomProperties.DriveItemId)){#If Opp, update DriveItemId and add to flagForReprocessing (to include in next run)
                Write-Host "`t`tOpportunity Term [$($thisOppTerm.Name)] found with .CustomProperties.DriveItemId [$($thisOppTerm.CustomProperties.DriveItemId)] - updating Project Term [$($thisProjTerm.Name)].CustomProperties.DriveItemId to match"
                $thisProjTerm.SetCustomProperty("DriveItemId",$thisOppTerm.CustomProperties.DriveItemId)
                try{
                    Write-Verbose "`tTrying: [$($thisProjTerm.Name)].SetCustomProperty(DriveItemId,[$($updatedFolder.id)])"
                    $thisProjTerm.Context.ExecuteQuery()
                    [array]$flagForReprocessing += $thisProjTerm
                    }
                catch{
                    Write-Error "Error updating DriveItemId CustomProperty [$($thisProjTerm.CustomProperties.DriveItemId))] to [$($updatedFolder.id)] for Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisProjTerm.Name)] in sync-netsuiteToManagedMetaData()"
                    [array]$flagForReprocessing += $thisNewClient
                    }
                }
            else{#If no Opp, Create new folders
                Write-Host "`t`tNo corresponding Opportunity Term [$($thisOppTerm.Name)] or DriveItemId [$($thisOppTerm.CustomProperties.DriveItemId)] found - creating new set of Project folders"
                $thisClientTerm = $allClientTerms | ? {$_.CustomProperties.NetSuiteId -eq $thisProjTerm.CustomProperties.NetSuiteClientId}
                [array]$customisedFolderList = $thisProjTerm.Name
                $customisedFolderList += $listOfLeadProjSubFolders | % {"$($thisProjTerm.Name)\$_"}
                $newProjectFolders = add-graphArrayOfFoldersToDrive -graphDriveId $thisClientTerm.CustomProperties.GraphDriveId -foldersAndSubfoldersArray $customisedFolderList -tokenResponse $tokenResponseSharePointBot -conflictResolution Fail #-ErrorAction SilentlyContinue
                $thisProjTerm.SetCustomProperty("DriveItemId",$newProjectFolders[1].id)
                try{
                    Write-Verbose "`tTrying: [$($thisProjTerm.Name)].SetCustomProperty(DriveItemId,[$($updatedFolder.id)])"
                    $thisProjTerm.Context.ExecuteQuery()
                    }
                catch{
                    Write-Error "Error updating DriveItemId CustomProperty [$($thisProjTerm.CustomProperties.DriveItemId))] to [$($updatedFolder.id)] for Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisProjTerm.Name)] in sync-netsuiteToManagedMetaData()"
                    [array]$flagForReprocessing += $thisNewClient
                    }
                }
            }
        }
        #############################
        #Update flagForReprocessing
        #############################
        $projTermsToCheck | % {
            $thisProjToUpdate = $_
            if($flagForReprocessing -notcontains $thisProjToUpdate){ #If the process above worked as expected, update SharePointLastModifiedDate to prevent it from being re-processed next time
                Write-Host "[$($thisProjToUpdate.Name)] was processed successfully - updating flagForReprocessing to [$false]"
                $thisProjToUpdate.SetCustomProperty("flagForReprocessing",$false)
                }
            else{
                Write-Warning "Something went wrong with [$($thisProjToUpdate.Name)] - flagging for reprocessing"
                $thisProjToUpdate.SetCustomProperty("flagForReprocessing",$true)
                }
            try{
                Write-Verbose "`tTrying: [$($thisProjToUpdate.Name)][$($thisProjToUpdate.Id)].SetCustomProperty(SharePointLastModifiedDate,$($now))"
                $thisProjToUpdate.Context.ExecuteQuery()
                }
            catch{
                Write-Error "Error setting CustomProperty SharePointLastModifiedDate on Term [$($thisProjToUpdate.Name)][$($thisProjToUpdate.Id)] in sync-netsuiteToManagedMetaData()"
                $_
                }
            }
        }

    #endregion


    ###########################################
    #If the script hasn't borked completely, update the LastSpoSyncRun timestamp
    Write-Host "Setting Term [Anthesis][IT][LastModified] CustomProperty LastSpoSyncRun = [$(Get-Date $now -f s)]"
    $lastProcessedTerm = Get-PnPTerm -TermGroup "Anthesis" -TermSet "IT" -Identity "LastModified" -Includes CustomProperties
    $lastProcessedTerm.SetCustomProperty("LastSpoSyncRun",$(Get-Date $now -f s))
    try{
        $lastProcessedTerm.Context.ExecuteQuery()
        }
    catch{
        #Pfft.
        }
    }


Write-Host "sync-netsuiteManagedMetaDataToSharePoint completed in [$($fullSyncTime.TotalSeconds)] seconds"

Stop-Transcript