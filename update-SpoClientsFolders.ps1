Start-Transcript "$($MyInvocation.MyCommand.Definition)_$(Get-Date -Format "yyMMdd").log" -Append

Import-Module .\_REST_Library-SPO.psm1
Import-Module .\_REST_Library-Kimble.psm1

##################################
#
#Get ready
#
##################################

$serverUrl = "https://anthesisllc.sharepoint.com" 
$sitePath = "/clients"
$listOfClientFolders = @("_Kimble automatically creates Lead & Project folders","Background","Non-specific BusDev")
$listOfOppProjSubFolders = @("Admin & contracts", "Analysis","Data & refs","Meetings","Proposal","Reports","Summary (marketing) - end of project")

$o365user = "kevin.maitland@anthesisgroup.com"
$o365Pass = ConvertTo-SecureString (Get-Content 'C:\New Text Document.txt') -AsPlainText -Force
$credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $o365user, $o365Pass

Set-SPORestCredentials -Credential $credential
get-newDigest -serverUrl $serverUrl -sitePath $sitePath

#Log what
$logErrors = $true
$logActions = $true
$logResults = $true
$verboseLogging = $false
#Log to where
$logToScreen = $true
$logToFile = $true
$logfile = "C:\Scripts\Logs\update-spoClientsAndProjectFolders_$(Get-Date -Format "yyMMdd").log"
#$logFile = "C:\Scripts\Logs\update-SpoProjectsFolders.log"
$logErrorsToEmail = $true
$smtpServer = "anthesisgroup-com.mail.protection.outlook.com"
#$mailFrom = $MyInvocation.MyCommand.Name+"@sustain.co.uk"
$mailFrom = "scriptrobot@sustain.co.uk"
$mailTo = "kevin.maitland@anthesisgroup.com"

#region Sync Clients
#$allSpClients = get-itemsInList -sitePath $sitePath -listName "Kimble Clients"
#$dirtySpClients = get-itemsInList -sitePath $sitePath -listName "Kimble Clients" -oDataQuery "?&`$filter=Title eq 'Hussmann Corporation'"
$dirtySpClients = get-itemsInList -sitePath $sitePath -listName "Kimble Clients" -oDataQuery "?&`$filter=IsDirty eq 1"

#$recreateAllFolders = $true
foreach($dirtyClient in $dirtySpClients){
    if((!$dirtyClient.PreviousName -and !$dirtyClient.PreviousDescription) -OR $recreateAllFolders -eq $true){
        #Create a new Library and hope the user hasn't just deleted the Description
        try{
            log-action "new-library -sitePath $sitePath -libraryName $($dirtyClient.Title) -libraryDesc $($dirtyClient.ClientDescription)"
            $newLibrary = new-library -sitePath $sitePath -libraryName $dirtyClient.Title -libraryDesc $dirtyClient.ClientDescription
            if($newLibrary){
                log-result "SUCCESS: Library is there!"
                foreach($sub in $listOfClientFolders){
                    log-action "new-FolderInLibrary -site $sitePath -libraryName (/$($dirtyClient.Title)) -folderName $sub"
                    $newFolder = new-FolderInLibrary -site $sitePath -libraryName ("/"+$dirtyClient.Title) -folderName $sub
                    if ($newFolder){log-result "SUCCESS: $($dirtyClient.Title)\$sub created!"}
                    else{log-result "FAILURE: $($dirtyClient.Title)\$sub was not created!"}
                    }
                }
            else{log-result "FAILURE: $($dirtyClient.Title) was not created!"}
            }
        catch{log-error $_ -myFriendlyMessage "Failed to create new Library or subfolder for $($dirtyClient.Title)"}
        #Now try to add the new ClientName to the TermStore
        try{
            log-action "add-termToStore -pGroup Kimble -pSet Clients -pTerm $($dirtyClient.Title)"
            add-termToStore -pGroup "Kimble" -pSet "Clients" -pTerm $($dirtyClient.Title)
            log-result "SUCCESS: $($dirtyClient.Title) added to Managed MetaData Term Store"
            }
        catch{log-error $_ -myFriendlyMessage "Failed to add $($dirtyClient.Title) to Term Store"}
        #If we've got this far, try to update the Client in [Kimble Clients]
        try{
            if((get-library -sitePath $sitePath -libraryName $dirtyClient.Title) -ne $false){ #If it's worked, set the IsDirty flag to $false to prevent it reprocessing
                log-result "Successfully validated!"
                log-action "update-itemInList -serverUrl $serverUrl -sitePath $sitePath -listName `"Kimble Clients`" -predeterminedItemType $($dirtyClient.__metadata.type) -itemId $($dirtyClient.Id) -hashTableOfItemData @{IsDirty=$false}"
                update-itemInList -serverUrl $serverUrl -sitePath $sitePath -listName "Kimble Clients" -predeterminedItemType $dirtyClient.__metadata.type -itemId $dirtyClient.Id -hashTableOfItemData @{IsDirty=$false} | Out-Null
                log-result "Successfully updated!"
                }
            else{log-result -myMessage "Uh-oh, I couldn't find the Library I (allegedly) just created: [$($dirtyClient.Title)] this will stay as IsDirty=true forever :("}
            }
        catch{log-error $_ -myFriendlyMessage "Failed to update $($dirtyClient.Title) in [Kimble Clients] List - this will stay as IsDirty=true forever :("}
        }
    elseif(($dirtyClient.PreviousName) -and ($dirtyClient.PreviousName -ne $dirtyClient.Title)){
        #Update the folder name
        try{
            log-action "update-list -sitePath $sitePath -listName $($dirtyClient.PreviousName) -hashTableOfUpdateData @{Title=$($dirtyClient.Title)}"
            update-list -sitePath $sitePath -listName $dirtyClient.PreviousName -hashTableOfUpdateData @{Title=$dirtyClient.Title} | Out-Null
            log-result "SUCCESS: $($dirtyClient.PreviousName) updated to $($dirtyClient.Title)"
            }
        catch{log-error $Error[0] -myFriendlyMessage "Failed to update Library Title $($dirtyClient.PreviousName) to $($dirtyClient.Title)"}
        #Update teh Managed MetaData in the TermStore
        try{
            log-action "rename-termInStore -pGroup Kimble -pSet Clients -pOldTerm $($dirtyClient.PreviousName) -pNewTerm $($dirtyClient.Title)"
            rename-termInStore -pGroup "Kimble" -pSet "Clients" -pOldTerm $($dirtyClient.PreviousName) -pNewTerm $($dirtyClient.Title)
            log-result "SUCCESS: Term $($dirtyClient.PreviousName) renamed to $($dirtyClient.Title)"
            }
        catch{log-error $Error[0] -myFriendlyMessage "Failed to rename ManagedMetadata term $($dirtyClient.PreviousName) to $($dirtyClient.Title)"}
        #Update the Client in [Kimble Clients]
        try{
            if((get-list -sitePath $sitePath -listName $dirtyClient.Title) -ne $false){ #If it's worked, set the IsDirty flag to $false to prevent it reprocessing
                update-itemInList -serverUrl $serverUrl -sitePath $sitePath -listName "Kimble Clients" -predeterminedItemType $dirtyClient.__metadata.type -itemId $dirtyClient.Id -hashTableOfItemData @{IsDirty=$false} | Out-Null
                }
            else{log-result -myMessage "Uh-oh, I couldn't find the Library I (allegedly) just updated: [$($dirtyClient.Title)] this will stay as IsDirty=true forever :("}
            }
        catch{log-error $Error[0] -myFriendlyMessage "Error while renaming $($dirtyClient.PreviousName) to $($dirtyClient.Title) in [Kimble Clients] List - this will stay as IsDirty=true forever :("}
        }
    elseif(((sanitise-stripHtml $dirtyClient.PreviousDescription) -ne (sanitise-stripHtml $dirtyClient.ClientDescription)) -or ((sanitise-stripHtml $dirtyClient.ClientDescription) -ne ($dirtyClient.ClientDescription))){
        #Update the Library's Description
        try{
            update-list -sitePath $sitePath -listName $dirtyClient.Title -hashTableOfUpdateData @{Description=$(sanitise-stripHtml $dirtyClient.ClientDescription)} | Out-Null
            if(sanitise-stripHtml $((get-list -sitePath $sitePath -listName $dirtyClient.Title).ClientDescription) -eq $(sanitise-stripHtml $dirtyClient.ClientDescription)){ #If it's worked, set the IsDirty flag to $false to prevent it reprocessing
                update-itemInList -serverUrl $serverUrl -sitePath $sitePath -listName "Kimble Clients" -predeterminedItemType $dirtyClient.__metadata.type -itemId $dirtyClient.Id -hashTableOfItemData @{IsDirty=$false} | Out-Null
                }
            else{log-result -myMessage "Uh-oh, I couldn't find the Library I (allegedly) just updated: [$($dirtyClient.Title)] this will stay as IsDirty=true forever :("}
            }
        catch{log-error $Error[0] -myFriendlyMessage "Failed to update Library Description $($dirtyClient.PreviousName) to $($dirtyClient.Title)"}
        }
    
    }
#endregion

#region Sync Projects
    #$dirtySpProjects = get-itemsInList -sitePath $sitePath -listName "Kimble Projects" -oDataQuery "?&`$filter=Title eq '210717_Equitix_Gaia/Wrexham GMO wood waste sales strategy (E002556)'"
    $dirtySpProjects = get-itemsInList -sitePath $sitePath -listName "Kimble Projects" -oDataQuery "?&`$filter=IsDirty eq 1"
    
    $spClients = get-itemsInList -sitePath $sitePath -listName "Kimble Clients"
    $kimbleClientHashTable = @{}
    foreach ($spC in $spClients){$kimbleClientHashTable.Add($spC.KimbleId,$(sanitise-forSharePointListName $spc.Title))}

foreach($dirtyProject in $dirtySpProjects){
    if(!$dirtyProject.PreviousName -and (!$dirtyProject.PreviousKimbleClientId -or $dirtyProject.PreviousKimbleClientId -eq $dirtyProject.KimbleClientId)){
        #Create a new folder tree under the Client Library
        Write-Host -ForegroundColor Magenta "Creating New Project folders for $($kimbleClientHashTable[$dirtyProject.KimbleClientId])"
        try{
            log-action "new-folderInLibrary -sitePath $sitePath -libraryName $("/"+$kimbleClientHashTable[$dirtyProject.KimbleClientId]) -folderName $($dirtyProject.Title)"
            if ($kimbleClientHashTable[$dirtyProject.KimbleClientId] -eq $null){
                log-result "FAILURE: Client could not be found in [Kimble Clients]";
                log-error -myError $null -myFriendlyMessage "The Client with Id:$($dirtyProject.KimbleClientId) could not be determined for project [$($dirtyProject.Title)] (Id:$($dirtyProject.KimbleId))" -doNotLogToEmail $true
                continue
                }
            $foo = new-folderInLibrary -sitePath $sitePath -libraryName ("/"+ $kimbleClientHashTable[$dirtyProject.KimbleClientId]) -folderName $dirtyProject.Title.Replace("/","")
            if($foo){
                log-result "SUCCESS: Folder is there!"
                foreach($sub in $listOfOppProjSubFolders){
                    log-action "new-FolderInLibrary -site $sitePath -libraryName $("/"+$kimbleClientHashTable[$dirtyProject.KimbleClientId]) -folderPath $("/"+$dirtyProject.Title.Replace("/",'')) -folderName $sub"
                    $foo = new-FolderInLibrary -site $sitePath -libraryName ("/"+$kimbleClientHashTable[$dirtyProject.KimbleClientId]) -folderPath ("/"+$dirtyProject.Title.Replace("/","")) -folderName $sub
                    if($foo){log-result "SUCCESS: $($kimbleClientHashTable[$dirtyProject.KimbleClientId]+"\"+$dirtyProject.Title)\$sub is there!"}
                    else{log-result "FAILURE: SubFolder $sub was not created/retrievable"
                        log-Error -myError $null -myFriendlyMessage "new-folderInLibrary -sitePath $sitePath -libraryname $("/"+$kimbleClientHashTable[$dirtyProject.KimbleClientId]) -folderPath $("/"+$dirtyProject.Title.Replace("/",'')) -folderName $($sub) failed to create a retrieavble folder"
                        }
                    }
                }
            else{log-result "FAILURE: Folder was not created/retrievable"
                log-Error -myError $null -myFriendlyMessage "new-folderInLibrary -sitePath $sitePath -libraryName $("/"+$kimbleClientHashTable[$dirtyProject.KimbleClientId]) -folderName $($dirtyProject.Title) failed to create a retrieavble folder"
                }
                
            if((get-folderInLibrary -sitePath $sitePath -libraryName $kimbleClientHashTable[$dirtyProject.KimbleClientId] -folderName $dirtyProject.Title) -ne $false){ #If it's worked, set the IsDirty flag to $false to prevent it reprocessing
                log-result "Project folder successfully validated!"
                log-action "update-itemInList -serverUrl $serverUrl -sitePath $sitePath -listName `"Kimble Projects`" -predeterminedItemType $($dirtyProject.__metadata.type) -itemId $($dirtyProject.Id) -hashTableOfItemData @{IsDirty=$false}"
                update-itemInList -serverUrl $serverUrl -sitePath $sitePath -listName "Kimble Projects" -predeterminedItemType $dirtyProject.__metadata.type -itemId $dirtyProject.Id -hashTableOfItemData @{IsDirty=$false} | Out-Null
                log-result "Successfully updated!"
                }
            else{log-result -myMessage "Uh-oh, I couldn't find the Project Folder I (allegedly) just created: [$($dirtyProject.Title)] this will stay as IsDirty=true forever :("}
            }
        catch{log-error $_ -myFriendlyMessage "Failed to create new Folder or subfolder for $($dirtyProject.Title)"}
        }
    elseif(($dirtyProject.PreviousName) -and ($dirtyProject.PreviousName -ne $dirtyProject.Title)){
        #Update the folder name
        Write-Host -ForegroundColor Magenta "Updating Project name"
        try{
            update-list -sitePath $sitePath -listName $dirtyProject.PreviousName -hashTableOfUpdateData @{Title=$dirtyProject.Title} | Out-Null
            if((get-folderInLibrary -sitePath $sitePath -libraryName $kimbleClientHashTable[$dirtyProject.KimbleClientId] -folderName $dirtyProject.Title) -ne $false){ #If it's worked, set the IsDirty flag to $false to prevent it reprocessing
                update-itemInList -serverUrl $serverUrl -sitePath $sitePath -listName "Kimble Projects" -predeterminedItemType $dirtyProject.__metadata.type -itemId $dirtyProject.Id -hashTableOfItemData @{IsDirty=$false} | Out-Null
                }
            else{log-result -myMessage "Uh-oh, I couldn't find the Library I (allegedly) just updated: [$($dirtyProject.Title)] this will stay as IsDirty=true forever :("}
            }
        catch{log-error $Error[0] -myFriendlyMessage "Failed to update Library Title $($dirtyProject.PreviousName) to $($dirtyProject.Title)"}
        }
    elseif($dirtyProject.PreviousKimbleClientId -ne $dirtyProject.KimbleClientId){
        #Move the folder to the new Client
        Write-Host -ForegroundColor Magenta "Moving Project to different Client"
        try{
            #Yeah Kev, you actually need to write some code to *do* this. Move $kimbleClientHashTable[$dirtyProject.PreviousKimbleClientId]/$dirtyProject.Title to $kimbleClientHashTable[$dirtyProject.KimbleClientId]
            if((get-folderInLibrary -sitePath $sitePath -libraryName $kimbleClientHashTable[$dirtyProject.KimbleClientId] -folderName $dirtyProject.Title) -ne $false){ #If it's worked, set the IsDirty flag to $false to prevent it reprocessing
                #update-itemInList -serverUrl $serverUrl -sitePath $sitePath -listName "Kimble Projects" -predeterminedItemType $dirtyProject.__metadata.type -itemId $dirtyProject.Id -hashTableOfItemData @{IsDirty=$false} | Out-Null
                }
            else{log-result -myMessage "Uh-oh, I couldn't find the Library I (allegedly) just updated: [$($dirtyProject.Title)] this will stay as IsDirty=true forever :("}
            }
        catch{log-error $Error[0] -myFriendlyMessage "Failed to update Library Description $($dirtyProject.PreviousName) to $($dirtyProject.Title)"}
        }
    }
#endregion

#region Sync Leads
    $dirtySpLeads = get-itemsInList -sitePath $sitePath -listName "Kimble Leads" -oDataQuery "?&`$filter=IsDirty eq 1"
    
    #$spClients = get-itemsInList -sitePath $sitePath -listName "Kimble Clients" #No need to requery this after the Projects region
    #$kimbleClientHashTable = @{}
    #foreach ($spC in $spClients){$kimbleClientHashTable.Add($spC.KimbleId,$(sanitise-forSharePointListName $spc.Title))}

foreach($dirtyLead in $dirtySpLeads){
    $leadFoldername = "BD_"+$dirtyLead.Title
    if(!$dirtyLead.PreviousName -and (!$dirtyLead.PreviousKimbleClientId -or $dirtyLead.PreviousKimbleClientId -eq $dirtyLead.KimbleClientId)){
        #Create a new folder tree under the Client Library
        Write-Host -ForegroundColor Magenta "Creating New Lead folders"
        try{
            log-action "new-folderInLibrary -sitePath $sitePath -libraryName $("/"+$kimbleClientHashTable[$dirtyLead.KimbleClientId]) -folderName $leadFoldername"
            if ($kimbleClientHashTable[$dirtyLead.KimbleClientId] -eq $null){
                log-result "FAILURE: Client could not be found in [Kimble Clients]";
                log-error -myError $null -myFriendlyMessage "The Client with Id:$($dirtyLead.KimbleClientId) could not be determined for Lead [$($dirtyLead.Title)] (Id:$($dirtyLead.KimbleId))" -doNotLogToEmail $true
                continue
                }
            $foo = new-folderInLibrary -sitePath $sitePath -libraryName ("/"+ $kimbleClientHashTable[$dirtyLead.KimbleClientId]) -folderName $leadFoldername
            if($foo){
                log-result "SUCCESS: Folder is there!"
                foreach($sub in $listOfOppProjSubFolders){
                    log-action "new-FolderInLibrary -site $sitePath -libraryName $("/"+$kimbleClientHashTable[$dirtyLead.KimbleClientId]) -folderPath $("/"+$leadFoldername.Replace("/",'')) -folderName $sub"
                    $foo = new-FolderInLibrary -site $sitePath -libraryName ("/"+$kimbleClientHashTable[$dirtyLead.KimbleClientId]) -folderPath ("/"+$leadFoldername.Replace("/","")) -folderName $sub
                    if($foo){log-result "SUCCESS: $($kimbleClientHashTable[$dirtyLead.KimbleClientId]+"\"+$leadFoldername)\$sub is there!"}
                    else{log-result "FAILURE: SubFolder $sub was not created/retrievable"
                        log-Error -myError $null -myFriendlyMessage "new-folderInLibrary -sitePath $sitePath -libraryname $("/"+$kimbleClientHashTable[$dirtyLead.KimbleClientId]) -folderPath $("/"+$leadFoldername.Replace("/",'')) -folderName $($sub) failed to create a retrieavble folder"
                        }
                    }
                }
            else{log-result "FAILURE: Folder was not created/retrievable"
                log-Error -myError $null -myFriendlyMessage "new-folderInLibrary -sitePath $sitePath -libraryName $("/"+$kimbleClientHashTable[$dirtyLead.KimbleClientId]) -folderName $($leadFoldername) failed to create a retrieavble folder"
                }
                
            if((get-folderInLibrary -sitePath $sitePath -libraryName $kimbleClientHashTable[$dirtyLead.KimbleClientId] -folderName $leadFoldername) -ne $false){ #If it's worked, set the IsDirty flag to $false to prevent it reprocessing
                log-result "Lead folder successfully validated!"
                log-action "update-itemInList -serverUrl $serverUrl -sitePath $sitePath -listName `"Kimble Leads`" -predeterminedItemType $($dirtyLead.__metadata.type) -itemId $($dirtyLead.Id) -hashTableOfItemData @{IsDirty=$false}"
                update-itemInList -serverUrl $serverUrl -sitePath $sitePath -listName "Kimble Leads" -predeterminedItemType $dirtyLead.__metadata.type -itemId $dirtyLead.Id -hashTableOfItemData @{IsDirty=$false} | Out-Null
                log-result "Successfully updated!"
                }
            else{log-result -myMessage "Uh-oh, I couldn't find the Lead Folder I (allegedly) just created: [$($leadFoldername)] this will stay as IsDirty=true forever :("}
            }
        catch{log-error $_ -myFriendlyMessage "Failed to create new Folder or subfolder for $($leadFoldername)"}
        }
    elseif(($dirtyLead.PreviousName) -and ($dirtyLead.PreviousName -ne $dirtyLead.Title)){
        #Update the folder name
        Write-Host -ForegroundColor Magenta "Updating Lead name"
        try{
            #update-list -sitePath $sitePath -listName $dirtyLead.PreviousName -hashTableOfUpdateData @{Title=$dirtyLead.Title} | Out-Null
            if((get-folderInLibrary -sitePath $sitePath -libraryName $kimbleClientHashTable[$dirtyLead.KimbleClientId] -folderName $leadFoldername) -ne $false){ #If it's worked, set the IsDirty flag to $false to prevent it reprocessing
                update-itemInList -serverUrl $serverUrl -sitePath $sitePath -listName "Kimble Leads" -predeterminedItemType $dirtyLead.__metadata.type -itemId $dirtyLead.Id -hashTableOfItemData @{IsDirty=$false} | Out-Null
                }
            else{log-result -myMessage "Uh-oh, I couldn't find the Library I (allegedly) just updated: [$($dirtyLead.Title)] this will stay as IsDirty=true forever :("}
            }
        catch{log-error $Error[0] -myFriendlyMessage "Failed to update Library Title $($dirtyLead.PreviousName) to $($dirtyLead.Title)"}
        }
    elseif($dirtyLead.PreviousKimbleClientId -ne $dirtyLead.KimbleClientId){
        #Move the folder to the new Client
        Write-Host -ForegroundColor Magenta "Moving Lead to different Client"
        try{
            #Yeah Kev, you actually need to write some code to *do* this. Move $kimbleClientHashTable[$dirtyLead.PreviousKimbleClientId]/$dirtyLead.Title to $kimbleClientHashTable[$dirtyLead.KimbleClientId]
            if((get-folderInLibrary -sitePath $sitePath -libraryName $kimbleClientHashTable[$dirtyLead.KimbleClientId] -folderName $dirtyLead.Title) -ne $false){ #If it's worked, set the IsDirty flag to $false to prevent it reprocessing
                #update-itemInList -serverUrl $serverUrl -sitePath $sitePath -listName "Kimble Leads" -predeterminedItemType $dirtyLead.__metadata.type -itemId $dirtyLead.Id -hashTableOfItemData @{IsDirty=$false} | Out-Null
                }
            else{log-result -myMessage "Uh-oh, I couldn't find the Library I (allegedly) just updated: [$($dirtyLead.Title)] this will stay as IsDirty=true forever :("}
            }
        catch{log-error $Error[0] -myFriendlyMessage "Failed to update Library Description $($dirtyLead.PreviousName) to $($dirtyLead.Title)"}
        }
    }
#endregion



#new-folderInLibrary -sitePath "/clients" -libraryAndFolderPath "/Waste & Resources Action Programme (WRAP)" -folderName "20160705_WRAP_Raw_Materials_Tool_Workshop (E001372)"
#new-library -sitePath "/clients" -libraryName "Waste & Resources Action Programme (WRAP)" -libraryDesc "Client folder for Waste & Resources Action Programme (WRAP)"

<#$usefulShit =@()
foreach($proj in $dirtySpProjects){
    $obj = New-Object psobject
    $obj | Add-Member NoteProperty AccountName $kimbleClientHashTable[$proj.KimbleClientId]
    $obj | Add-Member NoteProperty AccountId $kimbleClientHashTable[$proj.KimbleClientId]["Id"]
    $obj | Add-Member NoteProperty AccountType $kimbleClientHashTable[$proj.KimbleClientId]["Type"]
    $obj | Add-Member NoteProperty IsClient $kimbleClientHashTable[$proj.KimbleClientId]["IsClient"]
    $obj | Add-Member NoteProperty ProjectName $proj.Name
    $obj | Add-Member NoteProperty ProjectId $proj.Id

    $usefulShit += $obj
    }
#>
Stop-Transcript