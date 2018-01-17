Start-Transcript "$($MyInvocation.MyCommand.Definition)_$(Get-Date -Format "yyMMdd").log" -Append

Import-Module .\_CSOM_Library-SPO.psm1
Import-Module .\_REST_Library-Kimble.psm1
Import-Module .\_REST_Library-SPO.psm1

##################################
#
#Get ready
#
##################################
$serverUrl = "https://anthesisllc.sharepoint.com" 
$sitePath = "/subs"
$o365user = "kevin.maitland@anthesisgroup.com"
$o365Pass = ConvertTo-SecureString (Get-Content 'C:\New Text Document.txt') -AsPlainText -Force
$credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $o365user, $o365Pass



Set-SPORestCredentials -Credential $credential
get-newDigest -serverUrl $serverUrl -sitePath $sitePath
$digest.GetContextWebInformation.FormDigestValue
#Log what
$logErrors = $true
$logActions = $true
$logResults = $true
$verboseLogging = $false
#Log to where
$logToScreen = $true
$logToFile = $true
$logfile = "C:\Scripts\Logs\update-spoCSupplierstFolders_$(Get-Date -Format "yyMMdd").log"
#$logFile = "C:\Scripts\Logs\update-SpoProjectsFolders.log"
$logErrorsToEmail = $true
$smtpServer = "anthesisgroup-com.mail.protection.outlook.com"
#$mailFrom = $MyInvocation.MyCommand.Name+"@sustain.co.uk"
$mailFrom = "scriptrobot@sustain.co.uk"
$mailTo = "kevin.maitland@anthesisgroup.com"
$listOfSupplierFolders = @("_Kimble automatically creates Supplier & Subcontractor folders","Background")
$listOfOppProjSubFolders = @("Admin & contracts", "Analysis","Data & refs","Meetings","Proposal","Reports","Summary (marketing) - end of project")


#region Sync Suppliers
#$allSpSuppliers = get-itemsInList -sitePath $sitePath -listName "Kimble Suppliers"
$dirtySpSuppliers = get-itemsInList -sitePath $sitePath -listName "Kimble Suppliers" -oDataQuery "?&`$filter=IsDirty eq 1"

#$recreateAllFolders = $true
foreach($dirtySupplier in $dirtySpSuppliers){
    if((!$dirtySupplier.PreviousName -and !$dirtySupplier.PreviousDescription) -OR $recreateAllFolders -eq $true){
        #Create a new Library and hope the user hasn't just deleted the Description
        try{
            log-action "new-library -sitePath $sitePath -libraryName $($dirtySupplier.Title) -libraryDesc $($dirtySupplier.SupplierDescription)"
            $newLibrary = new-library -sitePath $sitePath -libraryName $dirtySupplier.Title -libraryDesc $dirtySupplier.SupplierDescription
            if($newLibrary){
                log-result "SUCCESS: Library is there!"
                foreach($sub in $listOfSupplierFolders){
                    log-action "new-FolderInLibrary -site $sitePath -libraryName (/$($dirtySupplier.Title)) -folderName $sub"
                    $newFolder = new-FolderInLibrary -site $sitePath -libraryName ("/"+$dirtySupplier.Title) -folderName $sub
                    if ($newFolder){log-result "SUCCESS: $($dirtySupplier.Title)\$sub created!"}
                    else{log-result "FAILURE: $($dirtySupplier.Title)\$sub was not created!"}
                    }
                }
            else{log-result "FAILURE: $($dirtySupplier.Title) was not created!"}
            }
        catch{log-error $_ -myFriendlyMessage "Failed to create new Library or subfolder for $($dirtySupplier.Title)"}
        #Now try to add the new SupplierName to the TermStore
        try{
            log-action "add-termToStore -pGroup Kimble -pSet Subcontractors -pTerm $($dirtySupplier.Title)"
            add-termToStore -pGroup "Kimble" -pSet "Subcontractors" -pTerm $($dirtySupplier.Title)
            log-result "SUCCESS: $($dirtySupplier.Title) added to Managed MetaData Term Store"
            }
        catch{log-error $_ -myFriendlyMessage "Failed to add $($dirtySupplier.Title) to Term Store"}
        #If we've got this far, try to update the Supplier in [Kimble Suppliers]
        try{
            if((get-library -sitePath $sitePath -libraryName $dirtySupplier.Title) -ne $false){ #If it's worked, set the IsDirty flag to $false to prevent it reprocessing
                log-result "Successfully validated!"
                log-action "update-itemInList -serverUrl $serverUrl -sitePath $sitePath -listNameOrGuid `"Kimble Suppliers`" -predeterminedItemType $($dirtySupplier.__metadata.type) -itemId $($dirtySupplier.Id) -hashTableOfItemData @{IsDirty=$false}"
                update-itemInList -serverUrl $serverUrl -sitePath $sitePath -listNameOrGuid "Kimble Suppliers" -predeterminedItemType $dirtySupplier.__metadata.type -itemId $dirtySupplier.Id -hashTableOfItemData @{IsDirty=$false} | Out-Null
                log-result "Successfully updated!"
                }
            else{log-result -myMessage "Uh-oh, I couldn't find the Library I (allegedly) just created: [$($dirtySupplier.Title)] this will stay as IsDirty=true forever :("}
            }
        catch{log-error $_ -myFriendlyMessage "Failed to update $($dirtySupplier.Title) in [Kimble Suppliers] List - this will stay as IsDirty=true forever :("}
        }
    elseif(($dirtySupplier.PreviousName) -and ($dirtySupplier.PreviousName -ne $dirtySupplier.Title)){
        #Update the folder name
        try{
            log-action "update-list -sitePath $sitePath -listName $($dirtySupplier.PreviousName) -hashTableOfUpdateData @{Title=$($dirtySupplier.Title)}"
            update-list -sitePath $sitePath -listName $dirtySupplier.PreviousName -hashTableOfUpdateData @{Title=$dirtySupplier.Title} | Out-Null
            log-result "SUCCESS: $($dirtySupplier.PreviousName) updated to $($dirtySupplier.Title)"
            }
        catch{log-error $Error[0] -myFriendlyMessage "Failed to update Library Title $($dirtySupplier.PreviousName) to $($dirtySupplier.Title)"}
        #Update teh Managed MetaData in the TermStore
        try{
            log-action "rename-termInStore -pGroup Kimble -pSet Suppliers -pOldTerm $($dirtySupplier.PreviousName) -pNewTerm $($dirtySupplier.Title)"
            rename-termInStore -pGroup "Kimble" -pSet "Suppliers" -pOldTerm $($dirtySupplier.PreviousName) -pNewTerm $($dirtySupplier.Title)
            log-result "SUCCESS: Term $($dirtySupplier.PreviousName) renamed to $($dirtySupplier.Title)"
            }
        catch{log-error $Error[0] -myFriendlyMessage "Failed to rename ManagedMetadata term $($dirtySupplier.PreviousName) to $($dirtySupplier.Title)"}
        #Update the Supplier in [Kimble Suppliers]
        try{
            if((get-list -sitePath $sitePath -listName $dirtySupplier.Title) -ne $false){ #If it's worked, set the IsDirty flag to $false to prevent it reprocessing
                update-itemInList -serverUrl $serverUrl -sitePath $sitePath -listNameOrGuid "Kimble Suppliers" -predeterminedItemType $dirtySupplier.__metadata.type -itemId $dirtySupplier.Id -hashTableOfItemData @{IsDirty=$false} | Out-Null
                }
            else{log-result -myMessage "Uh-oh, I couldn't find the Library I (allegedly) just updated: [$($dirtySupplier.Title)] this will stay as IsDirty=true forever :("}
            }
        catch{log-error $Error[0] -myFriendlyMessage "Error while renaming $($dirtySupplier.PreviousName) to $($dirtySupplier.Title) in [Kimble Suppliers] List - this will stay as IsDirty=true forever :("}
        }
    elseif(((sanitise-stripHtml $dirtySupplier.PreviousDescription) -ne (sanitise-stripHtml $dirtySupplier.SupplierDescription)) -or ((sanitise-stripHtml $dirtySupplier.SupplierDescription) -ne ($dirtySupplier.SupplierDescription))){
        #Update the Library's Description
        try{
            update-list -sitePath $sitePath -listName $dirtySupplier.Title -hashTableOfUpdateData @{Description=$(sanitise-stripHtml $dirtySupplier.SupplierDescription)} | Out-Null
            if(sanitise-stripHtml $((get-list -sitePath $sitePath -listName $dirtySupplier.Title).SupplierDescription) -eq $(sanitise-stripHtml $dirtySupplier.SupplierDescription)){ #If it's worked, set the IsDirty flag to $false to prevent it reprocessing
                update-itemInList -serverUrl $serverUrl -sitePath $sitePath -listNameOrGuid "Kimble Suppliers" -predeterminedItemType $dirtySupplier.__metadata.type -itemId $dirtySupplier.Id -hashTableOfItemData @{IsDirty=$false} | Out-Null
                }
            else{log-result -myMessage "Uh-oh, I couldn't find the Library I (allegedly) just updated: [$($dirtySupplier.Title)] this will stay as IsDirty=true forever :("}
            }
        catch{log-error $Error[0] -myFriendlyMessage "Failed to update Library Description $($dirtySupplier.PreviousName) to $($dirtySupplier.Title)"}
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