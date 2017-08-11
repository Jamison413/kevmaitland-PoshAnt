Start-Transcript "$($MyInvocation.MyCommand.Definition)_$(Get-Date -Format "yyMMdd").log" -Append

Import-Module .\_CSOM_Library-SPO.psm1
Import-Module .\_REST_Library-Kimble.psm1
Import-Module .\_REST_Library-SPO.psm1

##################################
#
#Get ready
#
##################################
$o365user = "kevin.maitland@anthesisgroup.com"
$o365Pass = ConvertTo-SecureString (Get-Content 'C:\New Text Document.txt') -AsPlainText -Force
$credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $o365user, $o365Pass
$logfile = "C:\users\administrator.sustainltd\Desktop\provisionSpoClients.log"
$logErrors = $true
$logMethodMain = $true
$logFunctionCalls = $true
Set-SPORestCredentials -Credential $credential

$oAuthReqBody = Get-KimbleAuthorizationTokenWithUsernamePasswordFlowRequestBody -client_id $clientId -client_secret $clientSecret -user_name $username -pass_word $password -security_token $securityToken
try{$kimbleAccessToken=Invoke-RestMethod -Method Post -Uri $callbackUri -Body $oAuthReqBody} catch {Failure}
$kimbleRestHeaders = @{Authorization = "Bearer " + $kimbleAccessToken.access_token}


##################################
#
#Do Stuff
#
##################################

#region Kimble Sync
#Get the last Project modified in /Projects/lists/Kimble Projects to minimise the number of records to process
$serverUrl = "https://anthesisllc.sharepoint.com" 
$sitePath = "/clients"
$listName = "Kimble Projects"
get-newDigest -serverUrl $serverUrl -sitePath $sitePath
$kp = get-list -sitePath $sitePath -listName $listName
#$lastModifiedProject = get-lastModifiedItemInList -sitePath $sitePath -listName "Kimble Projects"

#Get the Kimble Projects that have been modifed since the last update
$cutoffDate = (Get-Date (Get-Date $kp.LastItemModifiedDate).AddHours(-1) -Format s) #Look one hour behind just in case there is ever a delay between polling Kimble and updating SharePoint
#$cutoffDate = (Get-Date (Get-Date $kp.LastItemModifiedDate).AddYears(-1) -Format s) #Bodge this once for the initial Sync
$soqlQuery = "SELECT Name,Id,KimbleOne__Account__c,LastModifiedDate,SystemModStamp,CreatedDate,IsDeleted FROM KimbleOne__DeliveryGroup__c WHERE LastModifiedDate > $cutoffDate`Z"
$kimbleModifiedProjects = Get-KimbleSoqlDataset -queryUri $queryUri -soqlQuery $soqlQuery -restHeaders $kimbleRestHeaders
$kimbleChangedProjects = $kimbleModifiedProjects | ?{$_.LastModifiedDate -lt $cutoffDate}
$kimbleNewProjects = $kimbleModifiedProjects | ?{$_.CreatedDate -ge $cutoffDate}
#Check any other Projects for changes
#At what point does it become more efficent to dump the whole [Kimble Projects] List from SP, rather than query individual items?
#SP pages results back in 100s, so when $spProject.Count/100 -gt $kimbleChangedProjects.Count, it will take fewer requests to query each $kimbleChangedProjects individually. This ought to happen most of the time (unless there is a batch update of Projects)

<# Use this is a batch update...
$spProjects = get-itemsInList -sitePath $sitePath -listName "Kimble Projects"
foreach($kimbleChangedProject in $kimbleChangedProjects){
    $spProject = $null
    $spProject = $spProjects | ?{$_.KimbleId -eq $kimbleChangedProject.Id}
    if($spProject){
        #Check whether spProject.Title = modProject.Name and update and mark IsDirty if necessary ;PreviousName=
        #if($spProject)
        }
    else{#Project is missing from SP, so add it
        $kimbleNewProjects += $kimbleChangedProject
        }
    }
#>
#Otherwise, use this:
foreach($kimbleChangedProject in $kimbleChangedProjects){
    $kpListItem = get-itemsInList -sitePath $sitePath -listName "Kimble Projects" -oDataQuery "?&`$filter=KimbleId eq `'$($kimbleChangedProject.Id)`'"
    if($kpListItem){
        #Check whether the data has changed
        if($kpListItem.Title -ne $kimbleChangedProject.Name `
            -or $kpListItem.KimbleClientId -ne $kimbleChangedProject.KimbleOne__Account__c `
            -or $kpListItem.IsDeleted -ne $kimbleChangedProject.IsDeleted){
            #Update the entry in [Kimble Projects]
            $updateData = @{PreviousName=$kpListItem.ProjectName;PreviousKimbleClientId=$kpListItem.KimbleClientId;Title=$kimbleChangedProject.Name;KimbleClientId=$kimbleChangedProject.KimbleOne__Account__c;IsDeleted=$kimbleChangedProject.IsDeleted;IsDirty=$true}
            try{update-itemInList -serverUrl $serverUrl -sitePath $sitePath -listName "Kimble Projects" -predeterminedItemType $kp.ListItemEntityTypeFullName -itemId $kpListItem.Id -hashTableOfItemData $updateData}
            catch{$false;log-error -myError $Error[0] -myFriendlyMessage "Failed to update [Kimble Projects].$($kimbleChangedProject.Id) with $updateData"}
            }
        }
    else{$kimbleNewProjects += $kimbleChangedProject} #The Library doesn't exist, so add it to the "to-be-created" array, even though we were expecting it to exist
    }


#Add the new Projects
foreach ($kimbleNewProject in $kimbleNewProjects){
#foreach ($kimbleNewProject in $kimbleNewProjects){
    $kimbleNewProjectData = @{KimbleId=$kimbleNewProject.Id;Title=$kimbleNewProject.Name;KimbleClientId=$kimbleNewProject.KimbleOne__Account__c;IsDeleted=$kimbleNewProject.IsDeleted;IsDirty=$true}
    try{new-itemInList -sitePath $sitePath -listName "Kimble Projects" -predeterminedItemType $kp.ListItemEntityTypeFullName -hashTableOfItemData $kimbleNewProjectData}
    catch{$false;log-error -myError $Error[0] -myFriendlyMessage "Failed to create new [Kimble Projects].$($kimbleNewProject.Id) with $kimbleNewProjectData"}
    }

#endregion



<##############################
#For building the initial Sync
###############################


$spProjects = get-itemsInList -sitePath $sitePath -listName "Kimble Projects" 
$remainingKimbleProjects = $kimbleModifiedProjects | ?{($spProjects | %{$_.KimbleId}) -notcontains $_.Id}

$remainingKimbleProjects = ,@();$j=0
foreach ($c in $kimbleModifiedProjects){
    $foundIt = $false
    foreach ($createdProject in $spProjects){
        if($c.Id -eq $createdProject.KimbleId){$foundIt= $true;break}
        }
    if(!$foundIt){$remainingKimbleProjects += $c}
    $j++
    if($j%100 -eq 0){$j}
    }

foreach ($kimbleNewProject in $remainingKimbleProjects){
#foreach ($kimbleNewProject in $kimbleNewProjects){
    $kimbleNewProjectData = @{KimbleId=$kimbleNewProject.Id;Title=$kimbleNewProject.Name;IsDeleted=$kimbleNewProject.IsDeleted;IsDirty=$true}
    switch ($kimbleNewProject.Description.Length){
        0 {break}
        {$_ -lt 255} {$kimbleNewProjectData.Add("ProjectDescription","$($kimbleNewProject.Description)");break}
        default {$kimbleNewProjectData.Add("ProjectDescription","$($kimbleNewProject.Description.Substring(0,254))")}
        }
    new-itemInList -sitePath $sitePath -listName "Kimble Projects" -predeterminedItemType $kp.ListItemEntityTypeFullName -hashTableOfItemData $kimbleNewProjectData
    }

$kimbleModifiedProjects.Count
$spProjects.Count
$remainingKimbleProjects.Count

#>
Stop-Transcript