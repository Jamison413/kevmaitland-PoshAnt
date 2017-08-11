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

$logfile = "C:\Scripts\Logs\sync-kimbleClientsToSpo_$(Get-Date -Format "yyMMdd").log"
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
#Get the last Client modified in /clients/lists/Kimble Clients to minimise the number of records to process
$serverUrl = "https://anthesisllc.sharepoint.com" 
$sitePath = "/clients"
$listName = "Kimble Clients"
get-newDigest -serverUrl $serverUrl -sitePath $sitePath
$kc = get-list -sitePath $sitePath -listName $listName
#$lastModifiedClient = get-lastModifiedItemInList -sitePath $sitePath -listName "Kimble Clients"

#Get the Kimble Clients that have been modifed since the last update
$cutoffDate = (Get-Date (Get-Date $kc.LastItemModifiedDate).AddHours(-1) -Format s) #Look one hour behind just in case there is ever a delay between polling Kimble and updating SharePoint
#$cutoffDate = (Get-Date (Get-Date $kc.LastItemModifiedDate).AddYears(-1) -Format s) #Bodge this once for the initial Sync
$soqlQuery = "SELECT Name,Id,Description,Type,KimbleOne__IsCustomer__c,LastModifiedDate,SystemModStamp,CreatedDate,IsDeleted FROM account WHERE ((LastModifiedDate > $cutoffDate`Z) AND ((KimbleOne__IsCustomer__c = TRUE) OR (Type = 'Client')))"
$kimbleModifiedClients = Get-KimbleSoqlDataset -queryUri $queryUri -soqlQuery $soqlQuery -restHeaders $kimbleRestHeaders
$kimbleChangedClients = $kimbleModifiedClients | ?{$_.LastModifiedDate -ge $cutoffDate}
$kimbleNewClients = $kimbleModifiedClients | ?{$_.CreatedDate -ge $cutoffDate}

#Check any other Clients for changes
#At what point does it become more efficent to dump the whole [Kimble Clients] List from SP, rather than query individual items?
#SP pages results back in 100s, so when $spClient.Count/100 -gt $kimbleChangedClients.Count, it will take fewer requests to query each $kimbleChangedClients individually. This ought to happen most of the time (unless there is a batch update of Clients)

<# Use this is a batch update...
$spClients = get-itemsInList -sitePath $sitePath -listName "Kimble Clients"

#foreach($kimbleChangedClient in $kimbleChangedClients){
#$kimbleNewClients = @()
for($j=0; $j -lt $kimbleChangedClients.Count;$j++){
    $spClient = $null
    if ($j -lt $kimbleChangedClients.Count/2){
        for ($i=0 ; $i -lt $spClients.Count;$i++){
            if ($spClient -ne $null){break}
            else{if($spClients[$i].KimbleId -eq $kimbleChangedClients[$j].Id){$spClient = $spClients[$i]}}
            }
        }
    else{
        for ($i=$spClients.Count-1 ; $i -ge 0;$i--){
            if ($spClient -ne $null){break}
            else{if($spClients[$i].KimbleId -eq $kimbleChangedClients[$j].Id){$spClient = $spClients[$i]}}
            }
        }
    if($spClient -eq $null){$kimbleNewClients += $kimbleChangedClients[$j]}
    if($j%100 -eq 0){Write-Host -ForegroundColor Magenta "$j / $($kimbleChangedClients.Count) _ $($kimbleNewClients.Count)"}
    }

#>
#Otherwise, use this:
foreach($kimbleChangedClient in $kimbleChangedClients){
    $kCListItem = get-itemsInList -sitePath $sitePath -listName "Kimble Clients" -oDataQuery "?&`$filter=KimbleId eq `'$($kimbleChangedClient.Id)`'"
    if($kCListItem){
        #Check whether the data has changed
        if($kCListItem.Title -ne $kimbleChangedClient.Name `
            -or $kCListItem.ClientDescription -ne $kimbleChangedClient.Description `
            -or $kCListItem.IsDeleted -ne $kimbleChangedClient.IsDeleted){
            #Update the entry in [Kimble Clients]
            $updateData = @{PreviousName=$kCListItem.ClientName;PreviousDescription=$kCListItem.ClientDescription;Title=$kimbleChangedClient.Name;ClientDescription=$kimbleChangedClient.Description;IsDeleted=$kimbleChangedClient.IsDeleted;IsDirty=$true}
            try{update-itemInList -serverUrl $serverUrl -sitePath $sitePath -listName "Kimble Clients" -predeterminedItemType $kc.ListItemEntityTypeFullName -itemId $kCListItem.Id -hashTableOfItemData $updateData}
            catch{$false;log-error -myError $Error[0] -myFriendlyMessage "Failed to update [Kimble Clients].$($kimbleChangedClient.Id) with $updateData"}
            }
        }
    else{$kimbleNewClients += $kimbleChangedClient} #The Library doesn't exist, so add it to the "to-be-created" array, even though we were expecting it to exist
    }


#Add the new Clients to the SPO List
foreach ($kimbleNewClient in $kimbleNewClients){
    $kimbleNewClientData = @{KimbleId=$kimbleNewClient.Id;Title=$kimbleNewClient.Name;IsDeleted=$kimbleNewClient.IsDeleted;IsDirty=$true}
    if($kimbleNewClient.Description){$kimbleNewClientData.Add("ClientDescription","$($kimbleNewClient.Description)")}
    try{new-itemInList -sitePath $sitePath -listName "Kimble Clients" -predeterminedItemType $kc.ListItemEntityTypeFullName -hashTableOfItemData $kimbleNewClientData}
    catch{$false;log-error -myError $Error[0] -myFriendlyMessage "Failed to create new [Kimble Clients].$($kimbleNewClient.Id) with $kimbleNewClientData"}
    }


#endregion



<##############################
#For building the initial Sync
###############################


$spClients = get-itemsInList -sitePath $sitePath -listName "Kimble Clients" 
$remainingKimbleClients = $kimbleModifiedClients | ?{($spClients | %{$_.KimbleId}) -notcontains $_.Id}

$remainingKimbleClients = ,@();$j=0
foreach ($c in $kimbleModifiedClients){
    $foundIt = $false
    foreach ($createdClient in $spClients){
        if($c.Id -eq $createdClient.KimbleId){$foundIt= $true;break}
        }
    if(!$foundIt){$remainingKimbleClients += $c}
    $j++
    if($j%100 -eq 0){$j}
    }

foreach ($kimbleNewClient in $remainingKimbleClients){
#foreach ($kimbleNewClient in $kimbleNewClients){
    $kimbleNewClientData = @{KimbleId=$kimbleNewClient.Id;Title=$kimbleNewClient.Name;IsDeleted=$kimbleNewClient.IsDeleted;IsDirty=$true}
    switch ($kimbleNewClient.Description.Length){
        0 {break}
        {$_ -lt 255} {$kimbleNewClientData.Add("ClientDescription","$($kimbleNewClient.Description)");break}
        default {$kimbleNewClientData.Add("ClientDescription","$($kimbleNewClient.Description.Substring(0,254))")}
        }
    new-itemInList -sitePath $sitePath -listName "Kimble Clients" -predeterminedItemType $kc.ListItemEntityTypeFullName -hashTableOfItemData $kimbleNewClientData
    }

$kimbleModifiedClients.Count
$spClients.Count
$remainingKimbleClients.Count

#>
Stop-Transcript
#$kimbleModifiedClients | ?{$_.Name -match "Link"} | Select Name, Type, KimbleOne__IsCustomer__c | sort Name
#$kimbleModifiedClients | Select Name, Type, KimbleOne__IsCustomer__c | sort Name
#$kimbleModifiedClients.Count | ?{$_.Name -match "Linked"}
