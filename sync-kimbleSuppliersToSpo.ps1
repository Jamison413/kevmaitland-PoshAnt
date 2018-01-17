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
$logfile = "C:\Scripts\Logs\sync-kimbleSuppliersToSpo_$(Get-Date -Format "yyMMdd").log"
$logErrors = $true
$logMethodMain = $true
$logFunctionCalls = $true
Set-SPORestCredentials -Credential $credential

$oAuthReqBody = Get-KimbleAuthorizationTokenWithUsernamePasswordFlowRequestBody -client_id $clientId -client_secret $clientsecret -user_name $username -pass_word $password -security_token $securityToken
try{$kimbleAccessToken=Invoke-RestMethod -Method Post -Uri $callbackUri -Body $oAuthReqBody} catch {Failure}
$kimbleRestHeaders = @{Authorization = "Bearer " + $kimbleAccessToken.access_token}


##################################
#
#Do Stuff
#
##################################

#region Kimble Sync
#Get the last Supplier modified in /Suppliers/lists/Kimble Suppliers to minimise the number of records to process
$serverUrl = "https://anthesisllc.sharepoint.com" 
$sitePath = "/Subs"
$listName = "Kimble Suppliers"
get-newDigest -serverUrl $serverUrl -sitePath $sitePath
$kc = get-list -sitePath $sitePath -listName $listName
#$lastModifiedSupplier = get-lastModifiedItemInList -sitePath $sitePath -listName "Kimble Suppliers"

#Get the Kimble Suppliers that have been modifed since the last update
$cutoffDate = (Get-Date (Get-Date $kc.LastItemModifiedDate).AddHours(-10) -Format s) #Look one hour behind just in case there is ever a delay between polling Kimble and updating SharePoint
#$cutoffDate = (Get-Date (Get-Date $kc.LastItemModifiedDate).AddYears(-1) -Format s) #Bodge this once for the initial Sync
$soqlQuery = "SELECT Name,Id,Description,Type,KimbleOne__IsCustomer__c,LastModifiedDate,SystemModStamp,CreatedDate,IsDeleted FROM account WHERE ((LastModifiedDate > $cutoffDate`Z) AND ((Is_Partner__c = TRUE) OR (Type = 'Partner') OR (Type = 'Partner/subcontractor') OR (Type = 'Supplier')))"
$kimbleModifiedSuppliers = Get-KimbleSoqlDataset -queryUri $queryUri -soqlQuery $soqlQuery -restHeaders $kimbleRestHeaders
$kimbleChangedSuppliers = $kimbleModifiedSuppliers | ?{$_.LastModifiedDate -ge $cutoffDate}
$kimbleNewSuppliers = $kimbleModifiedSuppliers | ?{$_.CreatedDate -ge $cutoffDate}

#Check any other Suppliers for changes
#At what point does it become more efficent to dump the whole [Kimble Suppliers] List from SP, rather than query individual items?
#SP pages results back in 100s, so when $spSupplier.Count/100 -gt $kimbleChangedSuppliers.Count, it will take fewer requests to query each $kimbleChangedSuppliers individually. This ought to happen most of the time (unless there is a batch update of Suppliers)

<# Use this is a batch update...
$spSuppliers = get-itemsInList -sitePath $sitePath -listName "Kimble Suppliers"

#foreach($kimbleChangedSupplier in $kimbleChangedSuppliers){
#$kimbleNewSuppliers = @()
for($j=0; $j -lt $kimbleChangedSuppliers.Count;$j++){
    $spSupplier = $null
    if ($j -lt $kimbleChangedSuppliers.Count/2){
        for ($i=0 ; $i -lt $spSuppliers.Count;$i++){
            if ($spSupplier -ne $null){break}
            else{if($spSuppliers[$i].KimbleId -eq $kimbleChangedSuppliers[$j].Id){$spSupplier = $spSuppliers[$i]}}
            }
        }
    else{
        for ($i=$spSuppliers.Count-1 ; $i -ge 0;$i--){
            if ($spSupplier -ne $null){break}
            else{if($spSuppliers[$i].KimbleId -eq $kimbleChangedSuppliers[$j].Id){$spSupplier = $spSuppliers[$i]}}
            }
        }
    if($spSupplier -eq $null){$kimbleNewSuppliers += $kimbleChangedSuppliers[$j]}
    if($j%100 -eq 0){Write-Host -ForegroundColor Magenta "$j / $($kimbleChangedSuppliers.Count) _ $($kimbleNewSuppliers.Count)"}
    }

#>
#Otherwise, use this:
foreach($kimbleChangedSupplier in $kimbleChangedSuppliers){
    $kCListItem = get-itemsInList -sitePath $sitePath -listName "Kimble Suppliers" -oDataQuery "?&`$filter=KimbleId eq `'$($kimbleChangedSupplier.Id)`'"
    if($kCListItem){
        #Check whether the data has changed
        if($kCListItem.Title -ne $kimbleChangedSupplier.Name `
            -or $kCListItem.SupplierDescription -ne $kimbleChangedSupplier.Description `
            -or $kCListItem.IsDeleted -ne $kimbleChangedSupplier.IsDeleted){
            #Update the entry in [Kimble Suppliers]
            $updateData = @{PreviousName=$kCListItem.SupplierName;PreviousDescription=$kCListItem.SupplierDescription;Title=$kimbleChangedSupplier.Name;SupplierDescription=$kimbleChangedSupplier.Description;IsDeleted=$kimbleChangedSupplier.IsDeleted;IsDirty=$true}
            try{update-itemInList -serverUrl $serverUrl -sitePath $sitePath -listNameOrGuid "Kimble Suppliers" -predeterminedItemType $kc.ListItemEntityTypeFullName -itemId $kCListItem.Id -hashTableOfItemData $updateData}
            catch{$false;log-error -myError $Error[0] -myFriendlyMessage "Failed to update [Kimble Suppliers].$($kimbleChangedSupplier.Id) with $updateData"}
            }
        }
    else{$kimbleNewSuppliers += $kimbleChangedSupplier} #The Library doesn't exist, so add it to the "to-be-created" array, even though we were expecting it to exist
    }


#Add the new Suppliers
foreach ($kimbleNewSupplier in $kimbleNewSuppliers){
    $kimbleNewSupplierData = @{KimbleId=$kimbleNewSupplier.Id;Title=$kimbleNewSupplier.Name;IsDeleted=$kimbleNewSupplier.IsDeleted;IsDirty=$true}
    if($kimbleNewSupplier.Description){$kimbleNewSupplierData.Add("SupplierDescription","$($kimbleNewSupplier.Description)")}
    try{new-itemInList -sitePath $sitePath -listName "Kimble Suppliers" -predeterminedItemType $kc.ListItemEntityTypeFullName -hashTableOfItemData $kimbleNewSupplierData}
    catch{$false;log-error -myError $Error[0] -myFriendlyMessage "Failed to create new [Kimble Suppliers].$($kimbleNewSupplier.Id) with $kimbleNewSupplierData"}
    }


#endregion



<##############################
#For building the initial Sync
###############################


$spSuppliers = get-itemsInList -sitePath $sitePath -listName "Kimble Suppliers" 
$remainingKimbleSuppliers = $kimbleModifiedSuppliers | ?{($spSuppliers | %{$_.KimbleId}) -notcontains $_.Id}

$remainingKimbleSuppliers = ,@();$j=0
foreach ($c in $kimbleModifiedSuppliers){
    $foundIt = $false
    foreach ($createdSupplier in $spSuppliers){
        if($c.Id -eq $createdSupplier.KimbleId){$foundIt= $true;break}
        }
    if(!$foundIt){$remainingKimbleSuppliers += $c}
    $j++
    if($j%100 -eq 0){$j}
    }

foreach ($kimbleNewSupplier in $remainingKimbleSuppliers){
#foreach ($kimbleNewSupplier in $kimbleNewSuppliers){
    $kimbleNewSupplierData = @{KimbleId=$kimbleNewSupplier.Id;Title=$kimbleNewSupplier.Name;IsDeleted=$kimbleNewSupplier.IsDeleted;IsDirty=$true}
    switch ($kimbleNewSupplier.Description.Length){
        0 {break}
        {$_ -lt 255} {$kimbleNewSupplierData.Add("SupplierDescription","$($kimbleNewSupplier.Description)");break}
        default {$kimbleNewSupplierData.Add("SupplierDescription","$($kimbleNewSupplier.Description.Substring(0,254))")}
        }
    new-itemInList -sitePath $sitePath -listName "Kimble Suppliers" -predeterminedItemType $kc.ListItemEntityTypeFullName -hashTableOfItemData $kimbleNewSupplierData
    }

$kimbleModifiedSuppliers.Count
$spSuppliers.Count
$remainingKimbleSuppliers.Count

#>
Stop-Transcript
#$kimbleModifiedSuppliers | ?{$_.Name -match "Link"} | Select Name, Type, KimbleOne__IsCustomer__c | sort Name
#$kimbleModifiedSuppliers | Select Name, Type, KimbleOne__IsCustomer__c | sort Name
#$kimbleModifiedSuppliers.Count | ?{$_.Name -match "Linked"}
