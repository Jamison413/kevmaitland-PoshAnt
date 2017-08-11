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
#Get the last Lead modified in /lists/Kimble Leads to minimise the number of records to process
$serverUrl = "https://anthesisllc.sharepoint.com" 
$sitePath = "/clients"
$listName = "Kimble Leads"
get-newDigest -serverUrl $serverUrl -sitePath $sitePath
$kp = get-list -sitePath $sitePath -listName $listName

#Get the Kimble Leads that have been modifed since the last update
$cutoffDate = (Get-Date (Get-Date $kp.LastItemModifiedDate).AddHours(-1) -Format s) #Look one hour behind just in case there is ever a delay between polling Kimble and updating SharePoint
#$cutoffDate = (Get-Date (Get-Date $kp.LastItemModifiedDate).AddYears(-1) -Format s) #Bodge this once for the initial Sync
$soqlQuery = "SELECT Name,Id,KimbleOne__Account__c,LastModifiedDate,SystemModStamp,CreatedDate,IsDeleted FROM KimbleOne__SalesOpportunity__c WHERE LastModifiedDate > $cutoffDate`Z"

$kimbleModifiedLeads = Get-KimbleSoqlDataset -queryUri $queryUri -soqlQuery $soqlQuery -restHeaders $kimbleRestHeaders
$kimbleChangedLeads = $kimbleModifiedLeads | ?{$_.LastModifiedDate -lt $cutoffDate}
$kimbleNewLeads = $kimbleModifiedLeads | ?{$_.CreatedDate -ge $cutoffDate}
#Check any other Leads for changes
#At what point does it become more efficent to dump the whole [Kimble Leads] List from SP, rather than query individual items?
#SP pages results back in 100s, so when $spLead.Count/100 -gt $kimbleChangedLeads.Count, it will take fewer requests to query each $kimbleChangedLeads individually. This ought to happen most of the time (unless there is a batch update of Leads)

<# Use this is a batch update...
$spLeads = get-itemsInList -sitePath $sitePath -listName "Kimble Leads"
foreach($kimbleChangedLead in $kimbleChangedLeads){
    $spLead = $null
    $spLead = $spLeads | ?{$_.KimbleId -eq $kimbleChangedLead.Id}
    if($spLead){
        #Check whether spLead.Title = modLead.Name and update and mark IsDirty if necessary ;PreviousName=
        #if($spLead)
        }
    else{#Lead is missing from SP, so add it
        $kimbleNewLeads += $kimbleChangedLead
        }
    }
#>
#Otherwise, use this:
foreach($kimbleChangedLead in $kimbleChangedLeads){
    $kpListItem = get-itemsInList -sitePath $sitePath -listName "Kimble Leads" -oDataQuery "?&`$filter=KimbleId eq `'$($kimbleChangedLead.Id)`'"
    if($kpListItem){
        #Check whether the data has changed
        if($kpListItem.Title -ne $kimbleChangedLead.Name `
            -or $kpListItem.KimbleClientId -ne $kimbleChangedLead.KimbleOne__Account__c `
            -or $kpListItem.IsDeleted -ne $kimbleChangedLead.IsDeleted){
            #Update the entry in [Kimble Leads]
            $updateData = @{PreviousName=$kpListItem.LeadName;PreviousKimbleClientId=$kpListItem.KimbleClientId;Title=$kimbleChangedLead.Name;KimbleClientId=$kimbleChangedLead.KimbleOne__Account__c;IsDeleted=$kimbleChangedLead.IsDeleted;IsDirty=$true}
            try{update-itemInList -serverUrl $serverUrl -sitePath $sitePath -listName "Kimble Leads" -predeterminedItemType $kp.ListItemEntityTypeFullName -itemId $kpListItem.Id -hashTableOfItemData $updateData}
            catch{$false;log-error -myError $Error[0] -myFriendlyMessage "Failed to update [Kimble Leads].$($kimbleChangedLead.Id) with $updateData"}
            }
        }
    else{$kimbleNewLeads += $kimbleChangedLead} #The Library doesn't exist, so add it to the "to-be-created" array, even though we were expecting it to exist
    }


#Add the new Leads
foreach ($kimbleNewLead in $kimbleNewLeads){
#foreach ($kimbleNewLead in $kimbleNewLeads){
    $kimbleNewLeadData = @{KimbleId=$kimbleNewLead.Id;Title=$kimbleNewLead.Name;KimbleClientId=$kimbleNewLead.KimbleOne__Account__c;IsDeleted=$kimbleNewLead.IsDeleted;IsDirty=$true}
    try{new-itemInList -sitePath $sitePath -listName "Kimble Leads" -predeterminedItemType $kp.ListItemEntityTypeFullName -hashTableOfItemData $kimbleNewLeadData}
    catch{$false;log-error -myError $Error[0] -myFriendlyMessage "Failed to create new [Kimble Leads].$($kimbleNewLead.Id) with $kimbleNewLeadData"}
    }

#endregion



<##############################
#For building the initial Sync
###############################


$spLeads = get-itemsInList -sitePath $sitePath -listName "Kimble Leads" 
$remainingKimbleLeads = $kimbleModifiedLeads | ?{($spLeads | %{$_.KimbleId}) -notcontains $_.Id}

$remainingKimbleLeads = ,@();$j=0
foreach ($c in $kimbleModifiedLeads){
    $foundIt = $false
    foreach ($createdLead in $spLeads){
        if($c.Id -eq $createdLead.KimbleId){$foundIt= $true;break}
        }
    if(!$foundIt){$remainingKimbleLeads += $c}
    $j++
    if($j%100 -eq 0){$j}
    }

foreach ($kimbleNewLead in $remainingKimbleLeads){
#foreach ($kimbleNewLead in $kimbleNewLeads){
    $kimbleNewLeadData = @{KimbleId=$kimbleNewLead.Id;Title=$kimbleNewLead.Name;IsDeleted=$kimbleNewLead.IsDeleted;IsDirty=$true}
    switch ($kimbleNewLead.Description.Length){
        0 {break}
        {$_ -lt 255} {$kimbleNewLeadData.Add("LeadDescription","$($kimbleNewLead.Description)");break}
        default {$kimbleNewLeadData.Add("LeadDescription","$($kimbleNewLead.Description.Substring(0,254))")}
        }
    new-itemInList -sitePath $sitePath -listName "Kimble Leads" -predeterminedItemType $kp.ListItemEntityTypeFullName -hashTableOfItemData $kimbleNewLeadData
    }

$kimbleModifiedLeads.Count
$spLeads.Count
$remainingKimbleLeads.Count

#>
Stop-Transcript