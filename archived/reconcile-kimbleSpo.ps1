﻿param(
    # Specifies whether we are updating Clients or Suppliers.
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Clients", "Suppliers","Projects")]
    [string]$objectType 
    )

$logFileLocation = "C:\ScriptLogs\"
if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)){
    $fullLogPathAndName = $logFileLocation+"reconcile-kimbleSpo_$objectType`_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = $logFileLocation+"reconcile-kimbleSpo_$objectType`_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
else{
    $fullLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_$objectType`_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_$objectType`_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
if($PSCommandPath){
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_Transcript_$(Get-Date -Format "yyMMdd").log"
    Start-Transcript $transcriptLogName -Append
    }

Import-Module _PS_Library_GeneralFunctionality
Import-Module _CSOM_Library-SPO.psm1
Import-Module _REST_Library-Kimble.psm1
Import-Module _REST_Library-SPO.psm1
Import-Module SharePointPnPPowerShellOnline
Import-Module _PNP_Library_SPO

#region Variables
##################################
#
#Get ready
#
########################################
$webUrl = "https://anthesisllc.sharepoint.com" 
$smtpServer = "anthesisgroup-com.mail.protection.outlook.com"
$mailFrom = "scriptrobot@sustain.co.uk"
$mailTo = "kevin.maitland@anthesisgroup.com"
$cacheFilePath = "$env:USERPROFILE\KimbleCache\"

#convertTo-localisedSecureString ""
$sharePointAdmin = "kimblebot@anthesisgroup.com"
#$sharePointAdminPass = ConvertTo-SecureString -String '' -AsPlainText -Force | ConvertFrom-SecureString
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Downloads\KimbleBot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
#$restCreds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($adminCreds.UserName,$adminCreds.Password)
#new-spoCred  -username $adminCreds.UserName -securePassword $adminCreds.Password
#$csomCreds = new-csomCredentials -username $adminCreds.UserName -password $adminCreds.Password
########################################
$kimbleCreds = Import-Csv "$env:USERPROFILE\Desktop\Kimble.txt"
$standardKimbleHeaders = get-kimbleHeaders -clientId $kimbleCreds.clientId -clientSecret $kimbleCreds.clientSecret -username $kimbleCreds.username -password $kimbleCreds.password -securityToken $kimbleCreds.securityToken -connectToLiveContext $true -verboseLogging $true
$standardKimbleQueryUri = get-kimbleQueryUri
#endregion

#region Functions
function try-newListItem($webUrl, $sitePath, $newSpoItemData, $spoListToAddTo, $restCreds, $clientsDigest, $fullLogPathAndName){
    try{
        log-action -myMessage "Creating new SPO List item [$($newSpoItemData["Title"])] in [$($spoListToAddTo.Title)]" -logFile $fullLogPathAndName
        $newItem = new-itemInList -serverUrl $webUrl -sitePath $sitePath -listName $spoListToAddTo.Title -predeterminedItemType $spoListToAddTo.ListItemEntityTypeFullName -hashTableOfItemData $newSpoItemData -restCreds $restCreds -digest $clientsDigest -verboseLogging $true -logFile $fullLogPathAndName
        #Check it's worked
        if($newItem){log-result -myMessage "SUCCESS: SPO [$($spoListToAddTo.Title)] item $($newItem.Title) created!" -logFile $fullLogPathAndName}
        else{
            log-result -myMessage "FAILED: SPO [$($spoListToAddTo.Title)] item $($newSpoItemData["Title"]) did not create!" -logFile $fullLogPathAndName
            #Bodge this with an e-mail alert as we don't want Projects going missing
            Send-MailMessage -SmtpServer $smtpServer -To $mailTo -From $mailFrom -Subject "[$($spoListToAddTo.Title)].[$($newSpoItemData["Title"])] could not be created in SPO" -Body "[$($spoListToAddTo.Title)]: $($newSpoItemData["KimbleId"])"
            }
        }
    catch{log-error -myError $_ -myFriendlyMessage "Failed to create new [Kimble Leads].$($kimbleLeadObject.Name) with @{$($($newSpoLeadData.Keys | % {$_+":"+$newSpoLeadData[$_]+","}) -join "`r")}" -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -smtpServer $smtpServer -mailTo $mailTo -mailFrom $mailFrom}
    
    }
function reconcile-objectsBetweenKimbleAndSpo($accountType, $standardKimbleQueryUri, $standardKimbleHeaders, $webUrl, $sitePath, $adminCreds, $verboseLogging){
    if($accountType -imatch "Client"){
        $sitePath = "/clients"
        $listName = "Kimble Clients"
        $cacheFileName = "kimbleClients.csv"
        #$soqlWhereStatement = "WHERE ((KimbleOne__IsCustomer__c = TRUE) OR (Type = 'Client') OR (Type = 'Potential Client'))"
        }
    elseif($accountType -imatch "Project"){
        $sitePath = "/clients"
        $listName = "Kimble Projects"
        $cacheFileName = "kimbleProjects.csv"
        }
    elseif($accountType -imatch "Supplier"){
        $sitePath = "/subs"
        $listName = "Kimble Suppliers"
        $cacheFileName = "kimbleSuppliers.csv"
        #$soqlWhereStatement = "WHERE ((Is_Partner__c = TRUE) OR (Type = 'Partner') OR (Type = 'Partner/subcontractor') OR (Type = 'Supplier'))"
        }
    else{break}
    #Get the full list of Kimble Clients
    $allKimbleAccountObjects = get-allKimbleAccounts -pQueryUri $standardKimbleQueryUri -pRestHeaders $standardKimbleHeaders -verboseLogging $verboseLogging
    #Get the full list of SPO Clients
    Connect-PnPOnline –Url $($webUrl+$sitePath) –Credentials $adminCreds

    $pnpList = Get-PnPList -Identity $listName -Includes ContentTypes #We need ContentTypes to create New Objects
    #$pnpAccountListItems = get-spoKimbleAccountListItems -pnpList $pnpList -spoCredentials $adminCreds -verboseLogging $verboseLogging

    #Retrieve (and update if necessary) the full Clients Cache as we'll need it to set up any new Leads/Projects
    $accountsCache = cache-spoKimbleAccountsList -pnpList $pnpList -kimbleListCachePathAndFileName $($cacheFilePath+$cacheFileName)

    
    #Check for uncreated spoAccountListItems 
    $missingAccounts = Compare-Object -ReferenceObject $allKimbleAccountObjects -DifferenceObject $accountsCache -Property "Id" -PassThru -CaseSensitive:$false -IncludeEqual
    $missingSpoAccounts = $missingAccounts | ?{$_.SideIndicator -eq "<="}

    $missingSpoAccounts | %{
        $kimbleAccountMissingFromSpo = $_
        log-action -myMessage "Creating missing $accountType [$($kimbleAccountMissingFromSpo.Name)]" -logFile $fullLogPathAndName
        $newSpoAccount = new-spoKimbleObjectListItem -kimbleObject $kimbleAccountMissingFromSpo -pnpKimbleObjectList $pnpList -fullLogPathAndName $fullLogPathAndName  -verboseLogging $verboseLogging
        if($newSpoAccount){log-result "SUCCESS: New $accountType [$($kimbleAccountMissingFromSpo.Name)] created in [$listName]" -logFile $fullLogPathAndName}
        else{log-result "FAILED: New $accountType [$($kimbleAccountMissingFromSpo.Name)] NOT created in [$listName]" -logFile $fullLogPathAndName}
        }

    #Check for orphaned spoAccountListItems, flag them as IsOprhaned=$true and isDirty=$false
    $missingKimbleAccounts = $missingAccounts | ?{$_.SideIndicator -eq "=>"}
    $missingKimbleAccounts | % {
        $orphanedSpoAccount = $_
        log-action -myMessage "Marking orphaned $accountType [$($orphanedSpoAccount. Name)]" -logFile $fullLogPathAndName
        #$updatedOrphanedAccount = update-spoKimbleObjectListItem -kimbleObject $orphanedSpoAccount -pnpKimbleObjectList $pnpList -overrideIsDirtyFalse $true -overrideIsOrphanedTrue $true -fullLogPathAndName $fullLogPathAndName  -verboseLogging $verboseLogging
        $updatedOrphanedAccount = update-spoKimbleObjectListItem -kimbleObject $orphanedSpoAccount -pnpKimbleObjectList $pnpList -overrideIsOrphanedTrue $true -fullLogPathAndName $fullLogPathAndName  -verboseLogging $verboseLogging #Removed isDirty Override now (I think) the excessive IsDirty flagging problem is resolved
        if($updatedOrphanedAccount.FieldValues.IsOrphaned -and !($updatedOrphanedAccount.FieldValues.IsDirty)){log-result "SUCCESS: Updated $accountType [$($updatedOrphanedAccount.FieldValues.Title)] marked as isOrphaned:`$true isDirty:`$false" -logFile $fullLogPathAndName}
        else{log-result "FAILED: Updated $accountType [$($orphanedSpoAccount.Name)] NOT updated in [$listName]" -logFile $fullLogPathAndName}
        }

    #Check for misclassified spoAccountListItems, flag them as IsMisclassified=$true and isDirty=$false
    $validKimbleAccounts = $missingAccounts | ?{$_.SideIndicator -eq "=="}
    if($accountType -imatch "Client"){
        $thisTypeOfKimbleAccounts = $validKimbleAccounts | ? {$_.KimbleOne__IsCustomer__c -eq $true -or $_.Type -eq "Client" -or $_.Type -eq "Potential Client"}
        }
    elseif($accountType -imatch "Supplier"){
        $thisTypeOfKimbleAccounts = $validKimbleAccounts | ? {$_.Is_Partner__c -eq $true -or $_.Type -eq "Partner" -or $_.Type -eq "Partner/subcontractor" -or $_.Type -eq "Supplier"}
        }

    $misclassifiedAccounts = Compare-Object -ReferenceObject $thisTypeOfKimbleAccounts -DifferenceObject $accountsCache -Property "Id" -PassThru -CaseSensitive:$false #This gives us all spoAccounts that are not valid Clients/Supplier
    $incorrectlyLablledMisclassifiedAccounts = $misclassifiedAccounts | ? {$_.isMisclassified -eq $false -or $_.IsDirty -eq $true} #This gives us all the ones that need processing
    $i = 1
    $incorrectlyLablledMisclassifiedAccounts | % {
        Write-Progress -Id 1000 -Status "Processing Misclassified Accounts" -Activity "$i/$($incorrectlyLablledMisclassifiedAccounts.Count)" -PercentComplete ($i*100/$incorrectlyLablledMisclassifiedAccounts.Count) #Display the overall progress
        $misclassifiedAccount = $_
        log-action -myMessage "Marking misclassified $accountType [$($misclassifiedAccount.Name)]" -logFile $fullLogPathAndName
        $updatedMisclassifiedAccount = update-spoKimbleObjectListItem -kimbleObject $misclassifiedAccount -pnpKimbleObjectList $pnpList -overrideIsDirtyFalse $true -overrideIsMisclassified $true -overrideIsOrphanedFalse $true -fullLogPathAndName $fullLogPathAndName  -verboseLogging $verboseLogging
        if($updatedMisclassifiedAccount.FieldValues.isMisclassified -and !($updatedMisclassifiedAccount.FieldValues.IsDirty)){log-result "SUCCESS: Updated $accountType [$($updatedMisclassifiedAccount.FieldValues.Title)] marked as isMisclassified:`$true isDirty:`$false" -logFile $fullLogPathAndName}
        else{log-result "FAILED: Updated $accountType [$($misclassifiedAccount.Name)] NOT updated in [$listName]" -logFile $fullLogPathAndName}
        $i++
        }
    }
function reconcile-clientsBetweenKimbleAndSpo($standardKimbleQueryUri, $standardKimbleHeaders, $webUrl, $sitePath, $adminCreds){
    #Get the full list of Kimble Clients
    $allKimbleClients = get-allKimbleAccounts -pQueryUri $standardKimbleQueryUri -pRestHeaders $standardKimbleHeaders -pWhereStatement "WHERE ((KimbleOne__IsCustomer__c = TRUE) OR (Type = 'Client') OR (Type = 'Potential Client'))" 
    #Get the full list of SPO Clients
    Connect-PnPOnline –Url $($webUrl+$sitePath) –Credentials $adminCreds
    $clientList = Get-PnPList -Identity "Kimble Clients" -Includes ContentTypes
    $clientListContentType = $clientList.ContentTypes | ? {$_.Name -eq "Item"}
    $clientListItems2 = Get-PnPListItem -List "Kimble Clients" -PageSize 5000 -Fields "Title","GUID","KimbleId","ClientDescription","IsDirty","IsDeleted","Modified","LastModifiedDate","PreviousName","PreviousDescription","Id"
    $clientListItems2.FieldValues | %{
        $thisClient = $_
        [array]$allSpoClients += New-Object psobject -Property $([ordered]@{"Id"=$thisClient["KimbleId"];"Name"=$thisClient["Title"];"GUID"=$thisClient["GUID"];"SPListItemID"=$thisClient["ID"];"IsDirty"=$thisClient["IsDirty"];"IsDeleted"=$thisClient["IsDeleted"];"LastModifiedDate"=$thisClient["LastModifiedDate"];"PreviousName"=$thisClient["PreviousName"];"PreviousDescription"=$thisClient["PreviousDescription"]})
        }

    $missingSpoClients = Compare-Object -ReferenceObject $allKimbleClients -DifferenceObject $allSpoClients -Property "Id" -PassThru -CaseSensitive:$false
    
    $missingSpoClients | ?{$_.SideIndicator -eq "<="} | %{
        $missingClient = $_
        log-action -myMessage "Creating missing client [$($missingClient.Name)]" -logFile $fullLogPathAndName
        $newClient = new-spoKimbleClientItem -kimbleClientObject $missingClient -pnpClientList $clientList -fullLogPathAndName $fullLogPathAndName  -verboseLogging $verboseLogging
        if($newClient){log-result "SUCCESS: New Client [$($missingClient.Name)] created in [Kimble Clients]" -logFile $fullLogPathAndName}
        else{log-result "FAILED: New Client [$($missingClient.Name)] NOT created in [Kimble Clients]" -logFile $fullLogPathAndName}
        #$newItem = Add-PnPListItem -List $clientList.Id -ContentType $clientListContentType.Id.StringValue -Values @{"Title"=$missingClient.Name;"KimbleId"=$missingClient.Id;"ClientDescription"=$missingClient.Description;"IsDirty"=$true;"IsDeleted"=$missingClient.IsDeleted;"LastModifiedDate"=$(Get-Date $missingClient.LastModifiedDate -Format "MM/dd/yyyy hh:mm")}
        }

    $updatedKimbleClients = Compare-Object -ReferenceObject $allKimbleClients -DifferenceObject $allSpoClients -Property @("Id","LastModifiedDate") -PassThru -CaseSensitive:$false
    $updatedKimbleClients | ?{$_.SideIndicator -eq "<="} | % {
        $updatedClient = $_
        $fixedClient = update-spoKimbleClientItem -kimbleClientObject $updatedClient -pnpClientList $clientList -fullLogPathAndName $fullLogPathAndName -verboseLogging $verboseLogging
<#        $spoClient = $allSpoClients | ? {$_.Id -eq $updatedClient.Id}
        $updatedValues = @{"IsDeleted"=$updatedClient.IsDeleted}
        if($updatedClient.LastModifiedDate -ne $null){
            $updatedValues.Add("LastModifiedDate",$(Get-Date $updatedClient.LastModifiedDate -Format "yyyy/MM/dd hh:mm:ss"))
            }
        if($updatedClient.Name -ne $spoClient.Name){
            $updatedValues.Add("Title",$updatedClient.Name)
            $updatedValues.Add("PreviousName",$spoClient.Name)
            $updatedValues.Add("IsDirty",$true)
            #Write-Host "$($spoClient.Name) renamed to $($updatedClient.Name)"
            $testName = $updatedValues
            }
        if((sanitise-stripHtml $updatedClient.Description) -ne $(sanitise-stripHtml $spoClient.Description)){
            $updatedValues.Add("ClientDescription",$(sanitise-stripHtml $updatedClient.Description))
            $updatedValues.Add("PreviousDescription", $spoClient.Description)
            #Write-Host "$($spoClient.Name) description change from $($spoClient.Description) to $($updatedClient.des)"
            if($updatedValues.Keys -notcontains "IsDirty"){$($updatedValues.Add("IsDirty",$true))}
            $testDesc = $updatedValues
            }
        Set-PnPListItem -List $clientList.Id -Identity $spoClient.SPListItemID -Values $updatedValues
        #Re-run for borked ClientDescription field
        <#if($updatedValues.Keys -contains "ClientDescription"){
            Write-Host -ForegroundColor Yellow "Updating $($spoClient.Name)"
            Set-PnPListItem -List $clientList.Id -Identity $spoClient.SPListItemID -Values $updatedValues
            }#>
        }
    }
function reconcile-leadsBetweenKimbleAndSpo($standardKimbleQueryUri, $standardKimbleHeaders, $webUrl, $sitePath, $adminCreds){
    #Get the full list of Kimble Leads
    $allKimbleLeads = get-allKimbleLeads -pQueryUri $standardKimbleQueryUri -pRestHeaders $standardKimbleHeaders
    
    Connect-PnPOnline –Url $($webUrl+$sitePath) –Credentials $adminCreds
    $leadList = Get-PnPList -Identity "Kimble Leads" -Includes ContentTypes
    $leadListContentType = $leadList.ContentTypes | ? {$_.Name -eq "Item"}
    $leadListItems = Get-PnPListItem -List "Kimble Leads" -PageSize 1000 -Fields "Title","GUID","KimbleId","leadDescription","KimbleClientId","PreviousKimbleClientId","IsDirty","IsDeleted","Modified","LastModifiedDate","PreviousName","PreviousDescription","Id"
    $leadListItems.FieldValues | %{
        $thisLead = $_
        [array]$allSpoLeads += New-Object psobject -Property $([ordered]@{"Id"=$thisLead["KimbleId"];"KimbleId"=$thisLead["KimbleId"];"KimbleClientId"=$thisLead["KimbleClientId"];"PreviousKimbleClientId"=$thisLead["PreviousKimbleClientId"];"Name"=$thisLead["Title"];"GUID"=$thisLead["GUID"];"SPListItemID"=$thisLead["ID"];"IsDirty"=$thisLead["IsDirty"];"IsDeleted"=$thisLead["IsDeleted"];"LastModifiedDate"=$thisLead["LastModifiedDate"];"PreviousName"=$thisLead["PreviousName"];"PreviousDescription"=$thisLead["PreviousDescription"]})
        }
    
    #Work out what's missing and create any omissions
    $allKimbleLeads | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name KimbleId -Value $_.Id}
    $missingLeads = Compare-Object -ReferenceObject $allKimbleLeads -DifferenceObject $allSpoLeads -Property "KimbleId" -PassThru -CaseSensitive:$false
    $missingLeads | % {
        if ($_.SideIndicator -eq "<="){new-spoLead -kimbleLeadObject $_ -pnpLeadsList $leadList -fullLogPathAndName $fullLogPathAndName -verboseLogging $verboseLogging}
        }

    $updatedKimbleLeads = Compare-Object -ReferenceObject $allKimbleLeads -DifferenceObject $allSpoLeads -Property @("Id","LastModifiedDate") -PassThru -CaseSensitive:$false
    $updatedKimbleLeads | ?{$_.SideIndicator -eq "<="} | % {
        $thisLead = $_
        $spoLead = $allSpoLeads | ? {$_.Id -eq $thisLead.Id}
        if($spoLead){
            #We've found the corresponding spoObject
            $updatedValues = @{"IsDeleted"=$thisLead.IsDeleted}
            if($thisLead.LastModifiedDate -ne $null){
                $updatedValues.Add("LastModifiedDate",$(Get-Date $thisLead.LastModifiedDate -Format "yyyy/MM/dd hh:mm:ss"))
                }
            if($thisLead.Name -ne $spoLead.Name){
                $updatedValues.Add("Title",$thisLead.Name)
                $updatedValues.Add("PreviousName",$thisLead.Name)
                $updatedValues.Add("IsDirty",$true)
                #Write-Host "$($spoClient.Name) renamed to $($updatedClient.Name)"
                }
            if($thisLead.KimbleOne__Account__c -ne $spoLead.KimbleClientId){
                $updatedValues.Add("KimbleClientId",$thisLead.KimbleOne__Account__c)
                $updatedValues.Add("PreviousKimbleClientId",$spoLead.KimbleClientId)
                $updatedValues.Add("IsDirty",$true)
                #Write-Host "$($spoClient.Name) renamed to $($updatedClient.Name)"
                }
            if($verboseLogging){Write-Host -ForegroundColor DarkYellow "Set-PnPListItem -List $($leadList.Id) -Identity $($spoLead.SPListItemID) -Values @{$(stringify-hashTable $updatedValues -interlimiter ":" -delimiter ", ")}"}
            Set-PnPListItem -List $leadList.Id -Identity $spoLead.SPListItemID -Values $updatedValues
            }
       
        }
    }
function reconcile-projectsBetweenKimbleAndSpo($standardKimbleQueryUri, $standardKimbleHeaders, $webUrl, $sitePath, $adminCreds){
    #Get the full list of Kimble Projects
    $allKimbleProjects = get-allKimbleProjects -pQueryUri $standardKimbleQueryUri -pRestHeaders $standardKimbleHeaders
    
    Connect-PnPOnline –Url $($webUrl+$sitePath) –Credentials $adminCreds
    $projectList = Get-PnPList -Identity "Kimble Projects" -Includes ContentTypes
    $projectListContentType = $projectList.ContentTypes | ? {$_.Name -eq "Item"}
    $projectListItems = Get-PnPListItem -List "Kimble Projects" -PageSize 1000 -Fields "Title","GUID","KimbleId","projectDescription","KimbleClientId","PreviousKimbleClientId","IsDirty","IsDeleted","Modified","LastModifiedDate","PreviousName","PreviousDescription","Id"
    $projectListItems.FieldValues | %{
        $thisProject = $_
        [array]$allSpoProjects += New-Object psobject -Property $([ordered]@{"Id"=$thisProject["KimbleId"];"KimbleId"=$thisProject["KimbleId"];"KimbleClientId"=$thisProject["KimbleClientId"];"PreviousKimbleClientId"=$thisProject["PreviousKimbleClientId"];"Name"=$thisProject["Title"];"GUID"=$thisProject["GUID"];"SPListItemID"=$thisProject["ID"];"IsDirty"=$thisProject["IsDirty"];"IsDeleted"=$thisProject["IsDeleted"];"LastModifiedDate"=$thisProject["LastModifiedDate"];"PreviousName"=$thisProject["PreviousName"];"PreviousDescription"=$thisProject["PreviousDescription"]})
        }
    
    #Work out what's missing and create any omissions
    $allKimbleProjects | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name KimbleId -Value $_.Id}
    $missingProjects = Compare-Object -ReferenceObject $allKimbleProjects -DifferenceObject $allSpoProjects -Property "KimbleId" -PassThru -CaseSensitive:$false
    $missingProjects | % {
        if ($_.SideIndicator -eq "<="){new-spoKimbleProjectItem -kimbleProjectObject $_ -pnpProjectList $projectList -fullLogPathAndName $fullLogPathAndName -verboseLogging $verboseLogging}
        }

    $updatedKimbleProjects = Compare-Object -ReferenceObject $allKimbleProjects -DifferenceObject $allSpoProjects -Property @("Id","LastModifiedDate") -PassThru -CaseSensitive:$false
    $updatedKimbleProjects | ?{$_.SideIndicator -eq "<="} | % {
        $thisProject = $_
        $spoProject = $allSpoProjects | ? {$_.Id -eq $thisProject.Id}
        if($spoProject){
            #We've found the corresponding spoObject
            $updatedValues = @{"IsDeleted"=$thisProject.IsDeleted}
            if($thisProject.LastModifiedDate -ne $null){
                $updatedValues.Add("LastModifiedDate",$(Get-Date $thisProject.LastModifiedDate -Format "yyyy/MM/dd hh:mm:ss"))
                }
            if($thisProject.Name -ne $spoProject.Name){
                $updatedValues.Add("Title",$thisProject.Name)
                $updatedValues.Add("PreviousName",$thisProject.Name)
                $updatedValues.Add("IsDirty",$true)
                #Write-Host "$($spoClient.Name) renamed to $($updatedClient.Name)"
                }
            if($thisProject.KimbleOne__Account__c -ne $spoProject.KimbleClientId){
                $updatedValues.Add("KimbleClientId",$thisProject.KimbleOne__Account__c)
                $updatedValues.Add("PreviousKimbleClientId",$spoProject.KimbleClientId)
                $updatedValues.Add("IsDirty",$true)
                #Write-Host "$($spoClient.Name) renamed to $($updatedClient.Name)"
                }
            if($verboseLogging){Write-Host -ForegroundColor DarkYellow "Set-PnPListItem -List $($projectList.Id) -Identity $($spoProject.SPListItemID) -Values @{$(stringify-hashTable $updatedValues -interlimiter ":" -delimiter ", ")}"}
            Set-PnPListItem -List $projectList.Id -Identity $spoProject.SPListItemID -Values $updatedValues
            }
        }
    }


#endregion


##################################
#
#Do Stuff
#
##################################

#reconcile-clientsBetweenKimbleAndSpo -standardKimbleQueryUri $standardKimbleQueryUri -standardKimbleHeaders $standardKimbleHeaders -webUrl $webUrl -sitePath $sitePath -adminCreds $adminCreds
#reconcile-leads
#reconcile-projectsBetweenKimbleAndSpo -standardKimbleQueryUri $standardKimbleQueryUri -standardKimbleHeaders $standardKimbleHeaders -webUrl $webUrl -sitePath $sitePath -adminCreds $adminCreds
reconcile-objectsBetweenKimbleAndSpo -accountType $objectType`
# reconcile-objectsBetweenKimbleAndSpo -accountType "Projects" #Needs more work
