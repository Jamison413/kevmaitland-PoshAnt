$logFileLocation = "C:\ScriptLogs\"
if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)){
    $fullLogPathAndName = $logFileLocation+"reconcile-kimbleFocalPoint_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = $logFileLocation+"reconcile-kimbleFocalPoint_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
else{
    $fullLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
if($PSCommandPath){
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_Transcript_$(Get-Date -Format "yyMMdd").log"
    Start-Transcript $transcriptLogName -Append
    }

Import-Module _PS_Library_GeneralFunctionality
Import-Module _PS_Library_Databases
Import-Module _REST_Library-Kimble.psm1

#Get set up
$sqlDbConn = connect-toSqlServer -SQLServer "sql.sustain.co.uk" -SQLDBName "SUSTAIN_LIVE"
$kimbleCreds = Import-Csv "$env:USERPROFILE\Desktop\Kimble.txt"
$standardKimbleHeaders = get-kimbleHeaders -clientId $kimbleCreds.clientId -clientSecret $kimbleCreds.clientSecret -username $kimbleCreds.username -password $kimbleCreds.password -securityToken $kimbleCreds.securityToken -connectToLiveContext $true -verboseLogging $true
$standardKimbleQueryUri = get-kimbleQueryUri


#region Accounts
$cachedAccounts = get-allFocalPointCachedKimbleAccounts -dbConnection $sqlDbConn -pWhereStatement $null #-verboseLogging $true
$kimbleAccounts = get-allKimbleAccounts -pQueryUri $standardKimbleQueryUri -pRestHeaders $standardKimbleHeaders #-pWhereStatement "WHERE Sync_to_FocalPoint__c = TRUE" 
$accountsDelta = Compare-Object -ReferenceObject $kimbleAccounts -DifferenceObject $cachedAccounts -Property Id -PassThru -CaseSensitive:$false -IncludeEqual
if(!$cachedAccounts){$accountsDelta = $kimbleAccounts}


#Create any new (uncached) Accounts
$accountsDelta | ? {$_.SideIndicator -eq "<=" } | % {
    $me = $_
    add-kimbleAccountToFocalPointCache -kimbleAccount $me -dbConnection $sqlDbConn
    } | % {Write-Host $me.Id $me.Name}
#Update all the Accounts based on the Kimble Data
if($i -ne $null){rv i}
$accountsDelta | ? {$_.SideIndicator -eq "==" } | % {
    Write-Progress -Activity "Updating cached Kimble Accounts" -status "$($_.Name)" -percentComplete ($i / $kimbleAccounts.count * 100)
    $thisAccount = $_
    $result = update-kimbleAccountToFocalPointCache -kimbleAccount $thisAccount -dbConnection $sqlDbConn #| Out-Null 
    if($result -ne 1){Write-Host "FAILED TO UPDATE [$($thisAccount.Id)] [$($thisAccount.Name)]"}
    $i++
    }
#These are now missing from Kimble, so mark them as Deleted
$accountsDelta | ? {$_.SideIndicator -eq "=>" } | % {
    $me = $_
    $me.IsDeleted = $true
    $me | Add-Member -MemberType NoteProperty -Name "IsMissingFromKimble" -Value $true -Force
    $result = update-kimbleAccountToFocalPointCache -kimbleAccount $me -dbConnection $sqlDbConn 
    if($result -ne 1){Write-Host "FAILED TO UPDATE [$($me.Id)] [$($me.Name)]"}
    }
#endregion

#region Opps
$cachedOpps = get-allFocalPointCachedKimbleOpps -dbConnection $sqlDbConn -pWhereStatement $null
$kimbleOpps = get-allKimbleLeads -pQueryUri $standardKimbleQueryUri -pRestHeaders $standardKimbleHeaders
$oppsDelta = Compare-Object -ReferenceObject $kimbleOpps -DifferenceObject $cachedOpps -Property Id -PassThru -CaseSensitive:$false -IncludeEqual
if(!$cachedOpps){$oppsDelta = $kimbleOpps}

$oppsDelta | ? {$_.SideIndicator -eq "<="} | % {
    $me = $_
    add-kimbleOppToFocalPointCache -kimbleOpp $me -dbConnection $sqlDbConn
    } | % {Write-Host $me.Id $me.Name}
#Update all the Accounts based on the Kimble Data
if($i -ne $null){rv i}
$oppsDelta | ? {$_.SideIndicator -eq "=="} | % {
    $me = $_
    Write-Progress -Activity "Updating cached Kimble Opps" -status "$($me.Name)" -percentComplete ($i / $kimbleOpps.count * 100)
    $result = update-kimbleOppToFocalPointCache -kimbleOpp $me -dbConnection $sqlDbConn -Verbose
    if($result -ne 1){Write-Host "FAILED TO UPDATE [$($me.Id)] [$($me.Name)]"}
    $i++
    } | % {Write-Host $me.Id $me.Name}
#These are now missing from Kimble. Not sure what to do with these...
$oppsDelta | ? {$_.SideIndicator -eq "=>" } | % {
    $me = $_
    $me.IsDeleted = $true
    update-kimbleOppToFocalPointCache -kimbleOpp $me -dbConnection $sqlDbConn
    if($result -ne 1){Write-Host "FAILED TO UPDATE [$($me.Id)] [$($me.Name)]"}
    }
<# This populates the data for the first time
if($i -ne $null){rv i}
$kimbleOpps | %{
    Write-Progress -Activity "Adding cached Kimble Opps" -status "$($_.Name)" -percentComplete ($i / $kimbleOpps.count * 100)
    add-kimbleOppToFocalPointCache -kimbleOpp $_ -dbConnection $sqlDbConn | Out-Null 
    Write-Host $_.Id $_.Name
    $i++
    }
#>

#endregion


#region Props
$cachedProps = get-allFocalPointCachedKimbleProps -dbConnection $sqlDbConn -pWhereStatement $null
$kimbleProps = get-allKimbleProposals -pQueryUri $standardKimbleQueryUri -pRestHeaders $standardKimbleHeaders
$propsDelta = Compare-Object -ReferenceObject $kimbleProps -DifferenceObject $cachedProps -Property Id -PassThru -CaseSensitive:$false -IncludeEqual
if(!$cachedProps){$propsDelta = $kimbleProps}

$propsToAdd = $propsDelta | ? {$_.SideIndicator -eq "<="} 
$propsToAdd | % {
    $me = $_
    add-kimbleProposalToFocalPointCache -kimbleProp $me -dbConnection $sqlDbConn
    } | % {Write-Host $me.Id $me.Name}
#Update all the matched Accounts based on the Kimble Data
if($i -ne $null){rv i}
$propsToUpdate = $propsDelta | ? {$_.SideIndicator -eq "=="}
$propsToUpdate | % {
    $me = $_
    Write-Progress -Activity "Updating cached Kimble Proposals" -status "$($me.Name)" -percentComplete ($i / $kimbleProps.count * 100)
    $result = update-kimbleProposalToFocalPointCache -kimbleProp $me -dbConnection $sqlDbConn 
    if($result -ne 1){Write-Host "FAILED TO UPDATE [$($me.Id)] [$($me.Name)]"}
    $i++
    }
#These are now missing from Kimble. Not sure what to do with these...
$propsToDelete = $propsDelta | ? {$_.SideIndicator -eq "=>" }
$propsToDelete | % {
    $me = $_
    $me.IsDeleted = $true
    $me | Add-Member -MemberType NoteProperty -Name "IsMissingFromKimble" -Value $true -Force
    $result = update-kimbleProposalToFocalPointCache -kimbleProp $me -dbConnection $sqlDbConn
    if($result -ne 1){Write-Host "FAILED TO UPDATE [$($me.Id)] [$($me.Name)]"}
    }
<# This populates the data for the first time
if($i -ne $null){rv i}
$kimbleOpps | %{
    Write-Progress -Activity "Adding cached Kimble Opps" -status "$($_.Name)" -percentComplete ($i / $kimbleOpps.count * 100)
    add-kimbleOppToFocalPointCache -kimbleOpp $_ -dbConnection $sqlDbConn | Out-Null 
    Write-Host $_.Id $_.Name
    $i++
    }
#>

#endregion



#region Engagements
$cachedEngagements = get-allFocalPointCachedKimbleEngagements -dbConnection $sqlDbConn -pWhereStatement $null
$kimbleEngagements = get-allKimbleEngagements -pQueryUri $standardKimbleQueryUri -pRestHeaders $standardKimbleHeaders -pWhereStatement $null
$engagementsDelta = Compare-Object -ReferenceObject $kimbleEngagements -DifferenceObject $cachedEngagements -Property Id -PassThru -CaseSensitive:$false -IncludeEqual

$engagementsToAdd = $engagementsDelta | ? {$_.SideIndicator -eq "<="} 
$engagementsToAdd | % {
    $me = $_
    add-kimbleEngagementToFocalPointCache -kimbleEngagement $me -dbConnection $sqlDbConn 
    } | % {Write-Host $me.Id $me.Name}
#Update all the matched Accounts based on the Kimble Data
$engagementsToUpdate = $engagementsDelta | ? {$_.SideIndicator -eq "=="}
if($i -ne $null){rv i}
$failed = @()
$engagementsToUpdate | % {
    $me = $_
    Write-Progress -Activity "Updating cached Kimble Engagments" -status "$($me.Name)" -percentComplete ($i / $engagementsToUpdate.count * 100)
    $result = update-kimbleEngagementToFocalPointCache -kimbleEngagement $me -dbConnection $sqlDbConn 
    if($result -ne 1){Write-Host "FAILED TO UPDATE [$($me.Id)] [$($me.Name)]";$failed += $me}
    $i++
    }
#These are now missing from Kimble. Not sure what to do with these...
$engagementsToDelete = $engagementsDelta | ? {$_.SideIndicator -eq "=>" }
$engagementsToDelete | % {
    $me = $_
    $me.IsDeleted = $true
    $me | Add-Member -MemberType NoteProperty -Name "IsMissingFromKimble" -Value $true -Force
    $result = update-kimbleEngagementToFocalPointCache -kimbleEngagement $me -dbConnection $sqlDbConn
    if($result -ne 1){Write-Host "FAILED TO UPDATE [$($me.Id)] [$($me.Name)]"}
    }

#endregion

#region Contacts
$cachedContacts = get-allFocalPointCachedKimbleContacts -dbConnection $sqlDbConn -pWhereStatement $null
$kimbleContacts = get-allKimbleContacts -pQueryUri $standardKimbleQueryUri -pRestHeaders $standardKimbleHeaders -pWhereStatement $null
$contactsDelta = Compare-Object -ReferenceObject $kimbleContacts -DifferenceObject $cachedContacts -Property Id -PassThru -CaseSensitive:$false -IncludeEqual

$contactsToAdd = $contactsDelta | ? {$_.SideIndicator -eq "<="} 
$contactsToAdd | % {
    $me = $_
    add-kimbleContactToFocalPointCache -kimbleContact $me -dbConnection $sqlDbConn 
    } | % {Write-Host $me.Id $me.Name}
#Update all the matched Accounts based on the Kimble Data
$contactsToUpdate = $contactsDelta | ? {$_.SideIndicator -eq "=="}
if($i -ne $null){rv i}
$failed = @()
$contactsToUpdate | % {
    $me = $_
    Write-Progress -Activity "Updating cached Kimble Engagments" -status "$($me.Name)" -percentComplete ($i / $engagementsToUpdate.count * 100)
    $result = update-kimbleContactToFocalPointCache -kimbleContact $me -dbConnection $sqlDbConn 
    if($result -ne 1){Write-Host "FAILED TO UPDATE [$($me.Id)] [$($me.Name)]";$failed += $me}
    $i++
    }
#These are now missing from Kimble. Not sure what to do with these...
$contactsToDelete = $contactsDelta | ? {$_.SideIndicator -eq "=>" }
$contactsToDelete | % {
    $me = $_
    $me.IsDeleted = $true
    $me | Add-Member -MemberType NoteProperty -Name "IsMissingFromKimble" -Value $true -Force
    $result = update-kimbleContactToFocalPointCache -kimbleContact $me -dbConnection $sqlDbConn
    if($result -ne 1){Write-Host "FAILED TO UPDATE [$($me.Id)] [$($me.Name)]"}
    }

#endregion

$sqlDbConn.close()
Stop-Transcript