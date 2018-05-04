$logFileLocation = "C:\ScriptLogs\"
$transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_Transcript_$(Get-Date -Format "yyMMdd").log"
if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)){
    $fullLogPathAndName = $logFileLocation+"sync-kimbleClientsToSpo_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = $logFileLocation+"sync-kimbleClientsToSpo_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
else{
    $fullLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
Start-Transcript $transcriptLogName -Append

Import-Module _PS_Library_GeneralFunctionality
Import-Module _PS_Library_Databases
Import-Module _REST_Library-Kimble.psm1

#Get set up
$sqlDbConn = connect-toSqlServer -SQLServer "sql.sustain.co.uk" -SQLDBName "SUSTAIN_LIVE"
$kimbleCreds = Import-Csv "$env:USERPROFILE\Desktop\Kimble.txt"
$standardKimbleHeaders = get-kimbleHeaders -clientId $kimbleCreds.clientId -clientSecret $kimbleCreds.clientSecret -username $kimbleCreds.username -password $kimbleCreds.password -securityToken $kimbleCreds.securityToken -connectToLiveContext $true -verboseLogging $true
$standardKimbleQueryUri = get-kimbleQueryUri


#region Contacts
$cachedContacts = get-allFocalPointCachedKimbleContacts -dbConnection $sqlDbConn -pWhereStatement $null -verboseLogging $true
$kimbleContacts = get-allKimbleAccounts -pQueryUri $standardKimbleQueryUri -pRestHeaders $standardKimbleHeaders #-pWhereStatement "WHERE Sync_to_FocalPoint__c = TRUE" 
$contactsDelta = Compare-Object -ReferenceObject $kimbleContacts -DifferenceObject $cachedContacts -Property Id -PassThru -CaseSensitive:$false
if(!$cachedContacts){$contactsDelta = $kimbleContacts}

#Create any new (uncached) Accounts
$contactsDelta | ? {$_.SideIndicator -eq "<="
    add-kimbleContactToFocalPointCache -kimbleContact $_ -dbConnection $sqlDbConn
    } | % {Write-Host $_.Id $_.Name}
#These are now missing from Kimble. Not sure what to do with these...
$contactsDelta | ? {$_.SideIndicator -eq "=>"
    
    } | % {Write-Host $_.Id $_.Name}

#Update all the Accounts based on the Kimble Data
if($i -ne $null){rv i}
$kimbleContacts | %{
    Write-Progress -Activity "Updating cached Kimble Contacts" -status "$($_.Name)" -percentComplete ($i / $kimbleContacts.count * 100)
    update-kimbleContactToFocalPointCache -kimbleContact $_ -dbConnection $sqlDbConn | Out-Null 
    Write-Host $_.Id $_.Name
    $i++
    }
<# This populates the data for the first time
if($i -ne $null){rv i}
$kimbleContacts | %{
    Write-Progress -Activity "Adding cached Kimble Accounts" -status "$($_.Name)" -percentComplete ($i / $kimbleContacts.count * 100)
    add-kimbleContactToFocalPointCache -kimbleContact $_ -dbConnection $sqlDbConn | Out-Null 
    Write-Host $_.Id $_.Name
    $i++
    }
#>
#endregion


$sqlDbConn.close()
Stop-Transcript