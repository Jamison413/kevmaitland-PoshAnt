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
$cachedAccounts = get-allFocalPointCachedKimbleAccounts -dbConnection $sqlDbConn -pWhereStatement $null -verboseLogging $true
$kimbleAccounts = get-allKimbleAccounts -pQueryUri $standardKimbleQueryUri -pRestHeaders $standardKimbleHeaders #-pWhereStatement "WHERE Sync_to_FocalPoint__c = TRUE" 
$accountsDelta = Compare-Object -ReferenceObject $kimbleAccounts -DifferenceObject $cachedAccounts -Property Id -PassThru -CaseSensitive:$false
if(!$cachedAccounts){$accountsDelta = $kimbleAccounts}

#Create any new (uncached) Accounts
$accountsDelta | ? {$_.SideIndicator -eq "<=" } | % {
    $me = $_
    add-kimbleAccountToFocalPointCache -kimbleAccount $_ -dbConnection $sqlDbConn
    } | % {Write-Host $me.Id $me.Name}
#Update all the Accounts based on the Kimble Data
if($i -ne $null){rv i}
$kimbleAccounts | %{
    Write-Progress -Activity "Updating cached Kimble Accounts" -status "$($_.Name)" -percentComplete ($i / $kimbleAccounts.count * 100)
    update-kimbleAccountToFocalPointCache -kimbleAccount $_ -dbConnection $sqlDbConn | Out-Null 
    Write-Host $_.Id $_.Name
    $i++
    }
#These are now missing from Kimble, so mark them as Deleted
$accountsDelta | ? {$_.SideIndicator -eq "=>" } | % {
    $me = $_
    $me.IsDeleted = $true
    update-kimbleAccountToFocalPointCache -kimbleAccount $_ -dbConnection $sqlDbConn # | Out-Null 
    } | % {Write-Host $me.Id $me.Name}

<# This populates the data for the first time
if($i -ne $null){rv i}
$kimbleAccounts | %{
    Write-Progress -Activity "Adding cached Kimble Accounts" -status "$($_.Name)" -percentComplete ($i / $kimbleAccounts.count * 100)
    add-kimbleAccountToFocalPointCache -kimbleAccount $_ -dbConnection $sqlDbConn | Out-Null 
    Write-Host $_.Id $_.Name
    $i++
    }
#>
#endregion

#region Opps
$cachedOpps = get-allFocalPointCachedKimbleOpps -dbConnection $sqlDbConn -pWhereStatement $null
$kimbleOpps = get-allKimbleLeads -pQueryUri $standardKimbleQueryUri -pRestHeaders $standardKimbleHeaders
$oppsDelta = Compare-Object -ReferenceObject $kimbleOpps -DifferenceObject $cachedOpps -Property Id -PassThru -CaseSensitive:$false
if(!$cachedOpps){$oppsDelta = $kimbleOpps}

$oppsDelta | ? {$_.SideIndicator -eq "<="
    add-kimbleOppToFocalPointCache -kimbleOpp $_ -dbConnection $sqlDbConn
    } | % {Write-Host $_.Id $_.Name}
#Update all the Accounts based on the Kimble Data
if($i -ne $null){rv i}
$oppsDelta | %{
    Write-Progress -Activity "Updating cached Kimble Opps" -status "$($_.Name)" -percentComplete ($i / $kimbleOpps.count * 100)
    update-kimbleOppToFocalPointCache -kimbleOpp $_ -dbConnection $sqlDbConn | Out-Null 
    Write-Host $_.Id $_.Name
    $i++
    }
#These are now missing from Kimble. Not sure what to do with these...
$oppsDelta | ? {$_.SideIndicator -eq "=>" } | % {
    $me = $_
    $me.IsDeleted = $true
    update-kimbleOppToFocalPointCache -kimbleOpp $me -dbConnection $sqlDbConn
    } | % {Write-Host $me.Id $me.Name}
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

$sqlDbConn.close()
Stop-Transcript