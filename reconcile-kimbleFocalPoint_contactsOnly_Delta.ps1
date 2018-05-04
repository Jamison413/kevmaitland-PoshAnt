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
Import-Module _REST_Library-Kimble


#Get set up
$sqlDbConn = connect-toSqlServer -SQLServer "sql.sustain.co.uk" -SQLDBName "SUSTAIN_LIVE"
$kimbleCreds = Import-Csv "$env:USERPROFILE\Desktop\Kimble.txt"
$standardKimbleHeaders = get-kimbleHeaders -clientId $kimbleCreds.clientId -clientSecret $kimbleCreds.clientSecret -username $kimbleCreds.username -password $kimbleCreds.password -securityToken $kimbleCreds.securityToken -connectToLiveContext $true -verboseLogging $true
$standardKimbleQueryUri = get-kimbleQueryUri

$lastCreatedAccountInDbSql = "SELECT MAX(CreatedDate) AS LastCreatedDate FROM SUS_Kimble_Accounts"
$lastCreatedAccountInDb = Execute-SQLQueryOnSQLDB -query $lastCreatedAccountInDbSql -queryType Scalar -sqlServerConnection $sqlDbConn
$lastModifiedAccountInDbSql = "SELECT MAX(LastModifiedDate) AS LastModifiedDate FROM SUS_Kimble_Accounts"
$lastModifiedAccountInDb = Execute-SQLQueryOnSQLDB -query $lastModifiedAccountInDbSql -queryType Scalar -sqlServerConnection $sqlDbConn
$cutoffModifiedAccountDate = Get-Date $lastModifiedAccountInDb -Format s #Does this need to account for Daylight Saving Time?

#Get all modified Accounts records since the last update (delta update) - this will necessarily include all created records
$modifiedKimbleAccounts = get-allKimbleAccounts -pQueryUri $standardKimbleQueryUri -pRestHeaders $standardKimbleHeaders -pWhereStatement "WHERE LastModifiedDate > $cutoffModifiedAccountDate`Z"
$modifiedKimbleAccounts | % {
    if((Get-Date $_.CreatedDate) -gt $lastCreatedAccountInDb){
        #Create any new Accounts
        if($verboseLogging){Write-Host -ForegroundColor DarkYellow Creating Account [$($_.Name)]:[$($_.Id)]}
        $result = add-kimbleAccountToFocalPointCache -kimbleAccount $_ -dbConnection $sqlDbConn -verboseLogging $verboseLogging
        if($result -ne 1){[array]$duffCreateAccounts += @($_,$result)}
        }
    else{
        #If it's not new, it must have been modified, so Update it
        if($verboseLogging){Write-Host -ForegroundColor DarkYellow Updating Opp [$($_.Name)]:[$($_.Id)]}
        $result = update-kimbleAccountToFocalPointCache -kimbleAccount $_ -dbConnection $sqlDbConn -verboseLogging $verboseLogging
        if($result -ne 1){[array]$duffModifyAccounts += @($_,$result)}
        }
    }

$lastModifiedOppInDbSql = "SELECT MAX(LastModifiedDate) AS LastModifiedDate FROM SUS_Kimble_Opps"
$lastModifiedOppInDb = Execute-SQLQueryOnSQLDB -query $lastModifiedOppInDbSql -queryType Scalar -sqlServerConnection $sqlDbConn
$cutoffModifiedOppDate = Get-Date $lastModifiedOppInDb -Format s #Does this need to account for Daylight Saving Time?
$lastCreatedOppInDbSql = "SELECT MAX(CreatedDate) AS LastCreatedDate FROM SUS_Kimble_Opps"
$lastCreatedOppInDb = Execute-SQLQueryOnSQLDB -query $lastCreatedOppInDbSql -queryType Scalar -sqlServerConnection $sqlDbConn

$modifiedKimbleOpps = get-allKimbleLeads -pQueryUri $standardKimbleQueryUri -pRestHeaders $standardKimbleHeaders -pWhereStatement "WHERE LastModifiedDate > $cutoffModifiedOppDate`Z"
$modifiedKimbleOpps | % {
    if((Get-Date $_.CreatedDate) -gt $lastCreatedOppInDb){
        #Create any new Opps
        if($verboseLogging){Write-Host -ForegroundColor DarkYellow Creating Opp [$($_.Name)]:[$($_.Id)]}
        $result = add-kimbleOppToFocalPointCache -kimbleOpp $_ -dbConnection $sqlDbConn -verboseLogging $verboseLogging
        if ($result -ne 1){[array]$duffCreateOpps += @($_,$result)}
        }
    else{
        #If it's not new, it must have been modified, so Update it
        if($verboseLogging){Write-Host -ForegroundColor DarkYellow Updating Opp [$($_.Name)]:[$($_.Id)]}
        $result = update-kimbleOppToFocalPointCache -kimbleOpp $_ -dbConnection $sqlDbConn -verboseLogging $verboseLogging
        if ($result -ne 1){[array]$duffModifyOpps += @($_,$result)}
        }
    }

$sqlDbConn.Close()
Stop-Transcript