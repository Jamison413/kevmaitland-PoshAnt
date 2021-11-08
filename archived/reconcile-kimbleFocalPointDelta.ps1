[cmdletbinding()]
$logFileLocation = "C:\ScriptLogs\"
$scriptName = "reconcile-kimbleFocalPointDelta"
$fullLogPathAndName = $logFileLocation+$scriptName+"_FullLog_$(Get-Date -Format "yyMMdd").log"
$errorLogPathAndName = $logFileLocation+$scriptName+"_ErrorLog_$(Get-Date -Format "yyMMdd").log"
if($PSCommandPath){
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_Transcript_$(Get-Date -Format "yyMMdd").log"
    Start-Transcript $transcriptLogName -Append
    }

$verboseLogging = $true

Import-Module _PS_Library_GeneralFunctionality
Import-Module _PS_Library_Databases
Import-Module _REST_Library-Kimble


#Get set up
$sqlDbConn = connect-toSqlServer -SQLServer "sql.sustain.co.uk" -SQLDBName "SUSTAIN_LIVE"
$kimbleCreds = Import-Csv "$env:USERPROFILE\Desktop\Kimble.txt"
$standardKimbleHeaders = get-kimbleHeaders -clientId $kimbleCreds.clientId -clientSecret $kimbleCreds.clientSecret -username $kimbleCreds.username -password $kimbleCreds.password -securityToken $kimbleCreds.securityToken -connectToLiveContext $true -verboseLogging $true
$standardKimbleQueryUri = get-kimbleQueryUri

log-action -myMessage " " -logFile $fullLogPathAndName
log-action -myMessage "************************************************************************" -logFile $fullLogPathAndName
log-action -myMessage "New $scriptName cycle" -logFile $fullLogPathAndName
log-action -myMessage "************************************************************************" -logFile $fullLogPathAndName

#region Get all modified Accounts records since the last update (delta update) - this will necessarily include all created records
$lastCreatedAccountInDbSql = "SELECT MAX(CreatedDate) AS LastCreatedDate FROM SUS_Kimble_Accounts"
$lastCreatedAccountInDb = Execute-SQLQueryOnSQLDB -query $lastCreatedAccountInDbSql -queryType Scalar -sqlServerConnection $sqlDbConn
$lastModifiedAccountInDbSql = "SELECT MAX(LastModifiedDate) AS LastModifiedDate FROM SUS_Kimble_Accounts"
$lastModifiedAccountInDb = Execute-SQLQueryOnSQLDB -query $lastModifiedAccountInDbSql -queryType Scalar -sqlServerConnection $sqlDbConn
$cutoffModifiedAccountDate = Get-Date $lastModifiedAccountInDb -Format s #Does this need to account for Daylight Saving Time?

$modifiedKimbleAccounts = get-allKimbleAccounts -pQueryUri $standardKimbleQueryUri -pRestHeaders $standardKimbleHeaders -pWhereStatement "WHERE LastModifiedDate > $cutoffModifiedAccountDate`Z" -verboseLogging $true
log-action -myMessage "[$($modifiedKimbleAccounts.Count)] Kimble Accounts require updating" -logFile $fullLogPathAndName
$modifiedKimbleAccounts | % {
    if((Get-Date $_.CreatedDate) -gt $lastCreatedAccountInDb){
        #Create any new Accounts
        Write-Verbose "Creating Accounts [$($_.Name)]:[$($_.Id)]"
        $result = add-kimbleAccountToFocalPointCache -kimbleAccount $_ -dbConnection $sqlDbConn
        if($result -ne 1){[array]$duffCreateAccounts += @($_,$result)}
        }
    else{
        #If it's not new, it must have been modified, so Update it
        Write-Verbose "Updating Accounts [$($_.Name)]:[$($_.Id)]"
        $result = update-kimbleAccountToFocalPointCache -kimbleAccount $_ -dbConnection $sqlDbConn
        if($result -ne 1){[array]$duffModifyAccounts += @($_,$result)}
        }
    }
if($modifiedKimbleAccounts){log-result -myMessage "[$($duffCreateAccounts.Count)] Accounts failed to create" -logFile $fullLogPathAndName}
if($duffCreateAccounts){log-result $($duffCreateAccounts | % {"["+$_[0].Id+"]["+$_[0].Name+"]["+$_[1]+"];"}) -logFile $fullLogPathAndName}
if($modifiedKimbleAccounts){log-result -myMessage "[$($duffModifyAccounts.Count)] Accounts failed to update" -logFile $fullLogPathAndName}
if($duffModifyAccounts){log-result $($duffModifyAccounts | % {"["+$_[0].Id+"]["+$_[0].Name+"]["+$_[1]+"];"}) -logFile $fullLogPathAndName}
#endregion

#region Now do the same for Opportunities
$lastModifiedOppInDbSql = "SELECT MAX(LastModifiedDate) AS LastModifiedDate FROM SUS_Kimble_Opps"
$lastModifiedOppInDb = Execute-SQLQueryOnSQLDB -query $lastModifiedOppInDbSql -queryType Scalar -sqlServerConnection $sqlDbConn
$cutoffModifiedOppDate = Get-Date $lastModifiedOppInDb -Format s #Does this need to account for Daylight Saving Time?
$lastCreatedOppInDbSql = "SELECT MAX(CreatedDate) AS LastCreatedDate FROM SUS_Kimble_Opps"
$lastCreatedOppInDb = Execute-SQLQueryOnSQLDB -query $lastCreatedOppInDbSql -queryType Scalar -sqlServerConnection $sqlDbConn

$modifiedKimbleOpps = get-allKimbleLeads -pQueryUri $standardKimbleQueryUri -pRestHeaders $standardKimbleHeaders -pWhereStatement "WHERE LastModifiedDate > $cutoffModifiedOppDate`Z"
log-action -myMessage "[$($modifiedKimbleProps.Count)] Kimble Opportunities require updating" -logFile $fullLogPathAndName
$modifiedKimbleOpps | % {
    if((Get-Date $_.CreatedDate) -gt $lastCreatedOppInDb){
        #Create any new Opps
        Write-Verbose "Creating Opportunity [$($_.Name)]:[$($_.Id)]"
        $result = add-kimbleOppToFocalPointCache -kimbleOpp $_ -dbConnection $sqlDbConn
        if ($result -ne 1){[array]$duffCreateOpps += @($_,$result)}
        }
    else{
        #If it's not new, it must have been modified, so Update it
        Write-Verbose "Updating Opportunity [$($_.Name)]:[$($_.Id)]"
        $result = update-kimbleOppToFocalPointCache -kimbleOpp $_ -dbConnection $sqlDbConn
        if ($result -ne 1){[array]$duffModifyOpps += @($_,$result)}
        }
    }
if($modifiedKimbleOpps){log-result -myMessage "[$($duffCreateOpps.Count)] Opportunities failed to create" -logFile $fullLogPathAndName}
if($duffCreateOpps){log-result $($duffCreateOpps | % {"["+$_[0].Id+"]["+$_[0].Name+"]["+$_[1]+"];"}) -logFile $fullLogPathAndName}
if($modifiedKimbleOpps){log-result -myMessage "[$($duffModifyOpps.Count)] Opportunities failed to update" -logFile $fullLogPathAndName}
if($duffModifyOpps){log-result $($duffModifyOpps | % {"["+$_[0].Id+"]["+$_[0].Name+"]["+$_[1]+"];"}) -logFile $fullLogPathAndName}
#endregion

#region Now do the same for Proposals
$lastModifiedPropInDbSql = "SELECT MAX(LastModifiedDate) AS LastModifiedDate FROM SUS_Kimble_Proposals"
$lastModifiedPropInDb = Execute-SQLQueryOnSQLDB -query $lastModifiedPropInDbSql -queryType Scalar -sqlServerConnection $sqlDbConn
$cutoffModifiedPropDate = Get-Date $lastModifiedPropInDb -Format s #Does this need to account for Daylight Saving Time?
$lastCreatedPropInDbSql = "SELECT MAX(CreatedDate) AS LastCreatedDate FROM SUS_Kimble_Proposals"
$lastCreatedPropInDb = Execute-SQLQueryOnSQLDB -query $lastCreatedPropInDbSql -queryType Scalar -sqlServerConnection $sqlDbConn

$modifiedKimbleProps = get-allKimbleProposals -pQueryUri $standardKimbleQueryUri -pRestHeaders $standardKimbleHeaders -pWhereStatement "WHERE LastModifiedDate > $cutoffModifiedPropDate`Z" -verboseLogging $true
log-action -myMessage "[$($modifiedKimbleProps.Count)] Kimble Proposals require updating" -logFile $fullLogPathAndName
$modifiedKimbleProps | % {
    if((Get-Date $_.CreatedDate) -gt $lastCreatedPropInDb){
        #Create any new Props
        Write-Verbose "Creating Proposal [$($_.Name)]:[$($_.Id)]"
        $result = add-kimbleProposalToFocalPointCache -kimbleProp $_ -dbConnection $sqlDbConn -verboseLogging $verboseLogging
        if ($result -ne 1){[array]$duffCreateProps += @($_,$result)}
        }
    else{
        #If it's not new, it must have been modified, so Update it
        Write-Verbose "Updating Proposal [$($_.Name)]:[$($_.Id)]"
        $result = update-kimbleProposalToFocalPointCache -kimbleProp $_ -dbConnection $sqlDbConn -verboseLogging $verboseLogging
        if ($result -ne 1){[array]$duffModifyProps += @($_,$result)}
        }
    }
if($modifiedKimbleProps){log-result -myMessage "[$($duffCreateProps.Count)] Proposals failed to create" -logFile $fullLogPathAndName}
if($duffCreateProps){log-result $($duffCreateProps | % {"["+$_[0].Id+"]["+$_[0].Name+"]["+$_[1]+"];"}) -logFile $fullLogPathAndName}
if($modifiedKimbleProps){log-result -myMessage "[$($duffModifyProps.Count)] Proposals failed to update" -logFile $fullLogPathAndName}
if($duffModifyProps){log-result $($duffModifyProps | % {"["+$_[0].Id+"]["+$_[0].Name+"]["+$_[1]+"];"}) -logFile $fullLogPathAndName}
#endregion

#region Now do the same for Engagements
$lastCreatedEngagementInDbSql = "SELECT MAX(CreatedDate) AS LastCreatedDate FROM SUS_Kimble_Engagements"
$lastCreatedEngagementInDb = Execute-SQLQueryOnSQLDB -query $lastCreatedEngagementInDbSql -queryType Scalar -sqlServerConnection $sqlDbConn
$lastModifiedEngagementInDbSql = "SELECT MAX(LastModifiedDate) AS LastModifiedDate FROM SUS_Kimble_Engagements"
$lastModifiedEngagementInDb = Execute-SQLQueryOnSQLDB -query $lastModifiedEngagementInDbSql -queryType Scalar -sqlServerConnection $sqlDbConn
$cutoffModifiedEngagementDate = Get-Date $lastModifiedEngagementInDb -Format s #Does this need to Engagement for Daylight Saving Time?

#Get all modified Engagements records since the last update (delta update) - this will necessarily include all created records
$modifiedKimbleEngagements = get-allKimbleEngagements -pQueryUri $standardKimbleQueryUri -pRestHeaders $standardKimbleHeaders -pWhereStatement "WHERE LastModifiedDate > $cutoffModifiedEngagementDate`Z" -verboseLogging $true
log-action -myMessage "[$($modifiedKimbleEngagements.Count)] Kimble Engagements require updating" -logFile $fullLogPathAndName
$modifiedKimbleEngagements | % {
    if((Get-Date $_.CreatedDate) -gt $lastCreatedEngagementInDb){
        #Create any new Engagements
        Write-Verbose "Creating Engagement [$($_.Name)]:[$($_.Id)]"
        $result = add-kimbleEngagementToFocalPointCache -kimbleEngagement $_ -dbConnection $sqlDbConn
        if($result -ne 1){[array]$duffCreateEngagements += @($_,$result)}
        }
    else{
        #If it's not new, it must have been modified, so Update it
        Write-Verbose "Updating Engagement [$($_.Name)]:[$($_.Id)]"
        $result = update-kimbleEngagementToFocalPointCache -kimbleEngagement $_ -dbConnection $sqlDbConn 
        if($result -ne 1){[array]$duffModifyEngagements += ,@($_,$result)}
        }
    }
if($modifiedKimbleEngagements){log-result -myMessage "[$($duffCreateEngagements.Count)] Engagements failed to create" -logFile $fullLogPathAndName}
if($duffCreateEngagements){log-result $($duffCreateEngagements | % {"["+$_[0].Id+"]["+$_[0].Name+"]["+$_[1]+"];"}) -logFile $fullLogPathAndName}
if($modifiedKimbleEngagements){log-result -myMessage "[$($duffModifyEngagements.Count)] Engagements failed to update" -logFile $fullLogPathAndName}
if($duffModifyEngagements){log-result $($duffModifyEngagements | % {"["+$_[0].Id+"]["+$_[0].Name+"]["+$_[1]+"];"}) -logFile $fullLogPathAndName}
#endregion

$sqlDbConn.Close()
Stop-Transcript