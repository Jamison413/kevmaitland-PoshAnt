#reconcile-sqlKimbleObjectsToSharePoint


$logFileLocation = "C:\ScriptLogs\"
$logFileName = "sync-kimbleSqlObjectsToSpoObjects"
$fullLogPathAndName = $logFileLocation+$logFileName+"_$objectType`_FullLog_$(Get-Date -Format "yyMMdd").log"
$errorLogPathAndName = $logFileLocatio+$logFileName+"_$objectType`_ErrorLog_$(Get-Date -Format "yyMMdd").log"
if($PSCommandPath){
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_$objectType`_Transcript_$(Get-Date -Format "yyMMdd").log"
    Start-Transcript $transcriptLogName -Append
    }

Import-Module _PS_Library_GeneralFunctionality
Import-Module _PNP_Library_SPO
Import-Module SharePointPnPPowerShellOnline

$webUrl = "https://anthesisllc.sharepoint.com"
$spoSite = "/clients"

$smtpServer = "anthesisgroup-com.mail.protection.outlook.com"
$mailFrom = "$(split-path $PSCommandPath -Leaf)_netmon@sustain.co.uk"
$mailTo = "kevin.maitland@anthesisgroup.com"
$recreateAllFolders = $false

$sharePointAdmin = "kimblebot@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Downloads\KimbleBot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass

$sqlAccountsTableName = "SUS_Kimble_Accounts"
$sqlEngagementsTableName = "SUS_Kimble_Engagements"

Connect-PnPOnline –Url $($webUrl+$spoSite) –Credentials $adminCreds #-RequestTimeout 7200000
$sqlDbConn = connect-toSqlServer -SQLServer "sql.sustain.co.uk" -SQLDBName "SUSTAIN_LIVE"

log-action -myMessage "Retrieving all Client Document Libraries from SPO" -logFile $fullLogPathAndName
$durationSpoDocLibs = Measure-Command {$allSpoClientDocLibs = Get-PnPList -Includes Description,ItemCount,LastItemDeletedDate,LastItemModifiedDate,LastItemUserModifiedDate}
log-result -myMessage "[$($allSpoClientDocLibs.Count)] Document Libraries retrieved from SPO in [$($durationSpoDocLibs.TotalSeconds)] seconds" -logFile $fullLogPathAndName
log-action -myMessage "Retrieving all Client Document Libraries from SQL" -logFile $fullLogPathAndName
$durationSqlAccounts = Measure-Command {$allSqlKimbleClients = get-allFocalPointCachedKimbleAccounts -dbConnection $sqlDbConn -pWhereStatement "WHERE Type LIKE '%Client%'"}
log-result -myMessage "[$($allsqlKimbleClients.Count)] cached Kimble records retrieved from SQL in [$($durationSqlAccounts.TotalSeconds)] seconds" -logFile $fullLogPathAndName

#Add new property so we can compare-object
$allSpoClientDocLibs | % {$_ | Add-Member -MemberType NoteProperty -Name DocumentLibraryGuid -Value $_.Id -Force}
$allSpoClientDocLibs | % {$_ | Add-Member -MemberType NoteProperty -Name MyGuid -Value $_.Id -Force}
$allSqlKimbleClients | % {$_ | Add-Member -MemberType NoteProperty -Name MyGuid -Value $_.DocumentLibraryGuid -Force}
$allSqlKimbleClientsWithGuids = $allsqlKimbleClients | ? {![string]::IsNullOrWhiteSpace($_.MyGuid)}
log-action -myMessage "Reconciling SQL vs SPO objects" -logFile $fullLogPathAndName
$durationSimpleGuid = measure-command {
    $docLibDelta = Compare-Object -ReferenceObject $allSqlKimbleClientsWithGuids -DifferenceObject $allSpoClientDocLibs -Property MyGuid -PassThru -CaseSensitive:$false -IncludeEqual
    }
log-result -myMessage "Delta processed in  [$($durationSimpleGuid.TotalSeconds)] seconds: Results contain [$($docLibDelta.Count)] records ([$($docLibDelta.Count - $allSpoClientDocLibs.Count)] more than `$allSpoClientDocLibs and [$($docLibDelta.Count - $allsqlKimbleClients.Count)] more than `$allsqlKimbleClients)" -logFile $fullLogPathAndName

$matchedByGuid =  $docLibDelta | ? {$_.SideIndicator -eq "=="}
log-result -myMessage "[$($matchedByGuid.Count)] records are matched between SPO & SQL" -logFile $fullLogPathAndName
$missingFromSpo = $docLibDelta | ? {$_.SideIndicator -eq "<="}
log-result -myMessage "[$($missingFromSpo.Count)] records are missing from SPO" -logFile $fullLogPathAndName
$missingFromSql = $docLibDelta | ? {$_.SideIndicator -eq "=>"}
log-result -myMessage "[$($missingFromSql.Count)] records are missing from SQL" -logFile $fullLogPathAndName

$missingFromSql | % {$_ | Add-Member -MemberType NoteProperty -Name Name -Value $_.Title -Force}
$2ndPass = Compare-Object -ReferenceObject $missingFromSql -DifferenceObject $allsqlKimbleClients -Property Name -PassThru -CaseSensitive:$false -IncludeEqual
$matchedByName = $2ndPass | ? {$_.SideIndicator -eq "=="}
log-result -myMessage "[$($matchedByName.Count)] missing records have been matched by name - will set GUIDs in SQL" -logFile $fullLogPathAndName

$matchedByName | % {
    $thisMatch = $_
    $kimbleObj = $allsqlKimbleClients | ? {$_.Name -eq $thisMatch.Name}
    $sql = "UPDATE SUS_Kimble_Accounts SET DocumentLibraryGuid = '$($thisMatch.Id)' WHERE Id = '$($kimbleObj.Id)'"
    $result = Execute-SQLQueryOnSQLDB -query $sql -queryType NonQuery -sqlServerConnection $sqlDbConn
    if($result -ne 1){log-result -myMessage "FAILED TO UPDATE: $sql" -logFile $fullLogPathAndName}
    }

$totallyLost = $2ndPass | ? {$_.SideIndicator -eq "<="}
if($totallyLost.Count -gt 0){log-result -myMessage ("The following DocLibs look like they're orphaned (no Guid, no name match in Kimble): `n`t`t" + $($totallyLost.RootFolder.ServerRelativeUrl -join "`n`t`t")) -logFile $fullLogPathAndName}

$sqlDbConn.Close()