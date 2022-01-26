$i = 0
$allClients | % {
    $thisClient = $_
    $i++
    $sql = "UPDATE SUS_Kimble_Accounts SET "
    $sql += "PreviousName = "+(format-asSqlValue -value $thisClient.PreviousName -dataType String)+","
    $sql += "Description = "+(format-asSqlValue -value $thisClient.ClientDescription -dataType String)+","
    $sql += "PreviousDescription = "+(format-asSqlValue -value $thisClient.PreviousDescription -dataType String)+","
    $sql += "DocumentLibraryGUID = "+(format-asSqlValue -value $thisClient.LibraryGUID -dataType Guid)
    $sql += " WHERE Id = '$(sanitise-forSql $thisClient.Id)'"
    try{$result = Execute-SQLQueryOnSQLDB -query $sql -queryType nonquery -sqlServerConnection $sqlDbConn}
    catch{write-host -ForegroundColor Red $_
        Write-Host -ForegroundColor darkred "$i`t$sql"
        break
        }
    if($result -ne 1){
        Write-Host -ForegroundColor Yellow "$i`t$sql"
        }
    }

$i = 0
$allProjects | % {
    $thisProject = $_
    $i++
    $sql = "UPDATE SUS_Kimble_Engagements SET "
    $sql += "PreviousName = "+(format-asSqlValue -value $thisProject.PreviousName -dataType String)+","
    $sql += "SuppressFolderCreation = "+(format-asSqlValue -value $thisProject.DoNotProcess -dataType Boolean)+","
    $sql += "PreviousKimbleClientId = "+(format-asSqlValue -value $thisProject.KimbleClientId -dataType String)+","
    $sql += "FolderGuid = "+(format-asSqlValue -value $thisProject.FolderGUID -dataType Guid)
    $sql += " WHERE Id = '$(sanitise-forSql $thisProject.Id)'"
    try{$result = Execute-SQLQueryOnSQLDB -query $sql -queryType nonquery -sqlServerConnection $sqlDbConn}
    catch{write-host -ForegroundColor Red $_
        Write-Host -ForegroundColor darkred "$i`t$sql"
        break
        }
    if($result -ne 1){
        Write-Host -ForegroundColor Yellow "$i`t$result`t$sql"
        }
    }


$sharePointAdmin = "kimblebot@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Downloads\KimbleBot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass

$webUrl = "https://anthesisllc.sharepoint.com"
$spoSite = "/clients"
Connect-PnPOnline –Url $($webUrl+$spoSite) –Credentials $adminCreds #-RequestTimeout 7200000
$sqlDbConn = connect-toSqlServer -SQLServer "sql.sustain.co.uk" -SQLDBName "SUSTAIN_LIVE"


$allClients = get-spoKimbleClientListItems -spoCredentials $adminCreds
$i = 0
$allClients | % {
    $thisClient = $_
    $i++
    $sql = "UPDATE SUS_Kimble_Accounts SET "
    $sql += "IsDirty = "+(format-asSqlValue -value $thisClient.IsDirty -dataType Boolean)
    $sql += " WHERE Id = '$(sanitise-forSql $thisClient.Id)'"
    try{$result = Execute-SQLQueryOnSQLDB -query $sql -queryType nonquery -sqlServerConnection $sqlDbConn}
    catch{write-host -ForegroundColor Red $_
        Write-Host -ForegroundColor darkred "$i`t$sql"
        break
        }
    if($result -ne 1){
        Write-Host -ForegroundColor Yellow "$i`t$sql"
        }
    }


$allProjects = get-spoKimbleProjectListItems -spoCredentials $adminCreds

$i = 0
$allProjects | % {
    $thisProject = $_
    $i++
    $sql = "UPDATE SUS_Kimble_Engagements SET "
    $sql += "IsDirty = "+(format-asSqlValue -value $thisProject.PreviousName -dataType Boolean)
    $sql += " WHERE Id = '$(sanitise-forSql $thisProject.Id)'"
    try{$result = Execute-SQLQueryOnSQLDB -query $sql -queryType nonquery -sqlServerConnection $sqlDbConn}
    catch{write-host -ForegroundColor Red $_
        Write-Host -ForegroundColor darkred "$i`t$sql"
        break
        }
    if($result -ne 1){
        Write-Host -ForegroundColor Yellow "$i`t$result`t$sql"
        }
    }


