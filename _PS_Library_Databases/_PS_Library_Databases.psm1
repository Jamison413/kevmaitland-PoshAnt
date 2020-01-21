function connect-toSqlServer($SQLServer,$SQLDBName){
    #SQL Server connection string
    $connDB = New-Object System.Data.SqlClient.SqlConnection
    $connDB.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True" #This relies on the current user having the appropriate Login/Role Membership ont he DB
    $connDB.Open()
    $connDB
    }
function connect-toAccessDB($dbPathAndName){
    #This needs to run in a 32-bit environemnt
    Write-Host -ForegroundColor DarkYellow "This needs to run in a 32-bit environemnt"
    $connDB = New-Object System.Data.OleDb.OleDbConnection
    $connDB.ConnectionString = "Provider=Microsoft.ACE.OLEDB.15.0; Data Source=`"$dbPathAndName`"; Persist Security Info=False;"
    $connDB.Open()
    $connDB
    }

function Execute-SQLQueryOnSQLDB([string]$query, [string]$queryType, $sqlServerConnection) { 
  # NonQuery - Insert/Update/Delete query where no return data is required
    $sql = New-Object System.Data.SqlClient.SqlCommand
    $sql.Connection = $sqlServerConnection
    $sql.CommandText = $query
    switch ($queryType){
        "NonQuery" {$sql.ExecuteNonQuery()}
        "Scalar" {$sql.ExecuteScalar()}
        "Reader" {    
            $oReader = $sql.ExecuteReader()
            $results = @()
            while ($oReader.Read()){
                $result = New-Object PSObject
                $skipMe = 0
                for ($i = 0; $oReader.FieldCount -gt $i; $i++){
                    $columnName = ($query.Replace(",","") -split '\s+')[$i+$skipMe+1]
                    if($columnName -eq "TOP"){$skipMe = 2;$columnName = ($query.Replace(",","") -split '\s+')[$i+$skipMe+1]}
                    if (1 -lt $columnName.Split(".").Length){$columnName = $columnName.Split(".")[1]} #Trim off any table names
                    $result | Add-Member NoteProperty $columnName $oReader[$i]
                    }
                 $results += $result
                }
            $oReader.Close()
            return $results
            }
        }
    }
function Execute-SQLQueryOnAccessDB([string]$query, [string]$queryType) { 
  # NonQuery - Insert/Update/Delete query where no return data is required
    $sql = New-Object System.Data.OleDb.OleDbCommand
    $sql.Connection = $connDB
    $sql.CommandText = $query
    switch ($queryType){
        "NonQuery" {$sql.ExecuteNonQuery()}
        "Scalar" {$sql.ExecuteScalar()}
        "Reader" {    
            $oReader = $sql.ExecuteReader()
            $results = @()
            while ($oReader.Read()){
                $result = New-Object PSObject
                for ($i = 0; $oReader.FieldCount -gt $i; $i++){
                        $columnName = ($query.Replace(",","") -split '\s+')[$i+1]
                        if (1 -lt $columnName.Split(".").Length){$columnName = $columnName.Split(".")[1]} #Trim off any table names
                        $result | Add-Member NoteProperty $columnName $oReader[$i]
                        }
                 $results += $result
                }
            $oReader.Close()
            return $results
            }
        }
    }
