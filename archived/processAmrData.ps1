$columnToTime = @{8="0030";9="0100";10="0130";11="0200";12="0230";13="0300";14="0330";15="0400";16="0430";17="0500";18="0530";19="0600";20="0630";21="0700";22="0730";23="0800";24="0830";25="0900";26="0930";27="1000";28="1030";29="1100";30="1130";31="1200";32="1230";33="1300";34="1330";35="1400";36="1430";37="1500";38="1530";39="1600";40="1630";41="1700";42="1730";43="1800";44="1830";45="1900";46="1930";47="2000";48="2030";49="2100";50="2130";51="2200";52="2230";53="2300";54="2330";55="0000";}
$outputCsv = "$env:USERPROFILE\Desktop\Output.csv"
$dbPathAndName = "$env:USERPROFILE\Downloads\Front End - Unite AMR Data - COPY.accdb"
<#$files = gci "X:\Clients\Unite Integrated Solutions plc\101754-EM_Unite_Platform_Development\Incoming\Energy Data" -Filter "*.csv"
Add-Content -Value "DataSouce,Company name,Site Name,Online Meter Name,MPAN,Type,Est,Date,MeterReadValue,MeterTimeStamp" -Path $outputCsv -Force
foreach ($file in $files){
    $name = $file.BaseName.Replace($file.BaseName.Split(" ")[$file.BaseName.Split(" ").Count-1],"")
    Get-Content -Path $file.FullName | % {
        if(!([string]::IsNullOrEmpty($_))){
            if($_.Replace('"',"").SubString(0,6) -eq "Ignite"){
                for($i=8; $i-lt 56;$i++){
                    $outString = "$name,".Replace(" ","")
                    $outString += $_.Replace('"',"").Split(",")[0]+","
                    $outString += $_.Replace('"',"").Split(",")[1]+","
                    $outString += $_.Replace('"',"").Split(",")[2]+","
                    $outString += $_.Replace('"',"").Split(",")[3]+","
                    $outString += $_.Replace('"',"").Split(",")[4]+","
                    $outString += $_.Replace('"',"").Split(",")[5]+","
                    $outString += $_.Replace('"',"").Split(",")[6]+","
                    $outString += $_.Replace('"',"").Split(",")[7]+","
                    $outString += $_.Replace('"',"").Split(",")[$i]+","
                    $outString += $columnToTime[$i]
                    Add-Content -Value $outString -Path $outputCsv
                    }
                }
            }
        }
    }
#>
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
                        $columnName = ($query.Replace(",","") -split '\s+')[$i+1].Replace("[","").Replace("]","")
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

#This needs to run in a 32-bit environemnt
$connDB = New-Object System.Data.OleDb.OleDbConnection
$connDB.ConnectionString = "Provider=Microsoft.ACE.OLEDB.15.0; Data Source=`"$dbPathAndName`"; Persist Security Info=False;"
$connDB.Open()


#now put it into Access
$tableName = "AMR-Elec"
$recordCountFile =  "$env:USERPROFILE\Desktop\Counter.txt"
$headersFile =  "$env:USERPROFILE\Desktop\Headers.txt"
$lastCommandFile = "$env:USERPROFILE\Desktop\iDidIt.txt"
$formattedHeaders = Get-Content $headersFile
[bigint]$i = 1
[bigint]$lastRecord = Get-Content $recordCountFile
Get-Content $outputCsv | %{
    if($i -eq 1){
        [array]$headers = $_.Split(",")
        $formattedHeaders = ""
        foreach($header in $headers){
            $formattedHeaders += "[$header],"
            }
        $formattedHeaders = $formattedHeaders.Substring(0,$formattedHeaders.Length-1)
        $formattedHeaders | Out-File -FilePath $headersFile
        }
    elseif($i -ge $lastRecord){
        $formattedValues = ""
        for ($j=0;$j -lt $_.Split(",").Count; $j++){
            if(@(9) -contains $j){
                if([string]::IsNullOrEmpty($_.Split(",")[$j]) -or $_.Split(",")[$j] -eq "-"){$formattedValues += "0,"}
                else{$formattedValues += $_.Split(",")[$j]+","}
                }#Format these fields as Numbers
            elseif(@(3) -contains $j){}#Skip these fields
            else{$formattedValues += "'"+$_.Split(",")[$j]+"',"}#Format everythign else as Strings
            }
        $formattedValues = $formattedValues.Substring(0,$formattedValues.Length-1)
        $sql = "INSERT INTO [$tableName] ($formattedHeaders) VALUES ($formattedValues)"
        #Write-Host -ForegroundColor Yellow $sql
        $sql | Out-File -FilePath $lastCommandFile
        $lastResult = Execute-SQLQueryOnAccessDB -query $sql -queryType NonQuery
        $($i+1) | Out-File -FilePath $recordCountFile
        }
    $i = $i+1
    if($i%10000 -eq 0){Write-Host -ForegroundColor Yellow "$i`t$_"}
    
    }
    