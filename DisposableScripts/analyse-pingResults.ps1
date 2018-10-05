<#$results = @{}
gci $env:USERPROFILE\Desktop\PingLogs | %{
    $thisLog = $_
    $results.Add($thisLog.BaseName,[ordered]@{})
    gc $thisLog.FullName | % {
        if ($_[0] -eq "`t"){
            if($_.Split(" ")[1] -eq "True"){
                $results[$thisLog.BaseName][$currentTimestamp]["true"] += $_.Split(" ")[0].Trim()
                }
            else{
                $results[$thisLog.BaseName][$currentTimestamp]["false"] += $_.Split(" ")[0].Trim()
                }

            #$helen[$current].Add($_.Split(" ")[0],$_.Split(" ")[1])
            }
        else{
            $currentTimestamp = $(Get-Date $_ -Format s)
            $results[$thisLog.BaseName].Add($currentTimestamp,@{})
            $results[$thisLog.BaseName][$currentTimestamp].Add("true",@())
            $results[$thisLog.BaseName][$currentTimestamp].Add("false",@())
            }
        }

    }
foreach($userLog in $results.Keys){
    foreach ($Key in ($results[$userLog].GetEnumerator() | ? {$_.Value["false"] -ne $null})){
        write-host "$userlog`t$($Key.name)`t$($results[$userLog][$Key.Name]["false"] -join ", ")"
        }
    }
#>

$results2 = @()
gci $env:USERPROFILE\Desktop\PingLogs -Filter "*.log" | %{
    $thisLog = $_
    $i=1;$j=1
    Write-Host -ForegroundColor Yellow "$($thisLog.BaseName)"
    gc $thisLog.FullName | % {
        if ($_[0] -ne "`t"){
            $results2 += $pingObj

            $currentTimestamp = $(Get-Date $_ -Format s)
            if($i%100 -eq 0){Write-Host -ForegroundColor DarkYellow "$($thisLog.BaseName)`tDateStamp:`t[$_][$i][$currentTimestamp]"}
            $pingObj = New-Object psobject -Property @{"DateStamp"=$currentTimestamp;"User"=$thisLog.BaseName;"Loopback"=$null;"Self"=$null;"192.168.254.248"=$null;"192.168.254.253"=$null;"192.168.254.254"=$null;"8.8.8.8"=$null}
            #$results2.Add($currentTimestamp,@())
            $j=1
            $i++
            }
        else{
            #Write-Host -ForegroundColor DarkYellow "PingLog:`t[$_][$($_.Split(" ")[0].Trim())][$j]"
            $pingResult = $_.Split(" ")[1]
            switch ($_.Split(" ")[0].Trim()){
                ("127.0.0.1")       {$pingObj.Loopback = $pingResult}
                ("192.168.254.248") {$pingObj.'192.168.254.248' = $pingResult}
                ("192.168.254.253") {$pingObj.'192.168.254.253' = $pingResult}
                ("192.168.254.254") {$pingObj.'192.168.254.254' = $pingResult}
                ("8.8.8.8")         {$pingObj.'8.8.8.8' = $pingResult}
                default             {$pingObj.Self = $pingResult}
                }
            $j++
            }
        }
    }



$results2 | ? {$_.'8.8.8.8' -ne $null}| sort DateStamp | select DateStamp,User,Loopback,Self,192.168.254.248,192.168.254.253,192.168.254.254,8.8.8.8 | Export-Csv -Path $env:USERPROFILE\Desktop\PingLogs\Anaysis.csv -NoTypeInformation