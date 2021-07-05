$inputFile = "$env:USERPROFILE\desktop\q_deemedScoresUnique.csv"
$inputFileCombined = "$env:USERPROFILE\desktop\q_deemedScoresUnique_combined.tsv"
#$outputfile = "$env:USERPROFILE\desktop\q_deemedScoresUnique_PASUplift.csv"

#$recordsInFile = 0
#Get-Content -Path $inputFile -ReadCount 100 | % {$recordsInFile += $_.Count}
#$outputArray = @($null) * $recordsInFile

#for ($i=0; $i -lt $recordsInFile; $i++){
    
#    }


$passThrough = Measure-Command {
    $startRamPassThru = [System.GC]::GetTotalMemory($false)
    Import-Csv $inputFile | % {
        $thisRecord = $_
        $thisRecord.upliftName = "PAS20:30 2019"
        $thisRecord.upliftValue = "1.2"
        $thisRecord.deemedCost = [math]::Round([decimal]$thisRecord.deemedCost * 1.2,2)
        #$outputArray[$i]
        #$thisRecord | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | Export-Csv -Path $env:USERPROFILE\desktop\q_deemedScoresUnique_PASUplift.csv -NoClobber -NoTypeInformation -Append
        $thisRecord | Export-Csv -Path $env:USERPROFILE\desktop\q_deemedScoresUnique_PASUplift_foreach.csv  -NoTypeInformation -Append -Force
        }
    $endRamPassThru = [System.GC]::GetTotalMemory($false)
    }


$passThrough = Measure-Command {
    $startRamPassThru = [System.GC]::GetTotalMemory($false)
    Import-Csv $inputFileCombined | % {
        $thisRecord = $_
        $thisRecord.upliftName = $thisRecord.upliftName+"_and_PAS20:30_2019"
        $thisRecord.upliftValue = [decimal]$thisRecord.upliftValue+1.2
        $thisRecord.deemedCost = [math]::Round([decimal]$thisRecord.deemedCost * $thisRecord.upliftValue,2)
        #$outputArray[$i]
        #$thisRecord | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | Export-Csv -Path $env:USERPROFILE\desktop\q_deemedScoresUnique_PASUplift.csv -NoClobber -NoTypeInformation -Append
        $thisRecord | Export-Csv -Path $env:USERPROFILE\desktop\q_deemedScoresUnique_PASUplift_foreach_combined.csv  -NoTypeInformation -Append -Force
        }
    $endRamPassThru = [System.GC]::GetTotalMemory($false)
    }


<#$inMemory = Measure-Command {
    $startRamInMem = [System.GC]::GetTotalMemory($false)
    $massiveInputCsv = Import-Csv $inputFile
    $massiveInputCsv | % {
        $thisRecord = $_
        $thisRecord.upliftName = "PAS20:30 2019"
        $thisRecord.upliftValue = "1.2"
        $thisRecord.deemedCost = [math]::Round([decimal]$thisRecord.deemedCost * 1.2,2)
        }
    $massiveInputCsv | % {Export-Csv -InputObject $_ -Path $env:USERPROFILE\desktop\q_deemedScoresUnique_PASUplift_inMemory.csv  -NoTypeInformation -Force -Append}
    $endRamInMem = [System.GC]::GetTotalMemory($false)
    }
#>    
Write-Host "PassThru:`t[$($passThrough.TotalSeconds)] seconds `t[$([math]::Round(($endRamPassThru-$startRamPassThru)/(1024*1024),2))] MB deltaRAM"
Write-Host "InMemory:`t[$($inMemory.TotalSeconds)] seconds `t[$([math]::Round(($endRamInMem-$startRamInMem)/(1024*1024),2))] MB deltaRAM"
