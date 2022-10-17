$logFileLocation = "C:\ScriptLogs\purge-RecycleBin\"
$transcriptLogName = "$($logFileLocation)purge-RecycleBins_Transcript_$(Get-Date -Format "yyyy-MM-dd").log"
Start-Transcript $transcriptLogName -Append

$sharePointAdmin = "kimblebot@anthesisgroup.com"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Downloads\KimbleBot.txt)
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/clients" -Credentials $adminCreds 

function Clear-RecycleBinItem {
    param(
        [Parameter(Mandatory)]
        [String[]]
        $Ids
    )
    
    $siteUrl = (Get-PnPSite).Url
    $apiCall = $siteUrl + "/clients/_api/site/RecycleBin/DeleteByIds"
    $body = "{""ids"":[""$($Ids -join '","')""]}"
    #$body = "{""ids"":[$($Ids -join ",")]}"

    Write-Verbose "Performing API Call to delete item from RecycleBin..."
    Write-Verbose $body
    try {
      Invoke-PnPSPRestMethod -Method Post -Url $apiCall -Content $body | Out-Null
    }
    catch {
        Write-Error "Unable to Delete ID {$Ids}"     
    }
}


$duration = Measure-Command {
    $rbin = Get-PnPRecycleBinItem -RowLimit 10000
    $rbinToDelete = $rbin | ? {$_.DeletedByEmail -eq "t0-kevin.maitland@anthesisgroup.com"}
    Write-Output "[$($rbinToDelete.Count)] items deleted by t0-kevin.maitland@anthesisgroup.com [$([Math]::Round($($rbinToDelete | Measure-Object -Property Size -Sum).Sum/1GB,2))]GB recoverable"
    Write-Verbose "[$($rbinToDelete.Count)] items deleted by t0-kevin.maitland@anthesisgroup.com [$([Math]::Round($($rbinToDelete | Measure-Object -Property Size -Sum).Sum/1GB,2))]GB recoverable"
    $storageRecoveredInTotal = 0
    for ($i=0; $i -lt $rbinToDelete.Count; $i++){
        Write-Progress -activity "Purging RecycleBin [$([Math]::Round($storageRecoveredInTotal/1GB,2))] GB" -Status "[$i/$($rbinToDelete.count)]" -PercentComplete ($i/ $rbinToDelete.count *100) -Id 1 -CurrentOperation "Purging [$($rbinToDelete[$i].DirName+"/"+$rbinToDelete[$i].LeafName)]"
        if($lastFilePurged -eq $($rbinToDelete[$i].DirName+"/"+$rbinToDelete[$i].LeafName) -and $versionsPurged -lt 99){
            $versionsPurged++
            [array]$idsToPurge += $rbinToDelete[$i].Id
            }
        else{
            Write-Output "Purging [$($lastFilePurged)]"
            Write-Verbose "Purging [$($lastFilePurged)]"
            $VerbosePreference = 0
            Clear-RecycleBinItem -Ids $idsToPurge
            $VerbosePreference = 2
            Write-Output "`t[$($versionsPurged)] versions purged [$([Math]::Round($storageRecoveredFromThisFile/1GB,2))] GB"
            Write-Verbose "`t[$($versionsPurged)] versions purged [$([Math]::Round($storageRecoveredFromThisFile/1GB,2))] GB"
            $storageRecoveredFromThisFile = 0
            $versionsPurged = 1
            $idsToPurge = @($rbinToDelete[$i].Id)
            $lastFilePurged = $($rbinToDelete[$i].DirName+"/"+$rbinToDelete[$i].LeafName)
            }
    
        #$rbinToDelete[$i] | Clear-PnPRecycleBinItem -Force 
        #Clear-RecycleBinItem -Ids $rbinToDelete[$i].Id
        $storageRecoveredInTotal += $rbinToDelete[$i].Size
        $storageRecoveredFromThisFile += $rbinToDelete[$i].Size
        }
    }
Write-Output "[$($rbinToDelete.Count)] items deleted, [$([Math]::Round($storageRecoveredInTotal/1GB,2))]GB recovered, in [$($duration.TotalMinutes)] minutes"

Stop-Transcript 
