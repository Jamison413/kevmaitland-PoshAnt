Function wait-dedupJobs {
    while ((Get-DedupJob).count -ne 0 ){
        Get-DedupJob
        Start-Sleep -Seconds 30
        }
    }

get-dedupjob | stop-dedupjob

foreach($item in Get-DedupVolume){
    wait-dedupJobs
    $item | Start-DedupJob -Type GarbageCollection -Priority High -Memory 80
    wait-dedupJobs
    $item | Start-DedupJob -Type GarbageCollection -Priority High -Memory 80 -Full
    wait-dedupJobs
    $item | Start-DedupJob -Type Optimization -Priority High -Memory 80
    wait-dedupJobs
    $item | Start-DedupJob -Type Optimization -Priority High -Memory 80 -Full
    wait-dedupJobs
    $item | Start-DedupJob -Type Scrubbing -Priority High -Memory 80
    wait-dedupJobs
    $item | Start-DedupJob -Type Scrubbing -Priority High -Memory 80 -Full
    wait-dedupJobs
    }

Get-DedupStatus