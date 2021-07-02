
Import-Module Deduplication
Start-DedupJob -Type Optimization -Volume D: -Memory 50 -Priority High -InputOutputThrottleLevel None
Start-DedupJob -Type GarbageCollection -Volume D: -Memory 50 -Priority High -InputOutputThrottleLevel None
Start-DedupJob -Type GarbageCollection -Volume D: -Memory 50 -Priority High -InputOutputThrottleLevel None -Full
Start-DedupJob -Type GarbageCollection -Volume D: -Memory 50 -Priority High -InputOutputThrottleLevel None
Start-DedupJob -Type Optimization -Volume D: -Memory 50 -Priority High -InputOutputThrottleLevel None -Full
Start-DedupJob -Type Optimization -Volume D: -Memory 50 -Priority High -InputOutputThrottleLevel None

Start-DedupJob -Type Optimization -Priority High -Volume d: -InputOutputThrottleLevel None
Start-DedupJob -Type Optimization -Priority High -Volume D: -Full -InputOutputThrottleLevel None
Start-DedupJob -Type Scrubbing -Priority High -Volume D: -Full -InputOutputThrottleLevel None


 
do 
{
    Write-Output "Dedup jobs are running.  Status:"
    $state = Get-DedupJob | Sort-Object StartTime -Descending 
    $state | ft
    Write-Host -ForegroundColor Yellow "Space on D:\ [$((Get-Volume D).SizeRemaining/(1024*1024*1024)) GB]"
    if ($state -eq $null) {Write-Output "Completing, please wait..."}
    sleep -s 5
} while ($state -ne $null)
 
#cls
Write-Output "Done DeDuping"
Get-DedupStatus | fl *