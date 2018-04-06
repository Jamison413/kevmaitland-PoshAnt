#test-connectivity
$ipAddresses = @("8.8.8.8","192.168.254.253","192.168.254.254","127.0.0.1")
$log = "$env:USERPROFILE\connectivity.log"

workflow test-Connectivity($ipAddresses, $log){
    $timestamp = Get-Date
    $timestamp.DateTime | Out-File -FilePath $log -Append
    foreach -parallel ($ip in $ipAddresses){
        $result = Test-Connection -ComputerName $ip -Count 1 -Quiet
        "`t$ip $result" | Out-File $log -Append
        }
    }

test-Connectivity -ipAddresses $ipAddresses -log $log
