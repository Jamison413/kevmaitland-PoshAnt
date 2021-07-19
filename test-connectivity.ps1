#test-connectivity
$myIntenetConnection = Get-NetIPConfiguration | ? {![string]::IsNullOrWhiteSpace($_.IPv4DefaultGateway)} #This won't work reliably if a device uses IPv6, or has multiple Default Gateways, but most won't
$externalIp = "8.8.8.8" #Google DNS server - leave this
$myDefaultGateway = $myIntenetConnection.IPv4DefaultGateway
$myDNS = $myIntenetConnection.DNSServer.ServerAddresses | select -First 1
$myIP = $myIntenetConnection.IPv4Address
$localLoopBack = "127.0.0.1"

$ipAddresses = @($externalIp,$myDefaultGateway,$myDNS,$myIP,$localLoopBack)
$log = "$env:USERPROFILE\connectivity.log"

workflow test-Connectivity($ipAddresses, $log){
    $timestamp = Get-Date
    $timestamp.DateTime | Out-File -FilePath $log -Append
    foreach -parallel ($ip in $ipAddresses){
        $result = Test-Connection -ComputerName $ip -Count 6000 -Quiet
        "`t$ip $result" | Out-File $log -Append
        }
    }

test-Connectivity -ipAddresses $ipAddresses -log $log

Get-NetIPConfiguration | Foreach IPv4DefaultGateway
