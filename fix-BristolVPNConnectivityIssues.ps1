Add-VpnConnectionRoute -ConnectionName "SSL VPN to Anthesis (GBR, Bristol)" -DestinationPrefix "192.168.91.0/24"
Add-VpnConnectionRoute -ConnectionName "SSL VPN to Anthesis (GBR, Bristol)" -DestinationPrefix "192.168.1.0/24"
Add-VpnConnectionRoute -ConnectionName "SSL VPN to Anthesis (GBR, Bristol)" -DestinationPrefix "192.168.22.0/24"
Get-VpnConnection | ? {$_.Name -eq "SSL VPN to Anthesis (GBR, Bristol)"} | % {
    if($_.ConnectionStatus -ne "Connected"){
        Write-host -f Yellow "VPN not connected, reconnecting now"
        rasdial "SSL VPN to Anthesis (GBR, Bristol)" /DISCONNECT
        rasdial "SSL VPN to Anthesis (GBR, Bristol)"
        }
    Start-Sleep -Seconds 5
    if(Test-Connection -IPAddress 192.168.91.13 -Count 1 -ErrorAction SilentlyContinue){
        Write-host -f Yellow "VPN working correctly"
        }
    else{
        Write-host -f Yellow "VPN not routing, reconnecting now"
        rasdial "SSL VPN to Anthesis (GBR, Bristol)" /DISCONNECT
        rasdial "SSL VPN to Anthesis (GBR, Bristol)"
        }
    }
Start-Sleep -seconds 5Start-Sleep -seconds 5