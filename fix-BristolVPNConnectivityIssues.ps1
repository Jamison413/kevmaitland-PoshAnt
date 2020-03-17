Add-VpnConnectionRoute -ConnectionName "Sustain AlwaysOn VPN" -DestinationPrefix "192.168.91.0/24"
Add-VpnConnectionRoute -ConnectionName "Sustain AlwaysOn VPN" -DestinationPrefix "192.168.1.0/24"
Add-VpnConnectionRoute -ConnectionName "Sustain AlwaysOn VPN" -DestinationPrefix "192.168.22.0/24"
Get-VpnConnection | ? {$_.Name -eq "Sustain AlwaysOn VPN"} | % {
    if($_.ConnectionStatus -ne "Connected"){
        Write-host -f Yellow "VPN not connected, reconnecting now"
        rasdial "Sustain AlwaysOn VPN" /DISCONNECT
        rasdial "Sustain AlwaysOn VPN"
        }
    Start-Sleep -Seconds 5
    if(Test-Connection -IPAddress 192.168.91.13 -Count 1 -ErrorAction SilentlyContinue){
        Write-host -f Yellow "VPN working correctly"
        }
    else{
        Write-host -f Yellow "VPN not routing, reconnecting now"
        rasdial "Sustain AlwaysOn VPN" /DISCONNECT
        rasdial "Sustain AlwaysOn VPN"
        }
    }
Start-Sleep -seconds 5