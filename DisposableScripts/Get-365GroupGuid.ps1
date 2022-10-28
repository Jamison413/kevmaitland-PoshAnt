connect-ToMsol -Interactive
connect-ToExo -Interactive

Uninstall-Module SharePointPnPPowerShellOnline
Install-Module PnP.PowerShell

Import-Module Pnp.PowerShell
Import-Module ExchangeOnlineManagement

$365GuidINeed = Get-UnifiedGroup -ResultSize Unlimited | Where {($_.DisplayName -match "External - Mattel EPR reporting dashboard")} | Select-Object -ExpandProperty Guid

