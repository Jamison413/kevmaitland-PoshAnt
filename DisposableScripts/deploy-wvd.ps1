Install-Module Microsoft.RDInfra.RDPowerShell
Import-Module Microsoft.RDInfra.RDPowerShell
Add-RdsAccount  -DeploymentUrl https://rdbroker.wvd.microsoft.com
New-RdsTenant -Name AADDS-WVD -AadTenantId a054308f-5864-479b-a71e-44ac3f05fd4f -AzureSubscriptionId 5fcade80-253b-4bda-91a6-9f1e9792bdfc