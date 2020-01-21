Install-Module -Name MSCommerce 
Import-Module MSCommerce
Connect-MSCommerce

Get-MSCommercePolicy -PolicyId AllowSelfServicePurchase | fl  
Get-MSCommerceProductPolicies -PolicyId AllowSelfServicePurchase 
Get-MSCommerceProductPolicies -PolicyId AllowSelfServicePurchase | % {Update-MSCommerceProductPolicy -PolicyId AllowSelfServicePurchase -ProductId $_.ProductId -Enabled $False}