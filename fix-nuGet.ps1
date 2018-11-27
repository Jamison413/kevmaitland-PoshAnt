Unregister-PackageSource -Name nuget.org 
Register-PackageSource -Location https://www.nuget.org/api/v2 -Name nuget.org -Trusted -ProviderName NuGet
Find-Package -Name Microsoft.SharePointOnline.CSOM | Install-Package
