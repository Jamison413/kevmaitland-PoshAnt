$sls = Get-WmiObject -Query 'SELECT * FROM SoftwareLicensingService'  

$sls.InstallProductKey($sls.OA3xOriginalProductKey) 

$sls.RefreshLicenseStatus() 