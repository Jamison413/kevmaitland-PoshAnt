$scriptToDeploy = "test-connectivity.ps1"
$scheduledTaskTemplate = "test-connectivity.xml"
$netlogonPath = "\\$env:USERDOMAIN\NETLOGON" #If the files are in a subfolder called \\BF\NETLOGON\Admin\DeploymentTemplates\, set this to "$env:USERDOMAIN\NETLOGON\Admin\DeploymentTemplates"

#Copy the PowerShell script that we want to run to the local computer (simplifies execution-policies)
Copy-Item "$netlogonPath\$scriptToDeploy" "$env:USERPROFILE\"

#Get the current user's details
$objuser = New-Object System.Security.Principal.NTAccount($env:USERNAME)
$sid = $objuser.Translate([System.Security.Principal.SecurityIdentifier])

#Read in the template XML from a DC, then customise it wuth the local user's details and save it locally
[xml]$xml = Get-Content "$netlogonPath\$scheduledTaskTemplate"
$xml.Task.Principals.InnerXml = $xml.Task.Principals.InnerXml.Replace("S-1-5-21-1963773607-3255835143-1045213775-3594",$sid)
$xml.Task.Triggers.InnerXml = $xml.Task.Triggers.InnerXml.Replace("DOMAIN\DummyUser","$env:USERDOMAIN\$env:USERNAME")
$xml.OuterXml | Out-File -FilePath "$env:USERPROFILE\$scheduledTaskTemplate"

#Create a localised Scheduled Task to run the copied PowerShell script
schtasks /Create /XML "$env:USERPROFILE\$scheduledTaskTemplate" /TN "Test Oxford Connectivity" | Out-Null
