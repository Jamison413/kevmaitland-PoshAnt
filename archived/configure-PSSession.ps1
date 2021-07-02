#Script to configure a PSSessionConfiguration to allow remote users to run the /reload-config command locally on the SharePoint server
#When this script is run, the end-users who will be modifying SharePoint accounts need to be granted Execute(Invoke) permissions. These can be restricted by using a specific Security Group, or can be granted generally by adding the "Domain Users" Security Group
#This alleviates the need for end-users to be local administrators on the SharePoint server by granting them elevated privileges in a restricted way
#
#This should be executed once on the SharePoint Server (and run again whenever the $sharePointServerLocalAdmin password changes)
#
#Requirements:
#- The RunAsCredential account needs to be a local administrator (otherwise the /reload-config command will fail)
#- The password needs to be provided in Plain Text here
#- The $psSessionConfigurationNameOnSharePointServer value should match the value used in the NewUser / ResetUser scripts
#
#If you like the script, feel free to pop a beer in the post c/o Kev Maitland - I work at the head office of www.sustain.co.uk :)


$sharePointServer = "FS03"
$pathToScripts = "C:\Scripts\WhoHasGotItOpen"
$scriptName = "WhoHasGotItOpen.ps1"
$psSessionConfigurationNameOnSharePointServer = "WhoHasGotItOpen"
$SPFarmCredential = Get-Credential -UserName "FS03\Administrator" -Message "Enter the password for SPFarm"

Register-PSSessionConfiguration -Name $psSessionConfigurationNameOnSharePointServer -ApplicationBase $pathToScripts -RunAsCredential $SPFarmCredential -ShowSecurityDescriptorUI -StartupScript "$pathToScripts\$scriptName"
Set-PSSessionConfiguration -Name $psSessionConfigurationNameOnSharePointServer -ApplicationBase $pathToScripts -RunAsCredential $SPFarmCredential -ShowSecurityDescriptorUI -StartupScript ""


#get-PSSessionConfiguration SharePointFarmScripts | Set-PSSessionConfiguration -ShowSecurityDescriptorUI