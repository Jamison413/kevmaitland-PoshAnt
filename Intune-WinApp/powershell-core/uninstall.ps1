$thisApp = "powershell-core"
#$thisApp = "%%PLACEHOLDERAPPNAME%%"
choco uninstall $thisApp -y
Unregister-ScheduledTask -TaskName "Anthesis IT - Choco IntallOrUpgrade $thisApp"  -Confirm:$false