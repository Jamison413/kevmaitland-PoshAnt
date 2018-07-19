(Get-SPOSite -Identity https://anthesisllc.sharepoint.com/teams/DummyTeam3_All_365).SharingCapability
ExistingExternalUserSharingOnly
https://anthesisllc.sharepoint.com/teams/DummyTeam3_All_365/

get-help set-sposite -Detailed

New-SPOSite

$settings = Get-AzureADDirectorySetting | where-object {$_.displayname -eq “Group.Unified”}
$settings["ClassificationList"] = "Internal,External,Confidential"
$settings["ClassificationDescriptions"] = "Internal:This is internal only,External:External users can access,Confidential:Highly secure"
Set-AzureADDirectorySetting -Id $settings.Id -DirectorySetting $settings