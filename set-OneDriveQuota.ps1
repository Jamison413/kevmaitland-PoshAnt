
Connect-SPOService -url 'https://anthesisllc-admin.sharepoint.com' -Credential $credential

$url = "https://anthesisllc-my.sharepoint.com/personal/kevin_maitland_anthesisgroup_com"
$url = "https://anthesisllc-my.sharepoint.com/personal/jono_adams_anthesisgroup_com"

Set-SPOSite -Identity $url -StorageQuota 1048576
