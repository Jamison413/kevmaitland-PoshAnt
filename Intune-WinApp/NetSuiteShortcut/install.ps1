#$thisApp = "NetSuiteShortcut"
function Add-Shortcut {
    param (
        [Parameter(Mandatory)]
        [String]$shortcutTargetPath,
        [Parameter(Mandatory)]
        [String] $destinationPath,
        [Parameter(Mandatory)]
        [String] $iconFileLocation
    )

        $WshShell = New-Object -ComObject ("WScript.Shell")
        $shortcut = $WshShell.CreateShortcut($destinationPath)
        $shortcut.TargetPath = $shortcutTargetPath
        $Shortcut.IconLocation = $iconFileLocation
        # Create the shortcut
        $Shortcut.Save()
        #cleanup
        [Runtime.InteropServices.Marshal]::ReleaseComObject($WshShell) | Out-Null

}

copy-item  '.\NetsuiteIcon.ico' -Destination "$env:USERPROFILE\" -Force
copy-item  'C:\Users\EmilyPressey\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules\Intune-WinApp\NetSuiteShortcut\NetsuiteIcon.ico' -Destination "$env:USERPROFILE\" -Force


#Install now
try{
$shortcutTargetPath = "https://account.activedirectory.windowsazure.com/applications/signin/e7c14b1d-0204-48de-ac8d-4ccb116cfe23?tenantId=271df584-ab64-437f-85b6-80ff9bef6c9f"
$destinationPath = "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\NetSuite (SSO).lnk"
$iconFileLocation = "$env:USERPROFILE\NetsuiteIcon.ico"
Add-Shortcut -DestinationPath $destinationPath -ShortcutTargetPath $ShortcutTargetPath -iconFileLocation $iconFileLocation
}
Catch{
Write-Host $error
}

