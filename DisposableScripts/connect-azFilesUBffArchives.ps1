$driveLetterToMap = "U"
$uncShareToMapTo = "https://gbrbff.file.core.windows.net/share"
$thisScriptName = "connect-azFilesUBffArchvies.ps1"
$ScriptDirectory = $env:APPDATA + "\Intune"
# Check if directory already exists.
if (!(Get-Item -Path $ScriptDirectory)) {
    New-Item -Path $env:APPDATA -Name "Intune" -ItemType "directory"
    }
$ScriptLogFilePath = $ScriptDirectory + "\$thisScriptName.log"



Write-Host -ForegroundColor Cyan "Anthesis IT: Connecting $driveLetterToMap`:\ drive to $uncShareToMapTo"
$connectTestResult = Test-NetConnection -ComputerName $($([uri]$uncShareToMapTo).Host) -Port 445
if ($connectTestResult.TcpTestSucceeded) {
    # Mount the drive
    try{New-PSDrive -Name $driveLetterToMap -PSProvider FileSystem -Root $uncShareToMapTo -Persist -ErrorAction Stop}
    catch{
        if($_.Exception -match "local device name is already in use"){<#Do nothing#>}
        else{Add-Content -Path $ScriptLogFilePath -Value $_}
        }
    }
else {Write-Error -Message 'Unable to reach the Azure storage account via port 445. Check to make sure your organization or ISP is not blocking port 445, or use Azure P2S VPN, Azure S2S VPN, or Express Route to tunnel SMB traffic over a different port.'}
If (Get-PSDrive -Name $driveLetterToMap) {
    Write-Host -ForegroundColor Cyan "t$driveLetterToMap`:\ drive mapped successfully."
    Add-Content -Path $ScriptLogFilePath -Value ((Get-Date).ToString() + ": " + "$driveLetterToMap`:\ drive mapped successfully.")
    }
Else {
    Write-Host -ForegroundColor Cyan "tFailed to map $driveLetterToMap`:\ drive."
    Add-Content -Path $ScriptLogFilePath -Value ((Get-Date).ToString() + ": " + "Failed to map $driveLetterToMap`:\ drive.")
    }
