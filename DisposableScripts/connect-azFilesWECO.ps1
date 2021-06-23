$driveLetterToMap = "W"
$uncShareToMapTo = "\\gbrenergy.file.core.windows.net\ecodata"
$thisScriptName = "connect-azFilesWECO.ps1"
$scheduledTaskScript = "Write-Host -ForegroundColor Cyan `"Anthesis IT: Connecting $driveLetterToMap`:\ drive to $uncShareToMapTo`"
`$connectTestResult = Test-NetConnection -ComputerName $($([uri]$uncShareToMapTo).Host) -Port 445
if (`$connectTestResult.TcpTestSucceeded) {
    # Mount the drive
    try{New-PSDrive -Name $driveLetterToMap -PSProvider FileSystem -Root $uncShareToMapTo -Persist -ErrorAction Stop}
    catch{
        if(`$_.Exception -match `"local device name is already in use`"){<#Do nothing#>}
        else{`$_}
        }
    }
else {Write-Error -Message 'Unable to reach the Azure storage account via port 445. Check to make sure your organization or ISP is not blocking port 445, or use Azure P2S VPN, Azure S2S VPN, or Express Route to tunnel SMB traffic over a different port.'}
If (Get-PSDrive -Name $driveLetterToMap) {
    Write-Host -ForegroundColor Cyan `"`t$driveLetterToMap`:\ drive mapped successfully.`"
    Add-Content -Path $ScriptLogFilePath -Value ((Get-Date).ToString() + `": `" + `"$driveLetterToMap`:\ drive mapped successfully.`")
    }
Else {
    Write-Host -ForegroundColor Cyan `"`tFailed to map $driveLetterToMap`:\ drive.`"
    Add-Content -Path $ScriptLogFilePath -Value ((Get-Date).ToString() + `": `" + `"Failed to map $driveLetterToMap`:\ drive.`")
    }"

function Test-Administrator {
    $User = [Security.Principal.WindowsIdentity]::GetCurrent();
    (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    }

$ScriptDirectory = $env:APPDATA + "\Intune"
# Check if directory already exists.
if (!(Get-Item -Path $ScriptDirectory)) {
    New-Item -Path $env:APPDATA -Name "Intune" -ItemType "directory"
}

# Logfile
$ScriptLogFilePath = $ScriptDirectory + "\$thisScriptName.log"

if (Test-Administrator) {
    # If running as administrator, create scheduled task as current user.
    Add-Content -Path $ScriptLogFilePath -Value ((Get-Date).ToString() + ": " + "Running as administrator.")

    $ScriptFilePath = $ScriptDirectory + "\$thisScriptName"
    

    $scheduledTaskScript | Out-File -FilePath $ScriptFilePath

    $PSexe = "powershell.exe"
    $Arguments = "-file $ScriptFilePath -WindowStyle Hidden -ExecutionPolicy Bypass"
    $CurrentUser = (Get-CimInstance –ClassName Win32_ComputerSystem | Select-Object -expand UserName)
    $Action = New-ScheduledTaskAction -Execute $PSexe -Argument $Arguments
    $Principal = New-ScheduledTaskPrincipal -UserId (Get-CimInstance –ClassName Win32_ComputerSystem | Select-Object -expand UserName)
    $Trigger = New-ScheduledTaskTrigger -AtLogOn -User $CurrentUser
    $Task = New-ScheduledTask -Action $Action -Trigger $Trigger -Principal $Principal

    Register-ScheduledTask "Anthesis IT - $thisScriptName" -Input $Task -Force
    Start-ScheduledTask "Anthesis IT - $thisScriptName"
    }

Else {
    # Not running as administrator. Connecting directly with Azure script.
    Add-Content -Path $ScriptLogFilePath -Value ((Get-Date).ToString() + ": " + "Not running as administrator.")
    Invoke-Expression $scheduledTaskScript
    }

