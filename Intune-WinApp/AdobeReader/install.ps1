$thisApp = "AdobeReader"
#$thisApp = "%%PLACEHOLDERAPPNAME%%"
#Install now
$localprograms = choco list --localonly
if ($localprograms -like "*$thisApp*"){
    choco upgrade $thisApp -y
    }
Else{
    choco install $thisApp -y
    }

#Add Scheduled Task to manage this app
$content = @"
    `$localprograms = choco list --localonly
    if (`$localprograms -like "*$thisApp*"){
        `$action = "upgrade"
        }
    Else{
        `$action = "install"
        }
    choco `$action $thisApp -y --logfile=$("`$env:ProgramData\CustomScripts\choco_$thisApp.log")
    if ((Get-WmiObject -Class Win32_NTEventLOgFile | Select-Object FileName, Sources | ForEach-Object -Begin { `$hash = @{}} -Process { `$hash[`$_.FileName] = `$_.Sources } -end { `$Hash })["Application"] -notcontains "Anthesis IT"){
        New-EventLog -LogName Application -Source "Anthesis IT" #Add Anthesis IT Application if required
        }
    switch (`$LASTEXITCODE){
        "0" {Write-EventLog -LogName Application -Source "Anthesis IT" -EntryType Information -Message "$thisApp `$action processed successfully" -EventId "26848"}
        "1641" {Write-EventLog -LogName Application -Source "Anthesis IT" -EntryType Warning -Message "$thisApp `$action processed successfully - reboot initiated" -EventId "26848"}
        "3010" {Write-EventLog -LogName Application -Source "Anthesis IT" -EntryType Information -Message "$thisApp `$action processed successfully - reboot required" -EventId "26848"}
        default {Write-EventLog -LogName Application -Source "Anthesis IT" -EntryType Error -Message "$thisApp `$action failed with Exit Code `$LASTEXITCODE" -EventId "26848"}
        }
"@

 
# create custom folder and write PS script
$path = $(Join-Path $env:ProgramData CustomScripts)
if (!(Test-Path $path)){
    New-Item -Path $path -ItemType Directory -Force -Confirm:$false
    }
Out-File -FilePath $(Join-Path $env:ProgramData CustomScripts\redo-choco$thisApp.ps1) -Encoding unicode -Force -InputObject $content -Confirm:$false
 
# register script as scheduled task
$Time = New-ScheduledTaskTrigger -At 12:00 -Daily
$User = "SYSTEM"
$Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ex bypass -file `"C:\ProgramData\CustomScripts\redo-choco$thisApp.ps1`""
Register-ScheduledTask -TaskName "Anthesis IT - Choco IntallOrUpgrade $thisApp" -Trigger $Time -User $User -Action $Action -Force
if(Get-ScheduledTask -TaskName "Anthesis IT - Choco IntallOrUpgrade $thisApp"){
    Out-File -FilePath $(Join-Path $env:ProgramData CustomScripts\redo-choco$thisApp-scheduledTaskCreated.log) -Encoding unicode -Force -InputObject $(Get-Date) -Confirm:$false
    }
