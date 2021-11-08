function add-registryValue(){
    [cmdletbinding()]
        <#
    .SYNOPSIS
    

    .DESCRIPTION
    

    .PARAMETER infile1
    

    .PARAMETER infile2
    

    .EXAMPLE
    #>

    param(
        [parameter(Mandatory=$true)]
            [String]$registryPath
        ,[parameter(Mandatory=$true)]
            [String]$registryKey
        ,[parameter(Mandatory=$true)]
            [String]$registryValue
        ,[parameter(Mandatory=$true)]
            [ValidateSet("String", "ExpandString", "Binary", "DWord", "MultiString", "QWord")] 
            [String]$registryType
        )

    $registryPath = $registryPath.Replace("Computer\","")
    $registryPath = $registryPath.Replace("HKEY_LOCAL_MACHINE","HKLM:")
    $registryPath = $registryPath.Replace("HKEY_CURRENT_USER","HKCU:")

    $registryPath -split "\\" | % { #Silently create any missing subkeys
        $thisStub = $thisStub += $_+"\"
        if((Test-Path $thisStub) -eq $false){
            Write-Verbose "Sliently creating [$($thisStub)]"
            New-Item -Path $thisStub | Out-Null
            }
        }

    try {$existingItem = Get-ItemProperty -Path $registryPath -name $registryKey -ErrorAction SilentlyContinue}
    catch{}
    
    Write-Verbose "`$registryPath = $registryPath"
    Write-Verbose "`$registryKey = $registryKey"
    Write-Verbose "`$registryValue = $registryValue"
    Write-Verbose "`$registryType = $registryType"
    Write-Verbose "`$existingItem = $existingItem"

    if([string]::IsNullOrWhiteSpace($existingItem.$registryKey)){
        Write-Verbose "Creating new Registry Value"
        New-ItemProperty -Path $registryPath -Name $registryKey -Value $registryValue -PropertyType $registryType
        }
    else{
        Write-Verbose "Updating existing Registry Value"
        Set-ItemProperty -Path $registryPath -Name $registryKey -Value $registryValue
        }
    
    
    }

#Test if VDI version is already installed:
$vdiRegKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Teams" -Name "IsWVDEnvironment" -ErrorAction SilentlyContinue

if($vdiRegKey -eq $null -or $vdiRegKey.IsWVDEnvironment -eq "0"){ #If not, reinstall Teams using the latest installer and the VDI switches
    #region Remove Teams Machine-Wide Installer
    Write-Host "Removing Teams Machine-wide Installer" -ForegroundColor Yellow
    $MachineWide = Get-WmiObject -Class Win32_Product | Where-Object{$_.Name -eq "Teams Machine-Wide Installer"}
    $MachineWide.Uninstall()
    #endregion

    #region Remove Teams for Current Users (from https://docs.microsoft.com/en-us/microsoftteams/scripts/powershell-script-deployment-cleanup)
    $TeamsPath = [System.IO.Path]::Combine($env:LOCALAPPDATA, 'Microsoft', 'Teams')
    $TeamsUpdateExePath = [System.IO.Path]::Combine($TeamsPath, 'Update.exe')
    try{
        if ([System.IO.File]::Exists($TeamsUpdateExePath)) {
            Write-Host "Uninstalling Teams process"

            # Uninstall app
            $proc = Start-Process $TeamsUpdateExePath "-uninstall -s" -PassThru
            $proc.WaitForExit()
            }
        Write-Host "Deleting Teams directory"
        Remove-Item –path $TeamsPath -recurse
        }
    catch{
        Write-Output "Uninstall failed with exception $($_.exception.message)"
        exit /b 1
        }
    #endregion
    
    #Download latest Teams installer (linked from from https://docs.microsoft.com/en-us/microsoftteams/teams-for-vdi#deploy-the-teams-desktop-app-to-the-vm)
    Invoke-WebRequest -Uri "https://teams.microsoft.com/downloads/desktopurl?env=production&plat=windows&arch=x64&managedInstaller=true&download=true" -OutFile "$env:TEMP\TeamsInstaller.msi"

    #region Set RegKey and install latest version of Teams with correct switches (Guide: https://docs.microsoft.com/en-us/microsoftteams/teams-for-vdi)
    add-registryValue -registryPath "HKLM:\SOFTWARE\Microsoft\Teams" -registryKey IsWVDEnvironment -registryValue 1 -registryType DWord
    try {
        $process = Start-Process -FilePath "$env:TEMP\TeamsInstaller.msi" -ArgumentList "/l*v teamsVdiInistaller.Log ALLUSER=1 ALLUSERS=1" -PassThru -Wait -ErrorAction STOP
        if ($process.ExitCode -ne 0){
            Write-Error "Uninstallation failed with exit code  $($process.ExitCode)."
            add-registryValue -registryPath "HKLM:\SOFTWARE\Microsoft\Teams" -registryKey IsWVDEnvironment -registryValue 0 -registryType DWord
            }
        }
    catch {
        Write-Error $_.Exception.Message
        add-registryValue -registryPath "HKLM:\SOFTWARE\Microsoft\Teams" -registryKey IsWVDEnvironment -registryValue 0 -registryType DWord
        }
    }
