#Generate install, uninstall and detect .PS1 scripts for Choco-managed software
$appFriendlyName = "Paint.Net (managed)"              #e.g. "Adobe Reader (managed)"
$appChocoName = "Paint.Net"                 #e.g. "AdobeReader"
$installedAppFileForDetection = "${env:ProgramFiles}\Paint.Net\PaintDotNet.exe" #e.g. "${env:ProgramFiles(x86)}\Adobe\Acrobat Reader DC"

#$test = Start-Process -FilePath "$env:ProgramData\chocolatey\choco.exe" -ArgumentList "list $appChocoName" -PassThru
#$test = &(choco list $appChocoName)

#region Create Files
if(!(Test-Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules\Intune-WinApp\$appChocoName")){
    New-Item -Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules\Intune-WinApp\" -Name $appChocoName -ItemType Directory
    }
if(Test-Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules\Intune-WinApp\$appChocoName"){
    if(Test-Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules\Intune-WinApp\$appChocoName\install.ps1"){Remove-Item "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules\Intune-WinApp\$appChocoName\install.ps1" -Force}
    Add-Content -Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules\Intune-WinApp\$appChocoName\install.ps1" -Value "`$thisApp = `"$appChocoName`""
    Get-Content -Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules\Intune-WinApp\_example\install.ps1" | Add-Content "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules\Intune-WinApp\$appChocoName\install.ps1"

    if(Test-Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules\Intune-WinApp\$appChocoName\uninstall.ps1"){Remove-Item "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules\Intune-WinApp\$appChocoName\uninstall.ps1" -Force}
    Add-Content -Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules\Intune-WinApp\$appChocoName\uninstall.ps1" -Value "`$thisApp = `"$appChocoName`""
    Get-Content -Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules\Intune-WinApp\_example\uninstall.ps1" | Add-Content "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules\Intune-WinApp\$appChocoName\uninstall.ps1"

    if(Test-Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules\Intune-WinApp\$appChocoName\detect.ps1"){Remove-Item "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules\Intune-WinApp\$appChocoName\detect.ps1" -Force}
    Add-Content -Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules\Intune-WinApp\$appChocoName\detect.ps1" -Value "`$thisApp = `"$appChocoName`""
    Add-Content -Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules\Intune-WinApp\$appChocoName\detect.ps1" -Value "`$filePathToTest = `"$installedAppFileForDetection`""
    Get-Content -Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules\Intune-WinApp\_example\detect.ps1" | Add-Content "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules\Intune-WinApp\$appChocoName\detect.ps1"

    Start-Process -NoNewWindow -FilePath  "$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules\Intune-WinApp\IntuneWinAppUtil.exe" -ArgumentList "-c `"$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules\Intune-WinApp\$appChocoName`"","-s install.ps1","-o `"$env:USERPROFILE\OneDrive - Anthesis LLC\Documents\WindowsPowerShell\Modules\Intune-WinApp\$appChocoName`"","-q"
    }
#endregion

#region Create Security Group
$tokenResponseTeamBot = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName TeamsBot) -grant_type client_credentials
$newSG = new-graphGroup -tokenResponse $tokenResponseTeamBot -groupDisplayName "Software - $appFriendlyName" -groupDescription "Security Group for deploying Choco-managed app [$appChocoName]" -groupType Security -membershipType Assigned
#endregion


#region Create Intune Win32 App
$tokenResponseIntuneBot = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName IntuneBot) -grant_type client_credentials
Write-Host -ForegroundColor Yellow "Manually create a Win32 app here: `r`n`thttps://endpoint.microsoft.com/#blade/Microsoft_Intune_DeviceSettings/AppsWindowsMenu/windowsApps"
<#The managedApp endpoint doesn't (as of 2020-10-19) let you create these via Graph (https://docs.microsoft.com/en-us/graph/api/resources/intune-apps-managedapp?view=graph-rest-1.0) but an object looks like this:
{
    "@odata.type":  "#microsoft.graph.win32LobApp",
    "id":  "c5ade2d4-2f6a-47df-b0ef-1b8da11b8089",
    "displayName":  "GitHub Desktop (managed)",
    "description":  "Code management tool",
    "publisher":  "Git",
    "largeIcon":  null,
    "createdDateTime":  "2020-10-15T13:00:21.9877002Z",
    "lastModifiedDateTime":  "2020-10-15T13:00:21.9877002Z",
    "isFeatured":  false,
    "privacyInformationUrl":  "",
    "informationUrl":  "https://desktop.github.com/",
    "owner":  "Anthesis",
    "developer":  "",
    "notes":  "",
    "publishingState":  "published",
    "committedContentVersion":  "1",
    "fileName":  "install.intunewin",
    "size":  1376,
    "installCommandLine":  "powershell.exe -executionpolicy bypass .\\install.ps1",
    "uninstallCommandLine":  "powershell.exe -executionpolicy bypass .\\uninstall.ps1",
    "applicableArchitectures":  "x86,x64",
    "minimumFreeDiskSpaceInMB":  null,
    "minimumMemoryInMB":  null,
    "minimumNumberOfProcessors":  null,
    "minimumCpuSpeedInMHz":  null,
    "msiInformation":  null,
    "setupFilePath":  "install.ps1",
    "minimumSupportedWindowsRelease":  "1803",
    "rules":  [
                  {
                      "@odata.type":  "#microsoft.graph.win32LobAppPowerShellScriptRule",
                      "ruleType":  "detection",
                      "displayName":  null,
                      "enforceSignatureCheck":  false,
                      "runAs32Bit":  false,
                      "runAsAccount":  null,
                      "scriptContent":  "77u/JHRoaXNBcHAgPSAiR2l0SHViLURlc2t0b3AiDQokZmlsZVBhdGhUb1Rlc3QgPSAiJGVudjpMT0NBTEFQUERBVEFcR2l0SHViRGVza3RvcFxHaXRIdWJEZXNrdG9wLmV4ZSINCiRzY2hlZHVsZWRUYXNrVG9UZXN0ID0gIkFudGhlc2lzIElUIC0gQ2hvY28gSW50YWxsT3JVcGdyYWRlICR0aGlzQXBwIg0KDQppZihUZXN0LV
BhdGggJGZpbGVQYXRoVG9UZXN0KXsNCiAgICBpZihHZXQtU2NoZWR1bGVkVGFzayAtVGFza05hbWUgJHNjaGVkdWxlZFRhc2tUb1Rlc3Qpew0KICAgICAgICBXcml0ZS1Ib3N0ICJQYXRoIFskKCRmaWxlUGF0aFRvVGVzdCldIGFuZCBUYXNrIFskKCRzY2hlZHVsZWRUYXNrVG9UZXN0KV0gYm90aCBleGlzdCEiDQogICAgICAgIH0NCiAgICB9DQplbHNlew0KICAgIFRocm93ICJQY
XRoIFskKCRmaWxlUGF0aFRvVGVzdCldIG5vdCBmb3VuZCAtIFskKCR0aGlzQXBwKV0gaXMgbm90IGluc3RhbGxlZCINCiAgICB9",
                      "operationType":  "notConfigured",
                      "operator":  "notConfigured",
                      "comparisonValue":  null
                  }
              ],
    "installExperience":  {
                              "runAsAccount":  "system",
                              "deviceRestartBehavior":  "allow"
                          },
    "returnCodes":  [
                        {
                            "returnCode":  0,
                            "type":  "success"
                        },
                        {
                            "returnCode":  1707,
                            "type":  "success"
                        },
                        {
                            "returnCode":  3010,
                            "type":  "softReboot"
                        },
                        {
                            "returnCode":  1641,
                            "type":  "hardReboot"
                        },
                        {
                            "returnCode":  1618,
                            "type":  "retry"
                        }
                    ]
} #>
#endregion