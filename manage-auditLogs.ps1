Import-Module _PS_Library_MSOL
connect-ToExo

$daysToLookBack = 90
if(!$toDate){$toDate = Get-Date}
if(!$fromDate){$fromDate = $toDate.AddDays(-$daysToLookBack)}
[array]$arrayOfUsers = @("sinead.fenton@anthesisgroup.com")
[array]$operations = @("PasswordLogonInitialAuthUsingPassword","UserLoggedIn")
$title = $arrayOfUsers[0]+" "+(get-date)

function parse-unifiedAuditLogToPsObjects($auditLogEntries){
    $auditLogEntries | %{
        [psobject]$event = [psobject]::new()
        $event | Add-Member -MemberType NoteProperty -Name TimeStamp -Value $_.CreationDate -Force
        $event | Add-Member -MemberType NoteProperty -Name User -Value $_.UserIds -Force
        $event | Add-Member -MemberType NoteProperty -Name Event -Value $_.Operations -Force
        foreach($prop in $_.AuditData.Replace("{","").Replace("}","").Replace("`"","").Split(",")){
            $event | Add-Member -MemberType NoteProperty -Name $($prop.Split(":")[0]) -Value $($prop.Replace($prop.Split(":")[0]+":","")) -Force
            }
        [array]$events += $event
        }
    $events
    }
function parse-unifiedAuditLogCsvToPsObjects($pathToAuditLogCsvFile){
    Get-Content -Path $pathToAuditLogCsvFile  | ?{$_.readCount -gt 1} | %{
        [psobject]$event = [psobject]::new()
        foreach($prop in $_.Replace("{","").Replace("}","").Replace("`"","").Split(",")){
            if($prop.Split(":")[0] -match "^[0-9]{4}-(0[1-9]|1[0-2])-(0[1-9]|[1-2][0-9]|3[0-1])"){
                $event | Add-Member -MemberType NoteProperty -Name TimeStamp -Value $prop -Force
                }
            elseif($prop.Split(":")[0] -match "@anthesisgroup.com"){
                $event | Add-Member -MemberType NoteProperty -Name User -Value $prop -Force
                }
            elseif($prop.Split(":")[1] -eq $null){
                $event | Add-Member -MemberType NoteProperty -Name Event -Value $prop -Force
                }
            else{$event | Add-Member -MemberType NoteProperty -Name $($prop.Split(":")[0]) -Value $($prop.Replace($prop.Split(":")[0]+":","")).Replace("\","") -Force}
            }
        [array]$events += $event
        }
    $events
    }


rv log
rv events
do{
    [int]$lastCount = $log.Count
    write-host $lastCount
    $log += Search-UnifiedAuditLog -StartDate $fromDate -EndDate $toDate -UserIds $arrayOfUsers -Operations $operations -SessionId $title -SessionCommand ReturnNextPreviewPage
    }
while($lastCount -ne $log.Count)

$usefulLog =  parse-unifiedAuditLogToPsObjects -auditLogEntries $log
$usefulLog | Out-GridView
$usefulLog  | Export-Csv -Path C:\Users\kevin.maitland\Desktop\logoutput3.csv -NoClobber -NoTypeInformation



$auditLog = "C:\Users\kevin.maitland\Downloads\AuditLog_2017-07-07_2017-10-06.csv"
$usefulLog = parse-unifiedAuditLogCsvToPsObjects -pathToAuditLogCsvFile $auditLog

$usefulLog | %{$_.psobject.Properties.Name | %{if($headers -notcontains $($_+":Dummy")){[array]$headers += $($_+":Dummy")}}}
$usefulLog | %{$_.psobject.Properties.Name | %{if($headers -notcontains $_){[array]$headers += $_}}}
[psobject]$headerObject  = [psobject]::new()
$headers | %{
    $headerObject | Add-Member -MemberType NoteProperty -Name $($_.Split(":")[0]) -Value $($_.Replace($_.Split(":")[0]+":","")).Replace("\","") -Force
    #$headerObject | Add-Member -MemberType NoteProperty -Name $_ -Value $null -Force
    }

[array]$formattedLog += $headerObject
$formattedLog += $usefulLog

$formattedLog  | Export-Csv -Path C:\Users\kevin.maitland\Desktop\UnifiedAuditLog_KayleeShalett_2017-07-07_2017-10-06.csv -NoClobber -NoTypeInformation

$formattedLog | Out-GridView
[System.String]::Join(",",$headers)
$usefulLog.Count