Import-Module _PS_Library_MSOL
connect-ToExo

$daysToLookBack = 5
if(!$toDate){$toDate = $(Get-Date).AddDays(1)}
$fromDate = $toDate.AddDays(-($daysToLookBack+1))
[array]$arrayOfUsers = @("Jim.FAva@anthesisgroup.com")
#[array]$operations = @("PasswordLogonInitialAuthUsingPassword","UserLoggedIn")
$title = "$($arrayOfUsers[0])_$(Get-Date -Format yyyy-MM-dd)"



function export-psobjectsToCSV($arrayOfPsobjectAuditEntries){
    $arrayOfAllProperties = @(
        $(New-Object psobject -Property @{"Name"="TimeStamp"}),
        $(New-Object psobject -Property @{"Name"="User"}),
        $(New-Object psobject -Property @{"Name"="Event"})
        )
    $arrayOfPsobjectAuditEntries | %{
        Compare-Object -ReferenceObject $arrayOfAllProperties -DifferenceObject $(Get-Member -InputObject $_ -MemberType NoteProperty) -Property Name -PassThru | ?{$_.SideIndicator -eq "=>"} | % {$arrayOfAllProperties += New-Object psobject -Property @{"Name"=$_.Name}}
        }
    $hashOfAllProperties = [ordered]@{} 
    $arrayOfAllProperties | % {$hashOfAllProperties.Add($_.Name,$null)}
    $fullyMemberedObject = New-Object psobject -Property $hashOfAllProperties
    $nicelyFormattedArray =@()
    $nicelyFormattedArray += $fullyMemberedObject
    $arrayOfPsobjectAuditEntries | % {$nicelyFormattedArray += $_}
    $nicelyFormattedArray | Export-Csv -Path $env:USERPROFILE\Desktop\AuditLogs\AuditLog_$title.csv -NoClobber -NoTypeInformation
    }
function parse-unifiedAuditLogToPsObjects($auditLogEntries){
    $auditLogEntries | %{
        [psobject]$event = [psobject]::new()
        $event | Add-Member -MemberType NoteProperty -Name TimeStamp -Value $_.CreationDate -Force
        $event | Add-Member -MemberType NoteProperty -Name User -Value $_.UserIds -Force
        $event | Add-Member -MemberType NoteProperty -Name Event -Value $_.Operations -Force
        foreach($prop in $_.AuditData.Replace('"ExtendedProperties":',"").Replace('"Name":',"").Replace(',"Value":',":").Replace("{","").Replace("[","").Replace("}","").Replace("]","").Replace("\/","/") -split ',(?=(?:[^"]|"[^"]*")*$)'){
            $event | Add-Member -MemberType NoteProperty -Name $($prop.Split(":")[0].Replace('"','')) -Value $($prop.Replace($prop.Split(":")[0]+":","").Replace('"','')) -Force
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


if ($log){rv log}
do{
    [int]$lastCount = $log.Count
    write-host $lastCount
    $log += Search-UnifiedAuditLog -StartDate $fromDate -EndDate $toDate -UserIds $arrayOfUsers -Operations $operations -SessionId $title -SessionCommand ReturnNextPreviewPage
    }
while($lastCount -ne $log.Count)

$usefulLog =  parse-unifiedAuditLogToPsObjects -auditLogEntries $log
#$usefulLog | Out-GridView
#$usefulLog  | Export-Csv -Path C:\Users\kevin.maitland\Desktop\AuditLog_$($arrayOfUsers[0])_$(Get-Date -Format yyyy-MM-dd).csv -NoClobber -NoTypeInformation
export-psobjectsToCSV -arrayOfPsobjectAuditEntries $usefulLog

<#
$auditLog = "C:\Users\kevin.maitland\Downloads\AuditLog_2017-09-26_2017-10-13.csv"
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

$formattedLog  | Export-Csv -Path C:\Users\kevin.maitland\Desktop\UnifiedAuditLog_MaryShort_2017-09-26_2017-10-13.csv -NoClobber -NoTypeInformation

$formattedLog | Out-GridView
[System.String]::Join(",",$headers)
$usefulLog.Count





        foreach($prop in $log[0].AuditData.Replace('"ExtendedProperties":',"").Replace('"Name":',"").Replace(',"Value":',":").Replace("{","").Replace("[","").Replace("}","").Replace("]","").Replace("\/","/") -split ',(?=(?:[^"]|"[^"]*")*$)')
            {
            Write-Host -ForegroundColor Yellow $($prop.Split(":")[0]) 
            Write-Host -ForegroundColor DarkYellow "`t"$($prop.Replace($prop.Split(":")[0]+":","").Replace("\/","/"))
            }


$RegexOptions = [System.Text.RegularExpressions.RegexOptions]
$csvSplit = '(,)(?=(?:[^"]|"[^"]*")*$)'
$splitColumns = [regex]::Split($log[12].AuditData.Replace("{","").Replace("[","").Replace("}","").Replace("]","").Replace("\/","/"), $csvSplit, $RegexOptions::ExplicitCapture)

$log[12].AuditData.Replace("{","").Replace("[","").Replace("}","").Replace("]","").Replace("\/","/") -split ',(?=(?:[^"]|"[^"]*")*$)'
 #>