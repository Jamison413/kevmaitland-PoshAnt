Import-Module _PS_Library_MSOL
connect-ToExo

$daysToLookBack = 2
if(!$toDate){$toDate = $(Get-Date).AddDays(1)}
$fromDate = $toDate.AddDays(-($daysToLookBack+1))
[array]$arrayOfUsers = @("t1-rae.victorio@anthesisgroup.com")
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
    $nicelyFormattedArray | Export-Csv -Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\AuditLogs\AuditLog_$title.csv" -NoClobber -NoTypeInformation
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
        $thisLog = $_
        [psobject]$event = [psobject]::new()
        $event | Add-Member -MemberType NoteProperty -Name TimeStamp -Value $thisLog.Split(",")[0] -Force
        $event | Add-Member -MemberType NoteProperty -Name User -Value $thisLog.Split(",")[1] -Force
        $event | Add-Member -MemberType NoteProperty -Name Event -Value $thisLog.Split(",")[2] -Force
        $remainingEvent = $thisLog.Substring($thisLog.Split(",")[0].Length + $thisLog.Split(",")[1].Length + $thisLog.Split(",")[2].Length +4).Trim('"')
        $remainingEvent = $remainingEvent.Replace('""""',"''")
        $remainingEvent = $remainingEvent.Replace('""','"')
        [array]$eventObjects += ConvertFrom-Json $remainingEvent
        }
    $eventObjects
    }
function export-psObjectsToBetterCsv(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
            [psobject[]]$inputPsObjects
        ,[parameter(Mandatory = $true)]
            [string]$outputCsvFilePath
        ,[parameter(Mandatory = $false)]
            [switch]$fullyFlattenObjects = $false
        )
    
    $headers = get-propertyNames -inputPsObjects $inputPsObjects -fullyFlattenObjects:$true | Sort-Object
    $headers | % {
        if($_ -match "."){
            $_ = {"$_.$($_)"}
            }
        }
    $inputPsObjects[0] | Select-Object -Property $headers | Export-Csv -NoTypeInformation -Path $outputCsvFilePath
    
    $headers[1]

    }
function Flatten-Array{
    $input | ForEach-Object{
        if ($_ -is [array]){$_ | Flatten-Array}else{$_}
    } | Where-Object{![string]::IsNullorEmpty($_)}
    # | Where-Object{$_} would also work.
}
function get-propertyNames(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
            [psobject[]]$inputPsObjects
        ,[parameter(Mandatory = $false)]
            [switch]$fullyFlattenObjects = $false
        ,[parameter(Mandatory = $false)]
            [string]$propertyPrefix
        )

    $properties = $inputPsObjects | Get-Member | ? {$_.MemberType -eq "NoteProperty"} | Select-Object -Unique | Flatten-Array 
    @($properties | Select-Object) | % {
        $thisProperty = $_
        <#if($fullyFlattenObjects){
            if($thisProperty.Definition -match 'Object\[\]'){
                get-propertyNames -inputPsObjects $inputPsObjects.$($thisProperty.Name) -fullyFlattenObjects -propertyPrefix "$($propertyPrefix)$($thisProperty.Name)."
                }
            }
        "$($propertyPrefix)$($thisProperty.Name)"
        }#>
        if($thisProperty.Definition -match 'Object\[\]'){
            $thisProperty -join ","
            get-propertyNames -inputPsObjects $inputPsObjects.$($thisProperty.Name) -fullyFlattenObjects -propertyPrefix "$($propertyPrefix)$($thisProperty.Name)."
            }
        else{
            "$($propertyPrefix)$($thisProperty.Name)"
            }
        }
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

$re3 = parse-unifiedAuditLogCsvToPsObjects -pathToAuditLogCsvFile C:\Users\kevinm\Desktop\AuditLogs\AuditLog_2018-10-31_2019-01-30.csv
export-psobjectsToCSV -arrayOfPsobjectAuditEntries $re