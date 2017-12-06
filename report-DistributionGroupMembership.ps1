Import-Module _PS_Library_MSOL.psm1
connect-ToExo
$outputFile = "$env:USERPROFILE\Desktop\DGs_after.csv"

$groups = Get-DistributionGroup
$groups | % {
    $memberDatum = [psobject]::new()
    $memberDatum | Add-Member -Name Group -Value $_ -MemberType NoteProperty
    $memberDatum | Add-Member -Name Members -Value $(Get-DistributionGroupMember -Identity $_.Identity) -MemberType NoteProperty
    [array]$memberData += $memberDatum
    }
$uGroups = Get-UnifiedGroup
$uGroups | % {
    $memberDatum = [psobject]::new()
    $memberDatum | Add-Member -Name Group -Value $_ -MemberType NoteProperty
    $memberDatum | Add-Member -Name Members -Value $(Get-UnifiedGroupLinks -Identity $_.Identity -LinkType Members) -MemberType NoteProperty
    [array]$memberData += $memberDatum
    }

$memberData = $memberData | Sort-Object {$_.Group.DisplayName}
rv thisLine
rv thatLine
rv whatLine
$memberData | % {
    if ($biggestGroup -lt $_.Members.Count){$biggestGroup = $_.Members.Count}
    $thisLine += $_.Group.DisplayName+","
    $thatLine += "`""+$_.Group.ManagedBy+"`","
    if ($_.Group.AccessType -eq $null){$whatLine += "`""+$_.Group.GroupType+"`","}
        else{$whatLine += "`"o365 - "+$_.Group.AccessType+"`","}
    }
Add-Content -Value $thisLine -Path $outputFile
Add-Content -Value $thatLine -Path $outputFile
Add-Content -Value $whatLine -Path $outputFile

$thisLine = ""
$memberData | % {
    if ($biggestGroup -lt $_.Members.Count){$biggestGroup = $_.Members.Count}
    $thisLine += $_.Group.DisplayName+","
    }
Add-Content -Value $thisLine -Path $outputFile

$thisLine = ""
$memberData | % {
    if($_.Members.Count -gt 1 ){$thisLine += $_.Members[0].Name+","}
    else{$thisLine += $_.Members.Name+","}
    }
Add-Content -Value $thisLine -Path $outputFile

for ($i=1;$i -lt $biggestGroup; $i++){
    $thisLine = ""
    $memberData | % {
        if ($i -lt $_.Members.Count){$thisLine += $_.Members[$i].Name}
        $thisLine += ","
        }
    Add-Content -Value $thisLine -Path $outputFile
    }


