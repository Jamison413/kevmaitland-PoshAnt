Remove-Variable i
Remove-Variable data
Remove-Variable linesInFiles

$files = Get-ChildItem -Path $env:USERPROFILE\Downloads\CNGFileAccessed
$files | ForEach-Object {
    $reader = New-Object IO.StreamReader $_.FullName
    while($reader.ReadLine() -ne $null){[int]$linesInFiles++ }
}
$reader.Close()

[array]$data = @($null) * $linesInFiles
$files | ForEach-Object {
    $thisFileCsvData = Import-Csv $_.FullName
    $thisFileCsvData | ForEach-Object {
        $data[[int]$i] = $_
        $i++ 
    }
}

[array]$externalUsers = $data | Where-Object {$_.UserId -match "#EXT#"}
[array]$externalUserData = @($null) * $externalUsers.Count
$i=0

$externalUsers | ForEach-Object {
    $auditData = $_.AuditData | ConvertFrom-Json
    $prettyData = [PSCustomObject]@{
        TimeStamp = $auditData.CreationTime
        File = $auditData.ObjectId
        Site = $auditData.SiteUrl
        RelativeUrl = $auditData.SourceRelativeUrl
        FileName = $auditData.SourceFileName
        User = $($($auditData.UserId).Replace("#ext#@climateneutralgroup.onmicrosoft.com","").Replace("#ext#@climateneutralgroup.com","").Replace("_","@"))
        UserApplication = $auditData.ApplicationDisplayName
        UserBrowser = $auditData.UserAgent
        UserIP = $auditData.ClientIP
    }
    $externalUserData[$i] = $prettyData
    $i++
}

$externalUserData | Export-Csv -Path $env:USERPROFILE\Downloads\CNGFileAccessed\PrettyOutput.csv -NoTypeInformation -Force

Remove-Variable fromToPairs
$emailData = Import-Csv "$env:USERPROFILE\Downloads\CNGFileAccessed\MTSummary_Message trace report - _2022-06-29T201716.208Z__554e106f-c67b-4fb2-a51f-ea4bf70d607b (3).csv"
$emailData | ForEach-Object {
    $thisEmail = $_
    $thisEmail.recipient_status.Replace("##Receive, Deliver","").Split(";") | ForEach-Object {
        [array]$fromToPairs += [PSCustomObject]@{
            From = $thisEmail.sender_address
            To = $_
        }
    }
}
$fromToPairs  | Export-Csv -Path $env:USERPROFILE\Downloads\CNGFileAccessed\PrettyOutputEmail.csv -NoTypeInformation -Force



