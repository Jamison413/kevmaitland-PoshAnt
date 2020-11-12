#Create a regular expression to match the field widths and capture the data.
$regex = [regex]'(.{10})(.{11})(.{9})(.{17})(.{9})(.{2})(.{17})(.{19})(.{5})(.{3})'

#create a filter to insert a pipe character between the captured groups.
filter PipeDelimit {$_ -replace $regex, '$1|$2|$3|$4|$5|$6|$7|$8|$9|$10'}


#Pipe the records thorough the filter in batches of 1000 
Get-Content $thisFile.PSPath | Select-Object -Skip 1 | Pipedelimit

[array]$logs = @()
gci "C:\Users\KevMaitland\OneDrive - Anthesis LLC\Desktop\CDR Logs" | % {
    $thisFile = $_
    Get-Content $thisFile.PSPath | Select-Object -Skip 1 | Pipedelimit | % {
        $thisEntry = $_.Split("|").Trim()
        if($thisEntry[1] -ne "Log"){
            $logEntry = New-Object -TypeName psobject -Property $([ordered]@{
                "LogID" = $thisEntry[0]
                "Date" = $thisEntry[1]
                "Time" = $thisEntry[2]
                "AnsweredBy" = $($thisEntry[3].Replace("105","VoiceMail"))
                "Duration" = $thisEntry[4]
                "Unknown2" = $thisEntry[5]
                "DialledNumber" = $thisEntry[6]
                "CallerID" = $thisEntry[7]
                "DailyID" = $thisEntry[8]
                })
            if($logEntry.DialledNumber -match "9\+"){Add-Member -InputObject $logEntry -MemberType NoteProperty -Name Direction -Value "Outbound"}
            else{Add-Member -InputObject $logEntry -MemberType NoteProperty -Name Direction -Value "Inbound"}
            $logs += $logEntry
            }
        }

    }
$usersReceivingCalls = $logs | ? {$_.DialledNumber -ne "Main" -and $_.Direction -eq "Inbound"}
$usersMakingCalls = $logs | ? {$_.Direction -eq "Outbound"}

$usersReceivingCalls | Group-Object -Property DialledNumber | Sort-Object Count -Descending
$usersReceivingCalls | Group-Object -Property DialledNumber | Sort-Object Name
$usersMakingCalls | Group-Object -Property DialledNumber | Sort-Object Count -Descending
