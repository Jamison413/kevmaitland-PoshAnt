$lastRunDateStamp = Get-Date $(Get-Content D:\jpegArchive.log)
$thisRunDateStamp = Get-Date

#$lastRunDateStamp = $thisRunDateStamp.AddMonths(-9)


gci D:\ -Recurse -Attributes !ReparsePoint -Directory | %{
    $thisFolder = $_
    gci -Path $_.FullName -Filter *.jpg | ?{$_.LastWriteTime -gt $lastRunDateStamp} | %{
        $thisFile = $_
        Write-Host $thisFile.FullName -ForegroundColor Yellow
        $result = $(& C:\Users\administrator.SUSTAINLTD\Desktop\jpeg-archive\jpeg-recompress.exe $thisFile.FullName $thisFile.FullName)
        }
    }


Set-Content -Path D:\jpegArchive.log -Value $(Get-Date $thisRunDateStamp -Format "yyyy-MM-dd hh:mm:ss") -Force