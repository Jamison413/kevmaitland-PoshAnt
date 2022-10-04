#Run as Administrator (to allow SeBackup permissions)

#Map the Z drive (for checking)
$connectTestResult = Test-NetConnection -ComputerName anthesislavola.file.core.windows.net -Port 445
if ($connectTestResult.TcpTestSucceeded) {
    # Save the password so the drive will persist on reboot
    cmd.exe /C "cmdkey /add:`"anthesislavola.file.core.windows.net`" /user:`"localhost\anthesislavola`" /pass:`"LrVLFIYO7kCqK0ZHt2oZXtPejPamH48i8317qRtDvgfW+KkfUZy+oh6gC80C/Ft3Mu9odzz3RI27+AStoYx+pw==`""
    # Mount the drive
    New-PSDrive -Name Z -PSProvider FileSystem -Root "\\anthesislavola.file.core.windows.net\lavolaarchives" -Persist
} else {
    Write-Error -Message "Unable to reach the Azure storage account via port 445. Check to make sure your organization or ISP is not blocking port 445, or use Azure P2S VPN, Azure S2S VPN, or Express Route to tunnel SMB traffic over a different port."
}

#Map teh G drive (for copying)
net use g: \\lavola.com\dades

#azcopy bench "$destination$sasToken" --size-per-file 10M --file-count 2000
<#Main Result:
  Code:   FileShareOrNetwork
  Desc:   Throughput may have been limited by File Share throughput limits, or by the network 
  Reason: No other factors were identified that are limiting performance, so the bottleneck is assumed to be either the throughput of the Azure File Share OR the availabl
e network bandwidth. To test whether the File Share or the network is the bottleneck, try running a benchmark to Blob Storage over the same network. If that is much faste
r, then the bottleneck in this run was probably the File Share. Check the published Azure File Share throughput targets for more info. In this run throughput of 508 Mega 
bits/sec was obtained with 256 concurrent connections.
#>

#azcopy copy "C:\Users\administrator.SUSTAINLTD\Desktop\Test2" "https://gbrenergy.file.core.windows.net/ecodata?sv=2020-02-10&ss=f&srt=sco&sp=rwdlc&se=2021-05-31T18:25:13Z&st=2021-05-24T10:25:13Z&sip=89.197.96.6&spr=https&sig=00b7rmIcV96b%2FcSws5KYVkNqEX6X7dUgtG6PfS92O5Q%3D" --recursive=true --put-md5 --preserve-smb-info --backup --cap-mbps 30
#robocopy "C:\Users\administrator.SUSTAINLTD\Desktop\Test2" "W:\Test2" /XF *.* /E /DCOPY:T 
#azcopy copy "E:\Internal" "https://gbrenergy.file.core.windows.net/ecodata?sv=2020-02-10&ss=f&srt=sco&sp=rwdlc&se=2021-05-24T18:25:13Z&st=2021-05-24T10:25:13Z&sip=89.197.96.6&spr=https&sig=DXqHgSyqhDfx2NVgSNslc57DGpshwBf4s8jUG57FUnM%3D" --recursive=true --put-md5 --preserve-smb-info --backup --cap-mbps 30

#"Sustain" data



$sources = Get-ChildItem g:\ | Where-Object {$_.Name -ne "ZJRLSWEJ" -and $_.Name -ne "TIC" -and $_.Name -ne "Escaner"}
#$destination = "https://gbrenergy.file.core.windows.net/ecodata"
#$sasToken = "sv=2020-02-10&ss=f&srt=sco&sp=rwdlc&se=2021-06-13T17:20:16Z&st=2021-05-30T09:20:16Z&sip=89.197.96.6&spr=https&sig=90ixyZkDRN6oEX1lqVrXbi6hMHxsw8c5eSNZ0SJgQCQ%3D"
#$destination = "https://gbrsustain.file.core.windows.net/x-drive"
#$sasToken = "?sv=2020-02-10&ss=f&srt=sco&sp=rwdlc&se=2021-07-05T14:41:43Z&st=2021-06-17T06:41:43Z&sip=89.197.96.6&spr=https&sig=RIi%2F8hwfPedQPHh%2FmTp7%2BUGC5K7URhWlrFcWCdPVVHQ%3D"

$destination = "https://anthesislavola.file.core.windows.net/lavolaarchives"
$sasToken = "?sv=2021-06-08&ss=bfqt&srt=sco&sp=rwdlacupyx&se=2022-10-12T21:49:46Z&st=2022-09-14T13:49:46Z&sip=194.30.115.122&spr=https&sig=zN6OUDeqADVJ7genzcVsg3Hnx0kgPiD%2BMCldPINKOhs%3D"


$sources | % {
    $thisSource = $_.FullName
    $thisDestination = $($destination).Replace('\','/')#+"/"+$_.PSChildName)
    #$thisDestination = (Split-Path $thisSource -Parent).Replace("X:",$destination).Replace("R:\X",$destination).Replace("R:",$destination).Replace("\","/") 
    #azcopy copy "$thisSource" "$thisDestination$sasToken" --recursive=true --put-md5 --backup --overwrite ifSourceNewer --cap-mbps 250 --preserve-smb-info
    azcopy copy "$thisSource" "$thisDestination$sasToken" --recursive=true --put-md5 --preserve-smb-info --backup 
    $thisDestinationOnZ = "Z:\$($_.PSChildName)"
    robocopy "$thisSource" "$thisDestinationOnZ" /XF *.* /E /DCOPY:T /XJ
    #continue
    }
