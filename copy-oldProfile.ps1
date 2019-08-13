$oldProfile = "Default"

$ownOldProfile = "takeown -f C:\Users\$oldProfile -r -d y"
Invoke-Expression $ownOldProfile

Get-ChildItem C:\Users\$oldProfile -Recurse | % {
    $thisThing = $_
    $CurrentACL = Get-Acl $thisThing.FullName
	#write-host ...Adding [$("REM\$env:USERNAME")] to [$($thisThing.FullName)] -Fore Yellow
	$SystemACLPermission = "$env:USERDOMAIN\$env:USERNAME","FullControl","None","None","Allow"
	$SystemAccessRule = new-object System.Security.AccessControl.FileSystemAccessRule $SystemACLPermission
	$CurrentACL.AddAccessRule($SystemAccessRule)
	Set-Acl -Path $thisThing.FullName -AclObject $CurrentACL 
    if(!$(Test-Path $thisThing.FullName.Replace("C:\Users\$oldProfile","$env:USERPROFILE"))){
        if($thisThing.GetType().Name -eq "FileInfo"){
            Copy-Item $thisThing.FullName -Destination $thisThing.FullName.Replace("C:\Users\$oldProfile","$env:USERPROFILE") -Confirm:$false #-Verbose
            }
        else{
            $thisDir = $_
            Copy-Item $thisDir.FullName -Destination $thisDir.FullName.Replace("C:\Users\$oldProfile","$env:USERPROFILE").TrimEnd((Split-Path $thisDir -leaf)) -Confirm:$false #-Verbose
            }
        }
    }
