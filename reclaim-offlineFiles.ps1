$ownCsc = "takeown -f c:\windows\csc -r -d y"
Invoke-Expression $ownCsc

Get-ChildItem -Recurse "C:\Windows\CSC\v2.0.6\namespace\lrsvr02\userfolderredirections`$\$env:USERNAME\My Documents" | % {
    $thisThing = $_
    $CurrentACL = Get-Acl $thisThing.FullName
	#write-host ...Adding [$("REM\$env:USERNAME")] to [$($thisThing.FullName)] -Fore Yellow
	$SystemACLPermission = "REM\$env:USERNAME","FullControl","None","None","Allow"
	$SystemAccessRule = new-object System.Security.AccessControl.FileSystemAccessRule $SystemACLPermission
	$CurrentACL.AddAccessRule($SystemAccessRule)
	Set-Acl -Path $thisThing.FullName -AclObject $CurrentACL 
    if($thisThing.GetType().Name -eq "FileInfo"){
        Copy-Item $thisThing.FullName -Destination $thisThing.FullName.Replace("C:\Windows\CSC\v2.0.6\namespace\lrsvr02\userfolderredirections$\$($env:USERNAME)\My Documents","$env:USERPROFILE\Documents") -Force #-Verbose
        }
    else{
        $thisDir = $_
        Copy-Item $thisDir.FullName -Destination $thisDir.FullName.Replace("C:\Windows\CSC\v2.0.6\namespace\lrsvr02\userfolderredirections$\$($env:USERNAME)\My Documents","$env:USERPROFILE\Documents").TrimEnd((Split-Path $thisDir -leaf)) -Force #-Verbose
        }
    }
Get-ChildItem -Recurse "C:\Windows\CSC\v2.0.6\namespace\lrsvr02\userfolderredirections`$\$env:USERNAME\Desktop" | % {
    $thisThing = $_
    $CurrentACL = Get-Acl $thisThing.FullName
	#write-host ...Adding [$("REM\$env:USERNAME")] to [$($thisThing.FullName)] -Fore Yellow
	$SystemACLPermission = "REM\$env:USERNAME","FullControl","None","None","Allow"
	$SystemAccessRule = new-object System.Security.AccessControl.FileSystemAccessRule $SystemACLPermission
	$CurrentACL.AddAccessRule($SystemAccessRule)
	Set-Acl -Path $thisThing.FullName -AclObject $CurrentACL 
    if($thisThing.GetType().Name -eq "FileInfo"){
        Copy-Item $thisThing.FullName -Destination $thisThing.FullName.Replace("C:\Windows\CSC\v2.0.6\namespace\lrsvr02\userfolderredirections$\$($env:USERNAME)\Desktop","$env:USERPROFILE\Desktop") -Force #-Verbose
        }
    else{
        $thisDir = $_
        Copy-Item $thisDir.FullName -Destination $thisDir.FullName.Replace("C:\Windows\CSC\v2.0.6\namespace\lrsvr02\userfolderredirections$\$($env:USERNAME)\Desktop","$env:USERPROFILE\Desktop").TrimEnd((Split-Path $thisDir -leaf)) -Force #-Verbose
        }
    }
