#Script to read in a list of folders to be archived, move the folders to the archive area, then create a SymLink
#to the archived folder
#
#Kev Maitland 19/02/2016
#
#SymbolicLinks and TakeOwn require Administrator privileges, so this needs to be Run As [Sustainltd\Administrator] or [SUSTAINLTD\FileSystem Manager]
#(Get-Item -Path "\\sustainltd.local\data\Internal\Residential\ECO\Compliance\Team folders\Fuchsia\Archive").GetDirectories() | ? {$_.Attributes.ToString() -notmatch "ReparsePoint"}

function Take-Ownership {
	param([String]$pFolder)
	#C:\takeown.exe /A /F $pFolder #Bodged to find real file :'(
    takeown.exe /A /F $pFolder 
	$CurrentACL = Get-Acl $pFolder
	write-host ...Adding NT Authority\SYSTEM to $pFolder -Fore DarkCyan
	$SystemACLPermission = "NT AUTHORITY\SYSTEM","FullControl","ContainerInherit,ObjectInherit","None","Allow"
	$SystemAccessRule = new-object System.Security.AccessControl.FileSystemAccessRule $SystemACLPermission
	$CurrentACL.AddAccessRule($SystemAccessRule)
	write-host ...Adding SUSTAINLTD\$($env:USERNAME) to $pFolder -Fore DarkCyan
	$AdminACLPermission = "SUSTAINLTD\$($env:USERNAME)","FullControl","ContainerInherit,ObjectInherit","None","Allow"
	$SystemAccessRule = new-object System.Security.AccessControl.FileSystemAccessRule $AdminACLPermission
	$CurrentACL.AddAccessRule($SystemAccessRule)
	Set-Acl -Path $pFolder -AclObject $CurrentACL
    }
function makeNewEmlText ([String]$pFolderPath){
    $emlFile = "To: AutomaticArchiveRetrieval@sustain.co.uk
CC: ICT Network Notifications DG <ICTNetworkNotificationsDG@sustain.co.uk>
Subject: Req: Unarchive $pFolderPath
Thread-Topic: Unarchive $pFolderPath
Thread-Index: AQHRheQ8w+cQTxhgfkG0XcFs0FZHDA==
Content-Language: en-US
X-MS-Has-Attach:
X-MS-TNEF-Correlator:
X-Unsent: 1
Content-Type: multipart/alternative;
	boundary=`"_000_7215a59208be4682b7fcfd571127332bEX02Sustainltdlocal_`"
MIME-Version: 1.0

--_000_7215a59208be4682b7fcfd571127332bEX02Sustainltdlocal_
Content-Type: text/plain; charset=`"us-ascii`"



--_000_7215a59208be4682b7fcfd571127332bEX02Sustainltdlocal_
Content-Type: text/html; charset=`us-ascii`"

<html>
<head>
<meta http-equiv=`"Content-Type`" content=`"text/html; charset=us-ascii`">
</head>
<body>

</body>
</html>

--_000_7215a59208be4682b7fcfd571127332bEX02Sustainltdlocal_--
"
    $emlFile
    }


if (!(Test-Path X:\)){New-PSDrive -Name "X" -PSProvider FileSystem -Root "\\Sustainltd.local\Data"}
$foldersToArchive = Get-Content "X:\Internal\ICT\Secure\DataStorage\X\160923_FoldersToArchive.txt"
$archiveDestinationStub = "\\Sustainltd.local\ArchivedData\Archives\X\"

foreach($folder in $foldersToArchive){
    $sourceFolder = $folder.Replace("`t","\").Replace("D:\","\\Sustainltd.local\Data\")
    $destinationFolder = $archiveDestinationStub + ($sourceFolder -Split 'Data\\')[1]
    if ($(Get-Item $sourceFolder).Attributes.ToString() -match "ReparsePoint"){Write-Host -ForegroundColor Yellow "Already archived:`t"$folder}
        else {
            Write-Host -ForegroundColor Yellow "Archiving:`t"$sourceFolder
            Take-Ownership $sourceFolder #Take ownership of the source Folder first
            foreach ($subFolder in Get-ChildItem $sourceFolder -Recurse -Directory){Take-Ownership $subFolder.FullName} #Then take Ownership of all subfolders
            
            $roboCopyOptions = @("/E","/B","/COPY:DAT","/MOVE","/DCOPY:DAT","/NP")
            $robocopyArgs = @($sourceFolder, $destinationFolder,$roboCopyOptions)
            RoboCopy @robocopyArgs

            New-Item -ItemType SymbolicLink -Path $($sourceFolder.Substring(0,$($sourceFolder.Length-$($sourceFolder.Split("\")[$sourceFolder.Split("\").Count-1]).Length -1))) -Name $($sourceFolder.Split("\")[$sourceFolder.Split("\").Count-1]) -Value $destinationFolder
            Add-Content -Value $(makeNewEmlText -pFolderPath $destinationFolder) -Path "$destinationFolder\Request Unarchive.eml"
            }
    }

#$folder = $foldersToArchive[3]