#param([string]$rootFolderPath,[string]$fileglob);
function Take-Ownership {
	param(
		[String]$Folder
	)
	takeown.exe /A /F $Folder
	$CurrentACL = Get-Acl $Folder
	write-host ...Adding NT Authority\SYSTEM to $Folder -Fore Yellow
	$SystemACLPermission = "NT AUTHORITY\SYSTEM","FullControl","ContainerInherit,ObjectInherit","None","Allow"
	$SystemAccessRule = new-object System.Security.AccessControl.FileSystemAccessRule $SystemACLPermission
	$CurrentACL.AddAccessRule($SystemAccessRule)
	write-host ...Adding AdminGroup to $Folder -Fore Yellow
	$AdminACLPermission = "SUSTAINLTD\ICT Team","FullControl","ContainerInherit,ObjectInherit","None","Allow"
	$SystemAccessRule = new-object System.Security.AccessControl.FileSystemAccessRule $AdminACLPermission
	$CurrentACL.AddAccessRule($SystemAccessRule)
	Set-Acl -Path $Folder -AclObject $CurrentACL
}

function Test-Folder($FolderToTest){
	$error.Clear()
	Get-ChildItem $FolderToTest -Recurse -ErrorAction SilentlyContinue | Select FullName
	if ($error) {
		foreach ($err in $error) {
			if($err.FullyQualifiedErrorId -eq "DirUnauthorizedAccessError,Microsoft.PowerShell.Commands.GetChildItemCommand") {
				Write-Host Unable to access $err.TargetObject -Fore Red
				Write-Host Attempting to take ownership of $err.TargetObject -Fore Yellow
				Take-Ownership($err.TargetObject)
				Test-Folder($err.TargetObject)
			}
		}
	}
}

Clear-Host
$rootFolderPath = "F:"
#$fileglob = "*.pst"
$saveToPath = "\\hv06\Exsustainus\Mail\Archived Mail"

Start-Transcript $Log
Take-OwnerShip ("$rootFolderPath\Users\")
Test-Folder("$rootFolderPath\Users\")
Stop-Transcript


$dirOutput = cmd.exe /c "dir $rootFolderPath\*.pst /s"
foreach ($line in $dirOutput){
    #$line.Split(" ")
    if ($line.split(" ")[1] -eq "Directory"){
        $copyFromPath = ""
        $copyFromPath += $line.Split(" ")[3]
        $i = 4
        do {
            $copyFromPath += " " + $line.Split(" ")[$i] 
            $i++
            }
        while ($i -lt $line.split(" ").Count)
        $copyFromPath += "\"
        }
    #if ([datetime]::ParseExact($line.Split(" ")[0],"dd/MM/yyyy", $null) -gt (Get-Date).addYears(-10)){}
    if (($line.Split(" ")[0].split("/").count -eq 3) -and ($line.Split(" ")[$line.Split(" ").count -1] -match ".pst")){ #If the line starts with a date and ends with .pst
        $filename = ""
        $i = $line.Split(" ").count -1
        do { #Start at the end and work backwards adding chunks
            $filename = $line.Split(" ")[$i] + " " +$filename
            $i--
            }
        until ($line.Split(" ")[$i].split(",")[0] -match "^[0-9]*$") #until we get to something that only conatins numbers

        Write-Host -ForegroundColor DarkCyan $copyFromPath$filename 
    
        $currentFolderStructure = $copyFromPath.Split("\") #Start trying to find the username from the current path
        if (($currentFolderStructure[2] -like "User*") -or ($currentFolderStructure[2] -like "Docume*")) {$userCalled = $currentFolderStructure[3]} #As we're searching recursively, we need to cater for old Windows installations
        else {$userCalled = $currentFolderStructure[2]} #Otherwise, just pull the username from the string
        
        if (-not (Test-Path $saveToPath\$userCalled)) {New-Item $("$saveToPath\$userCalled") -ItemType Directory} #Make a folder for the new user if necessary

        $toFileName = $filename
        $i = 1
        while (Test-Path $saveToPath\$userCalled\$toFilename) #Make sure we're not overwriting an existing file by iteratively changing the filename until we find one that's not is use
            {
            $toFilename = $toFilename.Substring(0,$toFilename.Length-4)+$i+'.pst'
            $i++
            }

        Copy-Item -path ($copyFromPath + $filename) -Destination ($saveToPath + "\" + $userCalled + "\" + $toFileName) 
        }
    }
