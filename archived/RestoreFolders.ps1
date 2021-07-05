$foldersToRestore = Get-Content X:\Internal\ICT\Secure\DataStorage\X\160222_FoldersToRestore.txt
$archiveDestination = "\\Sustainltd.local\ArchivedData\Archives\X\Clients\"
$restoreDestination = "\\Sustainltd.local\Data\Clients\"
$folder = " X:\Clients\Premier Estates Limited\101456-RE_HNMBR Compliance - Premier Estates"

foreach($folder in $foldersToRestore){
    $folder = $folder.Replace("`t","\")
    Write-Host -ForegroundColor Yellow "Restoring:`t"$folder
    if ($(Get-Item $($restoreDestination+$($folder.Split("\")[2])+"\"+$($folder.Split("\")[3]))).Attributes.ToString() -match "ReparsePoint"){cmd /c rmdir $($restoreDestination+$($folder.Split("\")[2])+"\"+$($folder.Split("\")[3]))}

    $roboCopyOptions = @("/E","/B","/COPY:DAT","/MOVE","/DCOPY:DAT","/NP")
    $robocopyArgs = @($($archiveDestination+$folder.Split("\")[2]+"\"+$folder.Split("\")[3]),$($restoreDestination+$folder.Split("\")[2]+"\"+$folder.Split("\")[3]),$roboCopyOptions)
    RoboCopy @robocopyArgs

	foreach ($subfolder in $(Get-ChildItem $($restoreDestination+$folder.Split("\")[2]+"\"+$folder.Split("\")[3]) -Directory))  {
        $CurrentACL = Get-Acl $subfolder.FullName
	    write-host ...Adding SUSTAINLTD\X-Clients-ClientData to $subfolder -Fore Yellow
	    $SystemACLPermission = "SUSTAINLTD\X-Clients-ClientData","FullControl","ContainerInherit,ObjectInherit","None","Allow"
	    $SystemAccessRule = new-object System.Security.AccessControl.FileSystemAccessRule $SystemACLPermission
	    $CurrentACL.AddAccessRule($SystemAccessRule)
    	Set-Acl -Path $subfolder.FullName -AclObject $CurrentACL
        }

    }