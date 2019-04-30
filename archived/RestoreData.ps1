#################################################################
#
# Script to automate the restoration of Archived folders
# 
# Kev Maitland 29/03/2016
#
# SymbolicLinks require Administrator privileges, so this needs to be Run As [Sustainltd\Administrator] or [SUSTAINLTD\FileSystem Manager]
#
# Example parameter:
# $pFolderPathToRestore = "\\sustainltd.local\archiveddata\Archives\X\Clients\Sustain Limited\101406.003-SmartHeat campaign management"
#
#################################################################

#param($pFolderPathToRestore)
$archiveRestoreDestinationStub = "\\Sustainltd.local\Data\"

if (($pFolderPathToRestore.Split("\")[6] -eq "Clients") `
    -or ($pFolderPathToRestore.Split("\")[6] -eq "Suppliers")){ #Other folders should not be restored automatically
    $sourceFolder = $pFolderPathToRestore
    $destinationFolder = $archiveRestoreDestinationStub + ($pFolderPathToRestore -Split 'X\\')[1]
    
    if ($(Get-Item $destinationFolder).Attributes.ToString() -match "ReparsePoint"){cmd /c rmdir $destinationFolder} #If there is a symbolic link to the archive location, remove that first
    
    $roboCopyOptions = @("/E","/B","/COPY:DAT","/MOVE","/DCOPY:DAT","/NP")
    $robocopyArgs = @($sourceFolder, $destinationFolder, $roboCopyOptions)
    RoboCopy @robocopyArgs #Robocaopy the data back to the X:\ drive
    
    if ($pFolderPathToRestore.Split("\")[6] -eq "Clients"){ #We need to set some extra permissions on a Clients folder too
        foreach ($subfolder in $(Get-ChildItem $destinationFolder -Directory))  {
            $CurrentACL = Get-Acl $subfolder.FullName
	        write-host ...Adding SUSTAINLTD\X-Clients-ClientData to $subfolder -Fore Yellow
	        $SystemACLPermission = "SUSTAINLTD\X-Clients-ClientData","FullControl","ContainerInherit,ObjectInherit","None","Allow"
	        $SystemAccessRule = new-object System.Security.AccessControl.FileSystemAccessRule $SystemACLPermission
	        $CurrentACL.AddAccessRule($SystemAccessRule)
    	    Set-Acl -Path $subfolder.FullName -AclObject $CurrentACL
            }
        }
    }

