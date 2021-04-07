function grant-ownership {
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
		    [string]$fullPath
        ,[parameter(Mandatory = $true)]
            [ValidateSet("File","Folder")] 
		    [string]$itemType
        ,[parameter(Mandatory = $true)]
            [string[]]$securityPrincipalsToGrantOwnershipTo
        ,[parameter(Mandatory = $false)]
            [switch]$alsoGrantSystemAccountOwnership
        ,[parameter(Mandatory = $false)]
            [switch]$recursive
        ,[parameter(Mandatory = $false)]
            [switch]$seizeFilesIndividually
        )
    Write-Verbose "Seizing ownership of [$itemType] [$fullPath]"
    switch($itemType){
        "File"   {$fullControlPermissions = "FullControl","Allow"}
        "Folder" {$fullControlPermissions = "FullControl","ContainerInherit,ObjectInherit","None","Allow"}
        default  {}
        }
            
	& "$env:SystemRoot\system32\takeown.exe" /A /F $fullPath | Out-Null
    
    if($alsoGrantSystemAccountOwnership){
        $securityPrincipalsToGrantOwnership.Add("NT AUTHORITY\SYSTEM")
        }

	$currentAcl = Get-Acl $fullPath
    $securityPrincipalsToGrantOwnershipTo | % {
        $thisSecurityPrincipal = $_
        Write-Verbose "Granting FullControl to [$($thisSecurityPrincipal)] on [$itemType] [$($fullPath)]"
        $aclPermission = @($thisSecurityPrincipal)
        $fullControlPermissions | % {$aclPermission += $_}
        $aclAccessRule = new-object System.Security.AccessControl.FileSystemAccessRule $aclPermission
        $currentAcl.AddAccessRule($aclAccessRule)
        } 
    Set-Acl -Path $fullPath -AclObject $currentAcl 
    if($recursive){
        Get-ChildItem -Path $fullPath | % {
            if($_.Mode -match "d"){grant-ownership -fullPath $_.FullName -itemType Folder -securityPrincipalsToGrantOwnershipTo $securityPrincipalsToGrantOwnershipTo -recursive -seizeFilesIndividually:$seizeFilesIndividually -Verbose:$VerbosePreference}
            elseif($seizeFilesIndividually){grant-ownership -fullPath $_.FullName -itemType File -securityPrincipalsToGrantOwnershipTo $securityPrincipalsToGrantOwnershipTo -seizeFilesIndividually:$seizeFilesIndividually -Verbose:$VerbosePreference}
            }
        }
    }

$localProfiles = Get-ChildItem $(Split-Path -Path $env:USERPROFILE -Parent) | Sort-Object LastWriteTime -Descending
$newProfile = $localProfiles[0]
$oldProfile = $localProfiles[1]

grant-ownership -fullPath $oldProfile.FullName -itemType Folder -securityPrincipalsToGrantOwnershipTo $("AzureAD\$($newProfile.Name)") -recursive -Verbose -seizeFilesIndividually

Copy-Item "$($oldProfile.FullName)\AppData\Local\Google\Chrome\User Data\Default\Bookmarks" "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\" -Force
Copy-Item "$($oldProfile.FullName)\AppData\Roaming\Microsoft\Signatures" "$env:APPDATA\Microsoft\" -Recurse
if((Test-Path "$($newProfile.FullName)\OneDrive - Anthesis LLC\Desktop") -eq $true){& "$env:SystemRoot\system32\robocopy.exe" "$($oldProfile.FullName)\Desktop" "$($newProfile.FullName)\OneDrive - Anthesis LLC\Desktop" /E /XN /XO /R:0 /W:1}
else{& "$env:SystemRoot\system32\robocopy.exe" "$($oldProfile.FullName)\Desktop" "$($newProfile.FullName)\Desktop" /E /XN /XO /R:0 /W:1}
if((Test-Path "$($newProfile.FullName)\OneDrive - Anthesis LLC\Documents") -eq $true){& "$env:SystemRoot\system32\robocopy.exe" "$($oldProfile.FullName)\Documents" "$($newProfile.FullName)\OneDrive - Anthesis LLC\Documents" /E /XN /XO /R:0 /W:1}
else{& "$env:SystemRoot\system32\robocopy.exe" "$($oldProfile.FullName)\Documents" "$($newProfile.FullName)\Documents" /E /XN /XO /R:0 /W:1}
& "$env:SystemRoot\system32\robocopy.exe" "$($oldProfile.FullName)\Pictures" "$($newProfile.FullName)\OneDrive - Anthesis LLC\Pictures" /E /XN /XO /R:0 /W:1
& "$env:SystemRoot\system32\robocopy.exe" "$($oldProfile.FullName)\Downloads" "$($newProfile.FullName)\Downloads" /E /XN /XO /R:0 /W:1
& "$env:SystemRoot\system32\robocopy.exe" "$($oldProfile.FullName)\Music" "$($newProfile.FullName)\Music" /E /XN /XO /R:0 /W:1
& "$env:SystemRoot\system32\robocopy.exe" "$($oldProfile.FullName)\Videos" "$($newProfile.FullName)\Videos" /E /XN /XO /R:0 /W:1
& "$env:SystemRoot\system32\robocopy.exe" "$($oldProfile.FullName)\AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState" "$($newProfile.FullName)\AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState" /E /XN /XO /R:0 /W:1

