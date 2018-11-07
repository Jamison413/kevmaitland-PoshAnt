
$all365Groups = Get-UnifiedGroup
$all365Groups | %{
    $this365Group = $_
    Write-Host -ForegroundColor Yellow "Processing 365Group [$($this365Group.DisplayName)]`t`t`t`t`t`t[$($this365Group.ExternalDirectoryObjectId)]"
    $associatedAdmsGroups = Get-AzureADMSGroup -Filter "StartsWith(DisplayName,'$($this365Group.DisplayName.Replace(" (All)",''))')" | ? {$_.GroupTypes -notcontains "Unified"} #Cannot set this on Unified Groups [The current operation is not supported on GroupMailbox.]
    $managersAadGroup = $associatedAdmsGroups | ? {$_.DisplayName -match "- Managers"}
    $membersAadGroup = $associatedAdmsGroups | ? {$_.DisplayName -match "- 365 Mirror"}
    $overallAadGroup = $associatedAdmsGroups | ? {$_.DisplayName -notmatch "Managers" -and $_.DisplayName -notmatch "365 Mirror"}
    $this365Group | Set-UnifiedGroup -CustomAttribute1 $this365Group.ExternalDirectoryObjectId
    if($managersAadGroup.Count -eq 1){$this365Group | Set-UnifiedGroup -CustomAttribute2 $managersAadGroup.Id}
        else{[array]$duffers+=@($this365Group.DisplayName, "No Managers Group")}
    if($membersAadGroup.Count -eq 1){$this365Group | Set-UnifiedGroup -CustomAttribute3 $membersAadGroup.Id}
        else{[array]$duffers+=@($this365Group.DisplayName, "No Members Group")}
    if($overallAadGroup.Count -eq 1){$this365Group | Set-UnifiedGroup -CustomAttribute4 $overallAadGroup.Id}
        else{[array]$duffers+=@($this365Group.DisplayName, "No Overall Group")}

    $this365Group = Get-UnifiedGroup $this365Group.ExternalDirectoryObjectId
    $foundManagersGroup = Get-DistributionGroup -Filter "ExternalDirectoryObjectId -eq '$($this365Group.CustomAttribute2)'"
    $foundMembersGroup = Get-DistributionGroup -Filter "ExternalDirectoryObjectId -eq '$($this365Group.CustomAttribute3)'"
    $foundOverallGroup = Get-DistributionGroup -Filter "ExternalDirectoryObjectId -eq '$($this365Group.CustomAttribute4)'"

    Write-Host -ForegroundColor DarkYellow "365Group [$($this365Group.DisplayName)]`t`t`t`t`t`t[$($this365Group.ExternalDirectoryObjectId)]"
    Write-Host -ForegroundColor DarkYellow "`tManagers`t[$($foundManagersGroup.DisplayName)]`t`t[$($this365Group.CustomAttribute2)]"
    Write-Host -ForegroundColor DarkYellow "`tMembers`t`t[$($foundMembersGroup.DisplayName)]`t[$($this365Group.CustomAttribute3)]"
    Write-Host -ForegroundColor DarkYellow "`tOveralls`t[$($foundOverallGroup.DisplayName)]`t`t`t`t`t[$($this365Group.CustomAttribute4)]"
    }

