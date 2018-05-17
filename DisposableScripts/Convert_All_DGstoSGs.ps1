#Convert "All" DGs to SGs
Import-Module _PS_Library_Groups

$dgs = Get-DistributionGroup -Filter "DisplayName -startswith 'All'"
$toUpdate = $dgs | ? {$_.DisplayName -match "All " -and $_.GRoupType -eq "Universal"}
$toUpdate | % {
    $thisGroup = $_
    [array]$members = $(Get-DistributionGroupMember -Identity $thisGroup.Identity).Name

    $memberOf = Get-DistributionGroup -Filter "Members -eq '$($thisGroup.DistinguishedName)'" 

    $newGroup = new-mailEnabledDistributionGroup -dgDisplayName $thisGroup.DisplayName -description $("Geographical Group for $($thisGroup.DisplayName)") -members $members -memberOf $memberOf -hideFromGal $false -blockExternalMail $true -owners $thisGroup.ManagedBy
    if($newGroup){
        #Move e-mail address
        $addressToMove = $thisGroup.EmailAddresses | ? {$_ -cmatch "SMTP:"}
        if ($thisGroup.EmailAddresses.Count -lt 2){
            #add a duff e-mail address
            $thisGroup | Set-DistributionGroup -EmailAddresses $("$($thisGroup.Guid)@anthesisgroup.com")
            }
        else{$thisGroup | Set-DistributionGroup -EmailAddresses $($thisGroup.EmailAddresses | ? {$_ -cnotmatch "SMTP:"})}
        Set-DistributionGroup -Identity $newGroup.ExternalDirectoryObjectId -EmailAddresses @{add=$addressToMove.Replace("SMTP:","")}
        #Hide from GAL
        Set-DistributionGroup -Identity $thisGroup.ExternalDirectoryObjectId -HiddenFromAddressListsEnabled $true
        }
    }