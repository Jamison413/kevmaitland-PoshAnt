$allGroups = Get-AzureADMSGroup -All:$true

$allGroups.Count
$allGroups | % {
    $thisGroup = $_
    $groupStub = New-Object psobject -Property $([ordered]@{"Name"=$thisGroup.DisplayName;"Type"=$null;"Owners"=@();"Members"=@();"ObjectId"=$thisGroup.Id})
    if($thisGroup.MailEnabled -eq $true -and $thisGroup.SecurityEnabled -eq $false -and $thisGroup.GroupTypes -notcontains "Unified"){$groupStub.Type = "Distribution"}
    elseif($thisGroup.MailEnabled -eq $true -and $thisGroup.SecurityEnabled -eq $true -and $thisGroup.GroupTypes -notcontains "Unified"){$groupStub.Type = "Mail-Enabled Security"}
    elseif($thisGroup.MailEnabled -eq $false -and $thisGroup.SecurityEnabled -eq $true -and $thisGroup.GroupTypes -notcontains "Unified"){$groupStub.Type = "Security Only"}
    elseif($thisGroup.GroupTypes -contains "Unified"){$groupStub.Type = "Unified"}
    else{$groupStub.Type = "Unknown"}
    if(@("Unified","Security Only") -contains $groupStub.Type){Get-AzureADGroupOwner -All:$true -ObjectId $thisGroup.Id | %{$groupStub.Owners += $_.UserPrincipalName}}
    else{$groupstub.Owners = $(Get-DistributionGroup -Identity $thisGroup.Id).ManagedBy}
    Get-AzureADGroupMember -All:$true -ObjectId $thisGroup.Id | %{$groupStub.Members += $_.UserPrincipalName}

    switch ($groupStub.Type){
        "Distribution" {}
        "Mail-Enabled Security"{}
        "Security Only" {}
        default {}
        }

    [array]$allGroupStubs += $groupStub
    }

    $allGroupStubs | ?{$_.Type -eq "Unknown"}

    $thisGroup = $allGroups |? {$_.DisplayName -eq "View all Sustain Calendars"}

    Get-MsolGroup -ObjectId $thisGroup.Id | fl


$mappings = Import-Csv C:\Users\kevinm\Desktop\GroupMembership_mapping.csv


$mappings | % {
    if([string]::IsNullOrWhiteSpace($_.NewGroupName)){$updatedGroupName = "∂_"+$_.GroupName}
    else{$updatedGroupName = $_.NewGroupName}
    $g = Get-AzureADMSGroup -Id $_.Id
    write-host -ForegroundColor Yellow "Changing $($g.DisplayName) to $updatedGroupName"
    if (@("Unified") -contains $_.GroupType){
        if($_.HideFromGAL -eq  "TRUE"){Set-UnifiedGroup -Identity $_.Id -DisplayName $updatedGroupName -HiddenFromAddressListsEnabled $true}
        elseif($_.HideFromGAL -eq  "FALSE"){Set-UnifiedGroup -Identity $_.Id -DisplayName $updatedGroupName -HiddenFromAddressListsEnabled $false}
        else{write-host -ForegroundColor Magenta "Well, that's weird. HideFromGAL ($($_.HideFromGAL)) is neither TRUE nor FALSE"}
        }
    elseif(@("Distribution","Mail-Enabled Security") -contains $_.GroupType){
        if($_.HideFromGAL -eq  "TRUE"){Set-DistributionGroup -Identity $_.Id -DisplayName $updatedGroupName -HiddenFromAddressListsEnabled $true}
        elseif($_.HideFromGAL -eq  "FALSE"){Set-DistributionGroup -Identity $_.Id -DisplayName $updatedGroupName -HiddenFromAddressListsEnabled $false}
        else{write-host -ForegroundColor Magenta "Well, that's weird. HideFromGAL ($($_.HideFromGAL)) is neither TRUE nor FALSE"}
        }
    }