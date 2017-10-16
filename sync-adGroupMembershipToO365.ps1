Import-Module _PS_Library_MSOL
Import-Module ActiveDirectory

#region functions
function new-mailEnabledDistributionGroup($displayName, $members, $memberOf, $hideFromGal, $blockExternalMail){
    #Needs Exchange Cmdlets loaded
    New-DistributionGroup -Name $displayName -Type Security -Members $members -PrimarySmtpAddress $($displayName.Replace(" ","")+"@anthesisgroup.com")
    Set-DistributionGroup $displayName -HiddenFromAddressListsEnabled $hideFromGal -RequireSenderAuthenticationEnabled $blockExternalMail
    }

function add-groupAdMembersToO365Group($groupNameToAddFrom, $groupNameToAddTo){
    $currentAdMembership = Get-ADGroupMember $groupNameToAddFrom 
    $currentO365Membership = Get-DistributionGroupMember -Identity $groupNameToAddTo
    $currentAdMembership | % {
        if ($_.distinguishedName -notmatch "OU=External Users"){#Ignore external users
            if($_.objectclass -eq "group"){add-groupAdMembersToO365Group -groupNameToAddFrom $_.Name -groupNameToAddTo $groupNameToAddTo} #Recurse for groups
            if($currentO365Membership.Name -notcontains $_.SamAccountName){Add-DistributionGroupMember -Id $groupNameToAddTo -Member $_.SamAccountName} #Add if missing
            }
        }    
    }

function remove-groupO365MembersBasedOnAdGroup($groupNameOfMasterList, $groupNameToRemoveFrom, $overrideConfirm){
    $currentAdMembership = Get-ADGroupMember $groupNameOfMasterList 
    $currentO365Membership = Get-DistributionGroupMember -Identity $groupNameToRemoveFrom
    $currentO365Membership | % {
        if ($currentAdMembership.SamAccountName -notcontains $_.Name){Remove-DistributionGroupMember -Id $groupNameToRemoveFrom -Member $_.Name -Confirm:$overrideConfirm}
        }
    }
#endregion

$creds = set-MsolCredentials
connect-ToExo -credential $creds
$hashOfGroupsToSync = @{
    "ESE Team"="Energy Systems Engineering Team";
    "EM Team"="Energy Management Team";
    "Decentralised Energy Team"="Decentralised Energy Team"
    }

foreach($securityGroup in $hashOfGroupsToSync.Keys){
    if($(try{Get-DistributionGroup $hashOfGroupsToSync[$securityGroup]}catch{$false}) -eq $false){
        new-mailEnabledDistributionGroup -displayName $hashOfGroupsToSync[$securityGroup] -members $null -memberOf "AllSustain" -hideFromGal $false -blockExternalMail $true
        }

    #Remove users first (this will remove subgroup inheritors)
    remove-groupO365MembersBasedOnAdGroup -groupNameOfMasterList $securityGroup -groupNameToRemoveFrom $hashOfGroupsToSync[$securityGroup]
    #Add Users (this will reinstate subgroup inheritors)
    add-groupAdMembersToO365Group -groupNameToAddFrom $securityGroup -groupNameToAddTo $hashOfGroupsToSync[$securityGroup]
    }

