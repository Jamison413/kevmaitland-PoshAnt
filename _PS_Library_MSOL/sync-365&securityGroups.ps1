#Sync Office 365 Group membership to correspnoding security group membership

Import-Module _PS_Library_MSOL.psm1
Import-Module _PS_Library_GeneralFunctionality
$msolCredentials = set-MsolCredentials #Set these once as a PSCredential object and use that to build the CSOM SharePointOnlineCredentials object and set the creds for REST
$restCredentials = new-spoCred -username $msolCredentials.UserName -securePassword $msolCredentials.Password
$csomCredentials = new-csomCredentials -username $msolCredentials.UserName -password $msolCredentials.Password
connect-ToMsol -credential $msolCredentials
connect-toAAD -credential $msolCredentials
connect-ToExo -credential $msolCredentials

#Start with the 365 Groups and synchronise back to EXO
Get-AzureADMSGroup -All:$true | ?{$_.GroupTypes -contains "Unified"} | %{
    $365Group = $_
    #Look for the corresponding Security Group
    $securityGroup = Get-AzureADMSGroup -SearchString $365Group.DisplayName | ?{$_.MailEnabled -eq $true -and $_.SecurityEnabled -eq $true -and $_.GroupTypes -notcontains "Unified"}
    if($securityGroup.Count -eq 1){
        #Get the members for the 365 Group from AAD
        $365GroupMembers = @() #Not only do we /never/ want to add users to the wrong group, having an intantiated empty array helps with compare-object later
        $secGroupMembers = @()
        Get-AzureADGroupMember -All:$true -ObjectId  $365Group.Id | %{[array]$365GroupMembers += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $_.UserPrincipalName;"objectId"=$_.ObjectId})}
        #Get the members of the Security Group (this currently has to be done via Exchange for mail-enabled security groups)
        Get-DistributionGroupMember -Identity $securityGroup.Id | %{[array]$secGroupMembers += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $_.WindowsLiveId;"objectId"=$_.Guid})}

        #Update the Security Group membership based on the 365 Group membership
        $membersDelta = Compare-Object -ReferenceObject $365GroupMembers -DifferenceObject $secGroupMembers -Property userPrincipalName -PassThru 
        #Add extra members in the 365 Group
        $membersDelta | ?{$_.SideIndicator -eq "<="} | %{ 
            $userStub = $_
            try {
                #Add-DistributionGroupMember -Identity $securityGroup.Id -Member $userStub.objectId
                [array]$membersChanged += (New-Object psobject -Property $([ordered]@{"Change"="Added";"SecGroupName"=$securityGroup.DisplayName;"UPN"=$userStub.userPrincipalName;"Result"="Succeeded";"ErrorMessage"=$null}))
                }
            catch {
                [array]$membersChanged += (New-Object psobject -Property $([ordered]@{"Change"="Added";"SecGroupName"=$securityGroup.DisplayName;"UPN"=$userStub.userPrincipalName;"Result"="Failed";"ErrorMessage"=$_}))
                }
            }
        #Remove "removed" members in the 365 Group
        $membersDelta | ?{$_.SideIndicator -eq "=>"} | %{ 
             try {
                #Remove-DistributionGroupMember -Identity $securityGroup.Id -Member $_.userPrincipalName -Confirm:$false
                $membersRemoved.Add($securityGroup.DisplayName,@($_.userPrincipalName,"Succeeded","Woopwoop"))
                }
            catch {
                $membersRemoved.Add($securityGroup.DisplayName,@($_.userPrincipalName,"Failed",$_))
                }
            }

        #Get the members for the 365 Group from AAD
        $365GroupOwners = @()
        $secGroupOwners = @()
        Get-AzureADGroupOwner -All:$true -ObjectId $365Group.Id | %{[array]$365GroupOwners += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $_.UserPrincipalName;"objectId"=$_.ObjectId})}
        #region Getting the "owners" of a mail-enabled distribution group is more complicated
        $managedBy = $(Get-DistributionGroup -Identity $securityGroup.Id).ManagedBy 
        $managedBy | % {
            $hopefullyASingleUserMailbox = @()
            $hopefullyASingleUserMailbox += Get-Mailbox -Identity $_ #Get all the mailboxes that match each entry in the ManagedBy property
            if($hopefullyASingleUserMailbox.Count -eq 1){ #If there's only one, carry on processding
                [array]$secGroupOwners += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $hopefullyASingleUserMailbox[0].UserPrincipalName;"objectId"=$hopefullyASingleUserMailbox[0].ExternalDirectoryObjectId})
                }
            else{ #If there isn't exactly one, log an error
                if($hopefullyASingleUserMailbox.Count -lt 1){[array]$ownersChanged += (New-Object psobject -Property $([ordered]@{"Change"="FindSecGroupOwner";"365GroupName"=$365Group.DisplayName;"UPN"=$_;"Result"="Failed";"ErrorMessage"="No mailbox matches ManagedBy Alias - WTF?!?"}))}
                else{
                    [string]$multipleAliasMatches = $hopefullyASingleUserMailbox.Alias -join ","
                    [array]$ownersChanged += (New-Object psobject -Property $([ordered]@{"Change"="FindSecGroupOwner";"365GroupName"=$365Group.DisplayName;"UPN"=$managedByStub;"Result"="Failed";"ErrorMessage"="Multiple mailbox match ManagedBy Alias [$_] @($multipleAliasMatches)"}))
                    }
                }
            }
        #endregion
        #Update the 365 Group ownership based on the Security Group ownership (the opposite direction to Members)
        $ownersDelta = Compare-Object -ReferenceObject $secGroupOwners -DifferenceObject $365GroupOwners -Property userPrincipalName -PassThru 
        #Add extra members in the 365 Group
        $ownersDelta | ?{$_.SideIndicator -eq "<="} | %{
            $userStub = $_
            try {
                #Add-AzureADGroupOwner -ObjectId $365Group.Id -RefObjectId $userStub.objectId
                [array]$ownersChanged += (New-Object psobject -Property $([ordered]@{"Change"="Added";"365GroupName"=$365Group.DisplayName;"UPN"=$userStub.userPrincipalName;"Result"="Succeeded";"ErrorMessage"=$null}))
                }
            catch {
                [array]$ownersChanged += (New-Object psobject -Property $([ordered]@{"Change"="Added";"365GroupName"=$365Group.DisplayName;"UPN"=$userStub.userPrincipalName;"Result"="Failed";"ErrorMessage"=$_}))
                }
            }
        #Remove "removed" members in the 365 Group
        $ownersDelta | ?{$_.SideIndicator -eq "=>"} | %{ 
             try {
                #Remove-AzureADGroupOwner -ObjectId $365Group.Id -OwnerId $_.objectId 
                [array]$ownersChanged += (New-Object psobject -Property $([ordered]@{"Change"="Removed";"365GroupName"=$365Group.DisplayName;"UPN"=$userStub.userPrincipalName;"Result"="Succeeded";"ErrorMessage"=$null}))
                }
            catch {
                [array]$ownersChanged += (New-Object psobject -Property $([ordered]@{"Change"="Removed";"365GroupName"=$365Group.DisplayName;"UPN"=$userStub.userPrincipalName;"Result"="Failed";"ErrorMessage"=$_}))
                }
            }
        }
    else{
        if($securityGroup){[array]$multiMatched365Groups += $365Group.DisplayName}
        else{[array]$unmatched365Groups += $365Group.DisplayName}
        #Create a Mail-enabled Security Group and populate it based on 365 Group Owners/Memebers
        #Nah - let's just alert for now.
        }
    }

