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
    if($securityGroup){
        #Get the members and owners for the 365 Group from AAD
        $365GroupMembers = @()
        $365GroupOwners = @()
        Get-AzureADGroupMember -All:$true -ObjectId  $365Group.Id | %{[array]$365GroupMembers += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $_.UserPrincipalName;"objectId"=$_.ObjectId})}
        Get-AzureADGroupOwner -All:$true -ObjectId $365Group.Id | %{[array]$365GroupOwners += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $_.UserPrincipalName;"objectId"=$_.ObjectId})}
        #Get the members of the Security Group (this currently has to be done via Exchange for mail-enabled security groups)
        $secGroupMembers = @()
        $secGroupOwners = @()
        Get-DistributionGroupMember -Identity $securityGroup.Id | %{[array]$secGroupMembers += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $_.WindowsLiveId;"objectId"=$_.Guid})}
        Get-DistributionGroupMember -Identity $securityGroup.Id | %{[array]$secGroupOwners += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $_.WindowsLiveId;"objectId"=$_.Guid})}

        #Update the Security Group membership based on the 365 Group membership
        $membersDelta = Compare-Object -ReferenceObject $365GroupMembers -DifferenceObject $secGroupMembers -Property userPrincipalName -PassThru 
        $membersDelta | ?{$_.SideIndicator -eq "<="} | %{ #Add extra members in the 365 Group
            Add-DistributionGroupMember -Identity $securityGroup.Id -Member $_.objectId
            }
        $membersDelta | ?{$_.SideIndicator -eq "=>"} | %{ #Remove "removed" members in the 365 Group
            Remove-DistributionGroupMember -Identity $securityGroup.Id -Member $_.userPrincipalName -Confirm:$false
            }

        #Update the 365 Group ownership based on the Security Group ownership
        }
    else{
        #Create a Mail-enabled Security Group and populate it based on 365 Group Owners/Memebers
        }

    }

