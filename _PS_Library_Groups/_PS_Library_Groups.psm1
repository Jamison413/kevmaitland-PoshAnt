#Sync Office 365 Group membership to correspnoding security group membership

Import-Module _PS_Library_MSOL.psm1
Import-Module _PS_Library_GeneralFunctionality
#Import-Module *pnp*

function guess-aliasFromDisplayName($displayName){
    Write-Host -ForegroundColor Magenta "guess-aliasFromDisplayName($displayName)"
    if(![string]::IsNullOrWhiteSpace($displayName)){$guessedAlias = $displayName.replace(" ","_").Replace("(","").Replace(")","").Replace(",","")}
    if($guessedAlias.length -gt 64){$guessedAlias = $guessedAlias.SubString(0,64)} 
    Write-Debug -Message "guess-aliasFromDisplayName($displayName) = [$guessedAlias]"
    $guessedAlias
    }
function guess-shorterAliasFromDisplayName($displayName){
    Write-Host -ForegroundColor Magenta "guess-aliasFromDisplayName($displayName)"
    if(![string]::IsNullOrWhiteSpace($displayName)){$guessedAlias = $displayName.replace(" ","").Replace("(","").Replace(")","").Replace(",","").Replace("&","")}
    if($guessedAlias.length -gt 64){$guessedAlias = $guessedAlias.SubString(0,64)} 
    Write-Debug -Message "guess-shorterAliasFromDisplayName($displayName) = [$(guess-aliasFromDisplayName($displayName))]"
    $guessedAlias
    }
function enumerate-groupMemberships(){
    Write-Host -ForegroundColor Magenta "enumerate-groupMemberships()"
    Get-AzureADMSGroup -All:$true | % {
        $thisGroup = $_
        $groupStub = New-Object psobject -Property $([ordered]@{"Name"=$thisGroup.DisplayName;"Type"=$null;"Owners"=@();"Members"=@();"ObjectId"=$thisGroup.Id})
        if($thisGroup.MailEnabled -eq $true -and $thisGroup.SecurityEnabled -eq $false -and $thisGroup.GroupTypes -notcontains "Unified"){$groupStub.Type = "Distribution"}
        elseif($thisGroup.MailEnabled -eq $true -and $thisGroup.SecurityEnabled -eq $true -and $thisGroup.GroupTypes -notcontains "Unified"){$groupStub.Type = "Mail-Enabled Security"}
        elseif($thisGroup.MailEnabled -eq $false -and $thisGroup.SecurityEnabled -eq $true -and $thisGroup.GroupTypes -notcontains "Unified"){$groupStub.Type = "Security Only"}
        elseif($thisGroup.GroupTypes -contains "Unified"){$groupStub.Type = "Unified"}
        else{$groupStub.Type = "Unknown"}
        if(@("Unified","Security Only") -contains $groupStub.Type){
            Get-AzureADGroupOwner -All:$true -ObjectId $thisGroup.Id | %{
                if($_.ObjectType -eq "User"){$groupStub.Owners += $_.UserPrincipalName}
                else{$groupStub.Owners += "["+$_.DisplayName+"]"}
                }
            }
        else{$groupstub.Owners = $(Get-DistributionGroup -Identity $thisGroup.Id).ManagedBy}
        Get-AzureADGroupMember -All:$true -ObjectId $thisGroup.Id | %{
            if($_.ObjectType -eq "User"){$groupStub.Members += $_.UserPrincipalName}
            else{$groupStub.Members += "["+$_.DisplayName+"]"}
            }

        [array]$allGroupStubs += $groupStub
        }   
    $allGroupStubs
    }
function new-365Group($displayName, $description, $managers, $teamMembers, $memberOf, $hideFromGal, $blockExternalMail, $isPublic, $autoSubscribe, $additionalEmailAddress, $groupClassification, $ownersAreRealManagers){
    #Groups created look like this:
    # [Dummy Team (All)] - Mail-enabled Security Group (DisplayName)
    # [Dummy Team (All)] - Unified Group (DisplayName)
    # [Dummy_Team_All] - Mail-enabled Security Group (Alias)
    # [Dummy_Team_All_365] - Unified Group (Alias)
    # [Shared Mailbox - Dummy Team (All)] - Shared Mailbox (for bodging DG membership)
    # [Dummy Team (All) - Data Managers Subgroup] - Mail-enabled Security Group for Managers
    # [Dummy Team (All) - Members Subgroup] - Mail-enabled Security Group Mirroring Unified Group Members
    #$UnifiedGroupObject.CustomAttribute1 = Own ExternalDirectoryObjectId
    #$UnifiedGroupObject.CustomAttribute2 = Data Managers Subgroup ExternalDirectoryObjectId
    #$UnifiedGroupObject.CustomAttribute3 = Members Subgroup ExternalDirectoryObjectId
    #$UnifiedGroupObject.CustomAttribute4 = Mail-Enabled Security Group ExternalDirectoryObjectId
    #$UnifiedGroupObject.CustomAttribute5 = Shared Mailbox ExternalDirectoryObjectId

    Write-Host -ForegroundColor Magenta "new-365Group($displayName, $description, $managers, $teamMembers, $memberOf, $hideFromGal, $blockExternalMail, $isPublic, $autoSubscribe, $additionalEmailAddress, $groupClassification, $ownersAreRealManagers)"
    $shortName = $displayName.Replace(" (All)","")
    #Firstly, check whether we have already created a Unified Group for this DisplayName
    $365MailAlias = $(guess-aliasFromDisplayName "$displayName 365")

    $365Group = Get-UnifiedGroup -Filter "DisplayName -eq `'$displayName`'"
    if(!$365Group){$365Group = Get-UnifiedGroup -Filter "Alias -eq `'$365MailAlias`'"} #If we can't find it by the DisplayName, check the Alias as this is less mutable

    #If we have a UG, check whether we can find the associated groups (we certainly should be able to!)
    if($365Group){
        if(![string]::IsNullOrWhiteSpace($365Group.CustomAttribute2)){
            $managersSg = Get-DistributionGroup -Filter "ExternalDirectoryObjectId -eq `'$($365Group.CustomAttribute2)`'"
            if(!$managersSg){Write-Warning "Data Managers Group [$($365Group.CustomAttribute2)] for UG [$($365Group.DisplayName)] could not be retrieved"}
            }
        else{Write-Warning "365 Group [$($365Group.DisplayName)] found, but no CustomAttribute2 (Data Managers Subgroup) property set!"}
        if(![string]::IsNullOrWhiteSpace($365Group.CustomAttribute3)){
            $membersSg = Get-DistributionGroup -Filter "ExternalDirectoryObjectId -eq '$($365Group.CustomAttribute3)'"
            if(!$membersSg){Write-Warning "Members Group [$($365Group.CustomAttribute3)] for UG [$($365Group.DisplayName)] could not be retrieved"}
            }
        else{Write-Warning "365 Group [$($365Group.DisplayName)] found, but no CustomAttribute3 (Members Subgroup) property set!"}
        if(![string]::IsNullOrWhiteSpace($365Group.CustomAttribute4)){
            $combinedSg = Get-DistributionGroup -Filter "ExternalDirectoryObjectId -eq '$($365Group.CustomAttribute4)'"
            if(!$combinedSg){Write-Warning "Combined Group [$($365Group.CustomAttribute4)] for UG [$($365Group.DisplayName)] could not be retrieved"}
            }
        else{Write-Warning "365 Group [$($365Group.DisplayName)] found, but no CustomAttribute4 (Combined Subgroup) property set!"}
        if(![string]::IsNullOrWhiteSpace($365Group.CustomAttribute5)){
            $sharedMailbox = Get-Mailbox -Filter "ExternalDirectoryObjectId -eq '$($365Group.CustomAttribute5)'"
            if(!$sharedMailbox){Write-Warning "Shared Mailbox [$($365Group.CustomAttribute5)] for UG [$($365Group.DisplayName)] could not be retrieved"}
            }
        else{Write-Warning "365 Group [$($365Group.DisplayName)] found, but no CustomAttribute5 (Shared Mailbox) property set!"}
        Write-Information "Pre-existing 365 Group found [$($365Group.DisplayName)] with CA1=[$($365Group.CustomAttribute1)], CA2=[$($365Group.CustomAttribute2)], CA3=[$($365Group.CustomAttribute3)], CA4=[$($365Group.CustomAttribute4)], CA5=[$($365Group.CustomAttribute5)]"
        }
    else{
        $combinedSgDisplayName = $displayName
        $managersSgDisplayName = "$displayName - Data Managers Subgroup"
        $membersSgDisplayName = "$displayName - Members Subgroup"
        $sharedMailboxDisplayName = "Shared Mailbox - $displayName"

        #Check whether any of these MESG exist
        $combinedSg = Get-DistributionGroup -Filter "DisplayName -eq `'$combinedSgDisplayName`'"
        if(!$combinedSg){$combinedSg = Get-DistributionGroup -Filter "Alias -eq `'$(guess-aliasFromDisplayName $combinedSgDisplayName)`'"} #If we can't find it by the DisplayName, check the Alias as this is less mutable
        if($combinedSg.Count -gt 1){#If we get too many results (e.g. we've collided with an existing group name) try again using the Alias
            $tryAgain = Get-DistributionGroup -Filter "Alias -eq `'$(guess-aliasFromDisplayName $combinedSgDisplayName)`'"
            if($tryAgain -ne $null -and !($tryAgain.Count -gt 1)){$combinedSg = $tryAgain}
            else{
                Write-Warning "Multiple Groups matched for Combined Group [$combinedSgDisplayName]`r`n`t $($combinedSg.PrimarySmtpAddress)"
                $combinedSg = $null
                }
            } 
        $managersSg = Get-DistributionGroup -Filter "DisplayName -eq `'$managersSgDisplayName`'"
        if(!$managersSg){$managersSg = Get-DistributionGroup -Filter "Alias -eq `'$(guess-aliasFromDisplayName $managersSgDisplayName)`'"} #If we can't find it by the DisplayName, check the Alias as this is less mutable
        $membersSg = Get-DistributionGroup -Filter "DisplayName -eq `'$membersSgDisplayName`'"
        if(!$membersSg){$membersSg = Get-DistributionGroup -Filter "Alias -eq `'$(guess-aliasFromDisplayName $membersSgDisplayName)`'"} #If we can't find it by the DisplayName, check the Alias as this is less mutable
        $sharedMailbox = Get-Mailbox -Filter "DisplayName -eq `'$sharedMailboxDisplayName`'"
        if(!$sharedMailbox){$sharedMailbox = Get-DistributionGroup -Filter "Alias -eq `'$(guess-aliasFromDisplayName $sharedMailboxDisplayName)`'"} #If we can't find it by the DisplayName, check the Alias as this is less mutable

        #Create any groups that don't already exist
        if(!$combinedSg){
            Write-Host -ForegroundColor Yellow "Creating Combined Security Group [$combinedSgDisplayName]"
            try{$combinedSg = new-mailEnabledSecurityGroup -dgDisplayName $combinedSgDisplayName -members $null -memberOf $memberOf -hideFromGal $false -blockExternalMail $true -owners "ITTeamAll@anthesisgroup.com" -description "Mail-enabled Security Group for $displayName"}
            catch{$Error}
            }

        if($combinedSg){ #If we now have a Combined SG
            if(!$managersSg){ #Create a Managers SG if required
                Write-Host -ForegroundColor Yellow "Creating Data Managers Security Group [$managersSgDisplayName]"
                $managersMemberOf =@($combinedSg.ExternalDirectoryObjectId)
                if($ownersAreRealManagers){$managersMemberOf += "Managers (All)"}
                try{$managersSg = new-mailEnabledSecurityGroup -dgDisplayName $managersSgDisplayName -members $managers -memberOf $managersMemberOf -hideFromGal $false -blockExternalMail $true -owners "ITTeamAll@anthesisgroup.com" -description "Mail-enabled Security Group for $shortName Data Managers"}
                catch{$Error}
                }

            if(!$membersSg){ #And create a Members SG if required
                Write-Host -ForegroundColor Yellow "Creating Members Security Group [$membersSgDisplayName]"
                try{$membersSg = new-mailEnabledSecurityGroup -dgDisplayName $("$membersSgDisplayName") -members $teamMembers -memberOf $combinedSg.ExternalDirectoryObjectId -hideFromGal $false -blockExternalMail $true -owners "ITTeamAll@anthesisgroup.com" -description "Mail-enabled Security Group for mirroring membership of $shortName Unified Group"}
                catch{$Error}
                }

            #Check that everything's worked so far
            if(!$managersSg){Write-Error "Managers Security Group [$managersSgDisplayName] not available. Cannot proceed with UnifiedGroup creation"}
            if(!$membersSg){Write-Error "Members Security Group [$membersSgDisplayName] not available. Cannot proceed with UnifiedGroup creation"}
            if($managersSg -and $membersSg){#If we now have a Managers SG & Members SG
                #Create a UG
                try{
                    Write-Host -ForegroundColor DarkMagenta "Creating Unified 365 Group [$displayName]"
                    if($isPublic){$accessType = "Public"}else{$accessType = "Private"}
                    if([string]::IsNullOrWhiteSpace($description)){$description = "Unified 365 Group for $displayName"}
                    #Create the UG
                    Write-Host -ForegroundColor DarkMagenta "`tNew-UnifiedGroup -DisplayName $displayName -Name $365MailAlias -Alias $365MailAlias -Notes $description -AccessType $accessType -Owner $($managers[0]) -RequireSenderAuthenticationEnabled $blockExternalMail -AutoSubscribeNewMembers:$autoSubscribe -AlwaysSubscribeMembersToCalendarEvents:$autoSubscribe -Members $teamMembers   -Classification $groupClassification" 
                    $365Group = New-UnifiedGroup -DisplayName $displayName -Name $365MailAlias -Alias $365MailAlias -Notes $description -AccessType $accessType -Owner $managers[0] -RequireSenderAuthenticationEnabled $blockExternalMail -AutoSubscribeNewMembers:$autoSubscribe -AlwaysSubscribeMembersToCalendarEvents:$autoSubscribe -Members $teamMembers   -Classification $groupClassification
                    #Set the additional Properties and associations
                    Write-Host -ForegroundColor DarkMagenta "`tSet-UnifiedGroup -Identity $($365Group.ExternalDirectoryObjectId) -HiddenFromAddressListsEnabled $true -CustomAttribute1 [$($365Group.ExternalDirectoryObjectId)] -CustomAttribute2 [$($managersSg.ExternalDirectoryObjectId)] -CustomAttribute3 [$($membersSg.ExternalDirectoryObjectId)] -CustomAttribute4 [$($combinedSg.ExternalDirectoryObjectId)"] 
                    Set-UnifiedGroup -Identity $365Group.ExternalDirectoryObjectId -HiddenFromAddressListsEnabled $true -CustomAttribute1 $365Group.ExternalDirectoryObjectId -CustomAttribute2 $managersSg.ExternalDirectoryObjectId -CustomAttribute3 $membersSg.ExternalDirectoryObjectId -CustomAttribute4 $combinedSg.ExternalDirectoryObjectId
                    if($managers.Count -gt 1){Add-UnifiedGroupLinks -Identity $ug.Identity -LinkType Owner -Links $managers -Confirm:$false}
                    }
                catch{$Error}
                
                if($365Group){ #If we now have a 365 UG, create a Shared Mailbox (if required) and configure it
                    if(!$sharedMailbox){
                        Write-Host -ForegroundColor DarkMagenta  "Creating Shared Mailbox [$sharedMailboxDisplayName]"
                        try{$sharedMailbox = New-Mailbox -Shared -DisplayName $sharedMailboxDisplayName -Name $sharedMailboxDisplayName -Alias $(guess-aliasFromDisplayName ($sharedMailboxDisplayName)) -ErrorAction Continue}
                        catch{$Error}
                        }

                    if($sharedMailbox){
                        Set-Mailbox -Identity $sharedMailbox.ExternalDirectoryObjectId -HiddenFromAddressListsEnabled $true -RequireSenderAuthenticationEnabled $false -ForwardingAddress $365Group.PrimarySmtpAddress -DeliverToMailboxAndForward $true -ForwardingSmtpAddress $365Group.PrimarySmtpAddress -Confirm:$false
                        Set-user -Identity $sharedMailbox.ExternalDirectoryObjectId -Manager kevin.maitland #For want of someone better....
                        #DeliverToMailboxAndForward has to be true, otherwise it just doesn't forward :/
                        #Assign the Shared Mailbox as a member of the Security Group
                        Add-DistributionGroupMember -Identity $combinedSg.ExternalDirectoryObjectId -Member $sharedMailbox.ExternalDirectoryObjectId -BypassSecurityGroupManagerCheck
                        Set-UnifiedGroup -Identity $365Group.ExternalDirectoryObjectId -CustomAttribute5 $sharedMailbox.ExternalDirectoryObjectId
                        }
                    else{Write-Error "Shared Mailbox not available. Cannot complete UG setup."}
                    }
                else{Write-Error "Unified Group [$displayName] not available. Cannot proceed with Shared Mailbox creation."}
                }
            else{Write-Error "Managers/Members Security Group [$managersSgDisplayName]/[$membersSgDisplayName] not available. Cannot proceed with UnifiedGroup creation"}        

            }
        else{Write-Error "Combined Security Group [$combinedSgDisplayName] not available. Cannot proceed with SubGroup creation"}        
        Write-Information "New 365 Group created [$($365Group.DisplayName)] with CA1=[$($365Group.CustomAttribute1)], CA2=[$($365Group.CustomAttribute2)], CA3=[$($365Group.CustomAttribute3)], CA4=[$($365Group.CustomAttribute4)], CA5=[$($365Group.CustomAttribute5)]"
        }
    $365Group
    }
function new-mailEnabledSecurityGroup($dgDisplayName, $description, $members, $memberOf, $hideFromGal, $blockExternalMail, $owners){
    Write-Host -ForegroundColor Magenta "new-mailEnabledSecurityGroup($dgDisplayName, $description, $members, $memberOf, $hideFromGal, $blockExternalMail, $owners)"
    $mailAlias = guess-aliasFromDisplayName $dgDisplayName
    $shortMailAlias = guess-shorterAliasFromDisplayName $dgDisplayName #There are two ways now to guess the Alias because Kev is rubbish
    $mailName = $dgDisplayName
    if($mailName.length -gt 64){$mailName = $mailName.SubString(0,64)}

    #Check to see if this already exists. This is based on Alias, which is mutable :(    
    $mesg = Get-DistributionGroup -Filter "Alias -eq `'$mailAlias`'"
    if(!$mesg){$mesg = Get-DistributionGroup -Filter "Alias -eq `'$shortMailAlias`'"}
    if($mesg){ #If the group already exists, add the new Members (ignore any removals)
        $members  | % {
            Write-Host -ForegroundColor DarkMagenta "Adding TeamMembers Add-DistributionGroupMember $mailAlias -Member $_ -Confirm:$false -BypassSecurityGroupManagerCheck"
            Add-DistributionGroupMember $mailAlias -Member $_ -Confirm:$false -BypassSecurityGroupManagerCheck
            }
        }
    else{ #If the group doesn't exist, try creating it
        try{
            Write-Host -ForegroundColor DarkMagenta "New-DistributionGroup -Name $mailName -DisplayName $dgDisplayName -Type Security -Members $members -PrimarySmtpAddress $($mailAlias+"@anthesisgroup.com") -Notes $description -Alias $mailAlias"
            $mesg = New-DistributionGroup -Name $mailName -DisplayName $dgDisplayName -Type Security -Members $members -PrimarySmtpAddress $($mailAlias+"@anthesisgroup.com") -Notes $description -Alias $mailAlias #| Out-Null
            }
        catch{$Error}
        }

    if(!$mesg){Write-Error "Mail-Enabled Security Group [$dgDisplayName] neither found, nor created :/"}
    else{ #Now set the additional properties and MemberOf
        Write-Host -ForegroundColor DarkMagenta "Set-DistributionGroup -Identity $mailAlias -HiddenFromAddressListsEnabled $hideFromGal -RequireSenderAuthenticationEnabled $blockExternalMail -ManagedBy $owners"
        Set-DistributionGroup -Identity $mesg.ExternalDirectoryObjectId -HiddenFromAddressListsEnabled $hideFromGal -RequireSenderAuthenticationEnabled $blockExternalMail -ManagedBy $owners
        $memberOf | % {
            if(![string]::IsNullOrWhiteSpace($_)){
                Write-Host -ForegroundColor DarkYellow "Adding As MemberOf Add-DistributionGroupMember [$_] -Member [$mailAlias] -Confirm:$false -BypassSecurityGroupManagerCheck"
                Add-DistributionGroupMember $_ -Member $mailAlias -Confirm:$false -BypassSecurityGroupManagerCheck
                }
            }
        }
    $mesg
    }
function new-externalGroup(){}
function new-symGroup($displayName, $description, $managers, $teamMembers, $memberOf, $additionalEmailAddress){
    Write-Host -ForegroundColor Magenta "new-symGroup($displayName, $description, $managers, $teamMembers, $memberOf, $additionalEmailAddress)"
    $hideFromGal = $false
    $blockExternalMail = $true
    $isPublic = $true 
    $autoSubscribe = $true
    $groupClassification = "Internal"
    new-365Group -displayName $displayName -description $description -managers $managers -teamMembers $teamMembers -memberOf $memberOf -hideFromGal $hideFromGal -blockExternalMail $blockExternalMail -isPublic $isPublic -autoSubscribe $autoSubscribe -additionalEmailAddress $additionalEmailAddress -groupClassification $groupClassification -ownersAreRealManagers $false
    }
function new-teamGroup($displayName, $description, $managers, $teamMembers, $memberOf, $additionalEmailAddress){
    Write-Host -ForegroundColor Magenta "new-teamGroup($displayName, $description, $managers, $teamMembers, $memberOf, $additionalEmailAddress)"
    $hideFromGal = $false
    $blockExternalMail = $true
    $isPublic = $false 
    $autoSubscribe = $true
    $groupClassification = "Internal"
    new-365Group -displayName $displayName -description $description -managers $managers -teamMembers $teamMembers -memberOf $memberOf -hideFromGal $hideFromGal -blockExternalMail $blockExternalMail -isPublic $isPublic -autoSubscribe $autoSubscribe -additionalEmailAddress $additionalEmailAddress -groupClassification $groupClassification -ownersAreRealManagers $true} 
function report-groupMembershipEnumeration($allGroupStubs,$filePathAndName){
    Write-Host -ForegroundColor Magenta "report-groupMembershipEnumeration($allGroupStubs,$filePathAndName)"
    $allGroupStubs | % {
        [array]$formattedGroupStubs += New-Object psobject -Property $([ordered]@{"GroupName"=$_.Name;"GroupType"=$_.Type;"Owners"=$($_.Owners -join "`r`n");"Members"=$($_.Members -join "`r`n");"Id"=$_.ObjectId})
        }
    $formattedGroupStubs | Sort-Object GroupName | Export-Csv -Path $filePathAndName -Encoding UTF8 -NoTypeInformation -Append
    }
function report-groupMembershipSync([array]$groupChangesArray,[boolean]$changesAreToGroupOwners,[boolean]$actionedGroupIs365,$emailAddressForOverviewReport){
    Write-Host -ForegroundColor Magenta "report-groupMembershipSync($($groupChangesArray.Count) Users changed,[boolean]$changesAreToGroupOwners,[boolean]$actionedGroupIs365,$emailAddressForOverviewReport"
    #$groupChangesArray = $ownersChanged
    if($actionedGroupIs365){$groupChangesArray = $groupChangesArray | Sort-Object ActionedGroupName,Result,Change,DisplayName}
    else{$groupChangesArray = $groupChangesArray | Sort-Object SourceGroupName,Result,Change,DisplayName}
    $groupChangesArray | %{
        $thisChange = $_
        if($current365Group.Mail -ne $thisChange.SourceGroupName -and $current365Group.Mail -ne $thisChange.ActionedGroupName){
            #We need to start another report, so send the current one before we start again
            if($ownerReport){
                Write-Host $ownerReport
                send-membershipEmailReport -ownerReport $ownerReport -changesAreToGroupOwners $changesAreToGroupOwners -emailAddressForOverviewReport $emailAddressForOverviewReport
                }
            #Start new ownerReport
            $ownerReport = New-Object psobject -Property $([ordered]@{"To"=@();"groupName"=$null;"added"=@();"removed"=@();"problems"=@();"fullMemberList"=@()})
            if($actionedGroupIs365){$current365Group = Get-AzureADMSGroup -Filter "Mail eq '$($thisChange.ActionedGroupName)'"}
            else{$current365Group = Get-AzureADMSGroup -Filter "Mail eq '$($thisChange.SourceGroupName)'"}
            $ownerReport.groupName = $current365Group.DisplayName
            #Get the owners' e-mail addresses
            #[array]$owners = $current365Group | % {$(Get-AzureADGroupOwner -All:$true -ObjectId $_.Id).UserPrincipalName} #This gets the 365 Group Owners
            [array]$owners = $(Get-AzureADGroupMember -ObjectId $(Get-UnifiedGroup -Identity $current365Group.Id).CustomAttribute2).UserPrincipalName #This gets the Data Managers Subgroup members
            
            if($owners){$ownerReport.To = $owners}
            else{
                $ownerReport.To = $emailAddressForOverviewReport
                $ownerReport.groupName = "***Unowned Group*** $current365GroupName"
                }
            #Get the members' (or owners' if we're reporting on group Ownership) DisplayNames
            if($changesAreToGroupOwners){
                #[array]$members = Get-AzureADMSGroup -SearchString $current365GroupName | ? {$_.GroupTypes -contains "Unified"} | % {$(Get-AzureADGroupOwner -All:$true -ObjectId $thisChange.Id).DisplayName}
                [array]$members = $current365Group | % {$(Get-AzureADGroupOwner -All:$true -ObjectId $_.Id).DisplayName}
                $members = $($members | Sort-Object)
                if($members){$ownerReport.fullMemberList = $members}
                }
            else{
                [array]$members = $current365Group | % {$(Get-AzureADGroupMember -All:$true -ObjectId $_.Id).DisplayName}
                $members = $($members | Sort-Object)
                if($members){$ownerReport.fullMemberList = $members}
                }
            }
        #Add any processed changes
        if($thisChange.Result -eq "Succeeded"){
            if($thisChange.Change -eq "Added"){$ownerReport.added += $thisChange.DisplayName}
            else{$ownerReport.Removed += $thisChange.DisplayName}
            }
        #Add any failures as problems to be investigated manually
        else{$ownerReport.problems += $thisChange.DisplayName}
        }
    #Finally, send the last reports too
    Write-Host $ownerReport
    Write-Host "To: " + $ownerReport.To
    send-membershipEmailReport -ownerReport $ownerReport -changesAreToGroupOwners $changesAreToGroupOwners  -emailAddressForOverviewReport $emailAddressForOverviewReport
    }
function send-membershipChangeReportToManagers(){
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [psobject]$UnifiedGroup
        ,[Parameter(Mandatory=$true, Position=1)]
        [ValidateSet("Members", "Owners")]
        [string]$changesAreTo
        ,[Parameter(Mandatory=$false, Position=2)]
        [array]$usersAddedArray
        ,[Parameter(Mandatory=$false, Position=3)]
        [array]$usersRemovedArray
        ,[Parameter(Mandatory=$false, Position=4)]
        [array]$usersWithProblemsArray
        ,[Parameter(Mandatory=$false, Position=5)]
        [array]$usersInGroupAfterChanges
        ,[Parameter(Mandatory=$false, Position=6)]
        [string[]]$adminEmailAddresses
        ,[Parameter(Mandatory=$false, Position=7)]
        [string[]]$ownersEmailAddresses
        )

    #The recipient of the report is only interested in changes to the 365 Group, so we'll just portray all changes as being to the 365 group reagrdless of which group was actualyl updated

    $subject = "$($UnifiedGroup.DisplayName) $($changesAreTo)hip updated"
    $body = "<HTML><FONT FACE=`"Calibri`">Hello Data Managers for <B>$($UnifiedGroup.DisplayName)</B>,`r`n`r`n<BR><BR>"
    $body += "Changes have been made to the <B><U>$($changesAreTo.TrimEnd("s"))</U>ship</B> of $($UnifiedGroup.DisplayName)`r`n`r`n<BR><BR>"
    if($usersAddedArray.Count -gt 0){
        $usersAddedArray  = $usersAddedArray | Sort-Object DisplayName
        $body += "The following users have been <B>added</B> as Team <B>$($changesAreTo)</B>:      `r`n`t<BR><PRE>&#9;$($usersAddedArray.DisplayName -join     "`r`n`t")</PRE>`r`n`r`n<BR>"
        }
    if($usersRemovedArray.Count -gt 0){
        $usersRemovedArray = $usersRemovedArray | Sort-Object DisplayName
        $body += "The following users have been <B>removed</B> from the Group <B>$($changesAreTo)</B>:  `r`n`t<BR><PRE>&#9;$($usersRemovedArray.DisplayName -join   "`r`n`t")</PRE>`r`n`r`n<BR>"
        }
    if($usersWithProblemsArray.Count -gt 0){
        $usersWithProblemsArray = $usersWithProblemsArray | Sort-Object DisplayName
        $body += "The were some problems processing changes to these users (but IT have been notified):`r`n`t<BR><PRE>&#9;$($usersWithProblemsArray.DisplayName -join "`r`n`t")</PRE>`r`n`r`n<BR>"
        }
    if($usersInGroupAfterChanges.Count -gt 0){
        $usersInGroupAfterChanges = $usersInGroupAfterChanges | Sort-Object DisplayName
        $body += "The full list of group $($changesAreTo) looks like this:`r`n`t<BR><PRE>&#9;$($usersInGroupAfterChanges.DisplayName -join "`r`n`t")</PRE>`r`n`r`n<BR>"
        }
    else{$body += "It looks like the group is now empty...`r`n`r`n<BR><BR>"}
    if($changesAreTo -eq "Owners"){$body += "To help us all remain compliant and secure, group <I>ownership</I> is still managed centrally by your IT Team, and you will need to liaise with them to make changes to group ownership.`r`n`r`n<BR><BR>"}
    $body += "As an owner, you can manage the membership of this group (and there is a <A HREF=`"https://anthesisllc.sharepoint.com/sites/Resources-IT/_layouts/15/DocIdRedir.aspx?ID=HXX7CE52TSD2-1759992947-6`">guide available to help you</A>), or you can contact the IT team for your region,`r`n`r`n<BR><BR>"
    $body += "Love,`r`n`r`n<BR><BR>The Helpful Groups Robot</FONT></HTML>"
    
    if($PSCmdlet.ShouldProcess($("$changesAreTo [$($UnifiedGroup.DisplayName)]"))){#Fudges -WhatIf as it's not suppoerted natively by Send-MailMessage
        Send-MailMessage -To $ownersEmailAddresses -From "thehelpfulgroupsrobot@anthesisgroup.com" -cc $adminEmailAddresses -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
        }
    else{
        Send-MailMessage -To "kevin.maitland@anthesisgroup.com" -From "thehelpfulgroupsrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
        }

    }
function send-membershipEmailReport($ownerReport,[boolean]$changesAreToGroupOwners,$emailAddressForOverviewReport){
    #######################################################################################################################################
    #
    #
    #   Deprecated - use send-membershipChangeReportToManagers instead
    #
    #
    #######################################################################################################################################
    Write-Warning "send-membershipEmailReport is deprecated - use send-membershipChangeReportToManagers instead"
    Write-Host -ForegroundColor Magenta "send-membershipEmailReport($ownerReport,[boolean]$changesAreToGroupOwners,$emailAddressForOverviewReport"
    #Write and send e-mail
    if($changesAreToGroupOwners){$type = "owner"}
    else{$type = "member"}
    $subject = "$($ownerReport.groupName) $($type)ship updated"
    $body = "<HTML><FONT FACE=`"Calibri`">Hello Data Managers for <B>$($ownerReport.groupName)</B>,`r`n`r`n<BR><BR>"
    #$body += $ownerReport.To+"`r`n`r`n<BR><BR>"
    $body += "Changes have been made to the <B><U>$($type)</U>ship</B> of $($ownerReport.groupName)`r`n`r`n<BR><BR>"
    if($ownerReport.added)  {$body += "The following users have been <B>added</B> as Team <B>$($type)s</B>:      `r`n`t<BR><PRE>&#9;$($ownerReport.added -join     "`r`n`t")</PRE>`r`n`r`n<BR>"}
    if($ownerReport.removed){$body += "The following users have been <B>removed</B> from the Group <B>$($type)s</B>:  `r`n`t<BR><PRE>&#9;$($ownerReport.removed -join   "`r`n`t")</PRE>`r`n`r`n<BR>"}
    if($ownerReport.problems){
        $body += "The were some problems processing changes to these users (but IT have been notified):`r`n`t<BR><PRE>&#9;$($ownerReport.problems -join "`r`n`t")</PRE>`r`n`r`n<BR>"
        $ownerReport.To += $emailAddressForOverviewReport
        }
    if($ownerReport.fullMemberList){$body += "The full list of group $($type)s looks like this:`r`n`t<BR><PRE>&#9;$($ownerReport.fullMemberList -join "`r`n`t")</PRE>`r`n`r`n<BR>"}
    else{$body += "It looks like the group is now empty...`r`n`r`n<BR><BR>"}
    if($type -eq "owner"){$body += "To help us all remain compliant and secure, group <I>ownership</I> is still managed centrally by your IT Team, and you will need to liaise with them to make changes to group ownership.`r`n`r`n<BR><BR>"}
    $body += "As an owner, you can manage the membership of this group (and there is a <A HREF=`"https://anthesisllc.sharepoint.com/sites/Resources-IT/_layouts/15/DocIdRedir.aspx?ID=HXX7CE52TSD2-1759992947-6`">guide available to help you</A>), or you can contact the IT team for your region,`r`n`r`n<BR><BR>"
    $body += "Love,`r`n`r`n<BR><BR>The Helpful Groups Robot</FONT></HTML>"
    #Send-MailMessage -To "kevin.maitland@anthesisgroup.com" -From "thehelpfulgroupsrobot@anthesisgroup.com" -cc "kevin.maitland@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
    Send-MailMessage -To $ownerReport.To -From "thehelpfulgroupsrobot@anthesisgroup.com" -cc "kevin.maitland@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
    #$body
    }
function send-membershipChangeProblemReportToAdmins(){
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [psobject]$UnifiedGroup
        ,[Parameter(Mandatory=$true, Position=1)]
        [ValidateSet("Members", "Owners")]
        [string]$changesAreTo
        ,[Parameter(Mandatory=$false, Position=2)]
        [array]$usersWithProblemsArray
        ,[Parameter(Mandatory=$false, Position=3)]
        [array]$usersIn365GroupAfterChanges
        ,[Parameter(Mandatory=$false, Position=4)]
        [array]$usersInAADGroupAfterChanges
        ,[Parameter(Mandatory=$false, Position=5)]
        [string[]]$adminEmailAddresses
        )

    $subject = "Error managing $($changesAreTo)hip for Group [$($UnifiedGroup.DisplayName)]"
    $body = "<HTML><FONT FACE=`"Calibri`">Hello 365 Group Admins,`r`n`r`n<BR><BR>"
    if($usersWithProblemsArray.Count -gt 0){
        $usersWithProblemsArray = $usersWithProblemsArray | Sort-Object Change,DisplayName
        #$body += "The were some problems processing changes to these users (but IT have been notified):`r`n`t<BR><PRE>&#9;$($usersWithProblemsArray.DisplayName -join "`r`n`t")</PRE>`r`n`r`n<BR>"
        $body += "I encountered some problems automatically managing the $($changesAreTo)hip for [$($UnifiedGroup.DisplayName)][$($UnifiedGroup.ExternalDirectoryObjectId)]:`r`n`t<BR><PRE>&#9;"
        $usersWithProblemsArray | % {$body += "$($_.Change):`t<B>$($_.DisplayName)</B>`r`n$($_.Error)`r`n`t"}
        $body += "</PRE>`r`n`r`n<BR>"
        }
    if($usersIn365GroupAfterChanges.Count -gt 0){
        $usersIn365GroupAfterChanges = $usersIn365GroupAfterChanges | Sort-Object DisplayName
        $body += "The full list of 365 group $($changesAreTo) looks like this:`r`n`t<BR><PRE>&#9;$($usersIn365GroupAfterChanges.DisplayName -join "`r`n`t")</PRE>`r`n`r`n<BR>"
        }
    else{$body += "It looks like the group is now empty...`r`n`r`n<BR><BR>"}
    if($usersInAADGroupAfterChanges.Count -gt 0){
        $usersInAADGroupAfterChanges = $usersInAADGroupAfterChanges | Sort-Object DisplayName
        $body += "The full list of AAD group $($changesAreTo) looks like this:`r`n`t<BR><PRE>&#9;$($usersInAADGroupAfterChanges.DisplayName -join "`r`n`t")</PRE>`r`n`r`n<BR>"
        }
    else{$body += "It looks like the group is now empty...`r`n`r`n<BR><BR>"}
    $body += "Love,`r`n`r`n<BR><BR>The Helpful Groups Robot</FONT></HTML>"
    
    if($PSCmdlet.ShouldProcess($("$changesAreTo [$($UnifiedGroup.DisplayName)]"))){#Fudges -WhatIf as it's not suppoerted natively by Send-MailMessage
        Send-MailMessage -To $adminEmailAddresses -From "thehelpfulgroupsrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
        }
    else{
        Send-MailMessage -To "kevin.maitland@anthesisgroup.com" -From "thehelpfulgroupsrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
        }

    }
function send-noOwnersForGroupAlertToAdmins(){
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [psobject]$UnifiedGroup
        ,[Parameter(Mandatory=$false, Position=3)]
        [array]$currentOwners
        ,[Parameter(Mandatory=$false, Position=5)]
        [string[]]$adminEmailAddresses
        )

    $subject = "Unowned 365 Group found: [$($UnifiedGroup.DisplayName)]"
    $body = "<HTML><FONT FACE=`"Calibri`">Hello 365 Group Admins,`r`n`r`n<BR><BR>"
    $body += "365 Group [$($UnifiedGroup.DisplayName)][$($UnifiedGroup.ExternalDirectoryObjectId)] has no active owners:`r`n`t<BR><PRE>&#9;"

    if($currentOwners.Count -gt 0){
        $currentOwners = $currentOwners | Sort-Object DisplayName
        $body += "The full list of 365 group Owners looks like this:`r`n`t<BR><PRE>&#9;$($usersIn365GroupAfterChanges.DisplayName -join "`r`n`t")</PRE>`r`n`r`n<BR>"
        }
    else{$body += "It looks like the Owners group is now empty...`r`n`r`n<BR><BR>"}
    $body += "Love,`r`n`r`n<BR><BR>The Helpful Groups Robot</FONT></HTML>"    

    if([string]::IsNullOrWhiteSpace($adminEmailAddresses)){$adminEmailAddresses = get-groupAdminRoleEmailAddresses}

    if($PSCmdlet.ShouldProcess($("[$($UnifiedGroup.DisplayName)]"))){#Fudges -WhatIf as it's not suppoerted natively by Send-MailMessage
        Send-MailMessage -To $adminEmailAddresses -From "thehelpfulgroupsrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 -Priority High
        }
    else{
        Send-MailMessage -To "kevin.maitland@anthesisgroup.com" -From "thehelpfulgroupsrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 -Priority High
        }
    
    }
function sync-365GroupMembersToMirroredSecurityGroup($unifiedGroupObject,[boolean]$reallyDoIt,[boolean]$dontSendEmailReport,$fullLogFile, $errorLogFile){
    #######################################################################################################################################
    #
    #
    #   Deprecated - use sync-groupMemberships instead
    #
    #
    #######################################################################################################################################
    Write-Warning "sync-365GroupMembersToMirroredSecurityGroup is deprecated - use sync-groupMemberships instead"
    Write-Host -ForegroundColor Magenta "sync-365GroupMembersToMirroredSecurityGroup($($unifiedGroupObject.DisplayName),[boolean]$reallyDoIt,[boolean]$dontSendEmailReport"
    #$unifiedGroupObject = Get-UnifiedGroup "Energy Engineering Team (All)"
    $itAdminEmailAddress = "kevin.maitland@anthesisgroup.com"

    #$foundManagersGroup = Get-AzureADMSGroup -Id $($unifiedGroupObject.CustomAttribute2)
    $foundMembersGroup = Get-AzureADMSGroup -Id $($unifiedGroupObject.CustomAttribute3)
    #$foundOverallGroup = Get-AzureADMSGroup -id $($unifiedGroupObject.CustomAttribute4)
    if(![string]::IsNullOrWhiteSpace($365GroupMembers)){rv 365GroupMembers}
    if(![string]::IsNullOrWhiteSpace($membersDelta)){rv membersDelta}
    if(![string]::IsNullOrWhiteSpace($secGroupMembers)){rv secGroupMembers}
    if(![string]::IsNullOrWhiteSpace($membersChanged)){rv membersChanged}
    if(![string]::IsNullOrWhiteSpace($userStub)){rv userStub}

    if($foundMembersGroup){
        #Get the members for the 365 Group from AAD
        $365GroupMembers = @() #Not only do we /never/ want to add users to the wrong group, having an intantiated empty array helps with compare-object later
        $secGroupMembers = @()
        Get-AzureADGroupMember -All:$true -ObjectId $unifiedGroupObject.ExternalDirectoryObjectId | %{[array]$365GroupMembers += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $_.UserPrincipalName;"displayName"=$_.DisplayName;"objectId"=$_.ObjectId})}
        #Get the members of the Security Group (this currently has to be done via Exchange for mail-enabled security groups)
        Get-DistributionGroupMember -Identity $foundMembersGroup.Id | %{[array]$secGroupMembers += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $_.WindowsLiveId;"displayName"=$_.DisplayName;"objectId"=$_.Guid})}

        #Update the Security Group membership based on the 365 Group membership
        $membersDelta = Compare-Object -ReferenceObject $365GroupMembers -DifferenceObject $secGroupMembers -Property userPrincipalName -PassThru 
        #Add extra members in the 365 Group
        $membersDelta | ?{$_.SideIndicator -eq "<="} | %{ 
            $userStub = $_
            try {
                log-action -myMessage "Attempting to add new 365 Group Member [$($userStub.displayName) | $($userStub.objectId)] to AAD Group [$($foundMembersGroup.DisplayName)]" -logFile $fullLogFile
                if($reallyDoIt){
                    Add-DistributionGroupMember -Identity $foundMembersGroup.Id -Member $userStub.objectId -BypassSecurityGroupManagerCheck:$true
                    log-result -myMessage "Success! (or, at least no error!)" -logFile $fullLogFile
                    }
                else{log-result -myMessage "We're only pretending to do this anyway..." -logFile $fullLogFile}
                [array]$membersChanged += (New-Object psobject -Property $([ordered]@{"Change"="Added";"ActionedGroupName"=$foundMembersGroup.Mail;"SourceGroupName"=$unifiedGroupObject.PrimarySmtpAddress;"UPN"=$userStub.userPrincipalName;"DisplayName"=$userStub.displayName;"Result"="Succeeded";"ErrorMessage"=$null}))
                }
            catch {
                [array]$membersChanged += (New-Object psobject -Property $([ordered]@{"Change"="Added";"ActionedGroupName"=$foundMembersGroup.Mail;"SourceGroupName"=$unifiedGroupObject.PrimarySmtpAddress;"UPN"=$userStub.userPrincipalName;"DisplayName"=$userStub.displayName;"Result"="Failed";"ErrorMessage"=$_}))
                log-error -myError $_ -myFriendlyMessage "Failed to add new 365 Group Member [$($userStub.displayName) | $($userStub.objectId)] to [$($unifiedGroupObject.DisplayName)]" -fullLogFile $fullLogFile -errorLogFile $errorLogFile
                }
            }
        #Remove "removed" members in the 365 Group
        $membersDelta | ?{$_.SideIndicator -eq "=>"} | %{ 
            $userStub = $_
            try {
                log-action -myMessage "Attempting to remove incorrect 365 Group Owner [$($userStub.displayName)] from 365 Group [$($unifiedGroupObject.DisplayName)]" -logFile $fullLogFile
                if($reallyDoIt){
                    Remove-DistributionGroupMember -Identity $foundMembersGroup.Id -Member $_.userPrincipalName -Confirm:$false -BypassSecurityGroupManagerCheck:$true
                    log-result -myMessage "Success! (or, at least no error!)" -logFile $fullLogFile
                    }
                else{log-result -myMessage "We're only pretending to do this anyway..." -logFile $fullLogFile}
                [array]$membersChanged += (New-Object psobject -Property $([ordered]@{"Change"="Removed";"ActionedGroupName"=$foundMembersGroup.Mail;"SourceGroupName"=$unifiedGroupObject.WindowsEmailAddress;"UPN"=$userStub.userPrincipalName;"DisplayName"=$userStub.displayName;"Result"="Succeeded";"ErrorMessage"=$null}))
                }
            catch {
                [array]$membersChanged += (New-Object psobject -Property $([ordered]@{"Change"="Removed";"ActionedGroupName"=$foundMembersGroup.Mail;"SourceGroupName"=$unifiedGroupObject.WindowsEmailAddress;"UPN"=$userStub.userPrincipalName;"DisplayName"=$userStub.displayName;"Result"="Failed";"ErrorMessage"=$_}))
                log-error -myError $_ -myFriendlyMessage "Failed to remove incorrect 365 Group Member [$($userStub.displayName)] from AAD Group [$($foundMembersGroup.DisplayName)]" -fullLogFile $fullLogFile -errorLogFile $errorLogFile
                }
            }
        }  
    if(!$dontSendEmailReport -and $membersChanged){report-groupMembershipSync -groupChangesArray $membersChanged -changesAreToGroupOwners $false -actionedGroupIs365 $false -emailAddressForOverviewReport $itAdminEmailAddress}
    if(!$dontSendEmailReport -and !$ownersChanged){log-result -myMessage "No Changes to Members" -logFile $fullLogFile}
    if($dontSendEmailReport){log-result -myMessage "Report specifically not requested" -logFile $fullLogFile}
    }
function sync-groupMemberships(){
    [CmdletBinding(SupportsShouldProcess=$true )]
    param(
        [Parameter(Mandatory=$true,ParameterSetName="365GroupObjectSupplied")]
        [Parameter(Mandatory=$false,ParameterSetName="AADGroupObjectSupplied")]
        [PSObject]$UnifiedGroup
        ,[Parameter(Mandatory=$false,ParameterSetName="365GroupObjectSupplied")]
        [Parameter(Mandatory=$true,ParameterSetName="AADGroupObjectSupplied")]
        [Microsoft.Open.MSGraph.Model.MsGroup]$AADGroup
        ,[Parameter(Mandatory=$true,ParameterSetName="365GroupIdOnly")]
        [Parameter(Mandatory=$false,ParameterSetName="AADGroupIdOnly")]
        [string]$UnifiedGroupId
        ,[Parameter(Mandatory=$false,ParameterSetName="365GroupIdOnly")]
        [Parameter(Mandatory=$true,ParameterSetName="AADGroupIdOnly")]
        [string]$AADGroupId
        ,[Parameter(Mandatory=$true,ParameterSetName="365GroupObjectSupplied")]
        [Parameter(Mandatory=$true,ParameterSetName="AADGroupObjectSupplied")]
        [Parameter(Mandatory=$true,ParameterSetName="365GroupIdOnly")]
        [Parameter(Mandatory=$true,ParameterSetName="AADGroupIdOnly")]
        [ValidateSet("Members", "Owners")]
        [string]$syncWhat
        ,[Parameter(Mandatory=$true,ParameterSetName="365GroupObjectSupplied")]
        [Parameter(Mandatory=$true,ParameterSetName="AADGroupObjectSupplied")]
        [Parameter(Mandatory=$true,ParameterSetName="365GroupIdOnly")]
        [Parameter(Mandatory=$true,ParameterSetName="AADGroupIdOnly")]
        [ValidateSet("365", "AAD")]
        [string]$sourceGroup
        ,[Parameter(Mandatory=$false,ParameterSetName="365GroupObjectSupplied")]
        [Parameter(Mandatory=$false,ParameterSetName="AADGroupObjectSupplied")]
        [Parameter(Mandatory=$false,ParameterSetName="365GroupIdOnly")]
        [Parameter(Mandatory=$false,ParameterSetName="AADGroupIdOnly")]
        [bool]$dontSendEmailReport = $false
        ,[Parameter(Mandatory=$false,ParameterSetName="365GroupObjectSupplied")]
        [Parameter(Mandatory=$false,ParameterSetName="AADGroupObjectSupplied")]
        [Parameter(Mandatory=$false,ParameterSetName="365GroupIdOnly")]
        [Parameter(Mandatory=$false,ParameterSetName="AADGroupIdOnly")]
        [string[]]$adminEmailAddresses = $false
        )

    #region Get $UnifiedGroup and $AADGroup, regardless of which parameters we've been given
    switch ($PsCmdlet.ParameterSetName){
        “365GroupIdOnly”  {
            Write-Verbose "We've been given a 365 Id, so we need the Group objects"
            $UnifiedGroup = Get-UnifiedGroup $UnifiedGroupId
            if(!$UnifiedGroup){
                Write-Error "Could not retrieve Unified Group from ID [$UnifiedGroupId]"
                break
                }
            if(![string]::IsNullOrWhiteSpace($AADGroupId)){
                Write-Verbose "Getting AAD Group from ID [$AADGroupId]"
                $AADGroup = Get-AzureADMSGroup -Id $AADGroupId
                if(!$AADGroup){
                    Write-Error "Could not retrieve AAD Group with ID [$AADGroupId]. Exiting without attempting to find the AAD Group associated with 365 Group [$UnifiedGroupId].."
                    break
                    }
                }
            }
        “AADGroupIdOnly”  {
            Write-Verbose "We've been given an AAD Id, so we need the Group objects"
            $AADGroup = Get-AzureADMSGroup -Id $AADGroupId
            if(!$AADGroup){
                Write-Error "Could not retrieve AAD Group from ID [$AADGroupId]. Cannot continue."
                break
                }
            if(![string]::IsNullOrWhiteSpace($UnifiedGroupId)){
                $UnifiedGroup = Get-UnifiedGroup -Identity $UnifiedGroupId
                if(!$UnifiedGroup){
                    Write-Error "Could not retrieve 365 Group with Id [$UnifiedGroupId]. Exiting without attempting to find the 365 Group associated with MESG [$AADGroupId]."
                    break
                    }
                }
            }
        #Now we've definitely got either $UnifiedGroup or $AADGroup, get the other one if it hasn't been supplied as a parameter
        {$_ -in "365GroupIdOnly","365GroupObjectSupplied"}  {
            if([string]::IsNullOrWhiteSpace($AADGroup)){
                switch ($syncWhat){
                    "Members" {
                        Write-Verbose "No `$AADGroup or `$AADGroupId provided - looking for Members group with Id [$($UnifiedGroup.CustomAttribute3)] linked to UG [$($UnifiedGroup.DisplayName)][$($UnifiedGroup.ExternalDirectoryObjectId)]"
                        $AADGroup = Get-AzureADMSGroup -Id ($UnifiedGroup.CustomAttribute3)
                        if(!$AADGroup){
                            Write-Error "Could not retrieve AAD Members Group from ID [$($UnifiedGroup.CustomAttribute3)]. Cannot continue."
                            break
                            }
                        }
                    "Owners"  {
                        Write-Verbose "No `$AADGroup or `$AADGroupId provided - looking for Owners group with Id [$($UnifiedGroup.CustomAttribute2)] linked to UG [$($UnifiedGroup.DisplayName)][$($UnifiedGroup.ExternalDirectoryObjectId)]"
                        $AADGroup = Get-AzureADMSGroup -Id ($UnifiedGroup.CustomAttribute2)
                        if(!$AADGroup){
                            Write-Error "Could not retrieve AAD Owners Group from ID [$($UnifiedGroup.CustomAttribute2)]. Cannot continue."
                            break
                            }
                        }
                    }
                }            
            }
        {$_ -in "AADGroupIdOnly","AADGroupObjectSupplied"}  {
            if([string]::IsNullOrWhiteSpace($UnifiedGroup)){
                switch($syncWhat){
                    "Members" {
                        Write-Verbose "No `$UnifiedGroup or `$UnifiedGroupId provided - looking for associated 365 Group with `$UnifiedGroup.CustomAttribute3 -eq [$($AADGroup.Id)]"
                        $UnifiedGroup = Get-UnifiedGroup -Filter "CustomAttribute3 -eq '$($AADGroup.Id)'"
                        }
                    "Owners" {
                        Write-Verbose "No `$UnifiedGroup or `$UnifiedGroupId provided - looking for associated 365 Group with `$UnifiedGroup.CustomAttribute2 -eq [$($AADGroup.Id)]"
                        $UnifiedGroup = Get-UnifiedGroup  -Filter "CustomAttribute2 -eq '$($AADGroup.Id)'"
                        }
                    }
                if(!$UnifiedGroup){
                    Write-Error "Could not retrieve 365 Group based on $syncWhat AADGroupID [$($AADGroup.Id)]. Cannot continue."
                    break
                    }
                }
            
            }
        }
    #endregion
    
    if($AADGroup -and $UnifiedGroup){ #If we've got an AAD and a 365 Group to compare...
        $ugUsers = @()
        $aadgUsers = @()
        switch ($syncWhat){
            "Members" {
                Get-AzureADGroupMember -All:$true -ObjectId $UnifiedGroup.ExternalDirectoryObjectId | %{[array]$ugUsers += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $_.UserPrincipalName;"displayName"=$_.DisplayName;"objectId"=$_.ObjectId})}
                Get-AzureADGroupMember -All:$true -ObjectId $AADGroup.Id | %{[array]$aadgUsers += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $_.UserPrincipalName;"displayName"=$_.DisplayName;"objectId"=$_.ObjectId})}
                }
            "Owners" {
                Get-AzureADGroupOwner -All:$true -ObjectId $UnifiedGroup.ExternalDirectoryObjectId | %{[array]$ugUsers += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $_.UserPrincipalName;"displayName"=$_.DisplayName;"objectId"=$_.ObjectId})}
                Get-AzureADGroupMember -All:$true -ObjectId $AADGroup.Id | %{[array]$aadgUsers += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $_.UserPrincipalName;"displayName"=$_.DisplayName;"objectId"=$_.ObjectId})}
                }
            }

        $usersDelta = Compare-Object -ReferenceObject $ugUsers -DifferenceObject $aadgUsers -Property userPrincipalName -PassThru -IncludeEqual
         $($usersDelta | % {Write-Verbose $_})

        $usersAdded = @()
        $usersRemoved = @()
        $usersFailed = @()

        switch($sourceGroup){
            "365" {
                #Add extra users from UG to MESG
                $usersDelta | ?{$_.SideIndicator -eq "<="} | %{
                    $userToBeChanged = $_
                    Write-Verbose "`tAdding [$($userToBeChanged.userPrincipalName)] to [$($AADGroup.DisplayName)][$($AADGroup.Id)] MESG"
                    try{
                        Add-DistributionGroupMember -Identity $AADGroup.Id -Member $userToBeChanged.objectId -BypassSecurityGroupManagerCheck:$true -WhatIf:$WhatIfPreference -ErrorAction Stop
                        [array]$usersAdded += (New-Object psobject -Property $([ordered]@{"UPN"=$userToBeChanged.userPrincipalName;"DisplayName"=$userToBeChanged.displayName}))
                        }
                    catch{
                        Write-Warning "Failed to add [$($userToBeChanged.userPrincipalName)] to MESG [$($AADGroup.DisplayName)][$($AADGroup.Id)]" 
                        [array]$usersFailed += (New-Object psobject -Property $([ordered]@{"Change"="Added";"UPN"=$userToBeChanged.userPrincipalName;"DisplayName"=$userToBeChanged.displayName;"ErrorMessage"=$_}))
                        }
                    }

                #Remove "removed" users from MESG
                $usersDelta | ?{$_.SideIndicator -eq "=>"} | %{ 
                    $userToBeChanged = $_
                    Write-Verbose "`tRemoving [$($userToBeChanged.userPrincipalName)] from [$($AADGroup.DisplayName)][$($AADGroup.Id)] MESG"
                    try{
                        Remove-DistributionGroupMember -Identity $AADGroup.Id -Member $userToBeChanged.objectId -BypassSecurityGroupManagerCheck:$true -Confirm:$false -WhatIf:$WhatIfPreference -ErrorAction Stop
                        [array]$usersRemoved += (New-Object psobject -Property $([ordered]@{"Change"="Removed";"UPN"=$userToBeChanged.userPrincipalName;"DisplayName"=$userToBeChanged.displayName}))
                        }
                    catch{
                        Write-Warning "Failed to remove [$($userToBeChanged.userPrincipalName)] from MESG [$($AADGroup.DisplayName)][$($AADGroup.Id)]"
                        [array]$usersFailed += (New-Object psobject -Property $([ordered]@{"Change"="Removed";"UPN"=$userToBeChanged.userPrincipalName;"DisplayName"=$userToBeChanged.displayName;"ErrorMessage"=$_}))
                        }
                    }                
                }
            "AAD" {
                #Add extra users from MESG to UG
                $usersDelta | ?{$_.SideIndicator -eq "=>"} | %{
                    $userToBeChanged = $_
                    Write-Verbose "`tAdding [$($userToBeChanged.userPrincipalName)] to [$($UnifiedGroup.DisplayName)][$($UnifiedGroup.Id)] UG $syncWhat"
                    try{
                        Add-UnifiedGroupLinks -Identity $UnifiedGroup.ExternalDirectoryObjectId -Links $userToBeChanged.objectId -LinkType $syncWhat -Confirm:$false -WhatIf:$WhatIfPreference -ErrorAction Stop
                        [array]$usersAdded += (New-Object psobject -Property $([ordered]@{"UPN"=$userToBeChanged.userPrincipalName;"DisplayName"=$userToBeChanged.displayName}))
                        }
                    catch{
                        Write-Warning "Failed to add [$($userToBeChanged.userPrincipalName)] to UG $syncWhat [$($UnifiedGroup.DisplayName)][$($UnifiedGroup.Id)]" 
                        [array]$usersFailed += (New-Object psobject -Property $([ordered]@{"Change"="Added";"UPN"=$userToBeChanged.userPrincipalName;"DisplayName"=$userToBeChanged.displayName;"ErrorMessage"=$_}))
                        }
                    }

                #Remove "removed" users from UG
                $usersDelta | ?{$_.SideIndicator -eq "<="} | %{ 
                    $userToBeChanged = $_
                    Write-Verbose "`tRemoving [$($userToBeChanged.userPrincipalName)] from [$($AADGroup.DisplayName)][$($AADGroup.Id)] UG $syncWhat"
                    try{
                        Remove-UnifiedGroupLinks -Identity $UnifiedGroup.ExternalDirectoryObjectId -Links $userToBeChanged.objectId -LinkType $syncWhat -Confirm:$false -WhatIf:$WhatIfPreference -ErrorAction Stop
                        [array]$usersRemoved += (New-Object psobject -Property $([ordered]@{"Change"="Removed";"UPN"=$userToBeChanged.userPrincipalName;"DisplayName"=$userToBeChanged.displayName}))
                        }
                    catch{
                        Write-Warning "Failed to remove [$($userToBeChanged.userPrincipalName)] from UG $syncWhat [$($AADGroup.DisplayName)][$($AADGroup.Id)]"
                        [array]$usersFailed += (New-Object psobject -Property $([ordered]@{"Change"="Removed";"UPN"=$userToBeChanged.userPrincipalName;"DisplayName"=$userToBeChanged.displayName;"ErrorMessage"=$_}))
                        }
                    }                
                }
            }

        #Now report any problems/changes    
        if(!$dontSendEmailReport){
            Write-Verbose "Preparing 365 to MESG $syncWhat sync report to send to Admins & Owners"
            if($usersFailed.Count -ne 0){
                Write-Verbose "Found [$($usersFailed.Count)] problems - notifying 365 Group Admins"
                $ugUsersAfterChanges = Get-UnifiedGroupLinks -Identity $UnifiedGroup.ExternalDirectoryObjectId -LinkType $syncWhat  #Get-AzureADGroupMember is too slow and doesn't pick up the changes we've made above.
                $ugUsersAfterChanges | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name userPrincipalName -Value $_.WindowsLiveID}
                $aadgUsersAfterChanges = Get-DistributionGroupMember -Identity $AADGroup.Id
                $aadgUsersAfterChanges | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name userPrincipalName -Value $_.WindowsLiveID}
                send-membershipChangeProblemReportToAdmins -UnifiedGroup $UnifiedGroup -changesAreTo $syncWhat -usersWithProblemsArray $usersFailed -usersIn365GroupAfterChanges $ugUsersAfterChanges -usersInAADGroupAfterChanges $aadgUsersAfterChanges -adminEmailAddresses $adminEmailAddresses -WhatIf:$WhatIfPreference
                }
            else{Write-Verbose "No problems adding/removing users, not sending problem report e-mail to Admins"}

            switch($syncWhat){
                "Members" {
                    Write-Verbose "Gettings group Owners for [$($UnifiedGroup.DisplayName)]"
                    $owners = Get-AzureADGroupOwner -ObjectId $UnifiedGroup.ExternalDirectoryObjectId
                    }
                "Owners"  {
                    $owners = $ugUsers
                    }
                }            

            if($($owners.DisplayName | ? {$_ -notmatch "Ω"}).Count -eq 0){
                Write-Verbose "No active owners for 365 Group [$($UnifiedGroup.DisplayName)] - notifying Admins so that this doesn't get auto-deleted"
                send-noOwnersForGroupAlertToAdmins -UnifiedGroup $UnifiedGroup -currentOwners $owners -adminEmailAddresses $adminEmailAddresses -WhatIf:$WhatIfPreference
                }
            else{Write-Verbose "Owners look normal, not sending problem report e-mail to Admins"}

            if($usersAdded.Count -ne 0 -or $usersRemoved.Count -ne 0){
                Write-Verbose "[$($usersAdded.Count + $usersRemoved.Count)] changes made - sending the change report to managers and admins"
                $ownersEmailAddresses = $owners.UserPrincipalName
                if($syncWhat -eq "Owners"){
                    Write-Verbose "Getting all group Owners (both added and removed) for [$($UnifiedGroup.DisplayName)]"
                    $ownersEmailAddresses += $usersAdded.UPN
                    $ownersEmailAddresses += $usersRemoved.UPN
                    $ownersEmailAddresses = $ownersEmailAddresses | Select-Object -Unique
                    }
                $ugUsersAfterChanges = Get-UnifiedGroupLinks -Identity $UnifiedGroup.ExternalDirectoryObjectId -LinkType $syncWhat  #Get-AzureADGroupMember is too slow and doesn't pick up the changes we've made above.
                $ugUsersAfterChanges | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name userPrincipalName -Value $_.WindowsLiveID}
                send-membershipChangeReportToManagers -UnifiedGroup $UnifiedGroup -changesAreTo $syncWhat -usersAddedArray $usersAdded -usersRemovedArray $usersRemoved -usersWithProblemsArray $usersFailed -usersInGroupAfterChanges $ugUsersAfterChanges -adminEmailAddresses $adminEmailAddresses -ownersEmailAddresses $ownersEmailAddresses -WhatIf:$WhatIfPreference
                }
            else{Write-Verbose "No membership changes - not sending report to Mangers & Admins"}
            }
        }
    else{
        if(!$AADGroup){
            Write-Error "No AAD group found for UG [$($UnifiedGroup.DisplayName)][$($UnifiedGroup.ExternalDirectoryObjectId)]"
            break
            }
        elseif(!$UnifiedGroup){
            Write-Error "No 365 group found for AAD Group [$($AADGroup.DisplayName)][$($AADGroup.Id)]"
            break
            }
        }
    }
function sync-managersTo365GroupOwners($unifiedGroupObject,[boolean]$reallyDoIt,[boolean]$dontSendEmailReport,$fullLogFile, $errorLogFile){
    #######################################################################################################################################
    #
    #
    #   Deprecated - use sync-groupMemberships instead
    #
    #
    #######################################################################################################################################
    Write-Warning "sync-managersTo365GroupOwners is deprecated - use sync-groupMemberships instead"
    Write-Host -ForegroundColor Magenta "sync-managersTo365GroupOwners($($unifiedGroupObject.DisplayName),[boolean]$reallyDoIt,[boolean]$dontSendEmailReport)"
    log-action -myMessage "Syncronising Manager/Owner members for [$($unifiedGroupObject.DisplayName)]" -logFile $fullLogFile

    #$unifiedGroupObject = Get-UnifiedGroup "IT Team (All)"
    $itAdminEmailAddress = "kevin.maitland@anthesisgroup.com"

    $foundManagersGroup = Get-AzureADMSGroup -Id $($unifiedGroupObject.CustomAttribute2)
    if(![string]::IsNullOrWhiteSpace($365GroupOwners)){rv 365GroupOwners}
    if(![string]::IsNullOrWhiteSpace($ownersDelta)){rv ownersDelta}
    if(![string]::IsNullOrWhiteSpace($managerGroupMembers)){rv managerGroupMembers}
    if(![string]::IsNullOrWhiteSpace($ownersChanged)){rv ownersChanged}
    if(![string]::IsNullOrWhiteSpace($userStub)){rv userStub}

    if($foundManagersGroup){
        #Get the members for the 365 Group from AAD
        $365GroupOwners = @() #Not only do we /never/ want to add users to the wrong group, having an intantiated empty array helps with compare-object later
        $managerGroupMembers = @()
        Get-AzureADGroupOwner -All:$true -ObjectId $unifiedGroupObject.ExternalDirectoryObjectId | %{[array]$365GroupOwners += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $_.UserPrincipalName;"displayName"=$_.DisplayName;"objectId"=$_.ObjectId})}
        #Get the members of the Security Group (this currently has to be done via Exchange for mail-enabled security groups)
        Get-DistributionGroupMember -Identity $foundManagersGroup.Id | %{[array]$managerGroupMembers += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $_.WindowsLiveId;"displayName"=$_.DisplayName;"objectId"=$_.ExternalDirectoryObjectId})}

        #Update the Security Group membership based on the 365 Group membership
        $ownersDelta = Compare-Object -ReferenceObject $365GroupOwners -DifferenceObject $managerGroupMembers -Property userPrincipalName -PassThru 
        #Add extra members in the AD Managers Group
        $ownersDelta | ?{$_.SideIndicator -eq "=>"} | %{ 
            $userStub = $_
            try {
                log-action -myMessage "Attempting to add new 365 Group Owner [$($userStub.displayName) | $($userStub.objectId)] to 365 Group [$($unifiedGroupObject.DisplayName)]" -logFile $fullLogFile
                if($reallyDoIt){
                    Add-AzureADGroupOwner -ObjectId $unifiedGroupObject.ExternalDirectoryObjectId -RefObjectId $userStub.objectId
                    log-result -myMessage "Success! (or, at least no error!)" -logFile $fullLogFile
                    }
                else{log-result -myMessage "We're only pretending to do this anyway..." -logFile $fullLogFile}
                [array]$ownersChanged += (New-Object psobject -Property $([ordered]@{"Change"="Added";"ActionedGroupName"=$foundManagersGroup.Mail;"SourceGroupName"=$unifiedGroupObject.WindowsEmailAddress;"UPN"=$userStub.userPrincipalName;"DisplayName"=$userStub.displayName;"Result"="Succeeded";"ErrorMessage"=$null}))
                }
            catch {
                [array]$ownersChanged += (New-Object psobject -Property $([ordered]@{"Change"="Added";"ActionedGroupName"=$foundManagersGroup.Mail;"SourceGroupName"=$unifiedGroupObject.WindowsEmailAddress;"UPN"=$userStub.userPrincipalName;"DisplayName"=$userStub.displayName;"Result"="Failed";"ErrorMessage"=$_}))
                log-error -myError $_ -myFriendlyMessage "Failed to add new 365 Group Owner [$($userStub.displayName) | $($userStub.objectId)] to [$($unifiedGroupObject.DisplayName)]" -fullLogFile $fullLogFile -errorLogFile $errorLogFile
                }
            }
        #Remove unexpected Owners from the 365 Group
        $ownersDelta | ?{$_.SideIndicator -eq "<="} | %{ 
            $userStub = $_
            try {
                log-action -myMessage "Attempting to remove incorrect 365 Group Owner [$($userStub.displayName)] from 365 Group [$($unifiedGroupObject.DisplayName)]" -logFile $fullLogFile
                if($reallyDoIt){
                    Remove-AzureADGroupOwner -ObjectId $unifiedGroupObject.ExternalDirectoryObjectId -OwnerId $userStub.objectId
                    log-result -myMessage "Success! (or, at least no error!)" -logFile $fullLogFile
                    }
                else{log-result -myMessage "We're only pretending to do this anyway..." -logFile $fullLogFile}
                [array]$ownersChanged += (New-Object psobject -Property $([ordered]@{"Change"="Removed";"ActionedGroupName"=$foundManagersGroup.Mail;"SourceGroupName"=$unifiedGroupObject.WindowsEmailAddress;"UPN"=$userStub.userPrincipalName;"DisplayName"=$userStub.displayName;"Result"="Succeeded";"ErrorMessage"=$null}))
                }
            catch {
                [array]$ownersChanged += (New-Object psobject -Property $([ordered]@{"Change"="Removed";"ActionedGroupName"=$foundManagersGroup.Mail;"SourceGroupName"=$unifiedGroupObject.WindowsEmailAddress;"UPN"=$userStub.userPrincipalName;"DisplayName"=$userStub.displayName;"Result"="Failed";"ErrorMessage"=$_}))
                log-error -myError $_ -myFriendlyMessage "Failed to remove incorrect 365 Group Owner [$($userStub.displayName)] from [$($unifiedGroupObject.DisplayName)]" -fullLogFile $fullLogFile -errorLogFile $errorLogFile
                }
            }
        }  
    if(!$dontSendEmailReport -and $ownersChanged){report-groupMembershipSync -groupChangesArray $ownersChanged -changesAreToGroupOwners $true -actionedGroupIs365 $false -emailAddressForOverviewReport $itAdminEmailAddress}
    if(!$dontSendEmailReport -and !$ownersChanged){log-result -myMessage "No Changes to Managers" -logFile $fullLogFile}
    if($dontSendEmailReport){log-result -myMessage "Report specifically not requested" -logFile $fullLogFile}

    }
<#
$msolCredentials = set-MsolCredentials #Set these once as a PSCredential object and use that to build the CSOM SharePointOnlineCredentials object and set the creds for REST
$restCredentials = new-spoCred -username $msolCredentials.UserName -securePassword $msolCredentials.Password
$csomCredentials = new-csomCredentials -username $msolCredentials.UserName -password $msolCredentials.Password
connect-ToMsol -credential $msolCredentials
connect-toAAD -credential $msolCredentials
connect-ToExo -credential $msolCredentials


$displayName = "Finance Team (DEU)"
$description = ""
$managers = @("Marie.Jones")
$teamMembers = @("Marie.Jones")
$memberOf = @("Finance Team (All)")
$hideFromGal = $false
$blockExternalMail = $false
$isPublic = $false
$autoSubscribe = $true

$displayName = "Energy Management Team (All)"
$description = $null
$managers = @("Matt.Whitehead")
$teamMembers = convertTo-arrayOfEmailAddresses "Amy.Dartington@anthesisgroup.com
Ben.Lynch@anthesisgroup.com
Duncan.Faulkes@anthesisgroup.com
Georgie.Edwards@anthesisgroup.com
James.Carberry@anthesisgroup.com
Josep.Porta@anthesisgroup.com
Matt.Landick@anthesisgroup.com
Matt.Whitehead@anthesisgroup.com
Matthew.Gitsham@anthesisgroup.com
Stuart.Gray@anthesisgroup.com
Tom.Willis@anthesisgroup.com"
#$memberOf = @("ITTeam")
$hideFromGal = $false
$blockExternalMail = $true
$isPublic = $false
$autoSubscribe = $true


$displayName = "Working Group - Collaboration Improvement"
$description = ""
$managers = @("Dee.Moloney")
$teamMembers = @("Dee.Moloney","kevin.maitland","helen.tyrrell","ian.forrester","craig.simmons","rosanna.collorafi","laura.thompson")


$displayName = "Working Group - Plastics"
$description = ""
$managers = @("Pearl.Nemeth")
$teamMembers = @("Pearl.Nemeth","Beth.Simpson")
$memberOf =$null

$displayName = "Energy Engineering Team (All)"
$description = $null
$managers = @("Ben.Lynch","Chris.Jennings")
$teamMembers = convertTo-arrayOfEmailAddresses "Alex.Matthews@anthesisgroup.com
Ben.Lynch@anthesisgroup.com
Chris.Jennings@anthesisgroup.com
Duncan.Faulkes@anthesisgroup.com
Gavin.Way@anthesisgroup.com
josep.porta@anthesisgroup.com
Matt.Landick@anthesisgroup.com
Matthew.Gitsham@anthesisgroup.com
Stuart.Miller@anthesisgroup.com
Thomas.Milne@anthesisgroup.com
Ben.Lynch@anthesisgroup.com
Huw.Blackwell@anthesisgroup.com
Laurie.Eldridge@anthesisgroup.com
Stuart.Miller@anthesisgroup.com
pete.best@anthesisgroup.com
Thomas.Milne@anthesisgroup.com"
#$memberOf = @("ITTeam")
$hideFromGal = $false
$blockExternalMail = $true
$isPublic = $false
$autoSubscribe = $true
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers


new-365Group -displayName $displayName -description $null -managers $managers -teamMembers $teamMembers -memberOf $memberOf -hideFromGal $false -blockExternalMail $true -isPublic $false -autoSubscribe $autoSubscribe


Cristina.Knapp@anthesisgroup.com


	
$displayName = "Health & Safety Team (GBR)"
$description = $null
$managers = @("Andy.Marsh")
$teamMembers = convertTo-arrayOfStrings "Amanda.Cox
Andy.Marsh
Ben.Hardman
Ian.Forrester
Nigel.Arnott
Sophie.Taylor
Wai.Cheung"

new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Corporate Social Responsibility (CSR) Team (All)"
$description = $null
$managers = @("Helen.Tyrrell")
$teamMembers = convertTo-arrayOfStrings "Helen.Tyrrell"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Confidential Human Resources (HR) Team (GBR)"
$description = $null
$managers = @("Helen.Tyrrell")
$teamMembers = convertTo-arrayOfStrings "Helen.Tyrrell"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Confidential Finance Team (GBR - Energy)"
$description = $null
$managers = @("kevin.maitland")
$teamMembers = convertTo-arrayOfStrings "Greg.Francis
Kath.Addison-Scott"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Finance Team (GBR - Energy)"
$description = $null
$managers = @("Greg.Francis")
$teamMembers = convertTo-arrayOfStrings "Greg.Francis
Kath.Addison-Scott"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Administration Team (GBR)"
$description = $null
$managers = @("Helen.Tyrrell")
$teamMembers = convertTo-arrayOfStrings "amanda.cox
laura.thompson
elle.smith
wai.cheung"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers


$displayName = "Working Group - Kimble"
$description = $null
$managers = @("Laura.Thompson")
$teamMembers = convertTo-arrayOfEmailAddresses "Craig Simmons <Craig.Simmons@anthesisgroup.com>; Greg Francis <Greg.Francis@anthesisgroup.com>; Ian Forrester <Ian.Forrester@anthesisgroup.com>; Jason Urry <Jason.Urry@anthesisgroup.com>; John Heckman <john.heckman@anthesisgroup.com>; Kev Maitland <kevin.maitland@anthesisgroup.com>; Laura Thompson <Laura.Thompson@anthesisgroup.com>; Maggie Weglinski <maggie.weglinski@anthesisgroup.com>; Phil Harrison <Phil.Harrison@anthesisgroup.com>; Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>; Rosanna Collorafi <Rosanna.Collorafi@anthesisgroup.com>; Sophie Taylor <Sophie.Taylor@anthesisgroup.com>; Tobias Parker <Tobias.Parker@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Senior Leadership Team (GBR)"
$description = $null
$managers = @("elle.wright","Stuart.McLachlan")
$teamMembers = convertTo-arrayOfEmailAddresses "Stuart McLachlan <Stuart.McLachlan@anthesisgroup.com>; Elle Wright <Elle.Wright@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Administration Team (North America)"
$description = $null
$managers = @("rosanna.collorafi","maggie.weglinski")
$teamMembers = convertTo-arrayOfEmailAddresses "Rosanna Collorafi <Rosanna.Collorafi@anthesisgroup.com>; Maggie Weglinski <maggie.weglinski@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Finance Team (North America)"
$description = $null
$managers = @("rosanna.collorafi")
$teamMembers = convertTo-arrayOfEmailAddresses "Rosanna Collorafi <Rosanna.Collorafi@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers


$displayName = "Human Resources (HR) Team (North America)"
$description = $null
$managers = @("rosanna.collorafi")
$teamMembers = convertTo-arrayOfEmailAddresses "Rosanna Collorafi <Rosanna.Collorafi@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Software Team (GBR)"
$description = $null
$managers = @("kevin.maitland","praveenaa.kathirvasan")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>; Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Software Team (PHI)"
$description = $null
$managers = @("kevin.maitland","praveenaa.kathirvasan")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>; Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Contracts & Project Management Team (All)"
$description = $null
$managers = @("kevin.maitland","praveenaa.kathirvasan")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>; Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Senior Leadership Team (GBR)"
$description = $null
$managers = @("kevin.maitland","praveenaa.kathirvasan")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>; Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Senior Management Team (Energy Division)"
$description = $null
$managers = @("kevin.maitland","praveenaa.kathirvasan")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>; Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Senior Management Team (Sustainability Division)"
$description = $null
$managers = @("kevin.maitland","praveenaa.kathirvasan")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>; Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Strategy & Communications (S&C) Team (All)"
$description = $null
$managers = @("kevin.maitland","praveenaa.kathirvasan")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>; Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Waste & Resource Sustainability (WRS) Team (All)"
$description = $null
$managers = @("kevin.maitland","praveenaa.kathirvasan")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>; Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Transaction & Corporate Services (TCS) Team (All)"
$description = $null
$managers = @("kevin.maitland","praveenaa.kathirvasan")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>; Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Sustainable Chemistry Team (All)"
$description = $null
$managers = @("kevin.maitland","praveenaa.kathirvasan")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>; Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Senior Management Team (GBR)"
$description = $null
$managers = @("kevin.maitland","praveenaa.kathirvasan")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>; Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Energy & Carbon Consulting, Analysts & Software (ECCAST) Community"
$description = $null
$managers = @("kevin.maitland","praveenaa.kathirvasan")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>; Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Analysts Team (All)"
$description = $null
$managers = @("kevin.maitland","praveenaa.kathirvasan")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>; Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Sales Team (All)"
$description = $null
$managers = @("kevin.maitland","praveenaa.kathirvasan")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>; Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Sales Team (GBR)"
$description = $null
$managers = @("kevin.maitland","praveenaa.kathirvasan")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>; Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Sales Team (North America)"
$description = $null
$managers = @("kevin.maitland","praveenaa.kathirvasan")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>; Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Diversity & Inclusivity (GBR)"
$description = $null
$managers = @("kevin.maitland","praveenaa.kathirvasan")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>; Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Data Visualisation Team (All)"
$description = $null
$managers = @("kevin.maitland","praveenaa.kathirvasan")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>; Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Carbon Consulting Team (All)"
$description = $null
$managers = @("kevin.maitland","praveenaa.kathirvasan")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>; Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Environmental, Social & Governance (ESG) Team (All)"
$description = $null
$managers = @("kevin.maitland","praveenaa.kathirvasan")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>; Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "STEP Team (All)"
$description = $null
$managers = @("kevin.maitland","praveenaa.kathirvasan")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>; Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Pulse Team (All)"
$description = $null
$managers = @("kevin.maitland","praveenaa.kathirvasan")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>; Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Diversity & Inclusivity (GBR)"
$description = $null
$managers = @("kevin.maitland","praveenaa.kathirvasan")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>; Praveenaa Kathirvasan <Praveenaa.Kathirvasan@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Health & Safety Team (North America)"
$description = $null
$managers = @("kevin.maitland")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Recruitment Team (GBR)"
$description = $null
$managers = @("kevin.maitland")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Administration & Human Resources (HR) Team (ITA)"
$description = $null
$managers = @("kevin.maitland")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Management Team (ITA)"
$description = $null
$managers = @("kevin.maitland")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Environmental Planning and Permitting Team (All)"
$description = $null
$managers = @("kevin.maitland")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Recruitment Team (SWE)"
$description = $null
$managers = @("kevin.maitland")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Finance Team (SWE)"
$description = $null
$managers = @("kevin.maitland")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Finance Team (FIN)"
$description = $null
$managers = @("emily.pressey")
$teamMembers = convertTo-arrayOfEmailAddresses "Emily Pressey <emily.pressey@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers

$displayName = "Lavola Integration Team (All)"
$description = $null
$managers = @("kevin.maitland")
$teamMembers = convertTo-arrayOfEmailAddresses "Kev Maitland <kevin.maitland@anthesisgroup.com>"
new-teamGroup -displayName $displayName -managers $managers -teamMembers $teamMembers
#>

