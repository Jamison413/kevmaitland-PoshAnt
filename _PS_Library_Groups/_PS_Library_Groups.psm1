#Sync Office 365 Group membership to correspnoding security group membership

Import-Module _PS_Library_MSOL.psm1
Import-Module _PS_Library_GeneralFunctionality
#Import-Module *pnp*

#region Functions
function guess-aliasFromDisplayName($displayName){
    Write-Host -ForegroundColor Magenta "guess-aliasFromDisplayName($displayName)"
    if(![string]::IsNullOrWhiteSpace($displayName)){$guessedAlias = $displayName.replace(" ","_").Replace("(","").Replace(")","").Replace(",","")}
    if($guessedAlias.length -gt 64){$guessedAlias = $guessedAlias.SubString(0,64)} 
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
    # [Shared Mailbox Bodge - Dummy Team (All)] - Shared Mailbox (for bodging DG membership)
    # [Dummy Team (All) - Managers] - Mail-enabled Security Group for Managers
    # [Dummy Team (All) - 365 Mirror] - Mail-enabled Security Group Mirroring Unified Group Members
    Write-Host -ForegroundColor Magenta "new-365Group($displayName, $description, $managers, $teamMembers, $memberOf, $hideFromGal, $blockExternalMail, $isPublic, $autoSubscribe, $additionalEmailAddress, $groupClassification, $ownersAreRealManagers)"
    $shortName = $displayName.Replace(" (All)","")
    try{
        #First, create a corresponding mail-enabled Security group
        Write-Host -ForegroundColor Yellow "Creating Mail-Enabled Security Group [$displayName]"
        if(Get-DistributionGroup -Identity $(guess-aliasFromDisplayName $displayName) -ErrorAction SilentlyContinue){Write-Host -ForegroundColor DarkYellow "`tCorresponding Security Group already exists";$onlyUpdate = $true}
        else{$onlyUpdate = $false}
        $sg = new-mailEnabledDistributionGroup -dgDisplayName $displayName -members $teamMembers -memberOf $memberOf -hideFromGal $false -blockExternalMail $true -owners "ITTeamAll@anthesisgroup.com" -description "Mail-enabled Security Group for $displayName" -onlyUpdate $onlyUpdate
        
        try{
            #Then create a Managers mail-enabled Security group
            Write-Host -ForegroundColor Yellow "Creating Managers Mail-Enabled Security Group [$displayName - Managers]"
            if(Get-DistributionGroup -Identity $("$displayName - Managers") -ErrorAction SilentlyContinue){Write-Host -ForegroundColor DarkYellow "`tManagers Security Group already exists";$onlyUpdate = $true}
            else{$onlyUpdate = $false}
            
            $managersMemberOf =@($sg.ExternalDirectoryObjectId)
            if($ownersAreRealManagers){$managersMemberOf += "Managers (All)"}
            $managerSG = new-mailEnabledDistributionGroup -dgDisplayName $("$displayName - Managers") -members $managers -memberOf $managersMemberOf -hideFromGal $true -blockExternalMail $true -owners "ITTeamAll@anthesisgroup.com" -description "Mail-enabled Security Group for $shortName Managers" -onlyUpdate $onlyUpdate

            Write-Host -ForegroundColor Yellow "Creating 365 Mirror Mail-Enabled Security Group [$displayName - 365 Mirror]"
            if(Get-DistributionGroup -Identity $("$displayName - 365 Mirror") -ErrorAction SilentlyContinue){Write-Host -ForegroundColor DarkYellow "`t365 Mirror Security Group already exists";$onlyUpdate = $true}
            else{$onlyUpdate = $false}
            $mirrorSG = new-mailEnabledDistributionGroup -dgDisplayName $("$displayName - 365 Mirror") -members $teamMembers -memberOf $sg.ExternalDirectoryObjectId -hideFromGal $true -blockExternalMail $true -owners "ITTeamAll@anthesisgroup.com" -description "Mail-enabled Security Group for mirroring membership of $shortName Unified Group" -onlyUpdate $onlyUpdate

            try{
                #Then, if that's worked, create the 365 Group
                Write-Host -ForegroundColor Yellow "Creating Unified 365 Group [$displayName]"
                if($isPublic){$accessType = "Public"}else{$accessType = "Private"}
                $mailAlias = $(guess-aliasFromDisplayName "$displayName 365")
                if($mailAlias.length -gt 64){$mailAlias = $mailAlias.substring(0,63)}
                if([string]::IsNullOrWhiteSpace($description)){$description = "Unified 365 Group for $displayName"}
                if(Get-UnifiedGroup -Identity $mailAlias  -ErrorAction SilentlyContinue){Write-Host -ForegroundColor Yellow "Unified Group already exists - not recreating!"}
                else{New-UnifiedGroup -DisplayName $displayName -Name $mailAlias -Alias $mailAlias -Notes $description -AccessType $accessType -Owner $managers[0] -RequireSenderAuthenticationEnabled $blockExternalMail -AutoSubscribeNewMembers:$autoSubscribe -AlwaysSubscribeMembersToCalendarEvents:$autoSubscribe -Members $teamMembers   -Classification $groupClassification | Set-UnifiedGroup -HiddenFromAddressListsEnabled $true}
                $ug = Get-UnifiedGroup -Identity $mailAlias
                if($managers.Count -gt 1){Add-UnifiedGroupLinks -Identity $ug.Identity -LinkType Owner -Links $managers -Confirm:$false}

                #Create a Shared Mailbox and autoforward mail to the Unified Group
                Write-Host -ForegroundColor Yellow "Creating Shared Mailbox [Shared Mailbox Bodge - $displayName]"
                $sm = New-Mailbox -Shared -DisplayName "Shared Mailbox Bodge - $displayName" -Name "Shared Mailbox Bodge - $displayName" -Alias $(guess-aliasFromDisplayName ("Shared Mailbox Bodge - $displayName")) -ErrorAction Continue
                sleep -Seconds 15
                if($sm -eq $null){
                    Write-Host -ForegroundColor DarkYellow "Shared Mailbox could not create - trying to retrieve instead"
                    $sm = Get-Mailbox $(guess-aliasFromDisplayName ("Shared Mailbox Bodge - $displayName"))
                    }
                if($sm){
                    #DeliverToMailboxAndForward has to be true, otherwise it just doesn't forward :/
                    Set-Mailbox -Identity $sm.ExternalDirectoryObjectId -HiddenFromAddressListsEnabled $true -RequireSenderAuthenticationEnabled $false -ForwardingAddress $ug.PrimarySmtpAddress -DeliverToMailboxAndForward $true -ForwardingSmtpAddress $ug.PrimarySmtpAddress -Confirm:$false
                    Set-user -Identity $sm.ExternalDirectoryObjectId -Manager kevin.maitland #For want of someone better....
                    #Assign the Shared Mailbox as a member of the Security Group
                    Add-DistributionGroupMember -Identity $sg.ExternalDirectoryObjectId -Member $sm.ExternalDirectoryObjectId -BypassSecurityGroupManagerCheck
                    }
                else{Write-Host -ForegroundColor DarkMagenta "Failed to create 365 Shared Mailbox "}
                }
            catch{
                Write-Host -ForegroundColor DarkMagenta "Failed to create 365 Group (but Security groups seemed to work)"
                $_
                }
            }
        catch{
            Write-Host -ForegroundColor DarkMagenta "Failed to create the Managers security group (but the basic Security group seemed to work) (not attempting 365 group)"
            $_
            }
        }
    catch{
        Write-Host -ForegroundColor DarkMagenta "Failed to create the corresponding security group (not attempting Manager or 365 group)"
        $_
        }
    }
function new-mailEnabledDistributionGroup($dgDisplayName, $description, $members, $memberOf, $hideFromGal, $blockExternalMail, $owners, [boolean]$onlyUpdate){
    Write-Host -ForegroundColor Magenta "new-mailEnabledDistributionGroup($dgDisplayName, $description, $members, $memberOf, $hideFromGal, $blockExternalMail, $owners, [boolean]$onlyUpdate)"
    $mailAlias = guess-aliasFromDisplayName $dgDisplayName
    $mailName = $dgDisplayName
    if($mailName.length -gt 64){$mailName = $mailName.SubString(0,64)}
    if($onlyUpdate){
        $members  | % {
            Write-Host -ForegroundColor DarkMagenta "Adding TeamMembers Add-DistributionGroupMember $mailAlias -Member $_ -Confirm:$false -BypassSecurityGroupManagerCheck"
            Add-DistributionGroupMember $mailAlias -Member $_ -Confirm:$false -BypassSecurityGroupManagerCheck
            }
        }
    else{
        try{
            Write-Host -ForegroundColor DarkMagenta "New-DistributionGroup -Name $mailName -DisplayName $dgDisplayName -Type Security -Members $members -PrimarySmtpAddress $($mailAlias+"@anthesisgroup.com") -Notes $description -Alias $mailAlias"
            New-DistributionGroup -Name $mailName -DisplayName $dgDisplayName -Type Security -Members $members -PrimarySmtpAddress $($mailAlias+"@anthesisgroup.com") -Notes $description -Alias $mailAlias #| Out-Null
            }
        catch{$Error}
        }
    Write-Host -ForegroundColor DarkMagenta "Set-DistributionGroup -Identity $mailAlias -HiddenFromAddressListsEnabled $hideFromGal -RequireSenderAuthenticationEnabled $blockExternalMail -ManagedBy $owners"
    Set-DistributionGroup -Identity $mailAlias -HiddenFromAddressListsEnabled $hideFromGal -RequireSenderAuthenticationEnabled $blockExternalMail -ManagedBy $owners
    $memberOf | % {
        if(![string]::IsNullOrWhiteSpace($_)){
            Write-Host -ForegroundColor DarkYellow "Adding As MemberOf Add-DistributionGroupMember [$_] -Member [$mailAlias] -Confirm:$false -BypassSecurityGroupManagerCheck"
            Add-DistributionGroupMember $_ -Member $mailAlias -Confirm:$false -BypassSecurityGroupManagerCheck
            }
        }
    Write-Host -ForegroundColor DarkMagenta "Get-DistributionGroup $mailAlias | ? {$_.alias -eq $mailAlias}"
    Get-DistributionGroup $mailAlias | ? {$_.alias -eq $mailAlias}
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
function report-groupMembershipSync($groupChangesArray,[boolean]$changesAreToGroupOwners,[boolean]$actionedGroupIs365,$emailAddressForOverviewReport){
    Write-Host -ForegroundColor Magenta "report-groupMembershipSync($groupChangesArray,[boolean]$changesAreToGroupOwners,[boolean]$actionedGroupIs365,$emailAddressForOverviewReport"
    #$groupChangesArray = $ownersChanged
    if($actionedGroupIs365){$groupChangesArray = $groupChangesArray | Sort-Object ActionedGroupName,Result,Change,DisplayName}
    else{$groupChangesArray = $groupChangesArray | Sort-Object SourceGroupName,Result,Change,DisplayName}
    $groupChangesArray | %{
        if($current365Group.Mail -ne $_.SourceGroupName -and $current365Group.Mail -ne $_.ActionedGroupName){
            #We need to start another report, so send the current one before we start again
            if($ownerReport){
                Write-Host $ownerReport
                send-membershipEmailReport -ownerReport $ownerReport -changesAreToGroupOwners $changesAreToGroupOwners -emailAddressForOverviewReport $emailAddressForOverviewReport
                }
            #Start new ownerReport
            $ownerReport = New-Object psobject -Property $([ordered]@{"To"=@();"groupName"=$null;"added"=@();"removed"=@();"problems"=@();"fullMemberList"=@()})
            if($actionedGroupIs365){$current365Group = Get-AzureADMSGroup -Filter "Mail eq '$($_.ActionedGroupName)'"}
            else{$current365Group = Get-AzureADMSGroup -Filter "Mail eq '$($_.SourceGroupName)'"}
            $ownerReport.groupName = $current365Group.DisplayName
            #Get the owners' e-mail addresses
            #[array]$owners = Get-AzureADMSGroup -SearchString $current365GroupName | ? {$_.GroupTypes -contains "Unified"} | % {$(Get-AzureADGroupOwner -All:$true -ObjectId $_.Id).UserPrincipalName}
            [array]$owners = $current365Group | % {$(Get-AzureADGroupOwner -All:$true -ObjectId $_.Id).UserPrincipalName}
            
            if($owners){$ownerReport.To = $owners}
            else{
                $ownerReport.To = $emailAddressForOverviewReport
                $ownerReport.groupName = "***Unowned Group*** $current365GroupName"
                }
            #Get the members' (or owners' if we're reporting on group Ownership) DisplayNames
            if($changesAreToGroupOwners){
                #[array]$members = Get-AzureADMSGroup -SearchString $current365GroupName | ? {$_.GroupTypes -contains "Unified"} | % {$(Get-AzureADGroupOwner -All:$true -ObjectId $_.Id).DisplayName}
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
        if($_.Result -eq "Succeeded"){
            if($_.Change -eq "Added"){$ownerReport.added += $_.DisplayName}
            else{$ownerReport.Removed += $_.DisplayName}
            }
        #Add any failures as problems to be investigated manually
        else{$ownerReport.problems += $_.DisplayName}
        }
    #Finally, send the last reports too
    Write-Host $ownerReport
    Write-Host "To: " + $ownerReport.To
    send-membershipEmailReport -ownerReport $ownerReport -changesAreToGroupOwners $changesAreToGroupOwners  -emailAddressForOverviewReport $emailAddressForOverviewReport
    }
function send-membershipEmailReport($ownerReport,[boolean]$changesAreToGroupOwners,$emailAddressForOverviewReport){
    Write-Host -ForegroundColor Magenta "send-membershipEmailReport($ownerReport,[boolean]$changesAreToGroupOwners,$emailAddressForOverviewReport"
    #Write and send e-mail
    if($changesAreToGroupOwners){$type = "owner"}
    else{$type = "member"}
    $subject = "$($ownerReport.groupName) $($type)ship updated"
    $body = "<HTML><FONT FACE=`"Calibri`">Hello owners of <B>$($ownerReport.groupName)</B>,`r`n`r`n<BR><BR>"
    $body += "Changes have been made to the <B><U>$($type)</U>ship</B> of $($ownerReport.groupName)`r`n`r`n<BR><BR>"
    if($ownerReport.added)  {$body += "The following users have been <B>added</B> as Group <B>$($type)s</B>:      `r`n`t<BR><PRE>&#9;$($ownerReport.added -join     "`r`n`t")</PRE>`r`n`r`n<BR>"}
    if($ownerReport.removed){$body += "The following users have been <B>removed</B> from the Group <B>$($type)s</B>:  `r`n`t<BR><PRE>&#9;$($ownerReport.removed -join   "`r`n`t")</PRE>`r`n`r`n<BR>"}
    if($ownerReport.problems){
        $body += "The were some problems processing changes to these users (but IT have been notified):`r`n`t<BR><PRE>&#9;$($ownerReport.problems -join "`r`n`t")</PRE>`r`n`r`n<BR>"
        $ownerReport.To += $emailAddressForOverviewReport
        }
    if($ownerReport.fullMemberList){$body += "The full list of group $($type)s looks like this:`r`n`t<BR><PRE>&#9;$($ownerReport.fullMemberList -join "`r`n`t")</PRE>`r`n`r`n<BR>"}
    else{$body += "It looks like the group is now empty...`r`n`r`n<BR><BR>"}
    if($type -eq "owner"){$body += "To help us all remain compliant and secure, group <I>ownership</I> is still managed centrally by your IT Team, and you will need to liaise with them to make changes to group ownership.`r`n`r`n<BR><BR>"}
    $body += "As an owner, you can manage the membership of this group (and there is a <A HREF=`"https://anthesisllc.sharepoint.com/sites/Resources-IT/SitePages/Group-membership-management-(for-Team-Managers).aspx`">guide available to help you</A>), or you can contact the IT team for your region,`r`n`r`n<BR><BR>"
    $body += "Love,`r`n`r`n<BR><BR>The Helpful Groups Robot</FONT></HTML>"
    #Send-MailMessage -To "kevin.maitland@anthesisgroup.com" -From "thehelpfulgroupsrobot@anthesisgroup.com" -cc "kevin.maitland@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
    Send-MailMessage -To $ownerReport.To -From "thehelpfulgroupsrobot@anthesisgroup.com" -cc "kevin.maitland@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
    #$body
    }
function sync-all365GroupMembersToMirroredSecurityGroups([boolean]$reallyDoIt,[boolean]$dontSendEmailReport){
    Write-Host -ForegroundColor Magenta "sync-all365GroupMembersToMirroredSecurityGroups([boolean]$reallyDoIt,[boolean]$dontSendEmailReport"
    $itAdminEmailAddress = "kevin.maitland@anthesisgroup.com"
    Get-AzureADMSGroup -All:$true | ?{$_.GroupTypes -contains "Unified"} | %{
        $365Group = $_
        #Look for the corresponding Security Group (search by alias by swapping the "_365" suffix of the 365 group for the "-365Mirror" suffix of the Mirror Groups as Owners can alter DisplayName)
        $securityGroup = Get-AzureADMSGroup -Filter "MailNickname eq '$($365Group.MailNickname.Replace("_365","_-_365") + "_Mirror")'" | ?{$_.MailEnabled -eq $true -and $_.SecurityEnabled -eq $true -and $_.GroupTypes -notcontains "Unified"}
        if($securityGroup.Count -eq 1){
            #Get the members for the 365 Group from AAD
            $365GroupMembers = @() #Not only do we /never/ want to add users to the wrong group, having an intantiated empty array helps with compare-object later
            $secGroupMembers = @()
            Get-AzureADGroupMember -All:$true -ObjectId  $365Group.Id | %{[array]$365GroupMembers += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $_.UserPrincipalName;"displayName"=$_.DisplayName;"objectId"=$_.ObjectId})}
            #Get the members of the Security Group (this currently has to be done via Exchange for mail-enabled security groups)
            Get-DistributionGroupMember -Identity $securityGroup.Id | %{[array]$secGroupMembers += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $_.WindowsLiveId;"displayName"=$_.DisplayName;"objectId"=$_.Guid})}

            #Update the Security Group membership based on the 365 Group membership
            $membersDelta = Compare-Object -ReferenceObject $365GroupMembers -DifferenceObject $secGroupMembers -Property userPrincipalName -PassThru 
            #Add extra members in the 365 Group
            $membersDelta | ?{$_.SideIndicator -eq "<="} | %{ 
                $userStub = $_
                try {
                    if($reallyDoIt){Add-DistributionGroupMember -Identity $securityGroup.Id -Member $userStub.objectId -BypassSecurityGroupManagerCheck:$true}
                    [array]$membersChanged += (New-Object psobject -Property $([ordered]@{"Change"="Added";"ActionedGroupName"=$securityGroup.Mail;"SourceGroupName"=$365Group.Mail;"UPN"=$userStub.userPrincipalName;"DisplayName"=$userStub.displayName;"Result"="Succeeded";"ErrorMessage"=$null}))
                    }
                catch {
                    [array]$membersChanged += (New-Object psobject -Property $([ordered]@{"Change"="Added";"ActionedGroupName"=$securityGroup.Mail;"SourceGroupName"=$365Group.Mail;"UPN"=$userStub.userPrincipalName;"DisplayName"=$userStub.displayName;"Result"="Failed";"ErrorMessage"=$_}))
                    }
                }
            #Remove "removed" members in the 365 Group
            $membersDelta | ?{$_.SideIndicator -eq "=>"} | %{ 
                $userStub = $_
                 try {
                    if($reallyDoIt){Remove-DistributionGroupMember -Identity $securityGroup.Id -Member $_.userPrincipalName -Confirm:$false -BypassSecurityGroupManagerCheck:$true}
                    [array]$membersChanged += (New-Object psobject -Property $([ordered]@{"Change"="Removed";"ActionedGroupName"=$securityGroup.Mail;"SourceGroupName"=$365Group.Mail;"UPN"=$userStub.userPrincipalName;"DisplayName"=$userStub.displayName;"Result"="Succeeded";"ErrorMessage"=$null}))
                    }
                catch {
                    [array]$membersChanged += (New-Object psobject -Property $([ordered]@{"Change"="Removed";"ActionedGroupName"=$securityGroup.Mail;"SourceGroupName"=$365Group.Mail;"UPN"=$userStub.userPrincipalName;"DisplayName"=$userStub.displayName;"Result"="Failed";"ErrorMessage"=$_}))
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
    if(!$dontSendEmailReport -and $membersChanged){report-groupMembershipSync -groupChangesArray $membersChanged -changesAreToGroupOwners $false -actionedGroupIs365 $false -emailAddressForOverviewReport $itAdminEmailAddress}
    }
function sync-allSecurityGroupOwnersTo365Groups([boolean]$reallyDoIt,[boolean]$dontSendEmailReport){
    Write-Host -ForegroundColor Magenta "sync-allSecurityGroupOwnersTo365Groups([boolean]$reallyDoIt,[boolean]$dontSendEmailReport)"
    $itAdminEmailAddress = "kevin.maitland@anthesisgroup.com"
    #This should be less important now as Owners cannot add Owners, it should just synchronise IT-led changes to the [Dummy Team (All) - Managers] group
    #Start with the 365 Groups as there are fewer of them
    Get-AzureADMSGroup -All:$true | ?{$_.GroupTypes -contains "Unified"} | %{
        $365Group = $_
        #Look for the corresponding Security Group
        $securityGroup = Get-AzureADMSGroup -Filter "MailNickname eq '$($365Group.MailNickname.Replace("_365",'') + "_-_Managers")'" | ?{$_.MailEnabled -eq $true -and $_.SecurityEnabled -eq $true -and $_.GroupTypes -notcontains "Unified"}
        if($securityGroup.Count -eq 1){
            #Get the owners for the 365 Group from AAD
            $365GroupOwners = @()
            $secGroupOwners = @()
            Get-AzureADGroupOwner -All:$true -ObjectId $365Group.Id | %{[array]$365GroupOwners += New-Object psobject -Property $([ordered]@{"windowsLiveID"= $_.UserPrincipalName;"displayName"=$_.DisplayName;"objectId"=$_.ObjectId})}
            #region Getting the "owners" of a mail-enabled distribution group is more complicated
            <#What are you blithering about you dopey bastard? We're never letting users manage the Managers groups - we want the /Members/ of the Managers Group
            $managedBy = $(Get-DistributionGroup -Identity $securityGroup.Id).ManagedBy 
            $managedBy | % {
                $hopefullyASingleUserMailbox = @()
                $hopefullyASingleUserMailbox += Get-Mailbox -Identity $_ #Get all the mailboxes that match each entry in the ManagedBy property
                if($hopefullyASingleUserMailbox.Count -eq 1){ #If there's only one, carry on processding
                    [array]$secGroupOwners += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $hopefullyASingleUserMailbox[0].UserPrincipalName;"displayName"=$hopefullyASingleUserMailbox[0].DisplayName;"objectId"=$hopefullyASingleUserMailbox[0].ExternalDirectoryObjectId})
                    }
                else{ #If there isn't exactly one, log an error
                    if($hopefullyASingleUserMailbox.Count -lt 1){[array]$ownersChanged += (New-Object psobject -Property $([ordered]@{"Change"="FindSecGroupOwner";"365GroupName"=$365Group.DisplayName;"UPN"=$_;"DisplayName"=$_;"Result"="Failed";"ErrorMessage"="No mailbox matches ManagedBy Alias - WTF?!?"}))}
                    else{
                        [string]$multipleAliasMatches = $hopefullyASingleUserMailbox.Alias -join ","
                        [array]$ownersChanged += (New-Object psobject -Property $([ordered]@{"Change"="FindSecGroupOwner";"365GroupName"=$365Group.DisplayName;"UPN"=$managedByStub;"DisplayName"=$managedByStub;"Result"="Failed";"ErrorMessage"="Multiple mailbox match ManagedBy Alias [$_] @($multipleAliasMatches)"}))
                        }
                    }
                }
            #>
            $membersOfManagersGroup = Get-DistributionGroupMember -Identity $securityGroup.Id
            #endregion
            #Update the 365 Group ownership based on the Security Group ownership (the opposite direction to Members)
            $ownersDelta = Compare-Object -ReferenceObject $membersOfManagersGroup -DifferenceObject $365GroupOwners -Property windowsLiveID -PassThru 
            #Add extra members in the 365 Group
            $ownersDelta | ?{$_.SideIndicator -eq "<="} | %{
                $userStub = $_
                try {
                    if($reallyDoIt){Add-AzureADGroupOwner -ObjectId $365Group.Id -RefObjectId $userStub.objectId}
                    [array]$ownersChanged += (New-Object psobject -Property $([ordered]@{"Change"="Added";"ActionedGroupName"=$365Group.Mail;"SourceGroupName"=$securityGroup.Mail;"UPN"=$userStub.windowsLiveID;"DisplayName"=$userStub.displayName;"Result"="Succeeded";"ErrorMessage"=$null}))
                    }
                catch {
                    [array]$ownersChanged += (New-Object psobject -Property $([ordered]@{"Change"="Added";"ActionedGroupName"=$365Group.Mail;"SourceGroupName"=$securityGroup.Mail;"UPN"=$userStub.windowsLiveID;"DisplayName"=$userStub.displayName;"Result"="Failed";"ErrorMessage"=$_}))
                    }
                }
            #Remove "removed" members in the 365 Group
            $ownersDelta | ?{$_.SideIndicator -eq "=>"} | %{ 
                $userStub = $_
                 try {
                    if($reallyDoIt){Remove-AzureADGroupOwner -ObjectId $365Group.Id -OwnerId $_.objectId}
                    [array]$ownersChanged += (New-Object psobject -Property $([ordered]@{"Change"="Removed";"ActionedGroupName"=$365Group.Mail;"SourceGroupName"=$securityGroup.Mail;"UPN"=$userStub.windowsLiveID;"DisplayName"=$userStub.displayName;"Result"="Succeeded";"ErrorMessage"=$null}))
                    }
                catch {
                    [array]$ownersChanged += (New-Object psobject -Property $([ordered]@{"Change"="Removed";"ActionedGroupName"=$365Group.Mail;"SourceGroupName"=$securityGroup.Mail;"UPN"=$userStub.windowsLiveID;"DisplayName"=$userStub.displayName;"Result"="Failed";"ErrorMessage"=$_}))
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
    if(!$dontSendEmailReport -and $ownersChanged){report-groupMembershipSync -groupChangesArray $ownersChanged -changesAreToGroupOwners $true -actionedGroupIs365 $true -emailAddressForOverviewReport $itAdminEmailAddress}
    }
#endregion

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


#>




#new-mailEnabledDistributionGroup -dgDisplayName "Software Team (PHI)" -members @("soren.mateo@anthesisgroup.com","michael.malate@anthesisgroup.com","gerber.manalo@anthesisgroup.com") -memberOf "Software Team (All)" -hideFromGal $false -blockExternalMail $true -owners "IT Team (All)"


