function add-SPOSitetoSharepointTeamsTermStore{
[CmdletBinding()]
Param ($displayName)

    If( ($displayName -notmatch "Sym") -and ($displayName -notmatch "Working Group") ){
        Write-Host "This isn't a Sym or Working Group, adding to the Team Term Store" -ForegroundColor Magenta 
        New-PnPTerm -TermSet "Live Sharepoint Teams" -TermGroup "Anthesis" -Name $displayName -Lcid 1033
        }
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
function enumerate-nestedAADGroups(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true,ParameterSetName="aadGroupObjectSupplied")]
        [PSObject]$aadGroupObject
        ,[parameter(Mandatory = $true,ParameterSetName="aadGroupIdOnly")]
        [string]$aadGroupId
        )
    switch ($PsCmdlet.ParameterSetName){
        “aadGroupIdOnly”  {
            if($VerbosePreference -eq "Continue"){
                Write-Verbose "We've been given an Id, but we need the AAD Group object for Verbose reporting"
                $aadGroupObject = Get-AzureADGroup -ObjectId $aadGroupId
                }
            else{
                #If we're not using -Verbose, there's no need to get the AAD Group Object too (we only use it DisplayName to help with troubleshooting)
                }
            }
        "aadGroupObjectSupplied" {
            Write-Verbose "We've already got the AAD Group object"
            $aadGroupId = $aadGroupObject.ObjectId
            }
        }

    Write-Verbose "enumerate-nestedAADGroups($($aadGroupObject.DisplayName))"
    $immediateMembers = Get-AzureADGroupMember -ObjectId $aadGroupId
    $userObjects = @()
    $immediateMembers | % {
        $thisMember = $_
        switch($thisMember.ObjectType){
            ("User") {
                [array]$userObjects += $thisMember
                Write-Verbose "`$userObjects.Count: [$($userObjects.Count)]`tAADUser [$($thisMember.DisplayName)] is a member of [$($aadGroupObject.DisplayName)]"
                }
            ("Group"){
                $subAadGroup = Get-AzureADGroup -ObjectId $thisMember.ObjectId
                Write-Verbose "Retrieved Subgroup [$($subAadGroup.DisplayName)]"
                [array]$subUserObjects = enumerate-nestedAADGroups -aadGroupObject $subAadGroup -Verbose:$VerbosePreference
                Write-Verbose "`$UserObjects.Count = $($userObjects.Count) `t`$subUserObjects.Count = $($subUserObjects.Count)"
                $userObjects += $subUserObjects
                }
            default {}
            }
        }

    $userObjects = $userObjects  | Select-Object -Unique
    $userObjects
    }
function enumerate-nestedDistributionGroups(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true,ParameterSetName="distributionGroupObjectSupplied")]
        [PSObject]$distributionGroupObject
        ,[parameter(Mandatory = $true,ParameterSetName="distributionGroupIdOnly")]
        [string]$distributionGroupId
        )
    switch ($PsCmdlet.ParameterSetName){
        “distributionGroupIdOnly”  {
            if($VerbosePreference -eq "Continue"){
                Write-Verbose "We've been given an Id, but we need the Distribution Group object for Verbose reporting"
                $distributionGroupObject = Get-DistributionGroup -Identity $distributionGroupId
                }
            else{
                #If we're not using -Verbose, there's no need to get the Distribution Group Object (we only use it for .DisplayName to help with troubleshooting)
                }
            }
        "distributionGroupObjectSupplied" {
            Write-Verbose "We've already got the Distribution Group object"
            $distributionGroupId = $distributionGroupObject.ExternalDirectoryObjectId
            }
        }
    Write-Verbose "enumerate-distributionGroupId($($distributionGroupObject.DisplayName))"

    $immediateMembers = Get-DistributionGroupMember -Identity $distributionGroupId
    $userObjects = @()
    $immediateMembers | % {
        $thisMember = $_
        switch($thisMember.RecipientTypeDetails){
            ("UserMailbox") {
                [array]$userObjects += $thisMember
                Write-Verbose "`$userObjects.Count: [$($userObjects.Count)]`tEXOUser [$($thisMember.DisplayName)] is a member of [$($distributionGroupObject.DisplayName)]"
                }
            ("MailUniversalSecurityGroup"){
                $subDistributionGroup = Get-DistributionGroup -Identity $thisMember.ExternalDirectoryObjectId
                [array]$subUserObjects = enumerate-nestedDistributionGroups -distributionGroupObject $subDistributionGroup -Verbose:$VerbosePreference
                Write-Verbose "`$userObjects.Count = [$($userObjects.Count)] `t`$subUserObjects.Count = [$($subUserObjects.Count)]"
                $userObjects += $subUserObjects
                }
            default {}
            }
        }
    $userObjects | sort Name | Get-Unique -AsString
    }
function get-membersGroup(){
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true,ParameterSetName="GroupNameOnly")]
            [string]$groupName
        ,[Parameter(Mandatory=$true,ParameterSetName="GroupIdOnly")]
            [ValidateLength(36,36)]
            [string]$groupId
        ,[Parameter(Mandatory=$true,ParameterSetName="GroupObject")]
            $groupObject
        )

    switch ($PsCmdlet.ParameterSetName){
        'GroupNameOnly' {
            if($groupName.EndsWith("Members Subgroup")){ #Nice and easy
                Write-Verbose "This looks like a Members Subgroup - searching for it directly"
                [array]$result = Get-DistributionGroup $groupName
                }
            else{ #Otherwise, assume it's a 365 group and find it that way
                Write-Verbose "I don't know what sort of group this is, so I'm going to assume it's a 365 group and find the Members Subgroup via that"
                $group = Get-UnifiedGroup -Identity $groupName -ErrorAction SilentlyContinue
                if($group){[array]$result = Get-DistributionGroup $ug.CustomAttribute3}
                else{
                    Write-Verbose "Well, that didn't work. Trying it as a Distribution Group."
                    $group = Get-DistributionGroup $groupName -ErrorAction SilentlyContinue
                    if($group){get-membersGroup -groupObject $group}
                    else{Write-Warning "[$groupName] doesn't appear to be a group :/"}
                    }
                }
            }
        'GroupIdOnly' {
            #See if it's a UG
            $group = Get-UnifiedGroup -Identity $groupId
            #Otherwise, try a DG
            if(!$group){Get-DistributionGroup $groupId}
            #Otherwise, try an AADG
            if(!$group){Get-AzureADGroup -ObjectId $groupId}
            if($group){get-membersGroup -groupObject $group}
            else{Write-Warning "Could not find a group with ID [$groupId]"}
            }
        'GroupObject' {
            switch($groupObject.RecipientTypeDetails){
                "GroupMailbox" { #It's a UG
                    Write-Verbose "This looks like a 365 group - finding it's associated Members Subgroup"
                    [array]$result = Get-DistributionGroup $groupObject.CustomAttribute3
                    break
                    }
                {[string]::IsNullOrWhiteSpace($_)}{#Assume it's an AAD Group and find the comparable DG
                    $groupObject = Get-DistributionGroup $groupObject.DisplayName
                    }
                default{#Assume it's (now) a DG/MESG and find the corresponding UG first
                    $ug = Get-UnifiedGroup -Filter "CustomAttribute2 -eq `'$($groupObject.ExternalDirectoryObjectId)`'"
                    if(!$ug){$ug = Get-UnifiedGroup -Filter "CustomAttribute3 -eq `'$($groupObject.ExternalDirectoryObjectId)`'"}
                    if(!$ug){$ug = Get-UnifiedGroup -Filter "CustomAttribute4 -eq `'$($groupObject.ExternalDirectoryObjectId)`'"}
                    [array]$result = Get-DistributionGroup $ug.CustomAttribute3
                    break
                    }
                }
            }
        }
    if($result){
        if($result.Count -gt 1){
            switch ($PsCmdlet.ParameterSetName){
                'GroupNameOnly' {Write-Warning "Multiple Members Groups found - searched using DisplayName [$groupName]"}
                'GroupIdOnly' {Write-Warning "Multiple Members Groups found - searched using Id [$groupId]"}
                'GroupObject' {Write-Warning "Multiple Members Groups found - searched using object [$($groupObject.DisplayName)]"}
                }

            $result
            }
        else{
            Write-Verbose "[$($result.Count)] results found"
            $result[0]
            }
        }
    }
function guess-aliasFromDisplayName(){
    [CmdletBinding()]
    Param (
        [parameter(Mandatory = $true)]
        [string]$displayName
        ,[parameter(Mandatory = $false)]
        [string]$fixedSuffix
        )
    #Write-Host -ForegroundColor Magenta "guess-aliasFromDisplayName($displayName)"
    if(![string]::IsNullOrWhiteSpace($displayName)){$guessedAlias = $displayName.replace(" ","_").Replace("(","").Replace(")","").Replace(",","").Replace("@","").Replace("\","").Replace("[","").Replace("]","").Replace("`"","").Replace(";","").Replace(":","").Replace("<","").Replace(">","")}
    $guessedAlias = set-suffixAndMaxLength -string $guessedAlias -suffix $fixedSuffix -maxLength 64
    $guessedAlias = sanitise-forMicrosoftEmailAddress -dirtyString $guessedAlias
    $guessedAlias = remove-diacritics -String $guessedAlias
    Write-Verbose -Message "guess-aliasFromDisplayName($displayName) = [$guessedAlias]"
    $guessedAlias
    }
function new-365Group(){
    #Groups created look like this:
    # [Dummy Team (All)] - Combined Mail-enabled Security Group (DisplayName)
    # [Dummy Team (All) - Data Managers Subgroup] - Mail-enabled Security Group for Managers
    # [Dummy Team (All) - Members Subgroup] - Mail-enabled Security Group Mirroring Unified Group Members
    # [Dummy Team (All)] - Unified Group (DisplayName)
    # [Shared Mailbox - Dummy Team (All)] - Shared Mailbox (for bodging DG membership)
    #$UnifiedGroupObject.CustomAttribute1 = Own ExternalDirectoryObjectId
    #$UnifiedGroupObject.anthesisgroup_UGSync.dataManagerGroupId = Data Managers Subgroup ExternalDirectoryObjectId
    #$UnifiedGroupObject.anthesisgroup_UGSync.memberGroupId = Members Subgroup ExternalDirectoryObjectId
    #$UnifiedGroupObject.anthesisgroup_UGSync.combinedGroupId = Combined Mail-Enabled Security Group ExternalDirectoryObjectId
    #$UnifiedGroupObject.anthesisgroup_UGSync.sharedMailboxId = Shared Mailbox ExternalDirectoryObjectId
    #$UnifiedGroupObject.anthesisgroup_UGSync.masterMembershipList = [string] "365"|"AAD" Is membership driven by the 365 Group or the associated AAD group?
    #$UnifiedGroupObject.anthesisgroup_UGSync.classification = [string] "Internal"|"External"|"Confidential" Intended Site Classification (used to reset in the event of unauthorised change)
    #$UnifiedGroupObject.anthesisgroup_UGSync.privacy = [string] "Public"|"Private" Intended Site Privacy (used to reset in the event of unauthorised change)
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true)]
            [string]$displayName
        ,[Parameter(Mandatory=$false)]
            [string]$description
        ,[Parameter(Mandatory=$true)]
            [string[]]$managerUpns
        ,[Parameter(Mandatory=$true)]
            [string[]]$teamMemberUpns
        ,[Parameter(Mandatory=$false)]
            [string[]]$memberOf
        ,[Parameter(Mandatory=$true)]
            [bool]$hideFromGal
        ,[Parameter(Mandatory=$true)]
            [bool]$blockExternalMail = $true
        ,[Parameter(Mandatory=$true)]
            [ValidateSet("Public", "Private")]
            [string]$accessType
        ,[Parameter(Mandatory=$true)]
            [bool]$autoSubscribe = $true
        ,[Parameter(Mandatory=$false)]
            [string[]]$additionalEmailAddresses
        ,[Parameter(Mandatory=$true)]
            [string]$groupClassification
        ,[Parameter(Mandatory=$false)]
            [string]$ownersAreRealManagers
        ,[Parameter(Mandatory=$true)]
        [ValidateSet("365", "AAD")]
            [string]$membershipManagedBy = "365"
        ,[Parameter(Mandatory=$true)]
            [PSCustomObject]$tokenResponse
        ,[Parameter(Mandatory=$false)]
            [bool]$alsoCreateTeam = $false
        ,[Parameter(Mandatory = $true)]
        [pscredential]$pnpCreds
        )

    Write-Verbose "new-365Group($displayName, $description, $managerUpns, $teamMemberUpns, $memberOf, $hideFromGal, $blockExternalMail, $isPublic, $autoSubscribe, $additionalEmailAddresses, $groupClassification, $ownersAreRealManagers,$membershipmanagedBy)"
    $displayName = $displayName.Replace("&","and") #Having ampersands in URLs causes problems, so the simplest option is to prevent them entirely.
    $shortName = $displayName.Replace(" (All)","")
    $365MailAlias = $(guess-aliasFromDisplayName "$displayName 365")
    $combinedSgDisplayName = $displayName
    $managersSgDisplayNameSuffix = " - Data Managers Subgroup"
    $managersSgDisplayName = "$displayName$managersSgDisplayNameSuffix"
    $membersSgDisplayNameSuffix = " - Members Subgroup"
    $membersSgDisplayName = "$displayName$membersSgDisplayNameSuffix"
    $sharedMailboxDisplayName = "Shared Mailbox - $displayName"

    #Firstly, check whether we have already created a Unified Group for this DisplayName
    $graphGroupExtended = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -filterDisplayName $(sanitise-forSql $displayName)
    #$365Group = Get-UnifiedGroup -Filter "DisplayName -eq `'$(sanitise-forSql $displayName)`'"
    if(!$graphGroupExtended){
        #$365Group = Get-UnifiedGroup -Filter "Alias -eq `'$365MailAlias`'" #If we can't find it by the DisplayName, check the Alias as this is less mutable
        $graphGroupExtended = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -filterUpn "$365MailAlias@anthesisgroup.com"
        } 

    #If we have a UG, check whether we can find the associated groups 
    if($graphGroupExtended){
        Write-Verbose "Pre-existing 365 Group found [$($graphGroupExtended.DisplayName)] with id=[$($graphGroupExtended.id)], dataManagerGroupId=[$($graphGroupExtended.anthesisgroup_UGSync.dataManagerGroupId)], memberGroupId=[$($graphGroupExtended.anthesisgroup_UGSync.memberGroupId)], combinedGroupId=[$($graphGroupExtended.anthesisgroup_UGSync.combinedGroupId)], sharedMailboxId=[$($graphGroupExtended.anthesisgroup_UGSync.sharedMailboxId)], masterMembershipList=[$($graphGroupExtended.anthesisgroup_UGSync.masterMembershipList)]"
        if(![string]::IsNullOrWhiteSpace($graphGroupExtended.anthesisgroup_UGSync.dataManagerGroupId)){
            #$managersSg = Get-DistributionGroup -Filter "ExternalDirectoryObjectId -eq `'$($graphGroupExtended.anthesisgroup_UGSync.dataManagerGroupId)`'"
            $managersSg = get-graphGroups -tokenResponse $tokenResponse -filterId $graphGroupExtended.anthesisgroup_UGSync.dataManagerGroupId 
            if(!$managersSg){
                Write-Warning "Data Managers Group [$($graphGroupExtended.anthesisgroup_UGSync.dataManagerGroupId)] for UG [$($graphGroupExtended.DisplayName)] could not be retrieved"
                $missingGroupId = $true
                }
            }
        else{Write-Warning "365 Group [$($graphGroupExtended.DisplayName)] found, but no anthesisgroup_UGSync.dataManagerGroupId (Data Managers Subgroup) property set!"}
        if(![string]::IsNullOrWhiteSpace($graphGroupExtended.anthesisgroup_UGSync.memberGroupId)){
            #$membersSg = Get-DistributionGroup -Filter "ExternalDirectoryObjectId -eq '$($graphGroupExtended.anthesisgroup_UGSync.memberGroupId)'"
            $membersSg = get-graphGroups -tokenResponse $tokenResponse -filterId $graphGroupExtended.anthesisgroup_UGSync.memberGroupId 
            if(!$membersSg){
                Write-Warning "Members Group [$($graphGroupExtended.anthesisgroup_UGSync.memberGroupId)] for UG [$($graphGroupExtended.DisplayName)] could not be retrieved"}
                $missingGroupId = $true
            }
        else{Write-Warning "365 Group [$($graphGroupExtended.DisplayName)] found, but no anthesisgroup_UGSync.memberGroupId (Members Subgroup) property set!"}
        if(![string]::IsNullOrWhiteSpace($graphGroupExtended.anthesisgroup_UGSync.combinedGroupId)){
            #$combinedSg = Get-DistributionGroup -Filter "ExternalDirectoryObjectId -eq '$($graphGroupExtended.anthesisgroup_UGSync.combinedGroupId)'"
            $combinedSg = get-graphGroups -tokenResponse $tokenResponse -filterId $graphGroupExtended.anthesisgroup_UGSync.combinedGroupId 
            if(!$combinedSg){
                Write-Warning "Combined Group [$($graphGroupExtended.anthesisgroup_UGSync.combinedGroupId)] for UG [$($graphGroupExtended.DisplayName)] could not be retrieved"
                $missingGroupId = $true
                }
            }
        else{Write-Warning "365 Group [$($graphGroupExtended.DisplayName)] found, but no anthesisgroup_UGSync.combinedGroupId (Combined Subgroup) property set!"}
        if(![string]::IsNullOrWhiteSpace($graphGroupExtended.anthesisgroup_UGSync.sharedMailboxId)){
            $sharedMailbox = Get-Mailbox -Filter "ExternalDirectoryObjectId -eq '$($graphGroupExtended.anthesisgroup_UGSync.sharedMailboxId)'"
            if(!$sharedMailbox){
                Write-Warning "Shared Mailbox [$($graphGroupExtended.anthesisgroup_UGSync.sharedMailboxId)] for UG [$($graphGroupExtended.DisplayName)] could not be retrieved"
                $missingGroupId = $true
                }
            }
        else{
            Write-Warning "365 Group [$($graphGroupExtended.DisplayName)] found, but no anthesisgroup_UGSync.sharedMailboxId (Shared Mailbox) property set!"
            $sharedMailboxDisplayName = "Shared Mailbox - $displayName"
            }
        if($missingGroupId -eq $true){
            repair-graphGroupUGSyncSchemaExtensions -tokenResponse $tokenResponse -graphGroup $graphGroupExtended -groupClassifcation $graphGroupExtended.anthesisgroup_UGSync.classification -masterMembership $graphGroupExtended.anthesisgroup_UGSync.masterMembershipList -createGroupsIfMissing -Verbose
            }
        }
    else{
        Write-Verbose "No pre-existing 365 group found - checking for AAD Groups."

        #Check whether any of these MESG exist based on names (just in case we're re-creating a 365 group and want to retain the AAD Groups, or something failed in the original creation and EXO hasn't syncronised with AAD yet)
        $combinedSg = get-graphGroups -tokenResponse $tokenResponse -filterDisplayName $combinedSgDisplayName
        if($combinedSg){Write-Verbose "Combined Group [$($combinedSg.DisplayName)] found"}
        else{
            $combinedSg = rummage-forDistributionGroup -displayName $combinedSgDisplayName
            if($combinedSg){Write-Verbose "Combined Group [$($combinedSg.DisplayName)] found"}
            else{Write-Warning "Combined Group [$($combinedSgDisplayName)] not found"}
            }
        $managersSg = get-graphGroups -tokenResponse $tokenResponse -filterDisplayName $managersSgDisplayName 
        if($managersSg){Write-Verbose "Managers Group [$($managersSg.DisplayName)] found"}
        else{
            $managersSg = rummage-forDistributionGroup -displayName $managersSgDisplayName
            if($managersSg){Write-Verbose "Managers Group [$($managersSg.DisplayName)] found"}
            else{Write-Warning "Managers Group [$($managersSgDisplayName)] not found"}
            }
        $membersSg  = get-graphGroups -tokenResponse $tokenResponse -filterDisplayName $membersSgDisplayName 
        if($membersSg){Write-Verbose "Members Group [$($membersSg.DisplayName)] found"}
        else{
            $membersSg = rummage-forDistributionGroup -displayName $membersSgDisplayName
            if($membersSg){Write-Verbose "Members Group [$($membersSg.DisplayName)] found"}
            else{Write-Warning "Members Group [$($membersSgDisplayName)] not found"}
            }
        $sharedMailbox = Get-Mailbox -Filter "DisplayName -eq `'$(sanitise-forSql $sharedMailboxDisplayName)`'"
        if(!$sharedMailbox){$sharedMailbox = Get-Mailbox -Filter "Alias -eq `'$(guess-aliasFromDisplayName $sharedMailboxDisplayName)`'"} #If we can't find it by the DisplayName, check the Alias as this is less mutable
        if($sharedMailbox){Write-Verbose "Shared Mailbox [$($sharedMailbox.DisplayName)] found"}else{Write-Verbose "Mailbox not found"}

        #Create any groups that don't already exist
        if(!$combinedSg){
            Write-Verbose "Creating Combined Security Group [$combinedSgDisplayName]"
            try{
                $combinedSg = new-mailEnabledSecurityGroup -dgDisplayName $combinedSgDisplayName -membersUpns $null -hideFromGal $false -blockExternalMail $true -ownersUpns "ITTeamAll@anthesisgroup.com" -description "Mail-enabled Security Group for $displayName" -WhatIf:$WhatIfPreference
                #Sadly, EXO is just too slow to replicate to Graph
                #$combinedSg = get-graphGroups -tokenResponse $tokenResponse -filterId $combinedSg.ExternalDirectoryObjectId -retryCount 5 #We can't create MESGs with Graph, but we can switch to Graph objects to simplify things later 
                }
            catch{Write-Error $_}
            }

        if($combinedSg -or $WhatIfPreference){ #If we now have a Combined SG
            if(!$managersSg){ #Create a Managers SG if required
                Write-Verbose "Creating Data Managers Security Group [$managersSgDisplayName]"
                $managersMemberOf = @($combinedSg.ExternalDirectoryObjectId)
                if($ownersAreRealManagers){$managersMemberOf += "Managers (All)"}
                try{
                    $managersSg = new-mailEnabledSecurityGroup -dgDisplayName $managersSgDisplayName -fixedSuffix $managersSgDisplayNameSuffix -membersUpns $managerUpns -memberOf $managersMemberOf -hideFromGal $false -blockExternalMail $true -ownersUpns "ITTeamAll@anthesisgroup.com" -description "Mail-enabled Security Group for $shortName Data Managers" -WhatIf:$WhatIfPreference -Verbose
                    #Sadly, EXO is just too slow to replicate to Graph
                    #$managersSg = get-graphGroups -tokenResponse $tokenResponse -filterId $managersSg.ExternalDirectoryObjectId -retryCount 5 #We can't create MESGs with Graph, but we can switch to Graph objects to simplify things later 
                    }
                catch{Write-Error $_}
                }

            if(!$membersSg){ #And create a Members SG if required
                Write-Verbose "Creating Members Security Group [$membersSgDisplayName]"
                try{
                    $membersSg = new-mailEnabledSecurityGroup -dgDisplayName $membersSgDisplayName -fixedSuffix $membersSgDisplayNameSuffix -membersUpns $teamMemberUpns -memberOf $combinedSg.ExternalDirectoryObjectId -hideFromGal $false -blockExternalMail $true -ownersUpns "ITTeamAll@anthesisgroup.com" -description "Mail-enabled Security Group for mirroring membership of $shortName Unified Group" -WhatIf:$WhatIfPreference
                    #Sadly, EXO is just too slow to replicate to Graph
                    #$membersSg = get-graphGroups -tokenResponse $tokenResponse -filterId $membersSg.ExternalDirectoryObjectId -retryCount 5 #We can't create MESGs with Graph, but we can switch to Graph objects to simplify things later 
                    if(![string]::IsNullOrWhiteSpace($memberOf)){
                        $memberOf | % { #We now nest membership via Members groups, rather than Combined Groups, so this is a little more complicated now.
                            $parentGroup = get-membersGroup -groupName $_
                            #Sadly, EXO is just too slow to replicate to Graph
                            $parentGroup = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -filterDisplayName $_
                            Add-DistributionGroupMember -Identity $parentGroup.ExternalDirectoryObjectId -BypassSecurityGroupManagerCheck -Member $membersSg.id -Confirm:$false
                            #Add-DistributionGroupMember -Identity $parentGroup.anthesisgroup_UGSync.memberGroupId -BypassSecurityGroupManagerCheck -Member $membersSg.id -Confirm:$false
                            }
                        }
                    }
                catch{Write-Error $_}
                }
            }
        else{Write-Error "Combined Security Group [$combinedSgDisplayName] not available. Cannot proceed with SubGroup creation"}        
        }

    #Check that everything's worked so far
    if(!$combinedSg){
        if($WhatIfPreference){Write-Verbose "Combined Security Group [$combinedSgDisplayName] not created because we're only pretending."}
        else{Write-Error "Combined Security Group [$combinedSgDisplayName] not available. Cannot proceed with UnifiedGroup creation";break}}
    if(!$managersSg){
        if($WhatIfPreference){Write-Verbose "Managers Security Group [$combinedSgDisplayName] not created because we're only pretending."}
        else{Write-Error "Managers Security Group [$managersSgDisplayName] not available. Cannot proceed with UnifiedGroup creation"; break}}
    if(!$membersSg){
        if($WhatIfPreference){Write-Verbose "Members Security Group [$combinedSgDisplayName] not created because we're only pretending."}
        else{Write-Error "Members Security Group [$membersSgDisplayName] not available. Cannot proceed with UnifiedGroup creation";break}}
    if(!$graphGroupExtended -or $WhatIfPreference){
        if(($combinedSg -and $managersSg -and $membersSg)){#If we now have all the prerequisite groups, create a UG
            try{
                $groupIsNew = $true
                Write-Verbose "All MESGs found - creating Unified 365 Group [$displayName]"
                if([string]::IsNullOrWhiteSpace($description)){$description = "Unified 365 Group for $displayName"}
                #Create the UG
                # Example of json for POST https://graph.microsoft.com/v1.0/groups
                # https://docs.microsoft.com/en-us/graph/api/group-post-groups?view=graph-rest-1.0
                [array]$owners = @()
                $managerUpns | % {[string[]]$owners += ("https://graph.microsoft.com/v1.0/users/$_").ToLower()}
                [array]$members = @()
                $teamMemberUpns | % {[string[]]$members += ("https://graph.microsoft.com/v1.0/users/$_").ToLower()}
                $members = $($members+$owners) | Sort-Object | Get-Unique -AsString 

                $ugSyncExtensionHash = @{
                    "extensionType" = "UGSync"
                    "dataManagerGroupId" = $managersSg.ExternalDirectoryObjectId 
                    "memberGroupId" = $membersSg.ExternalDirectoryObjectId
                    "combinedGroupId" = $combinedSg.ExternalDirectoryObjectId
                    #"sharedMailboxId" = $unifiedGroup.anthesisgroup_UGSync.sharedMailboxId
                    "masterMembershipList" = $membershipManagedBy
                    "classification" = $groupClassification
                    "privacy" = $accessType
                    }

                $groupHash = @{
                    "displayName"          = "$(sanitise-forJson $displayName)"
                    "groupTypes"           = @("Unified")
                    "mailNickname"         = $365MailAlias
                    "mailEnabled"          = $true
                    "securityEnabled"      = $true
                    "owners@odata.bind"    = $owners
                    "members@odata.bind"   = $members
                    "classification"       = $groupClassification
                    "visibility"           = $accessType
                    "anthesisgroup_UGSync" = $ugSyncExtensionHash
                    }

                $graphGroup = invoke-graphPost -tokenResponse $tokenResponse -graphQuery "/groups" -graphBodyHashtable $groupHash
                $graphGroupExtended = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -filterId $graphGroup.id 
                Write-Verbose $graphGroupExtended

                #Set Invite Guests for Unified Group (SharePoint Site sharing needs handling separately)
                set-graphUnifiedGroupGuestSettings -tokenResponse $tokenResponse -graphUnifiedGroupExtended $graphGroupExtended -classificationOverride $groupClassification | Out-Null #Set and forget

                }
            catch{Write-Error $_}
            }
        else{Write-Error "Combined/Managers/Members Security Group [$combinedSgDisplayName]/[$managersSgDisplayName]/[$membersSgDisplayName] not available. Cannot proceed with UnifiedGroup creation";break}
        }

    if($graphGroupExtended){ #If we now have a 365 UG, create a Shared Mailbox (if required) and configure it
        Write-Verbose ""
        if(!$sharedMailbox){
            Write-Verbose "Creating Shared Mailbox [$sharedMailboxDisplayName]: New-Mailbox -Shared -DisplayName $sharedMailboxDisplayName -Name $sharedMailboxDisplayName -Alias $(guess-aliasFromDisplayName -displayName $sharedMailboxDisplayName) -ErrorAction Continue -WhatIf:$WhatIfPreference "
            try{
            $sharedMailbox = New-Mailbox -Shared -DisplayName $sharedMailboxDisplayName -Name $sharedMailboxDisplayName -Alias $(guess-aliasFromDisplayName ($sharedMailboxDisplayName)) -ErrorAction Continue -WhatIf:$WhatIfPreference 
            Set-User -Identity $sharedMailbox.UserPrincipalName -AuthenticationPolicy "Allow IMAP"
            }
            catch{$Error}
            }

        if($sharedMailbox){
            Write-Verbose "Mailbox [$($sharedMailbox.DisplayName)][$($sharedMailbox.ExternalDirectoryObjectId)] found: Set-Mailbox -Identity $($sharedMailbox.ExternalDirectoryObjectId) -HiddenFromAddressListsEnabled $true -RequireSenderAuthenticationEnabled $false -ForwardingAddress $($graphGroupExtended.Mail) -DeliverToMailboxAndForward $true -ForwardingSmtpAddress $($graphGroupExtended.Mail) -Confirm:$false -WhatIf:$WhatIfPreference"
            Set-Mailbox -Identity $sharedMailbox.ExternalDirectoryObjectId -HiddenFromAddressListsEnabled $true -RequireSenderAuthenticationEnabled $false -Confirm:$false -WhatIf:$WhatIfPreference #-ForwardingAddress $graphGroupExtended.Mail -DeliverToMailboxAndForward $true  #I don't think we want to forward from the Shared Mailbox to the 365 group. If anything, we want the forwarding to work in reverse, and with the advent of Teams, Shared Mailboxes are liekly to be become less useful.
            Set-user -Identity $sharedMailbox.ExternalDirectoryObjectId -Manager groupbot -WhatIf:$WhatIfPreference  #For want of someone better....
            #Assign the Shared Mailbox as a member of the Security Group
            try{Add-DistributionGroupMember -Identity $combinedSg.ExternalDirectoryObjectId -Member $sharedMailbox.ExternalDirectoryObjectId -BypassSecurityGroupManagerCheck -WhatIf:$WhatIfPreference -ErrorAction Stop}
            catch{
                if('-2146233087' -eq $_.Exception.HResult){Write-Warning "Shared Mailbox [$($sharedMailbox.DisplayName)] is already a member of [$($combinedSg.DisplayName)]"}
                else{Write-Error $_}
                }
            set-graphGroupUGSyncSchemaExtensions -tokenResponse $tokenResponse -groupId $graphGroupExtended.id -sharedMailboxId $sharedMailbox.ExternalDirectoryObjectId  | Out-Null
            }
        else{Write-Error "Shared Mailbox not available. Cannot complete UG setup."}
        }
    else{Write-Error "Unified Group [$displayName] not available. Cannot proceed with Shared Mailbox creation."}

    if($groupIsNew){Write-Verbose "New 365 Group created: [$($graphGroupExtended.DisplayName)] with id=[$($graphGroupExtended.id)], dataManagerGroupId=[$($graphGroupExtended.anthesisgroup_UGSync.dataManagerGroupId)], memberGroupId=[$($graphGroupExtended.anthesisgroup_UGSync.memberGroupId)], combinedGroupId=[$($graphGroupExtended.anthesisgroup_UGSync.combinedGroupId)], sharedMailboxId=[$($graphGroupExtended.anthesisgroup_UGSync.sharedMailboxId)], masterMembershipList=[$($graphGroupExtended.anthesisgroup_UGSync.masterMembershipList)]"}
    elseif($graphGroupExtended){Write-Verbose "Pre-existing 365 Group found: [$($graphGroupExtended.DisplayName)]  with id=[$($graphGroupExtended.id)], dataManagerGroupId=[$($graphGroupExtended.anthesisgroup_UGSync.dataManagerGroupId)], memberGroupId=[$($graphGroupExtended.anthesisgroup_UGSync.memberGroupId)], combinedGroupId=[$($graphGroupExtended.anthesisgroup_UGSync.combinedGroupId)], sharedMailboxId=[$($graphGroupExtended.anthesisgroup_UGSync.sharedMailboxId)], masterMembershipList=[$($graphGroupExtended.anthesisgroup_UGSync.masterMembershipList)]"}
    else{Write-Verbose "It doesn't look like there's a [$displayName] 365 Group available..."}

    #Provision MS Team if requested
    if($alsoCreateTeam -and $graphGroupExtended){
        Write-Verbose "Provisioning new MS Team (as requested)"
        $graphTeam = new-graphTeam -tokenResponse $tokenResponse -groupId $graphGroupExtended.id -allowMemberCreateUpdateChannels $true -allowMemberDeleteChannels $false -Verbose:$VerbosePreference -ErrorAction Continue #Create the Team if it doesn't already exist
        if(!$graphTeam){write-warning "Failed to provision Team [$($graphGroupExtended.DisplayName)] via Graph after 3 attempts. Try again later."}
        }
    else{Write-Verbose "_NOT_ attempting to provision new MS Team"}


    $graphGroupExtended

    #Shifted to the end to minimise a race condition where delays in provisioning speed were causing failures.    
    Write-Host -f DarkYellow "`tCreate, share and delete a dummy folder in this Site to trigger the SharedWith Site Column (this can fail if there is a delay provisioning the Drive)"
    do{
        $i++
        Start-Sleep -Seconds 5
        Write-Verbose "Drive not available. Retry in 5 seconds. ($i/50)"
        try{$graphDrive = get-graphDrives -tokenResponse $tokenResponse -groupGraphId $graphGroupExtended.id -returnOnlyDefaultDocumentsLibrary}# -Verbose:$VerbosePreference}
        catch{if($_.Exception -match "Couldn't find object" -or $_.Exception -match "Resource provisioning is in progress"){<#Do nothing - object not provisioned yet#>}}
        if($i -eq 50){break}
        }
    while($graphDrive -eq $null)
    $dummyFolder = add-graphArrayOfFoldersToDrive -graphDriveId $graphDrive.id -foldersAndSubfoldersArray "DummyFolder" -tokenResponse $tokenResponse -conflictResolution Replace
    grant-graphSharing -tokenResponse $tokenResponse -driveId $graphDrive.id -itemId $dummyFolder.id -sharingRecipientsUpns @($managerUpns[0]) -requireSignIn $true -sendInvitation $false -role Write -Verbose | Out-Null
    delete-graphDriveItem -tokenResponse $tokenResponse -graphDriveId $graphDrive.id -graphDriveItemId $dummyFolder.id -eTag $dummyFolder.eTag | Out-Null

    }
function new-aliasFromDisplayName(){
    [CmdletBinding()]
    Param (
        [parameter(Mandatory = $true)]
        [string]$displayName
        ,[parameter(Mandatory = $false)]
        [string]$fixedSuffix
        )
    #Write-Host -ForegroundColor Magenta "guess-aliasFromDisplayName($displayName)"
    $newGuid = [GUID]::NewGuid()
    if(![string]::IsNullOrWhiteSpace($displayName)){$guessedAlias = $displayName.replace(" ","_").Replace("(","").Replace(")","").Replace(",","").Replace("@","").Replace("\","").Replace("[","").Replace("]","").Replace("`"","").Replace(";","").Replace(":","").Replace("<","").Replace(">","")}
    $newSuffix =  set-suffixAndMaxLength -string $newGuid.Guid -suffix $fixedSuffix -maxLength 64
    $newAlias = set-suffixAndMaxLength -string $displayName -suffix $newSuffix -maxLength 64
    $newAlias = sanitise-forMicrosoftEmailAddress -dirtyString $newAlias
    $newAlias = remove-diacritics -String $newAlias
    Write-Verbose -Message "new-aliasFromDisplayName($displayName) = [$newAlias]"
    $newAlias
    }
function new-mailEnabledSecurityGroup(){
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true)]
            [string]$dgDisplayName
        ,[Parameter(Mandatory=$false)]
            [string]$fixedSuffix
        ,[Parameter(Mandatory=$false)]
            [string]$description
        ,[Parameter(Mandatory=$false)]
            [string[]]$ownersUpns
        ,[Parameter(Mandatory=$false)]
            [string[]]$membersUpns
        ,[Parameter(Mandatory=$false)]
            [string[]]$memberOf
        ,[Parameter(Mandatory=$true)]
            [bool]$hideFromGal
        ,[Parameter(Mandatory=$true)]
            [bool]$blockExternalMail
        )
    Write-Verbose "new-mailEnabledSecurityGroup([$dgDisplayName], [$description], [$($ownersUpns -join ", ")], [$($membersUpns -join ", ")], [$($memberOf -join ", ")], $hideFromGal, $blockExternalMail)"
    #$mailName = set-suffixAndMaxLength -string $dgDisplayName -suffix $fixedSuffix -maxLength 64 -Verbose
    #write-verbose "#####Mailname: $($mailName)"

    #Check to see if this already exists. This is based on Alias, which is mutable :(    
    $mesg = rummage-forDistributionGroup -displayName $dgDisplayName
    if($mesg){ #If the group already exists, add the new Members (ignore any removals - we'll let sync-groupMembership figure that out)
        $members  | % {
            Write-Verbose "Adding TeamMember Add-DistributionGroupMember $($mesg.ExternalDirectoryObjectId) -Member $_ -Confirm:$false -BypassSecurityGroupManagerCheck"
            Add-DistributionGroupMember $mesg.ExternalDirectoryObjectId -Member $_ -Confirm:$false -BypassSecurityGroupManagerCheck -WhatIf:$WhatIfPreference
            }
        }
    else{ #If the group doesn't exist, try creating it
        try{
            $mailAlias = $(new-aliasFromDisplayName $dgDisplayName)
            Write-Verbose "New-DistributionGroup -Name [$mailAlias] -DisplayName [$dgDisplayName] -Type Security -Members [$($membersUpns -join ", ")] -PrimarySmtpAddress $($(sanitise-forMicrosoftEmailAddress -dirtyString $(set-suffixAndMaxLength -string $dgDisplayName -suffix $fixedSuffix -maxLength 100))+"@anthesisgroup.com") -Notes [$description] -Alias [$mailAlias] -WhatIf:$WhatIfPreference"
            $mesg = New-DistributionGroup -Name $mailAlias -DisplayName $dgDisplayName -Type Security -Members $membersUpns -PrimarySmtpAddress $($(sanitise-forMicrosoftEmailAddress -dirtyString $(set-suffixAndMaxLength -string $dgDisplayName -suffix $fixedSuffix -maxLength 100))+"@anthesisgroup.com") -Notes $description -Alias $mailAlias -WhatIf:$WhatIfPreference -ErrorAction Stop
            }
        catch{
            if($_ -match "is already being used by the proxy addresses or LegacyExchangeDN of"){ #Name collision, but no DisplayName collision
                #Create the DG with a temporary Guid in the Name/Alias to eliminate the collision
                $tempGuid = $([guid]::NewGuid().Guid)
                $tempMailName = set-suffixAndMaxLength -string $dgDisplayName -suffix $tempGuid -maxLength 64 
                $tempMailAlias = guess-aliasFromDisplayName -displayName $dgDisplayName -fixedSuffix $tempGuid
                Write-Verbose "`t2nd attempt: New-DistributionGroup -Name [$tempMailName] -DisplayName [$dgDisplayName] -Type Security -Members [$($membersUpns -join ", ")] -PrimarySmtpAddress $($tempMailAlias+"@anthesisgroup.com") -Notes [$description] -Alias [$tempMailAlias] -WhatIf:$WhatIfPreference"
                $mesg = New-DistributionGroup -Name $tempMailName -DisplayName $dgDisplayName -Type Security -Members $membersUpns -PrimarySmtpAddress $($(guess-aliasFromDisplayName -displayName $dgDisplayName -fixedSuffix $tempGuid)+"@anthesisgroup.com") -Notes $description -Alias $tempMailAlias -WhatIf:$WhatIfPreference
                #Then use the ExternalDirectoryObjectId property to re-set the Name and Alias properties to a "useful" Guid
                $newMailName = set-suffixAndMaxLength -string $dgDisplayName -suffix $mesg.ExternalDirectoryObjectId -maxLength 64
                $newmailAlias = guess-aliasFromDisplayName -displayName $dgDisplayName -fixedSuffix $mesg.ExternalDirectoryObjectId
                $mesg | Set-DistributionGroup -Name $newMailName -Alias $newmailAlias -PrimarySmtpAddress $($newmailAlias+"@anthesisgroup.com")
                $mesg = Get-DistributionGroup -Identity $mesg.ExternalDirectoryObjectId
                }
            else{
                Write-Error "Error creating new Distribution Group [$($dgDisplayName)] in new-mailEnabledSecurityGroup()"
                get-errorSummary -errorToSummarise $_
                }
            }
        }

    if(!$mesg){Write-Error "Mail-Enabled Security Group [$dgDisplayName] neither found, nor created :/"}
    else{ #Now set the additional properties and MemberOf
        Write-Verbose "Set-DistributionGroup -Identity $mailAlias -HiddenFromAddressListsEnabled $hideFromGal -RequireSenderAuthenticationEnabled $blockExternalMail -ManagedBy [$($ownersUpns -join ", ")]"
        Set-DistributionGroup -Identity $mesg.ExternalDirectoryObjectId -HiddenFromAddressListsEnabled $hideFromGal -RequireSenderAuthenticationEnabled $blockExternalMail -ManagedBy $ownersUpns -WhatIf:$WhatIfPreference
        $memberOf | % {
            if(![string]::IsNullOrWhiteSpace($_)){
                Write-Verbose "Adding As MemberOf Add-DistributionGroupMember [$_] -Member [$($mesg.ExternalDirectoryObjectId)] -Confirm:$false -BypassSecurityGroupManagerCheck"
                Add-DistributionGroupMember -Identity $_ -Member $mesg.ExternalDirectoryObjectId -Confirm:$false -BypassSecurityGroupManagerCheck -WhatIf:$WhatIfPreference
                }
            }
        }
    $mesg
    }
function new-symGroup($displayName, $description, $managers, $teamMembers, $memberOf, $additionalEmailAddress){
    Write-Host -ForegroundColor Magenta "new-symGroup($displayName, $description, $managers, $teamMembers, $memberOf, $additionalEmailAddress)"
    $hideFromGal = $false
    $blockExternalMail = $true
    $isPublic = $true 
    $autoSubscribe = $true
    $groupClassification = "Internal"
    new-365Group -displayName $displayName -description $description -managerUpns $managers -teamMemberUpns $teamMembers -memberOf $memberOf -hideFromGal $hideFromGal -blockExternalMail $blockExternalMail -isPublic $isPublic -autoSubscribe $autoSubscribe -additionalEmailAddresses $additionalEmailAddress -groupClassification $groupClassification -ownersAreRealManagers $false
    }
function remove-DataManagerFromGroup(){
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true)]
            [string]$dataManagerGroupId
        ,[Parameter(Mandatory=$true)]
            [string]$upnToRemove
        )
    Write-Verbose "remove-DataManagerFromGroup([$upnToRemove],[$dataManagerGroupId])"
    $dataManagerGroupMembers = Get-DistributionGroupMember -Identity $dataManagerGroupId

    if($dataManagerGroupMembers.WindowsLiveID -notcontains $upnToRemove){
        Write-Warning "remove-DataManagerFromGroup could not remove [$upnToRemove] from [$dataManagerGroupId] because it is not a member"
        return
        }

    ,[array]$otherDataManagers = $dataManagerGroupMembers | ? {$_.WindowsLiveID -ne $upnToRemove}
    if($otherDataManagers.Count -eq 0){ #If this user is the last Data Manager, add GroupBot to prevent this Data Manager group from becoming empty
        Add-DistributionGroupMember -Identity $dataManagerGroupId -Member groupbot@anthesisgroup.com -BypassSecurityGroupManagerCheck -Confirm:$false
        }

    Remove-DistributionGroupMember -Identity $dataManagerGroupId -Member $upnToRemove -BypassSecurityGroupManagerCheck -Confirm:$false 
    }
function rummage-forDistributionGroup(){
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
            [string]$displayName
        ,[Parameter(Mandatory=$false)]
            [string]$alias
        )

    Write-Verbose "rummage-forDistributionGroup([$displayName],[$alias])"
    [array]$dg = Get-DistributionGroup -Filter "DisplayName -eq `'$(sanitise-forSql $displayName)`'"
    if($dg.Count -ne 1){
#        if($alias){
#            Write-Verbose "Trying to get DG by alias [$alias]"
#            [array]$dg = Get-DistributionGroup -Filter "Alias -eq `'$alias`'" #If we can't find it by the DisplayName, check the Alias as this is less mutable
#            }
        #if($dg.Count -ne 1){
            if($dg.Count -gt 1){Write-Warning "Multiple Groups matched for Distribution Group [$displayName]`r`n`t $($dg.PrimarySmtpAddress -join "`r`n`t")"}
            if($dg.Count -eq 0){Write-Verbose "No Distribution Group found"}
            $dg = $null
        #    }
        } 
    $dg
    }
function send-dataManagerReassignmentRequest(){
    [CmdletBinding(SupportsShouldProcess=$true )]
    param(
        [Parameter(Mandatory = $true)]
            [psobject]$tokenResponse
        ,[Parameter(Mandatory=$true)]
            [psobject]$unifiedGroup
        ,[Parameter(Mandatory=$false)]
            [array]$currentOwners
        ,[Parameter(Mandatory=$false)]
            [string[]]$adminEmailAddresses
        )

    if([string]::IsNullOrWhiteSpace($currentOwners.manager)){
        #If _none_ of the Ex-Employees have Line Managers, send an alert to the IT Admins to sort this Team/Site out manually
        send-noOwnersForGroupAlertToAdmins -tokenResponse $tokenResponse -UnifiedGroup $unifiedGroup -currentOwners $currentOwners -adminEmailAddresses $adminEmailAddresses
        return
        }
    else{
        $authorisedDataManagers = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId 'daf56fbd-ebce-457e-a10a-4fce50a2f99c' -memberType Members -returnOnlyLicensedUsers
        $otherMembers = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $unifiedGroup.id -memberType Members -returnOnlyLicensedUsers
        $potentialDataManagers = Compare-Object -ReferenceObject $authorisedDataManagers -DifferenceObject $otherMembers -Property userPrincipalName -ExcludeDifferent -IncludeEqual -PassThru
        $groupSite = get-graphSite -tokenResponse $tokenResponse -groupId $UnifiedGroup.id
    }


    $currentOwners | ForEach-Object {
        if([string]::IsNullOrWhiteSpace($_.manager)){<#Do nothing - leave it to one of the Ex-Employees who does have a Line Manager assigned#>}
        else{
            $subject = "Unmanaged Site/Team Group found: [$($UnifiedGroup.DisplayName)]"
            $body = "<HTML><FONT FACE=`"Calibri`">Hello $($_.manager.givenName),`r`n`r`n<BR><BR>"
            $body += "The 365 Site/Team [<A HREF=`"$($groupSite.webUrl)`"><B>$($UnifiedGroup.DisplayName)</B></A>] is currently managed by one of your ex-reportees:`r`n`t<BR>"
            $body += "<PRE>&#9;$($_.displayName)</PRE>`r`n`r`n<BR>"
            $body += "Please could you let the <A HREF='mailto:IT_Team_GBR@anthesisgroup.com'>IT Team</A> know who to reassign this to?`r`n`r`n<BR><BR>"
            if($potentialDataManagers.Count -gt 0){
                $body += "These Members have already completed Data Manager training:`r`n`t<BR><PRE>&#9;$($potentialDataManagers.DisplayName -join "`r`n`t")</PRE>`r`n`r`n<BR>"

            }
            else{
                    $body+= "Unfortunately, no other Members of the Team currently have Data Manager training. You can find a list of everyone with current Data Manager training by <A HREF='https://anthesisllc.sharepoint.com/sites/Resources-IT/_layouts/15/DocIdRedir.aspx?ID=HXX7CE52TSD2-1759992947-296'>expanding the <B>Data Manager - Authorised (All)</B> group in Outlook</A>.`r`n`r`n<BR><BR>"
            }
            $body += "<B>If you can tell the IT Team who to reassign this to, you will stop receiving these emails</B>.`r`n`r`n<BR><BR>"
            $body += "Love,`r`n`r`n<BR><BR>The Helpful Groups Robot</FONT></HTML>"
            send-graphMailMessage -tokenResponse $tokenResponse -fromUpn "groupbot@anthesisgroup.com" -toAddresses $_.manager.mail -bccAddresses "t0-kevin.maitland@anthesisgroup.com" -subject $subject -bodyHtml $body
            #send-graphMailMessage -tokenResponse $tokenResponse -fromUpn "groupbot@anthesisgroup.com" -toAddresses "kevin.maitland@anthesisgroup.com" -subject $subject -bodyHtml $body
            }
        }

    
    }
function send-membershipChangeReportToManagers(){
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory = $true)]
            [psobject]$tokenResponse
        ,[Parameter(Mandatory=$true, Position=0)]
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
        Write-verbose "To [$($ownersEmailAddresses -join "; ")]"
        Write-verbose "CC [$($adminEmailAddresses -join "; ")]"
        send-graphMailMessage -tokenResponse $tokenResponse -fromUpn "groupbot@anthesisgroup.com" -toAddresses $ownersEmailAddresses -bccAddresses "t0-kevin.maitland@anthesisgroup.com" -subject $subject -bodyHtml $body
        #send-graphMailMessage -tokenResponse $tokenResponse -fromUpn "groupbot@anthesisgroup.com" -toAddresses $ownersEmailAddresses -ccAddresses $adminEmailAddresses -subject $subject -bodyHtml $body
        #Send-MailMessage -To $ownersEmailAddresses -From "thehelpfulgroupsrobot@anthesisgroup.com" -cc $adminEmailAddresses -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
        }
    else{
        send-graphMailMessage -tokenResponse $tokenResponse -fromUpn "groupbot@anthesisgroup.com" -toAddresses "kevin.maitland@anthesisgroup.com" -ccAddresses $adminEmailAddresses -subject $subject -bodyHtml $body
        #Send-MailMessage -To "kevin.maitland@anthesisgroup.com" -From "thehelpfulgroupsrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
        }

    }
function send-membershipChangeProblemReportToAdmins(){
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory = $true)]
            [psobject]$tokenResponse
        ,[Parameter(Mandatory=$true, Position=0)]
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
        send-graphMailMessage -tokenResponse $tokenResponse -fromUpn "groupbot@anthesisgroup.com" -toAddresses $adminEmailAddresses -subject $subject -bodyHtml $body
        #Send-MailMessage -To $adminEmailAddresses -From "thehelpfulgroupsrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
        }
    else{
        send-graphMailMessage -tokenResponse $tokenResponse -fromUpn "groupbot@anthesisgroup.com" -toAddresses "kevin.maitland@anthesisgroup.com" -subject $subject -bodyHtml $body
        #Send-MailMessage -To "kevin.maitland@anthesisgroup.com" -From "thehelpfulgroupsrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
        }

    }
function send-noOwnersForGroupAlertToAdmins(){
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory = $true)]
            [psobject]$tokenResponse
        ,[Parameter(Mandatory=$true)]
            [psobject]$UnifiedGroup
        ,[Parameter(Mandatory=$false)]
            [array]$currentOwners
        ,[Parameter(Mandatory=$false)]
            [string[]]$adminEmailAddresses
        )

    $groupSite = get-graphSite -tokenResponse $tokenResponse -groupId $UnifiedGroup.id
    $groupMembers = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $UnifiedGroup.id -memberType Members -returnOnlyLicensedUsers -includeLineManager
    $subject = "Unowned 365 Group found: [$($UnifiedGroup.DisplayName)]"
    $body = "<HTML><FONT FACE=`"Calibri`">Hello 365 Group Admins,`r`n`r`n<BR><BR>"
    $body += "There are no active owners for [<B>$($UnifiedGroup.DisplayName)</B>][$($UnifiedGroup.id)][<A HREF=`"$($groupSite.webUrl)`">Site URL</A>][<A HREF=`"https://admin.microsoft.com/AdminPortal/Home#/groups/:/TeamDetails/$($UnifiedGroup.id)`">365</A>][<A HREF=`"https://portal.azure.com/#blade/Microsoft_AAD_IAM/GroupDetailsMenuBlade/Overview/groupId/$($UnifiedGroup.id)`">AAD<A/>][<A HREF=`"https://portal.azure.com/#blade/Microsoft_AAD_IAM/GroupDetailsMenuBlade/Overview/groupId/$($UnifiedGroup.anthesisgroup_UGSync.dataManagerGroupId)`">Data Manager Group</A>] `r`n`r`n<BR><BR>"
    
    if($currentOwners.Count -gt 0){
        $currentOwners = $currentOwners | Sort-Object DisplayName
        $body += "The full list of 365 group Owners looks like this:`r`n`t<BR><PRE>&#9;$($currentOwners.DisplayName -join "`r`n`t")</PRE>`r`n`r`n<BR>"
        $body += "<B>These owners either have no Line Manager, or their Line Manager has also been deactivated!</B>`r`n`r`n<BR><BR>"
    }
    else{$body += "It looks like the Owners group is now empty...`r`n`r`n<BR><BR>"}
    
    $body += "<B>This can only be fixed manually!</B>`r`n`r`n<BR><BR>"
    
    if($groupMembers.Count -gt 0){
        $groupMembers = $groupMembers | Sort-Object DisplayName
        $body += "The remaining members of the group, or their [Line Manager]s might be able to help:`r`n`t<BR><PRE>&#9;"
        $groupMembers | ForEach-Object {
            $body += "$($_.userPrincipalName)`t[$($_.manager.userPrincipalName)]`r`n`t"
        }
    }
    $body += "</PRE>`r`n`r`n<BR>Love,`r`n`r`n<BR><BR>The Helpful Groups Robot</FONT></HTML>"    

    if([string]::IsNullOrWhiteSpace($adminEmailAddresses)){$adminEmailAddresses = get-groupAdminRoleEmailAddresses}

    if($PSCmdlet.ShouldProcess($("[$($UnifiedGroup.DisplayName)]"))){#Fudges -WhatIf as it's not suppoerted natively by Send-MailMessage
        send-graphMailMessage -tokenResponse $tokenResponse -fromUpn groupbot@anthesisgroup.com -toAddresses $adminEmailAddresses -subject $subject -bodyHtml $body -priority high
        #Send-MailMessage -To $adminEmailAddresses -From "thehelpfulgroupsrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 -Priority High
        }
    else{
        send-graphMailMessage -tokenResponse $tokenResponse -fromUpn groupbot@anthesisgroup.com -toAddresses "kevin.maitland@anthesisgroup.com" -subject $subject -bodyHtml $body -priority high
        #Send-MailMessage -To "kevin.maitland@anthesisgroup.com" -From "thehelpfulgroupsrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 -Priority High
        }
    
    }
function set-guestAccessForUnifiedGroup(){
    param(
        [Parameter(Mandatory=$true)]
        [PSObject]$unifiedGroup
        )
    Write-Verbose "set-guestAccessForUnifiedGroup ([$($unifiedGroup.DisplayName)])"
    switch($unifiedGroup.CustomAttribute7){
        "External" {
            #Allow external sharing
            $allowToAddGuests = $true
            }
        "Internal" {
            #Block all external sharing
            $allowToAddGuests = $false
            }
        "Confidential" {
            #Block all external sharing
            $allowToAddGuests = $false
            }
        default {
            $allowToAddGuests = $false
            }
        }
    $preExistingSettings = Get-AzureADObjectSetting -TargetType Groups -TargetObjectId $unifiedGroup.ExternalDirectoryObjectId
    $template = Get-AzureADDirectorySettingTemplate | ? {$_.displayname -eq "group.unified.guest"}
    $settingsCopy = $template.CreateDirectorySetting()
    $settingsCopy["AllowToAddGuests"]=$allowToAddGuests
    
    if($preExistingSettings){
        Write-Verbose "Set-AzureADObjectSetting -TargetType Groups -TargetObjectId $($unifiedGroup.ExternalDirectoryObjectId) -DirectorySetting $settingsCopy"
        Set-AzureADObjectSetting -TargetType Groups -TargetObjectId $unifiedGroup.ExternalDirectoryObjectId -DirectorySetting $settingsCopy -Id $preExistingSettings.Id
        }
    else{
        Write-Verbose "New-AzureADObjectSetting -TargetType Groups -TargetObjectId $($unifiedGroup.ExternalDirectoryObjectId) -DirectorySetting $settingsCopy"
        New-AzureADObjectSetting -TargetType Groups -TargetObjectId $unifiedGroup.ExternalDirectoryObjectId -DirectorySetting $settingsCopy
        }
    }
function set-unifiedGroupCustomAttributes(){
    param(
        [Parameter(Mandatory=$true)]
            [PSObject]$unifiedGroup
        ,[Parameter(Mandatory=$true)]
            [ValidateSet ("Internal","Confidential","External","Sym")]
            [string]$groupType
        ,[Parameter(Mandatory=$true)]
            [ValidateSet ("AAD","365")]
            [string]$masterMembership
        ,[Parameter(Mandatory=$false)]
            [switch]$createGroupsIfMissing
        )

    $sgs = Get-AzureADGroup -SearchString $unifiedGroup.DisplayName

    $dataManagerSG = @()
    $membersSG = @()
    $combinedSG = @()
    $smb = @()

    $dataManagerSG += $sgs | ? {$_.DisplayName -match "data managers"}
    $membersSG += $sgs | ? {$_.DisplayName -match "members"}
    $combinedSG += $sgs | ? {$_.DisplayName -eq $unifiedGroup.DisplayName -and $_.ObjectId -ne $unifiedGroup.ExternalDirectoryObjectId}
    $smb += Get-Mailbox -Filter "DisplayName -like `'*$unifiedGroup.DisplayName*`'"

    switch($groupType){
        "Internal" {
            $pubPriv = "Private"
            $inEx = "Internal"
            }
        "Confidential" {
            $pubPriv = "Private"
            $inEx = "Internal"
            }
        "External" {
            $pubPriv = "Private"
            $inEx = "External"
            }
        "Sym" {
            $pubPriv = "Public"
            $inEx = "Internal"
            }
        }

    if(!$unifiedGroup){
        Write-Warning "Unified Group not found - cannot continue"
        break
        }
    if($dataManagerSG.Count -ne 1){
        Write-Warning "[$($dataManagerSG.Count)] Potential Data Manager groups identified [$($dataManagerSG.DisplayName -join ",")]. Cannot automatically resolve this problem."
        $bigProblem = $true
        }
    if($membersSG.Count -ne 1){
        Write-Warning "[$($membersSG.Count)] Potential Member groups identified [$($membersSG.DisplayName -join ",")]. Cannot automatically resolve this problem."
        $bigProblem = $true
        }
    if($combinedSG.Count -ne 1){
        Write-Warning "[$($combinedSG.Count)] Potential Combined groups identified [$($combinedSG.DisplayName -join ",")]. Cannot automatically resolve this problem."
        $bigProblem = $true
        }
    if($smb.Count -ne 1){
        Write-Warning "[$($smb.Count)] Potential Shared Mailboxes identified [$($smb.DisplayName -join ",")]. Cannot automatically resolve this problem."
        }

    if($bigProblem){
        if($createGroupsIfMissing){
            Write-Warning "Couldn't automatically identify the required groups to fix this. Will attempt to create missing groups"
            if(!$combinedSG){
                Write-Verbose "`tCreating Combined Security Group [$($unifiedGroup.DisplayName)]"
                try{
                    $combinedSg = new-mailEnabledSecurityGroup -dgDisplayName $unifiedGroup.DisplayName -membersUpns $null -hideFromGal $false -blockExternalMail $true -ownersUpns "ITTeamAll@anthesisgroup.com" -description "Mail-enabled Security Group for $($unifiedGroup.DisplayName)" -WhatIf:$WhatIfPreference
                    }
                catch{Write-Error $_}
                }
            if($combinedSG){#Dont try creating the subgroups if the Combined Group isn't available
                Set-UnifiedGroup -Identity $unifiedGroup.ExternalDirectoryObjectId -CustomAttribute4 $combinedSg.ExternalDirectoryObjectId
                if(!$dataManagerSG){ #Create a Managers SG if required
                    Write-Verbose "Creating Data Managers Security Group [$($unifiedGroup.DisplayName) - Data Managers Subgroup]"
                    try{$dataManagerSG = new-mailEnabledSecurityGroup -dgDisplayName "$($unifiedGroup.DisplayName) - Data Managers Subgroup" -fixedSuffix " - Data Managers Subgroup" -membersUpns $null -memberOf @($combinedSg.ExternalDirectoryObjectId,$combinedSG[0].ObjectId)-hideFromGal $false -blockExternalMail $true -ownersUpns "ITTeamAll@anthesisgroup.com" -description "Mail-enabled Security Group for $($unifiedGroup.DisplayName) Data Managers" -WhatIf:$WhatIfPreference -Verbose}
                    catch{Write-Error $_}
                    }
                if($dataManagerSG){Set-UnifiedGroup -Identity $unifiedGroup.ExternalDirectoryObjectId -CustomAttribute2 $dataManagerSG.ExternalDirectoryObjectId}

                if(!$membersSg){ #And create a Members SG if required
                    Write-Verbose "Creating Members Security Group [$($unifiedGroup.DisplayName) - Members Subgroup]"
                    try{$membersSg = new-mailEnabledSecurityGroup -dgDisplayName "$($unifiedGroup.DisplayName) - Members Subgroup" -fixedSuffix " - Members Subgroup" -membersUpns $null -memberOf @($combinedSg.ExternalDirectoryObjectId,$combinedSG[0].ObjectId) -hideFromGal $false -blockExternalMail $true -ownersUpns "ITTeamAll@anthesisgroup.com" -description "Mail-enabled Security Group for mirroring membership of $($unifiedGroup.DisplayName) Unified Group" -WhatIf:$WhatIfPreference -Verbose}
                    catch{Write-Error $_}
                    }
                if($membersSG){Set-UnifiedGroup -Identity $unifiedGroup.ExternalDirectoryObjectId -CustomAttribute3 $membersSg.ExternalDirectoryObjectId}
                
                }
            break
            }
        else{
            Write-Warning "Couldn't automatically identify the required groups to fix this. Will attempt to set remaining CustomAttributes then exit"
            Set-UnifiedGroup -Identity $unifiedGroup.ExternalDirectoryObjectId -CustomAttribute6 $masterMembership -CustomAttribute7 $groupType -CustomAttribute8 $pubPriv
            break
            }
        }

    if($smb.Count -eq 1){
        Write-Verbose "Set-UnifiedGroup -Identity [$($unifiedGroup.ExternalDirectoryObjectId)] -CustomAttribute1 [$($unifiedGroup.ExternalDirectoryObjectId)] -CustomAttribute2 [$($dataManagerSG[0].ObjectId)] -CustomAttribute3)] [$($membersSG[0].ObjectId)] -CustomAttribute4 [$($combinedSG[0].ObjectId)] -CustomAttribute5 [$($smb.ExternalDirectoryObjectId)] -CustomAttribute6 [$($pubPriv)] -CustomAttribute7 [$($groupType)] -CustomAttribute8 [$($masterMembership)]"
        Set-UnifiedGroup -Identity $unifiedGroup.ExternalDirectoryObjectId -CustomAttribute1 $unifiedGroup.ExternalDirectoryObjectId -CustomAttribute2 $dataManagerSG[0].ObjectId -CustomAttribute3 $membersSG[0].ObjectId -CustomAttribute4 $combinedSG[0].ObjectId -CustomAttribute5 $smb.ExternalDirectoryObjectId -CustomAttribute6 $masterMembership -CustomAttribute7 $groupType -CustomAttribute8 $pubPriv
        }
    else{
        Write-Verbose "Set-UnifiedGroup -Identity [$($unifiedGroup.ExternalDirectoryObjectId)] -CustomAttribute1 [$($unifiedGroup.ExternalDirectoryObjectId)] -CustomAttribute2 [$($dataManagerSG[0].ObjectId)] -CustomAttribute3)] [$($membersSG[0].ObjectId)] -CustomAttribute4 [$($combinedSG[0].ObjectId)] -CustomAttribute6 [$($pubPriv)] -CustomAttribute7 [$($groupType)] -CustomAttribute8 [$($masterMembership)]"
        Set-UnifiedGroup -Identity $unifiedGroup.ExternalDirectoryObjectId -CustomAttribute1 $unifiedGroup.ExternalDirectoryObjectId -CustomAttribute2 $dataManagerSG[0].ObjectId -CustomAttribute3 $membersSG[0].ObjectId -CustomAttribute4 $combinedSG[0].ObjectId -CustomAttribute6 $masterMembership -CustomAttribute7 $groupType -CustomAttribute8 $pubPriv
        }
    }
function sync-groupMemberships(){
    [CmdletBinding(SupportsShouldProcess=$true )]
    param(
        [Parameter(Mandatory = $true)]
            [psobject]$tokenResponseSmtp
        ,[Parameter(Mandatory=$true,ParameterSetName="365GroupObjectSupplied")]
            [Parameter(Mandatory=$false,ParameterSetName="AADGroupObjectSupplied")]
            [PSObject]$graphExtendedUG
        ,[Parameter(Mandatory=$false,ParameterSetName="365GroupObjectSupplied")]
            [Parameter(Mandatory=$true,ParameterSetName="AADGroupObjectSupplied")]
            [PSObject]$graphMesg
        ,[Parameter(Mandatory=$true,ParameterSetName="365GroupIdOnly")]
            [Parameter(Mandatory=$false,ParameterSetName="AADGroupIdOnly")]
            [string]$graphExtendedUGUpn
        ,[Parameter(Mandatory=$false,ParameterSetName="365GroupIdOnly")]
            [Parameter(Mandatory=$true,ParameterSetName="AADGroupIdOnly")]
            [string]$graphMesgUpn
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
            [psobject]$tokenResponse
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
            [string[]]$adminEmailAddresses
        ,[Parameter(Mandatory=$false,ParameterSetName="365GroupObjectSupplied")]
            [Parameter(Mandatory=$false,ParameterSetName="AADGroupObjectSupplied")]
            [Parameter(Mandatory=$false,ParameterSetName="365GroupIdOnly")]
            [Parameter(Mandatory=$false,ParameterSetName="AADGroupIdOnly")]
            [bool]$enumerateSubgroups = $false
        ,[Parameter(Mandatory=$false,ParameterSetName="365GroupObjectSupplied")]
            [Parameter(Mandatory=$false,ParameterSetName="AADGroupObjectSupplied")]
            [Parameter(Mandatory=$false,ParameterSetName="365GroupIdOnly")]
            [Parameter(Mandatory=$false,ParameterSetName="AADGroupIdOnly")]
            [psobject]$syncException
        )

    #region Get $graphExtendedUG and $graphMesg, regardless of which parameters we've been given
    switch ($PsCmdlet.ParameterSetName){
        “365GroupIdOnly”  {
            Write-Verbose "We've been given a 365 UPN, so we need the Group objects"
            $graphExtendedUG = get-graphGroups -tokenResponse $tokenResponse -filterUpn $graphExtendedUGUpn -selectAllProperties
            if(!$graphExtendedUG){
                Write-Error "Could not retrieve Unified Group from UPN [$graphExtendedUGUpn]"
                break
                }
            }
        “AADGroupIdOnly”  {
            Write-Verbose "We've been given an AAD UPN, so we need the Group objects"
            $graphMesg = get-graphGroups -tokenResponse $tokenResponse -filterUpn $graphMesgUpn
            if(!$graphMesg){
                Write-Error "Could not retrieve AAD Group from UPN [$graphMesgUpn]. Cannot continue."
                break
                }
            }
        #Now we've definitely got either $graphExtendedUG or $graphMesg, get the other one if it hasn't been supplied as a parameter
        {$_ -in "365GroupIdOnly","365GroupObjectSupplied"}  {
            if([string]::IsNullOrWhiteSpace($graphMesg)){
                switch ($syncWhat){
                    "Members" {
                        Write-Verbose "No `$graphMesg or `$graphMesgUpn provided - looking for Members group with Id [$($graphExtendedUG.anthesisgroup_UGSync.memberGroupId)] linked to UG [$($graphExtendedUG.DisplayName)][$($graphExtendedUG.id)]"
                        $graphMesg = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/groups/$($graphExtendedUG.anthesisgroup_UGSync.memberGroupId)"
                        if(!$graphMesg){
                            Write-Error "Could not retrieve AAD Members Group from ID [$($graphExtendedUG.anthesisgroup_UGSync.memberGroupId)]. Cannot continue."
                            break
                            }
                        }
                    "Owners"  {
                        Write-Verbose "No `$graphMesg or `$graphMesgUpn provided - looking for Owners group with Id [$($graphExtendedUG.anthesisgroup_UGSync.dataManagerGroupId)] linked to UG [$($graphExtendedUG.DisplayName)][$($graphExtendedUG.ExternalDirectoryObjectId)]"
                        $graphMesg = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/groups/$($graphExtendedUG.anthesisgroup_UGSync.dataManagerGroupId)"
                        if(!$graphMesg){
                            Write-Error "Could not retrieve AAD Owners Group from ID [$($graphExtendedUG.anthesisgroup_UGSync.dataManagerGroupId)]. Cannot continue."
                            break
                            }
                        }
                    }
                }            
            }
        {$_ -in "AADGroupIdOnly","AADGroupObjectSupplied"}  {
            if([string]::IsNullOrWhiteSpace($graphExtendedUG)){
                switch($syncWhat){
                    "Members" {
                        Write-Verbose "No `$graphExtendedUG or `$graphExtendedUGUpn provided - looking for associated 365 Group with `$graphExtendedUG.anthesisgroup_UGSync.memberGroupId -eq [$($graphMesg.Id)]"
                        #$graphExtendedUG = Get-UnifiedGroup -Filter "anthesisgroup_UGSync.memberGroupId -eq '$($graphMesg.Id)'"
                        $graphExtendedUG = get-graphGroups -tokenResponse $tokenResponseTeams -filterMembersGroupId $graphMesg.Id -selectAllProperties
                        }
                    "Owners" {
                        Write-Verbose "No `$graphExtendedUG or `$graphExtendedUGUpn provided - looking for associated 365 Group with `$graphExtendedUG.anthesisgroup_UGSync.dataManagerGroupId -eq [$($graphMesg.Id)]"
                        #$graphExtendedUG = Get-UnifiedGroup  -Filter "anthesisgroup_UGSync.dataManagerGroupId -eq '$($graphMesg.Id)'"
                        $graphExtendedUG = get-graphGroups -tokenResponse $tokenResponseTeams -filterDataManagersGroupId $graphMesg.Id -selectAllProperties
                        }
                    }
                if(!$graphExtendedUG){
                    Write-Error "Could not retrieve 365 Group based on $syncWhat AADGroupID [$($graphMesg.Id)]. Cannot continue."
                    break
                    }
                }
            
            }
        }
    #endregion
    
    if($graphMesg -and $graphExtendedUG){ #If we've got an AAD and a 365 Group to compare...
        $ugUsersBeforeChanges = @()
        $aadgUsersBeforeChanges = @()
        if($enumerateSubgroups){
            get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $graphMesg.Id -memberType TransitiveMembers -returnOnlyUsers | %{[array]$aadgUsersBeforeChanges += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $_.UserPrincipalName;"displayName"=$_.DisplayName;"objectId"=$_.Id})}
            }
        else{get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $graphMesg.Id -memberType Members -returnOnlyUsers | %{[array]$aadgUsersBeforeChanges += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $_.UserPrincipalName;"displayName"=$_.DisplayName;"objectId"=$_.Id})}}
        switch ($syncWhat){
            "Members" {
                get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $graphExtendedUG.Id -memberType Members -returnOnlyUsers | %{[array]$ugUsersBeforeChanges += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $_.UserPrincipalName;"displayName"=$_.DisplayName;"objectId"=$_.Id})}
                #if($sourceGroup -eq "AAD"){Get-AzureADGroupMember -All:$true -ObjectId $graphExtendedUG.anthesisgroup_UGSync.dataManagerGroupId | %{[array]$aadgUsersBeforeChanges += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $_.UserPrincipalName;"displayName"=$_.DisplayName;"objectId"=$_.ObjectId})}} #Add DataManagers too (to fix issue with Communities)
                }
            "Owners" {
                get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $graphExtendedUG.Id -memberType Owners -returnOnlyUsers | %{[array]$ugUsersBeforeChanges += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $_.UserPrincipalName;"displayName"=$_.DisplayName;"objectId"=$_.Id})}
                }
            }
        #If exception group, remove users from passed exception object from the UG array to omit from being compared to the AADG (which they will not be in, these groups in the geographic structure will be tied to other infrastructure we do not want applied to exception users)
        If($syncException){
        write-verbose "This AAD group is an exception - not syncing members from the defined exception user group"
            $syncExceptionUsers = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $syncException.exceptionTargetUserGroupToOmit -memberType Members -returnOnlyLicensedUsers
            $syncExceptionUsers | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name "objectID" -Value $_.Id}
            $ugUsersBeforeChanges =  Compare-Object -ReferenceObject $ugUsersBeforeChanges -DifferenceObject $syncExceptionUsers -Property objectId -PassThru -IncludeEqual | Where-Object {$_.SideIndicator -ne "=="}
            Write-Verbose "$($ugUsersBeforeChanges.userPrincipalName)"
        }
        

        $usersDelta = Compare-Object -ReferenceObject @($ugUsersBeforeChanges | select-object) -DifferenceObject @($aadgUsersBeforeChanges | select-object) -Property userPrincipalName -PassThru -IncludeEqual
         $($usersDelta | % {Write-Verbose "$_"})

        $usersAdded = @()
        $usersRemoved = @()
        $usersFailed = @()

        switch($sourceGroup){
            "365" {
                #Add extra users from UG to MESG
                $usersDelta | ?{$_.SideIndicator -eq "<="} | %{
                    $userToBeChanged = $_
                    Write-Verbose "`tAdding [$($userToBeChanged.userPrincipalName)] to [$($graphMesg.DisplayName)][$($graphMesg.Id)] MESG"
                    try{
                        #Unbelievbly, you still can't manage MESGs via Graph.
                        #add-graphUsersToGroup -tokenResponse $tokenResponse -graphGroupId $graphMesg.Id -memberType Members -graphUserIds $userToBeChanged.objectId -WhatIf:$WhatIfPreference -ErrorAction Stop #We always add to members regardless of $syncWhat because we're dealing with the MESGs. $syncWhat will already have set either the Data Managers MESG or Members MESG as $graphMesg
                        Add-DistributionGroupMember -Identity $graphMesg.Id -Member $userToBeChanged.objectId -BypassSecurityGroupManagerCheck -WhatIf:$WhatIfPreference -ErrorAction Stop
                        [array]$usersAdded += (New-Object psobject -Property $([ordered]@{"UPN"=$userToBeChanged.userPrincipalName;"DisplayName"=$userToBeChanged.displayName}))
                        }
                    catch{
                        Write-Warning "Failed to add [$($userToBeChanged.userPrincipalName)] to MESG [$($graphMesg.DisplayName)][$($graphMesg.Id)]" 
                        $_
                        [array]$usersFailed += (New-Object psobject -Property $([ordered]@{"Change"="Added";"UPN"=$userToBeChanged.userPrincipalName;"DisplayName"=$userToBeChanged.displayName;"ErrorMessage"=$_}))
                        }
                    }

                #Remove "removed" users from MESG
                $usersDelta | ?{$_.SideIndicator -eq "=>"} | %{ 
                    $userToBeChanged = $_
                    Write-Verbose "`tRemoving [$($userToBeChanged.userPrincipalName)] from [$($graphMesg.DisplayName)][$($graphMesg.Id)] MESG"
                    try{
                        #Unbelievbly, you still can't manage MESGs via Graph.
                        #remove-graphUsersFromGroup -tokenResponse $tokenResponse -graphGroupId $graphMesg.Id -memberType Members -graphUserIds $userToBeChanged.objectId -WhatIf:$WhatIfPreference -ErrorAction Stop
                        Remove-DistributionGroupMember -Identity $graphMesg.Id -Member $userToBeChanged.objectId -BypassSecurityGroupManagerCheck -Confirm:$false -WhatIf:$WhatIfPreference -ErrorAction Stop
                         [array]$usersRemoved += (New-Object psobject -Property $([ordered]@{"Change"="Removed";"UPN"=$userToBeChanged.userPrincipalName;"DisplayName"=$userToBeChanged.displayName}))
                        }
                    catch{
                        Write-Warning "Failed to remove [$($userToBeChanged.userPrincipalName)] from MESG [$($graphMesg.DisplayName)][$($graphMesg.Id)]"
                        [array]$usersFailed += (New-Object psobject -Property $([ordered]@{"Change"="Removed";"UPN"=$userToBeChanged.userPrincipalName;"DisplayName"=$userToBeChanged.displayName;"ErrorMessage"=$_}))
                        }
                    }                
                }
            "AAD" {
                #Add extra users from MESG to UG
                $usersDelta | ?{$_.SideIndicator -eq "=>" -and $_.DisplayName -notmatch "Shared Mailbox"} | %{
                    $userToBeChanged = $_
                    Write-Verbose "`tAdding [$($userToBeChanged.userPrincipalName)] to [$($graphExtendedUG.DisplayName)][$($graphExtendedUG.Id)] UG $syncWhat"
                    try{
                        #We want to add Data Managers as Members too, so we add to Members regardless of $syncWhat. However, we don't really need GroupBot as a Member, so we exclude this one exception
                        if($userToBeChanged.objectId -ne "00aa81e4-2e8f-4170-bc24-843b917fd7cf"){
                            try{
                                add-graphUsersToGroup -tokenResponse $tokenResponse -graphGroupId $graphExtendedUG.Id -memberType Members -graphUserIds $userToBeChanged.objectId -ErrorAction Stop -Verbose:$VerbosePreference
                                }
                            catch{
                                if($_.Message -match "One or more added object references already exist for the following modified properties"){continue}
                                else{$_}
                                }
                            }
                        if($syncWhat -eq "Owners"){ #If we are syncing Owners, we _don't_ want to exclude GroupBot (or anyone else)
                            add-graphUsersToGroup -tokenResponse $tokenResponse -graphGroupId $graphExtendedUG.Id -memberType Owners -graphUserIds $userToBeChanged.objectId -ErrorAction Stop -Verbose:$VerbosePreference
                            }
                        [array]$usersAdded += (New-Object psobject -Property $([ordered]@{"UPN"=$userToBeChanged.userPrincipalName;"DisplayName"=$userToBeChanged.displayName}))
                        }
                    catch{
                        Write-Warning "Failed to add [$($userToBeChanged.userPrincipalName)] to UG $syncWhat [$($graphExtendedUG.DisplayName)][$($graphExtendedUG.Id)]" 
                        [array]$usersFailed += (New-Object psobject -Property $([ordered]@{"Change"="Added";"UPN"=$userToBeChanged.userPrincipalName;"DisplayName"=$userToBeChanged.displayName;"ErrorMessage"=$_}))
                        }
                    }

                #Remove "removed" users from UG
                $usersDelta | ?{$_.SideIndicator -eq "<="} | %{ 
                    $userToBeChanged = $_
                    Write-Verbose "`tRemoving [$($userToBeChanged.userPrincipalName)] from [$($graphExtendedUG.DisplayName)][$($graphExtendedUG.Id)] UG $syncWhat"
                    try{
                        #if($syncWhat -eq "Owners"){start-sleep -Seconds 2} #Pause briefly if we're removing Owners because we can't remove the last owner from a group, and it takes a moment for any new owners we've added above to filter through.
                        remove-graphUsersFromGroup -tokenResponse $tokenResponse -graphGroupId $graphExtendedUG.Id -memberType $syncWhat -graphUserIds $userToBeChanged.objectId -ErrorAction Stop  -Verbose:$VerbosePreference
                        [array]$usersRemoved += (New-Object psobject -Property $([ordered]@{"Change"="Removed";"UPN"=$userToBeChanged.userPrincipalName;"DisplayName"=$userToBeChanged.displayName}))
                        }
                    catch{
                        Write-Warning "Failed to remove [$($userToBeChanged.userPrincipalName)] from UG $syncWhat [$($graphExtendedUG.DisplayName)][$($graphExtendedUG.Id)]"
                        [array]$usersFailed += (New-Object psobject -Property $([ordered]@{"Change"="Removed";"UPN"=$userToBeChanged.userPrincipalName;"DisplayName"=$userToBeChanged.displayName;"ErrorMessage"=$_}))
                        }
                    }                
                }
            }

        #Now report any problems/changes    
        if(!$dontSendEmailReport){
            $ownersAfterChanges = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $graphExtendedUG.id -memberType Owners -returnOnlyUsers -includeLineManager
            if(@($ownersAfterChanges | Select-Object).Count -eq 0){
               Write-Warning "No owners for 365 Group [$($graphExtendedUG.DisplayName)] - adding GroupBot"
               Add-DistributionGroupMember -Identity $graphExtendedUG.id -Member groupbot@anthesisgroup.com -BypassSecurityGroupManagerCheck -Confirm:$false #Add GroupBot in here instead of pissing everyone off with e-mail alerts.
                }
            if(@($ownersAfterChanges.DisplayName | ? {$_ -match "Ω"} | Select-Object).Count -eq @($ownersAfterChanges | Select-Object).Count){
                $now = Get-Date
                if($now.Hour -eq 8 -and $now.Minute -ge 0 -and $now.Minute -lt 30){#This function is run every 30 minutes, so this should generate 1 alert per day
                    Write-Warning "No active owners for 365 Group [$($graphExtendedUG.DisplayName)] - Notifying Line Managers to request reassignment"
                    send-dataManagerReassignmentRequest -tokenResponse $tokenResponseSmtp -UnifiedGroup $graphExtendedUG -currentOwners $ownersAfterChanges -adminEmailAddresses $adminEmailAddresses -WhatIf:$WhatIfPreference
                    }
                else {Write-Warning "No active owners for 365 Group [$($graphExtendedUG.DisplayName)] - Suppressing notification to Line Managers (to cut down on spam)"}
                }
            else{#If there is a problem _other than_ groups managed by deprovisioned users, notify IT
                Write-Verbose "Preparing 365<>MESG $syncWhat sync report to send to Admins & Owners"
                if($usersFailed.Count -ne 0){
                    Write-Warning "Found [$($usersFailed.Count)] problems - notifying 365 Group Admins"
                    $ugUsersAfterChanges = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $graphExtendedUG.id -memberType $syncWhat -returnOnlyUsers
                    $aadgUsersAfterChanges = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $graphMesg.id -memberType $syncWhat -returnOnlyUsers
                    send-membershipChangeProblemReportToAdmins  -tokenResponse $tokenResponseSmtp -UnifiedGroup $graphExtendedUG -changesAreTo $syncWhat -usersWithProblemsArray $usersFailed -usersIn365GroupAfterChanges $ugUsersAfterChanges -usersInAADGroupAfterChanges $aadgUsersAfterChanges -adminEmailAddresses $adminEmailAddresses -WhatIf:$WhatIfPreference
                    }
                else{Write-Verbose "No problems adding/removing users, not sending problem report e-mail to Admins"}                
                }

            if($usersAdded.Count -ne 0 -or $usersRemoved.Count -ne 0){
                Write-Verbose "[$($usersAdded.Count + $usersRemoved.Count)] changes made - sending the change report to managers and admins"
                $ownersEmailAddresses = @($ownersAfterChanges.UserPrincipalName)
                if($syncWhat -eq "Owners"){
                    Write-Verbose "Getting all group Owners (both added and removed) for [$($graphExtendedUG.DisplayName)]"
                    $ownersEmailAddresses += $usersAdded.UPN
                    $ownersEmailAddresses += $usersRemoved.UPN
                    $ownersEmailAddresses = $ownersEmailAddresses | Select-Object -Unique
                    }
                $ugUsersAfterChanges = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $graphExtendedUG.id -memberType $syncWhat -returnOnlyUsers
                send-membershipChangeReportToManagers -tokenResponse $tokenResponseSmtp -UnifiedGroup $graphExtendedUG -changesAreTo $syncWhat -usersAddedArray $usersAdded -usersRemovedArray $usersRemoved -usersWithProblemsArray $usersFailed -usersInGroupAfterChanges $ugUsersAfterChanges -adminEmailAddresses $adminEmailAddresses -ownersEmailAddresses $ownersEmailAddresses -WhatIf:$WhatIfPreference
                }
            else{Write-Verbose "No membership changes - not sending report to Mangers & Admins"}
            }
        }
    else{
        if(!$graphMesg){
            Write-Error "No AAD group found for UG [$($graphExtendedUG.DisplayName)][$($graphExtendedUG.ExternalDirectoryObjectId)]"
            break
            }
        elseif(!$graphExtendedUG){
            Write-Error "No 365 group found for AAD Group [$($graphMesg.DisplayName)][$($graphMesg.Id)]"
            break
            }
        }
    }
