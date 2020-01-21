#Sync Office 365 Group membership to correspnoding security group membership

Import-Module _PS_Library_MSOL.psm1
Import-Module _PS_Library_GeneralFunctionality
#Import-Module *pnp*


function addto-SharepointTeamsTermStore{
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
    if(![string]::IsNullOrWhiteSpace($displayName)){$guessedAlias = $displayName.replace(" ","_").Replace("(","").Replace(")","").Replace(",","")}
    $guessedAlias = set-suffixAndMaxLength -string $guessedAlias -suffix $fixedSuffix -maxLength 64
    $guessedAlias = sanitise-forMicrosoftEmailAddress -dirtyString $guessedAlias
    $guessedAlias = remove-diacritics -String $guessedAlias
    Write-Verbose -Message "guess-aliasFromDisplayName([$displayName],[$fixedSuffix]) = [$guessedAlias]"
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
    #$UnifiedGroupObject.CustomAttribute2 = Data Managers Subgroup ExternalDirectoryObjectId
    #$UnifiedGroupObject.CustomAttribute3 = Members Subgroup ExternalDirectoryObjectId
    #$UnifiedGroupObject.CustomAttribute4 = Combined Mail-Enabled Security Group ExternalDirectoryObjectId
    #$UnifiedGroupObject.CustomAttribute5 = Shared Mailbox ExternalDirectoryObjectId
    #$UnifiedGroupObject.CustomAttribute6 = [string] "365"|"AAD" Is membership driven by the 365 Group or the associated AAD group?
    #$UnifiedGroupObject.CustomAttribute7 = [string] "Internal"|"External"|"Confidential" Intended Site Classification (used to reset in the event of unauthorised change)
    #$UnifiedGroupObject.CustomAttribute8 = [string] "Public"|"Private" Intended Site Privacy (used to reset in the event of unauthorised change)
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
    $365MailAlias = $(guess-aliasFromDisplayName "$displayName 365")
    $managersSgDisplayNameSuffix = " - Data Managers Subgroup"
    $membersSgDisplayNameSuffix = " - Members Subgroup"
    $sharedMailboxDisplayName = "Shared Mailbox - $displayName"

    #Firstly, check whether we have already created a Unified Group for this DisplayName
    $365Group = Get-UnifiedGroup -Filter "DisplayName -eq `'$(sanitise-forSql $displayName)`'"
    if(!$365Group){$365Group = Get-UnifiedGroup -Filter "Alias -eq `'$365MailAlias`'"} #If we can't find it by the DisplayName, check the Alias as this is less mutable

    #If we have a UG, check whether we can find the associated groups (we certainly should be able to!)
    if($365Group){
        Write-Verbose "Pre-existing 365 Group found [$($365Group.DisplayName)] with: `r`n`t`tCA1=[$($365Group.CustomAttribute1)], `r`n`t`tCA2=[$($365Group.CustomAttribute2)], `r`n`t`tCA3=[$($365Group.CustomAttribute3)], `r`n`t`tCA4=[$($365Group.CustomAttribute4)], `r`n`t`tCA5=[$($365Group.CustomAttribute5)], `r`n`t`tCA6=[$($365Group.CustomAttribute6)], `r`n`t`tCA7=[$($365Group.CustomAttribute7)], `r`n`t`tCA8=[$($365Group.CustomAttribute8)]"
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
        }
    else{
        Write-Verbose "No pre-existing 365 group found - checking for AAD Groups."

        #Check whether any of these MESG exist based on names (just in case we're re-creating a 365 group and want to retain the AAD Groups)
        $combinedSg = rummage-forDistributionGroup -displayName $displayName
        if($combinedSg){Write-Verbose "`tCombined Group [$($combinedSg.DisplayName)] found"}else{Write-Verbose "`tCombined group not found"}
        $managersSg = rummage-forDistributionGroup -displayName ($displayName+$managersSgDisplayNameSuffix)
        if($managersSg){Write-Verbose "`tManagers Group [$($managersSg.DisplayName)] found"}else{Write-Verbose "`tManagers group not found"}
        $membersSg  = rummage-forDistributionGroup -displayName ($displayName +$membersSgDisplayNameSuffix)
        if($membersSg){Write-Verbose "`tMembers Group [$($membersSg.DisplayName)] found"}else{Write-Verbose "`tMembers group not found"}
        $sharedMailbox = Get-Mailbox -Filter "DisplayName -eq `'$(sanitise-forSql $sharedMailboxDisplayName)`'"
        if(!$sharedMailbox){$sharedMailbox = Get-Mailbox -Filter "Alias -eq `'$(guess-aliasFromDisplayName $sharedMailboxDisplayName)`'"} #If we can't find it by the DisplayName, check the Alias as this is less mutable
        if($sharedMailbox){Write-Verbose "`tShared Mailbox [$($sharedMailbox.DisplayName)] found"}else{Write-Verbose "`tMailbox not found"}

        #Create any groups that don't already exist
        if(!$combinedSg){
            Write-Verbose "Creating Combined Security Group [$displayName]"
            try{$combinedSg = new-mailEnabledSecurityGroup -dgDisplayName $displayName -membersUpns $null -hideFromGal $false -blockExternalMail $true -ownersUpns "ITTeamAll@anthesisgroup.com" -description "Mail-enabled Security Group for $displayName" -WhatIf:$WhatIfPreference}
            catch{Write-Error $_}
            }

        if($combinedSg -or $WhatIfPreference){ #If we now have a Combined SG
            if(!$managersSg){ #Create a Managers SG if required
                Write-Verbose "Creating Data Managers Security Group [$displayName$managersSgDisplayNameSuffix]"
                $managersMemberOf = @($combinedSg.ExternalDirectoryObjectId)
                if($ownersAreRealManagers){$managersMemberOf += "Managers (All)"}
                try{$managersSg = new-mailEnabledSecurityGroup -dgDisplayName $displayName -fixedSuffix $managersSgDisplayNameSuffix -membersUpns $managerUpns -memberOf $managersMemberOf -hideFromGal $false -blockExternalMail $true -ownersUpns "ITTeamAll@anthesisgroup.com" -description "Mail-enabled Security Group for managing Ownership of $displayName Unified Group" -WhatIf:$WhatIfPreference -Verbose}
                catch{Write-Error $_}
                }

            if(!$membersSg){ #And create a Members SG if required
                Write-Verbose "Creating Members Security Group [$displayName$membersSgDisplayNameSuffix]"
                try{
                    $membersSg = new-mailEnabledSecurityGroup -dgDisplayName $displayName -fixedSuffix $membersSgDisplayNameSuffix -membersUpns $teamMemberUpns -memberOf $combinedSg.ExternalDirectoryObjectId -hideFromGal $false -blockExternalMail $true -ownersUpns "ITTeamAll@anthesisgroup.com" -description "Mail-enabled Security Group for mirroring membership of $displayName$members Unified Group" -WhatIf:$WhatIfPreference
                    if(![string]::IsNullOrWhiteSpace($memberOf)){
                        $memberOf | % { #We now nest membership via Members groups, rather than Combined Groups, so this is a little more complicated now.
                            $parentGroup = get-membersGroup -groupName $_
                            Add-DistributionGroupMember -Identity $parentGroup.ExternalDirectoryObjectId -BypassSecurityGroupManagerCheck:$true -Member $membersSg.ExternalDirectoryObjectId -Confirm:$false
                            }
                        }
                    }
                catch{Write-Error $_}
                }
            }
        else{Write-Error "Combined Security Group [$displayName] not available. Cannot proceed with SubGroup creation"}        
        }

    #Check that everything's worked so far
    if(!$combinedSg){
        if($WhatIfPreference){Write-Verbose "Combined Security Group [$displayName] not created because we're only pretending."}
        else{Write-Error "Combined Security Group [$displayName] not available. Cannot proceed with UnifiedGroup creation";break}}
    if(!$managersSg){
        if($WhatIfPreference){Write-Verbose "Managers Security Group [$displayName$managersSgDisplayNameSuffix] not created because we're only pretending."}
        else{Write-Error "Managers Security Group [$displayName$managersSgDisplayNameSuffix] not available. Cannot proceed with UnifiedGroup creation";break}}
    if(!$membersSg){
        if($WhatIfPreference){Write-Verbose "Members Security Group [$displayName$membersSgDisplayNameSuffix] not created because we're only pretending."}
        else{Write-Error "Members Security Group [$displayName$membersSgDisplayNameSuffix] not available. Cannot proceed with UnifiedGroup creation";break}}
    if(!$365Group -or $WhatIfPreference){
        if(($combinedSg -and $managersSg -and $membersSg)){#If we now have all the prerequisite groups, create a UG
            try{
                $groupIsNew = $true
                Write-Verbose "All MESGs found - creating Unified 365 Group [$displayName]"
                if([string]::IsNullOrWhiteSpace($description)){$description = "Unified 365 Group for $displayName"}
                #Create the UG
                # Example of json for POST https://graph.microsoft.com/v1.0/groups
                # https://docs.microsoft.com/en-us/graph/api/group-post-groups?view=graph-rest-1.0
                $owners = @()
                $managerUpns | % {[string[]]$owners += ("https://graph.microsoft.com/v1.0/users/$_").ToLower()}
                $members = @()
                $teamMemberUpns | % {[string[]]$members += ("https://graph.microsoft.com/v1.0/users/$_").ToLower()}
                $members = $($members+$owners) | Sort-Object | Get-Unique -AsString 

                $creategroup = "{`
                    `"displayName`": `"$(sanitise-forJson $displayName)`",
                    `"groupTypes`": [
                      `"Unified`"
                    ],
                    `"mailEnabled`": true,
                    `"mailNickname`": `"$365MailAlias`",
                    `"securityEnabled`": true,
                    `"owners@odata.bind`": [
                        `"$($owners -join "``",`r`n``"")`"
                      ],
                    `"members@odata.bind`": [
                        `"$($members -join "``",`r`n``"")`"
                      ]
                    }"
                Write-Verbose $creategroup
                $creategroup = [System.Text.Encoding]::UTF8.GetBytes($creategroup) #To ensure Non-English characters are encoded correctly
                $response = Invoke-RestMethod -Uri https://graph.microsoft.com/v1.0/groups -Body $creategroup -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Post
                
                Connect-PnPOnline -AccessToken $tokenResponse.access_token
                #Set-PnPUnifiedGroup -Identity $response.id -DisplayName $displayName <# Graph API doesn't handle accents/diacritics properly and replaces them with �, so we have to set the DisplayName again via Pnp #>
                do{
                    Write-Verbose "Waiting for Unified Group to provision..."
                    $pnp365Group = Get-PnPUnifiedGroup -Identity $response.id -ErrorAction SilentlyContinue -WarningAction SilentlyContinue #This is (allegedly) the bit that triggers Site Collection creation
                    $365Group = Get-UnifiedGroup -Identity $response.id -ErrorAction SilentlyContinue -WarningAction SilentlyContinue 
                    Start-Sleep -Seconds 5
                    }
                while([string]::IsNullOrWhiteSpace($365Group))
                }
            catch{Write-Error $_}
            }
        else{Write-Error "Combined/Managers/Members Security Group [$displayName]/[$displayName$managersSgDisplayNameSuffix]/[$displayName$membersSgDisplayNameSuffix] not available. Cannot proceed with UnifiedGroup creation";break}
        }

    if($365Group){ #If we now have a 365 UG, set the CustomAttributes, and create a Shared Mailbox (if required) and configure it
        if($groupIsNew){$ca5Disclaimer = "It's okay if this is empty at the moment"}
        else{$ca5Disclaimer = "This shouldn't be empty now"}
        Write-Verbose "`tSet-UnifiedGroup -Identity [$($365Group.ExternalDirectoryObjectId)] -HiddenFromAddressListsEnabled [$true] `r`n`t`t-CustomAttribute1 [$($365Group.ExternalDirectoryObjectId)] `r`n`t`t-CustomAttribute2 [$($managersSg.ExternalDirectoryObjectId)] `r`n`t`t-CustomAttribute3 [$($membersSg.ExternalDirectoryObjectId)] `r`n`t`t-CustomAttribute4 [$($combinedSg.ExternalDirectoryObjectId)] `r`n`t`t-CustomAttribute5 [$($sharedMailbox.ExternalDirectoryObjectId)] - $ca5Disclaimer `r`n`t`t-CustomAttribute6 [$($membershipmanagedBy)] `r`n`t`t-CustomAttribute7 [$($groupClassification)] `r`n`t`t-CustomAttribute8 [$($accessType)] `r`n`t`t-WhatIf:[$($WhatIfPreference)] -AccessType [$($accessType)] -RequireSenderAuthenticationEnabled [$($blockExternalMail)] -AutoSubscribeNewMembers:[$($autoSubscribe)] -AlwaysSubscribeMembersToCalendarEvents:[$($autoSubscribe)] -Classification [$($groupClassification)]"
        Set-UnifiedGroup -Identity $365Group.ExternalDirectoryObjectId -HiddenFromAddressListsEnabled $true -CustomAttribute1 $365Group.ExternalDirectoryObjectId -CustomAttribute2 $managersSg.ExternalDirectoryObjectId -CustomAttribute3 $membersSg.ExternalDirectoryObjectId -CustomAttribute4 $combinedSg.ExternalDirectoryObjectId -CustomAttribute6 $membershipmanagedBy -CustomAttribute7 $groupClassification -CustomAttribute8 $accessType -WhatIf:$WhatIfPreference -AccessType $accessType -RequireSenderAuthenticationEnabled $blockExternalMail -AutoSubscribeNewMembers:$autoSubscribe -AlwaysSubscribeMembersToCalendarEvents:$autoSubscribe -Classification $groupClassification
        $365Group = Get-UnifiedGroup $365Group.ExternalDirectoryObjectId
        
        if(!$sharedMailbox){
            Write-Verbose "Creating Shared Mailbox [$sharedMailboxDisplayName]: New-Mailbox -Shared -DisplayName $sharedMailboxDisplayName -Name $sharedMailboxDisplayName -ErrorAction Continue -WhatIf:$WhatIfPreference "
            try{$sharedMailbox = New-Mailbox -Shared -DisplayName $sharedMailboxDisplayName -Name $(guess-aliasFromDisplayName $sharedMailboxDisplayName) -ErrorAction Stop -WhatIf:$WhatIfPreference}
            catch{
                if($_.Exception.Message -match "is already being used by the proxy addresses or LegacyExchangeDN. Please choose another proxy address."){ #Shit, but it returns generic HResult -2146233087, which is even less help.
                    Write-Warning $_.Exception.Message
                    try{
                        $alternativeAlias = guess-aliasFromDisplayName -displayName $sharedMailboxDisplayName -fixedSuffix $("_"+([guid]::NewGuid()).Guid)
                        Write-Verbose "`tTrying again with unique Name/Alias/Email: [$sharedMailboxDisplayName][$($alternativeAlias)]"
                        $sharedMailbox = New-Mailbox -Shared -DisplayName $sharedMailboxDisplayName -Name $alternativeAlias -Alias $alternativeAlias -PrimarySmtpAddress ($alternativeAlias+"@anthesisgroup.com") -ErrorAction Continue -WhatIf:$WhatIfPreference
                        }
                    catch{$_}
                    }
                else{$_} 
                }
            }

        if($sharedMailbox){
            Write-Verbose "Found Mailbox: [$($sharedMailbox.DisplayName)][$($sharedMailbox.ExternalDirectoryObjectId)] `r`n`t`tSet-Mailbox -Identity $($sharedMailbox.ExternalDirectoryObjectId) -HiddenFromAddressListsEnabled $true -RequireSenderAuthenticationEnabled $false -ForwardingAddress $($365Group.PrimarySmtpAddress) -DeliverToMailboxAndForward $true -Confirm:$false -WhatIf:$WhatIfPreference"
            Set-Mailbox -Identity $sharedMailbox.ExternalDirectoryObjectId -HiddenFromAddressListsEnabled $true -RequireSenderAuthenticationEnabled $false -ForwardingAddress $365Group.PrimarySmtpAddress -DeliverToMailboxAndForward $true  -Confirm:$false -WhatIf:$WhatIfPreference 
            Set-user -Identity $sharedMailbox.ExternalDirectoryObjectId -Manager kevin.maitland -WhatIf:$WhatIfPreference  #For want of someone better....
            #Assign the Shared Mailbox as a member of the Security Group
            try{Add-DistributionGroupMember -Identity $combinedSg.ExternalDirectoryObjectId -Member $sharedMailbox.ExternalDirectoryObjectId -BypassSecurityGroupManagerCheck -WhatIf:$WhatIfPreference -ErrorAction Stop}
            catch{
                if('-2146233087' -eq $_.Exception.HResult){Write-Verbose "Shared Mailbox [$($sharedMailbox.DisplayName)] is already a member of [$($combinedSg.DisplayName)]"}
                else{Write-Error $_}
                }
            Set-UnifiedGroup -Identity $365Group.ExternalDirectoryObjectId -CustomAttribute5 $sharedMailbox.ExternalDirectoryObjectId -WhatIf:$WhatIfPreference 
            }
        else{Write-Error "Shared Mailbox not available. Cannot complete UG setup."}
        }
    else{Write-Error "Unified Group [$displayName] not available. Cannot proceed with Shared Mailbox creation."}

    #Now we've set everything, retrieve the object to confirm the changes were applied successfully:
    $365Group = Get-UnifiedGroup -Identity $365Group.ExternalDirectoryObjectId
    if($groupIsNew){Write-Verbose "New 365 Group created: [$($365Group.DisplayName)] with: `r`n`t`tCA1=[$($365Group.CustomAttribute1)], `r`n`t`tCA2=[$($365Group.CustomAttribute2)], `r`n`t`tCA3=[$($365Group.CustomAttribute3)], `r`n`t`tCA4=[$($365Group.CustomAttribute4)], `r`n`t`tCA5=[$($365Group.CustomAttribute5)], `r`n`t`tCA6=[$($365Group.CustomAttribute6)], `r`n`t`tCA7=[$($365Group.CustomAttribute7)], `r`n`t`tCA8=[$($365Group.CustomAttribute8)]"}
    elseif($365Group){Write-Verbose "Pre-existing 365 Group found: [$($365Group.DisplayName)] with: `r`n`t`tCA1=[$($365Group.CustomAttribute1)], `r`n`t`tCA2=[$($365Group.CustomAttribute2)], `r`n`t`tCA3=[$($365Group.CustomAttribute3)], `r`n`t`tCA4=[$($365Group.CustomAttribute4)], `r`n`t`tCA5=[$($365Group.CustomAttribute5)], `r`n`t`tCA6=[$($365Group.CustomAttribute6)], `r`n`t`tCA7=[$($365Group.CustomAttribute7)], `r`n`t`tCA8=[$($365Group.CustomAttribute8)]"}
    else{Write-Warning "It doesn't look like there's a [$displayName] 365 Group available...";break}

    #Provision MS Team if requested
    if($alsoCreateTeam -and $365Group){
        Write-Verbose "Provisioning new MS Team (as requested)"
        if(Get-Team -GroupId $365Group.ExternalDirectoryObjectId){
            Write-Warning "Existing Team found - not attempting to reprovision"
            }
        else{New-Team -GroupId $365Group.ExternalDirectoryObjectId -AllowGuestCreateUpdateChannels $false -AllowGuestDeleteChannels $false}
        }
    else{Write-Verbose "_NOT_ attempting to provision new MS Team"}

    do{
        Write-Verbose "Waiting for Unified Group Site to provision..."
        Connect-PnPOnline -AccessToken $tokenResponse.access_token
        if($response){
            Write-Verbose "Get-PnPUnifiedGroup -Identity [$($response.id)] (`$GraphResponse)"
            $pnp365Group = Get-PnPUnifiedGroup -Identity $response.id -ErrorAction SilentlyContinue -WarningAction SilentlyContinue <#This is (allegedly) the bit that triggers Site Collection creation#>
            }
        elseif($365Group){
            Write-Verbose "Get-PnPUnifiedGroup -Identity [$($365Group.ExternalDirectoryObjectId)] (`$365Group)"
            $pnp365Group = Get-PnPUnifiedGroup -Identity $365Group.ExternalDirectoryObjectId -ErrorAction SilentlyContinue -WarningAction SilentlyContinue <#This is (allegedly) the bit that triggers Site Collection creation#>
            }
        else{Write-Warning "I haven't got a `$response or `$365Group object, so I can't check whether the Site has been provisioned!"}
        Start-Sleep -Seconds 5
        }
    while([string]::IsNullOrWhiteSpace($pnp365Group.SiteUrl))
    
    <#Is this a good idea here? It will immediately remove the executing user from the Site Collections Admins, which means they won;t be able to set a Homepage, or enable Site features. 
    Probably better if it's run as part of the new-externalSpoSites or new-internalSpoSites scripts. There's also a scheduled script that runs this too, so it won't stay wrong for long even if there's an error further through the provisioning scripts...
    Write-Verbose "set-standardTeamSitePermissions -teamSiteAbsoluteUrl [$($pnp365Group.SiteUrl)]"
    set-standardSitePermissions -unifiedGroupObject $365Group -tokenResponse $tokenResponse -pnpCreds $pnpCreds
    #>
    $365Group

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
    $mailName = guess-aliasFromDisplayName -displayName $dgDisplayName -fixedSuffix $fixedSuffix

    #Check to see if this already exists
    $mesg = rummage-forDistributionGroup -displayName ($dgDisplayName+$fixedSuffix)
    if($mesg){ #If the group already exists, add the new Members (ignore any removals - we'll let sync-groupMembership figure that out)
        Write-Verbose "Existing Mail-Enabled Distribution Group [$($mesg.DisplayName)][$($mesg.ExternalDirectoryObjectId)] retrieved"
        $members  | % {
            Write-Verbose "Adding TeamMember Add-DistributionGroupMember $($mesg.ExternalDirectoryObjectId) -Member $_ -Confirm:$false -BypassSecurityGroupManagerCheck"
            Add-DistributionGroupMember $mesg.ExternalDirectoryObjectId -Member $_ -Confirm:$false -BypassSecurityGroupManagerCheck -WhatIf:$WhatIfPreference
            }
        }
    else{ #If the group doesn't exist, try creating it
        $possibleNameCollision = rummage-forDistributionGroup -name $mailName
        if($possibleNameCollision){
            if(($possibleNameCollision.DisplayName -ne ($dgDisplayName+$fixedSuffix)) -and ($possibleNameCollision.Name -ne $mailName)){
                $alternativeName = set-suffixAndMaxLength -string $mailName -suffix $("_"+([guid]::NewGuid()).Guid) -maxLength $mailName.Length
                Write-Warning "This looks like a `"Name`" collision. Searching for [$mailName] has returned: `r`n`t[$($possibleNameCollision.DisplayName)][$($possibleNameCollision.ExternalDirectoryObjectId)]`r`n`tInstead of: `r`n`t[$($dgDisplayName+$fixedSuffix)]`r`n`tWill create new group with an alternative name [$alternativeName]"
                $mailName = $alternativeName
                }
            }
        try{
            Write-Verbose "New-DistributionGroup -Name $mailName -DisplayName ($dgDisplayName+$fixedSuffix) -Type Security -Members [$($membersUpns -join ", ")] -PrimarySmtpAddress $($mailName+"@anthesisgroup.com") -Notes $description -Alias $mailName -WhatIf:$WhatIfPreference"
            $mesg = New-DistributionGroup -Name $mailName -DisplayName ($dgDisplayName+$fixedSuffix) -Type Security -Members $membersUpns -PrimarySmtpAddress $($mailName+"@anthesisgroup.com") -Notes $description -Alias $mailName -WhatIf:$WhatIfPreference -ErrorAction Stop
            Write-Verbose "New Mail-Enabled Distribution Group [$($mesg.DisplayName)][$($mesg.ExternalDirectoryObjectId)] created"
            }
        catch{
            if($_.Exception.Message -match "is already being used by the proxy addresses or LegacyExchangeDN"){ #Shit, but it returns generic HResult -2146233087, which is even less help.
                Write-Warning $_.Exception.Message
                try{
                    $alternativeName = guess-aliasFromDisplayName -displayName ($dgDisplayName+$fixedSuffix) -fixedSuffix $("_"+([guid]::NewGuid()).Guid)
                    Write-Verbose "`tTrying New-DistributionGroup again with unique Name/Alias/Email: [($dgDisplayName$fixedSuffix)][$($alternativeName)]"
                    $mesg = New-DistributionGroup -Name $alternativeName -DisplayName ($dgDisplayName+$fixedSuffix) -Type Security -Members $membersUpns -PrimarySmtpAddress $($alternativeName+"@anthesisgroup.com") -Notes $description -Alias $alternativeName -WhatIf:$WhatIfPreference -ErrorAction Stop
                    if(!$mesg){
                        Write-Error "Could not create New-DistributionGroup again with unique Name/Alias/Email: [($dgDisplayName$fixedSuffix)][$($alternativeName)]"
                        return
                        }
                    }
                catch{$_}
                }
            else{
                Write-Error "Error creating new Distribution Group [$($dgDisplayName+$fixedSuffix)] in new-mailEnabledSecurityGroup()"
                $_
                } 
            }
        }

    if(!$mesg){Write-Error "Mail-Enabled Security Group [$dgDisplayName$fixedSuffix] neither found, nor created :/"}
    else{ #Now set the additional properties and MemberOf
        Write-Verbose "Set-DistributionGroup -Identity [$($mesg.ExternalDirectoryObjectId)] -HiddenFromAddressListsEnabled $hideFromGal -RequireSenderAuthenticationEnabled $blockExternalMail -ManagedBy [$($ownersUpns -join ", ")]"
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
function rummage-forDistributionGroup(){
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,ParameterSetName="DisplayName")]
            [string]$displayName
        ,[Parameter(Mandatory=$false,ParameterSetName="DisplayName")]
        [Parameter(Mandatory=$true,ParameterSetName="ObjectName")]
            [string]$name
        )

    Write-Verbose "rummage-forDistributionGroup([$displayName],[$name])"
    if(![string]::IsNullOrWhiteSpace($displayName)){
        Write-Verbose "`tTrying to get DG by DisplayName [$displayName]"
        [array]$dg = Get-DistributionGroup -Filter "DisplayName -eq `'$(sanitise-forSql $displayName)`'"
        }
    if($dg.Count -ne 1){
        if(![string]::IsNullOrWhiteSpace($name)){
            Write-Verbose "`tTrying to get DG by (Object) Name [$name]"
            [array]$dg = Get-DistributionGroup -Filter "Name -eq `'$name`'" #If we can't find it by the DisplayName, check the Name as this is less mutable and we might be testing to see whether we can create a new Group with this Name.
            }
        } 
    if($dg.Count -ne 1){
        $dg = $null
        if($dg.Count -gt 1){Write-Warning "`t`tMultiple Groups matched for Distribution Group [$displayName]`r`n`t $($dg.PrimarySmtpAddress -join "`r`n`t")"}
        if($dg.Count -eq 0){Write-Verbose "`t`tNo Distribution Group found"}
        }
    $dg
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
        ,[Parameter(Mandatory=$false, Position=1)]
        [array]$currentOwners
        ,[Parameter(Mandatory=$false, Position=2)]
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
        [string[]]$adminEmailAddresses
        ,[Parameter(Mandatory=$false,ParameterSetName="365GroupObjectSupplied")]
        [Parameter(Mandatory=$false,ParameterSetName="AADGroupObjectSupplied")]
        [Parameter(Mandatory=$false,ParameterSetName="365GroupIdOnly")]
        [Parameter(Mandatory=$false,ParameterSetName="AADGroupIdOnly")]
        [bool]$enumerateSubgroups = $false
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
        $ugUsersBeforeChanges = @()
        $aadgUsersBeforeChanges = @()
        if($enumerateSubgroups){enumerate-nestedAADGroups -aadGroupId $AADGroup.Id -Verbose:$VerbosePreference  | %{[array]$aadgUsersBeforeChanges += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $_.UserPrincipalName;"displayName"=$_.DisplayName;"objectId"=$_.ObjectId})}}
        else{Get-AzureADGroupMember -All:$true -ObjectId $AADGroup.Id | %{[array]$aadgUsersBeforeChanges += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $_.UserPrincipalName;"displayName"=$_.DisplayName;"objectId"=$_.ObjectId})}}
        switch ($syncWhat){
            "Members" {
                Get-AzureADGroupMember -All:$true -ObjectId $UnifiedGroup.ExternalDirectoryObjectId | %{[array]$ugUsersBeforeChanges += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $_.UserPrincipalName;"displayName"=$_.DisplayName;"objectId"=$_.ObjectId})}
                }
            "Owners" {
                Get-AzureADGroupOwner -All:$true -ObjectId $UnifiedGroup.ExternalDirectoryObjectId | %{[array]$ugUsersBeforeChanges += New-Object psobject -Property $([ordered]@{"userPrincipalName"= $_.UserPrincipalName;"displayName"=$_.DisplayName;"objectId"=$_.ObjectId})}
                }
            }

        $usersDelta = Compare-Object -ReferenceObject $ugUsersBeforeChanges -DifferenceObject $aadgUsersBeforeChanges -Property userPrincipalName -PassThru -IncludeEqual
         $($usersDelta | % {Write-Verbose "$_"})

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
                        Add-UnifiedGroupLinks -Identity $UnifiedGroup.ExternalDirectoryObjectId -Links $userToBeChanged.objectId -LinkType Members -Confirm:$false -WhatIf:$WhatIfPreference -ErrorAction Stop
                        if($syncWhat -eq "Owners"){
                            Add-UnifiedGroupLinks -Identity $UnifiedGroup.ExternalDirectoryObjectId -Links $userToBeChanged.objectId -LinkType Owners -Confirm:$false -WhatIf:$WhatIfPreference -ErrorAction Stop
                            } #Only Members can be Owners of a group. Please add 'User.Name' first as members before adding them as owners.
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
                    $owners = $ugUsersBeforeChanges
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

