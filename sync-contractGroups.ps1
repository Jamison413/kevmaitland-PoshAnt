if($PSCommandPath){
    $InformationPreference = 2
    $VerbosePreference = 0
    $logFileLocation = "C:\ScriptLogs\"
    #$suffix = "_fullSync"
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))$suffix`_Transcript_$(Get-Date -Format "yyyy-MM-dd").log"
    Start-Transcript $transcriptLogName -Append
    }

function sync-Members() {
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [array]$intendedMembers
        ,[parameter(Mandatory = $true)]
            [string]$targetGroupId
        )
    
    try{
        $currentMembers = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $targetGroupId -memberType TransitiveMembers -selectAllProperties -ErrorAction Stop
        }
    catch{
        get-errorSummary $_
        return
        }

    #Add new/missing intended members to group
    $deltaMembers = Compare-Object -ReferenceObject @($intendedMembers | Select-Object) -DifferenceObject @($currentMembers | Select-Object) -Property userPrincipalName -PassThru
    $toAdd = $deltaMembers | ? {$_.SideIndicator -eq "<="} 
    Write-Host "Adding [$($toAdd.Count)] members to group [$($targetGroupId)]"
    Write-Host "`t[$($toAdd.userPrincipalName -join '`r`n')]"
    $toAdd | % {
        if($addThese.count -lt 20){ #Bulk add only supports 20 users
            [array]$addThese += $_
            }
        else{
            try {
                add-graphUsersToGroup -tokenResponse $tokenResponseTeamsBot -graphGroupId $targetGroupId -memberType Members -graphUserIds @($addThese.id)
                $addThese = @($_)
                }
            catch{get-errorSummary $_}
            }
        }
    if($addThese.Count -ne 0){ #Add the final batch
        try{add-graphUsersToGroup -tokenResponse $tokenResponseTeamsBot -graphGroupId $targetGroupId -memberType Members -graphUserIds @($addThese.id)}
        catch{get-errorSummary $_}
        }

    #Remove extra/inapproprpriate members from group
    $toRemove = $deltaMembers | ? {$_.SideIndicator -eq "=>"} 
    Write-Host "Removing [$($toRemove.Count)] members to group [$($targetGroupId)]"
    Write-Host "`t[$($toAdd.userPrincipalName -join '`r`n')]"
    $toRemove | % {
        if($removeThese.count -lt 20){ #Bulk add only supports 20 users
            [array]$removeThese += $_
            }
        else{
            try{
                remove-graphUsersFromGroup -tokenResponse $tokenResponseTeamsBot -graphGroupId $targetGroupId -memberType Members -graphUserIds @($removeThese.id)
                $removeThese = @($_)
                }
            catch{get-errorSummary $_}
            }
        }
    if($removeThese.Count -ne 0){ #Add the final batch
        try{add-graphUsersToGroup -tokenResponse $tokenResponseTeamsBot -graphGroupId $targetGroupId -memberType Members -graphUserIds @($removeThese.id)}
        catch{get-errorSummary $_}
        }
    }

 
$teambotDetails = get-graphAppClientCredentials -appName TeamsBot
$tokenResponseTeamsBot = get-graphTokenResponse -aadAppCreds $teambotDetails

$allActiveEmployees = get-graphUsers -tokenResponse $tokenResponseTeamsBot -filterLicensedUsers -selectAllProperties
#get-graphGroups -tokenResponse $tokenResponseTeamsBot -filterDisplayNameStartsWith "Contract"

$intendedEmployees = $allActiveEmployees | ? {$_.anthesisgroup_employeeInfo.contractType -eq "Employee"}
sync-Members -tokenResponse $tokenResponseTeamsBot -intendedMembers @($intendedEmployees) -targetGroupId b32f993e-b113-48c7-bd88-9b804c901be3

$intendedServiceAccounts = $allActiveEmployees | ? {$_.anthesisgroup_employeeInfo.contractType -eq "ServiceAccount"}
sync-Members -tokenResponse $tokenResponseTeamsBot -intendedMembers @($intendedServiceAccounts) -targetGroupId 384cccb4-e23e-4949-99d4-2211b23b55df

$intendedSubcontractors = $allActiveEmployees | ? {$_.anthesisgroup_employeeInfo.contractType -eq "ServiceAccount"}
sync-Members -tokenResponse $tokenResponseTeamsBot -intendedMembers @($recordedSubcontractors | Select-Object) -targetGroupId 795cc658-79f7-4476-9e64-d852b10b504c




Stop-Transcript