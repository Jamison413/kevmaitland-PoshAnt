Import-Module _PS_Library_GeneralFunctionality

$teamBotDetails = get-graphAppClientCredentials -appName TeamsBot
$tokenResponse = get-graphTokenResponse -aadAppCreds $teamBotDetails

$coworkerEmails = "emily.pressey@anthesisgroup.com, Andrew.ost@anthesisgroup.com, george.gaisford@anthesisgroup.com, Rae.Victorio@anthesisgroup.com"
$coworkerEmails = convertTo-arrayOfEmailAddresses -blockOfText $coworkerEmails


$master365GroupList = @()
$allOtherGroups = @()
ForEach($coworker in $coworkerEmails){

$userGroups = get-graphUserGroupMembership -tokenResponse $tokenResponse -userUpn $coworker -Verbose
Write-Host "365 Groups for: $($coworker)":
$userGroups.displayName

    ForEach($foundGroup in $userGroups){
        
        If($foundGroup.groupTypes -eq "unified"){
            If(($foundGroup.displayName -notcontains $master365GroupList.displayName) -and (($foundGroup.displayName).Split(" ")[0] -ne "All")){
                $foundGroup | Add-Member -MemberType NoteProperty -Name "User" -Value $coworker -Force
                $foundGroup | Add-Member -MemberType NoteProperty -Name "What it is" -Value "Unified"  -Force
                $master365GroupList += $foundGroup
                }
            Else{
            Write-Host "Group is dupe" -ForegroundColor Red
            }
        }
        Else{
            If(($foundGroup.displayName -notcontains $allOtherGroups.displayName) -and (($foundGroup.displayName).Split(" ")[0] -ne "All")){
            $foundGroup | Add-Member -MemberType NoteProperty -Name "User" -Value $coworker -Force
            $foundGroup | Add-Member -MemberType NoteProperty -Name "What it is" -Value "Security"  -Force
            $allOtherGroups += $foundGroup
            }
            Else{
            Write-Host "Group is dupe" -ForegroundColor Red
            }

        }

    }
}


$finalMasterList365 = $master365GroupList | Select-Object displayName,User,'What it is',id
$finalMasterListgroups = $allOtherGroups | Select-Object displayName,User,'What it is',id

$final = @()
$final += $finalMasterList365
$final += $finalMasterListgroups

$final | export-csv -Path 'C:\groupdata.csv'

