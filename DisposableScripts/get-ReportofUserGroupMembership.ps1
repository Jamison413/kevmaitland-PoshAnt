param(
    [CmdletBinding()]
    [parameter(Mandatory = $true)]
    [ValidateNotNull()]
    [ValidatePattern(".[@].")]
    [string]$coworkerEmails
    )

Import-Module _PS_Library_GeneralFunctionality

$teamBotDetails = get-graphAppClientCredentials -appName UserBot
$tokenResponse = get-graphTokenResponse -aadAppCreds $teamBotDetails


#$coworkerEmails = "example.a@anthesisgroup.com,example.b@anthesisgroup.com"
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
$CurrentDate = Get-Date
$CurrentDate = $CurrentDate.ToString('yyyyMMdd-hhmmss')

$final | export-csv -Path C:\Users\$env:USERNAME\Downloads\CoworkersGroupList_$CurrentDate.csv
