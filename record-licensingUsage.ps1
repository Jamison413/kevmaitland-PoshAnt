$teamBotDetails = get-graphAppClientCredentials -appName TeamsBot
$tokenResponse = get-graphTokenResponse -aadAppCreds $teamBotDetails

$itSite = get-graphSite -tokenResponse $tokenResponse -serverRelativeUrl "/teams/IT_Team_All_365" -Verbose
$licensingList = get-graphList -tokenResponse $tokenResponse -graphSiteId $itSite.id -listName "365 Licensing Logs" -Verbose
$licensingListItem = get-graphListItems -tokenResponse $tokenResponse -graphSiteId $itSite.id -listId $licensingList.id -expandAllFields -filterId 10000 -Verbose

$allLicensedUsers = get-graphUsersWithEmployeeInfoExtensions -tokenResponse $tokenResponse -selectAllProperties -filterNone


$allLicensedUsers | % {
    $thisUser = $_
    write-host -f Yellow "[$($thisUser.displayName)]"
    $thisUser.assignedLicenses | % {
        $thisLicense = $_ #;break}}
        write-host -f DarkYellow "[$(get-microsoftProductInfo -getType intY -fromType GUID -fromValue $thisLicense.skuId -Verbose:$VerbosePreference)]"
        $licenseRecordHash = @{
            Title=$thisUser.displayName
            LicenseName=$(get-microsoftProductInfo -getType intY -fromType GUID -fromValue $thisLicense.skuId -Verbose:$VerbosePreference)
            BusinessUnit=$thisUser.companyName
            Country=$thisUser.country
            UserPrincipalName=$thisUser.userPrincipalName
            ContractType=$thisUser.anthesisgroup_employeeInfo.contractType
            TimeStamp=$(Get-Date -Format "yyyy-MM-dd")
            }
        do{
            try{
                $newLicenseRecord = new-graphListItem -tokenResponse $tokenResponse -graphSiteId $itSite.id -listId $licensingList.id -listItemFieldValuesHash $licenseRecordHash -Verbose
                $retry = $false
                }
            catch{
                if($_.ErrorDetails.Message -match "TooManyRequests"){
                    Write-Warning "TooManyRequests"
                    $retry = $true
                    Start-Sleep -Seconds 5
                    }
                }
            }
        while ($retry -eq $true)

        }
    
    }