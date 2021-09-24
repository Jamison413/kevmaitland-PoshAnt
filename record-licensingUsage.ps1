start-transcriptLog -thisScriptName "record-licensingUsage"

$teamBotDetails = get-graphAppClientCredentials -appName TeamsBot
$tokenResponseTeamBot = get-graphTokenResponse -aadAppCreds $teamBotDetails

$itSite = get-graphSite -tokenResponse $tokenResponseTeamBot -serverRelativeUrl "/teams/IT_Team_All_365" -Verbose
$licensingList = get-graphList -tokenResponse $tokenResponseTeamBot -graphSiteId $itSite.id -listName "365 Licensing Logs" -Verbose
$licensingListItem = get-graphListItems -tokenResponse $tokenResponseTeamBot -graphSiteId $itSite.id -listId $licensingList.id -expandAllFields -filterId 10000 -Verbose

$allLicensedUsers = get-graphUsers -tokenResponse $tokenResponseTeamBot -selectAllProperties -filterLicensedUsers


$allLicensedUsers | % {
    $thisUser = $_
    write-host -f Yellow "[$($thisUser.displayName)]"
    $thisUser.assignedLicenses | % {
        $thisLicense = $_ #;break}}
        write-host -f DarkYellow "[$(get-microsoftProductInfo -getType intY -fromType GUID -fromValue $thisLicense.skuId -Verbose:$VerbosePreference)]"
        $licenseRecordHash = @{
            Title=$thisUser.displayName
            LicenseName=$(get-microsoftProductInfo -getType intY -fromType GUID -fromValue $thisLicense.skuId -Verbose:$VerbosePreference)
            FriendlyLicenseName=$(get-microsoftProductInfo -getType FriendlyName -fromType GUID -fromValue $thisLicense.skuId -Verbose:$VerbosePreference)
            LicenseCostUSD=$(get-microsoftProductInfo -getType Cost -fromType GUID -fromValue $thisLicense.skuId -Verbose:$VerbosePreference)
            BusinessUnit=$thisUser.companyName
            Country=$thisUser.country
            UserPrincipalName=$thisUser.userPrincipalName
            ContractType=$thisUser.anthesisgroup_employeeInfo.contractType
            TimeStamp=$(Get-Date -Format "yyyy-MM-dd")
            }
        do{
            try{
                $newLicenseRecord = new-graphListItem -tokenResponse $tokenResponseTeamBot -graphSiteId $itSite.id -listId $licensingList.id -listItemFieldValuesHash $licenseRecordHash -Verbose
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

Stop-Transcript