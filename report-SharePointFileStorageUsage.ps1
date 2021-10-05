start-transcriptLog -thisScriptName "report-SharePointFileStorageUsage"

$sharePointBotDetails = get-graphAppClientCredentials -appName SharePointBot
$tokenResponseSharePointBot = get-graphTokenResponse -aadAppCreds $sharePointBotDetails -grant_type client_credentials
$sitesToProcess = @("/clients")

$sitesToProcess | % {
    $thisSite = get-graphSite -tokenResponse $tokenResponseSharePointBot -serverRelativeUrl $_
    Write-Host "Processing Site [$($thisSite.name)][$($thisSite.id)][$($thisSite.webUrl)]"
    $theseDrives = get-graphDrives -tokenResponse $tokenResponseSharePointBot -siteGraphId $thisSite.id
    $theseDrives | % {
        $tokenResponseSharePointBot = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponseSharePointBot -renewTokenExpiringInSeconds 60 -aadAppCreds $sharePointBotDetails
        $thisDrive = $_
        Write-Host "`tProcessing Drive [$($thisDrive.name)][$($thisDrive.id)][$($thisDrive.webUrl)]"
        $theseDriveItems = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $thisDrive.id -returnWhat Children -includePreviousVersions
        Write-Host "`t`tProcessing [$($theseDriveItems.Count)] DriveItems"
        $output = @($null)*$theseDriveItems.Count
        for ($i=0;$i -lt $output.Count; $i++){
            $output[$i] = [pscustomobject][ordered]@{
                Name=$theseDriveItems[$i].name
                DocLib=$thisDrive.name
                Type=$(
                    if(![string]::IsNullOrWhiteSpace($theseDriveItems[$i].folder)){"folder"}
                    elseif(![string]::IsNullOrWhiteSpace($theseDriveItems[$i].file)){"file"}
                    else{"unknown"})
                Size=$theseDriveItems[$i].size
                WebUrl=$theseDriveItems[$i].webUrl
                PreviousVersionCount=@($theseDriveItems[$i].PreviousVersions | Select-Object).Count
                PreviousVersionSize=$(
                    if($($theseDriveItems[$i].PreviousVersions | Select-Object).Count -gt 0){$($($theseDriveItems[$i].PreviousVersions | Measure-Object -Property size -Sum).Sum)}
                    else{0}
                    )
                }
            }

        $output | Sort-Object WebUrl | Select-Object * | Export-Csv  -Path "$env:USERPROFILE\Downloads\$($thisSite.name)_$((Get-Date -f u).Split(" ")[0]).csv" -NoTypeInformation -Force -Append
        #Write-Host "`t`tOutput written to [$("$env:USERPROFILE\Downloads\$($thisSite.name)_$($thisDrive.name)_$((Get-Date -f u).Split(" ")[0]).csv")]"
        Write-Host "`t`tOutput written to [$("$env:USERPROFILE\Downloads\$($thisSite.name)_$($thisDrive.name)_$((Get-Date -f u).Split(" ")[0]).csv")]"
        Write-Host
        }
    }


Stop-Transcript