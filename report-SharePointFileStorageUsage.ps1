$tokenResponseSharePointBot = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName SharePointBot) -grant_type client_credentials
$sitesToProcess = @("/clients")

$sitesToProcess | % {
    $thisSite = get-graphSite -tokenResponse $tokenResponseSharePointBot -serverRelativeUrl $_
    $theseDrives = get-graphDrives -tokenResponse $tokenResponseSharePointBot -siteGraphId $thisSite.id
    $theseDrives | % {
        $thisDrive = $_
        $theseDriveItems = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $thisDrive.id -returnWhat Children -includePreviousVersions
        $output = @($null)*$theseDriveItems.Count
        for ($i=0;$i -lt $output.Count; $i++){
            $output[$i] = [pscustomobject][ordered]@{
                Name=$theseDriveItems[$i].name
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

        $output | Sort-Object WebUrl | Select-Object * | Export-Csv  -Path "$env:USERPROFILE\Downloads\$($thisSite.name)_$($thisDrive.name)_$((Get-Date -f u).Split(" ")[0]).csv" -NoTypeInformation -Force
        }
    }

