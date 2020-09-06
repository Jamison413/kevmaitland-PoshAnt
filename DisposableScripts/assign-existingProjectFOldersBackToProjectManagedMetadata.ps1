$sharePointAdmin = "t0-kevin.maitland@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content 'C:\Users\KevMaitland\OneDrive - Anthesis LLC\Desktop\KimbleBot.txt') 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com" -Credentials $adminCreds

$sharePointBotDetails = get-graphAppClientCredentials -appName SharePointBot
$tokenResponseSharePointBot = get-graphTokenResponse -aadAppCreds $sharePointBotDetails
$clientSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,faed84bc-70be-4e35-bfbf-cdab31aeeb99"
$allClientDrives = get-graphDrives -tokenResponse $tokenResponseSharePointBot -siteGraphId $clientSiteId
$allClientDrives | % {
    Add-Member -InputObject $_ -MemberType NoteProperty -Name Name2 -Value $($_.name).Replace("&","").Replace("＆","").Replace("  "," ") -Force
    Add-Member -InputObject $_ -MemberType NoteProperty -Name DriveId -Value $($_.id) -Force
    }

$pnpTermGroup = "Kimble"
$pnpTermSet = "Clients"
$allClientTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes CustomProperties | ? {$_.IsDeprecated -eq $false}
$allClientTerms | % {
    Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteId -Value $_.CustomProperties.NetSuiteId -Force
    Add-Member -InputObject $_ -MemberType NoteProperty -Name Name2 -Value $($_.Name).Replace("&","").Replace("＆","").Replace("  "," ") -Force
    }


$pnpTermGroup = "Kimble"
$pnpTermSet = "Projects"
$allProjTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes CustomProperties | ? {$_.IsDeprecated -eq $false}
$allProjTerms | % {
    Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteId -Value $_.CustomProperties.NetSuiteProjId -Force
    }
$projTermsToCheck = $allProjTerms | ? {$_.LastModifiedDate -gt $lastProcessed -and ![string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteProjId) -and ![string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteClientId)}
#$projTermsToCheck = $allProjTerms | ? {$_.LastModifiedDate -gt $lastProcessed -and ![string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteProjId) -and ![string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteClientId) -and [string]::IsNullOrWhiteSpace($_.CustomProperties.DriveItemId)}

$projTermsToCheck | ? {$reallyNoFolder -notcontains $_} |% {
    $thisProjTerm = $_
    $thisClientTerm = $allClientTerms | ? {$_.NetSuiteId -eq $thisProjTerm.CustomProperties.NetSuiteClientId}
    Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes CustomProperties | ? {$_.CustomProperties.NetSuiteId -eq 4387}

    if(![string]::IsNullOrWhiteSpace($thisClientTerm.CustomProperties.GraphDriveId)){
        $thisClientDrive = get-graphDrives -driveId $thisClientTerm.CustomProperties.GraphDriveId -tokenResponse $tokenResponseSharePointBot
        $rootFolders = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $thisClientDrive.id
        $preExistingProjectFolder = $rootFolders | ? {$_.name -match $(get-kimbleEngagementCodeFromString $thisProjTerm.name)}
        if($preExistingProjectFolder){
            $thisProjTerm.SetCustomProperty("DriveItemId",$preExistingProjectFolder.id)
            try{
                Write-Verbose "`tTrying: `$thisProjTerm.SetCustomProperty(DriveItemId,$($preExistingProjectFolder.id)) [$($thisProjTerm.Name)]"
                $thisProjTerm.Context.ExecuteQuery()
                Write-Verbose "`tSUCCESS!"
                }
            catch{
                Write-Error "Error updating `$thisProjTerm.SetCustomProperty(DriveItemId,$($preExistingProjectFolder.id)) [$($thisProjTerm.Name)]"
                [array]$problems += $thisProjTerm
                }
            }
        else{
            Write-Warning "Could not find pre-existing Project Folder for [$($thisProjTerm.Name)]"
            [array]$noExistingFolder += $thisProjTerm
            }        
        }
    else{
        Write-Warning "Could not find Drive for Project [$($thisProjTerm.Name)]"
        [array]$noDrive += $thisProjTerm
        }
    }

$noExistingFolder | % {
    if($_.Name -eq "P-1000242 BG ECO Delivery 2020"){break}
    [array]$reallyNoFolder += $_
    }
    $reallyNoFolder.Count