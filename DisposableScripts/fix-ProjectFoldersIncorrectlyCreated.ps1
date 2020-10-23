$sharePointBotDetails = get-graphAppClientCredentials -appName SharePointBot
$tokenResponseSharePointBot = get-graphTokenResponse -aadAppCreds $sharePointBotDetails
$clientSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,faed84bc-70be-4e35-bfbf-cdab31aeeb99"
#$supplierSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,9fb8ecd6-c87d-485d-a488-26fd18c62303"
#$devSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,8ba7475f-dad0-4d16-bdf5-4f8787838809"
$allClientDrives = get-graphDrives -tokenResponse $tokenResponseSharePointBot -siteGraphId $clientSiteId

$sharePointAdmin = "kimblebot@anthesisgroup.com"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content "$env:USERPROFILE\Desktop\KimbleBot.txt") 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com" -Credentials $adminCreds

$pnpTermGroup = "Kimble"
$pnpTermSet = "Clients"
$allClientTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes CustomProperties | ? {$_.IsDeprecated -eq $false}



$pnpTermGroup = "Kimble"
$pnpTermSet = "Opportunities"
$allOppTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes CustomProperties | ? {$_.IsDeprecated -eq $false}
$allOppTerms | % {
    Add-Member -InputObject $_ -MemberType NoteProperty -Name ProjectId -Value $_.CustomProperties.NetSuiteProjectId
    }

$pnpTermGroup = "Kimble"
$pnpTermSet = "Projects"
$allProjTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes CustomProperties | ? {$_.IsDeprecated -eq $false -and $(![string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteClientId))}
$allProjTerms | % {
    Add-Member -InputObject $_ -MemberType NoteProperty -Name ProjectId -Value $_.CustomProperties.NetSuiteProjId
    }


$wonOpps = $allOppTerms | ? {![string]::IsNullOrWhiteSpace($_.ProjectId)}
$wonProjs = $allProjTerms |  ? {$allOppTerms.ProjectId -contains $_.ProjectId}
$delta = Compare-Object -ReferenceObject $wonOpps -DifferenceObject $wonProjs -Property ProjectId -IncludeEqual -PassThru
$wonOppsMatched = Compare-Object -ReferenceObject $wonOpps -DifferenceObject $allProjTerms -Property ProjectId -IncludeEqual -ExcludeDifferent -PassThru
$wonOppsWrongFolders = Compare-Object -ReferenceObject $wonOppsMatched -DifferenceObject $wonProjs -Property ProjectId,DriveItemId -PassThru

$problems = @()
$wonProjs | % {
    Write-Host
    $thisProject = $_
    $thisProjCode = $thisProject.Name.Split(" ")[0]
    $thisOpp = $allOppTerms | ? {$_.CustomProperties.NetSuiteProjectId -eq $thisProject.CustomProperties.NetSuiteProjId}
    $thisOppCode = $thisOpp.Name.Split(" ")[0]
    $thisClient = $allClientTerms | ? {$_.CustomProperties.NetSuiteId -eq $thisProject.CustomProperties.NetSuiteClientId}
    $theseRootFolders = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $thisClient.CustomProperties.GraphDriveId -returnWhat Children
    $oppFolderByName = $theseRootFolders | ? {$_.name -match $thisOppCode}
    $oppFolderById = $theseRootFolders | ? {$_.id -match $thisOpp.CustomProperties.DriveItemId}
    if($oppFolderById.Count -eq $theseRootFolders.Count){$oppFolderById = $null}
    $projFolderByName = $theseRootFolders | ? {$_.name -match $thisProjCode}
    $projFolderById = $theseRootFolders | ? {$_.id -match $thisProject.CustomProperties.DriveItemId}
    if($projFolderById.Count -eq $theseRootFolders.Count){$projFolderById = $null}
    Write-Host "Project [$($thisProject.Name)][$($thisProject.Id)] for Client [$($thisClient.Name)][$($thisProject.CustomProperties.NetSuiteClientId)]"
    $errorValue = [ordered]@{
        missingClientId = $false
        missingClientTerm = $false
        missingClientDriveId = $false
        missingOppDriveId = $false
        missingProjDriveId = $false
        separateFolders = $false
        oppFolderMissingByName = $false
        oppFolderMissingById = $false
        projFolderMissingByName = $false
        projFolderMissingById = $false
        }
    if(![string]::IsNullOrWhiteSpace($thisProject.CustomProperties.NetSuiteClientId)){}else{Write-Host -ForegroundColor Magenta "`tNetSuiteClientId missing for Proj [$($thisProject.Name)][$($thisProject.Id)]";$errorValue["missingClientId"] = $true}
    if($thisClient){}else{Write-Host -ForegroundColor Magenta "`tClient Term with missing NetSuiteClientId [$($thisProject.CustomProperties.NetSuiteClientId)] for Project [$($thisProject.Name)][$($thisProject.Id)]";$errorValue["missingClientTerm"] = $true}
    if(![string]::IsNullOrWhiteSpace($thisClient.CustomProperties.GraphDriveId)){}else{Write-Host -ForegroundColor Magenta "`tGraphDriveId missing for Client [$($thisClient.Name)][$($thisClient.Id)]";$errorValue["missingClientDriveId"] = $true}
    if(![string]::IsNullOrWhiteSpace($thisOpp.CustomProperties.DriveItemId)){}else{Write-Host -ForegroundColor Magenta "`tDriveItemId missing for Opp [$($thisOpp.Name)][$($thisOpp.Id)]";$errorValue["missingOppDriveId"] = $true}
    if(![string]::IsNullOrWhiteSpace($thisProject.CustomProperties.GraphDriveId)){}else{Write-Host -ForegroundColor Magenta "`tDriveItemId missing for Project [$($thisProject.Name)][$($thisProject.Id)]";$errorValue["missingProjDriveId"] = $true}
    if($thisOpp.CustomProperties.DriveItemId -eq $thisProject.CustomProperties.DriveItemId -and ![string]::IsNullOrWhiteSpace($thisProject.CustomProperties.DriveItemId)){Write-Host "`tOpp & Proj folder IDs match [$($thisOpp.CustomProperties.DriveItemId)]"}else{Write-Host -ForegroundColor Magenta "`tOpp & Proj folder IDs DO NOT match [$($thisOpp.CustomProperties.DriveItemId)] [$($thisProject.CustomProperties.DriveItemId)]";$errorValue["separateFolders"] = $true}
    if($oppFolderByName){Write-Host "`t`tOPP folder is present (matched by name) [$($oppFolderByName.name)][$($oppFolderByName.id)]"}else{Write-Host -ForegroundColor Magenta "`t`tOPP folder is NOT present (matched by name)";$errorValue["oppFolderMissingByName"] = $true}
    if($oppFolderById){Write-Host "`t`tOPP folder is present (matched by id) [$($oppFolderById.name)][$($oppFolderById.id)]"}else{Write-Host -ForegroundColor Magenta "`t`tOPP folder is NOT present (matched by id)";$errorValue["oppFolderMissingById"] = $true}
    if($projFolderByName){Write-Host "`t`tPROJ folder is present (matched by name) [$($projFolderByName.name)][$($projFolderByName.id)]"}else{Write-Host -ForegroundColor Magenta "`t`tPROJ folder is NOT present (matched by name)";$errorValue["projFolderMissingByName"] = $true}
    if($projFolderById){Write-Host "`t`tPROJ folder is present (matched by id) [$($projFolderById.name)][$($projFolderById.id)]"}else{Write-Host -ForegroundColor Magenta "`t`tPROJ folder is NOT present (matched by id)";$errorValue["projFolderMissingById"] = $true}
    [array]$problems += New-Object psobject -ArgumentList @{Client=$thisClient;Opp=$thisOpp;Proj=$thisProject;Issue=$errorValue}

    #If no ClientTerm
        #FFS

    #If $oppFolderByName.Id -ne $projFolderByName.Id > Check if folders empty
        #Both empty: Keep Project, set $opp and $proj.CustomProperties.DriveItemId to kept value
        #One emtpy: Keep non-empty, set $opp and $proj.CustomProperties.DriveItemId to kept value
        #Neither empty: Error
    
    if($oppFolderByName -and $projFolderByName -and $oppFolderByName.id -ne $projFolderByName.id){
        if($oppFolderByName.size -eq 0 -and $projFolderByName.size -eq 0){
            $thisOpp.SetCustomProperty("DriveItemId",$projFolderByName.id)
            $thisOpp.Context.ExecuteQuery()
            $thisProject.SetCustomProperty("DriveItemId",$projFolderByName.id)
            $thisProject.Context.ExecuteQuery()
            delete-graphDriveItem -tokenResponse $tokenResponseSharePointBot -graphDriveId $thisClient.CustomProperties.GraphDriveId -graphDriveItemId $oppFolderByName.id -Verbose
            }
        elseif($oppFolderByName.size -eq 0 -or $projFolderByName.size -eq 0){
            if($oppFolderByName.size -eq 0){$keepMe = $projFolderByName; $binMe = $oppFolderByName}
            if($projFolderByName.size -eq 0){$keepMe = $oppFolderByName; $binMe = $projFolderByName}
            $thisOpp.SetCustomProperty("DriveItemId",$keepMe.id)
            $thisOpp.Context.ExecuteQuery()
            $thisProject.SetCustomProperty("DriveItemId",$keepMe.id)
            $thisProject.Context.ExecuteQuery()
            delete-graphDriveItem -tokenResponse $tokenResponseSharePointBot -graphDriveId $thisClient.CustomProperties.GraphDriveId -graphDriveItemId $binMe.id -Verbose
            }
        else{}
        }
    }

