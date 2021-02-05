Remove-Module _REST_Library_NetSuite;Import-Module _REST_Library_NetSuite
Remove-Module _PS_Library_GeneralFunctionality;Import-Module _PS_Library_GeneralFunctionality
Remove-Module _PS_Library_Graph;Import-Module _PS_Library_Graph
$tokenResponseSharePointBot = get-graphTokenResponse -aadAppCreds $appCredsSharePointBot


function get-duplicatesByProperty(){
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
            [psobject[]]$arrayOfTerms 
        ,[Parameter(Mandatory = $true, Position = 1)]
            [String]$propertyToTest 
        ,[Parameter(Mandatory = $false, Position = 2)]
            [string[]]$alternativeFeedbackProperties
        ,[Parameter(Mandatory = $false, Position = 3)]
            [switch]$ignoreNullOrEmpty
        ,[Parameter(Mandatory = $false, Position = 4)]
            [switch]$alsoReturnOriginals
        )
    $allDuplicateTerms = @()
    $dupTermsByProperty = $arrayOfTerms | Group-Object -Property {$_.$propertyToTest} | ? {$_.Count -gt 1}
    if($ignoreNullOrEmpty){$dupTermsByProperty = $dupTermsByProperty | ? {![string]::IsNullOrWhiteSpace($_.Name)}}
    $dupTermsByProperty | % {
        $originalTerm = $_.Group | Sort-Object {$_.CreatedDate} -Descending | select -Last 1
        [array]$duplicateTerms = $_.Group | Sort-Object {$_.CreatedDate} -Descending | select -First $($_.Group.Count-1)
        if(($arrayOfTerms[0].GetType()).BaseType -match "TermSetItem"){Write-Host -ForegroundColor Yellow "[$($originalTerm.TermSet.Group.Name)][$($originalTerm.TermSet.Name)][$($originalTerm.Name)][$($originalTerm.$propertyToTest)] (original Term) duplicated [$($duplicateTerms.Count)] times as:"}
        elseif($alternativeFeedbackProperties){
            $alternativeOutput = ""
            $alternativeFeedbackProperties | % {$alternativeOutput += "[$($originalTerm.$_)]"}
            Write-Host -ForegroundColor Yellow "$($alternativeOutput) (original) duplicated [$($duplicateTerms.Count)] times as:"
            }
        else{Write-Host -ForegroundColor Yellow "[$($originalTerm.Name)][$($originalTerm.Id)][$($originalTerm.$propertyToTest)] (original) duplicated [$($propertyToTest)] [$($duplicateTerms.Count)] times as:"}
        $duplicateTerms | % {
            $thisDupTerm = $_
            if($alternativeFeedbackProperties){
                $alternativeOutput = ""
                $alternativeFeedbackProperties | % {$alternativeOutput += "[$($thisDupTerm.$_)]"}
                Write-Host -ForegroundColor DarkYellow "`t$($alternativeOutput)"
                }
            else{Write-Host -ForegroundColor DarkYellow "`t[$($_.Name)][$($_.Id)][$($_.$propertyToTest)]"}
            }
        $allDuplicateTerms += $duplicateTerms
        }
    if($alsoReturnOriginals){$dupTermsByProperty}
    else{$allDuplicateTerms}
    }
#region Terms with missing values
$clientsWithNoId = $allClientTerms | ? {[string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteId)}
$clientsWithNoId | % {
    $thisDuffClient = $_
    Write-Host -f Yellow "Client [$($thisDuffClient.UniversalClientName)] has no Id"
    #Remove-PnPTerm -Identity $thisDuffClient.Id
    }

$oppsWithNoId = $allOppTerms |  ? {[string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteOppId)}
$oppsWithNoId | % {
    $thisDuffOpp = $_
    Write-Host -f Cyan "Opp [$($thisDuffOpp.UniversalOppName)][$($thisDuffOpp.Id)] has no OppId"
    #Remove-PnPTerm -Identity $thisDuffOpp.Id
    }
$oppsWithNoClient = $allOppTerms |  ? {[string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteClientId)}
$oppsWithNoClient | % {
    $thisDuffOpp = $_
    Write-Host -f DarkCyan "Opp [$($thisDuffOpp.UniversalOppName)][$($thisDuffOpp.id)] has no Client (CreatedDate = [$($thisDuffOpp.CreatedDate)], LastModifiedDate = [$($thisDuffOpp.LastModifiedDate)]"
    #Remove-PnPTerm -Identity $thisDuffOpp.Id
    }
$duffProjs = $allProjTerms |  ? {[string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteProjId)}
$duffProjs | % {
    $thisDuffProj = $_
    Write-Host -f Magenta "Proj [$($thisDuffProj.UniversalProjName)][$($thisDuffProj.id)] has no ProjId"
    #Remove-PnPTerm -Identity $thisDuffProj.Id
    }
$projsWithNoClient = $allProjTerms |  ? {[string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteClientId)}
$projsWithNoClient | % {
    $thisDuffOpp = $_
    Write-Host -f DarkMagenta "Proj [$($thisDuffOpp.UniversalProjName)][$($thisDuffOpp.id)] has no Client (CreatedDate = [$($thisDuffOpp.CreatedDate)], LastModifiedDate = [$($thisDuffOpp.LastModifiedDate)]"
    #Remove-PnPTerm -Identity $thisDuffOpp.Id
    }
#endregion 
#region Terms with duplicate values
$dupClientsByNetSuiteId  = get-duplicatesByProperty -arrayOfTerms $allClientTerms -propertyToTest NetSuiteClientId -alsoReturnOriginals
$dupClientsByDriveItemId = get-duplicatesByProperty -arrayOfTerms $allClientTerms -propertyToTest DriveClientId -alsoReturnOriginals -ignoreNullOrEmpty
$dupOppsByOppId          = get-duplicatesByProperty -arrayOfTerms $allOppTerms -propertyToTest NetSuiteOppId -alsoReturnOriginals
$dupOppsByOppCode        = get-duplicatesByProperty -arrayOfTerms $allOppTerms -propertyToTest UniversalOppCode -alsoReturnOriginals
$dupOppsByDriveItemId    = get-duplicatesByProperty -arrayOfTerms $allOppTerms -propertyToTest DriveItemId -alsoReturnOriginals -ignoreNullOrEmpty
$dupProjsByProjId        = get-duplicatesByProperty -arrayOfTerms $allProjTerms -propertyToTest NetSuiteProjectId -alsoReturnOriginals
$dupProjsByProjCode      = get-duplicatesByProperty -arrayOfTerms $allProjTerms -propertyToTest UniversalProjCode -alsoReturnOriginals
$dupProjsByDriveItemId   = get-duplicatesByProperty -arrayOfTerms $allProjTerms -propertyToTest DriveItemId -alsoReturnOriginals -ignoreNullOrEmpty
#endregion
#region Folders with duplicate values
$dupClientDrivesByName  = get-duplicatesByProperty -arrayOfTerms $allClientDrives -propertyToTest Name -alsoReturnOriginals
$dupOppFoldersByCode    = get-duplicatesByProperty -arrayOfTerms $driveItemsOppFolders -propertyToTest UniversalOppCode -alsoReturnOriginals -alternativeFeedbackProperties @("UniversalOppName","DriveClientName","DriveItemId","DriveClientId","DriveItemSize")
$dupProjFoldersByCode   = get-duplicatesByProperty -arrayOfTerms $driveItemsProjFolders -propertyToTest UniversalProjCode -alsoReturnOriginals -alternativeFeedbackProperties @("UniversalProjName","DriveClientName","DriveItemId","DriveClientId","DriveItemSize")
#endregion

#region Fix missing values
Write-Host "Fixing [$($oppsWithNoClient.Count)] Opps with no Client"
$correspondingNetOpps = compare-object -ReferenceObject $netSuiteOppsToCheck -DifferenceObject $oppsWithNoClient -Property NetSuiteOppId -PassThru -IncludeEqual -ExcludeDifferent
$oppsWithNoClientComparison = process-comparison -subsetOfNetObjects $correspondingNetOpps -allTermObjects $oppsWithNoClient -idInCommon NetSuiteOppId -propertyToTest id -validate -Verbose
$termOppsMissingTheirId = $oppsWithNoClientComparison["=>"]
$netOppsWithTheMIssingId = $oppsWithNoClientComparison["<="]
for($j=0;$j -lt $termOppsMissingTheirId.Count;$j++){
    if($termOppsMissingTheirId[$j].NetSuiteOppId -eq $netOppsWithTheMIssingId[$j].NetSuiteOppId){ #This should always be the case anyway, but just to avoid breaking things even worse
        Write-Host "Fixing missing NetSuiteClientId [$($netOppsWithTheMIssingId[$j].entity.id)] on [$($termOppsMissingTheirId[$j].UniversalOppName)]"
        $termOppsMissingTheirId[$j].SetCustomProperty("NetSuiteClientId",$netOppsWithTheMIssingId[$j].entity.id)
        $termOppsMissingTheirId[$j].Context.ExecuteQuery()
        }
    else{Write-Warning "Something funky happened aligning [$($termOppsMissingTheirId[$j].UniversalOppName)] with it's Id"}
    }

Write-Host "Fixing [$($projsWithNoClient.Count)] Projs with no Client"
$correspondingNetProjs = compare-object -ReferenceObject $netSuiteProjsToCheck -DifferenceObject $projsWithNoClient -Property NetSuiteProjectId -PassThru -IncludeEqual -ExcludeDifferent
$projsWithNoClientComparison = process-comparison -subsetOfNetObjects $correspondingNetProjs -allTermObjects $projsWithNoClient -idInCommon NetSuiteProjectId -propertyToTest id -validate -Verbose
$termProjsMissingTheirId = $projsWithNoClientComparison["=>"]
$netProjsWithTheMIssingId = $projsWithNoClientComparison["<="]
for($j=0;$j -lt $termProjsMissingTheirId.Count;$j++){
    if($termProjsMissingTheirId[$j].NetSuiteProjectId -eq $netProjsWithTheMIssingId[$j].NetSuiteProjectId){ #This should always be the case anyway, but just to avoid breaking things even worse
        Write-Host "Fixing missing NetSuiteClientId [$($netProjsWithTheMIssingId[$j].parent.id)] on [$($termProjsMissingTheirId[$j].UniversalProjName)]"
        $termProjsMissingTheirId[$j].SetCustomProperty("NetSuiteClientId",$netProjsWithTheMIssingId[$j].entity.id)
        $termProjsMissingTheirId[$j].Context.ExecuteQuery()
        }
    else{Write-Warning "Something funky happened aligning [$($termProjsMissingTheirId[$j].UniversalProjName)] with it's Id"}
    }

#endregion

#region FixDuplicates
#region Clients
for($i=0; $i -lt $dupClientDrivesByName.Count;$i++){
    $thisGroup = $dupClientDrivesByName[$i]
    $thisClient = $allClientTerms | ? {$_.name -eq $thisGroup.Name}
    $realDrive = get-graphDrives -tokenResponse $tokenResponseSharePointBot -driveId $thisClient.DriveClientId -ErrorAction Stop
    $correctDrive = $thisGroup.Group | ? {$_.DriveClientId -eq $realDrive.id}
    [array]$incorrectDrives = $thisGroup.Group | ? {$_.DriveClientId -ne $realDrive.id}
    Write-Host -f Cyan "Moving [$($incorrectDrives.Count)] Drives Client [$($thisClient.UniversalClientName)]"
    Write-Host -f DarkCyan "`tTo:`t`t[$($correctDrive.webUrl)]"
    $incorrectDrives | % {Write-Host -f DarkMagenta "`tFrom:`t[$($_.webUrl)]"}

    #Be careful runnign this next section as not all Duplicate cleints need merging - some just need renaming, and some just need deleting.
    for($j=0;$j -lt $incorrectDrives.Count;$j++){
        $thisIncorrectDrive = $incorrectDrives[$j]
        $incorrectFolders = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $thisIncorrectDrive.id -returnWhat Children
        $incorrectFolders | ? {$_.size -gt 0} | % {
            $thisIncorrectFolder = $_
            try{
                move-graphDriveItem -tokenResponse $tokenResponseSharePointBot -driveGraphIdSource $thisIncorrectFolder.parentReference.driveId -itemGraphIdSource $thisIncorrectFolder.id -driveGraphIdDestination $correctDrive.DriveClientId -parentItemGraphIdDestination "root" -ErrorAction Stop
                }
            catch{
                if($_.exception -match "409" -or $_.InnerException -match "409"){
                    if($thisIncorrectSubfolder.size -eq 0){}#DoNothing
                    else{#Try again with a new name
                        try{
                            move-graphDriveItem -tokenResponse $tokenResponseSharePointBot -driveGraphIdSource $thisIncorrectFolder.parentReference.driveId -itemGraphIdSource $thisIncorrectFolder.id -driveGraphIdDestination $correctDrive.DriveClientId -parentItemGraphIdDestination "root" -ErrorAction Stop -newItemName "$($thisIncorrectFolder.name)_fromOpp$j"
                            }
                        catch{
                            Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)
                            }
                        }
                    }
                elseif($_.exception -match "404" -or $_.InnerException -match "404"){}#Do nothing, we're probably just re-runnign the same set of folders if somethign was locked
                elseif($_.exception -match "423" -or $_.InnerException -match "423"){Write-Warning "Folder [$($thisIncorrectFolder.name)][$($thisIncorrectFolder.webUrl)] is locked. Try again later."}
                else{
                    Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)
                    }
                }
            }
            ###
        Write-Host "`t`tChecking:  $($thisIncorrectFolder.webUrl)"
        try{
            $updatedIncorrectedDrive = get-graphDrives -tokenResponse $tokenResponseSharePointBot -driveId $thisIncorrectDrive.id -ErrorAction Stop
            if($updatedIncorrectedDrive.quota.used -eq 0){
                Write-Host -ForegroundColor Yellow "`t`t`tEmpty - removing it. Check  $($correctDrive.webUrl)"
                $list = get-graphList -tokenResponse $tokenResponseSharePointBot -graphDriveId $updatedIncorrectedDrive.id -ErrorAction Stop
                #invoke-graphDelete -tokenResponse $tokenResponse -graphQuery "/sites/$($list.parentReference.siteId)/lists/$($list.id)" -ErrorAction Stop
                if(![string]::IsNullOrWhiteSpace($list.id)){
                    Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/clients" -Credentials $adminCreds -ErrorAction Stop
                    Remove-PnPList -Identity $list.id -Recycle -Force 
                    Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com" -Credentials $adminCreds
                    }
                else{Write-Host -f DarkRed "List not found for [$($updatedIncorrectedDrive.name)][$($updatedIncorrectedDrive.id)]"}
                }
            }
        catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
        }
    }
#endregion
#region Opportunities
for($j=0;$j -lt $dupOppFoldersByCode.Count;$j++){
    $thisGroup = $dupOppFoldersByCode[$j]
    $thisOpp = $allOppTerms | ? {$_.name -match $thisgroup.Name}
    $thisOpp = set-standardisedClientDriveProperties -rawOppOrProjTerm $thisOpp -allClientTerms $allClientTerms
    try{
        $realDriveItem = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $thisOpp.DriveClientId -itemGraphId $thisOpp.DriveItemId -returnWhat Item -ErrorAction Stop
        $correctDestinationFolder = $thisGroup.Group | ? {$_.DriveItemId -eq $realDriveItem.id}
        $correctDestinationSubfolders = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $correctDestinationFolder.DriveClientId -itemGraphId $correctDestinationFolder.DriveItemId -returnWhat Children -ErrorAction Stop
        [array]$incorrectFolders  = $thisGroup.Group | ? {$_.DriveItemId -ne $realDriveItem.id}
        Write-Host -f Cyan "Moving [$($incorrectFolders.Count)] folders for Project [$($thisOpp.UniversalOppName)][$($thisOpp.UniversalClientName)]"
        Write-Host -f DarkCyan "`tTo:`t`t[$($correctDestinationFolder.DriveItemUrl)]"
        $incorrectFolders | % {Write-Host -f DarkMagenta "`tFrom:`t[$($_.DriveItemUrl)]"}

        if($correctDestinationSubfolders.Count -ge 7){
            $correctDestinationSubfolders | % {
                if($_.size -eq 0){delete-graphDriveItem -tokenResponse $tokenResponseSharePointBot -graphDriveId $correctDestinationFolder.DriveClientId -graphDriveItemId $_.id}
                }
            }
        for($i=0;$i -lt $incorrectFolders.Count;$i++){
            $thisIncorrectFolder = $incorrectFolders[$i]
            $incorrectSubfolders = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $thisIncorrectFolder.DriveClientId -itemGraphId $thisIncorrectFolder.DriveItemId -returnWhat Children
            $incorrectSubfolders | % {
                $thisIncorrectSubfolder = $_
                try{
                    move-graphDriveItem -tokenResponse $tokenResponseSharePointBot -driveGraphIdSource $thisIncorrectFolder.DriveClientId -itemGraphIdSource $thisIncorrectSubfolder.id -driveGraphIdDestination $correctDestinationFolder.DriveClientId -parentItemGraphIdDestination $correctDestinationFolder.DriveItemId -ErrorAction Stop
                    }
                catch{
                                                            if($_.exception -match "409" -or $_.InnerException -match "409"){
                    if($thisIncorrectSubfolder.size -eq 0){}#DoNothing
                    else{#Try again with a new name
                        try{
                            move-graphDriveItem -tokenResponse $tokenResponseSharePointBot -driveGraphIdSource $thisIncorrectFolder.DriveClientId -itemGraphIdSource $thisIncorrectSubfolder.id -driveGraphIdDestination $correctDestinationFolder.DriveClientId -parentItemGraphIdDestination $correctDestinationFolder.DriveItemId -ErrorAction Stop -newItemName "$($thisIncorrectSubfolder.name)_fromOpp$i"
                            }
                        catch{
                            Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)
                            }
                        }
                    }
                    elseif($_.exception -match "404" -or $_.InnerException -match "404"){}#Do nothing, we're probably just re-runnign the same set of folders if somethign was locked
                    elseif($_.exception -match "423" -or $_.InnerException -match "423"){Write-Warning "Subfolder [$($thisIncorrectSubfolder.name)][$($thisIncorrectSubfolder.webUrl)] is locked. Try again later."}
                    else{
                        Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)
                        }
                    }
                }
            Write-Host "`t`tChecking:  $($thisIncorrectFolder.DriveItemUrl)"
            try{
                $updatedIncorrectedFolder = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $thisIncorrectFolder.DriveClientId -itemGraphId $thisIncorrectFolder.DriveItemId -returnWhat Item -ErrorAction Stop
                if($updatedIncorrectedFolder.size -eq 0){
                    Write-Host -ForegroundColor Yellow "`t`t`tEmpty - removing it. Check  $($correctDestinationFolder.DriveItemUrl)"
                    delete-graphDriveItem -tokenResponse $tokenResponseSharePointBot -graphDriveId $thisIncorrectFolder.DriveClientId -graphDriveItemId $thisIncorrectFolder.DriveItemId
                    }
                }
            catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
            }
        }
    catch{
        Write-Host -f Red "Error processing duplicate Project Folders for [$($thisGroup.Name)]"
        Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)
        }
    }
#endregion
#region Projects
for($j=0;$j -lt $dupProjFoldersByCode.Count;$j++){
    $thisGroup = $dupProjFoldersByCode[$j]
    $thisProj = $allProjTerms | ? {$_.name -match $thisgroup.Name}
    $thisProj = set-standardisedClientDriveProperties -rawOppOrProjTerm $thisProj -allClientTerms $allClientTerms
    try{
        $realDriveItem = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $thisProj.DriveClientId -itemGraphId $thisProj.DriveItemId -returnWhat Item -ErrorAction Stop
        $correctDestinationFolder = $thisGroup.Group | ? {$_.DriveItemId -eq $realDriveItem.id}
        $correctDestinationSubfolders = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $correctDestinationFolder.DriveClientId -itemGraphId $correctDestinationFolder.DriveItemId -returnWhat Children -ErrorAction Stop
        [array]$incorrectFolders  = $thisGroup.Group | ? {$_.DriveItemId -ne $realDriveItem.id}
        Write-Host -f Cyan "Moving [$($incorrectFolders.Count)] folders for Project [$($thisProj.UniversalProjName)][$($thisProj.UniversalClientName)]"
        Write-Host -f DarkCyan "`tTo:`t`t[$($correctDestinationFolder.DriveItemUrl)]"
        $incorrectFolders | % {Write-Host -f DarkMagenta "`tFrom:`t[$($_.DriveItemUrl)]"}

        if($correctDestinationSubfolders.Count -ge 7){
            $correctDestinationSubfolders | % {
                if($_.size -eq 0){delete-graphDriveItem -tokenResponse $tokenResponseSharePointBot -graphDriveId $correctDestinationFolder.DriveClientId -graphDriveItemId $_.id}
                }
            }
        for($i=0;$i -lt $incorrectFolders.Count;$i++){
            $thisIncorrectFolder = $incorrectFolders[$i]
            $incorrectSubfolders = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $thisIncorrectFolder.DriveClientId -itemGraphId $thisIncorrectFolder.DriveItemId -returnWhat Children
            $incorrectSubfolders | % {
                $thisIncorrectSubfolder = $_
                try{
                    move-graphDriveItem -tokenResponse $tokenResponseSharePointBot -driveGraphIdSource $thisIncorrectFolder.DriveClientId -itemGraphIdSource $thisIncorrectSubfolder.id -driveGraphIdDestination $correctDestinationFolder.DriveClientId -parentItemGraphIdDestination $correctDestinationFolder.DriveItemId -ErrorAction Stop
                    }
                catch{
                                                            if($_.exception -match "409" -or $_.InnerException -match "409"){
                    if($thisIncorrectSubfolder.size -eq 0){}#DoNothing
                    else{#Try again with a new name
                        try{
                            move-graphDriveItem -tokenResponse $tokenResponseSharePointBot -driveGraphIdSource $thisIncorrectFolder.DriveClientId -itemGraphIdSource $thisIncorrectSubfolder.id -driveGraphIdDestination $correctDestinationFolder.DriveClientId -parentItemGraphIdDestination $correctDestinationFolder.DriveItemId -ErrorAction Stop -newItemName "$($thisIncorrectSubfolder.name)_fromOpp$i"
                            }
                        catch{
                            Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)
                            }
                        }
                    }
                    elseif($_.exception -match "404" -or $_.InnerException -match "404"){}#Do nothing, we're probably just re-runnign the same set of folders if somethign was locked
                    elseif($_.exception -match "423" -or $_.InnerException -match "423"){Write-Warning "Subfolder [$($thisIncorrectSubfolder.name)][$($thisIncorrectSubfolder.webUrl)] is locked. Try again later."}
                    else{
                        Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)
                        }
                    }
                }
            Write-Host "`t`tChecking:  $($thisIncorrectFolder.DriveItemUrl)"
            try{
                $updatedIncorrectedFolder = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $thisIncorrectFolder.DriveClientId -itemGraphId $thisIncorrectFolder.DriveItemId -returnWhat Item -ErrorAction Stop
                if($updatedIncorrectedFolder.size -eq 0){
                    Write-Host -ForegroundColor Yellow "`t`t`tEmpty - removing it. Check  $($correctDestinationFolder.DriveItemUrl)"
                    delete-graphDriveItem -tokenResponse $tokenResponseSharePointBot -graphDriveId $thisIncorrectFolder.DriveClientId -graphDriveItemId $thisIncorrectFolder.DriveItemId
                    }
                }
            catch{Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)}
            }
        }
    catch{
        Write-Host -f Red "Error processing duplicate Project Folders for [$($thisGroup.Name)]"
        Write-Host -ForegroundColor Red $(get-errorSummary -errorToSummarise $_)
        }
    }
#endregion
#endregion

#region Realign Existing Proj folders to Terms
function realign-existingProjFoldersToTerms(){
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
            $tokenResponse 
        ,[Parameter(Mandatory = $true, Position = 1)]
            [Microsoft.SharePoint.Client.Taxonomy.TermSetItem]$clientTerm 
        ,[Parameter(Mandatory = $true, Position = 1)]
            [Microsoft.SharePoint.Client.Taxonomy.TermSetItem[]]$alloppsOrProjs 
        )
    $clientDrive = get-graphDrives -tokenResponse $tokenResponse -driveId $clientTerm.DriveClientId
    $clientDrive = set-standardisedClientDriveProperties -rawClientDrive $clientDrive
    $folders = get-graphDriveItems -tokenResponse $tokenResponse -driveGraphId $clientDrive.id -returnWhat Children
    $folders | % {
        Add-Member -InputObject $_ -MemberType NoteProperty -Name Code -Value $($_.name.Split(" ")[0]) -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name DriveItemId -Value $($_.id) -Force
        }

    $relevantOppsOrProjs = $alloppsOrProjs | ? {$_.CustomProperties.NetSuiteClientId -eq $clientTerm.NetSuiteClientId}
    $relevantOppsOrProjs | % {
        Add-Member -InputObject $_ -MemberType NoteProperty -Name Code -Value $("$($_.UniversalProjCode)$($_.UniversalOppCode)") -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name DriveItemId -Value $($_.CustomProperties.DriveItemId) -Force
        }
    $relevantOppsOrProjs = Compare-Object -ReferenceObject $relevantOppsOrProjs -DifferenceObject $folders -Property Code -ExcludeDifferent -IncludeEqual -PassThru
    $codeAndItemIdComparison = process-comparison -subsetOfNetObjects $relevantOppsOrProjs -allTermObjects $folders -idInCommon Code -propertyToTest DriveItemId -Verbose
    for($i=0; $i -lt $codeAndItemIdComparison["<="].Count; $i++){
        if($codeAndItemIdComparison["<="][$i].Code -eq $codeAndItemIdComparison["=>"][$i].Code){
            Write-Host "Updating DriveItemId from [$($codeAndItemIdComparison["<="][$i].DriveItemId)] to [$($codeAndItemIdComparison["=>"][$i].DriveItemId)] for Term [$($codeAndItemIdComparison["<="][$i].Code)][$($codeAndItemIdComparison["<="][$i].id)] "
            $codeAndItemIdComparison["<="][$i].SetCustomProperty("DriveItemId",$codeAndItemIdComparison["=>"][$i].DriveItemId)
            try{$codeAndItemIdComparison["<="][$i].Context.ExecuteQuery()}
            catch{Write-Host -f Red $(get-errorSummary $_)}
            }
        }
    }
#endregion


#region Remove duff Client folders in Opp/Proj folders
$overCreatedFolders = $driveItemsOppFolders | ? {[int]$_.DriveItemChildCountForFolders -ge 10}
$overCreatedFolders += $driveItemsProjFolders | ? {[int]$_.DriveItemChildCountForFolders -ge 10}
$overCreatedFolders | % {
    $thisPossibleFolder = $_
    $thesePossibleSubfolders = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $thisPossibleFolder.DriveClientId -itemGraphId $thisPossibleFolder.DriveItemId -returnWhat Children
    $thesePossibleSubfolders | ? {@("_NetSuite automatically creates Opportunity & Project folders","_Kimble automatically creates Lead & Project folders","Background","Non-specific BusDev") -contains $_.name} | % {
        if($_.size -eq 0){
            Write-Host "Removing empty folder [$($_.name)][$($_.weburl)]"
            delete-graphDriveItem -tokenResponse $tokenResponseSharePointBot -graphDriveId $_.parentReference.driveId -graphDriveItemId $_.id
            }
        }
    }

#endregion

#region Re-add Standard Project Folders
$partiallyCreatedFolders = $driveItemsOppFolders | ? {[int]$_.DriveItemChildCountForFolders -lt 4}
$partiallyCreatedFolders | % {
    $thisDuffFolder = $_
    $correspondingOpp = $allOppTerms | ? {$_.name -match $thisDuffFolder.DriveItemFirstWord}
    $correspondingOpp = set-standardisedClientDriveProperties -rawOppOrProjTerm $correspondingOpp -allClientTerms $allClientTerms
    if(![string]::IsNullOrWhiteSpace($correspondingOpp.id)){
        new-oppProjFolders -tokenResponse $tokenResponseSharePointBot -oppProjTermWithClientInfo $correspondingOpp
        }
    else{Write-Warning "No Opp for [$($thisDuffFolder.DriveItemFirstWord)]"}
    }

$partiallyCreatedFolders = $driveItemsProjFolders | ? {[int]$_.DriveItemChildCountForFolders -lt 4}
for($i=0;$i -lt $partiallyCreatedFolders.count;$i++) {
    write-progress "Reprovisioning ProjFolders" -Status "[$($i)]/[$($partiallyCreatedFolders.count)]"
    $thisDuffFolder = $partiallyCreatedFolders[$i]
    $correspondingProj = $allProjTerms | ? {$_.name -match $thisDuffFolder.DriveItemFirstWord}
    $correspondingProj = set-standardisedClientDriveProperties -rawOppOrProjTerm $correspondingProj -allClientTerms $allClientTerms
    if(![string]::IsNullOrWhiteSpace($correspondingProj.id)){
        new-oppProjFolders -tokenResponse $tokenResponseSharePointBot -oppProjTermWithClientInfo $correspondingProj
        }
    else{Write-Warning "No Opp for [$($thisDuffFolder.DriveItemFirstWord)]"}
    }
#endregion

#region Change Kimble references to NetSuite
$toplevelfolders | ? {@("_Kimble automatically creates Project folders","_Kimble automatically creates Lead & Project folders","_Kimble automatically creates Lead & Project folders1","_Kimble automatically creates Lead & Project folders2") -contains $_.DriveItemName} | % {
    #set-graphDriveItem -tokenResponse $tokenResponseSharePointBot -driveId $_.DriveClientId -driveItemId $_.DriveItemId -driveItemPropertyHash @{name="_NetSuite automatically creates Opportunity & Project folders"}
    Write-Host "$($_.DriveClientName) $($_.DriveItemName)"
    }
#endregion


#region Weird Sanitised FolderNames (should never see this again)
$sanitisedCodes = $allOppTerms | % {sanitise-forNetsuiteIntegration $_.UniversalOppCode}
$sanitisedCodes += $allProjTerms | % {sanitise-forNetsuiteIntegration $_.UniversalProjCode}
$sanitisedFolders = $topLevelFolders | ? {$sanitisedCodes -contains $_.DriveItemFirstWord.Substring(0,8)}
$sanitisedFolders | ? {$_.DriveItemSize -eq 0} | % {
    Write-Host "Removing [$($_.DriveItemName)][$($_.DriveItemUrl)]"
    #delete-graphDriveItem -tokenResponse $tokenResponseSharePointBot -graphDriveId $_.DriveClientId -graphDriveItemId $_.DriveItemId
    }
#endregion
