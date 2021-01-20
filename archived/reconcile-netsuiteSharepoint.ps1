


#NetSuiteClient #ClientTerm #ClientDrive
$ClientReconcile = Measure-Command {
    $netQuery =  "?q=companyName CONTAIN_NOT `"Anthesis`"" #Excludes any Companies with "Anthesis" in the companyName
    $netQuery += " AND companyName CONTAIN_NOT `"intercompany project`"" #Excludes any Companies with "(intercompany project)" in the companyName
    $netQuery += " AND companyName START_WITH_NOT `"x `"" #Excludes any Companies that begin with "x " in the companyName
    $netQuery += " AND entityStatus ANY_OF_NOT [6, 7]" #Excludes LEAD-Unqualified and LEAD-Qualified (https://XXX.app.netsuite.com/app/crm/sales/customerstatuslist.nl?whence=)
    [array]$allNetSuiteClients = get-netSuiteClientsFromNetSuite -query $netQuery -netsuiteParameters $(get-netSuiteParameters -connectTo Production)
    @($allNetSuiteClients | Select-Object) | % {
        Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteClientId -Value $($_.Id) -Force
        }

    $sharePointAdmin = "kimblebot@anthesisgroup.com"
    #convertTo-localisedSecureString "KimbleBotPasswordHere"
    $sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\KimbleBot.txt) 
    $adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
    Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com" -Credentials $adminCreds
    $pnpTermGroup = "Kimble"
    $pnpTermSet = "Clients"
    $allClientTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes CustomProperties | ? {$_.IsDeprecated -eq $false}
    $allClientTerms | % {
        Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteClientId -Value $($_.CustomProperties.NetSuiteId) -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name DriveClientId -Value $($_.CustomProperties.GraphDriveId) -Force
        }
    $relevantClientTerms = $allClientTerms | ? {![string]::isnullorwhitespace($_.CustomProperties.NetSuiteId)}

    $sharePointBotDetails = get-graphAppClientCredentials -appName SharePointBot
    $tokenResponseSharePointBot = get-graphTokenResponse -aadAppCreds $sharePointBotDetails
    $clientSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,faed84bc-70be-4e35-bfbf-cdab31aeeb99"
    #$supplierSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,9fb8ecd6-c87d-485d-a488-26fd18c62303"
    #$devSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,8ba7475f-dad0-4d16-bdf5-4f8787838809"
    $allClientDrives = get-graphDrives -tokenResponse $tokenResponseSharePointBot -siteGraphId $clientSiteId
    $allClientDrives | % {
        Add-Member -InputObject $_ -MemberType NoteProperty -Name DriveClientId -Value $($_.id) -Force
        }

    $combinedClients = @($null) * $allNetSuiteClients.Count
    for($i=0; $i -lt $allNetSuiteClients.Count; $i++){
        $combinedClients[$i] = New-Object PSObject -Property @{
            NetSuiteClientId=$allNetSuiteClients[$i].id
            NetSuiteClientName=$allNetSuiteClients[$i].companyName
            TermClientId=$null
            TermClientName=$null
            DriveClientId=$null
            DriveClientName=$null
            DriveClientUrl=$null
            Problems=$null
            }
        }
    for ($i=0; $i -lt $combinedClients.count; $i++){
        write-progress -activity "Processing NetSuite Clients" -Status "[$i/$($combinedClients.count)]" -PercentComplete ($i/ $combinedClients.count *100)
        $thisNetSuiteClient = $combinedClients[$i]
        $correspondingTermClient = Compare-Object -ReferenceObject $relevantClientTerms -DifferenceObject $thisNetSuiteClient -Property NetSuiteClientId -PassThru -IncludeEqual -ExcludeDifferent
        $correspondingTermClient = $allClientTerms | ? {$_.NetSuiteClientId -eq $thisNetSuiteClient.NetSuiteClientId}
        if($correspondingTermClient.Count -gt 1){
            Write-Warning "`t[$($thisNetSuiteClient.NetSuiteClientName)][$($thisNetSuiteClient.NetSuiteClientId)] matches multiple Terms: {$($correspondingTermClient | % {"[$($_.Name)][$($_.id)], "})}"
            }
        elseif([string]::IsNullOrWhiteSpace($correspondingTermClient.id)){
            Write-Warning "`t[$($thisNetSuiteClient.NetSuiteClientName)][$($thisNetSuiteClient.NetSuiteClientId)] matched no Terms"
            }
        else{
            $thisNetSuiteClient.TermClientName = $correspondingTermClient.Name
            $thisNetSuiteClient.TermClientId = $correspondingTermClient.Id
            $correspondingDriveClient = Compare-Object -ReferenceObject $allClientDrives -DifferenceObject $correspondingTermClient -Property DriveClientId -PassThru -IncludeEqual -ExcludeDifferent
            if($correspondingDriveClient.Count -gt 1){
                Write-Warning "`t`t[$($correspondingTermClient.Name)][$($correspondingTermClient.DriveClientId)] matches multiple Drives: {$($correspondingDriveClient | % {"[$($_.Name)][$($_.id)], "})}"
                }
            elseif([string]::IsNullOrWhiteSpace($correspondingDriveClient.id)){
                Write-Warning "`t`t[$($correspondingTermClient.Name)][$($correspondingTermClient.DriveClientId)] matched no Drives"
                }
            else{
                $thisNetSuiteClient.DriveClientId = $correspondingTermClient.DriveClientId
                $thisNetSuiteClient.DriveClientName = $correspondingDriveClient.name
                $thisNetSuiteClient.DriveClientUrl = $correspondingDriveClient.webUrl
                }

            }
    
        }
    $now = $(Get-Date -f FileDateTimeUniversal)
    $combinedClients = $combinedClients | ? {![string]::IsNullOrWhiteSpace($_.NetSuiteClientId)}
    $combinedClients |  % {Export-Csv -InputObject $_ -Path "$env:USERPROFILE\Desktop\NetRec_Clients_$now.csv" -Append -NoTypeInformation -Encoding UTF8}
    }
Write-Host "Client reconcilliation completed in [$($ClientReconcile.TotalMinutes)] minutes"
#NetSuiteOpp #ClientTerm #ClientDrive #OppTerm #ProjTerm #OppFolder #ProjFolder



$oppsReconcile = measure-command {
    [array]$allNetSuiteOpps = get-netSuiteOpportunityFromNetSuite -netsuiteParameters $(get-netSuiteParameters -connectTo Production)
    $allNetSuiteOpps = $allNetSuiteOpps | ? {![string]::IsNullOrWhiteSpace($_.id)}
    @($allNetSuiteOpps | Select-Object) | % {
        Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteClientId -Value $($_.entity.id) -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteOppId -Value $($_.id) -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteProjectId -Value $($_.custbody_project_created.id) -Force
        }
    $pnpTermGroup = "Kimble"
    $pnpTermSet = "Opportunities"
    $allOppTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes CustomProperties | ? {$_.IsDeprecated -eq $false}
    $allOppTerms | % {
        Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteOppId -Value $($_.CustomProperties.NetSuiteOppId) -Force
        #Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteProjectId -Value $($_.CustomProperties.NetSuiteProjectId) -Force
        }

    [array]$allNetSuiteProjs = get-netSuiteProjectFromNetSuite -netsuiteParameters $(get-netSuiteParameters -connectTo Production)
    $allNetSuiteProjs = $allNetSuiteProjs | ? {![string]::IsNullOrWhiteSpace($_.id)}
    $allNetSuiteProjs | % {
        Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteProjectId -Value $($_.id) -Force
        }
    $pnpTermSet = "Projects"
    $allProjTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes CustomProperties | ? {$_.IsDeprecated -eq $false -and $(![string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteClientId))}
    $allProjTerms | % {
        Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteProjectId -Value $($_.CustomProperties.NetSuiteProjId) -Force
        }

    $combinedOpps = @($null) * $allNetSuiteOpps.Count
    for($i=0; $i -lt $allNetSuiteOpps.Count; $i++){
        $combinedOpps[$i] = New-Object PSObject -Property @{
            NetSuiteOppId = $allNetSuiteOpps[$i].Id
            NetSuiteOppLabel = "$($allNetSuiteOpps[$i].tranId) $($allNetSuiteOpps[$i].title)"
            NetSuiteClientId=$allNetSuiteOpps[$i].entity.id
            NetSuiteProjectId = $allNetSuiteOpps[$i].custbody_project_created.id
            NetSuiteProjectName = $null

            NetSuiteClientName=$null
            TermClientId=$null
            TermClientName=$null
            DriveClientId=$null
            DriveClientName=$null
            DriveClientUrl=$null
            
            TermOppId = $null
            TermOppLabel = $null
            TermProjId = $null
            TermProjLabel = $null
            DriveItemOppId = $null
            DriveItemOppName = $null
            DriveItemOppUrl = $null
            DriveItemProjId = $null
            DriveItemProjName = $null
            DriveItemProjUrl = $null
            }
        }
    for ($i=0; $i -lt $combinedOpps.count; $i++){
        write-progress -activity "Processing NetSuite Opps" -Status "[$i/$($combinedOpps.count)]" -PercentComplete ($i/ $combinedOpps.count *100)
        $thisNetSuiteOpp = $combinedOpps[$i]
        $correspondingClient = compare-object $combinedClients $thisNetSuiteOpp -Property NetSuiteClientId -PassThru -IncludeEqual -ExcludeDifferent
        if($correspondingClient.Count -gt 1){
            Write-Warning "`t[$($thisNetSuiteOpp.NetSuiteOppLabel)][$($thisNetSuiteOpp.NetSuiteOppId)] matches multiple Clients: {$($correspondingClient | % {"[$($_.NetSuiteClientName)][$($_.NetSuiteClientId)], "})}"
            $combinedOpps[$i].TermClientId = "Multiple"
            }
        elseif([string]::IsNullOrWhiteSpace($correspondingClient.NetSuiteClientId)){
            Write-Warning "`t[$($thisNetSuiteOpp.NetSuiteOppLabel)][$($thisNetSuiteOpp.NetSuiteOppId)] matched no Clients"
            }
        else{
            $combinedOpps[$i].NetSuiteClientName = $correspondingClient.NetSuiteClientName
            $combinedOpps[$i].TermClientId = $correspondingClient.TermClientId
            $combinedOpps[$i].TermClientName = $correspondingClient.TermClientName
            $combinedOpps[$i].DriveClientId = $correspondingClient.DriveClientId
            $combinedOpps[$i].DriveClientName = $correspondingClient.DriveClientName
            $combinedOpps[$i].DriveClientUrl = $correspondingClient.DriveClientUrl

            $correspondingTermOpp = compare-object $allOppTerms $thisNetSuiteOpp -Property NetSuiteOppId -PassThru -IncludeEqual -ExcludeDifferent
            if($correspondingTermOpp.Count -gt 1){
                Write-Warning "`t[$($thisNetSuiteOpp.NetSuiteOppLabel)][$($thisNetSuiteOpp.NetSuiteOppId)] matches multiple OppTerms: {$($correspondingTermOpp | % {"[$($_.Name)][$($_.id)], "})}"
                $combinedOpps[$i].TermOppId = "Multiple"
                }
            elseif([string]::IsNullOrWhiteSpace($correspondingTermOpp.id)){
                Write-Warning "`t[$($thisNetSuiteOpp.NetSuiteOppLabel)][$($thisNetSuiteOpp.NetSuiteOppId)] matched no OppTerms"
                }
            else{
                $combinedOpps[$i].TermOppId = $correspondingTermOpp.id
                $combinedOpps[$i].TermOppLabel = $correspondingTermOpp.Name
                $combinedOpps[$i].DriveItemOppId = $correspondingTermOpp.CustomProperties.DriveItemId

                $tokenResponseSharePointBot = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponseSharePointBot -renewTokenExpiringInSeconds 60 -aadAppCreds $sharePointBotDetails
                $correspondingDriveItemOpp = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $combinedOpps[$i].DriveClientId -itemGraphId $combinedOpps[$i].DriveItemOppId -returnWhat Item
                $combinedOpps[$i].DriveItemOppName = $correspondingDriveItemOpp.name
                $combinedOpps[$i].DriveItemOppUrl = $correspondingDriveItemOpp.webUrl
                }

            $correspondingTermProj = compare-object $allProjTerms $thisNetSuiteOpp -Property NetSuiteProjectId -PassThru -IncludeEqual -ExcludeDifferent
            if($correspondingTermProj.Count -gt 1){
                Write-Warning "`t[$($thisNetSuiteOpp.NetSuiteOppLabel)][$($thisNetSuiteOpp.NetSuiteOppId)] matches multiple ProjTerms: {$($correspondingTermProj | % {"[$($_.Name)][$($_.id)], "})}"
                $combinedOpps[$i].TermOppId = "Multiple"
                }
            elseif([string]::IsNullOrWhiteSpace($correspondingTermProj.id)){
                Write-Warning "`t[$($thisNetSuiteOpp.NetSuiteOppLabel)][$($thisNetSuiteOpp.NetSuiteOppId)] matched no ProjTerms"
                }
            else{
                $combinedOpps[$i].TermProjId = $correspondingTermProj.id
                $combinedOpps[$i].TermProjLabel = $correspondingTermProj.Name
                $combinedOpps[$i].DriveItemProjId = $correspondingTermProj.CustomProperties.DriveItemId

                $correspondingDriveItemProj = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $combinedOpps[$i].DriveClientId -itemGraphId $combinedOpps[$i].DriveItemProjId -returnWhat Item
                $combinedOpps[$i].DriveItemProjName = $correspondingDriveItemProj.name
                $combinedOpps[$i].DriveItemProjUrl = $correspondingDriveItemProj.webUrl
                }

            $correspondingNetSuiteProj = compare-object $allNetSuiteProjs $thisNetSuiteOpp -Property NetSuiteProjectId -PassThru -IncludeEqual -ExcludeDifferent
            if($correspondingNetSuiteProj.Count -gt 1){
                Write-Warning "`t[$($thisNetSuiteOpp.NetSuiteOppLabel)][$($thisNetSuiteOpp.NetSuiteOppId)] matches multiple NetProjs: {$($correspondingNetSuiteProj | % {"[$($_.Name)][$($_.id)], "})}"
                $combinedOpps[$i].TermOppId = "Multiple"
                }
            elseif([string]::IsNullOrWhiteSpace($correspondingNetSuiteProj.id)){
                Write-Warning "`t[$($thisNetSuiteOpp.NetSuiteOppLabel)][$($thisNetSuiteOpp.NetSuiteOppId)] matched no NetProjs"
                }
            else{
                $combinedOpps[$i].NetSuiteProjectName = $correspondingNetSuiteProj.entityId
                }
            }

        }
    
    $now = $(Get-Date -f FileDateTimeUniversal)
    $combinedOpps = $combinedOpps | ? {![string]::IsNullOrWhiteSpace($_.NetSuiteOppId)}
    $combinedOpps |  % {Export-Csv -InputObject $_ -Path "$env:USERPROFILE\Desktop\NetRec_Opps_$now.csv" -Append -NoTypeInformation -Encoding UTF8}
    }
Write-Host "Opps reconcilliation completed in [$($oppsReconcile.TotalMinutes)] minutes"
#NetSuiteProj #ProjTerm 



#List of all top-levlel folders in each drive
$now = $(Get-Date -f FileDateTimeUniversal)
$enumerateFolders = Measure-Command {
    for($i=0; $i-lt $allClientDrives.Count; $i++){
        write-progress -activity "Enumerating Drives contents" -Status "[$i/$($allClientDrives.count)]" -PercentComplete ($i/ $allClientDrives.count *100)
        $thisClientDrive = $allClientDrives[$i]
        $tokenResponseSharePointBot = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponseSharePointBot -renewTokenExpiringInSeconds 60 -aadAppCreds $sharePointBotDetails
        try{
            $theseTopLevelFolders = get-graphDriveItems -tokenResponse $tokenResponseSharePointBot -driveGraphId $thisClientDrive.DriveClientId -returnWhat Children
            }
        catch{
            write-warning "`tCould not retrieve DriveItems for Client [$($thisClientDrive.NetSuiteClientName)][$($thisClientDrive.NetSuiteClientId)]"
            Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
            }
        #$thisCombinedClient = $combinedClients | ? {$_.NetSuiteClientId -eq }
        @($theseTopLevelFolders | Select-Object) | % {
            $folderObject = New-Object PSObject -Property ([ordered]@{
                #NetSuiteClientId = $thisClientDrive.NetSuiteClientId
                #NetSuiteClientName = $thisClientDrive.NetSuiteClientName
                #TermClientId = $thisClientDrive.TermClientId
                #TermClientName = $thisClientDrive.TermClientName
                DriveClientId = $thisClientDrive.DriveClientId
                DriveClientName = $thisClientDrive.DriveClientName
                DriveClientUrl = $thisClientDrive.DriveClientUrl
                DriveItemName = $_.name
                DriveItemId = $_.Id
                DriveItemUrl = $_.weburl
                DriveItemCreatedDateTime = $_.createdDateTime
                DriveItemLastModifiedDateTime = $_.lastModifiedDateTime
                DriveItemSize = $_.size
                DriveItemChildCountForFolders = $_.folder.childCount
                DriveItemFirstWord = $null
                })
            $folderObject.DriveItemFirstWord = ([uri]::UnescapeDataString($(Split-Path $folderObject.DriveItemUrl -Leaf)) -split " ")[0]
            $folderObject | Export-Csv -Path "$env:USERPROFILE\Desktop\NetRec_AllFolders_$now.csv" -Append -NoTypeInformation -Encoding UTF8 -Force
            }
        }
    $combinedFolders = import-csv "$env:USERPROFILE\Desktop\NetRec_AllFolders_$now.csv"
    }
Write-Host "ClientDrive folders enumerated in [$($enumerateFolders.TotalMinutes)]"

$allOppProjFolders = $combinedFolders | ? {$_.DriveItemFirstWord -match '^[OP]-10'}
$allOppFolders = $combinedFolders | ? {$_.DriveItemFirstWord -match '^O-10'}
$allProjFolders = $combinedFolders | ? {$_.DriveItemFirstWord -match '^P-10'}




$validateOppFolders = Measure-Command {
    $now = $(Get-Date -f FileDateTimeUniversal)
    for($i=0; $i-lt $allOppFolders.Count; $i++){
        write-progress -activity "Validating Opp folders" -Status "[$i/$($allOppFolders.count)]" -PercentComplete ($i/ $allOppFolders.count *100)
        $thisOppFolder = $allOppFolders[$i]
        $correspondingOpp = $combinedOpps | ? {($_.NetSuiteOppLabel -split " ")[0] -eq $thisOppFolder.DriveItemFirstWord}
        $thisOppFolder | Add-Member -MemberType NoteProperty -Name "ProjectCodeFromNetSuite" -Value $null -Force
        $thisOppFolder | Add-Member -MemberType NoteProperty -Name "ClientDriveIdMatches" -Value $null -Force
        $thisOppFolder | Add-Member -MemberType NoteProperty -Name "OppDriveItemIdMatches" -Value $null -Force
        if($correspondingOpp){
            if(![string]::IsNullOrWhiteSpace($correspondingOpp.NetSuiteProjectName)){$thisOppFolder.ProjectCodeFromNetSuite = $(($correspondingOpp.NetSuiteProjectName -split " ")[0])}
            else{$thisOppFolder.ProjectCodeFromNetSuite = "None"} 
            if($correspondingOpp.DriveClientId -eq $thisOppFolder.DriveClientId){$thisOppFolder.ClientDriveIdMatches = $true}
            else{$thisOppFolder.ClientDriveIdMatches = $false}
            if($correspondingOpp.DriveItemOppId -eq $thisOppFolder.DriveItemId){$thisOppFolder.OppDriveItemIdMatches = $true}
            else{$thisOppFolder.OppDriveItemIdMatches = $false}
            }
        else{write-warning "[$($thisOppFolder.DriveItemName)] did not have a corresponding Opp"}
        $thisOppFolder | Export-Csv -Path "$env:USERPROFILE\Desktop\NetRec_ValidatedOppFolders_$now.csv" -Append -NoTypeInformation -Encoding UTF8 -Force
        }
    }
Write-Host "ClientDrive Opp folders validated in [$($validateOppFolders.TotalMinutes)]"
$validatedOppFolders = Import-Csv "$env:USERPROFILE\Desktop\NetRec_ValidatedOppFolders_$now.csv"



$validateProjFolders = Measure-Command {
    $now = $(Get-Date -f FileDateTimeUniversal)
    for($i=0; $i-lt $allProjFolders.Count; $i++){
        write-progress -activity "Validating Proj folders" -Status "[$i/$($allProjFolders.count)]" -PercentComplete ($i/ $allProjFolders.count *100)
        $thisProjFolder = $allProjFolders[$i]
         $correspondingProj = $combinedOpps | ? {($_. -split " ")[0] -eq $thisProjFolder.DriveItemFirstWord}
        $thisProjFolder | Add-Member -MemberType NoteProperty -Name "ClientDriveIdMatches" -Value $null -Force
        $thisProjFolder | Add-Member -MemberType NoteProperty -Name "ProjDriveItemIdMatches" -Value $null -Force
        if($correspondingProj){
            if($correspondingProj.DriveClientId -eq $thisProjFolder.DriveClientId){$thisProjFolder.ClientDriveIdMatches = $true}
            else{$thisProjFolder.ClientDriveIdMatches = $false}
            if($correspondingProj.DriveItemProjId -eq $thisProjFolder.DriveItemId){$thisProjFolder.ProjDriveItemIdMatches = $true}
            else{$thisProjFolder.ProjDriveItemIdMatches = $false}
            }
        else{write-warning "[$($thisProjFolder.DriveItemName)] did not have a corresponding Opp"}
        $thisProjFolder | Export-Csv -Path "$env:USERPROFILE\Desktop\NetRec_ValidatedProjFolders_$now.csv" -Append -NoTypeInformation -Encoding UTF8 -Force
        }
    }
Write-Host "ClientDrive Proj folders validated in [$($validateProjFolders.TotalMinutes)]"
$validatedProjFolders = Import-Csv "$env:USERPROFILE\Desktop\NetRec_ValidatedProjFolders_$now.csv"






#########################
#
#   Now fix some of the broken stuff
#
#########################

$realOppFoldersMistakenlyOrphaned = $validatedOppFolders |? {$_.ClientDriveIdMatches -eq $true -and $_.OppDriveItemIdMatches -eq $false}
$realOppFoldersMistakenlyOrphaned | % {
    #set term.DriveItemId  to this correct value
    }

$oppFoldersThatShouldHaveMergedIntoProjectFolders = $validatedOppFolders |? {![string]::IsNullOrWhiteSpace($_.ProjectCodeFromNetSuite) -and $_.ProjectCodeFromNetSuite -ne "None"}
$oppFoldersThatShouldHaveMergedIntoProjectFolders | % {
    $thisDuffer = $_
    if($thisDuffer.DriveItemSize -eq 0){
        #delete-graphDriveItem -tokenResponse $tokenResponseSharePointBot -graphDriveId $thisDuffer.DriveClientId -graphDriveItemId $thisDuffer.DriveItemId -Verbose
        }
    else{
        #validate Project folder, then merge
        }
    }

$oppFoldersinTheWrongPlaces = $validatedOppFolders |? {$_.ClientDriveIdMatches -eq $false}
$oppFoldersinTheWrongPlaces | % {
    $thisDuffer = $_
    if($thisDuffer.DriveItemSize -eq 0){
        #delete-graphDriveItem -tokenResponse $tokenResponseSharePointBot -graphDriveId $thisDuffer.DriveClientId -graphDriveItemId $thisDuffer.DriveItemId -Verbose
        }
    else{
        #validate Opp folder, then merge
        }
    }

<#
$duffers = Import-csv "$env:USERPROFILE\Desktop\EmptyOppFoldersWithProjects_toRemove.csv"
$duffers | % {
    $thisDuffer = $_
    if($thisDuffer.DriveItemSize -eq 0){
        delete-graphDriveItem -tokenResponse $tokenResponseSharePointBot -graphDriveId $thisDuffer.DriveClientId -graphDriveItemId $thisDuffer.DriveItemId -Verbose
        }
    
    }




























#NetSuiteClient #ClientTerm #ClientDrive


<#The names of these DocLibs don't match their URLs (the company may have been misspelled originally, the campany may have rebranded, or they may have been incorrectly corss-linked to another company)
$mismatches = $allClientDrives | ? {$_.webUrl -notmatch $($_.name -match '(?<=^[\s\?@]*)(\w+)' | out-null ;$Matches[0] -replace '[^a-zA-Z0-9]')}
$mismatches | select name, weburl
$urlMismatches = $combinedClients | ? {$mismatches.id -contains $_.DriveClientId} 
        
for($j=0;$j -lt $urlMismatches.Count; $j++){
    write-progress -activity "Processing NetSuite Opps" -Status "[$j/$($urlMismatches.count)]" -PercentComplete ($j/ $urlMismatches.count *100)
    $thisMismatch = $urlMismatches[$j]
    $possibleDriveMatch = $allClientDrives | ? {$_.webUrl -match $($thisMismatch.NetSuiteClientName -match '(?<=^[\s\?@]*)(\w+)' | out-null ;$Matches[0] -replace '[^a-zA-Z0-9]')}
    @($possibleDriveMatch | Select-Object) | % {
        $thisPossibleMatch = $_
        $i=0
        while(![string]::IsNullOrWhiteSpace($thisMismatch."possibleMatch$i")){$i++}
        $thisMismatch | Add-Member -MemberType NoteProperty -Name "possibleMatchId$i" -Value $thisPossibleMatch.id -Force
        $thisMismatch | Add-Member -MemberType NoteProperty -Name "possibleMatchUrl$i" -Value $thisPossibleMatch.webUrl -Force
        }
    
    }
$now = $(Get-Date -f FileDateTimeUniversal)
$urlMismatches |  % {Export-Csv -InputObject $_ -Path "$env:USERPROFILE\Desktop\NetRec_Clients_UrlMismatches_$now.csv" -Append -NoTypeInformation -Encoding UTF8 -Force}


#These DocLibs have duplicated names
$uniqueNames = $allClientDrives.name | Sort-Object | select -Unique
$duplicates = Compare-Object -ReferenceObject ($allClientDrives.name | Sort-Object) -DifferenceObject $uniqueNames -PassThru

#These Terms have duplicate NetSuiteIds


#These Terms have duplicate GraphDriveIds


#These Terms have no NetSuiteId
$termsWithNoNetSuiteId = $allClientTerms | ? {[string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteId)}
#These Terms have no GraphDriveId
$termsWithNoDriveId = $allClientTerms | ? {[string]::IsNullOrWhiteSpace($_.CustomProperties.GraphDriveId)}
#These Terms have neither:
$termsWithNoNetSuiteIdOrDriveId = $termsWithNoNetSuiteId | ? {[string]::IsNullOrWhiteSpace($_.CustomProperties.GraphDriveId)}
Write-Host "No NetSuiteId:`t[$($termsWithNoNetSuiteId.Count)/[$($allClientTerms.Count)]"
Write-Host "No DriveId:`t`t[$($termsWithNoDriveId.Count)/[$($allClientTerms.Count)]"
Write-Host "Neither:`t`t[$($termsWithNoNetSuiteIdOrDriveId.Count)/[$($allClientTerms.Count)]"
$termsWithNoNetSuiteIdOrDriveId | % {
    #Deprecate these Terms
    }




#Orphaned Terms:
    #These Client Terms have no NetSuiteClientID

    #These Opp Terms have no NetSuiteClientID
    #These Opp Terms have no NetSuiteOppID

    #These Proj Terms have no NetSuiteClientID
    #These Proj Terms have no NetSuiteProjID


#Orphaned Drives:
    #These Drives' Ids do not appear in Terms
    $mismatchedDrivesAndTerms = Compare-Object -ReferenceObject $allClientDrives -DifferenceObject $allClientTerms -Property DriveId -PassThru
    $drivesWithNoTerm = $mismatchedDrivesAndTerms | ? {$_.SideIndicator -eq "<="}
    $drivesWithNoTerm | select Name,webUrl

#Orphaned Folders:
    #These O- folders do not have Opps associated with them
    #These P- folders do not have Projects associated with them

#>