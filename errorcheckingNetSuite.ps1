#region Retrieve and prepare NetObjects
$importClientsTime = Measure-Command {
    $allNetSuiteClients = import-encryptedCsv -pathToEncryptedCsv $env:TEMP\Client.csv
    }
Write-Host "[$($allNetSuiteClients.Count)] Clients read from cache in $(format-measureCommandResults $importClientsTime)"
$importOppsTime = Measure-Command {
    $allNetSuiteOpps = import-encryptedCsv -pathToEncryptedCsv $env:TEMP\Opportunity.csv
    }
Write-Host "[$($allNetSuiteOpps.Count)] Opps read from cache in $(format-measureCommandResults $importOppsTime)"
$importProjTime = Measure-Command {
    $allNetSuiteProjs = import-encryptedCsv -pathToEncryptedCsv $env:TEMP\project.csv
    }
Write-Host "[$($allNetSuiteProjs.Count)] Projs read from cache in $(format-measureCommandResults $importProjTime)"
Write-Host "NetObject data import from cache took $(format-measureCommandResults $($importClientsTime+$importOppsTime+$importProjTime)) in total" -f Yellow
#endregion

#region Retrieve and prepare FoldObjects
$importFoldersTime = Measure-Command {
    $topLevelFolders = import-encryptedCsv -pathToEncryptedCsv $env:TEMP\folders.csv
    $topLevelFolders | % { #Generate the URL for the Document Library if it's missing (this seems to be a common problem)
        $thisTlf = $_
        if([string]::IsNullOrWhiteSpace($thisTlf.DriveClientUrl)){
            $thisTlf.DriveClientUrl = $($thisTlf.DriveItemUrl -replace [regex]::Escape($($thisTlf.DriveItemUrl -replace '^([^/]*\/){4}[^/]*')))
            }
        }
    #Create an array of Client Document Library objects (so we can differentiate these from Opp/Proj folders later)
    $allFolderClients = $topLevelFolders | Select-Object DriveClientId, DriveClientName, DriveClientUrl
    [array]$allFolderClientsUnique = $allFolderClients | Group-Object -Property DriveClientId | % {$_.Group | Select-Object -First 1}
    #Filter out all folders other than Opp/Proj ones
    $allOppProjFolders = $topLevelFolders | Where-Object {$_.DriveItemFirstWord -match 'O-\d\d\d\d\d\d|P-\d\d\d\d\d\d'}
    $allOppProjFolders | % {$_ | Add-Member -Name FolderClientDriveId -MemberType AliasProperty -Value DriveClientId -Force}
    }
Write-Host "[$($topLevelFolders.Count)] Top-level Folders read from cache, filtered to [$($allOppProjFolders.Count)] Opp/Proj folders in $(format-measureCommandResults $importFoldersTime)"
#$topLevelFolders = Import-Csv C:\Users\KEVMAI~1\AppData\Local\Temp\NetRec_AllFolders_20211206T1053496934Z.csv
Write-Host "FoldObject data import and preparation from cache took $(format-measureCommandResults $($importFoldersTime)) in total" -f Yellow
#endregion

#region Retrieve and prepare TermObjects
$appCredsSharePointBot = $(get-graphAppClientCredentials -appName SharePointBot)
$tokenResponseSharePointBot = get-graphTokenResponse -aadAppCreds $appCredsSharePointBot
$termSite = get-graphSite -tokenResponse $tokenResponseSharePointBot -serverRelativeUrl "/"
$termGroup = get-graphTermGroup -tokenResponse $tokenResponseSharePointBot -graphSiteId $termSite.id -graphTermGroupName "Kimble"
$termSets = get-graphTermSet -tokenResponse $tokenResponseSharePointBot -graphSiteId $termSite.id -graphTermGroupId $termGroup.id 

$getProjTerms = Measure-Command {
    $termSetProjects = $termSets | Where-Object {"Projects" -contains $_.localizedNames.name}
    $allProjTermsRaw = get-graphTerm -tokenResponse $tokenResponseSharePointBot -graphSiteId $termSite.id -graphTermGroupId $termGroup.id -graphTermSetId $termSetProjects.id -selectAllProperties
    }
Write-Host "[$($allProjTermsRaw.Count)] Projs retrieved in $(format-measureCommandResults $getProjTerms)"
$allProjTerms = @($null)*$allProjTermsRaw.Count
$standardiseProjs = Measure-Command {
    for($i=0; $i -lt $allProjTerms.Count; $i++){
        Write-Progress -activity "Standardising Projs" -Status "[$i/$($allProjTerms.count)]" -PercentComplete ($($i/ $allProjTerms.count) *100)
        $allProjTerms[$i] = New-Object -TypeName psobject -Property @{
            #NetSuiteProjLastModifiedDate = $($allProjTermsRaw[$i].properties | Where-Object {$_.key -eq "NetSuiteProjLastModifiedDate"}).value 
            TermNetSuiteProjId = $($allProjTermsRaw[$i].properties | Where-Object {$_.key -eq "NetSuiteProjId"}).value 
            TermProjClientId = $($allProjTermsRaw[$i].properties | Where-Object {$_.key -eq "NetSuiteClientId"}).value 
            TermProjCode = $(($($allProjTermsRaw[$i].labels | Where-Object {$_.isDefault -eq $true}).Name -split " ")[0]) 
            TermProjDriveItemId = $($allProjTermsRaw[$i].properties | Where-Object {$_.key -eq "DriveItemId"}).value 
            TermProjFlaggedForReprocessing = $($allProjTermsRaw[$i].properties | Where-Object {$_.key -eq "flagForReprocessing"}).value 
            TermProjId = $($allProjTermsRaw[$i].Id) 
            TermProjLastModifiedDate = $allProjTermsRaw[$i].lastModifiedDateTime
            TermProjName = $($allProjTermsRaw[$i].labels | Where-Object {$_.isDefault -eq $true}).Name 
            #UniversalProjCode = $(($($allProjTermsRaw[$i].labels | Where-Object {$_.isDefault -eq $true}).Name -split " ")[0]) 
            #UniversalProjName = $($allProjTermsRaw[$i].labels | Where-Object {$_.isDefault -eq $true}).Name 
            #UniversalProjNameSanitised = $(sanitise-forNetsuiteIntegration $($allProjTermsRaw[$i].labels | Where-Object {$_.isDefault -eq $true}).Name) 
            }
        }
    }
Write-Host "[$($allProjTerms.Count)] Projs standardised in $(format-measureCommandResults $standardiseProjs)"
Write-Host "Proj Term retrieval and preparation took $(format-measureCommandResults $($getProjTerms+$standardiseProjs)) in total" -f Yellow

$getOppTerms = Measure-Command {
    $termSetOpportunities = $termSets | Where-Object {"Opportunities" -contains $_.localizedNames.name}
    $allOppTermsRaw = get-graphTerm -tokenResponse $tokenResponseSharePointBot -graphSiteId $termSite.id -graphTermGroupId $termGroup.id -graphTermSetId $termSetOpportunities.id -selectAllProperties
    }
Write-Host "[$($allOppTermsRaw.Count)] Opps retrieved in $(format-measureCommandResults $getOppTerms)"
$allOppTerms = @($null)*$allOppTermsRaw.Count
$standardiseOpps = Measure-Command {
    for($i=0; $i -lt $allOppTerms.Count; $i++){
        Write-Progress -activity "Standardising Opps" -Status "[$i/$($allOppTerms.count)]" -PercentComplete ($($i/ $allOppTerms.count) *100)
        $allOppTerms[$i] = New-Object -TypeName psobject -Property @{
            #NetSuiteClientId = $($allOppTermsRaw[$i].properties | Where-Object {$_.key -eq "NetSuiteClientId"}).value 
            TermNetSuiteOppId = $($allOppTermsRaw[$i].properties | Where-Object {$_.key -eq "NetSuiteOppId"}).value
            #NetSuiteOppLastModifiedDate = $($allOppTermsRaw[$i].properties | Where-Object {$_.key -eq "NetSuiteOppLastModifiedDate"}).value 
            #NetSuiteOppProjId = $($allOppTermsRaw[$i].properties | Where-Object {$_.key -eq "NetSuiteProjectId"}).value 
            TermOppClientId = $($allOppTermsRaw[$i].properties | Where-Object {$_.key -eq "NetSuiteClientId"}).value 
            TermOppCode = $(($($allOppTermsRaw[$i].labels | Where-Object {$_.isDefault -eq $true}).Name -split " ")[0]) 
            TermOppDriveItemId = $($allOppTermsRaw[$i].properties | Where-Object {$_.key -eq "DriveItemId"}).value 
            TermOppFlaggedForReprocessing = $($allOppTermsRaw[$i].properties | Where-Object {$_.key -eq "flagForReprocessing"}).value 
            TermOppId = $($allOppTermsRaw[$i].Id) 
            TermOppLastModifiedDate = $allOppTermsRaw[$i].lastModifiedDateTime
            TermOppName = $($allOppTermsRaw[$i].labels | Where-Object {$_.isDefault -eq $true}).Name 
            TermOppProjId = $($allOppTermsRaw[$i].properties | Where-Object {$_.key -eq "NetSuiteProjectId"}).value 
            #UniversalOppName = $($allOppTermsRaw[$i].labels | Where-Object {$_.isDefault -eq $true}).Name 
            #UniversalOppNameSanitised = $(sanitise-forNetsuiteIntegration $($allOppTermsRaw[$i].labels | Where-Object {$_.isDefault -eq $true}).Name) 
            }
        }
    }
Write-Host "[$($allOppTerms.Count)] Opps standardised in $(format-measureCommandResults $standardiseOpps)"
Write-Host "Opp Term retrieval and preparation took $(format-measureCommandResults $($getOppTerms+$standardiseOpps)) in total" -f Yellow

$getClientTerms = Measure-Command {
    $termSetClients = $termSets | Where-Object {"Clients" -contains $_.localizedNames.name}
    $allClientTermsRaw = get-graphTerm -tokenResponse $tokenResponseSharePointBot -graphSiteId $termSite.id -graphTermGroupId $termGroup.id -graphTermSetId $termSetClients.id -selectAllProperties
    }
Write-Host "[$($allClientTermsRaw.Count)] Clients retrieved in $(format-measureCommandResults $getClientTerms)"
$allClientTerms = @($null)*$allClientTermsRaw.Count
$standardiseClients = Measure-Command {
    for($i=0; $i -lt $allClientTerms.Count; $i++){
        Write-Progress -activity "Standardising Clients" -Status "[$i/$($allClientTerms.count)]" -PercentComplete ($($i/ $allClientTerms.count) *100)
        $allClientTerms[$i] = New-Object -TypeName psobject -Property @{
            TermNetSuiteClientId = $($allClientTermsRaw[$i].properties | Where-Object {$_.key -eq "NetSuiteId"}).value 
            #DriveClientId = $($allClientTermsRaw[$i].properties | Where-Object {$_.key -eq "NetSuiteId"}).value 
            TermClientDriveId = $($allClientTermsRaw[$i].properties | Where-Object {$_.key -eq "GraphDriveId"}).value 
            TermClientId = $($allClientTermsRaw[$i].Id) 
            TermClientName = $($allClientTermsRaw[$i].labels | Where-Object {$_.isDefault -eq $true}).Name 
            TermClientLastModifiedDate = $allClientTermsRaw[$i].lastModifiedDateTime
            #NetSuiteLastModifiedDate = $($allClientTermsRaw[$i].properties | Where-Object {$_.key -eq "NetSuiteLastModifiedDate"}).value 
            }
        }
    }
Write-Host "[$($allClientTerms.Count)] Clients standardised in $(format-measureCommandResults $standardiseClients)"
Write-Host "Client Term retrieval and preparation took $(format-measureCommandResults $($getClientTerms+$standardiseClients)) in total" -f Yellow
Write-Host "TermObject retrieval and preparation took $(format-measureCommandResults $($getClientTerms+$standardiseClients+$getOppTerms+$standardiseOpps+$getProjTerms+$standardiseProjs)) in total" -f Magenta
#endregion

#region PrettyObject schemas
function get-linqLeftJoin(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
            $dataSet1
        ,[parameter(Mandatory = $true)]
            [string]$dataSet1PropertyToCompare
        ,[parameter(Mandatory = $true)]
            $dataSet2
        ,[parameter(Mandatory = $true)]
            [string]$dataSet2PropertyToCompare
        ,[parameter(Mandatory = $true)]
            [hashtable]$dataSet1PropertiesToReturn
        ,[parameter(Mandatory = $true)]
            [hashtable]$dataSet2PropertiesToReturn
        )

    #Write-Verbose "Return object hashtable: $(stringify-hashTable -hashtable $returnObjectHash)"
    $linqLeftJoinedData = [System.Linq.Enumerable]::GroupJoin(
        $dataSet1,
        $dataSet2,
        [System.Func[Object,string]] {param ($x);$x.$dataSet1PropertyToCompare},
        [System.Func[Object,string]]{param ($y);$y.$dataSet2PropertyToCompare},
        [System.Func[Object,Collections.Generic.IEnumerable[Object],Object]]{
            param ($x,$y);
            $returnObjectHash = [ordered]@{}
            $dataSet1PropertiesToReturn.Keys | Sort-Object | ForEach-Object {
                $returnObjectHash.Add($_,$($x.$($dataSet1PropertiesToReturn[$_] -replace '^(.*?)\.','')))
                }
            $dataSet2PropertiesToReturn.Keys | Sort-Object | ForEach-Object {
                $thisKey = $_
                try{$returnObjectHash.Add($thisKey,$($y.$($dataSet2PropertiesToReturn[$thisKey] -replace '^(.*?)\.','')))}
                catch{
                    if($_.Exception -match "Item has already been added"){
                        $returnObjectHash.Add("$thisKey"+"_1",$($y.$($dataSet2PropertiesToReturn[$thisKey] -replace '^(.*?)\.','')))
                        }
                    }
                }
            New-Object -TypeName psobject -Property $returnObjectHash
            }
        )
    [System.Linq.Enumerable]::ToArray($linqLeftJoinedData)
    }
function get-linqJoin(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
            $dataSet1
        ,[parameter(Mandatory = $true)]
            [string]$dataSet1PropertyToCompare
        ,[parameter(Mandatory = $true)]
            $dataSet2
        ,[parameter(Mandatory = $true)]
            [string]$dataSet2PropertyToCompare
        ,[parameter(Mandatory = $true)]
            [hashtable]$dataSet1PropertiesToReturn
        ,[parameter(Mandatory = $true)]
            [hashtable]$dataSet2PropertiesToReturn
        )

    #Write-Verbose "Return object hashtable: $(stringify-hashTable -hashtable $returnObjectHash)"

    $linqJoinedData = [System.Linq.Enumerable]::Join(
        $dataSet1,
        $dataSet2,
        [System.Func[Object,string]] {param ($x);$x.$dataSet1PropertyToCompare},
        [System.Func[Object,string]]{param ($y);$y.$dataSet2PropertyToCompare},
        [System.Func[Object,Object,Object]]{
            param ($x,$y);
            $returnObjectHash = [ordered]@{}
            $dataSet1PropertiesToReturn.Keys | ForEach-Object {
                $returnObjectHash.Add($_,$($x.$($dataSet1PropertiesToReturn[$_] -replace '^(.*?)\.','')))
                }
            $dataSet2PropertiesToReturn.Keys | ForEach-Object {
                $thisKey = $_
                try{$returnObjectHash.Add($thisKey,$($y.$($dataSet2PropertiesToReturn[$thisKey] -replace '^(.*?)\.','')))}
                catch{
                    if($_.Exception -match "Item has already been added"){
                        $returnObjectHash.Add("$thisKey"+"_1",$($y.$($dataSet2PropertiesToReturn[$thisKey] -replace '^(.*?)\.','')))
                        }
                    }
                }
            New-Object -TypeName psobject -Property $returnObjectHash
            }
        )
    [System.Linq.Enumerable]::ToArray($linqJoinedData)
    }
function get-prettyNetClientHash(){
    $([ordered]@{
        NetSuiteClientId = "NetSuiteClientId"
        NetSuiteClientCode = "NetSuiteClientCode"
        NetSuiteClientName = "NetSuiteClientName"
        NetSuiteClientLastModifiedDate = "NetSuiteClientLastModifiedDate"
        })
    }
function get-prettyNetOppHash(){
    $([ordered]@{
        NetSuiteOppId = "NetSuiteOppId"
        NetSuiteOppName = "NetSuiteOppName"
        NetSuiteOppProjId = "NetSuiteOppProjId"
        NetSuiteOppLastModifiedDate = "NetSuiteOppLastModifiedDate"
        })
    }
function get-prettyNetProjHash(){
    $([ordered]@{
        NetSuiteProjId = "NetSuiteProjId"
        NetSuiteProjName = "NetSuiteProjName"
        NetSuiteProjLastModifiedDate = "NetSuiteProjLastModifiedDate"
        })
    }
function get-prettyTermClientHash(){
    $([ordered]@{
        TermClientId = "TermClientId"
        TermClientName = "TermClientName"
        TermClientDriveId = "TermClientDriveId"
        TermClientLastModifiedDate = "TermClientLastModifiedDate"
        })
    }
function get-prettyTermOppHash(){
    $([ordered]@{
        #TermClientId = "TermClientId"
        #TermClientName = "TermClientName"
        #TermClientDriveId = "TermClientDriveId"
        #TermClientLastModifiedDate = "TermClientLastModifiedDate"
        TermOppClientId = "TermOppClientId"
        TermOppId = "TermOppId"
        TermOppName = "TermOppName"
        TermOppProjId = "TermOppProjId"
        TermOppDriveItemId = "TermOppDriveItemId"
        TermOppFlaggedForReprocessing = "TermOppFlaggedForReprocessing"
        TermOppLastModifiedDate = "TermOppLastModifiedDate"
        })
    }
function get-prettyTermProjHash(){
    $([ordered]@{
        #TermClientId = "TermClientId"
        #TermClientName = "TermClientName"
        #TermClientDriveId = "TermClientDriveId"
        #TermClientLastModifiedDate = "TermClientLastModifiedDate"
        TermProjClientId = "TermProjClientId"
        TermProjId = "TermProjId"
        TermProjName = "TermProjName"
        TermProjDriveItemId = "TermProjDriveItemId"
        TermProjFlaggedForReprocessing = "TermProjFlaggedForReprocessing"
        TermProjLastModifiedDate = "TermProjLastModifiedDate"
        })
    }
function get-prettyFolderClientHash(){
    $([ordered]@{
        FolderClientDriveId = "FolderClientDriveId"
        FolderClientName = "FolderClientName"
        FolderClientUrl = "FolderClientUrl"
        })
    }
function get-prettyFolderOppHash(){
    $([ordered]@{
        #FolderClientDriveId = "FolderClientDriveId"
        #FolderClientName = "FolderClientName"
        #FolderClientUrl = "FolderClientUrl"
        FolderOppDriveItemId = "FolderOppDriveItemId"
        FolderOppName = "FolderOppName"
        FolderOppUrl = "FolderOppUrl"
        })
    }
function get-prettyFolderProjHash(){
    $([ordered]@{
        #FolderClientDriveId = "FolderClientDriveId"
        #FolderClientName = "FolderClientName"
        #FolderClientUrl = "FolderClientUrl"
        FolderProjDriveItemId = "FolderProjDriveItemId"
        FolderProjName = "FolderProjName"
        FolderProjUrl = "FolderProjUrl"
        })
    }
function get-uglyNetClientHash(){
    $([ordered]@{
        NetSuiteClientId = "id"
        NetSuiteClientCode = "entityId"
        NetSuiteClientName = "companyName"
        NetSuiteClientLastModifiedDate = "lastModifiedDate"
        })
    }
function get-uglyNetOppHash(){
    $([ordered]@{
        NetSuiteOppId = "id"
        NetSuiteOppName = "NetSuiteOppLabel"
        NetSuiteOppProjId = "NetSuiteProjectId"
        NetSuiteOppLastModifiedDate = "lastModifiedDate"
        })
    }
function get-uglyNetProjHash(){
    $([ordered]@{
        NetSuiteProjId = "id"
        NetSuiteProjName = "entityId"
        NetSuiteProjLastModifiedDate = "lastModifiedDate"
        })
    }
function get-uglyTermClientHash(){
    $([ordered]@{
        TermClientId = "TermClientId"
        TermClientName = "TermClientName"
        TermClientDriveId = "FolderClientDriveId"
        TermClientLastModifiedDate = "TermLastModifiedDate"
        })
    }
function get-uglyTermOppHash(){
    $([ordered]@{
        TermClientId = "TermClientId"
        TermClientName = "TermClientName"
        TermClientDriveId = "FolderClientDriveId"
        TermClientLastModifiedDate = "TermLastModifiedDate"
        TermOppClientId = "NetSuiteClientId"
        TermOppId = "NetSuiteOppId"
        TermOppName = "TermOppLabel"
        TermOppProjId = "TermProjId"
        TermOppDriveItemId = "DriveItemId"
        TermOppFlaggedForReprocessing = "TermOppFlaggedForReprocessing"
        TermOppLastModifiedDate = "TermOppLastModifiedDate"
        })
    }
function get-uglyFolderClientHash(){
    $([ordered]@{
        FolderClientDriveId = "DriveClientId"
        FolderClientName = "DriveClientName"
        FolderClientUrl = "DriveClientUrl"
        })
    }
function get-uglyFolderOppHash(){
    $([ordered]@{
        FolderClientDriveId = "DriveClientId"
        FolderClientName = "DriveClientName"
        FolderClientUrl = "DriveClientUrl"
        FolderOppDriveItemId = "DriveItemId"
        FolderOppName = "DriveItemName"
        FolderOppUrl = "DriveItemUrl"
        })
    }
function get-uglyFolderProjHash(){
    $([ordered]@{
        FolderClientDriveId = "DriveClientId"
        FolderClientName = "DriveClientName"
        FolderClientUrl = "DriveClientUrl"
        FolderProjDriveItemId = "DriveItemId"
        FolderProjName = "DriveItemName"
        FolderProjUrl = "DriveItemUrl"
        })
    }
#endregion

#region Cross-reference NetSuite (Net), Term Store (Term) and SharePoint (Folder) data to build "Pretty" objects
$preparingClients = Measure-Command {
    #Get the NetClients LEFT OUTER JOIN TermClients data
    $prettyClientsInterim = get-linqLeftJoin -dataSet1 $allNetSuiteClients -dataSet1PropertyToCompare id -dataSet2 $allClientTerms -dataSet2PropertyToCompare TermNetSuiteClientId -dataSet1PropertiesToReturn $(get-uglyNetClientHash) -dataSet2PropertiesToReturn $(get-prettyTermClientHash)
    #Get the NetClients RIGHT OUTER JOIN TermClients data
    $prettyClientsInterimRight = get-linqLeftJoin -dataSet1 $allClientTerms -dataSet1PropertyToCompare TermNetSuiteClientId -dataSet2 $allNetSuiteClients -dataSet2PropertyToCompare id -dataSet1PropertiesToReturn $(get-prettyTermClientHash) -dataSet2PropertiesToReturn $(get-uglyNetClientHash)
    #Remove the INNER JOIN results from the NetClients RIGHT OUTER JOIN TermClients data
    $prettyTermClientsWithoutNetClients = $prettyClientsInterimRight | Where-Object {[string]::IsNullOrWhiteSpace($_.NetSuiteClientId)} #These are "Orphaned" Client Terms
    Write-Verbose "[$($prettyClientsInterim.Count)] NetClients linked to TermClients"
    Write-Verbose "[$($prettyTermClientsWithoutNetClients.Count)] TermClients without NetClients identified (orphaned TermClients)"
    Write-Verbose "[$($prettyClientsInterim.Count + $prettyTermClientsWithoutNetClients.Count)] ([$($prettyClientsInterim.Count)]+[$($prettyTermClientsWithoutNetClients.Count)]) total NetTermClients"
    #Recombine the LEFT OUTER JOIN and (RIGHT OUTER JOIN - INNER JOIN) data
    $prettyClientsInterim += $prettyTermClientsWithoutNetClients 
    }
Write-Host "[$($prettyClientsInterim.Count)] NetClients linked to TermClients in $(format-measureCommandResults $preparingClients)"
$preparingClients2 = Measure-Command {
    #Get the NetTermClients LEFT OUTER JOIN FoldClients data
    $prettyClients = get-linqLeftJoin -dataSet1 $prettyClientsInterim -dataSet1PropertyToCompare TermClientDriveId -dataSet2 $allFolderClientsUnique -dataSet2PropertyToCompare DriveClientId -dataSet1PropertiesToReturn $($(get-prettyNetClientHash) + $(get-prettyTermClientHash)) -dataSet2PropertiesToReturn $(get-uglyFolderClientHash)
    #Get the NetTermClients RIGHT OUTER JOIN FoldClients data
    $prettyClientsRight = get-linqLeftJoin -dataSet1 $allFolderClientsUnique -dataSet1PropertyToCompare DriveClientId -dataSet2 $prettyClientsInterim -dataSet2PropertyToCompare TermClientDriveId -dataSet1PropertiesToReturn $(get-uglyFolderClientHash) -dataSet2PropertiesToReturn $($(get-prettyNetClientHash) + $(get-prettyTermClientHash))
    #Remove the INNER JOIN results from the NetTermClients RIGHT OUTER JOIN FoldClients data
    $prettyFoldClientsWithoutNetTermClients = $prettyClientsRight | Where-Object {[string]::IsNullOrWhiteSpace($_.TermClientDriveId)} #These are "Orphaned" Client Folders
    Write-Verbose "[$($prettyClients.Count)] NetTermClients linked to FoldClients"
    Write-Verbose "[$($prettyFoldClientsWithoutNetTermClients.Count)] FoldClients without NetTermClients identified (orphaned FoldClients)"
    Write-Verbose "[$($prettyClients.Count + $prettyFoldClientsWithoutNetTermClients.Count)] ([$($prettyClients.Count)]+[$($prettyFoldClientsWithoutNetTermClients.Count)]) total NetTermFoldClients (`"PrettyClients`")"
    #Recombine the LEFT OUTER JOIN and (RIGHT OUTER JOIN - INNER JOIN) data
    $prettyClients += $prettyFoldClientsWithoutNetTermClients
    }
Write-Host "[$($prettyClients.Count)] NetTermClients linked to FoldClients in $(format-measureCommandResults $preparingClients2)"
Write-Host "PrettyClients preparation took $(format-measureCommandResults $($preparingClients+$preparingClients2)) in total" -f Yellow

$preparingOpps = Measure-Command {
    #Get the NetOpps LEFT OUTER JOIN TermOpps data
    $prettyOppsInterim = get-linqLeftJoin -dataSet1 $allNetSuiteOpps -dataSet1PropertyToCompare id -dataSet2 $allOppTerms -dataSet2PropertyToCompare TermNetSuiteOppId -dataSet1PropertiesToReturn $(get-uglyNetOppHash) -dataSet2PropertiesToReturn $(get-prettyTermOppHash)
    #Get the NetOpps RIGHT OUTER JOIN TermOpps data
    $prettyOppsInterimRight = get-linqLeftJoin -dataSet1 $allOppTerms -dataSet1PropertyToCompare TermNetSuiteOppId -dataSet2 $allNetSuiteOpps -dataSet2PropertyToCompare id -dataSet1PropertiesToReturn $(get-prettyTermOppHash) -dataSet2PropertiesToReturn $(get-uglyNetOppHash)
    #Remove the INNER JOIN results from the NetOpps RIGHT OUTER JOIN TermOpps data
    $prettyTermOppsWithoutNetOpps = $prettyOppsInterimRight | Where-Object {[string]::IsNullOrWhiteSpace($_.NetSuiteOppId)} #These are "Orphaned" Opp Terms
    Write-Verbose "[$($prettyOppsInterim.Count)] NetOpps linked to TermOpps"
    Write-Verbose "[$($prettyTermOppsWithoutNetOpps.Count)] TermOpps without NetOpps identified (orphaned TermOpps)"
    Write-Verbose "[$($prettyOppsInterim.Count + $prettyTermOppsWithoutNetOpps.Count)] ([$($prettyOppsInterim.Count)]+[$($prettyTermOppsWithoutNetOpps.Count)]) total NetTermOpps"
    #Recombine the LEFT OUTER JOIN and (RIGHT OUTER JOIN - INNER JOIN) data
    $prettyOppsInterim = $prettyOppsInterim + $prettyTermOppsWithoutNetOpps
    }
Write-Host "[$($prettyOppsInterim.Count)] NetOpps linked to TermOpps in $(format-measureCommandResults $preparingOpps)"
$preparingOpps2 = Measure-Command {
    #Get the NetTermOpps LEFT OUTER JOIN FoldOpps data
    $prettyOpps = get-linqLeftJoin -dataSet1 $prettyOppsInterim -dataSet1PropertyToCompare TermOppDriveItemId -dataSet2 $allOppProjFolders -dataSet2PropertyToCompare DriveItemId -dataSet1PropertiesToReturn $($(get-prettyNetOppHash) + $(get-prettyTermOppHash)) -dataSet2PropertiesToReturn $(get-uglyFolderOppHash)
    #Get the NetTermOpps RIGHT OUTER JOIN FoldOpps data
    $prettyOppsRight = get-linqLeftJoin -dataSet1 $allOppProjFolders -dataSet1PropertyToCompare DriveItemId -dataSet2 $prettyOppsInterim -dataSet2PropertyToCompare TermOppDriveItemId -dataSet1PropertiesToReturn $(get-uglyFolderOppHash) -dataSet2PropertiesToReturn $($(get-prettyNetOppHash) + $(get-prettyTermOppHash))
    #Remove the INNER JOIN results from the NetTermOpps RIGHT OUTER JOIN FoldOpps data
    $prettyFoldOppsWithoutNetTermOpps = $prettyOppsRight | Where-Object {[string]::IsNullOrWhiteSpace($_.TermOppDriveItemId)} #These are "Orphaned" Opp Folders
    Write-Verbose "[$($prettyOpps.Count)] NetTermOpps linked to FoldOpps"
    Write-Verbose "[$($prettyFoldOppsWithoutNetTermOpps.Count)] FoldOpps without NetTermOpps identified (orphaned FoldOpps)"
    Write-Verbose "[$($prettyOpps.Count + $prettyFoldOppsWithoutNetTermOpps.Count)] ([$($prettyOpps.Count)]+[$($prettyFoldOppsWithoutNetTermOpps.Count)]) total NetTermFoldOpps (`"PrettyOpps`")"
    #Recombine the LEFT OUTER JOIN and (RIGHT OUTER JOIN - INNER JOIN) data
    $prettyOpps += $prettyFoldOppsWithoutNetTermOpps
    }
Write-Host "[$($prettyOpps.Count)] NetTermOpps linked to FoldOpps in $(format-measureCommandResults $preparingOpps2)"
Write-Host "PrettyOpps preparation took $(format-measureCommandResults $($preparingOpps+$preparingOpps2)) in total" -f Yellow

$preparingProjs = Measure-Command {
    #Get the NetProjs LEFT OUTER JOIN TermProjs data
    $prettyProjsInterim = get-linqLeftJoin -dataSet1 $allNetSuiteProjs -dataSet1PropertyToCompare id -dataSet2 $allProjTerms -dataSet2PropertyToCompare TermNetSuiteProjId -dataSet1PropertiesToReturn $(get-uglyNetProjHash) -dataSet2PropertiesToReturn $(get-prettyTermProjHash)
    #Get the NetProjs RIGHT OUTER JOIN TermProjs data
    $prettyProjsInterimRight = get-linqLeftJoin -dataSet1 $allProjTerms -dataSet1PropertyToCompare TermNetSuiteProjId -dataSet2 $allNetSuiteProjs -dataSet2PropertyToCompare id -dataSet1PropertiesToReturn $(get-prettyTermProjHash) -dataSet2PropertiesToReturn $(get-uglyNetProjHash)
    #Remove the INNER JOIN results from the NetProjs RIGHT OUTER JOIN TermProjs data
    $prettyTermProjsWithoutNetProjs = $prettyProjsInterimRight | Where-Object {[string]::IsNullOrWhiteSpace($_.NetSuiteProjId)} #These are "Orphaned" Proj Terms
    Write-Verbose "[$($prettyProjsInterim.Count)] NetProjs linked to TermProjs"
    Write-Verbose "[$($prettyTermProjsWithoutNetProjs.Count)] TermProjs without NetProjs identified (orphaned TermProjs)"
    Write-Verbose "[$($prettyProjsInterim.Count + $prettyTermProjsWithoutNetProjs.Count)] ([$($prettyProjsInterim.Count)]+[$($prettyTermProjsWithoutNetProjs.Count)]) total NetTermProjs"
    #Recombine the LEFT OUTER JOIN and (RIGHT OUTER JOIN - INNER JOIN) data
    $prettyProjsInterim = $prettyProjsInterim + $prettyTermProjsWithoutNetProjs
    }
Write-Host "[$($prettyProjsInterim.Count)] NetProjs linked to TermProjs in $(format-measureCommandResults $preparingProjs)"
$preparingProjs2 = Measure-Command {
    #Get the NetTermProjs LEFT OUTER JOIN FoldProjs data
    $prettyProjs = get-linqLeftJoin -dataSet1 $prettyProjsInterim -dataSet1PropertyToCompare TermProjDriveItemId -dataSet2 $allOppProjFolders -dataSet2PropertyToCompare DriveItemId -dataSet1PropertiesToReturn $($(get-prettyNetProjHash) + $(get-prettyTermProjHash)) -dataSet2PropertiesToReturn $(get-uglyFolderProjHash)
    #Get the NetTermProjs RIGHT OUTER JOIN FoldProjs data
    $prettyProjsRight = get-linqLeftJoin -dataSet1 $allOppProjFolders -dataSet1PropertyToCompare DriveItemId -dataSet2 $prettyProjsInterim -dataSet2PropertyToCompare TermProjDriveItemId -dataSet1PropertiesToReturn $(get-uglyFolderProjHash) -dataSet2PropertiesToReturn $($(get-prettyNetProjHash) + $(get-prettyTermProjHash))
    #Remove the INNER JOIN results from the NetTermProjs RIGHT OUTER JOIN FoldProjs data
    $prettyFoldProjsWithoutNetTermProjs = $prettyProjsRight | Where-Object {[string]::IsNullOrWhiteSpace($_.TermProjDriveItemId)} #These are "Orphaned" Proj Folders
    Write-Verbose "[$($prettyProjs.Count)] NetTermProjs linked to FoldProjs"
    Write-Verbose "[$($prettyFoldProjsWithoutNetTermProjs.Count)] FoldProjs without NetTermProjs identified (orphaned FoldProjs)"
    Write-Verbose "[$($prettyProjs.Count + $prettyFoldProjsWithoutNetTermProjs.Count)] ([$($prettyProjs.Count)]+[$($prettyFoldProjsWithoutNetTermProjs.Count)]) total NetTermFoldProjs (`"PrettyProjs`")"
    #Recombine the LEFT OUTER JOIN and (RIGHT OUTER JOIN - INNER JOIN) data
    $prettyProjs += $prettyFoldProjsWithoutNetTermProjs
    }
Write-Host "[$($prettyProjs.Count)] NetTermProjs linked to FoldProjs in $(format-measureCommandResults $preparingProjs2)"
Write-Host "PrettyProjs preparation took $(format-measureCommandResults $($preparingProjs+$preparingProjs2)) in total" -f Yellow
Write-Host "PrettyObject preparation took $(format-measureCommandResults $($preparingClients+$preparingClients2+$preparingOpps+$preparingOpps2+$preparingProjs+$preparingProjs2)) in total" -f Magenta

rv prettyClientsInterim
rv prettyClientsInterimRight
rv prettyTermClientsWithoutNetClients
rv prettyClientsRight
rv prettyFoldClientsWithoutNetTermClients
rv prettyOppsInterim
rv prettyOppsInterimRight
rv prettyTermOppsWithoutNetOpps
rv prettyOppsRight
rv prettyFoldOppsWithoutNetTermOpps
rv prettyProjsInterim
rv prettyProjsInterimRight
rv prettyTermProjsWithoutNetProjs
rv prettyProjsRight
rv prettyFoldProjsWithoutNetTermProjs
#endregion

#region Cross-reference PrettyClients, PrettyOpps & PrettyProjs to build "Beautiful" objects
#Linq LEFT OUTER JOIN doesn't behave as expected: Clients LEFT OUTER JOIN Opps only returns exactly 1 Client-Opp match even if 1:Many (several Opps for the same Client) 
# :. we need to INNER JOIN, the add the (LEFT OUTER JOIN - INNER JOIN) and (RIGHT OUTER JOIN - INNER JOIN) results to retain all unlinked Clients and Opps. 
#In practice, an Opp can only have 0-1 Clients, so our RIGHT OUTER JOIN will correctly include all INNER JOIN results (meaning fewer queries). Then we need to find all Clients without Opps (LEFT OUTER JOIN - INNER JOIN) and re-add them.
$makingPrettyClientOpps = measure-Command {
    $prettyClientsAndOppsRight = get-linqLeftJoin -dataSet1 $prettyOpps -dataSet1PropertyToCompare TermOppClientId -dataSet2 $prettyClients -dataSet2PropertyToCompare NetSuiteClientId -dataSet1PropertiesToReturn $($(get-prettyNetOppHash) + $(get-prettyTermOppHash) + $(get-prettyFolderOppHash)) -dataSet2PropertiesToReturn $($(get-prettyNetClientHash) + $(get-prettyTermClientHash) + $(get-prettyFolderClientHash)) #RIGHT OUTER JOIN = LEFT OUTER JOIN with the tables reversed
    $prettyClientsAndOppsLeft = get-linqLeftJoin -dataSet1 $prettyClients -dataSet1PropertyToCompare NetSuiteClientId -dataSet2 $prettyOpps -dataSet2PropertyToCompare TermOppClientId -dataSet1PropertiesToReturn $($(get-prettyNetClientHash) + $(get-prettyTermClientHash) + $(get-prettyFolderClientHash)) -dataSet2PropertiesToReturn $($(get-prettyNetOppHash) + $(get-prettyTermOppHash) + $(get-prettyFolderOppHash) + $(get-prettyNetClientHash) + $(get-prettyTermClientHash) + $(get-prettyFolderClientHash)) #LEFT OUTER JOIN
    $prettyClientsWithoutOpps = $prettyClientsAndOppsLeft | Where-Object {[string]::IsNullOrWhiteSpace($_.NetSuiteOppId)} #This is "Clients LEFT OUTER JOIN Opps" without the INNER JOIN results
    Write-Verbose "[$($prettyClientsAndOppsRight.Count)] PrettyOpps linked to PrettyClients"
    Write-Verbose "[$($prettyClientsWithoutOpps.Count)] PrettyClients without PrettyOpps identified (not orphaned, as not all Clients _should_ have Opps)"
    Write-Verbose "[$($prettyClientsAndOppsRight.Count + $prettyClientsWithoutOpps.Count)] ([$($prettyClientsAndOppsRight.Count)]+[$($prettyClientsWithoutOpps.Count)]) total PrettyClientOpps"
    $prettyClientOpps = $prettyClientsAndOppsRight + $prettyClientsWithoutOpps
    }
Write-Host "[$($prettyClientOpps.Count)] PrettyClients linked to PrettyOpps in $(format-measureCommandResults $makingPrettyClientOpps)"

#Next we need to match ClientOpps to Projs (and get Projs without Opps or Clients)
#This is easiest to understand with a 3-circle Venn diagram. We have the FULL OUTER JOIN of Clients and Opps. If Projs intersects with either/both, then the only records we are missing are the RIGHT MINUS JOIN
Write-Host "This next bit will take a while..."
$makingPrettyClientOppProjs = Measure-Command {
    $prettyClientOppsProjsLeft = get-linqLeftJoin -dataSet1 $prettyClientOpps -dataSet1PropertyToCompare TermOppProjId -dataSet2 $prettyProjs -dataSet2PropertyToCompare NetSuiteProjId -dataSet1PropertiesToReturn $((get-prettyNetClientHash) + $(get-prettyTermClientHash) + $(get-prettyFolderClientHash) + $(get-prettyNetOppHash) + $(get-prettyTermOppHash) + $(get-prettyFolderOppHash)) -dataSet2PropertiesToReturn $($(get-prettyNetProjHash) + $(get-prettyTermProjHash) + $(get-prettyFolderProjHash))
    Write-Verbose "[$($prettyClientOppsProjsLeft.Count)] PrettyClientOpps linked to PrettyProjs"
    }
Write-Host "[$($prettyClientOppsProjsLeft.Count)] PrettyClientOpps linked to PrettyProjs in $(format-measureCommandResults $makingPrettyClientOppProjs)"

$makingPrettyClientOppProjs2 = Measure-Command {
    $prettyClientOppsProjsRight = get-linqLeftJoin -dataSet1 $prettyProjs -dataSet1PropertyToCompare NetSuiteProjId -dataSet2 $prettyClientOpps -dataSet2PropertyToCompare TermOppProjId -dataSet1PropertiesToReturn $($(get-prettyNetProjHash) + $(get-prettyTermProjHash) + $(get-prettyFolderProjHash)) -dataSet2PropertiesToReturn $((get-prettyNetClientHash) + $(get-prettyTermClientHash) + $(get-prettyFolderClientHash) + $(get-prettyNetOppHash) + $(get-prettyTermOppHash) + $(get-prettyFolderOppHash))
    $prettyProjsWithoutClientOpps = $prettyClientOppsProjsRight | Where-Object {[string]::IsNullOrWhiteSpace($_.TermOppProjId) -and [string]::IsNullOrWhiteSpace($_.NetSuiteProjId)}
    Write-Verbose "[$($prettyClientOppsProjsRight.Count)] PrettyProjs without PrettyClientOpps identified (not orphaned, as (weirdly) not all Projs have Clients/Opps)"
    Write-Verbose "[$($prettyClientOppsProjsRight.Count + $prettyProjsWithoutClientOpps.Count)] ([$($prettyClientOppsProjsRight.Count)]+[$($prettyProjsWithoutClientOpps.Count)]) total PrettyClientOppProjs"
    $prettyClientOppProjs = $prettyClientOppsProjsLeft + $prettyProjsWithoutClientOpps
    }
Write-Host "[$($prettyClientOppsProjsLeft.Count)] PrettyClientOpps linked to PrettyProjs in $(format-measureCommandResults $($makingPrettyClientOppProjs + $makingPrettyClientOppProjs2))" -f Yellow

Write-Host "BeautifulObject preparation took $(format-measureCommandResults $($makingPrettyClientOpps+$makingPrettyClientOppProjs+$makingPrettyClientOppProjs2)) in total" -f Magenta

#endregion
#Confirmed $prettyOppsWithoutClients are included in $prettyOppsAndClientsLeft
#$prettyOppsWithoutClients = $prettyOppsAndClientsLeft | Where-Object {[string]::IsNullOrWhiteSpace($_.NetSuiteClientId)} 
#Write-Host "`$prettyOppsWithoutClients.Count = [$($prettyOppsWithoutClients.Count)]"

$prettyProjsAndClientsLeft = get-linqLeftJoin -dataSet1 $prettyProjs -dataSet1PropertyToCompare TermProjClientId -dataSet2 $prettyClients -dataSet2PropertyToCompare NetSuiteClientId -dataSet1PropertiesToReturn $($(get-prettyNetProjHash) + $(get-prettyTermProjHash) + $(get-prettyFolderProjHash)) -dataSet2PropertiesToReturn $($(get-prettyNetOppHash) + $(get-prettyTermOppHash) + $(get-prettyFolderOppHash) + $(get-prettyNetClientHash) + $(get-prettyTermClientHash) + $(get-prettyFolderClientHash))
$prettyProjsWithoutClients = $prettyProjsAndClientsLeft | Where-Object {[string]::IsNullOrWhiteSpace($_.TermProjClientId)} 

$prettyProjsAndOppsLeft = get-linqLeftJoin -dataSet1 $prettyProjs -dataSet1PropertyToCompare NetSuiteProjId -dataSet2 $prettyOpps -dataSet2PropertyToCompare TermOppProjId -dataSet1PropertiesToReturn $($(get-prettyNetProjHash) + $(get-prettyTermProjHash) + $(get-prettyFolderProjHash)) -dataSet2PropertiesToReturn $($(get-prettyNetOppHash) + $(get-prettyTermOppHash) + $(get-prettyFolderOppHash) + $(get-prettyNetClientHash) + $(get-prettyTermClientHash) + $(get-prettyFolderClientHash))
$prettyProjsWithoutOpps = $prettyProjsAndOppsLeft | Where-Object {[string]::IsNullOrWhiteSpace($_.TermOppProjId)} 
$prettyProjsAndOppsLeft2 = get-linqLeftJoin -dataSet1 $prettyProjs -dataSet1PropertyToCompare NetSuiteOppProjId -dataSet2 $prettyOpps -dataSet2PropertyToCompare TermOppProjId -dataSet1PropertiesToReturn $($(get-prettyNetProjHash) + $(get-prettyTermProjHash) + $(get-prettyFolderProjHash)) -dataSet2PropertiesToReturn $($(get-prettyNetOppHash) + $(get-prettyTermOppHash) + $(get-prettyFolderOppHash) + $(get-prettyNetClientHash) + $(get-prettyTermClientHash) + $(get-prettyFolderClientHash))
$prettyProjsWithoutOpps2 = $prettyProjsAndOppsLeft2 | Where-Object {[string]::IsNullOrWhiteSpace($_.TermOppProjId)} 
$prettyProjsWithoutOpps.Count
$prettyProjsWithoutOpps2.Count

$prettyClientsAndOppsAndProjects = get-linqLeftJoin -dataSet1 $prettyClientsAndOpps -dataSet1PropertyToCompare TermOppProjId -dataSet2 $prettyProjs -dataSet2PropertyToCompare NetSuiteProjId -dataSet1PropertiesToReturn $($(get-prettyNetOppHash) + $(get-prettyTermOppHash) + $(get-prettyFolderOppHash) + $(get-prettyNetClientHash) + $(get-prettyTermClientHash) + $(get-prettyFolderClientHash)) -dataSet2PropertiesToReturn $($(get-prettyNetProjHash) + $(get-prettyTermProjHash) + $(get-prettyFolderProjHash))

$prettyOppsAndProjects | % {
    $thisPrettyLittleThing = $_
    $thisPrettyLittleThing | Add-Member -MemberType NoteProperty -Name Errors -Value @()
    if([string]::IsNullOrWhiteSpace($thisPrettyLittleThing.TermClientDriveId)){$thisPrettyLittleThing.Errors += ""}
    }

<#region Types of problems:
Clients
    ClientTerm cannot be created from NetClient
        Duplicate Client Term already exists
            NetClient.Name -eq TermClient.Name -and NetClient.id -ne TermClient.TermClientId
    ClientFolder cannot be created from ClientTerm
        Can this happen?

Opps
    NetOpp does not belong to a client
        NetOpp.entity (parent) -eq $null

    OppTerm cannot be created from NetOpp
        Illegal characters in NetOpp Name
            test-validNameForTermStore(NetOpp.Name) -eq $false
        Duplicate OppTerm already exists (should be impossible - NetOpp Name contains unique code)
            NetOpp.Name -eq TermOpp.Name -and NetOpp.id -ne TermOpp.TermOppId

    OppFolder cannot be created from OppTerm
        TermClientId missing
            NetOpp.entity (parent) -ne $null -and TermOpp.TermClientId -eq $null
        Related ClientTerm missing
            ($TermClients | ? {$_.NetSuiteClientId -eq $thisOppTerm.TermClientId}) -eq $null
        Derived ClientFolderDriveId missing
            ($TermClients | ? {$_.NetSuiteClientId -eq $thisOppTerm.TermClientId}) -ne $null -and ($TermClients | ? {$_.NetSuiteClientId -eq $thisOppTerm.TermClientId}).ClientFolderDriveId -eq $null
        ClientFolderDriveItemId missing
            TermOpp.ClientFolderDriveItemId -eq $null
        Illegal characters in OppTerm
            test-validNameForSharePointFolder(TermOpp) -eq $false
    
    OppFolder exists in wrong ClientFolder
        OppFolder FolderDriveItemId -ne OppTerm TermDriveItemId
    OppFolder exists in addition to ProjFolder
        OppFolder.Name -eq OppTerm.Name -and $OppTerm.ProjectId -ne $null
    

Projs
    NetProj does not belong to a client
        NetProj.parent -eq $null
    ProjTerm cannot be created from NetProj
        Illegal characters in NetProj Name
            test-validNameForTermStore(NetProj.Name) -eq $false
        Duplicate ProjTerm already exists (should be impossible - NetProj Name contains unique code)
            NetProj.Name -eq TermProj.Name -and NetProj.id -ne TermProj.TermOppId
    ProjFolder cannot be created from ProjTerm
        TermClientId missing
        Related ClientTerm missing
        Derived ClientFolderDriveId missing
        Derived ClientFolder missing
        Illegal characters in ProjTerm

    ProjFolder exists in wrong ClientFolder
        Duplicate ProjFolder
        Misplaced ProjFolder



#>

#Error Checking
[array]$netProjectsWithNoClientId = $allNetSuiteProjs | ? {[string]::IsNullOrWhiteSpace($_.parent.id)} #Find the weird Projects with no Client (as we can't create folders for them anyway)
[array]$netProjsWithOvertlyDuffNames = $allNetSuiteProjs | ? {$(test-validNameForTermStore -stringToTest $_.entityId) -eq $false}

[array]$netOppsWithNoClientId = $allNetSuiteOpps | ? {[string]::IsNullOrWhiteSpace($_.parent.id)} #Find the weird Projects with no Client (as we can't create folders for them anyway)
[array]$netOppsWithOvertlyDuffNames = $allNetSuiteOpps | ? {$(test-validNameForTermStore -stringToTest $_.title) -eq $false}

[array]$netClientsWithOvertlyDuffNames = $allNetSuiteClients | ? {$(test-validNameForTermStore -stringToTest $_.companyName) -eq $false}

$netProjectsWithNoClientId.entityId
Write-Warning "[$($netProjectsWithNoClientId.Count)] Projects with no Clients (there are at least 74 known internal/broken projects like this)"
$netProjsWithOvertlyDuffNames.entityId
if($netProjsWithOvertlyDuffNames.Count -gt 0){
    Write-Warning "[$($netProjsWithOvertlyDuffNames.Count)] Projects have illegal characters in their name, which means the Term cannot be created:"#`r`n`t[$($netProjsWithOvertlyDuffNames.entityId -join ']`r`n`t`[')]"
    $netProjsWithOvertlyDuffNames | % {Write-Warning "`t[$($_.entityId)] contains [$(test-validNameForTermStore -stringToTest $_.entityId -returnSpecifcProblem)]"}
    }
$netOppsWithNoClientId.entityId
$netOppsWithOvertlyDuffNames.entityId
if($netOppsWithOvertlyDuffNames.Count -gt 0){
    Write-Warning "[$($netOppsWithOvertlyDuffNames.Count)] Projects have illegal characters in their name (; < > \ | `t), which means the Term cannot be created:"#`r`n`t[$($netOppsWithOvertlyDuffNames.title -join "]$([Environment]::NewLine)`t[")]"
    $netOppsWithOvertlyDuffNames | % {Write-Warning "`t[$($_.tranId) $($_.title)] contains [$(test-validNameForTermStore -stringToTest $_.title -returnSpecifcProblem)]"}
    }
$netClientsWithOvertlyDuffNames.entityId


# Num Unique Opps
# + Projects without Opps
# + Clients without Opps or Projects
$numOpps = $($allNetSuiteOpps | select id -Unique).Count
$numProjectsWithoutOpps = compare-object $allNetSuiteOpps $allNetSuiteProjs -Property "" -PassThru | ? {$_.SideIndicator -eq "=>"}
$numClientsWithoutOpps = $(compare-object $allNetSuiteOpps $allNetSuiteClients -Property "" -PassThru | ? {$_.SideIndicator -eq "=>"}).Count
$numClientsWithoutProjs = $(compare-object $allNetSuiteProjs $allNetSuiteClients -Property "" -PassThru | ? {$_.SideIndicator -eq "=>"}).Count
$numClientsWithoutOppsOrProjs = $(compare-object $numClientsWithoutOpps $numClientsWithoutProjs -Property "" -PassThru -IncludeEqual -ExcludeDifferent | ? {$_.SideIndicator -eq "=="}).Count
