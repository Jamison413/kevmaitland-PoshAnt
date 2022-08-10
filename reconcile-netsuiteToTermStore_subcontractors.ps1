[cmdletbinding()]
param(
    [Parameter(Mandatory = $false, Position = 0)]
        [string]$deltaSync = $true #Specifies whether we are doing a full or incremental sync.
    )

if($PSCommandPath){
    $InformationPreference = 2
    $VerbosePreference = 0
    $logFileLocation = "C:\ScriptLogs\"
    if($deltaSync -eq $true){$suffix = "_deltaSync"}
    else{$suffix = "_fullSync"}
    #$suffix = "_fullSync"
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))$suffix`_Transcript_$(Get-Date -Format "yyyy-MM-dd").log"
    Start-Transcript $transcriptLogName -Append
    }

function export-encryptedCache(){
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
            [AllowNull()]
            [array]$netObjects 
        ,[Parameter(Mandatory = $true, Position = 1)]
            [ValidateSet("Client","Subcontractor","Employee","Opportunity","Project")]
            [array]$netObjectType 
        )
    
    $netObjectSchema = [ordered]@{}
    $netObjects | % {
        $thisNetObject = $_
        Compare-Object -ReferenceObject @($($netObjectSchema.Keys) | % {$_.ToString()} | Select-Object) -DifferenceObject $thisNetObject.PSObject.Properties.Name  | ? {$_.SideIndicator -eq "=>"} | % {
            $netObjectSchema.Add($_.InputObject,$null)# | Add-Member -MemberType NoteProperty -Name $_ -Value $null
            #Write-Host "Adding [$($_.InputObject)] from [$($thisNetObject.id)]"
            }
        }
    $prettyNetSuiteObjects = @($null)*$netObjects.Count
    $i=0
    $netObjects | %{
        $thisNetObject = $_
        $prettyNetSuiteObjects[$i] = New-Object -TypeName PSCustomObject -Property $netObjectSchema
        $thisNetObject.PSObject.Properties.Name | % {
            $prettyNetSuiteObjects[$i].$_ = $(convertTo-localisedSecureString $thisNetObject.$_)
            }
        $i++
        }
        
    $prettyNetSuiteObjects | Select-Object @($($netObjectSchema.Keys) | % {$_.ToString()} | Select-Object) | Export-Csv -Path "$env:TEMP\Net$netObjectType.csv" -NoTypeInformation -Force -Encoding UTF8
    }
function process-comparison(){
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
            [AllowNull()]
            [array]$subsetOfNetObjects 
        ,[Parameter(Mandatory = $true, Position = 1)]
            [AllowNull()]
            [array]$allTermObjects 
        ,[Parameter(Mandatory = $true, Position = 2)]
            [string]$idInCommon 
        ,[Parameter(Mandatory = $true, Position = 3)]
            [string]$propertyToTest
        ,[Parameter(Mandatory = $false, Position = 4)]
            [switch]$validate
        )

    #compare-object jiggery-pokery documented with pictures on IT Site: https://anthesisllc.sharepoint.com/sites/Resources-IT/_layouts/15/DocIdRedir.aspx?ID=HXX7CE52TSD2-1759992947-266
    #Prerequisite: $subsetOfNetObjects, $allTermObjects
    [array]$correspondingSubsetOfTermObjects = Compare-Object -ReferenceObject @($allTermObjects | Select-Object) -DifferenceObject @($subsetOfNetObjects | Select-Object) -Property $idInCommon -PassThru -IncludeEqual -ExcludeDifferent
    [array]$comparisonOfPropertyToTest = Compare-Object -ReferenceObject @($subsetOfNetObjects | Select-Object) -DifferenceObject @($correspondingSubsetOfTermObjects | Select-Object) -Property $idInCommon,$propertyToTest -PassThru -IncludeEqual
    [array]$netObjectsWithMismatchedProperty = $comparisonOfPropertyToTest | ? {$_.SideIndicator -eq "<="} | Sort-Object $idInCommon
    [array]$correspondingTermObjectsWithMismatchedProperty = $comparisonOfPropertyToTest | ? {$_.SideIndicator -eq "=>"} | Sort-Object $idInCommon
    [array]$netObjectsWithMatchingProperty = $comparisonOfPropertyToTest | ? {$_.SideIndicator -eq "=="} 
    
    Write-Verbose "subsetOfNetObjects.Count = `t`t`t`t[$($subsetOfNetObjects.Count)]";Write-Verbose "correspondingSubsetOfTermObjects.Count = `t[$($correspondingSubsetOfTermObjects.Count)] (should be equal)";Write-Verbose "comparisonOfPropertyToTest.Count = `t`t[$($comparisonOfPropertyToTest.Count)] (<=[$(($netObjectsWithMismatchedProperty).Count)]  ==[$($netObjectsWithMatchingProperty.Count)]  =>[$(($correspondingTermObjectsWithMismatchedProperty).Count)]) (<= should equal =>)"
    if($validate){
        if($netObjectsWithMismatchedProperty.Count -ne $correspondingTermObjectsWithMismatchedProperty.Count){
            Write-Verbose "`"<=`" array Count [$($netObjectsWithMismatchedProperty.Count)] does not equal `"=>`" array Count [$($correspondingTermObjectsWithMismatchedProperty.Count)]: Invalid output"
            $invalid = $true
            }
        for($i=0; $i -lt $correspondingTermObjectsWithMismatchedProperty.Count; $i++){
            if($correspondingTermObjectsWithMismatchedProperty[$i]."$idInCommon" -ne $netObjectsWithMismatchedProperty[$i]."$idInCommon"){
                Write-Verbose "Property [$propertyToTest] for array `"<=`" item [$i] [$($netObjectsWithMismatchedProperty[$i]."$idInCommon")] does not equal `"=>`" array item [$($correspondingTermObjectsWithMismatchedProperty[$i]."$idInCommon")]: Invalid output"
                $invalid = $true
                }
            }
        if($invalid -eq $true){return $false} #Return $false instead of the invalid comparison data
        }
    @{"<=" = $netObjectsWithMismatchedProperty
        "==" = $netObjectsWithMatchingProperty
        "=>" = $correspondingTermObjectsWithMismatchedProperty
        }
    }
function process-orphanedTerm(){
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
            [Microsoft.SharePoint.Client.Taxonomy.TermSetItem]$orphanedTerm 
        )
    #First confirm Term is definitely orphaned by re-querying NetSuite (the API seems a little unreliable ~1/3000 errors)
    switch($orphanedTerm.TermSet.Name){
        "Clients" {
            $netRecord = get-netSuiteClientsFromNetSuite -clientId $orphanedTerm.NetSuiteClientId -netsuiteParameters $netSuiteProductionParamaters
            $netRecord | Add-Member -MemberType NoteProperty -Name "UniversalName" -Value $orphanedTerm.UniversalClientName
            }
        "Opportunities" {
            $netRecord = get-netSuiteOpportunityFromNetSuite -query "?q=Id EQUAL $($orphanedTerm.NetSuiteOppId)" -netsuiteParameters $netSuiteProductionParamaters
            $netRecord | Add-Member -MemberType NoteProperty -Name "UniversalName" -Value $orphanedTerm.UniversalOppName
            }
        "Projects" {
            $netRecord = get-netSuiteProjectFromNetSuite -query "?q=Id EQUAL $($orphanedTerm.NetSuiteProjectId)" -netsuiteParameters $netSuiteProductionParamaters
            $netRecord | Add-Member -MemberType NoteProperty -Name "UniversalName" -Value $orphanedTerm.UniversalProjName
            }
        "Subcontractors" {
            $netRecord = get-netSuiteSubcontractorsFromNetSuite -query "?q=Id EQUAL $($orphanedTerm.NetSuiteSubcontractorId)" -netsuiteParameters $netSuiteProductionParamaters
            $netRecord | Add-Member -MemberType NoteProperty -Name "UniversalName" -Value $orphanedTerm.UniversalProjName
            }
        }
    
    if(![string]::IsNullOrWhiteSpace($netRecord.id)){ #This prevents dropped GET requests from NetSuite being incorrectly deleted
        #Record exists in NetSuite - looks like an initial retrieval problem.
        Write-Warning "NetSuite [$($orphanedTerm.TermSet.Name)] record [$($netRecord.UniversalName)] was flagged as Orphan, but still exists in NetSuite! Terminating process-orphanedTerm()"
        return $false
        }


    do{
        try{
            #Copy Term to OrphanedTerms
            Write-Host "`t`tBacking up orphaned Term [$($orphanedTerm.TermSet.Group.Name)][$($orphanedTerm.TermSet.Name)][$($orphanedTerm.Name)][$($orphanedTerm.id)] to [$($orphanedTerm.TermSet.Group.Name)][Orphaned$($orphanedTerm.TermSet.Name)][$($orphanedTerm.Name)$i]"
            $backedUpTerm = New-PnPTerm -TermGroup $($orphanedTerm.TermSet.Group.Name) -TermSet "Orphaned$($orphanedTerm.TermSet.Name)" -Name $("$($orphanedTerm.Name)$i")  -Lcid 1033 -CustomProperties $([hashtable]::new($orphanedTerm.CustomProperties)) -ErrorAction Stop
            if(![string]::IsNullOrWhiteSpace($backedUpTerm.Name)){
                $success = $true
                }
            }
        catch{
            if($_.Exception -match "TermStoreErrorCodeEx:There is already a term with the same default label and parent term."){
                Write-Verbose $_.Exception
                #Do nothing - just continue through the loop, incrementing $i until we find an empty value
                }
            else{ #If we get a different error, report it and move on
                return $(get-errorSummary -errorToSummarise $_)
                }
            }
        if($backedUpTerm){
            if($backedUpTerm.Name -match [Regex]::Escape($orphanedTerm.Name)){
                #Delete original Term
                try{
                    Write-Host "`t`tDeleting orphaned Term [$($orphanedTerm.TermSet.Group.Name)][$($orphanedTerm.TermSet.Name)][$($orphanedTerm.Name)][$($orphanedTerm.id)][$($orphanedTerm.NetSuiteClientId)]"
                    Remove-PnPTaxonomyItem -TermPath "$($orphanedTerm.TermSet.Group.Name)|$($orphanedTerm.TermSet.Name)|$($orphanedTerm.Name)" -Force -Verbose
                    return $true
                    }
                catch{
                    return $(get-errorSummary -errorToSummarise $_)
                    }
                }
            }
        #else{return "Weird - the Term [$($orphanedTerm.Name)] was backed up, but its new name is [$($backedUpTerm.Name)], which doesn't look right. Not deleting the original Term."}
        $i++
        }
    until($success -eq $true)
    }
function test-validNameForTermStore(){
    [cmdletbinding()]
    param(
         [Parameter(Mandatory = $true,Position=0)]
            [string]$stringToTest
        )
    
    #if($stringToTest -match ';|"|<|>|\||\t'){$false} #These are Microsoft's list of invalid characters
    if($stringToTest -match ';|<|>|\||\t'){$false} #New-PnPTerm silently recodes " and \ but \ will still break foldernames later
    else{$true}
    }

$timeForFullCycle = Measure-Command {

    #region GetData
    $sharePointAdmin = "kimblebot@anthesisgroup.com"
    #convertTo-localisedSecureString "KimbleBotPasswordHere"
    $sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Downloads\KimbleBot.txt) 
    $adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
    Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com" -Credentials $adminCreds
    $netSuiteProductionParamaters = $(get-netSuiteParameters -connectTo Production)

        #region getSubcontractorData
        $termSubcontractorRetrieval = Measure-Command {
            $pnpTermGroup = "Kimble"
            $pnpTermSet = "Subcontractors"
            $allSubcontractorTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes TermSet,TermSet.Group,TermStore,CustomProperties | ? {$_.IsDeprecated -eq $false}
            @($allSubcontractorTerms | Select-Object) | % {
                Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteSubcontractorId -Value $($_.CustomProperties.NetSuiteId) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name DriveSubcontractorId -Value $($_.CustomProperties.GraphDriveId) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name TermSubcontractorId -Value $($_.Id) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name TermSubcontractorName -Value $($_.name) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteLastModifiedDate -Value $($_.CustomProperties.NetSuiteLastModifiedDate) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalSubcontractorName -Value $($_.Name) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalSubcontractorNameSanitised -Value $(sanitise-forNetsuiteIntegration $_.Name) -Force #This helps to avoid weird encoding, diacritic and special character problems when comparing strings
                }
            }
        Write-Host "[$($allSubcontractorTerms.Count)] subcontractors retrieved from TermStore in [$($termSubcontractorRetrieval.TotalSeconds)] seconds"

        $netSubcontractorRetrieval = Measure-Command {
            $netQuery =  "?q=companyName CONTAIN_NOT `"Anthesis`"" #Excludes any Companies with "Anthesis" in the companyName
            $netQuery += " AND companyName CONTAIN_NOT `"Best Foot Forward Ltd`"" #Excludes any Companies with "(intercompany project)" in the companyName
            $netQuery += " AND companyName CONTAIN_NOT `"Caleb Management Service`"" #Excludes any Companies with "(intercompany project)" in the companyName
            $netQuery += " AND companyName CONTAIN_NOT `"LRS Consultancy Ltd`"" #Excludes any Companies with "(intercompany project)" in the companyName
            $netQuery += " AND companyName CONTAIN_NOT `"LRS Environmental Ltd`"" #Excludes any Companies with "(intercompany project)" in the companyName
            $netQuery += " AND companyName CONTAIN_NOT `"Lavola 1981 SAU`"" #Excludes any Companies with "(intercompany project)" in the companyName
            $netQuery += " AND companyName CONTAIN_NOT `"Lavola Andora SA`"" #Excludes any Companies with "(intercompany project)" in the companyName
            $netQuery += " AND companyName CONTAIN_NOT `"Lavola Columbia`"" #Excludes any Companies with "(intercompany project)" in the companyName
            $netQuery += " AND companyName CONTAIN_NOT `"Lavola Sucursal Colombia`"" #Excludes any Companies with "(intercompany project)" in the companyName
            $netQuery += " AND companyName CONTAIN_NOT `"Media4Change`"" #Excludes any Companies with "(intercompany project)" in the companyName
            $netQuery += " AND companyName CONTAIN_NOT `"The Goodbrand Works Ltd`"" #Excludes any Companies with "(intercompany project)" in the companyName
            #$netQuery += " AND companyName START_WITH_NOT `"x `"" #Excludes any Companies that begin with "x " in the companyName
            #$netQuery += " AND isPerson IS $false" #Exclude Individuals (until we figure out how to deal with them) # We need to _include_ them for Subcontractors
            #$netQuerySubcontractors = "?q=isPerson IS $false" #Exclude Individuals (until we figure out how to deal with them)
            #$netQuerySubcontractors += " AND entityStatus ANY_OF_NOT [6, 7]" #Excludes LEAD-Unqualified and LEAD-Qualified (https://XXX.app.netsuite.com/app/crm/sales/customerstatuslist.nl?whence=)
            if($deltaSync -eq $true){
                [datetime]$lastProcessed = $($allSubcontractorTerms | sort {$_.CustomProperties.NetSuiteLastModifiedDate} | select -Last 1).CustomProperties.NetSuiteLastModifiedDate
                $netQuerySubcontractors += " AND lastModifiedDate ON_OR_AFTER `"$($(Get-Date $lastProcessed -Format g))`"" #Excludes any Companies that haven;t been updated since X
                }
            #[array]$netSuiteSubcontractorsToCheck = get-netSuiteSubcontractorsFromNetSuite -query $netQuerySubcontractors -netsuiteParameters $netSuiteProductionParamaters
            #[psobject[]]$netSuiteSubcontractorsToCheck = get-netSuiteSubcontractorsFromNetSuite -query $netQuerySubcontractors -netsuiteParameters $netSuiteProductionParamaters
            #$netSuiteSubcontractorsToCheck = @()
            $allNetSuiteSubcontractors = get-netSuiteSubcontractorsFromNetSuite -query $netQuerySubcontractors -netsuiteParameters $netSuiteProductionParamaters
            if([string]::IsNullOrEmpty($allNetSuiteSubcontractors.Count)){$allNetSuiteSubcontractors = @()}
            @($allNetSuiteSubcontractors | Select-Object) | % {
                Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteSubcontractorId -Value $($_.Id) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteSubcontractorName -Value $($_.companyName) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteLastModifiedDate -Value $($_.lastModifiedDate) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalSubcontractorName -Value $($_.companyName) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalSubcontractorNameSanitised -Value $(sanitise-forNetsuiteIntegration $_.companyName) -Force
                }
            }
        Write-Host "[$($allNetSuiteSubcontractors.Count)] subcontractors retrieved from NetSuite in [$($netSubcontractorRetrieval.TotalSeconds)] seconds or [$($netSuiteSubcontractorsToCheck.Count / $netSubcontractorRetrieval.TotalMinutes)] per minute"
        if($allNetSuiteSubcontractors.Count -gt 0 -and $deltaSync -eq $false){export-encryptedCache -netObjects $allNetSuiteSubcontractors -netObjectType Subcontractor}

        $netSuiteSubcontractorsToCheck = @($allNetSuiteSubcontractors | Select-Object) #Remove any $nulls that 401'ed/disappeared in transit
        if($deltaSync -eq $true){
            [array]$processedAtExactlyLastTimestamp = $netSuiteSubcontractorsToCheck | ? {$(Get-Date $_.lastModifiedDate) -ge $(Get-Date $lastProcessed)} #Find how many Subcontractors match the $lastProcessed timestamp exactly
            if($processedAtExactlyLastTimestamp.Count -eq 1){$netSuiteSubcontractorsToCheck = $netSuiteSubcontractorsToCheck | ? {$netSuiteSubcontractorsToCheck.id -notcontains $processedAtExactlyLastTimestamp[0].id}} #If it's exactly one, exclude it from processing (as we've already processed it on a previous cycle)
            }

    #Weed out the duff records now to prevent errors later on.
        [array]$netSuiteSubcontractorsToCheckWithoutDuffGets = $netSuiteSubcontractorsToCheck | ? {$_ -notmatch "PSMessageDetails"}
        if($netSuiteSubcontractorsToCheckWithoutDuffGets.Count -ne $netSuiteSubcontractorsToCheck.Count){
            Write-Host "`t[$($netSuiteSubcontractorsToCheck.Count-$netSuiteSubcontractorsToCheckWithoutDuffGets.Count)] REST errors occured retrieving data from NetSuite. Discarding these records, leaving [$($netSuiteSubcontractorsToCheck.Count)] records to process."
            $netSuiteSubcontractorsToCheck = $netSuiteSubcontractorsToCheckWithoutDuffGets
            }

        [array]$newSubcontractorsWithoutOvertlyDuffNames = $netSuiteSubcontractorsToCheck | ? {$(test-validNameForTermStore -stringToTest $_.UniversalSubcontractorName) -eq $true}
        if($newSubcontractorsWithoutOvertlyDuffNames.Count -ne $netSuiteSubcontractorsToCheck.Count){
            Write-Host "`t[$($netSuiteSubcontractorsToCheck.Count - $newSubcontractorsWithoutOvertlyDuffNames.Count)] of these contain illegal characters for Terms, so I'll just process the remaining [$($newSubcontractorsWithoutOvertlyDuffNames.Count)]"
            $netSuiteSubcontractorsToCheck = $newSubcontractorsWithoutOvertlyDuffNames
            }


        #endregion
    #endregion

    $SubcontractorReconcile = Measure-Command {

        $subcontractorComparison = Compare-Object -ReferenceObject @($netSuiteSubcontractorsToCheck | Select-Object) -DifferenceObject @($allSubcontractorTerms | Select-Object) -Property NetSuiteSubcontractorId -IncludeEqual -PassThru #Wrapped in @($ | select) to remove $nulls
        if($deltaSync -eq $false){[array]$orphanedSubcontractors = $subcontractorComparison | ? {$_.SideIndicator -eq "=>"}}
        [array]$newSubcontractors = $subcontractorComparison | ? {$_.SideIndicator -eq "<="}
        [array]$existingSubcontractors = $subcontractorComparison | ? {$_.SideIndicator -eq "=="}

        #region orphanedSubcontractors
            #Copy Term to OrphanedTerms
            #Delete original Term
        if($deltaSync -eq $false){
            Write-Host "`tProcessing [$($orphanedSubcontractors.Count)] orphaned Subcontractors"
            @($orphanedSubcontractors | select-object) | % {
                $thisOrphanedTerm = $_
                $processedOrphanedTerm = process-orphanedTerm -orphanedTerm $thisOrphanedTerm
                if($processedOrphanedTerm -ne $true){
                    [array]$duffOrphanedSubcontractors += $thisOrphanedTerm
                    }
                }
            }
        #endregion

        #region newSubcontractors
            #Create new Term
        Write-Host "`tProcessing [$($newSubcontractors.Count)] new Subcontractors"
        if($deltaSync -eq $false){
            #Fisrt exclude any duplicates companies in NetSuite 
            [array]$duplicateNetSuiteSubcontractors = $($netSuiteSubcontractorsToCheck | Group-Object -Property {$_.companyName} | ? {$_.Count -ne 1}).Group
            $newSubcontractorsDeduped = $newSubcontractors | ? {$duplicateNetSuiteSubcontractors.id -notcontains $_.id}
            Write-Host "`t`tExcluding [$($newSubcontractors.Count - $newSubcontractorsDeduped.count)] of these as they are duplicates of existing Subcontractors in NetSuite"
            #Send e-mail report too!!
            #Send e-mail report too!!
            #Send e-mail report too!!
            $newSubcontractors = $newSubcontractorsDeduped
            }

        @($newSubcontractors | select-object) | % {
            $pnpTermGroup = "Kimble"
            $pnpTermSet = "Subcontractors"
            $thisNewSubcontractor = $_
            Write-Host "`t`tProcessing new Subcontractor [$($thisNewSubcontractor.companyName)]"
            try{
                Write-Host "`t`t`tCreating new Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisNewSubcontractor.companyName)][@{NetSuiteId=$($thisNewSubcontractor.id);NetSuiteLastModifiedDate=$($thisNewSubcontractor.lastModifiedDate);flagForReprocessing=$true]"
                $newTerm = New-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Name $thisNewSubcontractor.companyName -Lcid 1033 -CustomProperties @{NetSuiteId=$thisNewSubcontractor.id;NetSuiteLastModifiedDate=$thisNewSubcontractor.lastModifiedDate;flagForReprocessing=$true} -ErrorAction Stop
                }
            catch{ #We don't handle any specific errors here. If there's already a term with this Subcontractor's name then the new NetSuite subcontractor is a duplicate (and we'll probably want to keep the older record)
                Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                [array]$duffnewSubcontractors += @($thisNewSubcontractor,$(get-errorSummary -errorToSummarise $_))
                }
            }
        #endregion

        #region existingSubcontractors
            #Update Term
                #Has Name changed?
                    #Yes: Update TermSubcontractorName, NetSuiteLastModifiedDate & flagForReproccessing
                    #No: Update NetSuiteLastModifiedDate
    
                #Has Name changed?
        $subcontractorNameComparison = process-comparison -subsetOfNetObjects $existingSubcontractors -allTermObjects $allSubcontractorTerms -idInCommon NetSuiteSubcontractorId -propertyToTest UniversalSubcontractorNameSanitised -validate
        [array]$existingNetSubcontractorsWithChangedNames  = $subcontractorNameComparison["<="]
        [array]$existingTermSubcontractorsWithChangedNames  = $subcontractorNameComparison["=>"]
        #Write-Host "existingSubcontractors.Count = `t`t`t`t`t`t`t[$($existingSubcontractors.Count)]";Write-Host "subcontractorNameComparison.Count = `t[$($subcontractorNameComparison.Count)] (<=[$(($existingNetSubcontractorsWithChangedNames).Count)]  ==[$(($subcontractorNameComparison["=="]).Count)]  =>[$(($existingTermSubcontractorsWithChangedNames).Count)])"

        Write-Host "`tProcessing [$($existingSubcontractors.Count)] existing Subcontractors"
        Write-Host "`t`tProcessing [$($existingTermSubcontractorsWithChangedNames.Count)] existing Subcontractors with changed names"
        for($i=0;$i -lt $existingTermSubcontractorsWithChangedNames.Count; $i++){
                    #Yes: Update TermSubcontractorName, NetSuiteLastModifiedDate & flagForReproccessing
            Write-Host "`t`t`tRenaming Term `t[$($existingTermSubcontractorsWithChangedNames[$i].Name)][$($existingTermSubcontractorsWithChangedNames[$i].Id)][$($existingTermSubcontractorsWithChangedNames[$i].NetSuiteSubcontractorId)]"
            Write-Host "`t`t`tto:`t`t`t`t[$($existingNetSubcontractorsWithChangedNames[$i].UniversalSubcontractorName)][$($existingNetSubcontractorsWithChangedNames[$i].NetSuiteSubcontractorId)]"
            $existingTermSubcontractorsWithChangedNames[$i].Name = $existingNetSubcontractorsWithChangedNames[$i].UniversalSubcontractorName
            $existingTermSubcontractorsWithChangedNames[$i].SetCustomProperty("NetSuiteLastModifiedDate",$existingNetSubcontractorsWithChangedNames[$i].NetSuiteLastModifiedDate)
            $existingTermSubcontractorsWithChangedNames[$i].SetCustomProperty("flagForReprocessing",$true)
            try{
                Write-Verbose "`tTrying: [$($existingTermSubcontractorsWithChangedNames[$i].UniversalSubcontractorName)].Name = [$($existingNetSubcontractorsWithChangedNames[$i].UniversalSubcontractorName)]"
                $existingTermSubcontractorsWithChangedNames[$i].Context.ExecuteQuery()
                }
            catch{
                if($_.Exception -match "TermStoreErrorCodeEx:There is already a term with the same default label and parent term."){
                    #A NetSuite subcontractor has been renamed and the new name collides with an existing NetSuite subcontractor. This is a NetSuite problem, and the Subcontractors need to be merged there first.
                    Write-Warning "There is already a Term called [$($existingNetSubcontractorsWithChangedNames[$i].UniversalSubcontractorName)] - cannot rename Term [$($existingTermSubcontractorsWithChangedNames[$i].Name)]"
                    [array]$duffUpdatedSubcontractors += @($thisNewSubcontractor,"TermStoreErrorCodeEx:There is already a term with the same default label and parent term. Subcontractor Term rename [$($existingNetSubcontractorsWithChangedNames[$i].UniversalSubcontractorName)] -> [$($existingTermSubcontractorsWithChangedNames[$i].UniversalSubcontractorName)] failed.")
                    }
                else{
                    Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                    [array]$duffUpdatedSubcontractors += @($thisNewSubcontractor,$(get-errorSummary -errorToSummarise $_))
                    }
                }       
            }
    
        [array]$existingNetSubcontractorsWithOriginalNames = $subcontractorNameComparison["=="]
        $existingSubcontractorsWithOriginalNamesComparison = process-comparison -subsetOfNetObjects $existingNetSubcontractorsWithOriginalNames -allTermObjects $allSubcontractorTerms -idInCommon NetSuiteSubcontractorId -propertyToTest NetSuiteLastModifiedDate -validate 
        [array]$existingNetSubcontractorsWithOriginalNames =  $existingSubcontractorsWithOriginalNamesComparison["<="] #This is the same as above, but ordered by NetSuiteSubcontractorId
        [array]$existingTermSubcontractorsWithOriginalNames = $existingSubcontractorsWithOriginalNamesComparison["=>"]
        Write-Host "`t`tProcessing [$($existingTermSubcontractorsWithOriginalNames.Count)] existing Subcontractors without changed names, but have been updated in another way"
        for($i=0;$i -lt $existingTermSubcontractorsWithOriginalNames.Count; $i++){
            #No: Update NetSuiteLastModifiedDate
            if($i%1000 -eq 0){Write-Host "`t`t`tUpdating Term [$($i+1)]/[$($existingTermSubcontractorsWithOriginalNames.Count)]: [$($existingTermSubcontractorsWithOriginalNames[$i].UniversalSubcontractorName)]"}
            $thisExistingTermSubcontractorWithOriginalName = $existingTermSubcontractorsWithOriginalNames[$i]
            $thisExistingTermSubcontractorWithOriginalName.SetCustomProperty("NetSuiteLastModifiedDate",$existingNetSubcontractorsWithOriginalNames[$i].NetSuiteLastModifiedDate)
            try{
                Write-Verbose "`t`t`tTrying: [$($thisExistingTermSubcontractorWithOriginalName[$i].UniversalSubcontractorName)].NetSuiteLastModifiedDate = [$($existingNetSubcontractorsWithOriginalNames[$i].NetSuiteLastModifiedDate)]"
                if(($i%10 -eq 0) -or ($i -eq $existingTermSubcontractorsWithOriginalNames.Count-1)){$thisExistingTermSubcontractorWithOriginalName.Context.ExecuteQuery()} #ExecuteQuery() every 10th iteration, and on the last run (to improve efficiency)
                }
            catch{
                Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                [array]$duffUpdatedSubcontractors += @($thisExistingTermSubcontractorWithOriginalName,$(get-errorSummary -errorToSummarise $_)) #This won't necessarily catch the problematic Term, but hopefully the error message with give us a good clue
                }
            }

        #endregion
        }
    Write-Host "Subcontractor reconcilliation completed in [$($SubcontractorReconcile.TotalMinutes)] minutes ([$($SubcontractorReconcile.TotalSeconds)] seconds)"
    Write-Host

    }

Write-Host "Processing complete at [$(get-date -Format s)] in [$($timeForFullCycle.TotalMinutes)] minutes ([$($timeForFullCycle.TotalSeconds)] seconds)"

Stop-Transcript

    #>