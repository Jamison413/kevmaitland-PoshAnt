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

        #region getProjData
        $termProjRetrieval = Measure-Command {
            $pnpTermGroup = "Kimble"
            $pnpTermSet = "Projects"
            $allProjTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes TermSet,TermSet.Group,TermStore,CustomProperties | ? {$_.IsDeprecated -eq $false}
            $allProjTerms | % {
                Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteProjectId -Value $($_.CustomProperties.NetSuiteProjId) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteClientId -Value $($_.CustomProperties.NetSuiteClientId) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name LastModifiedDate -Value $($_.CustomProperties.NetSuiteProjLastModifiedDate) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name TermProjName -Value $($_.name) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name TermProjCode -Value $(($_.name -split " ")[0]) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name TermProjId -Value $($_.Id) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name DriveItemId -Value $($_.CustomProperties.DriveItemId) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalProjCode -Value $(($_.name -split " ")[0]) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalProjName -Value $_.name -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalProjNameSanitised -Value $(sanitise-forNetsuiteIntegration $_.name) -Force
                }
            }
        Write-Host "[$($allProjTerms.Count)] Projects retrieved from TermStore in [$($termProjRetrieval.TotalSeconds)] seconds"

        $netProjRetrieval = Measure-Command {
            if($deltaSync -eq $true){
                [datetime]$lastProcessed = $($allProjTerms | sort {$_.CustomProperties.NetSuiteProjLastModifiedDate} | select -Last 1).CustomProperties.NetSuiteProjLastModifiedDate
                $netQueryProjects = "?q=lastModifiedDate ON_OR_AFTER `"$($(Get-Date $lastProcessed -Format g))`"" #Excludes any Opps that haven't been updated since X
                }
            #$netSuiteProjsToCheck = get-netSuiteProjectFromNetSuite -netsuiteParameters $netSuiteProductionParamaters -query $netQueryProjects 
            #Write-Host "`$netSuiteProjsToCheck.GetType() = [$($netSuiteProjsToCheck.GetType())]"
            #[array]$netSuiteProjsToCheck = get-netSuiteProjectFromNetSuite -netsuiteParameters $netSuiteProductionParamaters -query $netQueryProjects 
            #[System.Object[]]$netSuiteProjsToCheck = get-netSuiteProjectFromNetSuite -netsuiteParameters $netSuiteProductionParamaters -query $netQueryProjects 
            #[psobject[]]$netSuiteProjsToCheck = get-netSuiteProjectFromNetSuite -netsuiteParameters $netSuiteProductionParamaters -query $netQueryProjects 
            $allNetSuiteProjs = get-netSuiteProjectFromNetSuite -netsuiteParameters $netSuiteProductionParamaters -query $netQueryProjects 
            if([string]::IsNullOrEmpty($allNetSuiteProjs.Count)){$allNetSuiteProjs = @()}
            @($allNetSuiteProjs | Select-Object) | % {
                Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteClientId -Value $($_.parent.id) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteProjectId -Value $($_.id) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteProjLastModifiedDate -Value $($_.lastModifiedDate) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteProjCode -Value  $($($_.entityId -split " ")[0]) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteProjName -Value $($_.entityId) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalProjName -Value $($_.entityId) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalProjNameSanitised -Value $(sanitise-forNetsuiteIntegration $_.entityId) -Force
                }
            }
        Write-Host "[$($allNetSuiteProjs.Count)] Projects retrieved from NetSuite in [$($netProjRetrieval.TotalSeconds)] seconds ([$($netProjRetrieval.TotalMinutes)] minutes or [$($netSuiteProjsToCheck.Count/$netProjRetrieval.TotalMinutes)] per minute)"
        if($allNetSuiteProjs.Count -gt 0){export-encryptedCache -netObjects $allNetSuiteProjs -netObjectType Project}

        $script:netSuiteProjsToCheck = @($allNetSuiteProjs | Select-Object) #Remove any $nulls that 401'ed/disappeared in transit
        $netSuiteProjsToCheck = $netSuiteProjsToCheck | ? {![string]::IsNullOrWhiteSpace($_.parent.id)} #Remove the weird Projects with no Client (as we can't create folders for them anyway)
        if($deltaSync -eq $true){
            [array]$processedAtExactlyLastTimestamp = $netSuiteProjsToCheck | ? {$(Get-Date $_.lastModifiedDate) -ge $(Get-Date $lastProcessed)} #Find how many Clients match the $lastProcessed timestamp exactly
            if($processedAtExactlyLastTimestamp.Count -eq 1){$script:netSuiteProjsToCheck = $netSuiteProjsToCheck | ? {$netSuiteProjsToCheck.id -notcontains $processedAtExactlyLastTimestamp[0].id}} #If it's exactly one, exclude it from processing (as we've already processed it on a previous cycle)
            }

    #Weed out the duff records now to prevent errors later on.
        [array]$netSuiteProjsToCheckWithoutDuffGets = $netSuiteProjsToCheck | ? {$_ -notmatch "PSMessageDetails"}
        if($netSuiteProjsToCheckWithoutDuffGets.Count -ne $netSuiteProjsToCheck.Count){
            Write-Host "`t[$($netSuiteProjsToCheck.Count-$netSuiteProjsToCheckWithoutDuffGets.Count)] REST errors occured retrieving data from NetSuite. Discarding these records, leaving [$($netSuiteProjsToCheck.Count)] records to process."
            $netSuiteProjsToCheck = $netSuiteProjsToCheckWithoutDuffGets
            }

        [array]$newProjsWithoutOvertlyDuffNames = $netSuiteProjsToCheck | ? {$(test-validNameForTermStore -stringToTest $_.UniversalProjName) -eq $true}
        if($newProjsWithoutOvertlyDuffNames.Count -ne $netSuiteProjsToCheck.Count){
            Write-Host "`t[$($netSuiteProjsToCheck.Count - $newProjsWithoutOvertlyDuffNames.Count)] of these contain illegal characters for Terms, so I'll just process the remaining [$($newProjsWithoutOvertlyDuffNames.Count)]"
            $netSuiteProjsToCheck = $newProjsWithoutOvertlyDuffNames
            }

        #endregion

        #region getOppData
        $termOppRetrieval = Measure-Command {
            $pnpTermGroup = "Kimble"
            $pnpTermSet = "Opportunities"
            $allOppTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes TermSet,TermSet.Group,TermStore,CustomProperties | ? {$_.IsDeprecated -eq $false}
            $allOppTerms | % {
                Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteOppId -Value $($_.CustomProperties.NetSuiteOppId) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteClientId -Value $($_.CustomProperties.NetSuiteClientId) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name LastModifiedDate -Value $($_.CustomProperties.NetSuiteOppLastModifiedDate) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name TermOppLabel -Value $($_.name) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name TermOppCode -Value $($_.name) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name TermProjId -Value $($_.CustomProperties.NetSuiteProjectId) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name DriveItemId -Value $($_.CustomProperties.DriveItemId) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalOppCode -Value $(($_.name -split " ")[0]) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalOppName -Value $($_.name) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalOppNameSanitised -Value $(sanitise-forNetsuiteIntegration $_.name) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteProjectId -Value $($_.CustomProperties.NetSuiteProjectId) -Force
                }
            }
        Write-Host "[$($allOppTerms.Count)] Opportunities retrieved from TermStore in [$($termOppRetrieval.TotalSeconds)] seconds"

        $netOppRetrieval = Measure-Command {
            if($deltaSync -eq $true){
                [datetime]$lastProcessed = $($allOppTerms | sort {$_.CustomProperties.NetSuiteOppLastModifiedDate} | select -Last 1).CustomProperties.NetSuiteOppLastModifiedDate
                $netQueryOpps = "?q=lastModifiedDate ON_OR_AFTER `"$($(Get-Date $lastProcessed -Format g))`"" #Excludes any Opps that haven;t been updated since X
                }
            #[array]$netSuiteOppsToCheck = get-netSuiteOpportunityFromNetSuite -netsuiteParameters $netSuiteProductionParamaters -query $netQueryOpps
            #[psobject[]]$netSuiteOppsToCheck = get-netSuiteOpportunityFromNetSuite -netsuiteParameters $netSuiteProductionParamaters -query $netQueryOpps
            #$netSuiteOppsToCheck = @()
            $allNetSuiteOpps = get-netSuiteOpportunityFromNetSuite -netsuiteParameters $netSuiteProductionParamaters -query $netQueryOpps
            if([string]::IsNullOrEmpty($allNetSuiteOpps.Count)){$allNetSuiteOpps = @()}
            @($allNetSuiteOpps | Select-Object) | % {
                Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteClientId -Value $($_.entity.id) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteOppId -Value $($_.id) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteProjectId -Value $($_.custbody_project_created.id) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteOppLastModifiedDate -Value $($_.lastModifiedDate) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteOppCode -Value $("$($_.tranId)") -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteOppLabel -Value $("$($_.tranId) $($_.title)") -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalOppName -Value $("$($_.tranId) $($_.title)") -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalOppNameSanitised -Value $(sanitise-forNetsuiteIntegration $("$($_.tranId) $($_.title)")) -Force
                }
            }
        Write-Host "[$($allNetSuiteOpps.Count)] opportunities retrieved from NetSuite in [$($netOppRetrieval.TotalSeconds)] seconds ([$($netOppRetrieval.TotalMinutes)] minutes or [$($netSuiteOppsToCheck.Count / $netOppRetrieval.TotalMinutes)] per minute)"
        if($allNetSuiteOpps.Count -gt 0){export-encryptedCache -netObjects $allNetSuiteOpps -netObjectType Opportunity}

        $netSuiteOppsToCheck = @($allNetSuiteOpps | Select-Object) #Remove any $nulls that 401'ed/disappeared in transit
        if($deltaSync -eq $true){
            [array]$processedAtExactlyLastTimestamp = $netSuiteOppsToCheck | ? {$(Get-Date $_.lastModifiedDate) -ge $(Get-Date $lastProcessed)} #Find how many Clients match the $lastProcessed timestamp exactly
            if($processedAtExactlyLastTimestamp.Count -eq 1){$netSuiteOppsToCheck = $netSuiteOppsToCheck | ? {$netSuiteOppsToCheck.id -notcontains $processedAtExactlyLastTimestamp[0].id}} #If it's exactly one, exclude it from processing (as we've already processed it on a previous cycle)
            }

    #Weed out the duff records now to prevent errors later on.
        [array]$netSuiteOppsToCheckWithoutDuffGets = $netSuiteOppsToCheck | ? {$_ -notmatch "PSMessageDetails"}
        if($netSuiteOppsToCheckWithoutDuffGets.Count -ne $netSuiteOppsToCheck.Count){
            Write-Host "`t[$($netSuiteOppsToCheck.Count-$netSuiteOppsToCheckWithoutDuffGets.Count)] REST errors occured retrieving data from NetSuite. Discarding these records, leaving [$($netSuiteOppsToCheck.Count)] records to process."
            $netSuiteOppsToCheck = $netSuiteOppsToCheckWithoutDuffGets
            }

        [array]$newOppsWithoutOvertlyDuffNames = $netSuiteOppsToCheck | ? {$(test-validNameForTermStore -stringToTest $_.UniversalOppName) -eq $true}
        if($newOppsWithoutOvertlyDuffNames.Count -ne $netSuiteOppsToCheck.Count){
            Write-Host "`t[$($netSuiteOppsToCheck.Count - $newOppsWithoutOvertlyDuffNames.Count)] of these contain illegal characters for Terms, so I'll just process the remaining [$($newOppsWithoutOvertlyDuffNames.Count)]"
            $netSuiteOppsToCheck = $newOppsWithoutOvertlyDuffNames
            }



        #endregion

        #region getClientData
        $termClientRetrieval = Measure-Command {
            $pnpTermGroup = "Kimble"
            $pnpTermSet = "Clients"
            $allClientTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes TermSet,TermSet.Group,TermStore,CustomProperties | ? {$_.IsDeprecated -eq $false}
            @($allClientTerms | Select-Object) | % {
                Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteClientId -Value $($_.CustomProperties.NetSuiteId) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name DriveClientId -Value $($_.CustomProperties.GraphDriveId) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name TermClientId -Value $($_.Id) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name TermClientName -Value $($_.name) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteLastModifiedDate -Value $($_.CustomProperties.NetSuiteLastModifiedDate) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalClientName -Value $($_.Name) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalClientNameSanitised -Value $(sanitise-forNetsuiteIntegration $_.Name) -Force #This helps to avoid weird encoding, diacritic and special character problems when comparing strings
                }
            }
        Write-Host "[$($allClientTerms.Count)] clients retrieved from TermStore in [$($termClientRetrieval.TotalSeconds)] seconds"

        $netClientRetrieval = Measure-Command {
            #$netQuery =  "?q=companyName CONTAIN_NOT `"Anthesis`"" #Excludes any Companies with "Anthesis" in the companyName
            #$netQuery += " AND companyName CONTAIN_NOT `"intercompany project`"" #Excludes any Companies with "(intercompany project)" in the companyName
            #$netQuery += " AND companyName START_WITH_NOT `"x `"" #Excludes any Companies that begin with "x " in the companyName
            #$netQuery += " AND isPerson IS $false" #Exclude Individuals (until we figure out how to deal with them)
            $netQueryClients = "?q=isPerson IS $false" #Exclude Individuals (until we figure out how to deal with them)
            $netQueryClients += " AND entityStatus ANY_OF_NOT [6, 7]" #Excludes LEAD-Unqualified and LEAD-Qualified (https://XXX.app.netsuite.com/app/crm/sales/customerstatuslist.nl?whence=)
            if($deltaSync -eq $true){
                [datetime]$lastProcessed = $($allClientTerms | sort {$_.CustomProperties.NetSuiteLastModifiedDate} | select -Last 1).CustomProperties.NetSuiteLastModifiedDate
                $netQueryClients += " AND lastModifiedDate ON_OR_AFTER `"$($(Get-Date $lastProcessed -Format g))`"" #Excludes any Companies that haven;t been updated since X
                }
            #[array]$netSuiteClientsToCheck = get-netSuiteClientsFromNetSuite -query $netQueryClients -netsuiteParameters $netSuiteProductionParamaters
            #[psobject[]]$netSuiteClientsToCheck = get-netSuiteClientsFromNetSuite -query $netQueryClients -netsuiteParameters $netSuiteProductionParamaters
            #$netSuiteClientsToCheck = @()
            $allNetSuiteClients = get-netSuiteClientsFromNetSuite -query $netQueryClients -netsuiteParameters $netSuiteProductionParamaters
            if([string]::IsNullOrEmpty($allNetSuiteClients.Count)){$allNetSuiteClients = @()}
            @($allNetSuiteClients | Select-Object) | % {
                Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteClientId -Value $($_.Id) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteClientName -Value $($_.companyName) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteLastModifiedDate -Value $($_.lastModifiedDate) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalClientName -Value $($_.companyName) -Force
                Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalClientNameSanitised -Value $(sanitise-forNetsuiteIntegration $_.companyName) -Force
                }
            }
        Write-Host "[$($allNetSuiteClients.Count)] clients retrieved from NetSuite in [$($netClientRetrieval.TotalSeconds)] seconds or [$($netSuiteClientsToCheck.Count / $netClientRetrieval.TotalMinutes)] per minute"
        if($allNetSuiteClients.Count -gt 0){export-encryptedCache -netObjects $allNetSuiteClients -netObjectType Client}

        $netSuiteClientsToCheck = @($allNetSuiteClients | Select-Object) #Remove any $nulls that 401'ed/disappeared in transit
        if($deltaSync -eq $true){
            [array]$processedAtExactlyLastTimestamp = $netSuiteClientsToCheck | ? {$(Get-Date $_.lastModifiedDate) -ge $(Get-Date $lastProcessed)} #Find how many Clients match the $lastProcessed timestamp exactly
            if($processedAtExactlyLastTimestamp.Count -eq 1){$netSuiteClientsToCheck = $netSuiteClientsToCheck | ? {$netSuiteClientsToCheck.id -notcontains $processedAtExactlyLastTimestamp[0].id}} #If it's exactly one, exclude it from processing (as we've already processed it on a previous cycle)
            }

    #Weed out the duff records now to prevent errors later on.
        [array]$netSuiteClientsToCheckWithoutDuffGets = $netSuiteClientsToCheck | ? {$_ -notmatch "PSMessageDetails"}
        if($netSuiteClientsToCheckWithoutDuffGets.Count -ne $netSuiteClientsToCheck.Count){
            Write-Host "`t[$($netSuiteClientsToCheck.Count-$netSuiteClientsToCheckWithoutDuffGets.Count)] REST errors occured retrieving data from NetSuite. Discarding these records, leaving [$($netSuiteClientsToCheck.Count)] records to process."
            $netSuiteClientsToCheck = $netSuiteClientsToCheckWithoutDuffGets
            }

        [array]$newClientsWithoutOvertlyDuffNames = $netSuiteClientsToCheck | ? {$(test-validNameForTermStore -stringToTest $_.UniversalClientName) -eq $true}
        if($newClientsWithoutOvertlyDuffNames.Count -ne $netSuiteClientsToCheck.Count){
            Write-Host "`t[$($netSuiteClientsToCheck.Count - $newClientsWithoutOvertlyDuffNames.Count)] of these contain illegal characters for Terms, so I'll just process the remaining [$($newClientsWithoutOvertlyDuffNames.Count)]"
            $netSuiteClientsToCheck = $newClientsWithoutOvertlyDuffNames
            }


        #endregion

    #endregion

    $ClientReconcile = Measure-Command {

        $clientComparison = Compare-Object -ReferenceObject @($netSuiteClientsToCheck | Select-Object) -DifferenceObject @($allClientTerms | Select-Object) -Property NetSuiteClientId -IncludeEqual -PassThru #Wrapped in @($ | select) to remove $nulls
        if($deltaSync -eq $false){[array]$orphanedClients = $clientComparison | ? {$_.SideIndicator -eq "=>"}}
        [array]$newClients = $clientComparison | ? {$_.SideIndicator -eq "<="}
        [array]$existingClients = $clientComparison | ? {$_.SideIndicator -eq "=="}

        #region orphanedClients
            #Copy Term to OrphanedTerms
            #Delete original Term
        if($deltaSync -eq $false){
            Write-Host "`tProcessing [$($orphanedClients.Count)] orphaned Clients"
            @($orphanedClients | select-object) | % {
                $thisOrphanedTerm = $_
                $processedOrphanedTerm = process-orphanedTerm -orphanedTerm $thisOrphanedTerm
                if($processedOrphanedTerm -ne $true){
                    [array]$duffOrphanedClients += $thisOrphanedTerm
                    }
                }
            }
        #endregion

        #region newClients
            #Create new Term
        Write-Host "`tProcessing [$($newClients.Count)] new Clients"
        if($deltaSync -eq $false){
            #Fisrt exclude any duplicates companies in NetSuite 
            [array]$duplicateNetSuiteClients = $($netSuiteClientsToCheck | Group-Object -Property {$_.companyName} | ? {$_.Count -ne 1}).Group
            $newClientsDeduped = $newClients | ? {$duplicateNetSuiteClients.id -notcontains $_.id}
            Write-Host "`t`tExcluding [$($newClients.Count - $newClientsDeduped.count)] of these as they are duplicates of existing Clients in NetSuite"
            #Send e-mail report too!!
            #Send e-mail report too!!
            #Send e-mail report too!!
            $newClients = $newClientsDeduped
            }

        @($newClients | select-object) | % {
            $pnpTermGroup = "Kimble"
            $pnpTermSet = "Clients"
            $thisNewClient = $_
            Write-Host "`t`tProcessing new Client [$($thisNewClient.companyName)]"
            try{
                Write-Host "`t`t`tCreating new Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisNewClient.companyName)][@{NetSuiteId=$($thisNewClient.id);NetSuiteLastModifiedDate=$($thisNewClient.lastModifiedDate);flagForReprocessing=$true]"
                $newTerm = New-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Name $thisNewClient.companyName -Lcid 1033 -CustomProperties @{NetSuiteId=$thisNewClient.id;NetSuiteLastModifiedDate=$thisNewClient.lastModifiedDate;flagForReprocessing=$true} -ErrorAction Stop
                }
            catch{ #We don't handle any specific errors here. If there's already a term with this Client's name then the new NetSuite client is a duplicate (and we'll probably want to keep the older record)
                Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                [array]$duffnewClients += @($thisNewClient,$(get-errorSummary -errorToSummarise $_))
                }
            }
        #endregion

        #region existingClients
            #Update Term
                #Has Name changed?
                    #Yes: Update TermClientName, NetSuiteLastModifiedDate & flagForReproccessing
                    #No: Update NetSuiteLastModifiedDate
    
                #Has Name changed?
        $clientNameComparison = process-comparison -subsetOfNetObjects $existingClients -allTermObjects $allClientTerms -idInCommon NetSuiteClientId -propertyToTest UniversalClientNameSanitised -validate
        [array]$existingNetClientsWithChangedNames  = $clientNameComparison["<="]
        [array]$existingTermClientsWithChangedNames  = $clientNameComparison["=>"]
        #Write-Host "existingClients.Count = `t`t`t`t`t`t`t[$($existingClients.Count)]";Write-Host "clientNameComparison.Count = `t[$($clientNameComparison.Count)] (<=[$(($existingNetClientsWithChangedNames).Count)]  ==[$(($clientNameComparison["=="]).Count)]  =>[$(($existingTermClientsWithChangedNames).Count)])"

        Write-Host "`tProcessing [$($existingClients.Count)] existing Clients"
        Write-Host "`t`tProcessing [$($existingTermClientsWithChangedNames.Count)] existing Clients with changed names"
        for($i=0;$i -lt $existingTermClientsWithChangedNames.Count; $i++){
                    #Yes: Update TermClientName, NetSuiteLastModifiedDate & flagForReproccessing
            Write-Host "`t`t`tRenaming Term `t[$($existingTermClientsWithChangedNames[$i].Name)][$($existingTermClientsWithChangedNames[$i].Id)][$($existingTermClientsWithChangedNames[$i].NetSuiteClientId)]"
            Write-Host "`t`t`tto:`t`t`t`t[$($existingNetClientsWithChangedNames[$i].UniversalClientName)][$($existingNetClientsWithChangedNames[$i].NetSuiteClientId)]"
            $existingTermClientsWithChangedNames[$i].Name = $existingNetClientsWithChangedNames[$i].UniversalClientName
            $existingTermClientsWithChangedNames[$i].SetCustomProperty("NetSuiteLastModifiedDate",$existingNetClientsWithChangedNames[$i].NetSuiteLastModifiedDate)
            $existingTermClientsWithChangedNames[$i].SetCustomProperty("flagForReprocessing",$true)
            try{
                Write-Verbose "`tTrying: [$($existingTermClientsWithChangedNames[$i].UniversalClientName)].Name = [$($existingNetClientsWithChangedNames[$i].UniversalClientName)]"
                $existingTermClientsWithChangedNames[$i].Context.ExecuteQuery()
                }
            catch{
                if($_.Exception -match "TermStoreErrorCodeEx:There is already a term with the same default label and parent term."){
                    #A NetSuite client has been renamed and the new name collides with an existing NetSuite client. This is a NetSuite problem, and the Clients need to be merged there first.
                    Write-Warning "There is already a Term called [$($existingNetClientsWithChangedNames[$i].UniversalClientName)] - cannot rename Term [$($existingTermClientsWithChangedNames[$i].Name)]"
                    [array]$duffUpdatedClients += @($thisNewClient,"TermStoreErrorCodeEx:There is already a term with the same default label and parent term. Client Term rename [$($existingNetClientsWithChangedNames[$i].UniversalClientName)] -> [$($existingTermClientsWithChangedNames[$i].UniversalClientName)] failed.")
                    }
                else{
                    Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                    [array]$duffUpdatedClients += @($thisNewClient,$(get-errorSummary -errorToSummarise $_))
                    }
                }       
            }
    
        [array]$existingNetClientsWithOriginalNames = $clientNameComparison["=="]
        $existingClientsWithOriginalNamesComparison = process-comparison -subsetOfNetObjects $existingNetClientsWithOriginalNames -allTermObjects $allClientTerms -idInCommon NetSuiteClientId -propertyToTest NetSuiteLastModifiedDate -validate 
        [array]$existingNetClientsWithOriginalNames =  $existingClientsWithOriginalNamesComparison["<="] #This is the same as above, but ordered by NetSuiteClientId
        [array]$existingTermClientsWithOriginalNames = $existingClientsWithOriginalNamesComparison["=>"]
        Write-Host "`t`tProcessing [$($existingTermClientsWithOriginalNames.Count)] existing Clients without changed names, but have been updated in another way"
        for($i=0;$i -lt $existingTermClientsWithOriginalNames.Count; $i++){
            #No: Update NetSuiteLastModifiedDate
            if($i%1000 -eq 0){Write-Host "`t`t`tUpdating Term [$($i+1)]/[$($existingTermClientsWithOriginalNames.Count)]: [$($existingTermClientsWithOriginalNames[$i].UniversalClientName)]"}
            $thisExistingTermClientWithOriginalName = $existingTermClientsWithOriginalNames[$i]
            $thisExistingTermClientWithOriginalName.SetCustomProperty("NetSuiteLastModifiedDate",$existingNetClientsWithOriginalNames[$i].NetSuiteLastModifiedDate)
            try{
                Write-Verbose "`t`t`tTrying: [$($thisExistingTermClientWithOriginalName[$i].UniversalClientName)].NetSuiteLastModifiedDate = [$($existingNetClientsWithOriginalNames[$i].NetSuiteLastModifiedDate)]"
                if(($i%10 -eq 0) -or ($i -eq $existingTermClientsWithOriginalNames.Count-1)){$thisExistingTermClientWithOriginalName.Context.ExecuteQuery()} #ExecuteQuery() every 10th iteration, and on the last run (to improve efficiency)
                }
            catch{
                Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                [array]$duffUpdatedClients += @($thisExistingTermClientWithOriginalName,$(get-errorSummary -errorToSummarise $_)) #This won't necessarily catch the problematic Term, but hopefully the error message with give us a good clue
                }
            }

        #endregion
        }
    Write-Host "Client reconcilliation completed in [$($ClientReconcile.TotalMinutes)] minutes ([$($ClientReconcile.TotalSeconds)] seconds)"
    Write-Host


    $oppsReconcile = Measure-Command {

        $oppComparison = Compare-Object -ReferenceObject @($netSuiteOppsToCheck | Select-Object) -DifferenceObject @($allOppTerms | Select-Object) -Property NetSuiteOppId -IncludeEqual -PassThru
        [array]$newOpps = $oppComparison | ? {$_.SideIndicator -eq "<="}
        [array]$existingOpps = $oppComparison | ? {$_.SideIndicator -eq "=="}
        if($deltaSync -eq $false){[array]$orphanedOpps = $oppComparison | ? {$_.SideIndicator -eq "=>"}}

        #region orphanedOpps
        if($deltaSync -eq $false){
            Write-Host "`tProcessing [$($orphanedOpps.Count)] orphaned Opportunities"
            @($orphanedOpps | select-object) | % {
                $thisOrphanedTerm = $_
                $processedOrphanedTerm = process-orphanedTerm -orphanedTerm $thisOrphanedTerm
                if($processedOrphanedTerm -ne $true){
                    [array]$duffOrphanedOpps += @($thisOrphanedTerm,$processedOrphanedTerm)
                    }
                }
            }
        #endregion

        #region newOpps
        Write-Host "`tProcessing [$($newOpps.Count)] new Opportunities"
        @($newOpps | select-object) | % {
            $pnpTermGroup = "Kimble"
            $pnpTermSet = "Opportunities"
            #Create new Term
            $thisNewOpp = $_
            Write-Host "`t`tProcessing new Opp [$($thisNewOpp.UniversalOppName)][$($thisNewOpp.entity.refName)][$($thisNewOpp.NetSuiteClientId)]"
            try{
                Write-Host "`t`t`tCreating new Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisNewOpp.UniversalOppName)][@{NetSuiteOppId=$($thisNewOpp.id);NetSuiteOppLastModifiedDate=$($thisNewOpp.lastModifiedDate);NetSuiteClientId=$($thisNewOpp.entity.id);NetSuiteProjectId=$($thisNewOpp.custbody_project_created.id);flagForReprocessing=$true]"
                $newTerm = New-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Name $thisNewOpp.UniversalOppName -Lcid 1033 -CustomProperties @{NetSuiteOppId=$thisNewOpp.id;NetSuiteOppLastModifiedDate=$thisNewOpp.lastModifiedDate;NetSuiteClientId=$thisNewOpp.entity.id;NetSuiteProjectId=$thisNewOpp.custbody_project_created.id;flagForReprocessing=$true}
                }
            catch{ #We don't handle any specific errors here - the OppLabel should be unique in NetSuite
                Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                [array]$duffNewOpps += @($thisNewOpp,$(get-errorSummary -errorToSummarise $_))
                }
            }
        #endregion

        #region existingOpps
            #Update Term
                #Does this Opp have a TermProjId?
                    #Yes: 
                        #Has Project changed?
                            #Yes: Update TermProjId, NetSuiteOppLastModifiedDate & flagForReproccessing
                            #No: Update NetSuiteOppLastModifiedDate
                    #No:
                        #Has Name changed?
                            #Yes: Update TermOppName, NetSuiteOppLastModifiedDate & flagForReproccessing
                            #No: Update NetSuiteOppLastModifiedDate
                        #Has Client changed?
                            #Yes: Update NetSuiteClientId, NetSuiteOppLastModifiedDate & flagForReproccessing
                            #No: Update NetSuiteOppLastModifiedDate

        Write-Host "`tProcessing [$($existingOpps.Count)] existing Opportunities"
            #Update Term
                #Does this Opp have a TermProjId?
        [array]$existingNetOppsWithProjId = $existingOpps    | ? {![string]::IsNullOrWhiteSpace($_.NetSuiteProjectId)}
                    #Yes: 
                        #Has Project changed?
        $existingOppsWithProjIdComparison = process-comparison -subsetOfNetObjects $existingNetOppsWithProjId -allTermObjects $allOppTerms -idInCommon NetSuiteOppId -propertyToTest NetSuiteProjectId -validate
        if($existingOppsWithProjIdComparison -eq $false){[array]$comparisonErrors+="`$existingOppsWithProjIdComparison"}
        [array]$existingNetOppsWithProjIdWithChangedProjects  = $existingOppsWithProjIdComparison["<="]
        [array]$existingTermOppsWithProjIdWithChangedProjects  = $existingOppsWithProjIdComparison["=>"]
                            #Yes: Update TermProjId, NetSuiteOppLastModifiedDate & flagForReproccessing
        Write-Host "`t`tProcessing [$($existingTermOppsWithProjIdWithChangedProjects.Count)] existing Opportunities with changed Projects"
        for($i=0;$i -lt $existingTermOppsWithProjIdWithChangedProjects.Count; $i++){
            Write-Host "`t`t`Updating NetSuiteProjectId `t[$($existingTermOppsWithProjIdWithChangedProjects[$i].TermProjId)] for Term `t`t[$($existingTermOppsWithProjIdWithChangedProjects[$i].UniversalOppName)][$($existingTermOppsWithProjIdWithChangedProjects[$i].Id)][$($existingTermOppsWithProjIdWithChangedProjects[$i].NetSuiteClientId)]"
            Write-Host "`t`t`tto:`t`t`t`t`t`t[$($existingNetOppsWithProjIdWithChangedProjects[$i].NetSuiteProjectId)] from NetOpp `t[$($existingNetOppsWithProjIdWithChangedProjects[$i].UniversalOppName)][$($existingNetOppsWithProjIdWithChangedProjects[$i].NetSuiteOppId)][$($existingNetOppsWithProjIdWithChangedProjects[$i].entity.refName)][$($existingNetOppsWithProjIdWithChangedProjects[$i].entity.id)]"
            $existingTermOppsWithProjIdWithChangedProjects[$i].SetCustomProperty("NetSuiteProjectId",$existingNetOppsWithProjIdWithChangedProjects[$i].NetSuiteProjectId)
            $existingTermOppsWithProjIdWithChangedProjects[$i].SetCustomProperty("NetSuiteOppLastModifiedDate",$existingNetOppsWithProjIdWithChangedProjects[$i].NetSuiteOppLastModifiedDate)
            $existingTermOppsWithProjIdWithChangedProjects[$i].SetCustomProperty("flagForReprocessing",$true)
            try{
                Write-Verbose "`tTrying: [$($existingTermOppsWithProjIdWithChangedProjects[$i].UniversalOppName)].NetSuiteProjectId = [$($existingNetOppsWithProjIdWithChangedProjects[$i].NetSuiteProjectId)]"
                $existingTermOppsWithProjIdWithChangedProjects[$i].Context.ExecuteQuery()
                }
            catch{
                Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                [array]$duffUpdatedOpps += @($existingTermOppsWithProjIdWithChangedProjects[$i],$(get-errorSummary -errorToSummarise $_))
                }
            }

                            #No: Update NetSuiteOppLastModifiedDate
        [array]$existingNetOppsWithProjIdWithOriginalProjects = $existingOppsWithProjIdComparison["=="]
        $existingNetOppsWithProjIdWithOriginalProjectsComparison = process-comparison -subsetOfNetObjects $existingNetOppsWithProjIdWithOriginalProjects -allTermObjects $allOppTerms -idInCommon NetSuiteOppId -propertyToTest LastModifiedDate -validate
        [array]$existingNetOppsWithProjIdWithOriginalProjects  = $existingNetOppsWithProjIdWithOriginalProjectsComparison["<="]
        [array]$existingTermOppsWithProjIdWithOriginalProjects = $existingNetOppsWithProjIdWithOriginalProjectsComparison["=>"]
        #Write-Host "existingNetOppsWithProjIdWithOriginalProjectsComparison = [$($existingNetOppsWithProjIdWithOriginalProjectsComparison.Count)] (<=[$(($existingNetOppsWithProjIdWithOriginalProjects).Count)]  ==[$(($existingNetOppsWithProjIdWithOriginalProjectsComparison).Count)]  =>[$(($existingTermOppsWithProjIdWithOriginalProjects).Count)])"
        Write-Host "`t`tProcessing [$($existingTermOppsWithProjIdWithOriginalProjects.Count)] existing Opportunities with original Projects"
        for($i=0;$i -lt $existingTermOppsWithProjIdWithOriginalProjects.Count; $i++){
            if($i%1000 -eq 0){Write-Host "`t`t`tUpdating Term [$($i+1)]/[$($existingTermOppsWithProjIdWithOriginalProjects.Count)]: [$($existingTermOppsWithProjIdWithOriginalProjects[$i].UniversalOppName)]"}
            Write-Verbose "Updating Term [$($existingTermOppsWithProjIdWithOriginalProjects[$i].UniversalOppName)][$($existingTermOppsWithProjIdWithOriginalProjects[$i].id)][$($existingTermOppsWithProjIdWithOriginalProjects[$i].NetSuiteClientId)].NetSuiteOppLastModifiedDate to [$($existingNetOppsWithProjIdWithOriginalProjects[$i].lastModifiedDate)] from Opp [$($existingNetOppsWithProjIdWithOriginalProjects[$i].UniversalOppName)][$($existingNetOppsWithProjIdWithOriginalProjects[$i].id)][$($existingNetOppsWithProjIdWithOriginalProjects[$i].entity.refName)][$($existingNetOppsWithProjIdWithOriginalProjects[$i].entity.id)]"
            $existingTermOppsWithProjIdWithOriginalProjects[$i].SetCustomProperty("NetSuiteOppLastModifiedDate",$existingNetOppsWithProjIdWithOriginalProjects[$i].lastModifiedDate)
            try{
                Write-Verbose "`tTrying: [$($existingTermOppsWithProjIdWithOriginalProjects[$i].UniversalOppName)].NetSuiteOppLastModifiedDate = [$($existingNetOppsWithProjIdWithOriginalProjects[$i].NetSuiteOppLastModifiedDate)]"
                $existingTermOppsWithProjIdWithOriginalProjects[$i].Context.ExecuteQuery()
                }
            catch{
                Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                [array]$duffUpdatedOpps += @($existingTermOppsWithProjIdWithOriginalProjects[$i],$(get-errorSummary -errorToSummarise $_))
                }
            }

                    #No:
        [array]$existingNetOppsWithoutProjId = $existingOpps | ? { [string]::IsNullOrWhiteSpace($_.NetSuiteProjectId)}
                        #Has Name changed?
        $existingOppsWithoutProjIdNameComparison = process-comparison -subsetOfNetObjects $existingNetOppsWithoutProjId -allTermObjects $allOppTerms -idInCommon NetSuiteOppId -propertyToTest UniversalOppNameSanitised -validate
        [array]$existingNetOppsWithoutProjIdWithMismatchedUniversalOppName = $existingOppsWithoutProjIdNameComparison["<="]
        [array]$existingTermOppsWithoutProjIdWithMismatchedUniversalOppName = $existingOppsWithoutProjIdNameComparison["=>"]
        #Write-Host "existingNetOppsWithoutProjId.Count = `t`t`t`t`t`t`t[$($existingNetOppsWithoutProjId.Count)]";Write-Host "correspondingExistingTermOppsWithoutProjId.Count = `t`t`t`t[$($correspondingExistingTermOppsWithoutProjId.Count)] (should be equal)";Write-Host "existingOppsWithoutProjIdNameComparison.Count = `t[$($existingOppsWithoutProjIdNameComparison.Count)] (<=[$(($existingNetOppsWithoutProjIdWithMismatchedUniversalOppName).Count)]  ==[$($existingNetOppsWithoutProjIdWithOriginalUniversalOppName.Count)]  =>[$(($existingTermOppsWithoutProjIdWithMismatchedUniversalOppName).Count)])"
                            #Yes: Update TermOppName, NetSuiteOppLastModifiedDate & flagForReproccessing
        Write-Host "`t`tProcessing [$($existingTermOppsWithoutProjIdWithMismatchedUniversalOppName.Count)] existing Opportunities without Projects, with changed Names"
        for($i=0;$i -lt $existingTermOppsWithoutProjIdWithMismatchedUniversalOppName.Count; $i++){
            Write-Host "`t`t`Renaming Term `t`t`t[$($existingTermOppsWithoutProjIdWithMismatchedUniversalOppName[$i].UniversalOppName)][$($existingTermOppsWithoutProjIdWithMismatchedUniversalOppName[$i].Id)][$($existingTermOppsWithoutProjIdWithMismatchedUniversalOppName[$i].NetSuiteClientId)]"
            Write-Host "`t`t`tto:`t`t`t`t`t[$($existingNetOppsWithoutProjIdWithMismatchedUniversalOppName[$i].UniversalOppName)][$($existingNetOppsWithoutProjIdWithMismatchedUniversalOppName[$i].NetSuiteOppId)][$($existingNetOppsWithoutProjIdWithMismatchedUniversalOppName[$i].entity.refName)][$($existingNetOppsWithoutProjIdWithMismatchedUniversalOppName[$i].entity.id)]"
            $existingTermOppsWithoutProjIdWithMismatchedUniversalOppName[$i].Name = $existingNetOppsWithoutProjIdWithMismatchedUniversalOppName[$i].UniversalOppName
            $existingTermOppsWithoutProjIdWithMismatchedUniversalOppName[$i].SetCustomProperty("NetSuiteOppLastModifiedDate",$existingNetOppsWithoutProjIdWithMismatchedUniversalOppName[$i].NetSuiteOppLastModifiedDate)
            $existingTermOppsWithoutProjIdWithMismatchedUniversalOppName[$i].SetCustomProperty("flagForReprocessing",$true)
            try{
                Write-Verbose "`tTrying: [$($existingTermOppsWithoutProjIdWithMismatchedUniversalOppName[$i].UniversalOppName)].Name = [$($existingNetOppsWithoutProjIdWithMismatchedUniversalOppName[$i].UniversalOppName)]"
                $existingTermOppsWithoutProjIdWithMismatchedUniversalOppName[$i].Context.ExecuteQuery()
                }
            catch{
                Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                [array]$duffUpdatedOpps += @($existingTermOppsWithoutProjIdWithMismatchedUniversalOppName[$i],$(get-errorSummary -errorToSummarise $_))
                }
            }
                            #No: Update NetSuiteOppLastModifiedDate
        [array]$existingNetOppsWithoutProjIdWithOriginalUniversalOppName = $existingOppsWithoutProjIdNameComparison["=="]
        $existingOppsWithoutProjIdWithOriginalUniversalOppNameComparison = process-comparison -subsetOfNetObjects $existingNetOppsWithoutProjIdWithOriginalUniversalOppName -allTermObjects $allOppTerms -idInCommon NetSuiteOppId -propertyToTest LastModifiedDate -validate
        [array]$existingNetOppsWithoutProjIdWithOriginalUniversalOppName  = $existingOppsWithoutProjIdWithOriginalUniversalOppNameComparison["<="]
        [array]$existingTermOppsWithoutProjIdWithOriginalUniversalOppName = $existingOppsWithoutProjIdWithOriginalUniversalOppNameComparison["=>"]
        Write-Host "`t`tProcessing [$($existingTermOppsWithoutProjIdWithOriginalUniversalOppName.Count)] existing Opportunities with original Names"
        for($i=0;$i -lt $existingTermOppsWithoutProjIdWithOriginalUniversalOppName.Count; $i++){
            if($i%1000 -eq 0){Write-Host "`t`t`tUpdating Term [$($i+1)]/[$($existingTermOppsWithoutProjIdWithOriginalUniversalOppName.Count)]: [$($existingTermOppsWithoutProjIdWithOriginalUniversalOppName[$i].UniversalOppName)]"}
            Write-Verbose "Updating Term [$($existingTermOppsWithoutProjIdWithOriginalUniversalOppName[$i].UniversalOppName)][$($existingTermOppsWithoutProjIdWithOriginalUniversalOppName[$i].id)][$($existingTermOppsWithoutProjIdWithOriginalUniversalOppName[$i].NetSuiteClientId)].NetSuiteOppLastModifiedDate to [$($existingNetOppsWithoutProjIdWithOriginalUniversalOppName[$i].lastModifiedDate)] from Opp [$($existingNetOppsWithoutProjIdWithOriginalUniversalOppName[$i].UniversalOppName)][$($existingNetOppsWithoutProjIdWithOriginalUniversalOppName[$i].id)][$($existingNetOppsWithoutProjIdWithOriginalUniversalOppName[$i].entity.refName)][$($existingNetOppsWithoutProjIdWithOriginalUniversalOppName[$i].entity.id)]"
            $existingTermOppsWithoutProjIdWithOriginalUniversalOppName[$i].SetCustomProperty("NetSuiteOppLastModifiedDate",$existingNetOppsWithoutProjIdWithOriginalUniversalOppName[$i].NetSuiteOppLastModifiedDate)
            try{
                Write-Verbose "`tTrying: [$($existingTermOppsWithoutProjIdWithOriginalUniversalOppName[$i].UniversalOppName)].NetSuiteOppLastModifiedDate = [$($existingNetOppsWithoutProjIdWithOriginalUniversalOppName[$i].NetSuiteOppLastModifiedDate)]"
                $existingTermOppsWithoutProjIdWithOriginalUniversalOppName[$i].Context.ExecuteQuery()
                }
            catch{
                Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                [array]$duffUpdatedOpps += @($existingTermOppsWithoutProjIdWithOriginalUniversalOppName[$i],$(get-errorSummary -errorToSummarise $_))
                }
           }

                        #Has Client changed?
        $existingOppsWithoutProjIdClientComparison = process-comparison -subsetOfNetObjects $existingNetOppsWithoutProjId -allTermObjects $allOppTerms -idInCommon NetSuiteOppId -propertyToTest NetSuiteClientId -validate
        [array]$existingNetOppsWithoutProjIdWithMismatchedClient = $existingOppsWithoutProjIdClientComparison["<="]
        [array]$existingTermOppsWithoutProjIdWithMismatchedClient = $existingOppsWithoutProjIdClientComparison["=>"]
                            #Yes: Update NetSuiteClientId, NetSuiteOppLastModifiedDate & flagForReproccessing
        Write-Host "`t`tProcessing [$($existingTermOppsWithoutProjIdWithMismatchedClient.Count)] existing Opportunities without Projects, with changed Clients"
        for($i=0;$i -lt $existingTermOppsWithoutProjIdWithMismatchedClient.Count; $i++){
            Write-Host "`t`t`Updating NetSuiteClientId `t[$($existingTermOppsWithoutProjIdWithMismatchedClient[$i].NetSuiteClientId)] for Term `t[$($existingTermOppsWithoutProjIdWithMismatchedClient[$i].UniversalOppName)][$($existingTermOppsWithoutProjIdWithMismatchedClient[$i].Id)][$($existingTermOppsWithoutProjIdWithMismatchedClient[$i].NetSuiteClientId)]"
            Write-Host "`t`t`tto:`t`t`t`t`t`t[$($existingNetOppsWithoutProjIdWithMismatchedClient[$i].NetSuiteClientId)] from NetOpp `t[$($existingNetOppsWithoutProjIdWithMismatchedClient[$i].UniversalOppName)][$($existingNetOppsWithoutProjIdWithMismatchedClient[$i].NetSuiteOppId)][$($existingNetOppsWithoutProjIdWithMismatchedClient[$i].entity.refName)][$($existingNetOppsWithoutProjIdWithMismatchedClient[$i].entity.id)]"
            $existingTermOppsWithoutProjIdWithMismatchedClient[$i].SetCustomProperty("NetSuiteClientId",$existingNetOppsWithoutProjIdWithMismatchedClient[$i].NetSuiteClientId)
            $existingTermOppsWithoutProjIdWithMismatchedClient[$i].SetCustomProperty("NetSuiteOppLastModifiedDate",$existingNetOppsWithoutProjIdWithMismatchedClient[$i].NetSuiteOppLastModifiedDate)
            $existingTermOppsWithoutProjIdWithMismatchedClient[$i].SetCustomProperty("flagForReprocessing",$true)
            try{
                Write-Verbose "`tTrying: [$($existingTermOppsWithoutProjIdWithMismatchedClient[$i].UniversalOppName)].NetSuiteClientId = [$($existingNetOppsWithoutProjIdWithMismatchedClient[$i].NetSuiteClientId)]"
                $existingTermOppsWithoutProjIdWithMismatchedClient[$i].Context.ExecuteQuery()
                }
            catch{
                Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                [array]$duffUpdatedOpps += @($existingTermOppsWithoutProjIdWithMismatchedClient[$i],$(get-errorSummary -errorToSummarise $_))
                }
            }
                            #No: Update NetSuiteOppLastModifiedDate
        [array]$existingNetOppsWithoutProjIdWithOriginalClient = $existingOppsWithoutProjIdClientComparison["=="]
        $existingOppsWithoutProjIdWithOriginalClientComparison = process-comparison -subsetOfNetObjects $existingNetOppsWithoutProjIdWithOriginalClient -allTermObjects $allOppTerms -idInCommon NetSuiteOppId -propertyToTest LastModifiedDate -validate
        [array]$existingNetOppsWithoutProjIdWithOriginalClient  = $existingOppsWithoutProjIdWithOriginalClientComparison["<="]
        [array]$existingTermOppsWithoutProjIdWithOriginalClient = $existingOppsWithoutProjIdWithOriginalClientComparison["=>"]
        Write-Host "`t`tProcessing [$($existingTermOppsWithoutProjIdWithOriginalClient.Count)] existing Opportunities with original Clients"
        #We've probably already processed some of these $existingTermOppsWithoutProjIdWithOriginalUniversalOppName, so we could exclude some to make the process more efficient
            $dedupedNetOppsWithoutProjIdWithOriginalClient = $existingNetOppsWithoutProjIdWithOriginalClient | ? {$existingNetOppsWithoutProjIdWithOriginalUniversalOppName.id -notcontains $_.id}
            $dedupedComparison =  process-comparison -subsetOfNetObjects $dedupedNetOppsWithoutProjIdWithOriginalClient -allTermObjects $allOppTerms -idInCommon NetSuiteOppId -propertyToTest LastModifiedDate -validate
            $dedupedTermOppsWithoutProjIdWithOriginalClient = $dedupedComparison["=>"]
        Write-Host "`t`tProcessing [$($dedupedTermOppsWithoutProjIdWithOriginalClient.Count)] existing Opportunities with original Clients (after deduplicating, I reduced the number by [$($existingTermOppsWithoutProjIdWithOriginalClient.Count - $dedupedTermOppsWithoutProjIdWithOriginalClient.Count)])"
        for($i=0;$i -lt $dedupedTermOppsWithoutProjIdWithOriginalClient.Count; $i++){
            Write-Verbose "Updating Term [$($dedupedTermOppsWithoutProjIdWithOriginalClient[$i].UniversalOppName)][$($dedupedTermOppsWithoutProjIdWithOriginalClient[$i].id)][$($dedupedTermOppsWithoutProjIdWithOriginalClient[$i].NetSuiteClientId)].NetSuiteOppLastModifiedDate to [$($existingNetOppsWithoutProjIdWithOriginalClient[$i].lastModifiedDate)] from Opp [$($existingNetOppsWithoutProjIdWithOriginalClient[$i].UniversalOppName)][$($existingNetOppsWithoutProjIdWithOriginalClient[$i].id)][$($existingNetOppsWithoutProjIdWithOriginalClient[$i].entity.refName)][$($existingNetOppsWithoutProjIdWithOriginalClient[$i].entity.id)]"
            $dedupedTermOppsWithoutProjIdWithOriginalClient[$i].SetCustomProperty("NetSuiteOppLastModifiedDate",$existingNetOppsWithoutProjIdWithOriginalClient[$i].NetSuiteOppLastModifiedDate)
            try{
                Write-Verbose "`tTrying: [$($dedupedTermOppsWithoutProjIdWithOriginalClient[$i].UniversalOppName)].NetSuiteOppLastModifiedDate = [$($existingNetOppsWithoutProjIdWithOriginalClient[$i].NetSuiteOppLastModifiedDate)]"
                $dedupedTermOppsWithoutProjIdWithOriginalClient[$i].Context.ExecuteQuery()
                }
            catch{
                Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                [array]$duffUpdatedOpps += @($dedupedTermOppsWithoutProjIdWithOriginalClient[$i],$(get-errorSummary -errorToSummarise $_))
                }
           }

        #endregion
        }
    Write-Host "Opportunity reconcilliation completed in [$($oppsReconcile.TotalMinutes)] minutes ([$($oppsReconcile.TotalSeconds)] seconds)"
    Write-Host


    $projReconcile = Measure-Command{

        $projComparison = Compare-Object -ReferenceObject @($netSuiteProjsToCheck | Select-Object) -DifferenceObject @($allProjTerms | Select-Object) -Property NetSuiteProjectId -IncludeEqual -PassThru
        [array]$newProjs = $projComparison | ? {$_.SideIndicator -eq "<="}
        [array]$existingProjs = $projComparison | ? {$_.SideIndicator -eq "=="}
        if($deltaSync -eq $false){[array]$orphanedProjs = $projComparison | ? {$_.SideIndicator -eq "=>"}}

        #region orphanedProjs
        if($deltaSync -eq $false){
            Write-Host "`tProcessing [$($orphanedProjs.Count)] orphaned Projects"
            @($orphanedProjs | select-object) | % {
                $thisOrphanedTerm = $_
                $processedOrphanedTerm = process-orphanedTerm -orphanedTerm $thisOrphanedTerm
                if($processedOrphanedTerm -ne $true){
                    [array]$duffOrphanedProjects += @($thisOrphanedTerm,$processedOrphanedTerm)
                    }
                }
            }
        #endregion

        #region newProjs
        Write-Host "`tProcessing [$($newProjs.Count)] new Projects"
        @($newProjs | select-object) | % {
            $pnpTermGroup = "Kimble"
            $pnpTermSet = "Projects"
            #Create new Term
            $thisNewProj = $_
            Write-Host "`t`tProcessing new Proj [$($thisNewProj.UniversalProjName)][$($thisNewProj.NetSuiteProjectId)][$($thisNewProj.entity.refName)][$($thisNewProj.NetSuiteClientId)]"
            try{
                Write-Host "`t`t`tCreating new Term [$($pnpTermGroup)][$($pnpTermSet)][$($thisNewProj.UniversalProjName)][@{NetSuiteProjId=$($thisNewProj.NetSuiteProjectId);NetSuiteProjLastModifiedDate=$($thisNewProj.NetSuiteProjLastModifiedDate);NetSuiteClientId=$($thisNewProj.NetSuiteClientId);flagForReprocessing=$true]"
                $newTerm = New-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Name $thisNewProj.UniversalProjName -Lcid 1033 -CustomProperties @{NetSuiteProjId=$thisNewProj.NetSuiteProjectId;NetSuiteProjLastModifiedDate=$thisNewProj.NetSuiteProjLastModifiedDate;NetSuiteClientId=$thisNewProj.NetSuiteClientId;flagForReprocessing=$true}
                }
            catch{ #We don't handle any specific errors here - the OppLabel should be unique in NetSuite
                Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                [array]$duffNewProjs += @($thisNewProj,$(get-errorSummary -errorToSummarise $_))
                }
            }
        #endregion

        #region existingProjs
            #Update Term
            #Has Name changed?
                #Yes: Update TermProjName, NetSuiteProjLastModifiedDate & flagForReproccessing
                #No: Update NetSuiteProjLastModifiedDate
            #Has Client changed?
                #Yes: Update NetSuiteClientId, NetSuiteProjLastModifiedDate & flagForReproccessing
                #No: Update NetSuiteProjLastModifiedDate

        Write-Host "`tProcessing [$($existingProjs.Count)] existing Projects"

        $existingProjNameComparison = process-comparison -subsetOfNetObjects $existingProjs -allTermObjects $allProjTerms -idInCommon NetSuiteProjectId -propertyToTest UniversalProjNameSanitised -validate
        [array]$existingNetProjsWithMismatchedUniversalProjName = $existingProjNameComparison["<="]
        [array]$existingTermProjsWithMismatchedUniversalProjName = $existingProjNameComparison["=>"]
        #Write-Host "existingProjs.Count = `t`t`t`t`t`t`t[$($existingProjs.Count)]";Write-Host "correspondingExistingTermOppsWithoutProjId.Count = `t`t`t`t[$($correspondingExistingTermOppsWithoutProjId.Count)] (should be equal)";Write-Host "existingProjNameComparison.Count = `t[$($existingProjNameComparison.Count)] (<=[$(($existingNetProjsWithMismatchedUniversalProjName).Count)]  ==[$($existingNetProjsWithOriginalUniversalProjName.Count)]  =>[$(($existingTermProjsWithMismatchedUniversalProjName).Count)])"
                            #Yes: Update TermOppName, NetSuiteProjLastModifiedDate & flagForReproccessing
        Write-Host "`t`tProcessing [$($existingTermProjsWithMismatchedUniversalProjName.Count)] existing Projects, with changed Names"
        for($i=0;$i -lt $existingTermProjsWithMismatchedUniversalProjName.Count; $i++){
            Write-Host "`t`t`Renaming Term `t[$($existingTermProjsWithMismatchedUniversalProjName[$i].UniversalProjName)][$($existingTermProjsWithMismatchedUniversalProjName[$i].Id)][$($existingTermProjsWithMismatchedUniversalProjName[$i].NetSuiteClientId)]"
            Write-Host "`t`t`tto:`t`t`t[$($existingNetProjsWithMismatchedUniversalProjName[$i].UniversalProjName)][$($existingNetProjsWithMismatchedUniversalProjName[$i].NetSuiteProjectId)][$($existingNetProjsWithMismatchedUniversalProjName[$i].entity.refName)][$($existingNetProjsWithMismatchedUniversalProjName[$i].entity.id)]"
            $existingTermProjsWithMismatchedUniversalProjName[$i].Name = $existingNetProjsWithMismatchedUniversalProjName[$i].UniversalProjName
            $existingTermProjsWithMismatchedUniversalProjName[$i].SetCustomProperty("NetSuiteProjLastModifiedDate",$existingNetProjsWithMismatchedUniversalProjName[$i].NetSuiteProjLastModifiedDate)
            $existingTermProjsWithMismatchedUniversalProjName[$i].SetCustomProperty("flagForReprocessing",$true)
            try{
                Write-Verbose "`tTrying: [$($existingTermProjsWithMismatchedUniversalProjName[$i].UniversalProjName)].Name = [$($existingNetProjsWithMismatchedUniversalProjName[$i].UniversalProjName)]"
                $existingTermProjsWithMismatchedUniversalProjName[$i].Context.ExecuteQuery()
                }
            catch{
                Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                [array]$duffUpdatedProjs += @($existingTermProjsWithMismatchedUniversalProjName[$i],$(get-errorSummary -errorToSummarise $_))
                }
            }
                            #No: Update NetSuiteProjLastModifiedDate
        [array]$existingNetProjsWithOriginalUniversalProjName = $existingProjNameComparison["=="]
        $existingProjsWithOriginalUniversalProjNameComparison = process-comparison -subsetOfNetObjects $existingNetProjsWithOriginalUniversalProjName -allTermObjects $allProjTerms -idInCommon NetSuiteProjectId -propertyToTest LastModifiedDate -validate
        [array]$existingNetProjsWithOriginalUniversalProjName  = $existingProjsWithOriginalUniversalProjNameComparison["<="]
        [array]$existingTermProjsWithOriginalUniversalProjName = $existingProjsWithOriginalUniversalProjNameComparison["=>"]
        Write-Host "`t`tProcessing [$($existingTermProjsWithOriginalUniversalProjName.Count)] existing Projects with original Names"
        for($i=0;$i -lt $existingTermProjsWithOriginalUniversalProjName.Count; $i++){
            if($i%1000 -eq 0){Write-Host "`t`t`tUpdating Term [$($i+1)]/[$($existingTermProjsWithOriginalUniversalProjName.Count)]: [$($existingTermProjsWithOriginalUniversalProjName[$i].UniversalProjName)]"}
            Write-Verbose "Updating Term [$($existingTermProjsWithOriginalUniversalProjName[$i].UniversalProjName)][$($existingTermProjsWithOriginalUniversalProjName[$i].id)][$($existingTermProjsWithOriginalUniversalProjName[$i].NetSuiteClientId)].NetSuiteProjLastModifiedDate to [$($existingNetProjsWithOriginalUniversalProjName[$i].lastModifiedDate)] from Opp [$($existingNetProjsWithOriginalUniversalProjName[$i].UniversalProjName)][$($existingNetProjsWithOriginalUniversalProjName[$i].id)][$($existingNetProjsWithOriginalUniversalProjName[$i].entity.refName)][$($existingNetProjsWithOriginalUniversalProjName[$i].entity.id)]"
            $existingTermProjsWithOriginalUniversalProjName[$i].SetCustomProperty("NetSuiteProjLastModifiedDate",$existingNetProjsWithOriginalUniversalProjName[$i].NetSuiteProjLastModifiedDate)
            try{
                Write-Verbose "`tTrying: [$($existingTermProjsWithOriginalUniversalProjName[$i].UniversalProjName)].NetSuiteProjLastModifiedDate = [$($existingNetProjsWithOriginalUniversalProjName[$i].NetSuiteProjLastModifiedDate)]"
                $existingTermProjsWithOriginalUniversalProjName[$i].Context.ExecuteQuery()
                }
            catch{
                Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                [array]$duffUpdatedProjs += @($existingTermProjsWithOriginalUniversalProjName[$i],$(get-errorSummary -errorToSummarise $_))
                }
           }

                        #Has Client changed?
        $existingProjClientComparison = process-comparison -subsetOfNetObjects $existingProjs -allTermObjects $allProjTerms -idInCommon NetSuiteProjectId -propertyToTest NetSuiteClientId -validate
        [array]$existingNetProjsWithMismatchedClient = $existingProjClientComparison["<="]
        [array]$existingTermProjWithMismatchedClient = $existingProjClientComparison["=>"]
                            #Yes: Update NetSuiteClientId, NetSuiteProjLastModifiedDate & flagForReproccessing
        Write-Host "`t`tProcessing [$($existingTermProjWithMismatchedClient.Count)] existing Projects, with changed Clients"
        for($i=0;$i -lt $existingTermProjWithMismatchedClient.Count; $i++){
            Write-Host "`t`t`Updating NetSuiteClientId `t[$($existingTermProjWithMismatchedClient[$i].NetSuiteClientId)] for Term `t[$($existingTermProjWithMismatchedClient[$i].UniversalProjName)][$($existingTermProjWithMismatchedClient[$i].Id)][$($existingTermProjWithMismatchedClient[$i].NetSuiteClientId)]"
            Write-Host "`t`t`tto:`t`t`t`t`t`t[$($existingNetProjsWithMismatchedClient[$i].NetSuiteProjectId)] from NetOpp `t[$($existingNetProjsWithMismatchedClient[$i].UniversalProjName)][$($existingNetProjsWithMismatchedClient[$i].NetSuiteProjectId)][$($existingNetProjsWithMismatchedClient[$i].entity.refName)][$($existingNetProjsWithMismatchedClient[$i].entity.id)]"
            $existingTermProjWithMismatchedClient[$i].SetCustomProperty("NetSuiteClientId",$existingNetProjsWithMismatchedClient[$i].NetSuiteClientId)
            $existingTermProjWithMismatchedClient[$i].SetCustomProperty("NetSuiteProjLastModifiedDate",$existingNetProjsWithMismatchedClient[$i].NetSuiteProjLastModifiedDate)
            $existingTermProjWithMismatchedClient[$i].SetCustomProperty("flagForReprocessing",$true)
            try{
                Write-Verbose "`tTrying: [$($existingTermProjWithMismatchedClient[$i].UniversalProjName)].Name = [$($existingNetProjsWithMismatchedClient[$i].UniversalProjName)]"
                $existingTermProjWithMismatchedClient[$i].Context.ExecuteQuery()
                }
            catch{
                Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                [array]$duffUpdatedProjs += @($existingTermProjWithMismatchedClient[$i],$(get-errorSummary -errorToSummarise $_))
                }
            }
                            #No: Update NetSuiteProjLastModifiedDate
        [array]$existingNetProjsWithOriginalClient = $existingProjClientComparison["=="]
        $existingProjsWithOriginalClientComparison = process-comparison -subsetOfNetObjects $existingNetProjsWithOriginalClientWithMismatchedTimestamps -allTermObjects $allProjTerms -idInCommon NetSuiteProjectId -propertyToTest LastModifiedDate -validate
        [array]$existingNetProjsWithOriginalClientWithMismatchedTimestamps  = $existingProjsWithOriginalClientComparison["<="]
        [array]$existingTermProjsWithOriginalClientWithMismatchedTimestamps = $existingProjsWithOriginalClientComparison["=>"]
        Write-Host "`t`tProcessing [$($existingTermProjsWithOriginalClientWithMismatchedTimestamps.Count)] existing Projects with original Clients that have been updated in some other way (that we're not worried about)"
        #We've probably already processed some of these $existingTermProjsWithOriginalUniversalProjName, so we could exclude some to make the process more efficient
            $dedupedNetProjsWithOriginalClient = $existingNetProjsWithOriginalClientWithMismatchedTimestamps | ? {$existingNetProjsWithOriginalUniversalProjName.id -notcontains $_.id}
            $dedupedComparison =  process-comparison -subsetOfNetObjects $dedupedNetProjsWithOriginalClient -allTermObjects $allProjTerms -idInCommon NetSuiteProjectId -propertyToTest LastModifiedDate -validate
            $dedupedTermProjsWithOriginalClient = $dedupedComparison["=>"]
        Write-Host "`t`tProcessing [$($dedupedTermProjsWithOriginalClient.Count)] existing Projects with original Clients (after deduplicating, I reduced the number by [$($existingTermProjsWithOriginalClientWithMismatchedTimestamps.Count - $dedupedTermProjsWithOriginalClient.Count)])"
        for($i=0;$i -lt $dedupedTermProjsWithOriginalClient.Count; $i++){
            if($i%1000 -eq 0){Write-Host "`t`t`tUpdating Term [$($i+1)]/[$($dedupedTermProjsWithOriginalClient.Count)]: [$($dedupedTermProjsWithOriginalClient[$i].UniversalProjName)]"}
            Write-Verbose "Updating Term [$($dedupedTermProjsWithOriginalClient[$i].UniversalProjName)][$($dedupedTermProjsWithOriginalClient[$i].id)][$($dedupedTermProjsWithOriginalClient[$i].NetSuiteClientId)].NetSuiteProjLastModifiedDate to [$($existingNetProjsWithOriginalClientWithMismatchedTimestamps[$i].lastModifiedDate)] from Opp [$($existingNetProjsWithOriginalClientWithMismatchedTimestamps[$i].UniversalProjName)][$($existingNetProjsWithOriginalClientWithMismatchedTimestamps[$i].id)][$($existingNetProjsWithOriginalClientWithMismatchedTimestamps[$i].entity.refName)][$($existingNetProjsWithOriginalClientWithMismatchedTimestamps[$i].entity.id)]"
            $dedupedTermProjsWithOriginalClient[$i].SetCustomProperty("NetSuiteProjLastModifiedDate",$existingNetProjsWithOriginalClientWithMismatchedTimestamps[$i].NetSuiteProjLastModifiedDate)
            try{
                Write-Verbose "`tTrying: [$($dedupedTermProjsWithOriginalClient[$i].UniversalProjName)].NetSuiteProjLastModifiedDate = [$($existingNetProjsWithOriginalClientWithMismatchedTimestamps[$i].NetSuiteProjLastModifiedDate)]"
                $dedupedTermProjsWithOriginalClient[$i].Context.ExecuteQuery()
                }
            catch{
                Write-Host -ForegroundColor Red "`t`t`t$(get-errorSummary -errorToSummarise $_)"
                [array]$duffUpdatedProjs += @($dedupedTermProjsWithOriginalClient[$i],$(get-errorSummary -errorToSummarise $_))
                }
           }

        #endregion
    
        }
    Write-Host "Project reconcilliation completed in [$($projReconcile.TotalMinutes)] minutes ([$($projReconcile.TotalSeconds)] seconds)"
    }

Write-Host "Processing complete at [$(get-date -Format s)] in [$($timeForFullCycle.TotalMinutes)] minutes ([$($timeForFullCycle.TotalSeconds)] seconds)"

Stop-Transcript

    #>