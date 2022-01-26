$errorLogFile = Get-Content C:\ScriptLogs\reconcile-netsuiteToTermStore.ps1_fullSync_Transcript_2021-09-23.log

$orphanedTerms = @()

$errorLogFile | ForEach-Object {
    if($_ -match "Deleting orphaned Term "){
        $orphanedTerms += $_
        }
    } 

$orphanedTermsCleaned =@()
$orphanedTerms | ForEach-Object {
    #$cleanTerm = $_ -match '\[[a-zA-Z]{3}\]'
    $orphanedTermsCleaned += [regex]::Matches($_, '(?<=\[).*?(?=\])')[2].value
    }


    $sharePointAdmin = "kimblebot@anthesisgroup.com"
    #convertTo-localisedSecureString "KimbleBotPasswordHere"
    $sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Downloads\KimbleBot.txt) 
    $adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
    Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com" -Credentials $adminCreds

        $termClientRetrieval = Measure-Command {
            $pnpTermGroup = "Kimble"
            $pnpTermSet = "OrphanedClients"
            $allClientTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes TermSet,TermSet.Group,TermStore,CustomProperties | ? {$_.IsDeprecated -eq $false}
            }


$orphanedTermsCleaned | ForEach-Object {
    $orphanedTerm = $_
    try{
        #Copy Term to OrphanedTerms
        
        $termToRecover =  $allClientTerms  | Where-Object {$_.Name -match $([regex]::Escape($orphanedTerm))} | Sort-Object Name -Descending | Select-Object -First 1
        Write-Host "`tRecovering orphaned Term [Kimble][OrphanedClients][$($termToRecover.Name)] to [Kimble][Clients][$orphanedTerm]"
        #$recoveredTerm = New-PnPTerm -TermGroup "Kimble" -TermSet "Clients" -Name $orphanedTerm -Lcid 1033 -CustomProperties $([hashtable]::new($termToRecover.CustomProperties)) -ErrorAction Stop
        if(![string]::IsNullOrWhiteSpace($recoveredTerm.Name)){
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
    #if($recoveredTerm){
        #if($recoveredTerm.Name -match [Regex]::Escape($orphanedTerm.Name)){
            #Delete original Term
            try{
                Write-Host "`t`tDeleting unorphaned Term [$($termToRecover.TermSet.Group.Name)][$($termToRecover.TermSet.Name)][$($termToRecover.Name)][$($termToRecover.id)][$($termToRecover.NetSuiteClientId)]"
                Remove-PnPTaxonomyItem -TermPath "$($termToRecover.TermSet.Group.Name)|$($termToRecover.TermSet.Name)|$($termToRecover.Name)" -Force -Verbose
                return $true
                }
            catch{
                return $(get-errorSummary -errorToSummarise $_)
                }
          #  }
        #}#>
    }

