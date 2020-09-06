$allClientTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes CustomProperties | ? {$_.IsDeprecated -eq $false}
$allClientTerms | % {
    Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteId -Value $_.CustomProperties.NetSuiteId -Force
    Add-Member -InputObject $_ -MemberType NoteProperty -Name Name2 -Value $($_.Name).Replace("&","").Replace("＆","").Replace("  "," ") -Force
    }

#$allClientTerms | ? {![string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteId)} | select Name,{$_.CustomProperties.NetSuiteId} 
$duplicateIds = $allClientTerms | ? {![string]::IsNullOrWhiteSpace($_.CustomProperties.NetSuiteId)} | Group-Object -Property {$_.CustomProperties.NetSuiteId} | Where-Object Count -GT 1
$duplicateIds.Name | % {
    $thisId = $_
    $netClient = $matchedId | ? {$_.NetSuiteId -eq $thisId}
    $matchingTerms = $allClientTerms | ? {$_.NetSuiteId -eq $thisId}
    $primaryTerm = $matchingTerms | sort {$_.CustomProperties.count} | select -Last 1
    $matchingTerms | ? {$_.Id -ne $primaryTerm.Id} | % {
        if($($primaryTerm.Name).Replace("&","").Replace("＆","").Replace("  "," ") -eq $($netClient.companyName).Replace("&","").Replace("＆","").Replace("  "," ")){
            merge-pnpTerms -termToBeRetained $primaryTerm -termToBeMerged $_ -setDefaultLabelTo Retained -pnpTermGroup $pnpTermGroup -pnpTermSet $pnpTermSet
            }
        elseif($($_.Name).Replace("&","").Replace("＆","").Replace("  "," ") -eq $($netClient.companyName).Replace("&","").Replace("＆","").Replace("  "," ")){
            merge-pnpTerms -termToBeRetained $primaryTerm -termToBeMerged $_ -setDefaultLabelTo Merged -pnpTermGroup $pnpTermGroup -pnpTermSet $pnpTermSet
            }
        else{
            Write-Warning "Neither [$($primaryTerm.Name)] nor [$($_.Name)] seem to match [$($netClient.companyName)] - cannot merge"
            }
        
        }
    }