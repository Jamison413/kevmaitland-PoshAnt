reconcile
Import-Module _PS_Library_Databases


$sqlDbConn = connect-toSqlServer -SQLServer "sql.sustain.co.uk" -SQLDBName "SUSTAIN_LIVE"

$filesToProcess = @("Scrape_mark.sayers@anthesisgroup.com_filtered.csv")

foreach ($file in $filesToProcess){
    gci $env:USERPROFILE\Desktop\Scrapes\ | ?{$_.Name -match $file}   | % {
        $scrapeFile = $_
        rv allContacts
        $headers = gc -Path $scrapeFile.FullName -First 1
        $notFirstLine = $false
        gc $scrapeFile.FullName | %{
            $thisLine = $_
            if($notFirstLine){
                #Do stuff
                $thisContact = New-Object psobject
                $i = 0
                $headers.Split(",") | % {
                    $thisContact | Add-Member -Name $_ -MemberType NoteProperty -Value $thisLine.Split(",")[$i]
                    $i++
                    }
                $thisContact | Add-Member -Name "KimbleId" -MemberType NoteProperty -Value $null
                $thisContact | Add-Member -Name "KimbleName" -MemberType NoteProperty -Value $null
                $thisContact | Add-Member -Name "AccountId" -MemberType NoteProperty -Value $null
                $thisContact | Add-Member -Name "Results" -MemberType NoteProperty -Value $null
                $thisContact | Add-Member -Name "ActionToTake" -MemberType NoteProperty -Value "Do not import"
                $thisContact | Add-Member -Name "ContactNameToImport" -MemberType NoteProperty -Value $null
                $thisContact | Add-Member -Name "ContactEmailToImport" -MemberType NoteProperty -Value $null
                $thisContact | Add-Member -Name "CompanyNameToImport" -MemberType NoteProperty -Value $null

                if($thisContact.mailbox -eq $thisContact.from){$theirEmailAddress = $thisContact.to}
                else{$theirEmailAddress = $thisContact.from}
            
                $sql = "SELECT Id, Email, Name, AccountId FROM SUS_Kimble_Contacts WHERE Email = '$(sanitise-forSql $theirEmailAddress)'"
                if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`t`$query = $sql"}
                $results = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $sqlDbConn
                if($results){
                    $thisContact.KimbleId = $results.Id
                    $thisContact.AccountId = $results.AccountId
                    $thisContact.KimbleName = $results.Name
                    $thisContact.Results = "Contact already exists in Kimble"
                    $thisContact.ContactNameToImport = "N/A"
                    $thisContact.ContactEmailToImport = "$theirEmailAddress"
                    $thisContact.CompanyNameToImport = "N/A"
                    }
                else{
                    #Try to find out more about the contact
                    $sql = "SELECT AccountId, Domain, CompanyName FROM SUS_VW_Kimble_AccountDomains WHERE Domain = '$(sanitise-forSql $thisContact.theirDomain)'"
                    $results = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $sqlDbConn
                    if($results){
                        if($results.count -gt 1){
                            #We've matched the domain to multiple companies, so we can't proceed
                            $thisContact.Results = "Contact's e-mail domain matches more than one company ($($results.CompanyName -join ", "))"
                            $thisContact.ContactNameToImport = $thisContact.guessedName
                            if($thisContact.to -match $thisContact.theirDomain){$thisContact.ContactEmailToImport = $thisContact.to}else{$thisContact.ContactEmailToImport = $thisContact.from}
                            $thisContact.CompanyNameToImport = $($results.CompanyName -join " OR ")+"[Delete as appropriate]"
                            }
                        else{
                            $thisContact.AccountId = $results.AccountId
                            $thisContact.Results = "Based on the Contact's e-mail domain, they probably work at $($results.CompanyName)"
                            $thisContact.ContactNameToImport = $thisContact.guessedName
                            if($thisContact.to -match $thisContact.theirDomain){$thisContact.ContactEmailToImport = $thisContact.to}else{$thisContact.ContactEmailToImport = $thisContact.from}
                            $thisContact.CompanyNameToImport = $results.CompanyName
                            }
                        }
                    else{ #If it doesn't match a known domain, we can't do much
                        $thisContact.Results = "Contact's e-mail domain doesn't match any companies in Kimble. Please provide a new Company Name if importing."
                        $thisContact.ContactNameToImport = $thisContact.guessedName
                        if($thisContact.to -match $thisContact.theirDomain){$thisContact.ContactEmailToImport = $thisContact.to}else{$thisContact.ContactEmailToImport = $thisContact.from}
                        }
                    }
                [array]$allContacts += $thisContact
                }
            else{$notFirstLine = $true}
            }
        $allContacts | %{
            $_ | Add-Member -MemberType NoteProperty -Name lastInboundDate -Value $_.inboundDate
            $_ | Add-Member -MemberType NoteProperty -Name lastOutboundDate -Value $_.outboundDate
            if($thisContact.directionOutbound){$_ | Add-Member -MemberType NoteProperty -Name lastMessageWas -Value "Outbound"}
            else{$_ | Add-Member -MemberType NoteProperty -Name lastMessageWas -Value "Inbound"}
            $_ | Add-Member -MemberType NoteProperty -Name totalMessageCount -Value $($_.inboundMessageCount + $_.outboundMessageCount)
            }
        $allContacts | Select-Object Results,ActionToTake,ContactNameToImport,ContactEmailToImport,CompanyNameToImport,lastInboundDate,lastOutboundDate,lastMessageWas,inboundMessageCount,outboundMessageCount,totalMessageCount,lastSubject,KimbleId,AccountId |  Export-Csv -Path "$env:USERPROFILE\Desktop\Scrapes\Scrape_$($thisContact.mailbox)_ForValidation.csv" -NoTypeInformation
        }
    }
$sqlDbConn.close()
