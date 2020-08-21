$netSuiteParams = get-netSuiteParameters -connectTo Sandbox
$allNClientsProd = get-netSuiteClientsFromNetSuite -netsuiteParameters $netSuiteParams
$customers = invoke-netsuiteRestMethod -requestType GET -url "$($netSuiteParams.uri)/customer$query" -netsuiteParameters $netSuiteParams #-Verbose 
$customers.Count

$sharePointBotDetails = get-graphAppClientCredentials -appName SharePointBot
$tokenResponseSharePointBot = get-graphTokenResponse -aadAppCreds $sharePointBotDetails

$sharePointAdmin = "kimblebot@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\KimbleBot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass

Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com" -Credentials $adminCreds
get-help Get-PnPTaxonomyItem -full

$pnpTermGroup = "Kimble"
$pnpTermSet = "Clients"

$allKimbleClientTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet
$allClientTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet


$allKimbleClientTerms[0]


$allNClientsSand[0]


$usefulTerms = $allKimbleClientTerms | ? {![string]::IsNullOrWhiteSpace($_.CustomProperties["KimbleId"])}
$usefulTerms | sort Name | select Name,{$_.CustomProperties["KimbleId"]}


$allNClientsProd | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name Name -Value $_.companyName -Force}
$allNClientsProd | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name Name2 -Value $($_.companyName).Replace("&","").Replace("＆","") -Force}
$netTerms  | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name Name2 -Value $($_.Name).Replace("&","").Replace("＆","") -Force}
$delta = Compare-Object -ReferenceObject $allNClientsSand -DifferenceObject $allKimbleClientTerms -Property Name2 -PassThru -IncludeEqual

$matched3 = $delta | ? {$_.SideIndicator -eq "=="}
$matched | % {
    $thisMatch = $_
    $allKimbleClientTerms | ? {$_.Name -eq $thisMatch.Name} | % {
        $matchedTerm = $_
        $matchedTerm.SetCustomProperty("NetSuiteId",$thisMatch.id)
        $matchedTerm.Context.ExecuteQuery()
        }
    }

$missingFromMMD2 = $delta | ? {$_.SideIndicator -eq "<="}
$missingFromMMD2 | % {
    $thisMiss = $_
    Write-Host -ForegroundColor Yellow $thisMiss.companyName 
    $alredyTerm = Get-PnPTerm -Identity $thisMiss.companyName -TermSet "Clients" -TermGroup "Kimble" 
    $alredyTerm.SetCustomProperty("NetSuiteId",$thisMiss.id)
    $alredyTerm.Context.ExecuteQuery()
    }

    $allKimbleClientTerms | ? {$_.Name -eq $thisMatch.Name} | % {
        $matchedTerm = $_
        $matchedTerm.SetCustomProperty("NetSuiteId",$thisMatch.id)
        $matchedTerm.Context.ExecuteQuery()
        }
    }
#$netSuiteParamsProd = get-netSuiteParameters -connectTo Production
#$allNClientsProd = get-netSuiteClientsFromNetSuite -netsuiteParameters $netSuiteParamsProd
#get netsuiteparams for Prod (currenlt only valid for Sandbox)