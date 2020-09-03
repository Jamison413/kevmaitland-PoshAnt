$netSuiteParams = get-netSuiteParameters -connectTo Sandbox
$allNClientsProd = get-netSuiteClientsFromNetSuite -netsuiteParameters $(get-netSuiteParameters -connectTo Production)
#$allNContactsSand = get-netSuiteClientsFromNetSuite -netsuiteParameters $(get-netSuiteParameters -connectTo Sandbox)
#$customers = invoke-netsuiteRestMethod -requestType GET -url "$($netSuiteParams.uri)/customer$query" -netsuiteParameters $netSuiteParams #-Verbose 
#$customers.Count

$sharePointBotDetails = get-graphAppClientCredentials -appName SharePointBot
$tokenResponseSharePointBot = get-graphTokenResponse -aadAppCreds $sharePointBotDetails

$sharePointAdmin = "kimblebot@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\KimbleBot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass

Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com" -Credentials $adminCreds

$pnpTermGroup = "Kimble"
$pnpTermSet = "Clients"

$allClientTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet


$allNClientsProd | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name Name2 -Value $($_.companyName).Replace("&","").Replace("＆","") -Force}
$allClientTerms  | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name Name2 -Value $($_.Name).Replace("&","").Replace("＆","") -Force}
$delta = Compare-Object -ReferenceObject $allNClientsProd -DifferenceObject $allClientTerms -Property Name2 -PassThru -IncludeEqual

$matched = $delta | ? {$_.SideIndicator -eq "=="}
$matched | % {
    $thisMatch = $_
    $allClientTerms | ? {$_.Name -eq $thisMatch.Name} | % {
        $matchedTerm = $_
        $matchedTerm.SetCustomProperty("NetSuiteId",$thisMatch.id)
        $matchedTerm.Context.ExecuteQuery()
        }
    }

$deltaKimbleId = Compare-Object -ReferenceObject $allNClientsProd -DifferenceObject $allClientTerms -Property Name2 -PassThru -IncludeEqual

$missingFromMMD = $delta | ? {$_.SideIndicator -eq "<="}
$missingFromMMD | % {
    $thisMiss = $_
    Write-Host -ForegroundColor Yellow $thisMiss.companyName 
    $alredyTerm = Get-PnPTerm -Identity $thisMiss.companyName -TermSet "Clients" -TermGroup "Kimble" 
    $alredyTerm.SetCustomProperty("NetSuiteId",$thisMiss.id)
    $alredyTerm.Context.ExecuteQuery()
    }

#$netSuiteParamsProd = get-netSuiteParameters -connectTo Production
#$allNClientsProd = get-netSuiteClientsFromNetSuite -netsuiteParameters $netSuiteParamsProd
#get netsuiteparams for Prod (currenlt only valid for Sandbox)