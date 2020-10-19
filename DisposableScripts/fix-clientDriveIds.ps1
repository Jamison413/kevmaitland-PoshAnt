$sharePointBotDetails = get-graphAppClientCredentials -appName SharePointBot
$tokenResponseSharePointBot = get-graphTokenResponse -aadAppCreds $sharePointBotDetails
$clientSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,faed84bc-70be-4e35-bfbf-cdab31aeeb99"
#$supplierSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,9fb8ecd6-c87d-485d-a488-26fd18c62303"
#$devSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,8ba7475f-dad0-4d16-bdf5-4f8787838809"
$allClientDrives = get-graphDrives -tokenResponse $tokenResponseSharePointBot -siteGraphId $clientSiteId
$allClientDrives | % {
    Add-Member -InputObject $_ -MemberType NoteProperty -Name Name2 -Value $($_.name).Replace("&","").Replace("＆","").Replace("  "," ") -Force
    Add-Member -InputObject $_ -MemberType NoteProperty -Name DriveId -Value $($_.id) -Force
    }


$sharePointAdmin = "kimblebot@anthesisgroup.com"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content "$env:USERPROFILE\Desktop\KimbleBot.txt") 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com" -Credentials $adminCreds

$pnpTermGroup = "Kimble"
$pnpTermSet = "Clients"
$allClientTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes CustomProperties | ? {$_.IsDeprecated -eq $false}
$allClientTerms | % {
    Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteId -Value $_.CustomProperties.NetSuiteId -Force
    Add-Member -InputObject $_ -MemberType NoteProperty -Name Name2 -Value $($_.Name).Replace("&","").Replace("＆","").Replace("  "," ").Trim() -Force
    Add-Member -InputObject $_ -MemberType NoteProperty -Name DriveId -Value $_.CustomProperties.GraphDriveId -Force
    }


$matchName2 = Compare-Object -ReferenceObject @($allClientDrives | Select-Object) -DifferenceObject @($allClientTerms  | Select-Object) -Property Name2 -ExcludeDifferent -IncludeEqual -PassThru
$matchName2Reversed = Compare-Object -ReferenceObject @($allClientTerms  | Select-Object) -DifferenceObject @($matchName2 | Select-Object) -Property Name2 -ExcludeDifferent -IncludeEqual -PassThru
$mismatchId = Compare-Object -ReferenceObject @($matchName2Reversed  | Select-Object) -DifferenceObject @($matchName2 | Select-Object) -Property DriveId -PassThru | ? {$_.SideINdicator -eq "<="}


$mismatchId.Count

$mismatchId | % {
    $thisClientTerm = $_
    $thisClientDrive = $allClientDrives | ? {$_.Name2 -eq $thisClientTerm.Name2}
    $oldClientDrive = $allClientDrives | ? {$_.id -eq $thisClientTerm.DriveId}
    Write-Host "Updating [$($thisClientTerm.Name)]"
    if($thisClientDrive){
        Write-Host "`tChanging DriveId from [$($thisClientTerm.CustomProperties.GraphDriveId)][$($oldClientDrive.name)] to [$($thisClientDrive.id)][$($thisClientDrive.name)]"
        $thisClientTerm.SetCustomProperty("GraphDriveId",$thisClientDrive.id)
        $thisClientTerm.Context.ExecuteQuery()
        }
    else{
        [array]$problems += $thisClientTerm
        }
    }

$bustDriveIds = $allClientTerms | ? {$_.customproperties.GraphDriveId -match " "}; $bustDriveIds.Count
$bustDriveIds | % {
    $thisClientTerm = $_
    $thisClientDrive = $allClientDrives | ? {$_.Name -eq $thisClientTerm.Name}
    if($thisClientDrive.Count -gt 1){
        $thisClientDrive = $thisClientDrive | Sort-Object createdDateTime | Select-Object -First 1
        [array]$toDudupeDrives += $thisClientTerm
        }
    Write-Host "Updating [$($thisClientTerm.Name)]"
    if($thisClientDrive){
        Write-Host "`tChanging DriveId to [$($thisClientDrive.id)][$($thisClientDrive.name)]"
        $thisClientTerm.SetCustomProperty("GraphDriveId",$thisClientDrive.id)
        $thisClientTerm.Context.ExecuteQuery()
        }
    else{
        [array]$problems += $thisClientTerm
        }
    }
