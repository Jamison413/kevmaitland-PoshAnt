[cmdletbinding()]
param(
    [Parameter(Mandatory = $false, Position = 0)]
        [string]$deltaSync = $false #Specifies whether we are doing a full or incremental sync.
    )

if($PSCommandPath){
    $InformationPreference = 2
    $VerbosePreference = 0
    $logFileLocation = "C:\ScriptLogs\"
    if($deltaSync){$suffix = "_deltaSync"}
    else{$suffix = "_fullSync"}
    #$suffix = "_fullSync"
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))$suffix`_Transcript_$(Get-Date -Format "yyyy-MM-dd").log"
    Start-Transcript $transcriptLogName -Append
    }


#region Clients
    #region getData
$termClientRetrieval = Measure-Command {
    $sharePointAdmin = "kimblebot@anthesisgroup.com"
    #convertTo-localisedSecureString "KimbleBotPasswordHere"
    $sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\KimbleBot.txt) 
    $adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
    Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com" -Credentials $adminCreds
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

$driveClientRetrieval = Measure-Command {
    $tokenResponseSharePointBot = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName SharePointBot )
    $clientSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,faed84bc-70be-4e35-bfbf-cdab31aeeb99"
    $allClientDrives = get-graphDrives -tokenResponse $tokenResponseSharePointBot -siteGraphId $clientSiteId
    $allClientDrives | % {
        Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalClientName -Value $_.name -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalClientNameSanitised -Value $(sanitise-forNetsuiteIntegration $_.name) -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteClientId -Value $_.id -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name DriveClientName -Value $_.name -Force
        }
    }
Write-Host "[$($allClientDrives.Count)] Client Drives retrieved from SharePoint in [$($driveClientRetrieval.TotalSeconds)] seconds ([$($driveClientRetrieval.totalMinutes)] minutes)"

    #endregion

    #Does Term have a DriveClientId?
        #Yes: Update the DriveName
        #No: Create a new Drive

#endregion

#region Opportunities
    #regionGetData
$termOppRetrieval = Measure-Command {
    $pnpTermGroup = "Kimble"
    $pnpTermSet = "Opportunities"
    $allOppTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes TermSet,TermSet.Group,TermStore,CustomProperties | ? {$_.IsDeprecated -eq $false}
    $allOppTerms | % {
        Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteOppId -Value $($_.CustomProperties.NetSuiteOppId) -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteClientId -Value $($_.CustomProperties.NetSuiteClientId) -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteOppLastModifiedDate -Value $($_.CustomProperties.NetSuiteOppLastModifiedDate) -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name TermOppLabel -Value $($_.name) -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name TermOppCode -Value $($_.name) -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name TermProjId -Value $($_.CustomProperties.NetSuiteProjectId) -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name DriveItemId -Value $($_.CustomProperties.DriveItemId) -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalOppName -Value $($_.name) -Force
        Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteProjectId -Value $($_.CustomProperties.NetSuiteProjectId) -Force
        }
    }
Write-Host "[$($allOppTerms.Count)] Opportunities retrieved from TermStore in [$($termOppRetrieval.TotalSeconds)] seconds"

$now = $(Get-Date -f FileDateTimeUniversal)
$topLevelFolderRetrieval = Measure-Command {
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
            if($folderObject.DriveItemFirstWord -match "^O-"){$folderObject | add-member -MemberType NoteProperty -Name UniversalOppName -Value $($_.name) -Force}
            elseif($folderObject.DriveItemFirstWord -match "^P-"){$folderObject | add-member -MemberType NoteProperty -Name UniversalProjName -Value $($_.name) -Force}
            $folderObject | Export-Csv -Path "$env:TEMP\NetRec_AllFolders_$now.csv" -Append -NoTypeInformation -Encoding UTF8 -Force #There are going to be a _lot_ of these, but the number is unknown. Rather than += an array (which will get very inefficient at large numbers), append the data to a CSV and import the CSV once the enumeration is complete
            }
        }
    $topLevelFolders = import-csv "$env:TEMP\NetRec_AllFolders_$now.csv"
    }
Write-Host "ClientDrive top-level folders enumerated in [$($topLevelFolderRetrieval.TotalMinutes)] minutes"
    #endregion
    #Does Term have a TermProjId?
        #Yes: Find Opp Folder and rename to match Project, & set flagForReproccessing = $false
        #No:
            #Does the Term have a DriveItemId?
                #No: Create a new DriveItem, & set flagForReproccessing = $false
                #Yes:
                    #Id the DriveItem there?
                    
                    #Has the Name changed?
                        #Yes: Update the DriveItemName, & set flagForReproccessing = $false
                        #No: Set flagForReproccessing = $false
                    #Has the Client changed?
                        #Yes: Update the NetSuiteClientId, & set flagForReproccessing = $false
                        #No: Dedupe & set flagForReproccessing = $false
#endregion

#region Projects
    #region GetData
    $termProjRetrieval = Measure-Command {
        $pnpTermGroup = "Kimble"
        $pnpTermSet = "Projects"
        $allProjTerms = Get-PnPTerm -TermGroup $pnpTermGroup -TermSet $pnpTermSet -Includes TermSet,TermSet.Group,TermStore,CustomProperties | ? {$_.IsDeprecated -eq $false}
        $allProjTerms | % {
            Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteProjectId -Value $($_.CustomProperties.NetSuiteProjId) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteClientId -Value $($_.CustomProperties.NetSuiteClientId) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name NetSuiteProjLastModifiedDate -Value $($_.CustomProperties.NetSuiteProjLastModifiedDate) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name TermProjName -Value $($_.name) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name TermProjCode -Value $(($_.name -split " ")[0]) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name TermProjId -Value $($_.Id) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name DriveItemId -Value $($_.CustomProperties.DriveItemId) -Force
            Add-Member -InputObject $_ -MemberType NoteProperty -Name UniversalProjName -Value $($_.name) -Force
            }
        }
    Write-Host "[$($allProjTerms.Count)] Projects retrieved from TermStore in [$($termProjRetrieval.TotalSeconds)] seconds"
    #endregion

    #Does the Term have a DriveItemId?
        #No: Create a new DriveItem, & set flagForReproccessing = $false
        #Yes:
            #Has the Name changed?
                #Yes: Update the DriveItemName, & set flagForReproccessing = $false
                #No: Set flagForReproccessing = $false
            #Has the Client changed?
                #Yes: Update the NetSuiteClientId, & set flagForReproccessing = $false
                #No: Dedupe & set flagForReproccessing = $false
#endregion