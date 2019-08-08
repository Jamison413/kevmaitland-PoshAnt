<#************************************************************************
Set Logs
************************************************************************#>
$logFileLocation = "C:\ScriptLogs\"
$fullLogPathAndName = $logFileLocation+"Set-Folder Restriction`_FullLog_$(Get-Date -Format "yyMMdd").log"
$errorLogPathAndName = $logFileLocation+"Set-Folder Restriction`_ErrorLog_$(Get-Date -Format "yyMMdd").log"
if($PSCommandPath){
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))`_Transcript_$(Get-Date -Format "yyMMdd").log"
    Start-Transcript $transcriptLogName -Append
    }

    
<#************************************************************************
Import all the modules
************************************************************************#>

Import-Module _PS_Library_GeneralFunctionality
Import-Module _PNP_Library_SPO
Import-Module SharePointPnPPowerShellOnline
Import-Module _REST_Library-Kimble

<#************************************************************************
Connect to all the services/DB's
************************************************************************#>

connect-ToSpo
$sqlDbConn = connect-toSqlServer -SQLServer "sql.sustain.co.uk" -SQLDBName "SUSTAIN_LIVE"

#Set variables for the Clients Site connection & connect via PnPOnline
$webUrl = "https://anthesisllc.sharepoint.com"
$spoSite = "/clients"
Connect-PnPOnline -Url $($webUrl+$spoSite) -Credentials (Get-Credential)


<#************************************************************************
Set the Enumbers in a variable
************************************************************************#>

#Set variable to contain for each type of project - there is a much better way to do this, but for now fastest way to figure this out
#$ConfidentialProjectFiles = @('E004124')
#$ConfidentialProjectFiles += 'E004118'
#$ConfidentialProjectFiles +='E004187'
#$ConfidentialProjectFiles +='E004189'
#$ConfidentialProjectFiles +='E004196'
#$ConfidentialProjectFiles +='E004579'
#$ConfidentialProjectFiles +='E004177'
#$ConfidentialProjectFiles +='E004203'
#$ConfidentialProjectFiles +='E004256'
#$ConfidentialProjectFiles +='E004418'
#$ConfidentialProjectFiles +='E004199'
#$ConfidentialProjectFiles +='E004494'
#$ConfidentialProjectFiles +='E004192'
#$ConfidentialProjectFiles +='E004179'
#$ConfidentialProjectFiles +='E004198'
#Amended permissions manually - $ConfidentialProjectFiles +='E003842'
#$ConfidentialProjectFiles +='E004125'
#$ConfidentialProjectFiles +='E004126'
#$ConfidentialProjectFiles +='E004129'
#$ConfidentialProjectFiles +='E004117'
#$ConfidentialProjectFiles +='E004173'
#$ConfidentialProjectFiles +='E004182'
#$ConfidentialProjectFiles +='E004191'
#$ConfidentialProjectFiles +='E004240'
#$ConfidentialProjectFiles +='E004113'
#amended permissions manually - this is a duplicate with the same enumber twice that needs investigation    -   $ConfidentialProjectFiles +='E004128'
#$ConfidentialProjectFiles +='E004172'
#$ConfidentialProjectFiles +='E004674'
#$ConfidentialProjectFiles +='E004675'
#$ConfidentialProjectFiles +='E004676'
#$ConfidentialProjectFiles +='E004200'
#$ConfidentialProjectFiles +='E004127'
#$ConfidentialProjectFiles +='E004176'
#$ConfidentialProjectFiles +='E004178'
#$ConfidentialProjectFiles +='E004184'
#$ConfidentialProjectFiles +='E004197'
#$ConfidentialProjectFiles +='E004237'
#$ConfidentialProjectFiles +='E004252'
#$ConfidentialProjectFiles +='E004348'
#$ConfidentialProjectFiles +='E003288'
#$ConfidentialProjectFiles +='E004180'
#$ConfidentialProjectFiles +='E004181'
#$ConfidentialProjectFiles +='E004194'
#$ConfidentialProjectFiles +='E004175'
#$ConfidentialProjectFiles +='E004542'
#$ConfidentialProjectFiles +='E004163'
#$ConfidentialProjectFiles +='E004253'
#$ConfidentialProjectFiles +='E004347'
#$ConfidentialProjectFiles +='E004112'
#$ConfidentialProjectFiles +='E004413'
#$ConfidentialProjectFiles +='E004642'
#$ConfidentialProjectFiles +='E004365'
#$ConfidentialProjectFiles +='E004910'
#$ConfidentialProjectFiles +='E004256'
#$ConfidentialProjectFiles +='E004203'
#$ConfidentialProjectFiles +='E004203'
#$ConfidentialProjectFiles +='E005064'
#$ConfidentialProjectFiles +='E005065'
#Does not exist? Folder no longer exists but data in sql? $ConfidentialProjectFiles +='E005303'
#$ConfidentialProjectFiles +='E005346'
#$ConfidentialProjectFiles +='E005074'
#$ConfidentialProjectFiles +='E005187'
#$ConfidentialProjectFiles +='E005148'
#Missing for clients library $ConfidentialProjectFiles +='E005266'
#$ConfidentialProjectFiles +='E004406'
#$ConfidentialProjectFiles +='E004714'
#$ConfidentialProjectFiles +='E004908'
#$ConfidentialProjectFiles +='E005418'
#$ConfidentialProjectFiles +='E004721'
#$ConfidentialProjectFiles +='E005272'
#$ConfidentialProjectFiles +='E005328'
#$ConfidentialProjectFiles +='E005306'
#$ConfidentialProjectFiles +='E005073'
#$ConfidentialProjectFiles +='E004909'
#$ConfidentialProjectFiles +='E004813'
#$ConfidentialProjectFiles +='E005066'
#$ConfidentialProjectFiles +='E004718'




#Testing hacks!

#This will allow you to search an array with a where clause select by criteria

#doesnt work for E004908, E003842,E004203(emily still left, TCS applied) - not a broken permissions thing), E003842, E004199, E004125

<#************************************************************************
Query SQL for data
************************************************************************#>

<##Function to SELECT key fields from both the SUS_Kimble_Engagements and SUS_Kimble_Accounts tables
#My problem is understanding this one is that currently in the sync scripts, the tables are selected and quiried seperately (I think) - can we make one long list of columns form both tables (or even just the necessary columns) and create a join from the get-go? Like the poor attempt below - noticed some of the column names are also the same, how does this work?
#function get-ClientsandEngagements{
   # [cmdletbinding()]
  #  Param (
   #     [parameter(Mandatory = $true)]
    #    [System.Data.Common.DbConnection]$dbConnection

    #    ,[parameter(Mandatory = $false)]
    #    [string]$pWhereStatement
    #    )
   # Write-Verbose "ClientsandEngagements"#>
    
    
    $pWhereStatement = " WHERE KimbleOne__Reference__c IN ('$($ConfidentialProjectFiles -join "','")')"
    $sql = "SELECT Id, KimbleOne__Account__c, KimbleOne__ShortName__c, KimbleOne__Reference__c, ClientName, EngName, FolderGuid, DocumentLibraryGuid FROM SUS_VW_Kimble_Engagements_with_Client $pWhereStatement"

    Write-Verbose "`t`$query = $sql"
    $results = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $sqlDbConn
    $results.count
    
    $results | select KimbleOne__Reference__c, DocumentLibraryGuid,ClientName
    #}

<#************************************************************************
Get the Client Doc Library GUIDs and Project GUIDs from sql server
************************************************************************


#Get the project folder details via sql
foreach ($eNumber in $ConfidentialProjectFiles){
    #Is this WHERE statement enough to grab all columns stated via the function? I think it should grab all the items with an Enumber from the Accounts table and then subsequently show the Account id 
$ConfidentialProjectLocations = get-ClientsandEngagements -dbConnection $sqlDbConn -pWhereStatement "WHERE KimbleOne__Reference__c = @($eNumber)"
}
#>

<#************************************************************************
Apply Restricted Permissions
************************************************************************#>


#Looking for missing DocLibGUIDs?      $results | ? {[string]::IsNullOrWhiteSpace($_.DocumentLibraryGuid)} | select KimbleOne__Reference__c, DocumentLibraryGuid,ClientName

$results2 | % {
    $result = $results2 #$_
    $list = Get-PnPList $result.DocumentLibraryGuid
    $item = get-spoProjectFolder -pnpList $list -kimbleEngagementCodeToLookFor $result.KimbleOne__Reference__c #-folderGuid $result.FolderGuid
    Set-PnPListItemPermission -List $list.Id -Identity $item.Id -User Transaction_and_Corporate_Services_FIN@anthesisgroup.com -AddRole "Contribute" -ClearExisting
    Set-PnPListItemPermission -List $list.Id -Identity $item.Id -User emily.pressey@anthesisgroup.com -RemoveRole "Full Control" 
    }



$results2 = $results | ? {@("E004128") -contains $_.KimbleOne__Reference__c } #check in the morning - have these applied? MArk them off above
$results2.ClientName
$results2.EngName


#This will select items between numbers using array logic
    for($i=0;$i -lt 4;$i++){
        [array]$results2.count += $results[$i]
        }


#Adding an extra bracket after the start like below (commented out) will pick the last thing in an array
        $results | % {
    $result = $_      
    $list = Get-PnPList $result.DocumentLibraryGuid
    $item = get-spoProjectFolder -pnpList $list -kimbleEngagementCodeToLookFor $result.KimbleOne__Reference__c -folderGuid $result.FolderGuid
    Set-PnPListItemPermission -List $list.Id -Identity $item.Id -User Transaction_and_Corporate_Services_FIN@anthesisgroup.com -AddRole Contribute -ClearExisting
    Set-PnPListItemPermission -List $list.Id -Identity $item.Id -User emily.pressey@anthesisgroup.com -RemoveRole "Full Control" 
    }
