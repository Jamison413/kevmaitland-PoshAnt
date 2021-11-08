#######################################################             
#                                                     #
#     Hubspot/Netsuite Object Matching                #
#                                                     #
#######################################################

#This script will match Netsuite contacts with Hubspot contacts:
#-get all Hubpspot and Netsuite contacts (only get a subest of fields form Hubspot as the API is slow to bring back 53k contacts - approx. 15 minutes)
#-for Clients, we do a bit of processing on the 'entityID' field to remove the artifical ID from the Client name. Contacts are fine as is.
#-we omit any Lavola contacts
#-we match Contacts on the email address and Clients on the name fields (email address for Contacts is mandatory in Hubspot so we shouldn't have anything going in without it - most unique identifier)
#-we match the Contact or Client name to the name in Hubspot, if found we update the NetsuiteId field in Hubspot with the InternalID of the Netsutie object - this should be numerical and should be unique per record type in Netsuite.
#The Contact/Client is thereby linked by the Netsuite InternalId field which we can query Hubspot with.

############################################                  
#                                          #
#                  Setup                   #
#                                          #
############################################

<#---Logging---#>
$Logname = "C:\Scripts" + "\Logs" + "\NetsuiteHuspotMatch $(Get-Date -Format "yyMMdd").log"
Start-Transcript -Path $Logname -Append
Write-Host "Script started:" (Get-date)
Write-host "**********************" -ForegroundColor White

<#---Import Modules---#>
Import-Module -Name 'C:\Users\Emily.Pressey\Documents\WindowsPowerShell\Modules\_PS_Library_GeneralFunctionality\_PS_Library_GeneralFunctionality.psm1'
Import-Module -Name 'C:\Users\Emily.Pressey\Documents\WindowsPowerShell\Modules\_REST_Library_NetSuite\_REST_Library_NetSuite.psm1'
Import-Module HubSpotCmdlets
$ss = Import-CliXml -Path  'C:\Users\Admin\Desktop\Hubspot.xml'
$key = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($ss))

#Get all contacts and companies from Netsuite function (slight amendment for a one off to remove the $query option)
function get-netSuiteContactsAllFromNetSuite(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $false)]
        [ValidatePattern('^?[\w+][=][\w+]')]
        [string]$query

        ,[parameter(Mandatory=$false)]
        [psobject]$netsuiteParameters
        )

    Write-Verbose "`tget-netSuiteContactFromNetSuite([$($query)])"
    if([string]::IsNullOrWhiteSpace($netsuiteParameters)){$netsuiteParameters = get-netsuiteParameters}

    $contacts = invoke-netsuiteRestMethod -requestType GET -url "https://3487287-sb1.suitetalk.api.netsuite.com/rest/platform/v1/record/contact" -netsuiteParameters $netsuiteParameters #-Verbose 
    $contactsEnumerated = [psobject[]]::new($contacts.count)
    for ($i=0; $i -lt $contacts.count;$i++) {
    $url = "$($contacts.items[$i].links[0].href)" + "/?expandSubResources=True"
    write-host $url -ForegroundColor white
        $contactsEnumerated[$i] = invoke-netsuiteRestMethod -requestType GET -url $url -netsuiteParameters $netsuiteParameters 
        }
    $contactsEnumerated
    }
function get-netSuiteClientsAllFromNetSuite(){
        [cmdletbinding()]
        Param (
            [parameter(Mandatory = $false)]
            [ValidatePattern('^?[\w+][=][\w+]')]
            [string]$query
    
            ,[parameter(Mandatory=$false)]
            [psobject]$netsuiteParameters
            )
    
        Write-Verbose "`tget-allNetSuiteClients([$($query)])"
        if([string]::IsNullOrWhiteSpace($netsuiteParameters)){$netsuiteParameters = get-netsuiteParameters}
    
        $customers = invoke-netsuiteRestMethod -requestType GET -url "https://3487287-sb1.suitetalk.api.netsuite.com/rest/platform/v1/record/customer" -netsuiteParameters $netsuiteParameters #-Verbose 
        $customers.items.links.href | out-file "C:\Users\Emily.Pressey\customers.txt"
        $customersEnumerated = [psobject[]]::new($customers.count)
        for ($i=0; $i -lt $customers.count;$i++) {
            $customersEnumerated[$i] = invoke-netsuiteRestMethod -requestType GET -url "$($customers.items[$i].links[0].href)/?expandSubResources=$true" -netsuiteParameters $netsuiteParameters 
            }
        $customersEnumerated
        }
    
############################################                  
#                                          #
#           Contact Matching               #
#                                          #
############################################

<#---Get all contacts from Hubspot---#>
$hubConn = Connect-HubSpot -ApiKey $key
$timetoexecutehub = Measure-Command {$HubContacts = Select-HubSpot -Connection $hubConn -Table "Contacts" -Columns @("[First name]","[Last name]","Email","[Company name]","VID","[NetsuiteId]")} #I take roughly 6 minutes to pull back 53k contacts
write-host "It took $($timetoexecutehub) to retrieve all NetsuiteID's from the Contacts table in Hubspot" -ForegroundColor White
write-host "$($HubContacts.count) Hubspot contacts returned" -ForegroundColor Yellow

<#---Get all contacts from Netsuite---#>
$timetoexecutenet = Measure-Command {$NetContacts = get-netSuiteContactFromNetSuite -Verbose}
write-host "It took $($timetoexecutenet) to retrieve all NetsuiteID's from the Contacts table in Hubspot" -ForegroundColor White
write-host "$($NetContacts.count) Netsuite contacts returned" -ForegroundColor Yellow

<#---Process and update Hubspot contacts with Netsuite InternalId---#>
ForEach($NetContact in $NetContacts){

    #Check it's not a Lavola contact
    If(("4" -ne ($NetContact.subsidiary.Id)) -and ("40" -ne ($NetContact.subsidiary.Id))){
        Write-Host "$($NetContact.entityId) - $($NetContact.subsidiary.refName): I'm NOT Lavola contact" -ForegroundColor Green
        #Look for it in Hubspot
        write-host "Searching for $($netcompanyname) in Hubspot..." -ForegroundColor Yellow
        $outcome = $HubContacts | Where-Object {$_.email -eq $NetContact.email}
            If($outcome){
                write-host "$($netcompanyname) found in Hubspot! Trying to update with Netsuite ID: $($NetClient.id)..." -ForegroundColor Yellow
                $success = (Update-HubSpot -Connection $hubConn -Table "Contacts" -Columns @("NetsuiteId") -Values @("$($NetContact.id)") -Where "VID = '$($outcome.VID)'")
                If($success){Write-Host "$($NetContact.email): I've been updated!"}
                Else{Write-Host "$($NetContact.email): I've failed to update, something has gone wrong..."}
            }
            Else{
            Write-Host "$($NetContact.email): I don't look like I'm in Hubspot...skipping"
            }
    }
    Else{
    Write-Host "$($NetContact.entityId) - $($NetContact.subsidiary.refName): I'm a Lavola contact" -ForegroundColor Red
    }
}

############################################                  
#                                          #
#           Client  Matching               #
#                                          #
############################################

<#---Get all clients from Hubspot---#>
$hubConn = Connect-HubSpot -ApiKey $key
$timetoexecutehub = Measure-Command {$HubClients = Select-HubSpot -Connection $hubConn -Table "Companies" -Columns @("Name","[Company ID]","[NetsuiteId]")}  #I take roughly 1 minutes to bring back 52k clients
write-host "It took $($timetoexecutehub) to retrieve all NetsuiteID's from the Companies table in Hubspot" -ForegroundColor White
write-host "$($HubClients.count) Hubspot contacts returned" -ForegroundColor Yellow

<#---Get all clients from Netsuite---#>
$timetoexecutenet = Measure-Command {$NetClients = get-netSuiteClientsAllFromNetSuite -Verbose} #I take roughly 10 minutes to bring back 600 clients
$timetoexecutenet
write-host "$($NetClients.count) Netsuite contacts returned" -ForegroundColor Yellow

#Process and update Hubspot clients with Netsuite InternalId
ForEach($NetClient in $NetClients){
    #Check it's not a Lavola client
    If(("4" -ne ($NetClient.subsidiary.Id)) -and ("40" -ne ($NetClient.subsidiary.Id))){
        Write-Host "$($NetClient.entityId) - $($NetClient.subsidiary.refName): I'm NOT Lavola client" -ForegroundColor Green
        
        #Remove the ID from the name and make it friendly
        $id, $otherstuff = ($NetClient.entityId).Split(" ")
        $netcompanyname = (($NetClient.entityId) -replace ($($id),"")).Trim()
        
        #Look for it in Hubspot
        write-host "Searching for $($netcompanyname) in Hubspot..." -ForegroundColor Yellow
        $outcome = $HubClients | Where-Object {$_.Name -eq $netcompanyname}
            If($outcome){
                write-host "$($netcompanyname) found in Hubspot! Trying to update with Netsuite ID: $($NetClient.id)..." -ForegroundColor Yellow
                $updatehubspot = (Update-HubSpot -Connection $hubConn -Table "Companies" -Columns @("NetsuiteId") -Values @("$($NetClient.id)") -Where "[Company ID] = '$($outcome.'Company ID')'")
                If($updatehubspot){Write-Host "$netcompanyname): I've been successfully updated with Nestuite ID $($NetClient.id)!" -ForegroundColor Green}
                Else{Write-Host "$($NetClient.entityId): I've failed to update, something has gone wrong..."}
            }
            Else{
            #For this script we don't want to amend the ones the aren't in Hubspot
            Write-Host "$($NetClient.entityId): I don't look like I'm in Hubspot...skipping"
            }
    }
    Else{
    #We don't want to amend anything that is a Lavola contact
    Write-Host "$($NetClient.entityId) - $($NetClient.subsidiary.refName): I'm a Lavola contact...skipping" -ForegroundColor Red
    }
}












