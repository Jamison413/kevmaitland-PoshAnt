###########################################                  
#                                         #      
#           HubspotNetsuite Sync          #
#                                         #
###########################################

<#
This script syncs data from Netsuite to Hubspot, using the Netsuite ID as the makeshift 'foreign key'. We sync two objects: Contacts and Companies and the NetsuiteID field features on both objects in Hubspot.
Due to the nature of the data, all data that is processed is held in memory and does not leave the session (it is not imported into an intermediary database). We sync the following fields one way currently:
+----------------+-----------------------+
| Contact Object |     Client Object     |
+----------------+-----------------------+
| Salutation     | Annual Revenue        | 
| First name*    | Client Sector         | 
| Last name*     | Company owner         |                        
| Job title*     | Generic email address |                        
| Email*         | Email                 |                        
| Contact owner* | Name*                 |                        
| Company name*  | Phone Number          |      
| Street Address |                       |                     
| City           |                       |                       
| State/Region   |                       |                       
| Country/Region |                       |                       
| Postal Code    |                       |                      
| Phone Number   |                       |                     
| Fax Number     |                       |                        
| Message        |                       |                       
+----------------+-----------------------+
(*mandatory field)
#>


###########################################                  
#                                         #      
#                 Setup                   #
#                                         #
###########################################

<#---Logging---#>
$Logname = "C:\Scripts" + "\Logs" + "\NetsuiteHuspotSync $(Get-Date -Format "yyMMdd").log"
Start-Transcript -Path $Logname -Append
Write-Host "Script started:" (Get-date)
Write-host "**********************" -ForegroundColor White

<#---Import Modules---#>
Import-Module -Name C:\Users\Emily.Pressey\Documents\WindowsPowerShell\PoshAnt\_PS_Library_GeneralFunctionality\_PS_Library_GeneralFunctionality.psm1
Import-Module -Name C:\Users\Emily.Pressey\Documents\WindowsPowerShell\PoshAnt\_REST_Library_NetSuite\_REST_Library_NetSuite.psm1
Import-Module HubSpotCmdlets
$ss = Import-CliXml -Path  'C:\Users\Admin\Desktop\Hubspot.xml'
$key = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($ss))

############################################                  
#                                          #
#               Contact Get                #
#                                          #
############################################

#Get todays date
$todaydate = (Get-Date -UFormat "%d/%m/%y")
#Get last run date
$lastrundate = Get-Content 'C:\Users\Emily.Pressey\OneDrive - Anthesis LLC\Documents\NetsuiteHubspotSync.txt' | Get-Date -UFormat "%d/%m/%y"

<#------------------------Get all Netsuite contacts - anything modified on or after the last run date------------------------#>
[string]$query = "?q=lastModifiedDate ON_OR_AFTER `"$($lastrundate)`""
$Netcontactsfull = get-netSuiteContactFromNetSuite -query $query -Verbose

#Get some more information about the contact and process the contacts into nice PS friendly objects
$Netcontacts = @()
ForEach($item in $Netcontactsfull){

#If it's not a Lavola contact (we don't want to sync these)
If(("4" -ne ($item.subsidiary.Id)) -and ("40" -ne ($item.subsidiary.Id))){
Write-Host "$($item.id) $($item.entityId) - $($item.subsidiary.refName): I'm NOT Lavola contact, so let's get more information about me" -ForegroundColor Green
$contactrecord = @()

    Write-Host "$($item.id) $($item.entityId) - $($item.subsidiary.refName): adding the original record" -ForegroundColor Yellow
    $contactrecord += New-Object PSObject $item #!might be some discrepency around Owner field

    #Get the address details - we will select the default billing address
    Write-Host "$($item.id) $($item.entityId) - $($item.subsidiary.refName): retrieving and processing the address record" -ForegroundColor Yellow
    $addressrecord = $item.addressbook.items | Where-Object -Property "defaultBilling" -EQ "True" | Select-Object -Property "addressbookaddress" #So we select the property with the juicy address data formed in a hashtable and throw it into a variable
    $contactrecord | Add-Member -MemberType NoteProperty -Name addr1 -Value $($addressrecord.addressbookaddress.addr1) -Force
    $contactrecord | Add-Member -MemberType NoteProperty -Name city -Value $($addressrecord.addressbookaddress.city) -Force
    $contactrecord | Add-Member -MemberType NoteProperty -Name state -Value $($addressrecord.addressbookaddress.state) -Force
    $contactrecord | Add-Member -MemberType NoteProperty -Name zip -Value $($addressrecord.addressbookaddress.zip) -Force
    #Get the correct country format - get list from codes in pre-made csv, Netsuite uses 2-letter code whereas Hubspot uses the full country name
    $countrydetails = import-csv 'C:\Users\Emily.Pressey\OneDrive - Anthesis LLC\Documents\Country Details.csv' #Need to move onto Netmon
    $country = $countrydetails | Where-Object {$_.'Short Name (Code)' -eq $addressrecord.addressbookaddress.country}
    $contactrecord | Add-Member -MemberType NoteProperty -Name country -Value $($country.Country) -Force
    
    #Get the company details for the Contact
    Write-Host "$($item.id) $($item.entityId) - $($item.subsidiary.refName): retrieving and processing the company record" -ForegroundColor Yellow
    [string]$companyrefname = $item.company.refName
    $cid,$companyrealname = "$($companyrefname)" -split " ",2
    If("C" -eq $companyrefname.Substring(0, [Math]::Min($companyrefname.Length, 1))){
    $query = "?q=Id EQUAL $($item.company.id)"
    $customer = get-netSuiteClientsFromNetSuite -query $query -Verbose
    $contactrecord | Add-Member -MemberType NoteProperty -Name annualrevenue -Value $($customer.custentity_esc_annual_revenue) -Force
    $contactrecord | Add-Member -MemberType NoteProperty -Name industry -Value $($customer.custentity_esc_industry.refName) -Force
    $contactrecord | Add-Member -MemberType NoteProperty -Name companyname -Value $($companyrealname) -Force
    $contactrecord | Add-Member -MemberType NoteProperty -Name cid -Value $($cid) -Force
    }
    Else{
    Write-Host "I've been attached to a project instead of a client...woops" -ForegroundColor Red
    }
    #Get the internal ID and Type - Internal ID's are only unique to their types, so good just to catch this in a property on it's own
    $contactrecord | Add-Member -MemberType NoteProperty -Name internalid -Value $($item.id) -Force
    $contactrecord | Add-Member -MemberType NoteProperty -Name recordType -Value "contact" -Force

    $Netcontacts += $contactrecord
}
Else{
Write-Host "$($item.entityId) - $($item.subsidiary.refName): I AM a Lavola contact...skipping" -ForegroundColor Red  
}
}

<#------------------------Get all Netsuite contacts in Hubspot by NetsuiteID column------------------------#>

#Get all the NetsuiteIDs from Hubspot because it's faster than querying all the fieds - this will give us all Netsuite contacts that currently exist in Hubspot
$hubConn = Connect-HubSpot -ApiKey $key
$timetoexecutehub = Measure-Command {$HubContacts = Select-HubSpot -Connection $hubConn -Table "Contacts" -Columns @("NetsuiteId")} #I take roughly 6 minutes to pull back 53k contacts
write-host "It took $($timetoexecutehub) to retrieve all NetsuiteID's from the Contacts table in Hubspot" -ForegroundColor White
write-host "$($HubContacts.count) Hubspot contacts returned" -ForegroundColor Yellow

#Process all contacts to get all the non-nulls (faster in Powershell opposed to querying Hubspot), this will give us a finalised list of NetsuiteIDs or it brings back everything
$NetsuiteIdsinHubspot = @()
Foreach($NetsuiteId in $HubContacts){
        $type = $NetsuiteId.NetsuiteId.GetType()
        if("DBNull" -ne $type.Name){
            write-host "Adding $($NetsuiteId.NetsuiteId) to NetsuiteIdsinHubspot"
            $NetsuiteIdsinHubspot += New-Object PSObject $NetsuiteId.NetsuiteId
            }  
}
Write-Host "***We currently have $($NetsuiteIdsinHubspot.count) NetsuiteIDs in Hubspot. Now we can check if a Netsuite contact already exists in Hubspot...if not we will create it...***" -ForegroundColor Yellow

<#------------------------Get all Hubspot Owners------------------------#>
#Get the Owners table (a Read-Only view in Hubspot - Owners must be created manually as the Hubspot API does not allow us to edit this table, only query it)
$HubOwners = Select-HubSpot -Connection $hubconn -Table "Owners" -Verbose

############################################                  
#                                          #
#           Contact Update                 #
#                                          #
############################################

<#------------------------See if Netsuite contact Id is already in the list of Hubspot NetsuiteIDs------------------------#>
#Check if it exists by comparing the NetsuiteId's in Hubspot to the Contacts NetsuiteId, if it does update all the fields - comparing each one would be more processing time so faster to update them all
Foreach($Netcontact in $Netcontacts){        
    $outcome = $NetsuiteIdsinHubspot | Where-Object {$_ -eq $Netcontact.internalid}
    #So, if I already exist, try to update me...
    If($outcome){
        write-host "Existing Contact: NetsuiteId $($Netcontact.internalid) found in Hubspot! Trying to update fields for $($Netcontact.email)..." -ForegroundColor Yellow
        #Match Owner to Hubspot
        $HubOwner = $HubOwners | Where-Object {$_.Email -eq $NetContact.owner}
        #Update fields 
        Try{
        $UpdateHubContact = Update-HubSpot -Connection $hubConn -Table "Contacts" -Columns @("Salutation","First name","Last name","Job title", "Email","Contact owner","Company name","Street Address","City","State/Region","Country/Region","Postal Code","Phone Number","Fax Number","Message") -Values @("$($Netcontact.salutation)","$($Netcontact.firstName)","$($Netcontact.lastName)","$($Netcontact.title)","$($Netcontact.email)","$($HubOwner.OwnerId)","$($Netcontact.companyname)","$Netcontact.addr1","$Netcontact.city","$Netcontact.state","$Netcontact.country","$Netcontact.zip", "$Netcontact.officePhone","$Netcontact.fax","$Netcontact.comments") -Where "NetsuiteId = '$($Netcontact.internalid)'"
        write-host "$($timetoexecuteupdatecontact.Seconds) seconds to update existing Hubspot Contact" -ForegroundColor White
        }
        Catch{
        $Error
        Write-Host "Failed to update fields for $($Netcontact.email)" -ForegroundColor Red
        }
    }
        #And if I don't exist, try to create me!
        #Match Owner to Hubspot
        $HubOwner = $HubOwners | Where-Object {$_.Email -eq $NetContact.owner}
        write-host "New Contact: NetsuiteId $($Netcontact.internalid) NOT found in Hubspot! Trying to create new contact in Hubspot for $($Netcontact.email)..." -ForegroundColor Yellow
        Try{
        #If the email is not unique, it does not create the contact and will fail
        $timetoexecutenewcontact = Measure-Command {$NewHubContact = Add-HubSpot -Connection $hubConn -Table "Contacts" -Columns @("Salutation","First name","Last name","Job title", "Email","Contact owner","Company name","Street Address","City","State/Region","Country/Region","Postal Code","Phone Number","Fax Number","Message","NetsuiteId") -Values @("$($Netcontact.salutation)","$($Netcontact.firstName)","$($Netcontact.lastName)","$($Netcontact.title)","$($Netcontact.email)","$($HubOwner.OwnerId)","$($Netcontact.companyname)","$($Netcontact.addr1)","$($Netcontact.city)","$($Netcontact.state)","United Kingdom","$($Netcontact.zip)", "$($Netcontact.officePhone)","$($Netcontact.fax)","$($Netcontact.comments)","$($Netcontact.internalid)")}
        write-host "$($timetoexecutenewcontact.Seconds) seconds to create new Hubspot Contact" -ForegroundColor White
        }
        Catch{
        #$Error
        Write-Host "Failed to create contact for $($Netcontact.email)" -ForegroundColor Red
        }
}
<#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#>

############################################                  
#                                          #
#                Client Get                #
#                                          #
############################################

#Get today's date
$todaydate = (Get-Date -Month 1 -UFormat "%d/%m/%y")
#Get last run date
$lastrundate = Get-Content 'C:\Users\Emily.Pressey\OneDrive - Anthesis LLC\Documents\NetsuiteHubspotSync.txt' | Get-Date -UFormat "%d/%m/%y"

<#------------------------Get all Netsuite contacts - anything modified after the last run date------------------------#>
[string]$query = "?q=lastModifiedDate ON_OR_AFTER `"$($lastrundate)`""
$Netclientsfull = get-netSuiteClientsFromNetSuite -query $query -Verbose

#Process the contacts into nice friendly objects
$Netclients = @()
ForEach($item in $Netclientsfull){

#If it's not a Lavola client
If(("4" -ne ($item.subsidiary.Id)) -and ("40" -ne ($item.subsidiary.Id))){
Write-Host "$($item.entityId) - $($item.subsidiary.refName): I'm NOT Lavola client" -ForegroundColor Green
$clientrecord = @()

    $clientrecord += New-Object PSObject $item
        
    #Sanitise Netsuite company name
    $id, $name = ($item.entityId).Split(" ")
    $netcompanyname = (($item.entityId) -replace ($($id),"")).Trim()
    $clientrecord | Add-Member -MemberType NoteProperty -Name companyname -Value $($netcompanyname) -Force

    ####NOTE: Contact owner isn't an option on the form currently, so we make it up
    $clientrecord | Add-Member -MemberType NoteProperty -Name owner -Value "emily.pressey@anthesisgroup.com" -Force

    #Get the internal ID and Type - Internal ID's are only unique to their types, so good just to catch this in a property on it's own
    $clientrecord | Add-Member -MemberType NoteProperty -Name internalid -Value $($item.id) -Force
    $clientrecord | Add-Member -MemberType NoteProperty -Name recordType -Value "client" -Force

    $Netclients += $clientrecord
}
Else{
Write-Host "$($item.entityId) - $($item.subsidiary.refName): I AM a Lavola client...skipping" -ForegroundColor Red  
}
}


<#------------------------Get all Netsuite contacts in Hubspot by NetsuiteID column------------------------#>

#Get all the NetsuiteIDs from Hubspot because it's faster than querying all the fieds - this will give us all Netsuite contacts that currently exist in Hubspot
$hubConn = Connect-HubSpot -ApiKey $key
$timetoexecutehub = Measure-Command {$HubClients = Select-HubSpot -Connection $hubConn -Table "Companies" -Columns @("NetsuiteId")}
write-host "It took $($timetoexecutehub) to retrieve all NetsuiteID's from the Company table in Hubspot" -ForegroundColor White
write-host "$($HubClients.count) Hubspot Clients returned" -ForegroundColor Yellow

#Process it to get all the non-nulls (faster in Powershell opposed to querying Hubspot), this will give us a finalised list or it brings back everything
$NetsuiteIdsinHubspot = @()
Foreach($NetsuiteId in $HubClients){
        $type = $NetsuiteId.NetsuiteId.GetType()
        if("DBNull" -ne $type.Name){
            write-host "Adding $($NetsuiteId.NetsuiteId) to NetsuiteIdsinHubspot"
            $NetsuiteIdsinHubspot += New-Object PSObject $NetsuiteId.NetsuiteId
            }  
}
Write-Host "***We currently have $($NetsuiteIdsinHubspot.count) NetsuiteIDs in Hubspot. Now we can check if a Netsuite client already exists in Hubspot...if not we will create it...***" -ForegroundColor Yellow

############################################                  
#                                          #
#             Client Update                #
#                                          #
############################################

<#------------------------See if Netsuite Client Id is already in the list of Hubspot NetsuiteIDs------------------------#>
#Check if it exists by comparing the NetsuiteId's in Hubspot to the Client's NetsuiteId, if it does update all the fields - comparing each one would be more processing time so faster to update them all
Foreach($Netclient in $Netclientsfull){        

    $outcome = $NetsuiteIdsinHubspot | Where-Object {$_ -eq $Netclient.internalid}
    #So, if I already exist, try to update me...
    If($outcome){
        write-host "Existing Client: NetsuiteId $($Netclient.id) found in Hubspot! Trying to update fields for $($Netclient.email)..." -ForegroundColor Yellow
        #Match Owner to Hubspot
        $HubOwner = $HubOwners | Where-Object {$_.Email -eq $Netclient.salesRep} #this might be different depending on fields/table after changes
        #Update fields 
        Try{
        $timetoexecuteupdateclient = Measure-Command {$UpdateHubClient = Update-HubSpot -Connection $hubConn -Table "Companies" @("Annual Revenue","Client Sector","Company owner","Generic email address", "Email","Name","Phone Number") -Values @("$($Netclient.custentity_esc_annual_revenue)","$($Netclient.custentity_esc_industry)","$($HubOwner.OwnerId)","$($Netclient.custentity_2663_email_address_notif)","$($netcompanyname)","$($Netclient.phone)") -Where "NetsuiteId = '$($Netclient.internalid)'"}
        write-host "$($timetoexecuteupdateclient.Seconds) seconds to update existing Hubspot Contact" -ForegroundColor Yellow
        }
        Catch{
        $Error
        Write-Host "Failed to update fields for $($netcompanyname)" -ForegroundColor Red
        }
    }
        #And if I don't exist, try to create me!
        write-host "New Client: NetsuiteId $($Netclient.id) NOT found in Hubspot! Trying to create new client in Hubspot for $($netcompanyname)..." -ForegroundColor Yellow
        #Match Owner to Hubspot
        $HubOwner = $HubOwners | Where-Object {$_.Email -eq $Netclient.salesRep} #this might be different depending on fields/table after changes
        #Sanitise Netsuite name
        $id, $name = ($Netclient.entityId).Split(" ")
        $netcompanyname = (($Netclient.entityId) -replace ($($id),"")).Trim()        
        Try{        
        $timetoexecutenewclient = Measure-Command {$NewHubClient = Add-HubSpot -Connection $hubConn -Table "Companies" @("Annual Revenue","Client Sector","Company owner","Generic email address", "Email","Name","Phone Number") -Values @("$($Netclient.custentity_esc_annual_revenue)","$($Netclient.custentity_esc_industry.refName)","$($HubOwner.OwnerId)","$($Netclient.custentity_2663_email_address_notif)","$($netcompanyname)","$($Netclient.phone)","$($Netclient.id)")}
        write-host "$($timetoexecutenewclient.Seconds) seconds to create new Hubspot Contact" -ForegroundColor Yellow
        }
        Catch{
        $Error
        Write-Host "Failed to create contact for $($netcompanyname)" -ForegroundColor Red
        }
}

############################################                  
#                                          #
#             Finishing Up                 #
#                                          #
############################################

#Set the last run date
Get-Date -UFormat "%d/%m/%y" | Out-File 'C:\Users\Emily.Pressey\OneDrive - Anthesis LLC\Documents\NetsuiteHubspotSync.txt'

Stop-Transcript