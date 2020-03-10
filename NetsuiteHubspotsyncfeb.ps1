####################                  
#                  #
#     Setup        #
#                  #
####################

<#---Import Modules---#>
Import-Module -Name 'C:\Users\Emily.Pressey\Documents\WindowsPowerShell\Modules\_PS_Library_GeneralFunctionality\_PS_Library_GeneralFunctionality.psm1'
Import-Module -Name 'C:\Users\Emily.Pressey\Documents\WindowsPowerShell\Modules\_REST_Library_NetSuite\_REST_Library_NetSuite.psm1'
Import-Module HubSpotCmdlets

#remove-module -name _REST_Library_NetSuite

############################################                  
#                                          #
#           Contact Get                    #
#                                          #
############################################
#$error[0] | Out-File 'C:\Users\Emily.Pressey\OneDrive - Anthesis LLC\Documents\NetsuiteHubspotSync.txt'

Get-Date -UFormat "%d/%m/%y" | Out-File 'C:\Users\Emily.Pressey\OneDrive - Anthesis LLC\Documents\NetsuiteHubspotSync.txt'

#Get todays date
$todaydate = (Get-Date -Month 1 -UFormat "%d/%m/%y")

#Get last run date
$lastrundate = Get-Content 'C:\Users\Emily.Pressey\OneDrive - Anthesis LLC\Documents\NetsuiteHubspotSync.txt' | Get-Date -UFormat "%d/%m/%y"

<#------------------------Get all Netsuite contacts - anything modified after the last run date------------------------#>
[string]$query = "?q=lastModifiedDate ON_OR_AFTER `"$($todaydate)`""
#[string]$query = "?q=middleName CONTAIN P"
$Netcontactsfull = get-netSuiteContactFromNetSuite -query $query -Verbose


#Process the contacts into nice friendly objects
$Netcontacts = @()
ForEach($item in $Netcontactsfull){

#If it's not a Lavola contact
If(("4" -ne ($item.subsidiary.Id)) -and ("40" -ne ($item.subsidiary.Id))){
Write-Host "$($item.entityId) - $($item.subsidiary.refName): I'm NOT Lavola contact" -ForegroundColor Green
$contactrecord = @()

    $contactrecord += New-Object PSObject $item
    #Get the address details - we will select the default billing address
    $addressrecord = $item.addressbook.items | Where-Object -Property "defaultBilling" -EQ "True" | Select-Object -Property "addressbookaddress" #So we select the property with the juicy address data formed in a hashtable and throw it into a variable
    $contactrecord | Add-Member -MemberType NoteProperty -Name addr1 -Value $($addressrecord.addressbookaddress.addr1) -Force
    $contactrecord | Add-Member -MemberType NoteProperty -Name city -Value $($addressrecord.addressbookaddress.city) -Force
    $contactrecord | Add-Member -MemberType NoteProperty -Name state -Value $($addressrecord.addressbookaddress.state) -Force
    $contactrecord | Add-Member -MemberType NoteProperty -Name zip -Value $($addressrecord.addressbookaddress.zip) -Force

#Function to switch out two letter country code and add full country name

    $contactrecord | Add-Member -MemberType NoteProperty -Name country -Value $($addressrecord.addressbookaddress.country) -Force
    $contactrecord | Add-Member -MemberType NoteProperty -Name companyname -Value $($item.company.refName) -Force

    ####NOTE: Contact owner isn't an option on the form currently, so we make it up
    $contactrecord | Add-Member -MemberType NoteProperty -Name owner -Value "emily.pressey@anthesisgroup.com" -Force

    
    #Get the company details
    $url = $($item.company.links.href)
    $customer = invoke-netsuiteRestMethod -requestType GET -url $url -Verbose
    $contactrecord | Add-Member -MemberType NoteProperty -Name annualrevenue -Value $($customer.custentity_esc_annual_revenue) -Force
    $contactrecord | Add-Member -MemberType NoteProperty -Name industry -Value $($customer.custentity_esc_industry.refName) -Force

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
$hubConn = Connect-HubSpot -ApiKey "4bf5a895-e0d1-4012-9fe6-e6341424920c"
$timetoexecutehub = Measure-Command {$HubContacts = Select-HubSpot -Connection $hubConn -Table "Contacts" -Columns @("NetsuiteId")} #I take roughly 6 minutes to pull back 53k contacts
$timetoexecutehub

<#------------------------Get all Hubspot Owners------------------------#>
$HubOwners = Select-HubSpot -Connection $hubconn -Table "Owners" #-Where "Email = 'emily.pressey@anthesisgroup.com'" -Verbose


write-host "$($HubContacts.count) Hubspot contacts returned" -ForegroundColor Yellow
#Process it to get all the non-nulls (faster in Powershell opposed to querying Hubspot), this will give us a finalised list or it brings back everything
$NetsuiteIdsinHubspot = @()
Foreach($NetsuiteId in $HubContacts){
        $type = $NetsuiteId.NetsuiteId.GetType()
        if("DBNull" -ne $type.Name){
            write-host "Adding $($NetsuiteId.NetsuiteId) to NetsuiteIdsinHubspot"
            $NetsuiteIdsinHubspot += New-Object PSObject $NetsuiteId.NetsuiteId
            }  
}
Write-Host "We currently have $($NetsuiteIdsinHubspot.count) NetsuiteIDs in Hubspot. Now we can check if a Netsuite contact already exists in Hubspot...if not we will create it..." -ForegroundColor Yellow
                                            

############################################                  
#                                          #
#           Contact Update                 #
#                                          #
############################################


<#------------------------See if Netsuite contact Id is already in the list of Hubspot NetsuiteIDs------------------------#>
#Check if it exists by comparing the NetsuiteId's in Hubspot to the Contacts NetsuiteId, if it does update all the fields - comparing each one would be more processing time so faster to update them all
Foreach($Netcontact in $Netcontacts[1]){        

    $outcome = $NetsuiteIdsinHubspot | Where-Object {$_ -eq $Netcontact.internalid}

    #So, if I already exist, try to update me...
    If($outcome){
        write-host "Existing Contact: NetsuiteId $($Netcontact.internalid) found in Hubspot! Trying to update fields for $($Netcontact.email)..." -ForegroundColor Yellow
        #Match Owner to Hubspot
        $HubOwner = $HubOwners | Where-Object {$_.Email -eq $NetContact.owner}
        #Update fields 
        Try{
        $timetoexecuteupdatecontact = Measure-Command {$UpdateHubContact = Update-HubSpot -Connection $hubConn -Table "Contacts" -Columns @("Salutation","First name","Last name","Job title", "Email","Contact owner","Company name","Street Address","City","State/Region","Country/Region","Postal Code","Phone Number","Fax Number","Message") -Values @("$($Netcontact.salutation)","$($Netcontact.firstName)","$($Netcontact.lastName)","$($Netcontact.title)","$($Netcontact.email)","$($HubOwner.OwnerId)","$($Netcontact.companyname)","$Netcontact.addr1","$Netcontact.city","$Netcontact.state","$Netcontact.country","$Netcontact.zip", "$Netcontact.officePhone","$Netcontact.fax","$Netcontact.comments") -Where "NetsuiteId = '$($Netcontact.internalid)'"}
        write-host "$($timetoexecuteupdatecontact.Seconds) seconds to update existing Hubspot Contact" -ForegroundColor Yellow
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
        $timetoexecutenewcontact = Measure-Command {$NewHubContact = Add-HubSpot -Connection $hubConn -Table "Contacts" -Columns @("Salutation","First name","Last name","Job title", "Email","Contact owner","Company name","Street Address","City","State/Region","Country/Region","Postal Code","Phone Number","Fax Number","Message","NetsuiteId") -Values @("$($Netcontact.salutation)","$($Netcontact.firstName)","$($Netcontact.lastName)","$($Netcontact.title)","$($Netcontact.email)","$($HubOwner.OwnerId)","$($Netcontact.companyname)","$($Netcontact.addr1)","$($Netcontact.city)","$($Netcontact.state)","United Kingdom","$($Netcontact.zip)", "$($Netcontact.officePhone)","$($Netcontact.fax)","$($Netcontact.comments)","$($Netcontact.internalid)")}
        write-host "$($timetoexecutenewcontact.Seconds) seconds to create new Hubspot Contact" -ForegroundColor Yellow
        }
        Catch{
        $Error
        Write-Host "Failed to create contact for $($Netcontact.email)" -ForegroundColor Red
        }
}


<#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#>

############################################                  
#                                          #
#                Client Get                #
#                                          #
############################################

$todaydate = (Get-Date -Month 1 -UFormat "%d/%m/%y")

#Get last run date
$lastrundate = Get-Content 'C:\Users\Emily.Pressey\OneDrive - Anthesis LLC\Documents\NetsuiteHubspotSync.txt' | Get-Date -UFormat "%d/%m/%y"

<#------------------------Get all Netsuite contacts - anything modified after the last run date------------------------#>
[string]$query = "?q=lastModifiedDate ON_OR_AFTER `"$($todaydate)`""
$Netclientsfull[0] = get-netSuiteClientsFromNetSuite -query $query -Verbose

#Process the contacts into nice friendly objects
$Netclients = @()
ForEach($item in $Netclientsfull){

#If it's not a Lavola client
If(("4" -ne ($item.subsidiary.Id)) -and ("40" -ne ($item.subsidiary.Id))){
Write-Host "$($item.entityId) - $($item.subsidiary.refName): I'm NOT Lavola client" -ForegroundColor Green
$clientrecord = @()

    $clientrecord += New-Object PSObject $item
        
    #Sanitise Netsuite name
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
$hubConn = Connect-HubSpot -ApiKey "4bf5a895-e0d1-4012-9fe6-e6341424920c"
$timetoexecutehub = Measure-Command {$HubClients = Select-HubSpot -Connection $hubConn -Table "Companies" -Columns @("NetsuiteId")}
$timetoexecutehub

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
Write-Host "We currently have $($NetsuiteIdsinHubspot.count) NetsuiteIDs in Hubspot. Now we can check if a Netsuite client already exists in Hubspot...if not we will create it..." -ForegroundColor Yellow


############################################                  
#                                          #
#             Client Update                #
#                                          #
############################################

<#------------------------See if Netsuite Client Id is already in the list of Hubspot NetsuiteIDs------------------------#>
#Check if it exists by comparing the NetsuiteId's in Hubspot to the Client's NetsuiteId, if it does update all the fields - comparing each one would be more processing time so faster to update them all
Foreach($Netclient in $Netclientsfull[0]){        

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

#Need to swap out api return with processed stuff

