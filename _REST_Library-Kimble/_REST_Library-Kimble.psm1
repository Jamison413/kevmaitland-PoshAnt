[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
#$callbackUri = "https://login.salesforce.com/services/oauth2/token"
#"https://test.salesforce.com/services/oauth2/token"
#$grantType = "password"
#$myInstance = "https://eu5.salesforce.com"
#$queryUri = "$myInstance/services/data/v39.0/query/?q="
#$querySuffixStub = " -H `"Authorization: Bearer "
#$kimbleLogin = Import-Csv "$env:USERPROFILE\Desktop\Kimble.txt"
#$clientId = $kimbleLogin.clientId
#$clientSecret = $kimbleLogin.clientSecret
#$username = $kimbleLogin.username
#$password = $kimbleLogin.password
#$securityToken = $kimbleLogin.securityToken
Import-Module _PS_Library_GeneralFunctionality

function add-kimbleAccountToFocalPointCache{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [PSCustomObject]$kimbleAccount 

        ,[parameter(Mandatory = $true)]
        [System.Data.Common.DbConnection]$dbConnection
        )
    Write-Verbose "add-kimbleAccountToFocalPointCache [$($kimbleAccount.Name)]"
    $sql = "SELECT Name,Id FROM SUS_Kimble_Accounts WHERE Id = '$($kimbleAccount.Id)'"
    Write-Verbose "`t$sql"
    $alreadyPresent = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $dbConnection
    if ($alreadyPresent){
        Write-Verbose "`tId [$($kimbleAccount.Id)] already present in Cache, not adding duplicate"
        -1 #Return "unsuccessful" as the record is in the DB, but we didn't add it and this might need investigating
        }
    else{
        Write-Verbose "`tId [$($kimbleAccount.Id)] not present in SQL, adding to [SUS_Kimble_Accounts]"
        $sql = "INSERT INTO SUS_Kimble_Accounts (attributes, Website, Type, SystemModstamp, Phone, ParentId, OwnerId, Name, LastModifiedDate, LastModifiedById, KimbleOne__IsCustomer__c, KimbleOne__BusinessUnit__c, Is_Partner__c, Is_Competitor__c, IsDeleted, Id, CreatedDate, CreatedById, Client_Sector__c, BillingStreet, BillingState, BillingPostalCode, BillingCountry, BillingCity, Description , DimCode, DimCode_Supplier, IsMissingFromKimble, IsMisclassified, IsDirty, isOrphaned, DocumentLibraryGuid) VALUES ("
        $sql += $(sanitise-forSqlValue -value $kimbleAccount.attributes.url -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.Website -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.Type -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.SystemModstamp -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.Phone -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.ParentId -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.OwnerId -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.Name -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.LastModifiedDate -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.LastModifiedById -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.KimbleOne__IsCustomer__c -dataType Boolean)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.KimbleOne__BusinessUnit__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.Is_Partner__c -dataType Boolean)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.Is_Competitor__c -dataType Boolean)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.IsDeleted -dataType Boolean)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.Id -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.CreatedDate -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.CreatedById -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.Client_Sector__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.BillingStreet -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.BillingState -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.BillingPostalCode -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.BillingCountry -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.BillingCity -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.Description -dataType String)
        $sql += ",NULL" #DimCode
        $sql += ",NULL" #DimCode_Supplier
        $sql += ",0" #IsMissingFromKimble
        $sql += ",0" #IsMisclassified
        $sql += ",1" #IsDirty
        $sql += ",0" #isOrphaned
        $sql += ",NULL" #DocumentLibraryGuid
        $sql += ")"
        Write-Verbose "`t$sql"
        $result = Execute-SQLQueryOnSQLDB -query $sql -queryType nonquery -sqlServerConnection $dbConnection
        if($result -eq 1){Write-Verbose "`t`tSUCCESS!"}
        else{Write-Verbose "`t`tFAILURE :( - Code: $result"}
        $result
        }
    }
function add-kimbleAccountMigrationDataToFocalPointCache{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [PSCustomObject]$kimbleAccount 

        ,[parameter(Mandatory = $true)]
        [System.Data.Common.DbConnection]$dbConnection
        )
    Write-Verbose "add-kimbleAccountToFocalPointCache [$($kimbleAccount.Name)]"
    $sql = "SELECT Id FROM SUS_Kimble_Accounts_MigrationData WHERE Id = '$($kimbleAccount.Id)'"
    Write-Verbose "`t$sql"
    $alreadyPresent = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $dbConnection
    if ($alreadyPresent){
        Write-Verbose "`tId [$($kimbleAccount.Id)] already present in Cache, not adding duplicate"
        -1 #Return "unsuccessful" as the record is in the DB, but we didn't add it and this might need investigating
        }
    else{
        Write-Verbose "`tId [$($kimbleAccount.Id)] not present in SQL, adding to [SUS_Kimble_Accounts_MigrationData]"
        $sql = "INSERT INTO SUS_Kimble_Accounts_MigrationData (Id,GenericEmail,GenericPhone,ParentId,CurrencyType,Terms,VatNumber,ShippingStreet,ShippingCity,ShippingState,ShippingPostalCode,ShippingCountry,Locale) VALUES ("
        $sql += $(sanitise-forSqlValue -value $kimbleAccount.Id -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.Generic_email_address__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.Phone -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.ParentId -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.CurrencyIsoCode -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.KimbleOne__InvoicePaymentTermDays__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.KimbleOne__TaxCodeReference__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.ShippingStreet -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.ShippingCity -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.ShippingState -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.ShippingPostalCode -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.ShippingCountry -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleAccount.KimbleOne__Locale__c -dataType String)
        $sql += ")"
        Write-Verbose "`t$sql"
        $result = Execute-SQLQueryOnSQLDB -query $sql -queryType nonquery -sqlServerConnection $dbConnection
        if($result -eq 1){Write-Verbose "`t`tSUCCESS!"}
        else{Write-Verbose "`t`tFAILURE :( - Code: $result"}
        $result
        }
    }
function add-kimbleContactToFocalPointCache($kimbleContact, $dbConnection, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Yellow "add-kimbleAccountToFocalPointCache"}
    $sql = "SELECT Name,Id FROM SUS_Kimble_Contacts WHERE Id = '$($kimbleContact.Id)'"
    $alreadyPresent = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $sqlDbConn
    if ($alreadyPresent){
        if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`Id already present in Cache, not adding duplicate"}
        -1 #Return "unsuccessful" as the record is in the DB, but we didn't add it and this might need investigating
        }
    else{
        $sql = "INSERT INTO SUS_Kimble_Contacts (AccountId,Anthesis_Events__c,AssistantName,AssistantPhone,attributes,Birthdate,Cleanup__c,Client_Type__c,CreatedById,CreatedDate,CurrencyIsoCode,Department,Description,Email,EmailBouncedDate,EmailBouncedReason,Fax,FirstName,Gender__c,Have_you_completed_this_section__c,HomePhone,Id,IsDeleted,IsEmailBounced,Jigsaw,JigsawContactId,Key_areas_of_interest__c,LastActivityDate,LastCURequestDate,LastCUUpdateDate,LastModifiedById,LastModifiedDate,LastName,LastReferencedDate,LastViewedDate,Lead_Source_Detail__c,LeadSource,Linked_In__c,MailingAddress,MailingCity,MailingCountry,MailingGeocodeAccuracy,MailingLatitude,MailingLongitude,MailingPostalCode,MailingState,MailingStreet,MasterRecordId,Met_At__c,MobilePhone,Name,Nee__c,Newsletters_Campaigns__c,Nickname__c,No_Show_Event_Attendees__c,Other_Email__c,OtherAddress,OtherCity,OtherCountry,OtherGeocodeAccuracy,OtherLatitude,OtherLongitude,OtherPhone,OtherPostalCode,OtherState,OtherStreet,OwnerId,Phone,PhotoUrl,Previous_Company__c,Region__c,ReportsToId,Role_Responsibilities__c,Salutation,Secondary_contact_owner__c,Skype__c,SystemModstamp,Title,Twitter__c,Unsubscribe_Newsletter_Campaigns__c) VALUES ("
        $sql += "'"+$(sanitise-forSql $kimbleContact.AccountId)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Anthesis_Events__c)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.AssistantName)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.AssistantPhone)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.attributes.url)+"',"
        $sql += "'"+$(sanitise-forSql $(Get-Date (smartReplace $kimbleContact.Birthdate "+0000" "") -Format s -ErrorAction SilentlyContinue))+"',"
        if($kimbleContact.Cleanup__c -eq $true){$sql += "1,"}else{$sql += "0,"}
        $sql += "'"+$(sanitise-forSql $kimbleContact.Client_Type__c)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.CreatedById)+"',"
        $sql += "'"+$(sanitise-forSql $(Get-Date (smartReplace $kimbleContact.CreatedDate "+0000" "") -Format s -ErrorAction SilentlyContinue))+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.CurrencyIsoCode)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Department)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Description)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Email)+"',"
        $sql += "'"+$(sanitise-forSql $(Get-Date (smartReplace $kimbleContact.EmailBouncedDate "+0000" "") -Format s -ErrorAction SilentlyContinue))+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.EmailBouncedReason)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Fax)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.FirstName)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Gender__c)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Have_you_completed_this_section__c)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.HomePhone)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Id)+"',"
        if($kimbleContact.IsDeleted -eq $true){$sql += "1,"}else{$sql += "0,"}
        if($kimbleContact.IsEmailBounced -eq $true){$sql += "1,"}else{$sql += "0,"}
        $sql += "'"+$(sanitise-forSql $kimbleContact.Jigsaw)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.JigsawContactId)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Key_areas_of_interest__c)+"',"
        $sql += "'"+$(sanitise-forSql $(Get-Date (smartReplace $kimbleContact.LastActivityDate "+0000" "") -Format s -ErrorAction SilentlyContinue))+"',"
        $sql += "'"+$(sanitise-forSql $(Get-Date (smartReplace $kimbleContact.LastCURequestDate "+0000" "") -Format s -ErrorAction SilentlyContinue))+"',"
        $sql += "'"+$(sanitise-forSql $(Get-Date (smartReplace $kimbleContact.LastCUUpdateDate "+0000" "") -Format s -ErrorAction SilentlyContinue))+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.LastModifiedById)+"',"
        $sql += "'"+$(sanitise-forSql $(Get-Date (smartReplace $kimbleContact.LastModifiedDate "+0000" "") -Format s -ErrorAction SilentlyContinue))+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.LastName)+"',"
        $sql += "'"+$(sanitise-forSql $(Get-Date (smartReplace $kimbleContact.LastReferencedDate "+0000" "") -Format s -ErrorAction SilentlyContinue))+"',"
        $sql += "'"+$(sanitise-forSql $(Get-Date (smartReplace $kimbleContact.LastViewedDate "+0000" "") -Format s -ErrorAction SilentlyContinue))+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Lead_Source_Detail__c)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.LeadSource)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Linked_In__c)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.MailingAddress)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.MailingCity)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.MailingCountry)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.MailingGeocodeAccuracy)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.MailingLatitude)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.MailingLongitude)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.MailingPostalCode)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.MailingState)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.MailingStreet)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.MasterRecordId)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Met_At__c)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.MobilePhone)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Name)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Nee__c)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Newsletters_Campaigns__c)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Nickname__c)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.No_Show_Event_Attendees__c)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Other_Email__c)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.OtherAddress)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.OtherCity)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.OtherCountry)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.OtherGeocodeAccuracy)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.OtherLatitude)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.OtherLongitude)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.OtherPhone)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.OtherPostalCode)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.OtherState)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.OtherStreet)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.OwnerId)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Phone)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.PhotoUrl)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Previous_Company__c)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Region__c)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.ReportsToId)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Role_Responsibilities__c)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Salutation)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Secondary_contact_owner__c)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Skype__c)+"',"
        $sql += "'"+$(sanitise-forSql $(Get-Date (smartReplace $kimbleContact.SystemModstamp "+0000" "") -Format s -ErrorAction SilentlyContinue))+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Title)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Twitter__c)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Unsubscribe_Newsletter_Campaigns__c)+"')"

        if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`t$sql"}
        $result = Execute-SQLQueryOnSQLDB -query $sql -queryType nonquery -sqlServerConnection $dbConnection
        if($verboseLogging){if($result -eq 1){Write-Host -ForegroundColor DarkYellow "`t`tSUCCESS!"}else{Write-Host -ForegroundColor DarkYellow "`t`tFAILURE :( - Code: $result"}}
        $result
        }
    }
function add-kimbleEngagementToFocalPointCache{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [PSCustomObject]$kimbleEngagement 

        ,[parameter(Mandatory = $true)]
        [System.Data.Common.DbConnection]$dbConnection
        )
    Write-Verbose "add-kimbleEngagementToFocalPointCache [$($kimbleEngagement.Name)]"
    $sql = "SELECT Name,Id FROM SUS_Kimble_Engagements WHERE Id = '$($kimbleEngagement.Id)'"
    Write-Verbose "`t$sql"
    $alreadyPresent = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $dbConnection
    if ($alreadyPresent){
        Write-Verbose "`tId [$($kimbleEngagement.Id)] already present in Cache, not adding duplicate"
        -1 #Return "unsuccessful" as the record is in the DB, but we didn't add it and this might need investigating
        }
    else{
        Write-Verbose "`tId [$($kimbleEngagement.Id)] not present in SQL, adding to [SUS_Kimble_Engagements]"
        $sql = "INSERT INTO SUS_Kimble_Engagements (attributes,Community__c,Competitive__c,CreatedById,CreatedDate,CurrencyIsoCode,Finance_Contact__c,Id,IsDeleted,Is_Engagement_Owner__c,KimbleOne__Account__c,KimbleOne__ActualUsage__c,KimbleOne__BaselineContractCost__c,KimbleOne__BaselineContractMarginAmount__c,KimbleOne__BaselineContractMargin__c,KimbleOne__BaselineContractRevenue__c,KimbleOne__BaselineUsage__c,KimbleOne__BusinessUnitGroup__c,KimbleOne__ContractCost__c,KimbleOne__ContractMarginAmount__c,KimbleOne__ContractMargin__c,KimbleOne__ContractRevenue__c,KimbleOne__DeliveryProgram__c,KimbleOne__DeliveryStage__c,KimbleOne__DeliveryStatus__c,KimbleOne__ExpectedCcyExpectedContractRevenue__c,KimbleOne__ExpectedCurrencyISOCode__c,KimbleOne__ExpectedEndDate__c,KimbleOne__ExpectedStartDate__c,KimbleOne__ForecastAtThisLevel__c,KimbleOne__ForecastStatus__c,KimbleOne__ForecastUsage__c,KimbleOne__InvoicingCcyExpensesInvoiceableAmount__c,KimbleOne__InvoicingCcyServicesInvoiceableAmount__c,KimbleOne__IsExpectedStartDateBeforeCloseDate__c,KimbleOne__LostReasonNarrative__c,KimbleOne__LostReason__c,KimbleOne__ProbabilityCodeEnum__c,KimbleOne__ProductGroup__c,KimbleOne__Proposal__c,KimbleOne__Reference__c,KimbleOne__RiskLevel__c,KimbleOne__SalesOpportunity__c,KimbleOne__ShortName__c,KimbleOne__WeightedContractRevenue__c,LastActivityDate,LastModifiedById,LastModifiedDate,LastReferencedDate,LastViewedDate,Name,OwnerId,Primary_Client_Contact__c,Project_Risk_Assessment_Completed_H_S__c,Project_Type__c,SystemModstamp,IsMissingFromKimble,IsDirty,SuppressFolderCreation) VALUES ("
        $sql += $(sanitise-forSqlValue -value $kimbleEngagement.attributes.url -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.Community__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.Competitive__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.CreatedById -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.CreatedDate -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.CurrencyIsoCode -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.Finance_Contact__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.Id -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.IsDeleted -dataType Boolean)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.Is_Engagement_Owner__c -dataType Boolean)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__Account__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ActualUsage__c -dataType Decimal)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__BaselineContractCost__c -dataType Decimal)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__BaselineContractMarginAmount__c -dataType Decimal)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__BaselineContractMargin__c -dataType Decimal)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__BaselineContractRevenue__c -dataType Decimal)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__BaselineUsage__c -dataType Decimal)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__BusinessUnitGroup__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ContractCost__c -dataType Decimal)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ContractMarginAmount__c -dataType Decimal)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ContractMargin__c -dataType Decimal)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ContractRevenue__c -dataType Decimal)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__DeliveryProgram__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__DeliveryStage__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__DeliveryStatus__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ExpectedCcyExpectedContractRevenue__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ExpectedCurrencyISOCode__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ExpectedEndDate__c -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ExpectedStartDate__c -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ForecastAtThisLevel__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ForecastStatus__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ForecastUsage__c -dataType Decimal)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__InvoicingCcyExpensesInvoiceableAmount__c -dataType Decimal)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__InvoicingCcyServicesInvoiceableAmount__c -dataType Decimal)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__IsExpectedStartDateBeforeCloseDate__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__LostReasonNarrative__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__LostReason__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ProbabilityCodeEnum__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ProductGroup__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__Proposal__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__Reference__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__RiskLevel__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__SalesOpportunity__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ShortName__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__WeightedContractRevenue__c -dataType Decimal)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.LastActivityDate -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.LastModifiedById -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.LastModifiedDate -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.LastReferencedDate -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.LastViewedDate -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.Name -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.OwnerId -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.Primary_Client_Contact__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.Project_Risk_Assessment_Completed_H_S__c -dataType Boolean)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.Project_Type__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleEngagement.SystemModstamp -dataType Date)
	    $sql += ",0" #IsMissing
	    $sql += ",1" #IsDirty
        $sql += ",0" #SuppressFolderCreation 
        $sql += ")"

        Write-Verbose "`t$sql"
        $result = Execute-SQLQueryOnSQLDB -query $sql -queryType nonquery -sqlServerConnection $dbConnection
        if($result -eq 1){Write-Verbose "`t`tSUCCESS!"}else{Write-Verbose "`t`tFAILURE :( - Code: $result"}
        $result
        }
    }
function add-kimbleOppToFocalPointCache{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [PSCustomObject]$kimbleOpp 

        ,[parameter(Mandatory = $true)]
        [System.Data.Common.DbConnection]$dbConnection
        )
    Write-Verbose "add-kimbleOppToFocalPointCache [$($kimbleOpp.Name)]"
    $sql = "SELECT Name,Id FROM SUS_Kimble_Accounts WHERE Id = '$($kimbleOpp.Id)'"
    Write-Verbose "`t$sql"
    $alreadyPresent = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $dbConnection
    if ($alreadyPresent){
        Write-Verbose "`tId [$($kimbleOpp.Id)] already present in Cache, not adding duplicate"
        -1 #Return "unsuccessful" as the record is in the DB, but we didn't add it and this might need investigating
        }
    else{
        if($verboseLogging){Write-Host -ForegroundColor Yellow "add-kimbleOppToFocalPointCache"}
        $sql = "INSERT INTO SUS_Kimble_Opps (attributes, SystemModstamp, Weighted_Net_Revenue__c, Proposal_Contract_Revenue__c, Project_Type__c, OwnerId, Name, LastModifiedDate, LastModifiedById, LastActivityDate, KimbleOne__WonLostReason__c, KimbleOne__WonLostNarrative__c, KimbleOne__ResponseRequiredDate__c, KimbleOne__Proposal__c, KimbleOne__OpportunityStage__c, KimbleOne__OpportunitySource__c, KimbleOne__ForecastStatus__c, KimbleOne__Description__c, KimbleOne__CloseDate__c, KimbleOne__Account__c, IsDeleted, Id, CreatedDate, CreatedById, Community__c, ANTH_SalesOpportunityStagesCount__c, ANTH_PipelineStage__c) VALUES ("
        $sql += $(sanitise-forSqlValue -value $kimbleOpp.attributes.url -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleOpp.SystemModstamp -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleOpp.Weighted_Net_Revenue__c -dataType Decimal)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleOpp.Proposal_Contract_Revenue__c -dataType Decimal)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleOpp.Project_Type__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleOpp.OwnerId -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleOpp.Name -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleOpp.LastModifiedDate -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleOpp.LastModifiedById -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleOpp.LastActivityDate -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleOpp.KimbleOne__WonLostReason__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleOpp.KimbleOne__WonLostNarrative__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleOpp.KimbleOne__ResponseRequiredDate__c -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleOpp.KimbleOne__Proposal__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleOpp.KimbleOne__OpportunityStage__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleOpp.KimbleOne__OpportunitySource__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleOpp.KimbleOne__ForecastStatus__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleOpp.KimbleOne__Description__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleOpp.KimbleOne__CloseDate__c -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleOpp.KimbleOne__Account__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleOpp.IsDeleted -dataType Boolean)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleOpp.Id -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleOpp.CreatedDate -dataType Date)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleOpp.CreatedById -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleOpp.Community__c -dataType String)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleOpp.ANTH_SalesOpportunityStagesCount__c -dataType Decimal)
        $sql += ","+$(sanitise-forSqlValue -value $kimbleOpp.ANTH_PipelineStage__c -dataType String)
        $sql += ")"
        Write-Verbose "`t$sql"
        $result = Execute-SQLQueryOnSQLDB -query $sql -queryType nonquery -sqlServerConnection $dbConnection
        if($result -eq 1){Write-Verbose "`t`tSUCCESS!"}else{Write-Verbose "`t`tFAILURE :( - Code: [$result]"}
        $result
        }
    }
function add-kimbleProposalToFocalPointCache($kimbleProp, $dbConnection, $verboseLogging){
    $sql = "SELECT Name,Id FROM SUS_Kimble_Opps WHERE Id = '$($kimbleProp.Id)'"
    $alreadyPresent = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $dbConnection
    if ($alreadyPresent){
        if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`Id already present in Cache, not adding duplicate"}
        -1 #Return "unsuccessful" as the record is in the DB, but we didn't add it and this might need investigating
        }
    else{
        if($verboseLogging){Write-Host -ForegroundColor Yellow "add-kimbleProposalToFocalPointCache"}
        $sql = "INSERT INTO SUS_Kimble_Proposals (attributes,SystemModstamp,KimbleOne__WeightedContractRevenue__c,Proposal_Contract_Revenue__c,Project_Type__c,OwnerId,Id,Name,LastModifiedDate,LastModifiedById,KimbleOne__HighLevelWeightedContractRevenue__c,KimbleOne__HighLevelContractRevenue__c,KimbleOne__HighLevelContractCost__c,KimbleOne__DetailedLevelWeightedContractRevenue__c,KimbleOne__DetailedLevelContractRevenue__c,KimbleOne__DetailedLevelContractCost__c,KimbleOne__ContractRevenue__c,KimbleOne__ContractMargin__c,KimbleOne__ContractMarginAmount__c,KimbleOne__ContractCost__c,KimbleOne__RiskLevel__c,KimbleOne__Reference__c,KimbleOne__Proposition__c,KimbleOne__ForecastStatus__c,KimbleOne__ForecastAtDetailedLevel__c,KimbleOne__Description__c,KimbleOne__BusinessUnit__c,KimbleOne__AcceptanceType__c,KimbleOne__AcceptanceDate__c,IsDeleted,CreatedDate,CreatedById,CurrencyIsoCode) VALUES ("
	    $sql += "'"+$kimbleProp.attributes.url+"',"
	    $sql += "'"+$(Get-Date (smartReplace $kimbleProp.SystemModstamp "+0000" "") -Format s -ErrorAction SilentlyContinue)+"',"
        if($kimbleProp.KimbleOne__WeightedContractRevenue__c -eq $null){$sql += "0,"}else{$sql += [string]$kimbleProp.KimbleOne__WeightedContractRevenue__c+ ","}
        if($kimbleProp.Proposal_Contract_Revenue__c -eq $null){$sql += "0,"}else{$sql += [string]$kimbleProp.Proposal_Contract_Revenue__c+ ","}
	    $sql += "'"+$(smartReplace $kimbleProp.Project_Type__c "'" "`'`'")+"',"
	    $sql += "'"+$(smartReplace $kimbleProp.OwnerId "'" "`'`'")+"',"
	    $sql += "'"+$(smartReplace $kimbleProp.Id "'" "`'`'")+"',"
	    $sql += "'"+$(smartReplace $kimbleProp.Name "'" "`'`'")+"',"
	    $sql += "'"+$(Get-Date (smartReplace $kimbleProp.LastModifiedDate "+0000" "") -Format s -ErrorAction SilentlyContinue)+"',"
	    $sql += "'"+$(smartReplace $kimbleProp.LastModifiedById "'" "`'`'")+"',"
        if($kimbleProp.KimbleOne__HighLevelWeightedContractRevenue__c -eq $null){$sql += "0,"}else{$sql += [string]$kimbleProp.KimbleOne__HighLevelWeightedContractRevenue__c+ ","}
        if($kimbleProp.KimbleOne__HighLevelContractRevenue__c -eq $null){$sql += "0,"}else{$sql += [string]$kimbleProp.KimbleOne__HighLevelContractRevenue__c+ ","}
        if($kimbleProp.KimbleOne__HighLevelContractCost__c -eq $null){$sql += "0,"}else{$sql += [string]$kimbleProp.KimbleOne__HighLevelContractCost__c+ ","}
        if($kimbleProp.KimbleOne__DetailedLevelWeightedContractRevenue__c -eq $null){$sql += "0,"}else{$sql += [string]$kimbleProp.KimbleOne__DetailedLevelWeightedContractRevenue__c+ ","}
        if($kimbleProp.KimbleOne__DetailedLevelContractRevenue__c -eq $null){$sql += "0,"}else{$sql += [string]$kimbleProp.KimbleOne__DetailedLevelContractRevenue__c+ ","}
        if($kimbleProp.KimbleOne__DetailedLevelContractCost__c -eq $null){$sql += "0,"}else{$sql += [string]$kimbleProp.KimbleOne__DetailedLevelContractCost__c+ ","}
        if($kimbleProp.KimbleOne__ContractRevenue__c -eq $null){$sql += "0,"}else{$sql += [string]$kimbleProp.KimbleOne__ContractRevenue__c+ ","}
        if($kimbleProp.KimbleOne__ContractMargin__c -eq $null){$sql += "0,"}else{$sql += [string]$kimbleProp.KimbleOne__ContractMargin__c+ ","}
        if($kimbleProp.KimbleOne__ContractMarginAmount__c -eq $null){$sql += "0,"}else{$sql += [string]$kimbleProp.KimbleOne__ContractMarginAmount__c+ ","}
        if($kimbleProp.KimbleOne__ContractCost__c -eq $null){$sql += "0,"}else{$sql += [string]$kimbleProp.KimbleOne__ContractCost__c+ ","}
	    $sql += "'"+$(smartReplace $kimbleProp.KimbleOne__RiskLevel__c "'" "`'`'")+"',"
	    $sql += "'"+$(smartReplace $kimbleProp.KimbleOne__Reference__c "'" "`'`'")+"',"
	    $sql += "'"+$(smartReplace $kimbleProp.KimbleOne__Proposition__c "'" "`'`'")+"',"
	    $sql += "'"+$(smartReplace $kimbleProp.KimbleOne__ForecastStatus__c "'" "`'`'")+"',"
        if($kimbleProp.KimbleOne__ForecastAtDetailedLevel__c -eq $true){$sql += "1,"}else{$sql += "0,"}
	    $sql += "'"+$(smartReplace $kimbleProp.KimbleOne__Description__c "'" "`'`'")+"',"
	    $sql += "'"+$(smartReplace $kimbleProp.KimbleOne__BusinessUnit__c "'" "`'`'")+"',"
	    $sql += "'"+$(smartReplace $kimbleProp.KimbleOne__AcceptanceType__c "'" "`'`'")+"',"
	    $sql += "'"+$(Get-Date (smartReplace $kimbleProp.KimbleOne__AcceptanceDate__c "+0000" "") -Format s -ErrorAction SilentlyContinue)+"',"
        if($kimbleProp.IsDeleted -eq $true){$sql += "1,"}else{$sql += "0,"}
	    $sql += "'"+$(Get-Date (smartReplace $kimbleProp.CreatedDate "+0000" "") -Format s -ErrorAction SilentlyContinue)+"',"
	    $sql += "'"+$(smartReplace $kimbleProp.CreatedById "'" "`'`'")+"',"
	    $sql += "'"+$(smartReplace $kimbleProp.CurrencyIsoCode "'" "`'`'")+"')"

        if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`t$sql"}
        $result = Execute-SQLQueryOnSQLDB -query $sql -queryType nonquery -sqlServerConnection $dbConnection
        if($verboseLogging){if($result -eq 1){Write-Host -ForegroundColor DarkYellow "`t`tSUCCESS!"}else{Write-Host -ForegroundColor DarkYellow "`t`tFAILURE :( - Code: $result"}}
        $result
        }
    }
function Get-KimbleAuthorizationTokenWithUsernamePasswordFlowRequestBody($client_id, $client_secret, $user_name, $pass_word, $security_token){
    Add-Type -AssemblyName System.Web
    $user_name = [System.Web.HttpUtility]::UrlEncode($user_name)
    $pass_word = [System.Web.HttpUtility]::UrlEncode($pass_word)
    #$requestBody = "grant_type=$grantType"
    $requestBody = "grant_type=password"
    $requestBody += "&client_id=$client_id"
    $requestBody += "&client_secret=$client_secret"
    $requestBody += "&username=$user_name"
    $requestBody += "&password=$pass_word$security_token"
    $requestBody += "&Content-Type=application/x-www-form-urlencoded"
    $requestBody
    #Write-Host "Body:" $requestBody

    #Invoke-RestMethod -Method Post -Uri $callbackUri -Body $requestBody
    #try{Invoke-RestMethod -Method Post -Uri $callbackUri -Body $requestBody} catch {Failure}
    }
function get-kimbleHeaders($clientId,$clientSecret,$username,$password,$securityToken,$connectToLiveContext,$verboseLogging){
    <######################################## The old way of doing it
    #Don't change these unless the Kimble account or App changes
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $callbackUri = "https://login.salesforce.com/services/oauth2/token" #"https://test.salesforce.com/services/oauth2/token"
    $grantType = "password"
    $myInstance = "https://eu5.salesforce.com"
    $queryUri = "$myInstance/services/data/v39.0/query/?q="
    $querySuffixStub = " -H `"Authorization: Bearer "
    $kimbleLogin = Import-Csv "$env:USERPROFILE\Desktop\KimbleAdmin.txt"
    $clientId = $kimbleLogin.clientId
    $clientSecret = $kimbleLogin.clientSecret
    $username = $kimbleLogin.username
    $password = $kimbleLogin.password
    $securityToken = $kimbleLogin.securityToken

    $oAuthReqBody = Get-KimbleAuthorizationTokenWithUsernamePasswordFlowRequestBody -client_id $clientId -client_secret $clientSecret -user_name $username -pass_word $password -security_token $securityToken
    try{$kimbleAccessToken=Invoke-RestMethod -Method Post -Uri $callbackUri -Body $oAuthReqBody} catch {Failure}
    $kimbleRestHeaders = @{Authorization = "Bearer " + $kimbleAccessToken.access_token}
    ########################################>
    if($verboseLogging){Write-Host -ForegroundColor Yellow "get-kimbleHeaders"}
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $oAuthReqBody = Get-KimbleAuthorizationTokenWithUsernamePasswordFlowRequestBody -client_id $clientId -client_secret $clientSecret -user_name $username -pass_word $password -security_token $securityToken
    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`t`$oAuthReqBody = $oAuthReqBody"}
    if($connectToLiveContext){$callbackUri = "https://login.salesforce.com/services/oauth2/token"}
    else{$callbackUri = "https://test.salesforce.com/services/oauth2/token"}
    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`t`$callbackUri = $callbackUri"}
    try{
        $kimbleAccessToken=Invoke-RestMethod -Method Post -Uri $callbackUri -Body $oAuthReqBody
        } 
    catch {
        Failure
        }
    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`t`$kimbleAccessToken = $kimbleAccessToken"}
    $kimbleRestHeaders = @{Authorization = "Bearer " + $kimbleAccessToken.access_token}
    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`t`$kimbleRestHeaders = $kimbleRestHeaders"}
    $kimbleRestHeaders
    }
function get-kimbleInstance(){
    "https://eu5.salesforce.com"
    }
function get-kimbleQueryUri($myInstance){
    if(!($myInstance)){$myInstance = get-kimbleInstance}
    $queryUri = "$myInstance/services/data/v39.0/query/?q="
    $queryUri
    }
function Failure {
    $global:helpme = $body
    $global:helpmoref = $moref
    $global:result = $_.Exception.Response.GetResponseStream()
    $global:reader = New-Object System.IO.StreamReader($global:result)
    $global:responseBody = $global:reader.ReadToEnd();
    Write-Host -BackgroundColor:Black -ForegroundColor:Red "Status: A system exception was caught."
    Write-Host -BackgroundColor:Black -ForegroundColor:Red $global:responsebody
    Write-Host -BackgroundColor:Black -ForegroundColor:Red "The request body has been saved to `$global:helpme"
    break
    }
function Get-KimbleSoqlDataset($queryUri, $soqlQuery, $restHeaders, $myInstance){
    if(!($myInstance)){$myInstance = get-kimbleInstance}
    $soqlQuery = $soqlQuery.Replace(" ","+")
    $lastIndex = 0
    $nextIndex = 0
    do{
        $lastIndex = $nextIndex
        if($lastIndex -eq 0){
            #Write-Host -ForegroundColor Magenta $partialDataSet.totalSize
            Write-Verbose "`$partialDataSet = Invoke-RestMethod -Uri [$queryUri$soqlQuery] -Headers [$(stringify-hashTable $restHeaders)]"
            $partialDataSet = Invoke-RestMethod -Uri $queryUri+$soqlQuery -Headers $restHeaders
            Write-Host -ForegroundColor Cyan $partialDataSet.totalSize
            $fullDataSet = New-Object object[] $partialDataSet.totalSize
            }
            else{$partialDataSet = Invoke-RestMethod -Uri $myInstance$($partialDataSet.nextRecordsUrl) -Headers $restHeaders}
        try{[int]$nextIndex = $partialDataSet.nextRecordsUrl.Split("-")[1]}catch{$nextIndex = $partialDataSet.totalSize-1}
        $j=0
        for($i=$lastIndex;$i -le $nextIndex;$i++){
            $fullDataSet[$i] = $partialDataSet.records[$j]
            $j++
            if($i%100 -eq 0){Write-Host -ForegroundColor DarkMagenta $i $j}
            }
        Write-Host -ForegroundColor Yellow $partialDataSet.nextRecordsUrl
        }
    while($partialDataSet.nextRecordsUrl -ne $null)
    $fullDataSet
    }
function get-allFocalPointCachedKimbleAccounts($dbConnection, $pWhereStatement, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Yellow "get-allFocalPointCachedKimbleAccounts"}
    $sql = "SELECT attributes, Website, Type, SystemModstamp, Phone, ParentId, OwnerId, Name, LastModifiedDate, LastModifiedById, KimbleOne__IsCustomer__c, KimbleOne__BusinessUnit__c, Is_Partner__c, Is_Competitor__c, IsDeleted, Id, CreatedDate, CreatedById, Client_Sector__c, BillingStreet, BillingState, BillingPostalCode, BillingCountry, BillingCity, DimCode, DimCode_Supplier, IsMissingFromKimble, IsMisclassified, IsDirty, isOrphaned, DocumentLibraryGuid, Description, PreviousDescription, PreviousName FROM SUS_Kimble_Accounts $pWhereStatement"
    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`t`$query = $sql"}
    $results = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $dbConnection
    $results
    }
function get-allFocalPointCachedKimbleContacts($dbConnection, $pWhereStatement, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Yellow "get-allFocalPointCachedKimbleContacts"}
    $sql = "SELECT AccountId, Anthesis_Events__c, AssistantName, AssistantPhone, attributes, Birthdate, Cleanup__c, Client_Type__c, CreatedById, CreatedDate, CurrencyIsoCode, Department, Description, Email, EmailBouncedDate, EmailBouncedReason, Fax, FirstName, Gender__c, Have_you_completed_this_section__c, HomePhone, Id, IsDeleted, IsEmailBounced, Jigsaw, JigsawContactId, Key_areas_of_interest__c, LastActivityDate, LastCURequestDate, LastCUUpdateDate, LastModifiedById, LastModifiedDate, LastName, LastReferencedDate, LastViewedDate, Lead_Source_Detail__c, LeadSource, Linked_In__c, MailingAddress, MailingCity, MailingCountry, MailingGeocodeAccuracy, MailingLatitude, MailingLongitude, MailingPostalCode, MailingState, MailingStreet, MasterRecordId, Met_At__c, MobilePhone, Name, Nee__c, Newsletters_Campaigns__c, Nickname__c, No_Show_Event_Attendees__c, Other_Email__c, OtherAddress, OtherCity, OtherCountry, OtherGeocodeAccuracy, OtherLatitude, OtherLongitude, OtherPhone, OtherPostalCode, OtherState, OtherStreet, OwnerId, Phone, PhotoUrl, Previous_Company__c, Region__c, ReportsToId, Role_Responsibilities__c, Salutation, Secondary_contact_owner__c, Skype__c, SystemModstamp, Title, Twitter__c, Unsubscribe_Newsletter_Campaigns__c FROM SUS_Kimble_Contacts $pWhereStatement"
    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`t`$query = $sql"}
    $results = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $dbConnection
    $results
    }
function get-allFocalPointCachedKimbleEngagements{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [System.Data.Common.DbConnection]$dbConnection

        ,[parameter(Mandatory = $false)]
        [string]$pWhereStatement
        )
    Write-Verbose "get-allFocalPointCachedKimbleEngagements"
    $sql = "SELECT attributes, Community__c, Competitive__c, CreatedById, CreatedDate, CurrencyIsoCode, Finance_Contact__c, Id, IsDeleted, Is_Engagement_Owner__c, KimbleOne__Account__c, KimbleOne__ActualUsage__c, KimbleOne__BaselineContractCost__c, KimbleOne__BaselineContractMarginAmount__c, KimbleOne__BaselineContractMargin__c, KimbleOne__BaselineContractRevenue__c, KimbleOne__BaselineUsage__c, KimbleOne__BusinessUnitGroup__c, KimbleOne__ContractCost__c, KimbleOne__ContractMarginAmount__c, KimbleOne__ContractMargin__c, KimbleOne__ContractRevenue__c, KimbleOne__DeliveryProgram__c, KimbleOne__DeliveryStage__c, KimbleOne__DeliveryStatus__c, KimbleOne__ExpectedCcyExpectedContractRevenue__c, KimbleOne__ExpectedCurrencyISOCode__c, KimbleOne__ExpectedEndDate__c, KimbleOne__ExpectedStartDate__c, KimbleOne__ForecastAtThisLevel__c, KimbleOne__ForecastStatus__c, KimbleOne__ForecastUsage__c, KimbleOne__InvoicingCcyExpensesInvoiceableAmount__c, KimbleOne__InvoicingCcyServicesInvoiceableAmount__c, KimbleOne__IsExpectedStartDateBeforeCloseDate__c, KimbleOne__LostReasonNarrative__c, KimbleOne__LostReason__c, KimbleOne__ProbabilityCodeEnum__c, KimbleOne__ProductGroup__c, KimbleOne__Proposal__c, KimbleOne__Reference__c, KimbleOne__RiskLevel__c, KimbleOne__SalesOpportunity__c, KimbleOne__ShortName__c, KimbleOne__WeightedContractRevenue__c, LastActivityDate, LastModifiedById, LastModifiedDate, LastReferencedDate, LastViewedDate, Name, OwnerId, Primary_Client_Contact__c, Project_Risk_Assessment_Completed_H_S__c, Project_Type__c, SystemModstamp, IsMissingFromKimble, IsDirty, FolderGuid, SuppressFolderCreation, PreviousKimbleClientId, PreviousName FROM SUS_Kimble_Engagements $pWhereStatement"
    Write-Verbose "`t`$query = $sql"
    $results = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $dbConnection
    $results
    }
function get-allFocalPointCachedKimbleOpps($dbConnection, $pWhereStatement, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Yellow "get-allFocalPointCachedKimbleOpps"}
    $sql = "SELECT attributes, SystemModstamp, Weighted_Net_Revenue__c, Proposal_Contract_Revenue__c, Project_Type__c, OwnerId, Name, LastModifiedDate, LastModifiedById, LastActivityDate, KimbleOne__WonLostReason__c, KimbleOne__WonLostNarrative__c, KimbleOne__ResponseRequiredDate__c, KimbleOne__Proposal__c, KimbleOne__OpportunityStage__c, KimbleOne__OpportunitySource__c, KimbleOne__ForecastStatus__c, KimbleOne__Description__c, KimbleOne__CloseDate__c, KimbleOne__Account__c, IsDeleted, Id, CreatedDate, CreatedById, Community__c, ANTH_SalesOpportunityStagesCount__c, ANTH_PipelineStage__c FROM SUS_Kimble_Opps $pWhereStatement"
    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`t`$query = $sql"}
    $results = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $dbConnection
    $results
    }
function get-allFocalPointCachedKimbleProps($dbConnection, $pWhereStatement, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Yellow "get-allFocalPointCachedKimbleProps"}
    $sql = "SELECT attributes, SystemModstamp, KimbleOne__WeightedContractRevenue__c, Proposal_Contract_Revenue__c, Project_Type__c, OwnerId, Id, Name, LastModifiedDate, LastModifiedById, KimbleOne__HighLevelWeightedContractRevenue__c, KimbleOne__HighLevelContractRevenue__c, KimbleOne__HighLevelContractCost__c, KimbleOne__DetailedLevelWeightedContractRevenue__c, KimbleOne__DetailedLevelContractRevenue__c, KimbleOne__DetailedLevelContractCost__c, KimbleOne__ContractRevenue__c, KimbleOne__ContractMargin__c, KimbleOne__ContractMarginAmount__c, KimbleOne__ContractCost__c, KimbleOne__RiskLevel__c, KimbleOne__Reference__c, KimbleOne__Proposition__c, KimbleOne__ForecastStatus__c, KimbleOne__ForecastAtDetailedLevel__c, KimbleOne__Description__c, KimbleOne__BusinessUnit__c, KimbleOne__AcceptanceType__c, KimbleOne__CloseDate__c, KimbleOne__AcceptanceDate__c, IsDeleted, CreatedDate, CreatedById, CurrencyIsoCode  FROM SUSTAIN_LIVE.dbo.SUS_Kimble_Proposals $pWhereStatement"
    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`t`$query = $sql"}
    $results = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $dbConnection
    $results
    }
function get-allKimbleAccounts($pQueryUri, $pRestHeaders, $pWhereStatement, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Yellow "get-allKimbleAccounts"}
    if($pWhereStatement){if($pWhereStatement.Split(" ")[0] -notmatch "WHERE"){$whereStatement = " WHERE "+$pWhereStatement}}
    else{$whereStatement = $pWhereStatement}
    $query = "Select a.Sync_to_FocalPoint__c, a.X2nd_Account_Owner__c, a.Website, a.Walmart_Supplier_Account__c, a.Walmart_Sub_Category__c, a.Walmart_Category_New__c, a.Type_of_Product__c, a.Type, a.SystemModstamp, a.Sym_Group__c, a.Supply_Since_Last_Update__c, a.Supplier_Sector__c, a.SicDesc, a.ShippingStreet, a.ShippingState, a.ShippingPostalCode, a.ShippingLongitude, a.ShippingLatitude, a.ShippingGeocodeAccuracy, a.ShippingCountry, a.ShippingCity, a.ShippingAddress, a.Received_Sust_Index_from_WM__c, a.Purchase_Type__c, a.PhotoUrl, a.Phone, a.Partner_Sector__c, a.ParentId, a.OwnerId, a.NumberOfEmployees, a.Name, a.Membership__c, a.MasterRecordId, a.LastViewedDate, a.LastReferencedDate, a.LastModifiedDate, a.LastModifiedById, a.LastActivityDate, a.KimbleOne__TaxCode__c, a.KimbleOne__TaxCodeReference__c, a.KimbleOne__PurchaseOrderRule__c, a.KimbleOne__Locale__c, a.KimbleOne__IsSupplier__c, a.KimbleOne__IsCustomer__c, a.KimbleOne__InvoicingCurrencyIsoCode__c, a.KimbleOne__InvoiceTemplate__c, a.KimbleOne__InvoicePaymentTermDays__c, a.KimbleOne__InvoiceFormat__c, a.KimbleOne__Code__c, a.KimbleOne__BusinessUnit__c, a.KimbleOne__BusinessUnitTradingEntity__c, a.KimbleOne__BillingParentAccount__c, a.KimbleOne__BillingContact__c, a.KimbleOne__AllowPartItemInvoicing__c, a.Key_Account__c, a.JigsawCompanyId, a.Jigsaw, a.Is_Partner__c, a.Is_Competitor__c, a.IsDeleted, a.Industry, a.Id, a.Have_you_completed_this_section__c, a.Generic_email_address__c, a.Fax, a.Description, a.D_U_N_S_Number__c, a.CurrencyIsoCode, a.CreatedDate, a.CreatedById, a.Client_Sector__c, a.BillingStreet, a.BillingState, a.BillingPostalCode, a.BillingLongitude, a.BillingLatitude, a.BillingGeocodeAccuracy, a.BillingCountry, a.BillingCity, a.BillingAddress, a.AnnualRevenue, a.Account_Team__c, a.Account_Director__c, a.AccountSource From Account a $pWhereStatement"
    #$query = "Select a.s2cor__VAT_Registration_Number__c, a.s2cor__Trading_Status__c, a.s2cor__Total_Pending__c, a.s2cor__Total_Overdue__c, a.s2cor__Total_Due_Today__c, a.s2cor__Total_Due_Soon__c, a.s2cor__Sage_UID__c, a.s2cor__Registration_Number_Type__c, a.s2cor__Registration_Number_Type_Code__c, a.s2cor__Record_T5018__c, a.s2cor__Record_T4A__c, a.s2cor__Record_1099__c, a.s2cor__Previous_Name__c, a.s2cor__Next_Key_Date__c, a.s2cor__Legal_Type__c, a.s2cor__Is_Current_User__c, a.s2cor__Is_Active__c, a.s2cor__HB_BC_CompanyId__c, a.s2cor__Gift_Aid__c, a.s2cor__Gift_Aid_Start_Date__c, a.s2cor__Email_address__c, a.s2cor__EU_Country_Code__c, a.s2cor__Date_of_Incorporation__c, a.s2cor__Country_Code__c, a.X2nd_Account_Owner__c, a.Website, a.Walmart_Supplier_Account__c, a.Walmart_Sub_Category__c, a.Walmart_Category_New__c, a.Type_of_Product__c, a.Type, a.SystemModstamp, a.Sync_to_FocalPoint__c, a.Sym_Group__c, a.Supply_Since_Last_Update__c, a.Supplier_Sector__c, a.SicDesc, a.ShippingStreet, a.ShippingState, a.ShippingPostalCode, a.ShippingLongitude, a.ShippingLatitude, a.ShippingGeocodeAccuracy, a.ShippingCountry, a.ShippingCity, a.ShippingAddress, a.SFSSDupeCatcher__Override_DupeCatcher__c, a.Received_Sust_Index_from_WM__c, a.Purchase_Type__c, a.PhotoUrl, a.Phone, a.Partner_Sector__c, a.ParentId, a.OwnerId, a.NumberOfEmployees, a.Name, a.Membership__c, a.MasterRecordId, a.LinkedIn_Company_Page__c, a.LastViewedDate, a.LastReferencedDate, a.LastModifiedDate, a.LastModifiedById, a.LastActivityDate, a.KimbleOne__TaxCode__c, a.KimbleOne__TaxCodeReference__c, a.KimbleOne__PurchaseOrderRule__c, a.KimbleOne__Locale__c, a.KimbleOne__IsSupplier__c, a.KimbleOne__IsCustomer__c, a.KimbleOne__InvoicingCurrencyIsoCode__c, a.KimbleOne__InvoiceTemplate__c, a.KimbleOne__InvoicePaymentTermDays__c, a.KimbleOne__InvoiceFormat__c, a.KimbleOne__Code__c, a.KimbleOne__BusinessUnit__c, a.KimbleOne__BusinessUnitTradingEntity__c, a.KimbleOne__BillingParentAccount__c, a.KimbleOne__BillingContact__c, a.KimbleOne__AllowPartItemInvoicing__c, a.Key_Account__c, a.JigsawCompanyId, a.Jigsaw, a.Is_Partner__c, a.Is_Competitor__c, a.IsDeleted, a.Industry, a.Id, a.Have_you_completed_this_section__c, a.Generic_email_address__c, a.Fax, a.Description, a.D_U_N_S_Number__c, a.CurrencyIsoCode, a.Credit_Status__c, a.Credit_Check_Date__c, a.Credit_Check_By__c, a.CreatedDate, a.CreatedById, a.Client_Sector__c, a.BillingStreet, a.BillingState, a.BillingPostalCode, a.BillingLongitude, a.BillingLatitude, a.BillingGeocodeAccuracy, a.BillingCountry, a.BillingCity, a.BillingAddress, a.AnnualRevenue, a.Account_Manager__c, a.Account_Director__c, a.AccountSource From Account a $whereStatement"

    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`t`$query = $query"}
    Get-KimbleSoqlDataset -queryUri $pQueryUri -soqlQuery $query -restHeaders $pRestHeaders -myInstance $(get-kimbleInstance)
    }
function get-allKimbleLeads($pQueryUri, $pRestHeaders, $pWhereStatement, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Yellow "get-allKimbleLeads"}
    $query = "Select k.Won_Reason__c, k.Weighted_Net_Revenue__c, k.SystemModstamp, k.Proposal_Contract_Revenue__c, k.Project_Type__c, k.OwnerId, k.Name, k.Lost_to_competitor_reason__c, k.LastViewedDate, k.LastReferencedDate, k.LastModifiedDate, k.LastModifiedById, k.LastActivityDate, k.KimbleOne__WonLostReason__c, k.KimbleOne__WonLostNarrative__c, k.KimbleOne__WeightedContractRevenue__c, k.KimbleOne__TaxCode__c, k.KimbleOne__ShortName__c, k.KimbleOne__ResponseRequiredDate__c, k.KimbleOne__Reference__c, k.KimbleOne__Proposition__c, k.KimbleOne__Proposal__c, k.KimbleOne__OpportunityStage__c, k.KimbleOne__OpportunitySource__c, k.KimbleOne__MarketingCampaign__c, k.KimbleOne__LostToCompetitor__c, k.KimbleOne__InvoicingCurrencyISOCode__c, k.KimbleOne__ForecastStatus__c, k.KimbleOne__Description__c, k.KimbleOne__CustomerAccount__c, k.KimbleOne__ContractRevenue__c, k.KimbleOne__ContractMargin__c, k.KimbleOne__ContractMarginAmount__c, k.KimbleOne__ContractCost__c, k.KimbleOne__CloseDate__c, k.KimbleOne__BusinessUnit__c, k.KimbleOne__Account__c, k.IsDeleted, k.Id, k.CurrencyIsoCode, k.CreatedDate, k.CreatedById, k.Country__c, k.Competitive__c, k.Community__c, k.ANTH_SalesOpportunityStagesCount__c, k.ANTH_PipelineStage__c From KimbleOne__SalesOpportunity__c k $pWhereStatement"
    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`t`$query = $query"}
    Get-KimbleSoqlDataset -queryUri $pQueryUri -soqlQuery $query -restHeaders $pRestHeaders
    }
function get-allKimbleEngagements($pQueryUri, $pRestHeaders, $pWhereStatement, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Yellow "get-allKimbleEngagements"}
    $query = "Select k.SystemModstamp, k.Project_Type__c, k.Primary_Client_Contact__c, k.OwnerId, k.Name, k.LastViewedDate, k.LastReferencedDate, k.LastModifiedDate, k.LastModifiedById, k.LastActivityDate, k.KimbleOne__WeightedContractRevenue__c, k.KimbleOne__ShortName__c, k.KimbleOne__SalesOpportunity__c, k.KimbleOne__RiskLevel__c, k.KimbleOne__Reference__c, k.KimbleOne__Proposal__c, k.KimbleOne__ProductGroup__c, k.KimbleOne__ProbabilityCodeEnum__c, k.KimbleOne__LostReason__c, k.KimbleOne__LostReasonNarrative__c, k.KimbleOne__IsExpectedStartDateBeforeCloseDate__c, k.KimbleOne__InvoicingCcyServicesInvoiceableAmount__c, k.KimbleOne__InvoicingCcyExpensesInvoiceableAmount__c, k.KimbleOne__ForecastUsage__c, k.KimbleOne__ForecastStatus__c, k.KimbleOne__ForecastAtThisLevel__c, k.KimbleOne__ExpectedStartDate__c, k.KimbleOne__ExpectedEndDate__c, k.KimbleOne__ExpectedCurrencyISOCode__c, k.KimbleOne__ExpectedCcyExpectedContractRevenue__c, k.KimbleOne__DeliveryStatus__c, k.KimbleOne__DeliveryStage__c, k.KimbleOne__DeliveryProgram__c, k.KimbleOne__ContractRevenue__c, k.KimbleOne__ContractMargin__c, k.KimbleOne__ContractMarginAmount__c, k.KimbleOne__ContractCost__c, k.KimbleOne__BusinessUnitGroup__c, k.KimbleOne__BaselineUsage__c, k.KimbleOne__BaselineContractRevenue__c, k.KimbleOne__BaselineContractMargin__c, k.KimbleOne__BaselineContractMarginAmount__c, k.KimbleOne__BaselineContractCost__c, k.KimbleOne__ActualUsage__c, k.KimbleOne__Account__c, k.Is_Engagement_Owner__c, k.IsDeleted, k.Id, k.Finance_Contact__c, k.CurrencyIsoCode, k.CreatedDate, k.CreatedById, k.Competitive__c, k.Community__c From KimbleOne__DeliveryGroup__c k $pWhereStatement"
    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`t`$query = $query"}
    Get-KimbleSoqlDataset -queryUri $pQueryUri -soqlQuery $query -restHeaders $pRestHeaders
    }
function get-allKimbleProposals($pQueryUri, $pRestHeaders, $pWhereStatement, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Yellow "get-allKimbleProposals"}
    $query = "Select k.SystemModstamp, k.OwnerId, k.Name, k.LastModifiedDate, k.LastModifiedById, k.KimbleOne__WeightedContractRevenue__c, k.KimbleOne__ShortName__c, k.KimbleOne__SalesOpportunity__c, k.KimbleOne__RiskLevel__c, k.KimbleOne__Reference__c, k.KimbleOne__Proposition__c, k.KimbleOne__HighLevelWeightedContractRevenue__c, k.KimbleOne__HighLevelContractRevenue__c, k.KimbleOne__HighLevelContractCost__c, k.KimbleOne__ForecastStatus__c, k.KimbleOne__ForecastAtDetailedLevel__c, k.KimbleOne__DetailedLevelWeightedContractRevenue__c, k.KimbleOne__DetailedLevelContractRevenue__c, k.KimbleOne__DetailedLevelContractCost__c, k.KimbleOne__Description__c, k.KimbleOne__ContractRevenue__c, k.KimbleOne__ContractMargin__c, k.KimbleOne__ContractMarginAmount__c, k.KimbleOne__ContractCost__c, k.KimbleOne__BusinessUnit__c, k.KimbleOne__Account__c, k.KimbleOne__AcceptanceType__c, k.KimbleOne__AcceptanceDate__c, k.IsDeleted, k.Id, k.CurrencyIsoCode, k.CreatedDate, k.CreatedById From KimbleOne__Proposal__c k $pWhereStatement"
    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`t`$query = $query"}
    Get-KimbleSoqlDataset -queryUri $pQueryUri -soqlQuery $query -restHeaders $pRestHeaders
    }
function get-allKimbleContacts($pQueryUri, $pRestHeaders, $pWhereStatement, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Yellow "get-allKimbleContacts"}
    $query = "Select c.Unsubscribe_Newsletter_Campaigns__c, c.Twitter__c, c.Title, c.SystemModstamp, c.Skype__c, c.Secondary_contact_owner__c, c.Salutation, c.Role_Responsibilities__c, c.ReportsToId, c.Region__c, c.Previous_Company__c, c.PhotoUrl, c.Phone, c.OwnerId, c.Other_Email__c, c.OtherStreet, c.OtherState, c.OtherPostalCode, c.OtherPhone, c.OtherLongitude, c.OtherLatitude, c.OtherGeocodeAccuracy, c.OtherCountry, c.OtherCity, c.OtherAddress, c.No_Show_Event_Attendees__c, c.Nickname__c, c.Newsletters_Campaigns__c, c.Nee__c, c.Name, c.MobilePhone, c.Met_At__c, c.MasterRecordId, c.MailingStreet, c.MailingState, c.MailingPostalCode, c.MailingLongitude, c.MailingLatitude, c.MailingGeocodeAccuracy, c.MailingCountry, c.MailingCity, c.MailingAddress, c.Linked_In__c, c.Lead_Source_Detail__c, c.LeadSource, c.LastViewedDate, c.LastReferencedDate, c.LastName, c.LastModifiedDate, c.LastModifiedById, c.LastCUUpdateDate, c.LastCURequestDate, c.LastActivityDate, c.Key_areas_of_interest__c, c.JigsawContactId, c.Jigsaw, c.IsEmailBounced, c.IsDeleted, c.Id, c.HomePhone, c.Have_you_completed_this_section__c, c.Gender__c, c.FirstName, c.Fax, c.EmailBouncedReason, c.EmailBouncedDate, c.Email, c.Description, c.Department, c.CurrencyIsoCode, c.CreatedDate, c.CreatedById, c.Client_Type__c, c.Cleanup__c, c.Birthdate, c.AssistantPhone, c.AssistantName, c.Anthesis_Events__c, c.AccountId From Contact c $pWhereStatement"
    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`t`$query = $query"}
    Get-KimbleSoqlDataset -queryUri $pQueryUri -soqlQuery $query -restHeaders $pRestHeaders
    }
function update-kimbleAccountToFocalPointCache{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [PSCustomObject]$kimbleAccount 

        ,[parameter(Mandatory = $true)]
        [System.Data.Common.DbConnection]$dbConnection
        )
    Write-Verbose "update-kimbleAccountToFocalPointCache [$($kimbleAccount.Name)]"

    #We need to check whether Name and/or Description have changed
    $sql = "SELECT Id, Name, Description FROM SUS_Kimble_Accounts WHERE Id = '$($kimbleAccount.Id)'"
    Write-Verbose "`t$sql"
    $existingRecord = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $dbConnection
    switch ($($i = 0; $existingRecord | % {$i++})){ #This doesn't support .Count >:(
        0 {
            #Id not in SUS_Kimble_Accounts. Might as well send it for creation as fail it
            Write-Verbose "`t[$($kimbleAccount.Id)][$($kimbleAccount.Name)] is missing (cannot update). Sending for recreation instead"
            add-kimbleAccountToFocalPointCache -kimbleAccount $kimbleAccount -dbConnection $dbConnection -verboseLogging $verboseLogging
            return
            }
        1 {
            #Expected result. Allows us to compare Name & Description fields later
            Write-Verbose "`t[$($kimbleAccount.Id)][$($kimbleAccount.Name)] found. Will update."
            }
        default {
            #Id matches >1. Shouldn't happen as there is a constraint on the SQL table to prevent this
            Write-Verbose "`t[$($kimbleAccount.Id)][$($kimbleAccount.Name)] has [$i] matches in [SUS_Kimble_Accounts] - check the constraints on the table as this shouldn't be possible"
            return
            }
        }

    Write-Verbose "`tGenerating SQL UPDATE statement"
    $sql = "UPDATE SUS_Kimble_Accounts "
    #(attributes, Website, Type, SystemModstamp, Phone, ParentId, OwnerId, Name, LastModifiedDate, LastModifiedById, KimbleOne__IsCustomer__c, KimbleOne__BusinessUnit__c, Is_Partner__c, Is_Competitor__c, IsDeleted, Id, CreatedDate, CreatedById, Client_Sector__c, BillingStreet, BillingState, BillingPostalCode, BillingCountry, BillingCity, DimCode, DimCode_Supplier, IsMissingFromKimble, IsMisclassified, IsDirty, isOrphaned, DocumentLibraryGuid) VALUES ("
    $sql += "SET attributes = "+$(sanitise-forSqlValue -dataType String -value $kimbleAccount.attributes.url)
    $sql += ",Website = "+$(sanitise-forSqlValue -dataType String -value $kimbleAccount.Website)
    $sql += ",Type = "+$(sanitise-forSqlValue -dataType String -value $kimbleAccount.Type)
    $sql += ",Name = "+$(sanitise-forSqlValue -dataType String -value $kimbleAccount.Name)
    $sql += ",Description = "+$(sanitise-forSqlValue -dataType String -value $kimbleAccount.Description)
    $sql += ",SystemModstamp = "+$(sanitise-forSqlValue -dataType Date -value $kimbleAccount.SystemModstamp)
    $sql += ",Phone = "+$(sanitise-forSqlValue -dataType String -value $kimbleAccount.Phone)
    $sql += ",ParentId = "+$(sanitise-forSqlValue -dataType String -value $kimbleAccount.ParentId)
    $sql += ",OwnerId = "+$(sanitise-forSqlValue -dataType String -value $kimbleAccount.OwnerId)
    $sql += ",LastModifiedDate = "+$(sanitise-forSqlValue -dataType Date -value $kimbleAccount.LastModifiedDate)
    $sql += ",LastModifiedById = "+$(sanitise-forSqlValue -dataType String -value $kimbleAccount.LastModifiedById)
    $sql += ",KimbleOne__IsCustomer__c = "+$(sanitise-forSqlValue -dataType Boolean -value $kimbleAccount.KimbleOne__IsCustomer__c)
    $sql += ",KimbleOne__BusinessUnit__c = "+$(sanitise-forSqlValue -dataType String -value $kimbleAccount.KimbleOne__BusinessUnit__c)
    $sql += ",Is_Partner__c = "+$(sanitise-forSqlValue -dataType Boolean -value $kimbleAccount.Is_Partner__c)
    $sql += ",Is_Competitor__c = "+$(sanitise-forSqlValue -dataType Boolean -value $kimbleAccount.Is_Competitor__c)
    $sql += ",IsDeleted = "+$(sanitise-forSqlValue -dataType Boolean -value $kimbleAccount.IsDeleted)
    #$sql += "'"+$kimbleAccount.Id
    $sql += ",CreatedDate = "+$(sanitise-forSqlValue -dataType Date -value $kimbleAccount.CreatedDate)
    $sql += ",CreatedById = "+$(sanitise-forSqlValue -dataType String -value $kimbleAccount.CreatedById)
    $sql += ",Client_Sector__c = "+$(sanitise-forSqlValue -dataType String -value $kimbleAccount.Client_Sector__c)
    $sql += ",BillingStreet = "+$(sanitise-forSqlValue -dataType String -value $kimbleAccount.BillingStreet)
    $sql += ",BillingState = "+$(sanitise-forSqlValue -dataType String -value $kimbleAccount.BillingState)
    $sql += ",BillingPostalCode = "+$(sanitise-forSqlValue -dataType String -value $kimbleAccount.BillingPostalCode)
    $sql += ",BillingCountry = "+$(sanitise-forSqlValue -dataType String -value $kimbleAccount.BillingCountry)
    $sql += ",BillingCity = "+$(sanitise-forSqlValue -dataType String -value $kimbleAccount.BillingCity)
    #These Properties need extra processing
    if((sanitise-forSql $existingRecord.Name) -ne (sanitise-forSql $kimbleAccount.Name)){
        $sql += ",PreviousName = "+$(sanitise-forSqlValue -dataType String -value $existingRecord.Name) #There is also a Trigger on SUS_Kimble_Accounts to record these changes over time
        $isDefinitelyDirtyNow = $true
        Write-Verbose "`tName has changed. Definitely Dirty now."
        }
    if((sanitise-forSql $existingRecord.Description) -ne (sanitise-forSql $kimbleAccount.Description)){
        $sql += ",PreviousDescription = "+$(sanitise-forSqlValue -dataType String -value $existingRecord.Description) #There is also a Trigger on SUS_Kimble_Accounts to record these changes over time
        $isDefinitelyDirtyNow = $true
        Write-Verbose "`tDescription has changed. Definitely Dirty now."
        }
    #These Properties are not native to Kimble, so may not be present:
    if($kimbleAccount.DimCode){$sql += ",DimCode = "+$(sanitise-forSqlValue -dataType String -value $kimbleAccount.DimCode)}
    if($kimbleAccount.DimCode_Supplier){$sql += ",DimCode_Supplier = "+$(sanitise-forSqlValue -dataType String -value $kimbleAccount.DimCode_Supplier)}
    if($kimbleAccount.IsMissingFromKimble){$sql += ",IsMissingFromKimble = "+$(sanitise-forSqlValue -dataType Boolean -value $kimbleAccount.IsMissingFromKimble)}
    if($kimbleAccount.IsMisclassified){$sql += ",IsMisclassified = "+$(sanitise-forSqlValue -dataType Boolean -value $kimbleAccount.IsMisclassified)}
    if($kimbleAccount.isOrphaned){$sql += ",isOrphaned = "+$(sanitise-forSqlValue -dataType Boolean -value $kimbleAccount.isOrphaned)}
    if($isDefinitelyDirtyNow){$sql += ",IsDirty = 1"}
    elseif($kimbleAccount.IsDirty){$sql += ",IsDirty = "+$(sanitise-forSqlValue -dataType Boolean -value $kimbleAccount.IsDirty)} #Don't include comma on final add
        #else{$sql += "IsDirty = 1"} #If this isn't supplied, assume IsDirty=$true (_Don't assume this as it marks everything as IsDirty=$true when we reconcile)

    $sql += " WHERE id = '$(sanitise-forSql $kimbleAccount.Id)'"
    Write-Verbose "`t$sql"
    $result = Execute-SQLQueryOnSQLDB -query $sql -queryType nonquery -sqlServerConnection $dbConnection
    if($result -eq 1){Write-Verbose "`t`tSUCCESS!"}
    else{Write-Verbose "`t`tFAILURE :( - Code: $result"}
    $result
    }
function update-kimbleContactToFocalPointCache($kimbleContact, $dbConnection, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Yellow "update-kimbleContactToFocalPointCache"}
    $sql = "UPDATE SUS_Kimble_Contacts "
    #(AccountId,Anthesis_Events__c,AssistantName,AssistantPhone,attributes,Birthdate,Cleanup__c,Client_Type__c,CreatedById,CreatedDate,CurrencyIsoCode,Department,Description,Email,EmailBouncedDate,EmailBouncedReason,Fax,FirstName,Gender__c,Have_you_completed_this_section__c,HomePhone,Id,IsDeleted,IsEmailBounced,Jigsaw,JigsawContactId,Key_areas_of_interest__c,LastActivityDate,LastCURequestDate,LastCUUpdateDate,LastModifiedById,LastModifiedDate,LastName,LastReferencedDate,LastViewedDate,Lead_Source_Detail__c,LeadSource,Linked_In__c,MailingAddress,MailingCity,MailingCountry,MailingGeocodeAccuracy,MailingLatitude,MailingLongitude,MailingPostalCode,MailingState,MailingStreet,MasterRecordId,Met_At__c,MobilePhone,Name,Nee__c,Newsletters_Campaigns__c,Nickname__c,No_Show_Event_Attendees__c,Other_Email__c,OtherAddress,OtherCity,OtherCountry,OtherGeocodeAccuracy,OtherLatitude,OtherLongitude,OtherPhone,OtherPostalCode,OtherState,OtherStreet,OwnerId,Phone,PhotoUrl,Previous_Company__c,Region__c,ReportsToId,Role_Responsibilities__c,Salutation,Secondary_contact_owner__c,Skype__c,SystemModstamp,Title,Twitter__c,Unsubscribe_Newsletter_Campaigns__c
    $sql += "SET "
    $sql += "AccountId = '"+$kimbleContact.AccountId+"',"
    $sql += "Anthesis_Events__c = '"+$kimbleContact.Anthesis_Events__c+"',"
    $sql += "AssistantName = '"+$kimbleContact.AssistantName+"',"
    $sql += "AssistantPhone = '"+$kimbleContact.AssistantPhone+"',"
    $sql += "attributes = '"+$kimbleContact.attributes.url+"',"
    $sql += "Birthdate = '"+$(Get-Date (smartReplace $kimbleContact.Birthdate "+0000" "") -Format s -ErrorAction SilentlyContinue)+"',"
    if($kimbleContact.Cleanup__c -eq $true){$sql += "Cleanup__c = 1,"}else{$sql += "Cleanup__c = 0,"}
    $sql += "Client_Type__c = '"+$kimbleContact.Client_Type__c+"',"
    $sql += "CreatedById = '"+$kimbleContact.CreatedById+"',"
    $sql += "CreatedDate = '"+$(Get-Date (smartReplace $kimbleContact.CreatedDate "+0000" "") -Format s -ErrorAction SilentlyContinue)+"',"
    $sql += "CurrencyIsoCode = '"+$kimbleContact.CurrencyIsoCode+"',"
    $sql += "Department = '"+$kimbleContact.Department+"',"
    $sql += "Description = '"+$kimbleContact.Description+"',"
    $sql += "Email = '"+$kimbleContact.Email+"',"
    $sql += "EmailBouncedDate = '"+$(Get-Date (smartReplace $kimbleContact.EmailBouncedDate "+0000" "") -Format s -ErrorAction SilentlyContinue)+"',"
    $sql += "EmailBouncedReason = '"+$kimbleContact.EmailBouncedReason+"',"
    $sql += "Fax = '"+$kimbleContact.Fax+"',"
    $sql += "FirstName = '"+$kimbleContact.FirstName+"',"
    $sql += "Gender__c = '"+$kimbleContact.Gender__c+"',"
    $sql += "Have_you_completed_this_section__c = '"+$kimbleContact.Have_you_completed_this_section__c+"',"
    $sql += "HomePhone = '"+$kimbleContact.HomePhone+"',"
    $sql += "Id = '"+$kimbleContact.Id+"',"
    if($kimbleContact.IsDeleted -eq $true){$sql += "IsDeleted = 1,"}else{$sql += "IsDeleted = 0,"}
    if($kimbleContact.IsEmailBounced -eq $true){$sql += "IsEmailBounced = 1,"}else{$sql += "IsEmailBounced = 0,"}
    $sql += "Jigsaw = '"+$kimbleContact.Jigsaw+"',"
    $sql += "JigsawContactId = '"+$kimbleContact.JigsawContactId+"',"
    $sql += "Key_areas_of_interest__c = '"+$kimbleContact.Key_areas_of_interest__c+"',"
    $sql += "LastActivityDate = '"+$(Get-Date (smartReplace $kimbleContact.LastActivityDate "+0000" "") -Format s)+"',"
    $sql += "LastCURequestDate = '"+$(Get-Date (smartReplace $kimbleContact.LastCURequestDate "+0000" "") -Format s)+"',"
    $sql += "LastCUUpdateDate = '"+$(Get-Date (smartReplace $kimbleContact.LastCUUpdateDate "+0000" "") -Format s)+"',"
    $sql += "LastModifiedById = '"+$kimbleContact.LastModifiedById+"',"
    $sql += "LastModifiedDate = '"+$(Get-Date (smartReplace $kimbleContact.LastModifiedDate "+0000" "") -Format s)+"',"
    $sql += "LastName = '"+$kimbleContact.LastName+"',"
    $sql += "LastReferencedDate = '"+$(Get-Date (smartReplace $kimbleContact.LastReferencedDate "+0000" "") -Format s)+"',"
    $sql += "LastViewedDate = '"+$(Get-Date (smartReplace $kimbleContact.LastViewedDate "+0000" "") -Format s)+"',"
    $sql += "Lead_Source_Detail__c = '"+$kimbleContact.Lead_Source_Detail__c+"',"
    $sql += "LeadSource = '"+$kimbleContact.LeadSource+"',"
    $sql += "Linked_In__c = '"+$kimbleContact.Linked_In__c+"',"
    $sql += "MailingAddress = '"+$kimbleContact.MailingAddress+"',"
    $sql += "MailingCity = '"+$kimbleContact.MailingCity+"',"
    $sql += "MailingCountry = '"+$kimbleContact.MailingCountry+"',"
    $sql += "MailingGeocodeAccuracy = '"+$kimbleContact.MailingGeocodeAccuracy+"',"
    $sql += "MailingLatitude = '"+$kimbleContact.MailingLatitude+"',"
    $sql += "MailingLongitude = '"+$kimbleContact.MailingLongitude+"',"
    $sql += "MailingPostalCode = '"+$kimbleContact.MailingPostalCode+"',"
    $sql += "MailingState = '"+$kimbleContact.MailingState+"',"
    $sql += "MailingStreet = '"+$kimbleContact.MailingStreet+"',"
    $sql += "MasterRecordId = '"+$kimbleContact.MasterRecordId+"',"
    $sql += "Met_At__c = '"+$kimbleContact.Met_At__c+"',"
    $sql += "MobilePhone = '"+$kimbleContact.MobilePhone+"',"
    $sql += "Name = '"+$kimbleContact.Name+"',"
    $sql += "Nee__c = '"+$kimbleContact.Nee__c+"',"
    $sql += "Newsletters_Campaigns__c = '"+$kimbleContact.Newsletters_Campaigns__c+"',"
    $sql += "Nickname__c = '"+$kimbleContact.Nickname__c+"',"
    $sql += "No_Show_Event_Attendees__c = '"+$kimbleContact.No_Show_Event_Attendees__c+"',"
    $sql += "Other_Email__c = '"+$kimbleContact.Other_Email__c+"',"
    $sql += "OtherAddress = '"+$kimbleContact.OtherAddress+"',"
    $sql += "OtherCity = '"+$kimbleContact.OtherCity+"',"
    $sql += "OtherCountry = '"+$kimbleContact.OtherCountry+"',"
    $sql += "OtherGeocodeAccuracy = '"+$kimbleContact.OtherGeocodeAccuracy+"',"
    $sql += "OtherLatitude = '"+$kimbleContact.OtherLatitude+"',"
    $sql += "OtherLongitude = '"+$kimbleContact.OtherLongitude+"',"
    $sql += "OtherPhone = '"+$kimbleContact.OtherPhone+"',"
    $sql += "OtherPostalCode = '"+$kimbleContact.OtherPostalCode+"',"
    $sql += "OtherState = '"+$kimbleContact.OtherState+"',"
    $sql += "OtherStreet = '"+$kimbleContact.OtherStreet+"',"
    $sql += "OwnerId = '"+$kimbleContact.OwnerId+"',"
    $sql += "Phone = '"+$kimbleContact.Phone+"',"
    $sql += "PhotoUrl = '"+$kimbleContact.PhotoUrl+"',"
    $sql += "Previous_Company__c = '"+$kimbleContact.Previous_Company__c+"',"
    $sql += "Region__c = '"+$kimbleContact.Region__c+"',"
    $sql += "ReportsToId = '"+$kimbleContact.ReportsToId+"',"
    $sql += "Role_Responsibilities__c = '"+$kimbleContact.Role_Responsibilities__c+"',"
    $sql += "Salutation = '"+$kimbleContact.Salutation+"',"
    $sql += "Secondary_contact_owner__c = '"+$kimbleContact.Secondary_contact_owner__c+"',"
    $sql += "Skype__c = '"+$kimbleContact.Skype__c+"',"
    $sql += "SystemModstamp = '"+$(Get-Date (smartReplace $kimbleContact.SystemModstamp "+0000" "") -Format s)+"',"
    $sql += "Title = '"+$kimbleContact.Title+"',"
    $sql += "Twitter__c = '"+$kimbleContact.Twitter__c+"',"
    $sql += "Unsubscribe_Newsletter_Campaigns__c = '"+$kimbleContact.Unsubscribe_Newsletter_Campaigns__c+"' "
    $sql += "WHERE id = '$(sanitise-forSql $kimbleContact.Id)'"
    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`t$sql"}
    $result = Execute-SQLQueryOnSQLDB -query $sql -queryType nonquery -sqlServerConnection $dbConnection
    if($verboseLogging){if($result -eq 1){Write-Host -ForegroundColor DarkYellow "`t`tSUCCESS!"}else{Write-Host -ForegroundColor DarkYellow "`t`tFAILURE :( - Code: [$result]"}}
    $result
    }
function update-kimbleEngagementToFocalPointCache{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [PSCustomObject]$kimbleEngagement 

        ,[parameter(Mandatory = $true)]
        [System.Data.Common.DbConnection]$dbConnection
        )
    Write-Verbose "update-kimbleEngagementToFocalPointCache"
    #We need to check whether Name and/or KimbleOne__Account__c have changed
    $sql = "SELECT Id, Name, KimbleOne__Account__c FROM SUS_Kimble_Engagements WHERE Id = '$($kimbleEngagement.Id)'"
    Write-Verbose "`t$sql"
    $existingRecord = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $dbConnection
    switch ($($i = 0; $existingRecord | % {$i++})){ #This doesn't support .Count >:(
        0 {
            #Id not in SUS_Kimble_Engagements. Might as well send it for creation as fail it
            add-kimbleEngagementToFocalPointCache -kimbleEngagement $kimbleEngagement -dbConnection $dbConnection
            Write-Verbose "`t[$($kimbleEngagement.Id)][$($kimbleEngagement.Name)] is missing (cannot update). Sending for recreation instead"
            return
            }
        1 {
            #Expected result. Allows us to compare Name & KimbleOne__Account__c fields later
            Write-Verbose "`t[$($kimbleEngagement.Id)][$($kimbleEngagement.Name)] found. Will update."
            }
        default {
            #Id matches >1. Shouldn't happen as there is a constraint on the SQL table to prevent this
            Write-Verbose "`t[$($kimbleEngagement.Id)][$($kimbleEngagement.Name)] has [$i] matches in [SUS_Kimble_Accounts] - check the constraints on the table as this shouldn't be possible"
            return
            }
        }
    
    $sql = "UPDATE SUS_Kimble_Engagements "
    #attributes,Community__c,Competitive__c,CreatedById,CreatedDate,CurrencyIsoCode,Finance_Contact__c,Id,IsDeleted,Is_Engagement_Owner__c,KimbleOne__Account__c,KimbleOne__ActualUsage__c,KimbleOne__BaselineContractCost__c,KimbleOne__BaselineContractMarginAmount__c,KimbleOne__BaselineContractMargin__c,KimbleOne__BaselineContractRevenue__c,KimbleOne__BaselineUsage__c,KimbleOne__BusinessUnitGroup__c,KimbleOne__ContractCost__c,KimbleOne__ContractMarginAmount__c,KimbleOne__ContractMargin__c,KimbleOne__ContractRevenue__c,KimbleOne__DeliveryProgram__c,KimbleOne__DeliveryStage__c,KimbleOne__DeliveryStatus__c,KimbleOne__ExpectedCcyExpectedContractRevenue__c,KimbleOne__ExpectedCurrencyISOCode__c,KimbleOne__ExpectedEndDate__c,KimbleOne__ExpectedStartDate__c,KimbleOne__ForecastAtThisLevel__c,KimbleOne__ForecastStatus__c,KimbleOne__ForecastUsage__c,KimbleOne__InvoicingCcyExpensesInvoiceableAmount__c,KimbleOne__InvoicingCcyServicesInvoiceableAmount__c,KimbleOne__IsExpectedStartDateBeforeCloseDate__c,KimbleOne__LostReasonNarrative__c,KimbleOne__LostReason__c,KimbleOne__ProbabilityCodeEnum__c,KimbleOne__ProductGroup__c,KimbleOne__Proposal__c,KimbleOne__Reference__c,KimbleOne__RiskLevel__c,KimbleOne__SalesOpportunity__c,KimbleOne__ShortName__c,KimbleOne__WeightedContractRevenue__c,LastActivityDate,LastModifiedById,LastModifiedDate,LastReferencedDate,LastViewedDate,Name,OwnerId,Primary_Client_Contact__c,Project_Risk_Assessment_Completed_H_S__c,Project_Type__c,SystemModstamp,IsMissingFromKimble,IsDirty
	$sql += "SET attributes = "+(sanitise-forSqlValue -value $kimbleEngagement.attributes.url -dataType String)
	$sql += ",Community__c = "+(sanitise-forSqlValue -value $kimbleEngagement.Community__c -dataType String)
	$sql += ",Competitive__c = "+(sanitise-forSqlValue -value $kimbleEngagement.Competitive__c -dataType String)
	$sql += ",CreatedById = "+(sanitise-forSqlValue -value $kimbleEngagement.CreatedById -dataType String)
	$sql += ",CreatedDate = "+(sanitise-forSqlValue -value $kimbleEngagement.CreatedDate -dataType Date)
	$sql += ",CurrencyIsoCode = "+(sanitise-forSqlValue -value $kimbleEngagement.CurrencyIsoCode -dataType String)
	$sql += ",Finance_Contact__c = "+(sanitise-forSqlValue -value $kimbleEngagement.Finance_Contact__c -dataType String)
	#$sql += ",Id = "+(sanitise-forSqlValue -value $kimbleEngagement.Id -dataType String)
	$sql += ",IsDeleted = "+(sanitise-forSqlValue -value $kimbleEngagement.IsDeleted -dataType Boolean)
	$sql += ",Is_Engagement_Owner__c = "+(sanitise-forSqlValue -value $kimbleEngagement.Is_Engagement_Owner__c -dataType Boolean)
	$sql += ",KimbleOne__Account__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__Account__c -dataType String)
	$sql += ",KimbleOne__ActualUsage__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ActualUsage__c -dataType Decimal)
	$sql += ",KimbleOne__BaselineContractCost__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__BaselineContractCost__c -dataType Decimal)
	$sql += ",KimbleOne__BaselineContractMarginAmount__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__BaselineContractMarginAmount__c -dataType Decimal)
	$sql += ",KimbleOne__BaselineContractMargin__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__BaselineContractMargin__c -dataType Decimal)
	$sql += ",KimbleOne__BaselineContractRevenue__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__BaselineContractRevenue__c -dataType Decimal)
	$sql += ",KimbleOne__BaselineUsage__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__BaselineUsage__c -dataType Decimal)
	$sql += ",KimbleOne__BusinessUnitGroup__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__BusinessUnitGroup__c -dataType String)
	$sql += ",KimbleOne__ContractCost__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ContractCost__c -dataType Decimal)
	$sql += ",KimbleOne__ContractMarginAmount__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ContractMarginAmount__c -dataType Decimal)
	$sql += ",KimbleOne__ContractMargin__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ContractMargin__c -dataType Decimal)
	$sql += ",KimbleOne__ContractRevenue__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ContractRevenue__c -dataType Decimal)
	$sql += ",KimbleOne__DeliveryProgram__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__DeliveryProgram__c -dataType String)
	$sql += ",KimbleOne__DeliveryStage__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__DeliveryStage__c -dataType String)
	$sql += ",KimbleOne__DeliveryStatus__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__DeliveryStatus__c -dataType String)
	$sql += ",KimbleOne__ExpectedCcyExpectedContractRevenue__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ExpectedCcyExpectedContractRevenue__c -dataType String)
	$sql += ",KimbleOne__ExpectedCurrencyISOCode__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ExpectedCurrencyISOCode__c -dataType String)
	$sql += ",KimbleOne__ExpectedEndDate__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ExpectedEndDate__c -dataType Date)
	$sql += ",KimbleOne__ExpectedStartDate__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ExpectedStartDate__c -dataType Date)
	$sql += ",KimbleOne__ForecastAtThisLevel__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ForecastAtThisLevel__c -dataType String)
	$sql += ",KimbleOne__ForecastStatus__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ForecastStatus__c -dataType String)
	$sql += ",KimbleOne__ForecastUsage__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ForecastUsage__c -dataType Decimal)
	$sql += ",KimbleOne__InvoicingCcyExpensesInvoiceableAmount__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__InvoicingCcyExpensesInvoiceableAmount__c -dataType Decimal)
	$sql += ",KimbleOne__InvoicingCcyServicesInvoiceableAmount__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__InvoicingCcyServicesInvoiceableAmount__c -dataType Decimal)
	$sql += ",KimbleOne__IsExpectedStartDateBeforeCloseDate__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__IsExpectedStartDateBeforeCloseDate__c -dataType String)
	$sql += ",KimbleOne__LostReasonNarrative__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__LostReasonNarrative__c -dataType String)
	$sql += ",KimbleOne__LostReason__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__LostReason__c -dataType String)
	$sql += ",KimbleOne__ProbabilityCodeEnum__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ProbabilityCodeEnum__c -dataType String)
	$sql += ",KimbleOne__ProductGroup__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ProductGroup__c -dataType String)
	$sql += ",KimbleOne__Proposal__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__Proposal__c -dataType String)
	$sql += ",KimbleOne__Reference__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__Reference__c -dataType String)
	$sql += ",KimbleOne__RiskLevel__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__RiskLevel__c -dataType String)
	$sql += ",KimbleOne__SalesOpportunity__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__SalesOpportunity__c -dataType String)
	$sql += ",KimbleOne__ShortName__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__ShortName__c -dataType String)
	$sql += ",KimbleOne__WeightedContractRevenue__c = "+(sanitise-forSqlValue -value $kimbleEngagement.KimbleOne__WeightedContractRevenue__c -dataType Decimal)
	$sql += ",LastActivityDate = "+(sanitise-forSqlValue -value $kimbleEngagement.LastActivityDate -dataType Date)
	$sql += ",LastModifiedById = "+(sanitise-forSqlValue -value $kimbleEngagement.LastModifiedById -dataType String)
	$sql += ",LastModifiedDate = "+(sanitise-forSqlValue -value $kimbleEngagement.LastModifiedDate -dataType Date)
	$sql += ",LastReferencedDate = "+(sanitise-forSqlValue -value $kimbleEngagement.LastReferencedDate -dataType Date)
	$sql += ",LastViewedDate = "+(sanitise-forSqlValue -value $kimbleEngagement.LastViewedDate -dataType Date)
	$sql += ",Name = "+(sanitise-forSqlValue -value $kimbleEngagement.Name -dataType String)
	$sql += ",OwnerId = "+(sanitise-forSqlValue -value $kimbleEngagement.OwnerId -dataType String)
	$sql += ",Primary_Client_Contact__c = "+(sanitise-forSqlValue -value $kimbleEngagement.Primary_Client_Contact__c -dataType String)
	$sql += ",Project_Risk_Assessment_Completed_H_S__c = "+(sanitise-forSqlValue -value $kimbleEngagement.Project_Risk_Assessment_Completed_H_S__c -dataType Boolean)
	$sql += ",Project_Type__c = "+(sanitise-forSqlValue -value $kimbleEngagement.Project_Type__c -dataType String)
	$sql += ",SystemModstamp = "+(sanitise-forSqlValue -value $kimbleEngagement.SystemModstamp -dataType Date)
    #These Properties need extra processing
    if((sanitise-forSqlValue -value $existingRecord.Name -dataType String) -ne (sanitise-forSqlValue -value  $kimbleEngagement.Name -dataType String)){
        $sql += ",PreviousName = "+$(sanitise-forSqlValue -dataType String -value $existingRecord.Name) #There is also a Trigger on SUS_Kimble_Engagements to record these changes over time
        $isDefinitelyDirtyNow = $true
        $dirtyReason += " [Name Change]"
        #Write-Debug "Name Change - Old:[$(sanitise-forSql $existingRecord.Name)] New:[$(sanitise-forSql $kimbleEngagement.Name)]"
        }
    if((sanitise-forSqlValue -value $existingRecord.KimbleOne__Account__c -dataType String) -ne (sanitise-forSqlValue -value  $kimbleEngagement.KimbleOne__Account__c -dataType String)){
        $sql += ",PreviousKimbleClientId = "+$(sanitise-forSqlValue -dataType String -value $existingRecord.KimbleOne__Account__c) #There is also a Trigger on SUS_Kimble_Engagements to record these changes over time
        $isDefinitelyDirtyNow = $true
        $dirtyReason += " [Client change]"
        #Write-Debug "Client Change - Old:[$(sanitise-forSql $existingRecord.KimbleOne__Account__c)] New:[$(sanitise-forSql $kimbleEngagement.KimbleOne__Account__c)]"
        }
    #These Properties are not native to Kimble, so may not be present:
    if($kimbleEngagement.SuppressFolderCreation){$sql += ",SuppressFolderCreation = "+$(sanitise-forSqlValue -dataType Boolean -value $kimbleEngagement.SuppressFolderCreation)}
    if($kimbleEngagement.IsMissingFromKimble){$sql += ",IsMissingFromKimble = "+$(sanitise-forSqlValue -dataType Boolean -value $kimbleEngagement.IsMissingFromKimble)}
    if($kimbleEngagement.FolderGuid){$sql += ",FolderGuid = "+$(sanitise-forSqlValue -dataType Guid -value $kimbleEngagement.FolderGuid)}
    if($kimbleEngagement.SuppressFolderCreation){$sql += ",SuppressFolderCreation = "+$(sanitise-forSqlValue -dataType Boolean -value $kimbleEngagement.SuppressFolderCreation)}
    if($kimbleEngagement.PreviousName){$sql += ",PreviousName = "+$(sanitise-forSqlValue -dataType String -value $kimbleEngagement.PreviousName)}
    if($kimbleEngagement.PreviousKimbleClientId){$sql += ",PreviousKimbleClientId = "+$(sanitise-forSqlValue -dataType String -value $PreviousKimbleClientId.isOrphaned)}
    if($kimbleEngagement.IsOrphaned){$sql += ",IsOrphaned = "+$(sanitise-forSqlValue -dataType Boolean -value $kimbleEngagement.IsOrphaned)}
    if($isDefinitelyDirtyNow){$sql += ",IsDirty = 1"} 
    elseif($kimbleEngagement.IsDirty){$sql += ",IsDirty = "+$(sanitise-forSqlValue -dataType Boolean -value $kimbleEngagement.IsDirty)}
    if($kimbleEngagement.IsDirty -eq $true){$dirtyReason += " [Already flagged as IsDirty]"}
        #else{$sql += "IsDirty = 1"} #If this isn't supplied, _don't_ assume IsDirty=$true as it marks everything as IsDirty=$true when we reconcile)

    $sql += " WHERE id = '$(sanitise-forSql $kimbleEngagement.Id)'"

    if(![string]::IsNullOrWhiteSpace($dirtyReason)){Write-Debug "update-kimbleEngagementToFocalPointCache flagged [$($kimbleEngagement.Name)] as `$IsDirty=`$true because $dirtyReason"}
    Write-Verbose "`t$sql"
    if($dirtyReason){Write-Debug "Engagement isDirty because: [$($dirtyReason)]"}
    $result = Execute-SQLQueryOnSQLDB -query $sql -queryType nonquery -sqlServerConnection $dbConnection
    if($result -eq 1){Write-Verbose "`t`tSUCCESS!"}else{Write-Verbose "`t`tFAILURE :( - Code: $result"}
#    Write-Debug "`$result = [$result]"
    $result
    }
function update-kimbleOppToFocalPointCache{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [PSCustomObject]$kimbleOpp 

        ,[parameter(Mandatory = $true)]
        [System.Data.Common.DbConnection]$dbConnection
        )
    Write-Verbose "update-kimbleOppToFocalPointCache [$($kimbleOpp.Name)]"
    #We don't need to check this for changes as we don't monitor them for anything. The data just needs to be up to date.

    $sql = "UPDATE SUS_Kimble_Opps "
    #(attributes, SystemModstamp, Weighted_Net_Revenue__c, Proposal_Contract_Revenue__c, Project_Type__c, OwnerId, Name, LastModifiedDate, LastModifiedById, LastActivityDate, KimbleOne__WonLostReason__c, KimbleOne__WonLostNarrative__c, KimbleOne__ResponseRequiredDate__c, KimbleOne__Proposal__c, KimbleOne__OpportunityStage__c, KimbleOne__OpportunitySource__c, KimbleOne__ForecastStatus__c, KimbleOne__Description__c, KimbleOne__CloseDate__c, KimbleOne__Account__c, IsDeleted, CreatedDate, CreatedById, Community__c, ANTH_SalesOpportunityStagesCount__c, ANTH_PipelineStage__c) VALUES ("
    $sql += "SET attributes = "+$(sanitise-forSqlValue -dataType String -value $kimbleOpp.attributes.url)
    $sql += ",SystemModstamp = "+$(sanitise-forSqlValue -dataType Date -value $kimbleOpp.SystemModstamp)
    $sql += ",Weighted_Net_Revenue__c = "+$(sanitise-forSqlValue -dataType Decimal -value $kimbleOpp.Weighted_Net_Revenue__c)
    $sql += ",Proposal_Contract_Revenue__c = "+$(sanitise-forSqlValue -dataType Decimal -value $kimbleOpp.Proposal_Contract_Revenue__c)
    $sql += ",Project_Type__c = "+$(sanitise-forSqlValue -dataType String -value $kimbleOpp.Project_Type__c)
    $sql += ",OwnerId = "+$(sanitise-forSqlValue -dataType String -value $kimbleOpp.OwnerId)
    $sql += ",Name = "+$(sanitise-forSqlValue -dataType String -value $kimbleOpp.Name)
    $sql += ",LastModifiedDate = "+$(sanitise-forSqlValue -dataType Date -value $kimbleOpp.LastModifiedDate)
    $sql += ",LastModifiedById = "+$(sanitise-forSqlValue -dataType String -value $kimbleOpp.LastModifiedById)
    $sql += ",LastActivityDate = "+$(sanitise-forSqlValue -dataType Date -value $kimbleOpp.LastActivityDate)
    $sql += ",KimbleOne__WonLostReason__c = "+$(sanitise-forSqlValue -dataType String -value $kimbleOpp.KimbleOne__WonLostReason__c)
    $sql += ",KimbleOne__WonLostNarrative__c = "+$(sanitise-forSqlValue -dataType String -value $kimbleOpp.KimbleOne__WonLostNarrative__c)
    $sql += ",KimbleOne__ResponseRequiredDate__c = "+$(sanitise-forSqlValue -dataType Date -value $kimbleOpp.KimbleOne__ResponseRequiredDate__c)
    $sql += ",KimbleOne__Proposal__c = "+$(sanitise-forSqlValue -dataType String -value $kimbleOpp.KimbleOne__Proposal__c)
    $sql += ",KimbleOne__OpportunityStage__c = "+$(sanitise-forSqlValue -dataType String -value $kimbleOpp.KimbleOne__OpportunityStage__c)
    $sql += ",KimbleOne__OpportunitySource__c = "+$(sanitise-forSqlValue -dataType String -value $kimbleOpp.KimbleOne__OpportunitySource__c)
    $sql += ",KimbleOne__ForecastStatus__c = "+$(sanitise-forSqlValue -dataType String -value $kimbleOpp.KimbleOne__ForecastStatus__c)
    $sql += ",KimbleOne__Description__c = "+$(sanitise-forSqlValue -dataType String -value $kimbleOpp.KimbleOne__Description__c)
    $sql += ",KimbleOne__CloseDate__c = "+$(sanitise-forSqlValue -dataType Date -value $kimbleOpp.KimbleOne__CloseDate__c)
    $sql += ",KimbleOne__Account__c = "+$(sanitise-forSqlValue -dataType String -value $kimbleOpp.KimbleOne__Account__c)
    $sql += ",IsDeleted = "+$(sanitise-forSqlValue -dataType Boolean -value $kimbleOpp.IsDeleted)
	#$sql += "'"+$(smartReplace $kimbleOpp.Id "'" "`'`'")+"',"
    $sql += ",CreatedDate = "+$(sanitise-forSqlValue -dataType Date -value $kimbleOpp.CreatedDate)
    $sql += ",CreatedById = "+$(sanitise-forSqlValue -dataType String -value $kimbleOpp.CreatedById)
    $sql += ",Community__c = "+$(sanitise-forSqlValue -dataType String -value $kimbleOpp.Community__c)
    $sql += ",ANTH_SalesOpportunityStagesCount__c = "+$(sanitise-forSqlValue -dataType String -value $kimbleOpp.ANTH_SalesOpportunityStagesCount__c)
    $sql += ",ANTH_PipelineStage__c = "+$(sanitise-forSqlValue -dataType String -value $kimbleOpp.ANTH_PipelineStage__c)
    $sql += " WHERE id = '$(sanitise-forSql $kimbleOpp.Id)'"
    Write-Verbose "`t$sql"
    $result = Execute-SQLQueryOnSQLDB -query $sql -queryType nonquery -sqlServerConnection $dbConnection
    if($result -eq 1){Write-Verbose "`t`tSUCCESS!"}else{Write-Verbose "`t`tFAILURE :( - Code: [$result]"}
    $result
    }
function update-kimbleProposalToFocalPointCache($kimbleProp, $dbConnection, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Yellow "update-kimbleOppToFocalPointCache"}
    $sql = "UPDATE SUS_Kimble_Proposals "
    #(attributes, SystemModstamp, Weighted_Net_Revenue__c, Proposal_Contract_Revenue__c, Project_Type__c, OwnerId, Name, LastModifiedDate, LastModifiedById, LastActivityDate, KimbleOne__WonLostReason__c, KimbleOne__WonLostNarrative__c, KimbleOne__ResponseRequiredDate__c, KimbleOne__Proposal__c, KimbleOne__OpportunityStage__c, KimbleOne__OpportunitySource__c, KimbleOne__ForecastStatus__c, KimbleOne__Description__c, KimbleOne__CloseDate__c, KimbleOne__Account__c, IsDeleted, CreatedDate, CreatedById, Community__c, ANTH_SalesOpportunityStagesCount__c, ANTH_PipelineStage__c) VALUES ("
	$sql += "SET attributes = '"+$kimbleProp.attributes.url+"',"
	$sql += "SystemModstamp ='"+$(Get-Date (smartReplace $kimbleProp.SystemModstamp "+0000" "") -Format s -ErrorAction SilentlyContinue)+"',"
    if($kimbleProp.KimbleOne__WeightedContractRevenue__c -eq $null){$sql += "KimbleOne__WeightedContractRevenue__c = 0,"}else{$sql += "KimbleOne__WeightedContractRevenue__c = " + [string]$kimbleProp.KimbleOne__WeightedContractRevenue__c+ ","}
    if($kimbleProp.Proposal_Contract_Revenue__c -eq $null){$sql += "Proposal_Contract_Revenue__c = 0,"}else{$sql += "Proposal_Contract_Revenue__c = " + [string]$kimbleProp.Proposal_Contract_Revenue__c+ ","}
	$sql += "Project_Type__c = '"+$(smartReplace $kimbleProp.Project_Type__c "'" "`'`'")+"',"
	$sql += "OwnerId = '"+$(smartReplace $kimbleProp.OwnerId "'" "`'`'")+"',"
	$sql += "Name = '"+$(smartReplace $kimbleProp.Name "'" "`'`'")+"',"
	$sql += "LastModifiedDate = '"+$(Get-Date (smartReplace $kimbleProp.LastModifiedDate "+0000" "") -Format s -ErrorAction SilentlyContinue)+"',"
	$sql += "LastModifiedById = '"+$(smartReplace $kimbleProp.LastModifiedById "'" "`'`'")+"',"
    if($kimbleProp.KimbleOne__HighLevelWeightedContractRevenue__c -eq $null){$sql += "KimbleOne__HighLevelWeightedContractRevenue__c = 0,"}else{$sql += "KimbleOne__HighLevelWeightedContractRevenue__c = " + [string]$kimbleProp.KimbleOne__HighLevelWeightedContractRevenue__c+ ","}
    if($kimbleProp.KimbleOne__HighLevelContractRevenue__c -eq $null){$sql += "KimbleOne__HighLevelContractRevenue__c = 0,"}else{$sql += "KimbleOne__HighLevelContractRevenue__c = " + [string]$kimbleProp.KimbleOne__HighLevelContractRevenue__c+ ","}
    if($kimbleProp.KimbleOne__HighLevelContractCost__c -eq $null){$sql += "KimbleOne__HighLevelContractCost__c = 0,"}else{$sql += "KimbleOne__HighLevelContractCost__c = " + [string]$kimbleProp.KimbleOne__HighLevelContractCost__c+ ","}
    if($kimbleProp.KimbleOne__DetailedLevelWeightedContractRevenue__c -eq $null){$sql += "KimbleOne__DetailedLevelWeightedContractRevenue__c = 0,"}else{$sql += "KimbleOne__DetailedLevelWeightedContractRevenue__c = " + [string]$kimbleProp.KimbleOne__DetailedLevelWeightedContractRevenue__c+ ","}
    if($kimbleProp.KimbleOne__DetailedLevelContractRevenue__c -eq $null){$sql += "KimbleOne__DetailedLevelContractRevenue__c = 0,"}else{$sql += "KimbleOne__DetailedLevelContractRevenue__c = " + [string]$kimbleProp.KimbleOne__DetailedLevelContractRevenue__c+ ","}
    if($kimbleProp.KimbleOne__DetailedLevelContractCost__c -eq $null){$sql += "KimbleOne__DetailedLevelContractCost__c = 0,"}else{$sql += "KimbleOne__DetailedLevelContractCost__c = " + [string]$kimbleProp.KimbleOne__DetailedLevelContractCost__c+ ","}
    if($kimbleProp.KimbleOne__ContractRevenue__c -eq $null){$sql += "KimbleOne__ContractRevenue__c = 0,"}else{$sql += "KimbleOne__ContractRevenue__c = " + [string]$kimbleProp.KimbleOne__ContractRevenue__c+ ","}
    if($kimbleProp.KimbleOne__ContractMargin__c -eq $null){$sql += "KimbleOne__ContractMargin__c = 0,"}else{$sql += "KimbleOne__ContractMargin__c = " + [string]$kimbleProp.KimbleOne__ContractMargin__c+ ","}
    if($kimbleProp.KimbleOne__ContractMarginAmount__c -eq $null){$sql += "KimbleOne__ContractMarginAmount__c = 0,"}else{$sql += "KimbleOne__ContractMarginAmount__c = " + [string]$kimbleProp.KimbleOne__ContractMarginAmount__c+ ","}
    if($kimbleProp.KimbleOne__ContractCost__c -eq $null){$sql += "KimbleOne__ContractCost__c = 0,"}else{$sql += "KimbleOne__ContractCost__c = " + [string]$kimbleProp.KimbleOne__ContractCost__c+ ","}
	$sql += "KimbleOne__RiskLevel__c = '"+$(smartReplace $kimbleProp.KimbleOne__RiskLevel__c "'" "`'`'")+"',"
	$sql += "KimbleOne__Reference__c = '"+$(smartReplace $kimbleProp.KimbleOne__Reference__c "'" "`'`'")+"',"
	$sql += "KimbleOne__Proposition__c = '"+$(smartReplace $kimbleProp.KimbleOne__Proposition__c "'" "`'`'")+"',"
	$sql += "KimbleOne__ForecastStatus__c = '"+$(smartReplace $kimbleProp.KimbleOne__ForecastStatus__c "'" "`'`'")+"',"
    if($kimbleProp.KimbleOne__ForecastAtDetailedLevel__c -eq $true){$sql += "KimbleOne__ForecastAtDetailedLevel__c = 1,"}else{$sql += "KimbleOne__ForecastAtDetailedLevel__c = 0,"}
	$sql += "KimbleOne__Description__c = '"+$(smartReplace $kimbleProp.KimbleOne__Description__c "'" "`'`'")+"',"
	$sql += "KimbleOne__BusinessUnit__c = '"+$(smartReplace $kimbleProp.KimbleOne__BusinessUnit__c "'" "`'`'")+"',"
	$sql += "KimbleOne__AcceptanceType__c = '"+$(smartReplace $kimbleProp.KimbleOne__AcceptanceType__c "'" "`'`'")+"',"
	$sql += "KimbleOne__AcceptanceDate__c = '"+$(smartReplace $kimbleProp.KimbleOne__AcceptanceDate__c "'" "`'`'")+"',"
    if($kimbleProp.IsDeleted -eq $true){$sql += "IsDeleted = 1,"}else{$sql += "IsDeleted = 0,"}
	#$sql += "'"+$(smartReplace $kimbleProp.Id "'" "`'`'")+"',"
	$sql += "CreatedDate = '"+$(Get-Date (smartReplace $kimbleProp.CreatedDate "+0000" "") -Format s -ErrorAction SilentlyContinue)+"',"
	$sql += "CreatedById = '"+$(smartReplace $kimbleProp.CreatedById "'" "`'`'")+"',"
	$sql += "CurrencyIsoCode = '"+$(smartReplace $kimbleProp.CurrencyIsoCode "'" "`'`'")+"'"
    $sql += " WHERE id = '$(sanitise-forSql $kimbleProp.Id)'"
    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`t$sql"}
    $result = Execute-SQLQueryOnSQLDB -query $sql -queryType nonquery -sqlServerConnection $dbConnection
    if($verboseLogging){if($result -eq 1){Write-Host -ForegroundColor DarkYellow "`t`tSUCCESS!"}else{Write-Host -ForegroundColor DarkYellow "`t`tFAILURE :( - Code: $result"}}
    $result
    }


