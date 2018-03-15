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


function add-kimbleAccountToFocalPointCache($kimbleAccount, $dbConnection, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Yellow "add-kimbleAccountToFocalPointCache"}
    $sql = "SELECT Name,Id FROM SUS_Kimble_Accounts WHERE Id = '$($kimbleAccount.Id)'"
    $alreadyPresent = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $dbConnection
    if ($alreadyPresent.Count -gt 0){
        if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`Id already present in Cache, not adding duplicate"}
        -1 #Return "unsuccessful" as the record is in the DB, but we didn't add it and this might need investigating
        }
    else{
        $sql = "INSERT INTO SUS_Kimble_Accounts (attributes, Website, Type, SystemModstamp, Phone, ParentId, OwnerId, Name, LastModifiedDate, LastModifiedById, KimbleOne__IsCustomer__c, KimbleOne__BusinessUnit__c, Is_Partner__c, Is_Competitor__c, IsDeleted, Id, CreatedDate, CreatedById, Client_Sector__c, BillingStreet, BillingState, BillingPostalCode, BillingCountry, BillingCity) VALUES ("
        $sql += "'"+$kimbleAccount.attributes+"',"
        $sql += "'"+$kimbleAccount.Website+"',"
        $sql += "'"+$kimbleAccount.Type+"',"
        $sql += "'"+$(smartReplace $kimbleAccount.SystemModstamp "+0000" "")+"',"
        $sql += "'"+$kimbleAccount.Phone+"',"
        $sql += "'"+$kimbleAccount.ParentId+"',"
        $sql += "'"+$kimbleAccount.OwnerId+"',"
        $sql += "'"+$(smartReplace $kimbleAccount.Name "'" "`'`'")+"',"
        $sql += "'"+$(smartReplace $kimbleAccount.LastModifiedDate "+0000" "")+"',"
        $sql += "'"+$kimbleAccount.LastModifiedById+"',"
        if($kimbleAccount.KimbleOne__IsCustomer__c -eq $true){$sql += "1,"}else{$sql += "0,"}
        $sql += "'"+$kimbleAccount.KimbleOne__BusinessUnit__c+"',"
        if($kimbleAccount.Is_Partner__c -eq $true){$sql += "1,"}else{$sql += "0,"}
        if($kimbleAccount.Is_Competitor__c -eq $true){$sql += "1,"}else{$sql += "0,"}
        if($kimbleAccount.IsDeleted -eq $true){$sql += "1,"}else{$sql += "0,"}
        $sql += "'"+$kimbleAccount.Id+"',"
        $sql += "'"+$(smartReplace $kimbleAccount.CreatedDate "+0000" "")+"',"
        $sql += "'"+$kimbleAccount.CreatedById+"',"
        $sql += "'"+$(smartReplace $kimbleAccount.Client_Sector__c "'" "`'`'")+"',"
        $sql += "'"+$(smartReplace $kimbleAccount.BillingStreet "'" "`'`'")+"',"
        $sql += "'"+$(smartReplace $kimbleAccount.BillingState "'" "`'`'")+"',"
        $sql += "'"+$(smartReplace $kimbleAccount.BillingPostalCode "'" "`'`'")+"',"
        $sql += "'"+$(smartReplace $kimbleAccount.BillingCountry "'" "`'`'")+"',"
        $sql += "'"+$(smartReplace $kimbleAccount.BillingCity "'" "`'`'")+"')"
        if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`t$sql"}
        $result = Execute-SQLQueryOnSQLDB -query $sql -queryType nonquery -sqlServerConnection $dbConnection
        if($verboseLogging){if($result -eq 1){Write-Host -ForegroundColor DarkYellow "`t`tSUCCESS!"}else{Write-Host -ForegroundColor DarkYellow "`t`tFAILURE :( - Code: $result"}}
        $result
        }
    }
function add-kimbleOppToFocalPointCache($kimbleOpp, $dbConnection, $verboseLogging){
    $sql = "SELECT Name,Id FROM SUS_Kimble_Opps WHERE Id = '$($kimbleOpp.Id)'"
    $alreadyPresent = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $dbConnection
    if ($alreadyPresent.Count -gt 0){
        if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`Id already present in Cache, not adding duplicate"}
        -1 #Return "unsuccessful" as the record is in the DB, but we didn't add it and this might need investigating
        }
    else{
        if($verboseLogging){Write-Host -ForegroundColor Yellow "add-kimbleOppToFocalPointCache"}
        $sql = "INSERT INTO SUS_Kimble_Opps (attributes, SystemModstamp, Weighted_Net_Revenue__c, Proposal_Contract_Revenue__c, Project_Type__c, OwnerId, Name, LastModifiedDate, LastModifiedById, LastActivityDate, KimbleOne__WonLostReason__c, KimbleOne__WonLostNarrative__c, KimbleOne__ResponseRequiredDate__c, KimbleOne__Proposal__c, KimbleOne__OpportunityStage__c, KimbleOne__OpportunitySource__c, KimbleOne__ForecastStatus__c, KimbleOne__Description__c, KimbleOne__CloseDate__c, KimbleOne__Account__c, IsDeleted, Id, CreatedDate, CreatedById, Community__c, ANTH_SalesOpportunityStagesCount__c, ANTH_PipelineStage__c) VALUES ("
	    $sql += "'"+$kimbleOpp.attributes+"',"
	    $sql += "'"+$(smartReplace $kimbleOpp.SystemModstamp "+0000" "")+"',"
        if($_.Weighted_Net_Revenue__ -eq $null){$sql += "0,"}else{$sql += $kimbleOpp.Weighted_Net_Revenue__+ ","}
	    $sql += [string]$kimbleOpp.Proposal_Contract_Revenue__c+ ","
	    $sql += "'"+$(smartReplace $kimbleOpp.Project_Type__c "'" "`'`'")+"',"
	    $sql += "'"+$(smartReplace $kimbleOpp.OwnerId "'" "`'`'")+"',"
	    $sql += "'"+$(smartReplace $kimbleOpp.Name "'" "`'`'")+"',"
	    $sql += "'"+$(smartReplace $kimbleOpp.LastModifiedDate "+0000" "")+"',"
	    $sql += "'"+$(smartReplace $kimbleOpp.LastModifiedById "'" "`'`'")+"',"
	    $sql += "'"+$(smartReplace $kimbleOpp.LastActivityDate "+0000" "")+"',"
	    $sql += "'"+$(smartReplace $kimbleOpp.KimbleOne__WonLostReason__c "'" "`'`'")+"',"
	    $sql += "'"+$(smartReplace $kimbleOpp.KimbleOne__WonLostNarrative__c "'" "`'`'")+"',"
	    $sql += "'"+$(smartReplace $kimbleOpp.KimbleOne__ResponseRequiredDate__c "+0000" "")+"',"
	    $sql += "'"+$(smartReplace $kimbleOpp.KimbleOne__Proposal__c "'" "`'`'")+"',"
	    $sql += "'"+$(smartReplace $kimbleOpp.KimbleOne__OpportunityStage__c "'" "`'`'")+"',"
	    $sql += "'"+$(smartReplace $kimbleOpp.KimbleOne__OpportunitySource__c "'" "`'`'")+"',"
	    $sql += "'"+$(smartReplace $kimbleOpp.KimbleOne__ForecastStatus__c "'" "`'`'")+"',"
	    $sql += "'"+$(smartReplace $kimbleOpp.KimbleOne__Description__c "'" "`'`'")+"',"
	    $sql += "'"+$(smartReplace $kimbleOpp.KimbleOne__CloseDate__c "+0000" "")+"',"
	    $sql += "'"+$(smartReplace $kimbleOpp.KimbleOne__Account__c "'" "`'`'")+"',"
        if($_.IsDeleted -eq $true){$sql += "1,"}else{$sql += "0,"}
	    $sql += "'"+$(smartReplace $kimbleOpp.Id "'" "`'`'")+"',"
	    $sql += "'"+$(smartReplace $kimbleOpp.CreatedDate "+0000" "")+"',"
	    $sql += "'"+$(smartReplace $kimbleOpp.CreatedById "'" "`'`'")+"',"
	    $sql += "'"+$(smartReplace $kimbleOpp.Community__c "'" "`'`'")+"',"
	    $sql += [string]$kimbleOpp.ANTH_SalesOpportunityStagesCount__c + ","
	    $sql += "'"+$kimbleOpp.ANTH_PipelineStage__c+"')"
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
        #Failure
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
    $sql = "SELECT attributes, Website, Type, SystemModstamp, Phone, ParentId, OwnerId, Name, LastModifiedDate, LastModifiedById, KimbleOne__IsCustomer__c, KimbleOne__BusinessUnit__c, Is_Partner__c, Is_Competitor__c, IsDeleted, Id, CreatedDate, CreatedById, Client_Sector__c, BillingStreet, BillingState, BillingPostalCode, BillingCountry, BillingCity FROM SUS_Kimble_Accounts $pWhereStatement"
    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`t`$query = $sql"}
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
function get-allKimbleAccounts($pQueryUri, $pRestHeaders, $pWhereStatement, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Yellow "get-allKimbleAccounts"}
    if($pWhereStatement){if($pWhereStatement.Split(" ")[0] -notmatch "WHERE"){$whereStatement = " WHERE "+$pWhereStatement}}
    else{$whereStatement = $pWhereStatement}
    $query = "Select a.Sync_to_FocalPoint__c, a.X2nd_Account_Owner__c, a.Website, a.Walmart_Supplier_Account__c, a.Walmart_Sub_Category__c, a.Walmart_Category_New__c, a.Type_of_Product__c, a.Type, a.SystemModstamp, a.Sym_Group__c, a.Supply_Since_Last_Update__c, a.Supplier_Sector__c, a.SicDesc, a.ShippingStreet, a.ShippingState, a.ShippingPostalCode, a.ShippingLongitude, a.ShippingLatitude, a.ShippingGeocodeAccuracy, a.ShippingCountry, a.ShippingCity, a.ShippingAddress, a.Received_Sust_Index_from_WM__c, a.Purchase_Type__c, a.PhotoUrl, a.Phone, a.Partner_Sector__c, a.ParentId, a.OwnerId, a.NumberOfEmployees, a.Name, a.Membership__c, a.MasterRecordId, a.LastViewedDate, a.LastReferencedDate, a.LastModifiedDate, a.LastModifiedById, a.LastActivityDate, a.KimbleOne__TaxCode__c, a.KimbleOne__TaxCodeReference__c, a.KimbleOne__PurchaseOrderRule__c, a.KimbleOne__Locale__c, a.KimbleOne__IsSupplier__c, a.KimbleOne__IsCustomer__c, a.KimbleOne__InvoicingCurrencyIsoCode__c, a.KimbleOne__InvoiceTemplate__c, a.KimbleOne__InvoicePaymentTermDays__c, a.KimbleOne__InvoiceFormat__c, a.KimbleOne__Code__c, a.KimbleOne__BusinessUnit__c, a.KimbleOne__BusinessUnitTradingEntity__c, a.KimbleOne__BillingParentAccount__c, a.KimbleOne__BillingContact__c, a.KimbleOne__AllowPartItemInvoicing__c, a.Key_Account__c, a.JigsawCompanyId, a.Jigsaw, a.Is_Partner__c, a.Is_Competitor__c, a.IsDeleted, a.Industry, a.Id, a.Have_you_completed_this_section__c, a.Generic_email_address__c, a.Fax, a.Description, a.D_U_N_S_Number__c, a.CurrencyIsoCode, a.CreatedDate, a.CreatedById, a.Client_Sector__c, a.BillingStreet, a.BillingState, a.BillingPostalCode, a.BillingLongitude, a.BillingLatitude, a.BillingGeocodeAccuracy, a.BillingCountry, a.BillingCity, a.BillingAddress, a.AnnualRevenue, a.Account_Manager__c, a.Account_Director__c, a.AccountSource From Account a $pWhereStatement"
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
function get-allKimbleProjects($pQueryUri, $pRestHeaders, $pWhereStatement, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Yellow "get-allKimbleProjects"}
    $query = "Select k.SystemModstamp, k.Project_Type__c, k.Primary_Client_Contact__c, k.OwnerId, k.Name, k.LastViewedDate, k.LastReferencedDate, k.LastModifiedDate, k.LastModifiedById, k.LastActivityDate, k.KimbleOne__WeightedContractRevenue__c, k.KimbleOne__ShortName__c, k.KimbleOne__SalesOpportunity__c, k.KimbleOne__RiskLevel__c, k.KimbleOne__Reference__c, k.KimbleOne__Proposal__c, k.KimbleOne__ProductGroup__c, k.KimbleOne__ProbabilityCodeEnum__c, k.KimbleOne__LostReason__c, k.KimbleOne__LostReasonNarrative__c, k.KimbleOne__IsExpectedStartDateBeforeCloseDate__c, k.KimbleOne__InvoicingCcyServicesInvoiceableAmount__c, k.KimbleOne__InvoicingCcyExpensesInvoiceableAmount__c, k.KimbleOne__ForecastUsage__c, k.KimbleOne__ForecastStatus__c, k.KimbleOne__ForecastAtThisLevel__c, k.KimbleOne__ExpectedStartDate__c, k.KimbleOne__ExpectedEndDate__c, k.KimbleOne__ExpectedCurrencyISOCode__c, k.KimbleOne__ExpectedCcyExpectedContractRevenue__c, k.KimbleOne__DeliveryStatus__c, k.KimbleOne__DeliveryStage__c, k.KimbleOne__DeliveryProgram__c, k.KimbleOne__ContractRevenue__c, k.KimbleOne__ContractMargin__c, k.KimbleOne__ContractMarginAmount__c, k.KimbleOne__ContractCost__c, k.KimbleOne__BusinessUnitGroup__c, k.KimbleOne__BaselineUsage__c, k.KimbleOne__BaselineContractRevenue__c, k.KimbleOne__BaselineContractMargin__c, k.KimbleOne__BaselineContractMarginAmount__c, k.KimbleOne__BaselineContractCost__c, k.KimbleOne__ActualUsage__c, k.KimbleOne__Account__c, k.Is_Engagement_Owner__c, k.IsDeleted, k.Id, k.Finance_Contact__c, k.CurrencyIsoCode, k.CreatedDate, k.CreatedById, k.Competitive__c, k.Community__c From KimbleOne__DeliveryGroup__c k $pWhereStatement"
    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`t`$query = $query"}
    Get-KimbleSoqlDataset -queryUri $pQueryUri -soqlQuery $query -restHeaders $pRestHeaders
    }
function update-kimbleAccountToFocalPointCache($kimbleAccount, $dbConnection, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Yellow "update-kimbleAccountToFocalPointCache"}
    $sql = "UPDATE SUS_Kimble_Accounts "
    #(attributes, Website, Type, SystemModstamp, Phone, ParentId, OwnerId, Name, LastModifiedDate, LastModifiedById, KimbleOne__IsCustomer__c, KimbleOne__BusinessUnit__c, Is_Partner__c, Is_Competitor__c, IsDeleted, Id, CreatedDate, CreatedById, Client_Sector__c, BillingStreet, BillingState, BillingPostalCode, BillingCountry, BillingCity) VALUES ("
    $sql += "SET attributes = '"+$kimbleAccount.attributes+"',"
    $sql += "Website = '"+$kimbleAccount.Website+"',"
    $sql += "Type = '"+$kimbleAccount.Type+"',"
    $sql += "SystemModstamp = '"+$(smartReplace $kimbleAccount.SystemModstamp "+0000" "")+"',"
    $sql += "Phone = '"+$kimbleAccount.Phone+"',"
    $sql += "ParentId = '"+$kimbleAccount.ParentId+"',"
    $sql += "OwnerId = '"+$kimbleAccount.OwnerId+"',"
    #$sql += "OwnerId = 'Kev',"
    $sql += "Name = '"+$(smartReplace $kimbleAccount.Name "'" "`'`'")+"',"
    $sql += "LastModifiedDate = '"+$(smartReplace $kimbleAccount.LastModifiedDate "+0000" "")+"',"
    $sql += "LastModifiedById = '"+$kimbleAccount.LastModifiedById+"',"
    if($kimbleAccount.KimbleOne__IsCustomer__c -eq $true){$sql += "KimbleOne__IsCustomer__c = 1,"}else{$sql += "KimbleOne__IsCustomer__c = 0,"}
    $sql += "KimbleOne__BusinessUnit__c = '"+$kimbleAccount.KimbleOne__BusinessUnit__c+"',"
    if($kimbleAccount.Is_Partner__c -eq $true){$sql += "Is_Partner__c = 1,"}else{$sql += "Is_Partner__c = 0,"}
    if($kimbleAccount.Is_Competitor__c -eq $true){$sql += "Is_Competitor__c = 1,"}else{$sql += "Is_Competitor__c = 0,"}
    if($kimbleAccount.IsDeleted -eq $true){$sql += "IsDeleted = 1,"}else{$sql += "IsDeleted = 0,"}
    #$sql += "'"+$kimbleAccount.Id+"',"
    $sql += "CreatedDate = '"+$(smartReplace $kimbleAccount.CreatedDate "+0000" "")+"',"
    $sql += "CreatedById = '"+$kimbleAccount.CreatedById+"',"
    $sql += "Client_Sector__c = '"+$(smartReplace $kimbleAccount.Client_Sector__c "'" "`'`'")+"',"
    $sql += "BillingStreet = '"+$(smartReplace $kimbleAccount.BillingStreet "'" "`'`'")+"',"
    $sql += "BillingState = '"+$(smartReplace $kimbleAccount.BillingState "'" "`'`'")+"',"
    $sql += "BillingPostalCode = '"+$(smartReplace $kimbleAccount.BillingPostalCode "'" "`'`'")+"',"
    $sql += "BillingCountry = '"+$(smartReplace $kimbleAccount.BillingCountry "'" "`'`'")+"',"
    $sql += "BillingCity = '"+$(smartReplace $kimbleAccount.BillingCity "'" "`'`'")+"'"
    $sql += "WHERE id = '$($kimbleAccount.Id)'"
    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`t$sql"}
    $result = Execute-SQLQueryOnSQLDB -query $sql -queryType nonquery -sqlServerConnection $dbConnection
    if($verboseLogging){if($result -eq 1){Write-Host -ForegroundColor DarkYellow "`t`tSUCCESS!"}else{Write-Host -ForegroundColor DarkYellow "`t`tFAILURE :( - Code: $result"}}
    $result
    }
function update-kimbleOppToFocalPointCache($kimbleOpp, $dbConnection, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Yellow "update-kimbleOppToFocalPointCache"}
    $sql = "UPDATE SUS_Kimble_Opps "
    #(attributes, SystemModstamp, Weighted_Net_Revenue__c, Proposal_Contract_Revenue__c, Project_Type__c, OwnerId, Name, LastModifiedDate, LastModifiedById, LastActivityDate, KimbleOne__WonLostReason__c, KimbleOne__WonLostNarrative__c, KimbleOne__ResponseRequiredDate__c, KimbleOne__Proposal__c, KimbleOne__OpportunityStage__c, KimbleOne__OpportunitySource__c, KimbleOne__ForecastStatus__c, KimbleOne__Description__c, KimbleOne__CloseDate__c, KimbleOne__Account__c, IsDeleted, CreatedDate, CreatedById, Community__c, ANTH_SalesOpportunityStagesCount__c, ANTH_PipelineStage__c) VALUES ("
	$sql += "SET attributes = '"+$kimbleOpp.attributes+"',"
	$sql += "SystemModstamp ='"+$(smartReplace $kimbleOpp.SystemModstamp "+0000" "")+"',"
    if($_.Weighted_Net_Revenue__ -eq $null){$sql += "Weighted_Net_Revenue__c = 0,"}else{$sql += "Weighted_Net_Revenue__c = " + $kimbleOpp.Weighted_Net_Revenue__+ ","}
	$sql += "Proposal_Contract_Revenue__c = " + [string]$kimbleOpp.Proposal_Contract_Revenue__c+ ","
	$sql += "Project_Type__c = '"+$(smartReplace $kimbleOpp.Project_Type__c "'" "`'`'")+"',"
	$sql += "OwnerId = '"+$(smartReplace $kimbleOpp.OwnerId "'" "`'`'")+"',"
	$sql += "Name = '"+$(smartReplace $kimbleOpp.Name "'" "`'`'")+"',"
	$sql += "LastModifiedDate = '"+$(smartReplace $kimbleOpp.LastModifiedDate "+0000" "")+"',"
	$sql += "LastModifiedById = '"+$(smartReplace $kimbleOpp.LastModifiedById "'" "`'`'")+"',"
	$sql += "LastActivityDate = '"+$(smartReplace $kimbleOpp.LastActivityDate "+0000" "")+"',"
	$sql += "KimbleOne__WonLostReason__c = '"+$(smartReplace $kimbleOpp.KimbleOne__WonLostReason__c "'" "`'`'")+"',"
	$sql += "KimbleOne__WonLostNarrative__c = '"+$(smartReplace $kimbleOpp.KimbleOne__WonLostNarrative__c "'" "`'`'")+"',"
	$sql += "KimbleOne__ResponseRequiredDate__c = '"+$(smartReplace $kimbleOpp.KimbleOne__ResponseRequiredDate__c "+0000" "")+"',"
	$sql += "KimbleOne__Proposal__c = '"+$(smartReplace $kimbleOpp.KimbleOne__Proposal__c "'" "`'`'")+"',"
	$sql += "KimbleOne__OpportunityStage__c = '"+$(smartReplace $kimbleOpp.KimbleOne__OpportunityStage__c "'" "`'`'")+"',"
	$sql += "KimbleOne__OpportunitySource__c = '"+$(smartReplace $kimbleOpp.KimbleOne__OpportunitySource__c "'" "`'`'")+"',"
	$sql += "KimbleOne__ForecastStatus__c = '"+$(smartReplace $kimbleOpp.KimbleOne__ForecastStatus__c "'" "`'`'")+"',"
	$sql += "KimbleOne__Description__c = '"+$(smartReplace $kimbleOpp.KimbleOne__Description__c "'" "`'`'")+"',"
	$sql += "KimbleOne__CloseDate__c = '"+$(smartReplace $kimbleOpp.KimbleOne__CloseDate__c "+0000" "")+"',"
	$sql += "KimbleOne__Account__c = '"+$(smartReplace $kimbleOpp.KimbleOne__Account__c "'" "`'`'")+"',"
    if($_.IsDeleted -eq $true){$sql += "IsDeleted = 1,"}else{$sql += "IsDeleted = 0,"}
	#$sql += "'"+$(smartReplace $kimbleOpp.Id "'" "`'`'")+"',"
	$sql += "CreatedDate = '"+$(smartReplace $kimbleOpp.CreatedDate "+0000" "")+"',"
	$sql += "CreatedById = '"+$(smartReplace $kimbleOpp.CreatedById "'" "`'`'")+"',"
	$sql += "Community__c = '"+$(smartReplace $kimbleOpp.Community__c "'" "`'`'")+"',"
	$sql += "ANTH_SalesOpportunityStagesCount__c = " + [string]$kimbleOpp.ANTH_SalesOpportunityStagesCount__c + ","
	$sql += "ANTH_PipelineStage__c = '"+$kimbleOpp.ANTH_PipelineStage__c+"' "
    $sql += "WHERE id = '$($kimbleOpp.Id)'"
    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`t$sql"}
    $result = Execute-SQLQueryOnSQLDB -query $sql -queryType nonquery -sqlServerConnection $dbConnection
    if($verboseLogging){if($result -eq 1){Write-Host -ForegroundColor DarkYellow "`t`tSUCCESS!"}else{Write-Host -ForegroundColor DarkYellow "`t`tFAILURE :( - Code: $result"}}
    $result
    }



