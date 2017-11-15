[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$callbackUri = "https://login.salesforce.com/services/oauth2/token"
#"https://test.salesforce.com/services/oauth2/token"
$grantType = "password"
$myInstance = "https://eu5.salesforce.com"
$queryUri = "$myInstance/services/data/v39.0/query/?q="
$querySuffixStub = " -H `"Authorization: Bearer "
$kimbleLogin = Import-Csv "$env:USERPROFILE\Desktop\Kimble.txt"
$clientId = $kimbleLogin.clientId
$clientSecret = $kimbleLogin.clientSecret
$username = $kimbleLogin.username
$password = $kimbleLogin.password
$securityToken = $kimbleLogin.securityToken

#region functions
function Get-KimbleAuthorizationTokenWithUsernamePasswordFlowRequestBody($client_id, $client_secret, $user_name, $pass_word, $security_token){
    Add-Type -AssemblyName System.Web
    $user_name = [System.Web.HttpUtility]::UrlEncode($user_name)
    $pass_word = [System.Web.HttpUtility]::UrlEncode($pass_word)
    $requestBody = "grant_type=$grantType"
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
function Get-KimbleSoqlDataset($queryUri, $soqlQuery, $restHeaders){
    $soqlQuery = $soqlQuery.Replace(" ","+")
    $lastIndex = 0
    $nextIndex = 0
    do{
        $lastIndex = $nextIndex
        if($lastIndex -eq 0){
            #Write-Host -ForegroundColor Magenta $partialDataSet.totalSize
            $partialDataSet = Invoke-RestMethod -Uri $queryUri+$soqlQuery -Headers $restHeaders
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
function get-allKimbleAccounts($pQueryUri, $pRestHeaders){
    $query = "Select a.X2nd_Account_Owner__c, a.Website, a.Walmart_Supplier_Account__c, a.Walmart_Sub_Category__c, a.Walmart_Category_New__c, a.Type_of_Product__c, a.Type, a.SystemModstamp, a.Sym_Group__c, a.Supply_Since_Last_Update__c, a.Supplier_Sector__c, a.SicDesc, a.ShippingStreet, a.ShippingState, a.ShippingPostalCode, a.ShippingLongitude, a.ShippingLatitude, a.ShippingGeocodeAccuracy, a.ShippingCountry, a.ShippingCity, a.ShippingAddress, a.Received_Sust_Index_from_WM__c, a.Purchase_Type__c, a.PhotoUrl, a.Phone, a.Partner_Sector__c, a.ParentId, a.OwnerId, a.NumberOfEmployees, a.Name, a.Membership__c, a.MasterRecordId, a.LastViewedDate, a.LastReferencedDate, a.LastModifiedDate, a.LastModifiedById, a.LastActivityDate, a.KimbleOne__TaxCode__c, a.KimbleOne__TaxCodeReference__c, a.KimbleOne__PurchaseOrderRule__c, a.KimbleOne__Locale__c, a.KimbleOne__IsSupplier__c, a.KimbleOne__IsCustomer__c, a.KimbleOne__InvoicingCurrencyIsoCode__c, a.KimbleOne__InvoiceTemplate__c, a.KimbleOne__InvoicePaymentTermDays__c, a.KimbleOne__InvoiceFormat__c, a.KimbleOne__Code__c, a.KimbleOne__BusinessUnit__c, a.KimbleOne__BusinessUnitTradingEntity__c, a.KimbleOne__BillingParentAccount__c, a.KimbleOne__BillingContact__c, a.KimbleOne__AllowPartItemInvoicing__c, a.Key_Account__c, a.JigsawCompanyId, a.Jigsaw, a.Is_Partner__c, a.Is_Competitor__c, a.IsDeleted, a.Industry, a.Id, a.Have_you_completed_this_section__c, a.Generic_email_address__c, a.Fax, a.Description, a.D_U_N_S_Number__c, a.CurrencyIsoCode, a.CreatedDate, a.CreatedById, a.Client_Sector__c, a.BillingStreet, a.BillingState, a.BillingPostalCode, a.BillingLongitude, a.BillingLatitude, a.BillingGeocodeAccuracy, a.BillingCountry, a.BillingCity, a.BillingAddress, a.AnnualRevenue, a.Account_Manager__c, a.Account_Director__c, a.AccountSource From Account a"
    Get-KimbleSoqlDataset -queryUri $pQueryUri -soqlQuery $query -restHeaders $pRestHeaders
    }
function get-allKimbleLeads($pQueryUri, $pRestHeaders){
    $query = "Select k.Won_Reason__c, k.Weighted_Net_Revenue__c, k.SystemModstamp, k.Proposal_Contract_Revenue__c, k.Project_Type__c, k.OwnerId, k.Name, k.Lost_to_competitor_reason__c, k.LastViewedDate, k.LastReferencedDate, k.LastModifiedDate, k.LastModifiedById, k.LastActivityDate, k.KimbleOne__WonLostReason__c, k.KimbleOne__WonLostNarrative__c, k.KimbleOne__WeightedContractRevenue__c, k.KimbleOne__TaxCode__c, k.KimbleOne__ShortName__c, k.KimbleOne__ResponseRequiredDate__c, k.KimbleOne__Reference__c, k.KimbleOne__Proposition__c, k.KimbleOne__Proposal__c, k.KimbleOne__OpportunityStage__c, k.KimbleOne__OpportunitySource__c, k.KimbleOne__MarketingCampaign__c, k.KimbleOne__LostToCompetitor__c, k.KimbleOne__InvoicingCurrencyISOCode__c, k.KimbleOne__ForecastStatus__c, k.KimbleOne__Description__c, k.KimbleOne__CustomerAccount__c, k.KimbleOne__ContractRevenue__c, k.KimbleOne__ContractMargin__c, k.KimbleOne__ContractMarginAmount__c, k.KimbleOne__ContractCost__c, k.KimbleOne__CloseDate__c, k.KimbleOne__BusinessUnit__c, k.KimbleOne__Account__c, k.IsDeleted, k.Id, k.CurrencyIsoCode, k.CreatedDate, k.CreatedById, k.Country__c, k.Competitive__c, k.Community__c, k.ANTH_SalesOpportunityStagesCount__c, k.ANTH_PipelineStage__c From KimbleOne__SalesOpportunity__c k"
    Get-KimbleSoqlDataset -queryUri $pQueryUri -soqlQuery $query -restHeaders $pRestHeaders
    }
function get-allKimbleProjects($pQueryUri, $pRestHeaders){
    $query = "Select k.SystemModstamp, k.Project_Type__c, k.Primary_Client_Contact__c, k.OwnerId, k.Name, k.LastViewedDate, k.LastReferencedDate, k.LastModifiedDate, k.LastModifiedById, k.LastActivityDate, k.KimbleOne__WeightedContractRevenue__c, k.KimbleOne__ShortName__c, k.KimbleOne__SalesOpportunity__c, k.KimbleOne__RiskLevel__c, k.KimbleOne__Reference__c, k.KimbleOne__Proposal__c, k.KimbleOne__ProductGroup__c, k.KimbleOne__ProbabilityCodeEnum__c, k.KimbleOne__LostReason__c, k.KimbleOne__LostReasonNarrative__c, k.KimbleOne__IsExpectedStartDateBeforeCloseDate__c, k.KimbleOne__InvoicingCcyServicesInvoiceableAmount__c, k.KimbleOne__InvoicingCcyExpensesInvoiceableAmount__c, k.KimbleOne__ForecastUsage__c, k.KimbleOne__ForecastStatus__c, k.KimbleOne__ForecastAtThisLevel__c, k.KimbleOne__ExpectedStartDate__c, k.KimbleOne__ExpectedEndDate__c, k.KimbleOne__ExpectedCurrencyISOCode__c, k.KimbleOne__ExpectedCcyExpectedContractRevenue__c, k.KimbleOne__DeliveryStatus__c, k.KimbleOne__DeliveryStage__c, k.KimbleOne__DeliveryProgram__c, k.KimbleOne__ContractRevenue__c, k.KimbleOne__ContractMargin__c, k.KimbleOne__ContractMarginAmount__c, k.KimbleOne__ContractCost__c, k.KimbleOne__BusinessUnitGroup__c, k.KimbleOne__BaselineUsage__c, k.KimbleOne__BaselineContractRevenue__c, k.KimbleOne__BaselineContractMargin__c, k.KimbleOne__BaselineContractMarginAmount__c, k.KimbleOne__BaselineContractCost__c, k.KimbleOne__ActualUsage__c, k.KimbleOne__Account__c, k.Is_Engagement_Owner__c, k.IsDeleted, k.Id, k.Finance_Contact__c, k.CurrencyIsoCode, k.CreatedDate, k.CreatedById, k.Competitive__c, k.Community__c From KimbleOne__DeliveryGroup__c k"
    Get-KimbleSoqlDataset -queryUri $pQueryUri -soqlQuery $query -restHeaders $pRestHeaders
    }
#endregion


