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

function add-kimbleAccountToFocalPointCache($kimbleAccount, $dbConnection, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Yellow "add-kimbleAccountToFocalPointCache"}
    $sql = "SELECT Name,Id FROM SUS_Kimble_Accounts WHERE Id = '$($kimbleAccount.Id)'"
    $alreadyPresent = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $dbConnection
    if ($alreadyPresent){
        if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`Id already present in Cache, not adding duplicate"}
        -1 #Return "unsuccessful" as the record is in the DB, but we didn't add it and this might need investigating
        }
    else{
        $sql = "INSERT INTO SUS_Kimble_Accounts (attributes, Website, Type, SystemModstamp, Phone, ParentId, OwnerId, Name, LastModifiedDate, LastModifiedById, KimbleOne__IsCustomer__c, KimbleOne__BusinessUnit__c, Is_Partner__c, Is_Competitor__c, IsDeleted, Id, CreatedDate, CreatedById, Client_Sector__c, BillingStreet, BillingState, BillingPostalCode, BillingCountry, BillingCity) VALUES ("
        $sql += "'"+$(sanitise-forSql $kimbleAccount.attributes)+"',"
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
        $sql += "'"+$(sanitise-forSql $kimbleContact.attributes)+"',"
        $sql += "'"+$(sanitise-forSql $(smartReplace $kimbleContact.Birthdate "+0000" ""))+"',"
        if($kimbleContact.Cleanup__c -eq $true){$sql += "1,"}else{$sql += "0,"}
        $sql += "'"+$(sanitise-forSql $kimbleContact.Client_Type__c)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.CreatedById)+"',"
        $sql += "'"+$(sanitise-forSql $(smartReplace $kimbleContact.CreatedDate "+0000" ""))+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.CurrencyIsoCode)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Department)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Description)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Email)+"',"
        $sql += "'"+$(sanitise-forSql $(smartReplace $kimbleContact.EmailBouncedDate "+0000" ""))+"',"
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
        $sql += "'"+$(sanitise-forSql $(smartReplace $kimbleContact.LastActivityDate "+0000" ""))+"',"
        $sql += "'"+$(sanitise-forSql $(smartReplace $kimbleContact.LastCURequestDate "+0000" ""))+"',"
        $sql += "'"+$(sanitise-forSql $(smartReplace $kimbleContact.LastCUUpdateDate "+0000" ""))+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.LastModifiedById)+"',"
        $sql += "'"+$(sanitise-forSql $(smartReplace $kimbleContact.LastModifiedDate "+0000" ""))+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.LastName)+"',"
        $sql += "'"+$(sanitise-forSql $(smartReplace $kimbleContact.LastReferencedDate "+0000" ""))+"',"
        $sql += "'"+$(sanitise-forSql $(smartReplace $kimbleContact.LastViewedDate "+0000" ""))+"',"
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
        $sql += "'"+$(sanitise-forSql $(smartReplace $kimbleContact.SystemModstamp "+0000" ""))+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Title)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Twitter__c)+"',"
        $sql += "'"+$(sanitise-forSql $kimbleContact.Unsubscribe_Newsletter_Campaigns__c)+"')"

        if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`t$sql"}
        $result = Execute-SQLQueryOnSQLDB -query $sql -queryType nonquery -sqlServerConnection $dbConnection
        if($verboseLogging){if($result -eq 1){Write-Host -ForegroundColor DarkYellow "`t`tSUCCESS!"}else{Write-Host -ForegroundColor DarkYellow "`t`tFAILURE :( - Code: $result"}}
        $result
        }
    }
function add-kimbleOppToFocalPointCache($kimbleOpp, $dbConnection, $verboseLogging){
    $sql = "SELECT Name,Id FROM SUS_Kimble_Opps WHERE Id = '$($kimbleOpp.Id)'"
    $alreadyPresent = Execute-SQLQueryOnSQLDB -query $sql -queryType Reader -sqlServerConnection $dbConnection
    if ($alreadyPresent){
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
function get-allFocalPointCachedKimbleContacts($dbConnection, $pWhereStatement, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Yellow "get-allFocalPointCachedKimbleContacts"}
    $sql = "SELECT AccountId, Anthesis_Events__c, AssistantName, AssistantPhone, attributes, Birthdate, Cleanup__c, Client_Type__c, CreatedById, CreatedDate, CurrencyIsoCode, Department, Description, Email, EmailBouncedDate, EmailBouncedReason, Fax, FirstName, Gender__c, Have_you_completed_this_section__c, HomePhone, Id, IsDeleted, IsEmailBounced, Jigsaw, JigsawContactId, Key_areas_of_interest__c, LastActivityDate, LastCURequestDate, LastCUUpdateDate, LastModifiedById, LastModifiedDate, LastName, LastReferencedDate, LastViewedDate, Lead_Source_Detail__c, LeadSource, Linked_In__c, MailingAddress, MailingCity, MailingCountry, MailingGeocodeAccuracy, MailingLatitude, MailingLongitude, MailingPostalCode, MailingState, MailingStreet, MasterRecordId, Met_At__c, MobilePhone, Name, Nee__c, Newsletters_Campaigns__c, Nickname__c, No_Show_Event_Attendees__c, Other_Email__c, OtherAddress, OtherCity, OtherCountry, OtherGeocodeAccuracy, OtherLatitude, OtherLongitude, OtherPhone, OtherPostalCode, OtherState, OtherStreet, OwnerId, Phone, PhotoUrl, Previous_Company__c, Region__c, ReportsToId, Role_Responsibilities__c, Salutation, Secondary_contact_owner__c, Skype__c, SystemModstamp, Title, Twitter__c, Unsubscribe_Newsletter_Campaigns__c FROM SUS_Kimble_Contacts $pWhereStatement"
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
function get-allKimbleProjects($pQueryUri, $pRestHeaders, $pWhereStatement, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Yellow "get-allKimbleProjects"}
    $query = "Select k.SystemModstamp, k.Project_Type__c, k.Primary_Client_Contact__c, k.OwnerId, k.Name, k.LastViewedDate, k.LastReferencedDate, k.LastModifiedDate, k.LastModifiedById, k.LastActivityDate, k.KimbleOne__WeightedContractRevenue__c, k.KimbleOne__ShortName__c, k.KimbleOne__SalesOpportunity__c, k.KimbleOne__RiskLevel__c, k.KimbleOne__Reference__c, k.KimbleOne__Proposal__c, k.KimbleOne__ProductGroup__c, k.KimbleOne__ProbabilityCodeEnum__c, k.KimbleOne__LostReason__c, k.KimbleOne__LostReasonNarrative__c, k.KimbleOne__IsExpectedStartDateBeforeCloseDate__c, k.KimbleOne__InvoicingCcyServicesInvoiceableAmount__c, k.KimbleOne__InvoicingCcyExpensesInvoiceableAmount__c, k.KimbleOne__ForecastUsage__c, k.KimbleOne__ForecastStatus__c, k.KimbleOne__ForecastAtThisLevel__c, k.KimbleOne__ExpectedStartDate__c, k.KimbleOne__ExpectedEndDate__c, k.KimbleOne__ExpectedCurrencyISOCode__c, k.KimbleOne__ExpectedCcyExpectedContractRevenue__c, k.KimbleOne__DeliveryStatus__c, k.KimbleOne__DeliveryStage__c, k.KimbleOne__DeliveryProgram__c, k.KimbleOne__ContractRevenue__c, k.KimbleOne__ContractMargin__c, k.KimbleOne__ContractMarginAmount__c, k.KimbleOne__ContractCost__c, k.KimbleOne__BusinessUnitGroup__c, k.KimbleOne__BaselineUsage__c, k.KimbleOne__BaselineContractRevenue__c, k.KimbleOne__BaselineContractMargin__c, k.KimbleOne__BaselineContractMarginAmount__c, k.KimbleOne__BaselineContractCost__c, k.KimbleOne__ActualUsage__c, k.KimbleOne__Account__c, k.Is_Engagement_Owner__c, k.IsDeleted, k.Id, k.Finance_Contact__c, k.CurrencyIsoCode, k.CreatedDate, k.CreatedById, k.Competitive__c, k.Community__c From KimbleOne__DeliveryGroup__c k $pWhereStatement"
    if($verboseLogging){Write-Host -ForegroundColor DarkYellow "`t`$query = $query"}
    Get-KimbleSoqlDataset -queryUri $pQueryUri -soqlQuery $query -restHeaders $pRestHeaders
    }
function get-allKimbleContacts($pQueryUri, $pRestHeaders, $pWhereStatement, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Yellow "get-allKimbleContacts"}
    $query = "Select c.Unsubscribe_Newsletter_Campaigns__c, c.Twitter__c, c.Title, c.SystemModstamp, c.Skype__c, c.Secondary_contact_owner__c, c.Salutation, c.Role_Responsibilities__c, c.ReportsToId, c.Region__c, c.Previous_Company__c, c.PhotoUrl, c.Phone, c.OwnerId, c.Other_Email__c, c.OtherStreet, c.OtherState, c.OtherPostalCode, c.OtherPhone, c.OtherLongitude, c.OtherLatitude, c.OtherGeocodeAccuracy, c.OtherCountry, c.OtherCity, c.OtherAddress, c.No_Show_Event_Attendees__c, c.Nickname__c, c.Newsletters_Campaigns__c, c.Nee__c, c.Name, c.MobilePhone, c.Met_At__c, c.MasterRecordId, c.MailingStreet, c.MailingState, c.MailingPostalCode, c.MailingLongitude, c.MailingLatitude, c.MailingGeocodeAccuracy, c.MailingCountry, c.MailingCity, c.MailingAddress, c.Linked_In__c, c.Lead_Source_Detail__c, c.LeadSource, c.LastViewedDate, c.LastReferencedDate, c.LastName, c.LastModifiedDate, c.LastModifiedById, c.LastCUUpdateDate, c.LastCURequestDate, c.LastActivityDate, c.Key_areas_of_interest__c, c.JigsawContactId, c.Jigsaw, c.IsEmailBounced, c.IsDeleted, c.Id, c.HomePhone, c.Have_you_completed_this_section__c, c.Gender__c, c.FirstName, c.Fax, c.EmailBouncedReason, c.EmailBouncedDate, c.Email, c.Description, c.Department, c.CurrencyIsoCode, c.CreatedDate, c.CreatedById, c.Client_Type__c, c.Cleanup__c, c.Birthdate, c.AssistantPhone, c.AssistantName, c.Anthesis_Events__c, c.AccountId From Contact c $pWhereStatement"
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
function update-kimbleContactToFocalPointCache($kimbleContact, $dbConnection, $verboseLogging){
    if($verboseLogging){Write-Host -ForegroundColor Yellow "update-kimbleContactToFocalPointCache"}
    $sql = "UPDATE SUS_Kimble_Contacts "
    #(AccountId,Anthesis_Events__c,AssistantName,AssistantPhone,attributes,Birthdate,Cleanup__c,Client_Type__c,CreatedById,CreatedDate,CurrencyIsoCode,Department,Description,Email,EmailBouncedDate,EmailBouncedReason,Fax,FirstName,Gender__c,Have_you_completed_this_section__c,HomePhone,Id,IsDeleted,IsEmailBounced,Jigsaw,JigsawContactId,Key_areas_of_interest__c,LastActivityDate,LastCURequestDate,LastCUUpdateDate,LastModifiedById,LastModifiedDate,LastName,LastReferencedDate,LastViewedDate,Lead_Source_Detail__c,LeadSource,Linked_In__c,MailingAddress,MailingCity,MailingCountry,MailingGeocodeAccuracy,MailingLatitude,MailingLongitude,MailingPostalCode,MailingState,MailingStreet,MasterRecordId,Met_At__c,MobilePhone,Name,Nee__c,Newsletters_Campaigns__c,Nickname__c,No_Show_Event_Attendees__c,Other_Email__c,OtherAddress,OtherCity,OtherCountry,OtherGeocodeAccuracy,OtherLatitude,OtherLongitude,OtherPhone,OtherPostalCode,OtherState,OtherStreet,OwnerId,Phone,PhotoUrl,Previous_Company__c,Region__c,ReportsToId,Role_Responsibilities__c,Salutation,Secondary_contact_owner__c,Skype__c,SystemModstamp,Title,Twitter__c,Unsubscribe_Newsletter_Campaigns__c
    $sql += "SET "
    $sql += "AccountId = '"+$kimbleContact.AccountId+"',"
    $sql += "Anthesis_Events__c = '"+$kimbleContact.Anthesis_Events__c+"',"
    $sql += "AssistantName = '"+$kimbleContact.AssistantName+"',"
    $sql += "AssistantPhone = '"+$kimbleContact.AssistantPhone+"',"
    $sql += "attributes = '"+$kimbleContact.attributes+"',"
    $sql += "Birthdate = '"+$(smartReplace $kimbleContact.Birthdate "+0000" "")+"',"
    if($kimbleContact.Cleanup__c -eq $true){$sql += "Cleanup__c = 1,"}else{$sql += "Cleanup__c = 0,"}
    $sql += "Client_Type__c = '"+$kimbleContact.Client_Type__c+"',"
    $sql += "CreatedById = '"+$kimbleContact.CreatedById+"',"
    $sql += "CreatedDate = '"+$(smartReplace $kimbleContact.CreatedDate "+0000" "")+"',"
    $sql += "CurrencyIsoCode = '"+$kimbleContact.CurrencyIsoCode+"',"
    $sql += "Department = '"+$kimbleContact.Department+"',"
    $sql += "Description = '"+$kimbleContact.Description+"',"
    $sql += "Email = '"+$kimbleContact.Email+"',"
    $sql += "EmailBouncedDate = '"+$(smartReplace $kimbleContact.EmailBouncedDate "+0000" "")+"',"
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
    $sql += "LastActivityDate = '"+$(smartReplace $kimbleContact.LastActivityDate "+0000" "")+"',"
    $sql += "LastCURequestDate = '"+$(smartReplace $kimbleContact.LastCURequestDate "+0000" "")+"',"
    $sql += "LastCUUpdateDate = '"+$(smartReplace $kimbleContact.LastCUUpdateDate "+0000" "")+"',"
    $sql += "LastModifiedById = '"+$kimbleContact.LastModifiedById+"',"
    $sql += "LastModifiedDate = '"+$(smartReplace $kimbleContact.LastModifiedDate "+0000" "")+"',"
    $sql += "LastName = '"+$kimbleContact.LastName+"',"
    $sql += "LastReferencedDate = '"+$(smartReplace $kimbleContact.LastReferencedDate "+0000" "")+"',"
    $sql += "LastViewedDate = '"+$(smartReplace $kimbleContact.LastViewedDate "+0000" "")+"',"
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
    $sql += "SystemModstamp = '"+$(smartReplace $kimbleContact.SystemModstamp "+0000" "")+"',"
    $sql += "Title = '"+$kimbleContact.Title+"',"
    $sql += "Twitter__c = '"+$kimbleContact.Twitter__c+"',"
    $sql += "Unsubscribe_Newsletter_Campaigns__c = '"+$kimbleContact.Unsubscribe_Newsletter_Campaigns__c+"' "
    $sql += "WHERE id = '$($kimbleContact.Id)'"
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


