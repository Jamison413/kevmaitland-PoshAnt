function combine-url($arrayOfStrings){ 
    $output = ""
    $arrayOfStrings | % {
        $output += $_.TrimStart("/").TrimEnd("/")+"/"
        }
    $output = $output.Substring(0,$output.Length-1)
    $output = $output.Replace("//","/").Replace("//","/").Replace("//","/")
    $output = $output.Replace("http:/","http://").Replace("https:/","https://")
    $output
    }
function compare-objectProperties {
    #https://blogs.technet.microsoft.com/janesays/2017/04/25/compare-all-properties-of-two-objects-in-windows-powershell/
    Param(
        [PSObject]$ReferenceObject,
        [PSObject]$DifferenceObject 
        )
    $objprops = $ReferenceObject | Get-Member -MemberType Property,NoteProperty | % Name
    $objprops += $DifferenceObject | Get-Member -MemberType Property,NoteProperty | % Name
    $objprops = $objprops | Sort | Select -Unique
    $diffs = @()
    foreach ($objprop in $objprops) {
        $diff = Compare-Object $ReferenceObject $DifferenceObject -Property $objprop
        if ($diff) {            
            $diffprops = @{
                PropertyName=$objprop
                RefValue=($diff | ? {$_.SideIndicator -eq '<='} | % $($objprop))
                DiffValue=($diff | ? {$_.SideIndicator -eq '=>'} | % $($objprop))
                }
            $diffs += New-Object PSObject -Property $diffprops
            }        
        }
    if ($diffs) {return ($diffs | Select PropertyName,RefValue,DiffValue)}     
    }
function convert-csvToSecureStrings(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
        [PSCustomObject]$rawCsvData        
        )
    
    $encryptedObject = New-Object psobject
    $rawCsvData.PSObject.Properties | ForEach-Object {
        $encryptedObject | Add-Member -MemberType NoteProperty -Name $_.Name -Value $(convertTo-localisedSecureString -plainText $_.Value)
        }
    $encryptedObject
    }
function convertTo-arrayOfEmailAddresses($blockOfText){
    [string[]]$addresses = @()
    $blockOfText | %{
        if(![string]::IsNullOrWhiteSpace($_)){
            foreach($blob in $_.Split(" ").Split("`r`n").Split(";").Split(",")){
                if($blob -match "@" -and $blob -match "."){$addresses += $blob.Replace("<","").Replace(">","").Replace(";","").Trim()}
                }
            }
        }
    $addresses
    }
function convertTo-arrayOfStrings($blockOfText){
    $strings = @()
    $blockOfText | %{
        foreach($blob in $_.Split(",").Split("`r`n")){
            if(![string]::IsNullOrEmpty($blob)){$strings += $blob}
            }
        }
    $strings
    }
function convertTo-exTimeZoneValue($pAmbiguousTimeZone){
    $singleResult = @()
    $tzs = get-timeZones
    if($pAmbiguousTimeZone -match '\('){
        $tryThis = $pAmbiguousTimeZone.Replace([regex]::Match($pAmbiguousTimeZone,"\(([^)]+)\)").Groups[0].Value,"").Trim() #Get everything not between "(" and ")"
        }
    else{$tryThis = $pAmbiguousTimeZone}
    [array]$singleResult = $tzs | ? {$_.PSChildName -eq $tryThis} #Match it to the registry timezone names
    if ($singleResult.Count -eq 1){$singleResult[0].PSChildName}
    else{
        #Try something else
        }
    }
function convertTo-localisedSecureString($plainText){
    if ($(Get-Module).Name -notcontains "_PS_Library_Forms"){Import-Module _PS_Library_Forms}
    if (!$plainText){$plainText = form-captureText -formTitle "PlainText" -formText "Enter the plain text to be converted to a secure string" -sizeX 300 -sizeY 200}
    ConvertTo-SecureString $plainText -AsPlainText -Force | ConvertFrom-SecureString
    }
function decrypt-SecureString($secureString){
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString)
    [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    }
function export-encryptedCsv(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true,ParameterSetName = "PreEncrypted")]
            [psobject]$encryptedCsvData        
        ,[parameter(Mandatory = $true,ParameterSetName = "NotEncrypted")]
            [psobject]$unencryptedCsvData
        ,[parameter(Mandatory = $true,ParameterSetName = "PreEncrypted")]
            [parameter(Mandatory = $true,ParameterSetName = "NotEncrypted")]
            [string]$pathToOutputCsv
        ,[parameter(Mandatory = $false,ParameterSetName = "PreEncrypted")]
            [parameter(Mandatory = $false,ParameterSetName = "NotEncrypted")]
            [switch]$force
        )
    if(!$encryptedCsvData){
        $encryptedCsvData = convert-csvToSecureStrings -rawCsvData $unencryptedCsvData
        }
    if(Test-Path $pathToOutputCsv){
        if($force){Remove-Item -Path $pathToOutputCsv -Force}
        else{Write-Error "File [$pathToOutputCsv] already exists";break}
        }
    Export-Csv -InputObject $encryptedCsvData -Path $pathToOutputCsv -NoTypeInformation -NoClobber
    remove-doubleQuotesFromCsv -inputFile $pathToOutputCsv
    }
function format-internationalPhoneNumber($pDirtyNumber,$p3letterIsoCountryCode,[boolean]$localise){
    if($pDirtyNumber.Length -gt 0){
        $dirtynumber = $pDirtyNumber.Split("ext")[0]
        $dirtynumber = $dirtyNumber.Trim() -replace '[^0-9]+',''
        switch ($p3letterIsoCountryCode){
            "ARE" {
                if($dirtyNumber.Length -eq 10 -and $dirtyNumber.Substring(0,1) -eq "0"){$dirtyNumber = $dirtyNumber.Substring(1,9)}
                if($dirtyNumber.Length -eq 12 -and $dirtyNumber.Substring(0,3) -eq "971"){$dirtyNumber = $dirtyNumber.Substring(3,9)}
                if($dirtyNumber.Length -eq 9){
                    if ($localise){}
                    else{$cleanNumber = "+971 $dirtyNumber"}
                    }
                }
            "CAN" {
                if($dirtyNumber.Length -eq 11 -and $dirtyNumber.Substring(0,1) -eq "1"){$dirtyNumber = $dirtyNumber.Substring(1,10)}
                if($dirtyNumber.Length -eq 10){
                    if ($localise){$cleanNumber = "+1 " + $dirtyNumber.Substring(1,3) + "-"+$dirtyNumber.Substring(4,3)+"-"+$dirtyNumber.Substring(7,4)}
                    else{$cleanNumber = "+1 $dirtyNumber"}
                    }
                }
            "CHN" {
                if($dirtyNumber.Length -eq 13 -and $dirtyNumber.Substring(0,2) -eq "86"){$dirtyNumber = $dirtyNumber.Substring(2,11)}
                if($dirtyNumber.Length -eq 11){
                    if ($localise){}
                    else{$cleanNumber = "+86 $dirtyNumber"}
                    }
                }
            "DEU" {
                if($dirtyNumber.Length -eq 12 -and $dirtyNumber.Substring(0,2) -eq "49"){$dirtyNumber = $dirtyNumber.Substring(2,10)}
                if($dirtyNumber.Length -eq 10){
                    if ($localise){}
                    else{$cleanNumber = "+49 $dirtyNumber"}
                    }
                }
            "ESP" {"ES"}
            "FIN" {
                if($dirtyNumber.Length -eq 12 -and $dirtyNumber.Substring(0,3) -eq "358"){$dirtyNumber = $dirtyNumber.Substring(3,9)}
                if($dirtyNumber.Length -eq 9){
                    if ($localise){}
                    else{$cleanNumber = "+358 $dirtyNumber"}
                    }
                }
            "GBR" {
                if($dirtyNumber.Length -eq 11 -and $dirtyNumber.Substring(0,1) -eq "0"){$dirtyNumber = $dirtyNumber.Substring(1,10)}
                if($dirtyNumber.Length -eq 12 -and $dirtyNumber.Substring(0,2) -eq "44"){$dirtyNumber = $dirtyNumber.Substring(2,10)}
                if($dirtyNumber.Length -eq 10){
                    if ($localise){}
                    else{$cleanNumber = "+44 $dirtyNumber"}
                    }
                }
            "IRL" {
                if($dirtyNumber.Substring(0,1) -eq "0"){$dirtyNumber = $dirtyNumber.Substring(1,$dirtyNumber.Length-1)}
                if($dirtyNumber.Substring(0,3) -eq "353"){$dirtyNumber = $dirtyNumber.Substring(3,$dirtyNumber.Length-3)}
                if ($localise){}
                else{$cleanNumber = "+353 $dirtyNumber"}
                }
            "PHL" {
                if($dirtyNumber.Length -eq 12 -and $dirtyNumber.Substring(0,2) -eq "63"){$dirtyNumber = $dirtyNumber.Substring(2,10)}
                if($dirtyNumber.Length -eq 10){
                    if ($localise){}
                    else{$cleanNumber = "+63 $dirtyNumber"}
                    }
                }
            "SWE" {
                if($dirtyNumber.Length -eq 11 -and $dirtyNumber.Substring(0,2) -eq "46"){$dirtyNumber = $dirtyNumber.Substring(2,9)}
                if($dirtyNumber.Length -eq 9){
                    if ($localise){}
                    else{$cleanNumber = "+46 $dirtyNumber"}
                    }
                }
            "USA" {
                if($dirtyNumber.Length -eq 11 -and $dirtyNumber.Substring(0,1) -eq 1){$dirtyNumber = $dirtyNumber.Substring(1,10)}
                if($dirtyNumber.Length -eq 10){
                    if ($localise){$cleanNumber = "+1 (" + $dirtyNumber.Substring(1,3) + ") "+$dirtyNumber.Substring(4,3)+"-"+$dirtyNumber.Substring(7,4)}
                    else{$cleanNumber = "+1 $dirtyNumber"}
                    }
                }
            }
        }
    if($cleanNumber -eq $null){$cleanNumber = $pDirtyNumber}
    $cleanNumber
    }
function get-3lettersInBrackets($stringMaybeContaining3LettersInBrackets,$verboseLogging){
    if($stringMaybeContaining3LettersInBrackets -match '\([a-zA-Z]{3}\)'){
        $Matches[0].Replace('(',"").Replace(')',"")
        if($verboseLogging){Write-Host -ForegroundColor DarkCyan "[$($Matches[0])] found in $stringMaybeContainingEngagementCode"}
        }
    else{if($verboseLogging){Write-Host -ForegroundColor DarkCyan "3 letters in brackets not found in $stringMaybeContainingEngagementCode"}}
    }
function get-3letterIsoCodeFromCountryName($pCountryName){
    switch ($pCountryName) {
        {@("UAE","UE","AE","ARE","United Arab Emirates","Dubai") -contains $_} {"ARE"}
        {@("BR","BRA","Brazil","Brasil") -contains $_} {"BRA"}
        {@("CA","CAN","Canada","Canadia") -contains $_} {"CAN"}
        {@("CN","CHN","China") -contains $_} {"CHN"}
        {@("DE","DEU","GE","GER","Germany","Deutschland","Deutchland") -contains $_} {"DEU"}
        {@("ES","ESP","SP","SPA","Spain","España","Espania") -contains $_} {"ESP"}
        {@("FI","FIN","Finland","Suomen","Suomen tasavalta") -contains $_} {"FIN"}
        {@("F","FR",,"FRA","France") -contains $_} {"FRA"}
        {@("UK","GB","GBR","United Kingdom","Great Britain","Scotland","England","Wales","Northern Ireland") -contains $_} {"GBR"}
        {@("IE","IRL","IR","IER","Ireland") -contains $_} {"IRL"}
        {@("PH","PHL","PHI","FIL","Philippenes","Phillippenes","Philipenes","Phillipenes") -contains $_} {"IRL"}
        {@("SE","SWE","SW","SWD","Sweden","Sweeden","Sverige") -contains $_} {"SWE"}
        {@("US","USA","United States","United States of America") -contains $_} {"USA"}
        {@("IT","ITA","Italy","Italia") -contains $_} {"ITA"}
        #Add more countries
        default {}
        }
    }
function get-2letterIsoCodeFrom3LetterIsoCode($p3letterIsoCode){
    switch ($p3letterIsoCode) {
        "ARE" {"AE"}
        "CAN" {"CA"}
        "CHN" {"CN"}
        "DEU" {"DE"}
        "ESP" {"ES"}
        "FIN" {"FI"}
        "GBR" {"GB"}
        "IRL" {"IE"}
        "ITA" {"IT"}
        "PHL" {"PH"}
        "SWE" {"SE"}
        "USA" {"US"}
        #Add more countries
        default {"Unknown"}
        }
    }
function get-2letterIsoCodeFromCountryName($pCountryName){
    $3letterCode = get-3letterIsoCodeFromCountryName -pCountryName $pCountryName
    get-2letterIsoCodeFrom3LetterIsoCode -p3letterIsoCode $3letterCode
    }
function get-available365licensecount{
        [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true,ParameterSetName="LicenseType")]
            [ValidateSet("E1", "E3", "EMS", "All")]
            [string[]]$licensetype
            )
            if(![string]::IsNullOrWhiteSpace($licensetype)){
                    switch ($licensetype){
                        "E1" {
                            $availableLicenses = Get-MsolAccountSku | Where-Object -Property "AccountSkuId" -EQ "AnthesisLLC:STANDARDPACK"
                        }
                        "E3" {
                            $availableLicenses = Get-MsolAccountSku | Where-Object -Property "AccountSkuId" -EQ "AnthesisLLC:ENTERPRISEPACK"
                        }
                        "EMS"{
                            $availableLicenses = Get-MsolAccountSku | Where-Object -Property "AccountSkuId" -EQ "AnthesisLLC:EMS"
                        }
                        "All"{
                            $availableE1Licenses = Get-MsolAccountSku | Where-Object -Property "AccountSkuId" -EQ "AnthesisLLC:STANDARDPACK"
                            $availableE3Licenses = Get-MsolAccountSku | Where-Object -Property "AccountSkuId" -EQ "AnthesisLLC:ENTERPRISEPACK"
                            $availableEMSLicenses = Get-MsolAccountSku | Where-Object -Property "AccountSkuId" -EQ "AnthesisLLC:EMS"
                            }
                        
                        }
                        If(("E1" -eq $licensetype) -or ("E3" -eq $licensetype) -or ("EMS" -eq $licensetype)){
                            Write-Host "$($licensetype)" "license count:" "$($availableLicenses.ConsumedUnits)"  "/"  "$($availableLicenses.ActiveUnits)" -ForegroundColor Yellow
                        }
                        Else{
                            Write-Host "Available E1 license count: "$($availableE1Licenses.ConsumedUnits)"  "/"  "$($availableE1Licenses.ActiveUnits)"" -ForegroundColor Yellow
                            Write-Host "Available E3 license count: "$($availableE3Licenses.ConsumedUnits)"  "/"  "$($availableE3Licenses.ActiveUnits)"" -ForegroundColor Yellow
                            Write-Host "Available EMS license count: "$($availableEMSLicenses.ConsumedUnits)"  "/"  "$($availableEMSLicenses.ActiveUnits)"" -ForegroundColor Yellow
                        }
            }
}
function get-azureAdBitlockerHeader{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $false)]
        [pscredential]$aadCreds
        )
    Write-Verbose "get-azureAdBitlockerHeader -aadCreds [$($aadCreds.UserName) | $($aadCreds.Password)]"
    #Test for connection to AzureRM
    Import-Module AzureRM.Profile
    try {    
        $context = Get-AzureRmContext -ErrorAction Stop -WarningAction Stop -InformationAction Stop
        if([string]::IsNullOrWhiteSpace($context)){throw [System.AccessViolationException] "Insuffient privileges to connect to Get-AzureRmContext"}
        }
    catch {
        connect-toAzureRm -aadCreds $aadCreds
        }
    finally {
        if([string]::IsNullOrWhiteSpace($context)){$context = Get-AzureRmContext}
        }

    #Then build header
    $tenantId = $context.Tenant.Id
    $refreshToken = @($context.TokenCache.ReadItems() | Where-Object {$_.tenantId -eq $tenantId -and $_.ExpiresOn -gt (Get-Date)})[0].RefreshToken
    $body = "grant_type=refresh_token&refresh_token=$($refreshToken)&resource=74658136-14ec-4630-ad9b-26e160ff0fc6"
    $apiToken = Invoke-RestMethod "https://login.windows.net/$tenantId/oauth2/token" -Method POST -Body $body -ContentType 'application/x-www-form-urlencoded'
    $header = @{
        'Authorization'          = 'Bearer ' + $apiToken.access_token
        'X-Requested-With'       = 'XMLHttpRequest'
        'x-ms-client-request-id' = [guid]::NewGuid()
        'x-ms-correlation-id'    = [guid]::NewGuid()
        }
    $header
    }
function get-azureAdBitLockerKeysForAllDevices{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $false)]
        [hashtable]$header
        ,[pscredential]$aadCreds
        )

    #Get Header if necessary
    if([string]::IsNullOrWhiteSpace($header)){
        $header = get-azureAdBitlockerHeader -aadCreds $aadCreds 
        }

    #Check if connected to AzureAD
    try{$allDevices = Get-AzureADDevice -All:$true -ErrorAction Stop -WarningAction Stop -InformationAction Stop}
    catch{
        connect-toAAD -credential $aadCreds
        }
    finally{
        if([string]::IsNullOrWhiteSpace($allDevices)){$allDevices = Get-AzureADDevice -All:$true -ErrorAction Stop -WarningAction Stop -InformationAction Stop}
        }

    $bitLockerKeys = @()

    foreach ($device in $allDevices) {
        $bitLockerKeysForThisDevice = get-azureADBitLockerKeysForDevice -adDevice $device -header $header -Verbose
        if(![string]::IsNullOrWhiteSpace($bitLockerKeysForThisDevice)){
            $bitLockerKeys += $bitLockerKeysForThisDevice
            }
        }
    $bitLockerKeys
    }
function get-azureAdBitLockerKeysForDevice{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [Microsoft.Open.AzureAD.Model.DirectoryObject]$adDevice
        ,[hashtable]$header
        )

    $deviceBitLockerKeys = @()
    $url = "https://main.iam.ad.ext.azure.com/api/Device/$($adDevice.objectId)"
    $deviceRecord = Invoke-RestMethod -Uri $url -Headers $header -Method Get
    if ($deviceRecord.bitlockerKey.count -ge 1) {
        $deviceBitLockerKeys += [PSCustomObject]@{
            Device      = $deviceRecord.displayName
            DriveType   = $deviceRecord.bitLockerKey.driveType
            KeyId       = $deviceRecord.bitLockerKey.keyIdentifier
            RecoveryKey = $deviceRecord.bitLockerKey.recoveryKey
            CreationTime= $deviceRecord.bitLockerKey.creationTime
            }
        }
    $deviceBitLockerKeys
    }
function get-azureAdBitLockerKeysForUser {
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [string]$SearchString
        ,[pscredential]$Credential
        )
 
    try{$userDevices = Get-AzureADUser -SearchString $SearchString | Get-AzureADUserRegisteredDevice -All:$true}
    catch{
        connect-toAAD -credential $aadCreds
        }
    finally{
        if([string]::IsNullOrWhiteSpace($userDevices)){$userDevices = Get-AzureADUser -SearchString $SearchString | Get-AzureADUserRegisteredDevice -All:$true}
        }
 
    #Get Header if necessary
    if([string]::IsNullOrWhiteSpace($header)){
        $header = get-azureAdBitlockerHeader -aadCreds $aadCreds
        }

    $bitLockerKeys = @()
    foreach ($device in $userDevices) {
        $bitLockerKeysForThisDevice = get-azureADBitLockerKeysForDevice -adDevice $device -header $header
        if(![string]::IsNullOrWhiteSpace($bitLockerKeysForThisDevice)){
            $bitLockerKeys += $bitLockerKeysForThisDevice
            }
        }

     $bitLockerKeys
    }
function get-groupAdminRoleEmailAddresses_deprecated(){
    [CmdletBinding()]
    param()
    $admins = @()
    Get-MsolRoleMember -RoleObjectId fe930be7-5e62-47db-91af-98c3a49a38b1 | % {$admins += $_.EmailAddress} #User Account Administrator
    Get-MsolRoleMember -RoleObjectId 29232cdf-9323-42fd-ade2-1d097af3e4de | % {$admins += $_.EmailAddress} #Exchange Service Administrator
    $admins | Sort-Object -Unique
    }
function get-keyFromValue($value, $hashTable){
    foreach ($Key in ($hashTable.GetEnumerator() | Where-Object {$_.Value -eq $value})){
        $Key.name}
    }
function get-keyFromValueViaAnotherKey($value, $interimKey, $hashTable){
    foreach ($Key in ($hashTable.GetEnumerator() | Where-Object {$_.Value[$interimKey] -eq $value})){
        $Key.name}
    }
function get-kimbleEngagementCodeFromString($stringMaybeContainingEngagementCode,$verboseLogging){
    if($stringMaybeContainingEngagementCode -match 'E(\d){6}'){
        $Matches[0]
        if($verboseLogging){Write-Host -ForegroundColor DarkCyan "[$($Matches[0])] found in $stringMaybeContainingEngagementCode"}
        }
    else{if($verboseLogging){Write-Host -ForegroundColor DarkCyan "Kimble Project Code not found in $stringMaybeContainingEngagementCode"}}
    }
function get-managersGroupNameFromTeamUrl($teamSiteUrl){
    if(![string]::IsNullOrWhiteSpace($teamSiteUrl)){
        $leaf = Split-Path $teamSiteUrl -Leaf
        $guess = $leaf.Replace("_","")
        if($guess.Substring($guess.Length-3,3) -eq "365"){
            $managerGuess = $guess.Substring(0,$guess.Length-3)+"-Managers"
            }
        else{
            Write-Warning "The URL [$teamSiteUrl] doesn't look like a standardised O365 Group Name - I can't guess this"
            }
        }
    $managerGuess
    }
function get-microsoftProductInfo(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
            [ValidateSet("FriendlyName","MSProductName","MSStringID","GUID")]
            [string]$getType
        ,[parameter(Mandatory = $true)]
            [ValidateSet("FriendlyName","MSProductName","MSStringID","GUID")]
            [string]$fromType
        ,[parameter(Mandatory = $true)]
            [string]$fromValue
        )
   
    #@(FriendlyName,MSName,MSStringID,GUID)
    switch($getType){
        "FriendlyName" {$getId = 0}
        "MSProductName" {$getId = 1}
        "MSStringID" {$getId = 2}
        "GUID" {$getId = 3}
        }
    switch($fromType){
        "FriendlyName" {$fromId = 0}
        "MSProductName" {$fromId = 1}
        "MSStringID" {$fromId = 2}
        "GUID" {$fromId = 3}
        }
    Write-Verbose "getId = [$getId]"
    Write-Verbose "fromId = [$fromId]"
    $productList = @(
    @("AudioConferencing","AUDIO CONFERENCING","MCOMEETADV","0c266dff-15dd-4b49-8397-2bb16070ed52"),
    @("AZURE ACTIVE DIRECTORY BASIC","AZURE ACTIVE DIRECTORY BASIC","AAD_BASIC","2b9c8e7c-319c-43a2-a2a0-48c5c6161de7"),
    @("AZURE ACTIVE DIRECTORY PREMIUM P1","AZURE ACTIVE DIRECTORY PREMIUM P1","AAD_PREMIUM","078d2b04-f1bd-4111-bbd4-b4b1b354cef4"),
    @("AZURE ACTIVE DIRECTORY PREMIUM P2","AZURE ACTIVE DIRECTORY PREMIUM P2","AAD_PREMIUM_P2","84a661c4-e949-4bd2-a560-ed7766fcaf2b"),
    @("AZURE INFORMATION PROTECTION PLAN 1","AZURE INFORMATION PROTECTION PLAN 1","RIGHTSMANAGEMENT","c52ea49f-fe5d-4e95-93ba-1de91d380f89"),
    @("DYNAMICS 365 CUSTOMER ENGAGEMENT PLAN ENTERPRISE EDITION","DYNAMICS 365 CUSTOMER ENGAGEMENT PLAN ENTERPRISE EDITION","DYN365_ENTERPRISE_PLAN1","ea126fc5-a19e-42e2-a731-da9d437bffcf"),
    @("DYNAMICS 365 FOR CUSTOMER SERVICE ENTERPRISE EDITION","DYNAMICS 365 FOR CUSTOMER SERVICE ENTERPRISE EDITION","DYN365_ENTERPRISE_CUSTOMER_SERVICE","749742bf-0d37-4158-a120-33567104deeb"),
    @("DYNAMICS 365 FOR FINANCIALS BUSINESS EDITION","DYNAMICS 365 FOR FINANCIALS BUSINESS EDITION","DYN365_FINANCIALS_BUSINESS_SKU","cc13a803-544e-4464-b4e4-6d6169a138fa"),
    @("DYNAMICS 365 FOR SALES AND CUSTOMER SERVICE ENTERPRISE EDITION","DYNAMICS 365 FOR SALES AND CUSTOMER SERVICE ENTERPRISE EDITION","DYN365_ENTERPRISE_SALES_CUSTOMERSERVICE","8edc2cf8-6438-4fa9-b6e3-aa1660c640cc"),
    @("DYNAMICS 365 FOR SALES ENTERPRISE EDITION","DYNAMICS 365 FOR SALES ENTERPRISE EDITION","DYN365_ENTERPRISE_SALES","1e1a282c-9c54-43a2-9310-98ef728faace"),
    @("DYNAMICS 365 FOR TEAM MEMBERS ENTERPRISE EDITION","DYNAMICS 365 FOR TEAM MEMBERS ENTERPRISE EDITION","DYN365_ENTERPRISE_TEAM_MEMBERS","8e7a3d30-d97d-43ab-837c-d7701cef83dc"),
    @("DYNAMICS 365 UNF OPS PLAN ENT EDITION","DYNAMICS 365 UNF OPS PLAN ENT EDITION","Dynamics_365_for_Operations","ccba3cfe-71ef-423a-bd87-b6df3dce59a9"),
    @("Security","ENTERPRISE MOBILITY + SECURITY E3","EMS","efccb6f7-5641-4e0e-bd10-b4976e1bf68e"),
    @("ENTERPRISE MOBILITY + SECURITY E5","ENTERPRISE MOBILITY + SECURITY E5","EMSPREMIUM","b05e124f-c7cc-45a0-a6aa-8cf78c946968"),
    @("EXCHANGE ONLINE (PLAN 1)","EXCHANGE ONLINE (PLAN 1)","EXCHANGESTANDARD","4b9405b0-7788-4568-add1-99614e613b69"),
    @("EXCHANGE ONLINE (PLAN 2)","EXCHANGE ONLINE (PLAN 2)","EXCHANGEENTERPRISE","19ec0d23-8335-4cbd-94ac-6050e30712fa"),
    @("EXCHANGE ONLINE ARCHIVING FOR EXCHANGE ONLINE","EXCHANGE ONLINE ARCHIVING FOR EXCHANGE ONLINE","EXCHANGEARCHIVE_ADDON","ee02fd1b-340e-4a4b-b355-4a514e4c8943"),
    @("EXCHANGE ONLINE ARCHIVING FOR EXCHANGE SERVER","EXCHANGE ONLINE ARCHIVING FOR EXCHANGE SERVER","EXCHANGEARCHIVE","90b5e015-709a-4b8b-b08e-3200f994494c"),
    @("EXCHANGE ONLINE ESSENTIALS","EXCHANGE ONLINE ESSENTIALS","EXCHANGEESSENTIALS","7fc0182e-d107-4556-8329-7caaa511197b"),
    @("EXCHANGE ONLINE ESSENTIALS","EXCHANGE ONLINE ESSENTIALS","EXCHANGE_S_ESSENTIALS","e8f81a67-bd96-4074-b108-cf193eb9433b"),
    @("Kiosk","EXCHANGE ONLINE KIOSK","EXCHANGEDESKLESS","80b2d799-d2ba-4d2a-8842-fb0d0f3a4b82"),
    @("EXCHANGE ONLINE POP","EXCHANGE ONLINE POP","EXCHANGETELCO","cb0a98a8-11bc-494c-83d9-c1b1ac65327e"),
    @("INTUNE","INTUNE","INTUNE_A","061f9ace-7d42-4136-88ac-31dc755f143f"),
    @("Microsoft 365 A1","Microsoft 365 A1","M365EDU_A1","b17653a4-2443-4e8c-a550-18249dda78bb"),
    @("Microsoft 365 A3 for faculty","Microsoft 365 A3 for faculty","M365EDU_A3_FACULTY","4b590615-0888-425a-a965-b3bf7789848d"),
    @("Microsoft 365 A3 for students","Microsoft 365 A3 for students","M365EDU_A3_STUDENT","7cfd9a2b-e110-4c39-bf20-c6a3f36a3121"),
    @("Microsoft 365 A5 for faculty","Microsoft 365 A5 for faculty","M365EDU_A5_FACULTY","e97c048c-37a4-45fb-ab50-922fbf07a370"),
    @("Microsoft 365 A5 for students","Microsoft 365 A5 for students","M365EDU_A5_STUDENT","46c119d4-0379-4a9d-85e4-97c66d3f909e"),
    @("MICROSOFT 365 BUSINESS","MICROSOFT 365 BUSINESS","SPB","cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46"),
    @("MICROSOFT 365 E3","MICROSOFT 365 E3","SPE_E3","05e9a617-0261-4cee-bb44-138d3ef5d965"),
    @("Microsoft 365 E5","Microsoft 365 E5","SPE_E5","06ebc4ee-1bb5-47dd-8120-11324bc54e06"),
    @("Microsoft 365 E3_USGOV_DOD","Microsoft 365 E3_USGOV_DOD","SPE_E3_USGOV_DOD","d61d61cc-f992-433f-a577-5bd016037eeb"),
    @("Microsoft 365 E3_USGOV_GCCHIGH","Microsoft 365 E3_USGOV_GCCHIGH","SPE_E3_USGOV_GCCHIGH","ca9d1dd9-dfe9-4fef-b97c-9bc1ea3c3658"),
    @("Microsoft 365 E5 Compliance","Microsoft 365 E5 Compliance","INFORMATION_PROTECTION_COMPLIANCE","184efa21-98c3-4e5d-95ab-d07053a96e67"),
    @("Microsoft 365 E5 Security","Microsoft 365 E5 Security","IDENTITY_THREAT_PROTECTION","26124093-3d78-432b-b5dc-48bf992543d5"),
    @("Microsoft 365 E5 Security for EMS E5","Microsoft 365 E5 Security for EMS E5","IDENTITY_THREAT_PROTECTION_FOR_EMS_E5","44ac31e7-2999-4304-ad94-c948886741d4"),
    @("Microsoft 365 F1","Microsoft 365 F1","SPE_F1","66b55226-6b4f-492c-910c-a3b7a3c9d993"),
    @("Microsoft Defender Advanced Threat Protection","Microsoft Defender Advanced Threat Protection","WIN_DEF_ATP","111046dd-295b-4d6d-9724-d52ac90bd1f2"),
    @("MICROSOFT DYNAMICS CRM ONLINE BASIC","MICROSOFT DYNAMICS CRM ONLINE BASIC","CRMPLAN2","906af65a-2970-46d5-9b58-4e9aa50f0657"),
    @("MICROSOFT DYNAMICS CRM ONLINE","MICROSOFT DYNAMICS CRM ONLINE","CRMSTANDARD","d17b27af-3f49-4822-99f9-56a661538792"),
    @("MS IMAGINE ACADEMY","MS IMAGINE ACADEMY","IT_ACADEMY_AD","ba9a34de-4489-469d-879c-0f0f145321cd"),
    @("Office 365 A5 for faculty","Office 365 A5 for faculty","ENTERPRISEPREMIUM_FACULTY","a4585165-0533-458a-97e3-c400570268c4"),
    @("Office 365 A5 for students","Office 365 A5 for students","ENTERPRISEPREMIUM_STUDENT","ee656612-49fa-43e5-b67e-cb1fdf7699df"),
    @("Office 365 Advanced Compliance","Office 365 Advanced Compliance","EQUIVIO_ANALYTICS","1b1b1f7a-8355-43b6-829f-336cfccb744c"),
    @("AdvancedSpam","Office 365 Advanced Threat Protection (Plan 1)","ATP_ENTERPRISE","4ef96642-f096-40de-a3e9-d83fb2f90211"),
    @("OFFICE 365 BUSINESS","OFFICE 365 BUSINESS","O365_BUSINESS","cdd28e44-67e3-425e-be4c-737fab2899d3"),
    @("OFFICE 365 BUSINESS","OFFICE 365 BUSINESS","SMB_BUSINESS","b214fe43-f5a3-4703-beeb-fa97188220fc"),
    @("OFFICE 365 BUSINESS ESSENTIALS","OFFICE 365 BUSINESS ESSENTIALS","O365_BUSINESS_ESSENTIALS","3b555118-da6a-4418-894f-7df1e2096870"),
    @("OFFICE 365 BUSINESS ESSENTIALS","OFFICE 365 BUSINESS ESSENTIALS","SMB_BUSINESS_ESSENTIALS","dab7782a-93b1-4074-8bb1-0e61318bea0b"),
    @("OFFICE 365 BUSINESS PREMIUM","OFFICE 365 BUSINESS PREMIUM","O365_BUSINESS_PREMIUM","f245ecc8-75af-4f8e-b61f-27d8114de5f3"),
    @("OFFICE 365 BUSINESS PREMIUM","OFFICE 365 BUSINESS PREMIUM","SMB_BUSINESS_PREMIUM","ac5cef5d-921b-4f97-9ef3-c99076e5470f"),
    @("E1","OFFICE 365 E1","STANDARDPACK","18181a46-0d4e-45cd-891e-60aabd171b4e"),
    @("OFFICE 365 E2","OFFICE 365 E2","STANDARDWOFFPACK","6634e0ce-1a9f-428c-a498-f84ec7b8aa2e"),
    @("E3","OFFICE 365 E3","ENTERPRISEPACK","6fd2c87f-b296-42f0-b197-1e91e994b900"),
    @("OFFICE 365 E3 DEVELOPER","OFFICE 365 E3 DEVELOPER","DEVELOPERPACK","189a915c-fe4f-4ffa-bde4-85b9628d07a0"),
    @("Office 365 E3_USGOV_DOD","Office 365 E3_USGOV_DOD","ENTERPRISEPACK_USGOV_DOD","b107e5a3-3e60-4c0d-a184-a7e4395eb44c"),
    @("Office 365 E3_USGOV_GCCHIGH","Office 365 E3_USGOV_GCCHIGH","ENTERPRISEPACK_USGOV_GCCHIGH","aea38a85-9bd5-4981-aa00-616b411205bf"),
    @("OFFICE 365 E4","OFFICE 365 E4","ENTERPRISEWITHSCAL","1392051d-0cb9-4b7a-88d5-621fee5e8711"),
    @("E5","OFFICE 365 E5","ENTERPRISEPREMIUM","c7df2760-2c81-4ef7-b578-5b5392b571df"),
    @("OFFICE 365 E5 WITHOUT AUDIO CONFERENCING","OFFICE 365 E5 WITHOUT AUDIO CONFERENCING","ENTERPRISEPREMIUM_NOPSTNCONF","26d45bd9-adf1-46cd-a9e1-51e9a5524128"),
    @("OFFICE 365 F1","OFFICE 365 F1","DESKLESSPACK","4b585984-651b-448a-9e53-3b10f069cf7f"),
    @("OFFICE 365 MIDSIZE BUSINESS","OFFICE 365 MIDSIZE BUSINESS","MIDSIZEPACK","04a7fb0d-32e0-4241-b4f5-3f7618cd1162"),
    @("OFFICE 365 PROPLUS","OFFICE 365 PROPLUS","OFFICESUBSCRIPTION","c2273bd0-dff7-4215-9ef5-2c7bcfb06425"),
    @("OFFICE 365 SMALL BUSINESS","OFFICE 365 SMALL BUSINESS","LITEPACK","bd09678e-b83c-4d3f-aaba-3dad4abd128b"),
    @("OFFICE 365 SMALL BUSINESS PREMIUM","OFFICE 365 SMALL BUSINESS PREMIUM","LITEPACK_P2","fc14ec4a-4169-49a4-a51e-2c852931814b"),
    @("OneDrive","ONEDRIVE FOR BUSINESS (PLAN 1)","WACONEDRIVESTANDARD","e6778190-713e-4e4f-9119-8b8238de25df"),
    @("ONEDRIVE FOR BUSINESS (PLAN 2)","ONEDRIVE FOR BUSINESS (PLAN 2)","WACONEDRIVEENTERPRISE","ed01faf2-1d88-4947-ae91-45ca18703a96"),
    @("POWER APPS PER USER PLAN","POWER APPS PER USER PLAN","POWERAPPS_PER_USER","b30411f5-fea1-4a59-9ad9-3db7c7ead579"),
    @("POWER BI FOR OFFICE 365 ADD-ON","POWER BI FOR OFFICE 365 ADD-ON","POWER_BI_ADDON","45bc2c81-6072-436a-9b0b-3b12eefbc402"),
    @("POWER BI PRO","POWER BI PRO","POWER_BI_PRO","f8a1db68-be16-40ed-86d5-cb42ce701560"),
    @("PROJECT FOR OFFICE 365","PROJECT FOR OFFICE 365","PROJECTCLIENT","a10d5e58-74da-4312-95c8-76be4e5b75a0"),
    @("PROJECT ONLINE ESSENTIALS","PROJECT ONLINE ESSENTIALS","PROJECTESSENTIALS","776df282-9fc0-4862-99e2-70e561b9909e"),
    @("PROJECT ONLINE PREMIUM","PROJECT ONLINE PREMIUM","PROJECTPREMIUM","09015f9f-377f-4538-bbb5-f75ceb09358a"),
    @("PROJECT ONLINE PREMIUM WITHOUT PROJECT CLIENT","PROJECT ONLINE PREMIUM WITHOUT PROJECT CLIENT","PROJECTONLINE_PLAN_1","2db84718-652c-47a7-860c-f10d8abbdae3"),
    @("Project","PROJECT ONLINE PROFESSIONAL","PROJECTPROFESSIONAL","53818b1b-4a27-454b-8896-0dba576410e6"),
    @("PROJECT ONLINE WITH PROJECT FOR OFFICE 365","PROJECT ONLINE WITH PROJECT FOR OFFICE 365","PROJECTONLINE_PLAN_2","f82a60b8-1ee3-4cfb-a4fe-1c6a53c2656c"),
    @("SHAREPOINT ONLINE (PLAN 1)","SHAREPOINT ONLINE (PLAN 1)","SHAREPOINTSTANDARD","1fc08a02-8b3d-43b9-831e-f76859e04e1a"),
    @("SHAREPOINT ONLINE (PLAN 2)","SHAREPOINT ONLINE (PLAN 2)","SHAREPOINTENTERPRISE","a9732ec9-17d9-494c-a51c-d6b45b384dcb"),
    @("SKYPE FOR BUSINESS CLOUD PBX","SKYPE FOR BUSINESS CLOUD PBX","MCOEV","e43b5b99-8dfb-405f-9987-dc307f34bcbd"),
    @("SKYPE FOR BUSINESS ONLINE (PLAN 1)","SKYPE FOR BUSINESS ONLINE (PLAN 1)","MCOIMP","b8b749f8-a4ef-4887-9539-c95b1eaa5db7"),
    @("SKYPE FOR BUSINESS ONLINE (PLAN 2)","SKYPE FOR BUSINESS ONLINE (PLAN 2)","MCOSTANDARD","d42c793f-6c78-4f43-92ca-e8f6a02b035f"),
    @("InternationalCalling","SKYPE FOR BUSINESS PSTN DOMESTIC AND INTERNATIONAL CALLING","MCOPSTN2","d3b4fe1f-9992-4930-8acb-ca6ec609365e"),
    @("DomesticCalling","SKYPE FOR BUSINESS PSTN DOMESTIC CALLING","MCOPSTN1","0dab259f-bf13-4952-b7f8-7db8f131b28d"),
    @("InternationalCalling","SKYPE FOR BUSINESS PSTN DOMESTIC CALLING (120 Minutes)","MCOPSTN5","54a152dc-90de-4996-93d2-bc47e670fc06"),
    @("VISIO ONLINE PLAN 1","VISIO ONLINE PLAN 1","VISIOONLINE_PLAN1","4b244418-9658-4451-a2b8-b5e2b364e9bd"),
    @("Visio","VISIO Online Plan 2","VISIOCLIENT","c5928f49-12ba-48f7-ada3-0d743a3601d5"),
    @("WINDOWS 10 ENTERPRISE E3","WINDOWS 10 ENTERPRISE E3","WIN10_PRO_ENT_SUB","cb10e6cd-9da4-4992-867b-67546b1db821"),
    @("Windows 10 Enterprise E5","Windows 10 Enterprise E5","WIN10_VDA_E5","488ba24a-39a9-4473-8ee5-19291e71b002"),
    @("PowerAutomateFree","FLOW_FREE","FLOW_FREE","f30db892-07e9-47e9-837c-80727f46fd3d")
    )
    $foundProduct = $productList | ? {$_[$fromId] -eq $fromValue} 
    $foundProduct[$getId]
    }
function get-timeZones(){
    $timeZones = Get-ChildItem "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Time zones" | foreach {Get-ItemProperty $_.PSPath}; $TimeZone | Out-Null
    $timeZones
    }
function get-timeZoneHashTable($timeZoneArray){
    if($timeZoneArray.Count -lt 1){$timeZones = get-timeZones}
        else {$timeZones = $timeZoneArray}
    $timeZoneHashTable = @{}
    $timeZones | % {$timeZoneHashTable.Add($_.PSChildName, ($_.Display.Split(" ")[0].Replace("(","").Replace(")","")))} | Out-Null
    $timeZoneHashTable.Add("","Unknown") | Out-Null
    $timeZoneHashTable
    }
function get-timeZoneSpsIdFromUnformattedTimeZone($pUnformattedTimeZone, $pTimeZoneHashTable, $pSpoTimeZoneHashTable){
    if ($pTimeZoneHashTable.Count -eq 0){$timeZoneHashTable = get-timeZoneHashTable}
        else{$timeZoneHashTable = $pTimeZoneHashTable}
    if ($pSpoTimeZoneHashTable.Count -eq 0){
        

        $spoTimeZoneHashTable = get-timeZoneHashTable
        }
        else{$spoTimeZoneHashTable = $pSpoTimeZoneHashTable}

    }
function get-trailing3LettersIfTheyLookLikeAnIsoCountryCode(){
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ambiguousString
        )
    if($ambiguousString -match ", [a-zA-Z]{3}$"){
        $ambiguousString.Substring($ambiguousString.Length-3,3)
        }
    }
function get-unformattedTimeZone ($pFormattedTimeZone){
    if ($pFormattedTimeZone -eq "" -or $pFormattedTimeZone -eq $null){"Unknown"}
    else{
        #$pFormattedTimeZone.Split("(")[1].Replace(")","").Trim()
        [regex]::Match($pFormattedTimeZone,"\(([^)]+)\)").Groups[1].Value #Get everything between "(" and ")"
        }
    }
function guess-languageCodeFromCountry($p3LetterCountryIsoCode){
    switch ($p3LetterCountryIsoCode){
        "ARE" {"en-GB"}
        "CAN" {"en-CA"}
        "CHN" {"en-US"}
        "DEU" {"de"}
        "ESP" {"es"}
        "FIN" {"fi"}
        "GBR" {"en-GB"}
        "IRL" {"en-GB"}
        "PHL" {"en-US"}
        "SWE" {"sv"}
        "USA" {"en-US"}
        }
    }
function guess-nameFromString([string]$ambiguousString){
    $lessAmbiguousString = $ambiguousString.Trim().Replace('"','')
    $leastAmbiguousString = $null
    #If it doesn't contain a space, see if it's an e-mail address
    if($lessAmbiguousString.Split(" ").Count -lt 2){
        if($lessAmbiguousString -match "@"){
            $lessAmbiguousString.Split("@")[0] | % {$_.Split(".")} | %{
                $blob = $_.Trim()
                $leastAmbiguousString += $($blob.SubString(0,1).ToUpper() + $blob.SubString(1,$blob.Length-1).ToLower()) + " " #Title Case
                }
            }
        else{$leastAmbiguousString = $lessAmbiguousString}#Do nothing - it's too weird.
        }
    else{
        if($lessAmbiguousString -match ","){#If Lastname, Firstname
            $lessAmbiguousString.Split(",") | %{
                $blob = $_.Trim()
                $leastAmbiguousString = $($blob.SubString(0,1).ToUpper() + $blob.SubString(1,$blob.Length-1).ToLower()) + " $leastAmbiguousString" #Prepend each blob as they're in reverse order
                }
            }
        else{
            $lessAmbiguousString.Split(" ") | %{ #If firstname lastname
                $blob = $_.Trim()
                $leastAmbiguousString += $($blob.SubString(0,1).ToUpper() + $blob.SubString(1,$blob.Length-1).ToLower()) + " "#Just Title Case it
                }
            }
        }
    $leastAmbiguousString.Trim()
    }
function import-encryptedCsv(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
        [string]$pathToEncryptedCsv        
        )

    $encryptedCsvData = import-csv $pathToEncryptedCsv
    $decryptedObject = New-Object psobject
    $encryptedCsvData.PSObject.Properties | ForEach-Object {
        $decryptedObject | Add-Member -MemberType NoteProperty -Name $_.Name -Value $(decrypt-SecureString -secureString $(ConvertTo-SecureString $_.Value))
        }
    $decryptedObject
    }
function log-action($myMessage, $logFile, $doNotLogToFile, $doNotLogToScreen){
    if(!$doNotLogToFile -or $logToFile){Add-Content -Value ((Get-Date -Format "yyyy-MM-dd HH:mm:ss")+"`tACTION:`t$myMessage") -Path $logFile}
    if(!$doNotLogToScreen -or $logToScreen){Write-Host -ForegroundColor Yellow $myMessage}
    }
function log-error($myError, $myFriendlyMessage, $fullLogFile, $errorLogFile, $doNotLogToFile, $doNotLogToScreen, $doNotLogToEmail, $smtpServer, $mailTo, $mailFrom){
    if(!$doNotLogToFile -or $logToFile){
        Add-Content -Value "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")`t`tERROR:`t$myFriendlyMessage" -Path $errorLogFile
        Add-Content -Value "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")`t$($myError.Exception.Message)" -Path $errorLogFile
        if($fullLogFile){
            Add-Content -Value "`t`tERROR:`t$myFriendlyMessage" -Path $fullLogFile
            Add-Content -Value "`t`t$($myError.Exception.Message)" -Path $fullLogFile
            }
        }
    if(!$doNotLogToScreen -or $logToScreen){
        Write-Host -ForegroundColor Red $myFriendlyMessage
        Write-Host -ForegroundColor Red $myError
        }
    if(!$doNotLogToEmail -or $logErrorsToEmail){
        if([string]::IsNullOrWhiteSpace($to)){$to = $env:USERNAME+"@anthesisgroup.com"}
        if([string]::IsNullOrWhiteSpace($mailFrom)){$mailFrom = $env:COMPUTERNAME+"@anthesisgroup.com"}
        if([string]::IsNullOrWhiteSpace($smtpServer)){$smtpServer= "anthesisgroup-com.mail.protection.outlook.com"}
        Send-MailMessage -To $mailTo -From $mailFrom -Subject "Error in automated script - $($myFriendlyMessage.SubString(0,20))" -Body ("$myError`r`n`r`n$myFriendlyMessage") -SmtpServer $smtpServer
        }
    }
function log-result($myMessage, $logFile, $doNotLogToFile, $doNotLogToScreen){
    if(!$doNotLogToFile -or $logToFile){Add-Content -Value ("`tRESULT:`t$myMessage") -Path $logfile}
    if(!$doNotLogToScreen -or $logToScreen){Write-Host -ForegroundColor DarkYellow "`t$myMessage"}
    }
function matchContains($term, $arrayOfStrings){
    # Turn wildcards into regexes
    # First escape all characters that might cause trouble in regexes (leaving out those we care about)
    $escaped = $arrayOfStrings -replace '[ #$()+.[\\^{]','\$&' # list taken from Regex.Escape
    # replace wildcards with their regex equivalents
    $regexes = $escaped -replace '\*','.*' -replace '\?','.'
    # combine them into one regex
    $singleRegex = ($regexes | %{ '^' + $_ + '$' }) -join '|'

    # match against that regex
    $term -match $singleRegex
    }
function remove-diacritics{
    PARAM ([string]$String)
    [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($String))
    }
function remove-doubleQuotesFromCsv(){
    param (
        [Parameter(Mandatory=$true)]
        [string]
        $inputFile,

        [string]
        $outputFile
        )

    if (-not $outputFile){
        $outputFile = $inputFile
        }

    $inputCsv = Import-Csv $inputFile
    $quotedData = $inputCsv | ConvertTo-Csv -NoTypeInformation
    $outputCsv = $quotedData | % {$_ -replace  `
        '\G(?<start>^|,)(("(?<output>[^,"]*?)"(?=,|$))|(?<output>".*?(?<!")("")*?"(?=,|$)))' `
        ,'${start}${output}'}
    $outputCsv | Out-File $outputFile -Encoding utf8 -Force
    }
function sanitise-forJson(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
        [string]$dirtyString
        )
    $cleanString = $dirtyString.Replace('"','\"')
    $cleanString
    }
function sanitise-forMicrosoftEmailAddress(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
        [string]$dirtyString
        )
    $cleanString = $dirtyString -creplace '[^a-zA-Z0-9_@\-\.]+', ''
    do{$cleanString = $cleanString.Replace("..",".")}
    While($cleanString -match "\.\.")
    $cleanString = $cleanString.Trim(".")
    $cleanString = $cleanString.Replace(".@","@")
    $cleanString = $cleanString.Replace("@.","@")
    $cleanString
    }
function sanitise-forPnpSharePoint($dirtyString){ 
    if([string]::IsNullOrWhiteSpace($dirtyString)){return}
    $cleanerString = sanitise-forSharePointStandard -dirtyString $dirtyString
    $cleanerString.Replace(":","").Replace("/","")
    if(@("."," ") -contains $dirtyString.Substring(($dirtyString.Length-1),1)){$dirtyString = $dirtyString.Substring(0,$dirtyString.Length-1)} #Trim trailing "."
    }
function sanitise-forSharePointStandard($dirtyString){
    $dirtyString = $dirtyString.Trim()
    $dirtyString = $dirtyString.Replace(" "," ") #Weird instance where a space character is not a space character...
    if(@("."," ") -contains $dirtyString.Substring(($dirtyString.Length-1),1)){$dirtyString = $dirtyString.Substring(0,$dirtyString.Length-1)} #Trim trailing "."
    $dirtyString.Replace("`"","").Replace("#","").Replace("%","").Replace("?","").Replace("<","").Replace(">","").Replace("\","").Replace("...","").Replace("..","").Replace("'","`'").Replace("`t","").Replace("`r","").Replace("`n","").Replace("*","")
    }
function sanitise-LibraryNameForUrl($dirtyString){
    $cleanerString = $dirtyString.Trim()
    $cleanerString = $dirtyString -creplace '[^a-zA-Z0-9 _/]+', ''
    $cleanerString
    }
function sanitise-forSharePointListName($dirtyString){ 
    $cleanerString = sanitise-forSharePointStandard $dirtyString
    $cleanerString.Replace("/","")
    }
function sanitise-forSharePointFileName($dirtyString){ 
    $cleanerString = sanitise-forSharePointStandard $dirtyString
    $cleanerString.Replace("/","").Replace(":","")
    }
function sanitise-forSharePointFileName2($dirtyString){ 
    $dirtyString = $dirtyString.Trim()
    $dirtyString.Replace("`"","").Replace("#","").Replace("%","").Replace("?","").Replace("<","").Replace(">","").Replace("\","").Replace("/","").Replace("...","").Replace("..","").Replace("'","`'")
    if(@("."," ") -contains $dirtyString.Substring(($dirtyString.Length-1),1)){$dirtyString = $dirtyString.Substring(0,$dirtyString.Length-1)} #Trim trailing "."
    }
function sanitise-forSharePointGroupName($dirtyString){ 
    #"The group name is empty, or you are using one or more of the following invalid characters: " / \ [ ] : | < > + = ; , ? * ' @"
    $dirtyString = $dirtyString.Trim()
    $dirtyString.Replace("`"_","_").Replace("/","_").Replace("\","_").Replace("[","_").Replace("]","_").Replace(":","_").Replace("|","_").Replace("<","_").Replace(">","_").Replace("+","_").Replace("=","_").Replace(";","_").Replace(",","_").Replace("?","_").Replace("*","_").Replace("`'","_").Replace("@","_")
    if(@("."," ") -contains $dirtyString.Substring(($dirtyString.Length-1),1)){$dirtyString = $dirtyString.Substring(0,$dirtyString.Length-1)} #Trim trailing "."
    }
function sanitise-forSharePointFolderPath($dirtyString){ 
    $cleanerString = sanitise-forSharePointStandard $dirtyString
    $cleanerString.Replace(":","")
    }
function sanitise-forSharePointUrl($dirtyString){ 
    $dirtyString = $dirtyString.Trim()
    $dirtyString = $dirtyString.Replace(" "," ") #Weird instance where a space character is not a space character...
    $dirtyString = $dirtyString -creplace '[^a-zA-Z0-9 _/]+', ''
    #$dirtyString = $dirtyString.Replace("`"","").Replace("#","").Replace("%","").Replace("?","").Replace("<","").Replace(">","").Replace("\","/").Replace("//","/").Replace(":","")
    #$dirtyString = $dirtyString.Replace("$","`$").Replace("``$","`$").Replace("(","").Replace(")","").Replace("-","").Replace(".","").Replace("&","").Replace(",","").Replace("'","").Replace("!","")
    $cleanString =""
    for($i= 0;$i -lt $dirtyString.Split("/").Count;$i++){ #Examine each virtual directory in the URL
        if($i -gt 0){$cleanString += "/"}
        if($dirtyString.Split("/")[$i].Length -gt 50){$tempString = $dirtyString.Split("/")[$i].SubString(0,50)} #Truncate long folder names to 50 characters
            else{$tempString = $dirtyString.Split("/")[$i]}
        if($tempString.Length -gt 0){
            if(@(".", " ") -contains $tempString.Substring(($tempString.Length-1),1)){$tempString = $tempString.Substring(0,$tempString.Length-1)} #Trim trailing "." and " ", even if this results in a truncation <50 characters
            }
        $cleanString += $tempString
        }
    $cleanString = $cleanString.Replace("//","/").Replace("https/","https://") #"//" is duplicated to catch trailing "/" that might now be duplicated. https is an exception that needs specific handling
    $cleanString
    }
function sanitise-forResourcePath($dirtyString){
    if($dirtyString.Length -gt 0){
        if(@("."," ") -contains $dirtyString.Substring(($dirtyString.Length-1),1)){$dirtyString = $dirtyString.Substring(0,$dirtyString.Length-1)} #Trim trailing "."
        $dirtyString = $dirtyString.trim().replace("`'","`'`'")
        $dirtyString = $dirtyString.replace("#","").replace("%","") #As of 2017-05-26, these characters are not supported by SharePoint (even though https://msdn.microsoft.com/en-us/library/office/dn450841.aspx suggests they should be)
        #$dirtyString = $dirtyString -creplace "[^a-zA-Z0-9 _/()`'&-@!]+", '' #No need to strip non-standard characters
        #[uri]::EscapeUriString($dirtyString) #No need to encode the URL
        $dirtyString
        }
    }
function sanitise-forSql([string]$dirtyString){
    if([string]::IsNullOrWhiteSpace($dirtyString)){}
    else{$dirtyString.Replace("'","`'`'").Replace("`'`'","`'`'")}
    }
function sanitise-forSqlValue{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [ValidateSet(“String”,”Int”,”Decimal”,"Boolean","Guid","Date","HTML")] 
        [string]$dataType

        ,[parameter(Mandatory = $false)]
        $value
        )
    switch($dataType){
        "String" {"`'$(smartReplace -mysteryString $value -findThis "'" -replaceWithThis "''")`'"}
        "HTML"   {"`'$(sanitise-forSqlValue -value $(sanitise-stripHtml $value ) -dataType String)`'"}
        "Int"    {if([string]::IsNullOrWhiteSpace($value)){"0"}else{$value}}
        "Decimal"{if([string]::IsNullOrWhiteSpace($value)){"0.0"}else{$value}}
        "Boolean"{if($value -eq $true){"1"}else{"0"}}
        "Guid"   {if([string]::IsNullOrWhiteSpace($value)){"NULL"}else{"`'$value`'"}} #This could be handled better
        "Date"   {if([string]::IsNullOrWhiteSpace($value)){"NULL"}else{"`'"+$(Get-Date (smartReplace -mysteryString $value -findThis "+0000" -replaceWithThis "") -Format s)+"`'"}}
        }
    }
function sanitise-forTermStore($dirtyString){
    #$dirtyString.Replace("\t", " ").Replace(";", ",").Replace("\", "\uFF02").Replace("<", "\uFF1C").Replace(">", "\uFF1E").Replace("|", "\uFF5C")
    $cleanerString = $dirtyString.Replace("`t", "").Replace(";", "").Replace("\", "").Replace("<", "").Replace(">", "").Replace("|", "").Replace("＆","&").Replace(" "," ").Trim()
    if($cleanerString.Length -gt 255){$cleanerString.Substring(0,254)}
    else{$cleanerString}
    }
function sanitise-stripHtml($dirtyString){
    if(![string]::IsNullOrWhiteSpace($dirtyString)){
        $cleanString = $dirtyString -replace '<[^>]+>',''
        $cleanString = [System.Web.HttpUtility]::HtmlDecode($cleanString)# -replace '&amp;','&'
        $cleanString
        }
    }
function set-suffixAndMaxLength(){
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory =$true)]
        [string]$string
        ,[Parameter(Mandatory =$false)]
        [string]$suffix
        ,[Parameter(Mandatory =$true)]
        [int]$maxLength
        )
    if($string.Length -gt ($maxLength-$suffix.length)){
        $outString = $string.Substring(0,$maxLength-$suffix.length)+$suffix
        }
    else{$outString = $string+$suffix}
    $outString
    }
function smartReplace($mysteryString,$findThis,$replaceWithThis){
    if([string]::IsNullOrEmpty($mysteryString)){$result = $mysteryString}
    else{$result = $mysteryString.ToString().Replace($findThis,$replaceWithThis)}
    $result
    }
function stringify-hashTable($hashtable,$interlimiter,$delimiter){
    if([string]::IsNullOrWhiteSpace($interlimiter)){$interlimiter = ":"}
    if([string]::IsNullOrWhiteSpace($delimiter)){$delimiter = ", "}
    if($hashtable.Count -gt 0){
        $dirty = $($($hashtable.Keys | % {$_+"$interlimiter"+$hashtable[$_]+"$delimiter"}) -join "`r")
        $dirty.Substring(0,$dirty.Length-$delimiter.length)
        }
    }
#endregion
