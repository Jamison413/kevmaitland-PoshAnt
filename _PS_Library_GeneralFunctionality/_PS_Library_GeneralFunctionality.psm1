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
