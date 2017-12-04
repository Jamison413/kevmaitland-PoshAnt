Import-Module _PS_Library_MSOL.psm1
Import-Module _CSOM_Library-SPO.psm1

$antUsersCsv = import-csv "C:\Users\Kevin.Maitland\Desktop\BulkUpload.csv" -Delimiter "`t" #This should be a tab-delimited file (e.g. a cut-and-paste from Excel to notepad) with the following headings:
# Employee	
# Role	
# Country	
# Principal Location	
# Nearest office	
# Community / Practise Area	
# Employee Line Manager	
# Principal Community Lead	
# Work mobile	
# DD	
# Email	
# Time zone	
# Business Entity


$anthesisMySite = 'https://anthesisllc-my.sharepoint.com/' # This needs to be the mySite where the userdata lives.
$anthesisAdminSite = 'https://anthesisllc-admin.sharepoint.com/' # This needs to be the "admin" site.
$badUsers = @()
$countryToLocaleHashTable = @{"Canada"="4105";"China"="2052";"Finland"="1035";"Germany"="1031";"Korea"="1042";"Spain"="1034";"Sri Lanka"="1097";"Philippines"="13321";"Sweden"="1053";"United Arab Emirates"="";"United Kingdom"="2057";"United States"="1033"}

#region functions
function convert-UPNToSpoAccountName($upn){
    "i:0#.f|membership|"+$upn
    }
function get-formattedTimeZone ($pTimeZone, $pTimeZoneHashTable){
    if ($pTimeZone -eq "" -or $pTimeZone -eq $null){"Unknown"}
        else{
            $timeZoneHashTable = @{}
            if ($pTimeZoneHashTable.Count -eq 0){$timeZoneHashTable = get-timeZoneHashTable}
                else{$timeZoneHashTable = $pTimeZoneHashTable}
            "$($timeZoneHashTable[$pTimeZone]) ($pTimeZone)"
            }
    }
function get-languageFromCountry(){}
function get-localeFromCountry(){}
function get-spoTimeZoneHashTable(){
    $ctx = new-csomContext -fullSitePath "https://anthesisllc.sharepoint.com" -sharePointCredentials $csomCreds
    $tz = $ctx.Web.RegionalSettings.TimeZones
    $ctx.Load($tz) | Out-Null
    $tzEnum = $ctx.Web.RegionalSettings.TimeZones.GetEnumerator()
    $ctx.ExecuteQuery() | Out-Null
    $spoTimeZones = @{}
    while($tzEnum.MoveNext()){
        $spoTimeZones.Add($tzEnum.Current.Description, $tzEnum.Current.Id) | Out-Null
        }
    $spoTimeZones
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
function get-unformattedTimeZone ($pFormattedTimeZone){
    if ($pFormattedTimeZone -eq "" -or $pFormattedTimeZone -eq $null){"Unknown"}
    else{
        #$pFormattedTimeZone.Split("(")[1].Replace(")","").Trim()
        [regex]::Match($pFormattedTimeZone,"\(([^)]+)\)").Groups[1].Value #Get everything between "(" and ")"
        }
    
    }
function get-upnFromDisplayName($displayName, $listOfMsolUsers){
    $user = $listOfMsolUsers | ? {($_.DisplayName) -imatch ($displayName)}
    $user.UserPrincipalName
    }
function get-upnFromEmail($emailAddress, $listOfMsolUsers){
    $user = $listOfMsolUsers | ? {($_.ProxyAddresses) -imatch ("smtp:"+$emailAddress)}
    $user.UserPrincipalName
    }
function remove-Diacritics {
    param ([String]$src = [String]::Empty)
    $normalized = $src.Normalize( [Text.NormalizationForm]::FormD )
    $sb = new-object Text.StringBuilder
    $normalized.ToCharArray() | % { 
        if( [Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {[void]$sb.Append($_)}
        }
    $sb.ToString()
    }
function sanitise-phoneNumber($dirtyString){
    $dirtyString.Trim() -replace '[^a-z0-9+() ]+',''
    }
#endregion


$o365Creds = set-MsolCredentials
connect-ToMsol -credential $o365Creds
connect-ToExo -credential $o365Creds
$csomCreds = new-csomCredentials -username $o365Creds.UserName -password $o365Creds.Password

$userContext = new-csomContext -fullSitePath $anthesisMySite -sharePointCredentials $csomCreds
$adminContext = new-csomContext -fullSitePath $anthesisAdminSite -sharePointCredentials $csomCreds
$spoUsers = New-Object Microsoft.SharePoint.Client.UserProfiles.PeopleManager($adminContext)

$msolUsers = Get-MsolUser -all #We'll use this to validate users
$timeZones = get-timeZones #We need this to translate between the differente TimeZone formats
$spoTimeZones = get-spoTimeZoneHashTable #We need this to translate between the differente TimeZone formats

foreach($antUser in $antUsersCsv){
    write-host -f Yellow $antUser.Employee
    $antUserInO365 = $msolUsers | ? {($_.ProxyAddresses) -imatch ("smtp:"+$antUser.Email)} #Match the spreadsheet user to an AD user via their e-mail address to get their UPN
    if ($antUserInO365){
        try{#Set the easy properties on the MSOL, Mail-User and Mailboxes
            Write-Host -ForegroundColor DarkYellow "`tUPN:"$antUserInO365.UserPrincipalName
            Set-MsolUser -UserPrincipalName $antUserInO365.UserPrincipalName -Title $antUser.Role.Trim() -Department $antUser.'Community / Practise Area'.Trim() -Country $antUser.Country.Trim() -City $antUser.'Principal Location'.Trim() -Office $antUser.'Nearest office'.Trim() -MobilePhone ($antUser.'Work mobile') -PhoneNumber ($antUser.DD)
            Set-Mailbox -Identity $antUserInO365.UserPrincipalName  -CustomAttribute1 $antUser.'Business Entity'
            Set-User -Identity $antUserInO365.UserPrincipalName -Company $antUser.'Business Entity'
            if($antUser.'Time zone' -notcontains @("",$null,"Unknown")){Set-MailboxRegionalConfiguration -Identity $antUserInO365.UserPrincipalName -TimeZone (get-unformattedTimeZone -pFormattedTimeZone $antUser.'Time zone')}
            }
        catch{$badUsers += $antUser;Write-Host -f Red "Failed to set data for $($antUser.Employee.Trim())`r`n$($Error[0])"}
        if($antUser.'Employee Line Manager' -ne ""){#
            try{#Try to validate the Line Manager exists as an account in SPO (via the User Context)
                $manager = $userContext.Web.EnsureUser($(convert-UPNToSpoAccountName -upn $(get-upnFromDisplayName -displayName ($antUser.'Employee Line Manager') -listOfMsolUsers $msolUsers)))
                $userContext.Load($manager)
                $userContext.ExecuteQuery()
                try{#If they do, try to set the value for the user (via the Admin Context)
                    $spoUsers.SetSingleValueProfileProperty($(convert-UPNToSpoAccountName -upn ($antUserInO365.UserPrincipalName)),"Manager", $manager.LoginName)
                    $adminContext.ExecuteQuery()
                    Write-Host -ForegroundColor DarkYellow "`tLine Manager: $($antUser.'Employee Line Manager')"
                    }
                catch{$badUsers += $antUser;Write-Host -f Red "Failed to add Line Manager $($manager.LoginName.Trim()) for $($antUser.Employee.Trim())`r`n$($Error[0])"}
                }
            catch{$badUsers += $antUser;Write-Host -f Red "Line Manager $($antUser.'Employee Line Manager'.Trim()) is invalid for $($antUser.Employee.Trim())`r`n$($Error[0])"}
            }
        if($antUser.'Principal Community Lead' -ne ""){
            try{#Try to validate the Dotted-line Manager exists as an account in SPO (via the User Context)
                $dottedLineManager = $userContext.Web.EnsureUser($(convert-UPNToSpoAccountName -upn (get-upnFromDisplayName -displayName ($antUser.'Principal Community Lead') -listOfMsolUsers $msolUsers)))
                $userContext.Load($dottedLineManager)
                $userContext.ExecuteQuery()
                try{#If they do, try to set the value for the user (via the Admin Context)
                    $spoUsers.SetSingleValueProfileProperty($(convert-UPNToSpoAccountName ($antUserInO365.UserPrincipalName)),"Manager", $dottedLineManager.LoginName)
                    $adminContext.ExecuteQuery()
                    Write-Host -ForegroundColor DarkYellow "`tDottedLine Manager: $($antUser.'Principal Community Lead')"
                    }
                catch{$badUsers += $antUser;Write-Host -f Red "Failed to add Dotted-Line Manager $($dottedLineManager.LoginName.Trim()) for $($antUser.Employee.Trim())`r`n$($Error[0])"}
                }
            catch{$badUsers += $antUser;Write-Host -f Red "Dotted-Line Manager $($antUser.'Principal Community Lead'.Trim()) is invalid for $($antUser.Employee.Trim())`r`n$($Error[0])"}
            }
        if($antUser.'Time zone' -notcontains @("",$null,"Unknown")){
            $spoUsers.SetSingleValueProfileProperty($(convert-UPNToSpoAccountName -upn $antUserInO365.UserPrincipalName), "SPS-RegionalSettings-Initialized", $true)
            $spoUsers.SetSingleValueProfileProperty($(convert-UPNToSpoAccountName -upn $antUserInO365.UserPrincipalName), "SPS-RegionalSettings-FollowWeb", $false)
            #Getting the TimeZoneID is a massive PITA:
            $tzUnformattedName = get-unformattedTimeZone $antUser.'Time zone' #Unformat the value on the spreadsheet
            $tz = $timeZones | ?{$_.PSChildName -eq $tzUnformattedName} #Look that up in the registry list
            $tzID = $spoTimeZones[$tz.Display.replace("+00:00","")] #Then match a different property of the registry object to the SPO object
            if($tzID.Length -gt 0){$spoUsers.SetSingleValueProfileProperty($(convert-UPNToSpoAccountName -upn $antUserInO365.UserPrincipalName), "SPS-TimeZone", $tzID)}
            if($countryToLocaleHashTable[$antUser.Country].length -gt 0){$spoUsers.SetSingleValueProfileProperty($(convert-UPNToSpoAccountName -upn $antUserInO365.UserPrincipalName), "SPS-Locale", $countryToLocaleHashTable[$antUser.Country])}
            #$spoUsers.SetSingleValueProfileProperty($(convert-UPNToSpoAccountName -upn $antUserInO365.UserPrincipalName), "SPS-MUILanguages", "en-GB")
            #$spoUsers.SetSingleValueProfileProperty($(convert-UPNToSpoAccountName -upn $antUserInO365.UserPrincipalName), "SPS-CalendarType", 1)
            #$spoUsers.SetSingleValueProfileProperty($(convert-UPNToSpoAccountName -upn $antUserInO365.UserPrincipalName), "SPS-AltCalendarType", 1)
            $adminContext.ExecuteQuery()
            Write-Host -ForegroundColor DarkYellow "`tTimeZone SPO values set"
            }
        }
        else{Write-Host -ForegroundColor DarkRed "Failed to match $($antUser.Employee) to an account in MSOL"}
    }


