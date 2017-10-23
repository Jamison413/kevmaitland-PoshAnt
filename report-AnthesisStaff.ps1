#Generate Anthesis Staff Lists
Import-Module _PS_Library_MSOL.psm1
Import-Module _CSOM_Library-SPO.psm1

$anthesisMySite = 'https://anthesisllc-my.sharepoint.com/' # This needs to be the mySite where the userdata lives.
$anthesisAdminSite = 'https://anthesisllc-admin.sharepoint.com/' # This needs to be the "admin" site.
$spoUsers = @()

#Unfuckulate it so that it can be scheduled.
if (!(Test-Path "$env:windir\System32\config\systemprofile\Desktop")){New-Item "$env:windir\System32\config\systemprofile\Desktop" -ItemType Directory}
if (!(Test-Path "$env:windir\SysWOW64\config\systemprofile\Desktop")){New-Item "$env:windir\SysWOW64\config\systemprofile\Desktop" -ItemType Directory}



#region functions
function get-displayNameFromSpoUserId($pSpoUserId, $pCompoundUserList){
    if ($pSpoUserId -ne "" -and $pSpoUserId -ne $null){
        $email = $pSpoUserId.Replace("i:0#.f|membership|","")
        $pCompoundUserList | ?{$_.Email -ieq $email} | %{$_.DisplayName}
        }
    }
function get-upnFromSpoUserId($pSpoUserId){
    if ($pSpoUserId -ne "" -and $pSpoUserId -ne $null){
        $pSpoUserId.Replace("i:0#.f|membership|","")
        }
    }
function get-formattedTimeZone ($pTimeZone){
    if ($pTimeZone -eq "" -or $pTimeZone -eq $null){"Unknown"}
    else{"$($timeZoneHashTable[$pTimeZone]) ($pTimeZone)"}
    }
function to-titleCase($dirtyString,$delimiter){
    $cleanString = ""
    $dirtyString.Split($delimiter) | % {
        $cleanString += $_.Substring(0,1).ToUpper()
        $cleanString += $_.Substring(1,$_.length-1).ToLower()
        $cleanString += $delimiter
        }
    $cleanString.SubString(0,$cleanString.length-1)
    }
function format-emailAddress($dirtyString){
    $cleanString = ""
    $cleanString += to-titleCase $dirtyString.Split("@")[0] -delimiter "."
    $cleanString += "@"
    $cleanString += $dirtyString.Split("@")[1].ToLower()
    $cleanString
    }
#endregion

$o365Creds = set-MsolCredentials
$csomCreds = new-csomCredentials -username $o365Creds.UserName -password $o365Creds.Password
connect-ToMsol -credential $o365Creds
connect-ToExo -credential $o365Creds

$userContext = new-csomContext -fullSitePath $anthesisMySite -sharePointCredentials $csomCreds
$adminContext = new-csomContext -fullSitePath $anthesisAdminSite -sharePointCredentials $csomCreds
$spoGetUsers = New-Object Microsoft.SharePoint.Client.UserProfiles.PeopleManager($userContext)
$spoSetUsers = New-Object Microsoft.SharePoint.Client.UserProfiles.PeopleManager($adminContext)

$spoUserList = $userContext.Web.SiteUsers
$userContext.Load($spoUserList)
$userContext.ExecuteQuery()

#Get the Client Context and Bind the Site Collection for the MySites first to get teh full list of users
#$context = New-Object Microsoft.SharePoint.Client.ClientContext($mySiteUrl)
#$context.Credentials = $spCredentials 
#Fetch the users in Site Collection
#$sharepointUsers = $context.Web.SiteUsers
#$context.Load($sharepointUsers)
#$context.ExecuteQuery()
#Create an Object [People Manager] to retrieve profile information
#$people = New-Object Microsoft.SharePoint.Client.UserProfiles.PeopleManager($context)

Write-host -f Yellow "Getting MSOL Users"
$msolUsers = Get-MsolUser | ?{$_.isLicensed -eq $true} | sort Country, City, DisplayName
Write-host -f Yellow "Getting SPO Users"
#ForEach($user in $sharepointUsers){
ForEach($spoUser in ($spoUserList | ? {$msolUsers.UserPrincipalName -contains $_.LoginName.Replace("i:0#.f|membership|","")})){
    $userprofile = $spoGetUsers.GetPropertiesFor($spoUser.LoginName)
    $userContext.Load($userprofile)
    $userContext.ExecuteQuery()
    #if($userprofile.AccountName -ne $null){
        $spoUsers += $userprofile
    #    } #Enumerate each user account and save it to the userProfileCollection array
    }
Write-host -f Yellow "Getting Mail Users"
$mailBoxes = Get-Mailbox #| ?{$msolUsers.UserPrincipalName -contains $_.MicrosoftOnlineServicesID}
$mailUsers = Get-User #| ?{$msolUsers.UserPrincipalName -contains $_.MicrosoftOnlineServicesID}

$timeZoneHashTable = @{"Dateline Standard Time"="UTC-12:00";"Samoa Standard Time"="UTC-11:00";"Hawaiian Standard Time"="UTC-10:00";"Alaskan Standard Time"="UTC-09:00";"Pacific Standard Time"="UTC-08:00";"Mountain Standard Time"="UTC-07:00";"Mexico Standard Time 2"="UTC-07:00";"US Mountain Standard Time"="UTC-07:00";"Central Standard Time"="UTC-06:00";"Canada Central Standard Time"="UTC-06:00";"Mexico Standard Time"="UTC-06:00";"Central America Standard Time"="UTC-06:00";"Eastern Standard Time"="UTC-05:00";"US Eastern Standard Time"="UTC-05:00";"SA Pacific Standard Time"="UTC-05:00";"Atlantic Standard Time"="UTC-04:00";"SA Western Standard Time"="UTC-04:00";"Pacific SA Standard Time"="UTC-04:00";"Newfoundland and Labrador Standard Time"="UTC-03:30";"E South America Standard Time"="UTC-03:00";"SA Eastern Standard Time"="UTC-03:00";"Greenland Standard Time"="UTC-03:00";"Mid-Atlantic Standard Time"="UTC-02:00";"Azores Standard Time"="UTC-01:00";"Cape Verde Standard Time"="UTC-01:00";"UTC Standard Time"="UTC+00:00";"UTC"="UTC+00:00";"Greenwich Standard Time"="UTC+00:00";"GMT Standard Time"="UTC+00:00";"Central Europe Standard Time"="UTC+01:00";"Central European Standard Time"="UTC+01:00";"Romance Standard Time"="UTC+01:00";"W Europe Standard Time"="UTC+01:00";"W. Europe Standard Time"="UTC+01:00";"W Central Africa Standard Time"="UTC+01:00";"E Europe Standard Time"="UTC+02:00";"Egypt Standard Time"="UTC+02:00";"FLE Standard Time"="UTC+02:00";"GTB Standard Time"="UTC+02:00";"Israel Standard Time"="UTC+02:00";"South Africa Standard Time"="UTC+02:00";"Russian Standard Time"="UTC+03:00";"Arab Standard Time"="UTC+03:00";"E Africa Standard Time"="UTC+03:00";"Arabic Standard Time"="UTC+03:00";"Iran Standard Time"="UTC+03:30";"Arabian Standard Time"="UTC+04:00";"Caucasus Standard Time"="UTC+04:00";"Transitional Islamic State of Afghanistan Standard Time"="UTC+04:30";"Ekaterinburg Standard Time"="UTC+05:00";"West Asia Standard Time"="UTC+05:00";"India Standard Time"="UTC+05:30";"Nepal Standard Time"="UTC+05:45";"Central Asia Standard Time"="UTC+06:00";"Sri Lanka Standard Time"="UTC+06:00";"N Central Asia Standard Time"="UTC+06:00";"Myanmar Standard Time"="UTC+06:30";"SE Asia Standard Time"="UTC+07:00";"North Asia Standard Time"="UTC+07:00";"China Standard Time"="UTC+08:00";"Singapore Standard Time"="UTC+08:00";"Taipei Standard Time"="UTC+08:00";"W Australia Standard Time"="UTC+08:00";"North Asia East Standard Time"="UTC+08:00";"Korea Standard Time"="UTC+09:00";"Tokyo Standard Time"="UTC+09:00";"Yakutsk Standard Time"="UTC+09:00";"AUS Central Standard Time"="UTC+09:30";"Cen Australia Standard Time"="UTC+09:30";"AUS Eastern Standard Time"="UTC+10:00";"E Australia Standard Time"="UTC+10:00";"Tasmania Standard Time"="UTC+10:00";"Vladivostok Standard Time"="UTC+10:00";"West Pacific Standard Time"="UTC+10:00";"Central Pacific Standard Time"="UTC+11:00";"Fiji Islands Standard Time"="UTC+12:00";"New Zealand Standard Time"="UTC+12:00";"Tonga Standard Time"="UTC+13:00";""="Unknown"}

$userHash = [ordered]@{}
$msolUsers | % {$userHash.Add($_.UserPrincipalName,@($_,$null,$null,$null))}
$spoUsers | % {$userHash[$_.AccountName.Replace("i:0#.f|membership|","")][1] = $_}
$mailBoxes | % {$userHash[$_.MicrosoftOnlineServicesID][2] = $_}
$mailUsers | % {$userHash[$_.MicrosoftOnlineServicesID][3] = $_}

<#
Write-host -f Yellow "Building compound Users"
Measure-Command {
    $antUsers=,@()
    for($i=0; $i -lt $msolUsers.Length;$i++){
        Write-Host -ForegroundColor DarkYellow "($i/$($msolUsers.Count)) $($msolUsers[$i].DisplayName)"
        $antUser = New-Object Object
        $antUser | Add-Member NoteProperty upn $msolUsers[$i].UserPrincipalName
        $antUser | Add-Member NoteProperty DisplayName $msolUsers[$i].DisplayName
        $antUser | Add-Member NoteProperty Title $msolUsers[$i].Title
        $antUser | Add-Member NoteProperty Country $msolUsers[$i].Country
        $antUser | Add-Member NoteProperty City $msolUsers[$i].City
        $antUser | Add-Member NoteProperty Office $msolUsers[$i].Office
        $antUser | Add-Member NoteProperty Department $msolUsers[$i].Department
        $spoUser= $null
        for($j=0; $j -lt $spoUsers.Count;$j++){if($spoUsers[$j].AccountName -contains $msolUsers[$i].UserPrincipalName){$spoUser = $spoUsers[$j];break}}
        #$spoUser = $spoUsers | ?{$_.UserProfileProperties."SPS-UserPrincipalName" -eq $msolUser.UserPrincipalName}
        $antUser | Add-Member NoteProperty Manager $spoUser.UserProfileProperties.Manager
        $antUser | Add-Member NoteProperty DottedLineManager $spoUser."SPS-Dotted-Line"
        $antUser | Add-Member NoteProperty WorkMobilePhone $msolUsers[$i].MobilePhone
        $antUser | Add-Member NoteProperty DDI $msolUsers[$i].PhoneNumber
        $antUser | Add-Member NoteProperty Email (format-emailAddress ($msolUsers[$i].ProxyAddresses | ?{$_ -cmatch "SMTP:"}).Replace("SMTP:","")
        $userRegionalData = Get-MailboxRegionalConfiguration -Identity $msolUsers[$i].UserPrincipalName 
        $antUser | Add-Member NoteProperty TimeZone $userRegionalData.TimeZone
        $antMailbox = get-mailbox -identity $msolUsers[$i].UserPrincipalName
        $antUser | Add-Member NoteProperty BusinessEntity $antMailbox.CustomAttribute1 
        $antUsers += $antUser
        }
    }
$antUsers = $antUsers | Sort-Object Country,City,DisplayName
$antUsers | select DisplayName,Country,City
#>

#endregion

#region Report Variables
$templatePath = "C:\Reports\"
$templateFile = "AnthesisGlobalStaffListTemplate.xlsx"
$outputPath = "C:\Reports\AntUsers\"
$firstRowOfUsers = 6
#endregion

#region Write the report
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
"`tOpening report template"
$workbook = $excel.Workbooks.Open("$templatePath$templateFile")
$sheet = $workbook.Sheets.Item('Anthesians')
$currentRowAnthesians = $firstRowOfUsers
$currentRowContractors = $firstRowOfUsers
$currentRowServiceAccounts = $firstRowOfUsers
$currentRowUnknowns = $firstRowOfUsers

$sheet.Range("B3") = $(Get-Date -Format u).Replace("Z","")+" (UTC)"
foreach($ant in $userHash.Keys){# | ?{$_.Community -ne "Contractors"}){
    if($userHash[$ant][0].Title -eq "" -or $userHash[$ant][0].Title -eq $null){
        $currentRow = $currentRowUnknowns
        $currentRowUnknowns++
        $sheet = $workbook.Sheets.Item('Unknown')
        }
    elseif($userHash[$ant][0].Title -eq "Generic Mailbox" -or $userHash[$ant][0].Title -eq "Service Account" -or $userHash[$ant][0].Department -eq "Unspecified"){
        $currentRow = $currentRowServiceAccounts
        $currentRowServiceAccounts++
        $sheet = $workbook.Sheets.Item('Service Accounts')
        }
    elseif($userHash[$ant][0].Department -eq "Contractors"){
        $currentRow = $currentRowContractors
        $currentRowContractors++
        $sheet = $workbook.Sheets.Item('Contractors')
        }
    else{
        $currentRow = $currentRowAnthesians
        $currentRowAnthesians++
        $sheet = $workbook.Sheets.Item('Anthesians')
        }
    Write-Host $userHash[$ant][0].DisplayName
    #$sheet.Cells.Item($currentRow,1) = $userHash[$ant][0].DisplayName
    #$sheet.Cells.Item($currentRow,2) = $userHash[$ant][0].Title
    #$sheet.Cells.Item($currentRow,3) = $userHash[$ant][0].Country
    #$sheet.Cells.Item($currentRow,4) = $userHash[$ant][0].City
    #$sheet.Cells.Item($currentRow,5) = $userHash[$ant][0].Office
    #$sheet.Cells.Item($currentRow,6) = $userHash[$ant][0].Department
    #$sheet.Cells.Item($currentRow,7) = if($userHash[$ant][1] -ne $null){if($userHash[$ant][1].UserProfileProperties.Manager -ne ""){$userHash[(get-upnFromSpoUserId $userHash[$ant][1].UserProfileProperties.Manager)][0].DisplayName}}
    #$sheet.Cells.Item($currentRow,8) = if($userHash[$ant][1] -ne $null){if($userHash[$ant][1].UserProfileProperties.'SPS-Dotted-line' -ne ""){$userHash[(get-upnFromSpoUserId $userHash[$ant][1].UserProfileProperties.'SPS-Dotted-line')][0].DisplayName}}
    #$sheet.Cells.Item($currentRow,9) = $userHash[$ant][0].MobilePhone
    #$sheet.Cells.Item($currentRow,10) = $userHash[$ant][0].PhoneNumber
    #$sheet.Cells.Item($currentRow,11) = format-emailAddress ($userHash[$ant][0].ProxyAddresses | ?{$_ -cmatch "SMTP:"}).Replace("SMTP:","")
    #if($userHash[$ant][2] -ne $null){$sheet.Cells.Item($currentRow,12) = get-formattedTimeZone (Get-MailboxRegionalConfiguration -Identity $userHash[$ant][0].UserPrincipalName).TimeZone}
    #$sheet.Cells.Item($currentRow,13) = $userHash[$ant][2].CustomAttribute1
    $sheet.Cells.Item($currentRow,14) = $(Get-MailboxStatistics -Identity $userHash[$ant][2].UserPrincipalName | select LastLogonTime).LastLogonTime
    }
$workbook.SaveAs($outputPath+"Anthesis Staff List_$((Get-Date).ToString("yyyy-MM-dd")).xlsx")
$workbook.Close($false)
$excel.Quit()
#endregion


#$msolUsers | sort LastPasswordChangeTimestamp | select Displayname, LastPasswordChangeTimestamp | export-csv -Path .\Desktop\AnthesisPasswordChanges.log
