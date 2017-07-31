$antUsersCsv = import-csv ".\Desktop\AntStaffList_manual_3 - copy.csv" -Encoding UTF7 -Delimiter ","
$anthesisAdmin = "kevin.maitland@anthesisgroup.com"
$anthesisPassword = Read-Host -Prompt "Enter password for $anthesisAdmin" -AsSecureString
$credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $anthesisAdmin, $anthesisPassword
$anthesisMySite = 'https://anthesisllc-my.sharepoint.com/' # This needs to be the mySite where the userdata lives.
$anthesisAdminSite = 'https://anthesisllc-admin.sharepoint.com/' # This needs to be the "admin" site.

Import-Module MSOnline
Connect-MsolService -Credential $credential

$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
Import-PSSession $ExchangeSession

# Download and install this: http://www.microsoft.com/en-us/download/details.aspx?id=42038
Import-Module 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.UserProfiles.dll' #CSOM for SPO User Profiles
Import-Module 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll' #CSOM for SharePoint Online
$anthesisSharePointCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($anthesisAdmin, $anthesisPassword)
#Get the Client Context and Bind the Site Collection for the MySites first to get the full list of users
$anthesisContext = New-Object Microsoft.SharePoint.Client.ClientContext($anthesisAdminSite)
$anthesisContext.Credentials = $anthesisSharePointCredentials
$anthesisPeople = New-Object Microsoft.SharePoint.Client.UserProfiles.PeopleManager($anthesisContext)

function convert-DisplayNameToUpn($displayName){
    (Remove-Diacritics ($displayName.Trim().Replace(" ",".").Replace("'","")+"@anthesisgroup.com"))
    }
function sanitise-phoneNumber($dirtyString){
    $dirtyString.Trim() -replace '[^a-z0-9+() ]+',''
    }
function convert-displayNameToSpoAccountName($displayName){
    "i:0#.f|membership|"+(convert-DisplayNameToUpn $displayName)
    }
function Remove-Diacritics {
    param ([String]$src = [String]::Empty)
    $normalized = $src.Normalize( [Text.NormalizationForm]::FormD )
    $sb = new-object Text.StringBuilder
    $normalized.ToCharArray() | % { 
        if( [Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {[void]$sb.Append($_)}
        }
    $sb.ToString()
    }
$badUsers = ,@()
foreach($antUser in $antUsersCsv){
    write-host -f Yellow $antUser.Employee
    try{$validUser = $null
        Get-MsolUser -UserPrincipalName ($antUser.UPN.Trim()) | Out-Null
        $validUser = $true
        }
    catch{$badUsers += $antUser;Write-Host -f Red "Failed to look up $($antUser.Employee.Trim())`r`n$($Error[0])"}
    if ($validUser){
        try{
            Set-MsolUser -UserPrincipalName ($antUser.upn.Trim()) -Title $antUser.Role.Trim() -Department $antUser.Community.Trim() -Country $antUser.Country.Trim() -City $antUser.'Principal Location'.Trim() -Office $antUser.'Nearest office'.Trim() -MobilePhone ($antUser.'Workmobile') -PhoneNumber ($antUser.DD)
            Set-Mailbox -Identity $antUser.upn  -CustomAttribute1 $antUser.TradingEntity
            }
        catch{$badUsers += $antUser;Write-Host -f Red "Failed to set data for $($antUser.Employee.Trim())`r`n$($Error[0])"}}}
        if($antUser.'Employee Line Manager' -ne ""){#
            try{
                $manager = $anthesisContext.Web.EnsureUser((convert-DisplayNameToUpn $antUser.'Employee Line Manager'))
                $anthesisContext.Load($manager)
                $anthesisContext.ExecuteQuery()
                $anthesisPeople.SetSingleValueProfileProperty("i:0#.f|membership|"+($antUser.upn),"Manager", $manager.LoginName)
                $anthesisContext.ExecuteQuery()
                }
            catch{$badUsers += $antUser;Write-Host -f Red "Failed to add Line Manager $($manager.LoginName.Trim()) for $($antUser.Employee.Trim())`r`n$($Error[0])"}
            }
        if($antUser.'Principal Community Lead' -ne ""){
            try{
                $dottedLineManager = $anthesisContext.Web.EnsureUser((convert-displayNameToSpoAccountName $antUser.'Principal Community Lead'))
                $anthesisContext.Load($dottedLineManager)
                $anthesisContext.ExecuteQuery()
                $anthesisPeople.SetSingleValueProfileProperty(("i:0#.f|membership|"+$antUser.upn),"SPS-Dotted-Line", $dottedLineManager.LoginName)
                $anthesisContext.ExecuteQuery()
                }
            catch{$badUsers += $antUser;Write-Host -f Red "Failed to add Dotted Manager $($manager.LoginName) for $($antUser.Employee)`r`n$($Error[0])"}
            }
       
        }
    }

    convert-DisplayNameToUpn "Tore Söderqvist"