$o365admin = 'kevin.maitland@anthesisgroup.com'
$o365password = Read-Host "Enter Password for $o365admin" -AsSecureString
$credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $o365admin, $o365password
$duffAccounts = @("Liesa Guttmann","Varex PEC","Sophie Sapienza","Christian Walzel","Marcus Rother","Shakila Gamage","Larry Cody","Terry Wood","Barry Holt","Ian Bailey","Malcolm Paul","Sara Angrill","Deby Stabler","Jill Stoneberg","John Hennessey","Kat Pephens","Leslie Macdougall","Luis Schaeffer","Matt Dion","Stephane N'Diaye","Todd Lindbergh","Admin O365","Anna Rengstedt","AnthesisUKFinance","Arul Subra","Conflict Minerals","Cray","Czech Anthesis","DE Info","Eaton Pec","endsight","Finance Support","France","George Davey","Italy","Jae Ryu","L3 Pec","Mahmoud  Abourich","Matthew Williams","Michael Hoffmann","Microsoft ECM","Ningwei  Dong","Spain","Target Pilot","Tharaka Naga","UK Careers","UK HR","Varian Conflict Minerals","WD PEC","Michelle Langefeld","ACS Support","Andrew Hennig","Thom Schumann")

Import-Module MSOnline
Connect-MsolService -Credential $credential
$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
Import-PSSession $ExchangeSession
$msolUsers = Get-MsolUser | ?{$duffAccounts -contains $_.DisplayName} | sort Country, City, DisplayName
$mailUsers = Get-Mailbox | ?{$msolUsers.UserPrincipalName -contains $_.MicrosoftOnlineServicesID}

$userHash = [ordered]@{}
$msolUsers | % {$userHash.Add($_.UserPrincipalName,@($_,$null))}
$mailUsers | % {$userHash[$_.MicrosoftOnlineServicesID][1] = $_}

$duffUserData = @($null)*$userHash.Keys.Count+1
$duffUserData[0] = "upn,DisplayName,Community,Country,City,BusinessEntity,Created,LastPasswordChange,MailboxLastAccessed"
for ($i=0; $i -lt $duffUserData.Length-1;$i++){
    Write-Host -ForegroundColor Yellow "($($i+1)/$($duffUserData.Length-1)) $($userHash[$i][0].DisplayName)"
    $duffUserDatum = ""
    $duffUserDatum += $userHash[$i][0].UserPrincipalName+","
    $duffUserDatum += $userHash[$i][0].DisplayName+","
    $duffUserDatum += $userHash[$i][0].Department+","
    $duffUserDatum += $userHash[$i][0].Country+","
    $duffUserDatum += if($userHash[$i][0].City -ne $null){$userHash[$i][0].City.Replace(",","")+","}else{","}
    $duffUserDatum += $userHash[$i][1].CustomAttribute1+","
    $duffUserDatum += $userHash[$i][0].WhenCreated.ToString()+","
    $duffUserDatum += $userHash[$i][0].LastPasswordChangeTimestamp.ToString()+","
    $duffUserDatum += if($userHash[$i][1] -ne $null){(Get-MailboxStatistics $userHash[$i][0].UserPrincipalName).LastLogonTime.ToString()}
    $duffUserData[$i+1] = $duffUserDatum
    }
$duffUserData | Out-File -FilePath C:\Reports\AntUsers\DuffUsers4.csv -Encoding utf8