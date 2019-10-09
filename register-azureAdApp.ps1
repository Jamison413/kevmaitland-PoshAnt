Add-Type -AssemblyName System.Web
$clientIDEncoded = [System.Web.HttpUtility]::UrlEncode($clientid)
$clientSecretEncoded = [System.Web.HttpUtility]::UrlEncode($secret)
$redirectUriEncoded =  [System.Web.HttpUtility]::UrlEncode($redirect)
$resourceEncoded = [System.Web.HttpUtility]::UrlEncode($resource)
$scopeEncoded = [System.Web.HttpUtility]::UrlEncode("https://outlook.office.com/user.readwrite.all")


# Function to popup Auth Dialog Windows Form
Function Get-AuthCode {
Add-Type -AssemblyName System.Windows.Forms
$form = New-Object -TypeName System.Windows.Forms.Form -Property @{Width=440;Height=640}
$web  = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width=420;Height=600;Url=($url -f ($Scope -join "%20")) }
$DocComp  = {
$Global:uri = $web.Url.AbsoluteUri        
if ($Global:uri -match "error=[^&]*|code=[^&]*") {$form.Close() }
}
$web.ScriptErrorsSuppressed = $true
$web.Add_DocumentCompleted($DocComp)
$form.Controls.Add($web)
$form.Add_Shown({$form.Activate()})
$form.ShowDialog() | Out-Null
$queryOutput = [System.Web.HttpUtility]::ParseQueryString($web.Url.Query)
$output = @{}
foreach($key in $queryOutput.Keys){
$output["$key"] = $queryOutput[$key]
}
$output
}
# Get AuthCode
$url = "https://login.microsoftonline.com/common/oauth2/authorize?response_type=code&redirect_uri=$redirectUriEncoded&client_id=$clientID&resource=$resourceEncoded&prompt=admin_consent&scope=$scopeEncoded"
Get-AuthCode
# Extract Access token from the returned URI
$regex = '(?<=code=)(.*)(?=&)'
$authCode  = ($uri | Select-string -pattern $regex).Matches[0].Value
Write-output "Received an authCode, $authCode"
#get Access Token
$body = "grant_type=authorization_code&redirect_uri=$redirect&client_id=$clientId&client_secret=$clientSecretEncoded&code=$authCode&resource=$resource"
$tokenResponse = Invoke-RestMethod https://login.microsoftonline.com/common/oauth2/token `
-Method Post -ContentType "application/x-www-form-urlencoded" `
-Body $body `
-ErrorAction STOP
