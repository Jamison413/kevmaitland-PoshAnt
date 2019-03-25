Import-Module AzureRM.Profile
if (Get-Module -Name "AzureADPreview" -ListAvailable) {
    Import-Module AzureADPreview
    } elseif (Get-Module -Name "AzureAD" -ListAvailable) {
    Import-Module AzureAD
    }

if ($msolCredentials) {
    Try {
        Connect-AzureAD -Credential $msolCredentials -ErrorAction Stop | Out-Null
        } 
    Catch {
        Write-Warning "Couldn't connect to Azure AD non-interactively, trying interactively."
        Connect-AzureAD -TenantId $(($msolCredentials.UserName.Split("@"))[1]) -ErrorAction Stop | Out-Null
        }
    
    Try {
        Login-AzureRmAccount -Credential $msolCredentials -ErrorAction Stop | Out-Null
        } 
    Catch {
        Write-Warning "Couldn't connect to Azure RM non-interactively, trying interactively."
        Login-AzureRmAccount -TenantId $(($msolCredentials.UserName.Split("@"))[1]) -ErrorAction Stop | Out-Null
        }
    }
else {
    Connect-AzureAD -ErrorAction Stop | Out-Null
    Login-AzureRmAccount -ErrorAction Stop | Out-Null
    }

$context = Get-AzureRmContext
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

#$userDevices = Get-AzureADUser -SearchString $SearchString | Get-AzureADUserRegisteredDevice -All:$true
$allDevices = Get-AzureADDevice -All:$true

$bitLockerKeys = @()

foreach ($device in $allDevices) {
    $url = "https://main.iam.ad.ext.azure.com/api/Device/$($device.objectId)"
    $deviceRecord = Invoke-RestMethod -Uri $url -Headers $header -Method Get
    if ($deviceRecord.bitlockerKey.count -ge 1) {
        $bitLockerKeys += [PSCustomObject]@{
            Device      = $deviceRecord.displayName
            DriveType   = $deviceRecord.bitLockerKey.driveType
            KeyId       = $deviceRecord.bitLockerKey.keyIdentifier
            RecoveryKey = $deviceRecord.bitLockerKey.recoveryKey
            CreationTime= $deviceRecord.bitLockerKey.creationTime
            }
        }
    }

$bitLockerKeys
