<#Required PS Modules:
Install-Module -Name Microsft.Graph -Scope CurrentUser
Install-Module -Name ExchangeOnlineManagement -RequiredVersion 3.0.0
Install-Module -Name MSAL.PS -Scope CurrentUser #Only used for testing

#The order of loading the modules affects our testing (2023-01-11 [https://stackoverflow.com/questions/72746783/get-msaltoken-the-property-authority-cannot-be-found-on-this-object])
Import-Module MSAL.PS -Force 
#Import-Module Microsoft.Graph
#Import-Module ExchangeOnlineManagement

#>
#Permissions/Roles required for script:
    #Exchange Admin (create Application Access Policy)
    #Application Admin (create new App Registration)

#Permsissions/Roles required for manual steps:
    #Global Administrator / Privileged Role Administrator ("Grant admin consent" for Application Permissions)

$appDisplayName = "Test Mailer"
$securityGroupNamingConventionForApplicationAccessPolicies = "AAP-"
$mailboxesToGrantAccessTo = @("test1@domain.com","test2@domain.com")
$mailboxesToGrantAccessTo = @("kevin.maitland@anthesisgroup.com","t1-kevin.maitland@anthesisgroup.com")

Connect-MgGraph #Authenticating account requires "Application Admin" Role
#Create the App Registration using Graph
$params = @{
    DisplayName = $appDisplayName
    RequiredResourceAccess = @{
        ResourceAppId = "00000003-0000-0000-c000-000000000000" #Graph API
        ResourceAccess = @(
            @{
                Id = "b633e1c5-b582-4048-a93e-9f11b44c7e96" #Mail.Send App Permission
                Type = "Role"
            }
            @{
                Id = "e2a3a72e-5f79-4c64-b1b1-878b674786c9" #Mail.ReadWrite App Permission
                Type = "Role"
            }
        )
    }
}
$appRegistration = New-MgApplication @params

$passwordCred = @{
   displayName = 'Created in PowerShell for testing'
   endDateTime = (Get-Date).AddMonths(6)
}
$clientSecret = Add-MgApplicationPassword -ApplicationId $appRegistration.Id -PasswordCredential $passwordCred

Connect-ExchangeOnline #Authenticating account requires "Exchange Admin" Role
#Create the Mail-Enabled Security Group using EXO cmdlets
$mailEnabledSecurityGroup = New-DistributionGroup `
    -DisplayName "$securityGroupNamingConventionForApplicationAccessPolicies$appDisplayName" `
    -Name "$securityGroupNamingConventionForApplicationAccessPolicies$appDisplayName" `
    -Notes "Mail-Enabled Security Group to manage which Mailboxes the App [$appDisplayName][$($appRegistration.id)] can read/write/send from" `
    -Type Security `
    -Members $mailboxesToGrantAccessTo


#Create the Application Access Policy using EXO cmdlets
$applicationAccessPolicy = New-ApplicationAccessPolicy `
    -AccessRight RestrictAccess `
    -AppId $appRegistration.id `
    -PolicyScopeGroupId $mailEnabledSecurityGroup.PrimarySmtpAddress `
    -Description "Restrict App [$appDisplayName][$($appRegistration.id)] access to mailboxes of members of [$($mailEnabledSecurityGroup.DisplayName)][$($mailEnabledSecurityGroup.ExternalDirectoryObjectId)]"

#Manually grant Application Permissions for new App - https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/RegisteredApps
Write-Output "Manually `"Grant admin consent`" for App Permissions (requires Global / Privileged Role Admin) - opening page in browser:"
Write-Output "https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$($appRegistration.AppId)/isMSAApp~/false"
start-Process https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$($appRegistration.AppId)/isMSAApp~/false #Authenticating account requires "Global / Privileged Role Admin" Role


#Test Mail.Send
$connectionDetails = @{
    'TenantId'     =  $(Invoke-Command -ScriptBlock {whoami /upn}).Split("@")[1]
    'ClientId'     = $($appRegistration.AppId)
    'ClientSecret' = $($clientSecret.SecretText) | ConvertTo-SecureString -AsPlainText -Force
}
$msalToken = Get-MsalToken @connectionDetails
#Reference to propreiatry Anthesis code for sending mail via the Graph API using 
$msalToken | Add-Member -MemberType NoteProperty -Name access_token -Value $msalToken.AccessToken
function invoke-graphPost(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [string]$graphQuery
        ,[parameter(Mandatory = $true)]
            [Hashtable]$graphBodyHashtable
        ,[parameter(Mandatory = $false)]
            [hashtable]$additionalHeaders
        )

    $sanitisedGraphQuery = $graphQuery.Trim("/")
    $headers = @{Authorization = "Bearer $($tokenResponse.access_token)"}
    if($additionalHeaders){
        $additionalHeaders.GetEnumerator() | %{
            $headers.Add($_.Key,$_.Value)
            }
        }
    Write-Verbose "https://graph.microsoft.com/v1.0/$sanitisedGraphQuery"
        
    $graphBodyJson = ConvertTo-Json -InputObject $graphBodyHashtable -Depth 10
    Write-Verbose $graphBodyJson
    $graphBodyJsonEncoded = [System.Text.Encoding]::UTF8.GetBytes($graphBodyJson)
    
    Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/$sanitisedGraphQuery" -Body $graphBodyJsonEncoded -ContentType "application/json; charset=utf-8" -Headers $headers -Method Post
    }
function send-graphMailMessage(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [psobject]$tokenResponse        
        ,[parameter(Mandatory = $true)]
            [ValidatePattern("@")]
            [string]$fromUpn
        ,[parameter(Mandatory = $true)]
            [ValidatePattern("@")]
            [string[]]$toAddresses
        ,[parameter(Mandatory = $false)]
            [ValidatePattern("@")]
            [string[]]$ccAddresses
        ,[parameter(Mandatory = $false)]
            [ValidatePattern("@")]
            [string[]]$bccAddresses
        ,[parameter(Mandatory = $true)]
            [string]$subject
        ,[parameter(Mandatory = $true,ParameterSetName = "text")]
            [string]$bodyText
        ,[parameter(Mandatory = $true,ParameterSetName = "HTML")]
            [string]$bodyHtml
        ,[parameter(Mandatory = $false)]
            [bool]$saveToSentItems = $true
        ,[parameter(Mandatory = $false)]
            [ValidateSet ("low","normal","high")]
            [string]$priority = "normal"
        )

    [array]$formattedToAddresses = $toAddresses | % {
        @{emailAddress=@{'address'=$_}}
        }
    [array]$formattedFromAddresses = $fromUpn | % {
        @{emailAddress=@{'address'=$_}}
        }
    $message = @{
        toRecipients = $formattedToAddresses
        subject = $subject
        importance=$priority
        #from = $formattedFromAddresses
        #sender = $formattedFromAddresses
        }

    if($ccAddresses){
        [array]$formattedCcAddresses = $ccAddresses | % {
            @{emailAddress=@{'address'=$_}}
            }
        $message.Add("ccRecipients",$formattedCcAddresses)
        }
    if($bccAddresses){
        [array]$formattedBccAddresses = $bccAddresses | % {
            @{emailAddress=@{'address'=$_}}
            }
        $message.Add("bccRecipients",$formattedBccAddresses)
        }
    if($bodyText){$message.Add("body",@{"contentType"="Text";"content"=$bodyText})}
    if($bodyHtml){$message.Add("body",@{"contentType"="HTML";"content"=$bodyHtml})}

    invoke-graphPost -tokenResponse $tokenResponse -graphQuery "/users/$fromUpn/sendMail" -graphBodyHashtable @{"message"=$message;"saveToSentItems"=$saveToSentItems}
    }
send-graphMailMessage -tokenResponse $msalToken -fromUpn $mailboxesToGrantAccessTo[0] -toAddresses kevin.maitland@sustain.co.uk -subject "Test" -bodyText "Hurrah!" -Verbose

#Test Mail.ReadWrite
##TBC

<# Remove all objects created

$(Get-MgApplication -Filter "Id eq '$($appRegistration.Id)'") | Remove-MgApplication
Get-DistributionGroup -Filter "ExternalDirectoryObjectId -eq '$($mailEnabledSecurityGroup.ExternalDirectoryObjectId)'" | Remove-DistributionGroup
Get-ApplicationAccessPolicy | ? {$_.AppId -eq $applicationAccessPolicy.AppId} | Remove-ApplicationAccessPolicy
#>


Disconnect-Graph
Disconnect-ExchangeOnline
