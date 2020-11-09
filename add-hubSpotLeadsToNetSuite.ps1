$apiKey = get-hubSpotApiKey

#$contactProperties = get-hubSpotObjectProperties -apiKey $apiKey.HubApiKey -objectType contacts
#$companyProperties = get-hubSpotObjectProperties -apiKey $apiKey.HubApiKey -objectType companies
$netSuiteProductionParams = get-netSuiteParameters -connectTo Production
#$metadata = get-netSuiteMetadata -netsuiteParameters $netSuiteProductionParams
$clientSectorValues = get-netSuiteCustomListValues -objectType clientSector -netsuiteParameters $netSuiteProductionParams
$clientStatusValues = get-netSuiteCustomListValues -objectType customerstatus -netsuiteParameters $netSuiteProductionParams
$clientTypeValues = get-netSuiteCustomListValues -objectType clientType -netsuiteParameters $netSuiteProductionParams
$clientRatingValues = get-netSuiteCustomListValues -objectType clientRating -netsuiteParameters $netSuiteProductionParams

$filterCompanyFlaggedForSync = @{
    propertyName="netsuite_sync_company_"
    operator="HAS_PROPERTY"
    }
$filterCompanyHasNoNetsuiteId = @{
    propertyName="netsuiteid"
    operator="NOT_HAS_PROPERTY"
    }

$companiesToCreate = get-hubspotObjects -apiKey $apiKey.HubApiKey -objectType companies -filterGroup1 @{filters=@($filterCompanyFlaggedForSync,$filterCompanyHasNoNetsuiteId)} -firstPageOnly

$companiesToCreate | Select-Object | % {
    $thisCompanyToCreate = $_
    $thisCompanyToCreate 
    }












$test = Invoke-RestMethod -Method Get -Uri "https://api.hubapi.com/crm/v3/objects/contacts?limit=10&archived=false&hapikey=$($apiKey.HubApiKey)"

$test  = invoke-hubSpotGet -apiKey $apiKey.HubApiKey -query "/objects/contacts?archived=false" -Verbose -pageSize 100
$test2 = invoke-hubSpotPost -apiKey $apiKey.HubApiKey -query "/objects/contacts/search?archived=false" -bodyHashtable $filterGroups -Verbose -pageSize 10
$test3 = get-hubspotContacts -apiKey $apiKey.HubApiKey -filterGroup1 @{filters=@($filter2)} -filterGroup2 @{filters=@($filter3)}
$test4 = get-hubspotRecords -apiKey $apiKey.HubApiKey -filterGroup1 @{filters=@($filter4)} -objectType companies
$test3 = Invoke-RestMethod -Uri "https://api.hubapi.com/crm/v3/objects/contacts/search?limit=10&archived=false`&hapikey=$apiKey" -ContentType "application/json; charset=utf-8" -Method POST -Body $(ConvertTo-Json -InputObject $filterGroups -Depth 10)



$url = "https://app.hubspot.com/oauth/authorize?client_id=d549e97c-1e76-49be-980f-f1d4dc35c7d0&redirect_uri=https://anthesisgroup.com/antsuite&scope=contacts%20content%20oauth%20integration-sync"
function Show-OAuth2AuthCodeWindow {
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory = $true, Position = 0, HelpMessage = "The OAuth2 authorization code URL pointing towards the oauth2/v2.0/authorize endpoint as documented here: https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-auth-code-flow")]
    [System.Uri] $URL
  )
  try {

    # create an Internet Explorer object to display the OAuth 2 authorization code browser window to authenticate
    $InternetExplorer = New-Object -ComObject InternetExplorer.Application
    $InternetExplorer.Width = "600"
    $InternetExplorer.Height = "500"
    $InternetExplorer.AddressBar = $false # disable the address bar
    $InternetExplorer.ToolBar = $false # disable the tool bar
    $InternetExplorer.StatusBar = $false # disable the status bar

    # store the Console Window Handle (HWND) of the created Internet Explorer object
    $InternetExplorerHWND = $InternetExplorer.HWND

    # make the browser window visible and navigate to the OAuth2 authorization code URL supplied in the $URL parameter
    $InternetExplorer.Navigate($URL)

    # give Internet Explorer some time to start up
    Start-Sleep -Seconds 1

    # get the Internet Explorer window as application object
    $InternetExplorerWindow = (New-Object -ComObject Shell.Application).Windows() | Where-Object {($_.LocationURL -match "(^https?://.+)") -and ($_.HWND -eq $InternetExplorerHWND)}

    # wait for the URL of the Internet Explorer window to hold the OAuth2 authorization code after a successful authentication and close the window
    while (($InternetExplorerWindow = (New-Object -ComObject Shell.Application).Windows() | Where-Object {($_.LocationURL -match "(^https?://.+)") -and ($_.HWND -eq $InternetExplorerHWND)})) {
      Write-Host $InternetExplorerWindow.LocationURL
      if (($InternetExplorerWindow.LocationURL).StartsWith($RedirectURI.ToString() + "?code=")) {
        $OAuth2AuthCode = $InternetExplorerWindow.LocationURL
        $OAuth2AuthCode = $OAuth2AuthCode -replace (".*code=") -replace ("&.*")
        $InternetExplorerWindow.Quit()
      }
    }

    # return the OAuth2 Authorization Code
    return $OAuth2AuthCode

  }
  catch {
    Write-Host -ForegroundColor Red "Could not create a browser window for the OAuth2 authentication"
  }
}


$filter1 = @{
    propertyName="firstname"
    operator="EQ"
    value="ΩLaura"
    }
$filter2 = @{
    propertyName="firstname"
    operator="CONTAINS_TOKEN"
    value="Rachel"
    }
$filter3 = @{
    propertyName="email"
    operator="CONTAINS_TOKEN"
    value="halford@wales"
    }

    
$filterGroup1 = @{filters=@($filter2)}
$filterGroup2 = @{filters=@($filter3)}

$filterGroups = @{
    filterGroups=@($filterGroup1,$filterGroup2)
    }
