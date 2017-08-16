Import-Module _CSOM_Library-SPO
$csomCreds = set-csomCredentials

function Get-SPOWebs(){
param(
   $Url = $(throw "Please provide a Site Collection Url"),
   $Credential = $(throw "Please provide a Credentials")
)

  $context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)  
  $context.Credentials = $Credential 
  $web = $context.Web
  $context.Load($web)
  $context.Load($web.Webs)
  $context.ExecuteQuery()
  foreach($web in $web.Webs)
  {
       Get-SPOWebs -Url $web.Url -Credential $Credential 
       $web
  }
}

$dummy = "https://anthesisllc.sharepoint.com/sites/external/fishwick-ikea" 
$allWebs = Get-SPOWebs -Url "https://anthesisllc.sharepoint.com/teams/communities" -Credential $csomCreds
$AllWebs | %{ Write-Host $_.Title }

$AllWebs | %{ 
    $ctx = new-csomContext -fullSitePath $_.url -sharePointCredentials $csomCreds
    $ctx = new-csomContext -fullSitePath $dummy -sharePointCredentials $csomCreds
    $list = $ctx.Web.Lists.GetByTitle("Access Requests")
    $ctx.Load($list)
    #$list.BreakRoleInheritance($true, $true)
    $list.ResetRoleInheritance()
    $list.Update()
    $ctx.ExecuteQuery()
    Write-Output $list.ID
    $ctx.Dispose()
    }

