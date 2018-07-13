Import-Module _CSOM_Library-SPO
$csomCreds = new-csomCredentials

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
$allWebs = Get-SPOWebs -Url "https://anthesisllc.sharepoint.com/sites/external" -Credential $csomCreds
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


#Get all groups in SiteCollection
$ctx = new-csomContext -fullSitePath ($webUrl+$siteCollection) -sharePointCredentials $csomCreds
$groups = $ctx.Web.SiteGroups
$ctx.Load($groups)
$ctx.ExecuteQuery()
$groups.GetEnumerator() | % {
    $group = $_
    $ctx.Load($group)
    $ctx.ExecuteQuery()
    [array]$allGroups += $group
    }

$allGroups | %{$_.Title}
$allGroups | %{
    $_.Title
    $owner = $ctx.web.SiteGroups.GetByName($($_.Title.Replace(" Members"," Owners").Replace(" Visitors"," Owners")))
    $ctx.Load($owner)
    $_.Owner = $owner
    $_.Update()
    $ctx.ExecuteQuery()
    }


$allWebs | %{
    $ctx = new-csomContext -fullSitePath $_.url -sharePointCredentials $csomCreds
    $_.Title
    set-SPOGroupAsDefault -credentials $credentials -webUrl $webUrl -siteCollection $siteCollection -sitePath $("/"+$_.Url.Split("/")[$_.Url.Split("/").Count-1]) -groupName "$($_.Title) Visitors" -defaultForWhat "Visitors"
    set-SPOGroupAsDefault -credentials $credentials -webUrl $webUrl -siteCollection $siteCollection -sitePath $("/"+$_.Url.Split("/")[$_.Url.Split("/").Count-1]) -groupName "$($_.Title) Members" -defaultForWhat "Members"
    set-SPOGroupAsDefault -credentials $credentials -webUrl $webUrl -siteCollection $siteCollection -sitePath $("/"+$_.Url.Split("/")[$_.Url.Split("/").Count-1]) -groupName "$($_.Title) Owners" -defaultForWhat "Owners"
    }