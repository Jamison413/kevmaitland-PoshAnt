$doNotDelete = convertTo-arrayOfStrings "Documents
Lists
Endo Enterprises (UK) Ltd
Lawmax Electrical Contractors Ltd
Site Pages
_catalogs
Site Assets
Sharing Links
Anthesis GmbH
Carol Sneddon
Stockholm Environment Institute
Grigoriou Interiours Lts
Fishwick Environmental
Anthesis (UK) Ltd
10:10"

$deleteMe = $allSupplierDrives | ? {$_.name -notin ($doNotDelete)}
$saveMe = $allSupplierDrives | ? {$_.name -in ($doNotDelete)}

$allSupplierDrives.Count
$deleteMe.Count

($deleteMe | sort name).name
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/subs" -Credentials $adminCreds

$deleteMe | % {
    $thisDocLib = $_
    Write-Output "Deleting DocLib [$($thisDocLib.Name)]"
    #invoke-graphDelete -tokenResponse $tokenResponseSharePointBot -graphQuery "/drives/$($deleteMe[1].id)" -Verbose
    $thisList = get-graphList -tokenResponse $tokenResponseSharePointBot -graphDriveId $thisDocLib
    if($thisList){
        #invoke-graphDelete -tokenResponse $tokenResponseSharePointBot -graphQuery "/sites/$supplierSiteId/lists/$($thisList.id)" -Verbose
        Remove-PnPList -Identity $thisList.id -Recycle -Force
    }
}