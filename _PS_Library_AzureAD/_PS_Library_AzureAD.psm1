function get-aadUsersWithLicense(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [array]$arrayOfAadUserObjects
        ,[parameter(Mandatory = $true)]
        [ValidateSet("E1","E3","EMS","K1","Project","Visio")]
        [string]$licenseType
        ,[parameter(Mandatory = $false)]
        [bool]$getUsersWithoutLicenseInstead
        )
    switch($licenseType){
        "E1" {$skuPartNumber = "STANDARDPACK"}
        "E3" {$skuPartNumber = "ENTERPRISEPACK"}
        "EMS" {$skuPartNumber = "EMS"}
        "K1" {$skuPartNumber = "EXCHANGEDESKLESS"}
        "Project" {$skuPartNumber = "PROJECTPROFESSIONAL"}
        "Visio" {$skuPartNumber = "VISIOCLIENT"}
        }

    $arrayOfAadUserObjects | % {
        $thisUser = $_
        $thisLicenseSet = Get-AzureADUserLicenseDetail -ObjectId $thisUser.ObjectId
        if($thisLicenseSet.SkuPartNumber -contains $skuPartNumber){
            [array]$usersWithLicense += $thisUser
            }
        else{[array]$usersWithoutLicense += $thisUser}
        }
    
    if($getUsersWithoutLicenseInstead){
        $usersWithoutLicense
        }
    else{$usersWithLicense}
    }
