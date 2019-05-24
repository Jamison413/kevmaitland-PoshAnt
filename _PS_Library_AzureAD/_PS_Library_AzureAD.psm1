function enumerate-nestedDistributionGroupsToAadUsers(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [PSObject]$distributionGroupObject
        )
    $immediateMembers = Get-DistributionGroupMember -Identity $distributionGroupObject.ExternalDirectoryObjectId
    $immediateMembers | % {
        $thisMember = $_
        switch($thisMember.RecipientTypeDetails){
            ("UserMailbox") {
                $aadUser = Get-AzureADUser -ObjectId $thisMember.WindowsLiveID
                Write-Verbose "AADUser [$($aadUser.DisplayName)] is a member of [$($distributionGroupObject.DisplayName)]"
                [array]$aadUserObjects += $aadUser
                }
            ("MailUniversalSecurityGroup"){
                $subDistributionGroup = Get-DistributionGroup -Identity $thisMember.ExternalDirectoryObjectId
                [array]$subAadUserObjects = enumerate-nestedDistributionGroupsToAadUsers -distributionGroupObject $subDistributionGroup
                Write-Host "`$aadUserObjects.Count = $($aadUserObjects.Count) `t`$subAadUserObjects.Count = $($subAadUserObjects.Count)"
                $aadUserObjects = $subAadUserObjects
                }
            default {}
            }
        }
    $aadUserObjects
    }
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
