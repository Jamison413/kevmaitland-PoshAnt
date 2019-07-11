$logFileLocation = "C:\ScriptLogs\" 
$logFileName = "enable-IntuneByodForUsers"
$fullLogPathAndName = $logFileLocation+$logFileName+"_$deviceType`_FullLog_$(Get-Date -Format "yyMMdd").log"
$errorLogPathAndName = $logFileLocation+$logFileName+"_$deviceType`_ErrorLog_$(Get-Date -Format "yyMMdd").log"

Import-Module _PS_Library_GeneralFunctionality
Import-Module _PS_Library_Intune

$intuneAdmin = "kevin.maitland@anthesisgroup.com"
#convertTo-localisedSecureString "IntuneAdminPasswordHere"
$intuneAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\Kev.txt) 
#$adminCreds = set-MsolCredentials -username $intuneAdmin
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $intuneAdmin, $intuneAdminPass
Connect-AzureAD -Credential $adminCreds
connect-ToExo -credential $adminCreds

$usersToEnable = convertTo-arrayOfEmailAddresses "Jake.Cowan@anthesisgroup.com, Charlotte.Moss@anthesisgroup.com"
#$usersToEnable = Get-DistributionGroupMember "All Bristol (GBR)" | % {$_.WindowsLiveID}
$mdmByodDistributionGroup = get-mdmByodDistributionGroup -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName

$usersToEnable | %{
    $thisUpn = $_
    $aadUser = Get-AzureADUser -ObjectId $thisUpn
    #$aadUser = Get-AzureADUser -SearchString $thisUpn.Replace("@anthesisgroup.com","")
    $licenses = Get-AzureADUserLicenseDetail -ObjectId $aadUser.ObjectId
    if($licenses.SkuPartNumber -notcontains "EMS"){
        license-msolUser -pUPN $thisUpn -licenseType EMS
        Send-MailMessage -To itteam@anthesisgroup.com -Subject "User [$thisUpn] requires EM+S E3 License" -From enable-IntuneByodForUsers@anthessigroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com"
        }
    else{
        add-userToMdmByodDistributionGroup -upn $thisUpn -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -mdmByodDistributionGroup $mdmByodDistributionGroup -Verbose
        disable-legacyMailboxProtocols -upn $thisUpn -fullLogFile $fullLogPathAndName -errorLogFile $errorLogPathAndName -Verbose
        }

    }

