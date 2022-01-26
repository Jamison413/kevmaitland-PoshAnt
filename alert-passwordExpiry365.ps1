$logFileLocation = "C:\ScriptLogs\"
$transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_Transcript_$(Get-Date -Format "yyMMdd").log"
if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)){
    $fullLogPathAndName = $logFileLocation+"alert-passwordExpiry365_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = $logFileLocation+"alert-passwordExpiry365_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
else{
    $fullLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
$debugLog = "$env:USERPROFILE\Desktop\debugdump.log"
Start-Transcript $transcriptLogName -Append


Import-Module _PS_Library_MSOL
Import-Module _PS_Library_GeneralFunctionality

$groupAdmin = "groupbot@anthesisgroup.com"
#convertTo-localisedSecureString ""
$groupAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Downloads\GroupBot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $groupAdmin, $groupAdminPass
connect-ToMsol -credential $adminCreds

$expiryUsers = Get-MsolUser -All | ?{($(get-date )-$(Get-Date $_.LastPasswordChangeTimestamp)).Days -ge 105 -and $_.islicensed}

$expiryUsers | % {
    $thisUser = $_
     write-host $thisUser.DisplayName `t $($($(Get-Date $thisUser.LastPasswordChangeTimestamp).AddDays(120) - $(Get-Date)).Days)
    if(@(1,7,15) -contains ($(get-date )-$(Get-Date $thisUser.LastPasswordChangeTimestamp)).Days){
    #if($($($(Get-Date $thisUser.LastPasswordChangeTimestamp).AddDays(120) - $(Get-Date)).Days) -ge 0){
        Write-Host -for Yellow "ExpiringIn" $thisUser.UserPrincipalName
        $subject = "365 password expiry in $($($(Get-Date $thisUser.LastPasswordChangeTimestamp).AddDays(120) - $(Get-Date)).Days) days"
        $body = "<HTML><FONT FACE=`"Calibri`">Hello $($thisUser.FirstName),`r`n`r`n<BR><BR>"
        $body += "Just to let you know, you last changed your Anthesis 365 password on $(Get-Date $thisUser.LastPasswordChangeTimestamp -Format D) at $(Get-Date $thisUser.LastPasswordChangeTimestamp -Format t) UTC. Anthesis 365 passwords automatically expire after 120 days, which is on $(Get-Date $(Get-Date $thisUser.LastPasswordChangeTimestamp).AddDays(120) -f D) (in $($($(Get-Date $thisUser.LastPasswordChangeTimestamp).AddDays(120) - $(Get-Date)).Days) days).`r`n`r`n<BR><BR>"
        $body += "It is recommended that you change your Anthesis password before it expires to minimise any disruption in your access to Anthesis' 365 services. If you have forgotten your password, you will need to contact your regional IT support person and ask for their assistance in resetting it. If you <I>do</I> know your password, you can change it yourself by:`r`n`r`n<BR><BR>"
        $body += "<OL>"
        $body += "<LI>Logging into https://office.com</LI>"
        $body += "<LI>Log in using your @anthesisgroup.com email address and current Office 365 password</LI>"
        $body += "<LI>Click on the <B>Cog</B> at the top of the screen</LI>"
        $body += "<LI>Click on <B>Change password</B></LI>"
        $body += "</OL>"
        $body += "Just so that you’re aware where this password is used, when you change it you will need to enter the new one:"
        $body += "<UL>"
        $body += "<LI>If/when you log into any web-based Office 365 service:</LI>"
        $body += "<UL>"
        $body += "<LI><A HREF=""https://anthesisllc.sharepoint.com/global"">SharePoint</A></LI>"
        $body += "<LI><A HREF=""https://anthesisllc-my.sharepoint.com/personal/"">OneDrive</A></LI>"
        $body += "<LI><A HREF=""https://outlook.office.com/"">Webmail</A></LI></UL>"
        $body += "<LI>On any mobile device (phone/tablet) that receives your @anthesisgroup.com e-mail</LI>"
        $body += "<LI>When logging into Skype for Business</LI>"
        $body += "</UL>"
        $body += "Love,`r`n`r`n<BR><BR>"
        $body += "The Helpful 365 Password Robot`r`n`r`n<BR><BR></FONT></HTML>"
        Send-MailMessage -To $thisUser.UserPrincipalName -From "thehelpful365passwordrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
        }
    elseif(($(get-date )-$(Get-Date $thisUser.LastPasswordChangeTimestamp)).Days -le 0 -and ($(get-date )-$(Get-Date $thisUser.LastPasswordChangeTimestamp)).Days %7 -eq 0){
    #else{
        Write-Host -for Magenta "Expired" $thisUser.UserPrincipalName
        $subject = "365 password expired $($($(Get-Date $thisUser.LastPasswordChangeTimestamp).AddDays(120) - $(Get-Date)).Days*-1) days ago"
        $body = "<HTML><FONT FACE=`"Calibri`">Hello $($thisUser.FirstName),`r`n`r`n<BR><BR>"
        $body += "Just to let you know, you last changed your Anthesis 365 password on $(Get-Date $thisUser.LastPasswordChangeTimestamp -Format D) at $(Get-Date $thisUser.LastPasswordChangeTimestamp -Format t) UTC. Anthesis 365 passwords automatically expire after 120 days, which was on $(Get-Date $(Get-Date $thisUser.LastPasswordChangeTimestamp).AddDays(120) -f D) ($($($(Get-Date $thisUser.LastPasswordChangeTimestamp).AddDays(120) - $(Get-Date)).Days*-1) days ago).`r`n`r`n<BR><BR>"
        $body += "It is recommended that you change your Anthesis password before it expires to minimise any disruption in your access to Anthesis' 365 services. If you have forgotten your password, you will need to contact your regional IT support person and ask for their assistance in resetting it. If you <I>do</I> know your password, you can change it yourself by:`r`n`r`n<BR><BR>"
        $body += "<OL>"
        $body += "<LI>Logging into https://office.com</LI>"
        $body += "<LI>Log in using your @anthesisgroup.com email address and current Office 365 password</LI>"
        $body += "<LI>Click on the <B>Cog</B> at the top of the screen</LI>"
        $body += "<LI>Click on <B>Change password</B></LI>"
        $body += "</OL>"
        $body += "Just so that you’re aware where this password is used, when you change it you will need to enter the new one:"
        $body += "<UL>"
        $body += "<LI>If/when you log into any web-based Office 365 service:</LI>"
        $body += "<UL>"
        $body += "<LI><A HREF=""https://anthesisllc.sharepoint.com/global"">SharePoint</A></LI>"
        $body += "<LI><A HREF=""https://anthesisllc-my.sharepoint.com/personal/"">OneDrive</A></LI>"
        $body += "<LI><A HREF=""https://outlook.office.com/"">Webmail</A></LI></UL>"
        $body += "<LI>On any mobile device (phone/tablet) that receives your @anthesisgroup.com e-mail</LI>"
        $body += "<LI>When logging into Skype for Business</LI>"
        $body += "</UL>"
        $body += "Love,`r`n`r`n<BR><BR>"
        $body += "The Helpful 365 Password Robot`r`n`r`n<BR><BR></FONT></HTML>"
        Send-MailMessage -To $thisUser.UserPrincipalName -From "thehelpful365passwordrobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
        }
    }

Stop-Transcript