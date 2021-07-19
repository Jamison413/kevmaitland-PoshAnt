$logFileLocation = "C:\ScriptLogs\"
$transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_Transcript_$(Get-Date -Format "yyMMdd").log"
if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)){
    $fullLogPathAndName = $logFileLocation+"update-phishingTransportRules.ps1_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = $logFileLocation+"update-phishingTransportRules.ps1_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
else{
    $fullLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
Start-Transcript $transcriptLogName -Append

$groupAdmin = "groupbot@anthesisgroup.com"
$groupAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Downloads\GroupBot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $groupAdmin, $groupAdminPass
connect-ToExo -credential $adminCreds

$smtpBotDetails = get-graphAppClientCredentials -appName SmtpBot
$tokenResponseSmtp = get-graphTokenResponse -aadAppCreds $smtpBotDetails

$ruleHtml = "<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 align=left width=`"100%`" style='width:100.0%;mso-cellspacing:0cm;mso-yfti-tbllook:1184; mso-table-lspace:2.25pt;mso-table-rspace:2.25pt;mso-table-anchor-vertical:paragraph;mso-table-anchor-horizontal:column;mso-table-left:left;mso-padding-alt:0cm 0cm 0cm 0cm'>  <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow:yes'><td style='background:#910A19;padding:5.25pt 1.5pt 5.25pt 1.5pt'></td><td width=`"100%`" style='width:100.0%;background:#FDF2F4;padding:5.25pt 3.75pt 5.25pt 11.25pt; word-wrap:break-word' cellpadding=`"7px 5px 7px 15px`" color=`"#212121`"><div><p class=MsoNormal style='mso-element:frame;mso-element-frame-hspace:2.25pt; mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal: column;mso-height-rule:exactly'><span style='font-size:9.0pt;font-family: `"Segoe UI`",sans-serif;mso-fareast-font-family:`"Times New Roman`";color:#212121'>This message was sent from outside the company by someone with a display name matching someone within Anthesis. This could be a coincidence, or it could indicate a <A HREF='https://anthesisllc.sharepoint.com/sites/Resources-IT/_layouts/15/DocIdRedir.aspx?ID=HXX7CE52TSD2-1759992947-268'>CEO Fraud</A> / phishing attempt. Please do not click links or open attachments unless you recognise the source of this email and know the content is safe.<BR><BR>Anthesis IT will never add green banners to persuade you that an e-mail is legitimate. Learn more about <A HREF=`"https://anthesisllc.sharepoint.com/sites/Resources-IT/_layouts/15/DocIdRedir.aspx?ID=HXX7CE52TSD2-1759992947-5`">spotting phishing e-mails</A> and how to <A HREF=`"https://anthesisllc.sharepoint.com/sites/Resources-IT/_layouts/15/DocIdRedir.aspx?ID=HXX7CE52TSD2-1759992947-8`">report them to Microsoft</A><o:p></o:p></span></p></div></td></tr></table>"

$mailboxes = Get-Mailbox -ResultSize Unlimited 
$liveUsers = $mailboxes | ? {$_.AccountDisabled -eq $false -and $_.RecipientTypeDetails -eq "UserMailbox"}
$alphabet = [char[]]([int][char]'Z'..[int][char]'A')
$alphabet | % {
    Write-Verbose "Anti-CEO Fraud rule [$_]"
    Write-Verbose "`tGenerating list of DisplayNames beginning with [$_]"
    $thisLetter = $_
    New-Variable -Name "displayNameList$_" -Value $($($liveUsers | ? {$_.DisplayName -match "^[$thisLetter]"}).DisplayName | Sort-Object)
    Write-Verbose "`t`t[$((Get-Variable -Name "displayNameList$_").Value.Count)] DisplayNames beginning with [$_]"
    #Write-Verbose "`tCreating new Transport Rule [$("External Senders with matching Display Names beginning with $_")]"
    #New-Variable -Name "transportRule$_" -Value $(New-TransportRule -Name "External Senders with matching Display Names beginning with $_" -Priority 0 -FromScope "NotInOrganization" -ApplyHtmlDisclaimerLocation "Prepend" -HeaderMatchesMessageHeader From -HeaderMatchesPatterns $((Get-Variable "displayNameList$_").Value) -ApplyHtmlDisclaimerText $ruleHtml -SentTo "kevin.maitland@anthesisgroup.com" -ExceptIfFrom @("noreply@email.teams.microsoft.com","no-reply@sharepointonline.com"))
    Write-Verbose "`tUpdating Transport Rule [$("External Senders with matching Display Names beginning with $_")]"
    try{Set-TransportRule -Identity "External Senders with matching Display Names beginning with $_" -Name "External Senders with matching Display Names beginning with $_" -Priority 0 -FromScope "NotInOrganization" -ApplyHtmlDisclaimerLocation "Prepend" -HeaderMatchesMessageHeader From -HeaderMatchesPatterns $((Get-Variable "displayNameList$_").Value) -ApplyHtmlDisclaimerText $ruleHtml -SentTo $null -ExceptIfFrom @("noreply@email.teams.microsoft.com","no-reply@sharepointonline.com") -ErrorAction Stop}
    catch{send-graphMailMessage -tokenResponse $tokenResponseSmtp -fromUpn $groupAdmin -toAddresses ITTeamAll-Managers@anthesisgroup.com -subject "Error updating [$("External Senders with matching Display Names beginning with $thisLetter")] in update-PhisingTransportRules.ps1 on [$env:COMPUTERNAME]" -bodyText $(get-errorSummary $_) -priority high}
    }

<#
$ukRuleName = "External Senders with matching UK Display Names"
$spainRuleName = "External Senders with matching Spain Display Names"
$restOfWorldRuleName = "External Senders with matching RoW Display Names"
$ukRule = Get-TransportRule | Where-Object {$_.Identity -contains $ukRuleName}
$spainRule = Get-TransportRule | Where-Object {$_.Identity -contains $spainRuleName}
$roWRule = Get-TransportRule | Where-Object {$_.Identity -contains $restOfWorldRuleName}
#$displayNames = (Get-Mailbox -ResultSize Unlimited).DisplayName


$ukUsers = $liveUsers | ? {$_.UsageLocation -eq "United Kingdom"} 
$ukDisplayNames = $ukUsers.DisplayName | Sort-Object
$spainUsers = $liveUsers | ? {$_.UsageLocation -eq "Spain"}
$spainDisplayNames = $spainUsers.DisplayName | Sort-Object
$everyoneElse = $liveUsers | ? {$_.UsageLocation -ne "United Kingdom" -and $_.UsageLocation -ne "Spain"}
$roWDisplayNames = $everyoneElse.DisplayName | Sort-Object

try{$ukRule | Set-TransportRule -Name $ukRuleName -Priority 0 -FromScope "NotInOrganization" -ApplyHtmlDisclaimerLocation "Prepend" -HeaderMatchesMessageHeader From -HeaderMatchesPatterns $ukDisplayNames -ApplyHtmlDisclaimerText $ruleHtml -SentTo $null -ExceptIfFrom @("noreply@email.teams.microsoft.com","no-reply@sharepointonline.com") -ErrorAction Stop}
catch{send-graphMailMessage -tokenResponse $tokenResponseSmtp -fromUpn $groupAdmin -toAddresses ITTeamAll-Managers@anthesisgroup.com -subject "Error updating `$ukRule in update-PhisingTransportRules.ps1 on [$env:COMPUTERNAME]" -bodyText $(get-errorSummary $_) -priority high}
try{$spainRule | Set-TransportRule -Name $spainRuleName -Priority 0 -FromScope "NotInOrganization" -ApplyHtmlDisclaimerLocation "Prepend" -HeaderMatchesMessageHeader From -HeaderMatchesPatterns $spainDisplayNames -ApplyHtmlDisclaimerText $ruleHtml -SentTo $null  -ExceptIfFrom @("noreply@email.teams.microsoft.com","no-reply@sharepointonline.com") -ErrorAction Stop}
catch{send-graphMailMessage -tokenResponse $tokenResponseSmtp -fromUpn $groupAdmin -toAddresses ITTeamAll-Managers@anthesisgroup.com -subject "Error updating `$spainRule in update-PhisingTransportRules.ps1 on [$env:COMPUTERNAME]" -bodyText $(get-errorSummary $_) -priority high}
try{$roWRule | Set-TransportRule -Name $restOfWorldRuleName -Priority 0 -FromScope "NotInOrganization" -ApplyHtmlDisclaimerLocation "Prepend" -HeaderMatchesMessageHeader From -HeaderMatchesPatterns $roWDisplayNames -ApplyHtmlDisclaimerText $ruleHtml -SentTo $null  -ExceptIfFrom @("noreply@email.teams.microsoft.com","no-reply@sharepointonline.com")} #kevin.maitland@anthesisgroup.com
catch{send-graphMailMessage -tokenResponse $tokenResponseSmtp -fromUpn $groupAdmin -toAddresses ITTeamAll-Managers@anthesisgroup.com -subject "Error updating `$roWRule in update-PhisingTransportRules.ps1 on [$env:COMPUTERNAME]" -bodyText $(get-errorSummary $_) -priority high}
#get-help New-TransportRule -Detailed
#>

Stop-Transcript

