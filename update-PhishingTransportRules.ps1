$ukRuleName = "External Senders with matching UK Display Names"
$nonUkRuleName = "External Senders with matching non-UK Display Names"
$ruleHtml = "<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 align=left width=`"100%`" style='width:100.0%;mso-cellspacing:0cm;mso-yfti-tbllook:1184; mso-table-lspace:2.25pt;mso-table-rspace:2.25pt;mso-table-anchor-vertical:paragraph;mso-table-anchor-horizontal:column;mso-table-left:left;mso-padding-alt:0cm 0cm 0cm 0cm'>  <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow:yes'><td style='background:#910A19;padding:5.25pt 1.5pt 5.25pt 1.5pt'></td><td width=`"100%`" style='width:100.0%;background:#FDF2F4;padding:5.25pt 3.75pt 5.25pt 11.25pt; word-wrap:break-word' cellpadding=`"7px 5px 7px 15px`" color=`"#212121`"><div><p class=MsoNormal style='mso-element:frame;mso-element-frame-hspace:2.25pt; mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal: column;mso-height-rule:exactly'><span style='font-size:9.0pt;font-family: `"Segoe UI`",sans-serif;mso-fareast-font-family:`"Times New Roman`";color:#212121'>This message was sent from outside the company by someone with a display name matching someone within Anthesis. This could be a coincidence, or it could indicate a phishing attempt. Please do not click links or open attachments unless you recognise the source of this email and know the content is safe. <o:p></o:p></span></p></div></td></tr></table>"

$ukRule = Get-TransportRule | Where-Object {$_.Identity -contains $ukRuleName}
$nonUkRule = Get-TransportRule | Where-Object {$_.Identity -contains $nonUkRuleName}
$displayNames = (Get-Mailbox -ResultSize Unlimited).DisplayName

$mailboxes = Get-Mailbox -ResultSize Unlimited 
$liveUsers = $mailboxes | ? {$_.AccountDisabled -eq $false -and $_.RecipientTypeDetails -eq "UserMailbox"}
$ukUsers = $liveUsers | ? {$_.CustomAttribute1 -match "UK"}
$ukDisplayNames = $ukUsers.DisplayName
$everyoneElse = $liveUsers | ? {$_.CustomAttribute1 -notmatch "UK"}
$nonUkDisplayNames = $everyoneElse.DisplayName

New-TransportRule -Name $ukRuleName -Priority 0 -FromScope "NotInOrganization" -ApplyHtmlDisclaimerLocation "Prepend" -HeaderMatchesMessageHeader From -HeaderMatchesPatterns $ukDisplayNames -ApplyHtmlDisclaimerText $ruleHtml -SentTo kevin.maitland@anthesisgroup.com
New-TransportRule -Name $nonUkRuleName -Priority 0 -FromScope "NotInOrganization" -ApplyHtmlDisclaimerLocation "Prepend" -HeaderMatchesMessageHeader From -HeaderMatchesPatterns $nonUkDisplayNames -ApplyHtmlDisclaimerText $ruleHtml -SentTo kevin.maitland@anthesisgroup.com
get-help New-TransportRule -Detailed