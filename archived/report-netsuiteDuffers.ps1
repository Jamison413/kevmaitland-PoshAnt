#region Report Opps & Projs with names containig illegal SharePoint folder characters
$illegallyNamedOpps  = $allNetSuiteOpps  | ? {$($_.UniversalOppName -eq $(sanitise-forSharePointFolderName $_.UniversalOppName)) -eq $false}
$illegallyNamedProjs = $allNetSuiteProjs | ? {$($_.UniversalProjName -eq $(sanitise-forSharePointFolderName $_.UniversalProjName)) -eq $false}

$duplicatedClientNames = $allNetSuiteClients | Group-Object -Property {$_.companyName} | ? {$_.Count -gt 1}

$duplicatedClientNames | % {
    $thisGroup = $_
    $originalRecord = $thisGroup.Group | ? {$_.entityStatus.refName -match "CLIENT"}
    if($originalRecord -eq $null){$originalRecord = $thisGroup.Group | ? {$_.entityStatus.refName -match "PROSPECT"}}
    if($originalRecord -eq $null){$originalRecord = $thisGroup.Group | ? {$_.entityStatus.refName -match "LEAD"}}
    $originalRecord = $originalRecord | Sort-Object dateCreated | Select-Object -First 1
    $duplicateRecords = $thisGroup.Group | ? {$_.id -ne $originalRecord.id}
    $duplicateRecords | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name DuplicateOf -Value $originalRecord.id -Force}
    }
$duplicateClientNames = $duplicatedClientNames.Group | ? {![string]::IsNullOrEmpty($_.DuplicateOf)} 
$oppsBlockedByDuplicateClients = $allNetSuiteOpps | ? {$duplicateClientNames.NetSuiteClientId -contains $_.NetSuiteClientId}
$projsBlockedByDuplicateClients = $allNetSuiteProjs | ? {$duplicateClientNames.NetSuiteClientId -contains $_.NetSuiteClientId}

$test = $illegallyNamedOpps | Group-Object {$_.salesRep.refName}

$illegallyNamedOpps.UniversalOppName

$body =  "<HTML>"
$body += "<P>Hi $user,</P>"
$body += "<P>Sorry, I can't create your Opportunity folders below because they contain `"special`" characters that can't be used in folder names in SharePoint. I've suggested an alterntive name that will let me create your folders for you, but I'm really bad at guessing whether this will cause you problems so you'll have to follow the links and make the changes yourself:</P>"
$body += "<UL>"
$illegallyNamedOpps | Sort-Object {$_.salesRep.refName},tranId | % {
    $body += "<LI><A HREF=`"https://3487287.app.netsuite.com/app/accounting/transactions/opprtnty.nl?id=$($_.id)`">$($_.UniversalOppName)</A> owned by <B>$($_.salesRep.refName)</B>, you could change this to: <B>$(sanitise-forSharePointFolderName $_.UniversalOppName)</B></LI>"
    }
$body += "</UL>"
$body += "<P>These characters can't be used in a SharePoint folder name: <B>`" * : < > ? / \ |</B>  and it also cannot end with a . </P>"
$body += "<P>Love,</P>"
$body += "<P>The Netsuite-SharePoint Sync Robot</P>"
$body += "</HTML>"
Send-MailMessage -To "kevin.maitland@anthesisgroup.com" -Subject "Duff Opportunity NetSuite Names" -Encoding UTF8 -BodyAsHtml $body -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -From "thenetsuitesharePointsyncrobot@anthesisgroup.com"


$body =  "<HTML>"
$body += "<P>Hi $user,</P>"
$body += "<P>Sorry, I can't create your Project folders below because they contain `"special`" characters that can't be used in folder names in SharePoint. I've suggested an alterntive name that will let me create your folders for you, but I'm really bad at guessing whether this will cause you problems so you'll have to follow the links and make the changes yourself:</P>"
$body += "<UL>"
$illegallyNamedProjs | Sort-Object {$_.projectManager.refName},NetSuiteProjCode | % {
    $body += "<LI><A HREF=`"https://3487287.app.netsuite.com/app/accounting/project/project.nl?id=$($_.id)`">$($_.UniversalProjName)</A> owned by <B>$($_.projectManager.refName)</B>, you could change this to: <B>$(sanitise-forSharePointFolderName $_.UniversalProjName)</B></LI>"
    }
$body += "</UL>"
$body += "<P>These characters can't be used in a SharePoint folder name: <B>`" * : < > ? / \ |</B>  and it also cannot end with a . </P>"
$body += "<P>Love,</P>"
$body += "<P>The Netsuite-SharePoint Sync Robot</P>"
$body += "</HTML>"
Send-MailMessage -To "kevin.maitland@anthesisgroup.com" -Subject "Duff Project NetSuite Names" -Encoding UTF8 -BodyAsHtml $body -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -From "thenetsuitesharePointsyncrobot@anthesisgroup.com"


$body =  "<HTML>"
$body += "<P>Hi $user,</P>"
$body += "<P>Sorry, I can't create your Opportunity folders below because they are attached to a duplicate Client/Prospect. You'll need to merge the Client/Prospect yourself in NetSuite (and I'm neither clever, nor reliable enough to do this for you):</P>"
$body += "<UL>"
$oppsBlockedByDuplicateClients | Sort-Object {$_.salesRep.refName},tranId | % {
    $body += "<LI><A HREF=`"https://3487287.app.netsuite.com/app/accounting/transactions/opprtnty.nl?id=$($_.id)`">$($_.UniversalOppName)</A> owned by <B>$($_.salesRep.refName)</B></LI>"
    }
$body += "</UL>"
$body += "<P>Love,</P>"
$body += "<P>The Netsuite-SharePoint Sync Robot</P>"
$body += "</HTML>"
Send-MailMessage -To "kevin.maitland@anthesisgroup.com" -Subject "Opportunities blocked by duplicate clients" -Encoding UTF8 -BodyAsHtml $body -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -From "thenetsuitesharePointsyncrobot@anthesisgroup.com"


$body =  "<HTML>"
$body += "<P>Hi $user,</P>"
$body += "<P>Sorry, I can't create your Project folders below because they are attached to a duplicate Client/Prospect. You'll need to merge the Client/Prospect yourself in NetSuite (and I'm neither clever, nor reliable enough to do this for you):</P>"
$body += "<UL>"
$projsBlockedByDuplicateClients | Sort-Object {$_.projectManager.refName},NetSuiteProjCode | % {
    $body += "<LI><A HREF=`"https://3487287.app.netsuite.com/app/accounting/project/project.nl?id=$($_.id)`">$($_.UniversalProjName)</A> owned by <B>$($_.projectManager.refName)</B></LI>"
    }
$body += "</UL>"
$body += "<P>Love,</P>"
$body += "<P>The Netsuite-SharePoint Sync Robot</P>"
$body += "</HTML>"
Send-MailMessage -To "kevin.maitland@anthesisgroup.com" -Subject "Projects blocked by duplicate clients" -Encoding UTF8 -BodyAsHtml $body -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -From "thenetsuitesharePointsyncrobot@anthesisgroup.com"


$body =  "<HTML>"
$body += "<P>Hi $user,</P>"
$body += "<P>Sorry, you seem to have created a duplicate Client/Prospect in NetSuite. I can't create a folder in SharePoint for this duplicate, so I can't create folders for Opportunities or Projects either. This will cause confusion for other people too, so you'll need to merge the Client/Prospect yourself in NetSuite (and unfortunately I'm neither clever, nor reliable enough to do this for you):</P>"
$body += "<UL>"
$duplicateClientNames | Sort-Object DuplicateOf,entityTitle | % {
    $body += "<LI><A HREF=`"https://3487287.app.netsuite.com/app/common/entity/custjob.nl?id=$($_.id)`">$($_.entityTitle)</A> owned by <B>$($_.salesRep.refName)</B> is a duplicate of <A HREF=`"https://3487287.app.netsuite.com/app/common/entity/custjob.nl?id=$($_.DuplicateOf)`">$($_.UniversalClientName)</A></LI>"
    }
$body += "</UL>"
$body += "<P>Love,</P>"
$body += "<P>The Netsuite-SharePoint Sync Robot</P>"
$body += "</HTML>"
Send-MailMessage -To "kevin.maitland@anthesisgroup.com" -Subject "Duplicate clients warning" -Encoding UTF8 -BodyAsHtml $body -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -From "thenetsuitesharePointsyncrobot@anthesisgroup.com"
