#Connect to site
Import-Module _PS_Library_GeneralFunctionality
Import-Module _PNP_Library_SPO

$smtpaddress = "groupbot@anthesisgroup.com"

$smtpBotDetails = get-graphAppClientCredentials -appName SmtpBot
$tokenResponseSmtp = get-graphTokenResponse -aadAppCreds $smtpBotDetails


$sharePointAdmin = "kimblebot@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content "$env:USERPROFILE\Downloads\KimbleBot.txt") 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass

Connect-PnPOnline -url "https://anthesisllc.sharepoint.com/teams/IT_Team_All_365/" -Credentials $adminCreds





#Process assignees for the week
$allassignees = Get-PnPListItem -List "Anthesis Acts of Kindness"
#Process into pscustom object
$processedassignees = @()
ForEach($unprocessedassignee in $allassignees){

$processedassignees += New-Object psobject -Property @{

    "SharepointID" = $($unprocessedassignee.FieldValues.ID)
    "email" = $($unprocessedassignee.FieldValues.Yourname.Email);
    "community" = $($unprocessedassignee.FieldValues.Community);
    "timezone" = $($unprocessedassignee.FieldValues.Youtimezone);
    "userID" = $($unprocessedassignee.FieldValues.Yourname.LookupId);
    "name" = $($unprocessedassignee.FieldValues.Yourname.LookupValue);
    "country" = $($unprocessedassignee.FieldValues.Yourcountry);
}
}



$thisweekslist = $processedassignees


#We create a multi-dimensional array to hold pairs of contacts by iterating through using two counters, one that starts from 0 and one that starts from half way through the array
$pairedArray = @($false)*[math]::Ceiling($processedassignees.length / 2)
#Set the second counter in the middle of the array to start, this needs to be reset every loop
$j = [math]::floor($processedassignees.length / 2)

for($r = 0; $r -lt 1; $r++){
[System.Collections.ArrayList]$matchIDs = @()
#[System.Collections.ArrayList]$matchIDsold = @()
$j = [math]::floor($processedassignees.length / 2)
#We want to run this once and then check the pairings again historical records 
write-host "Run $($r + 1)..." -ForegroundColor Yellow
    for ($i = 0; $i -lt [math]::floor($processedassignees.length / 2); $i++){
        $pmatchID1 = "$($processedassignees[$i].userID)" + "$($processedassignees[$j].userID)"
        $pmatchID2 = "$($processedassignees[$j].userID)" + "$($processedassignees[$i].userID)"
        #$pmatchID1old = "$($processedassignees[$i].userID)" + "$($processedassignees[$j].userID)"
        #$pmatchID2old = "$($processedassignees[$j].userID)" + "$($processedassignees[$i].userID)"
        $pairedArray[$i] = @($processedassignees[$i],$processedassignees[$j])
        $matchIDs += $pmatchID1
        $matchIDs += $pmatchID2
        #$matchIDsold += $pmatchID1old
        #$matchIDsold += $pmatchID2old
        $j++
    }
#Check output
write-host "Here is our output for run $($r + 1):"
$matchIDs
}

#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#


#Checking
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
ForEach($pair in $pairedArray){

write-host "$($pair[0].email) and $($pair[1].email)" -ForegroundColor Green
#write-host "$($pair[0].email)"
#write-host "$($pair[1].email)"

}





#Send emails
Write-host "Let's finish up and send out the emails!" -ForegroundColor Green
$summaryemailmatches = @()
$ITemail = "IT_Team_GBR@anthesisgroup.com"


ForEach($pair in $pairedArray){

#Construct the first email
$friendlyname0 = $($pair[0].name.Trim().Split(" ")[0].Trim())
$subject = "Anthesis Act of Kindness | You've been matched!"

$body0 = "<HTML><FONT FACE=`"Calibri`">Hi $($friendlyname0),`r`n`r`n<BR><BR>"
$body0 += "Thank you for signing up to the <font color='FF671F'><b>Anthesis Act of Kindness</b></font color> initiative for December 2022.`r`n`r`ntraducció castellana"
$body0 += "We’re excited to share that you have been matched with: <font color='FF671F'><b>$($pair[1].name)</b></font color> in <font color='FF671F'><b>$($pair[1].country)</b></font color>.`r`n<BR>"
$body0 += "`r`n`r`n<BR><BR>"
$body0 += "<font color='FF671F'><b>What do I do next?</b></font color>`r`n`r`n<BR>"
$body0 += "<UL><LI>Decide on your <b>‘Act of Kindness’</b> for the colleague you've been matched with</LI>
               <LI>Complete your <b>‘Act of Kindness’</b> between 6 - 23 December 2022</LI></UL> `r`n`r`n<BR>"
$body0 += "<font color='FF671F'><b>What is an ‘Act of Kindness’?</b></font color>`r`n`r`n<BR>"
$body0 += "There are so many small ways to show kindness as part of our everyday lives; good deeds needn’t take much time or cost money. Get some inspiration from these ideas:`r`n`r`n<BR><BR>"
$body0 += "<UL><LI>Share some positive feedback</LI>
				<LI>Send a fun GIF or photo</LI>
				<LI>Send an e-card</LI>
				<LI>Write a poem</LI>
				<LI>Plant a tree</LI>
				<LI>Make a charity donation on their behalf</LI>
                <LI>Create a personalised playlist to make them smile</LI></UL>`r`n<BR>"
$body0 += "<font color='FF671F'><b>Anything else?</b></font color>`r`n`r`n<BR><BR>"
$body0 += "We want to hear about your Acts of Kindness stories – please share any photos or snippets with <b>Elle Wright</b> from an act you have completed or that you have received.`r`n`r`n<BR><BR>"
$body0 += "We hope you enjoy taking part in this campaign and wish you all the best for the end of 2022.`r`n`r`n<BR><BR>"
$body0 += "With love,`r`n<BR>"
$body0 += "Act of Kindness Match Maker <3 `r`n`r`n<BR><BR>"
$body0 += "(Ps I’m managed by the IT Team, if I have broken or if you have any questions, please get in touch via $($ITemail))<BR><BR><BR>"
$body0 += "<b>Traducción Española</b><BR><BR>"
$body0 += "<HTML><FONT FACE=`"Calibri`">Hola $($friendlyname0),`r`n`r`n<BR><BR>"
$body0 += "Gracias por apuntarte a la iniciativa <font color='FF671F'><b>Actos de Bondad de Anthesis</b></font color> para diciembre de 2022.`r`n`r`n<BR><BR>"
$body0 += "El objetivo de la iniciativa es animar a cada participante a llevar a cabo un pequeño <b>'Acto de Bondad'</b> para su colega, con el fin de difundir un poco de alegría festiva.<BR><BR>"
$body0 += "Nos complace compartir que has sido emparejado con: <font color='FF671F'><b>$($pair[1].name)</b></font color> de <font color='FF671F'><b>$($pair[1].country)</b></font color>.`r`n<BR>"
$body0 += "`r`n`r`n<BR><BR>"
$body0 += "<font color='FF671F'><b>¿Qué hago ahora?</b></font color>`r`n`r`n<BR>"
$body0 += "<UL><LI>Decide tu <b>‘Acto de Bondad’</b> y complétalo con el colega con el que has sido emparejado entre el 6 y el 23 de diciembre</LI>
               <LI>Queremos que nos cuentes tus <b>‘Actos de Bondad’</b> - no dudes en compartir cualquier foto o fragmento con <b>Elle Wright</b></LI></UL> `r`n`r`n<BR>"
$body0 += "<font color='FF671F'><b>¿Qué es un 'acto de bondad'?</b></font color>`r`n`r`n<BR>"
$body0 += "Hay muchas pequeñas formas de mostrar bondad en nuestra vida cotidiana. Las buenas acciones no requieren mucho tiempo ni necesariamente cuestan dinero. Aquí hay algunas ideas:`r`n`r`n<BR><BR>"
$body0 += "<UL><LI>Comparte algún comentario positivo</LI>
			    <LI>Envía un GIF o una foto divertida</LI>
			    <LI>Envía una tarjeta electrónica</LI>
			    <LI>Escribe un poema</LI>
				<LI>Haz una donación benéfica en nombre del destinatario</LI>
				<LI>Crea una playlist personalizada</LI></UL>`r`n<BR>"                
$body0 += "Esperamos que disfrutes de esta campaña y te deseamos lo mejor para el final de 2022.`r`n`r`n<BR><BR>"
$body0 += "Saludos,`r`n<BR>"
$body0 += "Act of Kindness Match Maker <3 `r`n`r`n<BR><BR><BR>"
$body0 += "<b>Traducció Castellana</b><BR><BR>"
$body0 += "<HTML><FONT FACE=`"Calibri`">Hola $($friendlyname0),`r`n`r`n<BR><BR>"
$body0 += "Gràcies per apuntar-te a la iniciativa <font color='FF671F'><b>Actes de Bondat d'Anthesis</b></font color> per al desembre del 2022.`r`n`r`n<BR>"
$body0 += "L'objectiu de la iniciativa és animar cada participant a dur a terme un petit <b>'Acte de Bondat'</b> per al seu col·lega, per tal de difondre una mica d'alegria festiva.<BR><BR>"
$body0 += "Ens complau compartir que has estat aparellat amb: <font color='FF671F'><b>$($pair[1].name)</b></font color> de <font color='FF671F'><b>$($pair[1].country)</b></font color>.`r`n<BR>"
$body0 += "`r`n`r`n<BR><BR>"
$body0 += "<font color='FF671F'><b>Què faig ara?</b></font color>`r`n`r`n<BR>"
$body0 += "<UL><LI>Decideix el teu <b>‘Acte de Bondat’</b> i completa'l amb el col·lega amb què has estat aparellat entre el 6 i el 23 de desembre</LI>
               <LI>Volem que ens expliquis els teus <b>‘Actes de Bondat’</b> - no dubtis a compartir qualsevol foto o fragment amb l'<b>Elle Wright</b></LI></UL> `r`n`r`n<BR>"
$body0 += "<font color='FF671F'><b>Què és un 'acte de bondat'?</b></font color>`r`n`r`n<BR>"
$body0 += "Hi ha moltes petites formes de mostrar bondat a la nostra vida quotidiana. Les bones accions no requereixen molt de temps ni necessàriament costen diners. Aquí teniu algunes idees:`r`n`r`n<BR><BR>"
$body0 += "<UL><LI>Comparteix algun comentari positiu</LI>
			    <LI>Envia un GIF o una foto divertida</LI>
			    <LI>Envia una targeta electrònica</LI>
			    <LI>Escriu un poema</LI>
				<LI>Fes una donació benèfica en nom del destinatari</LI>
				<LI>Crea una playlist personalitzada</LI></UL>`r`n<BR>"                
$body0 += "Esperem que gaudiu d'aquesta campanya i us desitgem el millor per al final del 2022.`r`n`r`n<BR><BR>"
$body0 += "Salutacions,`r`n<BR>"
$body0 += "Act of Kindness Match Maker <3 `r`n`r`n<BR><BR>"





#Construct the second email
$friendlyname1 = $($pair[1].name.Trim().Split(" ")[0].Trim())
$subject = "Anthesis Act of Kindness | You've been matched!"

$body1 = "<HTML><FONT FACE=`"Calibri`">Hi $($friendlyname1),`r`n`r`n<BR><BR>"
$body1 += "Thank you for signing up to the <font color='FF671F'><b>Anthesis Act of Kindness</b></font color> initiative for December 2022.`r`n`r`n<BR><BR>"
$body1 += "We’re excited to share that you have been matched with: <font color='FF671F'><b>$($pair[0].name)</b></font color> in <font color='FF671F'><b>$($pair[0].country)</b></font color>.`r`n<BR>"
$body1 += "`r`n`r`n<BR><BR>"
$body1 += "<font color='FF671F'><b>What do I do next?</b></font color>`r`n`r`n<BR>"
$body1 += "<UL><LI>Decide on your <b>‘Act of Kindness’</b> for the colleague you've been matched with</LI>
               <LI>Complete your <b>‘Act of Kindness’</b> between 6 - 23 December 2022</LI></UL> `r`n`r`n<BR>"
$body1 += "<font color='FF671F'><b>What is an ‘Act of Kindness’?</b></font color>`r`n`r`n<BR>"
$body1 += "There are so many small ways to show kindness as part of our everyday lives; good deeds needn’t take much time or cost money. Get some inspiration from these ideas:`r`n`r`n<BR><BR>"
$body1 += "<UL><LI>Share some positive feedback</LI>
				<LI>Send a fun GIF or photo</LI>
				<LI>Send an e-card</LI>
				<LI>Write a poem</LI>
				<LI>Plant a tree</LI>
				<LI>Make a charity donation on their behalf</LI>
                <LI>Create a personalised playlist to make them smile</LI></UL>`r`n<BR>"
$body1 += "<font color='FF671F'><b>Anything else?</b></font color>`r`n`r`n<BR><BR>"
$body1 += "We want to hear about your Acts of Kindness stories – please share any photos or snippets with <b>Elle Wright</b> from an act you have completed or that you have received.`r`n`r`n<BR><BR>"
$body1 += "We hope you enjoy taking part in this campaign and wish you all the best for the end of 2022.`r`n`r`n<BR><BR>"
$body1 += "With love,`r`n<BR>"
$body1 += "Act of Kindness Match Maker <3 `r`n`r`n<BR><BR>"
$body1 += "(Ps I’m managed by the IT Team, if I have broken or if you have any questions, please get in touch via $($ITemail))<BR><BR><BR><BR>"
$body1 += "<b>Traducción Española</b><BR><BR>"
$body1 += "<HTML><FONT FACE=`"Calibri`">Hola $($friendlyname1),`r`n`r`n<BR><BR>"
$body1 += "Gracias por apuntarte a la iniciativa <font color='FF671F'><b>Actos de Bondad de Anthesis</b></font color> para diciembre de 2022.`r`n`r`n<BR>"
$body1 += "El objetivo de la iniciativa es animar a cada participante a llevar a cabo un pequeño <b>'Acto de Bondad'</b> para su colega, con el fin de difundir un poco de alegría festiva.<BR><BR>"
$body1 += "Nos complace compartir que has sido emparejado con: <font color='FF671F'><b>$($pair[0].name)</b></font color> de <font color='FF671F'><b>$($pair[0].country)</b></font color>.`r`n<BR>"
$body1 += "`r`n`r`n<BR><BR>"
$body1 += "<font color='FF671F'><b>¿Qué hago ahora?</b></font color>`r`n`r`n<BR>"
$body1 += "<UL><LI>Decide tu <b>‘Acto de Bondad’</b> y complétalo con el colega con el que has sido emparejado entre el 6 y el 23 de diciembre</LI>
               <LI>Queremos que nos cuentes <b>‘Actos de Bondad’</b> - no dudes en compartir cualquier foto o fragmento con <b>Elle Wright</b></LI></UL> `r`n`r`n<BR>"
$body1 += "<font color='FF671F'><b>¿Qué es un 'acto de bondad'?</b></font color>`r`n`r`n<BR>"
$body1 += "Hay muchas pequeñas formas de mostrar bondad en nuestra vida cotidiana. Las buenas acciones no requieren mucho tiempo ni necesariamente cuestan dinero. Aquí hay algunas ideas:`r`n`r`n<BR><BR>"
$body1 += "<UL><LI>Comparte algún comentario positivo</LI>
			    <LI>Envía un GIF o una foto divertida</LI>
			    <LI>Envía una tarjeta electrónica</LI>
			    <LI>Escribe un poema</LI>
				<LI>Haz una donación benéfica en nombre del destinatario</LI>
				<LI>Crea una playlist personalizada</LI></UL>`r`n<BR>"                
$body1 += "Esperamos que disfrutes de esta campaña y te deseamos lo mejor para el final de 2022.`r`n`r`n<BR><BR>"
$body1 += "Salutaciones,`r`n<BR>"
$body1 += "Act of Kindness Match Maker <3 `r`n`r`n<BR><BR><BR><BR>"
$body1 += "<b>Traducció Castellana</b><BR><BR>"
$body1 += "<HTML><FONT FACE=`"Calibri`">Hola $($friendlyname1),`r`n`r`n<BR><BR>"
$body1 += "Gràcies per apuntar-te a la iniciativa <font color='FF671F'><b>Actes de Bondat d'Anthesis</b></font color> per al desembre del 2022.`r`n`r`n<BR>"
$body1 += "L'objectiu de la iniciativa és animar cada participant a dur a terme un petit <b>'Acte de Bondat'</b> per al seu col·lega, per tal de difondre una mica d'alegria festiva.<BR><BR>"
$body1 += "Ens complau compartir que has estat aparellat amb: <font color='FF671F'><b>$($pair[0].name)</b></font color> de <font color='FF671F'><b>$($pair[0].country)</b></font color>.`r`n<BR>"
$body1 += "`r`n`r`n<BR><BR>"
$body1 += "<font color='FF671F'><b>Què faig ara?</b></font color>`r`n`r`n<BR>"
$body1 += "<UL><LI>Decideix el teu <b>‘Acte de Bondat’</b> i completa'l amb el col·lega amb què has estat aparellat entre el 6 i el 23 de desembre</LI>
               <LI>Volem que ens expliquis els teus <b>‘Actes de Bondat’</b> - no dudes en compartir cualquier foto o fragmento con <b>Elle Wright</b></LI></UL> `r`n`r`n<BR>"
$body1 += "<font color='FF671F'><b>Què és un 'acte de bondat'?</b></font color>`r`n`r`n<BR>"
$body1 += "Hi ha moltes petites formes de mostrar bondat a la nostra vida quotidiana. Les bones accions no requereixen molt de temps ni necessàriament costen diners. Aquí teniu algunes idees:`r`n`r`n<BR><BR>"
$body1 += "<UL><LI>Comparteix algun comentari positiu</LI>
			    <LI>Envia un GIF o una foto divertida</LI>
			    <LI>Envia una targeta electrónica</LI>
			    <LI>Escriu un poema</LI>
				<LI>Fes una donació benèfica en nom del destinatari</LI>
				<LI>Crea una playlist personalitzada</LI></UL>`r`n<BR>"                
$body1 += "Esperem que gaudiu d'aquesta campanya i us desitgem el millor per al final del 2022.`r`n`r`n<BR><BR>"
$body1 += "Salutacions,`r`n<BR>"
$body1 += "Act of Kindness Match Maker <3 `r`n`r`n<BR><BR>"






send-graphMailMessage -tokenResponse $tokenResponseSmtp -fromUpn AnthesisActsofKindness@anthesisgroup.com -toAddresses "$($pair[0].email)" -subject $subject -bodyHtml $body0 -priority high -Verbose
send-graphMailMessage -tokenResponse $tokenResponseSmtp -fromUpn AnthesisActsofKindness@anthesisgroup.com -toAddresses "$($pair[1].email)" -subject $subject -bodyHtml $body1 -priority high -Verbose

}




  