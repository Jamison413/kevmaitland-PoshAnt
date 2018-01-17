Import-Module _PS_Library_MSOL
Import-Module _PS_Library_GeneralFunctionality
$msolCredentials = set-MsolCredentials
connect-ToMsol -credential $msolCredentials

$toDelete = convertTo-arrayOfEmailAddresses -blockOfText "Luis.Schaeffer@anthesisgroup.com
Matteo.Cossutta@anthesisgroup.com
Elanor.Swan@anthesisgroup.com
Tom.Drzewiecki@anthesisgroup.com
Saro.Sivakumaran@anthesisgroup.com
Bethany.Field@anthesisgroup.com
CJ.Westrick@anthesisgroup.com
Lucy.Richardson@anthesisgroup.com
Michelle.Langefeld@anthesisgroup.com
Dries.Dhooghe@anthesisgroup.com
will.schreiber@anthesisgroup.com
Paul.McNeillis@anthesisgroup.com
cara.egbe@anthesisgroup.com
Rob.Ashdown@anthesisgroup.com
Maximus.tam@anthesisgroup.com
Luis.Ibhiabor@anthesisgroup.com
kathy.smith1@anthesisgroup.com
Stefano.Andreola@anthesisgroup.com
coramorena.fritz@anthesisgroup.com
Dave.Cooke@anthesisgroup.com
wdpec@anthesisgroup.com
nicola.jenkin@anthesisgroup.com
Fiona.Barker@anthesisgroup.com
Esther.Areikin@anthesisgroup.com
William.Fletcher@anthesisgroup.com
Pippa.Reid@anthesisgroup.com
Alice.Handscomb@anthesisgroup.com
nicola.peters@anthesisgroup.com
Daniel.Jianoran@anthesisgroup.com
Geoff.Green@anthesisgroup.com
Alan.Ritchie@anthesisgroup.com
Ruth.Norriss@anthesisgroup.com
Ashley.Retallack@anthesisgroup.com
Tim.Duke@anthesisgroup.com
WD@anthesisgroup.com
renan.navarro@anthesisgroup.com
Isabel.Terry@anthesisgroup.com
alastair.pattrick@anthesisgroup.com
Elise.Benjiman@anthesisgroup.com
lucy.welch@anthesisgroup.com
Cat.Hobbs@anthesisgroup.com
Rosie.Sibley@anthesisgroup.com
Pat.Glanville@anthesisgroup.com
erika.bata@anthesisgroup.com
Charles.Perry@anthesisgroup.com
David.Fellows@anthesisgroup.com"
$sustain = convertTo-arrayOfEmailAddresses -blockOfText "Kitbag-EEF1@anthesisgroup.com
Kitbag-ESOS1@anthesisgroup.com
Kitbag-ECOTechnicalMonitoring@anthesisgroup.com
Kitbag-ESOS2@anthesisgroup.com
Kitbag-EEF2@anthesisgroup.com
MeetingRoom.Lovins@anthesisgroup.com
MeetingRoom.Tesla@anthesisgroup.com
MeetingRoomWatt@anthesisgroup.com
BidsTenders@sustain.co.uk
SatNav-Shephard@anthesisgroup.com
SatNav-Hurley@anthesisgroup.com
iPad02@anthesisgroup.com
Camera-ThermalImaging2@anthesisgroup.com
Sustain.Recruitment@anthesisgroup.com
Kitbag-SmartHeat@anthesisgroup.com
Camera-ThermalImaging1@anthesisgroup.com
SatNav-Austen@anthesisgroup.com"

$toDisable2 = convertTo-arrayOfEmailAddresses -blockOfText "Ali.Mahdavi@anthesisgroup.com
OxfordBoardRoom@anthesisgroup.com
Macclesfieldmeetingroom@anthesisgroup.com
Ningwei.Dong@anthesisgroup.com
content@anthesisgroup.com
CDP-Support@anthesisgroup.com
IoSBusinessEnergy@anthesisgroup.com
unsubscribe@lrsconsultancy.com
Kitbag-EEF1@anthesisgroup.com
Mats.Ivarsson@anthesisgroup.com
nwsbq@anthesisgroup.com
Kitbag-ESOS1@anthesisgroup.com
OxfordTrainingRoom@anthesisgroup.com
UK.Info@anthesisgroup.com
Kitbag-ECOTechnicalMonitoring@anthesisgroup.com
BWI@anthesisgroup.com
wsfpbs@anthesisgroup.com
MeetingRoom.Lovins@anthesisgroup.com
contact@anthesisgroup.com
MAILLRS@anthesisgroup.com
mail@anthesisgroup.com
Kitbag-EEF2@anthesisgroup.com
Liesa.Guttmann2@anthesisgroup.com
Ulf.Siefker@anthesisgroup.com
Conference.Genie@lrsconsultancy.com
CDPsupport@anthesisgroup.com
UK.SysAdmin@anthesisgroup.com
MeetingRoomWatt@anthesisgroup.com
globalcalendars@anthesisgroup.com
SMO-AnthesisResources@anthesisgroup.com
test.users@anthesisgroup.com
David.Baker@anthesisgroup.com
Paul.Becker@anthesisgroup.com
quintiles-help@anthesisgroup.com
adminpcrrg.uk@anthesisgroup.com
CalebOR@anthesisgroup.com
Richard.Gibbs@anthesisgroup.com
Dan.Verdonik@anthesisgroup.com
ukanalysts@anthesisgroup.com
Apple@anthesisgroup.com
rita@anthesisgroup.com
walmart@anthesisgroup.com
m.hidary@anthesisgroup.com
ian@anthesisgroup.com
londontcsroom@anthesisgroup.com
Erik.Wallentin@anthesisgroup.com
OxfordMeeting@anthesisgroup.com
Emergen.Programme@anthesisgroup.com
calebconatcts@anthesisgroup.com
SatNav-Shephard@anthesisgroup.com
kimberly@anthesisgroup.com
brent.alarma@anthesisgroup.com
BidsTenders@sustain.co.uk
diageo-help@anthesisgroup.com
SatNav-Austen@anthesisgroup.com
Matthew.Williams@anthesisgroup.com
administrator@lrsconsultancy.com
ukbidsandtenders@anthesisgroup.com
Info123@anthesisgroup.com
anna.rengstedt2@anthesisgroup.com
tenders@anthesisgroup.com
Gillian.Phillips@anthesisgroup.com
jobs_de@anthesisgroup.com
iPad01@anthesisgroup.com
EORM@anthesisgroup.com
sustainability@anthesisgroup.com
MeetingRoom.Tesla@anthesisgroup.com
Tom.McKellarSmythe@anthesisgroup.com
SMO-AnthesisResources1@anthesisgroup.com
Support@anthesisgroup.com
Uk.Admin@anthesisgroup.com
Kitbag-ESOS2@anthesisgroup.com
John.Yates@anthesisgroup.com
Emma.Collen@anthesisgroup.com
NA-CareersAutoreply@anthesisgroup.com
ManchesterHotDesks@anthesisgroup.com
Tom.Parker@anthesisgroup.com
SatNav-Hurley@anthesisgroup.com
communications@anthesisgroup.com
Lukas.Fingerhut@anthesisgroup.com
iPad02@anthesisgroup.com
Camera-ThermalImaging2@anthesisgroup.com
Kimble@anthesisgroup.com
Kitbag-SmartHeat@anthesisgroup.com
Gregor.Pecnik@anthesisgroup.com
ann.durrant@anthesisgroup.com
Camera-ThermalImaging1@anthesisgroup.com
Collin.Marshall@anthesisgroup.com
Sustain.Recruitment@anthesisgroup.com
Copier@lrsconsultancy.com"
    FLEETSTMTGRMSNUG@anthesisgroup.com
    FLEETSTMTGRM1@anthesisgroup.com

$toDelete | %{Set-MsolUser -UserPrincipalName $_ -BlockCredential $true}
$sustain | %{Set-MsolUser -UserPrincipalName $_ -BlockCredential $true}

$sustain.count
$toDisable.count

$toDisable2 | %{
    Write-Host -ForegroundColor Yellow $_
    Set-MsolUser -UserPrincipalName $_ -BlockCredential $true
    }

$toDelete |%{Get-Mailbox -Identity $_ | %{Write-Host -ForegroundColor Yellow "$($_.DisplayName)`t$($_.RecipientTypeDetails)"}}
$dummy = Get-Mailbox -Identity $toDelete[$toDelete.Count-1] 
$dummy | fl

Remove-MsolUser -UserPrincipalName Matteo.Cossutta@anthesisgroup.com
$toDelete |%{Remove-MsolUser -UserPrincipalName $_ -Force}
