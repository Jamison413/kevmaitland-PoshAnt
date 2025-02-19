﻿function Convert-AzureAdObjectIdToSid {
<#
.SYNOPSIS
Convert an Azure AD Object ID to SID
 
.DESCRIPTION
Converts an Azure AD Object ID to a SID.
Author: Oliver Kieselbach (oliverkieselbach.com)
The script is provided "AS IS" with no warranties.
 
.PARAMETER ObjectID
The Object ID to convert
#>

    param([String] $ObjectId)

    $bytes = [Guid]::Parse($ObjectId).ToByteArray()
    $array = New-Object 'UInt32[]' 4

    [Buffer]::BlockCopy($bytes, 0, $array, 0, 16)
    $sid = "S-1-12-1-$array".Replace(' ', '-')

    return $sid
}
function Convert-AzureAdSidToObjectId {
<#
.SYNOPSIS
Convert a Azure AD SID to Object ID
 
.DESCRIPTION
Converts an Azure AD SID to Object ID.
Author: Oliver Kieselbach (oliverkieselbach.com)
The script is provided "AS IS" with no warranties.
 
.PARAMETER ObjectID
The SID to convert
#>

    param([String] $Sid)

    $text = $sid.Replace('S-1-12-1-', '')
    $array = [UInt32[]]$text.Split('-')

    $bytes = New-Object 'Byte[]' 16
    [Buffer]::BlockCopy($array, 0, $bytes, 0, 16)
    [Guid]$guid = $bytes

    return $guid
}
function convert-timeZone(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true,ParameterSetName = "FromCountry")]
            [parameter(Mandatory = $true,ParameterSetName = "FromISO3166")]
            [parameter(Mandatory = $true,ParameterSetName = "FromTimezone")]
            [parameter(Mandatory = $true,ParameterSetName = "FromUTC")]
            [parameter(Mandatory = $true,ParameterSetName = "FromTimezoneDescription")]
            [ValidateSet("Country","ISO3166","Timezone","UTC","TimezoneDescription")]
            [string]$getType
        #,[parameter(Mandatory = $true,ParameterSetName = "PreEncrypted")]
        #    [parameter(Mandatory = $true,ParameterSetName = "NotEncrypted")]
        ,[parameter(Mandatory = $true,ParameterSetName = "FromCountry")]
            [ValidateSet("Afghanistan","Åland Islands","Albania","Algeria","American Samoa","Andorra","Angola","Anguilla","Antarctica","Antigua and Barbuda","Argentina","Armenia","Aruba","Australia","Austria","Azerbaijan","Bahamas, The","Bahrain","Bangladesh","Barbados","Belarus","Belgium","Belize
","Benin","Bermuda","Bhutan","Bolivarian Republic of Venezuela","Bolivia","Bonaire, Sint Eustatius and Saba","Bosnia and Herzegovina","Botswana","Bouvet Island","Brazil","British Indian Ocean Territory","Brunei","Bulgaria","Burkina Faso","Burundi","Cabo Verde","Cambodia",
"Cameroon","Canada","Cayman Islands","Central African Republic","Chad","Chile","China","Christmas Island","Cocos (Keeling) Islands","Colombia","Comoros","Congo","Congo (DRC)","Cook Islands","Costa Rica","Côte d'Ivoire","Croatia","Cuba","Curaçao","Cyprus","Czech Republic",
"Democratic Republic of Timor-Leste","Denmark","Djibouti","Dominica","Dominican Republic","Ecuador","Egypt","El Salvador","Equatorial Guinea","Eritrea","Estonia","Ethiopia","Falkland Islands (Islas Malvinas)","Faroe Islands","Fiji Islands","Finland","France","French Guian
a","French Polynesia","French Southern and Antarctic Lands","Gabon","Gambia, The","Georgia","Germany","Ghana","Gibraltar","Greece","Greenland","Grenada","Guadeloupe","Guam","Guatemala","Guernsey","Guinea","Guinea-Bissau","Guyana","Haiti","Heard Island and McDonald Islands
","Honduras","Hong Kong SAR","Hungary","Iceland","India","Indonesia","Iran","Iraq","Ireland","Israel","Italy","Jamaica","Jan Mayen","Japan","Jersey","Jordan","Kazakhstan","Kenya","Kiribati","Korea","Kosovo","Kuwait","Kyrgyzstan","Laos","Latvia","Lebanon","Lesotho","Liberi
a","Libya","Liechtenstein","Lithuania","Luxembourg","Macao SAR","Macedonia, Former Yugoslav Republic of","Madagascar","Malawi","Malaysia","Maldives","Mali","Malta","Man, Isle of","Marshall Islands","Martinique","Mauritania","Mauritius","Mayotte","Mexico","Micronesia","Mol
dova","Monaco","Mongolia","Montenegro","Montserrat","Morocco","Mozambique","Myanmar","Namibia","Nauru","Nepal","Netherlands","New Caledonia","New Zealand","Nicaragua","Niger","Nigeria","Niue","Norfolk Island","North Korea","Northern Mariana Islands","Norway","Oman","Pakis
tan","Palau","Palestinian Authority","Panama","Papua New Guinea","Paraguay","Peru","Philippines","Pitcairn Islands","Poland","Portugal","Puerto Rico","Qatar","Reunion","Romania","Russia","Rwanda","Saint Barthélemy","Saint Helena, Ascension and Tristan da Cunha","Saint Kit
ts and Nevis","Saint Lucia","Saint Martin (French part)","Saint Pierre and Miquelon","Saint Vincent and the Grenadines","Samoa","San Marino","São Tomé and Príncipe","Saudi Arabia","Senegal","Serbia","Seychelles","Sierra Leone","Singapore","Sint Maarten (Dutch part)","Slov
akia","Slovenia","Solomon Islands","Somalia","South Africa","South Georgia and the South Sandwich Islands","South Sudan","Spain","Sri Lanka","Sudan","Suriname","Svalbard","Swaziland","Sweden","Switzerland","Syria","Taiwan","Tajikistan","Tanzania","Thailand","Togo","Tokela
u","Tonga","Trinidad and Tobago","Tunisia","Turkey","Turkmenistan","Turks and Caicos Islands","Tuvalu","U.S. Minor Outlying Islands","Uganda","Ukraine","United Arab Emirates","United Kingdom","United States","Uruguay","Uzbekistan","Vanuatu","Vatican City","Vietnam","Virgi
n Islands, British","Virgin Islands, U.S.","Wallis and Futuna","Yemen","Zambia","Zimbabwe")]
            [string]$fromCountry
        ,[parameter(Mandatory = $true,ParameterSetName = "FromISO3166")]
            [ValidateSet("AD","AE","AF","AG","AI","AL","AM","AO","AQ","AR","AS","AT","AU","AW","AX","AZ","BA","BB","BD","BE","BF","BG","BH","BI","BJ","BL","BM","BN","BO","BQ","BR","BS","BT","BV","BW","BY","BZ","CA","CC","CD","CF","CG","CH","CI","CK","CL","CM","CN","CO","CR","CU","CV","CW","CX","CY
","CZ","DE","DJ","DK","DM","DO","DZ","EC","EE","EG","ER","ES","ET","FI","FJ","FK","FM","FO","FR","GA","GB","GD","GE","GF","GG","GH","GI","GL","GM","GN","GP","GQ","GR","GS","GT","GU","GW","GY","HK","HM","HN","HR","HT","HU","ID","IE","IL","IM","IN","IO","IQ","IR","IS","IT",
"JE","JM","JO","JP","KE","KG","KH","KI","KM","KN","KP","KR","KW","KY","KZ","LA","LB","LC","LI","LK","LR","LS","LT","LU","LV","LY","MA","MC","MD","ME","MF","MG","MH","MK","ML","MM","MN","MO","MP","MQ","MR","MS","MT","MU","MV","MW","MX","MY","MZ","NA","NC","NE","NF","NG","N
I","NL","NO","NP","NR","NU","NZ","OM","PA","PE","PF","PG","PH","PK","PL","PM","PN","PR","PS","PT","PW","PY","QA","RE","RO","RS","RU","RW","SA","SB","SC","SD","SE","SG","SH","SI","SJ","SK","SL","SM","SN","SO","SR","SS","ST","SV","SX","SY","SZ","TC","TD","TF","TG","TH","TJ"
,"TK","TL","TM","TN","TO","TR","TT","TV","TW","TZ","UA","UG","UM","US","UY","UZ","VA","VC","VE","VG","VI","VN","VU","WF","WS","XK","YE","YT","ZA","ZM","ZW")]
            [string]$fromISO3166
        ,[parameter(Mandatory = $true,ParameterSetName = "FromTimezone")]
            [ValidateSet("Afghanistan Standard Time","Arab Standard Time","Arabian Standard Time","Arabic Standard Time","Argentina Standard Time","Atlantic Standard Time","AUS Eastern Standard Time","Azerbaijan Standard Time","Bangladesh Standard Time","Belarus Standard Time","Cape Verde Standard
 Time","Caucasus Standard Time","Central America Standard Time","Central Asia Standard Time","Central Europe Standard Time","Central European Standard Time","Central Pacific Standard Time","Central Standard Time (Mexico)","China Standard Time","E. Africa Standard Time","E
. Europe Standard Time","E. South America Standard Time","Eastern Standard Time","Egypt Standard Time","Fiji Standard Time","FLE Standard Time","Georgian Standard Time","GMT Standard Time","Greenland Standard Time","Greenwich Standard Time","GTB Standard Time","Hawaiian S
tandard Time","India Standard Time","Iran Standard Time","Israel Standard Time","Jordan Standard Time","Korea Standard Time","Mauritius Standard Time","Middle East Standard Time","Montevideo Standard Time","Morocco Standard Time","Myanmar Standard Time","Namibia Standard 
Time","Nepal Standard Time","New Zealand Standard Time","Pacific SA Standard Time","Pacific Standard Time","Pakistan Standard Time","Paraguay Standard Time","Romance Standard Time","Russian Standard Time","SA Eastern Standard Time","SA Pacific Standard Time","SA Western S
tandard Time","Samoa Standard Time","SE Asia Standard Time","Singapore Standard Time","South Africa Standard Time","Sri Lanka Standard Time","Syria Standard Time","Taipei Standard Time","Tokyo Standard Time","Tonga Standard Time","Turkey Standard Time","Ulaanbaatar Standa
rd Time","UTC","UTC+12","UTC-02","UTC-11","Venezuela Standard Time","W. Central Africa Standard Time","W. Europe Standard Time","West Asia Standard Time","West Pacific Standard Time")]
            [string]$fromTimezone
        ,[parameter(Mandatory = $true,ParameterSetName = "FromUTC")]
            [ValidateSet("(UTC)","(UTC+01:00)","(UTC+02:00)","(UTC+03:00)","(UTC+03:30)","(UTC+04:00)","(UTC+04:30)","(UTC+05:00)","(UTC+05:30)","(UTC+05:45)","(UTC+06:00)","(UTC+06:30)","(UTC+07:00)","(UTC+08:00)","(UTC+09:00)","(UTC+10:00)","(UTC+11:00)","(UTC+12:00)","(UTC+13:00)","(UTC-01:00)"
,"(UTC-02:00)","(UTC-03:00)","(UTC-04:00)","(UTC-04:30)","(UTC-05:00)","(UTC-06:00)","(UTC-08:00)","(UTC-10:00)","(UTC-11:00)")]
            [string]$fromUTC
        ,[parameter(Mandatory = $true,ParameterSetName = "FromTimezoneDescription")]
            [ValidateSet("Abu Dhabi, Muscat","Amman","Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna","Ashgabat, Tashkent","Astana","Asuncion","Athens, Bucharest","Atlantic Time (Canada)","Auckland, Wellington","Baghdad","Baku","Bangkok, Hanoi, Jakarta","Beijing, Chongqing, Hong Kong, Urumqi","B
eirut","Belgrade, Bratislava, Budapest, Ljubljana, Prague","Bogota, Lima, Quito, Rio Branco","Brasilia","Brussels, Copenhagen, Madrid, Paris","Cabo Verde Is.","Cairo","Canberra, Melbourne, Sydney","Caracas","Casablanca","Cayenne, Fortaleza","Central America","Chennai, Kol
kata, Mumbai, New Delhi","City of Buenos Aires","Coordinated Universal Time","Coordinated Universal Time+12","Coordinated Universal Time-02","Coordinated Universal Time-11","Damascus","Dhaka","Dublin, Edinburgh, Lisbon, London","E. Europe","Eastern Time (US & Canada)","Fi
ji","Georgetown, La Paz, Manaus, San Juan","Greenland","Guadalajara, Mexico City, Monterrey","Guam, Port Moresby","Harare, Pretoria","Harare, Pretoria","Hawaii","Helsinki, Kyiv, Riga, Sofia, Tallinn, Vilnius","Islamabad, Karachi","Istanbul","Jerusalem","Kabul","Kathmandu"
,"Kuala Lumpur, Singapore","Kuwait, Riyadh","Minsk","Monrovia, Reykjavik","Montevideo","Moscow, St. Petersburg, Volgograd (RTZ 2)","Nairobi","Nuku'alofa","Osaka, Sapporo, Tokyo","Pacific Time (US & Canada)","Port Louis","Samoa","Santiago","Sarajevo, Skopje, Warsaw, Zagreb
","Seoul","Solomon Is., New Caledonia","Sri Jayawardenepura","Taipei","Tbilisi","Tehran","Ulaanbaatar","West Central Africa","Windhoek","Yangon (Rangoon)","Yerevan")]
            [string]$fromTimezoneDescription
        )
    #Text comes from https://docs.microsoft.com/en-us/windows-hardware/manufacture/desktop/default-time-zones
    #The headers are: Country[0], ISO3166[1], Timezone[2], UTC[3], Timezone description[4]
    $rawText = "Afghanistan	AF	Afghanistan Standard Time	(UTC+04:30)	Kabul
Åland Islands	AX	FLE Standard Time	(UTC+02:00)	Helsinki, Kyiv, Riga, Sofia, Tallinn, Vilnius
Albania	AL	Central Europe Standard Time	(UTC+01:00)	Belgrade, Bratislava, Budapest, Ljubljana, Prague
Algeria	DZ	W. Central Africa Standard Time	(UTC+01:00)	West Central Africa
American Samoa	AS	UTC-11	(UTC-11:00)	Coordinated Universal Time-11
Andorra	AD	W. Europe Standard Time	(UTC+01:00)	Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna
Angola	AO	W. Central Africa Standard Time	(UTC+01:00)	West Central Africa
Anguilla	AI	SA Western Standard Time	(UTC-04:00)	Georgetown, La Paz, Manaus, San Juan
Antarctica	AQ	Pacific SA Standard Time	(UTC-03:00)	Santiago
Antigua and Barbuda	AG	SA Western Standard Time	(UTC-04:00)	Georgetown, La Paz, Manaus, San Juan
Argentina	AR	Argentina Standard Time	(UTC-03:00)	City of Buenos Aires
Armenia	AM	Caucasus Standard Time	(UTC+04:00)	Yerevan
Aruba	AW	SA Western Standard Time	(UTC-04:00)	Georgetown, La Paz, Manaus, San Juan
Australia	AU	AUS Eastern Standard Time	(UTC+10:00)	Canberra, Melbourne, Sydney
Austria	AT	W. Europe Standard Time	(UTC+01:00)	Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna
Azerbaijan	AZ	Azerbaijan Standard Time	(UTC+04:00)	Baku
Bahamas, The	BS	Eastern Standard Time	(UTC-05:00)	Eastern Time (US & Canada)
Bahrain	BH	Arab Standard Time	(UTC+03:00)	Kuwait, Riyadh
Bangladesh	BD	Bangladesh Standard Time	(UTC+06:00)	Dhaka
Barbados	BB	SA Western Standard Time	(UTC-04:00)	Georgetown, La Paz, Manaus, San Juan
Belarus	BY	Belarus Standard Time	(UTC+03:00)	Minsk
Belgium	BE	Romance Standard Time	(UTC+01:00)	Brussels, Copenhagen, Madrid, Paris
Belize	BZ	Central America Standard Time	(UTC-06:00)	Central America
Benin	BJ	W. Central Africa Standard Time	(UTC+01:00)	West Central Africa
Bermuda	BM	Atlantic Standard Time	(UTC-04:00)	Atlantic Time (Canada)
Bhutan	BT	Bangladesh Standard Time	(UTC+06:00)	Dhaka
Bolivarian Republic of Venezuela	VE	Venezuela Standard Time	(UTC-04:30)	Caracas
Bolivia	BO	SA Western Standard Time	(UTC-04:00)	Georgetown, La Paz, Manaus, San Juan
Bonaire, Sint Eustatius and Saba	BQ	SA Western Standard Time	(UTC-04:00)	Georgetown, La Paz, Manaus, San Juan
Bosnia and Herzegovina	BA	Central European Standard Time	(UTC+01:00)	Sarajevo, Skopje, Warsaw, Zagreb
Botswana	BW	South Africa Standard Time	(UTC+02:00)	Harare, Pretoria
Bouvet Island	BV	UTC	(UTC)	Coordinated Universal Time
Brazil	BR	E. South America Standard Time	(UTC-03:00)	Brasilia
British Indian Ocean Territory	IO	Central Asia Standard Time	(UTC+06:00)	Astana
Brunei	BN	Singapore Standard Time	(UTC+08:00)	Kuala Lumpur, Singapore
Bulgaria	BG	FLE Standard Time	(UTC+02:00)	Helsinki, Kyiv, Riga, Sofia, Tallinn, Vilnius
Burkina Faso	BF	Greenwich Standard Time	(UTC)	Monrovia, Reykjavik
Burundi	BI	South Africa Standard Time	(UTC+02:00)	Harare, Pretoria
Cabo Verde	CV	Cape Verde Standard Time	(UTC-01:00)	Cabo Verde Is.
Cambodia	KH	SE Asia Standard Time	(UTC+07:00)	Bangkok, Hanoi, Jakarta
Cameroon	CM	W. Central Africa Standard Time	(UTC+01:00)	West Central Africa
Canada	CA	Eastern Standard Time	(UTC-05:00)	Eastern Time (US & Canada)
Cayman Islands	KY	SA Pacific Standard Time	(UTC-05:00)	Bogota, Lima, Quito, Rio Branco
Central African Republic	CF	W. Central Africa Standard Time	(UTC+01:00)	West Central Africa
Chad	TD	W. Central Africa Standard Time	(UTC+01:00)	West Central Africa
Chile	CL	Pacific SA Standard Time	(UTC-03:00)	Santiago
China	CN	China Standard Time	(UTC+08:00)	Beijing, Chongqing, Hong Kong, Urumqi
Christmas Island	CX	SE Asia Standard Time	(UTC+07:00)	Bangkok, Hanoi, Jakarta
Cocos (Keeling) Islands	CC	Myanmar Standard Time	(UTC+06:30)	Yangon (Rangoon)
Colombia	CO	SA Pacific Standard Time	(UTC-05:00)	Bogota, Lima, Quito, Rio Branco
Comoros	KM	E. Africa Standard Time	(UTC+03:00)	Nairobi
Congo	CG	W. Central Africa Standard Time	(UTC+01:00)	West Central Africa
Congo (DRC)	CD	W. Central Africa Standard Time	(UTC+01:00)	West Central Africa
Cook Islands	CK	Hawaiian Standard Time	(UTC-10:00)	Hawaii
Costa Rica	CR	Central America Standard Time	(UTC-06:00)	Central America
Côte d'Ivoire	CI	Greenwich Standard Time	(UTC)	Monrovia, Reykjavik
Croatia	HR	Central European Standard Time	(UTC+01:00)	Sarajevo, Skopje, Warsaw, Zagreb
Cuba	CU	Eastern Standard Time	(UTC-05:00)	Eastern Time (US & Canada)
Curaçao	CW	SA Western Standard Time	(UTC-04:00)	Georgetown, La Paz, Manaus, San Juan
Cyprus	CY	E. Europe Standard Time	(UTC+02:00)	E. Europe
Czech Republic	CZ	Central Europe Standard Time	(UTC+01:00)	Belgrade, Bratislava, Budapest, Ljubljana, Prague
Democratic Republic of Timor-Leste	TL	Tokyo Standard Time	(UTC+09:00)	Osaka, Sapporo, Tokyo
Denmark	DK	Romance Standard Time	(UTC+01:00)	Brussels, Copenhagen, Madrid, Paris
Djibouti	DJ	E. Africa Standard Time	(UTC+03:00)	Nairobi
Dominica	DM	SA Western Standard Time	(UTC-04:00)	Georgetown, La Paz, Manaus, San Juan
Dominican Republic	DO	SA Western Standard Time	(UTC-04:00)	Georgetown, La Paz, Manaus, San Juan
Ecuador	EC	SA Pacific Standard Time	(UTC-05:00)	Bogota, Lima, Quito, Rio Branco
Egypt	EG	Egypt Standard Time	(UTC+02:00)	Cairo
El Salvador	SV	Central America Standard Time	(UTC-06:00)	Central America
Equatorial Guinea	GQ	W. Central Africa Standard Time	(UTC+01:00)	West Central Africa
Eritrea	ER	E. Africa Standard Time	(UTC+03:00)	Nairobi
Estonia	EE	FLE Standard Time	(UTC+02:00)	Helsinki, Kyiv, Riga, Sofia, Tallinn, Vilnius
Ethiopia	ET	E. Africa Standard Time	(UTC+03:00)	Nairobi
Falkland Islands (Islas Malvinas)	FK	SA Eastern Standard Time	(UTC-03:00)	Cayenne, Fortaleza
Faroe Islands	FO	GMT Standard Time	(UTC)	Dublin, Edinburgh, Lisbon, London
Fiji Islands	FJ	Fiji Standard Time	(UTC+12:00)	Fiji
Finland	FI	FLE Standard Time	(UTC+02:00)	Helsinki, Kyiv, Riga, Sofia, Tallinn, Vilnius
France	FR	Romance Standard Time	(UTC+01:00)	Brussels, Copenhagen, Madrid, Paris
French Guiana	GF	SA Eastern Standard Time	(UTC-03:00)	Cayenne, Fortaleza
French Polynesia	PF	Hawaiian Standard Time	(UTC-10:00)	Hawaii
French Southern and Antarctic Lands	TF	West Asia Standard Time	(UTC+05:00)	Ashgabat, Tashkent
Gabon	GA	W. Central Africa Standard Time	(UTC+01:00)	West Central Africa
Gambia, The	GM	Greenwich Standard Time	(UTC)	Monrovia, Reykjavik
Georgia	GE	Georgian Standard Time	(UTC+04:00)	Tbilisi
Germany	DE	W. Europe Standard Time	(UTC+01:00)	Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna
Ghana	GH	Greenwich Standard Time	(UTC)	Monrovia, Reykjavik
Gibraltar	GI	W. Europe Standard Time	(UTC+01:00)	Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna
Greece	GR	GTB Standard Time	(UTC+02:00)	Athens, Bucharest
Greenland	GL	Greenland Standard Time	(UTC-03:00)	Greenland
Grenada	GD	SA Western Standard Time	(UTC-04:00)	Georgetown, La Paz, Manaus, San Juan
Guadeloupe	GP	SA Western Standard Time	(UTC-04:00)	Georgetown, La Paz, Manaus, San Juan
Guam	GU	West Pacific Standard Time	(UTC+10:00)	Guam, Port Moresby
Guatemala	GT	Central America Standard Time	(UTC-06:00)	Central America
Guernsey	GG	GMT Standard Time	(UTC)	Dublin, Edinburgh, Lisbon, London
Guinea	GN	Greenwich Standard Time	(UTC)	Monrovia, Reykjavik
Guinea-Bissau	GW	Greenwich Standard Time	(UTC)	Monrovia, Reykjavik
Guyana	GY	SA Western Standard Time	(UTC-04:00)	Georgetown, La Paz, Manaus, San Juan
Haiti	HT	Eastern Standard Time	(UTC-05:00)	Eastern Time (US & Canada)
Heard Island and McDonald Islands	HM	Mauritius Standard Time	(UTC+04:00)	Port Louis
Honduras	HN	Central America Standard Time	(UTC-06:00)	Central America
Hong Kong SAR	HK	China Standard Time	(UTC+08:00)	Beijing, Chongqing, Hong Kong, Urumqi
Hungary	HU	Central Europe Standard Time	(UTC+01:00)	Belgrade, Bratislava, Budapest, Ljubljana, Prague
Iceland	IS	Greenwich Standard Time	(UTC)	Monrovia, Reykjavik
India	IN	India Standard Time	(UTC+05:30)	Chennai, Kolkata, Mumbai, New Delhi
Indonesia	ID	SE Asia Standard Time	(UTC+07:00)	Bangkok, Hanoi, Jakarta
Iran	IR	Iran Standard Time	(UTC+03:30)	Tehran
Iraq	IQ	Arabic Standard Time	(UTC+03:00)	Baghdad
Ireland	IE	GMT Standard Time	(UTC)	Dublin, Edinburgh, Lisbon, London
Israel	IL	Israel Standard Time	(UTC+02:00)	Jerusalem
Italy	IT	W. Europe Standard Time	(UTC+01:00)	Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna
Jamaica	JM	SA Pacific Standard Time	(UTC-05:00)	Bogota, Lima, Quito, Rio Branco
Jan Mayen	SJ	W. Europe Standard Time	(UTC+01:00)	Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna
Japan	JP	Tokyo Standard Time	(UTC+09:00)	Osaka, Sapporo, Tokyo
Jersey	JE	GMT Standard Time	(UTC)	Dublin, Edinburgh, Lisbon, London
Jordan	JO	Jordan Standard Time	(UTC+02:00)	Amman
Kazakhstan	KZ	Central Asia Standard Time	(UTC+06:00)	Astana
Kenya	KE	E. Africa Standard Time	(UTC+03:00)	Nairobi
Kiribati	KI	UTC+12	(UTC+12:00)	Coordinated Universal Time+12
Korea	KR	Korea Standard Time	(UTC+09:00)	Seoul
Kosovo	XK	Central European Standard Time	(UTC+01:00)	Sarajevo, Skopje, Warsaw, Zagreb
Kuwait	KW	Arab Standard Time	(UTC+03:00)	Kuwait, Riyadh
Kyrgyzstan	KG	Central Asia Standard Time	(UTC+06:00)	Astana
Laos	LA	SE Asia Standard Time	(UTC+07:00)	Bangkok, Hanoi, Jakarta
Latvia	LV	FLE Standard Time	(UTC+02:00)	Helsinki, Kyiv, Riga, Sofia, Tallinn, Vilnius
Lebanon	LB	Middle East Standard Time	(UTC+02:00)	Beirut
Lesotho	LS	South Africa Standard Time	(UTC+02:00)	Harare, Pretoria
Liberia	LR	Greenwich Standard Time	(UTC)	Monrovia, Reykjavik
Libya	LY	E. Europe Standard Time	(UTC+02:00)	E. Europe
Liechtenstein	LI	W. Europe Standard Time	(UTC+01:00)	Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna
Lithuania	LT	FLE Standard Time	(UTC+02:00)	Helsinki, Kyiv, Riga, Sofia, Tallinn, Vilnius
Luxembourg	LU	W. Europe Standard Time	(UTC+01:00)	Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna
Macao SAR	MO	China Standard Time	(UTC+08:00)	Beijing, Chongqing, Hong Kong, Urumqi
Macedonia, Former Yugoslav Republic of	MK	Central European Standard Time	(UTC+01:00)	Sarajevo, Skopje, Warsaw, Zagreb
Madagascar	MG	E. Africa Standard Time	(UTC+03:00)	Nairobi
Malawi	MW	South Africa Standard Time	(UTC+02:00)	Harare, Pretoria
Malaysia	MY	Singapore Standard Time	(UTC+08:00)	Kuala Lumpur, Singapore
Maldives	MV	West Asia Standard Time	(UTC+05:00)	Ashgabat, Tashkent
Mali	ML	Greenwich Standard Time	(UTC)	Monrovia, Reykjavik
Malta	MT	W. Europe Standard Time	(UTC+01:00)	Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna
Man, Isle of	IM	GMT Standard Time	(UTC)	Dublin, Edinburgh, Lisbon, London
Marshall Islands	MH	UTC+12	(UTC+12:00)	Coordinated Universal Time+12
Martinique	MQ	SA Western Standard Time	(UTC-04:00)	Georgetown, La Paz, Manaus, San Juan
Mauritania	MR	Greenwich Standard Time	(UTC)	Monrovia, Reykjavik
Mauritius	MU	Mauritius Standard Time	(UTC+04:00)	Port Louis
Mayotte	YT	E. Africa Standard Time	(UTC+03:00)	Nairobi
Mexico	MX	Central Standard Time (Mexico)	(UTC-06:00)	Guadalajara, Mexico City, Monterrey
Micronesia	FM	West Pacific Standard Time	(UTC+10:00)	Guam, Port Moresby
Moldova	MD	GTB Standard Time	(UTC+02:00)	Athens, Bucharest
Monaco	MC	W. Europe Standard Time	(UTC+01:00)	Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna
Mongolia	MN	Ulaanbaatar Standard Time	(UTC+08:00)	Ulaanbaatar
Montenegro	ME	Central European Standard Time	(UTC+01:00)	Sarajevo, Skopje, Warsaw, Zagreb
Montserrat	MS	SA Western Standard Time	(UTC-04:00)	Georgetown, La Paz, Manaus, San Juan
Morocco	MA	Morocco Standard Time	(UTC)	Casablanca
Mozambique	MZ	South Africa Standard Time	(UTC+02:00)	Harare, Pretoria
Myanmar	MM	Myanmar Standard Time	(UTC+06:30)	Yangon (Rangoon)
Namibia	NA	Namibia Standard Time	(UTC+01:00)	Windhoek
Nauru	NR	UTC+12	(UTC+12:00)	Coordinated Universal Time+12
Nepal	NP	Nepal Standard Time	(UTC+05:45)	Kathmandu
Netherlands	NL	W. Europe Standard Time	(UTC+01:00)	Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna
New Caledonia	NC	Central Pacific Standard Time	(UTC+11:00)	Solomon Is., New Caledonia
New Zealand	NZ	New Zealand Standard Time	(UTC+12:00)	Auckland, Wellington
Nicaragua	NI	Central America Standard Time	(UTC-06:00)	Central America
Niger	NE	W. Central Africa Standard Time	(UTC+01:00)	West Central Africa
Nigeria	NG	W. Central Africa Standard Time	(UTC+01:00)	West Central Africa
Niue	NU	UTC-11	(UTC-11:00)	Coordinated Universal Time-11
Norfolk Island	NF	Central Pacific Standard Time	(UTC+11:00)	Solomon Is., New Caledonia
North Korea	KP	Korea Standard Time	(UTC+09:00)	Seoul
Northern Mariana Islands	MP	West Pacific Standard Time	(UTC+10:00)	Guam, Port Moresby
Norway	NO	W. Europe Standard Time	(UTC+01:00)	Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna
Oman	OM	Arabian Standard Time	(UTC+04:00)	Abu Dhabi, Muscat
Pakistan	PK	Pakistan Standard Time	(UTC+05:00)	Islamabad, Karachi
Palau	PW	Tokyo Standard Time	(UTC+09:00)	Osaka, Sapporo, Tokyo
Palestinian Authority	PS	Egypt Standard Time	(UTC+02:00)	Cairo
Panama	PA	SA Pacific Standard Time	(UTC-05:00)	Bogota, Lima, Quito, Rio Branco
Papua New Guinea	PG	West Pacific Standard Time	(UTC+10:00)	Guam, Port Moresby
Paraguay	PY	Paraguay Standard Time	(UTC-04:00)	Asuncion
Peru	PE	SA Pacific Standard Time	(UTC-05:00)	Bogota, Lima, Quito, Rio Branco
Philippines	PH	Singapore Standard Time	(UTC+08:00)	Kuala Lumpur, Singapore
Pitcairn Islands	PN	Pacific Standard Time	(UTC-08:00)	Pacific Time (US & Canada)
Poland	PL	Central European Standard Time	(UTC+01:00)	Sarajevo, Skopje, Warsaw, Zagreb
Portugal	PT	GMT Standard Time	(UTC)	Dublin, Edinburgh, Lisbon, London
Puerto Rico	PR	SA Western Standard Time	(UTC-04:00)	Georgetown, La Paz, Manaus, San Juan
Qatar	QA	Arab Standard Time	(UTC+03:00)	Kuwait, Riyadh
Reunion	RE	Mauritius Standard Time	(UTC+04:00)	Port Louis
Romania	RO	GTB Standard Time	(UTC+02:00)	Athens, Bucharest
Russia	RU	Russian Standard Time	(UTC+03:00)	Moscow, St. Petersburg, Volgograd (RTZ 2)
Rwanda	RW	South Africa Standard Time	(UTC+02:00)	Harare, Pretoria
Saint Barthélemy	BL	SA Western Standard Time	(UTC-04:00)	Georgetown, La Paz, Manaus, San Juan
Saint Helena, Ascension and Tristan da Cunha	SH	Greenwich Standard Time	(UTC)	Monrovia, Reykjavik
Saint Kitts and Nevis	KN	SA Western Standard Time	(UTC-04:00)	Georgetown, La Paz, Manaus, San Juan
Saint Lucia	LC	SA Western Standard Time	(UTC-04:00)	Georgetown, La Paz, Manaus, San Juan
Saint Martin (French part)	MF	SA Western Standard Time	(UTC-04:00)	Georgetown, La Paz, Manaus, San Juan
Saint Pierre and Miquelon	PM	Greenland Standard Time	(UTC-03:00)	Greenland
Saint Vincent and the Grenadines	VC	SA Western Standard Time	(UTC-04:00)	Georgetown, La Paz, Manaus, San Juan
Samoa	WS	Samoa Standard Time	(UTC+13:00)	Samoa
San Marino	SM	W. Europe Standard Time	(UTC+01:00)	Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna
São Tomé and Príncipe	ST	Greenwich Standard Time	(UTC)	Monrovia, Reykjavik
Saudi Arabia	SA	Arab Standard Time	(UTC+03:00)	Kuwait, Riyadh
Senegal	SN	Greenwich Standard Time	(UTC)	Monrovia, Reykjavik
Serbia	RS	Central Europe Standard Time	(UTC+01:00)	Belgrade, Bratislava, Budapest, Ljubljana, Prague
Seychelles	SC	Mauritius Standard Time	(UTC+04:00)	Port Louis
Sierra Leone	SL	Greenwich Standard Time	(UTC)	Monrovia, Reykjavik
Singapore	SG	Singapore Standard Time	(UTC+08:00)	Kuala Lumpur, Singapore
Sint Maarten (Dutch part)	SX	SA Western Standard Time	(UTC-04:00)	Georgetown, La Paz, Manaus, San Juan
Slovakia	SK	Central Europe Standard Time	(UTC+01:00)	Belgrade, Bratislava, Budapest, Ljubljana, Prague
Slovenia	SI	Central Europe Standard Time	(UTC+01:00)	Belgrade, Bratislava, Budapest, Ljubljana, Prague
Solomon Islands	SB	Central Pacific Standard Time	(UTC+11:00)	Solomon Is., New Caledonia
Somalia	SO	E. Africa Standard Time	(UTC+03:00)	Nairobi
South Africa	ZA	South Africa Standard Time	(UTC+02:00)	Harare, Pretoria
South Georgia and the South Sandwich Islands	GS	UTC-02	(UTC-02:00)	Coordinated Universal Time-02
South Sudan	SS	E. Africa Standard Time	(UTC+03:00)	Nairobi
Spain	ES	Romance Standard Time	(UTC+01:00)	Brussels, Copenhagen, Madrid, Paris
Sri Lanka	LK	Sri Lanka Standard Time	(UTC+05:30)	Sri Jayawardenepura
Sudan	SD	E. Africa Standard Time	(UTC+03:00)	Nairobi
Suriname	SR	SA Eastern Standard Time	(UTC-03:00)	Cayenne, Fortaleza
Svalbard	SJ	W. Europe Standard Time	(UTC+01:00)	Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna
Swaziland	SZ	South Africa Standard Time	(UTC+02:00)	Harare, Pretoria
Sweden	SE	W. Europe Standard Time	(UTC+01:00)	Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna
Switzerland	CH	W. Europe Standard Time	(UTC+01:00)	Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna
Syria	SY	Syria Standard Time	(UTC+02:00)	Damascus
Taiwan	TW	Taipei Standard Time	(UTC+08:00)	Taipei
Tajikistan	TJ	West Asia Standard Time	(UTC+05:00)	Ashgabat, Tashkent
Tanzania	TZ	E. Africa Standard Time	(UTC+03:00)	Nairobi
Thailand	TH	SE Asia Standard Time	(UTC+07:00)	Bangkok, Hanoi, Jakarta
Togo	TG	Greenwich Standard Time	(UTC)	Monrovia, Reykjavik
Tokelau	TK	Tonga Standard Time	(UTC+13:00)	Nuku'alofa
Tonga	TO	Tonga Standard Time	(UTC+13:00)	Nuku'alofa
Trinidad and Tobago	TT	SA Western Standard Time	(UTC-04:00)	Georgetown, La Paz, Manaus, San Juan
Tunisia	TN	W. Central Africa Standard Time	(UTC+01:00)	West Central Africa
Turkey	TR	Turkey Standard Time	(UTC+02:00)	Istanbul
Turkmenistan	TM	West Asia Standard Time	(UTC+05:00)	Ashgabat, Tashkent
Turks and Caicos Islands	TC	Eastern Standard Time	(UTC-05:00)	Eastern Time (US & Canada)
Tuvalu	TV	UTC+12	(UTC+12:00)	Coordinated Universal Time+12
U.S. Minor Outlying Islands	UM	UTC-11	(UTC-11:00)	Coordinated Universal Time-11
Uganda	UG	E. Africa Standard Time	(UTC+03:00)	Nairobi
Ukraine	UA	FLE Standard Time	(UTC+02:00)	Helsinki, Kyiv, Riga, Sofia, Tallinn, Vilnius
United Arab Emirates	AE	Arabian Standard Time	(UTC+04:00)	Abu Dhabi, Muscat
United Kingdom	GB	GMT Standard Time	(UTC)	Dublin, Edinburgh, Lisbon, London
United States	US	Pacific Standard Time	(UTC-08:00)	Pacific Time (US & Canada)
Uruguay	UY	Montevideo Standard Time	(UTC-03:00)	Montevideo
Uzbekistan	UZ	West Asia Standard Time	(UTC+05:00)	Ashgabat, Tashkent
Vanuatu	VU	Central Pacific Standard Time	(UTC+11:00)	Solomon Is., New Caledonia
Vatican City	VA	W. Europe Standard Time	(UTC+01:00)	Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna
Vietnam	VN	SE Asia Standard Time	(UTC+07:00)	Bangkok, Hanoi, Jakarta
Virgin Islands, U.S.	VI	SA Western Standard Time	(UTC-04:00)	Georgetown, La Paz, Manaus, San Juan
Virgin Islands, British	VG	SA Western Standard Time	(UTC-04:00)	Georgetown, La Paz, Manaus, San Juan
Wallis and Futuna	WF	UTC+12	(UTC+12:00)	Coordinated Universal Time+12
Yemen	YE	Arab Standard Time	(UTC+03:00)	Kuwait, Riyadh
Zambia	ZM	South Africa Standard Time	(UTC+02:00)	Harare, Pretoria
Zimbabwe	ZW	South Africa Standard Time	(UTC+02:00)	Harare, Pretoria"
    $timeZoneArray = @()
    $rawText.Split("`n") | % {
        #$timezone = $_.Split("`t")
        $timeZoneArray += ,@($_.Split("`t")) #The , between += and @() prevents the array from becoming unrolled as it is added
        }
    
    switch ($PsCmdlet.ParameterSetName){
        {$_ -eq "FromCountry"} {
            Write-Verbose "convert-timeZone | From Country [$fromCountry] To [$getType]"
            $timeZoneArrayFromIndex = 0
            $fromValue = $fromCountry
            }
        {$_ -eq "FromISO3166"} {
            Write-Verbose "convert-timeZone | From ISO3166 [$fromISO3166] To [$getType]"
            $timeZoneArrayFromIndex = 1
            $fromValue = $fromISO3166
            }
        {$_ -eq "FromTimezone"} {
            Write-Verbose "convert-timeZone | From Timezone [$fromTimezone] To [$getType]"
            $timeZoneArrayFromIndex = 2
            $fromValue = $fromTimezone
            }
        {$_ -eq "FromUTC"} {
            Write-Verbose "convert-timeZone | From UTC [$fromUTC] To [$getType]"
            $timeZoneArrayFromIndex = 3
            $fromValue = $fromUTC
            }
        {$_ -eq "FromTimezoneDescription"} {
            Write-Verbose "convert-timeZone | From TimezoneDescription [$fromTimezoneDescription] To [$getType]"
            $timeZoneArrayFromIndex = 4
            $fromValue = $fromTimezoneDescription
            }
        }

    switch ($getType){
        "Country"             {$timeZoneArrayToIndex = 0}
        "ISO3166"             {$timeZoneArrayToIndex = 1}
        "Timezone"            {$timeZoneArrayToIndex = 2}
        "UTC"                 {$timeZoneArrayToIndex = 3}
        "TimezoneDescription" {$timeZoneArrayToIndex = 4}
        }

    #$timeZoneArray | ? {$_[$timeZoneArrayFromIndex] -eq $fromValue} | Write-Verbose "[$($_[$timeZoneArrayFromIndex])] TimeZone found"
    $foundTimezones #()
    $timeZoneArray | ? {$_[$timeZoneArrayFromIndex] -eq $fromValue} | % {$foundTimezones += ,$_}
    Write-Verbose "[$($foundTimezones.Count)] TimeZones found"

    $foundTimezones | % {$_[$timeZoneArrayToIndex]}
    #if($foundTimezones.Count -gt 1){$foundTimezones | % {$_[$timeZoneArrayToIndex]}}
    #else {$foundTimezones[$timeZoneArrayToIndex]}
    #write-host -f Magenta $foundTimezones[$timeZoneArrayToIndex]
    #write-host -f yellow $foundTimezones
    # This section helps to generate the ValidateSet conditions for the parameters
    #$($timeZoneArray | %{$_[0]} | sort -Unique) -join "`",`""
    #$($timeZoneArray | %{$_[1]} | sort -Unique) -join "`",`""
    #$($timeZoneArray | %{$_[2]} | sort -Unique) -join "`",`""
    #$($timeZoneArray | %{$_[3]} | sort -Unique) -join "`",`""
    #$($timeZoneArray | %{$_[4]} | sort -Unique) -join "`",`""

    }
function combine-url($arrayOfStrings){ 
    $output = ""
    $arrayOfStrings | % {
        $output += $_.TrimStart("/").TrimEnd("/")+"/"
        }
    $output = $output.Substring(0,$output.Length-1)
    $output = $output.Replace("//","/").Replace("//","/").Replace("//","/")
    $output = $output.Replace("http:/","http://").Replace("https:/","https://")
    $output
    }
function compare-objectProperties {
    #https://blogs.technet.microsoft.com/janesays/2017/04/25/compare-all-properties-of-two-objects-in-windows-powershell/
    Param(
        [PSObject]$ReferenceObject,
        [PSObject]$DifferenceObject 
        )
    $objprops = $ReferenceObject | Get-Member -MemberType Property,NoteProperty | % Name
    $objprops += $DifferenceObject | Get-Member -MemberType Property,NoteProperty | % Name
    $objprops = $objprops | Sort | Select -Unique
    $diffs = @()
    foreach ($objprop in $objprops) {
        $diff = Compare-Object $ReferenceObject $DifferenceObject -Property $objprop
        if ($diff) {            
            $diffprops = @{
                PropertyName=$objprop
                RefValue=($diff | ? {$_.SideIndicator -eq '<='} | % $($objprop))
                DiffValue=($diff | ? {$_.SideIndicator -eq '=>'} | % $($objprop))
                }
            $diffs += New-Object PSObject -Property $diffprops
            }        
        }
    if ($diffs) {return ($diffs | Select PropertyName,RefValue,DiffValue)}     
    }
function convert-csvToSecureStrings(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
        [PSCustomObject]$rawCsvData        
        )
    
    $encryptedObject = New-Object psobject
    $rawCsvData.PSObject.Properties | ForEach-Object {
        $encryptedObject | Add-Member -MemberType NoteProperty -Name $_.Name -Value $(convertTo-localisedSecureString -plainText $_.Value)
        }
    $encryptedObject
    }
    Function Convert-OutputForCSV {
        <#
            .SYNOPSIS
                Provides a way to expand collections in an object property prior
                to being sent to Export-Csv.
            .DESCRIPTION
                Provides a way to expand collections in an object property prior
                to being sent to Export-Csv. This helps to avoid the object type
                from being shown such as system.object[] in a spreadsheet.
            .PARAMETER InputObject
                The object that will be sent to Export-Csv
            .PARAMETER OutPropertyType
                This determines whether the property that has the collection will be
                shown in the CSV as a comma delimmited string or as a stacked string.
                Possible values:
                Stack
                Comma
                Default value is: Stack
            .NOTES
                Name: Convert-OutputForCSV
                Author: Boe Prox
                Created: 24 Jan 2014
                Version History:
                    1.1 - 02 Feb 2014
                        -Removed OutputOrder parameter as it is no longer needed; inputobject order is now respected 
                        in the output object
                    1.0 - 24 Jan 2014
                        -Initial Creation
            .EXAMPLE
                $Output = 'PSComputername','IPAddress','DNSServerSearchOrder'
                Get-WMIObject -Class Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'" |
                Select-Object $Output | Convert-OutputForCSV | 
                Export-Csv -NoTypeInformation -Path NIC.csv    
                
                Description
                -----------
                Using a predefined set of properties to display ($Output), data is collected from the 
                Win32_NetworkAdapterConfiguration class and then passed to the Convert-OutputForCSV
                funtion which expands any property with a collection so it can be read properly prior
                to being sent to Export-Csv. Properties that had a collection will be viewed as a stack
                in the spreadsheet.        
                
        #>
        #Requires -Version 3.0
        [cmdletbinding()]
        Param (
            [parameter(ValueFromPipeline)]
            [psobject]$InputObject,
            [parameter()]
            [ValidateSet('Stack','Comma')]
            [string]$OutputPropertyType = 'Stack'
        )
        Begin {
            $PSBoundParameters.GetEnumerator() | ForEach {
                Write-Verbose "$($_)"
            }
            $FirstRun = $True
        }
        Process {
            If ($FirstRun) {
                $OutputOrder = $InputObject.psobject.properties.name
                Write-Verbose "Output Order:`n $($OutputOrder -join ', ' )"
                $FirstRun = $False
                #Get properties to process
                $Properties = Get-Member -InputObject $InputObject -MemberType *Property
                #Get properties that hold a collection
                $Properties_Collection = @(($Properties | Where-Object {
                    $_.Definition -match "Collection|\[\]"
                }).Name)
                #Get properties that do not hold a collection
                $Properties_NoCollection = @(($Properties | Where-Object {
                    $_.Definition -notmatch "Collection|\[\]"
                }).Name)
                Write-Verbose "Properties Found that have collections:`n $(($Properties_Collection) -join ', ')"
                Write-Verbose "Properties Found that have no collections:`n $(($Properties_NoCollection) -join ', ')"
            }
     
            $InputObject | ForEach {
                $Line = $_
                $stringBuilder = New-Object Text.StringBuilder
                $Null = $stringBuilder.AppendLine("[pscustomobject] @{")
    
                $OutputOrder | ForEach {
                    If ($OutputPropertyType -eq 'Stack') {
                        $Null = $stringBuilder.AppendLine("`"$($_)`" = `"$(($line.$($_) | Out-String).Trim())`"")
                    } ElseIf ($OutputPropertyType -eq "Comma") {
                        $Null = $stringBuilder.AppendLine("`"$($_)`" = `"$($line.$($_) -join ', ')`"")                   
                    }
                }
                $Null = $stringBuilder.AppendLine("}")
     
                Invoke-Expression $stringBuilder.ToString()
            }
        }
        End {}
    }
    
    function convertTo-arrayOfEmailAddresses($blockOfText){
    [string[]]$addresses = @()
    $blockOfText | %{
        if(![string]::IsNullOrWhiteSpace($_)){
            foreach($blob in $_.Split(" ").Split("`r`n").Split(";").Split(",")){
                if($blob -match "@" -and $blob -match "."){$addresses += $blob.Replace("<","").Replace(">","").Replace(";","").Trim()}
                }
            }
        }
    $addresses
    }
function convertTo-arrayOfStrings($blockOfText){
    $strings = @()
    $blockOfText | %{
        foreach($blob in $_.Split(",").Split("`r`n")){
            if(![string]::IsNullOrEmpty($blob)){$strings += $blob}
            }
        }
    $strings
    }
function convertTo-exTimeZoneValue($pAmbiguousTimeZone){
    $singleResult = @()
    $tzs = get-timeZones
    if($pAmbiguousTimeZone -match '\('){
        $tryThis = $pAmbiguousTimeZone.Replace([regex]::Match($pAmbiguousTimeZone,"\(([^)]+)\)").Groups[0].Value,"").Trim() #Get everything not between "(" and ")"
        }
    else{$tryThis = $pAmbiguousTimeZone}
    [array]$singleResult = $tzs | ? {$_.PSChildName -eq $tryThis} #Match it to the registry timezone names
    if ($singleResult.Count -eq 1){$singleResult[0].PSChildName}
    else{
        #Try something else
        }
    }
function convertTo-localisedSecureString($plainText){
    #if ($(Get-Module).Name -notcontains "_PS_Library_Forms"){Import-Module _PS_Library_Forms}
    #if (!$plainText){$plainText = form-captureText -formTitle "PlainText" -formText "Enter the plain text to be converted to a secure string" -sizeX 300 -sizeY 200}
    if(![string]::IsNullOrWhitespace($plainText)){
        ConvertTo-SecureString $plainText -AsPlainText -Force | ConvertFrom-SecureString
        }
    }
function decrypt-SecureString($secureString){
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString)
    [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    }
function export-encryptedCache(){
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
            [AllowNull()]
            [array]$objects 
        ,[Parameter(Mandatory = $true, Position = 1)]
            [ValidateSet("Client","Subcontractor","Employee","Opportunity","Project","Folders")]
            [array]$objectType 
        ,[Parameter(Mandatory = $true, Position = 1)]
            [ValidateSet("NetSuite","TermStore","SharePoint","Pretty")]
            [array]$objectSource
        )
    
    $objectSchema = [ordered]@{}
    $objects | % {
        $thisObject = $_
        Compare-Object -ReferenceObject @($($objectSchema.Keys) | % {$_.ToString()} | Select-Object) -DifferenceObject $thisObject.PSObject.Properties.Name  | ? {$_.SideIndicator -eq "=>"} | % {
            $objectSchema.Add($_.InputObject,$null)# | Add-Member -MemberType NoteProperty -Name $_ -Value $null
            #Write-Host "Adding [$($_.InputObject)] from [$($thisobject.id)]"
            }
        }
    $prettyNetSuiteObjects = @($null)*$objects.Count
    $i=0
    $objects | %{
        $thisObject = $_
        $prettyNetSuiteObjects[$i] = New-Object -TypeName PSCustomObject -Property $objectSchema
        $thisObject.PSObject.Properties.Name | % {
            $prettyNetSuiteObjects[$i].$_ = $(convertTo-localisedSecureString $thisObject.$_)
            }
        $i++
        }
        
    $prettyNetSuiteObjects | Select-Object @($($netObjectSchema.Keys) | % {$_.ToString()} | Select-Object) | Export-Csv -Path "$env:TEMP\$($objectSource)_$($objectType).csv" -NoTypeInformation -Force -Encoding UTF8
    }
function export-encryptedCsv(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true,ParameterSetName = "PreEncrypted")]
            [psobject]$encryptedCsvData        
        ,[parameter(Mandatory = $true,ParameterSetName = "NotEncrypted")]
            [psobject]$unencryptedCsvData
        ,[parameter(Mandatory = $true,ParameterSetName = "PreEncrypted")]
            [parameter(Mandatory = $true,ParameterSetName = "NotEncrypted")]
            [string]$pathToOutputCsv
        ,[parameter(Mandatory = $false,ParameterSetName = "PreEncrypted")]
            [parameter(Mandatory = $false,ParameterSetName = "NotEncrypted")]
            [switch]$force
        )
    if(!$encryptedCsvData){
        $encryptedCsvData = convert-csvToSecureStrings -rawCsvData $unencryptedCsvData
        }
    if(Test-Path $pathToOutputCsv){
        if($force){Remove-Item -Path $pathToOutputCsv -Force}
        else{Write-Error "File [$pathToOutputCsv] already exists";break}
        }
    Export-Csv -InputObject $encryptedCsvData -Path $pathToOutputCsv -NoTypeInformation -NoClobber
    remove-doubleQuotesFromCsv -inputFile $pathToOutputCsv
    }
function format-internationalPhoneNumber($pDirtyNumber,$p3letterIsoCountryCode,[boolean]$localise){
    if($pDirtyNumber.Length -gt 0){
        $dirtynumber = $pDirtyNumber.Split("ext")[0]
        $dirtynumber = $dirtyNumber.Trim() -replace '[^0-9]+',''
        switch ($p3letterIsoCountryCode){
            "ARE" {
                if($dirtyNumber.Length -eq 10 -and $dirtyNumber.Substring(0,1) -eq "0"){$dirtyNumber = $dirtyNumber.Substring(1,9)}
                if($dirtyNumber.Length -eq 12 -and $dirtyNumber.Substring(0,3) -eq "971"){$dirtyNumber = $dirtyNumber.Substring(3,9)}
                if($dirtyNumber.Length -eq 9){
                    if ($localise){}
                    else{$cleanNumber = "+971 $dirtyNumber"}
                    }
                }
            "CAN" {
                if($dirtyNumber.Length -eq 11 -and $dirtyNumber.Substring(0,1) -eq "1"){$dirtyNumber = $dirtyNumber.Substring(1,10)}
                if($dirtyNumber.Length -eq 10){
                    if ($localise){$cleanNumber = "+1 " + $dirtyNumber.Substring(1,3) + "-"+$dirtyNumber.Substring(4,3)+"-"+$dirtyNumber.Substring(7,4)}
                    else{$cleanNumber = "+1 $dirtyNumber"}
                    }
                }
            "CHN" {
                if($dirtyNumber.Length -eq 13 -and $dirtyNumber.Substring(0,2) -eq "86"){$dirtyNumber = $dirtyNumber.Substring(2,11)}
                if($dirtyNumber.Length -eq 11){
                    if ($localise){}
                    else{$cleanNumber = "+86 $dirtyNumber"}
                    }
                }
            "DEU" {
                if($dirtyNumber.Length -eq 12 -and $dirtyNumber.Substring(0,2) -eq "49"){$dirtyNumber = $dirtyNumber.Substring(2,10)}
                if($dirtyNumber.Length -eq 10){
                    if ($localise){}
                    else{$cleanNumber = "+49 $dirtyNumber"}
                    }
                }
            "ESP" {"ES"}
            "FIN" {
                if($dirtyNumber.Length -eq 12 -and $dirtyNumber.Substring(0,3) -eq "358"){$dirtyNumber = $dirtyNumber.Substring(3,9)}
                if($dirtyNumber.Length -eq 9){
                    if ($localise){}
                    else{$cleanNumber = "+358 $dirtyNumber"}
                    }
                }
            "GBR" {
                if($dirtyNumber.Length -eq 11 -and $dirtyNumber.Substring(0,1) -eq "0"){$dirtyNumber = $dirtyNumber.Substring(1,10)}
                if($dirtyNumber.Length -eq 12 -and $dirtyNumber.Substring(0,2) -eq "44"){$dirtyNumber = $dirtyNumber.Substring(2,10)}
                if($dirtyNumber.Length -eq 10){
                    if ($localise){}
                    else{$cleanNumber = "+44 $dirtyNumber"}
                    }
                }
            "IRL" {
                if($dirtyNumber.Substring(0,1) -eq "0"){$dirtyNumber = $dirtyNumber.Substring(1,$dirtyNumber.Length-1)}
                if($dirtyNumber.Substring(0,3) -eq "353"){$dirtyNumber = $dirtyNumber.Substring(3,$dirtyNumber.Length-3)}
                if ($localise){}
                else{$cleanNumber = "+353 $dirtyNumber"}
                }
            "PHL" {
                if($dirtyNumber.Length -eq 12 -and $dirtyNumber.Substring(0,2) -eq "63"){$dirtyNumber = $dirtyNumber.Substring(2,10)}
                if($dirtyNumber.Length -eq 10){
                    if ($localise){}
                    else{$cleanNumber = "+63 $dirtyNumber"}
                    }
                }
            "SWE" {
                if($dirtyNumber.Length -eq 11 -and $dirtyNumber.Substring(0,2) -eq "46"){$dirtyNumber = $dirtyNumber.Substring(2,9)}
                if($dirtyNumber.Length -eq 9){
                    if ($localise){}
                    else{$cleanNumber = "+46 $dirtyNumber"}
                    }
                }
            "USA" {
                if($dirtyNumber.Length -eq 11 -and $dirtyNumber.Substring(0,1) -eq 1){$dirtyNumber = $dirtyNumber.Substring(1,10)}
                if($dirtyNumber.Length -eq 10){
                    if ($localise){$cleanNumber = "+1 (" + $dirtyNumber.Substring(1,3) + ") "+$dirtyNumber.Substring(4,3)+"-"+$dirtyNumber.Substring(7,4)}
                    else{$cleanNumber = "+1 $dirtyNumber"}
                    }
                }
            }
        }
    if($cleanNumber -eq $null){$cleanNumber = $pDirtyNumber}
    $cleanNumber
    }
function format-measureCommandResults(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
            [TimeSpan]$timeSpan
            )
    switch($timeSpan){
        {$_.TotalSeconds -lt 1} {"[$([int]$timeSpan.TotalMilliseconds)] milliseconds"}
        {$_.TotalSeconds -ge 1 -and $_.TotalMinutes -lt 1} {"[$([int]$timeSpan.TotalSeconds)] seconds"}
        {$_.TotalMinutes -ge 1 -and $_.TotalHours -lt 1} {"[$([int]$timeSpan.TotalMinutes)] minutes [$([int]$timeSpan.Seconds)] seconds"}
        {$_.TotalHours -ge 1 -and $_.TotalDays -lt 1} {"[$([int]$timeSpan.TotalHours)] hours [$([int]$timeSpan.Minutes)] minutes"}
        {$_.TotalDays -ge 1} {"[$([int]$timeSpan.TotalDays)] days [$([int]$timeSpan.Hours)] hours"}
        }
    }
function get-3lettersInBrackets($stringMaybeContaining3LettersInBrackets,$verboseLogging){
    if($stringMaybeContaining3LettersInBrackets -match '\([a-zA-Z]{3}\)'){
        $Matches[0].Replace('(',"").Replace(')',"")
        if($verboseLogging){Write-Host -ForegroundColor DarkCyan "[$($Matches[0])] found in $stringMaybeContainingEngagementCode"}
        }
    else{if($verboseLogging){Write-Host -ForegroundColor DarkCyan "3 letters in brackets not found in $stringMaybeContainingEngagementCode"}}
    }
function get-3letterIsoCodeFromCountryName($pCountryName){
    switch ($pCountryName) {
        {@("UAE","UE","AE","ARE","United Arab Emirates","Dubai") -contains $_} {"ARE"}
        {@("BR","BRA","Brazil","Brasil") -contains $_} {"BRA"}
        {@("CA","CAN","Canada","Canadia") -contains $_} {"CAN"}
        {@("CN","CHN","China") -contains $_} {"CHN"}
        {@("DE","DEU","GE","GER","Germany","Deutschland","Deutchland") -contains $_} {"DEU"}
        {@("ES","ESP","SP","SPA","Spain","España","Espania") -contains $_} {"ESP"}
        {@("FI","FIN","Finland","Suomen","Suomen tasavalta") -contains $_} {"FIN"}
        {@("F","FR",,"FRA","France") -contains $_} {"FRA"}
        {@("UK","GB","GBR","United Kingdom","Great Britain","Scotland","England","Wales","Northern Ireland") -contains $_} {"GBR"}
        {@("IE","IRL","IR","IER","Ireland") -contains $_} {"IRL"}
        {@("PH","PHL","PHI","FIL","Philippenes","Phillippenes","Philipenes","Phillipenes") -contains $_} {"IRL"}
        {@("SE","SWE","SW","SWD","Sweden","Sweeden","Sverige") -contains $_} {"SWE"}
        {@("US","USA","United States","United States of America") -contains $_} {"USA"}
        {@("IT","ITA","Italy","Italia") -contains $_} {"ITA"}
        #Add more countries
        default {}
        }
    }
function get-2letterIsoCodeFrom3LetterIsoCode($p3letterIsoCode){
    switch ($p3letterIsoCode) {
        "ARE" {"AE"}
        "CAN" {"CA"}
        "CHN" {"CN"}
        "DEU" {"DE"}
        "ESP" {"ES"}
        "FIN" {"FI"}
        "GBR" {"GB"}
        "IRL" {"IE"}
        "ITA" {"IT"}
        "PHL" {"PH"}
        "SWE" {"SE"}
        "USA" {"US"}
        #Add more countries
        default {"Unknown"}
        }
    }
function get-2letterIsoCodeFromCountryName($pCountryName){
    $3letterCode = get-3letterIsoCodeFromCountryName -pCountryName $pCountryName
    get-2letterIsoCodeFrom3LetterIsoCode -p3letterIsoCode $3letterCode
    }
function get-available365licensecount{
        [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true,ParameterSetName="LicenseType")]
            [ValidateSet("Office_E1", "Office_E3", "EMS_E3", "All")]
            [string[]]$licensetype
            )
            if(![string]::IsNullOrWhiteSpace($licensetype)){
                    switch ($licensetype){
                        "Office_E1" {
                            $availableLicenses = Get-MsolAccountSku | Where-Object -Property "AccountSkuId" -EQ "AnthesisLLC:STANDARDPACK"
                        }
                        "Office_E3" {
                            $availableLicenses = Get-MsolAccountSku | Where-Object -Property "AccountSkuId" -EQ "AnthesisLLC:ENTERPRISEPACK"
                        }
                        "EMS_E3"{
                            $availableLicenses = Get-MsolAccountSku | Where-Object -Property "AccountSkuId" -EQ "AnthesisLLC:EMS"
                        }
                        "All"{
                            $availableE1Licenses = Get-MsolAccountSku | Where-Object -Property "AccountSkuId" -EQ "AnthesisLLC:$(get-microsoftProductInfo -getType MSStringID -fromType FriendlyName "Office_E1")" #"AnthesisLLC:STANDARDPACK"
                            $availableE3Licenses = Get-MsolAccountSku | Where-Object -Property "AccountSkuId" -EQ "AnthesisLLC:$(get-microsoftProductInfo -getType MSStringID -fromType FriendlyName "Office_E3")" #"AnthesisLLC:ENTERPRISEPACK"
                            $availableE5Licenses = Get-MsolAccountSku | Where-Object -Property "AccountSkuId" -EQ "AnthesisLLC:$(get-microsoftProductInfo -getType MSStringID -fromType FriendlyName "Office_E5")" #"AnthesisLLC:ENTERPRISEPACK"
                            $availableEMSLicenses = Get-MsolAccountSku | Where-Object -Property "AccountSkuId" -EQ "AnthesisLLC:$(get-microsoftProductInfo -getType MSStringID -fromType FriendlyName "EMS_E3")" #"AnthesisLLC:EMS"
                            $availableMDELicenses = Get-MsolAccountSku | Where-Object -Property "AccountSkuId" -EQ "AnthesisLLC:$(get-microsoftProductInfo -getType MSStringID -fromType FriendlyName "MDE")" #"AnthesisLLC:EMS"
                            $availableAudioLicenses = Get-MsolAccountSku | Where-Object -Property "AccountSkuId" -EQ "AnthesisLLC:$(get-microsoftProductInfo -getType MSStringID -fromType FriendlyName "TeamsAudioConferencingSelect")" #"AnthesisLLC:EMS"
                            $availableWinE3Licenses = Get-MsolAccountSku | Where-Object -Property "AccountSkuId" -EQ "AnthesisLLC:$(get-microsoftProductInfo -getType MSStringID -fromType FriendlyName "Win_E3")" #"AnthesisLLC:EMS"
                            $availableME3Licenses = Get-MsolAccountSku | Where-Object -Property "AccountSkuId" -EQ "AnthesisLLC:$(get-microsoftProductInfo -getType MSStringID -fromType FriendlyName "Microsoft_E3")" #"AnthesisLLC:EMS"
                            $availableME5Licenses = Get-MsolAccountSku | Where-Object -Property "AccountSkuId" -EQ "AnthesisLLC:$(get-microsoftProductInfo -getType MSStringID -fromType FriendlyName "Microsoft_E5")" #"AnthesisLLC:EMS"
                            }
                        
                        }
                        If(("Office_E1" -eq $licensetype) -or ("Office_E3" -eq $licensetype) -or ("EMS_E3" -eq $licensetype)){
                            Write-Host "$($licensetype)" "license count:" "$($availableLicenses.ConsumedUnits)"  "/"  "$($availableLicenses.ActiveUnits)" -ForegroundColor Yellow
                        }
                        Else{
                            Write-Host "Available Office_E1 license count:`t`t"$($availableE1Licenses.ConsumedUnits)"`t/`t"$($availableE1Licenses.ActiveUnits)"" -ForegroundColor Yellow
                            Write-Host "Available Office_E3 license count:`t`t"$($availableE3Licenses.ConsumedUnits)"`t/`t"$($availableE3Licenses.ActiveUnits)"" -ForegroundColor Yellow
                            Write-Host "Available Office_E5 license count:`t`t"$($availableE5Licenses.ConsumedUnits)"`t/`t"$($availableE5Licenses.ActiveUnits)"" -ForegroundColor Yellow
                            Write-Host "Available EMS_E3 license count:`t`t`t"$($availableEMSLicenses.ConsumedUnits)"`t/`t"$($availableEMSLicenses.ActiveUnits)"" -ForegroundColor Yellow
                            Write-Host "Available MDE license count:`t`t`t"$($availableMDELicenses.ConsumedUnits)"`t/`t"$($availableMDELicenses.ActiveUnits)"" -ForegroundColor Yellow
                            Write-Host "Available TeamsAudioConf license count:`t"$($availableAudioLicenses.ConsumedUnits)"`t/`t"$($availableAudioLicenses.ActiveUnits)"" -ForegroundColor Yellow
                            Write-Host "Available Win_E3 license count:`t`t`t"$($availableWinE3Licenses.ConsumedUnits)"`t/`t"$($availableWinE3Licenses.ActiveUnits)"" -ForegroundColor Yellow
                            Write-Host "Available Microsoft_E3 license count:`t"$($availableME3Licenses.ConsumedUnits)"`t/`t"$($availableME3Licenses.ActiveUnits)"" -ForegroundColor Yellow
                            Write-Host "Available Microsoft_E5 license count:`t"$($availableME5Licenses.ConsumedUnits)"`t/`t"$($availableME5Licenses.ActiveUnits)"" -ForegroundColor Yellow
                        }
            }
}
function get-azureAdBitlockerHeader{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $false)]
        [pscredential]$aadCreds
        )
    Write-Verbose "get-azureAdBitlockerHeader -aadCreds [$($aadCreds.UserName) | $($aadCreds.Password)]"
    #Test for connection to AzureRM
    Import-Module AzureRM.Profile
    try {    
        $context = Get-AzureRmContext -ErrorAction Stop -WarningAction Stop -InformationAction Stop
        if([string]::IsNullOrWhiteSpace($context)){throw [System.AccessViolationException] "Insuffient privileges to connect to Get-AzureRmContext"}
        }
    catch {
        connect-toAzureRm -aadCreds $aadCreds
        }
    finally {
        if([string]::IsNullOrWhiteSpace($context)){$context = Get-AzureRmContext}
        }

    #Then build header
    $tenantId = $context.Tenant.Id
    $refreshToken = @($context.TokenCache.ReadItems() | Where-Object {$_.tenantId -eq $tenantId -and $_.ExpiresOn -gt (Get-Date)})[0].RefreshToken
    $body = "grant_type=refresh_token&refresh_token=$($refreshToken)&resource=74658136-14ec-4630-ad9b-26e160ff0fc6"
    $apiToken = Invoke-RestMethod "https://login.windows.net/$tenantId/oauth2/token" -Method POST -Body $body -ContentType 'application/x-www-form-urlencoded'
    $header = @{
        'Authorization'          = 'Bearer ' + $apiToken.access_token
        'X-Requested-With'       = 'XMLHttpRequest'
        'x-ms-client-request-id' = [guid]::NewGuid()
        'x-ms-correlation-id'    = [guid]::NewGuid()
        }
    $header
    }
function get-azureAdBitLockerKeysForAllDevices{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $false)]
        [hashtable]$header
        ,[pscredential]$aadCreds
        )

    #Get Header if necessary
    if([string]::IsNullOrWhiteSpace($header)){
        $header = get-azureAdBitlockerHeader -aadCreds $aadCreds 
        }

    #Check if connected to AzureAD
    try{$allDevices = Get-AzureADDevice -All:$true -ErrorAction Stop -WarningAction Stop -InformationAction Stop}
    catch{
        connect-toAAD -credential $aadCreds
        }
    finally{
        if([string]::IsNullOrWhiteSpace($allDevices)){$allDevices = Get-AzureADDevice -All:$true -ErrorAction Stop -WarningAction Stop -InformationAction Stop}
        }

    $bitLockerKeys = @()

    foreach ($device in $allDevices) {
        $bitLockerKeysForThisDevice = get-azureADBitLockerKeysForDevice -adDevice $device -header $header -Verbose
        if(![string]::IsNullOrWhiteSpace($bitLockerKeysForThisDevice)){
            $bitLockerKeys += $bitLockerKeysForThisDevice
            }
        }
    $bitLockerKeys
    }
function get-azureAdBitLockerKeysForDevice{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [Microsoft.Open.AzureAD.Model.DirectoryObject]$adDevice
        ,[hashtable]$header
        )

    $deviceBitLockerKeys = @()
    $url = "https://main.iam.ad.ext.azure.com/api/Device/$($adDevice.objectId)"
    $deviceRecord = Invoke-RestMethod -Uri $url -Headers $header -Method Get
    if ($deviceRecord.bitlockerKey.count -ge 1) {
        $deviceBitLockerKeys += [PSCustomObject]@{
            Device      = $deviceRecord.displayName
            DriveType   = $deviceRecord.bitLockerKey.driveType
            KeyId       = $deviceRecord.bitLockerKey.keyIdentifier
            RecoveryKey = $deviceRecord.bitLockerKey.recoveryKey
            CreationTime= $deviceRecord.bitLockerKey.creationTime
            }
        }
    $deviceBitLockerKeys
    }
function get-azureAdBitLockerKeysForUser {
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [string]$SearchString
        ,[pscredential]$Credential
        )
 
    try{$userDevices = Get-AzureADUser -SearchString $SearchString | Get-AzureADUserRegisteredDevice -All:$true}
    catch{
        connect-toAAD -credential $aadCreds
        }
    finally{
        if([string]::IsNullOrWhiteSpace($userDevices)){$userDevices = Get-AzureADUser -SearchString $SearchString | Get-AzureADUserRegisteredDevice -All:$true}
        }
 
    #Get Header if necessary
    if([string]::IsNullOrWhiteSpace($header)){
        $header = get-azureAdBitlockerHeader -aadCreds $aadCreds
        }

    $bitLockerKeys = @()
    foreach ($device in $userDevices) {
        $bitLockerKeysForThisDevice = get-azureADBitLockerKeysForDevice -adDevice $device -header $header
        if(![string]::IsNullOrWhiteSpace($bitLockerKeysForThisDevice)){
            $bitLockerKeys += $bitLockerKeysForThisDevice
            }
        }

     $bitLockerKeys
    }
function get-dateFormatExamples(){
    [CmdletBinding()]
    param()
    $uFormats = @("d","D","f","F","g","G","m","M","o","O","r","R","s","t","T","u","U","y","Y","FileDateTimeUniversal")
    Write-Host -f Yellow "Without .ToUniversalTime()"
    $date = Get-Date
    $uFormats | % {
        Write-Host "`tGet-Date -f $_ :`t$(Get-Date $date -f $_)"
        }
    Write-Host -f Yellow "With .ToUniversalTime()"
    $date = (Get-Date $date).ToUniversalTime()
    $uFormats | % {
        Write-Host "`tGet-Date -f $_ :`t$(Get-Date $date -f $_)"
        }
    }
function get-dateInIsoFormat(){
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
            [datetime]$dateTime
        ,[parameter(Mandatory = $false)]
            [ValidateSet("Minutes", "Seconds", "Milliseconds", "Ticks")]
            [string]$precision
        )
    $utc = (Get-Date $dateTime).ToUniversalTime()
    switch($precision){
        "Minutes"      {(Get-Date $utc -Format o).Substring(0,16)+"Z"}
        "Seconds"      {(Get-Date $utc -Format o).Substring(0,19)+"Z"}
        "Milliseconds" {(Get-Date $utc -Format o).Substring(0,23)+"Z"}
        "Ticks"        {(Get-Date $utc -Format o)}
        }
    }
function get-errorSummary(){
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
            [System.Management.Automation.ErrorRecord]$errorToSummarise
        )

    $($errorToSummarise | fl * -Force) | Out-String    
    }
function get-groupAdminRoleEmailAddresses_deprecated(){
    [CmdletBinding()]
    param()
    $admins = @()
    Get-MsolRoleMember -RoleObjectId fe930be7-5e62-47db-91af-98c3a49a38b1 | % {$admins += $_.EmailAddress} #User Account Administrator
    Get-MsolRoleMember -RoleObjectId 29232cdf-9323-42fd-ade2-1d097af3e4de | % {$admins += $_.EmailAddress} #Exchange Service Administrator
    $admins | Sort-Object -Unique
    }
function get-keyFromValue($value, $hashTable){
    foreach ($Key in ($hashTable.GetEnumerator() | Where-Object {$_.Value -eq $value})){
        $Key.name}
    }
function get-keyFromValueViaAnotherKey($value, $interimKey, $hashTable){
    foreach ($Key in ($hashTable.GetEnumerator() | Where-Object {$_.Value[$interimKey] -eq $value})){
        $Key.name}
    }
function get-kimbleEngagementCodeFromString($stringMaybeContainingEngagementCode,$verboseLogging){
    if($stringMaybeContainingEngagementCode -match 'E(\d){6}'){
        $Matches[0]
        if($verboseLogging){Write-Host -ForegroundColor DarkCyan "[$($Matches[0])] found in $stringMaybeContainingEngagementCode"}
        }
    else{if($verboseLogging){Write-Host -ForegroundColor DarkCyan "Kimble Project Code not found in $stringMaybeContainingEngagementCode"}}
    }
function get-managersGroupNameFromTeamUrl($teamSiteUrl){
    if(![string]::IsNullOrWhiteSpace($teamSiteUrl)){
        $leaf = Split-Path $teamSiteUrl -Leaf
        $guess = $leaf.Replace("_","")
        if($guess.Substring($guess.Length-3,3) -eq "365"){
            $managerGuess = $guess.Substring(0,$guess.Length-3)+"-Managers"
            }
        else{
            Write-Warning "The URL [$teamSiteUrl] doesn't look like a standardised O365 Group Name - I can't guess this"
            }
        }
    $managerGuess
    }
function get-microsoftProductInfo(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
            [ValidateSet("FriendlyName","MSProductName","MSStringID","GUID","intY","Cost")]
            [string]$getType
        ,[parameter(Mandatory = $true)]
            [ValidateSet("FriendlyName","MSProductName","MSStringID","GUID","intY")]
            [string]$fromType
        ,[parameter(Mandatory = $true)]
            [string]$fromValue
        )
   
    #@(FriendlyName,MSName,MSStringID,GUID)
    switch($getType){
        "FriendlyName" {$getId = 0}
        "MSProductName" {$getId = 1}
        "MSStringID" {$getId = 2}
        "GUID" {$getId = 3}
        "intY" {$getId = 4}
        "Cost" {$getId = 5}
        }
    switch($fromType){
        "FriendlyName" {$fromId = 0}
        "MSProductName" {$fromId = 1}
        "MSStringID" {$fromId = 2}
        "GUID" {$fromId = 3}
        "intY" {$fromId = 4}
        }
    Write-Verbose "getId = [$getId]"
    Write-Verbose "fromId = [$fromId]"
    $productList = @(
        @("TeamsAudioConferencing","AUDIO CONFERENCING","MCOMEETADV","0c266dff-15dd-4b49-8397-2bb16070ed52","Microsoft 365 Audio Conferencing","4"),
        @("AZURE ACTIVE DIRECTORY BASIC","AZURE ACTIVE DIRECTORY BASIC","AAD_BASIC","2b9c8e7c-319c-43a2-a2a0-48c5c6161de7","AZURE ACTIVE DIRECTORY BASIC",""),
        @("AZURE ACTIVE DIRECTORY PREMIUM P1","AZURE ACTIVE DIRECTORY PREMIUM P1","AAD_PREMIUM","078d2b04-f1bd-4111-bbd4-b4b1b354cef4","AZURE ACTIVE DIRECTORY PREMIUM P1",""),
        @("AZURE ACTIVE DIRECTORY PREMIUM P2","AZURE ACTIVE DIRECTORY PREMIUM P2","AAD_PREMIUM_P2","84a661c4-e949-4bd2-a560-ed7766fcaf2b","AZURE ACTIVE DIRECTORY PREMIUM P2",""),
        @("AZURE INFORMATION PROTECTION PLAN 1","AZURE INFORMATION PROTECTION PLAN 1","RIGHTSMANAGEMENT","c52ea49f-fe5d-4e95-93ba-1de91d380f89","AZURE INFORMATION PROTECTION PLAN 1",""),
        @("DYNAMICS 365 CUSTOMER ENGAGEMENT PLAN ENTERPRISE EDITION","DYNAMICS 365 CUSTOMER ENGAGEMENT PLAN ENTERPRISE EDITION","DYN365_ENTERPRISE_PLAN1","ea126fc5-a19e-42e2-a731-da9d437bffcf","DYNAMICS 365 CUSTOMER ENGAGEMENT PLAN ENTERPRISE EDITION",""),
        @("DYNAMICS 365 FOR CUSTOMER SERVICE ENTERPRISE EDITION","DYNAMICS 365 FOR CUSTOMER SERVICE ENTERPRISE EDITION","DYN365_ENTERPRISE_CUSTOMER_SERVICE","749742bf-0d37-4158-a120-33567104deeb","DYNAMICS 365 FOR CUSTOMER SERVICE ENTERPRISE EDITION",""),
        @("DYNAMICS 365 FOR FINANCIALS BUSINESS EDITION","DYNAMICS 365 FOR FINANCIALS BUSINESS EDITION","DYN365_FINANCIALS_BUSINESS_SKU","cc13a803-544e-4464-b4e4-6d6169a138fa","DYNAMICS 365 FOR FINANCIALS BUSINESS EDITION",""),
        @("DYNAMICS 365 FOR SALES AND CUSTOMER SERVICE ENTERPRISE EDITION","DYNAMICS 365 FOR SALES AND CUSTOMER SERVICE ENTERPRISE EDITION","DYN365_ENTERPRISE_SALES_CUSTOMERSERVICE","8edc2cf8-6438-4fa9-b6e3-aa1660c640cc","DYNAMICS 365 FOR SALES AND CUSTOMER SERVICE ENTERPRISE EDITION",""),
        @("DYNAMICS 365 FOR SALES ENTERPRISE EDITION","DYNAMICS 365 FOR SALES ENTERPRISE EDITION","DYN365_ENTERPRISE_SALES","1e1a282c-9c54-43a2-9310-98ef728faace","DYNAMICS 365 FOR SALES ENTERPRISE EDITION",""),
        @("DYNAMICS 365 FOR TEAM MEMBERS ENTERPRISE EDITION","DYNAMICS 365 FOR TEAM MEMBERS ENTERPRISE EDITION","DYN365_ENTERPRISE_TEAM_MEMBERS","8e7a3d30-d97d-43ab-837c-d7701cef83dc","DYNAMICS 365 FOR TEAM MEMBERS ENTERPRISE EDITION",""),
        @("DYNAMICS 365 UNF OPS PLAN ENT EDITION","DYNAMICS 365 UNF OPS PLAN ENT EDITION","Dynamics_365_for_Operations","ccba3cfe-71ef-423a-bd87-b6df3dce59a9","DYNAMICS 365 UNF OPS PLAN ENT EDITION",""),
        @("EMS_E3","ENTERPRISE MOBILITY + SECURITY E3","EMS","efccb6f7-5641-4e0e-bd10-b4976e1bf68e","Enterprise Mobility + Security (E3) (CSP)","8.36"),
        @("EMS_E5","ENTERPRISE MOBILITY + SECURITY E5","EMSPREMIUM","b05e124f-c7cc-45a0-a6aa-8cf78c946968","ENTERPRISE MOBILITY + SECURITY E5",""),
        @("EXCHANGE ONLINE (PLAN 1)","EXCHANGE ONLINE (PLAN 1)","EXCHANGESTANDARD","4b9405b0-7788-4568-add1-99614e613b69","EXCHANGE ONLINE (PLAN 1)",""),
        @("EXCHANGE ONLINE (PLAN 2)","EXCHANGE ONLINE (PLAN 2)","EXCHANGEENTERPRISE","19ec0d23-8335-4cbd-94ac-6050e30712fa","EXCHANGE ONLINE (PLAN 2)",""),
        @("EXCHANGE ONLINE ARCHIVING FOR EXCHANGE ONLINE","EXCHANGE ONLINE ARCHIVING FOR EXCHANGE ONLINE","EXCHANGEARCHIVE_ADDON","ee02fd1b-340e-4a4b-b355-4a514e4c8943","EXCHANGE ONLINE ARCHIVING FOR EXCHANGE ONLINE",""),
        @("EXCHANGE ONLINE ARCHIVING FOR EXCHANGE SERVER","EXCHANGE ONLINE ARCHIVING FOR EXCHANGE SERVER","EXCHANGEARCHIVE","90b5e015-709a-4b8b-b08e-3200f994494c","EXCHANGE ONLINE ARCHIVING FOR EXCHANGE SERVER",""),
        @("EXCHANGE ONLINE ESSENTIALS","EXCHANGE ONLINE ESSENTIALS","EXCHANGEESSENTIALS","7fc0182e-d107-4556-8329-7caaa511197b","EXCHANGE ONLINE ESSENTIALS",""),
        @("EXCHANGE ONLINE ESSENTIALS","EXCHANGE ONLINE ESSENTIALS","EXCHANGE_S_ESSENTIALS","e8f81a67-bd96-4074-b108-cf193eb9433b","EXCHANGE ONLINE ESSENTIALS",""),
        @("Kiosk_K1","EXCHANGE ONLINE KIOSK","EXCHANGEDESKLESS","80b2d799-d2ba-4d2a-8842-fb0d0f3a4b82","Exchange Online Kiosk (CSP)","1.9"),
        @("EXCHANGE ONLINE POP","EXCHANGE ONLINE POP","EXCHANGETELCO","cb0a98a8-11bc-494c-83d9-c1b1ac65327e","EXCHANGE ONLINE POP",""),
        @("INTUNE","INTUNE","INTUNE_A","061f9ace-7d42-4136-88ac-31dc755f143f","INTUNE",""),
        @("Microsoft 365 A1","Microsoft 365 A1","M365EDU_A1","b17653a4-2443-4e8c-a550-18249dda78bb","Microsoft 365 A1",""),
        @("Microsoft 365 A3 for faculty","Microsoft 365 A3 for faculty","M365EDU_A3_FACULTY","4b590615-0888-425a-a965-b3bf7789848d","Microsoft 365 A3 for faculty",""),
        @("Microsoft 365 A3 for students","Microsoft 365 A3 for students","M365EDU_A3_STUDENT","7cfd9a2b-e110-4c39-bf20-c6a3f36a3121","Microsoft 365 A3 for students",""),
        @("Microsoft 365 A5 for faculty","Microsoft 365 A5 for faculty","M365EDU_A5_FACULTY","e97c048c-37a4-45fb-ab50-922fbf07a370","Microsoft 365 A5 for faculty",""),
        @("Microsoft 365 A5 for students","Microsoft 365 A5 for students","M365EDU_A5_STUDENT","46c119d4-0379-4a9d-85e4-97c66d3f909e","Microsoft 365 A5 for students",""),
        @("MICROSOFT 365 BUSINESS","MICROSOFT 365 BUSINESS","SPB","cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46","MICROSOFT 365 BUSINESS",""),
        @("Microsoft_E3","MICROSOFT 365 E3","SPE_E3","05e9a617-0261-4cee-bb44-138d3ef5d965","Microsoft 365 Enterprise E3","32"),
        @("Microsoft_E5","Microsoft 365 E5","SPE_E5","06ebc4ee-1bb5-47dd-8120-11324bc54e06","Microsoft 365 E5",""),
        @("Microsoft 365 E3_USGOV_DOD","Microsoft 365 E3_USGOV_DOD","SPE_E3_USGOV_DOD","d61d61cc-f992-433f-a577-5bd016037eeb","Microsoft 365 E3_USGOV_DOD",""),
        @("Microsoft 365 E3_USGOV_GCCHIGH","Microsoft 365 E3_USGOV_GCCHIGH","SPE_E3_USGOV_GCCHIGH","ca9d1dd9-dfe9-4fef-b97c-9bc1ea3c3658","Microsoft 365 E3_USGOV_GCCHIGH",""),
        @("Microsoft 365 E5 Compliance","Microsoft 365 E5 Compliance","INFORMATION_PROTECTION_COMPLIANCE","184efa21-98c3-4e5d-95ab-d07053a96e67","Microsoft 365 E5 Compliance",""),
        @("Microsoft 365 E5 Security","Microsoft 365 E5 Security","IDENTITY_THREAT_PROTECTION","26124093-3d78-432b-b5dc-48bf992543d5","Microsoft 365 E5 Security",""),
        @("Microsoft 365 E5 Security for EMS E5","Microsoft 365 E5 Security for EMS E5","IDENTITY_THREAT_PROTECTION_FOR_EMS_E5","44ac31e7-2999-4304-ad94-c948886741d4","Microsoft 365 E5 Security for EMS E5",""),
        @("Microsoft 365 F1","Microsoft 365 F1","SPE_F1","66b55226-6b4f-492c-910c-a3b7a3c9d993","Microsoft 365 F1",""),
        @("MDE","Microsoft Defender Advanced Threat Protection","MDATP_XPLAT","b126b073-72db-4a9d-87a4-b17afe41d4ab","Microsoft Defender Advanced Threat Protection","5.2"),
        @("MICROSOFT DYNAMICS CRM ONLINE BASIC","MICROSOFT DYNAMICS CRM ONLINE BASIC","CRMPLAN2","906af65a-2970-46d5-9b58-4e9aa50f0657","MICROSOFT DYNAMICS CRM ONLINE BASIC",""),
        @("MICROSOFT DYNAMICS CRM ONLINE","MICROSOFT DYNAMICS CRM ONLINE","CRMSTANDARD","d17b27af-3f49-4822-99f9-56a661538792","MICROSOFT DYNAMICS CRM ONLINE",""),
        @("MS IMAGINE ACADEMY","MS IMAGINE ACADEMY","IT_ACADEMY_AD","ba9a34de-4489-469d-879c-0f0f145321cd","MS IMAGINE ACADEMY",""),
        @("Office 365 A5 for faculty","Office 365 A5 for faculty","ENTERPRISEPREMIUM_FACULTY","a4585165-0533-458a-97e3-c400570268c4","Office 365 A5 for faculty",""),
        @("Office 365 A5 for students","Office 365 A5 for students","ENTERPRISEPREMIUM_STUDENT","ee656612-49fa-43e5-b67e-cb1fdf7699df","Office 365 A5 for students",""),
        @("Office 365 Advanced Compliance","Office 365 Advanced Compliance","EQUIVIO_ANALYTICS","1b1b1f7a-8355-43b6-829f-336cfccb744c","Office 365 Advanced Compliance",""),
        @("AdvancedSpam","Office 365 Advanced Threat Protection (Plan 1)","ATP_ENTERPRISE","4ef96642-f096-40de-a3e9-d83fb2f90211","Office 365 Advanced Threat Protection (Plan 1)","1.9"),
        @("OFFICE 365 BUSINESS","OFFICE 365 BUSINESS","O365_BUSINESS","cdd28e44-67e3-425e-be4c-737fab2899d3","OFFICE 365 BUSINESS",""),
        @("OFFICE 365 BUSINESS","OFFICE 365 BUSINESS","SMB_BUSINESS","b214fe43-f5a3-4703-beeb-fa97188220fc","OFFICE 365 BUSINESS",""),
        @("OFFICE 365 BUSINESS ESSENTIALS","OFFICE 365 BUSINESS ESSENTIALS","O365_BUSINESS_ESSENTIALS","3b555118-da6a-4418-894f-7df1e2096870","OFFICE 365 BUSINESS ESSENTIALS",""),
        @("OFFICE 365 BUSINESS ESSENTIALS","OFFICE 365 BUSINESS ESSENTIALS","SMB_BUSINESS_ESSENTIALS","dab7782a-93b1-4074-8bb1-0e61318bea0b","OFFICE 365 BUSINESS ESSENTIALS",""),
        @("OFFICE 365 BUSINESS PREMIUM","OFFICE 365 BUSINESS PREMIUM","O365_BUSINESS_PREMIUM","f245ecc8-75af-4f8e-b61f-27d8114de5f3","OFFICE 365 BUSINESS PREMIUM",""),
        @("OFFICE 365 BUSINESS PREMIUM","OFFICE 365 BUSINESS PREMIUM","SMB_BUSINESS_PREMIUM","ac5cef5d-921b-4f97-9ef3-c99076e5470f","OFFICE 365 BUSINESS PREMIUM",""),
        @("Office_E1","OFFICE 365 E1","STANDARDPACK","18181a46-0d4e-45cd-891e-60aabd171b4e","Office 365 Enterprise E1 (CSP)","7.6"),
        @("OFFICE 365 E2","OFFICE 365 E2","STANDARDWOFFPACK","6634e0ce-1a9f-428c-a498-f84ec7b8aa2e","OFFICE 365 E2",""),
        @("Office_E3","OFFICE 365 E3","ENTERPRISEPACK","6fd2c87f-b296-42f0-b197-1e91e994b900","Office 365 Enterprise E3 (CSP)","19"),
        @("OFFICE 365 E3 DEVELOPER","OFFICE 365 E3 DEVELOPER","DEVELOPERPACK","189a915c-fe4f-4ffa-bde4-85b9628d07a0","OFFICE 365 E3 DEVELOPER",""),
        @("Office 365 E3_USGOV_DOD","Office 365 E3_USGOV_DOD","ENTERPRISEPACK_USGOV_DOD","b107e5a3-3e60-4c0d-a184-a7e4395eb44c","Office 365 E3_USGOV_DOD",""),
        @("Office 365 E3_USGOV_GCCHIGH","Office 365 E3_USGOV_GCCHIGH","ENTERPRISEPACK_USGOV_GCCHIGH","aea38a85-9bd5-4981-aa00-616b411205bf","Office 365 E3_USGOV_GCCHIGH",""),
        @("OFFICE 365 E4","OFFICE 365 E4","ENTERPRISEWITHSCAL","1392051d-0cb9-4b7a-88d5-621fee5e8711","OFFICE 365 E4",""),
        @("Office_E5","OFFICE 365 E5","ENTERPRISEPREMIUM","c7df2760-2c81-4ef7-b578-5b5392b571df","Office 365 Enterprise E5","35"),
        @("OFFICE 365 E5 WITHOUT AUDIO CONFERENCING","OFFICE 365 E5 WITHOUT AUDIO CONFERENCING","ENTERPRISEPREMIUM_NOPSTNCONF","26d45bd9-adf1-46cd-a9e1-51e9a5524128","OFFICE 365 E5 WITHOUT AUDIO CONFERENCING",""),
        @("OFFICE 365 F1","OFFICE 365 F1","DESKLESSPACK","4b585984-651b-448a-9e53-3b10f069cf7f","OFFICE 365 F1",""),
        @("OFFICE 365 MIDSIZE BUSINESS","OFFICE 365 MIDSIZE BUSINESS","MIDSIZEPACK","04a7fb0d-32e0-4241-b4f5-3f7618cd1162","OFFICE 365 MIDSIZE BUSINESS",""),
        @("OFFICE 365 PROPLUS","OFFICE 365 PROPLUS","OFFICESUBSCRIPTION","c2273bd0-dff7-4215-9ef5-2c7bcfb06425","OFFICE 365 PROPLUS",""),
        @("OFFICE 365 SMALL BUSINESS","OFFICE 365 SMALL BUSINESS","LITEPACK","bd09678e-b83c-4d3f-aaba-3dad4abd128b","OFFICE 365 SMALL BUSINESS",""),
        @("OFFICE 365 SMALL BUSINESS PREMIUM","OFFICE 365 SMALL BUSINESS PREMIUM","LITEPACK_P2","fc14ec4a-4169-49a4-a51e-2c852931814b","OFFICE 365 SMALL BUSINESS PREMIUM",""),
        @("OneDrive","ONEDRIVE FOR BUSINESS (PLAN 1)","WACONEDRIVESTANDARD","e6778190-713e-4e4f-9119-8b8238de25df","ONEDRIVE FOR BUSINESS (PLAN 1)",""),
        @("ONEDRIVE FOR BUSINESS (PLAN 2)","ONEDRIVE FOR BUSINESS (PLAN 2)","WACONEDRIVEENTERPRISE","ed01faf2-1d88-4947-ae91-45ca18703a96","ONEDRIVE FOR BUSINESS (PLAN 2)",""),
        @("POWER APPS PER USER PLAN","POWER APPS PER USER PLAN","POWERAPPS_PER_USER","b30411f5-fea1-4a59-9ad9-3db7c7ead579","POWER APPS PER USER PLAN",""),
        @("POWER BI FOR OFFICE 365 ADD_ON","POWER BI FOR OFFICE 365 ADD-ON","POWER_BI_ADDON","45bc2c81-6072-436a-9b0b-3b12eefbc402","POWER BI FOR OFFICE 365 ADD-ON",""),
        @("PowerBIFree","POWER BI FREE","POWER_BI_FREE","a403ebcc-fae0-4ca2-8c8c-7a907fd6c235","Zero cost Power BI licence",""),
        @("PowerBI_Pro","POWER BI PRO","POWER_BI_PRO","f8a1db68-be16-40ed-86d5-cb42ce701560","Power BI Pro (CSP)","10"),
        @("PROJECT FOR OFFICE 365","PROJECT FOR OFFICE 365","PROJECTCLIENT","a10d5e58-74da-4312-95c8-76be4e5b75a0","PROJECT FOR OFFICE 365",""),
        @("PROJECT ONLINE ESSENTIALS","PROJECT ONLINE ESSENTIALS","PROJECTESSENTIALS","776df282-9fc0-4862-99e2-70e561b9909e","PROJECT ONLINE ESSENTIALS",""),
        @("PROJECT ONLINE PREMIUM","PROJECT ONLINE PREMIUM","PROJECTPREMIUM","09015f9f-377f-4538-bbb5-f75ceb09358a","PROJECT ONLINE PREMIUM",""),
        @("PROJECT ONLINE PREMIUM WITHOUT PROJECT CLIENT","PROJECT ONLINE PREMIUM WITHOUT PROJECT CLIENT","PROJECTONLINE_PLAN_1","2db84718-652c-47a7-860c-f10d8abbdae3","PROJECT ONLINE PREMIUM WITHOUT PROJECT CLIENT",""),
        @("Project Online","PROJECT PLAN 1","PROJECT_P1","beb6439c-caad-48d3-bf46-0c82871e12be","Microsoft Project Plan 1","28.5"),
        @("Project","PROJECT ONLINE PROFESSIONAL","PROJECTPROFESSIONAL","53818b1b-4a27-454b-8896-0dba576410e6","Microsoft Project Plan 3","28.5"),
        @("PROJECT ONLINE WITH PROJECT FOR OFFICE 365","PROJECT ONLINE WITH PROJECT FOR OFFICE 365","PROJECTONLINE_PLAN_2","f82a60b8-1ee3-4cfb-a4fe-1c6a53c2656c","PROJECT ONLINE WITH PROJECT FOR OFFICE 365",""),
        @("SHAREPOINT ONLINE (PLAN 1)","SHAREPOINT ONLINE (PLAN 1)","SHAREPOINTSTANDARD","1fc08a02-8b3d-43b9-831e-f76859e04e1a","SHAREPOINT ONLINE (PLAN 1)",""),
        @("SHAREPOINT ONLINE (PLAN 2)","SHAREPOINT ONLINE (PLAN 2)","SHAREPOINTENTERPRISE","a9732ec9-17d9-494c-a51c-d6b45b384dcb","SHAREPOINT ONLINE (PLAN 2)",""),
        @("SKYPE FOR BUSINESS CLOUD PBX","SKYPE FOR BUSINESS CLOUD PBX","MCOEV","e43b5b99-8dfb-405f-9987-dc307f34bcbd","SKYPE FOR BUSINESS CLOUD PBX",""),
        @("SKYPE FOR BUSINESS ONLINE (PLAN 1)","SKYPE FOR BUSINESS ONLINE (PLAN 1)","MCOIMP","b8b749f8-a4ef-4887-9539-c95b1eaa5db7","SKYPE FOR BUSINESS ONLINE (PLAN 1)",""),
        @("SKYPE FOR BUSINESS ONLINE (PLAN 2)","SKYPE FOR BUSINESS ONLINE (PLAN 2)","MCOSTANDARD","d42c793f-6c78-4f43-92ca-e8f6a02b035f","SKYPE FOR BUSINESS ONLINE (PLAN 2)",""),
        @("TeamsCallingPlan_International","SKYPE FOR BUSINESS PSTN DOMESTIC AND INTERNATIONAL CALLING","MCOPSTN2","d3b4fe1f-9992-4930-8acb-ca6ec609365e","Skype for Business PSTN Domestic/Local and International Calling","24"),
        @("TeamsCallingPlan_Domestic","SKYPE FOR BUSINESS PSTN DOMESTIC CALLING","MCOPSTN1","0dab259f-bf13-4952-b7f8-7db8f131b28d","Skype for Business PSTN Domestic/Local Calling","12"),
        @("InternationalCalling","SKYPE FOR BUSINESS PSTN DOMESTIC CALLING (120 Minutes)","MCOPSTN5","54a152dc-90de-4996-93d2-bc47e670fc06","SKYPE FOR BUSINESS PSTN DOMESTIC CALLING (120 Minutes)",""),
        @("VISIO ONLINE PLAN 1","VISIO ONLINE PLAN 1","VISIOONLINE_PLAN1","4b244418-9658-4451-a2b8-b5e2b364e9bd","VISIO ONLINE PLAN 1",""),
        @("Visio","VISIO Online Plan 2","VISIOCLIENT","c5928f49-12ba-48f7-ada3-0d743a3601d5","Visio Pro for Office 365 (CSP)","14.25"),
        @("Win_E3","WINDOWS 10 ENTERPRISE E3","Win10_VDA_E3","6a0f6da5-0b87-4190-a6ae-9bb5a2b9546a","WINDOWS 10 ENTERPRISE E3","7"),
        @("Win_E5","Windows 10 Enterprise E5","WIN10_VDA_E5","488ba24a-39a9-4473-8ee5-19291e71b002","Windows 10 Enterprise E5",""),
        @("PowerAutomateFree","FLOW_FREE","FLOW_FREE","f30db892-07e9-47e9-837c-80727f46fd3d","Zero cost Power Automate (Flow) licence",""),
        @("Microsoft Stream Trial","STREAM","STREAM_TRIAL","1f2f344a-700d-42c9-9427-5cea1d5d7ba6","STREAM",""),
        @("TeamsAudioConferencingSelect","Microsoft Teams Audio Conferencing with dial-out to USA/CAN","Microsoft_Teams_Audio_Conferencing_select_dial_out","1c27243e-fb4d-42b1-ae8c-fe25c9616588","STREAM","0")
        )
    $foundProduct = $productList | ? {$_[$fromId] -eq $fromValue} 
    $foundProduct[$getId]
    }
function Get-Shortcut {
    param(
        $path = $null
        )

    $obj = New-Object -ComObject WScript.Shell

    if ($path -eq $null) {
        $pathUser = [System.Environment]::GetFolderPath('StartMenu')
        $pathCommon = $obj.SpecialFolders.Item('AllUsersStartMenu')
        $path = dir $pathUser, $pathCommon -Filter *.lnk -Recurse 
        }
    if ($path -is [string]) {
        $path = dir $path -Filter *.lnk
        }
    $path | ForEach-Object { 
        if ($_ -is [string]) {
            $_ = dir $_ -Filter *.lnk
            }
        if ($_) {
            $link = $obj.CreateShortcut($_.FullName)

            $info = @{}
            $info.Hotkey = $link.Hotkey
            $info.TargetPath = $link.TargetPath
            $info.LinkPath = $link.FullName
            $info.Arguments = $link.Arguments
            $info.Target = try {Split-Path $info.TargetPath -Leaf } catch { 'n/a'}
            $info.Link = try { Split-Path $info.LinkPath -Leaf } catch { 'n/a'}
            $info.WindowStyle = $link.WindowStyle
            $info.IconLocation = $link.IconLocation

            New-Object PSObject -Property $info
            }
        }
    }
function get-timeZones(){
    $timeZones = Get-ChildItem "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Time zones" | foreach {Get-ItemProperty $_.PSPath}; $TimeZone | Out-Null
    $timeZones
    }
function get-timeZoneHashTable($timeZoneArray){
    if($timeZoneArray.Count -lt 1){$timeZones = get-timeZones}
        else {$timeZones = $timeZoneArray}
    $timeZoneHashTable = @{}
    $timeZones | % {$timeZoneHashTable.Add($_.PSChildName, ($_.Display.Split(" ")[0].Replace("(","").Replace(")","")))} | Out-Null
    $timeZoneHashTable.Add("","Unknown") | Out-Null
    $timeZoneHashTable
    }
function get-timeZoneSpsIdFromUnformattedTimeZone($pUnformattedTimeZone, $pTimeZoneHashTable, $pSpoTimeZoneHashTable){
    if ($pTimeZoneHashTable.Count -eq 0){$timeZoneHashTable = get-timeZoneHashTable}
        else{$timeZoneHashTable = $pTimeZoneHashTable}
    if ($pSpoTimeZoneHashTable.Count -eq 0){
        

        $spoTimeZoneHashTable = get-timeZoneHashTable
        }
        else{$spoTimeZoneHashTable = $pSpoTimeZoneHashTable}

    }
function get-trailing3LettersIfTheyLookLikeAnIsoCountryCode(){
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ambiguousString
        )
    if($ambiguousString -match ", [a-zA-Z]{3}$"){
        $ambiguousString.Substring($ambiguousString.Length-3,3)
        }
    }
function get-unformattedTimeZone ($pFormattedTimeZone){
    if ($pFormattedTimeZone -eq "" -or $pFormattedTimeZone -eq $null){"Unknown"}
    else{
        #$pFormattedTimeZone.Split("(")[1].Replace(")","").Trim()
        [regex]::Match($pFormattedTimeZone,"\(([^)]+)\)").Groups[1].Value #Get everything between "(" and ")"
        }
    }
function grant-ownership {
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
		    [string]$fullPath
        ,[parameter(Mandatory = $true)]
            [ValidateSet(“File”,”Folder”)] 
		    [string]$itemType
        ,[parameter(Mandatory = $true)]
            [string[]]$securityPrincipalsToGrantOwnershipTo
        ,[parameter(Mandatory = $false)]
            [switch]$alsoGrantSystemAccountOwnership
        ,[parameter(Mandatory = $false)]
            [switch]$recursive
        ,[parameter(Mandatory = $false)]
            [switch]$seizeFilesIndividually
        )
    Write-Verbose "Seizing ownership of [$itemType] [$fullPath]"
    switch($itemType){
        “File”   {$fullControlPermissions = "FullControl","Allow"}
        ”Folder” {$fullControlPermissions = "FullControl","ContainerInherit,ObjectInherit","None","Allow"}
        default  {}
        }
            
	& "$env:SystemRoot\system32\takeown.exe" /A /F $fullPath | Out-Null
    
    if($alsoGrantSystemAccountOwnership){
        $securityPrincipalsToGrantOwnership.Add("NT AUTHORITY\SYSTEM")
        }

	$currentAcl = Get-Acl $fullPath
    $securityPrincipalsToGrantOwnershipTo | % {
        $thisSecurityPrincipal = $_
        Write-Verbose "Granting FullControl to [$($thisSecurityPrincipal)] on [$itemType] [$($fullPath)]"
        $aclPermission = @($thisSecurityPrincipal)
        $fullControlPermissions | % {$aclPermission += $_}
        $aclAccessRule = new-object System.Security.AccessControl.FileSystemAccessRule $aclPermission
        $currentAcl.AddAccessRule($aclAccessRule)
        } 
    Set-Acl -Path $fullPath -AclObject $currentAcl 
    if($recursive){
        Get-ChildItem -Path $fullPath | % {
            if($_.Mode -match "d"){grant-ownership -fullPath $_.FullName -itemType Folder -securityPrincipalsToGrantOwnershipTo $securityPrincipalsToGrantOwnershipTo -recursive -seizeFilesIndividually:$seizeFilesIndividually -Verbose:$VerbosePreference}
            elseif($siezeFilesIndividually){grant-ownership -fullPath $_.FullName -itemType File -securityPrincipalsToGrantOwnershipTo $securityPrincipalsToGrantOwnershipTo -seizeFilesIndividually:$seizeFilesIndividually -Verbose:$VerbosePreference}
            }
        }
    }
function guess-languageCodeFromCountry($p3LetterCountryIsoCode){
    switch ($p3LetterCountryIsoCode){
        "ARE" {"en-GB"}
        "CAN" {"en-CA"}
        "CHN" {"en-US"}
        "DEU" {"de"}
        "ESP" {"es"}
        "FIN" {"fi"}
        "GBR" {"en-GB"}
        "IRL" {"en-GB"}
        "PHL" {"en-US"}
        "SWE" {"sv"}
        "USA" {"en-US"}
        }
    }
function guess-nameFromString([string]$ambiguousString){
    $lessAmbiguousString = $ambiguousString.Trim().Replace('"','')
    $leastAmbiguousString = $null
    #If it doesn't contain a space, see if it's an e-mail address
    if($lessAmbiguousString.Split(" ").Count -lt 2){
        if($lessAmbiguousString -match "@"){
            $lessAmbiguousString.Split("@")[0] | % {$_.Split(".")} | %{
                $blob = $_.Trim()
                $leastAmbiguousString += $($blob.SubString(0,1).ToUpper() + $blob.SubString(1,$blob.Length-1).ToLower()) + " " #Title Case
                }
            }
        else{$leastAmbiguousString = $lessAmbiguousString}#Do nothing - it's too weird.
        }
    else{
        if($lessAmbiguousString -match ","){#If Lastname, Firstname
            $lessAmbiguousString.Split(",") | %{
                $blob = $_.Trim()
                $leastAmbiguousString = $($blob.SubString(0,1).ToUpper() + $blob.SubString(1,$blob.Length-1).ToLower()) + " $leastAmbiguousString" #Prepend each blob as they're in reverse order
                }
            }
        else{
            $lessAmbiguousString.Split(" ") | %{ #If firstname lastname
                $blob = $_.Trim()
                $leastAmbiguousString += $($blob.SubString(0,1).ToUpper() + $blob.SubString(1,$blob.Length-1).ToLower()) + " "#Just Title Case it
                }
            }
        }
    $leastAmbiguousString.Trim()
    }
function import-encryptedCsv(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
        [string]$pathToEncryptedCsv        
        )

    [array]$encryptedCsvData = import-csv $pathToEncryptedCsv
    [array]$decryptedCsvData = @($null)*$encryptedCsvData.Count
    for ($i=0; $i -lt $encryptedCsvData.Count; $i++){
        $decryptedObject = New-Object psobject
        $encryptedCsvData[$i].PSObject.Properties | ForEach-Object {
            if([string]::IsNullOrWhiteSpace($_.Value)){$decryptedObject | Add-Member -MemberType NoteProperty -Name $_.Name -Value $null}
            else{$decryptedObject | Add-Member -MemberType NoteProperty -Name $_.Name -Value $(decrypt-SecureString -secureString $(ConvertTo-SecureString $_.Value))}
            }
        $decryptedCsvData[$i] = $decryptedObject
        }
    $decryptedCsvData
    }
function log-action($myMessage, $logFile, $doNotLogToFile, $doNotLogToScreen){
    if(!$doNotLogToFile -or $logToFile){Add-Content -Value ((Get-Date -Format "yyyy-MM-dd HH:mm:ss")+"`tACTION:`t$myMessage") -Path $logFile}
    if(!$doNotLogToScreen -or $logToScreen){Write-Host -ForegroundColor Yellow $myMessage}
    }
function log-error($myError, $myFriendlyMessage, $fullLogFile, $errorLogFile, $doNotLogToFile, $doNotLogToScreen, $doNotLogToEmail, $smtpServer, $mailTo, $mailFrom){
    if(!$doNotLogToFile -or $logToFile){
        Add-Content -Value "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")`t`tERROR:`t$myFriendlyMessage" -Path $errorLogFile
        Add-Content -Value "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")`t$($myError.Exception.Message)" -Path $errorLogFile
        if($fullLogFile){
            Add-Content -Value "`t`tERROR:`t$myFriendlyMessage" -Path $fullLogFile
            Add-Content -Value "`t`t$($myError.Exception.Message)" -Path $fullLogFile
            }
        }
    if(!$doNotLogToScreen -or $logToScreen){
        Write-Host -ForegroundColor Red $myFriendlyMessage
        Write-Host -ForegroundColor Red $myError
        }
    if(!$doNotLogToEmail -or $logErrorsToEmail){
        if([string]::IsNullOrWhiteSpace($to)){$to = $env:USERNAME+"@anthesisgroup.com"}
        if([string]::IsNullOrWhiteSpace($mailFrom)){$mailFrom = $env:COMPUTERNAME+"@anthesisgroup.com"}
        if([string]::IsNullOrWhiteSpace($smtpServer)){$smtpServer= "anthesisgroup-com.mail.protection.outlook.com"}
        Send-MailMessage -To $mailTo -From $mailFrom -Subject "Error in automated script - $($myFriendlyMessage.SubString(0,20))" -Body ("$myError`r`n`r`n$myFriendlyMessage") -SmtpServer $smtpServer
        }
    }
function log-result($myMessage, $logFile, $doNotLogToFile, $doNotLogToScreen){
    if(!$doNotLogToFile -or $logToFile){Add-Content -Value ("`tRESULT:`t$myMessage") -Path $logfile}
    if(!$doNotLogToScreen -or $logToScreen){Write-Host -ForegroundColor DarkYellow "`t$myMessage"}
    }
function matchContains($term, $arrayOfStrings){
    # Turn wildcards into regexes
    # First escape all characters that might cause trouble in regexes (leaving out those we care about)
    $escaped = $arrayOfStrings -replace '[ #$()+.[\\^{]','\$&' # list taken from Regex.Escape
    # replace wildcards with their regex equivalents
    $regexes = $escaped -replace '\*','.*' -replace '\?','.'
    # combine them into one regex
    $singleRegex = ($regexes | %{ '^' + $_ + '$' }) -join '|'

    # match against that regex
    $term -match $singleRegex
    }
function new-clientAssertionWithCertificate{
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate]$X509cert
        ,[parameter(Mandatory = $true)]
        [string]$clientId
        ,[parameter(Mandatory = $true)]
        [string]$tenantId
        ,[parameter(Mandatory = $false)]
        [string]$resource='https://graph.microsoft.com'
        ,[parameter(Mandatory = $false)]
        [string]$loginEndpoint='https://login.microsoftonline.com'
        )

    <#
    .SYNOPSIS
    Generates a signed client_assertion using the provided certificate, ready to present to Graph API for authentication.
    **REQUIRES PowerShell 5**
    
    .DESCRIPTION
    #Adapted from https://gist.github.com/jformacek/aecc4f379b88b3a330ee19b045252462
    There are alternative methods for authenticating with Graph using certificates (see get-graphTokenResponse), but this is in-keping with our 5.1 codebase.
    
    .PARAMETER X509cert
    A certificate with a Private Key (to sign the assertion)

    .PARAMETER clientId
    The ClientId/ApplicationId of the App Registration to authenticate with

    .PARAMETER tenantId
    The Id of Azure tenant to authenticate with

    .PARAMETER resource
    Future compatibility to support alternative resources (default is [https://graph.microsoft.com])

    .PARAMETER loginEndpoint
    Future compatibility to support alternative auth endpoints (default is [https://login.microsoftonline.com])
    #>
        
    
    #load required assembly that is not loaded by default by PowerShell
    [System.Reflection.Assembly]::LoadWithPartialName('system.identitymodel') | Out-Null

    #retrieve certificate from cert store
    #$cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object {$_.Thumbprint -eq $certThumbprint}[0]

    # Create base64 hash of the certificate
    $certHash = [System.Convert]::ToBase64String($X509cert.GetCertHash())

    #JWT expiration timestamp - valid for 5 minutes - just to allow token request to complete
    $StartDate = (Get-Date "1970-01-01T00:00:00Z" ).ToUniversalTime()
    $JWTExpiration = [int]((New-TimeSpan -Start $StartDate -End (Get-Date).ToUniversalTime().AddMinutes(5)).TotalSeconds)

    #JWT Start timestamp - optional
    $NotBefore = [int]((New-TimeSpan -Start $StartDate -End ((Get-Date).ToUniversalTime())).TotalSeconds)

    # Create JWT header
    $JWTHeader = @{
        alg = "RS256"
        typ = "JWT"
        x5t = ($certHash -replace '\+','-' -replace '/','_' -replace '=')
        }

    #create request payload - notice that we do not include nbf (but we could, if needed)
    $JWTPayLoad = @{
        aud = "$loginEndpoint/$tenantId/oauth2/token"
        exp = $JWTExpiration
        iss = $clientId
        jti = [guid]::NewGuid()
        #nbf = $NotBefore
        sub = $clientId
        }

    # Convert header and payload to base64 and create JWT assertion
    $EncodedHeader = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(($JWTHeader | ConvertTo-Json)))
    $EncodedPayload = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(($JWTPayload | ConvertTo-Json)))
    $JWT = $EncodedHeader + "." + $EncodedPayload

    #now sign the assertion
    $dataToSign = [byte[]] [System.Text.Encoding]::UTF8.GetBytes($JWT)
    #original RSA CSP from cert is not guaranteed to support SHA256
    #so we need to create new CSP with proper params and with private key from cert for signing
    $algo = (new-object System.IdentityModel.Tokens.X509AsymmetricSecurityKey($X509cert)).GetAsymmetricAlgorithm("http://www.w3.org/2001/04/xmldsig-more#rsa-sha256", $true) -as [System.Security.Cryptography.RSA]

    if($algo -is [System.Security.Cryptography.RSACryptoServiceProvider]){
        #cert uses CryptoAPI CSP
        if(($algo.CspKeyContainerInfo.ProviderType -ne 1) -and ($algo.CspKeyContainerInfo.ProviderType -ne 12) -or $algo.CspKeyContainerInfo.HardwareDevice){
            #we have SHA256 compatible provider, just use it
            $csp = $algo -as [System.Security.Cryptography.RSACryptoServiceProvider]
            }
        else {
            #we have to create new compatible CSP with key from cert
            $cspParams = new-object System.Security.Cryptography.CspParameters
            $cspParams.ProviderType=24 #MS Enhanced RSA and AES CSP - this supports SHA256; see Computer\HKLM\SOFTWARE\Microsoft\Cryptography\Defaults\Provider Types\Type 024
            $cspParams.KeyContainerName=$algo.CspKeyContainerInfo.KeyContainerName
            $cspParams.KeyNumber = $algo.CspKeyContainerInfo.KeyNumber
            $cspParams.Flags = 'UseExistingkey'
            if($algo.CspKeyContainerInfo.MachineKeyStore) {$cspParams.Flags = $cspParams.Flags -bor 'UseMachineKeyStore'}

            $csp = new-object System.Security.Cryptography.RSACryptoServiceProvider($cspParams)
            }

        $sha256 = new-object System.Security.Cryptography.SHA256Cng

        # Create a signature of the JWT
        $Signature = [Convert]::ToBase64String($csp.SignData($dataToSign,$sha256))
        }
    else{
        #we will use CNG - use the provider
        $csp = $algo -as [System.Security.Cryptography.RsaCng]
        $sha256 = new-object System.Security.Cryptography.SHA256Cng
        #and create a signature
        $hash = $sha256.ComputeHash($dataToSign)
        $Signature = [Convert]::ToBase64String($csp.SignHash($hash,[System.Security.Cryptography.HashAlgorithmName]::SHA256,[System.Security.Cryptography.RSASignaturePadding]::Pkcs1))
        }

    # add signature to assertion
    $JWT = $JWT + "." + ($Signature -replace '\+','-' -replace '/','_' -replace '=')
    return $JWT
    }
function new-shortcut {
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
        [string]$name
        ,[parameter(Mandatory = $true)]
        [string]$path        
        ,[parameter(Mandatory = $true)]
        [string]$target        
        ,[parameter(Mandatory = $false)]
        [string]$runasUser
        )
    
    $path = $path.TrimEnd("\")
    $name = $name.TrimEnd(".lnk")
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut("$path\$name.lnk")
    if([string]::IsNullOrWhiteSpace($runasUser)){
        $Shortcut.TargetPath = $target
        $Shortcut.IconLocation = "$target,0"
        }
    else{
        $Shortcut.TargetPath = "runas.exe"
        $Shortcut.Arguments = "/user:$runasUser `"$target`""
        $Shortcut.IconLocation = "$target,0"
        
        }
    $Shortcut.Save()    
    }
function remove-diacritics{
    PARAM ([string]$String)
    [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($String))
    }
function remove-doubleQuotesFromCsv(){
    param (
        [Parameter(Mandatory=$true)]
        [string]
        $inputFile,

        [string]
        $outputFile
        )

    if (-not $outputFile){
        $outputFile = $inputFile
        }

    $inputCsv = Import-Csv $inputFile
    $quotedData = $inputCsv | ConvertTo-Csv -NoTypeInformation
    $outputCsv = $quotedData | % {$_ -replace  `
        '\G(?<start>^|,)(("(?<output>[^,"]*?)"(?=,|$))|(?<output>".*?(?<!")("")*?"(?=,|$)))' `
        ,'${start}${output}'}
    $outputCsv | Out-File $outputFile -Encoding utf8 -Force
    }
function sanitise-forJson(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
        [string]$dirtyString
        )
    $cleanString = $dirtyString.Replace('"','\"')
    $cleanString
    }
function sanitise-forMicrosoftEmailAddress(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
        [string]$dirtyString
        )
    $cleanString = $dirtyString -creplace '[^a-zA-Z0-9_@\-\.]+', ''
    do{$cleanString = $cleanString.Replace("..",".")}
    While($cleanString -match "\.\.")
    $cleanString = $cleanString.Trim(".")
    $cleanString = $cleanString.Replace(".@","@")
    $cleanString = $cleanString.Replace("@.","@")
    $cleanString
    }
function sanitise-forNetsuiteIntegration(){
    [cmdletbinding()]
    param(
        [Parameter(Mandatory =$true)]
        [string]$dodgyString
        )
    $lessDodgyString = remove-diacritics -String $dodgyString
    #$prettyGoodString = $lessDodgyString -replace "[^A-Za-z0-9_ ]",""     #Updated 2021-06-18 by Kev Maitland to *include* reserved SharePoint folder characters (otherwise removing these from NetSuite Opps/Projects won't trigger them to be processed as name changes)

    $trailingFullStopRegex = '[\.](?=.)'
    $prettyGoodString = $lessDodgyString -replace $trailingFullStopRegex  #Remove all . except final character

    $reservedSharePointFolderCharacterRegex = $([regex]::Escape('":*<>?/\|'))
    $evenBetterString = $prettyGoodString -replace "[^A-Za-z0-9_\. $reservedSharePointFolderCharacterRegex]"
    
    $evenBetterString.Replace(" ","").Replace(" ","").Replace(" ","")
    }
function sanitise-forPnpSharePoint($dirtyString){ 
    if([string]::IsNullOrWhiteSpace($dirtyString)){return}
    $cleanerString = sanitise-forSharePointStandard -dirtyString $dirtyString
    $cleanerString.Replace(":","").Replace("/","")
    if(@("."," ") -contains $dirtyString.Substring(($dirtyString.Length-1),1)){$dirtyString = $dirtyString.Substring(0,$dirtyString.Length-1)} #Trim trailing "."
    }
function sanitise-forSharePointFolderName(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
        [string]$dirtyString
        )
    $dirtyString.Replace("`"","").Replace("*","").Replace(":","").Replace("<","").Replace(">","").Replace("?","").Replace("/","").Replace("\","").Replace("|","").TrimEnd(".")
    }
function sanitise-forSharePointStandard($dirtyString){
    $dirtyString = $dirtyString.Trim()
    $dirtyString = $dirtyString.Replace(" "," ") #Weird instance where a space character is not a space character...
    if(@("."," ") -contains $dirtyString.Substring(($dirtyString.Length-1),1)){$dirtyString = $dirtyString.Substring(0,$dirtyString.Length-1)} #Trim trailing "."
    $dirtyString.Replace("`"","").Replace("#","").Replace("%","").Replace("?","").Replace("<","").Replace(">","").Replace("\","").Replace("...","").Replace("..","").Replace("'","`'").Replace("`t","").Replace("`r","").Replace("`n","").Replace("*","")
    }
function sanitise-LibraryNameForUrl($dirtyString){
    $cleanerString = $dirtyString.Trim()
    $cleanerString = $dirtyString -creplace '[^a-zA-Z0-9 _/]+', ''
    $cleanerString
    }
function sanitise-forSharePointListName($dirtyString){ 
    $cleanerString = sanitise-forSharePointStandard $dirtyString
    $cleanerString.Replace("/","")
    }
function sanitise-forSharePointFileName($dirtyString){ 
    $cleanerString = sanitise-forSharePointStandard $dirtyString
    $cleanerString.Replace("/","").Replace(":","")
    }
function sanitise-forSharePointFileName2($dirtyString){ 
    $dirtyString = $dirtyString.Trim()
    $dirtyString.Replace("`"","").Replace("#","").Replace("%","").Replace("?","").Replace("<","").Replace(">","").Replace("\","").Replace("/","").Replace("...","").Replace("..","").Replace("'","`'")
    if(@("."," ") -contains $dirtyString.Substring(($dirtyString.Length-1),1)){$dirtyString = $dirtyString.Substring(0,$dirtyString.Length-1)} #Trim trailing "."
    }
function sanitise-forSharePointGroupName($dirtyString){ 
    #"The group name is empty, or you are using one or more of the following invalid characters: " / \ [ ] : | < > + = ; , ? * ' @"
    $dirtyString = $dirtyString.Trim()
    $dirtyString.Replace("`"_","_").Replace("/","_").Replace("\","_").Replace("[","_").Replace("]","_").Replace(":","_").Replace("|","_").Replace("<","_").Replace(">","_").Replace("+","_").Replace("=","_").Replace(";","_").Replace(",","_").Replace("?","_").Replace("*","_").Replace("`'","_").Replace("@","_")
    if(@("."," ") -contains $dirtyString.Substring(($dirtyString.Length-1),1)){$dirtyString = $dirtyString.Substring(0,$dirtyString.Length-1)} #Trim trailing "."
    }
function sanitise-forSharePointFolderPath($dirtyString){ 
    $cleanerString = sanitise-forSharePointStandard $dirtyString
    $cleanerString.Replace(":","")
    }
function sanitise-forSharePointUrl($dirtyString){ 
    $dirtyString = $dirtyString.Trim()
    $dirtyString = $dirtyString.Replace(" "," ") #Weird instance where a space character is not a space character...
    $dirtyString = $dirtyString -creplace '[^a-zA-Z0-9 _/]+', ''
    #$dirtyString = $dirtyString.Replace("`"","").Replace("#","").Replace("%","").Replace("?","").Replace("<","").Replace(">","").Replace("\","/").Replace("//","/").Replace(":","")
    #$dirtyString = $dirtyString.Replace("$","`$").Replace("``$","`$").Replace("(","").Replace(")","").Replace("-","").Replace(".","").Replace("&","").Replace(",","").Replace("'","").Replace("!","")
    $cleanString =""
    for($i= 0;$i -lt $dirtyString.Split("/").Count;$i++){ #Examine each virtual directory in the URL
        if($i -gt 0){$cleanString += "/"}
        if($dirtyString.Split("/")[$i].Length -gt 50){$tempString = $dirtyString.Split("/")[$i].SubString(0,50)} #Truncate long folder names to 50 characters
            else{$tempString = $dirtyString.Split("/")[$i]}
        if($tempString.Length -gt 0){
            if(@(".", " ") -contains $tempString.Substring(($tempString.Length-1),1)){$tempString = $tempString.Substring(0,$tempString.Length-1)} #Trim trailing "." and " ", even if this results in a truncation <50 characters
            }
        $cleanString += $tempString
        }
    $cleanString = $cleanString.Replace("//","/").Replace("https/","https://") #"//" is duplicated to catch trailing "/" that might now be duplicated. https is an exception that needs specific handling
    $cleanString
    }
function sanitise-forResourcePath($dirtyString){
    if($dirtyString.Length -gt 0){
        if(@("."," ") -contains $dirtyString.Substring(($dirtyString.Length-1),1)){$dirtyString = $dirtyString.Substring(0,$dirtyString.Length-1)} #Trim trailing "."
        $dirtyString = $dirtyString.trim().replace("`'","`'`'")
        $dirtyString = $dirtyString.replace("#","").replace("%","") #As of 2017-05-26, these characters are not supported by SharePoint (even though https://msdn.microsoft.com/en-us/library/office/dn450841.aspx suggests they should be)
        #$dirtyString = $dirtyString -creplace "[^a-zA-Z0-9 _/()`'&-@!]+", '' #No need to strip non-standard characters
        #[uri]::EscapeUriString($dirtyString) #No need to encode the URL
        $dirtyString
        }
    }
function sanitise-forSql([string]$dirtyString){
    if([string]::IsNullOrWhiteSpace($dirtyString)){}
    else{$dirtyString.Replace("'","`'`'").Replace("`'`'","`'`'")}
    }
function sanitise-forSqlValue{
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [ValidateSet(“String”,”Int”,”Decimal”,"Boolean","Guid","Date","HTML")] 
        [string]$dataType

        ,[parameter(Mandatory = $false)]
        $value
        )
    switch($dataType){
        "String" {"`'$(smartReplace -mysteryString $value -findThis "'" -replaceWithThis "''")`'"}
        "HTML"   {"`'$(sanitise-forSqlValue -value $(sanitise-stripHtml $value ) -dataType String)`'"}
        "Int"    {if([string]::IsNullOrWhiteSpace($value)){"0"}else{$value}}
        "Decimal"{if([string]::IsNullOrWhiteSpace($value)){"0.0"}else{$value}}
        "Boolean"{if($value -eq $true){"1"}else{"0"}}
        "Guid"   {if([string]::IsNullOrWhiteSpace($value)){"NULL"}else{"`'$value`'"}} #This could be handled better
        "Date"   {if([string]::IsNullOrWhiteSpace($value)){"NULL"}else{"`'"+$(Get-Date (smartReplace -mysteryString $value -findThis "+0000" -replaceWithThis "") -Format s)+"`'"}}
        }
    }
function sanitise-forTermStore($dirtyString){
    #$dirtyString.Replace("\t", " ").Replace(";", ",").Replace("\", "\uFF02").Replace("<", "\uFF1C").Replace(">", "\uFF1E").Replace("|", "\uFF5C")
    $cleanerString = $dirtyString.Replace("`t", "").Replace(";", "").Replace("\", "").Replace("<", "").Replace(">", "").Replace("|", "").Replace("＆","&").Replace(" "," ").Trim()
    if($cleanerString.Length -gt 255){$cleanerString.Substring(0,254)}
    else{$cleanerString}
    }
function sanitise-stripHtml($dirtyString){
    if(![string]::IsNullOrWhiteSpace($dirtyString)){
        $cleanString = $dirtyString -replace '<[^>]+>',''
        $cleanString = [System.Web.HttpUtility]::HtmlDecode($cleanString)# -replace '&amp;','&'
        $cleanString
        }
    }
function set-suffixAndMaxLength(){
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory =$true)]
        [string]$string
        ,[Parameter(Mandatory =$false)]
        [string]$suffix
        ,[Parameter(Mandatory =$true)]
        [int]$maxLength
        )
    if($string.Length -gt ($maxLength-$suffix.length)){
        $outString = $string.Substring(0,$maxLength-$suffix.length)+$suffix
        }
    else{$outString = $string+$suffix}
    $outString
    }
function smartReplace($mysteryString,$findThis,$replaceWithThis){
    if([string]::IsNullOrEmpty($mysteryString)){$result = $mysteryString}
    else{$result = $mysteryString.ToString().Replace($findThis,$replaceWithThis)}
    $result
    }
function start-transcriptLog(){
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory =$true)]
            [string]$thisScriptName
        ,[Parameter(Mandatory =$false)]
            [string]$alternativeLogLocation 
        )

    #if(![string]::IsNullOrEmpty($MyInvocation.ScriptName)){
    #    $thisScriptName = $MyInvocation.ScriptName
    #    }

    if(![string]::IsNullOrWhiteSpace($alternativeLogLocation)){
        $logPath = "$alternativeLogLocation\$thisScriptName".Replace('\\','\')
        }
    else{$logPath = "C:\ScriptLogs\$thisScriptName".Replace('\\','\')}

    if((Test-Path $logPath) -eq $false){
        New-Item -Path $(split-Path $logPath -Parent) -Name $(split-Path $logPath -Leaf) -ItemType Directory -Force
        }
    $transcriptLogPath = "$($logPath+"\$thisScriptName")_Transcript_$(Get-Date -Format "yyMMdd").log".Replace('\\','\')
    Start-Transcript $transcriptLogPath -Append
    }
function stringify-hashTable($hashtable,$interlimiter,$delimiter){
    if([string]::IsNullOrWhiteSpace($interlimiter)){$interlimiter = ":"}
    if([string]::IsNullOrWhiteSpace($delimiter)){$delimiter = ", "}
    if($hashtable.Count -gt 0){
        $dirty = $($($hashtable.Keys | % {$_+"$interlimiter"+$hashtable[$_]+"$delimiter"}) -join "`r")
        $dirty.Substring(0,$dirty.Length-$delimiter.length)
        }
    }
function test-containsMatch{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory =$true)]
            [string[]]$arrayOfStrings
        ,[Parameter(Mandatory =$true)]
            [string]$regexToMatch
        ,[Parameter(Mandatory =$true)]
            [ValidateSet("All","Any")]
            [string]$matchType
        )
    <#
    .SYNOPSIS
    Tests each item in an array for -match and returns [boolean]
    
    .DESCRIPTION
    There are alternative methods for authenticating with Graph using certificates (see get-graphTokenResponse), but this is in-keping with our 5.1 codebase.
    
    .PARAMETER arrayOfStrings
    An array of Strings to test

    .PARAMETER regexToMatch
    Regex expression to -match against each item in $arrayOfStrings

    .PARAMETER matchType
    Test whether all Strings in $arrayOfStrings match $regexToMatch, or 1+ match
    #>    
    $isMatched = $false
    switch($matchType){
        ("Any") {
            $arrayOfStrings | ForEach-Object {
                $isMatched = $_ -match $regexToMatch
                Write-Verbose "$_ -match $regexToMatch = [$($_ -match $regexToMatch)][$isMatched]"
                if($isMatched -eq $true){continue}
            }
        }
        ("All") {
            $arrayOfStrings | ForEach-Object {
                $isMatched = $_ -match $regexToMatch
                Write-Verbose "$_ -match $regexToMatch = [$($_ -match $regexToMatch)][$isMatched]"
                if($isMatched -eq $false){continue}
            }
        }
    }
    return $isMatched
}
function test-isGuid(){
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$objectGuid
        )

    # Define verification regex
    [regex]$guidRegex = '(?im)^[{(]?[0-9A-F]{8}[-]?(?:[0-9A-F]{4}[-]?){3}[0-9A-F]{12}[)}]?$'

    # Check guid against regex
    return $objectGuid -match $guidRegex
    }
function test-mobileHandsetIsSupported(){
    Param (
        [parameter(Mandatory = $true)]
            [string]$modelCode
        )

    switch ($modelCode){
		"iPad Air"	{$true}
		"iPad Air 2"	{$true}
		"iPad Pro"	{$true}
		"iPhone 11"	{$true}
		"iPhone 11 Pro"	{$true}
		"iPhone 12"	{$true}
		"iPhone 12 mini"	{$true}
		"iPhone 12 Pro"	{$true}
		"iPhone 13 Pro"	{$true}
		"iPhone 6s"	{$true}
		"iPhone 7"	{$true}
		"iPhone 7 Plus"	{$true}
		"iPhone 8"	{$true}
		"iPhone 8 Plus"	{$true}
		"iPhone SE"	{$true}
		"iPhone X"	{$true}
		"iPhone XR"	{$true}
		"iPhone XS"	{$true}
		"Pixel 2"	{$false}
		"Pixel 3"	{$false}
		"Pixel 3a"	{$true}
		"Pixel 4a"	{$true}
		"Pixel 6"	{$true}
		"Nokia 6.1"	{$false}
		"TA-1012"	{$false}
		"ANE-LX1"	{$true}
		"CLT-L09"	{$true}
		"ELE-L09"	{$true}
		"FIG-LX1"	{$false}
		"Moto G (5S)"	{$false}
		"moto g(8) plus"	{$false}
		"ONEPLUS A3003"	{$false}
		"SM-A025G"	{$true}
		"SM-A217F"	{$true}
		"SM-A326B"	{$true}
		"SM-A505FN"	{$true}
		"SM-A515F"	{$true}
		"SM-A530F"	{$true}
		"SM-A705FN"	{$true}
		"SM-A715F"	{$true}
		"SM-A920F"	{$true}
		"SM-G780F"	{$true}
		"SM-G950F"	{$false}
		"SM-G950W"	{$false}
		"SM-G970F"	{$true}
		"SM-G973F"	{$true}
		"SM-G975F"	{$true}
		"SM-G977B"	{$true}
		"SM-G981B"	{$true}
		"SM-G996B"	{$true}
		"SM-N986B"	{$true}
		"SM-T720"	{$true}
		"M2007J20CG"	{$true}
		"Redmi Note 5"	{$false}
		
        }
    }

#endregion

