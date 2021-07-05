$destination = "https://gbrenergy.file.core.windows.net/ecodata"
$sasToken = "?sv=2020-02-10&ss=f&srt=sco&sp=rwdlc&se=2021-07-05T17:07:04Z&st=2021-06-17T09:07:04Z&sip=89.197.96.6&spr=https&sig=ixPdh%2F4La9Cr%2BaSZ0JeHVuxbXTdrldmCEJzuh1IRAXc%3D"

#ECO data
    $sources = @()
	$sources += "R:\FTP\ECO"
	$sources += "X:Internal\Residential"
	$sources += "X:Internal\ECO"
	$sources += "R:\X\Internal\ECO"
	$sources += "R:\X\Internal\Residential"
    $sources += "X:\Clients\British Gas Trading Limited"
	$sources += "X:\Clients\E.ON Energy Solutions Limited"
    $sources += "X:\Clients\EDF Energy"
	$sources += "X:\Clients\npower Ltd"
    $sources += "X:\Clients\Npower Northern Ltd"
    $sources += "X:\Clients\Ovo Energy"
    $sources += "X:\Clients\SSE Energy Supply Limited"
	$sources += "X:\Clients\SSE Homes Services Limited"
	$sources += "R:\X\Clients\British Gas Trading Limited"
	$sources += "R:\X\Clients\E.ON Energy Solutions Limited"
	$sources += "R:\X\Clients\EDF Energy"
    $sources = @()
	$sources += "R:\X\Clients\npower Ltd"
	$sources += "R:\X\Clients\Npower Northern Ltd"
    $sources += "R:\X\Clients\Ovo Energy"
	$sources += "R:\X\Clients\SSE Energy Supply Limited"
	$sources += "R:\X\Clients\SSE Homes Services"
	$sources += "R:\X\Clients\SSE Homes Services Limited"
    $sources += "X:\Suppliers\Addvertising Green Limited"
    $sources += "X:\Suppliers\Bartons of Duke Street"
    $sources += "X:\Suppliers\DB Insulations"
    $sources += "X:\Suppliers\Diamond Bead Ltd"
    $sources += "X:\Suppliers\DM Developments"
    $sources += "X:\Suppliers\Domestic & General"
    $sources += "X:\Suppliers\Eco-Worx Ltd"
    $sources += "X:\Suppliers\Energy Low Ltd"
    $sources += "X:\Suppliers\Evolve Home Energy Solutions"
    $sources += "X:\Suppliers\Green Eco Hub"
    $sources += "X:\Suppliers\Greener Skies Ltd"
    $sources += "X:\Suppliers\IGA Ltd"
    $sources += "X:\Suppliers\Insulate UK Ltd"
    $sources += "X:\Suppliers\Insulation (Whalley) Ltd"
    $sources += "X:\Suppliers\Insulation ECO Limited"
    $sources += "X:\Suppliers\JJ Crump & Sons"
    $sources += "X:\Suppliers\Thermabead Ltd"
    $sources += "X:\Suppliers\Think Go Green"
    $sources += "X:\Suppliers\Well Warm Ltd"
	$sources += "R:\X\Suppliers"
    $sources += "X:\Databases\ECO2"
    $sources += "R:\X\Databases\ECO"
    $sources += "R:\X\Databases\ECO2"
    $sources += "R:\X\Databases\Resi"

$sources = @()
$sources += "X:\Suppliers\Climate Insulation"
$sources += "X:\Suppliers\Invictus Energy Group Ltd"
$sources += "X:\Suppliers\Landlord Certs London Limited"
$sources += "X:\Suppliers\Marshall & McCourt Plumbing & Heating Contracts Ltd (2)"
$sources += "X:\Suppliers\The Green Deal Factory Ltd"
$sources += "X:\Suppliers\The Green Eco Company NW"
$sources += "X:\Suppliers\Titan Property Services Limited"
$sources += "X:\Suppliers\Western Isles Insulation Ltd"
$sources += "X:\Suppliers\ECO FRAMEWORK CONTRACTORS"
$sources += "X:\Suppliers\ECO Health and Safety Report"
$sources += "X:\Suppliers\Llewellyn Smith Ltd"
$sources += "X:\Suppliers\Pennington Choices"
$sources += "X:\Suppliers\THS"

    
$sources | % {
    $thisSource = $_
    $thisDestination = (Split-Path $thisSource -Parent).Replace("X:",$destination).Replace("R:\X",$destination).Replace("R:",$destination).Replace("\","/") 
    #azcopy copy "$thisSource" "$thisDestination$sasToken" --recursive=true --put-md5 --preserve-smb-info --backup --cap-mbps 35 
    azcopy copy "$thisSource" "$thisDestination$sasToken" --recursive=true --put-md5 --backup --cap-mbps 35 --overwrite ifSourceNewer
    robocopy "$thisSource" "$($thisSource.Replace("E\X:","W:").Replace("R\X:","W:").Replace("D:","W:").Replace("X:","W:"))" /XF *.* /E /DCOPY:T /XJ
    }
Send-MailMessage -From migrationbot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "AzureFiles migration completed!" -Body "$($sources -join "`r`n")`r`n`r`n`r`nThese ones are complete - I need topping up :)" -To @("kevin.maitland@anthesisgroup.com","emily.pressey@anthesisgroup.com","andrew.ost@anthesisgroup.com")  -Encoding UTF8 -Priority High
