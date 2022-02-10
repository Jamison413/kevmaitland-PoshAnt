Begin{
#connect to AzureAD
Write-Host "Conéctese a Azure con la cuenta T..." -ForegroundColor Green
Connect-AzureAD

#Set selections
$officeLocationsToSelect = @("Andorra, AND",
"Bogota, COL",
"Barcelona, ESP",
"Manlleu, ESP",
"Madrid, ESP",
"Homeworker, AND",
"Homeworker, COL",
"Homeworker, ESP")
$languagesToSelect = @("ar-SA",
"bg-BG",
"ca-ES",
"cs-CZ",
"da-DK",
"de-DE",
"el-GR",
"en-GB",
"en-US",
"es-ES",
"es-MX",
"et-EE",
"eu-ES",
"fi-FI",
"fr-CA",
"fr-FR",
"gl-ES",
"he-IL",
"hr-HR",
"hu-HU",
"id-ID",
"it-IT",
"ja-JP",
"ko-KR",
"lt-LT",
"lv-LV",
"nb-NO",
"nl-NL",
"pl-PL",
"pt-BR",
"pt-PT",
"ro-RO",
"ru-RU",
"sk-SK",
"sl-SI",
"sr-Latn-CS",
"sr-Latn-RS",
"sv-SE",
"th-TH",
"tr-TR",
"uk-UA",
"vi-VN",
"zh-CN",
"zh-HK",
"zh-TW"
)

#get ESP users
write-host "Recuperando a todos los usuarios de ESP..."
$allESPUsers = Get-AzureADUser -All 10000 | Where-Object {($_.userType -EQ "Member") -and (($_.usageLocation -eq "ES") -or ($_.usageLocation -eq "AD") -or ($_.usageLocation -eq "CO")) -and (($_.DisplayName)[0] -ne "Ω")}

#select user email addresses
write-host "Seleccionar usuarios de ESP..." -ForegroundColor Green
$userEmailAddresses = $allESPUsers | select {$_.userPrincipalName},{$_.usageLocation},{$_.PhysicalDeliveryOfficeName},{$_.PreferredLanguage} | Out-GridView -Title "Seleccionar usuarios de ESP" -PassThru

#Set the officeLocation
write-host "Ingrese la ubicación de la oficina..." -ForegroundColor Green
$officeLocation = $officeLocationsToSelect | Out-GridView -Title "Ingrese la ubicación de la oficina" -PassThru
If(($officeLocation | Measure-Object).Count -gt 1){
#error - too many offices selected
$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Demasiadas selecciones, seleccione una ubicación de oficina para aplicar",0,"Error",0x1)
write-host "Ingrese la ubicación de la oficina..." -ForegroundColor Green
$officeLocation = $officeLocationsToSelect | Out-GridView -Title "Ingrese la ubicación de la oficina" -PassThru
}
Else{}

#Set the preferedLanguage
write-host "Seleccione el idioma para aplicar..." -ForegroundColor Green
$preferredLanguage = $languagesToSelect | Out-GridView -Title "Seleccione el idioma para aplicar" -PassThru
If(($preferredLanguage | Measure-Object).Count -gt 1){
#error - too many offices selected
$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Demasiadas selecciones, seleccione un idioma para aplicar",0,"Error",0x1)
write-host "Seleccione el idioma para aplicar..." -ForegroundColor Green
$preferredLanguage = $languagesToSelect | Out-GridView -Title "Seleccione el idioma para aplicar" -PassThru
}
Else{}

}
process{

    #last check for too many selections, exit if so
    If((($preferredLanguage | Measure-Object).Count -gt 1) -or (($officeLocation | Measure-Object).Count -gt 1)){
    #error - too many offices and/or languages selected
    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup("Demasiadas selecciones, inténtelo de nuevo y seleccione solo una oficina e idioma...Exiting",0,"Error",0x1)
    write-host "Demasiadas selecciones, inténtelo de nuevo y seleccione solo una oficina e idioma..." -ForegroundColor Green
    Exit
}


    #format email addresses and create an array to iterate through
    $arrayofUserEmailAddresses = convertTo-arrayOfEmailAddresses -blockOfText $userEmailAddresses.'$_.userPrincipalName'
    #Set officeLocation and preferedLanguage for each user in the array
    ForEach($user in $arrayofUserEmailAddresses){

        Write-Host "Ajuste officeLocation y preferredLanguage para $($user)" -ForegroundColor Green
        Set-AzureADUser -ObjectId $user -PreferredLanguage $preferredLanguage -PhysicalDeliveryOfficeName $officeLocation -Verbose
}

}
End{
#Disconnect-AzureAD
Write-Host "Completa" -ForegroundColor White


$applicationOffice = "Aplicado officeLocation: " + $officeLocation + "`n"
$applicationLanguage = "Aplicado preferredLanguage: " + $preferredLanguage + "`n`n"

$output = @()
$output += $applicationOffice
$output += $applicationLanguage

$output += $arrayofUserEmailAddresses | Format-Table -AutoSize | Out-String
$window = [System.Windows.Forms.MessageBox]::Show($($output),"Resultados")

}

