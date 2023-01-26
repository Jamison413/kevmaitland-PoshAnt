
if($(Get-Module -ListAvailable -Name pnp.powershell) -ne $null){
    Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/sites/Resources-HR/" -Interactive
}
$trainingRecords = (Get-PnPListItem -List "User Training Records" -Query "<View><Query><Where><Gt><FieldRef Name='Date_x0020_of_x0020_training' /><Value Type='DateTime'>2022-01-01T12:00:00Z</Value></Gt></Where></Query></View>") |  % {Add-Member -InputObject $_ -MemberType NoteProperty -Name Guid -Value $_.FieldValues.GUID.Guid;$_}

$cng = $trainingRecords | Where-Object {$_.FieldValues.User.Email -match "climate"}

$prettyCng = 
$cng | Select-Object {$_.FieldValues.User.LookupValue}, {$_.FieldValues.User.Email}, {$_.FieldValues.Training_x0020_session.Label}, {$_.FieldValues.Date_x0020_of_x0020_training} | Export-Csv -Path "$env:USERPROFILE\Downloads\CNG_UserTrainingRecords.csv" -NoTypeInformation -Encoding UTF8 
