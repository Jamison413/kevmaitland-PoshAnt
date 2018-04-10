$outlookContacts = Import-Csv 'C:\Users\kevinm\Desktop\contacts.csv'
$gmContacts = Import-Csv C:\Users\kevinm\Desktop\GoldMine_AllCompanies_AllContactsWithEmailAddresses.csv

$exceptThisString = "@.-" #We don't want to strip out @ or . or - symbols from e-mail addresses

(remove-nonAlphaNumeric -dirtyString $gmContacts[0].CONTSUPREF -exceptThisString $exceptThisString) -eq  "phil.jobson@roadchef.com"

function remove-nonAlphaNumeric($dirtyString, $exceptThisString){
    [Regex]$rgx = "[^a-zA-Z0-9$exceptThisString]"
    [regex]::Replace($dirtyString,$rgx,"")
    }

rv cleanedGMContacts
$gmContacts | % {
    $gmContact = $_
    $displayName = $null
    $firstName = $null
    $lastName = $null
    $phone1 = $null
    $phone2 = $null
    $mobile = $null
    $businessPhone = $null
    $email = $null
    $emailDomain = $null
    $companyIsValidated = $false

    if(![string]::IsNullOrWhiteSpace($gmContact.CONTACT)){
        $displayName = $(remove-nonAlphaNumeric -dirtyString $gmContact.CONTACT -exceptThisString " \'\-").Trim();
        if($displayName[0] -eq "'"){$displayName = $displayName.Substring(1,$displayName.length-1)}
        if($displayName[$displayName.Length-1] -eq "'"){$displayName = $displayName.Substring(0,$displayName.length-1)}

        $firstName = $(remove-nonAlphaNumeric -dirtyString $gmContact.CONTACT.Split(" ")[0] -exceptThisString "\'\-").Trim()
        if($firstName[0] -eq "'"){$firstName = $firstName.Substring(1,$firstName.length-1)}

        $lastName = $(remove-nonAlphaNumeric -dirtyString $gmContact.CONTACT.Split(" ")[$gmContact.CONTACT.Split(" ").Count-1] -exceptThisString "\'\-").Trim()
        if($lastName[$lastName.Length-1] -eq "'"){$lastName = $lastName.Substring(0,$lastName.length-1)}
        }
    else{
        $displayName = $null
        $firstName = $null
        $lastName = $null
        }
    $phone1 = format-internationalPhoneNumber -pDirtyNumber $gmContact.PHONE1 -p3letterIsoCountryCode "GBR" -localise $false -doNotReturnDuffNumbers $true
    $phone2 = format-internationalPhoneNumber -pDirtyNumber $gmContact.PHONE2 -p3letterIsoCountryCode "GBR" -localise $false -doNotReturnDuffNumbers $true
    if(![string]::IsNullOrWhiteSpace($gmcontact.CONTSUPREF)){
        if($gmContact.CONTSUPREF -match "@"){
            $email = $(remove-nonAlphaNumeric -dirtyString $gmContact.CONTSUPREF -exceptThisString "@\-\.").Trim()
            $emailDomain = $(remove-nonAlphaNumeric -dirtyString $gmContact.CONTSUPREF -exceptThisString "@\-\.").Trim().Split("@")[1]
            }
        else{$email = $null;$emailDomain = $null}
        }
    else{$email = $null;$emailDomain = $null}

    if(![string]::IsNullOrWhiteSpace($phone2)){
        if($phone2.Substring(0,5) -eq "+44 7"){$mobile = $phone2}
        else{$businessPhone = $phone2}
        }
    if(![string]::IsNullOrWhiteSpace($phone1)){#Phone1 takes priority over Phone2
        if($phone1.Substring(0,5) -eq "+44 7"){$mobile = $phone2}
        else{$businessPhone = $phone1}
        }
    if(![string]::IsNullOrWhiteSpace($gmcontact.UDIMCODE)){$companyIsValidated = $true}
    else{$companyIsValidated = $false}

    $newContact = New-Object psobject -Property $([ordered]@{
        "displayName"=$displayName;
        "firstName" = $firstName;
        "lastName"=$lastName;
        "email1"=$email;
        "email2"=$null;
        "email3"=$null;
        "businessPhone"=$businessPhone;
        "mobile"=$mobile;
        "companyId"=$gmContact.CompanyAccountNo;
        "company"=$gmContact.COMPANY.Trim();
        "companyType"=$gmContact.URECTYPE;
        "companyIsValidated"=$companyIsValidated;
        "companyEmailDomain"=$emailDomain;
        "jobTitle"=$gmContact.TITLE.Trim();
        "address1"=$gmContact.ADDRESS1.Trim();
        "address2"=$gmContact.ADDRESS2.Trim();
        "address3"=$gmContact.ADDRESS3.Trim();
        "address4"=$gmContact.CITY.Trim();
        "postcode"=$gmContact.ZIP.Trim();
        "source"="GoldMine";
        "scrapedFrom"=$null
        })

    if($displayName -and ($email -or $mobile -or $businessPhone)){[array]$cleanedGMContacts += $newContact}
    }
$outlookContacts | % {
    $olContact = $_
    $phone1 = $null
    $phone2 = $null
    $mobile = $null
    $businessPhone = $null
    $emailDomain = $null

    if(![string]::IsNullOrWhiteSpace($olContact.displayName)){
        $displayName=$(remove-nonAlphaNumeric -dirtyString $olContact.displayName -exceptThisString " \'\-").Trim()
        if($displayName[0] -eq "'"){$displayName = $displayName.Substring(1,$displayName.length-1)}
        if($displayName[$displayName.Length-1] -eq "'"){$displayName = $displayName.Substring(0,$displayName.length-1)}
        }
    else{$displayName = $null}
    if(![string]::IsNullOrWhiteSpace($olContact.firstName)){
        $firstName = $(remove-nonAlphaNumeric -dirtyString $olContact.firstName -exceptThisString "\'\-").Trim()
        if($firstName[0] -eq "'"){$firstName = $firstName.Substring(1,$firstName.length-1)}
        }
    else{$firstName = $null}
    if(![string]::IsNullOrWhiteSpace($olContact.lastName)){
        $lastName = $(remove-nonAlphaNumeric -dirtyString $olContact.lastName -exceptThisString "\'\-").Trim()
        if($lastName[$lastName.Length-1] -eq "'"){$lastName = $lastName.Substring(0,$lastName.length-1)}
        }
    else{$lastName = $null}

    $phone1 = format-internationalPhoneNumber -pDirtyNumber $olContact.businessPhone -p3letterIsoCountryCode "GBR" -localise $false -doNotReturnDuffNumbers $true
    $phone2 = format-internationalPhoneNumber -pDirtyNumber $olContact.mobile -p3letterIsoCountryCode "GBR" -localise $false -doNotReturnDuffNumbers $true

    if(![string]::IsNullOrWhiteSpace($phone2)){
        if($phone2.Substring(0,5) -eq "+44 7"){$mobile = $phone2}
        else{$businessPhone = $phone2}
        }
    if(![string]::IsNullOrWhiteSpace($phone1)){#Phone1 takes priority over Phone2
        if($phone1.Substring(0,5) -eq "+44 7"){$mobile = $phone2}
        else{$businessPhone = $phone1}
        }

    if(![string]::IsNullOrWhiteSpace($olContact.email1)){$emailDomain = $(remove-nonAlphaNumeric -dirtyString $olContact.email1 -exceptThisString "@\-\.").Trim().Split("@")[1]}
    else{$emailDomain = $null}

    $newContact = New-Object psobject -Property $([ordered]@{
        "displayName"=$displayName;
        "firstName" = $olContact.firstName;
        "lastName"=$olContact.lastName;
        "email1"=$olContact.email1;
        "email2"=$olContact.email2;
        "email3"=$olContact.email3;
        "businessPhone"=$businessPhone;
        "mobile"=$mobile;
        "companyId"=$null;
        "company"=$olContact.company.Trim();
        "companyType"=$null;
        "companyIsValidated"=$false;
        "companyEmailDomain"=$emailDomain;
        "jobTitle"=$olContact.jobTitle.Trim();
        "address1"=$null;
        "address2"=$null;
        "address3"=$null;
        "address4"=$null;
        "postcode"=$null;
        "source"="OutlookContactScrape";
        "scrapedFrom"=$olContact.scrapedFrom
        })

    if(($displayName -or $firstName -or $lastName) -and ($olContact.email1 -or $mobile -or $businessPhone)){[array]$cleanedGMContacts += $newContact}
    }

$cleanedGMContacts | Export-Csv C:\Users\kevinm\Desktop\combinedContact.csv -NoTypeInformation

$pippaList = import-csv 'C:\Users\kevinm\Desktop\combined Mailchimp contacts after Sustain contact database.csv'
$pippaList | %{
    
    if(![string]::IsNullOrWhiteSpace($_.'First Name')){$_.'First Name' = $($_.'First Name'.SubString(0,1).toUpper()+$_.'First Name'.SubString(1,$_.'First Name'.Length-1).toLower()).Trim()}
    if(![string]::IsNullOrWhiteSpace($_.'Last Name')){$_.'Last Name' = $($_.'Last Name'.SubString(0,1).toUpper()+$_.'Last Name'.SubString(1,$_.'Last Name'.Length-1).toLower()).Trim()}
    if(!([string]::IsNullOrWhiteSpace($_.'First Name') -and [string]::IsNullOrWhiteSpace($_.'Last Name'))){$_.displayName = $($_.'First Name' + " " + $_.'Last Name').Trim()}
    if(![string]::IsNullOrWhiteSpace($_.businessPhone)){$_.businessPhone = format-internationalPhoneNumber -pDirtyNumber $_.businessPhone -p3letterIsoCountryCode "GBR" -localise $false -doNotReturnDuffNumbers $true}
    if(![string]::IsNullOrWhiteSpace($_.mobile)){$_.mobile = format-internationalPhoneNumber -pDirtyNumber $_.mobile -p3letterIsoCountryCode "GBR" -localise $false -doNotReturnDuffNumbers $true}
    if(![string]::IsNullOrWhiteSpace($_.email1)){
        $_.email1 = $(remove-nonAlphaNumeric -dirtyString $_.email1 -exceptThisString "@\-\.").Trim()
        $_.companyEmailDomain = $(remove-nonAlphaNumeric -dirtyString $_.email1 -exceptThisString "@\-\.").Trim().Split("@")[1]
        }
    $_.companyIsValidated = $false
    if($_.'subscribed/unsub' -eq "subscriber"){[array]$goodList += $_}
    else{[array]$badList += $_}
    }

$pippaList | Export-Csv 'C:\Users\kevinm\Desktop\combined Mailchimp contacts after Sustain contact database2.csv' -NoTypeInformation

$combinedList = import-csv 'C:\Users\kevinm\Desktop\combinedGM&OutlookContact_new.csv'

$toRemove = Compare-Object $combinedList -DifferenceObject $badList -Property email1 -PassThru -ExcludeDifferent -IncludeEqual
$combinedListRemoved = Compare-Object $combinedList -DifferenceObject $badList -Property email1 -PassThru
$combinedListRemoved = Compare-Object $badList -DifferenceObject $combinedList -Property email1 -PassThru
$badList | %{
    #Add-Member -InputObject $_ -MemberType NoteProperty -Name email2 -Value $_.email1
    #Add-Member -InputObject $_ -MemberType NoteProperty -Name email3 -Value $_.email1
    $_.email2 = $_.email1
    $_.email3 = $_.email1
    }
$combinedListRemoved2 = Compare-Object $badList -DifferenceObject $combinedList -Property email2 -PassThru
$combinedListRemoved3 = Compare-Object $badList -DifferenceObject $combinedList -Property email3 -PassThru

$combinedList.Count
$combinedListRemoved.Count + $badList.Count

$combinedList.Count - $combinedListRemoved.Count
$dummy = $combinedListRemoved | ?{$_.SideIndicator -eq "=>"}
$dummy2 = $combinedListRemoved2 | ?{$_.SideIndicator -eq "=>"}
$dummy3 = $combinedListRemoved3 | ?{$_.SideIndicator -eq "=>"}

$dummy.Count
$badList.Count

$bigList = $dummy

$goodList | % {
    $pippaContact = $_
    $newContact = New-Object psobject -Property $([ordered]@{
    "displayName"=$pippaContact.displayName;
    "firstName" = $pippaContact.firstName;
    "lastName"=$pippaContact.lastName;
    "email1"=$pippaContact.email1;
    "email2"=$pippaContact.email2;
    "email3"=$pippaContact.email3;
    "businessPhone"=$pippaContact.businessPhone;
    "mobile"=$pippaContact.mobile;
    "companyId"=$null;
    "company"=$pippaContact.company.Trim();
    "companyType"=$pippaContact.companyType;
    "companyIsValidated"=$false;
    "companyEmailDomain"=$pippaContact.companyEmailDomain;
    "jobTitle"=$pippaContact.jobTitle.Trim();
    "address1"=$pippaContact.address1;
    "address2"=$pippaContact.address2;
    "address3"=$pippaContact.address3;
    "address4"=$pippaContact.address4;
    "postcode"=$pippaContact.postcode;
    "source"=$pippaContact.source;
    "scrapedFrom"=$pippaContact.scrapedFrom})
    $bigList += $newContact
    }


$bigList | Export-Csv 'C:\Users\kevinm\Desktop\combinedSuperList.csv' -NoTypeInformation