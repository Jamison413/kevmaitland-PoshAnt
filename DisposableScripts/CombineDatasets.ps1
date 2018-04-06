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