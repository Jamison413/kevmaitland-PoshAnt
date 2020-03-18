$inputPbxFileName = "$env:USERPROFILE\Downloads\$(get-date -Format "yyyyMMdd").csv"
$filename = "duffPlaceholder"
while($(Test-Path $inputPbxFileName) -eq $false){
    Write-Host "`$inputPbxFileName = [$inputPbxFileName]"-ForegroundColor Yellow
    Write-Host "`$fileName = [$fileName]"-ForegroundColor Yellow
    $fileName = Read-Host -Prompt "$tryAgain`Paste in the the file name that's in your Downloads or Desktop (e.g. MyFile.csv )"
    if(![string]::IsNullOrWhiteSpace($filename)){
        if(Test-Path "$env:USERPROFILE\Downloads\$fileName"){$inputPbxFileName = "$env:USERPROFILE\Downloads\$fileName"}
        elseif(Test-Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\$fileName"){$inputPbxFileName = "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\$fileName"}
        elseif(Test-Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop 1\$fileName"){$inputPbxFileName = "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop 1\$fileName"}
        elseif(Test-Path "$env:USERPROFILE\Desktop\$fileName"){$inputPbxFileName = "$env:USERPROFILE\Desktop\$fileName"}
        }
    $tryAgain = "Nope, that's not right. "
    }
$outputFileName = "payrollFile.txt"

if(Test-Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop 1\"){$desktopPath = "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop 1\"}
elseif(Test-Path "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\"){$desktopPath = "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\"}
elseif(Test-Path "$env:USERPROFILE\Desktop\"){$desktopPath = "$env:USERPROFILE\Desktop\"}

if(Test-Path "$desktopPath$outputFileName"){
    Remove-Item "$desktopPath$outputFileName"
    }
gc $inputPbxFileName | %{
    $thisLine = $_
    Write-Host -ForegroundColor Yellow $thisLine
    if(![string]::IsNullOrWhiteSpace($dirtyValue)){rv dirtyValue}
    if(![string]::IsNullOrWhiteSpace($sortCodeAndAccountNumber)){rv sortCodeAndAccountNumber}
    if(![string]::IsNullOrWhiteSpace($sortCodePayee)){rv sortCodePayee}
    if(![string]::IsNullOrWhiteSpace($accountNumberPayee)){rv accountNumberPayee}
    if(![string]::IsNullOrWhiteSpace($newSortCodePayee)){rv newSortCodePayee}
    if(![string]::IsNullOrWhiteSpace($newAccountNumberPayee)){rv newAccountNumberPayee}
    if(![string]::IsNullOrWhiteSpace($newNameOfPayee)){rv newNameOfPayee}
    if(![string]::IsNullOrWhiteSpace($newValue)){rv newValue}
    if(![string]::IsNullOrWhiteSpace($closingBlurb)){rv closingBlurb}

    $dirtyValue = $thisLine.Split(",")[0]
    if(!$dirtyValue){$dirtyValue = "NoValue"}
    $sortCodeAndAccountNumber = $thisLine.Split(",")[1]
    $sortCodePayee = $sortCodeAndAccountNumber.Substring(0,6)
    if(!$sortCodePayee){$sortCodePayee = "NoSort"}
    $accountNumberPayee = $sortCodeAndAccountNumber.Substring(6,$sortCodeAndAccountNumber.Length-6)
    if(!$accountNumberPayee){$accountNumberPayee = "NoAccount"}
    $dirtyWord1 = $thisLine.Split(",")[3]
    $dirtyWord3 = $thisLine.Split(",")[2]
    if($dirtyWord3.Length -gt 18){$dirtyWord3 = $dirtyWord3.Substring(0,18)}

    $newSortCodePayee = $($sortCodePayee.Substring(0,2)+"-"+$sortCodePayee.Substring(2,2)+"-"+$sortCodePayee.Substring(4,2)).PadRight(10," ")
    $newAccountNumberPayee = $accountNumberPayee.PadRight(13," ")
    $newNameOfPayee = $dirtyWord3.PadRight(18)
    $newValue = $("{0:N2}" -f $([double]$dirtyValue)).ToString().Replace(",","").PadLeft(10," ")
    $closingBlurb = "ANTHESIS ENERGY UK      40-11-60  41130994  0  Anthesis Enegy UK "
    Add-Content -Value $($newSortCodePayee+$newAccountNumberPayee+$newNameOfPayee+$newValue+"  "+$closingBlurb) -Path "$desktopPath\$outputFileName"  -Encoding UTF8
    }

[System.IO.File]::WriteAllLines("$desktopPath\$outputFileName", $(Get-Content "$desktopPath\$outputFileName"))
Start-Process "$desktopPath$outputFileName"
