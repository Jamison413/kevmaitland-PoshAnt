$inputPbxFileName = "$env:USERPROFILE\Desktop\BASUS001_Employee.pbx"
$outputFileName = "payrollFile.txt"

Remove-Item "$env:USERPROFILE\Desktop\$outputFileName" 
gc $inputPbxFileName | %{
    $thisLine = $_
    $sortCodePayee = $thisLine.Substring(0,6)
    $accountNumberPayee = $thisLine.Substring(6,8)
    $paymentType = $thisLine.Substring(14,3)
    $sortCodePayer = $thisLine.Substring(17,6)
    $accountNumberPayer = $thisLine.Substring(23,8)
    $nfi = $thisLine.Substring(31,4)
    $dirtyValue = $thisLine.Substring(35,11)
    $dirtyWord1 = $thisLine.Substring(46,18)
    $dirtyWord2 = $thisLine.Substring(64,18)
    $dirtyWord3 = $thisLine.Substring(82,18)

    $newSortCodePayee = $($sortCodePayee.Substring(0,2)+"-"+$sortCodePayee.Substring(2,2)+"-"+$sortCodePayee.Substring(4,2)).PadRight(10," ")
    $newAccountNumberPayee = $accountNumberPayee.PadRight(13," ")
    $newNameOfPayee = $dirtyWord3
    $newValue = $("{0:N2}" -f $([double]$dirtyValue/100)).ToString().Replace(",","").PadLeft(10," ")
    $description =$dirtyWord1.PadRight(24," ")
    $newSortCodePayer = $($sortCodePayer.Substring(0,2)+"-"+$sortCodePayer.Substring(2,2)+"-"+$sortCodePayer.Substring(4,2)).PadRight(10," ")
    $newAccountNumberPayer = $accountNumberPayer.PadRight(10," ")
    $closingBlurb = "0  Main Current a/c  "

    Add-Content -Value $($newSortCodePayee+$newAccountNumberPayee+$newNameOfPayee+$newValue+"  "+$description+$newSortCodePayer+$newAccountNumberPayer+$closingBlurb) -Path "$env:USERPROFILE\Desktop\$outputFileName"  -Encoding UTF8
    }

[System.IO.File]::WriteAllLines("$env:USERPROFILE\Desktop\$outputFileName", $(Get-Content "$env:USERPROFILE\Desktop\$outputFileName"))

#$thisLine = $(gc $env:USERPROFILE\Desktop\BASUS001_Employee.pbx)[0]
