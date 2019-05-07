$csvFiles = @("C:\Users\kevinm\Desktop\eco3_deemed_scores_v1.2.1.csv","C:\Users\kevinm\Desktop\eco3_deemed_scores_2018_2.csv","C:\Users\kevinm\Desktop\eco3_deemed_scores_2018_3.csv")

$csvData1 = import-csv $csvFiles[0]
$csvData2 = import-csv $csvFiles[1]
$csvData3 = import-csv $csvFiles[2]
$csvDatas = @($csvData1,$csvData2,$csvData3)
$myThings = @(0) * ($csvData1.Count+$csvData2.Count+$csvData3.Count)
$j = 1
$i = 0
foreach ($csvDataBlob in $csvDatas) {
    Write-Progress -Activity "Processing file #$j"
    $csvDataBlob | %{
    Write-Progress -Activity "Processing $($csvDataBlob.Count) Records" -Status "Processing Row [$($i+1)] ($($i / $csvDataBlob.Count*100))" -PercentComplete $($i / $csvDataBlob.Count*100)
#Get-Content $csvFile | %{
    $thisRow = $_
    #$thisObject = New-Object -TypeName psobject -ArgumentList $([ordered]@{[string]"id"=$null;[string]"version"="ECO3";[string]"measure"=$null;[string]"measureVariant"=$null;[string]"propertyType"=$null;[string]"bedrooms"=$null;[string]"preHeatingSystem"=$null;[string]"postHeatingSystem"=$null;[string]"annualSaving"=$null;[string]"costSaving"=$null;[string]"uValueDelta"=$null;[string]"wallType"=$null;[string]"lifetime"=$null;[string]"meanPopt"=$null;[string]"ageBand"=$null;[string]"issMeasure"=$null;[string]"issMeasureName"=$null;[string]"issWallType"=$null;[string]"issThermalConductivity"=$null;[string]"issDoorType"=$null;[string]"issGlazing"=$null;[string]"issRiri"=$null;[string]"issLoftInsulation"=$null;[string]"issPreExistingHeatingControls"=$null;[string]"issPropertyType"=$null;[string]"issBedrooms"=$null;[string]"issExtWalls"=$null;[string]"issDetatchment"=$null})
    $thisObject = [pscustomobject]@{[string]"id"="";[string]"version"="ECO3";[string]"measure"="";[string]"measureVariant"="";[string]"propertyType"="";[string]"bedrooms"="";[string]"preHeatingSystem"="";[string]"postHeatingSystem"="";[string]"annualSaving"="";[string]"costSaving"="";[string]"uValueDelta"="";[string]"wallType"="";[string]"lifetime"="";[string]"meanPopt"="";[string]"ageBand"="";[string]"issMeasure"="";[string]"issMeasureName"="";[string]"issWallType"="";[string]"issThermalConductivity"="";[string]"issDoorType"="";[string]"issGlazing"="";[string]"issRiri"="";[string]"issLoftInsulation"="";[string]"issPreExistingHeatingControls"="";[string]"issPropertyType"="";[string]"issBedrooms"="";[string]"issExtWalls"="";[string]"issDetatchment"="";[string]"upliftName"="";[string]"upliftValue"=""}
    
    $thisObject.measure = $thisRow.'Measure Category'
    $thisObject.upliftName = $thisRow.Name_of_Uplift
    $thisObject.upliftValue = $thisRow.Uplift
    $thisObject.costSaving = $thisRow.'Cost_Score_(�)'
    $thisObject.annualSaving = $thisRow.'Annual Saving (�)'
    $thisObject.meanPopt = $thisRow.'average POPT factor'
    $thisObject.lifetime = $thisRow.L
    $thisObject.preHeatingSystem = $thisRow.Pre_Main_Heating_Source_for_the_Property
    
    switch($thisRow.'Measure Category'){
        ("Solid Wall Insulation") {
            $thisObject.measureVariant = $thisRow.Measure_Type.Split("_")[0]
            $thisObject.wallType = $thisRow.Measure_Type.Split("_")[1]
            $thisObject.uValueDelta = $thisRow.Measure_Type.Split("_")[2] + " -> " + $thisRow.Measure_Type.Split("_")[3]
            $thisObject.postHeatingSystem = $thisObject.null

            $thisObject.issMeasure = "Solid Wall Insulation"
            $thisObject.issMeasureName = $thisRow.Measure_Type.Split("_")[0]+"_"+$thisRow.Measure_Type.Split("_")[1]
            $thisObject.issWallType = "Please specify"
            }
        ("Cavity Wall Insulation"){
            $thisObject.measureVariant = $thisRow.Measure_Type
            $thisObject.wallType = $thisObject.wallType
            $thisObject.uValueDelta = $thisObject.uValueDelta
            $thisObject.postHeatingSystem = $thisObject.postHeatingSystem

            $thisObject.issWallType = "Cavity"
            $thisObject.issMeasure = "Cavity Wall Insulation"
            $thisObject.issMeasureName = $thisRow.Measure_Type.Split("_")[0]
            $thisObject.issThermalConductivity = $thisRow.Measure_Type.Split("_")[1]
            if($thisObject.issThermalConductivity = "Cavity"){$thisObject.issThermalConductivity = $thisObject.null}
            }
        ("Other Insulation"){
            $thisObject.measureVariant = $thisRow.Measure_Type
            $thisObject.wallType = $thisObject.null
            $thisObject.uValueDelta = $thisObject.null
            $thisObject.postHeatingSystem = $thisObject.null

            $thisObject.issMeasure = "Other Insulation"
            $thisObject.issMeasureName = $thisRow.Measure_Type.Split("_")[0]
            switch($thisObject.issMeasureName){
                "HPED" {
                    if($thisRow.Measure_Type.Split("_")[1] -eq "less"){$thisObject.issDoorType = "60% or less"}
                    elseif($thisRow.Measure_Type.Split("_")[1] -eq "greater"){$thisObject.issDoorType = "More than 60%"}
                    else{$thisObject.issDoorType = $thisObject.null}
                    }
                "WG" {
                    if($thisRow.Measure_Type.Split("_")[1] -eq "singletodouble"){$thisObject.issGlazing = "Single to Double Glazing"}
                    elseif($thisRow.Measure_Type.Split("_")[1] -eq "improveddouble"){$thisObject.issGlazing = "Improved Double Glazing"}
                    else{$thisObject.issGlazing = $thisObject.null}
                    }
                "RIRI" {
                    if($thisRow.Measure_Type.Split("_")[2] -eq "unin"){$thisObject.issRiri = "No"}
                    elseif($thisRow.Measure_Type.Split("_")[2] -eq "in"){$thisObject.issRiri = "Yes"}
                    else{$thisObject.issRiri = $thisObject.null}
                    }
                default{}
                }
            }
        ("Loft Insulation"){
            $thisObject.measureVariant = $thisRow.Measure_Type.Split("_")[1]
            $thisObject.wallType = $thisObject.null
            $thisObject.uValueDelta = $thisObject.null
            $thisObject.postHeatingSystem = $thisObject.null
            
            $thisObject.issMeasure = "Loft Insulation"
            $thisObject.issMeasureName = "LI"
            if($thisObject.measureVariant -eq "lessequal100"){$thisObject.issLoftInsulation = "100mm or less"}
            elseif($thisObject.measureVariant -eq "greater100"){$thisObject.issLoftInsulation = "More than 100mm"}
            else{$thisObject.issLoftInsulation = $thisObject.null}
            }
        ("Boiler"){
            $thisObject.measureVariant = $thisRow.Measure_Type
            $thisObject.wallType = $thisRow.Measure_Type.Split("_")[2]
            if($thisObject.wallType = "CH"){$thisObject.wallType = $thisRow.Measure_Type.Split("_")[3]}
            $thisObject.uValueDelta = $thisObject.null
            $thisObject.postHeatingSystem = $thisRow.Post_Main_Heating_Source_for_the_Property

            $thisObject.issMeasure = "Boiler"
            $thisObject.issMeasureName = "B_"+$thisRow.Measure_Type.Split("_")[1]
            if($thisObject.issMeasureName -eq "First"){$thisObject.issMeasureName = "B_FTCH"}
            if($thisObject.wallType = "solid"){$thisObject.issWallType = "Solid Wall"}
            if($thisObject.wallType = "cavity"){$thisObject.issWallType = "Cavity"}
            else{$thisObject.issWallType = ""}
            if($thisRow.Measure_Type.Split("_")[3] -eq "nopreHCs"){$thisObject.issPreExistingHeatingControls = "No"}
            elseif($thisRow.Measure_Type.Split("_")[3] -eq "preHCs"){$thisObject.issPreExistingHeatingControls = "Yes"}
            else{$thisObject.issPreExistingHeatingControls = $thisObject.null}
            }
        ("Other Heating"){
            $thisObject.measureVariant = $thisRow.Measure_Type
            $thisObject.wallType = $thisRow.Measure_Type.Split("_")[2]
            $thisObject.uValueDelta = $thisObject.null
            $thisObject.postHeatingSystem = $thisObject.null

            $thisObject.issMeasure = "Other Heating"
            $thisObject.issMeasureName = "Heating_controls"
            if($thisObject.wallType = "solid"){$thisObject.issWallType = "Solid Wall"}
            if($thisObject.wallType = "cavity"){$thisObject.issWallType = "Cavity"}
            else{$thisObject.issWallType = ""}
            }
        ("Micro-Generation"){
            $thisObject.measureVariant = $thisRow.Measure_Type
            $thisObject.wallType = $thisObject.null
            $thisObject.uValueDelta = $thisObject.null
            $thisObject.postHeatingSystem = $thisObject.null

            $thisObject.issMeasure = ""
            $thisObject.issMeasureName = "Solar_PV"
            }
        ("ESH"){
            $thisObject.measureVariant = $thisRow.Measure_Type
            $thisObject.wallType = $thisRow.Measure_Type.Split("_")[3]
            $thisObject.uValueDelta = $thisObject.null
            $thisObject.postHeatingSystem = $thisObject.null

            $thisObject.issMeasure = ""
            $thisObject.issMeasureName = $thisRow.Measure_Type.Split("_")[1]
            if($thisObject.wallType = "solid"){$thisObject.issWallType = "Solid Wall"}
            if($thisObject.wallType = "cavity"){$thisObject.issWallType = "Cavity"}
            else{$thisObject.issWallType = ""}
            }
        }
        
    #$thisRow
    if([string]::IsNullOrWhiteSpace($thisRow.Property_Type)){Write-Warning "$thisRow [$i] has a problem with ProprtyType"}
    if($thisRow.Property_Type -match "2W_Flat"){
        $thisObject.propertyType = "2 ext. Wall Flat"
        [string]$thisObject.bedrooms = $thisRow.Property_Type.Split("_")[2]
        $thisObject.issPropertyType = "Flat"
        $thisObject.issExtWalls = "2 or fewer"
        if([string]$thisObject.bedrooms -eq "3+"){[string]$thisObject.issBedrooms = "3 or more"}
        else{[string]$thisObject.issBedrooms = [string]$thisObject.bedrooms}
        }
    elseif($thisRow.Property_Type -match "3W_Flat"){
        $thisObject.propertyType = "3 ext. Wall Flat"
        [string]$thisObject.bedrooms = $thisRow.Property_Type.Split("_")[2]
        $thisObject.issPropertyType = "Flat"
        $thisObject.issExtWalls = "3 or more"
        if([string]$thisObject.bedrooms -eq "3+"){[string]$thisObject.issBedrooms = "3 or more"}
        else{[string]$thisObject.issBedrooms = [string]$thisObject.bedrooms}
        }
    elseif($thisRow.Property_Type -match "End-terrace"){
        $thisObject.propertyType = "End-terrace"
        [string]$thisObject.bedrooms = $thisRow.Property_Type.Split("_")[1]
        $thisObject.issPropertyType = "House"
        $thisObject.issDetatchment = "End-Terrace"
        if([string]$thisObject.bedrooms -eq "5+"){[string]$thisObject.issBedrooms = "5 or more"}
        else{[string]$thisObject.issBedrooms = [string]$thisObject.bedrooms}
        }
    elseif($thisRow.Property_Type -match "Mid-terrace"){
        $thisObject.propertyType = "Mid-terrace"
        [string]$thisObject.bedrooms = $thisRow.Property_Type.Split("_")[1]
        $thisObject.issPropertyType = "House"
        $thisObject.issDetatchment = "Mid-Terrace"
        if([string]$thisObject.bedrooms -eq "5+"){[string]$thisObject.issBedrooms = "5 or more"}
        else{[string]$thisObject.issBedrooms = [string]$thisObject.bedrooms}
        }
    elseif($thisRow.Property_Type.SubString(0,4) -eq "Semi"){
        $thisObject.propertyType = "Semi-detatched"
        [string]$thisObject.bedrooms = $thisRow.Property_Type.Split("_")[1]
        $thisObject.issPropertyType = "House"
        $thisObject.issDetatchment = "Semi-Detatched"
        if([string]$thisObject.bedrooms -eq "2-"){[string]$thisObject.issBedrooms = "2 or fewer"}
        elseif([string]$thisObject.bedrooms -eq "5+"){[string]$thisObject.issBedrooms = "5 or more"}
        else{[string]$thisObject.issBedrooms = [string]$thisObject.bedrooms}
        }
    elseif($thisRow.Property_Type.SubString(0,3) -eq "Det"){
        $thisObject.propertyType = "Detatched"
        [string]$thisObject.bedrooms = $thisRow.Property_Type.Split("_")[1]
        $thisObject.issPropertyType = "House"
        $thisObject.issDetatchment = "Detatched"
        if([string]$thisObject.bedrooms -eq "2-"){[string]$thisObject.issBedrooms = "2 or fewer"}
        elseif([string]$thisObject.bedrooms -eq "6+"){[string]$thisObject.issBedrooms = "6 or more"}
        else{[string]$thisObject.issBedrooms = [string]$thisObject.bedrooms}
        }
    elseif($thisRow.Property_Type -match "Bung_Semi"){
        $thisObject.propertyType = "Bungalow - semi detached & end terrace"
        [string]$thisObject.bedrooms = $thisRow.Property_Type.Split("_")[2]
        $thisObject.issPropertyType = "House"
        $thisObject.issDetatchment = "End-Terrace"
        if([string]$thisObject.bedrooms -eq "3+"){[string]$thisObject.issBedrooms = "3 or more"}
        else{[string]$thisObject.issBedrooms = [string]$thisObject.bedrooms}
        }
    elseif($thisRow.Property_Type -match "Bung_Det"){
        $thisObject.propertyType = "Bungalow - detached"
        [string]$thisObject.bedrooms = $thisRow.Property_Type.Split("_")[2]
        $thisObject.issPropertyType = "House"
        $thisObject.issDetatchment = "End-Terrace"
        if([string]$thisObject.bedrooms -eq "2-"){[string]$thisObject.issBedrooms = "2 or fewer"}
        elseif([string]$thisObject.bedrooms -eq "3+"){[string]$thisObject.issBedrooms = "3 or more"}
        else{[string]$thisObject.issBedrooms = [string]$thisObject.bedrooms}
        }
    elseif($thisRow.Property_Type -match "Bung_Mid"){
        $thisObject.propertyType = "Bungalow - mid terrace"
        [string]$thisObject.bedrooms = $thisRow.Property_Type.Split("_")[2]
        $thisObject.issPropertyType = "House"
        $thisObject.issDetatchment = "Mid-Terrace"
        if([string]$thisObject.bedrooms -eq "3+"){[string]$thisObject.issBedrooms = "3 or more"}
        else{[string]$thisObject.issBedrooms = [string]$thisObject.bedrooms}
        }
    elseif($thisRow.Property_Type -match "2W_Maisonette"){
        $thisObject.propertyType = "Maisonette 2 ext. Wall"
        [string]$thisObject.bedrooms = $thisRow.Property_Type.Split("_")[2]
        $thisObject.issPropertyType = "Maisonette"
        $thisObject.issExtWalls = "2 or fewer"
        if([string]$thisObject.bedrooms -eq "3+"){[string]$thisObject.issBedrooms = "3 or more"}
        else{[string]$thisObject.issBedrooms = "2 or fewer"}
        }
    elseif($thisRow.Property_Type -match "3W_Maisonette"){
        $thisObject.propertyType = "Maisonette 3 ext. Wall"
        [string]$thisObject.bedrooms = $thisRow.Property_Type.Split("_")[2]
        $thisObject.issPropertyType = "Maisonette"
        $thisObject.issExtWalls = "3 or more"
        if([string]$thisObject.bedrooms -eq "3+"){[string]$thisObject.issBedrooms = "3 or more"}
        else{[string]$thisObject.issBedrooms = "2 or fewer"}
        }

    switch($thisObject.uValueDelta.Split(" ")[0]){
        ("2.0"){
            switch($thisObject.wallType){
                ("system"){$thisObject.ageBand = "Before 1976 (E&W) / Before 1976 (S)"}
                ("timber"){$thisObject.ageBand = "Before 1949 (E&W) / Before 1949 (S)"}
                default {$thisObject.ageBand = $thisObject.null}
                }
            }
        ("1.7"){
            switch($thisObject.wallType){
                {$_ -in "solid","cavity","stone"}{$thisObject.ageBand = "Before 1976 (E&W) / Before 1976 (S)"}
                ("system"){$thisObject.ageBand = "1967 - 1975 (E&W) / 1965 - 1975 (S)"}
                default {$thisObject.ageBand = $thisObject.null}
                }
            }
        ("1.0"){
            switch($thisObject.wallType){
                {$_ -in "solid","cavity","stone"}{$thisObject.ageBand = "1976 - 1982 (E&W) / 1976 - 1983 (S)"}
                ("system"){$thisObject.ageBand = "1976 - 1982 (E&W) / 1976 - 1983 (S)"}
                ("timber"){$thisObject.ageBand = "1950 - 1966 (E&W) / 1950 - 1964 (S)"}
                default {$thisObject.ageBand = $thisObject.null}
                }
            }
        ("0.6"){
            switch($thisObject.wallType){
                {$_ -in "solid","cavity","stone","system"}{$thisObject.ageBand = "1983 - 1995 (E&W) / 1984 - 1991 (S)"}
                ("cob"){$thisObject.ageBand = "Pre 1996 (E&W) / Pre 1999 (S)"}
                ("timber"){$thisObject.ageBand = "1967 - 1975 (E&W) / 1965 - 1975 (S)"}
                ("filled"){$thisObject.ageBand = "Pre 1976 (E&W) / Pre 1976 (S)"}
                default {$thisObject.ageBand = $thisObject.null}
                }
            }
        ("0.45"){
            switch($thisObject.wallType){
                {$_ -in "solid","cavity","stone","system","cob"}{$thisObject.ageBand = "From 1996 (E&W) / From 1992 (S)"}
                {$_ -in "timber","filled"}{$thisObject.ageBand = "from 1976 (E&W) / from 1976 (S)"}
                default {$thisObject.ageBand = $thisObject.null}
                }
            }
        }
    $myThings[$i] = $thisObject
    #Add-Content -Value ",ECO3,$measure,$measureVariant,$propertyType,$bedrooms,$preHeatingSystem,$postHeatingSystem,$annualSaving,$costSaving,$uValueDelta,$wallType,$lifetime,$meanPopt,$ageBand,$issMeasure,$issMeasureName,$issWallType,$issThermalConductivity,$issDoorType,$issGlazing,$issRiri,$issLoftInsulation,$issPreExistingHeatingControls,$issPropertyType,$issBedrooms,$issExtWalls,$issDetatchment" -Path C:\Users\kevinm\Desktop\eco3_deemedScoresParsed6.csv
    $i++
    }
    }

$myThings | Export-Csv -Path c:\users\kevinm\desktop\deemed5.csv -NoTypeInformation

$sql = "INSERT INTO t_deemedScores()"
