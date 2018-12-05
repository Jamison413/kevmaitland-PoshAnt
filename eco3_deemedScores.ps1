$csvFile = "C:\Users\kevinm\Desktop\eco3_deemed_scores_2018.csv"

$csvData = import-csv $csvFile

$csvData | %{
#Get-Content $csvFile | %{
    $thisRow = $_
    $measure = $null
    $measureVariant = $null
    $lifetime = $null
    $costSaving = $null
    $annualSaving = $null
    $preHeatingSystem = $null
    $meanPopt = $null
    $wallType = $null
    $uValueDelta = $null
    $postHeatingSystem = $null
    $ageBand = $null

    $issMeasure = $null
    $issMeasureName = $null
    $issWallType = $null
    $issThermalConductivity = $null
    $issDoorType = $null
    $issGlazing = $null
    $issRiri = $null
    $issLoftInsulation = $null
    $issPreExistingHeatingControls = $null
    $issPropertyType = $null
    $issBedrooms = $null
    $issExtWalls = $null
    $issDetatchment = $null


    switch($thisRow.Measure_Category){
        ("Solid Wall Insulation") {
            $measure = $thisRow.Measure_Category
            $measureVariant = $thisRow.Measure_Type.Split("_")[0]
            $lifetime = $thisRow.L
            $costSaving = $thisRow.'Cost_Score_(�)'
            $annualSaving = $thisRow.'Annual Saving (�)'
            $preHeatingSystem = $thisRow.Pre_Main_Heating_Source_for_the_Property
            $meanPopt = $thisRow.'average POPT factor'

            $wallType = $thisRow.Measure_Type.Split("_")[1]
            $uValueDelta = $thisRow.Measure_Type.Split("_")[2] + " -> " + $thisRow.Measure_Type.Split("_")[3]
            $postHeatingSystem = $null

            $issMeasure = "Solid Wall Insulation"
            $issMeasureName = $thisRow.Measure_Type.Split("_")[0]+"_"+$thisRow.Measure_Type.Split("_")[1]
            $issWallType = "Please specify"
            }
        ("Cavity Wall Insulation"){
            $measure = $thisRow.Measure_Category
            $measureVariant = $thisRow.Measure_Type
            $lifetime = $thisRow.L
            $costSaving = $thisRow.'Cost_Score_(�)'
            $annualSaving = $thisRow.'Annual Saving (�)'
            $preHeatingSystem = $thisRow.Pre_Main_Heating_Source_for_the_Property
            $meanPopt = $thisRow.'average POPT factor'
            
            $wallType = $null
            $uValueDelta = $null
            $postHeatingSystem = $null

            $issWallType = "Cavity"
            $issMeasure = "Cavity Wall Insulation"
            $issMeasureName = $thisRow.Measure_Type.Split("_")[0]
            $issThermalConductivity = $thisRow.Measure_Type.Split("_")[1]
            if($issThermalConductivity = "Cavity"){$issThermalConductivity = $null}
            }
        ("Other Insulation"){
            $measure = $thisRow.Measure_Category
            $measureVariant = $thisRow.Measure_Type
            $lifetime = $thisRow.L
            $costSaving = $thisRow.'Cost_Score_(�)'
            $annualSaving = $thisRow.'Annual Saving (�)'
            $preHeatingSystem = $thisRow.Pre_Main_Heating_Source_for_the_Property
            $meanPopt = $thisRow.'average POPT factor'
            
            $wallType = $null
            $uValueDelta = $null
            $postHeatingSystem = $null

            $issMeasure = "Other Insulation"
            $issMeasureName = $thisRow.Measure_Type.Split("_")[0]
            switch($issMeasureName){
                "HPED" {
                    if($thisRow.Measure_Type.Split("_")[1] -eq "less"){$issDoorType = "60% or less"}
                    elseif($thisRow.Measure_Type.Split("_")[1] -eq "greater"){$issDoorType = "More than 60%"}
                    else{$issDoorType = $null}
                    }
                "WG" {
                    if($thisRow.Measure_Type.Split("_")[1] -eq "singletodouble"){$issGlazing = "Single to Double Glazing"}
                    elseif($thisRow.Measure_Type.Split("_")[1] -eq "improveddouble"){$issGlazing = "Improved Double Glazing"}
                    else{$issGlazing = $null}
                    }
                "RIRI" {
                    if($thisRow.Measure_Type.Split("_")[2] -eq "unin"){$issRiri = "No"}
                    elseif($thisRow.Measure_Type.Split("_")[2] -eq "in"){$issRiri = "Yes"}
                    else{$issRiri = $null}
                    }
                default{}
                }
            }
        ("Loft Insulation"){
            $measure = $thisRow.Measure_Category
            $measureVariant = $thisRow.Measure_Type.Split("_")[1]
            $lifetime = $thisRow.L
            $costSaving = $thisRow.'Cost_Score_(�)'
            $annualSaving = $thisRow.'Annual Saving (�)'
            $preHeatingSystem = $thisRow.Pre_Main_Heating_Source_for_the_Property
            $meanPopt = $thisRow.'average POPT factor'
            
            $wallType = $null
            $uValueDelta = $null
            $postHeatingSystem = $null
            
            $issMeasure = "Loft Insulation"
            $issMeasureName = "LI"
            if($measureVariant -eq "lessequal100"){$issLoftInsulation = "100mm or less"}
            elseif($measureVariant -eq "greater100"){$issLoftInsulation = "More than 100mm"}
            else{$issLoftInsulation = $null}
            }
        ("Boiler"){
            $measure = $thisRow.Measure_Category
            $measureVariant = $thisRow.Measure_Type
            $lifetime = $thisRow.L
            $costSaving = $thisRow.'Cost_Score_(�)'
            $annualSaving = $thisRow.'Annual Saving (�)'
            $preHeatingSystem = $thisRow.Pre_Main_Heating_Source_for_the_Property
            $meanPopt = $thisRow.'average POPT factor'
            
            $wallType = $thisRow.Measure_Type.Split("_")[2]
            if($wallType = "CH"){$wallType = $thisRow.Measure_Type.Split("_")[3]}
            $uValueDelta = $null
            $postHeatingSystem = $thisRow.Post_Main_Heating_Source_for_the_Property

            $issMeasure = "Boiler"
            $issMeasureName = "B_"+$thisRow.Measure_Type.Split("_")[1]
            if($issMeasureName -eq "First"){$issMeasureName = "B_FTCH"}
            if($wallType = "solid"){$issWallType = "Solid Wall"}
            if($wallType = "cavity"){$issWallType = "Cavity"}
            else{$issWallType = ""}
            if($thisRow.Measure_Type.Split("_")[3] -eq "nopreHCs"){$issPreExistingHeatingControls = "No"}
            elseif($thisRow.Measure_Type.Split("_")[3] -eq "preHCs"){$issPreExistingHeatingControls = "Yes"}
            else{$issPreExistingHeatingControls = $null}
            }
        ("Other Heating"){
            $measure = $thisRow.Measure_Category
            $measureVariant = $thisRow.Measure_Type
            $lifetime = $thisRow.L
            $costSaving = $thisRow.'Cost_Score_(�)'
            $annualSaving = $thisRow.'Annual Saving (�)'
            $preHeatingSystem = $thisRow.Pre_Main_Heating_Source_for_the_Property
            $meanPopt = $thisRow.'average POPT factor'
            
            $wallType = $thisRow.Measure_Type.Split("_")[2]
            $uValueDelta = $null
            $postHeatingSystem = $null

            $issMeasure = "Other Heating"
            $issMeasureName = "Heating_controls"
            if($wallType = "solid"){$issWallType = "Solid Wall"}
            if($wallType = "cavity"){$issWallType = "Cavity"}
            else{$issWallType = ""}
            }
        ("Micro-Generation"){
            $measure = $thisRow.Measure_Category
            $measureVariant = $thisRow.Measure_Type
            $lifetime = $thisRow.L
            $costSaving = $thisRow.'Cost_Score_(�)'
            $annualSaving = $thisRow.'Annual Saving (�)'
            $preHeatingSystem = $thisRow.Pre_Main_Heating_Source_for_the_Property
            $meanPopt = $thisRow.'average POPT factor'
            
            $wallType = $null
            $uValueDelta = $null
            $postHeatingSystem = $null

            $issMeasure = ""
            $issMeasureName = "Solar_PV"
            }
        ("ESH"){
            $measure = $thisRow.Measure_Category
            $measureVariant = $thisRow.Measure_Type
            $lifetime = $thisRow.L
            $costSaving = $thisRow.'Cost_Score_(�)'
            $annualSaving = $thisRow.'Annual Saving (�)'
            $preHeatingSystem = $thisRow.Pre_Main_Heating_Source_for_the_Property
            $meanPopt = $thisRow.'average POPT factor'
            
            $wallType = $thisRow.Measure_Type.Split("_")[3]
            $uValueDelta = $null
            $postHeatingSystem = $null

            $issMeasure = ""
            $issMeasureName = $thisRow.Measure_Type.Split("_")[1]
            if($wallType = "solid"){$issWallType = "Solid Wall"}
            if($wallType = "cavity"){$issWallType = "Cavity"}
            else{$issWallType = ""}
            }
        }

    switch($thisRow.Property_Type){
        {$_ -match "2W_Flat"}{
            $propertyType = "2 ext. Wall Flat"
            $bedrooms = $thisRow.Property_Type.Split("_")[2]
            $issPropertyType = "Flat"
            $issExtWalls = "2 or fewer"
            if($bedrooms -ge 3){$issBedrooms = "3 or more"}
            else{$issBedrooms = $bedrooms}
            }
        {$_ -match "3W_Flat"}{
            $propertyType = "3 ext. Wall Flat"
            $bedrooms = $thisRow.Property_Type.Split("_")[2]
            $issPropertyType = "Flat"
            $issExtWalls = "3 or more"
            if($bedrooms -ge 3){$issBedrooms = "3 or more"}
            else{$issBedrooms = $bedrooms}
            }
        {$_ -match "End-terrace"}{
            $propertyType = "End-terrace"
            $bedrooms = $thisRow.Property_Type.Split("_")[1]
            $issPropertyType = "House"
            $issDetatchment = "End-Terrace"
            if($bedrooms -ge 5){$issBedrooms = "5 or more"}
            else{$issBedrooms = $bedrooms}
            }
        {$_ -match "Mid-terrace"}{
            $propertyType = "Mid-terrace"
            $bedrooms = $thisRow.Property_Type.Split("_")[1]
            $issPropertyType = "House"
            $issDetatchment = "Mid-Terrace"
            if($bedrooms -ge 5){$issBedrooms = "5 or more"}
            else{$issBedrooms = $bedrooms}
            }
        {$_.SubString(0,4) -match "Semi"}{
            $propertyType = "Semi-detatched"
            $bedrooms = $thisRow.Property_Type.Split("_")[1]
            $issPropertyType = "House"
            $issDetatchment = "Semi-Detatched"
            if($bedrooms -le 2){$issBedrooms = "2 or fewer"}
            elseif($bedrooms -ge 5){$issBedrooms = "5 or more"}
            else{$issBedrooms = $bedrooms}
            }
        {$_.SubString(0,3) -match "Det"}{
            $propertyType = "Detatched"
            $bedrooms = $thisRow.Property_Type.Split("_")[1]
            $issPropertyType = "House"
            $issDetatchment = "Detatched"
            if($bedrooms -le 2){$issBedrooms = "2 or fewer"}
            elseif($bedrooms -ge 6){$issBedrooms = "6 or more"}
            else{$issBedrooms = $bedrooms}
            }
        {$_ -match "Bung_Semi"}{
            $propertyType = "Bungalow - semi detached & end terrace"
            $bedrooms = $thisRow.Property_Type.Split("_")[2]
            $issPropertyType = "House"
            }
        {$_ -match "Bung_Det"}{
            $propertyType = "Bungalow - detached"
            $bedrooms = $thisRow.Property_Type.Split("_")[2]
            }
        {$_ -match "Bung_Mid"}{
            $propertyType = "Bungalow - mid terrace"
            $bedrooms = $thisRow.Property_Type.Split("_")[2]
            }
        {$_ -match "2W_Maisonette"}{
            $propertyType = "Maisonette 2 ext. Wall"
            $bedrooms = $thisRow.Property_Type.Split("_")[2]
            $issPropertyType = "Maisonette"
            $issExtWalls = "2 or fewer"
            if($bedrooms -ge 3){$issBedrooms = "3 or more"}
            else{$issBedrooms = "2 or fewer"}
            }
        {$_ -match "3W_Maisonette"}{
            $propertyType = "Maisonette 3 ext. Wall"
            $bedrooms = $thisRow.Property_Type.Split("_")[2]
            $issPropertyType = "Maisonette"
            $issExtWalls = "3 or more"
            if($bedrooms -ge 3){$issBedrooms = "3 or more"}
            else{$issBedrooms = "2 or fewer"}
            }
        }

    switch($uValueDelta.Split(" ")[0]){
        ("2.0"){
            switch($wallType){
                ("system"){$ageBand = "Before 1976 (E&W) / Before 1976 (S)"}
                ("timber"){$ageBand = "Before 1949 (E&W) / Before 1949 (S)"}
                default {$ageBand = $null}
                }
            }
        ("1.7"){
            switch($wallType){
                {$_ -in "solid","cavity","stone"}{$ageBand = "Before 1976 (E&W) / Before 1976 (S)"}
                ("system"){$ageBand = "1967 - 1975 (E&W) / 1965 - 1975 (S)"}
                default {$ageBand = $null}
                }
            }
        ("1.0"){
            switch($wallType){
                {$_ -in "solid","cavity","stone"}{$ageBand = "1976 - 1982 (E&W) / 1976 - 1983 (S)"}
                ("system"){$ageBand = "1976 - 1982 (E&W) / 1976 - 1983 (S)"}
                ("timber"){$ageBand = "1950 - 1966 (E&W) / 1950 - 1964 (S)"}
                default {$ageBand = $null}
                }
            }
        ("0.6"){
            switch($wallType){
                {$_ -in "solid","cavity","stone","system"}{$ageBand = "1983 - 1995 (E&W) / 1984 - 1991 (S)"}
                ("cob"){$ageBand = "Pre 1996 (E&W) / Pre 1999 (S)"}
                ("timber"){$ageBand = "1967 - 1975 (E&W) / 1965 - 1975 (S)"}
                ("filled"){$ageBand = "Pre 1976 (E&W) / Pre 1976 (S)"}
                default {$ageBand = $null}
                }
            }
        ("0.45"){
            switch($wallType){
                {$_ -in "solid","cavity","stone","system","cob"}{$ageBand = "From 1996 (E&W) / From 1992 (S)"}
                {$_ -in "timber","filled"}{$ageBand = "from 1976 (E&W) / from 1976 (S)"}
                default {$ageBand = $null}
                }
            }
        }
    Add-Content -Value ",ECO3,$measure,$measureVariant,$propertyType,$bedrooms,$preHeatingSystem,$postHeatingSystem,$annualSaving,$costSaving,$uValueDelta,$wallType,$lifetime,$meanPopt,$ageBand,$issMeasure,$issMeasureName,$issWallType,$issThermalConductivity,$issDoorType,$issGlazing,$issRiri,$issLoftInsulation,$issPreExistingHeatingControls,$issPropertyType,$issBedrooms,$issExtWalls,$issDetatchment" -Path C:\Users\kevinm\Desktop\eco3_deemedScoresParsed3.csv
    }
