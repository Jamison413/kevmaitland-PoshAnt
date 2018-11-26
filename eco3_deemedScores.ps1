$csvFile = "C:\Users\kevinm\Desktop\eco3_deemed_scores_2018.csv"

$csvData = import-csv $csvFile

$csvData | %{
#Get-Content $csvFile | %{
    $thisRow = $_
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
            }
        ("Cavity Wall Insulation"){
            $measure = $thisRow.Measure_Category
            $measureVariant = "Thermal conductivity " + $thisRow.Measure_Type.Split("_")[1]
            $lifetime = $thisRow.L
            $costSaving = $thisRow.'Cost_Score_(�)'
            $annualSaving = $thisRow.'Annual Saving (�)'
            $preHeatingSystem = $thisRow.Pre_Main_Heating_Source_for_the_Property
            $meanPopt = $thisRow.'average POPT factor'
            
            $wallType = $null
            $uValueDelta = $null
            $postHeatingSystem = $null
            }
        {$_ -in "Other Insulation","Loft Insulation"}{
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
            }
        ("Boiler"){
            switch($thisRow.Measure_Type.Split("_")[1]){
                {$_ -in"Broken","Repair","Upgrade"}{
                    $measure = $thisRow.Measure_Category
                    $measureVariant = $thisRow.Measure_Type
                    $lifetime = $thisRow.L
                    $costSaving = $thisRow.'Cost_Score_(�)'
                    $annualSaving = $thisRow.'Annual Saving (�)'
                    $preHeatingSystem = $thisRow.Pre_Main_Heating_Source_for_the_Property
                    $meanPopt = $thisRow.'average POPT factor'
            
                    $wallType = $thisRow.Measure_Type.Split("_")[2]
                    $uValueDelta = $null
                    $postHeatingSystem = $thisRow.Post_Main_Heating_Source_for_the_Property
                    }
                ("first"){
                    $measure = $thisRow.Measure_Category
                    $measureVariant = $thisRow.Measure_Type
                    $lifetime = $thisRow.L
                    $costSaving = $thisRow.'Cost_Score_(�)'
                    $annualSaving = $thisRow.'Annual Saving (�)'
                    $preHeatingSystem = $thisRow.Pre_Main_Heating_Source_for_the_Property
                    $meanPopt = $thisRow.'average POPT factor'
            
                    $wallType = $thisRow.Measure_Type.Split("_")[3]
                    $uValueDelta = $null
                    $postHeatingSystem = $thisRow.Post_Main_Heating_Source_for_the_Property
                    }
                }
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
            }
        }

    switch($thisRow.Property_Type){
        {$_ -match "2W_Flat"}{
            $propertyType = "3 ext. Wall Flat"
            $bedrooms = $thisRow.Property_Type.Split("_")[2]
            }
        {$_ -match "3W_Flat"}{
            $propertyType = "3 ext. Wall Flat"
            $bedrooms = $thisRow.Property_Type.Split("_")[2]
            }
        {$_ -match "End-terrace"}{
            $propertyType = "End-terrace"
            $bedrooms = $thisRow.Property_Type.Split("_")[1]
            }
        {$_ -match "Mid-terrace"}{
            $propertyType = "Mid-terrace"
            $bedrooms = $thisRow.Property_Type.Split("_")[1]
            }
        {$_.SubString(0,4) -match "Semi"}{
            $propertyType = "Semi-detatched"
            $bedrooms = $thisRow.Property_Type.Split("_")[1]
            }
        {$_.SubString(0,3) -match "Det"}{
            $propertyType = "Detatched"
            $bedrooms = $thisRow.Property_Type.Split("_")[1]
            }
        {$_ -match "Bung_Semi"}{
            $propertyType = "Bungalow - semi detached & end terrace"
            $bedrooms = $thisRow.Property_Type.Split("_")[2]
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
            }
        {$_ -match "3W_Maisonette"}{
            $propertyType = "Maisonette 3 ext. Wall"
            $bedrooms = $thisRow.Property_Type.Split("_")[2]
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
    Add-Content -Value ",ECO3,$measure,$measureVariant,$propertyType,$bedrooms,$preHeatingSystem,$postHeatingSystem,$annualSaving,$costSaving,$uValueDelta,$wallType,$lifetime,$meanPopt,$ageBand" -Path C:\Users\kevinm\Desktop\eco3_deemedScoresParsed.csv
    }
