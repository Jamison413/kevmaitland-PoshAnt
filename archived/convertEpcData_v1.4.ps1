#****************************************************************************************************************
#Script to process EPC XML data, regardless of format, and output it in a variety of formats
#Manually strips out XML namespaces as different EPC providers use different namespaces for the same nodes.
#
#Kev Maitland
$versionNumber = 1.4
#02/07/14
#
#v1.1 Functions and tempates renamed to improve readablity, no functionality changes.
#04/11/16
#v1.2 Prompts user to select output template file (allows alternative output formats)
#11/11/16
#v1.3 Moves un/successful XMLs into separate folders 
#
#v1.4 Fixed use of $Global: variable usage (obsolete in PowerShell v5)
#20/01/17
#****************************************************************************************************************

#**********************These variables can be edited**********************
#$xmlDirectory = "C:\Users\kevind\Desktop\XMLConverter\xmlData"
$translateFrom = "epcXml"
#$translateTo = "elmhurst_code"
#**********************End of editable variables *************************

#**********************These variables should not be edited***************
$referenceDirectory = "X:\Internal\ECO\EPC_Processing\"
$failedSubfolderName = "Unprocessed"
$outputLogFile = "EPC_Processing_Results.log"
$xmlNameSpacesHashTable = @{HIP="DCLG-HIP"; SAP="DCLG-SAP"; SAP09="DCLG-SAP09"; CS="DCLG-HIP/CommonStructures"; REG="http://www.epcregister.com"}
$xmlNativeLanguage = "xmlTags" #This is the column header in the RosettaStone file to used to link the fromLanguage and toLanguage. Using "XML" allows for code re-use in the RosettaStone file, using "CSV" allows each CSV header to have its own unique lookup (but will require duplication of similar entries, e.g. Main-Fuel-Type, where multiple heating systems use the same lookup rules)
#$procedureTemplateFileName = "TechnicalMonitoringTemplate.csv"
$proceduresCsvExplicitXpathRow = 0 #The 1st row after the Headers in procedureTemplate.csv contains XPath queries declaring a (semi)explicit node structure (returns the correct values more reliably)
$proceduresCsvImplicitXpathRow = 1 #The 2nd row after the Headers in procedureTemplate.csv contains XPath queries declaring an implicit node structure (returns the values fractionally more quickly)
$proceduresCsvRulesRow = 2 #The 3rd row after the Headers in procedureTemplate.csv describes which rule/procedure should be applied (see the Switch statement in the main body of code for a full list of supported options)
$translationValuesFileName = "epcRosettaStone.csv"
$formTitle = "Sustain EPC XML conversion tool version $versionNumber"

$proceduresCsvXpathRowToUse = $proceduresCsvExplicitXpathRow #This decides whether the implict or explicit XML Node paths are used to locate the data. Implicit may be fractionally faster, but is more likely to return multiple nodes by mistake...
$proceduresCsvStaticRow = $proceduresCsvXpathRowToUse #This makes the "static" values come from the same row as the XPath query being executed. In theory, you might want different static values under different circumstances (but it's not been required to date)
#**********************End of non-editable variables *********************
cls

#region functions
function formatDirectoryPath([String]$dirtyPath) {
   if ($dirtyPath.EndsWith("\")) {$dirtyPath} else {$dirtyPath + "\"}
    }
function makeListOfHeaders([String]$fileName){
    (Get-Content ("$referenceDirectory"+"$fileName") -TotalCount 1).Split(",")
    }
function makeHashTableOfValues($listOfValues){
    $hashTable = @{}
    $i=0
    foreach ($value in $listOfValues){
        $hashTable[$value] = $i
        $i++
        }
    $hashTable
    }
function translateXmlValueTo([String]$xmlValue,[String]$fromLanguage,[String]$toLanguage, [String]$xmlNodeName, $translationCsv){
    $i = 0 #Rows
    $foundValidTranslation = $false
    while (($foundValidTranslation -eq $false) -and ($i -lt $translationCsv.Count)){ #Until we find an answer or get to the end of the file, go through each row in $translationCSV
        if ($translationCsv[$i].$xmlNativeLanguage -eq $xmlNodeName){ #Check whether the $CSV column value matches the xmlElementName that we're looking up. The $xmlNativeLanguage is a global variable
            if ($translationCsv[$i].$fromLanguage -eq $xmlValue){ #If so, check whether $fromLanguage value matches the $xmlValue that we're trying to translate
                $translationCsv[$i].$toLanguage #If so, return the translated $toLanguage
                $foundValidTranslation = $true #Then set $foundValidTranslation to $true to exit the translation function early
                }    
            }
        $i++
        }
    }
function sanitiseThatString([String]$dirtyString){
    $dirtyString.Replace("'","").Replace('"',"").Replace(",","")
    }
function formCaptureText([string]$formTitle, [string]$formText){
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

    $objForm = New-Object System.Windows.Forms.Form 
    $objForm.Text = $formTitle
    $objForm.Size = New-Object System.Drawing.Size(300,200) 
    $objForm.StartPosition = "CenterScreen"

    $objForm.KeyPreview = $True
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
        {$script:capturedText = $objTextBox.Text;$objForm.Close()}})
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
        {$objForm.Close();$script:capturedText = $null}})


    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Size(75,120)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = "OK"
    $OKButton.Add_Click({$script:capturedText=$objTextBox.Text;$objForm.Close()})
    $objForm.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(150,120)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text = "Cancel"
    $CancelButton.Add_Click({$objForm.Close();$script:capturedText = $null})
    $objForm.Controls.Add($CancelButton)

    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel.Location = New-Object System.Drawing.Size(10,20) 
    $objLabel.Size = New-Object System.Drawing.Size(280,40) 
    $objLabel.Text = $formText
    $objForm.Controls.Add($objLabel) 

    $objTextBox = New-Object System.Windows.Forms.TextBox 
    $objTextBox.Location = New-Object System.Drawing.Size(10,60) 
    $objTextBox.Size = New-Object System.Drawing.Size(260,20) 
    $objForm.Controls.Add($objTextBox) 

    $objForm.Topmost = $True

    $objForm.Add_Shown({$objForm.Activate()})
    [void] $objForm.ShowDialog()

    $capturedText
    }
function formCaptureSelection([string]$formTitle, [string]$formText, [string[]]$choices){
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

    $objForm = New-Object System.Windows.Forms.Form 
    $objForm.Text = $formTitle
    $objForm.Size = New-Object System.Drawing.Size(300,200) 
    $objForm.StartPosition = "CenterScreen"

    $objForm.KeyPreview = $True
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
        {$script:capturedSelection = $objTextBox.Text;$objForm.Close()}})
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
        {$objForm.Close();$script:capturedSelection = $null}})


    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Size(75,120)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = "OK"
    $OKButton.Add_Click({$script:capturedSelection=$objListBox.SelectedItem;$objForm.Close()})
    $objForm.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(150,120)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text = "Cancel"
    $CancelButton.Add_Click({$objForm.Close();$script:capturedSelection = $null})
    $objForm.Controls.Add($CancelButton)

    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel.Location = New-Object System.Drawing.Size(10,20) 
    $objLabel.Size = New-Object System.Drawing.Size(280,20) 
    $objLabel.Text = $formText
    $objForm.Controls.Add($objLabel) 

    $objListBox = New-Object System.Windows.Forms.ListBox 
    $objListBox.Location = New-Object System.Drawing.Size(10,40) 
    $objListBox.Size = New-Object System.Drawing.Size(260,20) 
    $objListBox.Height = 80
    foreach ($choice in $choices){
        [void] $objListBox.Items.Add($choice)
        }
    $objForm.Controls.Add($objListBox) 

    $objForm.Topmost = $True
    $objForm.Add_Shown({$objForm.Activate()})
    [void] $objForm.ShowDialog()

    $capturedSelection
    }
#endregion

$listOfTemplates = Get-ChildItem -Path $referenceDirectory -Filter "*.csv" | ?{@("epcRosettaStone.csv","NES RDSAP 2012 9.92 XML Specification.csv", "RdSAP_Survey_Template.csv", "procedureTemplate.csv") -notcontains $_.Name} | foreach {$_.Name}
$procedureTemplateFileName = formCaptureSelection -formTitle "Select output template" -formText "Choose the template that you would like to produce your output in" -choices $listOfTemplates

#Read in the CSV template and create a hashtable to make it easier to find the relevant columns in the output array
$proceduresCsv = Import-Csv ("$referenceDirectory"+"$procedureTemplateFileName")
$listOfProceduresCsvHeaders = makeListOfHeaders -fileName $procedureTemplateFileName
$proceduresCsvHeaderHashTable = makeHashTableOfValues -listOfValues $listOfProceduresCsvHeaders

#Read in the translation document
$translationCsv = Import-Csv ("$referenceDirectory"+"$translationValuesFileName")
$listOfTranslationLanguages = makeListOfHeaders -fileName $translationValuesFileName
$translationHeaderHashTable = makeHashTableOfValues -listOfValues $listOfTranslationLanguages

#Get the user to provide the paths and translation languages and tidy the data
$xmlDirectory = formCaptureText -formTitle $formTitle -formText "Please paste in the directory path to the XML files (e.g. X:\Clients\MyClient\1234-MyProject\Calcs\XML\)"
$translateTo = formCaptureSelection -formTitle $formTitle -formText "Please select the output language of the CSV file" -choices $listOfTranslationLanguages
$xmlDirectory = formatDirectoryPath -dirtyPath $xmlDirectory
$listOfXmlFiles = Get-ChildItem($xmlDirectory) | where-object {$_.Extension -like ".xml"} | Select-Object Name

#Make an array to store our processed data in
$processedDataArray = New-Object 'object[,]' $listOfProceduresCsvHeaders.Count,$listOfXmlFiles.Count

#Work through the list of EPC XML files
$i = 0 #Rows
$outcomes = @{}
foreach ($epcXmlFileName in $listOfXmlFiles){
    $outcomes.Add($epcXmlFileName.Name, "")
    #$now = Get-Date
    #Write-Host -ForegroundColor Yellow "Importing file "$epcXmlFileName.Name $now.Second $now.Millisecond
    #$now = Get-Date
    Write-Host -ForegroundColor DarkYellow "Processing file "$epcXmlFileName.Name #$now.Second $now.Millisecond
    try{
        $epcXml = [xml](Get-Content $xmlDirectory\$($epcXmlFileName.Name))
        foreach ($header in $listOfProceduresCsvHeaders){ #Work through each column in the output array
            #$now = Get-Date
            #Write-Host -ForegroundColor DarkYellow "Processing Header: $header "$now.Second $now.Millisecond
            if ((($header -eq "Ext1DateBuilt") -or ($header -eq "Ext2DateBuilt") -or ($header -eq "Ext3DateBuilt") -or ($header -eq "Ext4DateBuilt")) -and ((Select-Xml -Namespace $xmlNameSpacesHashTable -Xml $epcXml -XPath $proceduresCsv[$proceduresCsvXpathRowToUse].$header) -eq $null)) {break} #Try to exit early if there are no more extensions
            $j = $proceduresCsvHeaderHashTable[$header] #Columns
            switch ($proceduresCsv[$proceduresCsvRulesRow].$header){#Lookup to rule type for the current header
                'DoesThisValueExistInXml'                  {$processedDataArray[$j,$i] = (Select-Xml -Namespace $xmlNameSpacesHashTable -Xml $epcXml -XPath $proceduresCsv[$proceduresCsvXpathRowToUse].$header) -ne $null}
                'CountHowManyTimesThisNodeAppears'         {$processedDataArray[$j,$i] = (Select-Xml -Namespace $xmlNameSpacesHashTable -Xml $epcXml -XPath $proceduresCsv[$proceduresCsvXpathRowToUse].$header).Count}
                'JustLookUpValueInXml'                     {$processedDataArray[$j,$i] = (Select-Xml -Namespace $xmlNameSpacesHashTable -Xml $epcXml -XPath $proceduresCsv[$proceduresCsvXpathRowToUse].$header).Node.InnerText}
                'UseThisValueRegardless'                   {$processedDataArray[$j,$i] = $proceduresCsv[$proceduresCsvStaticRow].$header}
                'LookUpValueInXmlAndTranslateIt'           {$xPathQuery = $proceduresCsv[$proceduresCsvXpathRowToUse].$header
                        $xmlValue = (Select-Xml -Namespace $xmlNameSpacesHashTable -Xml $epcXml -XPath $xPathQuery).Node.InnerText
                        $xmlNodeNameWithNamespace = $xPathQuery.Split('"')[($xPathQuery.Split('"').length-2)] #Get the 2nd-to-last chunk of the XPath query (it contains the leaf node name, which contains the data we are about to translate)
                        $xmlNodeName = $xmlNodeNameWithNamespace.Split(":")[$xmlNodeNameWithNamespace.Split(":").length-1] #Get the last chunk of the Node name (this strips out any namespace information, which seems to vary between EPC XML standards)
                        $processedDataArray[$j,$i] = translateXmlValueTo -xmlValue $xmlValue -fromLanguage $translateFrom -toLanguage $translateTo -xmlNodeName $xmlNodeName -translationCsv $translationCsv
                        }
                default                                     {}
                }
            }
        $outcomes.Set_Item($epcXmlFileName.Name, "Succeeded")
        }
    catch{
        $outcomes.Set_Item($epcXmlFileName.Name, $_.Exception.Message)
        $somethingWentWrong = $true
        }
    $i++
    }


#Output the data to a CSV file
$outputFilename = "$translateTo`_output.csv"
Write-Host -ForegroundColor Yellow "Writing output file $xmlDirectory$outputFilename"
Get-Content ("$referenceDirectory"+"$procedureTemplateFileName") -TotalCount 1 | Out-File $xmlDirectory$outputFilename -Encoding utf8
$i = 0
while ($i -lt $listOfXmlFiles.Length){
    $j = 0
    #Write-Host "`$i = $i"
    while ($j -lt $listOfProceduresCsvHeaders.Length){
    #Write-Host "`$j = $j"
        if ($j -eq 0){$csvRow = sanitiseThatString -dirtyString $processedDataArray[$j,$i]}
        else {$csvRow += ","+(sanitiseThatString -dirtyString $processedDataArray[$j,$i])}
        $j++
        }
    $csvRow | Out-File $xmlDirectory$outputFilename -Append -Encoding utf8
    $i++
    }

Write-Host -ForegroundColor Yellow "Processing complete!"
Write-Host 
Write-Host 
if ($somethingWentWrong){
    Write-Host -ForegroundColor Red "These files had errors:"
    foreach($key in $outcomes.Keys){
        if ($outcomes.Get_Item($key) -ne "Succeeded"){Write-Host -ForegroundColor DarkRed "`t$key`t$($outcomes.Get_Item($key))"}
        }
    if ((formCaptureSelection -formTitle "Move failed XMLs?" -formText "Move all failed XMLs to a subfolder called $($failedSubfolderName)?" -choices @("No","Yes")) -eq "Yes"){
        if (!(Test-Path "$xmlDirectory\$failedSubfolderName")){New-Item -Path $xmlDirectory\$failedSubfolderName -ItemType Directory}
        foreach($key in $outcomes.Keys){
            if ($outcomes.Get_Item($key) -ne "Succeeded"){Move-Item -Path "$xmlDirectory\$key" -Destination $xmlDirectory\$failedSubfolderName\$key -Force -Confirm:$false}
            }
        }
    }

Add-Content -Value (Get-Date) -Path "$xmlDirectory\$outputLogFile"
foreach($key in $outcomes.Keys){Add-Content -Value "`t$key`t$($outcomes.Get_Item($key))" -Path "$xmlDirectory\$outputLogFile"}

Write-Host "Press any key to continue ..."

$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

#Select-Xml -Namespace $xmlNameSpacesHashTable -Xml $epcXml -XPath "//Report-Header/Property/HIP:Address/HIP:Postcode" 
#Select-Xml -Namespace $xmlNameSpacesHashTable -Xml $epcXml -XPath '//*[contains(name(), "Report-Header")]/*[contains(name(), "Property")]//*[contains(name(), "Postcode")]'
#Select-Xml -Namespace $xmlNameSpacesHashTable -Xml $epcXml -XPath '//*[contains(name(), "Report-Header")]/*[contains(name(), "Home-Inspector")]//*[(contains(name(), "Name")) and not(contains(name(), "-"))]'
#(Select-Xml -Namespace $xmlNameSpacesHashTable -Xml $epcXml -XPath '//*[contains(name(), "SAP-Data")]//*[contains(name(), "SAP-Building-Parts")]/*[contains(name(), "SAP-Building-Part")][1]/*[contains(name(), "SAP-Floor-Dimensions")]/*[contains(name(), "SAP-Floor-Dimension")]').COUNT
#Select-Xml -Namespace $xmlNameSpacesHashTable -Xml $epcXml -XPath '//*[contains(name(), "SAP-Data")]//*[contains(name(), "SAP-Flat-Details")]/*[contains(name(), "Level")]'
#Select-Xml -Namespace $xmlNameSpacesHashTable -Xml $epcXml -XPath '//*[contains(name(), "SAP-Data")]//*[contains(name(), "Door-Count") and (string-length(name()) < 20)]'
#Select-Xml -Namespace $xmlNameSpacesHashTable -Xml $epcXml -XPath '//*[contains(name(), "SAP-Data")]/*[contains(name(), "SAP-Property-Details")]/*[contains(name(), "SAP-Heating")]/*[contains(name(), "Main-Heating-Details")]/*[contains(name(), "Main-Heating")][1]/*[contains(name(), "Main-Heating-Category")]'
#Select-Xml -Namespace $xmlNameSpacesHashTable -Xml $epcXml -XPath  '//*[contains(name(), "SAP-Data")]/*[contains(name(), "SAP-Property-Details")]/*[contains(name(), "SAP-Heating")]/*[contains(name(), "WWHRS")]/*[contains(name(), "Rooms-With-Bath-And-Mixer-Shower")]'
#Select-Xml -Namespace $xmlNameSpacesHashTable -Xml $epcXml -XPath  '//*[contains(name(), "Energy-Assessment")]/*[contains(name(), "Property-Summary")]/*[contains(name(), "Has-Hot-Water-Cylinder")]'
#Select-Xml -Namespace $xmlNameSpacesHashTable -Xml $epcXml -XPath  '//*[contains(name(), "SAP-Data")]/*[contains(name(), "SAP-Property-Details")]/*[contains(name(), "SAP-Energy-Source")]/*[contains(name(), "Main-Gas")]'
#Select-Xml -Namespace $xmlNameSpacesHashTable -Xml $epcXml -XPath '//*[contains(name(), "SAP-Data")]//*[contains(name(), "SAP-Building-Parts")]/*[contains(name(), "SAP-Building-Part")][1]/*[contains(name(), "SAP-Floor-Dimensions")]/*[contains(name(), "SAP-Floor-Dimension")][1]/*[contains(name(), "Heat-Loss-Perimeter")]'
#Select-Xml -Namespace $xmlNameSpacesHashTable -Xml $epcXml -XPath '//*[contains(name(), "SAP-Data")]//*[contains(name(), "SAP-Building-Parts")]/*[contains(name(), "SAP-Building-Part")][1]/*[contains(name(), "SAP-Room-In-Roof")][1]/*[contains(name(), "Floor-Area")]'
#if(Select-Xml -Namespace $xmlNameSpacesHashTable -Xml $epcXml -XPath  '//*[contains(name(), "SAP-Data")]/*[contains(name(), "SAP-Property-Details")]/*[contains(name(), "SAP-Building-Parts")]/*[contains(name(), "SAP-Building-Part")][1]/*[contains(name(), "Wall-U-Value")]'){$true} else {$false}
#Select-Xml -Namespace $xmlNameSpacesHashTable -Xml $epcXml -XPath  '//*[contains(name(), "SAP-Data")]/*[contains(name(), "SAP-Property-Details")]/*[contains(name(), "SAP-Building-Parts")]/*[contains(name(), "SAP-Building-Part")][1]/*[contains(name(), "SAP-Alternative-Wall")]/*[contains(name(), "Sheltered-Wall")]'
#Select-Xml -Namespace $xmlNameSpacesHashTable -Xml $epcXml -XPath  '//*[contains(name(), "SAP-Data")]/*[contains(name(), "SAP-Property-Details")]/*[contains(name(), "SAP-Building-Parts")]/*[contains(name(), "SAP-Building-Part")][1]/*[contains(name(), "SAP-Room-In-Roof")][1]/*[contains(name(), "Roof-Room-Connected")]'
#Select-Xml -Namespace $xmlNameSpacesHashTable -Xml $epcXml -XPath  '//*[contains(name(), "SAP-Data")]/*[contains(name(), "SAP-Property-Details")]/*[contains(name(), "SAP-Building-Parts")]/*[contains(name(), "SAP-Building-Part")][1]/*[contains(name(), "Floor-Heat-Loss")]'




