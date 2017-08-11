function form-captureText([string]$formTitle, [string]$formText, $sizeX, $sizeY){
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
    if ($sizeX -eq $null -or $sizeX -eq ""){$sizeX=300}
    if ($sizeY -eq $null -or $sizeY -eq ""){$sizeY=300}

    $objForm = New-Object System.Windows.Forms.Form 
    $objForm.Text = $formTitle
    $objForm.Size = New-Object System.Drawing.Size($sizeX,$sizeY) 
    $objForm.StartPosition = "CenterScreen"

    $objForm.KeyPreview = $True
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
        {$script:capturedText = $objTextBox.Text;$objForm.Close()}})
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
        {$objForm.Close();$script:capturedText = $null}})


    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Size($sizeX/7,120)
    $OKButton.Size = New-Object System.Drawing.Size($sizeX/3,23)
    $OKButton.Text = "OK"
    $OKButton.Add_Click({$script:capturedText=$objTextBox.Text;$objForm.Close()})
    $objForm.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size($sizeX*4/7,120)
    $CancelButton.Size = New-Object System.Drawing.Size($sizeX/3,23)
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
function form-captureSelection([string]$formTitle, [string]$formText, [string[]]$choices, $sizeX, $sizeY){
    #Consider piping choices directly to: out-gridview -PassThru
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
    if ($sizeX -eq $null -or $sizeX -eq ""){$sizeX=300}
    if ($sizeY -eq $null -or $sizeY -eq ""){$sizeY=300}

    $objForm = New-Object System.Windows.Forms.Form 
    $objForm.Text = $formTitle
    $objForm.Size = New-Object System.Drawing.Size($sizeX,$sizeY) 
    $objForm.StartPosition = "CenterScreen"

    $objForm.KeyPreview = $True
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
        {$Script:capturedSelection = $objListBox.SelectedItem;$objForm.Close()}})
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
        {$objForm.Close()}})


    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Size(($sizeX*.11),($sizeY-75))
    $OKButton.Size = New-Object System.Drawing.Size(($sizeX/3),25)
    $OKButton.Text = "OK"
    $OKButton.Add_Click({$Script:capturedSelection = $objListBox.SelectedItem;$objForm.Close()})
    $objForm.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(($sizeX*.56),($sizeY-75))
    $CancelButton.Size = New-Object System.Drawing.Size(($sizeX/3),25)
    $CancelButton.Text = "Cancel"
    $CancelButton.Add_Click({$objForm.Close()})
    $objForm.Controls.Add($CancelButton)

    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel.Location = New-Object System.Drawing.Size(($sizeX*0.05),(20)) 
    $objLabel.Size = New-Object System.Drawing.Size(($sizeX*0.9),($sizeY*0.15))
    $objLabel.Height = $sizeY*0.2-75
    $objLabel.Text = $formText
    $objForm.Controls.Add($objLabel) 

    $objListBox = New-Object System.Windows.Forms.ListBox 
    $objListBox.Location = New-Object System.Drawing.Size(($sizeX*0.05),($sizeY*0.25-75)) 
    $objListBox.Size = New-Object System.Drawing.Size(($sizeX*0.9),($sizeY*0.75)) 
    $objListBox.Height = $sizeY*0.8-75
    foreach ($choice in $choices){
        [void] $objListBox.Items.Add($choice)
        }
    $objForm.Controls.Add($objListBox) 

    $objForm.Topmost = $True
    $objForm.Add_Shown({$objForm.Activate()})
    [void] $objForm.ShowDialog()

    $capturedSelection
    }


function form-confirmation($messageTitle,$messageBody){
    #Add-Type -AssemblyName PresentationCore,PresentationFramework
    $ButtonType = [System.Windows.MessageBoxButton]::OKCancel
    $MessageIcon = [System.Windows.MessageBoxImage]::Warning
     $result = [System.Windows.MessageBox]::Show($messageBody,$MessageTitle,$buttonType,$messageIcon)
    $result
    }