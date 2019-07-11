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

#$lockedFile = formCaptureText -formTitle "Which file?" -formText "Copy-and-paste the full path and filename"
$lockedFile = "X:\Databases\ECO2\DATA\EM V2 BE.accdb"
Invoke-Command -ComputerName FS03 -ConfigurationName WhoHasGotItOpen -ScriptBlock {param([string]$lockedFile)\\fs03\WhoHasGotItOpen\WhoHasGotItOpen.ps1 "$lockedFile"} -ArgumentList $lockedFile

Write-Host "Press any key to continue ..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

#$lockedFile = "X:\Internal\Operations\Resources tracker v1.1.xlsm"