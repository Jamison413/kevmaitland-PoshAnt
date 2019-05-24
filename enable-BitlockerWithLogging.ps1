<#
Script to enable Bitlocker on all non-removable drives for Windows 10 / PS5 PCs

Kev Maitland

v1.0
04/01/17
#>
[cmdletbinding()]
param(
    [parameter(Mandatory = $true)]
    [string]
    [ValidateSet("XtsAes256", "XtsAes128", "Aes256", "Aes128")]
    $encryptionCipher = "XtsAes256"
    ,[parameter(Mandatory = $true)]
    [bool]
    $requirePin
    ,[parameter(Mandatory = $false)]
    [string]
    $companyWebsite = "https://www.anthesisgroup.com"
    )

#region functions
function ensure-recoveryPasswordProtector(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
        [string]
        $mountPoint
        )
    $recoveryPasswordProtector = $(Get-BitLockerVolume $mountPoint).KeyProtector | ? {$_.KeyProtectorType -eq "RecoveryPassword"}
    if([string]::IsNullOrWhiteSpace($recoveryPasswordProtector)){
        write-eventLogEntry -Message "No RecoveryPasswordProtector set - adding new" -type Information
        $recoveryPasswordProtector = $(Add-BitLockerKeyProtector -MountPoint $mountPoint -RecoveryPasswordProtector).KeyProtector | ? {$_.KeyProtectorType -eq "RecoveryPassword"}
        }
    $recoveryPasswordProtector 
    }
function get-pinViaWindowsForm(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [string]$formTitle
        ,[parameter(Mandatory = $true)]
        [string]$formText
        ,[parameter(Mandatory = $false)]
        [string]$companyWebsite
        )
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

    $objForm = New-Object System.Windows.Forms.Form 
    $objForm.Text = $formTitle
    $objForm.Size = New-Object System.Drawing.Size(300,450) 
    $objForm.StartPosition = "CenterScreen"

    $objForm.KeyPreview = $True
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
        {$script:capturedText = $objTextBox.Text;$objForm.Close()}})
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
        {$objForm.Close();$script:capturedText = $null}})


    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel.Location = New-Object System.Drawing.Size(10,20) 
    $objLabel.Size = New-Object System.Drawing.Size(270,275) 
    $objLabel.Text = $formText
    $objForm.Controls.Add($objLabel) 

    $objTextBox = New-Object System.Windows.Forms.TextBox 
    $objTextBox.Location = New-Object System.Drawing.Size(10,300) 
    $objTextBox.Size = New-Object System.Drawing.Size(260,20) 
    $objForm.Controls.Add($objTextBox) 

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Size(75,350)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = "OK"
    $OKButton.Add_Click({$script:capturedText=$objTextBox.Text;$objForm.Close()})
    $objForm.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(150,350)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text = "Cancel"
    $CancelButton.Add_Click({$objForm.Close();$script:capturedText = $null})
    $objForm.Controls.Add($CancelButton)

    if($companyWebsite){
        $pictureBox = new-object Windows.Forms.PictureBox
        $pictureBox.ImageLocation = "$companyWebsite/favicon.ico"
        $pictureBox.Location = New-object System.Drawing.Size(($objForm.right-100),($objForm.bottom-440))
        $pictureBox.Size = New-object System.Drawing.Size(75,75)
        $objForm.controls.add($pictureBox)
        $pictureBox.BringToFront()
        }

    $objForm.Topmost = $True
    $objForm.Add_Shown({$objForm.Activate()})
    [void] $objForm.ShowDialog()

    $capturedText
    }
function invoke-bitlockerEncryption() {
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
        [string]
        $mountPoint
        ,[parameter(Mandatory = $true)]
        [string]
        [ValidateSet("XtsAes256", "XtsAes128", "Aes256", "Aes128")]
        $encryptionCipher = "XtsAes256"
        ,[parameter(Mandatory = $true)]
        [string]
        $requirePin
        ,[parameter(Mandatory = $false)]
        [string]
        $pin
        )

    # Test that TPM is present and ready
    try {
        if (test-tpmReady -eq $true) {write-eventLogEntry "TPM Present and Ready. Beginning encryption process"}
        else {write-eventLogEntry "TPM *NOT* Present and Ready - cannot encrypt" -type Error;return}
        }
    catch {throw write-eventLogEntry "Issue with TPM. Exiting script [$_]" -type error}

    # Encrypting drive
    try {
        write-eventLogEntry "Enabling Bitlocker with Recovery Password protector$(if($requirePin){", PIN" }) and method $encryptionCipher"
        #Get/Add RecoveryPasswordProtector
        $recoveryPasswordProtector = ensure-recoveryPasswordProtector -mountPoint $mountPoint
        Write-Verbose "`$recoveryPasswordProtector = [$recoveryPasswordProtector]"
        #Backup RecoveryPasswordProtector to AAD *before* encrypting
        try{
            Backup-BitLockerKeyProtector -MountPoint $mountPoint -KeyProtectorId $recoveryPasswordProtector.KeyProtectorId -ErrorAction Stop
            write-eventLogEntry "Recovery Password [$($recoveryPasswordProtector.RecoveryPassword)] backed up to AD" -type Information
            BackupToAAD-BitLockerKeyProtector -MountPoint $mountPoint -KeyProtectorId $recoveryPasswordProtector.KeyProtectorId -ErrorAction Stop
            write-eventLogEntry "Recovery Password [$($recoveryPasswordProtector.RecoveryPassword)] backed up to AAD" -type Information

            if($requirePin){
                try{
                    Enable-BitLocker -MountPoint $mountPoint -SkipHardwareTest -EncryptionMethod $encryptionCipher -TpmAndPinProtector $(ConvertTo-SecureString -AsPlainText -Force -String $pin) -ErrorAction Stop
                    write-eventLogEntry -Message "Bitlocker successfully enabled on drive [$mountPoint] using [$encryptionCipher] with TPM & PIN"
                    }
                catch{write-eventLogEntry "Error enabling Bitlocker on drive [$mountPoint] using [$encryptionCipher] with TPM & PIN`r`n`r`n$_" -type Error}
                }
            else{
                try{
                    Enable-BitLocker -MountPoint $mountPoint -SkipHardwareTest -EncryptionMethod $encryptionCipher -TpmProtector
                    write-eventLogEntry -Message "Bitlocker successfully enabled on drive [$mountPoint] using [$encryptionCipher] with TPM"
                    }
                catch{write-eventLogEntry "Error enabling Bitlocker on drive [$mountPoint] using [$encryptionCipher] with TPM`r`n`r`n$_" -type Error}
                }
            }
        catch{
            write-eventLogEntry "Error backing up Recovery Password [$($recoveryPasswordProtector.RecoveryPassword)] to AAD!`r`n`r`n$_" -type Error
            }
        }
    catch {throw write-eventLogEntry "Error enabling Bitlocker on [$mountPoint]. Exiting script" }
    }
function invoke-bitlockerDecryption() {
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
        [string]
        $mountPoint
         )
    try {
        write-eventLogEntry "Decrypting Bitlocker-enabled drive $mountPoint"
        Disable-BitLocker -MountPoint $mountPoint
        }
    catch {throw write-eventLogEntry "Error decrypting Bitlocker-enabled drive $mountPoint`r`n`r`n$_"}
    }
function validate-pin([string]$unvalidatedString){
    try{
        if(([long]$digitsOnly = $unvalidatedString) -and ($unvalidatedString.Length -ge 10)){$stringIsOkay = $true}
        else{$stringIsOkay = $false}
        }
    catch{
        #Error probably caused by casting a [string] as [long]
        $stringIsOkay = $false
        }
    $stringIsOkay
    }    
function test-tpmReady {
    # Returns true/false if TPM is ready
    $tpm = Get-Tpm
    if ($tpm.TpmReady -and $tpm.TpmPresent -eq $true) {
        return $true
    }
    else {
        return $false
    }
}
function write-eventLogEntry {
    [cmdletbinding()]
    param (
        [parameter(Mandatory = $true, Position = 0)]
        [String]
        $message,
        [parameter(Mandatory = $false, Position = 1)]
        [string]
        [ValidateSet("Information", "Error")]
        $type = "Information"
        )
    
    Write-Verbose $message
    # Specify Parameters
    $log_params = @{
        Logname   = "Application"
        Source    = "Intune Bitlocker Encryption Script"
        Entrytype = $type
        EventID   = $(
            if ($type -eq "Information") { write-output 500 }
            else { Write-Output 501 }
        )
        Message   = $message
        }
    Write-EventLog @log_params
    }
#endregion

#====================================================================================================
#                                           Initialize
#====================================================================================================
#region  Initialize

# Provision new source for Event log
New-EventLog -LogName Application -Source "Intune Bitlocker Encryption Script" -ErrorAction SilentlyContinue

#endregion  Initialize

foreach ($drive in Get-BitLockerVolume){
    $mountPoint = $($drive.MountPoint).Replace(':','')
    if ((Get-Volume $mountPoint).DriveType -eq "Fixed"){ #Process all fixed drives
        if(@("FullyEncrypted","EncryptionInProgress") -contains $drive.VolumeStatus){$volumeStatusValidity ="Valid"}else{$volumeStatusValidity ="Invalid"}
        if($drive.EncryptionMethod -eq $encryptionCipher){$encryptionMethodValidity = "Valid"}else{$encryptionMethodValidity = "Invalid"}
        if($requirePin){if(($drive.KeyProtector.KeyProtectorType -contains "TpmPin")){$pinValidity = "Valid";$pinSet=$true}else{$pinValidity = "Invalid";$pinSet=$false}}
        else{if(($drive.KeyProtector.KeyProtectorType -notcontains "TpmPin")){$pinValidity = "Valid";$pinSet=$false}else{$pinValidity = "Invalid";$pinSet=$true}}
        write-eventLogEntry -Message "Drive [$mountPoint] VolumeStatus = [$($drive.VolumeStatus)][$volumeStatusValidity]; EncryptionMethod=[$($drive.EncryptionMethod)][$encryptionMethodValidity]; PinRequired=[$requirePin] PinSet=[$pinSet][$pinValidity]" -type Information

        if($volumeStatusValidity -eq "Valid" -and $encryptionMethodValidity -eq "Valid" -and $pinValidity -eq "Valid"){
            write-eventLogEntry -Message "Bitlocker status for drive [$mountPoint] is valid - no further configuration required" -type Information
            #Backup the recovery key to AD (& AAD as this can be wiped by the user)
            $recoveryPasswordProtector = ensure-recoveryPasswordProtector -mountPoint $mountPoint
            Backup-BitLockerKeyProtector -MountPoint $mountPoint -KeyProtectorId $recoveryPasswordProtector.KeyProtectorId
            BackupToAAD-BitLockerKeyProtector -MountPoint $mountPoint -KeyProtectorId $recoveryPasswordProtector.KeyProtectorId
            continue
            }

        if($encryptionMethodValidity = "Invalid"){ # Decrypt OS drive
            try {
                invoke-bitlockerDecryption -mountPoint $mountPoint
                # Wait for decryption to finish 
                while((Get-BitLockerVolume).VolumeStatus -ne 'FullyDecrypted')
                    {Start-Sleep -Seconds 30} 
                write-eventLogEntry -Message "Drive [$mountPoint] has been fully decrypted"
                }
            catch {throw Write-EventLogEntry -Message "Error decrypting drive [$mountPoint]`r`n$_" -type error}            
            }

        if($requirePin){
            while (!$codesMatch){ #Make sure that the two passcodes match
                while (!$codeIsValid){ #Make sure that the passcode is valid
                    $pin1 = get-pinViaWindowsForm -formTitle "Create your Bitlocker code" -formText "`r`nWelcome to Bitlocker Setup`r`n`r`n`r`nThis is very important!`r`n`r`n`r`nPlease enter a code (numbers only) that will be used to decrypt your computer each time it starts.`r`n`r`nThe number needs to be 10+ digits long, so a memorable phone number (that isn't on a business card stuck to your laptop) is a good suggestion." -companyWebsite $companyWebsite
                    $codeIsValid = validate-pin -unvalidatedString $pin1
                    }
                $pin2 = get-pinViaWindowsForm -formTitle "Validate your Bitlocker code" -formText "`r`n`r`n`r`n`r`n`r`n`r`n`r`n`r`n`r`nPlease re-enter your code (still only numbers)" -companyWebsite $companyWebsite
                if ($pin1 -eq $pin2){$codesMatch = $true}
                else {
                    $codesMatch = $false
                    $codeIsValid = $false
                    }
                }
            }
        
        invoke-bitlockerEncryption -mountPoint $mountPoint -encryptionCipher $encryptionCipher -requirePin $requirePin -pin $pin1 -Verbose
        }
    }

