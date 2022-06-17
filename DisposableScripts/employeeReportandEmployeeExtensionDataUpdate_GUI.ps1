
function export-GBUserAccountReport(){



$userBotDetails = get-graphAppClientCredentials -appName UserBot
$tokenResponse = get-graphTokenResponse -aadAppCreds $userBotDetails

#get all licensed users, remove most shared mailboxes and bots
$usersarray = get-graphUsers -tokenResponse $tokenResponse -filterLicensedUsers:$true -filterUsageLocation GB -selectAllProperties:$true -Verbose

#filtering
$allgraphusers = remove-mailboxesandbots -usersarray $usersarray
$allgraphusers = $allgraphusers.Where({$_.anthesisgroup_employeeInfo.contractType  -ne "ServiceAccount"})

$allgraphusers | Export-Csv -Path  $env:USERPROFILE\Downloads\GBUserReport.csv
    $wshell = New-Object -ComObject Wscript.Shell
    $answer = $wshell.Popup("Csv report exported to $('\Downloads\GBUserReport.csv')")
}
function update-AnthesisEmployeeExtensionData(){
        


$userBotDetails = get-graphAppClientCredentials -appName UserBot
$tokenResponse = get-graphTokenResponse -aadAppCreds $userBotDetails


#get all licensed users, remove most shared mailboxes and bots
$usersarray = get-graphUsers -tokenResponse $tokenResponse -filterLicensedUsers:$true -filterUsageLocation GB -selectAllProperties:$true -Verbose

#filtering
$allgraphusers = remove-mailboxesandbots -usersarray $usersarray
$allgraphusers = $allgraphusers.Where({$_.anthesisgroup_employeeInfo.contractType  -ne "ServiceAccount"})



##################

#Update accounts with detail changes/missing information

##################

#####update Graph extension data (we use to filter on)



If($allgraphusers){


#select one email
$selectedEmail = @()
While(($selectedEmail | Measure-Object).Count -ne 1){
if($allgraphusers){[array]$selectedEmail = $allgraphusers | select {$_.userPrincipalName} | Out-GridView -PassThru -Title "Highlight ONE email to edit and click OK"}
}

#select one property
$extensionPropertyToUpdate = @()
While(($extensionPropertyToUpdate | Measure-Object).Count -ne 1){
[array]$extensionPropertyToUpdate = @("contractType","businessUnit","extensionType") | Out-GridView -PassThru -Title "Highlight ONE property to update to edit and click OK"
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Value to update with'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,120)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Value to update with:'
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,40)
$textBox.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox)

$form.Topmost = $true

$form.Add_Shown({$textBox.Select()})
$result = $form.ShowDialog()

if (($result -eq [System.Windows.Forms.DialogResult]::OK) -and (![string]::IsNullOrWhiteSpace($textBox.Text))){
    $newValueToUpdate = $textBox.Text
    $queryThatWillBeProcessed = "Update user $($selectedEmail[0].'$_.userPrincipalName')'s $($extensionPropertyToUpdate) to $($newValueToUpdate)?"

    $wshell = New-Object -ComObject Wscript.Shell
    $answer = $wshell.Popup("$($queryThatWillBeProcessed)",0," Confirm?",0x1)

    If($answer -eq 1){
    $userID = get-graphUsers -tokenResponse $tokenResponse -filterUpns $($selectedEmail[0].'$_.userPrincipalName') -selectAllProperties -Verbose

        If(($userID | Measure-Object).Count -eq 1){
        write-host "Updating user $($userEmail)"
        set-graphUser -tokenResponse $tokenResponse -userIdOrUpn $userID.Id -userEmployeeInfoExtensionHash @{"$($extensionPropertyToUpdate)" = "$($newValueToUpdate)"} -Verbose
        }
        Else{
            $wshell = New-Object -ComObject Wscript.Shell
            $wshell.Popup("Too many accounts found, please check email address",0," Confirm?",0x1)
            Exit
        }
    }
    Else{
    $wshell = New-Object -ComObject Wscript.Shell
    $answer = $wshell.Popup("Cancelled")
    }

    sleep -Seconds 5
        $userID = get-graphUsers -tokenResponse $tokenResponse -filterUpns $($selectedEmail[0].'$_.userPrincipalName') -selectAllProperties -Verbose
            If($userID.anthesisgroup_employeeInfo.$($extensionPropertyToUpdate) -eq $($newValueToUpdate)){
            #success
            $wshell = New-Object -ComObject Wscript.Shell
            $wshell.Popup("Success, $($selectedEmail[0].'$_.userPrincipalName')'s $($extensionPropertyToUpdate) updated to $($newValueToUpdate)")
            $selectedEmail = @()
            $extensionPropertyToUpdate = @()
            $textBox = @()
            }
            Else{
            $wshell = New-Object -ComObject Wscript.Shell
            $wshell.Popup("Failure: $($selectedEmail[0].'$_.userPrincipalName')'s $($extensionPropertyToUpdate) appears not to be updated to $($newValueToUpdate). Check values and authentication.")
            }


}

}
Else{
            $wshell = New-Object -ComObject Wscript.Shell
            $wshell.Popup("Failure: Graph not connected, check authentication.")

}

}

$selectedJob = @()
While(($selectedJob | Measure-Object).Count -ne 1){
[array]$selectedJob = @("Export GB User Report","Amend individual user account's employee extension data") | Out-GridView -PassThru -Title "Select a job to run"
}

Switch($selectedJob){
"Export GB User Report" {export-GBUserAccountReport}
"Amend individual user account's employee extension data" {update-AnthesisEmployeeExtensionData}
}