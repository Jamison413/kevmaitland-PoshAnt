$Logname = "C:\ScriptLogs" + "\sync-POPObjectives$(Get-Date -Format "yyMMdd").log" #Check this location before live
Start-Transcript -Path $Logname -Append
Write-Host "Script started:" (Get-date)

Import-Module _PNP_Library_SPO


$Admin = "kimblebot@anthesisgroup.com"
$AdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Downloads\kimblebot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin, $AdminPass

$exoCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin, $AdminPass
connect-ToExo -credential $exoCreds
Connect-AzureAD -credential $adminCreds
connect-toAAD -credential $adminCreds

function get-POPReviewPeriod(){
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
            [string]$submitdate
        )

$RawReviewPeriod = Switch($submitdate){
"01" {"November"}
"02" {"November"}
"03" {"November"}
"04" {"November"}
"05" {"May"}
"06" {"May"}
"07" {"May"}
"08" {"May"}
"09" {"May"}
"10" {"May"}
"11" {"November"}
"12" {"November"}
}
$RawReviewPeriod = $RawReviewPeriod + " " + "$(get-date -UFormat "%Y")"
$RawReviewPeriod
}


#Connect to the POP processing list
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/teams/People_Services_Team_All_365/" -Credentials $adminCreds
$unprocessedobjectives = Get-PnPListItem -List "POP Objectives Processing (UK)"

ForEach($objective in $unprocessedobjectives){

#Nulls cause issues, have placeholder text instead if empty on submitting form

If($objective.FieldValues.Status -eq "Open"){
    #Sort the Review Period First
    [string]$submitdate = get-date $($objective.FieldValues.Review_x0020_Period) -UFormat "%m"
    $ReviewPeriod = get-POPReviewPeriod -submitdate $submitdate
    #Create item in Live Objectives List
    Try{
    Write-Host "Adding new objective: $($objective.FieldValues.Objective)" -ForegroundColor Yellow
    #Create initial objective with "ignore" as the objective name, PowerApps will filter this out
    $newobjective = Add-PnPListItem -List "POP Objectives List (UK)" -Values @{"Employee_x0020_Name" = $($objective.FieldValues.Employee_x0020_Name.Email); "Line_x0020_Manager" = $($objective.FieldValues.Line_x0020_Manager.Email); "Review_x0020_Period" = $($ReviewPeriod); "Objective" = "ignore"; "ObjectiveDescription" = $($objective.FieldValues.ObjectiveDescription); "ManagerAssessment" = $($objective.FieldValues.ManagerAssessment); "EmployeeAssessment" = $($objective.FieldValues.ManagerAssessment); "Status" = "Open"; "Cluster" = $($objective.FieldValues.Cluster)}
    }
    Catch{
    $error
    Write-Host "Something went wrong creating a live objective: $($objective.FieldValues.Objective)" -ForegroundColor Red
    }
    #Set the permissions to just the manager and the employee, People Services will have full control anyway
    Try{
    $1 = Set-PnPListItemPermission -List "POP Objectives List (UK)" -Identity $($newobjective.Id) -Group "People Services Team (All) Owners" -AddRole "Full Control" -ClearExisting
    $2 = Set-PnPListItemPermission -List "POP Objectives List (UK)" -Identity $($newobjective.Id) -User $($objective.FieldValues.Employee_x0020_Name.Email) -AddRole Contribute
    $3 = Set-PnPListItemPermission -List "POP Objectives List (UK)" -Identity $($newobjective.Id) -User $($objective.FieldValues.Line_x0020_Manager.Email) -AddRole Contribute
    $27 = Set-PnPListItemPermission -List "POP Objectives List (UK)" -Identity $($newobjective.Id) -Group 1188 -AddRole "Full Control" #App Administrators for POP SPO group
    }
    Catch{
    $error
    Write-Host "Something went wrong adding permissions onto a live objective: $($objective.FieldValues.Objective)" -ForegroundColor Red
    }
    #Finally in the live list, change the objective name to make it available to the user - not "ignore"
    Try{
    $4 = Set-PnPListItem -List "POP Objectives List (UK)" -Identity $($newobjective.Id) -Values @{"Objective" = $($objective.FieldValues.Objective)}
    }
    Catch{
    $error
    Write-Host "Something went wrong finalising a live objective to make it available to the user: $($objective.FieldValues.Objective)" -ForegroundColor Red
    }
    #Finally in the processing list, if there is a corresponding newobjective.id, delete the objective to be processed
    If($($newobjective.Id)){
    $5 = Remove-PnPListItem -List "POP Objectives Processing (UK)" -Identity $($objective.Id) -Force
    }
}
Else{
    Try{
    #Create initial objective with "ignore" as the objective name, PowerApps will filter this out
    Write-Host "Adding complete objective: $($objective.FieldValues.Objective)" -ForegroundColor Yellow
    $completeobjective = Add-PnPListItem -List "POP Completed Objectives (UK)" -Values @{"Employee_x0020_Name" = $($objective.FieldValues.Employee_x0020_Name.Email); "Line_x0020_Manager" = $($objective.FieldValues.Line_x0020_Manager.Email); "Review_x0020_Period" = $($objective.FieldValues.Review_x0020_Period); "Objective" = "ignore"; "ObjectiveDescription" = $($objective.FieldValues.ObjectiveDescription); "ManagerAssessment" = $($objective.FieldValues.ManagerAssessment); "EmployeeAssessment" = $($objective.FieldValues.ManagerAssessment); "ManagerComments" = $($objective.FieldValues.ManagerComments); "EmployeeComments" = $($objective.FieldValues.EmployeeComments); "Status" = "Closed"; "Cluster" = $($objective.FieldValues.Cluster); "Complete_x0020_Date" = $($($objective.FieldValues.Created_x0020_Date).Split("T")[0])}
    }
    Catch{
    $error
    Write-Host "Something went wrong creating a complete objective: $($objective.FieldValues.Objective)" -ForegroundColor Red
    }
    #Set the permissions to just the manager and the employee, People Services will have full control anyway
    Try{
    $6 = Set-PnPListItemPermission -List "POP Completed Objectives (UK)" -Identity $($completeobjective.Id) -Group "People Services Team (All) Owners" -AddRole "Full Control" -ClearExisting
    $7 = Set-PnPListItemPermission -List "POP Completed Objectives (UK)" -Identity $($completeobjective.Id) -User $($objective.FieldValues.Employee_x0020_Name.Email) -AddRole Contribute
    $8 = Set-PnPListItemPermission -List "POP Completed Objectives (UK)" -Identity $($completeobjective.Id) -User $($objective.FieldValues.Line_x0020_Manager.Email) -AddRole Contribute
    $28 = Set-PnPListItemPermission -List "POP Completed Objectives (UK)" -Identity $($completeobjective.Id) -Group 1188 -AddRole "Full Control" #App Administrators for POP SPO group
    }
    Catch{
    $error
    Write-Host "Something went wrong adding permissions onto a live objective: $($objective.FieldValues.Objective)" -ForegroundColor Red
    }
    #Finally in the complete list, change the objective name to make it available to the user - not "ignore"
    Try{
    $9 = Set-PnPListItem -List "POP Completed Objectives (UK)" -Identity $($completeobjective.Id) -Values @{"Objective" = $($objective.FieldValues.Objective)}
    }
    Catch{
    }
    #Finally in the processing list, if there is a corresponding completeobjective.id, delete the objective to be processed
    If($($completeobjective.Id)){
    $10 = Remove-PnPListItem -List "POP Objectives Processing (UK)" -Identity $($objective.Id) -Force
    }


}

}






####################################
#POP cleaning for changing managers#
####################################


#Live Objectives
$allliveobjectives = Get-PnPListItem -List "POP Objectives List (UK)"
ForEach($liveobjective in $allliveobjectives){
$thisuserManager = ""
$thisuserManager = Get-AzureADUserManager -ObjectId "$($liveobjective.FieldValues.Employee_x0020_Name.Email)"
    #Check it matches the Manager on the Objective
    If($liveobjective.FieldValues.Line_x0020_Manager.Email -ne $thisuserManager.UserPrincipalName){
    Write-Host "Manager has changed from $($liveobjective.FieldValues.Line_x0020_Manager.Email)  ->  $($thisuserManager.UserPrincipalName). Updating live Objective: $($liveobjective.Id)" -ForegroundColor Yellow
    $13 = Set-PnPListItem -List "POP Objectives List (UK)" -Identity $liveobjective.Id -Values @{"Line_x0020_Manager" = $($thisuserManager.UserPrincipalName)}
    $14 = Set-PnPListItemPermission -List "POP Objectives List (UK)" -Identity $liveobjective.Id -Group "People Services Team (All) Owners" -AddRole "Full Control" -ClearExisting
    $15 = Set-PnPListItemPermission -List "POP Objectives List (UK)" -Identity $liveobjective.Id -User $($liveobjective.FieldValues.Employee_x0020_Name.Email) -AddRole Contribute
    $16 = Set-PnPListItemPermission -List "POP Objectives List (UK)" -Identity $liveobjective.Id -User $($thisuserManager.UserPrincipalName) -AddRole Contribute
    $17 = Set-PnPListItemPermission -List "POP Objectives List (UK)" -Identity $liveobjective.Id -User "emily.pressey@anthesisgroup.com" -RemoveRole "Full Control" 	
    $29 = Set-PnPListItemPermission -List "POP Objectives List (UK)" -Identity $($liveobjective.Id) -Group 1188 -AddRole "Full Control" #App Administrators for POP SPO group
    }
}

#Complete Objectives
$allcompleteobjectives = Get-PnPListItem -List "POP Completed Objectives (UK)"
ForEach($completeobjective in $allcompleteobjectives){
$thisuserManager = ""
$thisuserManager = Get-AzureADUserManager -ObjectId "$($completeobjective.FieldValues.Employee_x0020_Name.Email)"
    #Check it matches the Manager on the Objective
    If($completeobjective.FieldValues.Line_x0020_Manager.Email -ne $thisuserManager.UserPrincipalName){
    Write-Host "Manager has changed from $($completeobjective.FieldValues.Line_x0020_Manager.Email)  ->  $($thisuserManager.UserPrincipalName). Updating Complete Objective: $($completeobjective.Id)" -ForegroundColor Yellow
    $18 = Set-PnPListItem -List "POP Completed Objectives (UK)" -Identity $completeobjective.Id -Values @{"Line_x0020_Manager" = $($thisuserManager.UserPrincipalName)}    
    $19 = Set-PnPListItemPermission -List "POP Completed Objectives (UK)" -Identity $completeobjective.Id -Group "People Services Team (All) Owners" -AddRole "Full Control" -ClearExisting
    $20 = Set-PnPListItemPermission -List "POP Completed Objectives (UK)" -Identity $completeobjective.Id -User $($completeobjective.FieldValues.Employee_x0020_Name.Email) -AddRole Contribute
    $21 = Set-PnPListItemPermission -List "POP Completed Objectives (UK)" -Identity $completeobjective.Id -User $($thisuserManager.UserPrincipalName) -AddRole Contribute
    $22 = Set-PnPListItemPermission -List "POP Completed Objectives (UK)" -Identity $completeobjective.Id -User "emily.pressey@anthesisgroup.com" -RemoveRole "Full Control" 	
    $30 = Set-PnPListItemPermission -List "POP Completed Objectives (UK)" -Identity $($completeobjective.Id) -Group 1188 -AddRole "Full Control" #App Administrators for POP SPO group

    }
}

####################################
#POP cleaning for Leavers          #
####################################

$allliveobjectives = Get-PnPListItem -List "POP Objectives List (UK)"
$allcompleteobjectives = Get-PnPListItem -List "POP Completed Objectives (UK)"

#Live List
$deactivatedUserObjectives = $allliveobjectives.where({$_.FieldValues.Employee_x0020_Name.LookupValue -match "Ω_"})
If($deactivatedUserObjectives){

    ForEach($deactivatedUserObjective in $deactivatedUserObjectives){
    #Move them to the archive list where only People Services has access
    $23 = Add-PnPListItem -List "POP Archive (UK)" -Values @{

    "Employee_x0020_Name" = $deactivatedUserObjective.FieldValues.Employee_x0020_Name.Email;
    "Line_x0020_Manager" = $deactivatedUserObjective.FieldValues.Line_x0020_Manager.Email;
    "Review_x0020_Period" = $deactivatedUserObjective.FieldValues.Review_x0020_Period;                                                                                                                                                                                                                                                                                                           
    "Objective" = $deactivatedUserObjective.FieldValues.Objective;                                                                                                       
    "ObjectiveDescription" = $deactivatedUserObjective.FieldValues.ObjectiveDescription;
    "ManagerAssessment" = $deactivatedUserObjective.FieldValues.ManagerAssessment;                                                                                                                               
    "EmployeeAssessment" = $deactivatedUserObjective.FieldValues.EmployeeAssessment;                                                                                                                               
    "Status" = $deactivatedUserObjective.FieldValues.Status;                                                                                                                                              
    "Cluster" = $deactivatedUserObjective.FieldValues.Cluster;                                                                                                                                      
    "ManagerComments" = $deactivatedUserObjective.FieldValues.ManagerComments;                                                                                                                                                             
    "EmployeeComments" = $deactivatedUserObjective.FieldValues.EmployeeComments;
        }
    $24 = Remove-PnPListItem -List "POP Objectives List (UK)" -Identity $deactivatedUserObjective.Id -Force
    }
}

$deactivatedUserObjectives = ""
#Complete List
$deactivatedUserObjectives = $allcompleteobjectives.where({$_.FieldValues.Employee_x0020_Name.LookupValue -match "Ω_"})
If($deactivatedUserObjectives){

    ForEach($deactivatedUserObjective in $deactivatedUserObjectives){
    #Move them to the archive list where only People Services has access
    $25 = Add-PnPListItem -List "POP Archive (UK)" -Values @{

    "Employee_x0020_Name" = $deactivatedUserObjective.FieldValues.Employee_x0020_Name.Email;
    "Line_x0020_Manager" = $deactivatedUserObjective.FieldValues.Line_x0020_Manager.Email;
    "Review_x0020_Period" = $deactivatedUserObjective.FieldValues.Review_x0020_Period;                                                                                                                                                                                                                                                                                                           
    "Objective" = $deactivatedUserObjective.FieldValues.Objective;                                                                                                       
    "ObjectiveDescription" = $deactivatedUserObjective.FieldValues.ObjectiveDescription;
    "ManagerAssessment" = $deactivatedUserObjective.FieldValues.ManagerAssessment;                                                                                                                               
    "EmployeeAssessment" = $deactivatedUserObjective.FieldValues.EmployeeAssessment;                                                                                                                               
    "Status" = $deactivatedUserObjective.FieldValues.Status;                                                                                                                                              
    "Cluster" = $deactivatedUserObjective.FieldValues.Cluster;                                                                                                                                      
    "ManagerComments" = $deactivatedUserObjective.FieldValues.ManagerComments;                                                                                                                                                             
    "EmployeeComments" = $deactivatedUserObjective.FieldValues.EmployeeComments;
    "Complete_x0020_Date" =  $deactivatedUserObjective.FieldValues.EmployeeComments;
        }
    
    $26 = Remove-PnPListItem -List "POP Completed Objectives (UK)" -Identity $deactivatedUserObjective.Id -Force
    }
}

Stop-Transcript
