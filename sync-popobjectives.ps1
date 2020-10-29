$Logname = "C:\ScriptLogs" + "\sync-POPObjectives$(Get-Date -Format "yyMMdd").log" #Check this location before live
Start-Transcript -Path $Logname -Append
Write-Host "Script started:" (Get-date)


Import-Module _PNP_Library_SPO


$sharePointAdmin = "kimblebot@anthesisgroup.com"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content "$env:USERPROFILE\Desktop\KimbleBot.txt") 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass

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
    Set-PnPListItemPermission -List "POP Objectives List (UK)" -Identity $($newobjective.Id) -User $($objective.FieldValues.Employee_x0020_Name.Email) -AddRole Contribute
    Set-PnPListItemPermission -List "POP Objectives List (UK)" -Identity $($newobjective.Id) -User $($objective.FieldValues.Line_x0020_Manager.Email) -AddRole Contribute
    }
    Catch{
    $error
    Write-Host "Something went wrong adding permissions onto a live objective: $($objective.FieldValues.Objective)" -ForegroundColor Red
    }
    #Finally in the live list, change the objective name to make it available to the user - not "ignore"
    Try{
    Set-PnPListItem -List "POP Objectives List (UK)" -Identity $($newobjective.Id) -Values @{"Objective" = $($objective.FieldValues.Objective)}
    }
    Catch{
    $error
    Write-Host "Something went wrong finalising a live objective to make it available to the user: $($objective.FieldValues.Objective)" -ForegroundColor Red
    }
    #Finally in the processing list, if there is a corresponding newobjective.id, delete the objective to be processed
    If($($newobjective.Id)){
    Remove-PnPListItem -List "POP Objectives Processing (UK)" -Identity $($objective.Id) -Force
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
    Set-PnPListItemPermission -List "POP Completed Objectives (UK)" -Identity $($completeobjective.Id) -User $($objective.FieldValues.Employee_x0020_Name.Email) -AddRole Contribute
    Set-PnPListItemPermission -List "POP Completed Objectives (UK)" -Identity $($completeobjective.Id) -User $($objective.FieldValues.Line_x0020_Manager.Email) -AddRole Contribute
    }
    Catch{
    $error
    Write-Host "Something went wrong adding permissions onto a live objective: $($objective.FieldValues.Objective)" -ForegroundColor Red
    }
    #Finally in the complete list, change the objective name to make it available to the user - not "ignore"
    Try{
    Set-PnPListItem -List "POP Completed Objectives (UK)" -Identity $($completeobjective.Id) -Values @{"Objective" = $($objective.FieldValues.Objective)}
    }
    Catch{
    }
    #Finally in the processing list, if there is a corresponding completeobjective.id, delete the objective to be processed
    If($($completeobjective.Id)){
    Remove-PnPListItem -List "POP Objectives Processing (UK)" -Identity $($objective.Id) -Force
    }


}

}

Stop-Transcript