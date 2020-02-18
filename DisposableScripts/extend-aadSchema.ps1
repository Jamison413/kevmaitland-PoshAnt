$teamBotDetails = import-encryptedCsv -pathToEncryptedCsv "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\teambotdetails.txt"
$tokenResponse = get-graphTokenResponse -aadAppCreds $teamBotDetails

$newSchemaDefinition = @{
    id = "anthesisgroup_employeeInfo"
    description = "Additional information about Anthesis employees"
    targetTypes = @('User')
    properties=@(
        @{
            "name" = "contractType"
            "type" = "String"
            },
        @{
            "name" = "employeeId"
            "type" = "String"
            }
        )
    }
$newSchemaDefinition2 = @{
    id = "anthesisgroup_trainingRecord"
    description = "Employee training record data"
    targetTypes = @("User")
    properties=@(
        @{
            name = "ITUserTraining"
            type = "DateTime"
            }
        )
    }
$newSchemaDefinition3 = @{
    id = "anthesisgroup_trainingRecord"
    description = "Employee training record data"
    targetTypes = @("User")
    properties=@(
        @{
            name = "ITUserTraining"
            type = "DateTime"
            },
        @{
            "name" = "ITDataManagerTraining"
            "type" = "DateTime"
            }
        )
    }
invoke-graphPost -tokenResponse $tokenResponse -graphQuery "/schemaExtensions" -graphBodyHashtable $newSchemaDefinition -Verbose
invoke-graphPost -tokenResponse $tokenResponse -graphQuery "/schemaExtensions" -graphBodyHashtable $newSchemaDefinition2 -Verbose
invoke-graphPatch -tokenResponse $tokenResponse -graphQuery "/schemaExtensions/anthesisgroup_trainingRecord" -graphBodyHashtable $newSchemaDefinition3 -Verbose
$schemaExtension = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/schemaExtensions?`$filter=id eq 'anthesisgroup_trainingRecord'" -Verbose
$schemaExtensions = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/schemaExtensions" -Verbose

$teamBotDetails = import-encryptedCsv -pathToEncryptedCsv "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\teambotdetails.txt"
$teamBotToken = get-graphTokenResponse -aadAppCreds $teamBotDetails

$graphUsers = invoke-graphGet -tokenResponse $teamBotToken -graphQuery "/users?`$top=500" -Verbose
$graphUsersAll2 =  invoke-graphGet -tokenResponse $teamBotToken -graphQuery "/users?`$select=businessPhones,displayName,givenName,id,jobTitle,mail,mobilePhone,officeLocation,preferredLanguage,surname,userPrincipalName,Company,companyName,Department,creationType,directReports,manager,userType,anthesisgroup_employeeInfo" 
$graphUsers.value.Count
$graphUsers.'@odata.nextLink'

$graphUsersAll[10] | FL

$kev = $graphUsersAll |  ? {$_.DisplayName -match "Kev Maitland"}
$kev2 = $graphUsersAll2 |  ? {$_.DisplayName -match "Kev Maitland"}

invoke-graphPatch -tokenResponse $teamBotToken -graphQuery "/users/$($kev.id)" -graphBodyHashtable @{
    anthesisgroup_employeeInfo = @{
        contractType = "Employee"
        }
    } -Verbose

$contractors = convertTo-arrayOfEmailAddresses "Chris.Hazen@anthesisgroup.com
DeAnn.Sarver@anthesisgroup.com
Deby.Stabler@anthesisgroup.com
John.Hennessey@anthesisgroup.com
Leslie.Macdougall@anthesisgroup.com
Matt.Dion@anthesisgroup.com
Susan.Mazzarella@anthesisgroup.com
Therese.Karkowski@anthesisgroup.com
"
$contractors | % {
    $thisContractor = $_
    $graphUser = $graphUsersAll2 | ? {$_.userPrincipalName -eq $thisContractor}
    Write-Host $graphUser.displayName
    invoke-graphPatch -tokenResponse $teamBotToken -graphQuery "/users/$($graphUser.id)" -graphBodyHashtable @{
        anthesisgroup_employeeInfo = @{
            contractType = "Contractor"
            }
        } -Verbose    
    }