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