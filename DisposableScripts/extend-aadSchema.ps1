$schemaBotDetails = import-encryptedCsv -pathToEncryptedCsv "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\SchemaBot.txt"
$tokenResponse = get-graphTokenResponse -grant_type device_code -aadAppCreds $(get-graphAppClientCredentials -appName SchemaBot)
$newSchemaDefinition = @{
    id = "anthesisgroup_employeeInfo"
    description = "Additional information about Anthesis employees"
    targetTypes = @('User')
    properties=@(
        @{
            name = "extensionType"
            type = "String"
            },
        @{
            "name" = "contractType"
            "type" = "String"
            },
        @{
            "name" = "employeeId"
            "type" = "String"
            }
,
        @{
            "name" = "businessUnit"
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
    id = "anthesisgroup_UGSync"
    description = "Properties to manange synchronsiation between Unified Groups"
    targetTypes = @("Group")
    properties=@(
        @{
            name = "extensionType"
            type = "String"
            },
        @{
            name = "dataManagerGroupId"
            type = "String"
            },
        @{
            "name" = "memberGroupId"
            "type" = "String"
            },
        @{
            "name" = "combinedGroupId"
            "type" = "String"
            },
        @{
            "name" = "sharedMailboxId"
            "type" = "String"
            },
        @{
            "name" = "masterMembershipList"
            "type" = "String"
            },
        @{
            "name" = "classification"
            "type" = "String"
            },
        @{
            "name" = "privacy"
            "type" = "String"
            }
        @{
            "name" = "deviceGroupId"
            "type" = "String"
            }
        @{
            "name" = "powerBiWorkspaceId"
            "type" = "String"
            }
        @{
            "name" = "powerBiManagerGroupId"
            "type" = "String"
            }
        )
    }
$schemaBotDetails = import-encryptedCsv -pathToEncryptedCsv "$env:USERPROFILE\Downloads\SchemaBot.txt"
$tokenResponse = get-graphTokenResponse -grant_type device_code -aadAppCreds $(get-graphAppClientCredentials -appName SchemaBot)
invoke-graphPatch -tokenResponse $tokenResponse -graphQuery "/schemaExtensions/anthesisgroup_UGSync" -graphBodyHashtable $newSchemaDefinition3 -Verbose
invoke-graphPost -tokenResponse $tokenResponse -graphQuery "/schemaExtensions" -graphBodyHashtable $newSchemaDefinition -Verbose
invoke-graphPost -tokenResponse $tokenResponse -graphQuery "/schemaExtensions" -graphBodyHashtable $newSchemaDefinition3 -Verbose
invoke-graphPatch -tokenResponse $tokenResponse -graphQuery "/schemaExtensions/anthesisgroup_UGSync" -graphBodyHashtable $newSchemaDefinition3 -Verbose
invoke-graphPatch -tokenResponse $tokenResponse -graphQuery "/schemaExtensions/anthesisgroup_trainingRecord" -graphBodyHashtable $newSchemaDefinition3 -Verbose
$schemaExtension = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/schemaExtensions?`$filter=id eq 'anthesisgroup_trainingRecord'" -Verbose
$schemaExtensions = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/schemaExtensions" -Verbose
$schemaExtension = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/schemaExtensions?`$filter=id eq 'anthesisgroup_employeeInfo'" -Verbose
invoke-graphPatch -tokenResponse $tokenResponse -graphQuery "/schemaExtensions/anthesisgroup_employeeInfo" -graphBodyHashtable $newSchemaDefinition -Verbose
