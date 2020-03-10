##########################################################
#                                                        #
#                Notes on Graph and Examples             #
#                                                        #
#                                                        #
# - Microsoft Graph is a REST API                        #
# - Connection is made via pnponline with an access      #
#   token                                                #
# - Called via Get (query/find something) or Post        #
#   (do something)                                       #
# - Graph really likes ID's, try to drill down into      #
#   each part you are looking for and redefine the       #
#   origninal query to use ID's where possible           #
#                                                        #
#                                                        #
#                                                        #
#                                                        #
#                                                        #
#                                                        #
#                                                        #
##########################################################






#Get REST credentials (from already existing salted credential file)
$teamBotDetails = Import-Csv "$env:USERPROFILE\Desktop\teambotdetails.txt"
$resource = "https://graph.microsoft.com"
$tenantId = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.TenantId)
$clientId = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.ClientID)
$redirect = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.Redirect)
$secret   = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.Secret)

$ReqTokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    client_Id     = $clientID
    Client_Secret = $secret
    } 
$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody


#Connnect to sharepoint via REST, it works on access tokens with a 'flat' connection across the site, unlike PNPOnline cmdlets which rely on site context.
Connect-PnPOnline -AccessToken $tokenResponse.access_token


#Get the root Sharepoint site
$response = Invoke-RestMethod -Uri 'https://graph.microsoft.com/beta/sites?$select=siteCollection,webUrl&$filter=siteCollection/root%20ne%20null' -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Get
$response.value

#Get a Team Site - you can use this to get Site ID's which are needed to drill down into a team site
$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com:/teams/IT_Team_All_365" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Get
write-host "The ID for $($response.displayname) is $($response.id)"

#Connect to the team site directly and view it via 'Drive'
$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,1ead4d00-2e2e-421e-9176-15287b5c22ce,d21ddf81-fcef-4e36-94e6-edd6fb487a31/drives" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Get

#Get visible Document Libraries (there are hidden ones in Sharepoint)
$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,e43ccfa7-1258-4a83-a6a9-483577275b99,d21ddf81-fcef-4e36-94e6-edd6fb487a31/drives" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Get
$response.value

#Get a specific Document Library (I think?)
$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,e43ccfa7-1258-4a83-a6a9-483577275b99,d21ddf81-fcef-4e36-94e6-edd6fb487a31/drives/b!p8885FgSg0qmqUg1dydbmYHfHdLv_DZOlObt1vtIejFDr6vvuqdFTaTWzb63-TzY" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Get
$response

#See all lists within a team site using the Site ID, you can get List ID's this way
$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,e43ccfa7-1258-4a83-a6a9-483577275b99,d21ddf81-fcef-4e36-94e6-edd6fb487a31/lists" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Get
$response.value

#See all items within a list, you can get Folder ID's this way
$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,e43ccfa7-1258-4a83-a6a9-483577275b99,d21ddf81-fcef-4e36-94e6-edd6fb487a31/lists/efabaf43-a7ba-4d45-a4d6-cdbeb7f93cd8/items" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Get
$response.value

#Another way of seeing all items in a list (folders, files, and all)
$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,e43ccfa7-1258-4a83-a6a9-483577275b99,d21ddf81-fcef-4e36-94e6-edd6fb487a31/drives/b!p8885FgSg0qmqUg1dydbmYHfHdLv_DZOlObt1vtIejFDr6vvuqdFTaTWzb63-TzY/root/children" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Get

#Get all folders within a Document library
$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,e43ccfa7-1258-4a83-a6a9-483577275b99,d21ddf81-fcef-4e36-94e6-edd6fb487a31/drives/b!p8885FgSg0qmqUg1dydbmYHfHdLv_DZOlObt1vtIejFDr6vvuqdFTaTWzb63-TzY/items/01LLWAYUOPYQEQ4RYWDJGZ2MPQ6QENEXEN/children" -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Get
$response.value

#Create a folder within a subfolder. Specify folder parameters in JSON format and Post as UTF8 to reduce problems with characters.
$body = "{
    `"name`": `"$($folder.'Candidate Name')`",
    `"folder`": { },
    `"@microsoft.graph.conflictBehavior`": `"rename`"
}"
$body = [System.Text.Encoding]::UTF8.GetBytes($body)
$CandidateNameResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,e43ccfa7-1258-4a83-a6a9-483577275b99,d21ddf81-fcef-4e36-94e6-edd6fb487a31/drives/b!p8885FgSg0qmqUg1dydbmYHfHdLv_DZOlObt1vtIejFDr6vvuqdFTaTWzb63-TzY/items/01LLWAYUILOIXGORD4QBFYI6MMKVPW4HZI/children" -Body $body -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Post

#Copy a file - this example is from the IT site, copies an individual file into a folder
$body = "{
    `"parentReference`": {
        `"driveId`": `"b!AE2tHi4uHkKRdhUoe1wizoHfHdLv_DZOlObt1vtIejFDr6vvuqdFTaTWzb63-TzY`",
        `"id`": `"01V67YTVHO2Y3JHJUM35EZTMY3LNCLRVNR`"
         },
    `"name`": `"New Starter Checklist.xlsx`",
    `"@microsoft.graph.conflictBehavior`": `"rename`"
}"
$body = [System.Text.Encoding]::UTF8.GetBytes($body)
$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,1ead4d00-2e2e-421e-9176-15287b5c22ce,d21ddf81-fcef-4e36-94e6-edd6fb487a31/drives/b!AE2tHi4uHkKRdhUoe1wizoHfHdLv_DZOlObt1vtIejFDr6vvuqdFTaTWzb63-TzY/items/01V67YTVFEUVDEF5VSBRFLKMCSRJZZOCK6/copy" -Body $body -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Post
#More MS help: https://docs.microsoft.com/en-us/graph/api/driveitem-copy?view=graph-rest-1.0&tabs=http
