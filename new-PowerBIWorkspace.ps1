
#Setting up:

#Username for permissions add in creation of internal Workspace - **must have PowerBI Pro license**
$365Admin = Get-Credential #recommend T1 account
connect-ToExo

#Connect to Graph
#for queries to Sharepoint Online and 365
$teamBotDetails = get-graphAppClientCredentials -appName TeamsBot
$tokenResponse = get-graphTokenResponse -aadAppCreds $teamBotDetails
#for PowerBI Workspace creation
$powerBIBotDetails = get-graphAppClientCredentials -appName PowerBIBot
$powerBIBottokenResponse = get-powerBITokenResponse -aadAppCreds $powerBIBotDetails -Verbose

#for querying
$powerBIBotAdminDetails = get-graphAppClientCredentials -appName PowerBIAdminBot 
$powerBIBotAdmintokenResponse = get-powerBITokenResponse -aadAppCreds $powerBIBotAdminDetails -Verbose


#Build ad-hoc Power BI Workspace (v2) and tie it via Graph Schema to 365 group object

$365groupUPN = "" #add upn here

#get a 365 group with UGextensions
$target365Group = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -filterUpn $365groupUPN 

#create v2 Workspace
$newWorkspace = new-powerBIWorkspace -tokenResponse $powerBIBottokenResponse -workspacename $($target365Group.displayName + "Workspace") -version v2 -Verbose
#update 365group object Graph extension data with Workspace GUID, helps us to map between Workspace and 365 group
set-graphGroupUGSyncSchemaExtensions -tokenResponse $tokenResponse -groupId $target365Group.id -powerBiWorkspaceId $newWorkspace.id -Verbose
 

#add yourself to the Workspace - the service principle for the PowerBIBot should be added as an Admin during the creation process
update-powerBIWorkspaceUserPermissions -tokenResponse $powerBIBottokenResponse -workspaceID $newWorkspace.id -userPrincipalName $365Admin.UserName -groupUserAccessRight Admin -PrincipalType User -Verbose 
#Note the above is a bit picky, you can add yourself to a Workspace via Power BI Admin > Workspaces > select Workspace > access



#optional: create a PowerBI Data Managers group to add *Admins* to
$newPowerBIManagerGroup = new-mailEnabledSecurityGroup -dgDisplayName $($target365Group.displayName + " - PowerBI Managers Subgroup") -description "Mail-enabled Security Group for $($target365Group.displayName) Power BI Managers" -ownersUpns "ITTeamAll@anthesisgroup.com" -fixedSuffix " - PowerBI Managers Subgroup" -blockExternalMail $true -hideFromGal $true ###need to fix group upn to match formatting of the subgroups - currently too many spaces
$newPowerBIManagerGroupAADObject = get-graphGroups -tokenResponse $tokenResponse -filterUpn $newPowerBIManagerGroup.PrimarySmtpAddress
set-graphGroupUGSyncSchemaExtensions -tokenResponse $tokenResponse -groupId $target365Group.id -powerBiManagerGroupId $newPowerBIManagerGroupAADObject.Id -Verbose

#Add Power BI data managers group as Admin level (change permission level as needed - ideally only IT in this group at the moment for Admins)
$target365Group = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -filterUpn $365groupUPN #re-get 365group after update above
add-userToPowerBIWorkspace -tokenResponse $powerBIBottokenResponse -workspaceID $newWorkspace.id -aadObjectId $target365Group.anthesisgroup_UGSync.powerBiManagerGroupId -groupUserAccessRight Admin -PrincipalType Group -Verbose















<#code for requests to be finished


#Set names for Graph queries
$serverRelativeSiteUrl = "https://anthesisllc.sharepoint.com/sites/TeamHub"
$listName = "Request PowerBI Workspace" 

#Select the request:
#Query 'Request PowerBI Workspace' list for any items with a status of "Waiting"
$requests = get-graphListItems -tokenResponse $tokenResponse -serverRelativeSiteUrl $serverRelativeSiteUrl -listName $listName -expandAllFields
if($requests){[array]$selectedRequests = $requests | select {$_.Fields._x0033_65_x0020_Group_x0020_disp},{$_.Fields._x0033_65_x0020_Group_x0020_UPN},{$_.Fields._x0033_65_x0020_Group_x0020_GUID} | Out-GridView -PassThru -Title "Highlight any requests to process and click OK"}


#for the email to the requestor
$smtpBotDetails = get-graphAppClientCredentials -appName SmtpBot
$smtpBottokenResponse = get-graphTokenResponse -aadAppCreds $smtpBotDetails


#Process the Request:
ForEach($request in $requests){

$thisRequest = $requests | Where-Object {$_.id -eq $request.id}

$target365Group = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -filterId $thisRequest.fields._x0033_65_x0020_Group_x0020_GUID

#create workspace based on classifcation
Switch($target365Group.anthesisgroup_UGSync.classification){
    
    "Internal"{
        Try{
        #create v2 (modern experience) Workspace
        $newWorkspace = new-powerBIWorkspace -tokenResponse $powerBIBottokenResponse -workspacename $($target365Group.displayName + "Workspace") -version v2 -Verbose $verbosePreference
        #Add site members group as Contribute level
        add-userToPowerBIWorkspace -tokenResponse $powerBIBottokenResponse -workspaceID $newWorkspace.id -aadObjectId $target365Group.anthesisgroup_UGSync.memberGroupId -groupUserAccessRight Contributor -PrincipalType Group -Verbose $verbosePreference
        #Add IT user as Admin to create initial App and amend settings to allow users to publish app (but not amend permissions)
        update-powerBIWorkspaceUserPermissions -tokenResponse $powerBIBottokenResponse -workspaceID $newWorkspace.id -userPrincipalName $365user.UserName -groupUserAccessRight Admin -PrincipalType User -Verbose $verbosePreference
        #Set unified group graph extension data for Workspace ID to help us articifically tie 365 group and Workspace together
        set-graphGroupUGSyncSchemaExtensions -tokenResponse $tokenResponse -groupId $target365Group.id -powerBiManagerGroupId $newWorkspace.id -Verbose $verbosePreference
        }
        Catch{
        $error[0]
        }
    }
    "External"{
        #create PowerBIManager AAD group
        $newPowerBIManagerGroup = new-mailEnabledSecurityGroup -dgDisplayName $($target365Group.displayName + " - PowerBI Managers Subgroup") -description "Mail-enabled Security Group for $($target365Group.displayName) Power BI Managers" -ownersUpns "ITTeamAll@anthesisgroup.com" -fixedSuffix " - PowerBI Managers Subgroup" -blockExternalMail $true -hideFromGal $true ###need to fix group upn to match formatting of the subgroups - currently too many spaces
        #create v2 (modern experience) Workspace
        $newWorkspace = new-powerBIWorkspace -tokenResponse $powerBIBottokenResponse -workspacename $($target365Group.displayName + " Workspace") -version v2 -Verbose
        #Set unified group graph extension data for Workspace ID to help us articifically tie 365 group and Workspace together
        set-graphGroupUGSyncSchemaExtensions -tokenResponse $tokenResponse -groupId $target365Group.id -powerBiWorkspaceId $newWorkspace.id -Verbose $verbosePreference
        set-graphGroupUGSyncSchemaExtensions -tokenResponse $tokenResponse -groupId $target365Group.id -powerBiManagerGroupId $newPowerBIManagerGroup.Id -Verbose $verbosePreference

        #Re-get group for graph info
        $target365Group = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -filterId $thisRequest.fields._x0033_65_x0020_Group_x0020_GUID

        #Add site data managers group as Members level 
        add-userToPowerBIWorkspace -tokenResponse $powerBIBottokenResponse -workspaceID $newWorkspace.id -aadObjectId $target365Group.anthesisgroup_UGSync.powerBiManagerGroupId -groupUserAccessRight Admin -PrincipalType Group -Verbose
    }
   
}
#Open the Workspace - If this is for an internal team, add an example dataset to publish the app initially. All users of Internal Workspaces with be Contributors to prevent external sharing (which we can't restrict), they can then update the app after its first published. Also allow contributors to update the app (manual only, urgh) https://docs.microsoft.com/en-us/power-bi/collaborate-share/service-create-the-new-workspaces#allow-contributors-to-update-the-app
Start-Process "https://app.powerbi.com/groups/$($newWorkspace.id)"


#Update request with Created status and add Workspace url
update-graphListItem -tokenResponse $tokenResponse -graphSiteId $serverRelativeSiteUrl -listName $listName -listitemId $thisRequest.id -fieldHash @{"PowerBI_x0020_Workspace_x0020_UR" = "https://app.powerbi.com/groups/$($newWorkspace.id)";"Status" = "Created"} -Verbose $verbosePreference

#Generate email and send to Data Managers of group
$powerBIManagersGroup = get-graphGroups -tokenResponse $tokenResponse -filterId $target365Group.anthesisgroup_UGSync.powerBiManagerGroupId

$body = "<HTML><BODY><p> Hi $($thisUser.FieldValues.Author.LookupValue),</p>
<p>Your new [PowerBI Workspace] is available for you now. This is a ‘New’ type of Workspace which will allow you to have greater control over your PowerBI data and delivery.</p>
<p> We have a PowerBI Sharing Guide available <a href=`"https://www.youtube.com/watch?v=qEV9qoup2mQ&t=658s`">Here</a> and a troubleshooting access guide you can send to clients <a href=`"https://www.youtube.com/watch?v=_SsccRkLLzU&list=LL`">Here</a>.</p>

<p>We have some additional guides available Below if you would like to apply more granular permissions or make other changes to the design of your Workspace:</p>
<UL><LI><a href=`"https://youtu.be/NfqOQtiRWLA?list=LL`">Setting Workspace logo and description</a></LI>
<LI><a href=`"https://www.youtube.com/watch?v=4YVVlCAph4s`">Setting PowerBI permissions on data</a></LI></UL>

<p> Love,</p>
<p>The PowerBI Robot.</p>"


send-graphMailMessage -tokenResponse $tokenResponseSmtp -fromUpn "Shared_Mailbox_-_IT_Team_GBR@anthesisgroup.com" -toAddresses $powerBIManagersGroup.mail -subject "New PowerBI Workspace - $($target365Group.displayName)" -bodyHtml $Body

}



    $result = invoke-powerBIGet -tokenResponse $powerBIBottokenResponse -powerBIQuery "admin/groups/$($workspaceID)/users" -Verbose

    #>