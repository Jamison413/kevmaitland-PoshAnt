

$anthesisSpoAdminSite = 'https://anthesisllc-admin.sharepoint.com'
$anthesisSpoMySite = 'https://anthesisllc-my.sharepoint.com/' # This needs to be the mySite where the userdata lives.
$sustainSpMySite = 'http://my.sustain.co.uk/' # This needs to be the mySite where the userdata lives.
$anthesisSharePointAdmin = "kevin.maitland@anthesisgroup.com"
$anthesisSpCredentials = get-spoCredentials -sharePointAdminUsername $anthesisSharePointAdmin -sharePointAdminSecurePassword (Read-Host -AsSecureString -Prompt "Password for $anthesisSharePointAdmin")
$sustainSharePointAdmin = "sustainltd\administrator"
$sustainSpCredentials = Get-Credential -UserName $sustainSharePointAdmin -Message "Password for $sustainSharePointAdmin"



function connectTo-MsolMailToSetPhotos($credential){
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/?proxyMethod=RPS -Credential $credential -Authentication Basic -AllowRedirection 
    Import-PSSession $Session 
    }
function connectTo-MsolSpo($sharePointAdminSite,$sharePointAdminCredentials){
    Import-Module Microsoft.Online.Sharepoint.PowerShell
    Connect-SPOService -url $sharePointAdminSite -Credential $sharePointAdminCredentials
    }
function update-userPhoto($userSAM,$userPhotoPath){
    Set-UserPhoto -Identity $userSAM -PictureData ([System.IO.File]::ReadAllBytes($userPhotoPath)) -Confirm:$false
    }
function import-CsomModules(){
    Import-Module 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.UserProfiles.dll' #CSOM for SPO User Profiles
    Import-Module 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll' #CSOM for SharePoint Online
    }
function get-spoCredentials($sharePointAdminUsername, $sharePointAdminSecurePassword){
    $sharePointCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($sharePointAdminUsername, $sharePointAdminSecurePassword)
    $sharePointCredentials
    }
function new-csomContext($sharepointSite, $sharePointCredentials){
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($sharepointSite)
    $context.Credentials = $sharePointCredentials
    $context
    }
function new-csomPeopleManger($context){
    $peopleManager = New-Object Microsoft.SharePoint.Client.UserProfiles.PeopleManager($context)
    $peopleManager
    }
function get-csomSharePointAllUsers($context, $userSamArray){
    $people = new-csomPeopleManger -context $context
    $spProfileArray = ,@()
    foreach($userSAM in $userSamArray){
        $spProfileArray += get-csomSharePointSingleUserProfile -context $context -peopleManager $people -userSAM $userSAM
        }
    $spProfileArray
    }    

function update-spoProfileFromOnPremiseProfile($userSAM,$sustainSpCredentials,$anthesisSpCredentials){
    $sustainContext = new-csomContext -sharepointSite $sustainSpMySite -sharePointCredentials $sustainSpCredentials
    $sustainPeopleManager = new-csomPeopleManger -context $sustainContext
    $susProfile = get-csomSharePointSingleUserProfile -context $sustainContext -peopleManager $sustainPeopleManager -userSAM "sustainltd\$userSAM"    
    $anthesisContext = new-csomContext -sharepointSite $anthesisSpoAdminSite -sharePointCredentials $anthesisSpCredentials
    $anthesisPeopleManager = new-csomPeopleManger -context $anthesisContext
    $antProfile = get-csomSharePointSingleUserProfile -context $anthesisnContext -peopleManager $anthesisPeopleManager -userSAM "i:0#.f|membership|$userSAM@anthesisgroup.com"
    update-msolSharePointProfileFromAnotherProfile -sourceSpProfile $susProfile -destSpProfile $antProfile -destContext $anthesisContext -destPeopleManager $anthesisPeopleManager
    }


#$sustainContext = new-csomContext -sharepointSite "http://my.sustain.co.uk/" -sharePointCredentials $sustainSpCredentials
#$sustainPeopleManager = new-csomPeopleManger -context $sustainContext
#$sustainSpUserProfiles = get-csomSharePointAllUsers -context $sustainContext -userSamArray @("sustainltd\kevin.maitland","mary.short@anthesisgroup.com")
