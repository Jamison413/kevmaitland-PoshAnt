#if(!(Get-PSSnapin | Where-Object {$_.name -eq "Microsoft.Exchange.Management.PowerShell.SnapIn"})) {
#    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
#    }

#foreach ($Mailbox in $(Get-Mailbox -ResultSize Unlimited)){
    #C:\Scripts\UpdateSustainersContactFolder.ps1 -user $Mailbox.Alias 
    #C:\Scripts\tidyOldContacts.ps1 -user $Mailbox.Alias
    #C:\Scripts\moveSustainersContactsToMainFolder.ps1 -user $Mailbox.Alias 
#    C:\Scripts\UpdateSustainersContactFolder_v2.0.ps1 -user $Mailbox.Alias 
#    }

#Get-GlobalAddressList | Update-GlobalAddressList
#Get-OfflineAddressBook | Update-OfflineAddressBook    

Import-Module ActiveDirectory
foreach ($adUser in Get-ADUser -Filter * -SearchBase "OU=Users,OU=Sustain,DC=Sustainltd,DC=local" -Properties mail){
    C:\Scripts\UpdateSustainersContactFolder_v3.0.ps1 -userEmailAddress $adUser.mail
    }