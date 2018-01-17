$toDelete = @("AnthesisUKNew Starters","BD meeting-Europe","BF Mail","BFF Admin","BFF Analysts","BFF Board","BFF KAM","BFF Technical","Billing","Campaign-ECO2","Campaign-ECO2Private","Campaign-ECO2Social","Campaign-HeatingHealthCheck","Campaign-Warmth","CCE Group","Europe Delivery","European-SteeringGroup","Footprint Reporter Media","ICT Team","LRSBristol","LRSManchester","Oscar's e-mail","RESC BD","RESC BD+","RESC Community","sbsadmins","Technology Infrastructure","Test123","UK Software Team2","UK Strategy and Comms","UKCentral","UKSparke","US-Anthesis")

foreach($group in $toDelete){
    Write-Host $group -ForegroundColor Magenta
    try{
        $dg = Get-DistributionGroup $group
        if ($dg -ne $null){try{Remove-DistributionGroup -Identity $dg.Identity -Confirm:$false}catch{$_}}
        }
    catch{$_ | Out-Null}
    }

$oGroups = @("Gobstock 17","Gobstock 2017","Dec 15 Workshop presenters","Anthesis UMR","BD contacts","BSI UAE National Program development","Sharepoint")
foreach($group in $oGroups){
    Add-UnifiedGroupLinks -Identity $(Get-UnifiedGroup $group).Identity -Links "kevin.maitland" -LinkType Member
    }
