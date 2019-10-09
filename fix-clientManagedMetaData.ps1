
Connect-PnPOnline -Credentials $msolCredentials -Url "https://anthesisllc.sharepoint.com/clients/"
$kimble = Get-PnPTermGroup "Kimble"
$clients = Get-PnPTermSet -Identity "Clients" -TermGroup $kimble.id
$clientMMDTerms = Get-PnPTerm -TermSet $clients.id -TermGroup $kimble.id

$clientDocLibs = Get-PnPList
$clientDocLibs | % {Add-Member -InputObject $_ -MemberType NoteProperty -Name "Name" -Value $_.Title}

$comparisons = Compare-Object -ReferenceObject $clientMMDTerms -DifferenceObject $clientDocLibs -Property "Name" -IncludeEqual -PassThru
$equalComparisons = $comparisons | ? {$_.SideIndicator -eq "=="} 
$equalComparisons = $equalComparisons | ? {[string]::IsNullOrWhiteSpace($_.CustomProperties["DocLibId"])} 
$i=0
$max = $equalComparisons.Count
$equalComparisons | %{
    Write-Progress -Activity "Updating Managed Metadata custom properties" -Status "$i/$max" -PercentComplete ($i/$max)
    $thisTerm = $_
    $thisDocLib = $clientDocLibs | ? {$_.Name -eq $thisTerm.Name}
    if($thisDocLib.Count -eq 1){
        Write-Host -f Cyan "[$($thisTerm.Name)][$($thisTerm.Id.Guid)][$($thisDocLib.Id)]"
        $thisTerm.SetCustomProperty("DocLibId",$thisDocLib.Id)
        $thisTerm.Context.ExecuteQuery()
        }
    else{Write-Host -f Green "[$($thisTerm.Name)] matched [$($thisDocLib.Count)] DocLibs"}
    $i++
    }


<#
[Anthesis Consulting Group Ltd] matched [2] DocLibs
[Boots UK Ltd] matched [2] DocLibs
[British Telecommunications (BT) Plc] matched [2] DocLibs
[Coca Cola Enterprises Partnership] matched [2] DocLibs
[Connect Housing Association] matched [2] DocLibs
[Coopers Farm] matched [2] DocLibs
[Cradle to Cradle] matched [2] DocLibs
[Eco Integrate Limited] matched [2] DocLibs
[Electro Scientific Industries (esi)] matched [2] DocLibs
[Ener-G Switch2] matched [2] DocLibs
[Grain Craft Inc.] matched [2] DocLibs
[Grup Barcelonesa] matched [2] DocLibs
[HAVI Global Solutions Europe LTD] matched [2] DocLibs
[HSBC Bank Plc] matched [2] DocLibs
[J D Williams] matched [2] DocLibs
[LEMITOR Ochrona Srodowiska Sp. z o.o.] matched [2] DocLibs
[Linklaters LLP] matched [2] DocLibs
[Marylebone Cricket Club (MCC)] matched [2] DocLibs
[Moorfields Eye Hospital NHS Foundation Trust] matched [2] DocLibs
[Park Cake Bakery (Park Cakes Ltd)] matched [2] DocLibs
[Symphony Housing Group Ltd] matched [2] DocLibs
[ValPak Consulting Ltd] matched [2] DocLibs

#>