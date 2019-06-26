$yourSecurityGroup = "$env:USERDNSDOMAIN\$env:USERNAME" 

function get-esetRemoteAgentProductCode(){
    $agentKey = Get-ChildItem -Recurse -Path HKLM:\SOFTWARE\Classes\Installer | ?{$_.Name -imatch "786A20824144DB1449FA500C3A98D88D"}
    $agentKey.Property
    }
function get-esetUuidFromProductCode($productCode){
    $uuid = $productCode.Substring(7,1)
    $uuid += $productCode.Substring(6,1)
    $uuid += $productCode.Substring(5,1)
    $uuid += $productCode.Substring(4,1)
    $uuid += $productCode.Substring(3,1)
    $uuid += $productCode.Substring(2,1)
    $uuid += $productCode.Substring(1,1)
    $uuid += $productCode.Substring(0,1)+"-"
    $uuid += $productCode.Substring(11,1)
    $uuid += $productCode.Substring(10,1)
    $uuid += $productCode.Substring(9,1)
    $uuid += $productCode.Substring(8,1)+"-"
    $uuid += $productCode.Substring(15,1)
    $uuid += $productCode.Substring(14,1)
    $uuid += $productCode.Substring(13,1)
    $uuid += $productCode.Substring(12,1)+"-"
    $uuid += $productCode.Substring(17,1)
    $uuid += $productCode.Substring(16,1)
    $uuid += $productCode.Substring(19,1)
    $uuid += $productCode.Substring(18,1)+"-"
    $uuid += $productCode.Substring(21,1)
    $uuid += $productCode.Substring(20,1)
    $uuid += $productCode.Substring(23,1)
    $uuid += $productCode.Substring(22,1)
    $uuid += $productCode.Substring(25,1)
    $uuid += $productCode.Substring(24,1)
    $uuid += $productCode.Substring(27,1)
    $uuid += $productCode.Substring(26,1)
    $uuid += $productCode.Substring(29,1)
    $uuid += $productCode.Substring(28,1)
    $uuid += $productCode.Substring(31,1)
    $uuid += $productCode.Substring(30,1)
    $uuid
    }
function remove-regLeafKeyMatchingString([Microsoft.Win32.RegistryKey]$branchKey,[string]$regexExpressionString,[boolean]$areYouSure){
    $branchKey | Get-ItemProperty | Get-Member -MemberType Properties | %{
        if($_.Name -inotmatch "^PS"-and $_.Name -imatch $regexExpressionString){
            if($areYouSure){Remove-ItemProperty -Path $branchKey.PSPath -Name $_.Name}
            else{Remove-ItemProperty -Path $branchKey.PSPath -Name $_.Name -WhatIf}
            }
        }
    }
function remove-regBranchKeyMatchingString([Microsoft.Win32.RegistryKey]$branchKey,[string]$regexExpressionString,[boolean]$areYouSure){
    if ($(split-path $branchKey -Leaf) -match $regexExpressionString){
        if($areYouSure){$branchKey | Remove-Item -Recurse -Force}
        else{$branchKey | Remove-Item -Recurse -Force -WhatIf}
        }
    }
function take-ownership {
	param(
		[String]$Folder
	)
	takeown.exe /A /F $Folder
	$CurrentACL = Get-Acl $Folder
	write-host ...Adding NT Authority\SYSTEM to $Folder -Fore Yellow
	$SystemACLPermission = "NT AUTHORITY\SYSTEM","FullControl","ContainerInherit,ObjectInherit","None","Allow"
	$SystemAccessRule = new-object System.Security.AccessControl.FileSystemAccessRule $SystemACLPermission
	$CurrentACL.AddAccessRule($SystemAccessRule)
	write-host ...Adding AdminGroup to $Folder -Fore Yellow
	$AdminACLPermission = "$env:USERDNSDOMAIN\$env:USERNAME","FullControl","ContainerInherit,ObjectInherit","None","Allow"
	$SystemAccessRule = new-object System.Security.AccessControl.FileSystemAccessRule $AdminACLPermission
	$CurrentACL.AddAccessRule($SystemAccessRule)
	Set-Acl -Path $Folder -AclObject $CurrentACL
}

$productCode = get-esetRemoteAgentProductCode
$uuid = get-esetUuidFromProductCode -productCode $productCode

iex -Command "c:\windows\system32\sc.exe delete EraAgentSvc"
iex -Command "c:\windows\system32\taskkill /im:eraagent.exe /f"

Take-Ownership -folder "$env:ProgramFiles\ESET\RemoteAdministrator\Agent"
gci "$env:ProgramFiles\ESET\RemoteAdministrator\Agent" | % {Take-Ownership -Folder $_.FullName}
Remove-Item -Path "$env:ProgramFiles\ESET\RemoteAdministrator\Agent" -Recurse -Force
Take-Ownership -folder "$env:ProgramData\ESET\RemoteAdministrator\Agent"
gci "$env:ProgramData\ESET\RemoteAdministrator\Agent" | % {Take-Ownership -Folder $_.FullName}
Remove-Item -Path "$env:ProgramData\ESET\RemoteAdministrator\Agent" -Recurse -Force 

Remove-Item -Path "HKLM:\SOFTWARE\ESET\RemoteAdministrator\Agent" -Recurse -Force
Try {
    Get-ChildItem -Path HKLM:\SOFTWARE -Recurse | %{
        if($_.Name -match $productCode -or $_.Property -match $productCode){[array]$productItems+=$_}
        if($_.Name -match $uuid -or $_.Property -match $uuid){[array]$uuidItems+=$_}
        }
    }
Catch{}


foreach ($relevantKey in $($uuidItems | ? {$_.Name -inotmatch $uuid})){
    remove-regLeafKeyMatchingString -branchKey $relevantKey -regexExpressionString $uuid -areYouSure $true
    }
foreach ($relevantKey in $($uuidItems | ? {$_.Name -imatch $uuid})){
    remove-regBranchKeyMatchingString -branchKey $relevantKey -regexExpressionString $uuid -areYouSure $true
    }
foreach ($relevantKey in $($productItems | ? {$_.Name -inotmatch $productCode})){
    remove-regLeafKeyMatchingString -branchKey $relevantKey -regexExpressionString $productCode -areYouSure $true
    }
foreach ($relevantKey in $($productItems | ? {$_.Name -imatch $productCode})){
    remove-regBranchKeyMatchingString -branchKey $relevantKey -regexExpressionString $productCode -areYouSure $true
    }