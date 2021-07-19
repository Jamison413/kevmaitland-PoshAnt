function add-registryValue(){
    [cmdletbinding()]
        <#
    .SYNOPSIS
    

    .DESCRIPTION
    

    .PARAMETER infile1
    

    .PARAMETER infile2
    

    .EXAMPLE
    #>

    param(
        [parameter(Mandatory=$true)]
            [String]$registryPath
        ,[parameter(Mandatory=$true)]
            [String]$registryKey
        ,[parameter(Mandatory=$true)]
            [String]$registryValue
        ,[parameter(Mandatory=$true)]
            [ValidateSet("String", "ExpandString", "Binary", "DWord", "MultiString", "QWord")] 
            [String]$registryType
        )

    $registryPath = $registryPath.Replace("Computer\","")
    $registryPath = $registryPath.Replace("HKEY_LOCAL_MACHINE","HKLM:")
    $registryPath = $registryPath.Replace("HKEY_CURRENT_USER","HKCU:")

    $registryPath -split "\\" | % { #Silently create any missing subkeys
        $thisStub = $thisStub += $_+"\"
        if((Test-Path $thisStub) -eq $false){
            Write-Verbose "Sliently creating [$($thisStub)]"
            New-Item -Path $thisStub | Out-Null
            }
        }

    try {$existingItem = Get-ItemProperty -Path $registryPath -name $registryKey -ErrorAction SilentlyContinue}
    catch{}
    
    Write-Verbose "`$registryPath = $registryPath"
    Write-Verbose "`$registryKey = $registryKey"
    Write-Verbose "`$registryValue = $registryValue"
    Write-Verbose "`$registryType = $registryType"
    Write-Verbose "`$existingItem = $existingItem"

    if([string]::IsNullOrWhiteSpace($existingItem.$registryKey)){
        Write-Verbose "Creating new Registry Value"
        New-ItemProperty -Path $registryPath -Name $registryKey -Value $registryValue -PropertyType $registryType
        }
    else{
        Write-Verbose "Updating existing Registry Value"
        Set-ItemProperty -Path $registryPath -Name $registryKey -Value $registryValue
        }
    
    
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

