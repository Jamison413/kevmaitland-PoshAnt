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