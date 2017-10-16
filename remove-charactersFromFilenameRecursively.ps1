$rootPath = "$env:USERPROFILE\Anthesis LLC"
$pattern = "[`\~#%&*{}/:<>?|`"``]"
$replacement = ""

function remove-duffCharactersRecursively($path){
    gci $path | % {
        #Write-Host -ForegroundColor Yellow $_.FullName
        if ($_.mode -match "^d"){remove-duffCharactersRecursively $_.FullName}
        if ($_.Name -match $pattern){
            Write-Host -ForegroundColor Yellow "$($_.Name) > $(($_.Name) -replace $pattern,$replacement)"
            Rename-Item $_.FullName -NewName $(($_.Name) -replace $pattern,$replacement) -WhatIf
            
            }
        }
    }

remove-duffCharactersRecursively -path $rootPath
