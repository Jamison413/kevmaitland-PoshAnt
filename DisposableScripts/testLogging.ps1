Start-Transcript -Path C:\Transcript.log
 
$DebugPreference = 'Continue'
$VerbosePreference = 'Continue'
$InformationPreference = 'Continue'
 
Write-Output 'Output message'
Write-Warning -Message 'Warning message'
Write-Debug -Message 'Debug message'
Write-Verbose -Message 'Verbose message'
Write-Information -MessageData$oppd 'Information message'
Write-Host 'Host message'
Write-Error -Message 'Error message'
throw 'Throw exception'
 
Stop-Transcript #this line should never been reached


