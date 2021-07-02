$toArchive = @("D:\Clients\SSE Energy Supply Limited\SSE ECO2 Delivery 2015-16","D:\Clients\Places for People\_Business Development\ECO CBR 2015","D:\Clients\Places for People\_Business Development\ECO CBR 2016","D:\Clients\Johnson Controls Building Efficiency UK\101474-RE_ECM_ESOS_Fenwicks","D:\Clients\Anchor Trust\_Business Development\ECO\ECO CBR\CBR Programme 2015","D:\Clients\Anchor Trust\_Business Development\ECO\ECO CBR\CBR Programme 2016","D:\Clients\British Gas Trading Limited\101233.001-BG ECO 2 - Communal Heating\Reports\Submissions\_Archive","D:\Clients\British Gas Trading Limited\101571-BG-ECO2-HHCRO\Reports\Submissions\_Archive","D:\Clients\Unite Integrated Solutions plc\101417-RE_EECM_Unite_EPC Delivery","D:\Clients\Housing and Care 21\101620-RE_ESE HC21 SAP C Options")
"D:\Clients\AmicusHorizon Group\_Business Development\ECO\ECO CBR 2016"
"D:\Clients\AmicusHorizon Group\_Business Development\ECO\Archive"
$dest= $source.Replace("D:\","E:\X\")
$logFile = "E:\RoboLogs\$(get-date -Format yyyy-MM-dd)_Archive.log"

foreach ($source in $toArchive){
    $result = iex -Command "robocopy `"$source`" `"$dest`" /e /copyall /MOVE /R:1 /W:1 /LOG+:$logFile /NP /TEE"
    if (Test-Path $source){
        #"It hasn't worked"
        [array]$failures += $source
        }
        else{[array]$successes += $source}
    }

D:\Clients\AmicusHorizon Group\_Business Development\ECO\ECO CBR 2016
New-Item -Path "D:\Clients\British Gas Trading Limited\101233.001-BG ECO 2 - Communal Heating\Reports\Submissions\_Archive" -ItemType Directory
New-Item -Path "D:\Clients\British Gas Trading Limited\101571-BG-ECO2-HHCRO\Reports\Submissions\_Archive" -ItemType Directory
New-Item -Path "D:\Clients\British Gas Trading Limited\101233.001-BG ECO 2 - Communal Heating\Reports\Submissions\_Archive" -ItemType Directory
