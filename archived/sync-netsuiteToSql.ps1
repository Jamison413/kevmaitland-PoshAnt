$netSuiteParams = get-netSuiteParameters -connectTo Sandbox
$allNClientsSand = get-netSuiteClientsFromNetSuite -netsuiteParameters $netSuiteParams
$customers = invoke-netsuiteRestMethod -requestType GET -url "$($netSuiteParams.uri)/customer$query" -netsuiteParameters $netSuiteParams #-Verbose 
$customers.Count

$netSuiteParamsProd = get-netSuiteParameters -connectTo Production
$allNClientsProd = get-netSuiteClientsFromNetSuite -netsuiteParameters $netSuiteParams
#get netsuiteparams for Prod (currenlt only valid for Sandbox)