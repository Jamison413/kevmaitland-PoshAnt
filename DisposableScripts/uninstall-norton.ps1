#(Get-WmiObject -Class Win32_Product -Filter "Name='Symantec Endpoint Protection'" -ComputerName . ).Uninstall()
get-package | ? {$_.Name -match "Symantec"} | uninstall-package -Force
