get-package | ? {$_.Name -eq "Cylance PROTECT"} | uninstall-package -Force
