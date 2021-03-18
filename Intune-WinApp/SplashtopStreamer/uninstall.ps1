get-package | ? {$_.Name -eq "Splashtop Streamer"} | uninstall-package -Force
get-package | ? {$_.TagId -eq "B7C5EA94-B96A-41F5-BE95-25D78B486678"} | uninstall-package -Force
