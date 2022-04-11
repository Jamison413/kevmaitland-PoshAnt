try {Start-Process -FilePath .\NDP471-KB4033342-x86-x64-AllOS-ENU.exe -ArgumentList "-q -norestart" -Wait}
catch {throw "Failed to install .NET 4.7.1 update"}
try {Start-Process -FilePath .\vc_redist.x64.exe -ArgumentList "-q -norestart" -Wait}
catch{throw "Failed to install Visual C++ 2015"}
try {Start-Process -FilePath .\UninstallPreviousVersions.bat -Wait}
catch{throw "Failed to uninstall previous versions of BlueBeam"}

try{Start-Process -FilePath ".\Bluebeam Revu x64 20.msi" -ArgumentList 'TRANSFORMS=OpenLicense.mst /qn' -wait}
catch{throw "Failed to install BlueBeam ReVu 20"}
