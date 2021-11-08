$ComputerSystem = Get-WmiObject -ClassName Win32_ComputerSystem
$ComputerSystem.AutomaticManagedPagefile = $true
$ComputerSystem.Put()