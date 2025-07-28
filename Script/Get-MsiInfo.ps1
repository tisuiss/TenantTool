param (
    [string]$msiPath
)

function Get-MsiProperty {
    param (
        [string]$Path,
        [string]$Property
    )
    $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
    $Database = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($Path, 0))
    $View = $Database.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $Database, @("SELECT `Value` FROM `Property` WHERE `Property`='$Property'"))
    $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
    $Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
    return $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
}

$props = @{}
foreach ($prop in "ProductName", "ProductVersion", "Manufacturer") {
    try {
        $props[$prop] = Get-MsiProperty -Path $msiPath -Property $prop
    } catch {
        $props[$prop] = ""
    }
}

$props | ConvertTo-Json
