$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$desktop = [Environment]::GetFolderPath("Desktop")
$shell = New-Object -ComObject WScript.Shell

$shortcuts = @(
    @{ Name = "Avito Parser";            Launcher = "launch_avito_parser.bat";       Icon = "$env:SystemRoot\System32\SHELL32.dll,18" },
    @{ Name = "Yandex Maps Lead Parser"; Launcher = "launch_yandex_maps_parser.bat"; Icon = "$env:SystemRoot\System32\SHELL32.dll,220" },
    @{ Name = "HH.ru Parser";            Launcher = "launch_hh_parser.bat";          Icon = "$env:SystemRoot\System32\SHELL32.dll,265" },
    @{ Name = "Rusprofile Parser";       Launcher = "launch_rusprofile_parser.bat";  Icon = "$env:SystemRoot\System32\SHELL32.dll,21" }
)

foreach ($item in $shortcuts) {
    $shortcutPath = Join-Path $desktop ($item.Name + ".lnk")
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = Join-Path $scriptDir $item.Launcher
    $shortcut.WorkingDirectory = $scriptDir
    $shortcut.IconLocation = $item.Icon
    $shortcut.Save()
    Write-Output $shortcutPath
}
