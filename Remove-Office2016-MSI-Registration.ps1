$Office = Get-ChildItem `
"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
"HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"

foreach ($Key in $Office) {

    $App = Get-ItemProperty $Key.PSPath -ErrorAction SilentlyContinue

    if ($App.DisplayName -like "Microsoft Office*2016*" -and
        $Key.PSChildName -match "^\{.*\}$") {

        Write-Host "Removing registry entry: $($App.DisplayName)"
        Remove-Item $Key.PSPath -Recurse -Force
    }
}
