foreach ($script in (Get-ChildItem -Path $PSScriptRoot -Depth 2 | ? {$_.FullName.Contains(".ps1")})){
    Write-Output "Loading $script"
    . $script.FullName
}