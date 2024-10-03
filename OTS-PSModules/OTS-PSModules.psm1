foreach ($script in (Get-ChildItem -Path $PSScriptRoot -Depth 3 | ? {$_.FullName.Contains(".ps1")})){
    . $script.FullName
}
foreach ($script in (Get-ChildItem -Path $PSScriptRoot -Depth 3 | ? {$_.FullName.Contains(".ps1") -and $_.FullName.Contains("Public")})){
    Export-ModuleMember  $script.Name.trim(".ps1")
}
