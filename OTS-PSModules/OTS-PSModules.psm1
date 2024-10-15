foreach ($script in (Get-ChildItem -Path $PSScriptRoot -Depth 3 | ? {$_.FullName.Contains(".ps1")})){
    . $script.FullName
}
foreach ($dll in (Get-ChildItem -Path $PSScriptRoot -Depth 3 | ? {$_.FullName.Contains(".dll") -and ($_.FullName.Contains("Public") -or $_.FullName.Contains("Private"))})){
    Import-Module  $dll.Name
}
foreach ($script in (Get-ChildItem -Path $PSScriptRoot -Depth 3 | ? {$_.FullName.Contains(".ps1") -and $_.FullName.Contains("Public")})){
    Export-ModuleMember  $script.Name.trim(".ps1")
}
