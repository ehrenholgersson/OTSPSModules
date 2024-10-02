foreach ($script in (Get-ChildItem -Path $PSScriptRoot -Depth 2 | ? {$_.FullName.Contains(".ps1")})){
    . $_.FullName
}