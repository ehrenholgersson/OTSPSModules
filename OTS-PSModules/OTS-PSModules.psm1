foreach ($script in (Get-ChildItem -Path $PSScriptRoot -Depth 2 | ? {$_.FullName.Contains(".ps1")})){
    . $script.FullName
}