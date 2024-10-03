foreach ($script in (Get-ChildItem -Path $PSScriptRoot -Depth 2 | ? {$_.FullName.Contains(".ps1")})){
    Write-Output "Export $($script.FullName) please?"
    . $script.FullName
}
Export-ModuleMember -Function OTS-OfficeRepair