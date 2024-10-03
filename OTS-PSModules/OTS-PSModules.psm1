foreach ($script in (Get-ChildItem -Path ".\" -Depth 2 | ? {$_.FullName.Contains(".ps1")})){
    . $script.FullName
}
Export-ModuleMember -Function OTS-OfficeRepair