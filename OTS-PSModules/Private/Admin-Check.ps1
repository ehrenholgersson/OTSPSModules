function Admin-Check{
    param(
        [string]$ToRun
    )

    $isAdmin = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        $wshell = New-Object -ComObject Wscript.Shell
    if (!($isAdmin)){
        $Credential = New-Object System.Management.Automation.PSCredential "$((Get-ComputerInfo).CsDNSHostName)\Maintenence", (Get-SecCredentials "$env:ProgramData\OTS\data\1.dat")
        if (!($Credential.UserName.Contains("Maintenence"))){
            & $ToRun
        }
        return $false
    }
    return $true
}