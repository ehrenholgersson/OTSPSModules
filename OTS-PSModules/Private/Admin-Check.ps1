function Admin-Check{
    param(
        [srting]$ToRun
    )

    $isAdmin = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        $wshell = New-Object -ComObject Wscript.Shell
    if (!($isAdmin)){
        $Credential = New-Object System.Management.Automation.PSCredential "$((Get-ComputerInfo).CsDNSHostName)\Maintenence", (Get-SecCredentials "$env:ProgramData\OTS\data\1.dat")
        if (!($Credential.UserName.Contains("Maintenence"))){
            Start-Process "Powershell" -Credential $Credential -ArgumentList "Start-Process 'Powershell' -Verb RunAs -ArgumentList $ToRun"

        }
        return $false
    }
    return $true
}