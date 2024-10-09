function Admin-Check{
    param(
        [string]$ToRun
    )
    try {
        $isAdmin = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
            # $wshell = New-Object -ComObject Wscript.Shell
        if (!($isAdmin)){
            $Credential = New-Object System.Management.Automation.PSCredential "$((Get-ComputerInfo).CsDNSHostName)\Maintenence", (Get-SecCredentials "$env:ProgramData\OTS\data\1.dat")
            # if (!($Credential.UserName.Contains("Maintenence"))){
                return (Start-Process "Powershell" -Credential $Credential -ArgumentList -PassThru "Start-Process 'Powershell' -Verb RunAs -ArgumentList $ToRun")
            # }
            # return $null # shouldn't be here
        }
        return (Start-Process 'Powershell' -Verb RunAs -PassThru -ArgumentList $ToRun)
    }
    catch{
        throw "Not able to run with admin credentials : $_"
    }
}