function Admin-Check{
    param(
        [string]$Process,
        [string]$ArgList
    )
    try {
        $isAdmin = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
            # $wshell = New-Object -ComObject Wscript.Shell
        if (!($isAdmin)){
            $Credential = New-Object System.Management.Automation.PSCredential "$((Get-ComputerInfo).CsDNSHostName)\Maintenence", (Get-SecCredentials "$env:ProgramData\OTS\data\1.dat")
            # if (!($Credential.UserName.Contains("Maintenence"))){
                Start-Process "Powershell" -Credential $Credential -Wait -ArgumentList "Start-Process $process -Verb RunAs -Wait -ArgumentList $ArgList"
            # }
            # return $null # shouldn't be here
        }
        else {
            Start-Process $Process -Verb RunAs -Wait -ArgumentList $ArgList
        }
        return $true
    }
    catch{
        throw "Not able to run with admin credentials : $_"
    }
}