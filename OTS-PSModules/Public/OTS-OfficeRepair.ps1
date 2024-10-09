function OTS-OfficeRepair{

    # if (!(Admin-Check OTS-OfficeRepair)) {
    #     return
    # } # if we are not admin then rerun with correct credentials

    $isAdmin = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if (!$isAdmin){
        $Credential = New-Object System.Management.Automation.PSCredential "$((Get-ComputerInfo).CsDNSHostName)\Maintenence", (Get-SecCredentials "$env:ProgramData\OTS\data\1.dat")
        $soon = Get-Date -Second (((Get-Date).Second + 5)%60)
        $Sta =  New-ScheduledTaskAction -Execute "powershell.exe OTS-OfficeRepair"
        $Stt = New-ScheduledTaskTrigger -Once -At $soon
        Register-ScheduledTask OTS_Repair -Action $Sta -Trigger $Stt # User $Credential.UserName -Password (ConvertFrom-SecureString $Credential.Password)
        Write-Output "Relaunch as admin...."
        Start-Sleep -Seconds 5
        return
    }

    $wshell = New-Object -ComObject Wscript.Shell
    $response = $wshell.Popup("This Script will update/repair your Office instalation. It is intended to be run in the case Office applications will not start.`n `n Please save any open work and click OK. ",0,"Office Repair",0x1)
	$actionString = @("<Should not see this>","<Should not see this>")
	if (!($response -eq 1))
	{
		return
	}
    try {
        Write-Output "Get current Office version..."
        $c2rPath = "$($env:CommonProgramW6432)\Microsoft Shared\ClickToRun\OfficeClickToRun.exe"
        $currentVersion = $null
        Get-CimInstance -ClassName Win32_Product | ? { !($_.Name -eq $null) -and $_.Name.Contains("Office")} | ForEach-Object -Process {$currentVersion = $_.Version}
        
        Write-Output "Get latest version..."
        try{
            $targetBuild = Get-LatestOfficeVersion -channel "Current"
        }
        catch {
            throw "Error getting latest version: $_"
        }

        Write-Output "Compare Versions..."
        if ([long]($targetBuild | Remove-FromString -toRemove '.') -gt [long]($currentVersion | Remove-FromString -toRemove '.')){
            $actionString = @("update","updated")
            Write-Output "Starting Update to $targetBuild..."
            $process = Start-Process -FilePath $c2rPath -PassThru -ArgumentList "/update user updatetoversion=$($targetBuild)" -PassThru
        }
        else {
            $actionString = @("repair","repaired")
            Write-Output "Starting Repair..."
            $process = Start-Process -FilePath $c2rPath -PassThru -ArgumentList 'scenario=Repair platform=x64 culture=en-us forceappshutdown=True RepairType=FullRepair DisplayLevel=True'
        }
         if ($process -ne $null){
             Wait-Process -Id $process.Id
        }
        else {
            throw "We missed Click-to-Run Repair, did it run?"
        }

        #old wait process code
        # $process = Get-Process OfficeClickToRun -ErrorAction SilentlyContinue -ErrorVariable ev
        # if ($ev -ne $null) {
        #     [double]$timer = 0
        #     while ($err -ne $null){
        #         $ev = $null
        #         start-sleep -Seconds 0.2
        #         $timer += 0.2
        #         $process = Get-Process OfficeClickToRun -ErrorAction SilentlyContinue -ErrorVariable ev
        #         if ($timer -gt 3) {throw "We missed Click-to-Run Repair, did it run?"}
        #     }
        # }
        # do{ # process can restart so when it terminates we give it a sec to make sure ist really finished
        #     $ev = $null
        #     Wait-Process -Id $process.Id
        #     start-sleep -Seconds 1
        #     $process = Get-Process OfficeClickToRun -ErrorAction SilentlyContinue -ErrorVariable ev
        # } while ($ev = $null)

     }
    catch {
        $response = $wshell.Popup("The $($actionString[0]) did not complete succesfully. `n `nPlease contact Olympus service desk for help at service_desk@OlympusTech.com.au, on 1800 932 964, or by right clicking on the Olympus icon in the taskbar and selecting 'Create Ticket'. `n`nError: $_",0,"Failed",	0x30)
        return
    }

    $response = $wshell.Popup("Office has been $($actionString[1]).`n `nIf this does not resolve your issue then please contact Olympus service desk for help at service_desk@OlympusTech.com.au, on 1800 932 964, or by right clicking on the Olympus icon in the taskbar and selecting 'Create Ticket'.",0,"Complete",0x0)
    
}