function OTS-OfficeRepair{

    if (!(Admin-Check OTS-OfficeRepair)) {
        return
    } # if we are not admin then rerun with correct credentials

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
            Start-Process -FilePath $c2rPath -ArgumentList "/update user updatetoversion=$($targetBuild)"
        }
        else {
            $actionString = @("repair","repaired")
            Write-Output "Starting Repair..."
            Start-Process -FilePath $c2rPath -Verb RunAs -ArgumentList "scenario=Repair platform=x64 culture=en-us forceappshutdown=True RepairType=FullRepair DisplayLevel=True"
        }
        $process = Get-Process OfficeClickToRun -ErrorAction SilentlyContinue -ErrorVariable ev
        if ($ev -ne $null) {
            [double]$timer = 0
            while ($err -ne $null){
                $ev = $null
                start-sleep -Seconds 0.2
                $timer += 0.2
                $process = Get-Process OfficeClickToRun -ErrorAction SilentlyContinue -ErrorVariable ev
                if ($timer -gt 3) {throw "We missed Click-to-Run Repair, did it run?"}
            }
        }
        do{ # process can restart so when it terminates we give it a sec to make sure ist really finished
            $ev = $null
            Wait-Process -Id $process.Id
            start-sleep -Seconds 1
            $process = Get-Process OfficeClickToRun -ErrorAction SilentlyContinue -ErrorVariable ev
        } while ($ev = $null)
     }
    catch {
        $response = $wshell.Popup("The $($actionString[0]) did not complete succesfully. `n `nPlease contact Olympus service desk for help at service_desk@OlympusTech.com.au, on 1800 932 964, or by right clicking on the Olympus icon in the taskbar and selecting 'Create Ticket'. `n`nError: $_",0,"Failed",	0x30)
        return
    }

    $response = $wshell.Popup("Office has been $($actionString[1]).`n `nIf this does not resolve your issue then please contact Olympus service desk for help at service_desk@OlympusTech.com.au, on 1800 932 964, or by right clicking on the Olympus icon in the taskbar and selecting 'Create Ticket'.",0,"Complete",0x0)
    
}