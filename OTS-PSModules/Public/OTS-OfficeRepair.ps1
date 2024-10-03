function OTS-OfficeRepair{
    $wshell = New-Object -ComObject Wscript.Shell
    $response = $wshell.Popup("This Script will update/repair your Office instalation. It is intended to be run in the case Office applications will not start.`n `n Please save any open work and click the button. ",0,"Office Repair",0x1)
	$actionString = @("Nonthing","Ignored") # Shouldn't see this
	if (!($response -eq 1))
	{
		return
	}
    try {
        $c2rPath = "$($env:CommonProgramW6432)\Microsoft Shared\ClickToRun\officec2rclient.exe"
        $currentVersion = $null
        Get-CimInstance -ClassName Win32_Product | ? { !($_.Name -eq $null) -and $_.Name.Contains("Office")} | ForEach-Object -Process {$currentVersion = $_.Version}

        $targetBuild = Get-LatestOfficeVersion -channel "Current"

        if ([long]($targetBuild | Remove-FromString -toRemove '.') -gt [long]($currentVersion | Remove-FromString -toRemove '.')){
            $actionString = @("update","updated")
        Write-Output "Starting Update to $targetBuild..."
        Start-Process -FilePath $c2rPath -ArgumentList "/update user updatetoversion=$($targetBuild)"
        }
        else {
            Write-Output "Starting Repair..."
            Start-Process -FilePath $c2rPath -ArgumentList "scenario=Repair platform=x64 culture=en-us forceappshutdown=True RepairType=FullRepair DisplayLevel=True"
        }
        $process = Get-Process OfficeC2RClient -ErrorAction SilentlyContinue -ErrorVariable ev
        if ($ev -ne $null) {
            [double]$timer = 0
            while ($err -ne $null){
                Write-Output "Wait..."
                $ev = $null
                start-sleep -Seconds 0.2
                $timer += 0.2
                $process = Get-Process OfficeC2RClient -ErrorAction SilentlyContinue -ErrorVariable ev
                if ($timer -gt 3) {throw "We missed Click-to-Run Repair, did it run?"}
            }
        }

        Wait-Process -Id $process.Id
        $actionString = $null #intentional error for TS
     }
    catch {
        $response = $wshell.Popup("The $($actionString[0]) did not complete succesfully. `n `nPlease contact Olympus service desk for help at service_desk@OlympusTech.com.au, on 1800 932 964, or by right clicking on the Olympus icon in the taskbar and selecting 'Create Ticket'. `n`nError: $_",0,"Failed",	0x30)
        return
    }

    $response = $wshell.Popup("Office has been $($actionString[1]).`n `nIf this does not resolve your issue then please contact Olympus service desk for help at service_desk@OlympusTech.com.au, on 1800 932 964, or by right clicking on the Olympus icon in the taskbar and selecting 'Create Ticket'.",0,"Complete",0x0)
    
}