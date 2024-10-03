
function Remove-FromString{

    param (
        [Parameter(Mandatory,ValueFromPipeline)][string]$stringIn,
        $toRemove
    )

    [string]$result = ""

    for ($i=0;$i -lt $stringIn.Length;$i++) {
        if ($stringIn[$i] -ne $toRemove){
            $result += $stringIn[$i]
        }
    }
    return $result
}
function Run-OfficeUpdateOrFix{
    $wshell = New-Object -ComObject Wscript.Shell
    $response = $wshell.Popup("This Script will update/repair your Office instalation. It is intended to be run in the case Office applications will not start.`n `n Please save any open work and click OK. ",0,"Office Repair",0x1)
	$actionString = @("repair","repaired")
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
        Write-Output "Action: $actionString"
        $process = Get-Process OfficeC2RClient -ErrorAction SilentlyContinue -ErrorVariable err
        if ($err -ne $null) {throw "Office Click-to-Run process didn't seem to run."}
        Wait-Process -Id $process.Id
 
    }
    catch {
        $response = $wshell.Popup("The $($actionString[0]) did not complete succesfully. `n `nPlease contact Olympus service desk for help at service_desk@OlympusTech.com.au, on 1800 932 964, or by right clicking on the Olympus icon in the taskbar and selecting 'Create Ticket'. `n`nError: $_",0,"Failed",	0x30)
        return
    }

    $response = $wshell.Popup("Office has been $($actionString[1]).`n `nIf this does not resolve your issue then please contact Olympus service desk for help at service_desk@OlympusTech.com.au, on 1800 932 964, or by right clicking on the Olympus icon in the taskbar and selecting 'Create Ticket'.",0,"Complete",0x0)
    
}

function OLD{
    
}

function Get-LatestOfficeVersion {

    param([Parameter(Mandatory,ValueFromPipeline)][string]$channel)

    $catalogUrl = "https://officecdn.microsoft.com/pr/wsus/releasehistory.cab"
    $catDir = $env:ProgramData +"\OTS\Catalogs\"

    If(!(test-path -PathType container $catDir))
    {
        New-Item -ItemType Directory -Path $catDir
    }

    Invoke-WebRequest $catalogUrl -OutFile "$($catDir)releasehistory.cab"

    Start-Process -FilePath "cmd.exe" -ArgumentList "/c extrac32.exe /Y /E /L $($catDir) $($catDir)releasehistory.cab" | Out-Null

    $process = Get-Process extrac32 -ErrorAction SilentlyContinue -ErrorVariable err
    if ($err -ne $null) {throw "Office Click-to-Run process didn't seem to run."}
    Wait-Process -Id $process.Id

    $catalog = Get-Content -Path "$($catDir)releasehistory.xml" -Raw

    $targetChannel = (Select-Xml -Content $catalog -XPath "//UpdateChannel" | ? {$_.Node.Name -eq $channel} )
    if ($targetChannel.Length -lt 1)
    {
        throw "Failed to find specified release channel details"
    }

    $targetBuild = ($targetChannel |foreach {$_.Node.Update} | ? {$_.Latest -eq "True"}).LegacyVersion

    if ($targetBuild.Length -lt 1)
    {
        throw "Failed to locate a valid build number"
    }

    return $targetBuild
}

