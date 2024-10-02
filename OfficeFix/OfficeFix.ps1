


function Run-OfficeUpdateOrFix
{
    $wshell = New-Object -ComObject Wscript.Shell
    $response = $wshell.Popup("This Script will update/repair your Office instalation, then temprarily disable updates. It is intended to be run in the case Office applications will not start ",0,"Office Repair",0x1)
	
	if (!($response -eq 1))
	{
		return
	}

    try {
        $currentVersion = $null
        Get-CimInstance -ClassName Win32_Product | ? { !($_.Name -eq $null) -and $_.Name.Contains("Office")} | ForEach-Object -Process {$currentVersion = $_.Version}

        $targetBuild = Get-LatestOfficeVersion -channel "Current"
        $actionString = "Repair"

        if (!($targetBuild -eq $currentVersion)){
        $actionString = "Update"
        Write-Output "Starting Update..."
        Start-Process -FilePath "cmd.exe" -ArgumentList "'$($env:CommonProgramW6432)\Microsoft Shared\ClickToRun\ClickToRunofficec2rclient.exe' /update user updatetoversion=$($targetBuild)"
        }
        else {
            Write-Output "Starting Repair..."
            Start-Process -FilePath "cmd.exe" -ArgumentList "'$($env:CommonProgramW6432)\Microsoft Shared\ClickToRun\ClickToRunofficec2rclient.exe' /update user updatetoversion=$($targetBuild)"
        }
    }
    catch {
        $wshell.Popup("The $actionString did not complete succesfully. Please contact Olympus service desk for help at service_desk@OlympusTech.com.au, on 1800 932 964, or by right clicking on the Olympus icon in the taskbar and selecting 'Create Ticket'",0,"Repair Failed",	0x30)
        return
    }
    $wshell.Popup("Office has been $actionString" + "ed. If this does not resolve your issue then please contact Olympus service desk for help at service_desk@OlympusTech.com.au, on 1800 932 964, or by right clicking on the Olympus icon in the taskbar and selecting 'Create Ticket'",0,"Complete",0x0)
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

    Start-Process -FilePath "cmd.exe" -ArgumentList "/c extrac32.exe /Y /E /L $($catDir) releasehistory.cab" | Out-Null

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