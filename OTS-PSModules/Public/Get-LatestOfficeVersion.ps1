function Get-LatestOfficeVersion {

    param([Parameter(Mandatory,ValueFromPipeline)][string]$channel)
    Write-Output "Download Catalog..."
    $catalogUrl = "https://officecdn.microsoft.com/pr/wsus/releasehistory.cab"
    $catDir = $env:ProgramData +"\OTS\Catalogs"

    If(!(test-path -PathType container $catDir))
    {
        New-Item -ItemType Directory -Path $catDir
    }

    Invoke-WebRequest $catalogUrl -OutFile "$($catDir)\releasehistory.cab"

     Write-Output "Extract..."

    Start-Process -FilePath "cmd.exe" -ArgumentList "/c extrac32.exe /Y /E /L $($catDir) $($catDir)\releasehistory.cab" | Out-Null
    Start-Sleep -Seconds 0.5 #give time to extract
    if (!(test-path "$catDir\releasehistory.xml")) {
        throw "Failed to extract version information"
    }

    # $process = Get-Process extrac32 -ErrorAction SilentlyContinue -ErrorVariable err
    # # if ($err -ne $null) {
    # #     [double]$timer = 0
    # #     while ($err -ne $null){
    # #         Write-Output "Wait..."
    # #         $err = $null
    # #         start-sleep -Seconds 0.05
    # #         $timer += 0.05
    # #         $process = Get-Process extrac32 -ErrorAction SilentlyContinue -ErrorVariable err
    # #         if ($timer -gt 3) {throw "Did not see file extraction run"}
    # #     }
    # # }
    # # slee -Id $process.Id

    $catalog = Get-Content -Path "$($catDir)\releasehistory.xml" -Raw

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