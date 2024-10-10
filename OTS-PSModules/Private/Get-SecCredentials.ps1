function Get-SecCredentials{
    param(
        [Parameter(Mandatory)][string]$Path
    )

    $datetime = ([DateTimeOffset](Get-Date (Get-Item -Path $Path).CreationTime -Format d)).ToUnixTimeSeconds()

    [byte[]]$timeKey = @(0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)
    for ($i = 0;$i -lt 8; $i++){
        $timeKey[$i] = ($datetime -shr (8*$i)) -band 255
    }
    $result = (Get-Content -Path $path) | ConvertTo-SecureString -Key $timeKey
    return $result
}