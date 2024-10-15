function Get-SecCredentials{

    param(
        [Parameter(Mandatory)][string]$Path
    )

    $timeKey = Get-UnlockKey($Path)
    $result = (Get-Content -Path $path) | ConvertTo-SecureString -Key $timeKey
    return $result

}