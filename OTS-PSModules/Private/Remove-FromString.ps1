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