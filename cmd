Get-NetTCPConnection |
    ForEach-Object {
        $proc = Get-Process -Id $_.OwningProcess -ErrorAction SilentlyContinue
        [PSCustomObject]@{
            LocalAddress  = $_.LocalAddress
            LocalPort     = $_.LocalPort
            State         = $_.State
            ProcessName   = if ($proc) { $proc.ProcessName } else { "N/A" }
        }
    } |
    Sort-Object LocalPort, ProcessName -Unique
