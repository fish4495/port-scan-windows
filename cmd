Get-NetTCPConnection | Select-Object LocalAddress, LocalPort, State, OwningProcess |
    Sort-Object LocalPort |
    ForEach-Object {
        $proc = Get-Process -Id $_.OwningProcess -ErrorAction SilentlyContinue
        [PSCustomObject]@{
            LocalAddress  = $_.LocalAddress
            LocalPort     = $_.LocalPort
            State         = $_.State
            ProcessName   = if ($proc) { $proc.ProcessName } else { "N/A" }
        }
    }
