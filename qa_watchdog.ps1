# QA Hub Watchdog - checks if server on port 8092 and Vercel/Tunnel are responding
$workDir = 'C:\Users\AndronikLindgren\.gemini\antigravity\scratch\Projects\Trendcarpet_Interiors'
$python = 'C:\Users\AndronikLindgren\miniconda3\python.exe'
$logFile = Join-Path $workDir 'qa_watchdog.log'
$vercelJson = Join-Path $workDir 'vercel.json'

function Log-Msg($msg) {
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    "$ts - $msg" | Out-File -FilePath $logFile -Append -Encoding utf8
}

# 1. Hamta den aktiva tunnel-URL:en ur vercel.json
$tunnelUrl = $null
if (Test-Path $vercelJson) {
    try {
        $vData = Get-Content $vercelJson -Raw -Encoding utf8 | ConvertFrom-Json
        $dest = $vData.rewrites[0].destination
        if ($dest -match '^(https?://[^/]+)') {
            $tunnelUrl = $matches[1]
        }
    } catch {
        Log-Msg "[WARN] Kunde inte lasa vercel.json: $_"
    }
}

# 2. Fasta regeln: Om tunneln i vercel.json svarar far cloudflared ALDRIG dodas!
$tunnelIsUp = $false
if ($tunnelUrl) {
    try {
        $resT = Invoke-WebRequest -Uri "$tunnelUrl/api/batches" -UseBasicParsing -TimeoutSec 10
        if ($resT.StatusCode -eq 200) {
            $tunnelIsUp = $true
        }
    } catch {
        # Retry en gang vid tillfallig natverkssvacka
        Start-Sleep -Seconds 2
        try {
            $resT = Invoke-WebRequest -Uri "$tunnelUrl/api/batches" -UseBasicParsing -TimeoutSec 10
            if ($resT.StatusCode -eq 200) {
                $tunnelIsUp = $true
            }
        } catch {}
    }
}

if ($tunnelIsUp) {
    # Tunneln fungerar felfritt! Ror ingenting.
    Log-Msg "[OK] Tunnel i vercel.json svarar 200 OK ($tunnelUrl). Ingen omstart kravs."
    exit 0
}

# 3. Tunneln svarade inte. Kontrollera lokal server pa port 8092
$localIsUp = $false
try {
    $res = Invoke-WebRequest -Uri 'http://127.0.0.1:8092/api/batches' -UseBasicParsing -TimeoutSec 10
    if ($res.StatusCode -eq 200) {
        $localIsUp = $true
    }
} catch {
    Start-Sleep -Seconds 2
    try {
        $res = Invoke-WebRequest -Uri 'http://127.0.0.1:8092/api/batches' -UseBasicParsing -TimeoutSec 10
        if ($res.StatusCode -eq 200) {
            $localIsUp = $true
        }
    } catch {
        $localIsUp = $false
    }
}

# Kontrollera om start_hub_with_tunnel kors
$myPid = $PID
$hubProcs = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object { 
    $_.ProcessId -ne $myPid -and $_.CommandLine -like '*start_hub_with_tunnel*' -and $_.Name -like '*python*'
}

if ($localIsUp) {
    if ($hubProcs) {
        Log-Msg "[INFO] Lokal server ar 200 OK och start_hub_with_tunnel kors (PID $($hubProcs.ProcessId)). Latar interna hanteraren skota tunneln."
        exit 0
    } else {
        Log-Msg "[WARN] Lokal server svarar men start_hub_with_tunnel kors inte. Startar start_hub_with_tunnel.py..."
        Start-Process -FilePath $python -ArgumentList 'start_hub_with_tunnel.py' -WorkingDirectory $workDir -WindowStyle Hidden
        exit 0
    }
}

# 4. Varken tunnel eller lokal server svarar
if (-not $hubProcs) {
    Log-Msg "[ALERT] Varken tunnel eller port 8092 svarar, och process saknas. Startar om QA-hubben..."
    Get-Process -Name cloudflared -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Process -FilePath $python -ArgumentList 'start_hub_with_tunnel.py' -WorkingDirectory $workDir -WindowStyle Hidden
} else {
    Log-Msg "[WARN] Port 8092 svarar inte tillfalligt men start_hub_with_tunnel rullar redan (PID $($hubProcs.ProcessId)). Avvaktar."
}
