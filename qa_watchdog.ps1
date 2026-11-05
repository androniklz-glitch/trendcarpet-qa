# QA Hub Watchdog - checks if server on port 8092 and Vercel/Tunnel are responding
# Runs every minute (task Trendcarpet_QA_Watchdog). Rules (CEO 2026-11-05 "fixa hubben"):
#  - port 8092 dead + cloudflared alive -> start ONLY hatshop_server (same tunnel URL, cloudflared never touched)
#  - port 8092 dead + cloudflared dead  -> start_hub_with_tunnel.py (new tunnel + Vercel sync)
$workDir = 'C:\Users\AndronikLindgren\.gemini\antigravity\scratch\Projects\Trendcarpet_Interiors'
$python = 'C:\Users\AndronikLindgren\miniconda3\python.exe'
$logFile = Join-Path $workDir 'qa_watchdog.log'
$vercelJson = Join-Path $workDir 'vercel.json'
$streakFile = Join-Path $workDir 'qa_watchdog_tunnel_fail.txt'
$deployStamp = Join-Path $workDir 'qa_watchdog_vercel_deploy.txt'

function Log-Msg($msg) {
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    "$ts - $msg" | Out-File -FilePath $logFile -Append -Encoding utf8
}

function Test-Local {
    for ($i = 0; $i -lt 2; $i++) {
        try {
            if ((Invoke-WebRequest -Uri 'http://127.0.0.1:8092/api/batches' -UseBasicParsing -TimeoutSec 10).StatusCode -eq 200) { return $true }
        } catch {}
        Start-Sleep -Seconds 2
    }
    return $false
}

function Invoke-VercelDeploy($why) {
    # Max one production deploy per 15 min, whatever the cause
    if ((Test-Path $deployStamp) -and ((Get-Item $deployStamp).LastWriteTime -gt (Get-Date).AddMinutes(-15))) {
        Log-Msg "[INFO] $why - Vercel-deploy hoppas over (senaste < 15 min sedan)."
        return
    }
    Set-Content -Path $deployStamp -Value (Get-Date -Format 's') -Encoding ascii
    Log-Msg "[INFO] $why - kor npx.cmd vercel --prod --yes..."
    Push-Location $workDir
    try {
        & npx.cmd vercel --prod --yes 2>&1 | Out-Null
        Log-Msg "[OK] npx.cmd vercel --prod --yes klar."
    } catch {
        Log-Msg "[WARN] Fel vid npx.cmd vercel: $_"
    } finally {
        Pop-Location
    }
}

# 0. Port 8092 dod men cloudflared lever -> starta BARA servern (samma tunnel-URL, cloudflared rors aldrig)
$cfProcs = @(Get-CimInstance Win32_Process -Filter "Name='cloudflared.exe'" -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -like '*localhost:8092*' })
$localIsUp = Test-Local
if ($cfProcs.Count -gt 0 -and -not $localIsUp) {
    $srv = @(Get-CimInstance Win32_Process -Filter "Name='python.exe'" -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -like '*hatshop_server*' })
    foreach ($s in $srv) {
        if ($s.CreationDate -gt (Get-Date).AddMinutes(-3)) {
            Log-Msg "[INFO] Port 8092 dod, server PID $($s.ProcessId) startades nyss. Avvaktar."
            exit 0
        }
        Log-Msg "[WARN] Server PID $($s.ProcessId) svarar inte pa port 8092 (> 3 min gammal). Dodar den."
        Stop-Process -Id $s.ProcessId -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
    }
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $p = Start-Process -FilePath $python -ArgumentList '-u hatshop_server.py 8092' -WorkingDirectory $workDir -WindowStyle Hidden -PassThru `
        -RedirectStandardOutput (Join-Path $workDir "hub_server_$ts.log") -RedirectStandardError (Join-Path $workDir "hub_server_$ts.err.log")
    Log-Msg "[ALERT] Port 8092 dod, cloudflared PID $(($cfProcs | ForEach-Object { $_.ProcessId }) -join ',') lever. Startade ENDAST servern (PID $($p.Id)), tunneln orord."
    exit 0
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
    for ($i = 0; $i -lt 2 -and -not $tunnelIsUp; $i++) {
        try {
            if ((Invoke-WebRequest -Uri "$tunnelUrl/api/batches" -UseBasicParsing -TimeoutSec 10).StatusCode -eq 200) { $tunnelIsUp = $true }
        } catch { Start-Sleep -Seconds 2 }
    }
}

function Sync-VercelIfTunnelChanged($oldUrl) {
    for ($i = 0; $i -lt 15; $i++) {
        Start-Sleep -Seconds 2
        if (Test-Path $vercelJson) {
            try {
                $vData = Get-Content $vercelJson -Raw -Encoding utf8 | ConvertFrom-Json
                $dest = $vData.rewrites[0].destination
                if ($dest -match '^(https?://[^/]+)') {
                    $currUrl = $matches[1]
                    if ($currUrl -and ($currUrl -ne $oldUrl)) {
                        Invoke-VercelDeploy "Ny tunnel-URL upptackt ($currUrl)"
                        break
                    }
                }
            } catch {}
        }
    }
}

if ($tunnelIsUp) {
    if (Test-Path $streakFile) { Remove-Item $streakFile -Force -ErrorAction SilentlyContinue }
    $vercelIsUp = $false
    try {
        if ((Invoke-WebRequest -Uri 'https://trendcarpet-qa.vercel.app/api/batches' -UseBasicParsing -TimeoutSec 10).StatusCode -eq 200) { $vercelIsUp = $true }
    } catch {}

    if ($vercelIsUp) {
        # Tunneln fungerar felfritt! Ror ingenting.
        Log-Msg "[OK] Tunnel i vercel.json svarar 200 OK ($tunnelUrl). Ingen omstart kravs."
    } else {
        Invoke-VercelDeploy "Tunnel ($tunnelUrl) svarar 200 OK men Vercel svarar ej 200"
    }
    exit 0
}

# 3. Tunneln svarade inte. Kontrollera om start_hub_with_tunnel kors
$myPid = $PID
$hubProcs = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object {
    $_.ProcessId -ne $myPid -and $_.CommandLine -like '*start_hub_with_tunnel*' -and $_.Name -like '*python*'
}

if ($localIsUp) {
    if ($hubProcs) {
        Log-Msg "[INFO] Lokal server ar 200 OK och start_hub_with_tunnel kors (PID $($hubProcs.ProcessId)). Latar interna hanteraren skota tunneln."
        exit 0
    }
    if ($cfProcs.Count -gt 0) {
        # cloudflared lever men tunnel-URL:en svarar inte: kraver 3 missar i rad (~3 min) innan ny tunnel
        $streak = 1
        if (Test-Path $streakFile) { $streak = [int](Get-Content $streakFile -Raw) + 1 }
        Set-Content -Path $streakFile -Value $streak -Encoding ascii
        if ($streak -lt 3) {
            Log-Msg "[WARN] Tunnel ($tunnelUrl) svarar inte (miss $streak av 3), cloudflared lever. Avvaktar."
            exit 0
        }
        Log-Msg "[ALERT] Tunnel ($tunnelUrl) dod 3 kontroller i rad trots levande cloudflared. Byter tunnel."
        $cfProcs | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    }
    Remove-Item $streakFile -Force -ErrorAction SilentlyContinue
    Log-Msg "[WARN] Lokal server svarar men tunneln ar nere och start_hub_with_tunnel kors inte. Startar start_hub_with_tunnel.py..."
    Start-Process -FilePath $python -ArgumentList 'start_hub_with_tunnel.py' -WorkingDirectory $workDir -WindowStyle Hidden
    Sync-VercelIfTunnelChanged $tunnelUrl
    exit 0
}

# 4. Varken tunnel eller lokal server svarar (och ingen cloudflared for port 8092 lever)
if (-not $hubProcs) {
    Log-Msg "[ALERT] Varken tunnel eller port 8092 svarar, och process saknas. Startar om QA-hubben..."
    Start-Process -FilePath $python -ArgumentList 'start_hub_with_tunnel.py' -WorkingDirectory $workDir -WindowStyle Hidden
    Sync-VercelIfTunnelChanged $tunnelUrl
} else {
    Log-Msg "[WARN] Port 8092 svarar inte tillfalligt men start_hub_with_tunnel rullar redan (PID $($hubProcs.ProcessId)). Avvaktar."
}
