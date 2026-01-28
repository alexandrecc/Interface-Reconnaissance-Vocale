# Durable TCP listener: accepts many connections sequentially and never exits by itself
$ErrorActionPreference = 'Continue'

function Parse-KV {
  param([string]$Text)
  $map = @{}
  foreach ($pair in ($Text -split ';')) {
    if ([string]::IsNullOrWhiteSpace($pair)) { continue }
    $kv = $pair -split '=', 2
    if ($kv.Count -eq 2) { $map[$kv[0].Trim()] = $kv[1].Trim() }
  }
  return $map
}

function Send-ForwardBytes {
    param(
        [byte[]]$Bytes,
        [string]$DestHost,
        [int]$Port
    )
    if (-not $Bytes -or $Bytes.Length -eq 0) { return }

    $fw = $null; $fwStream = $null
    try {
        $fw = [System.Net.Sockets.TcpClient]::new()
        $fw.ReceiveTimeout = 10000
        $fw.SendTimeout    = 10000
        $fw.Connect($DestHost, $Port)
        $fwStream = $fw.GetStream()
        $fwStream.Write($Bytes, 0, $Bytes.Length)
        $fwStream.Flush()
        Write-Host ("Forwarded {0} bytes to {1}:{2}" -f $Bytes.Length, $DestHost, $Port)
    }
    catch {
        Write-Warning "Forwarding failed"
    }
    finally {
        if ($fwStream) { try { $fwStream.Dispose() } catch {} }
        if ($fw)       { try { $fw.Close() } catch {} }
    }
}
$script:current_accnb = $null
$script:current_proc  = $null
$script:seen_proc     = @()      # mémorise tous les ProcDesc du lot courant

$enableForward = $true
$forwardHost   = '127.0.0.1'
$forwardPort   = 9876

$port = 9879
$log  = Join-Path $env:USERPROFILE 'tcp_message.log'
# Fichier-flag qui contrôle la pause
$pauseFlag = Join-Path $env:USERPROFILE 'pause.flag'


$listener = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Any, $port)
$listener.Start()
Write-Host "Listening on $port (PID=$PID). Close the powershell to stop."
    if ($enableForward) {
       Write-Host "Forwarding on $forwardPort."
    }
Write-Host "Logging to: $log"

for (;;) {
  try {
    # Block here until a client connects
    $client = $listener.AcceptTcpClient()
    $remote = $client.Client.RemoteEndPoint
    Write-Host "Accepted $remote"

    $stream = $client.GetStream()
    # Optional: comment out next line to avoid server-side read timeouts
    # $stream.ReadTimeout = 30000

    $buf = New-Object byte[] 4096
    $mem = New-Object IO.MemoryStream
    while (($n = $stream.Read($buf, 0, $buf.Length)) -gt 0) {
      $mem.Write($buf, 0, $n)
    }

    $bytes = $mem.ToArray()
    $isPaused = Test-Path -LiteralPath $pauseFlag
    if ($isPaused -and $enableForward) {
        Write-Host ("Pause active -> Forward to {0}:{1} ({2} bytes)" -f $forwardHost, $forwardPort, $bytes.Length)
        Send-ForwardBytes -Bytes $bytes -DestHost $forwardHost -Port $forwardPort
    }
    # Prefer UTF8; if decoding yields empty string, fall back to ASCII
    $text  = [Text.Encoding]::UTF8.GetString($bytes)
    if ([string]::IsNullOrEmpty($text) -and $bytes.Length -gt 0) {
      $text = [Text.Encoding]::ASCII.GetString($bytes)
    }

    $fields = Parse-KV $text

    if ($fields.ContainsKey('ProcDesc')) {
    $first_exam = $true

    $proc   = $fields['ProcDesc']
    $mod    = $fields['ModalityType']
    $loc    = $fields['Other2']
    $accnb  = $fields['AccNb']           # e.g. RA202508615001
    $reqnb  = $fields['OrderNb']
    $patdos  = $fields['PatMRN']
    $patnom  = $fields['PatFullName']

    # === Vérifier le reset après réception du message ===
    $resetFlag = Join-Path $env:USERPROFILE 'tcp_listener.reset'
    if (Test-Path $resetFlag) {
        try { Remove-Item $resetFlag -Force } catch {}
        $script:current_accnb = $null
        $script:current_proc  = $null
        Write-Host "==> Reset appliqué (current_accnb & current_proc remis à null)"
    }

    if ($script:current_accnb -and $accnb) {
        $prev = $script:current_accnb
        $curr = $accnb

        if ($prev.Length -ge 2 -and $curr.Length -ge 2) {
            $prevPrefix = $prev.Substring(0, $prev.Length - 2)
            $currPrefix = $curr.Substring(0, $curr.Length - 2)
            if ($prevPrefix -eq $currPrefix) { $first_exam = $false }
        }
    }
    
    # === Gestion des doublons ProcDesc dans le même lot ===
    $skipExam = $false

    if ($first_exam) {
        # nouveau lot -> on vide la mémoire
        $script:seen_proc = @()
    } else {
        # même lot -> vérifier si ce ProcDesc a déjà été vu
        if ($script:seen_proc -contains $proc) {
            Write-Host "Examen déjà vu dans ce lot (AccNb=$accnb, ProcDesc=$proc) — ignoré."
            $skipExam = $true
           }
    }

# Si pas doublon, on l'ajoute à la mémoire
if (-not $skipExam) {
    $script:seen_proc += $proc
}

    
    $first  = if ($first_exam) { "1" } else { "0" }   # send a simple 1/0 flag
    
    # --- Si un reset vient d'être fait, forcer first=1 ---
    if ($script:force_first) {
        $first_exam = $true
        $script:force_first = $false
    }

    $script:current_accnb = $accnb  # update for next message
    $script:current_proc  = $proc

    # === Si doublon, on saute le lancement ===
    if (-not $skipExam) {
        if (-not $isPaused) {
            Write-Host "ProcDesc: $proc (first_exam=$first_exam)"
            & ".\Initialisation.ahk" "$proc" "$mod" "$loc" "$first" "$reqnb" "$patdos" "$patnom"
        }
        else {
            Write-Host "Pause active – Initialisation.ahk ignoré (forward déjà effectué)."
        }
    }
}

    $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
    $entry = "[${stamp}] From ${remote}`r`n${text}`r`n---`r`n"
    try {
      Add-Content -Path $log -Value $entry -Encoding UTF8
    } catch {
      Write-Warning "Log write failed: $($_.Exception.Message)"
    }

    Write-Host "Saved $($bytes.Length) bytes. Closed $remote."
  }
  catch {
    # Swallow and continue so the listener doesn't die
    Write-Warning "Loop error: $($_.Exception.Message)"
  }
  finally {
    if ($stream) { try { $stream.Dispose() } catch {} }
    if ($client) { try { $client.Close() } catch {} }
  }
}
