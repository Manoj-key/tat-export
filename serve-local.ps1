# serve-local.ps1 -- zero-install local web server for the KGSO Tableau extensions (Windows PowerShell 5+, no admin).
# Serves THIS folder (the launcher, tatexport\ and woexport\) at http://localhost:8766/ so Tableau Desktop can load them.  Loopback only: nothing on the
# network can reach it, and Windows Firewall does not prompt.  Port 8766 -- the TAT extension uses 8765, so both can run.
# Start it with serve-local.bat (visible window).  -Quiet suppresses the per-request log.
param([switch]$Quiet)
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$port = 8766
$types = @{ '.html'='text/html; charset=utf-8'; '.js'='application/javascript; charset=utf-8'; '.css'='text/css'; '.png'='image/png'; '.json'='application/json'; '.txt'='text/plain; charset=utf-8'; '.trex'='application/xml'; '.ico'='image/x-icon' }
function Say($m) { if (-not $Quiet) { Write-Host $m } }

# already running?  (double-clicking twice) -- say so and exit instead of failing
$listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $port)
try { $listener.Start() }
catch {
  $mine = $false
  try {
    $probe = New-Object System.Net.WebClient
    $mine = ($probe.DownloadString("http://localhost:$port/index.html")) -match 'KGSO Tableau extensions'
  } catch { $mine = $false }
  if ($mine) { Say "Already running at http://localhost:$port/index.html - nothing to do."; if (-not $Quiet) { Start-Sleep 3 }; exit 0 }
  Write-Host "Port $port is in use by another program. Close it, or tell the TAT and SOT Performance team so the extensions can be moved to another port."
  if (-not $Quiet) { Start-Sleep 20 }
  exit 1
}

Say "KGSO Tableau extensions are served at http://localhost:$port/index.html"
Say "Folder: $root"
Say "Leave this window open while you use the dashboard. Press Ctrl+C to stop."
while ($true) {
  $client = $listener.AcceptTcpClient()
  try {
    $stream = $client.GetStream(); $stream.ReadTimeout = 5000
    $reader = New-Object System.IO.StreamReader($stream)
    $requestLine = $reader.ReadLine()
    while ($true) { $h = $reader.ReadLine(); if ($null -eq $h -or $h -eq '') { break } }
    $path = '/'
    if ($requestLine -match '^(GET|HEAD)\s+(\S+)') { $path = $Matches[2] }
    $path = ($path -split '\?')[0]; $path = [uri]::UnescapeDataString($path)
    if ($path -eq '/' -or $path -eq '') { $path = '/index.html' }
    $file = Join-Path $root ($path -replace '/', '\')
    $full = [System.IO.Path]::GetFullPath($file)
    if ($full.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase) -and (Test-Path $full -PathType Leaf)) {
      $bytes = [System.IO.File]::ReadAllBytes($full)
      $ext = [System.IO.Path]::GetExtension($full).ToLower()
      $ctype = if ($types.ContainsKey($ext)) { $types[$ext] } else { 'application/octet-stream' }
      $head = "HTTP/1.1 200 OK`r`nContent-Type: $ctype`r`nContent-Length: $($bytes.Length)`r`nCache-Control: no-cache`r`nConnection: close`r`n`r`n"
      Say ("200 " + $path)
    } else {
      $bytes = [System.Text.Encoding]::UTF8.GetBytes("Not found: $path")
      $head = "HTTP/1.1 404 Not Found`r`nContent-Type: text/plain`r`nContent-Length: $($bytes.Length)`r`nConnection: close`r`n`r`n"
      Say ("404 " + $path)
    }
    $hb = [System.Text.Encoding]::ASCII.GetBytes($head)
    $stream.Write($hb, 0, $hb.Length); $stream.Write($bytes, 0, $bytes.Length); $stream.Flush()
  } catch { Say ("error: " + $_.Exception.Message) }
  finally { $client.Close() }
}
