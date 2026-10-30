$ErrorActionPreference = 'Stop'
$deploy = 'C:\Users\WIN10\first-leg-cost-deploy\first-leg-cost'
$logDir = 'C:\Users\WIN10\first-leg-cost-deploy\logs'
$env:Path = [System.Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [System.Environment]::GetEnvironmentVariable('Path','User')

# Copy updated files from deploy parent if present
foreach ($f in @('app.py','README.md')) {
  $src = Join-Path 'C:\Users\WIN10\first-leg-cost-deploy' $f
  if (Test-Path $src) { Copy-Item $src (Join-Path $deploy $f) -Force }
}

# Stop old
Get-CimInstance Win32_Process -Filter "Name='python.exe'" -ErrorAction SilentlyContinue |
  Where-Object { $_.CommandLine -match 'app\.py' } |
  ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
Get-Process cloudflared -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

# Start flask
$py = Join-Path $deploy '.venv\Scripts\python.exe'
$env:PORT = '8765'
Remove-Item Env:FLASK_DEBUG -ErrorAction SilentlyContinue
$flaskOut = Join-Path $logDir 'flask.out.log'
$flaskErr = Join-Path $logDir 'flask.err.log'
$flaskProc = Start-Process -FilePath $py -ArgumentList 'app.py' -WorkingDirectory $deploy -RedirectStandardOutput $flaskOut -RedirectStandardError $flaskErr -WindowStyle Hidden -PassThru
Set-Content (Join-Path $logDir 'flask.pid') $flaskProc.Id

$ok = $false
for ($i=0; $i -lt 30; $i++) {
  try {
    $r = Invoke-WebRequest -Uri 'http://127.0.0.1:8765/' -UseBasicParsing -TimeoutSec 2
    if ($r.StatusCode -eq 200) { $ok = $true; break }
  } catch { Start-Sleep -Milliseconds 500 }
}
if (-not $ok) { Get-Content $flaskErr -Tail 40; throw 'flask not ready' }

$cfOut = Join-Path $logDir 'cloudflared.out.log'
$cfErr = Join-Path $logDir 'cloudflared.err.log'
$cfProc = Start-Process -FilePath 'cloudflared' -ArgumentList @('tunnel','--url','http://127.0.0.1:8765') -WorkingDirectory $deploy -RedirectStandardOutput $cfOut -RedirectStandardError $cfErr -WindowStyle Hidden -PassThru
Set-Content (Join-Path $logDir 'cloudflared.pid') $cfProc.Id

$url = $null
for ($i=0; $i -lt 40; $i++) {
  Start-Sleep -Seconds 1
  $text = ''
  if (Test-Path $cfErr) { $text += Get-Content $cfErr -Raw -EA SilentlyContinue }
  if (Test-Path $cfOut) { $text += Get-Content $cfOut -Raw -EA SilentlyContinue }
  if ($text -match 'https://[a-zA-Z0-9.-]+\.trycloudflare\.com') { $url = $Matches[0]; break }
}
if (-not $url) { Get-Content $cfErr -Tail 40; throw 'no url' }
Set-Content (Join-Path $logDir 'public_url.txt') $url
Write-Host "PUBLIC_URL=$url"
