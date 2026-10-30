$ErrorActionPreference = 'Stop'
$deploy = 'C:\Users\WIN10\first-leg-cost-deploy\first-leg-cost'
$logDir = 'C:\Users\WIN10\first-leg-cost-deploy\logs'
New-Item -ItemType Directory -Force -Path $logDir | Out-Null

# Refresh PATH
$env:Path = [System.Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [System.Environment]::GetEnvironmentVariable('Path','User')

# Kill prior instances if any
Get-CimInstance Win32_Process -Filter "Name='python.exe'" -ErrorAction SilentlyContinue |
  Where-Object { $_.CommandLine -match 'app\.py' } |
  ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
Get-Process cloudflared -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

# Start Flask bound to 0.0.0.0:8765
$py = Join-Path $deploy '.venv\Scripts\python.exe'
$env:PORT = '8765'
$flaskOut = Join-Path $logDir 'flask.out.log'
$flaskErr = Join-Path $logDir 'flask.err.log'
$flaskProc = Start-Process -FilePath $py -ArgumentList 'app.py' -WorkingDirectory $deploy -RedirectStandardOutput $flaskOut -RedirectStandardError $flaskErr -WindowStyle Hidden -PassThru
Set-Content -Path (Join-Path $logDir 'flask.pid') -Value $flaskProc.Id
Write-Host "flask_pid=$($flaskProc.Id)"

# Wait for local health
$ok = $false
for ($i=0; $i -lt 30; $i++) {
  try {
    $r = Invoke-WebRequest -Uri 'http://127.0.0.1:8765/' -UseBasicParsing -TimeoutSec 2
    if ($r.StatusCode -eq 200) { $ok = $true; break }
  } catch { Start-Sleep -Milliseconds 500 }
}
if (-not $ok) {
  Write-Host 'Flask failed to become ready'
  Get-Content $flaskErr -ErrorAction SilentlyContinue | Select-Object -Last 40
  exit 1
}
Write-Host 'Flask ready on 8765'

# Start cloudflared quick tunnel
$cfOut = Join-Path $logDir 'cloudflared.out.log'
$cfErr = Join-Path $logDir 'cloudflared.err.log'
# cloudflared writes informational to stderr typically
$cfProc = Start-Process -FilePath 'cloudflared' -ArgumentList @('tunnel','--url','http://127.0.0.1:8765') -WorkingDirectory $deploy -RedirectStandardOutput $cfOut -RedirectStandardError $cfErr -WindowStyle Hidden -PassThru
Set-Content -Path (Join-Path $logDir 'cloudflared.pid') -Value $cfProc.Id
Write-Host "cloudflared_pid=$($cfProc.Id)"

$url = $null
for ($i=0; $i -lt 40; $i++) {
  Start-Sleep -Seconds 1
  $text = ''
  if (Test-Path $cfErr) { $text += Get-Content $cfErr -Raw -ErrorAction SilentlyContinue }
  if (Test-Path $cfOut) { $text += Get-Content $cfOut -Raw -ErrorAction SilentlyContinue }
  if ($text -match 'https://[a-zA-Z0-9.-]+\.trycloudflare\.com') {
    $url = $Matches[0]
    break
  }
}

if (-not $url) {
  Write-Host 'No tunnel URL yet'
  Write-Host '--- cf err ---'
  Get-Content $cfErr -ErrorAction SilentlyContinue | Select-Object -Last 50
  Write-Host '--- cf out ---'
  Get-Content $cfOut -ErrorAction SilentlyContinue | Select-Object -Last 50
  exit 2
}

Set-Content -Path (Join-Path $logDir 'public_url.txt') -Value $url
Write-Host "PUBLIC_URL=$url"
