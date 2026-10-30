$ErrorActionPreference = 'Stop'
$deploy = 'C:\Users\WIN10\first-leg-cost-deploy'
Set-Location $deploy

Write-Host '=== Install Python 3.12 (winget) ==='
winget install -e --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements --disable-interactivity
Write-Host '=== Install cloudflared (winget) ==='
winget install -e --id Cloudflare.cloudflared --accept-package-agreements --accept-source-agreements --disable-interactivity

Write-Host '=== Refresh PATH for this session ==='
$env:Path = [System.Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [System.Environment]::GetEnvironmentVariable('Path','User')

Write-Host '=== Versions ==='
python --version
cloudflared --version

Write-Host '=== Extract bundle ==='
# Prefer tar (Windows 10+ has it)
if (Test-Path "$deploy\first-leg-cost") { Remove-Item -Recurse -Force "$deploy\first-leg-cost" }
tar -xzf "$deploy\first-leg-cost-bundle.tgz" -C $deploy
Get-ChildItem "$deploy\first-leg-cost" | Select-Object Name

Write-Host '=== venv + deps ==='
Set-Location "$deploy\first-leg-cost"
python -m venv .venv
& .\.venv\Scripts\python.exe -m pip install -U pip
& .\.venv\Scripts\pip.exe install -r requirements.txt
Write-Host '=== setup done ==='
