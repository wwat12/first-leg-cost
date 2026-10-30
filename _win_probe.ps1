$ErrorActionPreference = 'Continue'
Write-Host '=== python ==='
Get-Command python, py -ErrorAction SilentlyContinue | Format-List Name, Source
try { python --version } catch { Write-Host "python fail: $_" }
try { py -3 --version } catch { Write-Host "py fail: $_" }
Write-Host '=== winget ==='
winget --version
Write-Host '=== cloudflared ==='
Get-Command cloudflared -ErrorAction SilentlyContinue | Format-List
Write-Host '=== local programs ==='
Get-ChildItem "$env:LOCALAPPDATA\Programs" -ErrorAction SilentlyContinue | Select-Object Name
Write-Host '=== python dirs ==='
Get-ChildItem "$env:LOCALAPPDATA\Programs\Python" -Recurse -Filter python.exe -ErrorAction SilentlyContinue | Select-Object -First 5 FullName
Write-Host '=== done ==='
