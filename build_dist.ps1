<#
Dong goi outlook-mcp thanh file zip de mang sang may khac.

    powershell -ExecutionPolicy Bypass -File .\build_dist.ps1

Ket qua: dist\outlook-mcp-<version>-windows.zip gom file .whl, install.ps1, README.md.
Can Internet lan dau (pip tai setuptools de build wheel).
#>
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$py = Join-Path $here ".venv\Scripts\python.exe"
if (-not (Test-Path $py)) { $py = "python" }

$version = (& $py -c "import sys; sys.path.insert(0, r'$here'); import outlook_mcp; print(outlook_mcp.__version__)").Trim()
$dist = Join-Path $here "dist"
$stage = Join-Path $dist "outlook-mcp-$version"
if (Test-Path $stage) { Remove-Item -Recurse -Force $stage }
New-Item -ItemType Directory -Force -Path $stage | Out-Null

& $py -m pip wheel $here --no-deps --disable-pip-version-check -w $stage
if ($LASTEXITCODE) { throw "Build wheel that bai" }
# don thu muc tam setuptools sinh ra trong ma nguon (giu egg-info cua ban cai -e trong .venv)
$p = Join-Path $here "build"
if (Test-Path $p) { Remove-Item -Recurse -Force $p }
Copy-Item (Join-Path $here "install.ps1"), (Join-Path $here "README.md"), (Join-Path $here "constraints.txt") $stage

# checksum tung file trong goi (install.ps1 tu kiem truoc khi cai)
$lines = Get-ChildItem $stage -File | Where-Object { $_.Name -ne "SHA256SUMS.txt" } | Sort-Object Name |
    ForEach-Object { "{0}  {1}" -f (Get-FileHash $_.FullName -Algorithm SHA256).Hash.ToLower(), $_.Name }
[IO.File]::WriteAllLines((Join-Path $stage "SHA256SUMS.txt"), $lines, (New-Object Text.UTF8Encoding $false))

$zip = Join-Path $dist "outlook-mcp-$version-windows.zip"
if (Test-Path $zip) { Remove-Item -Force $zip }
Compress-Archive -Path (Join-Path $stage "*") -DestinationPath $zip
# checksum cua zip: gui kem qua kenh khac (chat noi bo) de nguoi nhan doi chieu
$zipHash = (Get-FileHash $zip -Algorithm SHA256).Hash.ToLower()
[IO.File]::WriteAllText("$zip.sha256", "$zipHash  $(Split-Path $zip -Leaf)`n", (New-Object Text.UTF8Encoding $false))
Write-Host "Da dong goi: $zip"
Write-Host "SHA256: $zipHash"
Get-ChildItem $stage | Format-Table Name, Length -AutoSize
