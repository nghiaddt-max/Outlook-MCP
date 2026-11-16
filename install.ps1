<#
Cai dat outlook-mcp tren may Windows.

Chay trong thu muc giai nen (cung cho voi file outlook_mcp-*.whl):
    powershell -ExecutionPolicy Bypass -File .\install.ps1

Tuy chon:
    -InstallDir <thu muc>   Noi cai (mac dinh %LOCALAPPDATA%\outlook-mcp)
    -Python <python.exe>    Chi dinh Python neu khong dung py launcher
    -Wheelhouse <thu muc>   Cai offline tu thu muc chua san cac file .whl phu thuoc
    -Force                  Xoa moi truong cu trong InstallDir roi cai lai

Yeu cau: Outlook co dien (Office) da cai va co profile mail. New Outlook khong ho tro COM.
#>
param(
    [string]$InstallDir = (Join-Path $env:LOCALAPPDATA "outlook-mcp"),
    [string]$Python = "",
    [string]$Wheelhouse = "",
    [switch]$Force
)
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

$wheel = Get-ChildItem -Path $here -Filter "outlook_mcp-*.whl" | Sort-Object Name | Select-Object -Last 1
if (-not $wheel) { throw "Khong thay file outlook_mcp-*.whl canh install.ps1" }
$constraints = Join-Path $here "constraints.txt"
if (-not (Test-Path $constraints)) { throw "Thieu constraints.txt canh install.ps1" }

# --- kiem checksum cac file trong goi
$sums = Join-Path $here "SHA256SUMS.txt"
if (-not (Test-Path $sums)) { throw "Thieu SHA256SUMS.txt - goi khong day du, khong cai." }
foreach ($line in Get-Content $sums) {
    if ($line -notmatch '^([0-9a-f]{64})\s+(.+)$') { continue }
    $f = Join-Path $here $Matches[2]
    if (-not (Test-Path $f)) { throw "Thieu file $($Matches[2]) trong goi." }
    $h = (Get-FileHash $f -Algorithm SHA256).Hash.ToLower()
    if ($h -ne $Matches[1]) { throw "Checksum sai: $($Matches[2]). Goi da bi thay doi, khong cai." }
}
Write-Host "Checksum goi: OK"

# --- Outlook co dien (canh bao, khong chan: co the cai truoc roi moi cai Outlook)
$clsid = (Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\Outlook.Application\CLSID" -ErrorAction SilentlyContinue).'(default)'
if (-not $clsid) {
    Write-Warning "Chua thay Outlook co dien (Outlook.Application). Server chi chay duoc khi co Outlook ban Office."
}

# --- Python >= 3.10
if (-not $Python) {
    if (Get-Command py -ErrorAction SilentlyContinue) {
        $Python = (& py -3 -c "import sys; print(sys.executable)").Trim()
    } elseif (Get-Command python -ErrorAction SilentlyContinue) {
        $Python = (Get-Command python).Source
    } else {
        throw "Chua cai Python. Tai Python 3.10 tro len tai python.org (tick 'Add python.exe to PATH') roi chay lai."
    }
}
$ver = (& $Python -c "import sys; print('%d.%d' % sys.version_info[:2])").Trim()
if ([version]$ver -lt [version]"3.10") { throw "Can Python 3.10 tro len, dang co $ver ($Python)" }
Write-Host "Python $ver : $Python"

# --- moi truong ao rieng
$venv = Join-Path $InstallDir ".venv"
if (Test-Path $venv) {
    if ($Force) {
        Write-Host "Xoa moi truong cu: $venv"
        Remove-Item -Recurse -Force $venv
    } else {
        Write-Host "Da co $venv - cai de len (dung -Force de cai lai tu dau)"
    }
}
New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
if (-not (Test-Path $venv)) { & $Python -m venv $venv; if ($LASTEXITCODE) { throw "Tao venv that bai" } }
$py = Join-Path $venv "Scripts\python.exe"

# --- cai goi
$pipArgs = @("-m", "pip", "install", "--disable-pip-version-check", "--force-reinstall", "--no-deps", $wheel.FullName)
& $py @pipArgs
if ($LASTEXITCODE) { throw "pip install goi outlook-mcp that bai" }
# phu thuoc: dung phien ban da kiem thu, chi nhan wheel (khong chay setup.py cua goi nguon)
$depArgs = @("-m", "pip", "install", "--disable-pip-version-check", "--only-binary=:all:",
             "-c", $constraints, $wheel.FullName)
if ($Wheelhouse) { $depArgs += @("--no-index", "--find-links", $Wheelhouse) }
& $py @depArgs
if ($LASTEXITCODE) { throw "pip install phu thuoc that bai (kiem tra Internet hoac dung -Wheelhouse)" }

$exe = Join-Path $venv "Scripts\outlook-mcp.exe"
& $exe --version
if ($LASTEXITCODE) { throw "Khong chay duoc $exe" }

# --- file cau hinh MCP mau (UTF-8 khong BOM)
$cfg = [ordered]@{ mcpServers = [ordered]@{ outlook = [ordered]@{ command = $exe; args = @() } } }
$json = $cfg | ConvertTo-Json -Depth 5
$cfgPath = Join-Path $InstallDir "mcp-config.json"
[IO.File]::WriteAllText($cfgPath, $json, (New-Object Text.UTF8Encoding $false))

Write-Host ""
Write-Host "CAI DAT XONG."
Write-Host "  Lenh server : $exe"
Write-Host "  Cau hinh mau: $cfgPath"
Write-Host ""
Write-Host "Kiem tra ket noi (Outlook co dien phai dang nhap san profile mail):"
Write-Host "  & `"$exe`" --check"
Write-Host ""
Write-Host "Dang ky voi Claude Code (dung cho moi du an):"
Write-Host "  claude mcp add --scope user outlook -- `"$exe`""
Write-Host "Hoac chep noi dung $cfgPath vao .mcp.json cua du an / claude_desktop_config.json."
