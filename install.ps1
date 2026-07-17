# =============================================================================
# DOffice Auto - CAI DAT NHANH TREN WINDOWS (KHONG can Docker).
#
# Chay 1 lenh trong PowerShell tai thu muc project:
#     powershell -ExecutionPolicy Bypass -File .\install.ps1
#
# Script se: tao moi truong ao .venv, cai thu vien Python, cai trinh duyet
# Chromium cho Playwright. Sau do in ra 2 buoc tiep theo (dang nhap + chay web).
#
# Yeu cau: da cai Python 3.10+ (https://www.python.org/downloads/ - nho tick
# "Add python.exe to PATH" khi cai).
# =============================================================================
#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
Set-Location -Path $PSScriptRoot

Write-Host "== DOffice Auto - cai dat tren Windows ==" -ForegroundColor Cyan

# 1) Tim Python (uu tien Python Launcher 'py -3', roi 'python').
$pyExe = $null
$pyArgs = @()
if (Get-Command py -ErrorAction SilentlyContinue) {
    $pyExe = 'py'; $pyArgs = @('-3')
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    $pyExe = 'python'
} elseif (Get-Command python3 -ErrorAction SilentlyContinue) {
    $pyExe = 'python3'
}
if (-not $pyExe) {
    Write-Host "KHONG tim thay Python. Cai Python 3.10+ tai https://www.python.org/downloads/" -ForegroundColor Red
    Write-Host "(Nho tick 'Add python.exe to PATH' khi cai) roi chay lai install.ps1." -ForegroundColor Red
    exit 1
}
Write-Host "Dung Python: $pyExe $($pyArgs -join ' ')"

# 2) Tao moi truong ao .venv (neu chua co).
if (-not (Test-Path ".venv\Scripts\python.exe")) {
    Write-Host "Tao moi truong ao .venv ..."
    & $pyExe @pyArgs -m venv .venv
}
$vpy = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $vpy)) {
    Write-Host "Tao .venv that bai. Kiem tra lai ban cai Python." -ForegroundColor Red
    exit 1
}

# 3) Cai thu vien Python.
Write-Host "Cai thu vien Python (flask, openpyxl, playwright) ..."
& $vpy -m pip install --upgrade pip
& $vpy -m pip install -r requirements.txt

# 4) Cai trinh duyet Chromium cho Playwright (co the mat vai phut lan dau).
Write-Host "Cai trinh duyet Chromium cho Playwright (co the mat vai phut) ..."
& $vpy -m playwright install chromium

Write-Host ""
Write-Host "== XONG CAI DAT ==" -ForegroundColor Green
Write-Host "Cac buoc tiep theo:" -ForegroundColor Yellow
Write-Host "  1) Dang nhap DOffice 1 lan (mo cua so Chromium de ban dang nhap, luu phien):" -ForegroundColor Yellow
Write-Host "       .\.venv\Scripts\python.exe login_save_state.py"
Write-Host "  2) Chay bang dieu khien web (roi mo http://127.0.0.1:8877):" -ForegroundColor Yellow
Write-Host "       .\.venv\Scripts\python.exe run_web.py"
Write-Host "  3) (Tuy chon) De web dashboard TU CHAY NGAM moi khi dang nhap Windows, khong"
Write-Host "     can tu tay go 'run_web.py' moi lan:" -ForegroundColor Yellow
Write-Host "       powershell -ExecutionPolicy Bypass -File .\install_web_startup.ps1"
Write-Host "     (Sau do vao tab 'Lich chay' tren web de dat gio chay tu dong - se tu tao/"
Write-Host "     cap nhat 1 Windows Scheduled Task rieng, chay duoc du dashboard dang tat.)"
Write-Host ""
Write-Host "Hoac chay bang dong lenh: .\.venv\Scripts\python.exe run_doffice.py --all" -ForegroundColor DarkGray
