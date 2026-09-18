# Fatura Yonetim Sistemi - Windows derleme betigi
# Kullanim:
#   Cift tiklama icin derle.bat kullanin
#   powershell -ExecutionPolicy Bypass -File .\derle.ps1
#   powershell -ExecutionPolicy Bypass -File .\derle.ps1 -Temiz   (once build ve dist silinir)
# Tek dogruluk kaynagi FaturaYonetimSistemi.spec dosyasidir.
# Cikti: dist\FaturaYonetimSistemi.exe

param(
    [switch]$Temiz
)

$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot

function Invoke-Native {
    # PS 5.1'de native komut hatasi 'Stop'u tetiklemez; $LASTEXITCODE elle kontrol edilir.
    param(
        [Parameter(Mandatory = $true)][string]$Aciklama,
        [Parameter(Mandatory = $true)][string]$Komut,
        [string[]]$Argumanlar = @()
    )
    Write-Host ">> $Aciklama" -ForegroundColor Cyan
    $eskiTercih = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        & $Komut @Argumanlar
        $kod = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $eskiTercih
    }
    if ($kod -ne 0) {
        throw "HATA: '$Aciklama' basarisiz oldu (cikis kodu: $kod)."
    }
}

# --- Surum ---
$versionDosyasi = Join-Path $PSScriptRoot 'code\version.py'
if (-not (Test-Path -LiteralPath $versionDosyasi)) {
    throw "HATA: $versionDosyasi bulunamadi."
}
$icerik = Get-Content -LiteralPath $versionDosyasi -Raw
$eslesme = [regex]::Match($icerik, 'APP_VERSION\s*=\s*["'']([^"'']+)["'']')
if (-not $eslesme.Success) {
    throw "HATA: code\version.py icinde APP_VERSION bulunamadi."
}
$surum = $eslesme.Groups[1].Value
Write-Host "Surum: $surum" -ForegroundColor Green

# --- Temizlik (istege bagli) ---
if ($Temiz) {
    foreach ($klasor in @('build', 'dist')) {
        $yol = Join-Path $PSScriptRoot $klasor
        if (Test-Path -LiteralPath $yol) {
            Write-Host ">> Siliniyor: $yol" -ForegroundColor Cyan
            Remove-Item -LiteralPath $yol -Recurse -Force
        }
    }
}

# --- Python ---
function Test-Python {
    # Komut calisiyor mu? (stderr ciktisi PS 5.1'de 'Stop' ile hata firlatmasin diye yerel 'Continue')
    param([string]$Komut, [string[]]$Argumanlar = @())
    if (-not (Get-Command $Komut -ErrorAction SilentlyContinue)) { return $false }
    $ErrorActionPreference = 'Continue'
    try {
        & $Komut @Argumanlar --version *> $null
        return ($LASTEXITCODE -eq 0)
    }
    catch {
        return $false
    }
}

$pythonKomut = $null
$pythonArgs = @()
if (Test-Python -Komut 'py' -Argumanlar @('-3')) {
    $pythonKomut = 'py'
    $pythonArgs = @('-3')
}
elseif (Test-Python -Komut 'python') {
    # Windows'taki Microsoft Store yonlendiricisi 'python' calismaz; surum sorgusu bunu eler.
    $pythonKomut = 'python'
}
if (-not $pythonKomut) {
    throw "HATA: Python bulunamadi. https://www.python.org adresinden Python 3 kurun ('py' veya 'python' PATH'te olmali)."
}
Write-Host ("Python: " + (& $pythonKomut @pythonArgs --version 2>&1)) -ForegroundColor Green

# --- Sanal ortam ---
$venvKlasoru = Join-Path $PSScriptRoot '.venv'
if (-not (Test-Path -LiteralPath $venvKlasoru)) {
    Invoke-Native -Aciklama 'Sanal ortam olusturuluyor (.venv)' -Komut $pythonKomut -Argumanlar ($pythonArgs + @('-m', 'venv', $venvKlasoru))
}
$venvPython = Join-Path $venvKlasoru 'Scripts\python.exe'
if (-not (Test-Path -LiteralPath $venvPython)) {
    # Windows disi (test amacli) duzen
    $venvPython = Join-Path $venvKlasoru 'bin/python'
}
if (-not (Test-Path -LiteralPath $venvPython)) {
    throw "HATA: Sanal ortamin python'u bulunamadi: $venvKlasoru. .venv klasorunu silip tekrar deneyin."
}

Invoke-Native -Aciklama 'pip guncelleniyor' -Komut $venvPython -Argumanlar @('-m', 'pip', 'install', '--upgrade', 'pip')
Invoke-Native -Aciklama 'Bagimliliklar kuruluyor' -Komut $venvPython -Argumanlar @('-m', 'pip', 'install', '-r', 'requirements.txt', 'pyinstaller')

# --- Derleme ---
Invoke-Native -Aciklama 'Derleniyor (FaturaYonetimSistemi.spec)' -Komut $venvPython -Argumanlar @('-m', 'PyInstaller', '--noconfirm', '--clean', 'FaturaYonetimSistemi.spec')

# --- Sonuc ---
$exe = Join-Path $PSScriptRoot 'dist\FaturaYonetimSistemi.exe'
if (-not (Test-Path -LiteralPath $exe)) {
    throw "HATA: Derleme bitti ama $exe bulunamadi."
}
$dosya = Get-Item -LiteralPath $exe
$hash = (Get-FileHash -LiteralPath $exe -Algorithm SHA256).Hash

Write-Host ''
Write-Host 'Derleme tamamlandi.' -ForegroundColor Green
Write-Host "Dosya : $($dosya.FullName)"
Write-Host ("Boyut : {0:N1} MB ({1} bayt)" -f ($dosya.Length / 1MB), $dosya.Length)
Write-Host "SHA256: $hash"
Write-Host "Surum : $surum"
Write-Host ''
Write-Host 'HATIRLATMA:' -ForegroundColor Yellow
Write-Host "  - GitHub Release tag'i 'v$surum' olmali."
Write-Host "  - Release'e asset adi tam olarak 'FaturaYonetimSistemi.exe' olacak sekilde yukleyin."
