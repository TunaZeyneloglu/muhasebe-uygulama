#!/bin/bash
# macOS (Apple Silicon) derlemesi: Finder'da çift tıklayın.
# Çıktı: dist/FaturaYonetimSistemi.app ve dist/FaturaYonetimSistemi-macOS.zip
# Argümanlar:
#   --etkilesimsiz  Sonunda Enter beklemez (CI ortam değişkeni de aynı etkiyi yapar)
#   --temiz         Uyumluluk için kabul edilir; temizlik her zaman yapılır
cd "$(dirname "$0")"
set -euo pipefail

ETKILESIMLI=1
if [ -n "${CI:-}" ]; then
    ETKILESIMLI=0
fi
for arg in "$@"; do
    case "$arg" in
        --etkilesimsiz) ETKILESIMLI=0 ;;
        --temiz) ;;
        *) echo "Bilinmeyen argüman: $arg" ;;
    esac
done

bekle() {
    if [ "$ETKILESIMLI" = "1" ]; then
        echo
        read -r -p "Kapatmak için Enter'a basın..." _ || true
    fi
}

hata() {
    echo
    echo "HATA: $1"
    trap - EXIT
    bekle
    exit 1
}

trap 'kod=$?; if [ $kod -ne 0 ]; then echo; echo "HATA: Derleme başarısız oldu (çıkış kodu $kod). Yukarıdaki çıktıya bakın."; bekle; fi' EXIT

SURUM="$(sed -n 's/^APP_VERSION = "\(.*\)"$/\1/p' code/version.py)"
[ -n "$SURUM" ] || hata "code/version.py içinden sürüm okunamadı."
echo "Fatura Yönetim Sistemi v$SURUM - macOS derlemesi"
echo

if [ ! -x ".venv-mac/bin/python" ]; then
    command -v python3 >/dev/null 2>&1 || hata "python3 bulunamadı. python.org'dan Python 3 kurun ve tekrar deneyin."
    echo "Sanal ortam oluşturuluyor (.venv-mac)..."
    python3 -m venv .venv-mac
fi
PY=".venv-mac/bin/python"

echo "Bağımlılıklar kuruluyor..."
"$PY" -m pip install --upgrade pip
"$PY" -m pip install -r requirements.txt pyinstaller

echo "Eski derleme çıktıları temizleniyor..."
rm -rf build dist/FaturaYonetimSistemi.app dist/FaturaYonetimSistemi

echo "Derleniyor..."
"$PY" -m PyInstaller --noconfirm --clean FaturaYonetimSistemi.spec

APP="dist/FaturaYonetimSistemi.app"
ZIP="dist/FaturaYonetimSistemi-macOS.zip"
[ -d "$APP" ] || hata "$APP oluşmadı."

# python.org Tcl.framework'ünden kopyalanan dosyalar karantina özniteliği taşıyabilir
xattr -dr com.apple.quarantine "$APP" 2>/dev/null || true

rm -f "$ZIP"
ditto -c -k --keepParent "$APP" "$ZIP"

echo
echo "Derleme tamamlandı: $APP"
echo "Zip: $ZIP"
echo "Boyut: $(du -h "$ZIP" | cut -f1) ($(stat -f %z "$ZIP") bayt)"
echo "SHA256: $(shasum -a 256 "$ZIP" | cut -d' ' -f1)"
echo
echo "Release'e eklemek için:"
echo "  gh release upload v$SURUM $ZIP --clobber"

trap - EXIT
bekle
