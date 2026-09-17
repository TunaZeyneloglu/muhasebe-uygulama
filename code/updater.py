import urllib.request
import webbrowser
from pathlib import Path

import customtkinter as ctk

import state

# ----------------- Sürüm Kontrolü Ayarları -----------------

# version.txt depo kökünde, code/ klasörünün bir üstünde duruyor
YEREL_SURUM_DOSYASI = Path(__file__).resolve().parent.parent / "version.txt"
UZAK_SURUM_URL = "https://raw.githubusercontent.com/TunaZeyneloglu/muhasebe-uygulama/main/version.txt"
SURUMLER_URL = "https://github.com/TunaZeyneloglu/muhasebe-uygulama/releases/latest"
ZAMAN_ASIMI = 3  # saniye

# ----------------- Yardımcı Fonksiyonlar -----------------

def _parse_version(v: str) -> tuple[int, ...]:
    """'1.10.0' -> (1, 10, 0). Sayıya çevrilemezse ValueError fırlatır."""
    return tuple(int(parca) for parca in v.strip().split("."))

def _guncelleme_var_mi(yerel: str, uzak: str) -> bool:
    """Uzak sürüm yerelden büyükse True. Bozuk sürüm metninde False."""
    try:
        return _parse_version(uzak) > _parse_version(yerel)
    except Exception:
        # Bozuk sürüm metni -> güncelleme yok say
        return False

def _yerel_surum() -> str:
    """Depo kökündeki version.txt içeriğini oku"""
    return YEREL_SURUM_DOSYASI.read_text(encoding="utf-8").strip()

def _uzak_surum() -> str:
    """GitHub üzerindeki version.txt içeriğini indir"""
    with urllib.request.urlopen(UZAK_SURUM_URL, timeout=ZAMAN_ASIMI) as yanit:
        return yanit.read().decode("utf-8").strip()

# ----------------- Güncelleme Penceresi -----------------

def _guncelleme_penceresi(yerel: str, uzak: str):
    """Yeni sürüm bulunduğunda gösterilen basit bilgi penceresi"""
    pencere = ctk.CTkToplevel(state.app)
    pencere.title("Güncelleme Mevcut")
    pencere.geometry("380x180")
    pencere.resizable(False, False)
    pencere.transient(state.app)
    pencere.attributes("-topmost", True)

    ctk.CTkLabel(
        pencere,
        text=f"Yeni bir sürüm mevcut: v{uzak}\nMevcut sürümünüz: v{yerel}",
        justify="center",
    ).pack(pady=(30, 20), padx=20)

    ctk.CTkButton(
        pencere,
        text="Güncellemeleri Görüntüle",
        width=200,
        command=lambda: webbrowser.open(SURUMLER_URL),
    ).pack(pady=(0, 8))

    ctk.CTkButton(
        pencere,
        text="Kapat",
        width=200,
        command=pencere.destroy,
    ).pack(pady=(0, 12))

# ----------------- Ana Giriş Noktası -----------------

def check_for_update():
    """Açılışta sürüm kontrolü yapar. Hata olursa sessizce devam eder."""
    try:
        yerel = _yerel_surum()
        uzak = _uzak_surum()
        if _guncelleme_var_mi(yerel, uzak):
            _guncelleme_penceresi(yerel, uzak)
    except Exception as e:
        # Ağ hatası, zaman aşımı, eksik dosya vb. kullanıcıyı ilgilendirmiyor
        print(f"Güncelleme kontrolü başarısız: {e}")
