"""Arka planda otomatik güncelleme.

Kullanıcı güncelleme sürecini yalnızca "Güncelleme Hazır" pop-up'ı ile fark
eder. Kontrol, indirme ve doğrulama arka plan thread'inde yapılır; hatalar
kullanıcıya gösterilmez, yalnızca log dosyasına yazılır. İndirme ve kurulum
sadece Windows'ta PyInstaller ile paketlenmiş exe olarak çalışırken yapılır.
"""
import hashlib
import json
import logging
import os
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request
from logging.handlers import RotatingFileHandler
from pathlib import Path

import customtkinter as ctk

import icons
import state
import theme
from version import APP_VERSION

# ----------------- Güncelleme Ayarları -----------------

REPO = "TunaZeyneloglu/muhasebe-uygulama"
API_URL = f"https://api.github.com/repos/{REPO}/releases/latest"
EXE_ADI = "FaturaYonetimSistemi.exe"  # release asset adı, birebir
ZAMAN_ASIMI = 10          # saniye, ağ işlemleri için
ILK_KONTROL_GECIKMESI = 4000  # ms, açılıştan sonra ilk kontrol
POPUP_TEKRAR_ARALIGI = 3000   # ms, başka modal açıkken pop-up'ı tekrar deneme
TEMIZLIK_DENEME = 5       # eski .old / .part silme deneme sayısı
TEMIZLIK_ARALIGI = 2      # saniye, silme denemeleri arası
PARCA_BOYUTU = 64 * 1024  # indirme okuma parçası

# Windows süreç oluşturma bayrakları (macOS'ta subprocess'te tanımlı değiller)
DETACHED_PROCESS = getattr(subprocess, "DETACHED_PROCESS", 0x00000008)
CREATE_NEW_PROCESS_GROUP = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0x00000200)
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)

# Doğrulanmış ve kullanıcıya gösterilmiş güncelleme ("Sonra" seçilince kapanışta kurulur)
_bekleyen = None
_kapanis_sarildi = False
_logger = None

# ----------------- Log -----------------

def _log_yolu() -> Path:
    """Windows'ta %LOCALAPPDATA% altı, diğerlerinde ev dizini altı."""
    if sys.platform == "win32" and os.environ.get("LOCALAPPDATA"):
        return Path(os.environ["LOCALAPPDATA"]) / "FaturaYonetimSistemi" / "guncelleme.log"
    return Path.home() / ".fatura_yonetim_sistemi" / "guncelleme.log"


def _log() -> logging.Logger:
    """Ayrı bir logger döndürür; root logger'a dokunmaz. Kurulamazsa sessiz kalır."""
    global _logger
    if _logger is not None:
        return _logger
    logger = logging.getLogger("fatura_yonetim.guncelleme")
    logger.setLevel(logging.INFO)
    logger.propagate = False
    try:
        yol = _log_yolu()
        yol.parent.mkdir(parents=True, exist_ok=True)
        handler = RotatingFileHandler(yol, maxBytes=256 * 1024, backupCount=2,
                                      encoding="utf-8")
        handler.setFormatter(logging.Formatter(
            "%(asctime)s %(levelname)s [%(threadName)s] %(message)s"))
        # Yazma hatası stderr'e "Logging error" basmasın
        handler.handleError = lambda kayit: None
        logger.addHandler(handler)
    except Exception:
        logger.addHandler(logging.NullHandler())
    _logger = logger
    return logger

# ----------------- Sürüm Karşılaştırma -----------------

def _surum_ayristir(metin) -> tuple:
    """'v2.9.18' -> (2, 9, 18). Ayrıştırılamazsa ValueError fırlatır."""
    temiz = str(metin).replace("﻿", "").strip()
    if temiz[:1] in ("v", "V"):
        temiz = temiz[1:]
    parcalar = temiz.split(".")
    if not all(p and all(c in "0123456789" for c in p) for p in parcalar):
        raise ValueError(f"geçersiz sürüm: {metin!r}")
    return tuple(int(p) for p in parcalar)


def _daha_yeni_mi(uzak, yerel) -> bool:
    """Uzak sürüm yerelden büyükse True. 2.9 == 2.9.0. Bozuk sürümde False."""
    try:
        a = _surum_ayristir(uzak)
        b = _surum_ayristir(yerel)
    except ValueError as hata:
        _log().warning("Sürüm karşılaştırılamadı: %s", hata)
        return False
    uzunluk = max(len(a), len(b))
    a = a + (0,) * (uzunluk - len(a))
    b = b + (0,) * (uzunluk - len(b))
    return a > b

# ----------------- Yardımcı Fonksiyonlar -----------------

def _etkin_mi() -> bool:
    """İndirme/kurulum yalnızca Windows'ta paketlenmiş exe olarak yapılır."""
    return bool(getattr(sys, "frozen", False)) and sys.platform == "win32"


def _exe_yolu() -> Path:
    return Path(sys.executable)


def _eski_exe_yolu(exe: Path) -> Path:
    return exe.with_name(exe.name + ".old")


def _istek(url, accept):
    return urllib.request.Request(url, headers={
        "User-Agent": f"FaturaYonetimSistemi/{APP_VERSION}",
        "Accept": accept,
    })


def _yazilabilir_mi(klasor: Path) -> bool:
    """Klasörde geçici dosya oluşturup silerek yazma iznini dener."""
    try:
        fd, yol = tempfile.mkstemp(prefix=".yazma_testi_", dir=str(klasor))
        os.close(fd)
        os.remove(yol)
        return True
    except Exception:
        return False


def _dogrula(yol: Path, asset: dict) -> bool:
    """MZ imzası + (digest varsa SHA256, yoksa boyut) kontrolü."""
    try:
        with open(yol, "rb") as f:
            if f.read(2) != b"MZ":
                _log().warning("Doğrulama: MZ imzası yok (%s)", yol.name)
                return False
        digest = asset.get("digest") or ""
        if isinstance(digest, str) and digest.lower().startswith("sha256:"):
            beklenen = digest.split(":", 1)[1].strip().lower()
            ozet = hashlib.sha256()
            with open(yol, "rb") as f:
                for parca in iter(lambda: f.read(PARCA_BOYUTU), b""):
                    ozet.update(parca)
            if ozet.hexdigest() != beklenen:
                _log().warning("Doğrulama: SHA256 eşleşmedi (%s)", yol.name)
                return False
            return True
        boyut = asset.get("size")
        if not isinstance(boyut, int) or yol.stat().st_size != boyut:
            _log().warning("Doğrulama: boyut eşleşmedi (%s, beklenen=%r, gerçek=%s)",
                           yol.name, boyut, yol.stat().st_size)
            return False
        return True
    except Exception:
        _log().exception("Doğrulama sırasında hata")
        return False


def _sil(yol: Path):
    try:
        yol.unlink()
    except FileNotFoundError:
        pass
    except Exception:
        _log().warning("Silinemedi: %s", yol)

# ----------------- Açılış Temizliği -----------------

def _temizlik_thread():
    """Önceki güncellemeden kalan .old ve yarım .part dosyalarını siler.
    Eski süreç henüz kapanmamış olabilir; birkaç kez aralıkla tekrar denenir."""
    try:
        exe = _exe_yolu()
        for deneme in range(TEMIZLIK_DENEME):
            kalanlar = []
            adaylar = [_eski_exe_yolu(exe)] + list(exe.parent.glob(f"{exe.stem}.*.part"))
            for yol in adaylar:
                if not yol.exists():
                    continue
                try:
                    yol.unlink()
                    _log().info("Temizlendi: %s", yol.name)
                except Exception:
                    kalanlar.append(yol)
            if not kalanlar:
                return
            time.sleep(TEMIZLIK_ARALIGI)
        _log().warning("Temizlik tamamlanamadı: %s", ", ".join(y.name for y in kalanlar))
    except Exception:
        _log().exception("Temizlik sırasında hata")

# ----------------- Kontrol + İndirme (arka plan) -----------------

def _release_getir():
    """releases/latest JSON'unu döndürür; release yoksa (404) None."""
    try:
        with urllib.request.urlopen(_istek(API_URL, "application/vnd.github+json"),
                                    timeout=ZAMAN_ASIMI) as yanit:
            return json.loads(yanit.read().decode("utf-8"))
    except urllib.error.HTTPError as hata:
        if hata.code == 404:
            _log().info("Yayınlanmış release yok (404)")
            return None
        raise


def _indir(url: str, hedef: Path):
    """URL'yi hedef dosyaya parça parça indirir."""
    with urllib.request.urlopen(_istek(url, "application/octet-stream"),
                                timeout=ZAMAN_ASIMI) as yanit, open(hedef, "wb") as f:
        for parca in iter(lambda: yanit.read(PARCA_BOYUTU), b""):
            f.write(parca)


def _kontrol_thread():
    """Sürüm kontrolü, indirme ve doğrulama. Hazırsa pop-up'ı ana thread'e iletir."""
    try:
        if not _etkin_mi():
            _log().info("Güncelleme kontrolü atlandı (paketlenmiş Windows exe değil: "
                        "frozen=%s, platform=%s)", getattr(sys, "frozen", False), sys.platform)
            return

        release = _release_getir()
        if not release:
            return
        if release.get("draft") or release.get("prerelease"):
            _log().info("Taslak/ön sürüm yoksayıldı: %s", release.get("tag_name"))
            return

        etiket = str(release.get("tag_name") or "")
        if not _daha_yeni_mi(etiket, APP_VERSION):
            _log().info("Güncel sürüm kullanılıyor (yerel=%s, uzak=%s)", APP_VERSION, etiket)
            return
        surum = ".".join(str(p) for p in _surum_ayristir(etiket))

        asset = next((a for a in release.get("assets") or []
                      if isinstance(a, dict) and a.get("name") == EXE_ADI), None)
        if asset is None or not asset.get("browser_download_url"):
            _log().warning("Release %s içinde %s bulunamadı", etiket, EXE_ADI)
            return

        exe = _exe_yolu()
        if not _yazilabilir_mi(exe.parent):
            _log().warning("Exe klasörü yazılabilir değil, güncelleme atlandı: %s", exe.parent)
            return

        yeni = exe.parent / f"{exe.stem}.{surum}.new"
        if yeni.exists() and _dogrula(yeni, asset):
            _log().info("Geçerli %s zaten mevcut, yeniden indirilmedi", yeni.name)
        else:
            _sil(yeni)
            parca = exe.parent / f"{exe.stem}.{surum}.part"
            _log().info("İndiriliyor: %s -> %s", asset["browser_download_url"], parca.name)
            try:
                _indir(asset["browser_download_url"], parca)
            except Exception:
                _sil(parca)
                raise
            if not _dogrula(parca, asset):
                _sil(parca)
                _log().warning("İndirilen dosya doğrulanamadı, silindi")
                return
            os.replace(parca, yeni)
            _log().info("İndirildi ve doğrulandı: %s", yeni.name)

        body = release.get("body") if isinstance(release.get("body"), str) else ""
        bilgi = {"surum": surum, "dosya": yeni, "notlar": body}
        from pages import popup_menu
        popup_menu.ana_threadde_calistir(state.app, lambda: _popup_goster_dene(bilgi))
    except Exception:
        _log().exception("Güncelleme kontrolü başarısız")

# ----------------- Güncelleme Penceresi (ana thread) -----------------

def _popup_goster_dene(bilgi):
    """Başka bir modal/menü açıksa birkaç saniye sonra tekrar dener."""
    try:
        if not bilgi["dosya"].exists():
            _log().warning("Hazır dosya kaybolmuş: %s", bilgi["dosya"])
            return
        if state.app.grab_current() is not None or state.popup_aktif:
            state.app.after(POPUP_TEKRAR_ARALIGI, lambda: _popup_goster_dene(bilgi))
            return
        _guncelleme_penceresi(bilgi)
    except Exception:
        _log().exception("Güncelleme penceresi gösterilemedi")


def _not_ozeti(body: str) -> str:
    """Release notunun ilk 6 satırı / 400 karakteri."""
    satirlar = [s.rstrip() for s in (body or "").strip().splitlines()]
    ozet = "\n".join(satirlar[:6])
    kirpildi = len(satirlar) > 6
    if len(ozet) > 400:
        ozet = ozet[:400].rstrip()
        kirpildi = True
    return ozet + ("…" if kirpildi else "")


def _guncelleme_penceresi(bilgi):
    """Doğrulanmış güncelleme dosyası hazırken gösterilen modal pencere."""
    from pages.kontrol_popups import _popup_basligi

    popup = ctk.CTkToplevel(state.app)
    popup.title("Güncelleme Hazır")
    popup.geometry("440x300")
    popup.resizable(False, False)
    popup.attributes("-topmost", True)
    popup.grab_set()  # Modal yap
    popup.configure(fg_color=theme.BG_ROOT)

    _popup_basligi(popup, "Güncelleme Hazır",
                   f"Yeni sürüm v{bilgi['surum']} (mevcut v{APP_VERSION})",
                   theme.ACCENT, icons.onay)

    mesaj_karti = ctk.CTkFrame(popup, fg_color=theme.BG_SURFACE, corner_radius=theme.CORNER_POPUP,
                               border_width=theme.BORDER_WIDTH, border_color=theme.BORDER)
    mesaj_karti.pack(fill="both", expand=True, padx=26, pady=(0, 18))
    ctk.CTkLabel(mesaj_karti,
                 text="Güncelleme için uygulama kısa süreliğine kapanıp yeniden açılacak.",
                 font=theme.FONT_MESSAGE(), text_color=theme.TEXT_SECONDARY,
                 justify="left", wraplength=350).pack(padx=16, pady=(14, 6), anchor="w")
    ozet = _not_ozeti(bilgi.get("notlar"))
    if ozet:
        ctk.CTkLabel(mesaj_karti, text=ozet, font=theme.FONT_SMALL(),
                     text_color=theme.TEXT_MUTED, justify="left",
                     wraplength=350).pack(padx=16, pady=(0, 14), anchor="w")

    def simdi():
        try:
            popup.destroy()
        except Exception:
            pass
        _simdi_guncelle(bilgi)

    def sonra():
        try:
            popup.destroy()
        except Exception:
            pass
        _sonraya_birak(bilgi)

    btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
    btn_frame.pack(pady=(0, 20))
    ctk.CTkButton(btn_frame, text="Şimdi Güncelle", width=theme.BTN_W_MD, height=theme.BTN_H_MD,
                  corner_radius=theme.CORNER_BTN, font=theme.FONT_BODY_BOLD(),
                  fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                  text_color=theme.TEXT_ON_ACCENT, command=simdi).pack(side="left", padx=5)
    ctk.CTkButton(btn_frame, text="Sonra", width=theme.BTN_W_SM, height=theme.BTN_H_MD,
                  corner_radius=theme.CORNER_BTN, font=theme.FONT_SMALL(),
                  fg_color="transparent", border_width=theme.BORDER_WIDTH, border_color=theme.BORDER,
                  hover_color=theme.BG_HOVER, text_color=theme.TEXT_SECONDARY,
                  command=sonra).pack(side="left", padx=5)

    # Pencere kapatma (X) = "Sonra"
    popup.protocol("WM_DELETE_WINDOW", sonra)

    # Yükseklik içerikten hesaplanır, pencere ana uygulamaya göre ortalanır
    popup.update_idletasks()
    height = max(300, popup.winfo_reqheight())
    x = state.app.winfo_x() + (state.app.winfo_width() // 2) - 220
    y = state.app.winfo_y() + (state.app.winfo_height() // 2) - (height // 2)
    popup.geometry(f"440x{height}+{x}+{y}")
    _log().info("Güncelleme penceresi gösterildi (v%s)", bilgi["surum"])

# ----------------- Kurulum -----------------

def _kur(yeni: Path) -> bool:
    """Çalışan exe'yi .old'a taşır, yeni dosyayı yerine koyar. Hata olursa geri alır."""
    try:
        exe = _exe_yolu()
        eski = _eski_exe_yolu(exe)
        if not yeni.exists():
            _log().warning("Kurulacak dosya yok: %s", yeni)
            return False
        if eski.exists():
            _sil(eski)
        # Windows çalışan exe'nin yeniden adlandırılmasına izin verir
        os.replace(exe, eski)
        try:
            os.replace(yeni, exe)
        except Exception:
            _log().exception("Yeni exe yerine konamadı, geri alınıyor")
            try:
                os.replace(eski, exe)
            except Exception:
                _log().exception("Geri alma başarısız")
            return False
        _log().info("Güncelleme kuruldu: %s", yeni.name)
        return True
    except Exception:
        _log().exception("Kurulum başarısız")
        return False


def _temiz_ortam() -> dict:
    """PyInstaller'ın çocuk sürece aktarmaması gereken değişkenlerden arındırılmış ortam."""
    ortam = os.environ.copy()
    for anahtar in list(ortam):
        if anahtar.upper() == "_MEIPASS2" or anahtar.upper().startswith("_PYI_"):
            del ortam[anahtar]
    # Tcl/Tk yolları eski sürecin geçici klasörünü gösteriyorsa taşıma
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        for anahtar in ("TCL_LIBRARY", "TK_LIBRARY"):
            deger = ortam.get(anahtar)
            if deger and os.path.normcase(deger).startswith(os.path.normcase(meipass)):
                del ortam[anahtar]
    ortam["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
    return ortam


def _yeniden_baslat() -> bool:
    """Yeni exe'yi konsol penceresi açmadan, bağımsız süreç olarak başlatır."""
    try:
        exe = _exe_yolu()
        subprocess.Popen(
            [str(exe)], cwd=str(exe.parent), close_fds=True, env=_temiz_ortam(),
            stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            creationflags=DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP | CREATE_NO_WINDOW,
        )
        _log().info("Yeni sürüm başlatıldı")
        return True
    except Exception:
        _log().exception("Yeniden başlatma başarısız")
        return False


def _simdi_guncelle(bilgi):
    """'Şimdi Güncelle': kur, yeni süreci başlat, uygulamayı kapat."""
    try:
        if not _kur(bilgi["dosya"]):
            return  # Kurulamadı: uygulama normal çalışmaya devam eder
        if _yeniden_baslat():
            state.app.destroy()
        # Başlatılamadıysa yeni exe yerinde; bir sonraki açılış yeni sürümle olur
    except Exception:
        _log().exception("Şimdi güncelle başarısız")


def _sonraya_birak(bilgi):
    """'Sonra': dosya hazır kalır, uygulama kapanırken sessizce kurulur."""
    global _bekleyen, _kapanis_sarildi
    try:
        _bekleyen = bilgi
        _log().info("Güncelleme kapanışa ertelendi (v%s)", bilgi["surum"])
        if _kapanis_sarildi:
            return
        onceki = state.app.protocol("WM_DELETE_WINDOW")  # mevcut Tcl komutu ya da ""

        def kapanirken():
            try:
                if _bekleyen is not None:
                    _kur(_bekleyen["dosya"])
            except Exception:
                _log().exception("Kapanışta kurulum başarısız")
            if onceki:
                try:
                    state.app.tk.call(onceki)
                    return
                except Exception:
                    _log().exception("Önceki kapanış işleyicisi çalıştırılamadı")
            state.app.destroy()

        state.app.protocol("WM_DELETE_WINDOW", kapanirken)
        _kapanis_sarildi = True
    except Exception:
        _log().exception("Kapanış işleyicisi kurulamadı")

# ----------------- Ana Giriş Noktası -----------------

def _thread_baslat():
    try:
        if _etkin_mi():
            threading.Thread(target=_temizlik_thread, name="guncelleme-temizlik",
                             daemon=True).start()
        threading.Thread(target=_kontrol_thread, name="guncelleme-kontrol",
                         daemon=True).start()
    except Exception:
        _log().exception("Güncelleme thread'i başlatılamadı")


def baslat():
    """main.py'den çağrılır: kontrolü açılıştan kısa süre sonraya planlar ve hemen döner."""
    try:
        state.app.after(ILK_KONTROL_GECIKMESI, _thread_baslat)
    except Exception:
        _log().exception("Güncelleme kontrolü planlanamadı")
