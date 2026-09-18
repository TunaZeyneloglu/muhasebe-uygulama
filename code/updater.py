"""Arka planda otomatik güncelleme.

Kullanıcı güncelleme sürecini yalnızca "Güncelleme Hazır" pop-up'ı ile fark
eder. Kontrol, indirme ve doğrulama arka plan thread'inde yapılır; hatalar
kullanıcıya gösterilmez, yalnızca log dosyasına yazılır. İndirme ve kurulum
sadece PyInstaller ile paketlenmiş olarak çalışırken yapılır: Windows'ta exe,
macOS'ta .app paketi.
"""
import hashlib
import json
import logging
import os
import shutil
import ssl
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
MAC_ZIP_ADI = "FaturaYonetimSistemi-macOS.zip"  # macOS release asset adı, birebir
MAC_APP_ADI = "FaturaYonetimSistemi.app"        # zip kökündeki .app paketi
MAC_CALISTIRILABILIR = "FaturaYonetimSistemi"   # .app/Contents/MacOS altındaki dosya
DITTO_ZAMAN_ASIMI = 300   # saniye, macOS'ta zip açma için
ZAMAN_ASIMI = 10          # saniye, ağ işlemleri için
ILK_KONTROL_GECIKMESI = 500   # ms, açılıştan sonra ilk kontrol
POPUP_TEKRAR_ARALIGI = 3000   # ms, başka modal açıkken pop-up'ı tekrar deneme
TEMIZLIK_DENEME = 5       # eski .old / .part silme deneme sayısı
TEMIZLIK_ARALIGI = 2      # saniye, silme denemeleri arası
PARCA_BOYUTU = 64 * 1024  # indirme okuma parçası
ILERLEME_ARALIGI = 0.1    # saniye, ilerleme güncellemeleri arası en az süre (~10/sn)
BITTI_DENEME = 20         # pencere kurulurken indirme biterse sonucu işleme deneme sayısı

# Windows süreç oluşturma bayrakları (macOS'ta subprocess'te tanımlı değiller)
DETACHED_PROCESS = getattr(subprocess, "DETACHED_PROCESS", 0x00000008)
CREATE_NEW_PROCESS_GROUP = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0x00000200)
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)

# Doğrulanmış ve kullanıcıya gösterilmiş güncelleme ("Sonra" seçilince kapanışta kurulur)
_bekleyen = None
_kapanis_sarildi = False
_translocation_loglandi = False  # translocation uyarısı yalnızca bir kez yazılır
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
    """İndirme/kurulum yalnızca paketlenmiş olarak yapılır: Windows'ta exe, macOS'ta .app.
    macOS'ta App Translocation altından (salt okunur rastgele yol) çalışırken yapılmaz."""
    if not getattr(sys, "frozen", False):
        return False
    if sys.platform == "win32":
        return True
    if sys.platform == "darwin":
        if "/AppTranslocation/" in str(sys.executable):
            global _translocation_loglandi
            if not _translocation_loglandi:
                _translocation_loglandi = True
                _log().info("Güncelleme devre dışı: uygulama AppTranslocation altından "
                            "çalışıyor (%s)", sys.executable)
            return False
        return True
    return False


def _exe_yolu() -> Path:
    return Path(sys.executable)


def _hedef_yolu() -> Path:
    """Güncellemenin yerine konacağı yol: Windows'ta exe, macOS'ta .app paketi."""
    if sys.platform == "darwin":
        # .../FaturaYonetimSistemi.app/Contents/MacOS/FaturaYonetimSistemi -> .app
        return Path(sys.executable).resolve().parents[2]
    return _exe_yolu()


def _eski_exe_yolu(exe: Path) -> Path:
    return exe.with_name(exe.name + ".old")


def _istek(url, accept):
    return urllib.request.Request(url, headers={
        "User-Agent": f"FaturaYonetimSistemi/{APP_VERSION}",
        "Accept": accept,
    })


def _ssl_baglami():
    """macOS'ta certifi CA paketiyle SSL bağlamı; diğer platformlarda None.

    python.org Python'unun OpenSSL varsayılan cafile'ı paketlenmiş .app'te yok.
    certifi yüklenemezse log'a yazılır ve varsayılan bağlama düşülür (None)."""
    if sys.platform != "darwin":
        return None
    try:
        import certifi
        cafile = certifi.where()
        if not os.path.isfile(cafile):
            raise FileNotFoundError(cafile)
        return ssl.create_default_context(cafile=cafile)
    except Exception as hata:
        _log().warning("certifi CA paketi kullanılamadı, varsayılan SSL bağlamı "
                       "kullanılıyor: %r", hata)
        return None


def _urlopen(istek):
    """urlopen sarmalayıcısı. Windows'ta context geçirilmez (sistem deposu).
    Yönlendirmeler (objects.githubusercontent.com) aynı bağlamla doğrulanır."""
    baglam = _ssl_baglami()
    if baglam is None:
        return urllib.request.urlopen(istek, timeout=ZAMAN_ASIMI)
    return urllib.request.urlopen(istek, timeout=ZAMAN_ASIMI, context=baglam)


def _yazilabilir_mi(klasor: Path) -> bool:
    """Klasörde geçici dosya oluşturup silerek yazma iznini dener."""
    try:
        fd, yol = tempfile.mkstemp(prefix=".yazma_testi_", dir=str(klasor))
        os.close(fd)
        os.remove(yol)
        return True
    except Exception:
        return False


def _dogrula(yol: Path, asset: dict, imza: bytes = b"MZ") -> bool:
    """Dosya imzası (Windows exe: MZ, macOS zip: PK\\x03\\x04) + (digest varsa SHA256,
    yoksa boyut) kontrolü."""
    try:
        with open(yol, "rb") as f:
            if f.read(len(imza)) != imza:
                _log().warning("Doğrulama: %r imzası yok (%s)", imza, yol.name)
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


def _sil_agac(yol: Path):
    """Klasörü (içeriğiyle) ya da dosyayı/symlink'i siler; hata yalnızca log'a yazılır."""
    try:
        if yol.is_symlink() or yol.is_file():
            yol.unlink()
        elif yol.exists():
            shutil.rmtree(yol)
    except Exception:
        _log().warning("Silinemedi: %s", yol)


def _mac_hazir_mi(staging: Path) -> bool:
    """macOS: açılmış klasörde .app/Contents/MacOS/<çalıştırılabilir> var mı."""
    try:
        return (staging / MAC_APP_ADI / "Contents" / "MacOS" / MAC_CALISTIRILABILIR).is_file()
    except Exception:
        return False


def _mac_zip_ac(zip_yolu: Path, staging: Path) -> bool:
    """macOS: zip'i ditto ile staging klasörüne açar (symlink ve izinler korunur).
    Başarılıysa zip silinir; geçersizse staging silinir."""
    _sil_agac(staging)
    try:
        subprocess.run(["/usr/bin/ditto", "-x", "-k", str(zip_yolu), str(staging)],
                       stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                       stderr=subprocess.DEVNULL, check=True, timeout=DITTO_ZAMAN_ASIMI)
    except Exception:
        _log().exception("Zip açılamadı: %s", zip_yolu.name)
        _sil_agac(staging)
        return False
    if not _mac_hazir_mi(staging):
        _log().warning("Açılan zip geçersiz (%s/%s/Contents/MacOS/%s yok), silindi",
                       staging.name, MAC_APP_ADI, MAC_CALISTIRILABILIR)
        _sil_agac(staging)
        return False
    try:
        subprocess.run(["/usr/bin/xattr", "-dr", "com.apple.quarantine",
                        str(staging / MAC_APP_ADI)],
                       stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                       stderr=subprocess.DEVNULL, check=False, timeout=60)
    except Exception:
        pass
    _sil(zip_yolu)
    return True

# ----------------- Açılış Temizliği -----------------

def _temizlik_thread_mac():
    """macOS: önceki güncellemeden kalan .app.old klasörünü, güncel sürümden yeni
    olmayan .new klasörlerini ve yarım .zip.part dosyalarını siler."""
    try:
        app = _hedef_yolu()
        adaylar = [_eski_exe_yolu(app)] + list(app.parent.glob(f"{app.stem}.*.zip.part"))
        onek, sonek = f"{app.stem}.", ".new"
        for yol in app.parent.glob(f"{app.stem}.*.new"):
            try:
                surum = yol.name[len(onek):-len(sonek)]
                _surum_ayristir(surum)  # ayrıştırılamayan adlara dokunma
                if not _daha_yeni_mi(surum, APP_VERSION):
                    adaylar.append(yol)
            except Exception:
                pass
        for yol in adaylar:
            if not (yol.exists() or yol.is_symlink()):
                continue
            _sil_agac(yol)
            if not (yol.exists() or yol.is_symlink()):
                _log().info("Temizlendi: %s", yol.name)
    except Exception:
        _log().exception("Temizlik sırasında hata")


def _temizlik_thread():
    """Önceki güncellemeden kalan .old ve yarım .part dosyalarını siler.
    Eski süreç henüz kapanmamış olabilir; birkaç kez aralıkla tekrar denenir."""
    if sys.platform == "darwin":
        _temizlik_thread_mac()
        return
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
        with _urlopen(_istek(API_URL, "application/vnd.github+json")) as yanit:
            return json.loads(yanit.read().decode("utf-8"))
    except urllib.error.HTTPError as hata:
        if hata.code == 404:
            _log().info("Yayınlanmış release yok (404)")
            return None
        raise


def _indir(url: str, hedef: Path, ilerleme=None, toplam=None):
    """URL'yi hedef dosyaya parça parça indirir.

    ilerleme verilirse (indirilen, toplam, son) ile çağrılır (indirme thread'inde).
    toplam bilinmiyorsa Content-Length denenir; o da yoksa None kalır."""
    with _urlopen(_istek(url, "application/octet-stream")) as yanit, \
            open(hedef, "wb") as f:
        if not toplam:
            try:
                toplam = int(yanit.headers.get("Content-Length")) or None
            except Exception:
                toplam = None
        indirilen = 0
        if ilerleme:
            ilerleme(indirilen, toplam, False)
        for parca in iter(lambda: yanit.read(PARCA_BOYUTU), b""):
            f.write(parca)
            indirilen += len(parca)
            if ilerleme:
                ilerleme(indirilen, toplam, False)
        if ilerleme:
            ilerleme(indirilen, toplam, True)


def _kontrol_thread():
    """Sürüm kontrolü, indirme ve doğrulama. Hazırsa pop-up'ı ana thread'e iletir."""
    try:
        if not _etkin_mi():
            _log().info("Güncelleme kontrolü atlandı (paketlenmiş uygulama değil ya da "
                        "desteklenmeyen konum: frozen=%s, platform=%s)",
                        getattr(sys, "frozen", False), sys.platform)
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

        mac = sys.platform == "darwin"
        asset_adi = MAC_ZIP_ADI if mac else EXE_ADI
        asset = next((a for a in release.get("assets") or []
                      if isinstance(a, dict) and a.get("name") == asset_adi), None)
        if asset is None or not asset.get("browser_download_url"):
            _log().warning("Release %s içinde %s bulunamadı", etiket, asset_adi)
            return

        exe = _hedef_yolu()  # macOS'ta .app paketi
        if not _yazilabilir_mi(exe.parent):
            _log().warning("Exe klasörü yazılabilir değil, güncelleme atlandı: %s", exe.parent)
            return

        yeni = exe.parent / f"{exe.stem}.{surum}.new"  # macOS'ta klasör
        body = release.get("body") if isinstance(release.get("body"), str) else ""
        bilgi = {"surum": surum, "dosya": yeni, "notlar": body}
        from pages import popup_menu
        if (_mac_hazir_mi(yeni) if mac else yeni.exists() and _dogrula(yeni, asset)):
            _log().info("Geçerli %s zaten mevcut, yeniden indirilmedi", yeni.name)
        else:
            if mac:
                _sil_agac(yeni)
            else:
                _sil(yeni)
            parca = exe.parent / (f"{exe.stem}.{surum}.zip.part" if mac
                                  else f"{exe.stem}.{surum}.part")
            _log().info("İndiriliyor: %s -> %s", asset["browser_download_url"], parca.name)
            # "İndiriliyor" penceresi ana thread'de açılmaya çalışılır; indirme beklemez.
            # pencere sözlüğü yalnızca ana thread'de okunur/değiştirilir.
            pencere = {"durum": "bekliyor"}
            popup_menu.ana_threadde_calistir(
                state.app, lambda: _indirme_penceresi_ac(bilgi, pencere))
            boyut = asset.get("size")
            son_gonderim = [0.0]

            def ilerleme(indirilen, toplam, son):
                simdi = time.monotonic()
                if not son and simdi - son_gonderim[0] < ILERLEME_ARALIGI:
                    return
                son_gonderim[0] = simdi
                popup_menu.ana_threadde_calistir(
                    state.app, lambda: _ilerleme_goster(pencere, indirilen, toplam))

            basarili = False
            try:
                try:
                    _indir(asset["browser_download_url"], parca, ilerleme,
                           boyut if isinstance(boyut, int) and boyut > 0 else None)
                except Exception:
                    _sil(parca)
                    raise
                if not _dogrula(parca, asset, b"PK\x03\x04" if mac else b"MZ"):
                    _sil(parca)
                    _log().warning("İndirilen dosya doğrulanamadı, silindi")
                    return
                if mac:
                    if not _mac_zip_ac(parca, yeni):
                        _sil(parca)
                        return
                else:
                    try:
                        os.replace(parca, yeni)
                    except Exception:
                        _sil(parca)
                        raise
                _log().info("İndirildi ve doğrulandı: %s", yeni.name)
                basarili = True
            finally:
                popup_menu.ana_threadde_calistir(
                    state.app, lambda: _indirme_bitti(bilgi, pencere, basarili))
            return  # Sonraki adım (hazır pencere / erteleme) _indirme_bitti'de

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


HAZIR_MESAJI = "Güncelleme için uygulama kısa süreliğine kapanıp yeniden açılacak."
INDIRME_MESAJI = ("Yeni sürüm indiriliyor. Bu pencereyi kapatırsanız indirme arka planda "
                  "sürer; güncelleme hazır olduğunda uygulama kapanırken kurulur.")
HATA_MESAJI = ("Güncelleme şu an indirilemedi. Uygulamayı kullanmaya devam edebilirsiniz; "
               "bir sonraki açılışta tekrar denenecek.")


def _popup_yok_et(popup):
    """Pencereyi grab'ı bırakarak güvenle kapatır; hata yalnızca yoksayılır."""
    if popup is None:
        return
    try:
        popup.grab_release()
    except Exception:
        pass
    try:
        popup.destroy()
    except Exception:
        pass


def _guncelleme_penceresi(bilgi, indirme=None):
    """Doğrulanmış güncelleme dosyası hazırken gösterilen modal pencere.

    indirme verilirse (ana thread'e ait durum sözlüğü) pencere "indiriliyor"
    aşamasında açılır; widget'lar bu sözlüğe yazılır ve _indirme_bitti ile
    aynı pencere "hazır" ya da "hata" aşamasına geçirilir.
    Kurulum sırasında hata olursa pencere (grab bırakılarak) yok edilir ve
    hata yukarı iletilir."""
    popup = ctk.CTkToplevel(state.app)
    if indirme is not None:
        indirme["popup"] = popup  # Hata olursa çağıran da kapatabilsin
    try:
        _guncelleme_penceresi_kur(popup, bilgi, indirme)
    except Exception:
        _popup_yok_et(popup)
        raise


def _guncelleme_penceresi_kur(popup, bilgi, indirme):
    """_guncelleme_penceresi'nin içerik/buton kurulumu."""
    from pages.kontrol_popups import _popup_basligi

    baslik = "Güncelleme İndiriliyor" if indirme is not None else "Güncelleme Hazır"
    popup.title(baslik)
    popup.geometry("440x300")
    popup.resizable(False, False)
    popup.attributes("-topmost", True)
    popup.grab_set()  # Modal yap
    popup.configure(fg_color=theme.BG_ROOT)

    head = _popup_basligi(popup, baslik,
                          f"Yeni sürüm v{bilgi['surum']} (mevcut v{APP_VERSION})",
                          theme.ACCENT, icons.chevron_asagi if indirme is not None else icons.onay)
    ikon_etiketi, baslik_etiketi = _baslik_parcalari(head, baslik)

    mesaj_karti = ctk.CTkFrame(popup, fg_color=theme.BG_SURFACE, corner_radius=theme.CORNER_POPUP,
                               border_width=theme.BORDER_WIDTH, border_color=theme.BORDER)
    mesaj_karti.pack(fill="both", expand=True, padx=26, pady=(0, 18))
    mesaj = ctk.CTkLabel(mesaj_karti,
                         text=INDIRME_MESAJI if indirme is not None else HAZIR_MESAJI,
                         font=theme.FONT_MESSAGE(), text_color=theme.TEXT_SECONDARY,
                         justify="left", wraplength=350)
    mesaj.pack(padx=16, pady=(14, 6), anchor="w")
    ozet = _not_ozeti(bilgi.get("notlar"))
    notlar = None
    if ozet:
        notlar = ctk.CTkLabel(mesaj_karti, text=ozet, font=theme.FONT_SMALL(),
                              text_color=theme.TEXT_MUTED, justify="left", wraplength=350)
        notlar.pack(padx=16, pady=(0, 14), anchor="w")

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
        if indirme is not None and indirme.get("asama") == "indiriliyor":
            # İndirme arka planda sürer; sonucu _indirme_bitti ele alır
            indirme["durum"] = "sonra"
            _log().info("Güncelleme penceresi indirme sırasında kapatıldı, "
                        "indirme arka planda sürüyor (v%s)", bilgi["surum"])
            return
        _sonraya_birak(bilgi)

    if indirme is not None:
        ilerleme_satiri = ctk.CTkFrame(popup, fg_color="transparent")
        ilerleme_satiri.pack(fill="x", padx=26, pady=(0, 16))
        cubuk = ctk.CTkProgressBar(ilerleme_satiri, height=8, corner_radius=4,
                                   fg_color=theme.BG_ELEVATED, progress_color=theme.ACCENT,
                                   mode="determinate")
        cubuk.set(0)
        cubuk.pack(side="left", fill="x", expand=True)
        yuzde = ctk.CTkLabel(ilerleme_satiri, text="%0", font=theme.FONT_SMALL(),
                             text_color=theme.TEXT_SECONDARY, width=150, anchor="e")
        yuzde.pack(side="left", padx=(12, 0))

    btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
    btn_frame.pack(pady=(0, 20))
    simdi_btn = ctk.CTkButton(btn_frame, text="Şimdi Güncelle", width=theme.BTN_W_MD,
                              height=theme.BTN_H_MD,
                              corner_radius=theme.CORNER_BTN, font=theme.FONT_BODY_BOLD(),
                              fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                              text_color=theme.TEXT_ON_ACCENT, command=simdi)
    simdi_btn.pack(side="left", padx=5)
    sonra_btn = ctk.CTkButton(btn_frame, text="Sonra", width=theme.BTN_W_SM, height=theme.BTN_H_MD,
                              corner_radius=theme.CORNER_BTN, font=theme.FONT_SMALL(),
                              fg_color="transparent", border_width=theme.BORDER_WIDTH,
                              border_color=theme.BORDER,
                              hover_color=theme.BG_HOVER, text_color=theme.TEXT_SECONDARY,
                              command=sonra)
    sonra_btn.pack(side="left", padx=5)
    if indirme is not None:
        # İndirme bitene kadar pasif (kontrol_popups'taki pasif buton görünümü)
        simdi_btn.configure(state="disabled", fg_color=theme.BG_ELEVATED,
                            hover_color=theme.BG_ELEVATED, text_color=theme.TEXT_MUTED)

    # Pencere kapatma (X) = "Sonra"
    popup.protocol("WM_DELETE_WINDOW", sonra)

    # Yükseklik içerikten hesaplanır, pencere ana uygulamaya göre ortalanır
    popup.update_idletasks()
    height = max(300, popup.winfo_reqheight())
    x = state.app.winfo_x() + (state.app.winfo_width() // 2) - 220
    y = state.app.winfo_y() + (state.app.winfo_height() // 2) - (height // 2)
    popup.geometry(f"440x{height}+{x}+{y}")
    if indirme is None:
        _log().info("Güncelleme penceresi gösterildi (v%s)", bilgi["surum"])
        return
    indirme.update(asama="indiriliyor", popup=popup, baslik=baslik_etiketi, ikon=ikon_etiketi,
                   mesaj=mesaj, notlar=notlar, cubuk=cubuk, yuzde=yuzde, simdi_btn=simdi_btn,
                   sonra_btn=sonra_btn, ilerleme_satiri=ilerleme_satiri, btn_frame=btn_frame)
    _log().info("Güncelleme penceresi gösterildi (indiriliyor, v%s)", bilgi["surum"])


def _baslik_parcalari(head, baslik):
    """_popup_basligi çerçevesindeki (ikon, başlık) etiketlerini sıra varsaymadan bulur:
    başlık metni eşleşen etiket başlık, metinsiz ve görselli etiket ikondur."""
    ikon = etiket = None
    bekleyen = list(head.winfo_children())
    while bekleyen:
        w = bekleyen.pop()
        if isinstance(w, ctk.CTkLabel):
            if w.cget("text") == baslik:
                etiket = w
            elif not w.cget("text") and w.cget("image") is not None:
                ikon = w
        elif isinstance(w, ctk.CTkFrame):
            bekleyen.extend(w.winfo_children())
    return ikon, etiket


def _baslik_ikonu(indirme, glif, renk):
    """Başlık ikonunu ve ikon kutusu zeminini aşamaya göre değiştirir
    (ikon bulunamadıysa dokunmaz)."""
    if indirme.get("ikon") is not None:
        indirme["ikon"].configure(image=glif(theme.ICON_LG, renk), fg_color=theme.TINT(renk))


def _pencere_acik_mi(indirme) -> bool:
    """İndirme penceresi hâlâ açık ve kullanılabilir mi (yalnızca ana thread)."""
    try:
        return indirme.get("durum") == "acik" and bool(indirme["popup"].winfo_exists())
    except Exception:
        return False


def _mb(bayt) -> str:
    return f"{bayt / (1024 * 1024):.1f}".replace(".", ",")


def _indirme_penceresi_ac(bilgi, indirme):
    """Ana thread: 'indiriliyor' penceresini açar. Başka modal açıksa açmaz;
    indirme sessiz sürer ve bitince bugünkü hazır pencere akışı kullanılır."""
    try:
        if state.app.grab_current() is not None or state.popup_aktif:
            indirme["durum"] = "gosterilmedi"
            _log().info("Güncelleme penceresi şu an gösterilemedi, indirme sessiz sürüyor (v%s)",
                        bilgi["surum"])
            return
        _guncelleme_penceresi(bilgi, indirme)
        indirme["durum"] = "acik"
    except Exception:
        indirme["durum"] = "gosterilmedi"  # Pencere yok; indirme arka planda sürer
        _log().exception("Güncelleme penceresi gösterilemedi")
        _popup_yok_et(indirme.get("popup"))


def _ilerleme_goster(indirme, indirilen, toplam):
    """Ana thread: ilerleme çubuğunu ve etiketini günceller. Pencere kapandıysa yoksayar."""
    try:
        if not _pencere_acik_mi(indirme) or indirme.get("asama") != "indiriliyor":
            return
        cubuk, yuzde = indirme["cubuk"], indirme["yuzde"]
        if toplam:
            oran = min(1.0, indirilen / toplam)
            cubuk.set(oran)
            yuzde.configure(text=f"%{int(oran * 100)}  ({_mb(indirilen)} / {_mb(toplam)} MB)")
        else:
            if not indirme.get("belirsiz"):
                indirme["belirsiz"] = True
                cubuk.configure(mode="indeterminate")
                cubuk.start()
            yuzde.configure(text=f"{_mb(indirilen)} MB")
    except Exception:
        _log().exception("İlerleme gösterilemedi")


def _indirme_bitti(bilgi, indirme, basarili, deneme=0):
    """Ana thread: indirme sonucu. Açık pencereyi başarılıysa hazır, başarısızsa hata
    aşamasına geçirir; pencere kapatılmışsa kapanışta kurulumu kurar, hiç gösterilmediyse
    hazır pencereyi dener. Pencere açık değilken hata yalnızca log'a yazılır."""
    try:
        if indirme.get("durum") == "bekliyor":
            # Pencere henüz kuruluyor (CTkToplevel kurulumda update() çağırabilir); sonra tekrar
            if deneme >= BITTI_DENEME:
                _log().warning("İndirme sonucu işlenemedi: pencere kurulumu bitmedi (v%s)",
                               bilgi["surum"])
                return
            state.app.after(100, lambda: _indirme_bitti(bilgi, indirme, basarili, deneme + 1))
            return
        if _pencere_acik_mi(indirme):
            if basarili:
                _hazir_asamasina_gec(bilgi, indirme)
            else:
                try:
                    _hata_asamasina_gec(bilgi, indirme)
                except Exception:
                    indirme["durum"] = "kapandi"  # Hata ekranı kurulamadı: grab'ı bırakıp kapat
                    _popup_yok_et(indirme["popup"])
                    raise
            return
        if indirme.get("durum") == "sonra":
            if basarili:
                _log().info("İndirme arka planda tamamlandı, kapanışta kurulacak (v%s)",
                            bilgi["surum"])
                _sonraya_birak(bilgi)
            else:
                _log().info("İndirme arka planda başarısız oldu (v%s)", bilgi["surum"])
            return
        if basarili:
            _popup_goster_dene(bilgi)
    except Exception:
        _log().exception("İndirme sonucu işlenemedi")


def _boyutu_yenile(popup, en_az=300):
    """Aşama değişince yüksekliği içeriğe göre yeniler, konumu korur."""
    popup.update_idletasks()
    height = max(en_az, popup.winfo_reqheight())
    popup.geometry(f"440x{height}+{popup.winfo_x()}+{popup.winfo_y()}")


def _hazir_asamasina_gec(bilgi, indirme):
    """Aynı pencereyi 'Güncelleme Hazır' aşamasına geçirir; butonlar bugünkü gibi çalışır."""
    indirme["asama"] = "hazir"
    popup, cubuk = indirme["popup"], indirme["cubuk"]
    popup.title("Güncelleme Hazır")
    indirme["baslik"].configure(text="Güncelleme Hazır")
    _baslik_ikonu(indirme, icons.onay, theme.ACCENT)
    indirme["mesaj"].configure(text=HAZIR_MESAJI)
    if indirme.get("belirsiz"):
        cubuk.stop()
        cubuk.configure(mode="determinate")
    cubuk.set(1)
    # Sabit genişlik (150) yalnızca indirme metni için; "%100" kadar daralınca çubuk satırı doldurur
    indirme["yuzde"].configure(text="%100", width=0)
    indirme["simdi_btn"].configure(state="normal", fg_color=theme.ACCENT,
                                   hover_color=theme.ACCENT_HOVER,
                                   text_color=theme.TEXT_ON_ACCENT)
    _boyutu_yenile(popup)
    _log().info("Güncelleme penceresi gösterildi (v%s)", bilgi["surum"])


def _hata_asamasina_gec(bilgi, indirme):
    """Aynı pencereyi 'Güncelleme İndirilemedi' aşamasına geçirir. Teknik detay
    gösterilmez (log'da); 'Tamam'/X pencereyi kapatır, kurulum ertelenmez."""
    indirme["asama"] = "hata"
    popup, cubuk = indirme["popup"], indirme["cubuk"]

    def kapat():
        indirme["durum"] = "kapandi"
        _popup_yok_et(popup)

    popup.title("Güncelleme İndirilemedi")
    indirme["baslik"].configure(text="Güncelleme İndirilemedi")
    _baslik_ikonu(indirme, icons.uyari, theme.WARNING)
    indirme["mesaj"].configure(text=HATA_MESAJI)
    if indirme.get("notlar") is not None:
        indirme["notlar"].pack_forget()
    indirme["mesaj"].pack_configure(pady=14)
    if indirme.get("belirsiz"):
        cubuk.stop()
    indirme["ilerleme_satiri"].pack_forget()
    indirme["simdi_btn"].pack_forget()
    indirme["sonra_btn"].pack_forget()
    ctk.CTkButton(indirme["btn_frame"], text="Tamam", width=theme.BTN_W_SM,
                  height=theme.BTN_H_MD, corner_radius=theme.CORNER_BTN, font=theme.FONT_SMALL(),
                  fg_color="transparent", border_width=theme.BORDER_WIDTH,
                  border_color=theme.BORDER, hover_color=theme.BG_HOVER,
                  text_color=theme.TEXT_SECONDARY, command=kapat).pack(side="left", padx=5)
    popup.protocol("WM_DELETE_WINDOW", kapat)
    _boyutu_yenile(popup, en_az=0)  # Hata aşamasında yükseklik içerik kadar
    _log().info("İndirme başarısız, güncelleme penceresi hata aşamasında (v%s)",
                bilgi["surum"])

# ----------------- Kurulum -----------------

def _kur_mac(staging: Path) -> bool:
    """macOS: çalışan .app'i .app.old'a taşır, staging'deki yeni .app'i yerine koyar.
    Hata olursa geri alır; başarılıysa staging klasörünü siler."""
    try:
        app = _hedef_yolu()
        eski = _eski_exe_yolu(app)
        if not _mac_hazir_mi(staging):
            _log().warning("Kurulacak paket yok: %s", staging)
            return False
        if eski.exists() or eski.is_symlink():
            _sil_agac(eski)
        os.replace(app, eski)
        try:
            os.replace(staging / MAC_APP_ADI, app)
        except Exception:
            _log().exception("Yeni .app yerine konamadı, geri alınıyor")
            try:
                os.replace(eski, app)
            except Exception:
                _log().exception("Geri alma başarısız")
            return False
        _sil_agac(staging)
        _log().info("Güncelleme kuruldu: %s", staging.name)
        return True
    except Exception:
        _log().exception("Kurulum başarısız")
        return False


def _kur(yeni: Path) -> bool:
    """Çalışan exe'yi .old'a taşır, yeni dosyayı yerine koyar. Hata olursa geri alır."""
    if sys.platform == "darwin":
        return _kur_mac(yeni)
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
    """Yeni exe'yi konsol penceresi açmadan, bağımsız süreç olarak başlatır.
    macOS'ta yeni .app'i 'open -n' ile başlatır."""
    if sys.platform == "darwin":
        try:
            subprocess.Popen(
                ["/usr/bin/open", "-n", str(_hedef_yolu())],
                stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                start_new_session=True, close_fds=True, env=_temiz_ortam(),
            )
            _log().info("Yeni sürüm başlatıldı")
            return True
        except Exception:
            _log().exception("Yeniden başlatma başarısız")
            return False
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
        kurulum_denendi = False

        def kapanirken():
            nonlocal kurulum_denendi
            try:
                if _bekleyen is not None and not kurulum_denendi:
                    kurulum_denendi = True  # Birden fazla çağrılırsa bir kez kur
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
        if sys.platform == "darwin":
            # Cmd+Q / uygulama menüsü "Quit" WM_DELETE_WINDOW'u değil ::tk::mac::Quit'i çağırır
            state.app.createcommand("::tk::mac::Quit", kapanirken)
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
