import threading
from tkinter import filedialog

import customtkinter as ctk

import icons
import state
import theme
from logic.xml_alis import klasor_sec_ve_isle_alis
from logic.xml_hepsiburada import klasor_sec_hepsiburada_alis, klasor_sec_hepsiburada_satis
from logic.xml_mustahsil import klasor_sec_ve_isle_mustahsil
from logic.xml_mustahsil_makbuz import klasor_sec_ve_isle_mustahsil_xml
from logic.xml_mustahsil_muhtasar import klasor_sec_ve_isle_mustahsil_muhtasar
from logic.xml_trendyol import klasor_sec_trendyol_alis, klasor_sec_trendyol_satis
from logic.xml_uyumsoft import klasor_sec_ve_isle_uyumsoft
from pages import kontrol_popups

# ----------------- Arka Plan İşlem Altyapısı -----------------
# XML aktarma işlemleri dosya okuma/yazma yaptığı için ana (Tkinter) thread'inde
# çalıştırıldığında arayüz donuyordu. Aşağıdaki yardımcılar işi arka plan
# thread'ine taşır; arayüz güncellemeleri ise after(0, ...) ile ana thread'e
# geri taşınır (Tk widget'larına yalnızca ana thread'den dokunulabilir).

# Aynı anda yalnızca tek bir XML aktarma işi çalışsın
_calisan_is_var = False


def ana_threadde_calistir(widget, fonk):
    """Verilen fonksiyonu ana thread'in olay döngüsünde çalıştırır."""
    try:
        widget.after(0, fonk)
    except Exception:
        # Pencere kapatılmış olabilir
        pass


class _GuvenliWidget:
    """Arka plan thread'inden gelen configure çağrılarını ana thread'e taşır.

    Logic katmanındaki fonksiyonlar `label_sonuc` / `buton_ac` üzerinde yalnızca
    `configure()` çağırıyor; bu sarmalayıcı sayesinde logic dosyalarına
    dokunmadan bu çağrılar güvenli hale gelir.
    """

    def __init__(self, widget):
        self._widget = widget
        self.guncellendi = False

    def configure(self, *args, **kwargs):
        self.guncellendi = True
        widget = self._widget
        ana_threadde_calistir(widget, lambda: widget.configure(*args, **kwargs))

    def __getattr__(self, ad):
        return getattr(self._widget, ad)


_gercek_askdirectory = filedialog.askdirectory


def _ana_threadde_klasor_sec(*args, **kwargs):
    """askdirectory diyaloğunu her zaman ana thread'de açar."""
    app = getattr(state, "app", None)
    if app is None or threading.current_thread() is threading.main_thread():
        return _gercek_askdirectory(*args, **kwargs)

    sonuc = {}
    bitti = threading.Event()

    def ac():
        try:
            sonuc["deger"] = _gercek_askdirectory(*args, **kwargs)
        except BaseException as hata:
            sonuc["hata"] = hata
        finally:
            bitti.set()

    app.after(0, ac)
    bitti.wait()
    if "hata" in sonuc:
        raise sonuc["hata"]
    return sonuc.get("deger", "")


# Logic katmanı klasör diyaloğunu doğrudan `filedialog.askdirectory` ile açıyor.
# İş arka plan thread'ine taşındığı için diyaloğun ana thread'de açılmasını
# garanti etmek üzere tek seferlik sarmalanır (ana thread'den çağrılınca
# davranış birebir aynı kalır).
filedialog.askdirectory = _ana_threadde_klasor_sec


def xml_isini_baslat(buton, islem_func, label_sonuc, buton_ac):
    """XML aktarma işlemini arka planda çalıştırır; arayüz donmaz."""
    global _calisan_is_var
    if _calisan_is_var:
        return
    _calisan_is_var = True

    eski_metin = buton.cget("text")
    buton.configure(state="disabled", text="İşleniyor...")
    label_sonuc.configure(text="İşleniyor...", text_color=theme.TEXT_SECONDARY)

    guvenli_label = _GuvenliWidget(label_sonuc)
    guvenli_buton_ac = _GuvenliWidget(buton_ac)

    def bitir():
        global _calisan_is_var
        _calisan_is_var = False
        try:
            if buton.winfo_exists():
                buton.configure(state="normal", text=eski_metin)
        except Exception:
            pass
        # Klasör seçimi iptal edildiyse "İşleniyor..." yazısı ekranda kalmasın
        if not guvenli_label.guncellendi:
            try:
                if label_sonuc.winfo_exists():
                    label_sonuc.configure(text="")
            except Exception:
                pass

    def calis():
        try:
            islem_func(guvenli_label, guvenli_buton_ac)
        except Exception as hata:
            guvenli_label.configure(text=f"Hata oluştu: {hata}", text_color=theme.ERROR)
        finally:
            ana_threadde_calistir(buton, bitir)

    threading.Thread(target=calis, daemon=True).start()


# ----------------- Popup Menü Görünümü -----------------
# Açılır menülerin görünümü tek yerden kurulur: dış pencere saç teli
# çerçeve rengini, içindeki çerçeve de menü yüzeyini taşır.

def menu_olcusu(satir_sayisi):
    """Satır sayısına göre açılır menünün başlangıç (genişlik, yükseklik) değeri.

    Yalnızca pencere ilk açılırken kullanılan kaba tahmindir; satır butonları
    yerleştirildikten sonra kesin yükseklik `menu_boyutunu_ayarla` ile ölçülür.
    """
    ic_yukseklik = 8 + satir_sayisi * theme.BTN_H_MENU + (satir_sayisi - 1) * 4 + 8
    return theme.MENU_W, ic_yukseklik + 2


def menu_boyutunu_ayarla(pencere, x, y):
    """Satırlar yerleştikten sonra menünün gerçek yüksekliğini ölçüp uygular.

    Satır butonlarının gerçek yüksekliği yazı tipine göre `theme.BTN_H_MENU`
    değerinden büyük olabildiği için formülle hesaplanan yükseklik son satırı
    kırpıyordu. Kontrol pencerelerindeki ile aynı yöntem kullanılır: önce
    yerleştir, sonra ölç, en son pencere ölçüsünü ver.
    """
    pencere.update_idletasks()
    yukseklik = pencere.winfo_reqheight()
    pencere.geometry(f"{theme.MENU_W}x{yukseklik}+{x}+{y}")


def menu_govdesi(pencere):
    """Açılır menünün iç yüzeyini (çerçeveli gövde) oluşturur."""
    pencere.configure(fg_color=theme.BORDER)
    govde = ctk.CTkFrame(pencere, fg_color=theme.BG_SURFACE, corner_radius=0)
    govde.pack(fill="both", expand=True, padx=theme.BORDER_WIDTH, pady=theme.BORDER_WIDTH)
    return govde


def menu_butonu(govde, metin, command, ilk=False, son=False):
    """Açılır menü satırı - sol hizalı, ikonlu, hover'da yüzeyi aydınlanır."""
    buton = ctk.CTkButton(govde, text=metin,
                          image=icons.ok_sag(theme.ICON_XS, theme.TEXT_MUTED),
                          compound="right",
                          width=theme.BTN_W_MENU, height=theme.BTN_H_MENU,
                          corner_radius=theme.CORNER_BTN, font=theme.FONT_BODY(),
                          fg_color="transparent", hover_color=theme.BG_HOVER,
                          text_color=theme.TEXT_SECONDARY,
                          anchor="w", border_spacing=10,
                          command=command)
    buton.pack(pady=(8 if ilk else 0, 8 if son else 4), padx=8)
    return buton


# ----------------- Popup Menü Fonksiyonları -----------------

def popup_menu_kapat():
    state.popup_aktif = False
    if state.aktif_menu is not None:
        try:
            state.aktif_menu.destroy()
        except Exception:
            # Pencere zaten kapatılmış olabilir
            pass
        state.aktif_menu = None

def on_app_click(event):
    """Ana pencereye tıklanınca popup'ı kapat"""
    if state.popup_aktif and state.aktif_menu is not None:
        # Popup içine tıklanmadıysa kapat
        widget = event.widget
        # Popup'ın child'ı değilse kapat
        if str(widget).startswith(".!ctktoplevel"):
            return  # Popup içi, kapatma
        popup_menu_kapat()

def show_popup_menu(button, label_sonuc, buton_ac, alis_func, satis_func):
    """Genel popup menü fonksiyonu - Alış/Satış seçenekleri gösterir"""
    popup_menu_kapat()
    
    x = button.winfo_rootx()
    y = button.winfo_rooty() + button.winfo_height() + 5
    
    w, h = menu_olcusu(2)
    state.aktif_menu = ctk.CTkToplevel(state.app)
    state.aktif_menu.geometry(f"{w}x{h}+{x}+{y}")
    state.aktif_menu.overrideredirect(True)
    state.aktif_menu.attributes("-topmost", True)
    govde = menu_govdesi(state.aktif_menu)

    menu_butonu(govde, "Alış Faturası",
                lambda: [popup_menu_kapat(), xml_isini_baslat(button, alis_func, label_sonuc, buton_ac)],
                ilk=True)
    menu_butonu(govde, "Satış Faturası",
                lambda: [popup_menu_kapat(), xml_isini_baslat(button, satis_func, label_sonuc, buton_ac)],
                son=True)
    menu_boyutunu_ayarla(state.aktif_menu, x, y)

    # Popup aktif olarak işaretle (biraz gecikmeyle)
    def activate_popup():
        state.popup_aktif = True
    state.app.after(150, activate_popup)
    
    # Escape tuşu ile kapatma
    state.aktif_menu.bind("<Escape>", lambda e: popup_menu_kapat())

def show_hepsiburada_menu(button, label_sonuc, buton_ac):
    show_popup_menu(button, label_sonuc, buton_ac, klasor_sec_hepsiburada_alis, klasor_sec_hepsiburada_satis)

def show_trendyol_menu(button, label_sonuc, buton_ac):
    show_popup_menu(button, label_sonuc, buton_ac, klasor_sec_trendyol_alis, klasor_sec_trendyol_satis)

def show_hizlibilisim_xml_menu(button, label_sonuc, buton_ac):
    """Hızlıbilişim XML için alış/satış menüsü - Hepsiburada ile aynı fonksiyonları kullanır"""
    show_popup_menu(button, label_sonuc, buton_ac, klasor_sec_hepsiburada_alis, klasor_sec_hepsiburada_satis)

def show_mustahsil_menu(button, label_sonuc, buton_ac):
    """Müstahsil butonu için 4 seçenekli popup menü: Borsa, Alış XML, Müstahsil XML, Muhtasar (Zirve)"""
    popup_menu_kapat()
    
    x = button.winfo_rootx()
    y = button.winfo_rooty() + button.winfo_height() + 5
    
    w, h = menu_olcusu(4)
    state.aktif_menu = ctk.CTkToplevel(state.app)
    state.aktif_menu.geometry(f"{w}x{h}+{x}+{y}")
    state.aktif_menu.overrideredirect(True)
    state.aktif_menu.attributes("-topmost", True)
    govde = menu_govdesi(state.aktif_menu)

    menu_butonu(govde, "Müstahsil Borsa",
                lambda: [popup_menu_kapat(), xml_isini_baslat(button, klasor_sec_ve_isle_mustahsil, label_sonuc, buton_ac)],
                ilk=True)
    menu_butonu(govde, "Alış XML",
                lambda: [popup_menu_kapat(), xml_isini_baslat(button, klasor_sec_ve_isle_alis, label_sonuc, buton_ac)])
    menu_butonu(govde, "Müstahsil XML",
                lambda: [popup_menu_kapat(), xml_isini_baslat(button, klasor_sec_ve_isle_mustahsil_xml, label_sonuc, buton_ac)])
    menu_butonu(govde, "Muhtasar (Zirve)",
                lambda: [popup_menu_kapat(), xml_isini_baslat(button, klasor_sec_ve_isle_mustahsil_muhtasar, label_sonuc, buton_ac)],
                son=True)
    menu_boyutunu_ayarla(state.aktif_menu, x, y)

    # Popup aktif olarak işaretle (biraz gecikmeyle)
    def activate_popup():
        state.popup_aktif = True
    state.app.after(150, activate_popup)
    
    # Escape tuşu ile kapatma
    state.aktif_menu.bind("<Escape>", lambda e: popup_menu_kapat())

def show_uyumsoft_xml_menu(button, label_sonuc, buton_ac):
    """Uyumsoft XML için ACE/Turkay dropdown menüsü"""
    popup_menu_kapat()
    
    x = button.winfo_rootx()
    y = button.winfo_rooty() + button.winfo_height() + 5
    
    w, h = menu_olcusu(2)
    state.aktif_menu = ctk.CTkToplevel(state.app)
    state.aktif_menu.geometry(f"{w}x{h}+{x}+{y}")
    state.aktif_menu.overrideredirect(True)
    state.aktif_menu.attributes("-topmost", True)
    govde = menu_govdesi(state.aktif_menu)

    menu_butonu(govde, "ACE",
                lambda: [popup_menu_kapat(), kontrol_popups.show_info_popup("Bu özellik henüz aktif değil.")],
                ilk=True)
    menu_butonu(govde, "Turkay",
                lambda: [popup_menu_kapat(), xml_isini_baslat(button, klasor_sec_ve_isle_uyumsoft, label_sonuc, buton_ac)],
                son=True)
    menu_boyutunu_ayarla(state.aktif_menu, x, y)

    # Popup aktif olarak işaretle (biraz gecikmeyle)
    def activate_popup():
        state.popup_aktif = True
    state.app.after(150, activate_popup)
    
    # Escape tuşu ile kapatma
    state.aktif_menu.bind("<Escape>", lambda e: popup_menu_kapat())
