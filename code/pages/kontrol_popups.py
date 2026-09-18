import os
import threading

import pandas as pd
import customtkinter as ctk
from tkinter import filedialog

import icons
import state
import theme
from logic.common import dosyayi_ac
from logic.hb_trendyol import hb_ve_trn_karsilastir
from logic.hizlibilisim import hizlibilisim_karsilastir
from logic.izibiz import izibiz_karsilastir
from logic.logo_hizli_kombine import logo_hizli_kombine_karsilastir
from logic.uyumsoft import uyumsoft_verim_karsilastir, uyumsoft_zirve_karsilastir
from logic.vega import vega_karsilastir
from pages import popup_menu

# ----------------- Popup Görünüm Yardımcıları -----------------
# Popup'lar da karşılama ekranıyla aynı görsel dili kullanır: ikon kutulu
# başlık, saç teli çerçeveli kartlar ve yuvarlatılmış butonlar.

def _popup_basligi(popup, baslik, aciklama, renk, glif):
    """İkon kutusu + başlık + açıklamadan oluşan popup başlığı."""
    head = ctk.CTkFrame(popup, fg_color="transparent")
    head.pack(fill="x", padx=26, pady=(22, 18))
    ctk.CTkLabel(head, text="", image=glif(theme.ICON_LG, renk),
                 width=theme.HEAD_TILE, height=theme.HEAD_TILE,
                 corner_radius=theme.CORNER_TILE, fg_color=theme.TINT(renk)).pack(side="left")
    metin = ctk.CTkFrame(head, fg_color="transparent")
    metin.pack(side="left", padx=14)
    ctk.CTkLabel(metin, text=baslik, font=theme.FONT_POPUP_TITLE(),
                 text_color=theme.TEXT_PRIMARY).pack(anchor="w")
    ctk.CTkLabel(metin, text=aciklama, font=theme.FONT_SMALL(),
                 text_color=theme.TEXT_MUTED, justify="left").pack(anchor="w", pady=(3, 0))
    return head


def _dosya_karti(parent, sutun, baslik, renk):
    """Dosya seçim kartı: ikon, etiket ve durum satırı. (kart, durum) döndürür."""
    kart = ctk.CTkFrame(parent, fg_color=theme.BG_SURFACE, corner_radius=theme.CORNER_POPUP,
                        border_width=theme.BORDER_WIDTH, border_color=theme.BORDER)
    kart.grid(row=0, column=sutun, padx=(0, 10) if sutun == 0 else (10, 0), sticky="nsew")
    parent.grid_columnconfigure(sutun, weight=1)
    ctk.CTkLabel(kart, text="", image=icons.tablo(theme.ICON_MD, renk),
                 width=theme.TILE_SIZE, height=theme.TILE_SIZE,
                 corner_radius=theme.CORNER_TILE, fg_color=theme.TINT(renk)).pack(pady=(16, 8))
    ctk.CTkLabel(kart, text=baslik, font=theme.FONT_BODY_BOLD(),
                 text_color=theme.TEXT_PRIMARY).pack()
    durum = ctk.CTkLabel(kart, text="Dosya seçilmedi", font=theme.FONT_TINY(),
                         text_color=theme.TEXT_MUTED)
    durum.pack(pady=(4, 12))
    return kart, durum


def _sec_butonu(parent, metin, command):
    """Dosya kartı içindeki ikincil 'Seç' butonu."""
    return ctk.CTkButton(parent, text=metin, width=theme.BTN_W_SM_WIDE, height=theme.BTN_H_SM,
                         corner_radius=theme.CORNER_BTN, font=theme.FONT_SMALL(),
                         fg_color="transparent", hover_color=theme.BG_HOVER,
                         text_color=theme.TEXT_SECONDARY,
                         border_width=theme.BORDER_WIDTH, border_color=theme.BORDER_HI,
                         command=command)


def _iptal_butonu(popup):
    """Popup'ın alt kısmındaki 'İptal' butonu."""
    return ctk.CTkButton(popup, text="İptal", width=theme.BTN_W_SM, height=theme.BTN_H_BACK,
                         corner_radius=theme.CORNER_BTN, font=theme.FONT_SMALL(),
                         fg_color="transparent", border_width=theme.BORDER_WIDTH,
                         border_color=theme.BORDER, hover_color=theme.BG_HOVER,
                         text_color=theme.TEXT_SECONDARY, command=popup.destroy)


def _eylem_butonu(popup, command):
    """Pasif başlayan 'Kontrol Et' butonu."""
    return ctk.CTkButton(popup, text="Kontrol Et", width=theme.BTN_W_ACTION,
                         height=theme.BTN_H_ACTION, corner_radius=theme.CORNER_BTN,
                         font=theme.FONT_BODY_BOLD(), command=command, state="disabled",
                         fg_color=theme.BG_ELEVATED, hover_color=theme.BG_ELEVATED,
                         text_color=theme.TEXT_MUTED)


def _eylem_butonu_hazir(buton, renk):
    """Her iki dosya seçilince 'Kontrol Et' butonunu etkinleştirir."""
    buton.configure(state="normal", fg_color=renk,
                    hover_color=theme.karistir(renk, 0.78, "#ffffff"),
                    text_color=theme.TEXT_ON_ACCENT,
                    border_width=0)


# Kontrol popup fonksiyonu
def show_kontrol_popup(portal_adi, kontrol_func=None):
    """Excel karşılaştırma popup'ı göster"""
    popup = ctk.CTkToplevel(state.app)
    popup.title(f"{portal_adi} Kontrol")
    popup.geometry("520x440")
    popup.resizable(False, False)
    popup.attributes("-topmost", True)
    popup.grab_set()  # Modal yap
    popup.configure(fg_color=theme.BG_ROOT)
    
    # Pencereyi ortala
    popup.update_idletasks()
    x = state.app.winfo_x() + (state.app.winfo_width() // 2) - 260
    y = state.app.winfo_y() + (state.app.winfo_height() // 2) - 220
    popup.geometry(f"520x440+{x}+{y}")
    
    # Seçilen dosyaları takip et
    secilen_dosyalar = {"zirve": None, "portal": []}
    
    # Başlık
    _popup_basligi(popup, f"{portal_adi} Fatura Kontrol",
                   "Karşılaştırmak için Excel dosyası/dosyaları seçin",
                   theme.SUCCESS, icons.karsilastir)

    # Dosya seçim alanları
    files_frame = ctk.CTkFrame(popup, fg_color="transparent")
    files_frame.pack(fill="x", padx=26)

    # Sol: Zirve Excel
    zirve_frame, zirve_status = _dosya_karti(files_frame, 0, "Zirve Excel", theme.ACCENT)

    def select_zirve():
        # Popup'ı gizle
        popup.withdraw()
        popup.update_idletasks()
        
        dosya = filedialog.askopenfilename(
            title="Zirve Excel Dosyası Seçin",
            filetypes=[("Excel Dosyaları", "*.xlsx *.xls")]
        )
        
        # Popup'ı tekrar göster
        popup.deiconify()
        popup.lift()
        popup.focus_force()
        
        if dosya:
            secilen_dosyalar["zirve"] = dosya
            dosya_adi = os.path.basename(dosya)
            if len(dosya_adi) > 20:
                dosya_adi = dosya_adi[:17] + "..."
            zirve_status.configure(text=f"{dosya_adi}", text_color=theme.SUCCESS)
            check_ready()
    
    _sec_butonu(zirve_frame, "Dosya Seç", select_zirve).pack(pady=(0, 16))

    # Sağ: Portal Excel (Çoklu Seçim)
    portal_frame, portal_status = _dosya_karti(files_frame, 1, f"{portal_adi} Excel", theme.SUCCESS)

    def select_portal():
        # Popup'ı gizle
        popup.withdraw()
        popup.update_idletasks()
        
        dosyalar = filedialog.askopenfilenames(
            title=f"{portal_adi} Excel Dosyası(ları) Seçin (Çoklu seçim için Ctrl tuşu)",
            filetypes=[("Excel Dosyaları", "*.xlsx *.xls")]
        )
        
        # Popup'ı tekrar göster
        popup.deiconify()
        popup.lift()
        popup.focus_force()
        
        if dosyalar:
            secilen_dosyalar["portal"] = list(dosyalar)
            dosya_sayisi = len(dosyalar)
            
            if dosya_sayisi == 1:
                dosya_adi = os.path.basename(dosyalar[0])
                if len(dosya_adi) > 18:
                    dosya_adi = dosya_adi[:15] + "..."
                portal_status.configure(text=f"{dosya_adi}", text_color=theme.SUCCESS)
            else:
                portal_status.configure(text=f"{dosya_sayisi} dosya seçildi", text_color=theme.SUCCESS)
            
            check_ready()
    
    _sec_butonu(portal_frame, "Dosyalar Seç", select_portal).pack(pady=(0, 16))

    # Kontrol Et butonu (başlangıçta devre dışı)
    def run_kontrol():
        if secilen_dosyalar["zirve"] and secilen_dosyalar["portal"]:
            if kontrol_func:
                _kontrol_baslat(popup, kontrol_btn, kontrol_func,
                                secilen_dosyalar["zirve"], secilen_dosyalar["portal"])
            else:
                # Henüz implementasyon yok
                popup.destroy()
                show_info_popup("Bu özellik henüz aktif değil.")
    
    kontrol_btn = _eylem_butonu(popup, run_kontrol)
    kontrol_btn.pack(pady=(24, 12))

    def check_ready():
        # Zirve seçilmiş ve Portal'de en az 1 dosya seçilmişse
        if secilen_dosyalar["zirve"] and len(secilen_dosyalar["portal"]) > 0:
            _eylem_butonu_hazir(kontrol_btn, theme.SUCCESS)

    # İptal butonu
    _iptal_butonu(popup).pack()

def show_info_popup(mesaj):
    """Bilgi popup'ı göster - Dinamik boyutlandırma ve kopyalanabilir metin"""
    popup = ctk.CTkToplevel(state.app)
    popup.title("Bilgi")
    popup.attributes("-topmost", True)
    popup.grab_set()
    popup.configure(fg_color=theme.BG_ROOT)
    
    # Mesaj uzunluğuna göre dinamik boyut hesapla
    satir_sayisi = mesaj.count('\n') + 1
    karakter_sayisi = max(len(satir) for satir in mesaj.split('\n'))
    
    # Minimum ve maksimum boyutlar
    min_width, max_width = 350, 800
    min_height, max_height = 200, 600
    
    # Genişlik: karakter sayısına göre (ortalama 8 piksel/karakter)
    width = min(max(min_width, karakter_sayisi * 8 + 100), max_width)
    
    popup.resizable(True, True)  # Kullanıcı manuel olarak da büyütebilir
    
    # Uyarı başlığı (ikon kutusu + başlık)
    ust = ctk.CTkFrame(popup, fg_color="transparent")
    ust.pack(fill="x", padx=20, pady=(18, 12))
    ctk.CTkLabel(ust, text="", image=icons.uyari(theme.ICON_MD, theme.WARNING),
                 width=theme.TILE_SIZE, height=theme.TILE_SIZE,
                 corner_radius=theme.CORNER_TILE,
                 fg_color=theme.TINT(theme.WARNING)).pack(side="left")
    ctk.CTkLabel(ust, text="Bilgi", font=theme.FONT_POPUP_TITLE(),
                 text_color=theme.TEXT_PRIMARY).pack(side="left", padx=14)

    # Kopyalanabilir metin kutusu
    text_frame = ctk.CTkFrame(popup, fg_color="transparent")
    text_frame.pack(fill="both", expand=True, padx=20, pady=(0, 15))

    textbox = ctk.CTkTextbox(text_frame, font=theme.FONT_SMALL(),
                             wrap="word", activate_scrollbars=True,
                             fg_color=theme.BG_SURFACE, text_color=theme.TEXT_PRIMARY,
                             corner_radius=theme.CORNER_POPUP,
                             border_width=theme.BORDER_WIDTH, border_color=theme.BORDER)
    textbox.pack(fill="both", expand=True)
    textbox.insert("1.0", mesaj)
    textbox.configure(state="normal")  # Seçilebilir ve kopyalanabilir
    
    # Buton çerçevesi
    btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
    btn_frame.pack(pady=(0, 15))
    
    # Tümünü Kopyala butonu
    def kopyala():
        popup.clipboard_clear()
        popup.clipboard_append(mesaj)
        kopyala_btn.configure(text="Kopyalandı", fg_color=theme.SUCCESS, hover_color=theme.SUCCESS_HOVER)
        popup.after(1500, lambda: kopyala_btn.configure(text="Tümünü Kopyala", fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER))

    # width=140: uzun etiket için tek kullanımlık genişlik (token'a girmeyecek kadar özel)
    kopyala_btn = ctk.CTkButton(btn_frame, text="Tümünü Kopyala", width=140, height=theme.BTN_H_SM,
                                corner_radius=theme.CORNER_BTN, font=theme.FONT_SMALL(),
                                command=kopyala, fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                                text_color=theme.TEXT_ON_ACCENT)
    kopyala_btn.pack(side="left", padx=5)

    ctk.CTkButton(btn_frame, text="Tamam", width=theme.BTN_W_SM, height=theme.BTN_H_SM,
                  corner_radius=theme.CORNER_BTN, font=theme.FONT_SMALL(),
                  fg_color="transparent", border_width=theme.BORDER_WIDTH,
                  border_color=theme.BORDER, hover_color=theme.BG_HOVER,
                  text_color=theme.TEXT_SECONDARY,
                  command=popup.destroy).pack(side="left", padx=5)

    # Yükseklik: widget'lar yerleştikten sonra gerçek gereksinimden hesaplanır
    # (şerit + metin kutusu + buton satırı + tüm boşluklar). Böylece butonlar
    # kısa mesajlarda da hiçbir zaman kırpılmaz.
    popup.update_idletasks()
    gerekli_height = popup.winfo_reqheight()
    # Uzun mesajlarda metin kutusunu satır sayısına göre büyüt (~20 piksel/satır)
    ekstra_height = max(0, satir_sayisi * 20 - textbox.winfo_reqheight())
    height = min(max(min_height, gerekli_height + ekstra_height), max_height)
    height = max(height, gerekli_height)  # taban asla içeriğin altına inmesin

    # Pencereyi ortala
    x = state.app.winfo_x() + (state.app.winfo_width() // 2) - (width // 2)
    y = state.app.winfo_y() + (state.app.winfo_height() // 2) - (height // 2)
    popup.geometry(f"{width}x{height}+{x}+{y}")
    # Kullanıcı elle küçültse bile butonlar görünür kalsın
    popup.minsize(min_width, gerekli_height)

# Kontrol işlemini arka planda başlatan yardımcı
def _kontrol_baslat(popup, kontrol_btn, kontrol_func, *args):
    """Kontrol Et butonunu meşgul duruma alır ve işi arka planda başlatır.

    Popup sonuç ana thread'e döndüğünde kapatılır; böylece işlem sürerken
    kullanıcı ikinci kez gönderim yapamaz ve arayüz donmaz.
    """
    kontrol_btn.configure(state="disabled", text="İşleniyor...",
                          fg_color=theme.BG_PRESSED, hover_color=theme.BG_PRESSED,
                          text_color=theme.TEXT_SECONDARY)

    def bitti_callback():
        try:
            if popup.winfo_exists():
                popup.destroy()
        except Exception:
            # Pencere zaten kapatılmış olabilir
            pass

    kontrol_func(*args, bitti_callback=bitti_callback)


def _kontrol_sonucunu_goster(sonuc, bitti_callback=None):
    """Arka plandan dönen sonucu ana thread'de popup olarak gösterir"""
    if bitti_callback:
        bitti_callback()

    try:
        sonuc_dosyasi, mesaj = sonuc
    except (TypeError, ValueError):
        sonuc_dosyasi, mesaj = None, "Hata oluştu:\nBeklenmeyen bir sonuç alındı."

    if sonuc_dosyasi:
        popup = ctk.CTkToplevel(state.app)
        popup.title("Kontrol Sonucu")
        popup.geometry("420x360")
        popup.resizable(False, False)
        popup.attributes("-topmost", True)
        popup.grab_set()
        popup.configure(fg_color=theme.BG_ROOT)

        # Başarı başlığı (ikon kutusu + başlık)
        _popup_basligi(popup, "Kontrol Sonucu", "karşılaştırma tamamlandı",
                       theme.SUCCESS, icons.onay)

        mesaj_karti = ctk.CTkFrame(popup, fg_color=theme.BG_SURFACE, corner_radius=theme.CORNER_POPUP,
                                   border_width=theme.BORDER_WIDTH, border_color=theme.BORDER)
        mesaj_karti.pack(fill="both", expand=True, padx=26, pady=(0, 18))
        ctk.CTkLabel(mesaj_karti, text=mesaj, font=theme.FONT_MESSAGE(),
                     text_color=theme.TEXT_SECONDARY, justify="left",
                     wraplength=330).pack(padx=16, pady=14, anchor="w")

        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(pady=(0, 20))

        ctk.CTkButton(btn_frame, text="Dosyayı Aç",
                      image=icons.tablo(theme.ICON_SM, theme.TEXT_ON_ACCENT), compound="left",
                      width=theme.BTN_W_MD, height=theme.BTN_H_MD,
                      corner_radius=theme.CORNER_BTN, font=theme.FONT_BODY_BOLD(),
                      fg_color=theme.SUCCESS, hover_color=theme.SUCCESS_HOVER,
                      text_color=theme.TEXT_ON_ACCENT,
                      command=lambda: [popup.destroy(), dosyayi_ac()]).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Kapat", width=theme.BTN_W_SM, height=theme.BTN_H_MD,
                      corner_radius=theme.CORNER_BTN, font=theme.FONT_SMALL(),
                      fg_color="transparent", border_width=theme.BORDER_WIDTH, border_color=theme.BORDER,
                      hover_color=theme.BG_HOVER, text_color=theme.TEXT_SECONDARY,
                      command=popup.destroy).pack(side="left", padx=5)

        # Yükseklik widget'lar yerleştikten sonra gerçek gereksinimden hesaplanır;
        # özet mesajı uzadığında alttaki butonlar kırpılmaz (bkz. show_info_popup).
        popup.update_idletasks()
        height = max(360, popup.winfo_reqheight())
        x = state.app.winfo_x() + (state.app.winfo_width() // 2) - 210
        y = state.app.winfo_y() + (state.app.winfo_height() // 2) - (height // 2)
        popup.geometry(f"420x{height}+{x}+{y}")
    else:
        show_info_popup(mesaj)


# Kontrol callback fabrika fonksiyonu
def _kontrol_sonuc_popup_goster(karsilastir_func):
    """Karşılaştırmayı arka planda çalıştırıp sonucu popup olarak gösteren callback oluşturur"""
    def callback(zirve_path, portal_path, bitti_callback=None):
        def calis():
            try:
                sonuc = karsilastir_func(zirve_path, portal_path)
            except Exception as hata:
                import traceback
                sonuc = (None, f"Hata oluştu:\n{hata}\n\n{traceback.format_exc()}")
            # Tk widget'larına yalnızca ana thread'den dokunulabilir
            popup_menu.ana_threadde_calistir(
                state.app, lambda: _kontrol_sonucunu_goster(sonuc, bitti_callback))

        threading.Thread(target=calis, daemon=True).start()
    return callback

# Kontrol callback'leri - fabrika fonksiyonuyla oluştur
# Uyumsoft için Zirve/Verim callback'leri
def uyumsoft_zirve_callback(zirve_path, portal_paths):
    return uyumsoft_zirve_karsilastir(zirve_path, portal_paths)

def uyumsoft_verim_callback(zirve_path, portal_paths):
    return uyumsoft_verim_karsilastir(zirve_path, portal_paths)

uyumsoft_zirve_kontrol_callback = _kontrol_sonuc_popup_goster(uyumsoft_zirve_callback)
uyumsoft_verim_kontrol_callback = _kontrol_sonuc_popup_goster(uyumsoft_verim_callback)
vega_kontrol_callback = _kontrol_sonuc_popup_goster(vega_karsilastir)
izibiz_kontrol_callback = _kontrol_sonuc_popup_goster(izibiz_karsilastir)

# Hızlıbilişim için Alış/Satış callback'leri
def hizlibilisim_alis_callback(zirve_path, portal_path):
    return hizlibilisim_karsilastir(zirve_path, portal_path, fatura_yonu='alis')

def hizlibilisim_satis_callback(zirve_path, portal_paths):
    return hizlibilisim_karsilastir(zirve_path, portal_paths, fatura_yonu='satis')

hizlibilisim_alis_kontrol_callback = _kontrol_sonuc_popup_goster(hizlibilisim_alis_callback)
hizlibilisim_satis_kontrol_callback = _kontrol_sonuc_popup_goster(hizlibilisim_satis_callback)

# Hepsiburada/Trendyol için callback (otomatik tespit)
hb_ve_trn_kontrol_callback = _kontrol_sonuc_popup_goster(hb_ve_trn_karsilastir)

# Henüz aktif olmayan özellikler için ortak fonksiyon
def henuz_aktif_degil():
    show_info_popup("Bu özellik henüz aktif değil.")

# Kontrol fonksiyonları
def kontrol_uyumsoft_zirve():
    show_kontrol_popup("Uyumsoft - Zirve", uyumsoft_zirve_kontrol_callback)

def show_uyumsoft_verim_popup():
    """Uyumsoft Verim kontrol popup'ı - Verim ve Uyumsoft dosyaları"""
    popup = ctk.CTkToplevel(state.app)
    popup.title("Uyumsoft - Verim Kontrol")
    popup.geometry("540x440")
    popup.resizable(False, False)
    popup.attributes("-topmost", True)
    popup.grab_set()
    popup.configure(fg_color=theme.BG_ROOT)
    
    popup.update_idletasks()
    x = state.app.winfo_x() + (state.app.winfo_width() // 2) - 270
    y = state.app.winfo_y() + (state.app.winfo_height() // 2) - 220
    popup.geometry(f"540x440+{x}+{y}")
    
    secilen_dosyalar = {"verim": [], "portal": []}
    
    _popup_basligi(popup, "Uyumsoft - Verim Fatura Kontrol",
                   "Verim ve Uyumsoft Excel dosyalarını seçin",
                   theme.PROVIDER_COLORS["uyumsoft"], icons.karsilastir)

    files_frame = ctk.CTkFrame(popup, fg_color="transparent")
    files_frame.pack(fill="x", padx=26)

    # Sol: Verim Excel
    verim_frame, verim_status = _dosya_karti(files_frame, 0, "Verim Excel", theme.ACCENT)

    def select_verim():
        popup.withdraw()
        popup.update_idletasks()
        dosyalar = filedialog.askopenfilenames(title="Verim Excel Dosyaları Seçin", filetypes=[("Excel Dosyaları", "*.xlsx *.xls")])
        popup.deiconify()
        popup.lift()
        popup.focus_force()
        if dosyalar:
            secilen_dosyalar["verim"] = list(dosyalar)
            dosya_sayisi = len(secilen_dosyalar["verim"])
            if dosya_sayisi == 1:
                dosya_adi = os.path.basename(dosyalar[0])
                if len(dosya_adi) > 18:
                    dosya_adi = dosya_adi[:15] + "..."
                verim_status.configure(text=f"{dosya_adi}", text_color=theme.SUCCESS)
            else:
                verim_status.configure(text=f"{dosya_sayisi} dosya seçildi", text_color=theme.SUCCESS)
            check_ready()
    
    _sec_butonu(verim_frame, "Dosyalar Seç", select_verim).pack(pady=(0, 16))

    # Sağ: Uyumsoft Excel
    portal_frame, portal_status = _dosya_karti(files_frame, 1, "Uyumsoft Excel",
                                               theme.PROVIDER_COLORS["uyumsoft"])

    def select_portal():
        popup.withdraw()
        popup.update_idletasks()
        dosyalar = filedialog.askopenfilenames(title="Uyumsoft Excel Dosyaları Seçin", filetypes=[("Excel Dosyaları", "*.xlsx *.xls")])
        popup.deiconify()
        popup.lift()
        popup.focus_force()
        if dosyalar:
            secilen_dosyalar["portal"] = list(dosyalar)
            dosya_sayisi = len(secilen_dosyalar["portal"])
            if dosya_sayisi == 1:
                dosya_adi = os.path.basename(dosyalar[0])
                if len(dosya_adi) > 18:
                    dosya_adi = dosya_adi[:15] + "..."
                portal_status.configure(text=f"{dosya_adi}", text_color=theme.SUCCESS)
            else:
                portal_status.configure(text=f"{dosya_sayisi} dosya seçildi", text_color=theme.SUCCESS)
            check_ready()
    
    _sec_butonu(portal_frame, "Dosyalar Seç", select_portal).pack(pady=(0, 16))

    def run_kontrol():
        if secilen_dosyalar["verim"] and secilen_dosyalar["portal"]:
            _kontrol_baslat(popup, kontrol_btn, uyumsoft_verim_kontrol_callback,
                            secilen_dosyalar["verim"], secilen_dosyalar["portal"])

    kontrol_btn = _eylem_butonu(popup, run_kontrol)
    kontrol_btn.pack(pady=(24, 12))

    def check_ready():
        if secilen_dosyalar["verim"] and secilen_dosyalar["portal"]:
            _eylem_butonu_hazir(kontrol_btn, theme.SUCCESS)

    _iptal_butonu(popup).pack()

def kontrol_uyumsoft_verim():
    show_uyumsoft_verim_popup()

def kontrol_vega():
    show_kontrol_popup("Vega", vega_kontrol_callback)

def kontrol_izibiz():
    show_kontrol_popup("İzibiz", izibiz_kontrol_callback)

def kontrol_hb_ve_trn():
    show_kontrol_popup("Hepsiburada/Trendyol", hb_ve_trn_kontrol_callback)

# Hızlıbilişim için özel kontrol popup'ı (tek dosya seçimi)
def show_hizlibilisim_alis_popup():
    """Hızlıbilişim Alış kontrol popup'ı - tek dosya seçimi"""
    popup = ctk.CTkToplevel(state.app)
    popup.title("Hızlıbilişim Alış Kontrol")
    popup.geometry("520x440")
    popup.resizable(False, False)
    popup.attributes("-topmost", True)
    popup.grab_set()
    popup.configure(fg_color=theme.BG_ROOT)
    
    popup.update_idletasks()
    x = state.app.winfo_x() + (state.app.winfo_width() // 2) - 260
    y = state.app.winfo_y() + (state.app.winfo_height() // 2) - 220
    popup.geometry(f"520x440+{x}+{y}")
    
    secilen_dosyalar = {"zirve": None, "portal": None}
    
    _popup_basligi(popup, "Hızlıbilişim Alış Fatura Kontrol",
                   "Gelen faturalar için karşılaştırma (tek dosya)",
                   theme.PROVIDER_COLORS["hizlibilisim"], icons.karsilastir)

    files_frame = ctk.CTkFrame(popup, fg_color="transparent")
    files_frame.pack(fill="x", padx=26)

    # Zirve Excel
    zirve_frame, zirve_status = _dosya_karti(files_frame, 0, "Zirve Excel", theme.ACCENT)

    def select_zirve():
        popup.withdraw()
        popup.update_idletasks()
        dosya = filedialog.askopenfilename(title="Zirve Excel Dosyası Seçin", filetypes=[("Excel Dosyaları", "*.xlsx *.xls")])
        popup.deiconify()
        popup.lift()
        popup.focus_force()
        if dosya:
            secilen_dosyalar["zirve"] = dosya
            dosya_adi = os.path.basename(dosya)
            if len(dosya_adi) > 20:
                dosya_adi = dosya_adi[:17] + "..."
            zirve_status.configure(text=f"{dosya_adi}", text_color=theme.SUCCESS)
            check_ready()
    
    _sec_butonu(zirve_frame, "Dosya Seç", select_zirve).pack(pady=(0, 16))

    # Portal Excel (Tek dosya)
    portal_frame, portal_status = _dosya_karti(files_frame, 1, "Gelen Fatura Excel",
                                               theme.PROVIDER_COLORS["hizlibilisim"])

    def select_portal():
        popup.withdraw()
        popup.update_idletasks()
        dosya = filedialog.askopenfilename(title="Gelen Fatura Excel Dosyası Seçin", filetypes=[("Excel Dosyaları", "*.xlsx *.xls")])
        popup.deiconify()
        popup.lift()
        popup.focus_force()
        if dosya:
            secilen_dosyalar["portal"] = dosya
            dosya_adi = os.path.basename(dosya)
            if len(dosya_adi) > 18:
                dosya_adi = dosya_adi[:15] + "..."
            portal_status.configure(text=f"{dosya_adi}", text_color=theme.SUCCESS)
            check_ready()
    
    _sec_butonu(portal_frame, "Dosya Seç", select_portal).pack(pady=(0, 16))

    def run_kontrol():
        if secilen_dosyalar["zirve"] and secilen_dosyalar["portal"]:
            _kontrol_baslat(popup, kontrol_btn, hizlibilisim_alis_kontrol_callback,
                            secilen_dosyalar["zirve"], secilen_dosyalar["portal"])

    kontrol_btn = _eylem_butonu(popup, run_kontrol)
    kontrol_btn.pack(pady=(24, 12))

    def check_ready():
        if secilen_dosyalar["zirve"] and secilen_dosyalar["portal"]:
            _eylem_butonu_hazir(kontrol_btn, theme.SUCCESS)

    _iptal_butonu(popup).pack()

def show_hizlibilisim_satis_popup():
    """Hızlıbilişim Satış kontrol popup'ı - çoklu dosya seçimi"""
    show_kontrol_popup("Hızlıbilişim Satış", hizlibilisim_satis_kontrol_callback)

def logo_hizli_otomatik_ayir_callback(zirve_path, portal_paths):
    """Portal dosyalarını otomatik tespit ederek Logo ve Hızlıbilişim olarak ayırır"""
    logo_portal_paths = []
    hizli_portal_paths = []
    
    # Her dosyayı oku ve tipini tespit et
    for path in portal_paths:
        try:
            df = pd.read_excel(path, nrows=0)  # Sadece sütun isimlerini oku
            sutunlar = df.columns.tolist()
            
            # Logo tespiti: "Toplam Tutar" ve "KDV Toplamı" varsa Logo
            if 'Toplam Tutar' in sutunlar and 'KDV Toplamı' in sutunlar:
                logo_portal_paths.append(path)
            # Hızlıbilişim tespiti: "FaturaNo" veya "PayableAmount" veya "OdenecekTutar" varsa Hızlı
            elif 'FaturaNo' in sutunlar or 'PayableAmount' in sutunlar or 'OdenecekTutar' in sutunlar:
                hizli_portal_paths.append(path)
            else:
                # Belirsiz, Logo olarak kabul et
                logo_portal_paths.append(path)
        except:
            # Hata durumunda Logo olarak kabul et
            logo_portal_paths.append(path)
    
    # Kombine karşılaştırma yap
    return logo_hizli_kombine_karsilastir(zirve_path, logo_portal_paths, hizli_portal_paths)

logo_hizli_kontrol_callback = _kontrol_sonuc_popup_goster(logo_hizli_otomatik_ayir_callback)

def show_logo_kontrol_popup():
    """Logo kontrol popup'ı - standart arayüz (arka planda Hızlıbilişim de kontrol edilir)"""
    show_kontrol_popup("Logo", logo_hizli_kontrol_callback)

def show_hizlibilisim_menu(button):
    """Hızlıbilişim için Alış/Satış popup menüsü göster"""
    popup_menu.popup_menu_kapat()
    
    x = button.winfo_rootx()
    y = button.winfo_rooty() + button.winfo_height() + 5
    
    w, h = popup_menu.menu_olcusu(2)
    state.aktif_menu = ctk.CTkToplevel(state.app)
    state.aktif_menu.geometry(f"{w}x{h}+{x}+{y}")
    state.aktif_menu.overrideredirect(True)
    state.aktif_menu.attributes("-topmost", True)
    govde = popup_menu.menu_govdesi(state.aktif_menu)

    popup_menu.menu_butonu(govde, "Alış Faturası",
                           lambda: [popup_menu.popup_menu_kapat(), show_hizlibilisim_alis_popup()],
                           ilk=True)
    popup_menu.menu_butonu(govde, "Satış Faturası",
                           lambda: [popup_menu.popup_menu_kapat(), show_hizlibilisim_satis_popup()],
                           son=True)
    popup_menu.menu_boyutunu_ayarla(state.aktif_menu, x, y)

    def activate_popup():
        state.popup_aktif = True
    state.app.after(150, activate_popup)
    
    state.aktif_menu.bind("<Escape>", lambda e: popup_menu.popup_menu_kapat())

def show_uyumsoft_menu(button):
    """Uyumsoft için Zirve/Verim popup menüsü göster"""
    popup_menu.popup_menu_kapat()
    
    x = button.winfo_rootx()
    y = button.winfo_rooty() + button.winfo_height() + 5
    
    w, h = popup_menu.menu_olcusu(2)
    state.aktif_menu = ctk.CTkToplevel(state.app)
    state.aktif_menu.geometry(f"{w}x{h}+{x}+{y}")
    state.aktif_menu.overrideredirect(True)
    state.aktif_menu.attributes("-topmost", True)
    govde = popup_menu.menu_govdesi(state.aktif_menu)

    popup_menu.menu_butonu(govde, "Zirve",
                           lambda: [popup_menu.popup_menu_kapat(), kontrol_uyumsoft_zirve()],
                           ilk=True)
    popup_menu.menu_butonu(govde, "Verim",
                           lambda: [popup_menu.popup_menu_kapat(), kontrol_uyumsoft_verim()],
                           son=True)
    popup_menu.menu_boyutunu_ayarla(state.aktif_menu, x, y)

    def activate_popup():
        state.popup_aktif = True
    state.app.after(150, activate_popup)
    
    state.aktif_menu.bind("<Escape>", lambda e: popup_menu.popup_menu_kapat())
