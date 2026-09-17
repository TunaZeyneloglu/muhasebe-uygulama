import os
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog

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

# Kontrol popup fonksiyonu
def show_kontrol_popup(portal_adi, kontrol_func=None):
    """Excel karşılaştırma popup'ı göster"""
    popup = ctk.CTkToplevel(state.app)
    popup.title(f"{portal_adi} Kontrol")
    popup.geometry("500x360")
    popup.resizable(False, False)
    popup.attributes("-topmost", True)
    popup.grab_set()  # Modal yap
    
    # Pencereyi ortala
    popup.update_idletasks()
    x = state.app.winfo_x() + (state.app.winfo_width() // 2) - 250
    y = state.app.winfo_y() + (state.app.winfo_height() // 2) - 180
    popup.geometry(f"500x360+{x}+{y}")
    
    # Seçilen dosyaları takip et
    secilen_dosyalar = {"zirve": None, "portal": []}
    
    # Başlık
    ctk.CTkLabel(popup, text=f"🔍 {portal_adi} Fatura Kontrol", 
                 font=theme.FONT_POPUP_TITLE()).pack(pady=(20, 5))
    ctk.CTkLabel(popup, text="Karşılaştırmak için Excel dosyası/dosyaları seçin", 
                 font=theme.FONT_SMALL(), text_color=theme.TEXT_SECONDARY).pack(pady=(0, 20))
    
    # Dosya seçim alanları
    files_frame = ctk.CTkFrame(popup, fg_color="transparent")
    files_frame.pack(fill="x", padx=30)
    
    # Sol: Zirve Excel
    zirve_frame = ctk.CTkFrame(files_frame, fg_color=theme.BG_ELEVATED, corner_radius=theme.CORNER_POPUP)
    zirve_frame.grid(row=0, column=0, padx=(0, 10), sticky="nsew")
    files_frame.grid_columnconfigure(0, weight=1)
    
    ctk.CTkLabel(zirve_frame, text="Zirve Excel", font=theme.FONT_BODY_BOLD()).pack(pady=(15, 5))
    zirve_status = ctk.CTkLabel(zirve_frame, text="📄 Dosya seçilmedi", font=theme.FONT_TINY(), text_color=theme.TEXT_MUTED)
    zirve_status.pack(pady=(0, 10))
    
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
            zirve_status.configure(text=f"✅ {dosya_adi}", text_color=theme.SUCCESS)
            check_ready()
    
    ctk.CTkButton(zirve_frame, text="📂 Seç", width=theme.BTN_W_SM, height=theme.BTN_H_SM,
                  fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                  command=select_zirve).pack(pady=(0, 15))
    
    # Sağ: Portal Excel (Çoklu Seçim)
    portal_frame = ctk.CTkFrame(files_frame, fg_color=theme.BG_ELEVATED, corner_radius=theme.CORNER_POPUP)
    portal_frame.grid(row=0, column=1, padx=(10, 0), sticky="nsew")
    files_frame.grid_columnconfigure(1, weight=1)
    
    ctk.CTkLabel(portal_frame, text=f"{portal_adi} Excel", font=theme.FONT_BODY_BOLD()).pack(pady=(15, 5))
    portal_status = ctk.CTkLabel(portal_frame, text="📄 Dosya seçilmedi", font=theme.FONT_TINY(), text_color=theme.TEXT_MUTED)
    portal_status.pack(pady=(0, 10))
    
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
                portal_status.configure(text=f"✅ {dosya_adi}", text_color=theme.SUCCESS)
            else:
                portal_status.configure(text=f"✅ {dosya_sayisi} dosya seçildi", text_color=theme.SUCCESS)
            
            check_ready()
    
    ctk.CTkButton(portal_frame, text="📂 Dosyalar Seç", width=theme.BTN_W_SM_WIDE, height=theme.BTN_H_SM,
                  fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                  command=select_portal).pack(pady=(0, 15))
    
    # Kontrol Et butonu (başlangıçta devre dışı)
    def run_kontrol():
        if secilen_dosyalar["zirve"] and secilen_dosyalar["portal"]:
            popup.destroy()
            if kontrol_func:
                kontrol_func(secilen_dosyalar["zirve"], secilen_dosyalar["portal"])
            else:
                # Henüz implementasyon yok
                show_info_popup("Bu özellik henüz aktif değil.")
    
    kontrol_btn = ctk.CTkButton(popup, text="🔍 Kontrol Et", width=theme.BTN_W_ACTION, height=theme.BTN_H_ACTION, 
                                 command=run_kontrol, state="disabled",
                                 fg_color=theme.BG_PRESSED, hover_color=theme.BG_PRESSED)
    kontrol_btn.pack(pady=(25, 15))
    
    def check_ready():
        # Zirve seçilmiş ve Portal'de en az 1 dosya seçilmişse
        if secilen_dosyalar["zirve"] and len(secilen_dosyalar["portal"]) > 0:
            kontrol_btn.configure(state="normal", fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER)
    
    # İptal butonu
    ctk.CTkButton(popup, text="İptal", width=theme.BTN_W_SM, height=theme.BTN_H_BACK, 
                  fg_color="transparent", border_width=1, border_color=theme.BORDER,
                  hover_color=theme.BG_HOVER, command=popup.destroy).pack()

def show_info_popup(mesaj):
    """Bilgi popup'ı göster - Dinamik boyutlandırma ve kopyalanabilir metin"""
    popup = ctk.CTkToplevel(state.app)
    popup.title("Bilgi")
    popup.attributes("-topmost", True)
    popup.grab_set()
    
    # Mesaj uzunluğuna göre dinamik boyut hesapla
    satir_sayisi = mesaj.count('\n') + 1
    karakter_sayisi = max(len(satir) for satir in mesaj.split('\n'))
    
    # Minimum ve maksimum boyutlar
    min_width, max_width = 350, 800
    min_height, max_height = 200, 600
    
    # Genişlik: karakter sayısına göre (ortalama 8 piksel/karakter)
    width = min(max(min_width, karakter_sayisi * 8 + 100), max_width)
    # Yükseklik: satır sayısına göre (ortalama 25 piksel/satır) + ekstra boşluklar
    height = min(max(min_height, satir_sayisi * 25 + 180), max_height)
    
    popup.geometry(f"{width}x{height}")
    popup.resizable(True, True)  # Kullanıcı manuel olarak da büyütebilir
    
    # Pencereyi ortala
    popup.update_idletasks()
    x = state.app.winfo_x() + (state.app.winfo_width() // 2) - (width // 2)
    y = state.app.winfo_y() + (state.app.winfo_height() // 2) - (height // 2)
    popup.geometry(f"{width}x{height}+{x}+{y}")
    
    # İkon
    ctk.CTkLabel(popup, text="⚠️", font=theme.FONT_ICON_MD()).pack(pady=(15, 5))
    
    # Kopyalanabilir metin kutusu
    text_frame = ctk.CTkFrame(popup, fg_color="transparent")
    text_frame.pack(fill="both", expand=True, padx=20, pady=(0, 15))
    
    textbox = ctk.CTkTextbox(text_frame, font=theme.FONT_SMALL(), 
                             wrap="word", activate_scrollbars=True)
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
        kopyala_btn.configure(text="✓ Kopyalandı", fg_color=theme.SUCCESS)
        popup.after(1500, lambda: kopyala_btn.configure(text="📋 Tümünü Kopyala", fg_color=theme.ACCENT))
    
    # width=140: uzun etiket için tek kullanımlık genişlik (token'a girmeyecek kadar özel)
    kopyala_btn = ctk.CTkButton(btn_frame, text="📋 Tümünü Kopyala", width=140, height=theme.BTN_H_SM,
                                command=kopyala, fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER)
    kopyala_btn.pack(side="left", padx=5)
    
    ctk.CTkButton(btn_frame, text="Tamam", width=theme.BTN_W_SM, height=theme.BTN_H_SM,
                  fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                  command=popup.destroy).pack(side="left", padx=5)

# Kontrol callback fabrika fonksiyonu
def _kontrol_sonuc_popup_goster(karsilastir_func):
    """Kontrol sonucu popup'ı gösteren callback oluşturur"""
    def callback(zirve_path, portal_path):
        sonuc_dosyasi, mesaj = karsilastir_func(zirve_path, portal_path)
        
        if sonuc_dosyasi:
            popup = ctk.CTkToplevel(state.app)
            popup.title("Kontrol Sonucu")
            popup.geometry("420x320")
            popup.resizable(False, False)
            popup.attributes("-topmost", True)
            popup.grab_set()
            
            popup.update_idletasks()
            x = state.app.winfo_x() + (state.app.winfo_width() // 2) - 210
            y = state.app.winfo_y() + (state.app.winfo_height() // 2) - 160
            popup.geometry(f"420x320+{x}+{y}")
            
            ctk.CTkLabel(popup, text="🎉", font=theme.FONT_ICON_MD()).pack(pady=(20, 5))
            ctk.CTkLabel(popup, text=mesaj, font=theme.FONT_MESSAGE(), 
                         justify="left").pack(pady=(0, 20), padx=20)
            
            btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
            btn_frame.pack(pady=(0, 20))
            
            ctk.CTkButton(btn_frame, text="📂 Dosyayı Aç", width=theme.BTN_W_MD, height=theme.BTN_H_MD, 
                          fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                          command=lambda: [popup.destroy(), dosyayi_ac()]).pack(side="left", padx=5)
            ctk.CTkButton(btn_frame, text="Kapat", width=theme.BTN_W_SM, height=theme.BTN_H_MD, 
                          fg_color="transparent", border_width=1, border_color=theme.BORDER,
                          hover_color=theme.BG_HOVER,
                          command=popup.destroy).pack(side="left", padx=5)
        else:
            show_info_popup(mesaj)
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
    popup.geometry("520x380")
    popup.resizable(False, False)
    popup.attributes("-topmost", True)
    popup.grab_set()
    
    popup.update_idletasks()
    x = state.app.winfo_x() + (state.app.winfo_width() // 2) - 260
    y = state.app.winfo_y() + (state.app.winfo_height() // 2) - 190
    popup.geometry(f"520x380+{x}+{y}")
    
    secilen_dosyalar = {"verim": [], "portal": []}
    
    ctk.CTkLabel(popup, text="🔍 Uyumsoft - Verim Fatura Kontrol", 
                 font=theme.FONT_POPUP_TITLE()).pack(pady=(20, 5))
    ctk.CTkLabel(popup, text="Verim ve Uyumsoft Excel dosyalarını seçin", 
                 font=theme.FONT_SMALL(), text_color=theme.TEXT_SECONDARY).pack(pady=(0, 20))
    
    files_frame = ctk.CTkFrame(popup, fg_color="transparent")
    files_frame.pack(fill="x", padx=30)
    
    # Sol: Verim Excel
    verim_frame = ctk.CTkFrame(files_frame, fg_color=theme.BG_ELEVATED, corner_radius=theme.CORNER_POPUP)
    verim_frame.grid(row=0, column=0, padx=(0, 10), sticky="nsew")
    files_frame.grid_columnconfigure(0, weight=1)
    
    ctk.CTkLabel(verim_frame, text="Verim Excel", font=theme.FONT_BODY_BOLD()).pack(pady=(15, 5))
    verim_status = ctk.CTkLabel(verim_frame, text="📄 Dosya seçilmedi", font=theme.FONT_TINY(), text_color=theme.TEXT_MUTED)
    verim_status.pack(pady=(0, 10))
    
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
                verim_status.configure(text=f"✅ {dosya_adi}", text_color=theme.SUCCESS)
            else:
                verim_status.configure(text=f"✅ {dosya_sayisi} dosya seçildi", text_color=theme.SUCCESS)
            check_ready()
    
    ctk.CTkButton(verim_frame, text="📂 Dosyalar Seç", width=theme.BTN_W_SM_WIDE, height=theme.BTN_H_SM,
                  fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                  command=select_verim).pack(pady=(0, 15))
    
    # Sağ: Uyumsoft Excel
    portal_frame = ctk.CTkFrame(files_frame, fg_color=theme.BG_ELEVATED, corner_radius=theme.CORNER_POPUP)
    portal_frame.grid(row=0, column=1, padx=(10, 0), sticky="nsew")
    files_frame.grid_columnconfigure(1, weight=1)
    
    ctk.CTkLabel(portal_frame, text="Uyumsoft Excel", font=theme.FONT_BODY_BOLD()).pack(pady=(15, 5))
    portal_status = ctk.CTkLabel(portal_frame, text="📄 Dosya seçilmedi", font=theme.FONT_TINY(), text_color=theme.TEXT_MUTED)
    portal_status.pack(pady=(0, 10))
    
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
                portal_status.configure(text=f"✅ {dosya_adi}", text_color=theme.SUCCESS)
            else:
                portal_status.configure(text=f"✅ {dosya_sayisi} dosya seçildi", text_color=theme.SUCCESS)
            check_ready()
    
    ctk.CTkButton(portal_frame, text="📂 Dosyalar Seç", width=theme.BTN_W_SM_WIDE, height=theme.BTN_H_SM,
                  fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                  command=select_portal).pack(pady=(0, 15))
    
    def run_kontrol():
        if secilen_dosyalar["verim"] and secilen_dosyalar["portal"]:
            popup.destroy()
            uyumsoft_verim_kontrol_callback(secilen_dosyalar["verim"], secilen_dosyalar["portal"])
    
    kontrol_btn = ctk.CTkButton(popup, text="🔍 Kontrol Et", width=theme.BTN_W_ACTION, height=theme.BTN_H_ACTION, 
                                 command=run_kontrol, state="disabled",
                                 fg_color=theme.BG_PRESSED, hover_color=theme.BG_PRESSED)
    kontrol_btn.pack(pady=(25, 15))
    
    def check_ready():
        if secilen_dosyalar["verim"] and secilen_dosyalar["portal"]:
            kontrol_btn.configure(state="normal", fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER)
    
    ctk.CTkButton(popup, text="İptal", width=theme.BTN_W_SM, height=theme.BTN_H_BACK, 
                  fg_color="transparent", border_width=1, border_color=theme.BORDER,
                  hover_color=theme.BG_HOVER, command=popup.destroy).pack()

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
    popup.geometry("500x360")
    popup.resizable(False, False)
    popup.attributes("-topmost", True)
    popup.grab_set()
    
    popup.update_idletasks()
    x = state.app.winfo_x() + (state.app.winfo_width() // 2) - 250
    y = state.app.winfo_y() + (state.app.winfo_height() // 2) - 180
    popup.geometry(f"500x360+{x}+{y}")
    
    secilen_dosyalar = {"zirve": None, "portal": None}
    
    ctk.CTkLabel(popup, text="🔍 Hızlıbilişim Alış Fatura Kontrol", 
                 font=theme.FONT_POPUP_TITLE()).pack(pady=(20, 5))
    ctk.CTkLabel(popup, text="Gelen faturalar için karşılaştırma (tek dosya)", 
                 font=theme.FONT_SMALL(), text_color=theme.TEXT_SECONDARY).pack(pady=(0, 20))
    
    files_frame = ctk.CTkFrame(popup, fg_color="transparent")
    files_frame.pack(fill="x", padx=30)
    
    # Zirve Excel
    zirve_frame = ctk.CTkFrame(files_frame, fg_color=theme.BG_ELEVATED, corner_radius=theme.CORNER_POPUP)
    zirve_frame.grid(row=0, column=0, padx=(0, 10), sticky="nsew")
    files_frame.grid_columnconfigure(0, weight=1)
    
    ctk.CTkLabel(zirve_frame, text="Zirve Excel", font=theme.FONT_BODY_BOLD()).pack(pady=(15, 5))
    zirve_status = ctk.CTkLabel(zirve_frame, text="📄 Dosya seçilmedi", font=theme.FONT_TINY(), text_color=theme.TEXT_MUTED)
    zirve_status.pack(pady=(0, 10))
    
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
            zirve_status.configure(text=f"✅ {dosya_adi}", text_color=theme.SUCCESS)
            check_ready()
    
    ctk.CTkButton(zirve_frame, text="📂 Seç", width=theme.BTN_W_SM, height=theme.BTN_H_SM,
                  fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                  command=select_zirve).pack(pady=(0, 15))
    
    # Portal Excel (Tek dosya)
    portal_frame = ctk.CTkFrame(files_frame, fg_color=theme.BG_ELEVATED, corner_radius=theme.CORNER_POPUP)
    portal_frame.grid(row=0, column=1, padx=(10, 0), sticky="nsew")
    files_frame.grid_columnconfigure(1, weight=1)
    
    ctk.CTkLabel(portal_frame, text="Gelen Fatura Excel", font=theme.FONT_BODY_BOLD()).pack(pady=(15, 5))
    portal_status = ctk.CTkLabel(portal_frame, text="📄 Dosya seçilmedi", font=theme.FONT_TINY(), text_color=theme.TEXT_MUTED)
    portal_status.pack(pady=(0, 10))
    
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
            portal_status.configure(text=f"✅ {dosya_adi}", text_color=theme.SUCCESS)
            check_ready()
    
    ctk.CTkButton(portal_frame, text="📂 Dosya Seç", width=theme.BTN_W_SM_WIDE, height=theme.BTN_H_SM,
                  fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                  command=select_portal).pack(pady=(0, 15))
    
    def run_kontrol():
        if secilen_dosyalar["zirve"] and secilen_dosyalar["portal"]:
            popup.destroy()
            hizlibilisim_alis_kontrol_callback(secilen_dosyalar["zirve"], secilen_dosyalar["portal"])
    
    kontrol_btn = ctk.CTkButton(popup, text="🔍 Kontrol Et", width=theme.BTN_W_ACTION, height=theme.BTN_H_ACTION, 
                                 command=run_kontrol, state="disabled",
                                 fg_color=theme.BG_PRESSED, hover_color=theme.BG_PRESSED)
    kontrol_btn.pack(pady=(25, 15))
    
    def check_ready():
        if secilen_dosyalar["zirve"] and secilen_dosyalar["portal"]:
            kontrol_btn.configure(state="normal", fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER)
    
    ctk.CTkButton(popup, text="İptal", width=theme.BTN_W_SM, height=theme.BTN_H_BACK, 
                  fg_color="transparent", border_width=1, border_color=theme.BORDER,
                  hover_color=theme.BG_HOVER, command=popup.destroy).pack()

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
    
    state.aktif_menu = ctk.CTkToplevel(state.app)
    state.aktif_menu.geometry(f"160x90+{x}+{y}")
    state.aktif_menu.overrideredirect(True)
    state.aktif_menu.attributes("-topmost", True)
    state.aktif_menu.configure(fg_color=theme.BG_ELEVATED)
    
    ctk.CTkButton(state.aktif_menu, text="📥 Alış Faturası", width=theme.BTN_W_MD, height=theme.BTN_H_MD,
                  fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                  command=lambda: [popup_menu.popup_menu_kapat(), show_hizlibilisim_alis_popup()]).pack(pady=(8, 4), padx=5)
    ctk.CTkButton(state.aktif_menu, text="📤 Satış Faturası", width=theme.BTN_W_MD, height=theme.BTN_H_MD,
                  fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                  command=lambda: [popup_menu.popup_menu_kapat(), show_hizlibilisim_satis_popup()]).pack(pady=(0, 8), padx=5)
    
    def activate_popup():
        state.popup_aktif = True
    state.app.after(150, activate_popup)
    
    state.aktif_menu.bind("<Escape>", lambda e: popup_menu.popup_menu_kapat())

def show_uyumsoft_menu(button):
    """Uyumsoft için Zirve/Verim popup menüsü göster"""
    popup_menu.popup_menu_kapat()
    
    x = button.winfo_rootx()
    y = button.winfo_rooty() + button.winfo_height() + 5
    
    state.aktif_menu = ctk.CTkToplevel(state.app)
    state.aktif_menu.geometry(f"160x90+{x}+{y}")
    state.aktif_menu.overrideredirect(True)
    state.aktif_menu.attributes("-topmost", True)
    state.aktif_menu.configure(fg_color=theme.BG_ELEVATED)
    
    ctk.CTkButton(state.aktif_menu, text="📊 Zirve", width=theme.BTN_W_MD, height=theme.BTN_H_MD,
                  fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                  command=lambda: [popup_menu.popup_menu_kapat(), kontrol_uyumsoft_zirve()]).pack(pady=(8, 4), padx=5)
    ctk.CTkButton(state.aktif_menu, text="📊 Verim", width=theme.BTN_W_MD, height=theme.BTN_H_MD,
                  fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                  command=lambda: [popup_menu.popup_menu_kapat(), kontrol_uyumsoft_verim()]).pack(pady=(0, 8), padx=5)
    
    def activate_popup():
        state.popup_aktif = True
    state.app.after(150, activate_popup)
    
    state.aktif_menu.bind("<Escape>", lambda e: popup_menu.popup_menu_kapat())
