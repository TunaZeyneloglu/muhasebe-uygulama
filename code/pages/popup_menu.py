import customtkinter as ctk

import state
import theme
from logic.xml_alis import klasor_sec_ve_isle_alis
from logic.xml_hepsiburada import klasor_sec_hepsiburada_alis, klasor_sec_hepsiburada_satis
from logic.xml_mustahsil import klasor_sec_ve_isle_mustahsil
from logic.xml_mustahsil_makbuz import klasor_sec_ve_isle_mustahsil_xml
from logic.xml_trendyol import klasor_sec_trendyol_alis, klasor_sec_trendyol_satis
from logic.xml_uyumsoft import klasor_sec_ve_isle_uyumsoft
from pages import kontrol_popups

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
    
    state.aktif_menu = ctk.CTkToplevel(state.app)
    state.aktif_menu.geometry(f"160x90+{x}+{y}")
    state.aktif_menu.overrideredirect(True)
    state.aktif_menu.attributes("-topmost", True)
    state.aktif_menu.configure(fg_color=theme.BG_ELEVATED)
    
    ctk.CTkButton(state.aktif_menu, text="📥 Alış Faturası", width=theme.BTN_W_MD, height=theme.BTN_H_MD,
                  fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                  command=lambda: [popup_menu_kapat(), alis_func(label_sonuc, buton_ac)]).pack(pady=(8, 4), padx=5)
    ctk.CTkButton(state.aktif_menu, text="📤 Satış Faturası", width=theme.BTN_W_MD, height=theme.BTN_H_MD,
                  fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                  command=lambda: [popup_menu_kapat(), satis_func(label_sonuc, buton_ac)]).pack(pady=(0, 8), padx=5)
    
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
    """Müstahsil butonu için 3 seçenekli popup menü: Borsa, Alış XML, Müstahsil XML"""
    popup_menu_kapat()
    
    x = button.winfo_rootx()
    y = button.winfo_rooty() + button.winfo_height() + 5
    
    state.aktif_menu = ctk.CTkToplevel(state.app)
    state.aktif_menu.geometry(f"160x135+{x}+{y}")
    state.aktif_menu.overrideredirect(True)
    state.aktif_menu.attributes("-topmost", True)
    state.aktif_menu.configure(fg_color=theme.BG_ELEVATED)
    
    ctk.CTkButton(state.aktif_menu, text="🥔 Müstahsil Borsa", width=theme.BTN_W_MD, height=theme.BTN_H_MD,
                  fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                  command=lambda: [popup_menu_kapat(), klasor_sec_ve_isle_mustahsil(label_sonuc, buton_ac)]).pack(pady=(8, 4), padx=5)
    ctk.CTkButton(state.aktif_menu, text="📦 Alış XML", width=theme.BTN_W_MD, height=theme.BTN_H_MD,
                  fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                  command=lambda: [popup_menu_kapat(), klasor_sec_ve_isle_alis(label_sonuc, buton_ac)]).pack(pady=(0, 4), padx=5)
    ctk.CTkButton(state.aktif_menu, text="📋 Müstahsil XML", width=theme.BTN_W_MD, height=theme.BTN_H_MD,
                  fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                  command=lambda: [popup_menu_kapat(), klasor_sec_ve_isle_mustahsil_xml(label_sonuc, buton_ac)]).pack(pady=(0, 8), padx=5)
    
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
    
    state.aktif_menu = ctk.CTkToplevel(state.app)
    state.aktif_menu.geometry(f"160x90+{x}+{y}")
    state.aktif_menu.overrideredirect(True)
    state.aktif_menu.attributes("-topmost", True)
    state.aktif_menu.configure(fg_color=theme.BG_ELEVATED)
    
    ctk.CTkButton(state.aktif_menu, text="🏢 ACE", width=theme.BTN_W_MD, height=theme.BTN_H_MD,
                  fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                  command=lambda: [popup_menu_kapat(), kontrol_popups.show_info_popup("Bu özellik henüz aktif değil.")]).pack(pady=(8, 4), padx=5)
    ctk.CTkButton(state.aktif_menu, text="🏭 Turkay", width=theme.BTN_W_MD, height=theme.BTN_H_MD,
                  fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                  command=lambda: [popup_menu_kapat(), klasor_sec_ve_isle_uyumsoft(label_sonuc, buton_ac)]).pack(pady=(0, 8), padx=5)
    
    # Popup aktif olarak işaretle (biraz gecikmeyle)
    def activate_popup():
        state.popup_aktif = True
    state.app.after(150, activate_popup)
    
    # Escape tuşu ile kapatma
    state.aktif_menu.bind("<Escape>", lambda e: popup_menu_kapat())
