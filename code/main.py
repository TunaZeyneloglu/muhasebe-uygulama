import warnings

import customtkinter as ctk

import state
import theme
import updater
from pages import navigation, popup_menu, welcome, xml_page, kontrol_page

# openpyxl uyarılarını sustur
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Tema ayarı
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

state.app = ctk.CTk()
state.app.title("Fatura Yönetim Sistemi")
state.app.geometry("650x610")
state.app.resizable(False, False)
state.app.configure(fg_color=theme.BG_ROOT)

state.welcome_frame = welcome.build_welcome_page(state.app)
state.xml_frame = xml_page.build_xml_page(state.app)
state.kontrol_frame = kontrol_page.build_kontrol_page(state.app)

# Ana pencereye tıklanınca popup menüyü kapat
state.app.bind_all("<Button-1>", popup_menu.on_app_click)

# Açılışta sürüm kontrolü
updater.check_for_update()

# Karşılama sayfasını göster
navigation.show_welcome()

state.app.mainloop()
