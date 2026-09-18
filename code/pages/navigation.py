import state

# ----------------- Sayfa Yönetimi -----------------

def show_welcome():
    """Ana karşılama sayfasını göster"""
    state.xml_frame.pack_forget()
    state.kontrol_frame.pack_forget()
    state.welcome_frame.pack(expand=True, fill="both")

def show_xml():
    """XML aktarma sayfasını göster"""
    state.welcome_frame.pack_forget()
    state.kontrol_frame.pack_forget()
    state.xml_frame.pack(expand=True, fill="both")

def show_kontrol():
    """Kontrol sayfasını göster"""
    state.welcome_frame.pack_forget()
    state.xml_frame.pack_forget()
    state.kontrol_frame.pack(expand=True, fill="both")
