import customtkinter as ctk

# ================= TEMA TOKENLERİ =================
# Tüm renk / font / boyut değerleri burada tanımlıdır.
# Sayfa ve logic dosyalarında ham (hex veya "red"/"green") değer kullanılmaz.

# --- Yüzeyler (arka planlar) ---
BG_ROOT     = "#1e1f24"
BG_SURFACE  = "#26272e"
BG_ELEVATED = "#2f313a"
BG_HOVER    = "#3a3c47"
BG_PRESSED  = "#454852"
BORDER      = "#3d3f49"

# --- Vurgu rengi (accent) ---
ACCENT          = "#5b7fdb"
ACCENT_HOVER    = "#4a6bc4"
ACCENT_PRESSED  = "#3d59ab"

# --- Anlamsal renkler ---
SUCCESS        = "#4fb477"
SUCCESS_HOVER  = "#3f9a63"
WARNING        = "#e0a750"
ERROR          = "#e0616b"

# --- Metin tonları (3 kademe) ---
TEXT_PRIMARY   = "#f0f1f4"
TEXT_SECONDARY = "#b7bac2"
TEXT_MUTED     = "#7d808c"

# --- Boyutlandırma tokenleri ---
BTN_W_SM, BTN_H_SM       = 100, 32   # "📂 Seç", "Tamam"
BTN_W_SM_WIDE            = 110       # "📂 Dosyalar Seç" (uzun etiket, BTN_H_SM ile birlikte)
BTN_W_MD, BTN_H_MD       = 150, 35   # popup menü butonları, "📂 Dosyayı Aç"
BTN_W_LG, BTN_H_LG       = 160, 40   # sayfa üzerindeki ana işlem butonları
BTN_W_ACTION, BTN_H_ACTION = 200, 40 # popup içindeki "🔍 Kontrol Et"
BTN_W_WIDE, BTN_H_WIDE   = 280, 40   # "📂 Oluşturulan Excel Dosyasını Aç"
BTN_W_BACK, BTN_H_BACK   = 80, 30    # "← Geri" (yükseklik "İptal" butonlarında da kullanılır)

CORNER_CARD  = 15
CORNER_TILE  = 12
CORNER_POPUP = 10

PAD_PAGE = 40
PAD_CARD = 25
PAD_ROW  = 8

# --- Font token fonksiyonları (CTkFont nesneleri app oluşturulmadan yaratılamaz, bu yüzden fonksiyon) ---
def FONT_TITLE():
    """28 bold - karşılama ekranı ana başlığı"""
    return ctk.CTkFont(size=28, weight="bold")

def FONT_PAGE_TITLE():
    """24 bold - XML / Kontrol sayfa başlıkları"""
    return ctk.CTkFont(size=24, weight="bold")

def FONT_H1():
    """22 bold - kart başlıkları"""
    return ctk.CTkFont(size=22, weight="bold")

def FONT_POPUP_TITLE():
    """20 bold - popup pencere başlıkları"""
    return ctk.CTkFont(size=20, weight="bold")

def FONT_H2():
    """18 - sayfa içi bölüm etiketleri (Logo, Uyumsoft ...)"""
    return ctk.CTkFont(size=18)

def FONT_SUBTITLE():
    """16 - alt başlık ve sonuç mesajı etiketleri"""
    return ctk.CTkFont(size=16)

def FONT_BODY():
    return ctk.CTkFont(size=14)

def FONT_BODY_BOLD():
    return ctk.CTkFont(size=14, weight="bold")

def FONT_MESSAGE():
    """13 - kontrol sonucu popup'ındaki çok satırlı özet metni"""
    return ctk.CTkFont(size=13)

def FONT_SMALL():
    return ctk.CTkFont(size=12)

def FONT_TINY():
    return ctk.CTkFont(size=11)

def FONT_ICON_LG():
    return ctk.CTkFont(size=56)

def FONT_ICON_MD():
    return ctk.CTkFont(size=36)
