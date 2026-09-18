import customtkinter as ctk

# ================= TEMA TOKENLERİ =================
# Tüm renk / font / boyut değerleri burada tanımlıdır.
# Sayfa ve logic dosyalarında ham (hex veya "red"/"green") değer kullanılmaz.
#
# Görsel dil: "koyu hero" (V2) - derin koyu zemin, hafif yükseltilmiş yüzeyler,
# ince saç teli çerçeveler ve doygun mavi/yeşil vurgu renkleri.

# --- Yüzeyler (arka planlar) ---
BG_ROOT     = "#0d1014"   # pencere zemini
BG_SURFACE  = "#161a20"   # kart / araç çubuğu yüzeyi
BG_ELEVATED = "#1c2129"   # kart içi yükseltilmiş satır
BG_HOVER    = "#232a34"
BG_PRESSED  = "#2b3341"
BORDER      = "#262c35"   # saç teli çerçeve
BORDER_HI   = "#39414d"   # vurgulu çerçeve

# --- Vurgu rengi (accent) ---
ACCENT          = "#5b8cf5"
ACCENT_HOVER    = "#4a7ae8"
ACCENT_PRESSED  = "#3c66cc"
ACCENT_TINT     = "#1a2743"   # accent rengin koyu zemine karışmış hâli (ikon kutusu)

# --- Anlamsal renkler ---
SUCCESS        = "#35b98a"
SUCCESS_HOVER  = "#2ba077"
SUCCESS_TINT   = "#122c25"
WARNING        = "#e5a94e"
WARNING_TINT   = "#33260f"
ERROR          = "#f0616f"
ERROR_TINT     = "#351419"

# --- Metin tonları (3 kademe) ---
TEXT_PRIMARY   = "#eef1f6"
TEXT_SECONDARY = "#a3abb8"
TEXT_MUTED     = "#6b7480"
TEXT_ON_ACCENT = "#0d1014"   # doygun vurgu rengi üzerindeki metin

# --- Sağlayıcı renkleri (XML / Kontrol sayfasındaki satırlar) ---
# Aynı ikon ailesi, sağlayıcıya göre değişen ton: 8 rastgele glif yerine
# tek bir sistem olarak okunur.
PROVIDER_COLORS = {
    "logo":         "#5b8cf5",
    "uyumsoft":     "#7c6cf0",
    "mustahsil":    "#35b98a",
    "hizlibilisim": "#2fb3c9",
    "hepsiburada":  "#e5734e",
    "trendyol":     "#e5a94e",
    "vega":         "#b06cf0",
    "izibiz":       "#4fb0f5",
}


def karistir(renk, oran, zemin=BG_ROOT):
    """İki hex rengi karıştırır (oran=0 -> zemin, oran=1 -> renk)."""
    def ayir(h):
        h = h.lstrip("#")
        return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)

    r1, g1, b1 = ayir(zemin)
    r2, g2, b2 = ayir(renk)
    return "#%02x%02x%02x" % (
        round(r1 + (r2 - r1) * oran),
        round(g1 + (g2 - g1) * oran),
        round(b1 + (b2 - b1) * oran),
    )


def TINT(renk):
    """İkon kutusu için rengin koyu zemine karışmış tonu."""
    return karistir(renk, 0.20)


def TINT_HOVER(renk):
    """İkon kutusunun hover'daki biraz daha canlı tonu."""
    return karistir(renk, 0.30)


# --- Boyutlandırma tokenleri ---
BTN_W_SM, BTN_H_SM       = 100, 32   # "Seç", "Tamam"
BTN_W_SM_WIDE            = 110       # "Dosyalar Seç" (uzun etiket, BTN_H_SM ile birlikte)
BTN_W_MD, BTN_H_MD       = 150, 35   # popup menü butonları, "Dosyayı Aç"
BTN_W_LG, BTN_H_LG       = 160, 40   # sayfa üzerindeki ana işlem butonları
BTN_W_ACTION, BTN_H_ACTION = 200, 40 # popup içindeki "Kontrol Et"
BTN_W_WIDE, BTN_H_WIDE   = 280, 44   # "Oluşturulan Excel Dosyasını Aç"
BTN_W_BACK, BTN_H_BACK   = 92, 32    # "← Geri" (yükseklik "İptal" butonlarında da kullanılır)
BTN_W_ROW, BTN_H_ROW     = 118, 36   # sağlayıcı satırındaki "Aktar" / "Kontrol" butonu
BTN_W_MENU, BTN_H_MENU   = 186, 36   # açılır menü satırları
MENU_W                   = 204       # açılır menü penceresi genişliği

CORNER_CARD  = 16   # büyük kart
CORNER_TILE  = 12   # sağlayıcı satırı / ikon kutusu
CORNER_POPUP = 12
CORNER_BTN   = 10   # buton köşesi
CORNER_CHIP  = 8    # sürüm rozeti
CORNER_BAR   = 2    # popup üstündeki ince renkli şerit

BORDER_WIDTH = 1    # kart / satır / ikincil buton çerçeve kalınlığı
CHIP_SIZE    = 66   # karşılama kartındaki ikon kutusunun kenar uzunluğu
TILE_SIZE    = 44   # sağlayıcı satırındaki ikon kutusunun kenar uzunluğu
HEAD_TILE    = 52   # sayfa başlığındaki ikon kutusunun kenar uzunluğu
BAR_H        = 4    # popup üstündeki renkli şeridin yüksekliği
DIVIDER_H    = 1    # başlık altındaki ince ayırıcı çizgi
TOOLBAR_H    = 60   # üst uygulama çubuğu
STATUSBAR_H  = 32   # alt durum çubuğu

ICON_XL      = 34   # karşılama kartı ikonu
ICON_LG      = 26   # sayfa başlığı ikonu
ICON_MD      = 22   # sağlayıcı satırı ikonu
ICON_SM      = 14   # buton içi ok / chevron
ICON_XS      = 12   # madde işareti onayı

PAD_PAGE  = 40   # sayfa kenar boşluğu (eski sayfalarla uyum için korunuyor)
PAD_STAGE = 46   # yeni sayfa gövdesi yatay kenar boşluğu
PAD_CARD  = 26   # kart iç boşluğu
PAD_ROW   = 5    # satırlar arası dikey boşluk
PAD_TILE  = 14   # sağlayıcı satırının iç boşluğu


# --- Font token fonksiyonları (CTkFont nesneleri app oluşturulmadan yaratılamaz, bu yüzden fonksiyon) ---
def FONT_HERO():
    """34 bold - karşılama ekranındaki hero başlığı"""
    return ctk.CTkFont(size=34, weight="bold")

def FONT_TITLE():
    """28 bold - büyük başlık"""
    return ctk.CTkFont(size=28, weight="bold")

def FONT_PAGE_TITLE():
    """24 bold - XML / Kontrol sayfa başlıkları"""
    return ctk.CTkFont(size=24, weight="bold")

def FONT_H1():
    """21 bold - kart başlıkları"""
    return ctk.CTkFont(size=21, weight="bold")

def FONT_POPUP_TITLE():
    """20 bold - popup pencere başlıkları"""
    return ctk.CTkFont(size=20, weight="bold")

def FONT_H2():
    """16 bold - sağlayıcı satırı başlıkları (Logo, Uyumsoft ...)"""
    return ctk.CTkFont(size=16, weight="bold")

def FONT_SUBTITLE():
    """16 - alt başlık ve sonuç mesajı etiketleri"""
    return ctk.CTkFont(size=16)

def FONT_BODY():
    return ctk.CTkFont(size=14)

def FONT_BODY_BOLD():
    return ctk.CTkFont(size=14, weight="bold")

def FONT_LINK():
    """13 bold - kart altındaki 'Aç' bağlantı butonu"""
    return ctk.CTkFont(size=13, weight="bold")

def FONT_MESSAGE():
    """13 - kontrol sonucu popup'ındaki çok satırlı özet metni"""
    return ctk.CTkFont(size=13)

def FONT_SMALL():
    return ctk.CTkFont(size=12)

def FONT_TINY():
    return ctk.CTkFont(size=11)

def FONT_EYEBROW():
    """10 bold - hero üstündeki küçük bölüm etiketi (ÇALIŞMA ALANI)"""
    return ctk.CTkFont(size=10, weight="bold")

def FONT_BRAND():
    """14 bold - araç çubuğundaki uygulama adı"""
    return ctk.CTkFont(size=14, weight="bold")

def FONT_ICON_LG():
    """40 bold - büyük monogram harfi (geriye dönük uyumluluk)"""
    return ctk.CTkFont(size=40, weight="bold")

def FONT_ICON_MD():
    """28 bold - küçük monogram / rozet harfi (geriye dönük uyumluluk)"""
    return ctk.CTkFont(size=28, weight="bold")
