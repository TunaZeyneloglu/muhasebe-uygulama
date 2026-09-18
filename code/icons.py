"""Programatik ikon üretimi.

Uygulamada emoji veya harf monogramı kullanılmaz; tüm ikonlar burada
PIL.ImageDraw ile çizilir ve CTkImage olarak döndürülür. Çizim 4 kat
büyük yapılıp LANCZOS ile küçültülür, böylece kenarlar yumuşak olur.

Tüm fonksiyonlar (size, color) alır ve CTkImage döndürür. Sonuçlar
önbelleğe alınır; aynı ikon defalarca çizilmez.

CTkImage yalnızca bir Tk kök penceresi oluşturulduktan sonra
yaratılabileceği için ikonlar modül yüklenirken değil, sayfa kurulurken
(ilk çağrıda) üretilir.
"""
import customtkinter as ctk
from PIL import Image, ImageDraw

# Süper örnekleme katsayısı (4x çiz, 1x küçült)
_OLCEK = 4

# (fonksiyon adı, boyut, renk) -> CTkImage
_onbellek = {}


def _ciz(size, fn):
    """Verilen çizim fonksiyonunu 4x tuvalde çalıştırıp küçültür."""
    n = size * _OLCEK
    im = Image.new("RGBA", (n, n), (0, 0, 0, 0))
    fn(ImageDraw.Draw(im), n)
    return im.resize((size, size), Image.LANCZOS)


def _ctki(pil, size):
    return ctk.CTkImage(light_image=pil, dark_image=pil, size=(size, size))


def _ikon(ad, cizim, size, color):
    """Önbellekli ikon üretimi."""
    anahtar = (ad, size, color)
    if anahtar not in _onbellek:
        _onbellek[anahtar] = _ctki(_ciz(size, lambda d, n: cizim(d, n, color)), size)
    return _onbellek[anahtar]


def _kalinlik(n, bolen=16):
    return max(2, int(n / bolen))


# ----------------- Glifler -----------------

def _belge(d, n, c):
    """Belge + dışarı çıkan ok (XML aktarma)."""
    w = _kalinlik(n)
    x0, x1, y0, y1 = n * .10, n * .58, n * .08, n * .92
    kivrim = n * .18
    d.line([(x0, y0), (x1 - kivrim, y0), (x1, y0 + kivrim), (x1, y1), (x0, y1), (x0, y0)],
           fill=c, width=w, joint="curve")
    d.line([(x1 - kivrim, y0), (x1 - kivrim, y0 + kivrim), (x1, y0 + kivrim)], fill=c, width=w)
    for yy in (.42, .56, .70):
        d.line([(x0 + n * .10, n * yy), (x1 - n * .10, n * yy)], fill=c, width=w)
    d.line([(n * .66, n * .50), (n * .94, n * .50)], fill=c, width=w)
    d.line([(n * .82, n * .38), (n * .94, n * .50), (n * .82, n * .62)],
           fill=c, width=w, joint="curve")


def _karsilastir(d, n, c):
    """İki liste + ortada eşitlik işareti (fatura kontrol)."""
    w = _kalinlik(n)
    d.rounded_rectangle([n * .04, n * .14, n * .40, n * .86], n * .06, outline=c, width=w)
    d.rounded_rectangle([n * .60, n * .14, n * .96, n * .86], n * .06, outline=c, width=w)
    for yy in (.30, .44, .58, .72):
        d.line([(n * .12, n * yy), (n * .32, n * yy)], fill=c, width=w)
        d.line([(n * .68, n * yy), (n * .88, n * yy)], fill=c, width=w)
    d.line([(n * .44, n * .42), (n * .56, n * .42)], fill=c, width=w)
    d.line([(n * .44, n * .58), (n * .56, n * .58)], fill=c, width=w)


def _onay(d, n, c):
    w = max(2, int(n / 8))
    d.line([(n * .18, n * .52), (n * .42, n * .76), (n * .84, n * .24)],
           fill=c, width=w, joint="curve")


def _ok_sag(d, n, c):
    w = max(2, int(n / 9))
    d.line([(n * .12, n * .50), (n * .84, n * .50)], fill=c, width=w)
    d.line([(n * .60, n * .24), (n * .86, n * .50), (n * .60, n * .76)],
           fill=c, width=w, joint="curve")


def _ok_sol(d, n, c):
    w = max(2, int(n / 9))
    d.line([(n * .16, n * .50), (n * .88, n * .50)], fill=c, width=w)
    d.line([(n * .40, n * .24), (n * .14, n * .50), (n * .40, n * .76)],
           fill=c, width=w, joint="curve")


def _chevron_asagi(d, n, c):
    w = max(2, int(n / 8))
    d.line([(n * .22, n * .38), (n * .50, n * .66), (n * .78, n * .38)],
           fill=c, width=w, joint="curve")


def _portal(d, n, c):
    """Sunucu/portal katmanları - e-fatura sağlayıcıları."""
    w = _kalinlik(n)
    for y in (.10, .40, .70):
        d.rounded_rectangle([n * .08, n * y, n * .92, n * (y + .20)], n * .05,
                            outline=c, width=w)
        d.ellipse([n * .18, n * (y + .075), n * .26, n * (y + .125)], fill=c)


def _magaza(d, n, c):
    """Alışveriş çantası - pazaryeri sağlayıcıları."""
    w = _kalinlik(n)
    d.rounded_rectangle([n * .12, n * .32, n * .88, n * .92], n * .10, outline=c, width=w)
    d.arc([n * .30, n * .10, n * .70, n * .54], 180, 360, fill=c, width=w)


def _makbuz(d, n, c):
    """Zikzak kenarlı makbuz - müstahsil belgeleri."""
    w = _kalinlik(n)
    ust, alt = n * .08, n * .84
    sol, sag = n * .16, n * .84
    nokta = [(sol, alt), (sol, ust), (sag, ust), (sag, alt)]
    adim = (sag - sol) / 6
    for i in range(6):
        x = sag - adim * i
        nokta.append((x - adim / 2, n * .94))
        nokta.append((x - adim, alt))
    d.line(nokta + [nokta[0]], fill=c, width=w, joint="curve")
    for yy in (.32, .48):
        d.line([(sol + n * .10, n * yy), (sag - n * .10, n * yy)], fill=c, width=w)
    d.line([(sol + n * .10, n * .64), (sol + n * .36, n * .64)], fill=c, width=w)


def _klasor(d, n, c):
    w = _kalinlik(n)
    d.line([(n * .08, n * .82), (n * .08, n * .22), (n * .40, n * .22), (n * .50, n * .36),
            (n * .92, n * .36), (n * .92, n * .82), (n * .08, n * .82)],
           fill=c, width=w, joint="curve")


def _tablo(d, n, c):
    """Excel/tablo - dosya seçim kutuları."""
    w = _kalinlik(n)
    d.rounded_rectangle([n * .10, n * .14, n * .90, n * .86], n * .08, outline=c, width=w)
    d.line([(n * .10, n * .38), (n * .90, n * .38)], fill=c, width=w)
    d.line([(n * .44, n * .38), (n * .44, n * .86)], fill=c, width=w)


def _uyari(d, n, c):
    w = _kalinlik(n, 14)
    d.line([(n * .50, n * .10), (n * .94, n * .86), (n * .06, n * .86), (n * .50, n * .10)],
           fill=c, width=w, joint="curve")
    d.line([(n * .50, n * .40), (n * .50, n * .62)], fill=c, width=w)
    d.ellipse([n * .455, n * .70, n * .545, n * .79], fill=c)


def _marka(d, n, c):
    """Uygulama marka işareti - dolu yuvarlak kare içinde fatura satırları."""
    d.rounded_rectangle([0, 0, n - 1, n - 1], n * .27, fill=c)
    w = max(2, int(n / 13))
    d.line([(n * .28, n * .30), (n * .60, n * .30)], fill="#0d1014", width=w)
    d.line([(n * .28, n * .48), (n * .72, n * .48)], fill="#0d1014", width=w)
    d.line([(n * .28, n * .66), (n * .48, n * .66)], fill="#0d1014", width=w)
    d.line([(n * .56, n * .72), (n * .64, n * .80), (n * .80, n * .56)],
           fill="#bfe8d7", width=w, joint="curve")


# ----------------- Genel API -----------------

def belge(size, color):
    return _ikon("belge", _belge, size, color)


def karsilastir(size, color):
    return _ikon("karsilastir", _karsilastir, size, color)


def onay(size, color):
    return _ikon("onay", _onay, size, color)


def ok_sag(size, color):
    return _ikon("ok_sag", _ok_sag, size, color)


def ok_sol(size, color):
    return _ikon("ok_sol", _ok_sol, size, color)


def chevron_asagi(size, color):
    return _ikon("chevron_asagi", _chevron_asagi, size, color)


def portal(size, color):
    return _ikon("portal", _portal, size, color)


def magaza(size, color):
    return _ikon("magaza", _magaza, size, color)


def makbuz(size, color):
    return _ikon("makbuz", _makbuz, size, color)


def klasor(size, color):
    return _ikon("klasor", _klasor, size, color)


def tablo(size, color):
    return _ikon("tablo", _tablo, size, color)


def uyari(size, color):
    return _ikon("uyari", _uyari, size, color)


def marka(size, color):
    return _ikon("marka", _marka, size, color)


# Sağlayıcı anahtarı -> kategori glifi. Aynı aileden üç glif; kimliği
# renk ayırıyor (bkz. theme.PROVIDER_COLORS).
SAGLAYICI_GLIF = {
    "logo":         portal,
    "uyumsoft":     portal,
    "hizlibilisim": portal,
    "izibiz":       portal,
    "vega":         portal,
    "mustahsil":    makbuz,
    "hepsiburada":  magaza,
    "trendyol":     magaza,
}


def saglayici(anahtar, size, color):
    """Sağlayıcı anahtarına göre kategori ikonu döndürür."""
    return SAGLAYICI_GLIF.get(anahtar, portal)(size, color)
