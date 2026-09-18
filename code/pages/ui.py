"""Sayfalar arasında paylaşılan görsel bileşenler.

Karşılama, XML ve Kontrol sayfaları aynı görsel dili kullansın diye üst
araç çubuğu, sayfa başlığı, sağlayıcı satırı gibi tekrar eden parçalar
burada tek yerde tanımlıdır. Bu modül yalnızca görünüm üretir; hiçbir iş
mantığı içermez.
"""
import customtkinter as ctk

import icons
import theme


def _surum_metni():
    """version.py'deki sürümü rozet için döndürür; okunamazsa rozet gizlenir."""
    try:
        import version
        return "v" + version.APP_VERSION
    except Exception:
        return ""


def build_toolbar(parent, alt_baslik="Mali Müşavirlik"):
    """Sayfanın en üstündeki 60 piksellik uygulama çubuğu."""
    bar = ctk.CTkFrame(parent, height=theme.TOOLBAR_H, corner_radius=0,
                       fg_color=theme.BG_SURFACE)
    bar.pack(fill="x", side="top")
    bar.pack_propagate(False)
    ctk.CTkFrame(bar, height=theme.DIVIDER_H, fg_color=theme.BORDER).pack(
        side="bottom", fill="x")

    sol = ctk.CTkFrame(bar, fg_color="transparent")
    sol.pack(side="left", padx=(theme.PAD_STAGE - 20, 0), pady=12)
    ctk.CTkLabel(sol, text="", image=icons.marka(28, theme.ACCENT)).pack(side="left")
    ctk.CTkLabel(sol, text="  Fatura Yönetim Sistemi", font=theme.FONT_BRAND(),
                 text_color=theme.TEXT_PRIMARY).pack(side="left")
    ctk.CTkLabel(sol, text="   ·   " + alt_baslik, font=theme.FONT_SMALL(),
                 text_color=theme.TEXT_MUTED).pack(side="left")

    sag = ctk.CTkFrame(bar, fg_color="transparent")
    sag.pack(side="right", padx=(0, theme.PAD_STAGE - 20), pady=12)
    surum = _surum_metni()
    if surum:
        ctk.CTkLabel(sag, text=" " + surum + " ", font=theme.FONT_SMALL(), height=24,
                     corner_radius=theme.CORNER_CHIP, fg_color=theme.BG_ELEVATED,
                     text_color=theme.TEXT_SECONDARY).pack(side="left")
    return bar


def build_statusbar(parent, sol_metin, sag_metin="çevrimdışı çalışır"):
    """Sayfanın en altındaki 32 piksellik durum çubuğu."""
    bar = ctk.CTkFrame(parent, height=theme.STATUSBAR_H, corner_radius=0,
                       fg_color=theme.BG_SURFACE)
    bar.pack(fill="x", side="bottom")
    bar.pack_propagate(False)
    ctk.CTkFrame(bar, height=theme.DIVIDER_H, fg_color=theme.BORDER).pack(
        side="top", fill="x")
    etiket = ctk.CTkLabel(bar, text=sol_metin, font=theme.FONT_SMALL(),
                          text_color=theme.TEXT_MUTED)
    etiket.pack(side="left", padx=theme.PAD_STAGE)
    ctk.CTkLabel(bar, text=sag_metin, font=theme.FONT_SMALL(),
                 text_color=theme.TEXT_MUTED).pack(side="right", padx=theme.PAD_STAGE)
    return etiket


def build_back_button(parent, command):
    """Sayfa gövdesinin üstündeki ikincil '← Geri' butonu."""
    return ctk.CTkButton(
        parent, text="Geri", image=icons.ok_sol(theme.ICON_SM, theme.TEXT_SECONDARY),
        compound="left", width=theme.BTN_W_BACK, height=theme.BTN_H_BACK,
        corner_radius=theme.CORNER_BTN, font=theme.FONT_SMALL(),
        fg_color="transparent", hover_color=theme.BG_HOVER,
        text_color=theme.TEXT_SECONDARY,
        border_width=theme.BORDER_WIDTH, border_color=theme.BORDER,
        command=command)


def build_page_header(parent, baslik, aciklama, renk, glif):
    """İkon kutusu + başlık + açıklamadan oluşan sayfa başlığı bloğu."""
    head = ctk.CTkFrame(parent, fg_color="transparent")
    head.pack(anchor="w", fill="x")
    ctk.CTkLabel(head, text="", image=glif(theme.ICON_LG, renk),
                 width=theme.HEAD_TILE, height=theme.HEAD_TILE,
                 corner_radius=theme.CORNER_TILE,
                 fg_color=theme.TINT(renk)).pack(side="left")
    metin = ctk.CTkFrame(head, fg_color="transparent")
    metin.pack(side="left", padx=16)
    ctk.CTkLabel(metin, text=baslik, font=theme.FONT_H1(),
                 text_color=theme.TEXT_PRIMARY).pack(anchor="w")
    ctk.CTkLabel(metin, text=aciklama, font=theme.FONT_SMALL(),
                 text_color=theme.TEXT_MUTED, justify="left").pack(anchor="w", pady=(3, 0))
    return head


def build_provider_row(parent, anahtar, baslik, aciklama, buton_metni, menulu=False):
    """Sağlayıcı satırı: ikon kutusu + ad/açıklama + işlem butonu.

    Buton `command` olmadan üretilir; çağıran taraf popup konumlandırması
    için aynı buton değişkeni üzerinden `.configure(command=...)` yapar.
    (card, buton) döndürür.
    """
    renk = theme.PROVIDER_COLORS.get(anahtar, theme.ACCENT)

    card = ctk.CTkFrame(parent, fg_color=theme.BG_SURFACE, corner_radius=theme.CORNER_TILE,
                        border_width=theme.BORDER_WIDTH, border_color=theme.BORDER)

    # fill/expand: satır kartı dikeyde büyüdüğünde içerik ortalanır
    ic = ctk.CTkFrame(card, fg_color="transparent")
    ic.pack(fill="both", expand=True, padx=theme.PAD_TILE, pady=theme.PAD_TILE)

    tile = ctk.CTkLabel(ic, text="", image=icons.saglayici(anahtar, theme.ICON_MD, renk),
                        width=theme.TILE_SIZE, height=theme.TILE_SIZE,
                        corner_radius=theme.CORNER_TILE, fg_color=theme.TINT(renk))
    tile.pack(side="left")

    buton = ctk.CTkButton(
        ic, text=buton_metni,
        image=(icons.chevron_asagi(theme.ICON_SM, theme.TEXT_ON_ACCENT) if menulu
               else icons.ok_sag(theme.ICON_SM, theme.TEXT_ON_ACCENT)),
        compound="right", width=theme.BTN_W_ROW, height=theme.BTN_H_ROW,
        corner_radius=theme.CORNER_BTN, font=theme.FONT_BODY_BOLD(),
        fg_color=renk, hover_color=theme.karistir(renk, 0.78, "#ffffff"),
        text_color=theme.TEXT_ON_ACCENT)
    buton.pack(side="right")

    metin = ctk.CTkFrame(ic, fg_color="transparent")
    metin.pack(side="left", padx=(14, 10), fill="x", expand=True)
    ad = ctk.CTkLabel(metin, text=baslik, font=theme.FONT_H2(),
                      text_color=theme.TEXT_PRIMARY)
    ad.pack(anchor="w")
    alt = ctk.CTkLabel(metin, text=aciklama, font=theme.FONT_SMALL(),
                       text_color=theme.TEXT_MUTED, justify="left")
    alt.pack(anchor="w", pady=(2, 0))

    # Hover: kartın çerçevesi sağlayıcı rengine döner (tıklama butondadır)
    def gir(_=None):
        card.configure(border_color=renk, fg_color=theme.BG_ELEVATED)

    def cik(_=None):
        card.configure(border_color=theme.BORDER, fg_color=theme.BG_SURFACE)

    for w in (card, ic, tile, metin, ad, alt):
        w.bind("<Enter>", gir)
        w.bind("<Leave>", cik)
    buton.bind("<Enter>", gir)
    buton.bind("<Leave>", cik)

    return card, buton
