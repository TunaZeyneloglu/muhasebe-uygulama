import customtkinter as ctk

import icons
import theme
from pages import ui
from pages.kontrol_popups import kontrol_hb_ve_trn, kontrol_izibiz, kontrol_vega, show_hizlibilisim_menu, show_logo_kontrol_popup, show_uyumsoft_menu
from pages.navigation import show_welcome

def build_kontrol_page(app):
    # ================= KONTROL SAYFASI =================
    # Karşılama ekranıyla aynı görsel dil: araç çubuğu, sayfa başlığı ve
    # ikonlu sağlayıcı satırları.

    kontrol_frame = ctk.CTkFrame(app, fg_color=theme.BG_ROOT, corner_radius=0)

    ui.build_toolbar(kontrol_frame)
    ui.build_statusbar(kontrol_frame, "Fatura Kontrol  ·  bir portal seçin")

    wrap = ctk.CTkFrame(kontrol_frame, fg_color="transparent")
    wrap.pack(fill="both", expand=True, padx=theme.PAD_STAGE, pady=(24, 26))

    # Geri Butonu
    ui.build_back_button(wrap, show_welcome).pack(anchor="w", pady=(0, 20))

    # Başlık
    ui.build_page_header(
        wrap, "Fatura Kontrol",
        "Zirve ve portal faturalarını karşılaştırarak eksik veya hatalı girişleri tespit edin.",
        theme.SUCCESS, icons.karsilastir)

    # Sağlayıcı satırları (2 sütun x 3 satır)
    rows_frame = ctk.CTkFrame(wrap, fg_color="transparent")
    rows_frame.pack(fill="both", expand=True, pady=(24, 0))
    rows_frame.grid_columnconfigure((0, 1), weight=1, uniform="kontrol_col")
    rows_frame.grid_rowconfigure((0, 1, 2), weight=1, uniform="kontrol_col_row")

    def yerlestir(kart, satir, sutun):
        kart.grid(row=satir, column=sutun, sticky="nsew",
                  padx=(0, 6) if sutun == 0 else (6, 0),
                  pady=theme.PAD_ROW)

    # Satır 1: Logo ve Uyumsoft
    logo_kontrol_col, logo_kontrol_btn = ui.build_provider_row(
        rows_frame, "logo", "Logo", "Logo ve Hızlıbilişim listeleri birlikte", "Kontrol")
    yerlestir(logo_kontrol_col, 0, 0)
    logo_kontrol_btn.configure(command=show_logo_kontrol_popup)

    uyumsoft_kontrol_col, uyumsoft_kontrol_btn = ui.build_provider_row(
        rows_frame, "uyumsoft", "Uyumsoft", "Zirve / Verim kaynağını seçin", "Kontrol", menulu=True)
    yerlestir(uyumsoft_kontrol_col, 0, 1)
    uyumsoft_kontrol_btn.configure(command=lambda: show_uyumsoft_menu(uyumsoft_kontrol_btn))

    # Satır 2: Hızlıbilişim ve İzibiz
    hizli_kontrol_col, hizli_kontrol_btn = ui.build_provider_row(
        rows_frame, "hizlibilisim", "Hızlıbilişim", "Alış / satış faturasını seçin", "Kontrol", menulu=True)
    yerlestir(hizli_kontrol_col, 1, 0)
    hizli_kontrol_btn.configure(command=lambda: show_hizlibilisim_menu(hizli_kontrol_btn))

    izibiz_kontrol_col, izibiz_kontrol_btn = ui.build_provider_row(
        rows_frame, "izibiz", "İzibiz", "Zirve listesiyle karşılaştırır", "Kontrol")
    yerlestir(izibiz_kontrol_col, 1, 1)
    izibiz_kontrol_btn.configure(command=kontrol_izibiz)

    # Satır 3: Vega ve Hepsiburada/Trendyol
    vega_kontrol_col, vega_kontrol_btn = ui.build_provider_row(
        rows_frame, "vega", "Vega", "Zirve listesiyle karşılaştırır", "Kontrol")
    yerlestir(vega_kontrol_col, 2, 0)
    vega_kontrol_btn.configure(command=kontrol_vega)

    hb_ve_trn_kontrol_col, hb_ve_trn_kontrol_btn = ui.build_provider_row(
        rows_frame, "hepsiburada", "Hepsiburada / Trendyol", "Portal otomatik tespit edilir", "Kontrol")
    yerlestir(hb_ve_trn_kontrol_col, 2, 1)
    hb_ve_trn_kontrol_btn.configure(command=kontrol_hb_ve_trn)

    # Bilgi şeridi + sonuç mesajı alanı
    bilgi_paneli = ctk.CTkFrame(wrap, fg_color=theme.BG_SURFACE, corner_radius=theme.CORNER_TILE,
                                border_width=theme.BORDER_WIDTH, border_color=theme.BORDER)
    bilgi_paneli.pack(fill="x", pady=(22, 0))

    bilgi_ic = ctk.CTkFrame(bilgi_paneli, fg_color="transparent")
    bilgi_ic.pack(fill="x", padx=theme.PAD_TILE, pady=12)

    ctk.CTkLabel(bilgi_ic, text="", image=icons.uyari(theme.ICON_SM, theme.TEXT_MUTED)).pack(side="left")
    ctk.CTkLabel(bilgi_ic,
                 text="Karşılaştırma sonucu renkli bir Excel dosyası olarak masaüstüne kaydedilir.",
                 font=theme.FONT_SMALL(), text_color=theme.TEXT_MUTED,
                 justify="left").pack(side="left", padx=10)

    # Sonuç mesajı
    kontrol_label_sonuc = ctk.CTkLabel(bilgi_ic, text="", font=theme.FONT_SMALL(),
                                       text_color=theme.TEXT_SECONDARY, anchor="e")
    kontrol_label_sonuc.pack(side="right")

    return kontrol_frame
