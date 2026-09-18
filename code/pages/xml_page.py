import customtkinter as ctk

import icons
import theme
from logic.common import dosyayi_ac
from logic.xml_logo import klasor_sec_ve_isle_logo
from pages import ui
from pages.navigation import show_welcome
from pages.popup_menu import show_hepsiburada_menu, show_hizlibilisim_xml_menu, show_mustahsil_menu, show_trendyol_menu, show_uyumsoft_xml_menu, xml_isini_baslat

def build_xml_page(app):
    # ================= XML SAYFASI =================
    # Karşılama ekranıyla aynı görsel dil: araç çubuğu, sayfa başlığı ve
    # ikonlu sağlayıcı satırları.

    xml_frame = ctk.CTkFrame(app, fg_color=theme.BG_ROOT, corner_radius=0)

    ui.build_toolbar(xml_frame)
    ui.build_statusbar(xml_frame, "XML Aktarma  ·  bir kaynak seçin")

    wrap = ctk.CTkFrame(xml_frame, fg_color="transparent")
    wrap.pack(fill="both", expand=True, padx=theme.PAD_STAGE, pady=(24, 26))

    # Geri Butonu
    ui.build_back_button(wrap, show_welcome).pack(anchor="w", pady=(0, 20))

    # Başlık
    ui.build_page_header(
        wrap, "XML Aktarma",
        "XML formatındaki faturaları seçerek Excel çıktısı oluşturabilirsiniz.",
        theme.ACCENT, icons.belge)

    # Sağlayıcı satırları (2 sütun x 3 satır)
    rows_frame = ctk.CTkFrame(wrap, fg_color="transparent")
    rows_frame.pack(fill="both", expand=True, pady=(24, 0))
    rows_frame.grid_columnconfigure((0, 1), weight=1, uniform="xml_col")
    rows_frame.grid_rowconfigure((0, 1, 2), weight=1, uniform="xml_col_row")

    def yerlestir(kart, satir, sutun):
        kart.grid(row=satir, column=sutun, sticky="nsew",
                  padx=(0, 6) if sutun == 0 else (6, 0),
                  pady=theme.PAD_ROW)

    # Satır 1: Logo ve Uyumsoft
    logo_col, logo_buton = ui.build_provider_row(
        rows_frame, "logo", "Logo", "Klasördeki XML faturalarını aktarır", "Aktar")
    yerlestir(logo_col, 0, 0)

    uyumsoft_col, uyumsoft_xml_buton = ui.build_provider_row(
        rows_frame, "uyumsoft", "Uyumsoft", "ACE / Turkay kaynağını seçin", "Aktar", menulu=True)
    yerlestir(uyumsoft_col, 0, 1)

    # Satır 2: Müstahsil ve Hızlıbilişim
    mustahsil_col, mustahsil_buton = ui.build_provider_row(
        rows_frame, "mustahsil", "Müstahsil", "Borsa, alış, makbuz ve muhtasar", "Aktar", menulu=True)
    yerlestir(mustahsil_col, 1, 0)

    hizlibilisim_xml_col, hizlibilisim_xml_buton = ui.build_provider_row(
        rows_frame, "hizlibilisim", "Hızlıbilişim", "Alış / satış faturasını seçin", "Aktar", menulu=True)
    yerlestir(hizlibilisim_xml_col, 1, 1)

    # Satır 3: Hepsiburada ve Trendyol
    hb_col, hb_buton = ui.build_provider_row(
        rows_frame, "hepsiburada", "Hepsiburada", "Alış / satış faturasını seçin", "Aktar", menulu=True)
    yerlestir(hb_col, 2, 0)

    ty_col, ty_buton = ui.build_provider_row(
        rows_frame, "trendyol", "Trendyol", "Alış / satış faturasını seçin", "Aktar", menulu=True)
    yerlestir(ty_col, 2, 1)

    # Sonuç şeridi: durum mesajı solda, Excel'i açma butonu sağda
    sonuc_paneli = ctk.CTkFrame(wrap, fg_color=theme.BG_SURFACE, corner_radius=theme.CORNER_TILE,
                                border_width=theme.BORDER_WIDTH, border_color=theme.BORDER)
    sonuc_paneli.pack(fill="x", pady=(22, 0))

    sonuc_ic = ctk.CTkFrame(sonuc_paneli, fg_color="transparent")
    sonuc_ic.pack(fill="x", padx=theme.PAD_TILE, pady=theme.PAD_TILE)

    # Excel aç butonu
    xml_buton_ac = ctk.CTkButton(sonuc_ic, text="Oluşturulan Excel Dosyasını Aç",
                                 image=icons.tablo(theme.ICON_SM, theme.TEXT_ON_ACCENT),
                                 compound="left", command=dosyayi_ac, state="disabled",
                                 width=theme.BTN_W_WIDE, height=theme.BTN_H_WIDE,
                                 corner_radius=theme.CORNER_BTN, font=theme.FONT_BODY_BOLD(),
                                 fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                                 text_color=theme.TEXT_ON_ACCENT)
    xml_buton_ac.pack(side="right")

    # Sonuç mesajı
    xml_label_sonuc = ctk.CTkLabel(sonuc_ic, text="", font=theme.FONT_BODY(),
                                   text_color=theme.TEXT_SECONDARY, justify="left",
                                   wraplength=540, anchor="w")
    xml_label_sonuc.pack(side="left", fill="x", expand=True, padx=(2, 14))

    # Popup menü komutlarını bağla (label_sonuc ve buton_ac tanımlandıktan sonra)
    logo_buton.configure(command=lambda: xml_isini_baslat(logo_buton, klasor_sec_ve_isle_logo, xml_label_sonuc, xml_buton_ac))
    uyumsoft_xml_buton.configure(command=lambda: show_uyumsoft_xml_menu(uyumsoft_xml_buton, xml_label_sonuc, xml_buton_ac))
    mustahsil_buton.configure(command=lambda: show_mustahsil_menu(mustahsil_buton, xml_label_sonuc, xml_buton_ac))
    hb_buton.configure(command=lambda: show_hepsiburada_menu(hb_buton, xml_label_sonuc, xml_buton_ac))
    ty_buton.configure(command=lambda: show_trendyol_menu(ty_buton, xml_label_sonuc, xml_buton_ac))
    hizlibilisim_xml_buton.configure(command=lambda: show_hizlibilisim_xml_menu(hizlibilisim_xml_buton, xml_label_sonuc, xml_buton_ac))

    return xml_frame
