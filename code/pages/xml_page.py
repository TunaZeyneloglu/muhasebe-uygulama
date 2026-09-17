import customtkinter as ctk

import theme
from logic.common import dosyayi_ac
from logic.xml_logo import klasor_sec_ve_isle_logo
from pages.navigation import show_welcome
from pages.popup_menu import show_hepsiburada_menu, show_hizlibilisim_xml_menu, show_mustahsil_menu, show_trendyol_menu, show_uyumsoft_xml_menu

def build_xml_page(app):
    # ================= XML SAYFASI =================

    xml_frame = ctk.CTkFrame(app, fg_color="transparent")

    xml_inner = ctk.CTkFrame(xml_frame, fg_color=theme.BG_SURFACE, corner_radius=theme.CORNER_CARD)
    xml_inner.pack(expand=True, fill="both")

    # Geri Butonu
    xml_back_frame = ctk.CTkFrame(xml_inner, fg_color="transparent")
    xml_back_frame.pack(fill="x", padx=15, pady=(15, 0))

    xml_back_btn = ctk.CTkButton(
        xml_back_frame, 
        text="← Geri", 
        width=theme.BTN_W_BACK, 
        height=theme.BTN_H_BACK,
        fg_color="transparent",
        border_width=1,
        border_color=theme.BORDER,
        hover_color=theme.BG_HOVER,
        command=show_welcome
    )
    xml_back_btn.pack(side="left")

    # Başlık
    xml_title = ctk.CTkLabel(xml_inner, text="📄 XML Aktarma", font=theme.FONT_PAGE_TITLE())
    xml_title.pack(pady=(10, 5))

    xml_aciklama = ctk.CTkLabel(
        xml_inner,
        text="XML formatındaki faturaları seçerek Excel çıktısı oluşturabilirsiniz.",
        font=theme.FONT_BODY(),
        text_color=theme.TEXT_SECONDARY,
        justify="center"
    )
    xml_aciklama.pack(pady=(0, 20))

    # Satır 1: Logo ve Uyumsoft Butonları
    xml_row1 = ctk.CTkFrame(xml_inner, fg_color="transparent")
    xml_row1.pack(pady=theme.PAD_ROW)

    # Logo
    logo_col = ctk.CTkFrame(xml_row1, fg_color="transparent")
    logo_col.grid(row=0, column=0, padx=40)
    ctk.CTkLabel(logo_col, text="Logo", font=theme.FONT_H2()).pack(pady=(0, 10))
    ctk.CTkButton(logo_col, text="📂 Aktar", width=theme.BTN_W_LG, height=theme.BTN_H_LG,
                  fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                  command=lambda: klasor_sec_ve_isle_logo(xml_label_sonuc, xml_buton_ac)).pack()

    # Uyumsoft
    uyumsoft_col = ctk.CTkFrame(xml_row1, fg_color="transparent")
    uyumsoft_col.grid(row=0, column=1, padx=40)
    ctk.CTkLabel(uyumsoft_col, text="Uyumsoft", font=theme.FONT_H2()).pack(pady=(0, 10))
    uyumsoft_xml_buton = ctk.CTkButton(uyumsoft_col, text="🧾 Aktar", width=theme.BTN_W_LG, height=theme.BTN_H_LG,
                                       fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER)
    uyumsoft_xml_buton.pack()

    # Satır 2: Müstahsil ve Hızlıbilişim
    xml_row2 = ctk.CTkFrame(xml_inner, fg_color="transparent")
    xml_row2.pack(pady=theme.PAD_ROW)

    # Müstahsil
    mustahsil_col = ctk.CTkFrame(xml_row2, fg_color="transparent")
    mustahsil_col.grid(row=0, column=0, padx=40)
    ctk.CTkLabel(mustahsil_col, text="Müstahsil", font=theme.FONT_H2()).pack(pady=(0, 10))
    mustahsil_buton = ctk.CTkButton(mustahsil_col, text="🥔 Aktar", width=theme.BTN_W_LG, height=theme.BTN_H_LG,
                                    fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER)
    mustahsil_buton.pack()

    # Hızlıbilişim
    hizlibilisim_xml_col = ctk.CTkFrame(xml_row2, fg_color="transparent")
    hizlibilisim_xml_col.grid(row=0, column=1, padx=40)
    ctk.CTkLabel(hizlibilisim_xml_col, text="Hızlıbilişim", font=theme.FONT_H2()).pack(pady=(0, 10))
    hizlibilisim_xml_buton = ctk.CTkButton(hizlibilisim_xml_col, text="⚡ Aktar", width=theme.BTN_W_LG, height=theme.BTN_H_LG,
                                           fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER)
    hizlibilisim_xml_buton.pack()

    # Satır 3: Hepsiburada ve Trendyol
    xml_row3 = ctk.CTkFrame(xml_inner, fg_color="transparent")
    xml_row3.pack(pady=theme.PAD_ROW)

    # Hepsiburada
    hb_col = ctk.CTkFrame(xml_row3, fg_color="transparent")
    hb_col.grid(row=0, column=0, padx=40)
    ctk.CTkLabel(hb_col, text="Hepsiburada", font=theme.FONT_H2()).pack(pady=(0, 10))
    hb_buton = ctk.CTkButton(hb_col, text="🛒 Aktar", width=theme.BTN_W_LG, height=theme.BTN_H_LG,
                             fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER)
    hb_buton.pack()

    # Trendyol
    ty_col = ctk.CTkFrame(xml_row3, fg_color="transparent")
    ty_col.grid(row=0, column=1, padx=40)
    ctk.CTkLabel(ty_col, text="Trendyol", font=theme.FONT_H2()).pack(pady=(0, 10))
    ty_buton = ctk.CTkButton(ty_col, text="🛍️ Aktar", width=theme.BTN_W_LG, height=theme.BTN_H_LG,
                             fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER)
    ty_buton.pack()

    # Sonuç mesajı
    xml_label_sonuc = ctk.CTkLabel(xml_inner, text="", font=theme.FONT_SUBTITLE())
    xml_label_sonuc.pack(pady=10)

    # Excel aç butonu
    xml_buton_ac = ctk.CTkButton(xml_inner, text="📂 Oluşturulan Excel Dosyasını Aç", command=dosyayi_ac, state="disabled", width=theme.BTN_W_WIDE, height=theme.BTN_H_WIDE,
                                 fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER)
    xml_buton_ac.pack(pady=(10, 30))

    # Popup menü komutlarını bağla (label_sonuc ve buton_ac tanımlandıktan sonra)
    uyumsoft_xml_buton.configure(command=lambda: show_uyumsoft_xml_menu(uyumsoft_xml_buton, xml_label_sonuc, xml_buton_ac))
    mustahsil_buton.configure(command=lambda: show_mustahsil_menu(mustahsil_buton, xml_label_sonuc, xml_buton_ac))
    hb_buton.configure(command=lambda: show_hepsiburada_menu(hb_buton, xml_label_sonuc, xml_buton_ac))
    ty_buton.configure(command=lambda: show_trendyol_menu(ty_buton, xml_label_sonuc, xml_buton_ac))
    hizlibilisim_xml_buton.configure(command=lambda: show_hizlibilisim_xml_menu(hizlibilisim_xml_buton, xml_label_sonuc, xml_buton_ac))

    return xml_frame
