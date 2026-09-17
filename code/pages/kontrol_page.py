import customtkinter as ctk

import theme
from pages.kontrol_popups import kontrol_hb_ve_trn, kontrol_izibiz, kontrol_vega, show_hizlibilisim_menu, show_logo_kontrol_popup, show_uyumsoft_menu
from pages.navigation import show_welcome

def build_kontrol_page(app):
    # ================= KONTROL SAYFASI =================

    kontrol_frame = ctk.CTkFrame(app, fg_color="transparent")

    kontrol_inner = ctk.CTkFrame(kontrol_frame, fg_color=theme.BG_SURFACE, corner_radius=theme.CORNER_CARD)
    kontrol_inner.pack(expand=False, fill="x", padx=theme.PAD_PAGE, pady=30)

    # Geri Butonu
    kontrol_back_frame = ctk.CTkFrame(kontrol_inner, fg_color="transparent")
    kontrol_back_frame.pack(fill="x", padx=15, pady=(15, 10))

    kontrol_back_btn = ctk.CTkButton(
        kontrol_back_frame, 
        text="← Geri", 
        width=theme.BTN_W_BACK, 
        height=theme.BTN_H_BACK,
        fg_color="transparent",
        border_width=1,
        border_color=theme.BORDER,
        hover_color=theme.BG_HOVER,
        command=show_welcome
    )
    kontrol_back_btn.pack(side="left")

    # Başlık
    kontrol_title = ctk.CTkLabel(kontrol_inner, text="🔍 Fatura Kontrol", font=theme.FONT_PAGE_TITLE())
    kontrol_title.pack(pady=(5, 2))

    kontrol_aciklama = ctk.CTkLabel(
        kontrol_inner,
        text="Zirve ve portal faturalarını karşılaştırarak\neksik veya hatalı girişleri tespit edin.",
        font=theme.FONT_BODY(),
        text_color=theme.TEXT_SECONDARY,
        justify="center"
    )
    kontrol_aciklama.pack(pady=(2, 25))

    # Satır 1: Logo ve Uyumsoft
    kontrol_row1 = ctk.CTkFrame(kontrol_inner, fg_color="transparent")
    kontrol_row1.pack(pady=theme.PAD_ROW, expand=False)

    logo_kontrol_col = ctk.CTkFrame(kontrol_row1, fg_color="transparent")
    logo_kontrol_col.grid(row=0, column=0, padx=40)
    ctk.CTkLabel(logo_kontrol_col, text="Logo", font=theme.FONT_H2()).pack(pady=(0, 10))
    ctk.CTkButton(logo_kontrol_col, text="🔍 Kontrol", width=theme.BTN_W_LG, height=theme.BTN_H_LG,
                  fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                  command=show_logo_kontrol_popup).pack()

    uyumsoft_kontrol_col = ctk.CTkFrame(kontrol_row1, fg_color="transparent")
    uyumsoft_kontrol_col.grid(row=0, column=1, padx=40)
    ctk.CTkLabel(uyumsoft_kontrol_col, text="Uyumsoft", font=theme.FONT_H2()).pack(pady=(0, 10))
    uyumsoft_kontrol_btn = ctk.CTkButton(uyumsoft_kontrol_col, text="🔍 Kontrol", width=theme.BTN_W_LG, height=theme.BTN_H_LG,
                                         fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER)
    uyumsoft_kontrol_btn.pack()
    uyumsoft_kontrol_btn.configure(command=lambda: show_uyumsoft_menu(uyumsoft_kontrol_btn))

    # Satır 2: Hızlıbilişim ve İzibiz
    kontrol_row2 = ctk.CTkFrame(kontrol_inner, fg_color="transparent")
    kontrol_row2.pack(pady=theme.PAD_ROW, expand=False)

    hizli_kontrol_col = ctk.CTkFrame(kontrol_row2, fg_color="transparent")
    hizli_kontrol_col.grid(row=0, column=0, padx=40)
    ctk.CTkLabel(hizli_kontrol_col, text="Hızlıbilişim", font=theme.FONT_H2()).pack(pady=(0, 10))
    hizli_kontrol_btn = ctk.CTkButton(hizli_kontrol_col, text="🔍 Kontrol", width=theme.BTN_W_LG, height=theme.BTN_H_LG,
                                      fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER)
    hizli_kontrol_btn.pack()
    hizli_kontrol_btn.configure(command=lambda: show_hizlibilisim_menu(hizli_kontrol_btn))

    izibiz_kontrol_col = ctk.CTkFrame(kontrol_row2, fg_color="transparent")
    izibiz_kontrol_col.grid(row=0, column=1, padx=40)
    ctk.CTkLabel(izibiz_kontrol_col, text="İzibiz", font=theme.FONT_H2()).pack(pady=(0, 10))
    ctk.CTkButton(izibiz_kontrol_col, text="🔍 Kontrol", width=theme.BTN_W_LG, height=theme.BTN_H_LG,
                  fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                  command=kontrol_izibiz).pack()

    # Satır 3: Vega ve Hepsiburada/Trendyol
    kontrol_row3 = ctk.CTkFrame(kontrol_inner, fg_color="transparent")
    kontrol_row3.pack(pady=theme.PAD_ROW, expand=False)

    vega_kontrol_col = ctk.CTkFrame(kontrol_row3, fg_color="transparent")
    vega_kontrol_col.grid(row=0, column=0, padx=40)
    ctk.CTkLabel(vega_kontrol_col, text="Vega", font=theme.FONT_H2()).pack(pady=(0, 10))
    ctk.CTkButton(vega_kontrol_col, text="🔍 Kontrol", width=theme.BTN_W_LG, height=theme.BTN_H_LG,
                  fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                  command=kontrol_vega).pack()

    hb_ve_trn_kontrol_col = ctk.CTkFrame(kontrol_row3, fg_color="transparent")
    hb_ve_trn_kontrol_col.grid(row=0, column=1, padx=40)
    ctk.CTkLabel(hb_ve_trn_kontrol_col, text="Hepsiburada/Trendyol", font=theme.FONT_SUBTITLE()).pack(pady=(0, 10))
    ctk.CTkButton(hb_ve_trn_kontrol_col, text="🔍 Kontrol", width=theme.BTN_W_LG, height=theme.BTN_H_LG,
                  fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER,
                  command=kontrol_hb_ve_trn).pack()

    # Sonuç mesajı
    kontrol_label_sonuc = ctk.CTkLabel(kontrol_inner, text="", font=theme.FONT_SUBTITLE())
    kontrol_label_sonuc.pack(pady=(20, 0))

    return kontrol_frame
