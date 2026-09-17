import customtkinter as ctk

import theme
from pages.navigation import show_kontrol, show_xml

def build_welcome_page(app):
    # ================= KARŞILAMA SAYFASI =================

    welcome_frame = ctk.CTkFrame(app, fg_color="transparent")

    welcome_inner = ctk.CTkFrame(welcome_frame, fg_color=theme.BG_SURFACE, corner_radius=theme.CORNER_CARD)
    welcome_inner.pack(expand=True, fill="both")

    # Başlık
    welcome_title = ctk.CTkLabel(
        welcome_inner, 
        text="📄 Fatura Yönetim Sistemi", 
        font=theme.FONT_TITLE()
    )
    welcome_title.pack(pady=(30, 8))

    welcome_desc = ctk.CTkLabel(
        welcome_inner,
        text="Yapmak istediğiniz işlemi seçin",
        font=theme.FONT_SUBTITLE(),
        text_color=theme.TEXT_SECONDARY
    )
    welcome_desc.pack(pady=(0, 20))

    # Kart Container - tüm alanı kaplasın
    cards_frame = ctk.CTkFrame(welcome_inner, fg_color="transparent")
    cards_frame.pack(expand=True, fill="both", padx=theme.PAD_CARD, pady=(0, theme.PAD_CARD))
    cards_frame.grid_columnconfigure(0, weight=1)
    cards_frame.grid_columnconfigure(1, weight=1)
    cards_frame.grid_rowconfigure(0, weight=1)

    # Kart hover efekti ve tıklama için yardımcı fonksiyon
    def make_card_clickable(card, all_widgets, command):
        """Karta hover efekti ve tıklama özelliği ekler"""
        normal_color = theme.BG_ELEVATED
        hover_color = theme.BG_HOVER
        click_color = theme.BG_PRESSED

        def on_enter(e):
            card.configure(fg_color=hover_color)

        def on_leave(e):
            card.configure(fg_color=normal_color)

        def on_click(e):
            card.configure(fg_color=click_color)
            card.after(100, lambda: [card.configure(fg_color=normal_color), command()])

        # Kart ve tüm child widget'lara event bağla
        for widget in all_widgets:
            widget.bind("<Enter>", on_enter)
            widget.bind("<Leave>", on_leave)
            widget.bind("<Button-1>", on_click)
            widget.configure(cursor="hand2")

    # XML Kartı
    xml_card = ctk.CTkFrame(cards_frame, fg_color=theme.BG_ELEVATED, corner_radius=theme.CORNER_TILE)
    xml_card.grid(row=0, column=0, padx=(0, 10), sticky="nsew")

    xml_card_content = ctk.CTkFrame(xml_card, fg_color="transparent")
    xml_card_content.pack(expand=True)

    xml_icon = ctk.CTkLabel(xml_card_content, text="📁", font=theme.FONT_ICON_LG())
    xml_icon.pack(pady=(0, 15))

    xml_card_title = ctk.CTkLabel(xml_card_content, text="XML Aktarma", font=theme.FONT_H1())
    xml_card_title.pack(pady=(0, 10))

    xml_card_desc = ctk.CTkLabel(
        xml_card_content, 
        text="XML faturalarını\nExcel'e aktarın",
        font=theme.FONT_BODY(),
        text_color=theme.TEXT_SECONDARY,
        justify="center"
    )
    xml_card_desc.pack()

    # XML kartını tıklanabilir yap
    make_card_clickable(xml_card, [xml_card, xml_card_content, xml_icon, xml_card_title, xml_card_desc], show_xml)

    # Kontrol Kartı
    kontrol_card = ctk.CTkFrame(cards_frame, fg_color=theme.BG_ELEVATED, corner_radius=theme.CORNER_TILE)
    kontrol_card.grid(row=0, column=1, padx=(10, 0), sticky="nsew")

    kontrol_card_content = ctk.CTkFrame(kontrol_card, fg_color="transparent")
    kontrol_card_content.pack(expand=True)

    kontrol_icon = ctk.CTkLabel(kontrol_card_content, text="🔍", font=theme.FONT_ICON_LG())
    kontrol_icon.pack(pady=(0, 15))

    kontrol_card_title = ctk.CTkLabel(kontrol_card_content, text="Fatura Kontrol", font=theme.FONT_H1())
    kontrol_card_title.pack(pady=(0, 10))

    kontrol_card_desc = ctk.CTkLabel(
        kontrol_card_content, 
        text="Zirve ve portal\nfaturalarını karşılaştırın",
        font=theme.FONT_BODY(),
        text_color=theme.TEXT_SECONDARY,
        justify="center"
    )
    kontrol_card_desc.pack()

    # Kontrol kartını tıklanabilir yap
    make_card_clickable(kontrol_card, [kontrol_card, kontrol_card_content, kontrol_icon, kontrol_card_title, kontrol_card_desc], show_kontrol)

    return welcome_frame
