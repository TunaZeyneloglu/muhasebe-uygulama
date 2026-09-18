import customtkinter as ctk

import icons
import theme
from pages import ui
from pages.navigation import show_kontrol, show_xml

def build_welcome_page(app):
    # ================= KARŞILAMA SAYFASI =================
    # Görsel dil: üst araç çubuğu + hero başlık + iki zengin kart.

    welcome_frame = ctk.CTkFrame(app, fg_color=theme.BG_ROOT, corner_radius=0)

    ui.build_toolbar(welcome_frame)
    ui.build_statusbar(welcome_frame, "Hazır  ·  bir işlem seçin")

    # Gövde
    wrap = ctk.CTkFrame(welcome_frame, fg_color="transparent")
    wrap.pack(fill="both", expand=True, padx=theme.PAD_STAGE, pady=(38, 32))

    ctk.CTkLabel(wrap, text="ÇALIŞMA ALANI", font=theme.FONT_EYEBROW(),
                 text_color=theme.ACCENT).pack(anchor="w")
    ctk.CTkLabel(wrap, text="Hangi işlemi yapmak istersiniz?", font=theme.FONT_HERO(),
                 text_color=theme.TEXT_PRIMARY).pack(anchor="w", pady=(10, 8))
    ctk.CTkLabel(wrap,
                 text="Fatura verilerinizi aktarın ya da iki liste arasında karşılaştırma yapın.",
                 font=theme.FONT_BODY(), text_color=theme.TEXT_SECONDARY).pack(anchor="w")

    # Kart Container - tüm alanı kaplasın
    cards_frame = ctk.CTkFrame(wrap, fg_color="transparent")
    cards_frame.pack(fill="both", expand=True, pady=(28, 0))
    cards_frame.grid_columnconfigure((0, 1), weight=1, uniform="welcome_col")
    cards_frame.grid_rowconfigure(0, weight=1)

    def build_card(col, baslik, aciklama, maddeler, renk, glif, command):
        """Zengin karşılama kartı: ikon kutusu, başlık, açıklama, maddeler, 'Aç' butonu."""
        card = ctk.CTkFrame(cards_frame, fg_color=theme.BG_SURFACE, corner_radius=theme.CORNER_CARD,
                            border_width=theme.BORDER_WIDTH, border_color=theme.BORDER)
        card.grid(row=0, column=col, padx=(0, 11) if col == 0 else (11, 0), sticky="nsew")

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=theme.PAD_CARD, pady=theme.PAD_CARD)

        tile = ctk.CTkLabel(inner, text="", image=glif(theme.ICON_XL, renk),
                            width=theme.CHIP_SIZE, height=theme.CHIP_SIZE,
                            corner_radius=theme.CORNER_CARD, fg_color=theme.TINT(renk))
        tile.pack(anchor="w")

        title = ctk.CTkLabel(inner, text=baslik, font=theme.FONT_H1(),
                             text_color=theme.TEXT_PRIMARY)
        title.pack(anchor="w", pady=(20, 7))

        desc = ctk.CTkLabel(inner, text=aciklama, font=theme.FONT_BODY(),
                            text_color=theme.TEXT_SECONDARY, wraplength=330, justify="left")
        desc.pack(anchor="w")

        # Artan dikey boşluk buraya toplanır; böylece son madde ile "Aç" butonu
        # arasında boşluk kalmaz, kartın üst ve alt blokları ayrışır.
        bosluk = ctk.CTkFrame(inner, fg_color="transparent", height=1)
        bosluk.pack(fill="both", expand=True)

        sep = ctk.CTkFrame(inner, height=theme.DIVIDER_H, fg_color=theme.BORDER)
        sep.pack(fill="x", pady=(0, 15))

        widgets = [card, inner, tile, title, desc, bosluk, sep]
        for madde in maddeler:
            satir = ctk.CTkFrame(inner, fg_color="transparent")
            satir.pack(anchor="w", pady=3, fill="x")
            isaret = ctk.CTkLabel(satir, text="", image=icons.onay(theme.ICON_XS, renk))
            isaret.pack(side="left")
            etiket = ctk.CTkLabel(satir, text=madde, font=theme.FONT_SMALL(),
                                  text_color=theme.TEXT_SECONDARY, wraplength=300,
                                  justify="left")
            etiket.pack(side="left", padx=9)
            widgets += [satir, isaret, etiket]

        # "Aç" gerçek bir buton: kendi hover'ı ve tıklama davranışı var
        ac_btn = ctk.CTkButton(inner, text="Aç", image=icons.ok_sag(theme.ICON_SM, renk),
                               compound="right", width=104, height=36,
                               corner_radius=theme.CORNER_BTN, font=theme.FONT_LINK(),
                               fg_color="transparent", hover_color=theme.TINT_HOVER(renk),
                               text_color=renk, border_width=theme.BORDER_WIDTH,
                               border_color=theme.karistir(renk, 0.45),
                               command=command)
        ac_btn.pack(anchor="w", pady=(16, 0))

        def on_enter(_=None):
            card.configure(border_color=renk, fg_color=theme.BG_ELEVATED)

        def on_leave(_=None):
            card.configure(border_color=theme.BORDER, fg_color=theme.BG_SURFACE)

        def on_click(_=None):
            command()

        # Kartın tamamı tıklanabilir; "Aç" butonu kendi command'ını çalıştırır
        # (Tk'de üst widget binding'leri child olaylarında tetiklenmez).
        for widget in widgets:
            widget.bind("<Enter>", on_enter)
            widget.bind("<Leave>", on_leave)
            widget.bind("<Button-1>", on_click)
            widget.configure(cursor="hand2")
        ac_btn.bind("<Enter>", on_enter)
        ac_btn.bind("<Leave>", on_leave)

        return card, ac_btn

    # XML Kartı
    xml_card, xml_ac_btn = build_card(
        0, "XML Aktarma",
        "XML faturalarını okuyup tek bir Excel çalışma kitabına aktarır.",
        ["Logo, Uyumsoft, Hepsiburada, Trendyol …",
         "Alış, satış ve müstahsil belgeleri",
         "Hatalı dosyalar ayrıca raporlanır",
         "Çıktı doğrudan masaüstüne kaydedilir"],
        theme.ACCENT, icons.belge, show_xml)

    # Kontrol Kartı
    kontrol_card, kontrol_ac_btn = build_card(
        1, "Fatura Kontrol",
        "İki Excel kaynağındaki faturaları karşılaştırıp farkları raporlar.",
        ["Zirve ve portal listeleri",
         "Eksik, fazla ve tutar farkları",
         "Sonuçlar renkli Excel çıktısı olarak",
         "Raporu tek tıkla açabilirsiniz"],
        theme.SUCCESS, icons.karsilastir, show_kontrol)

    return welcome_frame
