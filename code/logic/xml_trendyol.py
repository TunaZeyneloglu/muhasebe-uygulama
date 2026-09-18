import os
import glob
import xml.etree.ElementTree as ET
from datetime import datetime
from tkinter import filedialog

from constants import MASAUSTU, XML_NS
import state
import theme
from logic.common import cari_adi_al, kaydet_excel

# ----------------- Trendyol XML -----------------

def xml_trendyol_satis_oku_ve_yaz(klasor, label_sonuc, buton_ac):
    veriler = []
    basarisiz_dosyalar = []
    ns = XML_NS
    
    for xml_dosya in glob.glob(os.path.join(klasor, "**", "*.xml"), recursive=True):
        try:
            tree = ET.parse(xml_dosya)
            root = tree.getroot()
            
            tarih_raw = root.findtext("cbc:IssueDate", namespaces=ns)
            tarih = datetime.strptime(tarih_raw, "%Y-%m-%d").strftime("%d.%m.%Y")
            evrakno = root.findtext("cbc:ID", namespaces=ns)
            
            # Satış faturası: Cari = Alıcı (Customer)
            customer_party = root.find("cac:AccountingCustomerParty/cac:Party", namespaces=ns)
            carikodu = customer_party.findtext("cac:PartyIdentification/cbc:ID", namespaces=ns) if customer_party is not None else ""
            cariadi = cari_adi_al(customer_party, ns) if customer_party is not None else "Bilinmeyen Cari"
            
            doviz = root.findtext("cbc:DocumentCurrencyCode", default="TRY", namespaces=ns)
            if doviz == "TRY": doviz = "TL"
            
            # Fatura düzeyinde toplam iskonto
            toplam_iskonto_element = root.findtext("cac:LegalMonetaryTotal/cbc:AllowanceTotalAmount", namespaces=ns)
            toplam_iskonto = float(toplam_iskonto_element.replace(",", ".")) if toplam_iskonto_element else 0.0
            
            lines = root.findall("cac:InvoiceLine", namespaces=ns)
            for idx, line in enumerate(lines):
                miktar = int(float(line.findtext("cbc:InvoicedQuantity", default="0", namespaces=ns)))
                fiyat = float(line.findtext("cac:Price/cbc:PriceAmount", default="0", namespaces=ns).replace(",", "."))
                stokadi = line.findtext("cac:Item/cbc:Name", namespaces=ns)
                stokkodu = line.findtext("cac:Item/cac:SellersItemIdentification/cbc:ID", namespaces=ns) or ""
                
                # İskonto: toplam iskontoyu sadece ilk satıra yaz, yoksa boş
                iskonto = toplam_iskonto if idx == 0 and toplam_iskonto else ""
                
                kdv_raw = line.findtext("cac:TaxTotal/cac:TaxSubtotal/cbc:Percent", namespaces=ns)
                try:
                    kdv_orani = int(float(kdv_raw))
                except (ValueError, TypeError):
                    kdv_orani = 0
                
                veriler.append({
                    "SIRANO": "",
                    "TARIH": tarih,
                    "EVRAKNO": evrakno,
                    "EVRAKTURU": "Satış Faturası",
                    "CARIKODU": carikodu,
                    "CARIADI": cariadi,
                    "STOKKODU": stokkodu,
                    "STOKADI": stokadi,
                    "BIRIM": "ADET",
                    "MIKTAR": miktar,
                    "DOVIZC": doviz,
                    "BIRIMFIYAT": fiyat,
                    "INDTL": iskonto,
                    "KDVY": kdv_orani,
                    "TUTARTL": fiyat * miktar,
                    "FATURACK": "",
                })
        except Exception as e:
            label_sonuc.configure(text=f"HATA: {os.path.basename(xml_dosya)} - {e}", text_color=theme.ERROR)
            basarisiz_dosyalar.append(os.path.basename(xml_dosya))

    if veriler:
        dosya_yolu = os.path.join(MASAUSTU, "trendyol_satis.xlsx")
        kaydet_excel(veriler, dosya_yolu)
        if basarisiz_dosyalar:
            label_sonuc.configure(text=f"trendyol_satis.xlsx kaydedildi, ancak {len(basarisiz_dosyalar)} dosya işlenemedi: {', '.join(basarisiz_dosyalar[:3])}{' ...' if len(basarisiz_dosyalar) > 3 else ''}", text_color=theme.WARNING)
        else:
            label_sonuc.configure(text="trendyol_satis.xlsx masaüstüne kaydedildi.", text_color=theme.SUCCESS)
        buton_ac.configure(state="normal")
        state.SON_DOSYA_YOLU = dosya_yolu
    elif basarisiz_dosyalar:
        label_sonuc.configure(text=f"Geçerli XML dosyası bulunamadı. ({len(basarisiz_dosyalar)} dosya hata verdi)", text_color=theme.ERROR)
    else:
        label_sonuc.configure(text="Geçerli XML dosyası bulunamadı.", text_color=theme.ERROR)

def klasor_sec_trendyol_alis(label_sonuc, buton_ac):
    # Henüz Trendyol Alış faturası desteği yok
    label_sonuc.configure(text="Trendyol Alış faturası henüz desteklenmiyor.", text_color=theme.WARNING)

def klasor_sec_trendyol_satis(label_sonuc, buton_ac):
    secilen_klasor = filedialog.askdirectory(title="Trendyol Satış Faturası klasörünü seçin")
    if secilen_klasor:
        xml_trendyol_satis_oku_ve_yaz(secilen_klasor, label_sonuc, buton_ac)
