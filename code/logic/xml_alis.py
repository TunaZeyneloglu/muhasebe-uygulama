import os
import glob
import xml.etree.ElementTree as ET
from datetime import datetime
from tkinter import filedialog

from constants import MASAUSTU, XML_NS
import state
import theme
from logic.common import cari_adi_al, kaydet_excel

# ----------------- Alış XML -----------------

def xml_alis_oku_ve_yaz(klasor, label_sonuc, buton_ac, fatura_tipi):
    veriler = []
    basarisiz_dosyalar = []

    for xml_dosya in glob.glob(os.path.join(klasor, "**", "*.xml"), recursive=True):
        try:
            tree = ET.parse(xml_dosya)
            root = tree.getroot()
            ns = XML_NS

            tarih_raw = root.findtext(".//cbc:IssueDate", namespaces=ns)
            tarih = datetime.strptime(tarih_raw, "%Y-%m-%d").strftime("%d.%m.%Y")
            evrakno = root.findtext(".//cbc:ID", namespaces=ns)

            # Alış faturası: Cari = Satıcı (Supplier)
            supplier_party = root.find(".//cac:AccountingSupplierParty/cac:Party", namespaces=ns)
            carikodu = supplier_party.findtext("cac:PartyIdentification/cbc:ID", namespaces=ns) if supplier_party is not None else ""
            cariadi = cari_adi_al(supplier_party, ns) if supplier_party is not None else "Bilinmeyen Cari"

            doviz = root.findtext(".//cbc:DocumentCurrencyCode", default="TRY", namespaces=ns)
            if doviz == "TRY": doviz = "TL"

            for line in root.findall(".//cac:InvoiceLine", namespaces=ns):
                miktar = int(float(line.findtext("cbc:InvoicedQuantity", default="0", namespaces=ns)))
                birim = "ADET"
                fiyat = float(line.findtext("cac:Price/cbc:PriceAmount", default="0", namespaces=ns).replace(",", "."))
                stokadi = line.findtext("cac:Item/cbc:Name", namespaces=ns)
                stokkodu = line.findtext("cac:Item/cac:SellersItemIdentification/cbc:ID", namespaces=ns) or ""

                iskonto_element = line.find("cac:AllowanceCharge/cbc:Amount", namespaces=ns)
                iskonto = float(iskonto_element.text.replace(",", ".")) if iskonto_element is not None else 0.0

                kdv_raw = line.findtext("cac:TaxTotal/cac:TaxSubtotal/cbc:Percent", namespaces=ns)
                try:
                    kdv_orani = int(float(kdv_raw))
                except (ValueError, TypeError):
                    kdv_orani = 0

                veriler.append({
                    "SIRANO": "",
                    "TARIH": tarih,
                    "EVRAKNO": evrakno,
                    "EVRAKTURU": "Alış Faturası",
                    "CARIKODU": carikodu,
                    "CARIADI": cariadi,
                    "STOKKODU": stokkodu,
                    "STOKADI": stokadi,
                    "BIRIM": birim,
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
        dosya_yolu = os.path.join(MASAUSTU, f"{fatura_tipi}.xlsx")
        kaydet_excel(veriler, dosya_yolu)
        if basarisiz_dosyalar:
            label_sonuc.configure(text=f"{fatura_tipi}.xlsx kaydedildi, ancak {len(basarisiz_dosyalar)} dosya işlenemedi: {', '.join(basarisiz_dosyalar[:3])}{' ...' if len(basarisiz_dosyalar) > 3 else ''}", text_color=theme.WARNING)
        else:
            label_sonuc.configure(text=f"{fatura_tipi}.xlsx masaüstüne kaydedildi.", text_color=theme.SUCCESS)
        buton_ac.configure(state="normal")
        state.SON_DOSYA_YOLU = dosya_yolu
    elif basarisiz_dosyalar:
        label_sonuc.configure(text=f"Geçerli XML dosyası bulunamadı. ({len(basarisiz_dosyalar)} dosya hata verdi)", text_color=theme.ERROR)
    else:
        label_sonuc.configure(text="Geçerli XML dosyası bulunamadı.", text_color=theme.ERROR)

def klasor_sec_ve_isle_alis(label_sonuc, buton_ac):
    secilen_klasor = filedialog.askdirectory(title="Alış XML klasörünü seçin")
    if secilen_klasor:
        xml_alis_oku_ve_yaz(secilen_klasor, label_sonuc, buton_ac, "mustahsil_alis")
