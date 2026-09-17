import os
import glob
import xml.etree.ElementTree as ET
from datetime import datetime
from tkinter import filedialog

from constants import MASAUSTU, XML_NS
import state
import theme
from logic.common import cari_adi_al, kaydet_excel

# ----------------- Hepsiburada XML -----------------

def xml_hepsiburada_oku_ve_yaz(klasor, label_sonuc, buton_ac, fatura_tipi):
    """
    Hepsiburada XML faturalarını okur ve Excel'e yazar.
    
    Args:
        fatura_tipi: "alis" veya "satis"
    """
    
    # Fatura tipine göre ayarlar
    if fatura_tipi == "alis":
        evrak_turu = "Alış Faturası"
        party_path = ".//cac:AccountingSupplierParty/cac:Party"
        dosya_adi = "hepsiburada_alis.xlsx"
    else:
        evrak_turu = "Satış Faturası"
        party_path = ".//cac:AccountingCustomerParty/cac:Party"
        dosya_adi = "hepsiburada_satis.xlsx"
    
    veriler = []
    basarisiz_dosyalar = []
    ns = XML_NS
    
    for xml_dosya in glob.glob(os.path.join(klasor, "**", "*.xml"), recursive=True):
        try:
            tree = ET.parse(xml_dosya)
            root = tree.getroot()
            
            tarih_raw = root.findtext(".//cbc:IssueDate", namespaces=ns)
            tarih = datetime.strptime(tarih_raw, "%Y-%m-%d").strftime("%d.%m.%Y")
            evrakno = root.findtext(".//cbc:ID", namespaces=ns)
            
            # Cari bilgisi (fatura tipine göre)
            party = root.find(party_path, namespaces=ns)
            carikodu = party.findtext("cac:PartyIdentification/cbc:ID", namespaces=ns) if party is not None else ""
            cariadi = cari_adi_al(party, ns) if party is not None else "Bilinmeyen Cari"
            
            doviz = root.findtext(".//cbc:DocumentCurrencyCode", default="TRY", namespaces=ns)
            if doviz == "TRY": doviz = "TL"
            
            for line in root.findall(".//cac:InvoiceLine", namespaces=ns):
                miktar = int(float(line.findtext("cbc:InvoicedQuantity", default="0", namespaces=ns)))
                fiyat = float(line.findtext("cac:Price/cbc:PriceAmount", default="0", namespaces=ns).replace(",", "."))
                stokadi = line.findtext("cac:Item/cbc:Name", namespaces=ns)
                stokkodu = line.findtext("cac:Item/cac:SellersItemIdentification/cbc:ID", namespaces=ns) or ""
                
                # İskonto
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
                    "EVRAKTURU": evrak_turu,
                    "CARIKODU": carikodu,
                    "CARIADI": cariadi,
                    "STOKKODU": stokkodu,
                    "STOKADI": stokadi,
                    "BIRIM": "ADET",
                    "MIKTAR": miktar,
                    "DOVIZC": doviz,
                    "BIRIMFIYAT": fiyat,
                    "INDTL": iskonto if iskonto else "",
                    "KDVY": kdv_orani,
                    "TUTARTL": fiyat * miktar,
                    "FATURACK": "",
                })
        except Exception as e:
            label_sonuc.configure(text=f"HATA: {os.path.basename(xml_dosya)} - {e}", text_color=theme.ERROR)
            basarisiz_dosyalar.append(os.path.basename(xml_dosya))

    if veriler:
        dosya_yolu = os.path.join(MASAUSTU, dosya_adi)
        kaydet_excel(veriler, dosya_yolu)
        if basarisiz_dosyalar:
            label_sonuc.configure(text=f"⚠️ {dosya_adi} kaydedildi, ancak {len(basarisiz_dosyalar)} dosya işlenemedi: {', '.join(basarisiz_dosyalar[:3])}{' ...' if len(basarisiz_dosyalar) > 3 else ''}", text_color=theme.WARNING)
        else:
            label_sonuc.configure(text=f"✅ {dosya_adi} masaüstüne kaydedildi.", text_color=theme.SUCCESS)
        buton_ac.configure(state="normal")
        state.SON_DOSYA_YOLU = dosya_yolu
    elif basarisiz_dosyalar:
        label_sonuc.configure(text=f"Geçerli XML dosyası bulunamadı. ({len(basarisiz_dosyalar)} dosya hata verdi)", text_color=theme.ERROR)
    else:
        label_sonuc.configure(text="Geçerli XML dosyası bulunamadı.", text_color=theme.ERROR)

def klasor_sec_hepsiburada_alis(label_sonuc, buton_ac):
    secilen_klasor = filedialog.askdirectory(title="Hepsiburada Alış Faturası klasörünü seçin")
    if secilen_klasor:
        xml_hepsiburada_oku_ve_yaz(secilen_klasor, label_sonuc, buton_ac, "alis")

def klasor_sec_hepsiburada_satis(label_sonuc, buton_ac):
    secilen_klasor = filedialog.askdirectory(title="Hepsiburada Satış Faturası klasörünü seçin")
    if secilen_klasor:
        xml_hepsiburada_oku_ve_yaz(secilen_klasor, label_sonuc, buton_ac, "satis")
