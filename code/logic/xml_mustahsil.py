import os
import glob
import xml.etree.ElementTree as ET
import pandas as pd
import unicodedata
from tkinter import filedialog

from constants import MASAUSTU, XML_NS
import state
import theme

# ----------------- Müstahsil XML -----------------

def temizle_kelime(kelime):
    kelime = kelime.strip().lower()
    kelime = kelime.replace("ı", "i").replace("ş", "s").replace("ç", "c").replace("ğ", "g").replace("ü", "u").replace("ö", "o")
    kelime = unicodedata.normalize("NFD", kelime).encode("ascii", "ignore").decode("utf-8")
    return kelime

def xml_mustahsil_oku_ve_yaz(klasor, label_sonuc, buton_ac, fatura_tipi):

    urun_kodlari = {
        "patates": "08120101",
        "kuru sogan": "08110101",
        "sarimsak": "08190101",
        "limon": "16030301"
    }

    ns = XML_NS

    veriler = []
    basarisiz_dosyalar = []

    for xml_dosya in glob.glob(os.path.join(klasor, "**", "*.xml"), recursive=True):
        try:
            tree = ET.parse(xml_dosya)
            root = tree.getroot()

            belge_no = root.findtext(".//cbc:ID", namespaces=ns)
            belge_tarih = root.findtext(".//cbc:IssueDate", namespaces=ns)

            # TC kimlik no
            tc_no = ""
            for elem in root.findall(".//cac:AccountingCustomerParty//cbc:ID", namespaces=ns):
                if elem.attrib.get("schemeID") == "TCKN":
                    tc_no = elem.text.strip()
                    break

            for line in root.findall(".//cac:CreditNoteLine", namespaces=ns):
                miktar = line.findtext(".//cbc:CreditedQuantity", namespaces=ns)
                birim_fiyat = line.findtext(".//cbc:PriceAmount", namespaces=ns)
                tutar = line.findtext(".//cbc:LineExtensionAmount", namespaces=ns)
                urun_adi_raw = line.findtext(".//cac:Item//cbc:Name", namespaces=ns)

                if not urun_adi_raw:
                    continue

                urun_adi = temizle_kelime(urun_adi_raw)
                urun_kodu = urun_kodlari.get(urun_adi, "")

                # Mahsül yılı: belge tarihinin yılından 1 çıkar (önceki yılın mahsulü)
                mahsul_yili = str(int(belge_tarih[:4]) - 1) if belge_tarih else ""
                
                veriler.append({
                    "kimlik tipi": "k",
                    "TC KİMLİK NO": tc_no,
                    "MİKTAR": int(float(miktar)) if miktar else "",
                    "BİRİM FİYATI": birim_fiyat.replace(".", ",") if birim_fiyat else "",
                    "SATIS_TUTAR": tutar.replace(".", ",") if tutar else "",
                    "BELGE NO": belge_no,
                    "BELGE TARİHİ": f"{int(belge_tarih[8:10])}/{int(belge_tarih[5:7])}/{belge_tarih[:4]}" if belge_tarih else "",
                    "ÜRÜN KODU": urun_kodu,
                    "MENŞE": "001",
                    "MAHSÜL YILI": mahsul_yili,
                    "SATIŞ ŞEKLİ": "011",
                    "BAĞKUR MİKTAR": "",
                    "YRD_MIKTAR": "",
                    "YRD_OLCU_BIRIM": ""
                })

        except Exception as e:
            label_sonuc.configure(text=f"HATA: {os.path.basename(xml_dosya)} - {e}", text_color=theme.ERROR)
            basarisiz_dosyalar.append(os.path.basename(xml_dosya))

    if veriler:
        df = pd.DataFrame(veriler)
        dosya_yolu = os.path.join(MASAUSTU, f"{fatura_tipi}.xlsx")
        df.to_excel(dosya_yolu, index=False)
        if basarisiz_dosyalar:
            label_sonuc.configure(text=f"⚠️ {fatura_tipi}.xlsx kaydedildi, ancak {len(basarisiz_dosyalar)} dosya işlenemedi: {', '.join(basarisiz_dosyalar[:3])}{' ...' if len(basarisiz_dosyalar) > 3 else ''}", text_color=theme.WARNING)
        else:
            label_sonuc.configure(text=f"✅ {fatura_tipi}.xlsx masaüstüne kaydedildi.", text_color=theme.SUCCESS)
        buton_ac.configure(state="normal")
        state.SON_DOSYA_YOLU = dosya_yolu
    elif basarisiz_dosyalar:
        label_sonuc.configure(text=f"Geçerli XML dosyası bulunamadı. ({len(basarisiz_dosyalar)} dosya hata verdi)", text_color=theme.ERROR)
    else:
        label_sonuc.configure(text="Geçerli XML dosyası bulunamadı.", text_color=theme.ERROR)

def klasor_sec_ve_isle_mustahsil(label_sonuc, buton_ac):
    secilen_klasor = filedialog.askdirectory(title="Müstahsil Borsa XML klasörünü seçin")
    if secilen_klasor:
        xml_mustahsil_oku_ve_yaz(secilen_klasor, label_sonuc, buton_ac, "mustahsil_listesi")
