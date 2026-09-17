import os
import glob
import xml.etree.ElementTree as ET
from datetime import datetime
from tkinter import filedialog

from constants import MASAUSTU, TARGET_VKN_FOR_ZERO_VAT, XML_NS
import state
import theme
from logic.common import cari_adi_al, kaydet_excel

# ----------------- Uyumsoft XML -----------------

def xml_uyumsoft_oku_ve_yaz(klasor, label_sonuc, buton_ac, fatura_tipi):
    veriler = []
    basarisiz_dosyalar = []

    for xml_dosya in glob.glob(os.path.join(klasor, "**", "*.xml"), recursive=True):
        try:
            tree = ET.parse(xml_dosya)
            root = tree.getroot()

            ns = XML_NS

            # Tarih ve Evrak No
            tarih_raw = root.findtext(".//cbc:IssueDate", namespaces=ns)
            tarih = datetime.strptime(tarih_raw, "%Y-%m-%d").strftime("%d.%m.%Y")
            evrakno = root.findtext(".//cbc:ID", namespaces=ns)

            # VKN kontrolü - eğer Satıcı VKN'si hedef VKN ise KDV sıfırlanacak
            target_vkn = root.findtext(".//cac:AccountingSupplierParty//cbc:ID[@schemeID='VKN']", namespaces=ns)
            force_kdv_zero = (target_vkn == TARGET_VKN_FOR_ZERO_VAT)

            # Satıcı VKN ve Adı
            supplier_party = root.find(".//cac:AccountingSupplierParty/cac:Party", namespaces=ns)
            carikodu = supplier_party.findtext("cac:PartyIdentification/cbc:ID", namespaces=ns)
            cariadi = cari_adi_al(supplier_party, ns)

            # Yardımcı fonksiyon: güvenli float dönüşümü
            def _to_float(x):
                return float(str(x).replace(",", ".")) if x is not None else 0.0

            # Belge seviyesi AllowanceCharge kontrolü
            document_level_discount = 0.0
            doc_allowance = root.find("cac:AllowanceCharge", namespaces=ns)
            if doc_allowance is not None:
                charge_indicator = (doc_allowance.findtext("cbc:ChargeIndicator", namespaces=ns) or "").strip().lower()
                if charge_indicator == "false":
                    document_level_discount = _to_float(doc_allowance.findtext("cbc:Amount", namespaces=ns))

            # KDV'ye göre gruplama
            kdv_dict = {}
            lines = root.findall(".//cac:InvoiceLine", namespaces=ns)

            for line in lines:
                # BaseAmount (önce AllowanceCharge.BaseAmount'u dene, yoksa LineExtensionAmount)
                base_amount = 0.0
                ac_nodes = line.findall("cac:AllowanceCharge", namespaces=ns)
                if ac_nodes:
                    for ac in ac_nodes:
                        ba = ac.findtext("cbc:BaseAmount", namespaces=ns)
                        if ba is not None:
                            base_amount = _to_float(ba)
                            break
                if base_amount == 0.0:
                    base_amount = _to_float(line.findtext("cbc:LineExtensionAmount", namespaces=ns))

                # İndirim toplamı: ChargeIndicator == "false" olanların Amount toplamı
                discount_total = 0.0
                for ac in ac_nodes:
                    ch = (ac.findtext("cbc:ChargeIndicator", namespaces=ns) or "").strip().lower()
                    if ch == "false":
                        discount_total += _to_float(ac.findtext("cbc:Amount", namespaces=ns))

                # Satır değeri hesapla (KDV hariç)
                line_value = base_amount - discount_total
                if line_value < 0:
                    line_value = 0.0

                # KDV oranı
                kdv_raw = line.findtext("cac:TaxTotal/cac:TaxSubtotal/cbc:Percent", namespaces=ns)
                try:
                    kdv_orani = int(float(kdv_raw))
                except (ValueError, TypeError):
                    kdv_orani = 0

                # VKN'ye göre KDV sıfırlama
                if force_kdv_zero:
                    kdv_orani = 0

                # Grupla (totalleri ve iskonto listelerini tut)
                if kdv_orani not in kdv_dict:
                    kdv_dict[kdv_orani] = {"total": 0.0, "line_discounts": []}
                kdv_dict[kdv_orani]["total"] += line_value
                kdv_dict[kdv_orani]["line_discounts"].append(discount_total)

            # Excel için veri satırlarını oluştur
            for kdv, grp_data in kdv_dict.items():
                toplam = grp_data["total"]
                
                # İskonto hesapla: Satır bazlı varsa onu kullan, yoksa belge bazlı kontrol et
                group_discount = sum(grp_data["line_discounts"])
                indtl = ""
                
                if group_discount > 0:
                    # Satır bazlı iskonto
                    indtl = round(group_discount, 2)
                elif document_level_discount > 0 and len(kdv_dict) == 1:
                    # Belge bazlı iskonto (sadece tek KDV oranı varsa)
                    indtl = round(document_level_discount, 2)
                
                veriler.append({
                    "SIRANO": "",
                    "TARIH": tarih,
                    "EVRAKNO": evrakno,
                    "EVRAKTURU": "Alış Faturası",
                    "CARIKODU": carikodu,
                    "CARIADI": cariadi,
                    "STOKKODU": str(kdv),    # KDV oranı = StokKodu
                    "STOKADI": "",           # Her zaman boş
                    "BIRIM": "ADET",         # Sabit
                    "MIKTAR": 1,             # Sabit
                    "DOVIZC": "TL",          # Sabit
                    "BIRIMFIYAT": round(toplam, 2),
                    "INDTL": indtl if indtl else "",
                    "KDVY": kdv,
                    "TUTARTL": round(toplam, 2),
                    "FATURACK": "",
                })

        except Exception as e:
            label_sonuc.configure(
                text=f"HATA: {os.path.basename(xml_dosya)} - {e}",
                text_color=theme.ERROR
            )
            basarisiz_dosyalar.append(os.path.basename(xml_dosya))

    # Kaydetme işlemi
    if veriler:
        dosya_yolu = os.path.join(MASAUSTU, f"{fatura_tipi}.xlsx")
        kaydet_excel(veriler, dosya_yolu)
        if basarisiz_dosyalar:
            label_sonuc.configure(
                text=f"⚠️ {fatura_tipi}.xlsx kaydedildi, ancak {len(basarisiz_dosyalar)} dosya işlenemedi: {', '.join(basarisiz_dosyalar[:3])}{' ...' if len(basarisiz_dosyalar) > 3 else ''}",
                text_color=theme.WARNING
            )
        else:
            label_sonuc.configure(
                text=f"✅ {fatura_tipi}.xlsx masaüstüne kaydedildi.",
                text_color=theme.SUCCESS
            )
        buton_ac.configure(state="normal")
        state.SON_DOSYA_YOLU = dosya_yolu
    elif basarisiz_dosyalar:
        label_sonuc.configure(
            text=f"Geçerli XML dosyası bulunamadı. ({len(basarisiz_dosyalar)} dosya hata verdi)",
            text_color=theme.ERROR
        )
    else:
        label_sonuc.configure(
            text="Geçerli XML dosyası bulunamadı.",
            text_color=theme.ERROR
        )

def klasor_sec_ve_isle_uyumsoft(label_sonuc, buton_ac):
    secilen_klasor = filedialog.askdirectory(title="Uyumsoft XML klasörünü seçin")
    if secilen_klasor:
        xml_uyumsoft_oku_ve_yaz(secilen_klasor, label_sonuc, buton_ac, "uyumsoft_faturalar")
