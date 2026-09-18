import os
import pandas as pd
import subprocess
import sys
from openpyxl.styles import Font

import state

# ----------------- Ortak Fonksiyonlar -----------------

def cari_adi_al(party, ns):
    cariadi = party.findtext("cac:PartyName/cbc:Name", namespaces=ns)
    if not cariadi:
        firstname = party.findtext("cac:Person/cbc:FirstName", namespaces=ns)
        lastname = party.findtext("cac:Person/cbc:FamilyName", namespaces=ns)
        if firstname and lastname:
            cariadi = f"{firstname} {lastname}"
    if not cariadi:
        cariadi = party.findtext("cac:PartyLegalEntity/cbc:RegistrationName", namespaces=ns)
    return cariadi if cariadi else "Bilinmeyen Cari"

def kaydet_excel(veriler, dosya_yolu):
    df = pd.DataFrame(veriler)[[
        "SIRANO", "TARIH", "EVRAKNO", "EVRAKTURU", "CARIKODU", "CARIADI",
        "STOKKODU", "STOKADI", "BIRIM", "MIKTAR", "DOVIZC",
        "BIRIMFIYAT", "INDTL", "KDVY", "TUTARTL", "FATURACK"
    ]]
    # Metin sütunları: 0 ile başlayan kodların korunması için
    df["EVRAKNO"] = df["EVRAKNO"].fillna("").astype(str)
    df["CARIKODU"] = df["CARIKODU"].fillna("").astype(str)
    df["STOKKODU"] = df["STOKKODU"].fillna("").astype(str)
    with pd.ExcelWriter(dosya_yolu, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sayfa1")

def dosyayi_ac():
    if state.SON_DOSYA_YOLU:
        try:
            if sys.platform == "win32":
                os.startfile(state.SON_DOSYA_YOLU)
            elif sys.platform == "darwin":
                subprocess.call(["open", state.SON_DOSYA_YOLU])
            else:
                subprocess.call(["xdg-open", state.SON_DOSYA_YOLU])
        except Exception:
            pass

# ----------------- Ortak Yardımcı Fonksiyonlar (Karşılaştırma) -----------------

def _temizle_fatura_no(fatura_no):
    """Fatura numarasını temizle ve normalize et"""
    if pd.isna(fatura_no):
        return ""
    return str(fatura_no).strip().replace("'", "").upper()

def _satir_bulucu_olustur(df, anahtar_sutun='Fatura_No_Temiz'):
    """Anahtar sütununa göre O(1) çalışan bir satır arayıcı üretir.
    
    Dönen fonksiyon, df[df[anahtar_sutun] == anahtar].iloc[0] ile birebir
    aynı satırı döndürür: yinelenen anahtarlarda ilk satır kazanır.
    """
    indeks = {}
    for konum, anahtar in enumerate(df[anahtar_sutun]):
        # setdefault: sonraki aynı anahtar ilkinin üzerine yazmaz
        indeks.setdefault(anahtar, konum)
    
    def satir_bul(anahtar):
        konum = indeks.get(anahtar)
        if konum is None:
            # İndekste yer almayan anahtar (ör. NaN) -> eski tarama davranışı
            return df[df[anahtar_sutun] == anahtar].iloc[0]
        return df.iloc[konum]
    
    return satir_bul

def _fazla_girisler_olustur(satir_bul, fatura_nolar, tutar_sutun='Tutar_TL', kdv_sutun='KDV_TL'):
    """Zirve'de olup portalda olmayan faturalar için 'Fazla Girişler' satırlarını üretir.
    
    Bu satır yapısı kontrol akışlarında birebir aynıdır; sadece tutar/KDV
    sütun adları akışa göre değişir.
    """
    fazla_girisler = []
    for fatura_no in fatura_nolar:
        if not fatura_no:
            continue
        zirve_row = satir_bul(fatura_no)
        fazla_girisler.append({
            'Fatura No': fatura_no,
            'Cari Unvanı': zirve_row.get('Cari Unvanı', ''),
            'Tutar TL': round(zirve_row[tutar_sutun], 2),
            'KDV TL': round(zirve_row[kdv_sutun], 2),
            'Tarih': zirve_row.get('Tarih', ''),
            'E.Kod': zirve_row.get('E.Kod', ''),
            'İşlem Türü': zirve_row.get('İşlem Türü', '')
        })
    return fazla_girisler

def _excel_dosyalari_oku(zirve_path, portal_paths):
    """Zirve ve portal Excel dosyalarını oku ve birleştir"""
    df_zirve = pd.read_excel(zirve_path)
    
    if isinstance(portal_paths, list):
        portal_dfs = [pd.read_excel(path) for path in portal_paths]
        df_portal = pd.concat(portal_dfs, ignore_index=True)
    else:
        df_portal = pd.read_excel(portal_paths)
    
    return df_zirve, df_portal

def _kontrol_sonuc_excel_yaz(sonuc_dosyasi, tutar_farklari, eksik_girisler, fazla_girisler):
    """Kontrol sonuçlarını Excel'e yaz"""
    with pd.ExcelWriter(sonuc_dosyasi, engine='openpyxl') as writer:
        # Sayfa 1: Tutar Farkları
        if tutar_farklari:
            df_tutar = pd.DataFrame(tutar_farklari).sort_values('Fatura No')
            df_tutar.to_excel(writer, sheet_name='Tutar Farkları', index=False)
            
            # Formatlama
            worksheet = writer.sheets['Tutar Farkları']
            headers = list(df_tutar.columns)
            tutar_farki_col = headers.index('Tutar Farkı') + 1
            kdv_farki_col = headers.index('KDV Farkı') + 1
            
            for row_idx in range(2, len(df_tutar) + 2):
                worksheet.cell(row=row_idx, column=tutar_farki_col).font = Font(bold=True, size=13)
                worksheet.cell(row=row_idx, column=kdv_farki_col).font = Font(bold=True, size=13)
        else:
            pd.DataFrame({'Mesaj': ['Tutar farkı bulunamadı.']}).to_excel(
                writer, sheet_name='Tutar Farkları', index=False)
        
        # Sayfa 2: Eksik Girişler
        if eksik_girisler:
            pd.DataFrame(eksik_girisler).sort_values('Fatura No').to_excel(
                writer, sheet_name='Eksik Girişler', index=False)
        else:
            pd.DataFrame({'Mesaj': ['Eksik giriş bulunamadı.']}).to_excel(
                writer, sheet_name='Eksik Girişler', index=False)
        
        # Sayfa 3: Fazla Girişler
        if fazla_girisler:
            pd.DataFrame(fazla_girisler).sort_values('Fatura No').to_excel(
                writer, sheet_name='Fazla Girişler', index=False)
        else:
            pd.DataFrame({'Mesaj': ['Fazla giriş bulunamadı.']}).to_excel(
                writer, sheet_name='Fazla Girişler', index=False)

def _kontrol_ozet_mesaj_olustur(portal_paths, df_portal, df_zirve, tutar_farklari, eksik_girisler, fazla_girisler):
    """Kontrol sonucu özet mesajı oluştur"""
    portal_dosya_sayisi = len(portal_paths) if isinstance(portal_paths, list) else 1
    mesaj = f"Kontrol tamamlandı!\n\n"
    mesaj += f"İşlenen Portal Dosyası: {portal_dosya_sayisi}\n"
    mesaj += f"Toplam Portal Fatura: {len(df_portal)}\n"
    mesaj += f"Toplam Zirve Fatura: {len(df_zirve)}\n\n"
    mesaj += f"Tutar Farkları: {len(tutar_farklari)}\n"
    mesaj += f"Eksik Girişler: {len(eksik_girisler)}\n"
    mesaj += f"Fazla Girişler: {len(fazla_girisler)}\n\n"
    mesaj += f"Sonuç dosyası masaüstüne kaydedildi."
    return mesaj
