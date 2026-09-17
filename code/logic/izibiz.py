import os
import pandas as pd

from constants import MASAUSTU
import state
from logic.common import _kontrol_ozet_mesaj_olustur, _kontrol_sonuc_excel_yaz, _temizle_fatura_no
from logic.hizlibilisim import _izibiz_normalize_portal

# ----------------- İzibiz Kontrol Fonksiyonu -----------------

def izibiz_karsilastir(zirve_path, portal_paths):
    """Zirve ve İzibiz Portal Excel dosyalarını karşılaştırır.
    
    Alış ve satış faturalarını otomatik tespit eder.
    """
    
    try:
        # Zirve dosyasını oku
        df_zirve = pd.read_excel(zirve_path)
        df_zirve['Fatura_No_Temiz'] = df_zirve['Fatura No'].apply(_temizle_fatura_no)
        df_zirve['Tutar_TL'] = pd.to_numeric(df_zirve['Tutar TL'], errors='coerce').fillna(0)
        df_zirve['KDV_TL'] = pd.to_numeric(df_zirve['KDV TL'], errors='coerce').fillna(0)
        
        # Portal dosyalarını oku ve normalize et
        portal_dfs = []
        if isinstance(portal_paths, list):
            for path in portal_paths:
                df = pd.read_excel(path)
                df = _izibiz_normalize_portal(df)
                portal_dfs.append(df)
            df_portal = pd.concat(portal_dfs, ignore_index=True)
        else:
            df_portal = pd.read_excel(portal_paths)
            df_portal = _izibiz_normalize_portal(df_portal)
        
        # Fatura_No_Temiz normalize fonksiyonunda oluşturuldu
        
        # Set'ler oluştur
        zirve_faturalar = set(df_zirve['Fatura_No_Temiz']) - {''}
        portal_faturalar = set(df_portal['Fatura_No_Temiz']) - {''}
        
        # 1. TUTAR FARKLARI
        tutar_farklari = []
        for fatura_no in zirve_faturalar.intersection(portal_faturalar):
            if not fatura_no:
                continue
            zirve_row = df_zirve[df_zirve['Fatura_No_Temiz'] == fatura_no].iloc[0]
            portal_row = df_portal[df_portal['Fatura_No_Temiz'] == fatura_no].iloc[0]
            
            tutar_farki = abs(zirve_row['Tutar_TL'] - portal_row['Portal_Tutar'])
            kdv_farki = abs(zirve_row['KDV_TL'] - portal_row['Portal_KDV'])
            
            if tutar_farki > 1 or kdv_farki > 1:
                tutar_farklari.append({
                    'Fatura No': fatura_no,
                    'Zirve Cari': zirve_row.get('Cari Unvanı', ''),
                    'Portal Cari': portal_row.get('Portal_Cari', ''),
                    'Zirve Tutar': round(zirve_row['Tutar_TL'], 2),
                    'Portal Tutar': round(portal_row['Portal_Tutar'], 2),
                    'Tutar Farkı': round(tutar_farki, 2),
                    'Zirve KDV': round(zirve_row['KDV_TL'], 2),
                    'Portal KDV': round(portal_row['Portal_KDV'], 2),
                    'KDV Farkı': round(kdv_farki, 2),
                    'Zirve Tarih': zirve_row.get('Tarih', ''),
                    'Portal Tarih': portal_row.get('Portal_Tarih', '')
                })
        
        # 2. EKSİK GİRİŞLER
        eksik_girisler = []
        for fatura_no in (portal_faturalar - zirve_faturalar):
            if not fatura_no:
                continue
            portal_row = df_portal[df_portal['Fatura_No_Temiz'] == fatura_no].iloc[0]
            eksik_girisler.append({
                'Fatura No': fatura_no,
                'VKN/TCKN': portal_row.get('Portal_VKN', ''),
                'Cari': portal_row.get('Portal_Cari', ''),
                'Ödenecek Tutar': round(portal_row['Portal_Tutar'], 2),
                'Toplam KDV': round(portal_row['Portal_KDV'], 2),
                'Fatura Tarihi': portal_row.get('Portal_Tarih', ''),
                'Fatura Durumu': portal_row.get('Portal_Durum', '')
            })
        
        # 3. FAZLA GİRİŞLER
        fazla_girisler = []
        for fatura_no in (zirve_faturalar - portal_faturalar):
            if not fatura_no:
                continue
            zirve_row = df_zirve[df_zirve['Fatura_No_Temiz'] == fatura_no].iloc[0]
            fazla_girisler.append({
                'Fatura No': fatura_no,
                'Cari Unvanı': zirve_row.get('Cari Unvanı', ''),
                'Tutar TL': round(zirve_row['Tutar_TL'], 2),
                'KDV TL': round(zirve_row['KDV_TL'], 2),
                'Tarih': zirve_row.get('Tarih', ''),
                'E.Kod': zirve_row.get('E.Kod', ''),
                'İşlem Türü': zirve_row.get('İşlem Türü', '')
            })
        
        # Excel dosyası oluştur
        sonuc_dosyasi = os.path.join(MASAUSTU, "Izibiz_Kontrol_Sonuc.xlsx")
        _kontrol_sonuc_excel_yaz(sonuc_dosyasi, tutar_farklari, eksik_girisler, fazla_girisler)
        
        state.SON_DOSYA_YOLU = sonuc_dosyasi
        mesaj = _kontrol_ozet_mesaj_olustur(portal_paths, df_portal, df_zirve, tutar_farklari, eksik_girisler, fazla_girisler)
        
        return (sonuc_dosyasi, mesaj)
        
    except Exception as e:
        import traceback
        return (None, f"❌ Hata oluştu:\n{str(e)}\n\n{traceback.format_exc()}")
