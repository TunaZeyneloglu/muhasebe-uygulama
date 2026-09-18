import os
import pandas as pd

from constants import MASAUSTU
import state
from logic.common import _fazla_girisler_olustur, _kontrol_ozet_mesaj_olustur, _kontrol_sonuc_excel_yaz, _satir_bulucu_olustur, _temizle_fatura_no
from logic.hizlibilisim import _hizlibilisim_normalize_portal, _logo_normalize_portal

# ----------------- Logo + Hızlıbilişim Kombine Kontrol Fonksiyonu -----------------

def logo_hizli_kombine_karsilastir(zirve_path, logo_portal_paths, hizli_portal_paths):
    """Zirve ile Logo ve Hızlıbilişim Portal dosyalarını kombine karşılaştırır.
    
    Args:
        zirve_path: Zirve Excel dosya yolu
        logo_portal_paths: Logo portal dosya yolları listesi (boş liste olabilir)
        hizli_portal_paths: Hızlıbilişim portal dosya yolları listesi (boş liste olabilir)
    """
    
    try:
        # Zirve dosyasını oku
        df_zirve = pd.read_excel(zirve_path)
        df_zirve['Fatura_No_Temiz'] = df_zirve['Fatura No'].apply(_temizle_fatura_no)
        df_zirve['Tutar_TL'] = pd.to_numeric(df_zirve['Tutar TL'], errors='coerce').fillna(0)
        df_zirve['KDV_TL'] = pd.to_numeric(df_zirve['KDV TL'], errors='coerce').fillna(0)
        # Dövizli faturalar için
        df_zirve['Tutar_Dvz'] = pd.to_numeric(df_zirve.get('Tutar Dvz', 0), errors='coerce').fillna(0)
        df_zirve['KDV_Dvz'] = pd.to_numeric(df_zirve.get('Kdv Dvz', 0), errors='coerce').fillna(0)
        
        # Logo portal dosyalarını normalize et ve birleştir
        all_portal_dfs = []
        
        if logo_portal_paths:
            for path in logo_portal_paths:
                df = pd.read_excel(path)
                df = _logo_normalize_portal(df)
                all_portal_dfs.append(df)
        
        # Hızlıbilişim portal dosyalarını normalize et ve birleştir
        if hizli_portal_paths:
            for path in hizli_portal_paths:
                df = pd.read_excel(path)
                # Otomatik alış/satış tespiti: PayableAmount varsa Gelen/Giden, OdenecekTutar varsa E-Arşiv
                if 'PayableAmount' in df.columns:
                    # Gelen/Giden format
                    fatura_yonu = 'alis' if 'GondericiVkn' in df.columns else 'satis'
                else:
                    # E-Arşiv format (her zaman satış)
                    fatura_yonu = 'satis'
                df = _hizlibilisim_normalize_portal(df, fatura_yonu)
                all_portal_dfs.append(df)
        
        # Tüm portal dosyalarını birleştir
        if not all_portal_dfs:
            return (None, "Hata: En az bir Logo veya Hızlıbilişim portal dosyası seçmelisiniz.")
        
        df_portal = pd.concat(all_portal_dfs, ignore_index=True)
        df_portal['Fatura_No_Temiz'] = df_portal['FaturaNo'].apply(_temizle_fatura_no)
        
        # Set'ler oluştur
        zirve_faturalar = set(df_zirve['Fatura_No_Temiz']) - {''}
        portal_faturalar = set(df_portal['Fatura_No_Temiz']) - {''}
        
        # Fatura no -> satır indeksleri (O(1) arama)
        zirve_satir_bul = _satir_bulucu_olustur(df_zirve)
        portal_satir_bul = _satir_bulucu_olustur(df_portal)
        
        # 1. TUTAR FARKLARI
        tutar_farklari = []
        for fatura_no in zirve_faturalar.intersection(portal_faturalar):
            if not fatura_no:
                continue
            zirve_row = zirve_satir_bul(fatura_no)
            portal_row = portal_satir_bul(fatura_no)
            
            # Para birimine göre Zirve'den doğru sütunları kullan
            para_birimi = str(portal_row.get('Portal_ParaBirimi', 'TRY')).upper()
            if para_birimi != 'TRY' and para_birimi != '':
                # Dövizli fatura
                zirve_tutar = zirve_row['Tutar_Dvz']
                zirve_kdv = zirve_row['KDV_Dvz']
            else:
                # TRY fatura
                zirve_tutar = zirve_row['Tutar_TL']
                zirve_kdv = zirve_row['KDV_TL']
            
            tutar_farki = abs(zirve_tutar - portal_row['Portal_Tutar'])
            kdv_farki = abs(zirve_kdv - portal_row['Portal_KDV'])
            
            if tutar_farki > 1 or kdv_farki > 1:
                tutar_farklari.append({
                    'Fatura No': fatura_no,
                    'Zirve Cari': zirve_row.get('Cari Unvanı', ''),
                    'Portal Cari': portal_row.get('Portal_Cari', ''),
                    'Zirve Tutar': round(zirve_tutar, 2),
                    'Portal Tutar': round(portal_row['Portal_Tutar'], 2),
                    'Tutar Farkı': round(tutar_farki, 2),
                    'Zirve KDV': round(zirve_kdv, 2),
                    'Portal KDV': round(portal_row['Portal_KDV'], 2),
                    'KDV Farkı': round(kdv_farki, 2),
                    'Zirve Tarih': zirve_row.get('Tarih', ''),
                    'Portal Tarih': portal_row.get('FaturaTarihi', ''),
                    'Para Birimi': para_birimi
                })
        
        # 2. EKSİK GİRİŞLER (Portalda var, Zirve'de yok)
        eksik_girisler = []
        for fatura_no in (portal_faturalar - zirve_faturalar):
            if not fatura_no:
                continue
            portal_row = portal_satir_bul(fatura_no)
            eksik_girisler.append({
                'Fatura No': fatura_no,
                'VKN/TCKN': portal_row.get('Portal_VKN', ''),
                'Cari': portal_row.get('Portal_Cari', ''),
                'Ödenecek Tutar': round(portal_row['Portal_Tutar'], 2),
                'Toplam KDV': round(portal_row['Portal_KDV'], 2),
                'Fatura Tarihi': portal_row.get('FaturaTarihi', ''),
                'Para Birimi': portal_row.get('Portal_ParaBirimi', 'TRY'),
                'Fatura Durumu': portal_row.get('Portal_Durum', '')
            })
        
        # 3. FAZLA GİRİŞLER (Zirve'de var, portalda yok)
        fazla_girisler = _fazla_girisler_olustur(
            zirve_satir_bul, zirve_faturalar - portal_faturalar)
        
        # Excel dosyası oluştur
        sonuc_dosyasi = os.path.join(MASAUSTU, "Logo_Zirve_Kontrol_Sonuc.xlsx")
        _kontrol_sonuc_excel_yaz(sonuc_dosyasi, tutar_farklari, eksik_girisler, fazla_girisler)
        
        state.SON_DOSYA_YOLU = sonuc_dosyasi
        
        # Standart mesaj oluştur (diğer kontroller gibi)
        mesaj = _kontrol_ozet_mesaj_olustur(
            logo_portal_paths + hizli_portal_paths,  # Tüm portal dosyaları
            df_portal, 
            df_zirve, 
            tutar_farklari, 
            eksik_girisler, 
            fazla_girisler
        )
        
        return (sonuc_dosyasi, mesaj)
        
    except Exception as e:
        import traceback
        return (None, f"Hata oluştu:\n{str(e)}\n\n{traceback.format_exc()}")
