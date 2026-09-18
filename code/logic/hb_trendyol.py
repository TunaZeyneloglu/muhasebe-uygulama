import os
import pandas as pd

from constants import MASAUSTU
import state
from logic.common import _fazla_girisler_olustur, _kontrol_ozet_mesaj_olustur, _kontrol_sonuc_excel_yaz, _satir_bulucu_olustur, _temizle_fatura_no

# ----------------- Portal Tipi Tespit Fonksiyonu -----------------

def _tespit_portal_tipi(df):
    """Portal Excel dosyasının tipini tespit eder (Hepsiburada veya Trendyol).
    
    Args:
        df: Portal DataFrame
        
    Returns:
        'trendyol' veya 'hepsiburada'
    """
    sutunlar = df.columns.tolist()
    
    # Trendyol tespiti: "Gönderici Ünvan" VE "KDV Toplamı" varsa
    if 'Gönderici Ünvan' in sutunlar and 'KDV Toplamı' in sutunlar:
        return 'trendyol'
    # Hepsiburada tespiti: ("Gönderici" VEYA "Alıcı") VE "Toplam KDV" varsa
    elif ('Gönderici' in sutunlar or 'Alıcı' in sutunlar) and 'Toplam KDV' in sutunlar:
        return 'hepsiburada'
    else:
        return 'bilinmiyor'

# ----------------- Hepsiburada/Trendyol Kontrol Fonksiyonu -----------------

def hb_ve_trn_karsilastir(zirve_path, portal_paths):
    """Zirve ve Hepsiburada/Trendyol Portal Excel dosyalarını karşılaştırır.
    Portal tipi ve Gelen/Giden otomatik tespit edilir.
    
    Args:
        zirve_path: Zirve Excel dosya yolu
        portal_paths: Portal Excel dosya yolu(ları)
    """
    
    try:
        # Zirve dosyasını oku
        df_zirve = pd.read_excel(zirve_path)
        df_zirve['Fatura_No_Temiz'] = df_zirve['Fatura No'].apply(_temizle_fatura_no)
        df_zirve['Tutar_TL'] = pd.to_numeric(df_zirve['Tutar TL'], errors='coerce').fillna(0)
        df_zirve['KDV_TL'] = pd.to_numeric(df_zirve['KDV TL'], errors='coerce').fillna(0)
        
        # Portal dosyalarını oku ve normalize et (otomatik tespit)
        portal_dfs = []
        if isinstance(portal_paths, list):
            for path in portal_paths:
                df = pd.read_excel(path)
                portal_tipi = _tespit_portal_tipi(df)
                if portal_tipi == 'trendyol':
                    df = _trendyol_normalize_portal(df)
                elif portal_tipi == 'hepsiburada':
                    df = _hepsiburada_normalize_portal(df)
                else:
                    raise ValueError(f"Bilinmeyen portal tipi: {portal_tipi}")
                portal_dfs.append(df)
            df_portal = pd.concat(portal_dfs, ignore_index=True)
        else:
            df_portal = pd.read_excel(portal_paths)
            portal_tipi = _tespit_portal_tipi(df_portal)
            if portal_tipi == 'trendyol':
                df_portal = _trendyol_normalize_portal(df_portal)
            elif portal_tipi == 'hepsiburada':
                df_portal = _hepsiburada_normalize_portal(df_portal)
            else:
                raise ValueError(f"Bilinmeyen portal tipi: {portal_tipi}")
        
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
            
            tutar_farki = abs(zirve_row['Tutar_TL'] - portal_row['Portal_Tutar'])
            kdv_farki = abs(zirve_row['KDV_TL'] - portal_row['Portal_KDV'])
            
            if tutar_farki > 1 or kdv_farki > 1:
                tutar_farklari.append({
                    'Fatura No': fatura_no,
                    'Portal Kaynak': portal_row.get('Portal_Kaynak', ''),
                    'Zirve Cari': zirve_row.get('Cari Unvanı', ''),
                    'Portal Gönderici': portal_row.get('Portal_Gonderici', ''),
                    'Portal Alıcı': portal_row.get('Portal_Alici', ''),
                    'Zirve Tutar': round(zirve_row['Tutar_TL'], 2),
                    'Portal Tutar': round(portal_row['Portal_Tutar'], 2),
                    'Tutar Farkı': round(tutar_farki, 2),
                    'Zirve KDV': round(zirve_row['KDV_TL'], 2),
                    'Portal KDV': round(portal_row['Portal_KDV'], 2),
                    'KDV Farkı': round(kdv_farki, 2),
                    'Zirve Tarih': zirve_row.get('Tarih', ''),
                    'Portal Tarih': portal_row.get('FaturaTarihi', '')
                })
        
        # 2. EKSİK GİRİŞLER
        eksik_girisler = []
        for fatura_no in (portal_faturalar - zirve_faturalar):
            if not fatura_no:
                continue
            portal_row = portal_satir_bul(fatura_no)
            eksik_girisler.append({
                'Fatura No': fatura_no,
                'Portal Kaynak': portal_row.get('Portal_Kaynak', ''),
                'Gönderici': portal_row.get('Portal_Gonderici', ''),
                'Alıcı': portal_row.get('Portal_Alici', ''),
                'VKN/TCKN': portal_row.get('Portal_VKN', ''),
                'Ödenecek Tutar': round(portal_row['Portal_Tutar'], 2),
                'Toplam KDV': round(portal_row['Portal_KDV'], 2),
                'Fatura Tarihi': portal_row.get('FaturaTarihi', ''),
                'Oluşturulma Tarihi': portal_row.get('OlusturmaTarihi', ''),
                'Fatura Türü': portal_row.get('FaturaTipi', ''),
                'Fatura Durumu': portal_row.get('Portal_Durum', '')
            })
        
        # 3. FAZLA GİRİŞLER
        fazla_girisler = _fazla_girisler_olustur(
            zirve_satir_bul, zirve_faturalar - portal_faturalar)
        
        # Excel dosyası oluştur
        sonuc_dosyasi = os.path.join(MASAUSTU, "HB_TRN_Kontrol_Sonuc.xlsx")
        _kontrol_sonuc_excel_yaz(sonuc_dosyasi, tutar_farklari, eksik_girisler, fazla_girisler)
        
        state.SON_DOSYA_YOLU = sonuc_dosyasi
        mesaj = _kontrol_ozet_mesaj_olustur(portal_paths, df_portal, df_zirve, tutar_farklari, eksik_girisler, fazla_girisler)
        
        return (sonuc_dosyasi, mesaj)
        
    except Exception as e:
        import traceback
        return (None, f"Hata oluştu:\n{str(e)}\n\n{traceback.format_exc()}")

def _hepsiburada_normalize_portal(df):
    """Hepsiburada portal dosyasını normalize et.
    Gelen (Gönderici) veya Giden (Alıcı) otomatik tespit edilir.
    
    Args:
        df: Portal DataFrame
    """
    # Portal kaynağı
    df['Portal_Kaynak'] = 'Hepsiburada'
    
    # Fatura No
    df['FaturaNo'] = df['Fatura No'] if 'Fatura No' in df.columns else ''
    
    # Tutarlar
    df['Portal_Tutar'] = pd.to_numeric(df['Ödenecek Tutar'], errors='coerce').fillna(0)
    df['Portal_KDV'] = pd.to_numeric(df['Toplam KDV'], errors='coerce').fillna(0)
    
    # Tarihler
    df['FaturaTarihi'] = df['Fatura Tarihi'] if 'Fatura Tarihi' in df.columns else ''
    df['OlusturmaTarihi'] = df['Oluşturulma Tarihi'] if 'Oluşturulma Tarihi' in df.columns else ''
    
    # Fatura bilgileri
    df['FaturaTipi'] = df['Fatura Tipi'] if 'Fatura Tipi' in df.columns else ''
    df['Portal_Durum'] = df['Fatura Durumu'] if 'Fatura Durumu' in df.columns else ''
    
    # Cari bilgileri - Otomatik tespit (Gönderici varsa Gelen, Alıcı varsa Giden)
    if 'Gönderici' in df.columns:
        # Gelen faturalar -> Gönderici
        df['Portal_VKN'] = df['Gönderici VKN/TCKN'] if 'Gönderici VKN/TCKN' in df.columns else ''
        df['Portal_Gonderici'] = df['Gönderici']
        df['Portal_Alici'] = ''
        df['Portal_Cari'] = df['Gönderici']  # Geriye uyumluluk için
    elif 'Alıcı' in df.columns:
        # Giden faturalar -> Alıcı
        df['Portal_VKN'] = df['Alıcı VKN/TCKN'] if 'Alıcı VKN/TCKN' in df.columns else ''
        df['Portal_Gonderici'] = ''
        df['Portal_Alici'] = df['Alıcı']
        df['Portal_Cari'] = df['Alıcı']  # Geriye uyumluluk için
    else:
        # Ne Gönderici ne de Alıcı sütunu yok
        df['Portal_VKN'] = ''
        df['Portal_Gonderici'] = ''
        df['Portal_Alici'] = ''
        df['Portal_Cari'] = ''
    
    return df

def _trendyol_normalize_portal(df):
    """Trendyol portal dosyasını normalize et.
    İki farklı format desteklenir:
    - Format 1: Durum, Zarf UUID
    - Format 2: Statü, Rapor Durumu
    
    Args:
        df: Portal DataFrame
    """
    # Portal kaynağı
    df['Portal_Kaynak'] = 'Trendyol'
    
    # Fatura No
    df['FaturaNo'] = df['Fatura No'] if 'Fatura No' in df.columns else ''
    
    # Tutarlar
    df['Portal_Tutar'] = pd.to_numeric(df['Ödenecek Tutar'], errors='coerce').fillna(0)
    df['Portal_KDV'] = pd.to_numeric(df['KDV Toplamı'], errors='coerce').fillna(0)
    
    # Tarihler
    df['FaturaTarihi'] = df['Fatura Tarihi'] if 'Fatura Tarihi' in df.columns else ''
    df['OlusturmaTarihi'] = df['Oluşturulma Tarihi'] if 'Oluşturulma Tarihi' in df.columns else ''
    
    # Fatura bilgileri
    df['FaturaTipi'] = df['Fatura Tipi'] if 'Fatura Tipi' in df.columns else ''
    
    # Fatura durumu - iki farklı format
    if 'Durum' in df.columns:
        df['Portal_Durum'] = df['Durum']
    elif 'Statü' in df.columns:
        df['Portal_Durum'] = df['Statü']
    else:
        df['Portal_Durum'] = ''
    
    # Cari bilgileri - Hem Gönderici hem Alıcı var
    df['Portal_Gonderici'] = df['Gönderici Ünvan'] if 'Gönderici Ünvan' in df.columns else ''
    df['Portal_Alici'] = df['Alıcı Ünvan'] if 'Alıcı Ünvan' in df.columns else ''
    
    # VKN bilgisi
    if 'Gönderici VKN/TCKN' in df.columns:
        df['Portal_VKN'] = df['Gönderici VKN/TCKN']
    elif 'Alıcı VKN/TCKN' in df.columns:
        df['Portal_VKN'] = df['Alıcı VKN/TCKN']
    else:
        df['Portal_VKN'] = ''
    
    # Portal_Cari - Gönderici veya Alıcı (hangisi dolu ise)
    df['Portal_Cari'] = df['Portal_Gonderici'].where(df['Portal_Gonderici'] != '', df['Portal_Alici'])
    
    return df
