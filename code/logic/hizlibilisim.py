import os
import pandas as pd

from constants import MASAUSTU
import state
from logic.common import _fazla_girisler_olustur, _kontrol_ozet_mesaj_olustur, _kontrol_sonuc_excel_yaz, _satir_bulucu_olustur, _temizle_fatura_no

# ----------------- Hızlıbilişim Kontrol Fonksiyonu -----------------

def hizlibilisim_karsilastir(zirve_path, portal_paths, fatura_yonu='satis'):
    """Zirve ve Hızlıbilişim Portal Excel dosyalarını karşılaştırır.
    
    Args:
        zirve_path: Zirve Excel dosya yolu
        portal_paths: Portal Excel dosya yolu(ları)
        fatura_yonu: 'alis' veya 'satis'
            - alis: Gelen faturalar (GondericiUnvan/Vkn kullanılır)
            - satis: E-Arşiv + Giden faturalar (AliciUnvan/Vkn kullanılır)
    """
    
    try:
        # Zirve dosyasını oku
        df_zirve = pd.read_excel(zirve_path)
        df_zirve['Fatura_No_Temiz'] = df_zirve['Fatura No'].apply(_temizle_fatura_no)
        df_zirve['Tutar_TL'] = pd.to_numeric(df_zirve['Tutar TL'], errors='coerce').fillna(0)
        df_zirve['KDV_TL'] = pd.to_numeric(df_zirve['KDV TL'], errors='coerce').fillna(0)
        # Dövizli faturalar için Tutar Dvz ve Kdv Dvz sütunlarını oku
        df_zirve['Tutar_Dvz'] = pd.to_numeric(df_zirve.get('Tutar Dvz', 0), errors='coerce').fillna(0)
        df_zirve['KDV_Dvz'] = pd.to_numeric(df_zirve.get('Kdv Dvz', 0), errors='coerce').fillna(0)
        
        # Portal dosyalarını oku ve normalize et
        portal_dfs = []
        if isinstance(portal_paths, list):
            for path in portal_paths:
                df = pd.read_excel(path)
                df = _hizlibilisim_normalize_portal(df, fatura_yonu)
                portal_dfs.append(df)
            df_portal = pd.concat(portal_dfs, ignore_index=True)
        else:
            df_portal = pd.read_excel(portal_paths)
            df_portal = _hizlibilisim_normalize_portal(df_portal, fatura_yonu)
        
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
            
            # ParaBirimi'ne göre Zirve'den doğru sütunları kullan
            para_birimi = str(portal_row.get('Portal_ParaBirimi', 'TRY')).upper()
            if para_birimi != 'TRY' and para_birimi != '':
                # Dövizli fatura - Tutar Dvz ve Kdv Dvz kullan
                zirve_tutar = zirve_row['Tutar_Dvz']
                zirve_kdv = zirve_row['KDV_Dvz']
            else:
                # TRY fatura - Tutar TL ve KDV TL kullan
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
        
        # 2. EKSİK GİRİŞLER
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
                'Oluşturulma Tarihi': portal_row.get('FaturaAlinmaTarihi', ''),
                'Fatura Türü': portal_row.get('FaturaTipi', ''),
                'Fatura Durumu': portal_row.get('Portal_Durum', '')
            })
        
        # 3. FAZLA GİRİŞLER
        fazla_girisler = _fazla_girisler_olustur(
            zirve_satir_bul, zirve_faturalar - portal_faturalar)
        
        # Excel dosyası oluştur
        dosya_adi = "Hizlibilisim_Alis_Kontrol_Sonuc.xlsx" if fatura_yonu == 'alis' else "Hizlibilisim_Satis_Kontrol_Sonuc.xlsx"
        sonuc_dosyasi = os.path.join(MASAUSTU, dosya_adi)
        _kontrol_sonuc_excel_yaz(sonuc_dosyasi, tutar_farklari, eksik_girisler, fazla_girisler)
        
        state.SON_DOSYA_YOLU = sonuc_dosyasi
        mesaj = _kontrol_ozet_mesaj_olustur(portal_paths, df_portal, df_zirve, tutar_farklari, eksik_girisler, fazla_girisler)
        
        return (sonuc_dosyasi, mesaj)
        
    except Exception as e:
        import traceback
        return (None, f"Hata oluştu:\n{str(e)}\n\n{traceback.format_exc()}")

def _hizlibilisim_normalize_portal(df, fatura_yonu='satis'):
    """Hızlıbilişim portal dosyasını normalize et.
    
    Args:
        df: Portal DataFrame
        fatura_yonu: 'alis' veya 'satis'
            - alis: Gelen faturalar -> GondericiUnvan/Vkn kullanılır
            - satis: E-Arşiv + Giden -> AliciUnvan/Vkn kullanılır
    """
    # Tutar belirleme - hangi sütun varsa onu kullan
    if 'OdenecekTutar' in df.columns:
        df['Portal_Tutar'] = pd.to_numeric(df['OdenecekTutar'], errors='coerce').fillna(0)
    elif 'PayableAmount' in df.columns:
        df['Portal_Tutar'] = pd.to_numeric(df['PayableAmount'], errors='coerce').fillna(0)
    else:
        df['Portal_Tutar'] = 0
    
    # KDV hesaplama - her durumda aynı (Kdv20 + Kdv10 + Kdv1)
    kdv20 = pd.to_numeric(df['Kdv20Tutar'], errors='coerce').fillna(0) if 'Kdv20Tutar' in df.columns else pd.Series([0] * len(df))
    kdv10 = pd.to_numeric(df['Kdv10Tutar'], errors='coerce').fillna(0) if 'Kdv10Tutar' in df.columns else pd.Series([0] * len(df))
    kdv1 = pd.to_numeric(df['Kdv1Tutar'], errors='coerce').fillna(0) if 'Kdv1Tutar' in df.columns else pd.Series([0] * len(df))
    df['Portal_KDV'] = kdv20 + kdv10 + kdv1
    
    # Para birimi
    df['Portal_ParaBirimi'] = df['ParaBirimi'] if 'ParaBirimi' in df.columns else 'TRY'
    
    # VKN ve Cari belirleme - fatura yönüne göre
    if fatura_yonu == 'alis':
        # Alış: Gelen faturalar -> GondericiUnvan/Vkn
        df['Portal_VKN'] = df['GondericiVkn'] if 'GondericiVkn' in df.columns else ''
        df['Portal_Cari'] = df['GondericiUnvan'] if 'GondericiUnvan' in df.columns else ''
    else:
        # Satış: Giden faturalar veya E-Arşiv -> AliciUnvan/Vkn
        df['Portal_VKN'] = df['AliciVkn'] if 'AliciVkn' in df.columns else ''
        df['Portal_Cari'] = df['AliciUnvan'] if 'AliciUnvan' in df.columns else ''
    
    # Durum sütunu - hem alış hem satış için
    if 'IptalDurumu' in df.columns:
        df['Portal_Durum'] = df['IptalDurumu']
    elif 'Durum' in df.columns:
        df['Portal_Durum'] = df['Durum']
    else:
        df['Portal_Durum'] = ''
    
    return df

def _logo_normalize_portal(df):
    """Logo portal dosyasını normalize et.
    
    Satış ve alış faturalarını otomatik tespit eder ve standart formata dönüştürür.
    """
    # Otomatik satış/alış tespiti
    # Gönderici VKN varsa alış, Alıcı Vkn/Tckn varsa satış
    is_alis = 'Gönderici VKN' in df.columns or 'Gönderici Adı' in df.columns
    
    # Fatura No
    df['FaturaNo'] = df['Fatura No'] if 'Fatura No' in df.columns else ''
    
    # Tutar ve KDV
    df['Portal_Tutar'] = pd.to_numeric(df['Toplam Tutar'], errors='coerce').fillna(0) if 'Toplam Tutar' in df.columns else 0
    df['Portal_KDV'] = pd.to_numeric(df['KDV Toplamı'], errors='coerce').fillna(0) if 'KDV Toplamı' in df.columns else 0
    
    # Para Birimi
    df['Portal_ParaBirimi'] = df['Para Birimi'] if 'Para Birimi' in df.columns else 'TRY'
    
    # VKN ve Cari (alış/satış'a göre)
    if is_alis:
        # Alış faturası - Gönderici bilgileri
        df['Portal_VKN'] = df['Gönderici VKN'] if 'Gönderici VKN' in df.columns else ''
        df['Portal_Cari'] = df['Gönderici Adı'] if 'Gönderici Adı' in df.columns else ''
    else:
        # Satış faturası - Alıcı bilgileri
        df['Portal_VKN'] = df['Alıcı Vkn/Tckn'] if 'Alıcı Vkn/Tckn' in df.columns else ''
        df['Portal_Cari'] = df['Alıcı Ünvanı'] if 'Alıcı Ünvanı' in df.columns else ''
    
    # Tarih
    df['FaturaTarihi'] = df['Fatura Tarihi'] if 'Fatura Tarihi' in df.columns else ''
    
    # Durum - Öncelik sırası: İptal/İtiraz Durumu > Durumu > Durum
    df['Portal_Durum'] = ''
    if 'İptal/İtiraz Durumu' in df.columns:
        # İptal/İtiraz Durumu varsa önce ona bak
        df['Portal_Durum'] = df['İptal/İtiraz Durumu'].fillna('')
        # Boş olanlar için diğer sütunlara bak
        if 'Durumu' in df.columns:
            mask = (df['Portal_Durum'] == '') | (df['Portal_Durum'].isna())
            df.loc[mask, 'Portal_Durum'] = df.loc[mask, 'Durumu'].fillna('')
        elif 'Durum' in df.columns:
            mask = (df['Portal_Durum'] == '') | (df['Portal_Durum'].isna())
            df.loc[mask, 'Portal_Durum'] = df.loc[mask, 'Durum'].fillna('')
    elif 'Durumu' in df.columns:
        df['Portal_Durum'] = df['Durumu'].fillna('')
    elif 'Durum' in df.columns:
        df['Portal_Durum'] = df['Durum'].fillna('')
    
    return df

def _izibiz_normalize_portal(df):
    """İzibiz portal dosyasını normalize et.
    
    Alış ve satış faturalarını otomatik tespit eder:
    - Alış: Fatura No + VKN/TCKN + Unvan → Gönderici bilgisi
    - Satış: Belge No + Gönderici Unvan + Alıcı Unvan → Alıcı bilgisi
    """
    # Otomatik alış/satış tespiti
    is_alis = ('VKN/TCKN' in df.columns and 'Unvan' in df.columns 
               and 'Fatura No' in df.columns and 'Belge No' not in df.columns)
    
    if is_alis:
        # ALIŞ FATURASI (GELEN)
        df['Fatura_No_Temiz'] = df['Fatura No']
        df['Portal_Tutar'] = pd.to_numeric(df['Tutar'], errors='coerce').fillna(0)
        df['Portal_KDV'] = 0  # Alış faturalarında ayrı KDV sütunu yok, tutara dahil
        df['Portal_VKN'] = df['VKN/TCKN']
        df['Portal_Cari'] = df['Unvan']
        df['Portal_Tarih'] = df['Tarih'] if 'Tarih' in df.columns else ''
        df['Portal_Durum'] = df['Durum'] if 'Durum' in df.columns else ''
    else:
        # SATIŞ FATURASI (GİDEN)
        df['Fatura_No_Temiz'] = df['Belge No']
        df['Portal_Tutar'] = pd.to_numeric(df['Ödenecek Tutar'], errors='coerce').fillna(0)
        df['Portal_KDV'] = pd.to_numeric(df['Toplam KDV'], errors='coerce').fillna(0)
        # Satış faturasında ALICI bilgisi kullanılmalı (müşteri)
        df['Portal_VKN'] = df.get('Alıcı VKN/TCKN', '')
        df['Portal_Cari'] = df.get('Alıcı Unvan', '')
        df['Portal_Tarih'] = df['Tarih'] if 'Tarih' in df.columns else ''
        df['Portal_Durum'] = df['Durum'] if 'Durum' in df.columns else ''
    
    return df
