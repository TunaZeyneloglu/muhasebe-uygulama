import os
import pandas as pd

from constants import MASAUSTU
import state
from logic.common import _excel_dosyalari_oku, _fazla_girisler_olustur, _kontrol_ozet_mesaj_olustur, _kontrol_sonuc_excel_yaz, _satir_bulucu_olustur, _temizle_fatura_no

# ----------------- Uyumsoft Kontrol Fonksiyonları -----------------

def uyumsoft_zirve_karsilastir(zirve_path, portal_paths):
    """Zirve ve Uyumsoft Portal Excel dosyalarını karşılaştırır (Zirve programı için)."""
    
    try:
        df_zirve, df_portal = _excel_dosyalari_oku(zirve_path, portal_paths)
        
        df_zirve['Fatura_No_Temiz'] = df_zirve['Fatura No'].apply(_temizle_fatura_no)
        df_portal['Fatura_No_Temiz'] = df_portal['Fatura No'].apply(_temizle_fatura_no)
        
        # Tutarları float'a çevir
        df_zirve['Tutar_TL'] = pd.to_numeric(df_zirve['Tutar TL'], errors='coerce').fillna(0)
        df_zirve['KDV_TL'] = pd.to_numeric(df_zirve['KDV TL'], errors='coerce').fillna(0)
        df_portal['Odenecek_Tutar'] = pd.to_numeric(df_portal['Ödenecek Tutar'], errors='coerce').fillna(0)
        df_portal['Toplam_KDV'] = pd.to_numeric(df_portal['Toplam KDV'], errors='coerce').fillna(0)
        
        # Set'ler oluştur
        zirve_faturalar = set(df_zirve['Fatura_No_Temiz'])
        portal_faturalar = set(df_portal['Fatura_No_Temiz'])
        
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
            
            tutar_farki = abs(zirve_row['Tutar_TL'] - portal_row['Odenecek_Tutar'])
            kdv_farki = abs(zirve_row['KDV_TL'] - portal_row['Toplam_KDV'])
            
            if tutar_farki > 1 or kdv_farki > 1:
                tutar_farklari.append({
                    'Fatura No': fatura_no,
                    'Zirve Cari': zirve_row.get('Cari Unvanı', ''),
                    'Portal Cari': portal_row.get('Gönderici', ''),
                    'Zirve Tutar': round(zirve_row['Tutar_TL'], 2),
                    'Portal Tutar': round(portal_row['Odenecek_Tutar'], 2),
                    'Tutar Farkı': round(tutar_farki, 2),
                    'Zirve KDV': round(zirve_row['KDV_TL'], 2),
                    'Portal KDV': round(portal_row['Toplam_KDV'], 2),
                    'KDV Farkı': round(kdv_farki, 2),
                    'Zirve Tarih': zirve_row.get('Tarih', ''),
                    'Portal Tarih': portal_row.get('Fatura Tarihi', '')
                })
        
        # 2. EKSİK GİRİŞLER
        eksik_girisler = []
        for fatura_no in (portal_faturalar - zirve_faturalar):
            if not fatura_no:
                continue
            portal_row = portal_satir_bul(fatura_no)
            eksik_girisler.append({
                'Fatura No': fatura_no,
                'Gönderici VKN/TCKN': portal_row.get('Gönderici VKN/TCKN', ''),
                'Gönderici': portal_row.get('Gönderici', ''),
                'Ödenecek Tutar': round(portal_row['Odenecek_Tutar'], 2),
                'Toplam KDV': round(portal_row['Toplam_KDV'], 2),
                'Fatura Tarihi': portal_row.get('Fatura Tarihi', ''),
                'Oluşturulma Tarihi': portal_row.get('Oluşturulma Tarihi', ''),
                'Fatura Türü': portal_row.get('Fatura Tipi', ''),
                'Fatura Durumu': portal_row.get('Fatura Durumu', '')
            })
        
        # 3. FAZLA GİRİŞLER
        fazla_girisler = _fazla_girisler_olustur(
            zirve_satir_bul, zirve_faturalar - portal_faturalar)
        
        # Excel dosyası oluştur
        sonuc_dosyasi = os.path.join(MASAUSTU, "Uyumsoft_Zirve_Kontrol_Sonuc.xlsx")
        _kontrol_sonuc_excel_yaz(sonuc_dosyasi, tutar_farklari, eksik_girisler, fazla_girisler)
        
        state.SON_DOSYA_YOLU = sonuc_dosyasi
        mesaj = _kontrol_ozet_mesaj_olustur(portal_paths, df_portal, df_zirve, tutar_farklari, eksik_girisler, fazla_girisler)
        
        return (sonuc_dosyasi, mesaj)
        
    except Exception as e:
        return (None, f"Hata oluştu:\n{str(e)}")

# Verim Excel sütun isimleri (A'dan N'ye)
VERIM_SUTUNLAR = ['FATURA TARİH', 'FATURA NO', 'HESAP KODU', 'HESAP ADI', 'TOPLAM', 'İND', 'ARA TOPLAM', '%1', '% 8', '% 10', '% 18', '% 20', 'NET TOPLAM', 'PB']

def _verim_excel_oku(path):
    """Verim Excel dosyasını oku ve sütun isimlerini zorla ata."""
    # Dosyayı header olmadan oku
    df = pd.read_excel(path, header=None)
    
    # İlk satıra sütun isimlerini ata (A'dan N'ye kadar 14 sütun)
    if len(df.columns) >= 14:
        df.columns = VERIM_SUTUNLAR + list(df.columns[14:])
    else:
        df.columns = VERIM_SUTUNLAR[:len(df.columns)]
    
    # İlk satırı sil (çünkü orada eski sütun isimleri var)
    df = df.iloc[1:].reset_index(drop=True)
    
    return df

# Geçersiz fatura numaraları (Export özet satırları)
GECERSIZ_FATURA_NO = ['BAŞLAMA:', 'BELGE SAYISI:', 'BİTİŞ:', 'BITIŞ:', 'SÜRE:']

def _gecersiz_fatura_no_mu(fatura_no):
    """Fatura numarasının geçersiz (export özet satırı) olup olmadığını kontrol et."""
    if not fatura_no:
        return True
    fatura_str = str(fatura_no).upper()
    for gecersiz in GECERSIZ_FATURA_NO:
        if gecersiz.upper() in fatura_str:
            return True
    return False

def _verim_normalize(df):
    """Verim Excel dosyasını normalize et - sütun adlarını standartlaştır."""
    df_normalized = df.copy()
    
    # FATURA NO -> Fatura No
    if 'FATURA NO' in df.columns:
        df_normalized['Fatura No'] = df['FATURA NO']
    
    # NET TOPLAM -> Tutar TL
    if 'NET TOPLAM' in df.columns:
        df_normalized['Tutar TL'] = pd.to_numeric(df['NET TOPLAM'], errors='coerce').fillna(0)
    
    # KDV Toplamı (%1 + % 8 + % 10 + % 18 + % 20)
    kdv_sum = 0
    for col in ['%1', '% 8', '% 10', '% 18', '% 20']:
        if col in df.columns:
            kdv_sum = kdv_sum + pd.to_numeric(df[col], errors='coerce').fillna(0)
    df_normalized['KDV TL'] = kdv_sum
    
    # HESAP ADI -> Cari Unvanı
    if 'HESAP ADI' in df.columns:
        df_normalized['Cari Unvanı'] = df['HESAP ADI']
    
    # FATURA TARİH -> Tarih
    if 'FATURA TARİH' in df.columns:
        df_normalized['Tarih'] = df['FATURA TARİH']
    
    # HESAP KODU -> E.Kod
    if 'HESAP KODU' in df.columns:
        df_normalized['E.Kod'] = df['HESAP KODU']
    
    return df_normalized

def uyumsoft_verim_karsilastir(verim_paths, portal_paths):
    """Verim ve Uyumsoft Portal Excel dosyalarını karşılaştırır (Verim programı için)."""
    
    try:
        # Verim dosyalarını oku ve normalize et
        if isinstance(verim_paths, list):
            verim_dfs = [_verim_normalize(_verim_excel_oku(path)) for path in verim_paths]
            df_verim = pd.concat(verim_dfs, ignore_index=True)
        else:
            df_verim = _verim_normalize(_verim_excel_oku(verim_paths))
        
        # Portal (Uyumsoft) dosyalarını oku
        if isinstance(portal_paths, list):
            portal_dfs = [pd.read_excel(path) for path in portal_paths]
            df_portal = pd.concat(portal_dfs, ignore_index=True)
        else:
            df_portal = pd.read_excel(portal_paths)
        
        df_verim['Fatura_No_Temiz'] = df_verim['Fatura No'].apply(_temizle_fatura_no)
        df_portal['Fatura_No_Temiz'] = df_portal['Fatura No'].apply(_temizle_fatura_no)
        
        # Tutarları float'a çevir
        df_verim['Tutar_TL'] = pd.to_numeric(df_verim['Tutar TL'], errors='coerce').fillna(0)
        df_verim['KDV_TL'] = pd.to_numeric(df_verim['KDV TL'], errors='coerce').fillna(0)
        df_portal['Odenecek_Tutar'] = pd.to_numeric(df_portal['Ödenecek Miktar'], errors='coerce').fillna(0)
        
        # Portal KDV Toplamı = Toplam Kdv 1 + Toplam Kdv 8 + Toplam Kdv 18 + Toplam KDV %10 Tutar + Toplam KDV %20 Tutar
        kdv_cols = ['Toplam Kdv 1', 'Toplam Kdv 8', 'Toplam Kdv 18', 'Toplam KDV %10 Tutar', 'Toplam KDV %20 Tutar']
        df_portal['Toplam_KDV'] = 0
        for col in kdv_cols:
            if col in df_portal.columns:
                df_portal['Toplam_KDV'] = df_portal['Toplam_KDV'] + pd.to_numeric(df_portal[col], errors='coerce').fillna(0)
        
        # Set'ler oluştur
        verim_faturalar = set(df_verim['Fatura_No_Temiz'])
        portal_faturalar = set(df_portal['Fatura_No_Temiz'])
        
        # Fatura no -> satır indeksleri (O(1) arama)
        verim_satir_bul = _satir_bulucu_olustur(df_verim)
        portal_satir_bul = _satir_bulucu_olustur(df_portal)
        
        # 1. TUTAR FARKLARI
        tutar_farklari = []
        for fatura_no in verim_faturalar.intersection(portal_faturalar):
            if not fatura_no:
                continue
            verim_row = verim_satir_bul(fatura_no)
            portal_row = portal_satir_bul(fatura_no)
            
            tutar_farki = abs(verim_row['Tutar_TL'] - portal_row['Odenecek_Tutar'])
            kdv_farki = abs(verim_row['KDV_TL'] - portal_row['Toplam_KDV'])
            
            if tutar_farki > 1 or kdv_farki > 1:
                # Portal tarih ve saat formatını ayır
                portal_datetime = portal_row.get('Fatura Oluşturma Tarihi', '')
                portal_tarih = ''
                portal_saat = ''
                if pd.notna(portal_datetime) and portal_datetime:
                    try:
                        dt = pd.to_datetime(portal_datetime)
                        portal_tarih = dt.strftime('%d.%m.%Y')
                        portal_saat = dt.strftime('%H:%M:%S')
                    except:
                        portal_tarih = str(portal_datetime)
                
                tutar_farklari.append({
                    'Fatura No': fatura_no,
                    'Verim Cari': verim_row.get('Cari Unvanı', ''),
                    'Portal Cari': portal_row.get('Cari Adı', ''),
                    'Verim Tutar': round(verim_row['Tutar_TL'], 2),
                    'Portal Tutar': round(portal_row['Odenecek_Tutar'], 2),
                    'Tutar Farkı': round(tutar_farki, 2),
                    'Verim KDV': round(verim_row['KDV_TL'], 2),
                    'Portal KDV': round(portal_row['Toplam_KDV'], 2),
                    'KDV Farkı': round(kdv_farki, 2),
                    'Verim Tarih': verim_row.get('Tarih', ''),
                    'Portal Tarih': portal_tarih,
                    'Portal Saat': portal_saat
                })
        
        # 2. EKSİK GİRİŞLER
        eksik_girisler = []
        for fatura_no in (portal_faturalar - verim_faturalar):
            if not fatura_no:
                continue
            # Geçersiz fatura numaralarını atla (export özet satırları)
            if _gecersiz_fatura_no_mu(fatura_no):
                continue
            portal_row = portal_satir_bul(fatura_no)
            eksik_girisler.append({
                'Fatura No': fatura_no,
                'Cari VKN/TCKN': portal_row.get('Cari Vkn/Tckn', ''),
                'Cari Adı': portal_row.get('Cari Adı', ''),
                'Ödenecek Tutar': round(portal_row['Odenecek_Tutar'], 2),
                'Toplam KDV': round(portal_row['Toplam_KDV'], 2),
                'Fatura Tarihi': portal_row.get('Fatura Tarihi', ''),
                'Oluşturulma Tarihi': portal_row.get('Fatura Oluşturma Tarihi', ''),
                'Fatura Türü': portal_row.get('Tür', ''),
                'Fatura Durumu': portal_row.get('Durum', '')
            })
        
        # 3. FAZLA GİRİŞLER
        fazla_girisler = []
        for fatura_no in (verim_faturalar - portal_faturalar):
            if not fatura_no:
                continue
            verim_row = verim_satir_bul(fatura_no)
            fazla_girisler.append({
                'Fatura No': fatura_no,
                'Cari Unvanı': verim_row.get('Cari Unvanı', ''),
                'Tutar TL': round(verim_row['Tutar_TL'], 2),
                'KDV TL': round(verim_row['KDV_TL'], 2),
                'Tarih': verim_row.get('Tarih', ''),
                'E.Kod': verim_row.get('E.Kod', '')
            })
        
        # Excel dosyası oluştur
        sonuc_dosyasi = os.path.join(MASAUSTU, "Uyumsoft_Verim_Kontrol_Sonuc.xlsx")
        _kontrol_sonuc_excel_yaz(sonuc_dosyasi, tutar_farklari, eksik_girisler, fazla_girisler)
        
        state.SON_DOSYA_YOLU = sonuc_dosyasi
        mesaj = _kontrol_ozet_mesaj_olustur(portal_paths, df_portal, df_verim, tutar_farklari, eksik_girisler, fazla_girisler)
        
        return (sonuc_dosyasi, mesaj)
        
    except Exception as e:
        return (None, f"Hata oluştu:\n{str(e)}")
