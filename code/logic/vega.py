import os
import pandas as pd

from constants import MASAUSTU
import state
from logic.common import _excel_dosyalari_oku, _fazla_girisler_olustur, _kontrol_ozet_mesaj_olustur, _kontrol_sonuc_excel_yaz, _satir_bulucu_olustur, _temizle_fatura_no

# ----------------- Vega Kontrol Fonksiyonu -----------------

def vega_karsilastir(zirve_path, portal_paths):
    """Zirve ve Vega Portal Excel dosyalarını karşılaştırır."""
    
    try:
        df_zirve, df_portal = _excel_dosyalari_oku(zirve_path, portal_paths)
        
        # Tutarları sayıya çevir (Vega'ya özgü format)
        def parse_tutar(value):
            if pd.isna(value):
                return 0
            if isinstance(value, (int, float)):
                return float(value)
            return float(str(value).replace('.', '').replace(',', '.'))
        
        df_zirve['Fatura_No_Temiz'] = df_zirve['Fatura No'].apply(_temizle_fatura_no)
        df_portal['Fatura_No_Temiz'] = df_portal['Fatura No'].apply(_temizle_fatura_no)
        
        df_zirve['Toplam_TL'] = df_zirve['Tutar TL'].apply(parse_tutar)
        df_zirve['KDV_TL'] = df_zirve['KDV TL'].apply(parse_tutar)
        
        # Portal tutarları - farklı sütun isimleri olabilir
        def get_portal_toplam(row):
            if 'Ödenecek' in df_portal.columns and pd.notna(row.get('Ödenecek')):
                return parse_tutar(row['Ödenecek'])
            elif 'Toplam' in df_portal.columns and pd.notna(row.get('Toplam')):
                return parse_tutar(row['Toplam'])
            return 0
        
        df_portal['Portal_KDV'] = df_portal['Vergi'].apply(parse_tutar)
        df_portal['Portal_Toplam'] = df_portal.apply(get_portal_toplam, axis=1)
        
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
            
            toplam_farki = abs(zirve_row['Toplam_TL'] - portal_row['Portal_Toplam'])
            kdv_farki = abs(zirve_row['KDV_TL'] - portal_row['Portal_KDV'])
            
            if toplam_farki > 1 or kdv_farki > 1:
                tutar_farklari.append({
                    'Fatura No': fatura_no,
                    'Zirve Cari': zirve_row.get('Cari Unvanı', ''),
                    'Portal Cari': portal_row.get('Unvan', ''),
                    'Zirve Tutar': round(zirve_row['Toplam_TL'], 2),
                    'Portal Tutar': round(portal_row['Portal_Toplam'], 2),
                    'Tutar Farkı': round(toplam_farki, 2),
                    'Zirve KDV': round(zirve_row['KDV_TL'], 2),
                    'Portal KDV': round(portal_row['Portal_KDV'], 2),
                    'KDV Farkı': round(kdv_farki, 2),
                    'Zirve Tarih': zirve_row.get('Tarih', ''),
                    'Portal Tarih': portal_row.get('Tarih', '')
                })
        
        # 2. EKSİK GİRİŞLER
        # Vega tarih formatını parse et (örn: '2025.12.0101.12.2025' -> '01.12.2025')
        def parse_vega_tarih(tarih_str):
            if pd.isna(tarih_str) or tarih_str == '':
                return ''
            tarih_str = str(tarih_str)
            if len(tarih_str) >= 10 and '.' in tarih_str:
                parts = tarih_str.split('.')
                if len(parts) >= 4:
                    try:
                        return f"{parts[-3][-2:]}.{parts[-2]}.{parts[-1]}"
                    except:
                        pass
            return tarih_str
        
        eksik_girisler = []
        for fatura_no in (portal_faturalar - zirve_faturalar):
            if not fatura_no:
                continue
            portal_row = portal_satir_bul(fatura_no)
            # Fatura Türü: Türü veya Fatura Türü sütunlarından al
            fatura_turu = ''
            if 'Türü' in portal_row.index and pd.notna(portal_row['Türü']):
                fatura_turu = str(portal_row['Türü'])
            elif 'Fatura Türü' in portal_row.index and pd.notna(portal_row['Fatura Türü']):
                fatura_turu = str(portal_row['Fatura Türü'])
            
            # Fatura Durumu: Kabul/Red, Red/Kabul veya Rapor Durumu sütunlarından al
            fatura_durumu = ''
            if 'Kabul/Red' in portal_row.index and pd.notna(portal_row['Kabul/Red']):
                fatura_durumu = str(portal_row['Kabul/Red'])
            elif 'Red/Kabul' in portal_row.index and pd.notna(portal_row['Red/Kabul']):
                fatura_durumu = str(portal_row['Red/Kabul'])
            elif 'Rapor Durumu' in portal_row.index and pd.notna(portal_row['Rapor Durumu']):
                fatura_durumu = str(portal_row['Rapor Durumu'])
            
            eksik_girisler.append({
                'Fatura No': fatura_no,
                'VKN/TCKN': portal_row.get('VKN/TCKN', ''),
                'Unvan': portal_row.get('Unvan', ''),
                'Ödenecek Tutar': round(portal_row['Portal_Toplam'], 2),
                'Toplam KDV': round(portal_row['Portal_KDV'], 2),
                'Fatura Tarihi': parse_vega_tarih(portal_row.get('Tarih', '')),
                'Oluşturulma Tarihi': portal_row.get('Oluşturulma Tarihi', portal_row.get('Kayıt Tarihi', '')),
                'Fatura Türü': fatura_turu,
                'Fatura Durumu': fatura_durumu
            })
        
        # 3. FAZLA GİRİŞLER
        fazla_girisler = _fazla_girisler_olustur(
            zirve_satir_bul, zirve_faturalar - portal_faturalar, tutar_sutun='Toplam_TL')
        
        # Excel dosyası oluştur
        sonuc_dosyasi = os.path.join(MASAUSTU, "Vega_Kontrol_Sonuc.xlsx")
        _kontrol_sonuc_excel_yaz(sonuc_dosyasi, tutar_farklari, eksik_girisler, fazla_girisler)
        
        state.SON_DOSYA_YOLU = sonuc_dosyasi
        mesaj = _kontrol_ozet_mesaj_olustur(portal_paths, df_portal, df_zirve, tutar_farklari, eksik_girisler, fazla_girisler)
        
        return (sonuc_dosyasi, mesaj)
        
    except Exception as e:
        import traceback
        return (None, f"Hata oluştu:\n{str(e)}\n\n{traceback.format_exc()}")
