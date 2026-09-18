# Sürüm Yayınlama

## Yeni sürüm nasıl yayınlanır

1. `code/version.py` içindeki sürümü artırın (satır biçimi tam olarak `APP_VERSION = "X.Y.Z"` olmalı).
2. Değişikliği commit edip push edin.
3. Aynı sürümle tag oluşturup gönderin:
   ```
   git tag vX.Y.Z && git push origin vX.Y.Z
   ```
4. GitHub Actions (`.github/workflows/release.yml`) tag ile `code/version.py` uyuşuyor mu kontrol eder, uyuşmuyorsa durur. Uyuşuyorsa `FaturaYonetimSistemi.exe`'yi konsolsuz, tek dosya olarak derler ve `vX.Y.Z` Release'ine ekler.
5. Açık uygulamalar yeni sürümü arka planda (GitHub Releases `latest`) bulur. Draft ve prerelease sürümler dikkate alınmaz.

## Yerelde derleme

```
pip install -r requirements.txt pyinstaller && pyinstaller FaturaYonetimSistemi.spec
```

Çıktı: `dist/FaturaYonetimSistemi.exe`

Windows'ta tek komutla (sanal ortamı `.venv` olarak kurar, derler, exe'nin boyutunu ve SHA256'sını yazdırır; `-Temiz` önce `build`/`dist`'i siler):

```
powershell -ExecutionPolicy Bypass -File .\derle.ps1
```

## Notlar

- Bu sistemden önceki (konsollu) exe'ler kendini güncelleyemez. Yeni exe kullanıcılara bir kez elle verilmelidir.
- Exe yazılabilir bir klasörde durmalıdır (ör. Masaüstü, Belgeler; `Program Files` değil). Aksi halde otomatik güncelleme sessizce atlanır.
