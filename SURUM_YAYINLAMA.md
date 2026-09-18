# Sürüm Yayınlama

## Derleme

Kök klasördeki `derle.bat` dosyasına çift tıklayın. Sanal ortamı (`.venv`) kurar, derler, exe'nin boyutunu ve SHA256'sını yazdırır. Çıktı: `dist\FaturaYonetimSistemi.exe`

## Yayın

Yayın normalde sunucudaki Claude Code ile yapılır; adımlar `CLAUDE.md` dosyasındadır.

### Elle yayın

1. `code/version.py` içindeki sürümü artırın (satır biçimi tam olarak `APP_VERSION = "X.Y.Z"` olmalı), commit edip push edin.
2. `derle.bat` ile derleyin.
3. GitHub Releases'te `vX.Y.Z` tag'iyle yeni bir Release oluşturun ve `dist\FaturaYonetimSistemi.exe` dosyasını bu adla yükleyin. Draft veya prerelease işaretlemeyin.

### Yedek yol: GitHub Actions

Actions -> Release -> Run workflow -> tag'i girin (ör. `v2.9.19`). Tag önceden push edilmiş olmalı ve `code/version.py` ile eşleşmelidir. Tag push'u Actions'ı kendiliğinden tetiklemez.

## Notlar

- Açık uygulamalar yeni sürümü arka planda (GitHub Releases `latest`) bulur. Draft ve prerelease sürümler dikkate alınmaz.
- Bu sistemden önceki (konsollu) exe'ler kendini güncelleyemez. Yeni exe kullanıcılara bir kez elle verilmelidir.
- Exe yazılabilir bir klasörde durmalıdır (ör. Masaüstü, Belgeler; `Program Files` değil). Aksi halde otomatik güncelleme sessizce atlanır.

## macOS (Apple Silicon)

- Her sürümde Windows exe'yi sunucu yayınlar (`CLAUDE.md` prosedürü). Ardından aynı commit'ten bir Mac'te kök klasördeki `derle_mac.command` dosyasına çift tıklayın. Sanal ortamı (`.venv-mac`) kurar, `dist/FaturaYonetimSistemi.app`'i derler, `dist/FaturaYonetimSistemi-macOS.zip`'i oluşturur, boyutunu ve SHA256'sını yazdırır.
- Zip'i aynı Release'e ekleyin:
  ```
  gh release upload vX.Y.Z dist/FaturaYonetimSistemi-macOS.zip --clobber
  ```
- Asset adları birebir `FaturaYonetimSistemi.exe` ve `FaturaYonetimSistemi-macOS.zip` olmalıdır.
- Mac kullanıcıları uygulamayı `Uygulamalar` klasörüne taşımalıdır. Uygulama imzasız olduğundan ilk açılışta sağ tık -> Aç (veya Sistem Ayarları -> Gizlilik ve Güvenlik -> Yine de Aç) gerekir. İndirilenler'den çalıştırılırsa macOS uygulamayı salt okunur bir konumdan çalıştırır ve otomatik güncelleme sessizce atlanır.
