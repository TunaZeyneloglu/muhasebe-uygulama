# CLAUDE.md

Bu dosya, Windows Server'da çalışan Claude Code için yayın (Release) prosedürüdür.

## Proje notu

- Uygulama kodu: `code/`
- Sürümün tek kaynağı: `code/version.py` (`APP_VERSION = "X.Y.Z"`)
- Derlemenin tek kaynağı: `FaturaYonetimSistemi.spec` (`derle.ps1` bunu kullanır)

## Önkoşullar

- git, Python 3.12+ (`py`), GitHub CLI (`gh`) kurulu olmalı.
- `gh auth status` başarılı olmalı. Değilse kullanıcıdan bir kez `gh auth login` yapmasını iste. Token isteme, token yazma.

## Yayın adımları

1. `git pull`
2. `code/version.py` içindeki `APP_VERSION`'ı artır. Kullanıcı sürüm belirtmediyse son yama numarasını +1 yap (ör. `2.9.18` -> `2.9.19`).
3. Commit + push.
4. Derle (derle.bat DEĞİL; o `pause` ile bekler):
   ```
   powershell -NoProfile -ExecutionPolicy Bypass -File .\derle.ps1 -Temiz
   ```
5. `dist\FaturaYonetimSistemi.exe` dosyasının var olduğunu kontrol et.
6. Tag:
   ```
   git tag vX.Y.Z
   git push origin vX.Y.Z
   ```
7. Release:
   ```
   gh release create vX.Y.Z dist\FaturaYonetimSistemi.exe --title "Fatura Yonetim Sistemi vX.Y.Z" --notes "<kisa Turkce degisiklik ozeti>"
   ```
   Notlar uygulama içi pop-up'ta düz metin olarak gösterilir: kısa Türkçe düz metin yaz, Markdown başlık/işaret kullanma.
8. Doğrula:
   ```
   gh release view vX.Y.Z --json assets,isDraft,isPrerelease
   ```
   Asset adı birebir `FaturaYonetimSistemi.exe` olmalı; `isDraft` ve `isPrerelease` `false` olmalı.

## Kurallar

- Asset adını asla değiştirme (`FaturaYonetimSistemi.exe`).
- Draft veya prerelease kullanma; updater bunları yoksayar.
- Aynı tag'e ikinci Release oluşturma. Hata varsa exe'yi yeniden yükle:
  ```
  gh release upload vX.Y.Z dist\FaturaYonetimSistemi.exe --clobber
  ```
- Tag (`vX.Y.Z`) ile `code/version.py` her zaman eşleşmeli.
- Updater logu: `%LOCALAPPDATA%\FaturaYonetimSistemi\guncelleme.log`
