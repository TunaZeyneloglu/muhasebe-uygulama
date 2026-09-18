# -*- mode: python ; coding: utf-8 -*-
# Derleme: pyinstaller --noconfirm --clean FaturaYonetimSistemi.spec
# Çıktı: dist/FaturaYonetimSistemi.exe (tek dosya, konsolsuz)
import os

from PyInstaller.utils.hooks import collect_data_files

KOK = SPECPATH
KOD = os.path.join(KOK, 'code')

# customtkinter tema JSON'ları ve asset'leri runtime'da dosyadan okunur.
# code/ altında runtime'da dosyadan okunan başka veri yok (ikonlar icons.py'de
# PIL ile programatik çiziliyor).
datas = collect_data_files('customtkinter')

a = Analysis(
    [os.path.join(KOD, 'main.py')],
    pathex=[KOD],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Boyut küçültmek için: uygulamanın çalışma zamanında kullanmadığı modüller.
    excludes=[
        'sqlite3', 'unittest', 'pydoc', 'doctest', 'pdb', 'lib2to3',
        'distutils', 'setuptools', 'pip', 'test',
        'matplotlib', 'scipy', 'IPython', 'pyarrow', 'numexpr', 'bottleneck',
        'jinja2', 'pytest', 'hypothesis',
        'numpy.fft', 'numpy.polynomial', 'numpy.testing', 'numpy.distutils',
        'numpy.f2py',
        'PIL.AvifImagePlugin', 'PIL._avif', 'PIL.WebPImagePlugin', 'PIL._webp',
        'PIL.ImageCms', 'PIL._imagingcms', 'PIL.PdfImagePlugin',
        'PIL.PdfParser', 'PIL.ImageMath', 'PIL._imagingmath', 'PIL.ImageShow',
        'PIL.ImageQt',
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FaturaYonetimSistemi',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join(KOK, 'icon', 'ikon.ico'),
)
