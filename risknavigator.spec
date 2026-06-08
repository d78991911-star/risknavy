# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller-спецификация сборки автономного .exe «РискНавигатор».

Сборка (на Windows):
    pip install -r requirements.txt
    pip install pyinstaller
    pyinstaller --noconfirm --clean risknavigator.spec

Результат: dist\\RiskNavigator.exe (один файл).

Примечание: PyInstaller НЕ является кросс-компилятором — .exe для Windows
нужно собирать на Windows. На macOS/Linux получится исполняемый файл для
соответствующей ОС.
"""
from PyInstaller.utils.hooks import collect_all, copy_metadata

# Файлы-данные, которые должны попасть внутрь сборки
datas = [
    ("app.py", "."),
]
binaries = []
hiddenimports = [
    "streamlit",
    "streamlit.web.cli",
    "streamlit.runtime.scriptrunner.magic_funcs",
]

# Полный сбор пакетов (код + данные + динамические импорты)
_packages = [
    "streamlit", "plotly", "pandas", "numpy", "networkx",
    "sqlalchemy", "openpyxl", "altair", "pyarrow", "PIL", "dotenv",
]
for pkg in _packages:
    try:
        d, b, h = collect_all(pkg)
        datas += d
        binaries += b
        hiddenimports += h
    except Exception:
        pass

# Метаданные пакетов (Streamlit проверяет свою версию во время выполнения)
for meta in ["streamlit", "plotly", "pandas", "numpy", "networkx",
             "sqlalchemy", "openpyxl", "altair", "pyarrow"]:
    try:
        datas += copy_metadata(meta)
    except Exception:
        pass

block_cipher = None

a = Analysis(
    ["launcher.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="RiskNavigator",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=True,          # окно консоли = «сервер»; закрытие останавливает приложение
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
