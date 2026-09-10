# -*- mode: python ; coding: utf-8 -*-
import os
import playwright

# Динамически находим путь к драйверу Playwright
playwright_path = os.path.dirname(playwright.__file__)
playwright_driver = os.path.join(playwright_path, 'driver')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Обязательно добавляем драйвер в сборку
        (playwright_driver, 'playwright/driver'),
    ],
    hiddenimports=[
        'playwright.sync_api',
        'requests',
        'docx'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Исключаем тяжелые сторонние библиотеки, если они установлены в системе
        'numpy', 'pandas', 'matplotlib', 'scipy', 'PyQt5', 'PyQt6', 'PySide6', 'tkinter.test'
    ],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='DSTU_Checker',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,          # Используем сжатие UPX (если установлен)
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,     # Отключаем черное окно консоли
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,         # Сюда можно вписать путь к иконке: 'icon.ico'
)