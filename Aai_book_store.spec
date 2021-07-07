# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['Aai_book_store.py'],
             pathex=['C:\\Users\\bhushaga\\PycharmProjects\\Aai_book_store'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=['matplotlib', 'numpy', 'gtts', 'gTTS', 'speech_recognition', 'pandas', 'gtts_token', 'pyttsx3', 'spicy', 'PyQt5', 'selenium'],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='Aai_book_store',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False )
