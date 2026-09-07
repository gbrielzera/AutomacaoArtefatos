# -*- mode: python ; coding: utf-8 -*-
"""
Receita de empacotamento do PyInstaller para o Gerador de Artefato Supravizio.

Gera um único .exe (modo onefile) que já carrega dentro de si:
  * o "Modelo de Artefato.docx" (o app o encontra via sys._MEIPASS, veja get_template_padrao);
  * os binários do tkdnd, exigidos pelo tkinterdnd2 para arrastar-e-soltar;
  * os templates internos do python-docx e os lexers do Pygments.

Uso:  pyinstaller supra.spec
Saída: dist/Gerador de Artefato Supravizio.exe

Logs e config/settings.json continuam sendo gravados AO LADO do .exe (e não na pasta
temporária), porque get_base_dir() usa sys.executable quando a aplicação está congelada.
"""

from PyInstaller.utils.hooks import collect_all, collect_data_files

# tkinterdnd2 distribui os binários do tkdnd como dados do pacote; sem eles o app abre,
# mas o arrastar-e-soltar falha em tempo de execução.
dados_dnd, binarios_dnd, imports_dnd = collect_all("tkinterdnd2")

datas = [
    ("Modelo de Artefato.docx", "."),
    *dados_dnd,
    *collect_data_files("docx"),  # templates .docx internos do python-docx
]

hiddenimports = [
    *imports_dnd,
    "pygments.lexers.python",  # importado de forma preguiçosa por `from pygments.lexers import PythonLexer`
]


a = Analysis(
    ["supra.py"],
    pathex=[],
    binaries=binarios_dnd,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["pytest", "_pytest"],
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
    name="Gerador de Artefato Supravizio",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # aplicação gráfica: sem janela de console
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon="icone.ico",  # descomente e coloque um .ico na raiz para personalizar o ícone
)
