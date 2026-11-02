# Reproducible build from reviewed Python source; no legacy bytecode.
from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

root=Path(SPECPATH)
datas=collect_data_files('rapidocr_onnxruntime')
hiddenimports=collect_submodules('rapidocr_onnxruntime')+['pdfplumber','numpy']
a=Analysis([str(root/'src/main.py')],pathex=[str(root/'src')],
    datas=datas,hiddenimports=hiddenimports,
    excludes=['pandas','scipy','matplotlib','IPython','pytest','PyQt5','PyQt6','PySide6',
              'jupyter','torch','sklearn','sympy','numba'],noarchive=False)
pyz=PYZ(a.pure)
exe=EXE(pyz,a.scripts,[],exclude_binaries=True,name='HybeFiller_v1.3.1',
        debug=False,strip=False,upx=False,console=False,disable_windowed_traceback=False)
cli=EXE(pyz,a.scripts,[],exclude_binaries=True,name='HybeFillerCLI',
        debug=False,strip=False,upx=False,console=True,disable_windowed_traceback=False)
coll=COLLECT(exe,cli,a.binaries,a.datas,strip=False,upx=False,name='HybeFiller_v1.3.1_portable')
