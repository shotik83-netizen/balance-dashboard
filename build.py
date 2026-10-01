"""Собирает автономный HTML. Сторонние библиотеки не нужны."""
from pathlib import Path
import base64
root=Path(__file__).parent
source=root/'data/source.xlsx'
shell=(root/'src/shell.html').read_text()
js='const BOOT_CONFIG = null;\n'+'\n'.join((root/f'src/{n}.js').read_text() for n in ['core','xlsx-reader','zip-writer','ai-search','bridge','advance','app','excel-export','print','admin'])
html=shell.replace('/*STYLES*/',(root/'src/styles.css').read_text()).replace('/*SOURCE*/',base64.b64encode(source.read_bytes()).decode() if source.exists() else '').replace('/*SCRIPTS*/',js)
(root/'index.html').write_text(html)
print('index.html',len(html.encode()),'bytes')
