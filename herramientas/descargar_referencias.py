"""Descarga las 23 imágenes técnicas públicas del repositorio preinca/dashboard."""
from pathlib import Path
from urllib.request import urlopen, Request
from urllib.parse import quote

FILES = [
    'manifold_investigacion_desarrollo.jpg','manifold_pick_heater.jpg',
    'manifold_vapor_margarina.jpg','manifold_vapor_principal.jpg',
    'margarina_fundicion_emulsificante_1.jpg','margarina_fundicion_emulsificante_2.jpg',
    'margarina_fundicion_emulsificante_3.jpg','margarina_fundicion_emulsificante_4.jpg',
    'margarina_pasteurizacion_1.jpg','margarina_pasteurizacion_2.jpg',
    'margarina_pasteurizacion_3.jpg','margarina_pasteurizacion_4.jpg',
    'mayonesa_cip.jpg','mayonesa_manifold.jpg','queso_cocina_1.jpg','queso_cocina_2.jpg',
    'queso_fase_oleosa.jpg','queso_intercambiador_cip.jpg',
    'queso_linea_alimentacion.jpg','queso_tapadora.jpg','queso_tina_fundicion.jpg',
    'tomate_intercambiador_calor.jpg','tomate_intercambiador_tubular_salsa_base.jpg',
]
DEST = Path(__file__).resolve().parents[1] / 'referencias_condensado'
BASE = 'https://raw.githubusercontent.com/preinca/dashboard/main/referencias_condensado/'
DEST.mkdir(parents=True, exist_ok=True)
for name in FILES:
    path = DEST / name
    if path.exists():
        print('Existe:', name)
        continue
    url = BASE + quote(name)
    try:
        req = Request(url, headers={'User-Agent': 'ChecklistVaporCondensado/1.1'})
        with urlopen(req, timeout=30) as response:
            data = response.read()
        if not data.startswith(b'\xff\xd8\xff'):
            raise ValueError('La respuesta no contiene un archivo JPEG')
        path.write_bytes(data)
        print('Descargado:', name)
    except Exception as e:
        print('ERROR', name, '-', e)
print('Listo. Fotografías locales:', len(list(DEST.glob('*.jpg'))), 'de', len(FILES))
