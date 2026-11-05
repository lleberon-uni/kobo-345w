import json
import shutil
import sys
import tempfile
import re
import unittest
from pathlib import Path
from unittest.mock import patch
import etl


class Historial(unittest.TestCase):
    def test_consulta_api_real(self):
        original = etl.BASE
        token = (etl.ROOT/'apikey.txt').read_text(encoding='utf-8-sig').strip()
        if token.lower().startswith('token '):
            token = token[6:].strip()
        url = re.search(r'WASH\s+(https://\S+)', (etl.ROOT/'endpoints.txt').read_text(encoding='utf-8-sig')).group(1)
        with tempfile.TemporaryDirectory(dir=original, prefix='prueba_api_') as carpeta:
            with patch.object(etl, 'BASE', Path(carpeta)):
                indices = etl.actualizar_indice('WASH', url, token)
                self.assertTrue(indices)
                indices.sort(key=lambda r: (r['_submission_time'], int(r['_id'])))
                guardados = json.loads((Path(carpeta)/'datos_kobo/WASH_indice.json').read_text(encoding='utf-8'))
                self.assertEqual(set(guardados[0]), {'_id','_uuid','_submission_time'})
                args = ['etl.py', '--sector', 'WASH', '--desde-id', str(indices[-1]['_id'])]
                with patch.object(sys, 'argv', args):
                    self.assertIn(etl.main(), (0, 2))
                resumen = json.loads(next((Path(carpeta)/'salidas'/'WASH').iterdir()).joinpath('resumen.json').read_text(encoding='utf-8'))
                self.assertEqual(resumen['formularios']['WASH']['envios'], 1)

    def test_seleccion_y_deduplicacion(self):
        original = etl.BASE
        with tempfile.TemporaryDirectory(dir=original, prefix='prueba_historial_') as carpeta:
            base = Path(carpeta)
            (base/'datos_kobo').mkdir()
            for sufijo in ('.json', '_formulario.json'):
                shutil.copyfile(original/'datos_kobo'/f'WASH{sufijo}', base/'datos_kobo'/f'WASH{sufijo}')
            registros = json.loads((base/'datos_kobo'/'WASH.json').read_text(encoding='utf-8'))
            registros.sort(key=lambda r: (r['_submission_time'], int(r['_id'])))
            inicio = str(registros[-2]['_id'])
            args = ['etl.py', '--local', '--sector', 'WASH', '--desde-id', inicio]
            with patch.object(etl, 'BASE', base), patch.object(sys, 'argv', args):
                self.assertEqual(etl.main(), 0)
            historia = json.loads((base/'historial.json').read_text(encoding='utf-8'))
            self.assertEqual(Path(historia['transformaciones'][0]['salida']).parent.name, 'WASH')
            self.assertEqual([str(i['id']) for i in historia['transformaciones'][0]['envios']], [str(r['_id']) for r in registros[-2:]])
            self.assertTrue(all(i['submission_time'] and i['uuid'] for i in historia['transformaciones'][0]['envios']))
            with patch.object(etl, 'BASE', base), patch.object(sys, 'argv', args):
                with self.assertRaisesRegex(ValueError, 'No hay envíos nuevos'):
                    etl.main()
            with patch.object(etl, 'BASE', base), patch.object(sys, 'argv', args + ['--reprocesar']):
                self.assertEqual(etl.main(), 0)
            historia = json.loads((base/'historial.json').read_text(encoding='utf-8'))
            self.assertEqual(len(historia['transformaciones']), 2)


if __name__ == '__main__':
    unittest.main()
