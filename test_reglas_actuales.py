import unittest
from pathlib import Path
from catalogos import cargar
from etl import actividad, catalogo_formulario, desagregar
import json

ROOT = Path(__file__).resolve().parent.parent

class Reglas(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.catalogos = cargar(ROOT)
        schema = json.loads((ROOT/'ETL/datos_kobo/Proteccion_formulario.json').read_text(encoding='utf-8'))
        cls.choices = catalogo_formulario(schema)

    def test_sca33(self):
        destino, row = actividad('ACT_43', self.choices, self.catalogos[1])
        self.assertEqual(destino, 'Proteccion')
        self.assertTrue(row['#activity+indicator+name2'].startswith('SCA33E/IC1/IE3:'))
        self.assertEqual(row['#coord+space'], 'NNA')

    def test_total_con_desglose(self):
        a = {'ben_total': '511', 'ninas_0_17_transv': '421'}
        self.assertEqual(desagregar(a, 'Transversal', [])[1], 421)

    def test_total_sin_desglose(self):
        a = {'ben_total': '19', 'ninas_0_17_transv': '0'}
        self.assertEqual(desagregar(a, 'Transversal', [])[1], 19)

if __name__ == '__main__':
    unittest.main()
