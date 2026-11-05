from pathlib import Path
from openpyxl import load_workbook
import json
from collections import Counter
import csv, difflib, sys
from catalogos import cargar, normalizar
root=Path(__file__).resolve().parent.parent
if len(sys.argv)>1:
 lugares=cargar(root)[3]
 with open(sys.argv[1],encoding='utf-8-sig',newline='') as f:
  rows=list(csv.DictReader(f))
 counts=Counter(r['motivo'].split('EAT sin coincidencia única en catálogo: ',1)[1].rsplit(' (candidatos:',1)[0] for r in rows if r['motivo'].startswith('EAT sin coincidencia'))
 for name,count in counts.most_common():
  matches=difflib.get_close_matches(normalizar(name),lugares.keys(),n=3,cutoff=.62)
  print(count,repr(name).encode('ascii','backslashreplace').decode(),[(m,[(p,a,v) for (p,a),v in lugares[m].items()]) for m in matches])
 sys.exit()
wb=load_workbook(root/'345w form fields v9.4.xlsx',read_only=True,data_only=True)
for name in ('Logframe Ind Esp EQ','Logframe Indicadores Específico','Logframe Indicadores EQ'):
 print(name)
 for r in wb[name].values:
  if any('SCA33E/IC1' in str(x) for x in r): print(repr(r).encode('ascii','backslashreplace').decode())
schema=json.loads((root/'ETL/datos_kobo/Proteccion_formulario.json').read_text(encoding='utf8'))
for c in schema['content']['choices']:
 if c.get('list_name')=='actividad' and 'SCA33' in str(c.get('label')): print('KOBO',repr(c).encode('ascii','backslashreplace').decode())
