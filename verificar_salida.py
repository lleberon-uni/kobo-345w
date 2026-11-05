from pathlib import Path
from openpyxl import load_workbook
import re
import csv
from catalogos import cargar

root = Path(__file__).resolve().parent
resumenes = list((root/'salidas').rglob('resumen.json'))
if not resumenes:
    raise SystemExit('No se encontraron salidas para verificar')
salida = max(resumenes, key=lambda p: p.parent.name).parent
catalogo = load_workbook(root.parent/'345w form fields v9.4.xlsx', read_only=True, data_only=True)
organizaciones = {(r[0], r[2]) for r in list(catalogo['Organizaciones'].values)[1:] if r[0] and r[2]}
_, _, _, lugares, _, vocab, admins, combos, eat_admin = cargar(root.parent)
establecimientos = {(code,name,adm3) for options in lugares.values() for (code,adm3),name in options.items()}
indicadores = set()
for hoja in ('Logframe Indicadores EQ','Logframe Indicadores'):
    for row in list(catalogo[hoja].values)[1:]:
        if row[6] and row[5]:
            indicadores.add((row[6],row[5]))
for path in sorted(salida.glob('*.xlsx')):
    wb = load_workbook(path)
    ws = wb.active
    assert ws.title == path.stem.lower()
    tags = [c.value for c in ws[1]]
    assert len(ws.tables) == 0 and not ws.auto_filter.ref
    assert len(wb.worksheets) == 1 and ws.freeze_panes is None
    assert not ws.column_dimensions and not ws.row_dimensions
    assert all(not cell.has_style for row in ws.iter_rows() for cell in row)
    assert tags[0] == '#date+month+cycle' and len(tags) == len(set(tags))
    for row in ws.iter_rows(min_row=2, values_only=True):
        d = dict(zip(tags,row))
        assert re.fullmatch(r'\d{4}-\d{2}', d['#date+month+cycle'])
        assert d['#activity+natural+disaster'] in ('Sí','No')
        assert d['#modalidad+tm'] == 'No'
        if '#org+impl+id' in d:
            assert (d['#org+impl'],d['#org+impl+id']) in organizaciones
        assert d['#location+type'] in vocab['#location+type']
        if d.get('#adm4+code'):
            assert (d['#adm4+code'],d['#adm4+name'],d['#adm3+code']) in establecimientos
        assert (d['#indicator+code'], d['#activity+indicator+name']) in indicadores
        for level in (1,2,3):
            assert d[f'#adm{level}+code'] in admins[level]
    print(path.name, ws.max_row-1, 'filas verificadas')
    wb.close()
catalogo.close()
with (salida/'trazabilidad.csv').open(encoding='utf-8-sig', newline='') as f:
    revisados = [r for r in csv.DictReader(f) if 'EAT sin coincidencia única' in r['advertencias']]
for archivo in {r['archivo'] for r in revisados}:
    book = load_workbook(salida/archivo,read_only=True,data_only=True)
    ws = book.active
    rows = list(ws.values)
    headers = rows[0]
    for item in (r for r in revisados if r['archivo'] == archivo):
        data = dict(zip(headers,rows[int(item['fila_excel'])-1]))
        assert data['#location+type'] == 'Otro' and not data.get('#adm4+code') and not data.get('#adm4+name') and data.get('#location')
    book.close()
print(f'{len(revisados)} EAT sin coincidencia comprobados')
