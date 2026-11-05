"""Lee los Excel originales como referencia, sin guardarlos ni modificarlos."""
from collections import defaultdict
from pathlib import Path
import re
import unicodedata
from openpyxl import load_workbook


def normalizar(texto):
    texto = unicodedata.normalize('NFKD', str(texto or '').casefold())
    return ''.join(c for c in texto if c.isalnum())


def codigo(texto):
    """Identificador para comparar versiones E y SIN/SI; no inventa salidas."""
    texto = str(texto or '').split(':')[0].strip()
    texto = re.sub(r'((?:SCA|CCA)\d+)E(?=/|$)', r'\1', texto)
    return re.sub(r'/SIN(?=\d)', '/SI', texto)


def leer_hoja(ws):
    rows = iter(ws.values)
    for row in rows:
        if '#date+month+cycle' in row:
            tags = list(row)
            if len([t for t in tags if t]) != len(set(t for t in tags if t)):
                raise ValueError(f'Encabezados duplicados: {ws.title}')
            return tags, [dict(zip(tags, r)) for r in rows if any(v is not None for v in r)]
    return None, []


def cargar(root):
    plantillas, filas, fuentes = {}, [], []
    for sector in ('Proteccion', 'WASH', 'Educacion', 'Transversal'):
        paths = sorted((root/'upload templates').glob(f'*{sector}*.xlsx'))
        if len(paths) != 1:
            raise ValueError(f'Se esperaba exactamente una plantilla para {sector}; hay {len(paths)}')
        path = paths[0]
        wb = load_workbook(path, read_only=True, data_only=True)
        tags, datos = leer_hoja(wb.worksheets[0])
        plantillas[sector] = (wb.worksheets[0].title, tags)
        filas.extend(datos)
        fuentes.append(path)
        wb.close()
    # Los ejemplos de las plantillas tienen prioridad. Los datos finales
    # complementan los códigos que no aparecen en esos lotes incrementales.
    for path in sorted((root/'final data').glob('*.xlsx')):
        wb = load_workbook(path, read_only=True, data_only=True)
        _, datos = leer_hoja(wb.worksheets[0])
        filas.extend(r for r in datos if normalizar(r.get('#validated')) == 'si')
        fuentes.append(path)
        wb.close()
    actividades, orgs, lugares = {}, {}, defaultdict(dict)
    for d in filas:
        ind = d.get('#indicator+code') or d.get('#indicator+name') or d.get('#activity+indicator+name')
        sub = d.get('#activity+indicator+name2')
        act = d.get('#activity+code') or (str(ind).rsplit('/', 1)[0] if ind else '')
        cluster = str(ind).split('/')[0] if str(ind).startswith('GCL') else str(act).split('/')[0]
        if ind:
            short = codigo(sub or ind)
            if short.startswith('GCL'):
                short = short.split('/', 1)[1]
            referencia = dict(d)
            if not referencia.get('#indicator+code') and cluster in ('GCLPR', 'GCLWA', 'GCLED'):
                # En las exportaciones finales el código forma parte del nombre.
                # Se extrae literalmente; no se toma ningún ID de registro 345w.
                literal = str(ind).split(':')[0].strip()
                referencia['#indicator+code'] = literal if literal.startswith(cluster+'/') else cluster+'/'+literal
            actividades.setdefault((cluster, short), referencia)
        nombre = d.get('#org+impl+name') or d.get('#org+impl')
        orgid = d.get('#org+impl+id') or d.get('#org+impl')
        if nombre and ' - ' in str(nombre):
            orgs.setdefault(normalizar(str(nombre).split(' - ')[0]), (nombre, orgid if ' - ' not in str(orgid) else str(nombre).split(' - ')[0]))
            orgs.setdefault(normalizar(nombre), orgs[normalizar(str(nombre).split(' - ')[0])])
        loc = d.get('#adm4+name') or d.get('#location')
        pcode = d.get('#adm4+code')
        if loc and pcode:
            # Un código observado en varias parroquias se conserva con todos sus
            # contextos. No se elige entre candidatos ni se corrige geografía.
            key = normalizar(limpiar_lugar(loc))
            lugares[key].setdefault((pcode, d.get('#adm3+code')), loc)
    # El catálogo oficial es la referencia principal de nombres y códigos.
    path = root/'345w form fields v9.4.xlsx'
    wb = load_workbook(path, read_only=False, data_only=True)
    orgs = {}
    ambiguos = set()
    for row in list(wb['Organizaciones'].values)[1:]:
        selected, nombre, ident, corto = row[:4]
        if not selected or not ident:
            continue
        for alias in (selected, nombre, corto):
            key = normalizar(alias)
            if not key:
                continue
            valor = (selected, ident)
            if key in orgs and orgs[key] != valor:
                ambiguos.add(key)
            else:
                orgs[key] = valor
    for key in ambiguos:
        orgs.pop(key, None)
    def registros(nombre):
        rows = iter(wb[nombre].values)
        tags = next(rows)
        return [dict(zip(tags, row)) for row in rows]
    acts = registros('Logframe Actividades EQ') + registros('Logframe Actividades')
    inds = registros('Logframe Indicadores EQ') + registros('Logframe Indicadores')
    specs = registros('Logframe Ind Esp EQ')
    # No presentar como válidos códigos históricos ausentes del catálogo vigente.
    actividades = {k:v for k,v in actividades.items() if k[0] not in ('GCLPR', 'GCLED', 'GCLWA', 'GCLCO')}
    for ind in inds:
        code = ind.get('#indicator+code')
        if not code or not str(code).startswith('GCL') or '/' not in str(code):
            continue
        cluster, short = str(code).split('/', 1)
        matching = [a for a in acts if a.get('#activity+code') == ind.get('#activity+code')]
        # Un indicador puede tener varias áreas; la elección se hace después.
        ref = dict(ind)
        if matching:
            ref.update({k:v for k,v in matching[0].items() if v is not None})
        ref['#indicator+name'] = ind.get('#activity+indicator+name')
        ref['#meta+unit'] = ind.get('#activity+indicator+unit')
        actividades.setdefault((cluster, codigo(short)), ref)
    for spec in specs:
        code = spec.get('#indicator+code2')
        parent = spec.get('#indicator+code')
        if not code or not parent or '/' not in str(code):
            continue
        cluster, short = str(code).split('/', 1)
        parent_short = str(parent).split('/', 1)[-1]
        ref = actividades.get((cluster, codigo(parent_short)))
        if ref:
            ref = dict(ref)
            ref.update({k:v for k,v in spec.items() if v is not None})
            actividades[(cluster, codigo(short))] = ref
    # EAT contiene la parroquia codificada en la columna etiquetada #adm3+name.
    from openpyxl.utils.cell import range_boundaries
    sheet = wb['Adm']
    admins = {1: {}, 2: {}, 3: {}}
    for row in sheet.iter_rows(min_row=2, values_only=True):
        for level, code_col, name_col, parent_col in ((1,1,0,None),(2,7,6,5),(3,13,12,11)):
            if row[code_col] and row[name_col]:
                admins[level][row[code_col]] = (row[name_col], row[parent_col] if parent_col is not None else None)
    table = sheet.tables.get('EAT')
    if table is None:
        raise ValueError('El catálogo no contiene la tabla EAT')
    x1, y1, x2, y2 = range_boundaries(table.ref)
    eat = list(sheet.iter_rows(min_row=y1, max_row=y2, min_col=x1, max_col=x2, values_only=True))
    headers = list(eat[0])
    def index(tag):
        if headers.count(tag) != 1:
            raise ValueError(f'Columna EAT ausente o duplicada: {tag}')
        return headers.index(tag)
    ci, ni, pi = (index(k) for k in ('#adm4+code', '#adm4+name', '#adm3+name'))
    lugares = defaultdict(dict)
    eat_admin = {}
    i1, i2 = index('#adm1+code'), index('#adm2+code')
    for row in eat[1:]:
        code4, name4, adm3 = row[ci], row[ni], row[pi]
        if code4 and name4 and adm3:
            lugares[normalizar(limpiar_lugar(name4))][(code4, adm3)] = name4
            eat_admin[(code4, adm3)] = (row[i1], row[i2], adm3)
    vocab = {}
    sheet = wb['Vocab']
    for table in sheet.tables.values():
        x1, y1, x2, y2 = range_boundaries(table.ref)
        for col in range(x1, x2 + 1):
            tag = sheet.cell(y1, col).value
            if isinstance(tag, str) and tag.startswith('#'):
                vocab.setdefault(tag, set()).update(sheet.cell(r, col).value for r in range(y1+1, y2+1) if sheet.cell(r, col).value is not None)
    wb.close()
    fuentes.append(path)
    combos = {(a.get('#activity+code'), a.get('#activity+area_type')) for a in acts if a.get('#activity+code') and a.get('#activity+area_type')}
    return plantillas, actividades, orgs, lugares, fuentes, vocab, admins, combos, eat_admin


def limpiar_lugar(nombre):
    return re.sub(r'\s*\((?:Campamento (?:ONU|Gobierno)|Asentamiento Informal)\)\s*$', '', str(nombre or ''), flags=re.I).strip()
