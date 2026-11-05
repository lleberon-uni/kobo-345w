"""Kobo -> archivos de carga 345w. Ejecutar: python ETL/etl.py"""
import argparse
import csv
from datetime import date, datetime, timezone
from decimal import Decimal, InvalidOperation
import hashlib
import json
from pathlib import Path
import re
import sys
import time
from urllib.error import HTTPError, URLError
from urllib.parse import urlsplit, urljoin, urlencode
from urllib.request import Request, build_opener, HTTPRedirectHandler

from openpyxl import Workbook
from catalogos import cargar, codigo, limpiar_lugar, normalizar
from mapeos import DESAGREGACIONES, LIDER, SECTORES, TIPOS_LUGAR, TRANSVERSAL, VULNERABILIDAD

BASE = Path(__file__).resolve().parent
ROOT = BASE.parent


class RedireccionKobo(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        if urlsplit(newurl).scheme != 'https' or urlsplit(newurl).netloc != urlsplit(req.full_url).netloc:
            raise ValueError('Kobo redirigió la petición a un origen diferente')
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def guardar_json(path, data):
    temporal = path.with_suffix('.tmp')
    temporal.write_text(json.dumps(data, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
    temporal.replace(path)


def obtener_json(url, token, origen):
    """Sólo envía la clave al origen Kobo configurado; reintenta fallos transitorios."""
    parsed = urlsplit(url)
    if parsed.scheme != 'https' or parsed.netloc != origen:
        raise ValueError('La URL de Kobo o su paginación cambió de origen HTTPS')
    request = Request(url, headers={'Authorization': f'Token {token}', 'Accept': 'application/json'})
    for intento in range(3):
        try:
            with build_opener(RedireccionKobo()).open(request, timeout=60) as response:
                return json.load(response)
        except HTTPError as exc:
            if exc.code not in (429, 500, 502, 503, 504) or intento == 2:
                raise RuntimeError(f'Kobo respondió HTTP {exc.code}; revisar acceso y endpoint') from None
        except (URLError, TimeoutError):
            if intento == 2:
                raise RuntimeError('No se pudo conectar con Kobo tras 3 intentos') from None
        time.sleep(2 ** intento)


def extraer(sector, url, token):
    cache = BASE/'datos_kobo'
    cache.mkdir(exist_ok=True)
    origen = urlsplit(url).netloc
    records, paginas, next_url, esperado = [], set(), url, None
    while next_url:
        if next_url in paginas:
            raise ValueError('Kobo devolvió una página repetida')
        paginas.add(next_url)
        data = obtener_json(next_url, token, origen)
        if not isinstance(data, dict) or not isinstance(data.get('results'), list):
            raise ValueError('Respuesta Kobo sin lista results')
        if esperado is None:
            esperado = data.get('count')
        records.extend(data['results'])
        next_url = urljoin(url, data['next']) if data.get('next') else None
    ids = [r.get('_id') for r in records]
    if None in ids or len(set(ids)) != len(ids) or esperado is not None and len(records) != esperado:
        raise ValueError('La descarga tiene IDs repetidos o cambió el conteo; repetir la extracción')
    schema = obtener_json(url.rsplit('/data', 1)[0]+'/', token, origen)
    if not schema.get('content', {}).get('choices'):
        raise ValueError('Kobo no devolvió el catálogo del formulario')
    guardar_json(cache/f'{sector}.json', records)
    guardar_json(cache/f'{sector}_formulario.json', schema)
    return records, schema


def consultar(url, token, parametros):
    origen = urlsplit(url).netloc
    next_url = url + ('&' if '?' in url else '?') + urlencode(parametros)
    paginas, records, esperado = set(), [], None
    while next_url:
        if next_url in paginas:
            raise ValueError('Kobo devolvió una página repetida')
        paginas.add(next_url)
        data = obtener_json(next_url, token, origen)
        if not isinstance(data, dict) or not isinstance(data.get('results'), list):
            raise ValueError('Respuesta Kobo sin lista results')
        if esperado is None:
            esperado = data.get('count')
        records.extend(data['results'])
        next_url = urljoin(url, data['next']) if data.get('next') else None
    ids = [r.get('_id') for r in records]
    if None in ids or len(set(ids)) != len(ids) or esperado is not None and len(records) != esperado:
        raise ValueError('La descarga tiene IDs repetidos o cambió el conteo; actualice el sector')
    return records


def actualizar_indice(sector, url, token):
    cache = BASE/'datos_kobo'
    cache.mkdir(exist_ok=True)
    registros = consultar(url, token, {'limit': 1000, 'fields': json.dumps(['_id', '_uuid', '_submission_time'])})
    if any(not r.get('_submission_time') for r in registros):
        raise ValueError('Kobo devolvió un envío sin fecha')
    esquema = obtener_json(url.rsplit('/data', 1)[0]+'/', token, urlsplit(url).netloc)
    if not esquema.get('content', {}).get('choices'):
        raise ValueError('Kobo no devolvió el catálogo del formulario')
    guardar_json(cache/f'{sector}_formulario.json', esquema)
    guardar_json(cache/f'{sector}_indice.json', [{k:r.get(k) for k in ('_id','_uuid','_submission_time')} for r in registros])
    return registros


def simple(record):
    """Acorta las rutas del grupo; falla si dos campos distintos tienen el mismo nombre."""
    result = {}
    for key, value in record.items():
        short = key.split('/')[-1]
        if short in result:
            raise ValueError(f'Campo ambiguo: {short}')
        result[short] = value
    return result


def numero(value, campo):
    if value in (None, ''):
        return None
    try:
        num = Decimal(str(value))
        if not num.is_finite() or num < 0 or num != num.to_integral_value():
            raise InvalidOperation
        return int(num)
    except InvalidOperation:
        raise ValueError(f'{campo}: se esperaba un entero no negativo') from None


def catalogo_formulario(schema):
    return {(c['list_name'], c['name']): c for c in schema['content']['choices']}


def etiqueta(choices, lista, valor):
    c = choices.get((lista, valor))
    if c is None:
        raise ValueError(f'Valor fuera del catálogo Kobo: {lista}={valor}')
    label = c.get('label', valor)
    return label[0] if isinstance(label, list) else label


def actividad(code, choices, actividades):
    label = etiqueta(choices, 'actividad', code)
    source = choices['actividad', code].get('filter_sector_act')
    if source == 'TRANSV':
        if code not in TRANSVERSAL:
            raise ValueError(f'Actividad transversal sin correspondencia confirmada: {code}')
        act, ind = TRANSVERSAL[code]
        cluster = act.split('/')[0]
        destino = 'Transversal'
    else:
        destino = {'PROT': 'Proteccion', 'EDU': 'Educacion', 'WASH': 'WASH'}.get(source)
        if destino is None:
            raise ValueError(f'Sector Kobo no reconocido: {source}')
        cluster = {'Proteccion': 'GCLPR', 'Educacion': 'GCLED', 'WASH': 'GCLWA'}[destino]
        ind = 'CCA3/CIN3' if destino == 'Educacion' and code == 'ACT_11' else label.split(':')[0].strip()
        if destino == 'Proteccion' and codigo(ind) == 'SCA33/IC1':
            ind = 'SCA33E/IC1/IE3'
    referencia = actividades.get((cluster, codigo(ind)))
    if referencia is None:
        raise ValueError(f'Indicador sin referencia en plantillas o datos validados: {cluster}/{ind}')
    # Sólo campos descriptivos de actividad, nunca cantidades, IDs o datos del ejemplo.
    campos = ('#activity+code', '#activity+name+selected', '#indicator+code', '#indicator+name',
              '#activity+indicator+name', '#activity+indicator+name2', '#meta+unit',
              '#activity+indicator+unit', '#activity+area_type', '#area+type', '#coord+space')
    result = {k: referencia.get(k) for k in campos}
    if destino == 'Proteccion':
        result['#coord+space'] = 'VBG' if ': VBG' in label else 'NNA' if ': NNA' in label else referencia.get('#coord+space')
    return destino, result


def desagregar(a, destino, advertencias):
    mapping = {k: list(v) for k, v in DESAGREGACIONES[destino].items()}
    if destino == 'WASH':
        for sexo, infancia, adulto in (('f', 'ninas', 'mujeres'), ('m', 'ninos', 'hombres')):
            # Versiones anteriores del formulario usaban tres tramos infantiles.
            if a.get(f'{infancia}_0_17') in (None, ''):
                mapping[f'#reached+{sexo}+children+age0_17'] = [f'{infancia}_{edad}' for edad in ('0_5', '6_11', '12_17')]
            if a.get(f'{adulto}_60_79') in (None, '') and a.get(f'{adulto}_60_mas') not in (None, ''):
                # No se puede distribuir 60+ entre 60-79 y 80+ sin inventar edades.
                if numero(a[f'{adulto}_60_mas'], f'{adulto}_60_mas') > 0:
                    raise ValueError('WASH antiguo: 60+ no se puede separar en 60-79 y 80+')
    if destino == 'Educacion' and any(k.startswith('cdi_') and a[k] not in (None, '', '0', 0) for k in a):
        raise ValueError('Desagregación CDI especial incompatible con los tramos de la plantilla')
    result = {}
    for tag, fields in mapping.items():
        values = [numero(a.get(f), f) for f in fields]
        result[tag] = sum(v for v in values if v is not None) if any(v is not None for v in values) else None
    total = numero(a.get('ben_total'), 'ben_total')
    suma = sum(v for v in result.values() if v is not None)
    if total is None:
        if all(v is not None for v in result.values()):
            total = suma
        else:
            raise ValueError('Falta ben_total y la desagregación no está completa')
    if suma != total and not (destino == 'Transversal' and suma == 0):
        advertencias.append(f'Total reemplazado por suma de desagregaciones: {total} -> {suma}')
        total = suma
    elif suma == 0 and destino == 'Transversal' and total:
        advertencias.append(f'Total original {total} conservado: desagregaciones suman 0')
    calc = numero(a.get('calc_total'), 'calc_total')
    if calc is not None and calc != total:
        advertencias.append(f'calc_total original {calc}; total exportado {total}')
    return result, total


def transformar(record, raw, choices, catalogs):
    _, actividades, orgs, lugares, _, vocab, admins, combos, eat_admin = catalogs
    r, a = simple(record), simple(raw)
    destino, out = actividad(a.get('activity_name'), choices, actividades)
    advertencias = []
    if a.get('activity_name') == 'GEN_2' and destino == 'Transversal':
        advertencias.append('Equivalencia solicitada GEN.02 -> GCLCO/CCA1E (PEAS); confirmar interpretación programática')
    if a.get('activity_name') == 'ACT_11' and destino == 'Educacion':
        advertencias.append('Equivalencia solicitada CDI -> CCA3/CIN3')
    out.update({'#meta+hrp': 'Respuesta al terremoto', '#fund+type': 'Ninguno de los anteriores',
                '#org+leader': LIDER, '#org+leader+id': 429, '#modalidad+tm': 'No',
                '#reached+type': 'Ninguna de las anteriores',
                '#activity+natural+disaster': 'Sí'})
    fecha = date.fromisoformat(str(r.get('date_reported', '')))
    out['#date+month+cycle'] = fecha.strftime('%Y-%m')
    org = r.get('org_impl_name')
    org_label = etiqueta(choices, 'orgs', org)
    if org == 'SOCIO':
        org_label = r.get('org_impl_otro_cual') or ''
    key = normalizar({'UNICEF_DI': 'UNICEF', 'Accion Contra El hambre': 'ACH'}.get(org, org_label))
    resolved = orgs.get(key) or orgs.get(normalizar(org))
    if resolved is None:
        raise ValueError(f'Organización sin equivalencia 345w confirmada: {org_label}')
    out['#org+impl'], out['#org+impl+id'] = resolved
    tipo = r.get('tipo_instalacion')
    es_eat = tipo in ('albergue', 'albergue_onu')
    for n, lista in ((1, 'estado'), (2, 'municipio'), (3, 'parroquia')):
        value = r.get(f'adm{n}_name')
        if not es_eat:
            etiqueta(choices, lista, value)
        if value not in admins[n] and not es_eat:
            raise ValueError(f'Código ADM{n} no está en catálogo 345w: {value}')
        name, parent345 = admins[n].get(value, (None, None))
        if not es_eat and n > 1 and parent345 != out[f'#adm{n-1}+code']:
            raise ValueError(f'Jerarquía ADM{n} no coincide con catálogo 345w')
        out[f'#adm{n}+name'] = name
        out[f'#adm{n}+code'] = value
        if n > 1 and not es_eat:
            choice = choices[lista, value]
            parent = choice.get('filter_estado' if n == 2 else 'filter_muni_parroquia')
            if parent and parent != out[f'#adm{n-1}+code']:
                raise ValueError('Jerarquía geográfica Kobo inconsistente')
    out['#location+type'] = TIPOS_LUGAR.get(destino, {}).get(tipo)
    if destino == 'Transversal':
        out['#location+type'] = {'comunidad':'Comunidad', 'escuela':'Escuela o centro educativo', 'albergue':'EAT - Espacio de alojamiento temporal (civil)', 'albergue_onu':'EAT - Espacio de alojamiento temporal (civil)'}.get(tipo)
    equivalencias = {'asentamientos_informales': 'Otro', 'asentamiento': 'Otro', 'centro_salud': 'Hospital', 'espacio_aprendizaje': 'Escuela o centro educativo', 'Gobie': 'Otro', 'otro_albergue': 'Otro'}
    if tipo in equivalencias:
        out['#location+type'] = equivalencias[tipo]
        advertencias.append(f'Tipo de instalación {tipo} -> {equivalencias[tipo]}')
    if not out['#location+type']:
        raise ValueError(f'Tipo de instalación sin equivalencia confirmada: {tipo}')
    location = r.get('loc_name') or r.get('campamento_otro')
    for field, lista in (('campamento_select', 'campamento'), ('campamento_onu_select', 'camp_onu'),
                         ('campamento_asent_select', 'camp_asent'), ('centro_salud_select', 'establecimiento')):
        value = r.get(field)
        if value and value != 'otro_albergue':
            location = etiqueta(choices, lista, value)
            break
    location_original = str(location or '').strip()
    location = limpiar_lugar(location)
    if not location:
        raise ValueError('No se encontró nombre del lugar')
    out['#location'] = location
    out['#adm4+name'] = location
    aliases_eat = {
        'estadiocesarnieves': 'campamentocesarnieves',
        'universidadsimonbolivarsedelitoral': 'universidadsimonbolivar',
        'grancolombia': 'uegrancolombia',
    }
    lookup = normalizar(location)
    lookup_catalogo = aliases_eat.get(lookup, lookup) if es_eat else lookup
    candidates = lugares.get(lookup_catalogo, {})
    validos = {(pcode, nombre) for (pcode, adm3), nombre in candidates.items() if adm3 == out['#adm3+code']}
    if es_eat:
        matches = [(pcode, adm3, nombre) for (pcode, adm3), nombre in candidates.items()]
        same = [entry for entry in matches if entry[1] == out['#adm3+code']]
        matches = same if len(same) == 1 else matches
        if len(matches) != 1:
            out['#location+type'] = 'Otro'
            out['#location'] = location_original
            out['#adm4+code'] = None
            out['#adm4+name'] = None
            advertencias.append(f'EAT sin coincidencia única ({len(matches)} candidatos): tipo Otro, ADM4 vacío; ubicación original conservada')
            for n in (1, 2, 3):
                code = out[f'#adm{n}+code']
                if code not in admins[n] or (n > 1 and admins[n][code][1] != out[f'#adm{n-1}+code']):
                    raise ValueError(f'Ubicación original ADM{n} inválida; no se puede exportar como Otro')
            es_eat = False
        else:
            code4, adm3, nombre = matches[0]
    if es_eat:
        codes = eat_admin[(code4, adm3)]
        if any(not code or code not in admins[n] for n, code in enumerate(codes, 1)):
            raise ValueError(f'EAT sin jerarquía administrativa completa: {location}')
        if admins[2][codes[1]][1] != codes[0] or admins[3][codes[2]][1] != codes[1]:
            raise ValueError(f'Jerarquía EAT inconsistente con catálogo administrativo: {location}')
        cambios = []
        for n, code in enumerate(codes, 1):
            old = out[f'#adm{n}+code']
            if old != code:
                cambios.append(f'ADM{n} {old} -> {code}')
            out[f'#adm{n}+code'] = code
            out[f'#adm{n}+name'] = admins[n][code][0]
        if cambios:
            advertencias.append('Ubicación ajustada según EAT: ' + ', '.join(cambios))
        if lookup_catalogo != lookup:
            advertencias.append(f'Nombre EAT equivalente: {location} -> {nombre}')
        out['#meta+check_adm1code_adm2name'] = f"{codes[0]}-{out['#adm2+name']}"
        out['#meta+check_adm2code_adm3name'] = f"{codes[1]}-{out['#adm3+name']}"
        out['#adm4+code'], out['#adm4+name'] = code4, nombre
    elif out['#location+type'] != 'Otro' and len(validos) == 1:
        out['#adm4+code'], out['#adm4+name'] = next(iter(validos))
    elif out['#location+type'] != 'Otro':
        advertencias.append('ADM4 sin coincidencia única en la parroquia: código vacío; nombre conservado')
    gps = str(r.get('gps_loc') or '').split()
    if gps:
        try:
            lat, lon = float(gps[0]), float(gps[1])
            if not (0.5 <= lat <= 12.3 and -74.4 <= lon <= -59.6):
                raise ValueError('fuera de rango')
            out['#location+geo+latitude'], out['#location+geo+longitude'] = lat, lon
        except (ValueError, IndexError):
            advertencias.append('Coordenadas inválidas o fuera de rango: ambas en blanco')
    recurrente = a.get('recurrent_activity')
    if recurrente not in ('si', 'no'):
        raise ValueError('Recurrencia no reconocida')
    out['#activity+recurrent'] = 'Sí' if recurrente == 'si' else 'No'
    breakdown, total = desagregar(a, destino, advertencias)
    if a.get('ben_total') in (None, ''):
        advertencias.append('Total derivado de la desagregación completa; ben_total vacío')
    out.update(breakdown)
    out['#reached'] = total
    for tag, field in VULNERABILIDAD.items():
        try:
            val = numero(a.get(field), field)
        except ValueError:
            if field not in ('vuln_embarazadas', 'vuln_indigena'):
                raise
            val = 0
            advertencias.append(f'{field} inválido reemplazado por 0')
        if val is not None and val > total:
            advertencias.append(f'{field} ({val}) supera el total recalculado ({total}); valor conservado')
        out[tag] = val
    # Separados y no acompañados pueden solaparse: no se suman.
    if destino == 'Proteccion':
        sep = numero(a.get('vuln_NNA_No_Sep_001_001'), 'NNA separados')
        try:
            aco = numero(a.get('vuln_NNA_No_Aco_001'), 'NNA no acompañados')
        except ValueError:
            aco = 0
            advertencias.append('NNA no acompañados inválido reemplazado por 0')
        if sep and aco:
            advertencias.append(f'NNA separados={sep}, no acompañados={aco}; posible solapamiento: se exporta separados, sin sumar')
        out['#separados'] = sep or aco
        if out['#separados'] is not None and out['#separados'] > total:
            advertencias.append('NNA separados/no acompañados supera el total recalculado; valor conservado')
    unit = out.get('#meta+unit') or out.get('#activity+indicator+unit')
    if unit == 'Personas (no requerido)':
        unit = 'Personas'
        out['#meta+unit'] = unit
        advertencias.append('Unidad histórica Personas (no requerido) normalizada a Personas')
    if destino == 'WASH':
        out['#activity+area_type'] = 'ASH en Instituciones' if tipo in ('albergue', 'albergue_onu', 'escuela', 'centro_salud') else 'ASH Comunitario'
        if (out.get('#activity+code'), out['#activity+area_type']) not in combos:
            areas = sorted(area for code, area in combos if code == out.get('#activity+code'))
            if len(areas) != 1:
                raise ValueError(f'Actividad WASH sin área única en catálogo: {areas}')
            advertencias.append(f'Área WASH corregida: {out["#activity+area_type"]} -> {areas[0]}')
            out['#activity+area_type'] = areas[0]
    if destino == 'Transversal' and a.get('activity_name') in ('ACT_23', 'ACT_75'):
        out['#activity+indicator+reached'] = numero(a.get('ben_total'), 'ben_total')
        advertencias.append(f'Cantidad de personas Kobo usada como cantidad de {unit}, por instrucción')
    elif unit == 'Personas':
        out['#activity+indicator+reached'] = total
    elif destino == 'WASH' and unit == 'Hogares':
        out['#activity+indicator+reached'] = int(Decimal(total) / Decimal('4.5'))
        advertencias.append('Hogares estimados: piso(personas / 4.5)')
    elif destino == 'WASH' and unit == 'Instituciones':
        out['#activity+indicator+reached'] = 1
        advertencias.append('Instituciones estimadas: 1 por actividad')
    else:
        raise ValueError(f'Unidad sin cantidad capturada ni regla confirmada: {unit}')
    if destino == 'Educacion' and any(numero(a.get(f'{s}_18_59'), s) or numero(a.get(f'personal_comunitario_{s}_18_59'), s) for s in ('mujeres', 'hombres')):
        advertencias.append('Equivalencia autorizada: edades 18-59 -> 20-59')
    comentarios = [str(a.get(k) or '').strip() for k in ('activity_detail', 'reporte_cualitativo')]
    out['#activity+comments'] = ' | '.join(dict.fromkeys(c for c in comentarios if c)) or 'Sin detalle registrado en el formulario'
    for tag in ('#location+type', '#fund+type', '#reached+type', '#coord+space'):
        val = out.get(tag)
        if val is not None and tag in vocab and val not in vocab[tag]:
            raise ValueError(f'{tag}: valor fuera del catálogo 345w: {val}')
    # La plantilla no captura kits, aulas ni centros: se dejan vacíos.
    return destino, out, advertencias


def escribir_excel(path, titulo, tags, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = titulo
    ws.append(tags)
    for row in rows:
        ws.append([row.get(tag) for tag in tags])
    for row in ws:
        for cell in row:
            if isinstance(cell.value, str):
                cell.data_type = 's'  # Mantiene códigos y texto, incluso si empieza por '='.
    wb.save(path)


def escribir_csv(path, rows, columns):
    with path.open('w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=columns)
        writer.writeheader()
        for row in rows:
            # Evita que Excel interprete texto de origen como fórmula al abrir CSV.
            writer.writerow({k: "'"+v if isinstance(v, str) and v.startswith(('=', '+', '-', '@')) else v for k, v in row.items()})


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--local', action='store_true', help='Transformar la última descarga local sin consultar Kobo')
    parser.add_argument('--desde', type=date.fromisoformat, help='Fecha de reporte mínima, inclusiva (AAAA-MM-DD)')
    parser.add_argument('--hasta', type=date.fromisoformat, help='Fecha de reporte máxima, inclusiva (AAAA-MM-DD)')
    parser.add_argument('--sector', choices=SECTORES, help='Procesar sólo este formulario')
    parser.add_argument('--desde-id', help='Incluir este envío y los posteriores por fecha de envío e ID')
    parser.add_argument('--reprocesar', action='store_true', help='Incluir envíos ya transformados en esta aplicación')
    args = parser.parse_args()
    if args.desde_id and not args.sector:
        parser.error('--desde-id requiere --sector')
    if args.desde and args.hasta and args.desde > args.hasta:
        parser.error('--desde debe ser menor o igual a --hasta')
    catalogs = cargar(ROOT)
    token = None
    endpoints = {}
    if not args.local:
        token = (ROOT/'apikey.txt').read_text(encoding='utf-8-sig').strip()
        if token.lower().startswith('token '):
            token = token[6:].strip()
        if not token or any(c.isspace() for c in token):
            raise ValueError('apikey.txt debe contener sólo la clave Kobo, opcionalmente precedida por Token')
        endpoints = dict(re.findall(r'(Proteccion|WASH|Educacion)\s+(https://\S+)', (ROOT/'endpoints.txt').read_text(encoding='utf-8-sig')))
        if set(endpoints) != set(SECTORES):
            raise ValueError('endpoints.txt debe contener una URL para cada uno de los tres sectores')
    stamp = datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S_%fZ')
    salida = BASE/'salidas'/(args.sector or 'Todos')/stamp
    salida.mkdir(parents=True)
    tablas = {s: [] for s in (*SECTORES, 'Transversal')}
    revision, trazabilidad, resumen, hashes = [], [], {}, {}
    fuentes = catalogs[4]
    hashes.update({str(p.relative_to(ROOT)): hashlib.sha256(p.read_bytes()).hexdigest() for p in fuentes})
    historial_path = BASE/'historial.json'
    historial = json.loads(historial_path.read_text(encoding='utf-8')) if historial_path.exists() else {'transformaciones': [], 'descargas': []}
    incluidos_historial = []
    seleccion = {}
    for sector in ((args.sector,) if args.sector else SECTORES):
        if args.local:
            records = json.loads((BASE/'datos_kobo'/f'{sector}.json').read_text(encoding='utf-8'))
            schema = json.loads((BASE/'datos_kobo'/f'{sector}_formulario.json').read_text(encoding='utf-8'))
        elif args.sector and args.desde_id:
            indice_path = BASE/'datos_kobo'/f'{sector}_indice.json'
            indice = json.loads(indice_path.read_text(encoding='utf-8'))
            inicio = [r for r in indice if str(r.get('_id')) == args.desde_id]
            if len(inicio) != 1 or not inicio[0].get('_submission_time'):
                raise ValueError('ID inicial no encontrado en el índice; actualice el sector')
            print(f'Consultando {sector} desde ID {args.desde_id}...', flush=True)
            records = consultar(endpoints[sector], token, {'limit': 1000, 'query': json.dumps({'_submission_time': {'$gte': inicio[0]['_submission_time']}})})
            schema = json.loads((BASE/'datos_kobo'/f'{sector}_formulario.json').read_text(encoding='utf-8'))
        else:
            print(f'Descargando {sector}...', flush=True)
            records, schema = extraer(sector, endpoints[sector], token)
        for suffix in (('_indice.json', '_formulario.json') if not args.local and args.sector and args.desde_id else ('.json', '_formulario.json')):
            p = BASE/'datos_kobo'/f'{sector}{suffix}'
            hashes[str(p.relative_to(ROOT))] = hashlib.sha256(p.read_bytes()).hexdigest()
        choices = catalogo_formulario(schema)
        def orden(r):
            return (r.get('_submission_time') or '', int(r['_id']))
        records.sort(key=orden)
        if args.desde_id:
            matches = [r for r in records if str(r.get('_id')) == args.desde_id]
            if len(matches) != 1 or not matches[0].get('_submission_time'):
                raise ValueError('ID inicial no encontrado o sin fecha de envío en el sector seleccionado')
            records = [r for r in records if orden(r) >= orden(matches[0])]
        anteriores = {str(item['id']) for run in historial['transformaciones'] for item in run.get('envios', []) if run.get('sector') == sector}
        repetidos = 0
        if args.sector and not args.reprocesar:
            repetidos = sum(str(r.get('_id')) in anteriores for r in records)
            records = [r for r in records if str(r.get('_id')) not in anteriores]
        if args.sector and not records:
            raise ValueError('No hay envíos nuevos para transformar desde ese registro')
        seleccion[sector] = records
        conteo = {'envios': len(records), 'omitidos_ya_transformados': repetidos, 'actividades': 0, 'fuera_de_rango': 0, 'exportadas': 0, 'revision': 0}
        seen = set()
        for r in records:
            revision_antes = len(revision)
            antes_tablas = {s: len(rows) for s, rows in tablas.items()}
            antes_trazas = len(trazabilidad)
            repeats = r.get('repeticion_actividades', [])
            if not repeats:
                revision.append({'formulario': sector, 'envio_id': r.get('_id'), 'actividad_indice': '', 'actividad_kobo': '', 'motivo': 'Envío sin actividades'})
            for i, raw in enumerate(repeats, 1):
                conteo['actividades'] += 1
                origen = {'formulario': sector, 'envio_id': r.get('_id'), 'actividad_indice': i,
                          'actividad_kobo': raw.get('repeticion_actividades/grp_what/activity_name')}
                try:
                    key = (r.get('_id'), i)
                    if key in seen or key[0] is None:
                        raise ValueError('ID Kobo ausente o actividad repetida en la descarga')
                    seen.add(key)
                    fecha = date.fromisoformat(r.get('grp_who/date_reported', ''))
                    if args.desde and fecha < args.desde or args.hasta and fecha > args.hasta:
                        conteo['fuera_de_rango'] += 1
                        continue
                    destino, row, notas = transformar(r, raw, choices, catalogs)
                    tablas[destino].append(row)
                    trazabilidad.append({**origen, 'archivo': destino+'.xlsx', 'fila_excel': len(tablas[destino])+1,
                                         'advertencias': ' | '.join(notas)})
                    conteo['exportadas'] += 1
                except (ValueError, TypeError, KeyError, IndexError) as exc:
                    revision.append({**origen, 'motivo': str(exc)})
                    conteo['revision'] += 1
            if len(revision) > revision_antes:
                # No exportar un envío a medias: al corregirlo puede repetirse sin duplicar filas previas.
                descartadas = sum(len(tablas[s]) - antes_tablas[s] for s in tablas)
                for s in tablas:
                    del tablas[s][antes_tablas[s]:]
                del trazabilidad[antes_trazas:]
                conteo['exportadas'] -= descartadas
            if args.sector and repeats and len(revision) == revision_antes and not args.desde and not args.hasta:
                incluidos_historial.append({'id': r.get('_id'), 'uuid': r.get('_uuid'), 'submission_time': r.get('_submission_time')})
        resumen[sector] = conteo
        print(f'{sector}: {conteo}', flush=True)
    for sector, rows in tablas.items():
        if args.sector and not rows:
            continue
        titulo, tags = catalogs[0][sector]
        tags = list(tags)
        # La primera columna es un código HXL con datos; #id está vacío en las
        # cargas nuevas y algunos importadores descartan o desplazan esa columna.
        tags = [tag for tag in tags if tag and tag != '#id']
        # La validación de 345w exige estos campos aunque no estén en la plantilla antigua.
        for tag in ('#modalidad+tm', '#indicator+name', '#activity+code', '#activity+indicator+name', '#activity+indicator+unit', '#activity+indicator+reached', '#adm4+name', '#org+impl+id'):
            if tag not in tags:
                tags.append(tag)
        escribir_excel(salida/f'{sector}.xlsx', sector.lower(), tags, rows)
    origen_cols = ['formulario', 'envio_id', 'actividad_indice', 'actividad_kobo']
    escribir_csv(salida/'revision.csv', revision, origen_cols+['motivo'])
    escribir_csv(salida/'trazabilidad.csv', trazabilidad, origen_cols+['archivo', 'fila_excel', 'advertencias'])
    guardar_json(salida/'resumen.json', {'ejecucion_utc': stamp, 'modo': 'local' if args.local else 'API',
                 'desde': args.desde, 'hasta': args.hasta, 'formularios': resumen,
                 'filas_por_archivo': {s: len(r) for s,r in tablas.items()},
                 'incidencias_revision': len(revision), 'sha256_entradas': hashes})
    if args.sector:
        historial['transformaciones'].append({'sector': args.sector, 'fecha_utc': datetime.now(timezone.utc).isoformat(),
            'desde_id': args.desde_id, 'salida': str(salida), 'resumen': resumen[args.sector], 'envios': incluidos_historial})
        guardar_json(historial_path, historial)
    print(f'Salida: {salida}\nRevisar revision.csv antes de cargar en 345w.', flush=True)
    return 2 if revision else 0


if __name__ == '__main__':
    try:
        sys.exit(main())
    except (OSError, ValueError, RuntimeError) as exc:
        print(f'ERROR: {exc}', file=sys.stderr)
        sys.exit(1)
