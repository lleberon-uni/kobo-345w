"""Ventana local para descargar y transformar un formulario Kobo a la vez."""
import json
import os
from pathlib import Path
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timezone

from etl import BASE, ROOT, actualizar_indice, guardar_json
from mapeos import SECTORES

HISTORIAL = BASE/'historial.json'


def historial():
    return json.loads(HISTORIAL.read_text(encoding='utf-8')) if HISTORIAL.exists() else {'transformaciones': [], 'descargas': []}


def descargar(sector):
    token = (ROOT/'apikey.txt').read_text(encoding='utf-8-sig').strip()
    if token.lower().startswith('token '):
        token = token[6:].strip()
    if not token or any(c.isspace() for c in token):
        raise ValueError('Revise apikey.txt: debe contener solamente la clave Kobo')
    endpoints = dict(re.findall(r'(Proteccion|WASH|Educacion)\s+(https://\S+)', (ROOT/'endpoints.txt').read_text(encoding='utf-8-sig')))
    if sector not in endpoints:
        raise ValueError('Falta el endpoint del sector seleccionado')
    records = actualizar_indice(sector, endpoints[sector], token)
    hist = historial()
    hist['descargas'].append({'sector': sector, 'fecha_utc': datetime.now(timezone.utc).isoformat(), 'envios': len(records)})
    guardar_json(HISTORIAL, hist)
    return f'Descarga terminada: {len(records)} envíos.'


def edad(fecha):
    dt = datetime.fromisoformat(fecha.replace('Z', '+00:00'))
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    dias = max(0, (datetime.now(timezone.utc)-dt).days)
    return 'hoy' if dias == 0 else f'hace {dias} día(s)'


class Aplicacion:
    def __init__(self, root):
        self.root = root
        self.root.title('Kobo a 345w')
        self.root.geometry('760x500')
        self.events = queue.Queue()
        self.busy = False
        self.sector = tk.StringVar(value='Proteccion')
        self.registro = tk.StringVar()
        self.reprocesar = tk.BooleanVar(value=False)
        self.estado = tk.StringVar(value='Seleccione un sector.')
        self.info = tk.StringVar()
        self.descarga_info = tk.StringVar()
        self.ids = {}
        self.salida = None
        frame = ttk.Frame(root, padding=20)
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text='Actualizar datos de Kobo para 345w', font=('Segoe UI', 15, 'bold')).pack(anchor='w', pady=(0, 16))
        ttk.Label(frame, text='1. Seleccione el formulario').pack(anchor='w')
        self.seleccion = ttk.Combobox(frame, textvariable=self.sector, values=SECTORES, state='readonly', width=30)
        self.seleccion.pack(anchor='w', pady=4)
        self.seleccion.bind('<<ComboboxSelected>>', lambda _: self.refrescar())
        self.btn_descarga = ttk.Button(frame, text='Descargar / actualizar datos del sector', command=self.iniciar_descarga)
        self.btn_descarga.pack(anchor='w', pady=(8, 3))
        ttk.Label(frame, textvariable=self.descarga_info, wraplength=700).pack(anchor='w')
        ttk.Separator(frame).pack(fill='x', pady=12)
        ttk.Label(frame, text='2. Seleccione el registro inicial (incluido)').pack(anchor='w')
        self.lista = ttk.Combobox(frame, textvariable=self.registro, state='readonly', width=65)
        self.lista.pack(anchor='w', pady=4)
        ttk.Label(frame, textvariable=self.info, wraplength=700).pack(anchor='w', pady=4)
        self.check = ttk.Checkbutton(frame, text='Incluir envíos ya transformados (puede generar duplicados)', variable=self.reprocesar)
        self.check.pack(anchor='w', pady=4)
        buttons = ttk.Frame(frame)
        buttons.pack(anchor='w', pady=12)
        self.btn_exportar = ttk.Button(buttons, text='Transformar desde este registro', command=self.iniciar_exportacion)
        self.btn_exportar.pack(side='left')
        self.btn_abrir = ttk.Button(buttons, text='Abrir resultados', command=self.abrir, state='disabled')
        self.btn_abrir.pack(side='left', padx=10)
        ttk.Label(frame, textvariable=self.estado, wraplength=700).pack(anchor='w', pady=4)
        self.refrescar()
        root.after(100, self.procesar_eventos)

    def refrescar(self):
        sector = self.sector.get()
        cache = BASE/'datos_kobo'/f'{sector}_indice.json'
        try:
            records = json.loads(cache.read_text(encoding='utf-8')) if cache.exists() else []
            records = sorted((r for r in records if r.get('_id') is not None and r.get('_submission_time')), key=lambda r:(r['_submission_time'], int(r['_id'])))
            self.ids = {f"{r['_id']} — {r['_submission_time']} UTC": str(r['_id']) for r in records}
            self.lista['values'] = list(self.ids)
            self.registro.set(next(iter(self.ids), ''))
            h = historial()
            downloads = [d for d in h['descargas'] if d['sector'] == sector]
            self.descarga_info.set('Última descarga: ' + (downloads[-1]['fecha_utc'][:19].replace('T',' ')+' UTC' if downloads else 'sin fecha registrada; actualice el sector.'))
            runs = [x for x in h['transformaciones'] if x['sector'] == sector and x.get('envios')]
            if runs:
                run = runs[-1]; first = run['envios'][0]; last = run['envios'][-1]
                self.info.set(f"Última transformación: {edad(run['fecha_utc'])}. Desde ID {first['id']} ({first['submission_time']} UTC), hasta ID {last['id']} ({last['submission_time']} UTC).")
                pendientes = [r for r in records if str(r['_id']) not in {str(i['id']) for x in runs for i in x['envios']}]
                if pendientes:
                    self.registro.set(f"{pendientes[0]['_id']} — {pendientes[0]['_submission_time']} UTC")
            else:
                self.info.set('No hay transformaciones registradas para este sector en esta aplicación.')
        except (OSError, ValueError, KeyError) as e:
            self.estado.set(f'No se pudo leer los datos locales: {e}')

    def bloquear(self, flag):
        self.busy = flag
        state = 'disabled' if flag else 'normal'
        self.btn_descarga.configure(state=state)
        self.btn_exportar.configure(state=state)
        self.check.configure(state=state)
        self.seleccion.configure(state='disabled' if flag else 'readonly')
        self.lista.configure(state='disabled' if flag else 'readonly')

    def trabajo(self, fn):
        def run():
            try:
                self.events.put(('ok', fn()))
            except Exception as e:
                self.events.put(('error', str(e)))
        self.bloquear(True)
        threading.Thread(target=run, daemon=True).start()

    def iniciar_descarga(self):
        sector = self.sector.get()
        self.estado.set(f'Descargando {sector} desde Kobo. Espere…')
        self.trabajo(lambda: descargar(sector))

    def iniciar_exportacion(self):
        ident = self.ids.get(self.registro.get())
        if not ident:
            messagebox.showwarning('Seleccione un registro', 'Actualice el sector y seleccione el registro inicial.')
            return
        sector = self.sector.get()
        opciones = [sys.executable, str(BASE/'etl.py'), '--sector', sector, '--desde-id', ident]
        if self.reprocesar.get():
            opciones.append('--reprocesar')
        self.estado.set(f'Transformando {sector} desde ID {ident}. Espere…')
        def ejecutar():
            resultado = subprocess.run(opciones, cwd=str(ROOT), text=True, encoding='utf-8', errors='replace', capture_output=True, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
            if resultado.returncode not in (0, 2):
                raise ValueError(resultado.stderr.strip() or resultado.stdout.strip() or 'No se pudo transformar')
            match = re.search(r'^Salida: (.+)$', resultado.stdout, re.MULTILINE)
            if not match:
                raise ValueError('No se identificó la carpeta de resultados')
            self.events.put(('salida', match.group(1).strip()))
            summary = json.loads((Path(match.group(1).strip())/'resumen.json').read_text(encoding='utf-8'))['formularios'][sector]
            return (f"Transformación terminada: {summary['actividades']} actividades; "
                    f"{summary['omitidos_ya_transformados']} envíos ya transformados omitidos; "
                    f"{summary['revision']} casos en revisión.")
        self.trabajo(ejecutar)

    def procesar_eventos(self):
        try:
            while True:
                kind, value = self.events.get_nowait()
                if kind == 'salida':
                    self.salida = value
                    self.btn_abrir.configure(state='normal')
                elif kind == 'ok':
                    self.bloquear(False); self.refrescar(); self.estado.set(value)
                elif kind == 'error':
                    self.bloquear(False); self.estado.set('No se completó la operación.'); messagebox.showerror('No se pudo completar', value)
        except queue.Empty:
            pass
        self.root.after(100, self.procesar_eventos)

    def abrir(self):
        if self.salida:
            os.startfile(self.salida)


if __name__ == '__main__':
    root = tk.Tk()
    Aplicacion(root)
    root.mainloop()
