import os
import re
import datetime
from copy import copy
from collections import defaultdict
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill
import urllib.request
import json
import time
import webbrowser

from rutinas import menu_rutinas
from analisis_a2admin import menu_analisis

BASE = os.path.dirname(os.path.abspath(__file__))
CARPETA_SOPORTES = os.path.join(BASE, "SOPORTES DE ZOOM")
CARPETA_GUIA = os.path.join(BASE, "GUIA PARA TRABAJAR")
CARPETA_SALIDAS = os.path.join(BASE, "Salidas")
RUTA_ENVIOS = os.path.join(CARPETA_GUIA, "Envios 2026 2.xlsx")

RESET = "\033[0m"
ROJO = "\033[1;31m"
VERDE = "\033[1;32m"
CYAN = "\033[1;36m"
AMARILLO = "\033[1;33m"
MAGENTA = "\033[1;35m"
AZUL = "\033[1;34m"
BLANCO = "\033[1;37m"

MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
         "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]

ENCABEZADOS = [None, 'Fecha ', 'Vendedor', 'Cliente', '#Factura ZM', '$ Factura ZM',
               'Peso KG', '#Guia ', 'Piezas    (Cajas)', 'Dia de salida (Entrega a zoom)',
               'Dia de llegada (Entrega a cliente)', 'Total de dias para la entrega', 'Factura Zoom',
               None, 'Costo del envio', 'Tasa a recibir el dólar ZM', 'Promedio $/costo de envio',
               'Precio P/Kg', None, None]

EXT_TEXTO = ('.txt', '.csv', '.tsv')
ENCABEZADO_ZOOM = {'envio': 'ENVIO', 'fecha': 'FECHA', 'cliente': 'DESTINATARIO',
                   'peso': 'PESO', 'fpo': 'FPO', 'total': 'TOTAL'}
FORMATO_FECHA_ZOOM = ('%d-%m-%Y', '%d-%m-%y', '%d/%m/%Y', '%d/%m/%y', '%Y-%m-%d',
                      '%Y/%m/%d', '%d.%m.%Y')


def limpiar_pantalla():
    os.system('cls' if os.name == 'nt' else 'clear')


def a_float(valor):
    if valor is None:
        return None
    s = str(valor).strip()
    if not s:
        return None
    s = s.replace('BS', '').replace('Bs', '').replace('bs', '').replace(' ', '')
    try:
        return float(s)
    except ValueError:
        pass
    if '.' in s and ',' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s:
        s = s.replace(',', '.')
    else:
        s = s.replace('.', '')
    try:
        return float(s)
    except ValueError:
        return None


def sin_acentos(texto):
    try:
        import unicodedata
        return ''.join(c for c in unicodedata.normalize('NFKD', str(texto))
                       if not unicodedata.combining(c))
    except Exception:
        return str(texto)


def normalizar_encabezado(texto):
    return re.sub(r'[^A-Z0-9]', '', sin_acentos(texto).upper())


def decodificar_texto(ruta):
    with open(ruta, 'rb') as f:
        datos = f.read()
    for cod in ('utf-8-sig', 'utf-16', 'cp1252', 'latin-1'):
        try:
            return datos.decode(cod)
        except (UnicodeDecodeError, LookupError):
            continue
    return datos.decode('latin-1', errors='ignore')


def a_fecha(valor):
    if valor is None or str(valor).strip() == '':
        return None
    if isinstance(valor, datetime.datetime):
        return valor
    if isinstance(valor, datetime.date):
        return datetime.datetime(valor.year, valor.month, valor.day)
    s = str(valor).strip()
    for fmt in FORMATO_FECHA_ZOOM:
        try:
            return datetime.datetime.strptime(s, fmt)
        except ValueError:
            continue
    return s


def buscar_columna(mapa, nombre):
    if nombre in mapa:
        return mapa[nombre]
    for clave, indice in mapa.items():
        if clave.startswith(nombre):
            return indice
    return None


def buscar_encabezado_zoom(lineas):
    """Busca la linea de encabezados de Zoom y devuelve (separador, indice, mapa)."""
    for sep in ('\t', ';', '|', ','):
        for i, linea in enumerate(lineas[:20]):
            codigos = [normalizar_encabezado(c.strip().strip('"')) for c in linea.split(sep)]
            if 'TOTAL' in codigos and any(c.startswith('ENVIO') for c in codigos):
                return sep, i, {c: j for j, c in enumerate(codigos) if c}
    return None, None, {}


def leer_soporte_zoom_txt(ruta):
    texto = decodificar_texto(ruta)
    lineas = texto.splitlines()
    sep, fila_encabezado, mapa = buscar_encabezado_zoom(lineas)
    if fila_encabezado is None:
        raise ValueError("El archivo de texto no tiene el encabezado de Zoom "
                         "(se espera una linea con N, ENVIO, FECHA, DESTINATARIO, PESO, FPO, TOTAL)")

    col_envio = buscar_columna(mapa, ENCABEZADO_ZOOM['envio'])
    col_fecha = buscar_columna(mapa, ENCABEZADO_ZOOM['fecha'])
    col_cliente = buscar_columna(mapa, ENCABEZADO_ZOOM['cliente'])
    col_peso = buscar_columna(mapa, ENCABEZADO_ZOOM['peso'])
    col_fpo = buscar_columna(mapa, ENCABEZADO_ZOOM['fpo'])
    col_total = buscar_columna(mapa, ENCABEZADO_ZOOM['total'])
    faltan = [nombre for nombre, col in (('ENVIO', col_envio), ('TOTAL', col_total))
              if col is None]
    if faltan:
        raise ValueError(f"Faltan columnas en el archivo de texto: {', '.join(faltan)}")

    filas = []
    for linea in lineas[fila_encabezado + 1:]:
        if not linea.strip():
            continue
        celdas = [c.strip().strip('"') for c in linea.split(sep)]

        def celda(indice):
            if indice is None or indice >= len(celdas):
                return None
            return celdas[indice]

        envio = celda(col_envio)
        total = a_float(celda(col_total))
        if envio is None or str(envio).strip() == '' or total is None:
            continue
        if str(envio).strip().isdigit():
            envio = int(envio)
        filas.append({
            'envio': envio,
            'fecha': a_fecha(celda(col_fecha)),
            'cliente': version_corta(celda(col_cliente)),
            'peso': a_float(celda(col_peso)),
            'fpo': a_float(celda(col_fpo)),
            'total': total,
        })
    return filas


def version_corta(nombre):
    nombre = str(nombre).strip()
    if not nombre:
        return nombre
    nombre = re.sub(r'\bC\s*\.?\s*A\s*\.?\b', '', nombre, flags=re.I)
    nombre = re.sub(r'\bC\s*,\s*A\b', '', nombre, flags=re.I)
    nombre = re.sub(r'\bS\s*\.?\s*A\s*\.?\b', '', nombre, flags=re.I)
    palabras = [p for p in re.split(r'\s+|\s*,\s*', nombre) if p]
    if len(palabras) > 3:
        palabras = palabras[-3:]
    return ' '.join(palabras).strip().rstrip(',. ')


def extraer_codigo_factura(nombre):
    m = re.search(r'(f-\d+)', nombre, re.I)
    return m.group(1).upper() if m else None


def consultar_api_zoom(guia, reintentos=2):
    url = f'https://api.zoom.red/canguroazul/getZoomTrackWs?tipo_busqueda=1&web=1&codigo={guia}'
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    for intento in range(reintentos + 1):
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                return json.loads(r.read().decode('utf-8-sig'))
        except Exception:
            if intento < reintentos:
                time.sleep(1.5)
    return None


def extraer_fechas_zoom(data):
    ent = (data or {}).get('entidadRespuesta') or {}
    iz = ent.get('infoZoom') or {}
    track = ent.get('tracking') or []
    fecha_salida = iz.get('fecha')
    fecha_llegada = None
    fecha_transito = None
    for t in track:
        nom = ((t.get('estatus') or {}).get('nombre') or '').upper()
        raw = t.get('fechahorareal') or ''
        dia = raw.split(' ')[0] if raw else None
        if 'ENTREGADO' in nom and fecha_llegada is None:
            fecha_llegada = dia
        if 'EN TRANSITO' in nom and dia:
            if fecha_transito is None or dia < fecha_transito:
                fecha_transito = dia
    if fecha_transito:
        fecha_salida = fecha_transito
    return fecha_salida, fecha_llegada


def seleccionar_soporte():
    from tkinter import filedialog, Tk
    try:
        root = Tk()
        root.withdraw()
        root.update()
        ruta = filedialog.askopenfilename(
            title="Selecciona el soporte de Zoom",
            initialdir=CARPETA_SOPORTES if os.path.isdir(CARPETA_SOPORTES) else BASE,
            filetypes=[("Soportes de Zoom", "*.xlsx *.txt *.csv *.tsv"),
                       ("Archivos Excel", "*.xlsx"),
                       ("Archivos de texto", "*.txt *.csv *.tsv"),
                       ("Todos", "*.*")])
        root.destroy()
        if ruta:
            return ruta
    except Exception as e:
        print(f"{AMARILLO}No se pudo abrir el explorador: {e}{RESET}")
    if os.path.isdir(CARPETA_SOPORTES):
        soportes = [os.path.join(CARPETA_SOPORTES, f) for f in os.listdir(CARPETA_SOPORTES)
                    if f.lower().endswith(('.xlsx',) + EXT_TEXTO) and not f.startswith('~$')]
        if soportes:
            return max(soportes, key=os.path.getmtime)
    return None


def leer_soporte_zoom(ruta):
    if os.path.splitext(ruta)[1].lower() in EXT_TEXTO:
        return leer_soporte_zoom_txt(ruta)
    return leer_soporte_zoom_xlsx(ruta)


def leer_soporte_zoom_xlsx(ruta):
    wb = load_workbook(ruta, data_only=False)
    ws = wb.active
    filas = []
    for fila in ws.iter_rows(min_row=2):
        envio = fila[1].value
        if envio is None or str(envio).strip() == '':
            continue
        fecha = fila[4].value
        destinatario = fila[6].value
        peso = a_float(fila[9].value)
        fpo = a_float(fila[16].value)
        total = a_float(fila[17].value)
        if total is None:
            continue
        filas.append({
            'envio': envio,
            'fecha': fecha,
            'cliente': version_corta(destinatario),
            'peso': peso,
            'fpo': fpo,
            'total': total,
        })
    wb.close()
    return filas


def pedir_tasa():
    while True:
        val = input(f"{BLANCO}Ingresa la tasa del dolar (ej: 920): {RESET}").strip()
        try:
            tasa = float(val)
            if tasa > 0:
                return tasa
        except ValueError:
            pass
        print(f"{ROJO}Valor invalido, intenta de nuevo.{RESET}")


def copiar_formato(ws_template):
    estilo_datos = {}
    for col in range(1, ws_template.max_column + 1):
        c = ws_template.cell(row=3, column=col)
        estilo_datos[col] = {
            'font': copy(c.font),
            'number_format': c.number_format,
            'alignment': copy(c.alignment),
            'border': copy(c.border),
            'fill': copy(c.fill),
        }
    estilo_total = {}
    for col in range(1, ws_template.max_column + 1):
        c = ws_template.cell(row=30, column=col)
        estilo_total[col] = {
            'font': copy(c.font),
            'number_format': c.number_format,
            'alignment': copy(c.alignment),
            'border': copy(c.border),
            'fill': copy(c.fill),
        }
    combinadas_total = []
    for mr in ws_template.merged_cells.ranges:
        if mr.min_row == 30:
            combinadas_total.append((mr.min_col, mr.max_col))
    return estilo_datos, estilo_total, combinadas_total


def nombre_mes_actual():
    return MESES[datetime.date.today().month - 1]


def anio_actual():
    return datetime.date.today().year


def ruta_maestro_anio(anio):
    return os.path.join(CARPETA_GUIA, f"Envios {anio} 2.xlsx")


def crear_maestro_anual(anio, ruta_nuevo):
    anio_prev = anio - 1
    ruta_prev = ruta_maestro_anio(anio_prev)
    if not os.path.exists(ruta_prev):
        candidatos = [os.path.join(CARPETA_GUIA, f)
                      for f in os.listdir(CARPETA_GUIA)
                      if re.match(r"Envios \d{4} 2\.xlsx$", f)]
        if candidatos:
            ruta_prev = max(candidatos, key=os.path.getmtime)
        else:
            ruta_prev = RUTA_ENVIOS
    import shutil
    shutil.copy2(ruta_prev, ruta_nuevo)
    wb = load_workbook(ruta_nuevo, data_only=False)
    ref = wb["SEPTIEMBRE"] if "SEPTIEMBRE" in wb.sheetnames else wb.active
    for nombre_mes in MESES:
        if nombre_mes in wb.sheetnames:
            indice = wb.index(wb[nombre_mes])
            del wb[nombre_mes]
            crear_hoja_mes(wb, nombre_mes, ref=ref, indice=indice)
    wb.save(ruta_nuevo)
    wb.close()
    return ruta_nuevo


def obtener_ruta_maestro():
    anio = anio_actual()
    ruta = ruta_maestro_anio(anio)
    if not os.path.exists(ruta):
        crear_maestro_anual(anio, ruta)
    return ruta


def hoja_referencia_estructura(wb):
    nombre = nombre_mes_actual()
    if nombre in wb.sheetnames:
        return wb[nombre]
    if "SEPTIEMBRE" in wb.sheetnames:
        return wb["SEPTIEMBRE"]
    for m in reversed(MESES):
        if m in wb.sheetnames and wb[m].max_row > 1:
            return wb[m]
    return wb.active


def _posicion_mes(wb, nombre_mes):
    indice = MESES.index(nombre_mes)
    pos = len(wb.sheetnames)
    ultimo_mes = -1
    for i, nombre in enumerate(wb.sheetnames):
        if nombre in MESES:
            idx = MESES.index(nombre)
            if idx > indice:
                if pos > i:
                    pos = i
            elif idx > ultimo_mes:
                ultimo_mes = i
    if ultimo_mes >= 0 and pos > ultimo_mes + 1:
        pos = ultimo_mes + 1
    return pos


def crear_hoja_mes(wb, nombre_mes, ref=None, indice=None):
    if ref is None:
        ref = hoja_referencia_estructura(wb)
    pos = indice if indice is not None else _posicion_mes(wb, nombre_mes)
    ws = wb.create_sheet(title=nombre_mes, index=pos)
    for col in range(1, ref.max_column + 1):
        for r in (1, 2):
            src = ref.cell(row=r, column=col)
            dst = ws.cell(row=r, column=col, value=src.value)
            dst.font = copy(src.font)
            dst.number_format = src.number_format
            dst.alignment = copy(src.alignment)
            dst.border = copy(src.border)
            dst.fill = copy(src.fill)
    for dim in ref.column_dimensions.values():
        ws.column_dimensions[dim.index].width = dim.width
    for mr in ref.merged_cells.ranges:
        if mr.min_row == 1:
            ws.merge_cells(start_row=mr.min_row, start_column=mr.min_col,
                           end_row=mr.max_row, end_column=mr.max_col)
    return ws


def obtener_hoja_mes(wb, nombre_mes):
    if nombre_mes in wb.sheetnames:
        return wb[nombre_mes]
    return crear_hoja_mes(wb, nombre_mes)


def aplicar_estilo(ws, fila, estilo):
    for col, fmt in estilo.items():
        c = ws.cell(row=fila, column=col)
        c.font = copy(fmt['font'])
        c.number_format = fmt['number_format']
        c.alignment = copy(fmt['alignment'])
        c.border = copy(fmt['border'])
        c.fill = copy(fmt['fill'])


def escribir_grupo(ws, fila_inicio, datos, tasa, estilos=None, codigo_factura=None):
    estilo_datos, estilo_total, combinadas_total = estilos if estilos else (None, None, None)
    f = fila_inicio
    for i, d in enumerate(datos, start=1):
        if estilo_datos:
            aplicar_estilo(ws, f, estilo_datos)
        ws.cell(row=f, column=1, value=i)
        if d['fecha'] is not None:
            c = ws.cell(row=f, column=2, value=d['fecha'])
            c.number_format = 'DD/MM/YYYY'
        ws.cell(row=f, column=4, value=d['cliente'])
        if d['peso'] is not None:
            ws.cell(row=f, column=7, value=d['peso'])
        ws.cell(row=f, column=8, value=d['envio'])
        if codigo_factura:
            ws.cell(row=f, column=13, value=codigo_factura)
        ws.cell(row=f, column=15, value=f"=S{f}/P{f}")
        ws.cell(row=f, column=16, value=tasa)
        ws.cell(row=f, column=17, value=f"=O{f}/F{f}")
        ws.cell(row=f, column=18, value=f"=O{f}/G{f}")
        ws.cell(row=f, column=19, value=d['total'])
        f += 1
    if estilo_total:
        aplicar_estilo(ws, f, estilo_total)
    ws.cell(row=f, column=14, value='TOTAL')
    ws.cell(row=f, column=6, value=f"=SUM(F{fila_inicio}:F{f-1})")
    ws.cell(row=f, column=7, value=f"=SUM(G{fila_inicio}:G{f-1})")
    ws.cell(row=f, column=15, value=f"=S{f}/P{f}")
    ws.cell(row=f, column=16, value=tasa)
    ws.cell(row=f, column=17, value=f"=O{f}/F{f}")
    ws.cell(row=f, column=18, value=f"=O{f}/G{f}")
    subtotal = sum(d['total'] for d in datos if d['total'] is not None)
    fpo_total = sum(d['fpo'] for d in datos if d['fpo'] is not None)
    iva = round((subtotal - fpo_total) * 0.16, 2)
    total_factura = round(subtotal + iva, 2)
    ws.cell(row=f, column=19, value=total_factura)
    if combinadas_total:
        for c_min, c_max in combinadas_total:
            ws.merge_cells(start_row=f, start_column=c_min,
                           end_row=f, end_column=c_max)
    return f + 1


def guardar_con_fallo_silencioso(wb, ruta):
    try:
        wb.save(ruta)
        return ruta
    except PermissionError:
        import time
        base, ext = os.path.splitext(ruta)
        ruta2 = f"{base}_nuevo_{time.strftime('%H%M%S')}{ext}"
        wb.save(ruta2)
        return ruta2


def generar_temporal(datos, tasa, estilo=None, codigo_factura=None):
    hoy = datetime.date.today()
    nombre = f"Archivo temporal de zoom del dia {hoy.day:02d}-{hoy.month:02d}-{hoy.year}.xlsx"
    ruta_salida = os.path.join(CARPETA_GUIA, nombre)

    wb = Workbook()
    ws = wb.active
    ws.title = MESES[hoy.month - 1]
    ws["B1"] = "ZOOM"
    ws["B1"].font = Font(bold=True)
    for i, titulo in enumerate(ENCABEZADOS, start=1):
        if titulo is not None:
            ws.cell(row=2, column=i, value=titulo).font = Font(bold=True)
    escribir_grupo(ws, 3, datos, tasa, estilo, codigo_factura)
    ruta_final = guardar_con_fallo_silencioso(wb, ruta_salida)
    wb.close()
    return ruta_final, len(datos)


def generar_salida(datos, tasa, estilos=None, codigo_factura=None):
    hoy = datetime.date.today()
    os.makedirs(CARPETA_SALIDAS, exist_ok=True)
    nombre = f"Envios {hoy.year} {hoy.day:02d}-{hoy.month:02d}-{hoy.year}.xlsx"
    ruta_salida = os.path.join(CARPETA_SALIDAS, nombre)

    ruta_maestro_base = obtener_ruta_maestro()
    wb = load_workbook(ruta_maestro_base, data_only=False)
    ws = obtener_hoja_mes(wb, nombre_mes_actual())
    ultima_fila = ws.max_row
    while ultima_fila > 0 and all(ws.cell(row=ultima_fila, column=c).value is None
                                  for c in range(1, ws.max_column + 1)):
        ultima_fila -= 1
    escribir_grupo(ws, ultima_fila + 1, datos, tasa, estilos, codigo_factura)
    ruta_final = guardar_con_fallo_silencioso(wb, ruta_salida)
    ruta_maestro = guardar_con_fallo_silencioso(wb, ruta_maestro_base)
    wb.close()
    return ruta_final, len(datos), ruta_maestro


def consultar_envios_zoom(guias):
    print(f"{CYAN}Consultando API de Zoom para dias de salida/llegada y piezas...{RESET}")
    resultados = {}
    for guia in guias:
        guia = str(int(guia)) if isinstance(guia, (int, float)) else str(guia).strip()
        print(f"  {AMARILLO}Guia {guia}...{RESET}", end=" ", flush=True)
        data = consultar_api_zoom(guia)
        if not data:
            print(f"{ROJO}sin respuesta{RESET}")
            continue
        cod = data.get('codrespuesta') or ''
        if cod != 'COD_000':
            print(f"{ROJO}{cod}{RESET}")
            continue
        ent = data.get('entidadRespuesta') or {}
        iz = ent.get('infoZoom') or {}
        piezas = iz.get('nropiezas')
        try:
            piezas = int(piezas)
        except (TypeError, ValueError):
            piezas = None
        if piezas is None or piezas < 1:
            piezas = 1
        fecha_salida, fecha_llegada = extraer_fechas_zoom(data)
        if fecha_salida or fecha_llegada:
            resultados[guia] = {'piezas': piezas,
                                'salida': fecha_salida,
                                'llegada': fecha_llegada}
        print(f"{VERDE}piezas {piezas} | salida {fecha_salida} | llegada {fecha_llegada}{RESET}")
    return resultados


def aplicar_datos_zoom(ruta, resultados, codigo_factura=None, nombre_hoja=None):
    wb = load_workbook(ruta, data_only=False)
    if nombre_hoja and nombre_hoja in wb.sheetnames:
        ws = wb[nombre_hoja]
    else:
        ws = wb.active
    formato = '%Y-%m-%d'
    for fila in range(3, ws.max_row + 1):
        guia = ws.cell(row=fila, column=8).value
        if guia is None or str(guia).strip() == '':
            continue
        guia = str(int(guia)) if isinstance(guia, (int, float)) else str(guia).strip()
        if codigo_factura:
            actual = ws.cell(row=fila, column=13).value
            if actual is None or str(actual).strip() == '':
                ws.cell(row=fila, column=13, value=codigo_factura)
        res = resultados.get(guia)
        if not res:
            continue
        ws.cell(row=fila, column=9, value=res.get('piezas'))
        fecha_salida = res.get('salida')
        fecha_llegada = res.get('llegada')
        if fecha_salida:
            try:
                c = ws.cell(row=fila, column=10, value=datetime.datetime.strptime(fecha_salida, formato))
                c.number_format = 'DD/MM/YYYY'
            except ValueError:
                ws.cell(row=fila, column=10, value=fecha_salida)
        if fecha_llegada:
            try:
                c = ws.cell(row=fila, column=11, value=datetime.datetime.strptime(fecha_llegada, formato))
                c.number_format = 'DD/MM/YYYY'
            except ValueError:
                ws.cell(row=fila, column=11, value=fecha_llegada)
        if fecha_salida and fecha_llegada:
            try:
                d1 = datetime.datetime.strptime(fecha_salida, formato)
                d2 = datetime.datetime.strptime(fecha_llegada, formato)
                ws.cell(row=fila, column=12, value=(d2 - d1).days)
            except ValueError:
                pass
    if ws.max_row >= 3:
        primera = None
        ultima = None
        for fila in range(3, ws.max_row + 1):
            if ws.cell(row=fila, column=8).value is not None:
                if primera is None:
                    primera = fila
                ultima = fila
        if primera is not None:
            rango = f"L{primera}:L{ultima}"
            ws.conditional_formatting.add(
                rango,
                CellIsRule(operator='greaterThan', formula=['2'],
                           fill=PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid'),
                           font=Font(color='9C0006')))
            ws.conditional_formatting.add(
                rango,
                CellIsRule(operator='between', formula=['1', '2'],
                           fill=PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid'),
                           font=Font(color='006100')))
            ws.conditional_formatting.add(
                rango,
                CellIsRule(operator='equal', formula=['0'],
                           fill=PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid'),
                           font=Font(color='9C6500')))
    guardar_con_fallo_silencioso(wb, ruta)
    wb.close()
    print(f"{VERDE}Columnas I, J, K, L y M actualizadas en {os.path.basename(ruta)}.{RESET}")


def consolidar_codigos_excel():
    """Consolidar códigos desde un Excel seleccionando manualmente archivo, hoja y columna."""
    from tkinter import filedialog, Tk
    try:
        limpiar_pantalla()
        print(f"{CYAN}=== CONSOLIDAR CODIGOS DESDE EXCEL ==={RESET}")

        # Seleccionar archivo
        root = Tk()
        root.withdraw()
        root.update()
        ruta = filedialog.askopenfilename(
            title="Selecciona el archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xlsm *.xltx"),
                       ("Todos los archivos", "*.*")])
        root.destroy()

        if not ruta:
            print(f"{ROJO}No se seleccionó ningún archivo.{RESET}")
            input(f"{BLANCO}Presiona Enter para continuar...{RESET}")
            return

        print(f"{AZUL}Archivo seleccionado: {os.path.basename(ruta)}{RESET}")

        # Cargar workbook
        try:
            wb = load_workbook(ruta, data_only=True)
        except Exception as e:
            print(f"{ROJO}Error al abrir el archivo: {e}{RESET}")
            input(f"{BLANCO}Presiona Enter para continuar...{RESET}")
            return

        # Mostrar hojas disponibles
        if not wb.sheetnames:
            print(f"{ROJO}El archivo no tiene hojas.{RESET}")
            wb.close()
            input(f"{BLANCO}Presiona Enter para continuar...{RESET}")
            return

        print(f"\n{CYAN}Hojas disponibles:{RESET}")
        for i, nombre in enumerate(wb.sheetnames, 1):
            ws_tmp = wb[nombre]
            print(f"{AZUL}  {i}. {AMARILLO}{nombre}{BLANCO} ({ws_tmp.max_row} filas x {ws_tmp.max_column} columnas){RESET}")

        # Seleccionar hoja
        sel_hoja = input(f"\n{BLANCO}Selecciona el número de hoja (1-{len(wb.sheetnames)}): {RESET}").strip()
        try:
            idx_hoja = int(sel_hoja)
            if idx_hoja < 1 or idx_hoja > len(wb.sheetnames):
                raise ValueError
        except ValueError:
            print(f"{ROJO}Selección inválida.{RESET}")
            wb.close()
            input(f"{BLANCO}Presiona Enter para continuar...{RESET}")
            return

        nombre_hoja = wb.sheetnames[idx_hoja - 1]
        ws = wb[nombre_hoja]
        print(f"{VERDE}Hoja seleccionada: {nombre_hoja}{RESET}")

        # Mostrar encabezados/primeras filas para ayudar a elegir columna
        print(f"\n{CYAN}Vista previa (primeras 10 filas):{RESET}")
        max_col_mostrar = min(ws.max_column, 15)
        for r in range(1, min(ws.max_row + 1, 11)):
            fila_str = []
            for c in range(1, max_col_mostrar + 1):
                val = ws.cell(r, c).value
                fila_str.append(f"{c}:{str(val)[:15] if val is not None else ''}")
            print(f"  F{r}: " + " | ".join(fila_str))
        if ws.max_column > 15:
            print(f"  ... ({ws.max_column - 15} columnas más)")

        # Seleccionar columna de códigos
        sel_col = input(f"\n{BLANCO}Ingresa el NÚMERO de columna donde están los CÓDIGOS (ej: 1 para A): {RESET}").strip()
        try:
            col_codigo = int(sel_col)
            if col_codigo < 1 or col_codigo > ws.max_column:
                raise ValueError
        except ValueError:
            print(f"{ROJO}Número de columna inválido.{RESET}")
            wb.close()
            input(f"{BLANCO}Presiona Enter para continuar...{RESET}")
            return

        # Leer y consolidar (mantener orden de primera aparición)
        consolidado = defaultdict(float)
        orden_aparicion = []
        filas_leidas = 0
        filas_con_codigo = 0

        print(f"\n{CYAN}Procesando...{RESET}")
        for fila in range(1, ws.max_row + 1):
            codigo = ws.cell(fila, col_codigo).value
            if codigo is None or str(codigo).strip() == '':
                continue
            codigo_str = str(codigo).strip()

            # Obtener cantidad de la columna a la derecha (col_codigo + 1)
            col_cantidad = col_codigo + 1
            cantidad = None
            if col_cantidad <= ws.max_column:
                cantidad = ws.cell(fila, col_cantidad).value

            # Convertir cantidad a float
            if cantidad is None:
                cantidad_val = 0.0
            else:
                try:
                    # Usar la función a_float existente
                    cantidad_val = a_float(cantidad) if isinstance(cantidad, str) else float(cantidad)
                    if cantidad_val is None:
                        cantidad_val = 0.0
                except Exception:
                    cantidad_val = 0.0

            if codigo_str not in consolidado:
                orden_aparicion.append(codigo_str)
            consolidado[codigo_str] += cantidad_val
            filas_con_codigo += 1
            filas_leidas += 1

        print(f"{VERDE}Filas leídas con código: {filas_con_codigo}{RESET}")
        print(f"{VERDE}Códigos únicos encontrados: {len(consolidado)}{RESET}")

        # Crear nuevo Excel con resultado
        wb_salida = Workbook()
        ws_salida = wb_salida.active
        ws_salida.title = "Consolidado"

        # Encabezados
        ws_salida.cell(row=1, column=1, value="Código")
        ws_salida.cell(row=1, column=2, value="Cantidad Total")
        for c in (1, 2):
            ws_salida.cell(row=1, column=c).font = Font(bold=True)

        # Escribir datos manteniendo orden original
        for idx, codigo in enumerate(orden_aparicion, start=2):
            total = consolidado[codigo]
            ws_salida.cell(row=idx, column=1, value=codigo)
            if total != int(total):
                ws_salida.cell(row=idx, column=2, value=total)
            else:
                ws_salida.cell(row=idx, column=2, value=int(total) if abs(total - int(total)) < 0.0001 else total)

        # Autoajustar columnas
        ws_salida.column_dimensions['A'].width = 20
        ws_salida.column_dimensions['B'].width = 15

        # Guardar
        directorio = os.path.dirname(ruta)
        nombre_base = os.path.splitext(os.path.basename(ruta))[0]
        fecha = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        ruta_salida_excel = os.path.join(directorio, f"{nombre_base}_Consolidado_{fecha}.xlsx")

        try:
            wb_salida.save(ruta_salida_excel)
            print(f"\n{VERDE}¡Consolidación completada!{RESET}")
            print(f"{VERDE}Archivo guardado: {ruta_salida_excel}{RESET}")
        except Exception as e:
            print(f"{ROJO}Error al guardar: {e}{RESET}")
            wb_salida.close()
            wb.close()
            input(f"{BLANCO}Presiona Enter para continuar...{RESET}")
            return

        wb_salida.close()
        wb.close()
        input(f"\n{BLANCO}Presiona Enter para continuar...{RESET}")

    except Exception as e:
        print(f"{ROJO}Ocurrió un error: {e}{RESET}")
        input(f"{BLANCO}Presiona Enter para continuar...{RESET}")

def opcion_leer_zoom():
    try:
        print(f"{CYAN}--- LECTURA DE SOPORTE DE ZOOM ---{RESET}")
        tasa = pedir_tasa()
        ruta = seleccionar_soporte()
        if not ruta:
            print(f"{ROJO}No se encontro ningun archivo de soporte.{RESET}")
            input(f"{BLANCO}Presiona Enter para continuar...{RESET}")
            return
        print(f"{AZUL}Archivo seleccionado: {os.path.basename(ruta)}{RESET}")
        codigo_factura = extraer_codigo_factura(os.path.basename(ruta))
        datos = leer_soporte_zoom(ruta)
        if not datos:
            print(f"{ROJO}No se encontraron registros en el soporte.{RESET}")
        else:
            ruta_maestro_base = obtener_ruta_maestro()
            wb_plantilla = load_workbook(ruta_maestro_base, data_only=False)
            hoja_plantilla = hoja_referencia_estructura(wb_plantilla)
            estilos = copiar_formato(hoja_plantilla)
            wb_plantilla.close()
            ruta_temporal, total_t = generar_temporal(datos, tasa, estilos, codigo_factura)
            ruta_salida, total_s, ruta_maestro = generar_salida(datos, tasa, estilos, codigo_factura)
            guias = [d['envio'] for d in datos if d['envio'] is not None]
            resultados = consultar_envios_zoom(guias)
            aplicar_datos_zoom(ruta_temporal, resultados, codigo_factura, nombre_mes_actual())
            aplicar_datos_zoom(ruta_salida, resultados, codigo_factura, nombre_mes_actual())
            subtotal = sum(d['total'] for d in datos)
            fpo_total = sum(d['fpo'] for d in datos if d['fpo'] is not None)
            iva = round((subtotal - fpo_total) * 0.16, 2)
            total_factura = round(subtotal + iva, 2)
            print(f"{VERDE}Se generaron {total_t} registros.{RESET}")
            print(f"{VERDE}Subtotal: {subtotal:.2f} | FPO: {fpo_total:.2f} | IVA: {iva:.2f}{RESET}")
            print(f"{VERDE}Total factura: {total_factura:.2f}{RESET}")
            print(f"{VERDE}Archivo temporal: {ruta_temporal}{RESET}")
            print(f"{VERDE}Salida (Envios {anio_actual()} 2 actualizado): {ruta_salida}{RESET}")
            if ruta_maestro == ruta_maestro_base:
                print(f"{VERDE}Maestro Envios {anio_actual()} 2 reemplazado correctamente.{RESET}")
            else:
                print(f"{AMARILLO}No se pudo reemplazar el maestro (¿abierto en Excel?). Se guardo como: {ruta_maestro}{RESET}")
    except Exception as e:
        print(f"{ROJO}Ocurrio un error: {e}{RESET}")
        print(f"{ROJO}Verifica que el archivo de Salidas no este abierto en Excel y que los archivos existan.{RESET}")
    input(f"{BLANCO}Presiona Enter para continuar...{RESET}")


LINK_EXCALIDRAW = ("https://excalidraw.com/#room=68a64cd20ddece5e0e32,"
                   "dLSHtT_m8ddw-gYzC2vZ1A")


def abrir_excalidraw():
    """Opcion oculta 30310383: abre la pizarra compartida de Excalidraw."""
    print(f"\n{MAGENTA}{'=' * 50}{RESET}")
    print(f"{MAGENTA}{'  EXCALIDRAW (opcion oculta)'}{RESET}")
    print(f"{MAGENTA}{'=' * 50}{RESET}")
    print(f"{BLANCO}  Link:{RESET} {CYAN}{LINK_EXCALIDRAW}{RESET}\n")
    try:
        webbrowser.open(LINK_EXCALIDRAW, new=2)
        print(f"{VERDE}  Se abrio en el navegador. Pegalo a quien quieras colaborar.{RESET}")
    except Exception as e:
        print(f"{ROJO}  No se pudo abrir el navegador: {e}{RESET}")
        print(f"{AMARILLO}  Copia el link a mano en el navegador.{RESET}")
    input(f"\n{BLANCO}Presiona Enter para volver al menu...{RESET}")


def actualizar_desde_github():
    """Si hay commits nuevos en GitHub, hace pull y reinicia el programa."""
    import subprocess
    try:
        git = subprocess.run(['git', 'fetch', 'origin', 'main'],
                             capture_output=True, cwd=BASE, timeout=30)
        if git.returncode != 0:
            return
        atras = subprocess.run(['git', 'rev-list', '--count', 'HEAD..origin/main'],
                               capture_output=True, text=True, cwd=BASE, timeout=30)
        n = int((atras.stdout or '0').strip() or '0')
        if n <= 0:
            return
        print()
        print(f"{AMARILLO}  Hay {n} actualizacion(es) en GitHub. Actualizando el programa...{RESET}")

        res = subprocess.run(['git', 'pull', '--ff-only', 'origin', 'main'],
                             capture_output=True, text=True, cwd=BASE, timeout=120)
        if res.returncode == 0:
            print(f"{VERDE}  Programa actualizado. Reiniciando...{RESET}")
            import sys
            os.execv(sys.executable, [sys.executable] + sys.argv)
            return
        import time
        ultimo = subprocess.run(['git', 'log', '-1', '--format=%ct', 'HEAD'],
                                capture_output=True, text=True, cwd=BASE, timeout=30)
        try:
            hora_ultimo_commit = int((ultimo.stdout or '0').strip())
        except ValueError:
            hora_ultimo_commit = 0
        horas_sin_commit = (time.time() - hora_ultimo_commit) / 3600
        if hora_ultimo_commit and horas_sin_commit >= 8:
            print()
            print(f"{AMARILLO}  El ultimo commit local tiene {horas_sin_commit:.1f} horas. Forzando actualizacion...{RESET}")
            forzado = subprocess.run(['git', 'reset', '--hard', 'origin/main'],
                                     capture_output=True, text=True, cwd=BASE, timeout=120)
            if forzado.returncode == 0:
                print(f"{VERDE}  Programa actualizado forzadamente. Reiniciando...{RESET}")
                import sys
                os.execv(sys.executable, [sys.executable] + sys.argv)
            else:
                msg = (forzado.stderr or forzado.stdout or '').strip()
                print(f"{ROJO}  No se pudo forzar la actualizacion.{RESET}")
                print("  Detalle: " + msg[:250])
        else:
            pass
    except FileNotFoundError:
        pass
    except Exception as e:
        print(f"{AMARILLO}  (No se pudieron revisar actualizaciones: {e}){RESET}")


if __name__ == "__main__":
    actualizar_desde_github()
    while True:
        limpiar_pantalla()
        print(CYAN + "=" * 50)
        print("        MENU PRINCIPAL")
        print("=" * 50 + RESET)
        print(f"{AMARILLO}  1. Leer soporte de Zoom (Excel o TXT){RESET}")
        print(f"{VERDE}  2. Rutinas de mouse (grabar/repetir){RESET}")
        print(f"{AZUL}  3. Analisis a2Admin (facturacion ZM){RESET}")
        print(f"{MAGENTA}  4. Consolidar códigos desde Excel{RESET}")
        print(f"{ROJO}  5. Opcion 5{RESET}")
        print(f"{BLANCO}  0. Salir{RESET}")
        print(CYAN + "=" * 50 + RESET)

        opcion = input(f"{BLANCO}Selecciona una opcion: {RESET}")

        if opcion == "1":
            opcion_leer_zoom()
        elif opcion == "2":
            menu_rutinas()
        elif opcion == "3":
            menu_analisis()
        elif opcion == "4":
            consolidar_codigos_excel()
        elif opcion == "5":
            input(f"{ROJO}Seleccionaste Opcion 5. Presiona Enter...{RESET}")
        elif opcion == "30310383":
            abrir_excalidraw()
        elif opcion == "0":
            print(f"{CYAN}¡Hasta luego!{RESET}")
            break
        else:
            input(f"{ROJO}Opcion no valida. Presiona Enter...{RESET}")