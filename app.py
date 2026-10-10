"""Analítica de facturación en salud. Ejecutar: streamlit run app.py.
Datos exclusivamente en memoria; los importes son aditivos por registro.
"""
from io import BytesIO, StringIO
from datetime import date
import csv
import hashlib
import re
import unicodedata
import uuid
import warnings

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from scipy import stats
import streamlit as st

FIN = ['Valor_Bruto', 'Descuento', 'Copago_Cuota', 'Valor_Neto', 'Glosa_Inicial', 'Valor_Pagado', 'Saldo_Cartera']
FECHAS = ['Fecha_Servicio', 'Fecha_Emision', 'Fecha_Radicacion', 'Fecha_Aceptacion', 'Fecha_Pago']
CATEGORIAS = ['Entidad_Pagadora', 'Regimen', 'Sede', 'Servicio', 'Linea_Servicio', 'Canal_Radicacion', 'Estado_Factura', 'Soporte_Completo']
CAMPOS = ['ID_Factura', 'Numero_Radicado'] + FECHAS + ['Periodo_Cierre', 'Codigo_Entidad'] + CATEGORIAS + ['Codigo_Servicio', 'Cantidad', 'Dias_Emision_Radicacion'] + FIN
RECOMENDADAS = ['Entidad_Pagadora', 'Estado_Factura', 'Valor_Pagado', 'Saldo_Cartera', 'Glosa_Inicial']
PAGINAS = ['Inicio y carga de datos', 'Calidad de datos', 'Resumen ejecutivo', 'Cierre mensual', 'Entidades pagadoras', 'Servicios y líneas', 'Sedes y regímenes', 'Estados de facturación', 'Glosas', 'Recaudo y cartera', 'Tiempos de radicación', 'Análisis estadístico', 'Valores atípicos', 'Insights y alertas', 'Explorador de datos', 'Descargas', 'Metodología']
MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
PALETA = ['#087f8c', '#235789', '#48a9a6', '#f0a34b', '#d1495b']


def normalizar_nombre_columna(nombre):
    """Compara encabezados sin tildes, espacios ni puntuación."""
    texto = unicodedata.normalize('NFKD', str(nombre).strip()).encode('ascii', 'ignore').decode().lower()
    return re.sub(r'[^a-z0-9]', '', texto)


ALIASES = {normalizar_nombre_columna(c): c for c in CAMPOS}
for campo, nombres in {
    'ID_Factura': ['factura', 'numero factura', 'nro factura', 'identificador factura'],
    'Entidad_Pagadora': ['pagador', 'entidad', 'eps'], 'Copago_Cuota': ['copago', 'cuota moderadora'],
    'Linea_Servicio': ['linea de servicio'], 'Estado_Factura': ['estado'],
    'Valor_Neto': ['neto', 'valor neto factura'], 'Valor_Pagado': ['pagos', 'recaudo'],
    'Saldo_Cartera': ['cartera', 'saldo'], 'Glosa_Inicial': ['glosa', 'valor glosa'],
    'Dias_Emision_Radicacion': ['dias radicacion'],
}.items():
    ALIASES.update({normalizar_nombre_columna(n): campo for n in nombres})


def configurar_pagina():
    st.set_page_config(page_title='Analítica de Facturación en Salud', page_icon='📊', layout='wide', initial_sidebar_state='expanded')
    st.markdown('''<style>
    .stApp {background-color:#f5f8fb;color:#17324d;}
    h1,h2,h3 {color:#17324d;} div[data-testid="stMetric"] {
    background:white;border:1px solid #dbe6ed;border-radius:12px;padding:16px;}
    div[data-testid="stMetricValue"] {color:#087f8c;}
    .block-container {padding-top:2rem;padding-bottom:3rem;}
    </style>''', unsafe_allow_html=True)
    st.title('ANALÍTICA DE FACTURACIÓN EN SALUD')
    st.caption('Facturación, radicación, glosas, recaudo y cartera')
    st.info('Los resultados dependen de la calidad y estructura del archivo cargado. Esta herramienta facilita el análisis exploratorio y no reemplaza una auditoría financiera, contable, médica o contractual.')


@st.cache_data(ttl=900, max_entries=8, show_spinner=False)
def leer_archivo(contenido, extension, hoja, separador, codificacion, sesion):
    """Caché temporal aislada por token de sesión; nunca se persiste en disco."""
    if extension == 'xlsx':
        return pd.read_excel(BytesIO(contenido), sheet_name=hoja, engine='openpyxl', dtype=object)
    errores = []
    for enc in (['utf-8-sig', 'utf-16', 'cp1252', 'latin1'] if codificacion == 'Automática' else [codificacion]):
        try:
            texto = contenido.decode(enc)
            if '\x00' in texto:
                continue
            sep = separador
            if sep == 'Automático':
                sep = csv.Sniffer().sniff(texto[:32000], delimiters=',;\t').delimiter
            return pd.read_csv(StringIO(texto), sep=sep, dtype=object)
        except (UnicodeError, ValueError, csv.Error, pd.errors.ParserError) as exc:
            errores.append(type(exc).__name__)
    raise ValueError('Revise la codificación, el separador y que el archivo tenga encabezados.')


def cargar_archivo():
    st.sidebar.header('Archivo de trabajo')
    archivo = st.sidebar.file_uploader('Cargue el archivo de facturación que desea analizar', type=['xlsx', 'csv'], key='archivo')
    if archivo is None:
        return None
    contenido = archivo.getvalue()
    huella = hashlib.sha256(contenido).hexdigest()
    if st.session_state.get('huella') != huella:
        for k in list(st.session_state):
            if k not in ['archivo', 'sesion']:
                del st.session_state[k]
        st.session_state.huella = huella
    extension = archivo.name.rsplit('.', 1)[-1].lower()
    hoja, separador, codificacion = None, 'Automático', 'Automática'
    try:
        if extension == 'xlsx':
            with pd.ExcelFile(BytesIO(contenido), engine='openpyxl') as libro:
                hojas = libro.sheet_names
            preferida = next((i for i, x in enumerate(hojas) if normalizar_nombre_columna(x) == 'facturacion'), 0)
            hoja = st.sidebar.selectbox('Hoja a analizar', hojas, index=preferida, key='hoja')
        else:
            separador = st.sidebar.selectbox('Separador CSV', ['Automático', ';', ',', '\t'], format_func=lambda x: 'Tabulación' if x == '\t' else x)
            codificacion = st.sidebar.selectbox('Codificación CSV', ['Automática', 'utf-8-sig', 'cp1252', 'latin1', 'utf-16'])
        identidad = (huella, hoja, separador, codificacion)
        if st.session_state.get('identidad') != identidad:
            for k in list(st.session_state):
                if k.startswith(('map_', 'f_', 'lim_', 'out_')) or k in ['mapeo', 'original', 'limpios', 'config_atipicos']:
                    del st.session_state[k]
            st.session_state.identidad = identidad
        if archivo.size > 30_000_000:
            st.sidebar.warning('Archivo mayor de 30 MB: la lectura y Excel pueden tardar. Los gráficos de detalle usan muestras de hasta 5.000 filas.')
        with st.spinner('Leyendo el archivo en memoria…'):
            df = leer_archivo(contenido, extension, hoja, separador, codificacion, st.session_state.sesion)
        if df.empty:
            st.warning('La hoja seleccionada o el archivo no contiene datos.'); return None
        df.columns = df.columns.astype(str)
        st.session_state.original = df
        st.session_state.metadatos = {'Archivo': archivo.name, 'Tamaño MB': round(archivo.size / 1e6, 2), 'Hoja': hoja or 'CSV'}
        return df
    except Exception:
        st.error('No fue posible leer el archivo. Revise que no esté protegido o dañado. Para CSV, cambie el separador o la codificación.'); return None


def mapear_columnas(df):
    """Propone correspondencias y exige resolver colisiones explícitamente."""
    candidatos = {c: [x for x in df if ALIASES.get(normalizar_nombre_columna(x)) == c] for c in CAMPOS}
    with st.expander('Configurar correspondencia de columnas', expanded=any(len(v) > 1 for v in candidatos.values())):
        st.caption('Revise las equivalencias propuestas antes de aplicarlas. Las selecciones se conservan durante la sesión.')
        propuestas = [{'Campo esperado': c, 'Columnas encontradas': ', '.join(v) or 'No disponible'} for c, v in candidatos.items()]
        st.dataframe(pd.DataFrame(propuestas), hide_index=True, use_container_width=True)
        with st.form('form_mapeo'):
            columnas = st.columns(3)
            seleccion = {}
            opciones = ['Columna no disponible'] + list(df.columns)
            for i, c in enumerate(CAMPOS):
                previo = st.session_state.get('mapeo', {}).get(c)
                propuesta = previo if previo in df else (candidatos[c][0] if len(candidatos[c]) == 1 else None)
                seleccion[c] = columnas[i % 3].selectbox(c.replace('_', ' '), opciones, index=opciones.index(propuesta) if propuesta else 0, key='map_' + c)
            aplicar = st.form_submit_button('Aplicar correspondencia', type='primary')
        if aplicar:
            usados = [v for v in seleccion.values() if v != opciones[0]]
            if len(usados) != len(set(usados)):
                st.error('Una columna de origen no puede asignarse a dos campos. Corrija la correspondencia.')
            else:
                st.session_state.mapeo = {k: v for k, v in seleccion.items() if v != opciones[0]}
                st.success('Correspondencia aplicada.')
    if 'mapeo' not in st.session_state:
        st.info('Abra «Configurar correspondencia de columnas», revise las equivalencias y pulse «Aplicar correspondencia».')
        return None
    mapeo = st.session_state.mapeo
    # Preserva columnas extra con un prefijo si su nombre colisiona con el destino.
    extras = df.drop(columns=list(mapeo.values())).copy()
    extras.columns = ['Original_' + c if c in CAMPOS else c for c in extras.columns]
    usados = set(mapeo)
    nombres = []
    for c in extras.columns:
        nombre = c
        while nombre in usados:
            nombre = 'Original_' + nombre
        usados.add(nombre); nombres.append(nombre)
    extras.columns = nombres
    resultado = pd.concat([pd.DataFrame({c: df[v] for c, v in mapeo.items()}), extras], axis=1)
    return resultado


def convertir_monetaria(serie, modo='Automático conservador'):
    """No adivina importes ambiguos de un único separador con tres dígitos finales."""
    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(serie, errors='coerce').replace([np.inf, -np.inf], np.nan), 0
    texto = serie.astype('string').str.strip()
    # Números ya tipados no deben reinterpretarse según la región.
    nativos = serie.map(lambda v: isinstance(v, (int, float, np.number)) and not isinstance(v, bool))
    s = texto.str.replace(r'(?i)COP|USD|EUR|\$|€|\s', '', regex=True)
    s = s.str.replace(r'^\((.*)\)$', r'-\1', regex=True)
    ambiguos = s.str.match(r'^[+-]?\d{1,3}[.,]\d{3}$', na=False) & ~nativos
    if modo == 'Decimal coma (1.234,56)':
        s = s.str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
    elif modo == 'Decimal punto (1,234.56)':
        s = s.str.replace(',', '', regex=False)
    else:
        ambos = s.str.contains(',', na=False) & s.str.contains('.', regex=False, na=False)
        coma_decimal = ambos & (s.str.rfind(',') > s.str.rfind('.'))
        punto_decimal = ambos & ~coma_decimal
        s.loc[coma_decimal] = s[coma_decimal].str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
        s.loc[punto_decimal] = s[punto_decimal].str.replace(',', '', regex=False)
        miles = s.str.match(r'^[+-]?\d{1,3}([.,]\d{3}){2,}$', na=False)
        s.loc[miles] = s[miles].str.replace(r'[.,]', '', regex=True)
        s = s.str.replace(',', '.', regex=False)
        s.loc[ambiguos] = pd.NA
    salida = pd.to_numeric(s, errors='coerce').astype(float)
    salida.loc[nativos] = pd.to_numeric(serie[nativos], errors='coerce')
    return salida.replace([np.inf, -np.inf], np.nan), int(ambiguos.sum()) if modo == 'Automático conservador' else 0


def convertir_fecha(serie, dia_primero=True):
    """Fechas textuales mixtas y seriales Excel; números fuera de rango se invalidan."""
    if pd.api.types.is_datetime64_any_dtype(serie):
        return pd.to_datetime(serie, errors='coerce')
    numerica = pd.to_numeric(serie, errors='coerce')
    serial = numerica.between(20000, 90000)
    resultado = pd.to_datetime(serie.where(numerica.isna()), errors='coerce', dayfirst=dia_primero, format='mixed')
    resultado.loc[serial] = pd.to_datetime(numerica[serial], unit='D', origin='1899-12-30', errors='coerce')
    # Devuelve fechas sin zona, apropiadas para registros contables diarios.
    return pd.to_datetime(resultado, errors='coerce', utc=True).dt.tz_convert(None)


def dividir(a, b):
    if isinstance(b, pd.Series):
        return (a / b.replace(0, np.nan)).replace([np.inf, -np.inf], np.nan)
    if isinstance(a, pd.Series):
        return (a / b).replace([np.inf, -np.inf], np.nan) if pd.notna(b) and b != 0 else pd.Series(np.nan, index=a.index)
    return float(a / b) if pd.notna(a) and pd.notna(b) and b != 0 else np.nan


def disponible(df, campo):
    return campo in df and pd.api.types.is_numeric_dtype(df[campo])


def sumar(df, campo):
    return df[campo].sum(min_count=1) if disponible(df, campo) else np.nan


def facturas(df):
    return df['ID_Factura'].dropna().nunique() if 'ID_Factura' in df else 0


def soporte_incompleto(serie):
    s = serie.astype('string').str.strip().str.lower()
    return s.map({'no': 1., 'incompleto': 1., 'false': 1., '0': 1., '0.0': 1., 'sí': 0., 'si': 0., 'completo': 0., 'true': 0., '1': 0., '1.0': 0.})


def derivar(df, cortes=(3, 7, 15)):
    df = df.copy()
    if 'Fecha_Radicacion' in df and pd.api.types.is_datetime64_any_dtype(df.Fecha_Radicacion):
        f = df.Fecha_Radicacion
        df['Año_Radicacion'] = f.dt.year
        df['Numero_Mes'] = f.dt.month
        df['Mes_Radicacion'] = f.dt.month.map(dict(enumerate(MESES, 1)))
        df['Trimestre_Radicacion'] = f.dt.quarter
        df['Año_Mes'] = f.dt.to_period('M').astype('string').mask(f.isna())
        df['Día_Semana'] = f.dt.dayofweek.map(dict(enumerate(['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo'])))
        if 'Periodo_Cierre' not in df:
            df['Periodo_Cierre'] = df['Año_Mes']
        for fecha, destino, invertir in [('Fecha_Emision', 'Dias_Emision_Radicacion', True), ('Fecha_Pago', 'Dias_Radicacion_Pago', False)]:
            if fecha in df and pd.api.types.is_datetime64_any_dtype(df[fecha]):
                calculado = (f - df[fecha]).dt.total_seconds() / 86400 if invertir else (df[fecha] - f).dt.total_seconds() / 86400
                if destino not in df:
                    df[destino] = calculado
                else:
                    df[destino + '_Calculado'] = calculado
    for nombre, numerador, denominador in [('Descuento', 'Descuento', 'Valor_Bruto'), ('Copago', 'Copago_Cuota', 'Valor_Bruto'), ('Glosa', 'Glosa_Inicial', 'Valor_Neto'), ('Recaudo', 'Valor_Pagado', 'Valor_Neto'), ('Cartera', 'Saldo_Cartera', 'Valor_Neto')]:
        if disponible(df, numerador) and disponible(df, denominador):
            df['Porcentaje_' + nombre] = dividir(df[numerador], df[denominador]) * 100
    for origen, destino in [('Glosa_Inicial', 'Tiene_Glosa'), ('Saldo_Cartera', 'Tiene_Saldo')]:
        if disponible(df, origen):
            df[destino] = df[origen].gt(0).astype('boolean').mask(df[origen].isna())
    if 'Estado_Factura' in df:
        df['Es_Pagada'] = df.Estado_Factura.astype('string').str.lower().str.strip().eq('pagada')
    if 'Soporte_Completo' in df:
        df['Soporte_Incompleto'] = soporte_incompleto(df.Soporte_Completo)
    if disponible(df, 'Dias_Emision_Radicacion'):
        a, b, c = cortes
        df['Rango_Tiempo_Radicacion'] = pd.cut(df.Dias_Emision_Radicacion, [-np.inf, -0.000001, a, b, c, np.inf], labels=['Negativo: revisar', f'0 a {a}', f'>{a} a {b}', f'>{b} a {c}', f'>{c}']).astype('string')
    return df


@st.cache_data(ttl=900, max_entries=8, show_spinner=False)
def transformar(original, opciones, sesion):
    df = original.copy().replace(r'^\s*$', np.nan, regex=True)
    conversiones, exclusiones = [], []
    if not opciones['aplicar']:
        return original.copy(), pd.DataFrame(), pd.DataFrame()
    for campo in FECHAS:
        if campo in df and opciones['fechas']:
            previo = df[campo].notna()
            df[campo] = convertir_fecha(df[campo], opciones['dia'])
            conversiones.append([campo, int((previo & df[campo].isna()).sum()), int(df[campo].isna().sum()), 0])
    if 'Periodo_Cierre' in df:
        previo = df.Periodo_Cierre.notna()
        if pd.api.types.is_datetime64_any_dtype(df.Periodo_Cierre):
            df['Periodo_Cierre'] = df.Periodo_Cierre.dt.strftime('%Y-%m')
        else:
            p = df.Periodo_Cierre.astype('string').str.strip()
            valido = p.str.match(r'^\d{4}-(0[1-9]|1[0-2])$', na=False)
            df['Periodo_Cierre'] = p.where(valido)
        conversiones.append(['Periodo_Cierre', int((previo & df.Periodo_Cierre.isna()).sum()), int(df.Periodo_Cierre.isna().sum()), 0])
    for campo in FIN + ['Cantidad', 'Dias_Emision_Radicacion']:
        if campo in df and opciones['moneda']:
            previo = df[campo].notna()
            df[campo], ambiguos = convertir_monetaria(df[campo], opciones['formato'])
            conversiones.append([campo, int((previo & df[campo].isna()).sum()), int(df[campo].isna().sum()), ambiguos])
    def quitar(mascara, razon):
        nonlocal df
        cantidad = int(mascara.sum())
        df = df.loc[~mascara].copy()
        exclusiones.append([razon, cantidad])
    if opciones['vacias']:
        # Se evalúa contra el original, antes de que conversiones fallen.
        quitar(original.loc[df.index].replace(r'^\s*$', np.nan, regex=True).isna().all(axis=1), 'Filas originalmente vacías')
    if opciones['duplicados']:
        quitar(original.loc[df.index].duplicated(), 'Duplicados exactos del archivo')
    if opciones['neto'] and disponible(df, 'Valor_Neto'):
        quitar(df.Valor_Neto.isna(), 'Valor neto inválido o ausente')
    if opciones['invalidas']:
        columnas = [c for c in FECHAS if c in df and pd.api.types.is_datetime64_any_dtype(df[c])]
        mascara = pd.Series(False, index=df.index)
        for c in columnas:
            mascara |= original.loc[df.index, c].notna() & df[c].isna()
        if 'Periodo_Cierre' in original:
            mascara |= original.loc[df.index, 'Periodo_Cierre'].notna() & df.Periodo_Cierre.isna()
        quitar(mascara, 'Fechas o periodos informados pero inválidos')
    if opciones['categorias']:
        for c in CATEGORIAS:
            if c in df:
                df[c] = df[c].fillna('Sin información').astype(str)
    df = derivar(df, opciones['cortes'])
    return df, pd.DataFrame(conversiones, columns=['Columna', 'No convertidos', 'Nulos finales', 'Ambiguos']), pd.DataFrame(exclusiones, columns=['Razón', 'Registros excluidos'])


def configurar_limpieza():
    with st.sidebar.expander('Limpieza controlada y formatos'):
        if st.button('Restaurar datos originales'):
            for k in list(st.session_state):
                if k.startswith('lim_'):
                    del st.session_state[k]
            st.session_state.lim_aplicar = False
            st.rerun()
        opciones = {}
        for clave, etiqueta, defecto in [('aplicar', 'Aplicar transformaciones', True), ('vacias', 'Eliminar filas completamente vacías', False), ('duplicados', 'Eliminar duplicados exactos', False), ('fechas', 'Convertir fechas', True), ('moneda', 'Convertir valores numéricos', True), ('categorias', 'Categorías vacías: Sin información', False), ('neto', 'Excluir valor neto inválido', False), ('invalidas', 'Excluir fechas informadas inválidas', False), ('dia', 'Fechas: día antes del mes', True)]:
            opciones[clave] = st.checkbox(etiqueta, value=defecto, key='lim_' + clave)
        opciones['formato'] = st.selectbox('Formato numérico del archivo', ['Automático conservador', 'Decimal coma (1.234,56)', 'Decimal punto (1,234.56)'], key='lim_formato')
        st.caption('Automático deja nulos los textos ambiguos como 1.234. Seleccione el formato regional para resolverlos.')
        opciones['tolerancia'] = st.number_input('Tolerancia de identidades financieras', 0.0, value=1.0, key='lim_tol')
        cortes = [st.number_input(f'Límite de radicación {i + 1} (días)', min_value=0, value=v, key=f'lim_corte{i}') for i, v in enumerate([3, 7, 15])]
        opciones['cortes'] = tuple(cortes) if cortes[0] < cortes[1] < cortes[2] else (3, 7, 15)
        if tuple(cortes) != opciones['cortes']:
            st.warning('Los límites deben aumentar. Se usan 3, 7 y 15 hasta corregirlos.')
    return opciones


def validar_datos(df, tolerancia=1):
    """Cada regla conserva su universo evaluable y la máscara de registros afectados."""
    filas, mascaras = [], {}
    def regla(nombre, mal, evaluable=None):
        mal = mal.fillna(False).astype(bool)
        evaluable = pd.Series(True, index=df.index) if evaluable is None else evaluable.fillna(False)
        mascaras[nombre] = mal & evaluable
        n = int(evaluable.sum())
        filas.append({'Regla': nombre, 'Casos': int(mascaras[nombre].sum()), 'Evaluables': n, '%': dividir(int(mascaras[nombre].sum()) * 100, n)})
    regla('Filas duplicadas (excedentes)', df.duplicated())
    for c in ['ID_Factura', 'Numero_Radicado']:
        if c in df:
            regla(c + ' repetido (excedentes)', df[c].duplicated(), df[c].notna())
    for c in FIN:
        if disponible(df, c):
            regla(c + ' negativo', df[c] < 0, df[c].notna())
    if disponible(df, 'Valor_Neto'):
        regla('Valor neto cero', df.Valor_Neto == 0, df.Valor_Neto.notna())
        for c in ['Saldo_Cartera', 'Valor_Pagado', 'Glosa_Inicial']:
            if disponible(df, c):
                regla(c + ' superior al neto', df[c] > df.Valor_Neto + tolerancia, df[[c, 'Valor_Neto']].notna().all(axis=1))
    if 'Entidad_Pagadora' in df:
        regla('Sin entidad pagadora', df.Entidad_Pagadora.isna() | df.Entidad_Pagadora.eq('Sin información'))
    if 'ID_Factura' in df:
        regla('Sin identificador de factura', df.ID_Factura.isna())
    regla('Sin periodo válido', df.Periodo_Cierre.isna() if 'Periodo_Cierre' in df else pd.Series(True, index=df.index))
    for a, b, nombre in [('Fecha_Radicacion', 'Fecha_Emision', 'Radicación anterior a emisión'), ('Fecha_Pago', 'Fecha_Radicacion', 'Pago anterior a radicación')]:
        if all(c in df and pd.api.types.is_datetime64_any_dtype(df[c]) for c in [a, b]):
            regla(nombre, df[a] < df[b], df[[a, b]].notna().all(axis=1))
    for destino, terminos, nombre in [('Valor_Neto', ['Valor_Bruto', 'Descuento', 'Copago_Cuota'], 'Identidad de valor neto'), ('Saldo_Cartera', ['Valor_Neto', 'Glosa_Inicial', 'Valor_Pagado'], 'Identidad de saldo (hipótesis analítica)')]:
        if all(disponible(df, c) for c in [destino] + terminos):
            esperado = df[terminos[0]] - df[terminos[1]] - df[terminos[2]]
            regla(nombre, (df[destino] - esperado).abs() > tolerancia, df[[destino] + terminos].notna().all(axis=1))
    if disponible(df, 'Dias_Emision_Radicacion_Calculado') and disponible(df, 'Dias_Emision_Radicacion'):
        regla('Días reportados distintos a fechas', (df.Dias_Emision_Radicacion - df.Dias_Emision_Radicacion_Calculado).abs() > 1)
    return pd.DataFrame(filas), pd.DataFrame(mascaras, index=df.index)


def calidad(df, original, conversiones, tolerancia):
    informe, mascaras = validar_datos(df, tolerancia)
    nulos = pd.DataFrame({'Nulos': df.isna().sum(), '% nulos': df.isna().mean() * 100}).reset_index(names='Columna')
    completitud = 100 * (1 - df.isna().mean().mean()) if len(df) else 0
    unicidad = 100 * (1 - df.duplicated().mean()) if len(df) else 0
    fallos = conversiones['No convertidos'].sum() if len(conversiones) else 0
    universo = len(original) * len(conversiones)
    validez = 100 * (1 - fallos / universo) if universo else 100
    consistencia = 100 * (1 - mascaras.any(axis=1).mean()) if len(df) else 0
    componentes = pd.Series({'Completitud': completitud, 'Unicidad de filas': unicidad, 'Validez de conversión': validez, 'Consistencia exploratoria': consistencia})
    return informe, mascaras, nulos, componentes


def reset_filtros():
    for clave in list(st.session_state):
        if clave.startswith('f_'):
            del st.session_state[clave]


def aplicar_filtros(df):
    """Todos los módulos reciben el mismo DataFrame filtrado."""
    mascara = pd.Series(True, index=df.index)
    with st.sidebar.expander('Filtros globales', expanded=False):
        st.button('Seleccionar todos', on_click=reset_filtros)
        st.button('Limpiar filtros', on_click=reset_filtros)
        fecha = df['Fecha_Radicacion'] if 'Fecha_Radicacion' in df and pd.api.types.is_datetime64_any_dtype(df.Fecha_Radicacion) else pd.to_datetime(df.get('Periodo_Cierre', pd.Series(pd.NaT, index=df.index)), errors='coerce')
        validas = fecha.dropna()
        incluir_nulos = st.checkbox('Incluir fechas y valores netos ausentes', True, key='f_nulos')
        if len(validas):
            rango = st.date_input('Rango de fechas de radicación / cierre', (validas.min().date(), validas.max().date()), key='f_fecha')
            if len(rango) == 2:
                mascara &= fecha.between(pd.Timestamp(rango[0]), pd.Timestamp(rango[1]) + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1)) | (fecha.isna() & incluir_nulos)
            for etiqueta, valores in [('Año', fecha.dt.year), ('Mes', fecha.dt.month)]:
                opciones = sorted(valores.dropna().unique().astype(int))
                elegidos = st.multiselect(etiqueta, opciones, default=opciones, key='f_' + etiqueta, format_func=(lambda x: MESES[x-1]) if etiqueta == 'Mes' else str)
                mascara &= valores.isin(elegidos) | (valores.isna() & incluir_nulos)
        for c in CATEGORIAS:
            if c in df:
                valores = df[c].astype('string').fillna('Sin información')
                opciones = sorted(valores.unique().tolist())
                elegidos = st.multiselect(c.replace('_', ' '), opciones, default=opciones, key='f_' + c)
                mascara &= valores.isin(elegidos)
        for campo, etiqueta in [('Glosa_Inicial', 'Facturas con glosa'), ('Saldo_Cartera', 'Facturas con saldo')]:
            if disponible(df, campo):
                elegido = st.selectbox(etiqueta, ['Todas', 'Sí', 'No'], key='f_' + campo)
                if elegido != 'Todas':
                    mascara &= (df[campo] > 0 if elegido == 'Sí' else df[campo] <= 0)
        if disponible(df, 'Valor_Neto') and df.Valor_Neto.notna().any():
            a, b = float(df.Valor_Neto.min()), float(df.Valor_Neto.max())
            inferior = st.number_input('Valor neto mínimo', value=a, key='f_min')
            superior = st.number_input('Valor neto máximo', value=b, key='f_max')
            mascara &= df.Valor_Neto.between(inferior, superior) | (df.Valor_Neto.isna() & incluir_nulos)
    filtrado = df.loc[mascara].copy()
    st.sidebar.caption(f'{len(filtrado):,} de {len(df):,} registros · {dividir(len(filtrado)*100, len(df)):.1f}% visibles')
    return filtrado


@st.cache_data(ttl=900, max_entries=24, show_spinner=False)
def agrupar(df, grupo, sesion=''):
    """Suma importes por fila; cuenta identificadores únicos dentro de cada grupo."""
    if grupo not in df:
        return pd.DataFrame()
    g = df.groupby(grupo, dropna=False, observed=True)
    salida = g.size().rename('Registros').to_frame()
    if 'ID_Factura' in df:
        salida['Facturas'] = g.ID_Factura.nunique()
    for c in FIN + ['Cantidad']:
        if disponible(df, c):
            salida[c] = g[c].sum(min_count=1)
    if 'Valor_Neto' in salida:
        salida['Ticket_Promedio'] = dividir(salida.Valor_Neto, salida.Facturas)
        salida['Participacion_%'] = dividir(salida.Valor_Neto, salida.Valor_Neto.sum(min_count=1)) * 100
        for c, nombre in [('Glosa_Inicial', 'Glosa_%'), ('Valor_Pagado', 'Recaudo_%'), ('Saldo_Cartera', 'Cartera_%')]:
            if c in salida:
                salida[nombre] = dividir(salida[c], salida.Valor_Neto) * 100
    for c in ['Dias_Emision_Radicacion', 'Soporte_Incompleto']:
        if disponible(df, c):
            salida[c] = g[c].mean() * (100 if c == 'Soporte_Incompleto' else 1)
    if 'Estado_Factura' in df and 'ID_Factura' in df:
        devueltas = df.loc[df.Estado_Factura.astype('string').str.lower().str.contains('devuelt', na=False)]
        salida['Facturas_Devueltas'] = devueltas.groupby(grupo, dropna=False).ID_Factura.nunique().reindex(salida.index, fill_value=0)
    return salida.reset_index()


def cierre_mensual(df, frecuencia='M'):
    if 'Periodo_Cierre' not in df:
        return pd.DataFrame()
    fechas = pd.to_datetime(df.Periodo_Cierre, errors='coerce')
    trabajo = df.loc[fechas.notna()].copy()
    if trabajo.empty:
        return pd.DataFrame()
    trabajo['Periodo'] = fechas.loc[trabajo.index].dt.to_period(frecuencia)
    resultado = agrupar(trabajo, 'Periodo', st.session_state.get('sesion', ''))
    # Reindexar deja meses faltantes como NaN, nunca como facturación cero inventada.
    resultado = resultado.set_index('Periodo').reindex(pd.period_range(trabajo.Periodo.min(), trabajo.Periodo.max(), freq=frecuencia))
    resultado.index.name = 'Periodo'
    neto = resultado.Valor_Neto
    resultado['Crecimiento_%'] = dividir(neto - neto.shift(), neto.shift()) * 100
    resultado['Acumulado_Anual'] = neto.groupby(resultado.index.year).cumsum()
    resultado['Promedio_Movil_3_Periodos'] = neto.rolling(3, min_periods=3).mean()
    resultado = resultado.reset_index()
    resultado.Periodo = resultado.Periodo.astype(str)
    return resultado


def calcular_kpis(df):
    neto = sumar(df, 'Valor_Neto')
    k = {c: sumar(df, c) for c in FIN + ['Cantidad']}
    k['Facturas'] = facturas(df)
    k['Ticket_Promedio'] = dividir(neto, k['Facturas'])
    for c, nombre in [('Glosa_Inicial', 'Glosa_%'), ('Valor_Pagado', 'Recaudo_%'), ('Saldo_Cartera', 'Cartera_%')]:
        k[nombre] = dividir(k[c], neto) * 100
    k['Dias_Emision_Radicacion'] = df.Dias_Emision_Radicacion.mean() if disponible(df, 'Dias_Emision_Radicacion') else np.nan
    k['Soporte_Incompleto'] = df.Soporte_Incompleto.mean() * 100 if disponible(df, 'Soporte_Incompleto') else np.nan
    return k


def formato(valor, monetario=False, porcentaje=False):
    if pd.isna(valor) or not np.isfinite(valor):
        return 'No disponible'
    texto = f'{valor:,.1f}' if porcentaje else f'{valor:,.0f}'
    texto = texto.replace(',', 'X').replace('.', ',').replace('X', '.')
    etiqueta = st.session_state.get('moneda', 'COP')
    prefijo = {'COP': '$ ', 'USD': 'USD ', 'EUR': 'EUR ', 'Sin símbolo': ''}.get(etiqueta, '')
    return (prefijo if monetario else '') + texto + (' %' if porcentaje else '')


def tabla(df, limite=300):
    configuracion = {}
    for c in df:
        if pd.api.types.is_datetime64_any_dtype(df[c]):
            configuracion[c] = st.column_config.DatetimeColumn(format='DD/MM/YYYY')
        elif disponible(df, c):
            configuracion[c] = st.column_config.NumberColumn(format='%.2f' if '%' in c or c.startswith('Porcentaje') else 'localized')
    st.dataframe(df.head(limite), use_container_width=True, hide_index=True, height=min(500, 40 + 35 * min(limite, len(df))), column_config=configuracion)
    if len(df) > limite:
        st.caption(f'Vista: {limite:,} de {len(df):,} filas. La descarga conserva todas.')


def grafico(fig):
    fig.update_layout(template='plotly_white', colorway=PALETA, font={'color': '#17324d'}, margin=dict(l=15, r=15, t=55, b=20), legend_title_text='')
    st.plotly_chart(fig, use_container_width=True)


def muestra(df):
    if len(df) > 5000:
        st.caption('Gráfico de detalle: muestra reproducible de 5.000 registros. Los indicadores usan todos los registros filtrados.')
        return df.sample(5000, random_state=42)
    return df


def evolucion(df, variables=None, titulo='Evolución por periodo de cierre'):
    resumen = cierre_mensual(df)
    variables = variables or ['Valor_Neto', 'Valor_Pagado', 'Saldo_Cartera']
    variables = [c for c in variables if c in resumen]
    if not resumen.empty and variables:
        grafico(px.line(resumen, x='Periodo', y=variables, markers=True, title=titulo, labels={'value': 'Valor', 'variable': 'Indicador'}))
    else:
        st.info('No hay periodos válidos o variables para esta evolución.')


def mostrar_kpis(df):
    k = calcular_kpis(df)
    anterior, actual = {}, {}
    if 'Periodo_Cierre' in df:
        periodos = pd.to_datetime(df.Periodo_Cierre, errors='coerce').dt.to_period('M')
        if periodos.notna().any():
            ultimo = periodos.max()
            if (periodos == ultimo - 1).any():
                actual = calcular_kpis(df.loc[periodos == ultimo]); anterior = calcular_kpis(df.loc[periodos == ultimo-1])
                st.caption(f'Tarjetas: total filtrado. Variaciones: {ultimo} frente a {ultimo-1}, dentro de los filtros actuales; los meses pueden estar incompletos.')
    campos = ['Valor_Bruto', 'Valor_Neto', 'Facturas', 'Cantidad', 'Valor_Pagado', 'Saldo_Cartera', 'Glosa_Inicial', 'Recaudo_%', 'Glosa_%', 'Ticket_Promedio', 'Dias_Emision_Radicacion']
    columnas = st.columns(4)
    for i, c in enumerate(campos):
        variacion = dividir(actual.get(c, np.nan) - anterior.get(c, np.nan), anterior.get(c, np.nan)) * 100
        columnas[i % 4].metric(c.replace('_', ' '), formato(k[c], c in FIN or c == 'Ticket_Promedio', '%' in c), delta=(formato(variacion, porcentaje=True) + ' último mes') if pd.notna(variacion) else None, delta_color='off')


def pareto(df, grupo, valor):
    resumen = agrupar(df, grupo, st.session_state.sesion)
    if valor not in resumen:
        return
    p = resumen[[grupo, valor]].dropna().sort_values(valor, ascending=False)
    if p.empty or (p[valor] < 0).any() or p[valor].sum() <= 0:
        st.info('Pareto requiere valores agregados no negativos y total positivo.'); return
    p['Participacion_%'] = p[valor] / p[valor].sum() * 100
    p['Acumulado_%'] = p['Participacion_%'].cumsum()
    n80 = int((p['Acumulado_%'] < 80).sum()) + 1
    st.write(f'{n80} de {len(p)} categorías reúnen al menos el 80 %. Concentración top 5: {p.head(5)["Participacion_%"].sum():.1f} %.')
    limite = st.number_input('Categorías visibles en Pareto', 1, max(1, len(p)), min(25, len(p)), key='pareto_lim_' + grupo + valor)
    v = p.head(limite)
    fig = make_subplots(specs=[[{'secondary_y': True}]])
    fig.add_trace(go.Bar(x=v[grupo], y=v[valor], name=valor), secondary_y=False)
    fig.add_trace(go.Scatter(x=v[grupo], y=v['Acumulado_%'], name='% acumulado'), secondary_y=True)
    fig.add_hline(y=80, secondary_y=True, line_dash='dash')
    fig.update_yaxes(range=[0, 105], title_text='% acumulado', secondary_y=True)
    fig.update_layout(title=f'Pareto · {valor} por {grupo}')
    grafico(fig); tabla(p)


def heatmap(df, filas, columnas, valor='Valor_Neto'):
    if all(c in df for c in [filas, columnas, valor]):
        principales = df.groupby(filas, observed=True)[valor].sum().nlargest(20).index
        p = df[df[filas].isin(principales)].pivot_table(index=filas, columns=columnas, values=valor, aggfunc='sum', observed=True)
        p = p.iloc[:, -36:]
        if not p.empty:
            grafico(go.Figure(go.Heatmap(z=p.values, x=p.columns.astype(str), y=p.index.astype(str), colorscale='Teal', colorbar_title=valor)).update_layout(title=f'{filas} × {columnas} · hasta 20 × 36 categorías'))


def mostrar_grupo(df, grupo):
    if grupo not in df:
        st.info(f'No se encontró la columna necesaria: {grupo}.'); return
    r = agrupar(df, grupo, st.session_state.sesion)
    tabla(r)
    opciones = [c for c in ['Valor_Neto', 'Saldo_Cartera', 'Glosa_Inicial', 'Valor_Pagado', 'Facturas', 'Dias_Emision_Radicacion'] if c in r]
    valor = st.selectbox('Ordenar ranking por', opciones, key='ranking_' + grupo)
    top = st.selectbox('Categorías en gráficos', ['Top 5', 'Top 10', 'Top 15', 'Todos'], index=1, key='top_' + grupo)
    n = len(r) if top == 'Todos' else int(top.split()[1])
    r = r.sort_values(valor, ascending=False, na_position='last')
    v = r.head(n).copy()
    v[grupo] = v[grupo].astype('string').fillna('Sin información')
    grafico(px.bar(v, x=valor, y=grupo, orientation='h', title=f'Ranking · {grupo}', height=max(350, min(1200, n * 28))))
    if 'Recaudo_%' in v:
        dispersion = v.dropna(subset=['Valor_Neto', 'Recaudo_%']).copy()
        if not dispersion.empty:
            dispersion['Tamaño_cartera'] = dispersion.get('Saldo_Cartera', pd.Series(1, index=dispersion.index)).fillna(0).clip(lower=0) + 1
            grafico(px.scatter(dispersion, x='Valor_Neto', y='Recaudo_%', size='Tamaño_cartera', color='Glosa_%' if 'Glosa_%' in dispersion and dispersion['Glosa_%'].notna().any() else None, hover_name=grupo, hover_data=['Facturas', 'Ticket_Promedio'], title='Facturación y recaudo · tamaño: cartera positiva'))
    positivos = v.loc[v.Valor_Neto > 0]
    if not positivos.empty:
        grafico(px.treemap(positivos, path=[grupo], values='Valor_Neto', title='Participación de valores netos positivos'))
    grafico(px.box(muestra(df[df[grupo].astype('string').fillna('Sin información').isin(v[grupo])]), x=grupo, y='Valor_Neto', title='Distribución de valor neto por registro', points=False))
    heatmap(df, grupo, 'Periodo_Cierre')
    if grupo in ['Servicio', 'Linea_Servicio']:
        heatmap(df, grupo, 'Entidad_Pagadora')
    if grupo in ['Sede', 'Regimen']:
        variables = [c for c in ['Valor_Neto', 'Valor_Pagado', 'Saldo_Cartera'] if c in v]
        grafico(px.bar(v, x=grupo, y=variables, barmode='group', title='Comparación de importes'))
        if 'Estado_Factura' in df:
            cruce = pd.crosstab(df[grupo], df.Estado_Factura, normalize='index') * 100
            grafico(px.bar(cruce, barmode='stack', title='Estados por grupo · % de registros'))
    seleccion = st.selectbox('Detalle de una categoría', sorted(df[grupo].dropna().astype(str).unique()), index=None, key='detalle_' + grupo)
    if seleccion is not None:
        detalle = df[df[grupo].astype(str) == seleccion]
        mostrar_kpis(detalle); evolucion(detalle)
    if grupo in ['Entidad_Pagadora', 'Servicio']:
        st.subheader('Concentración y Pareto')
        metrica = st.selectbox('Métrica de concentración', [c for c in ['Valor_Neto', 'Saldo_Cartera', 'Glosa_Inicial'] if disponible(df, c)], key='pareto_' + grupo)
        pareto(df, grupo, metrica)


def mostrar_cierre(df):
    frecuencia = st.radio('Periodicidad', ['Mensual', 'Trimestral', 'Anual'], horizontal=True)
    f = {'Mensual': 'M', 'Trimestral': 'Q', 'Anual': 'Y'}[frecuencia]
    r = cierre_mensual(df, f)
    if r.empty:
        st.info('No hay periodos válidos.'); return
    tabla(r)
    st.caption('Pagos, glosas y cartera agrupados por cierre de la factura; no representan flujos históricos de caja ni saldos reconstruidos. Promedio móvil: tres periodos completos disponibles. Los periodos ausentes quedan sin datos.')
    for campo in ['Valor_Neto', 'Valor_Pagado', 'Saldo_Cartera', 'Glosa_Inicial']:
        if campo in r and r[campo].notna().any():
            mayor = r.loc[r[campo].idxmax()]
            st.write(f'Máximo {campo}: {mayor.Periodo} · {formato(mayor[campo], True)}')
    minimo = r.loc[r.Valor_Neto.idxmin()] if r.Valor_Neto.notna().any() else None
    if minimo is not None:
        st.write(f'Menor facturación: {minimo.Periodo} · {formato(minimo.Valor_Neto, True)}')
    st.write('Variación del último periodo:', formato(r.iloc[-1]['Crecimiento_%'], porcentaje=True))
    variables = [c for c in ['Valor_Neto', 'Valor_Pagado', 'Saldo_Cartera', 'Promedio_Movil_3_Periodos'] if c in r]
    grafico(px.line(r, x='Periodo', y=variables, markers=True, title='Facturación, recaudo, cartera y promedio móvil'))
    fig = make_subplots(specs=[[{'secondary_y': True}]])
    fig.add_trace(go.Bar(x=r.Periodo, y=r.Valor_Neto, name='Valor neto'), secondary_y=False)
    fig.add_trace(go.Scatter(x=r.Periodo, y=r['Crecimiento_%'], name='Crecimiento %'), secondary_y=True)
    fig.update_layout(title='Facturación y crecimiento'); grafico(fig)
    mensual = cierre_mensual(df)
    mensual['Año'] = mensual.Periodo.str[:4]
    mensual['Mes'] = mensual.Periodo.str[5:7].astype(int)
    grafico(px.line(mensual, x='Mes', y='Valor_Neto', color='Año', markers=True, title='Comparación entre años').update_xaxes(tickvals=list(range(1, 13)), ticktext=MESES))


def mostrar_estados(df):
    if 'Estado_Factura' not in df:
        st.info('No se encontró Estado_Factura.'); return
    r = agrupar(df, 'Estado_Factura', st.session_state.sesion)
    r['Facturas_%_del_total'] = dividir(r.Facturas, facturas(df)) * 100
    st.caption('Un ID presente en más de un estado se cuenta en cada estado; los porcentajes pueden superar 100 %.')
    tabla(r)
    grafico(px.bar(r, x='Estado_Factura', y='Facturas', title='Facturas únicas por estado'))
    for c in ['Periodo_Cierre', 'Entidad_Pagadora', 'Sede']:
        if c in df:
            cruce = pd.crosstab(df[c], df.Estado_Factura)
            tabla(cruce.reset_index())
            grafico(px.bar(cruce.head(30), barmode='stack', title=f'Estados por {c} · número de registros'))
            normalizado = cruce.div(cruce.sum(axis=1).replace(0, np.nan), axis=0) * 100
            grafico(px.bar(normalizado.head(30), barmode='stack', title=f'Estados por {c} · % de registros'))
    heatmap(df, 'Estado_Factura', 'Periodo_Cierre')


def mostrar_glosas(df):
    if not disponible(df, 'Glosa_Inicial'):
        st.info('No se encontró Glosa_Inicial numérica.'); return
    positivos = df[df.Glosa_Inicial > 0]
    k = calcular_kpis(df)
    cols = st.columns(4)
    valores = [('Valor glosado', formato(k['Glosa_Inicial'], True)), ('Facturas con glosa', str(facturas(positivos))), ('Frecuencia de glosa', formato(dividir(facturas(positivos)*100, facturas(df)), porcentaje=True)), ('Severidad de glosa', formato(k['Glosa_%'], porcentaje=True))]
    for col, (etiqueta, valor) in zip(cols, valores): col.metric(etiqueta, valor)
    st.write('Entre registros con glosa positiva: promedio', formato(positivos.Glosa_Inicial.mean(), True), '· mediana', formato(positivos.Glosa_Inicial.median(), True), '· máximo', formato(positivos.Glosa_Inicial.max(), True))
    evolucion(df, ['Glosa_Inicial'])
    for grupo in ['Entidad_Pagadora', 'Servicio', 'Sede', 'Soporte_Completo']:
        if grupo in df:
            r = agrupar(df, grupo, st.session_state.sesion).sort_values('Glosa_Inicial', ascending=False).head(15)
            grafico(px.bar(r, x=grupo, y='Glosa_Inicial', title=f'Glosa por {grupo}'))
    grafico(px.histogram(muestra(df), x='Glosa_Inicial', title='Distribución de glosas'))
    grafico(px.box(muestra(df), y='Glosa_Inicial', x='Soporte_Completo' if 'Soporte_Completo' in df else None, title='Glosas y soporte', points=False))
    tabla(positivos.sort_values('Glosa_Inicial', ascending=False))
    st.caption('Una asociación entre soporte y glosa no demuestra causalidad. La frecuencia se calcula sobre ID únicos; los estadísticos de valor usan registros.')


def mostrar_cartera(df):
    if not any(disponible(df, c) for c in ['Valor_Pagado', 'Saldo_Cartera']):
        st.info('Se necesita Valor_Pagado o Saldo_Cartera.'); return
    mostrar_kpis(df); evolucion(df)
    por_factura = agrupar(df, 'ID_Factura', st.session_state.sesion)
    if 'Valor_Pagado' in por_factura:
        total = ((por_factura.Valor_Pagado >= por_factura.Valor_Neto) & (por_factura.Valor_Neto > 0)).sum()
        parcial = ((por_factura.Valor_Pagado > 0) & (por_factura.Valor_Pagado < por_factura.Valor_Neto)).sum()
        st.write(f'ID con pago ≥ neto positivo: {total}; con pago parcial: {parcial}. Clasificación por importes, independiente del estado declarado.')
        grafico(px.scatter(muestra(df), x='Valor_Neto', y='Valor_Pagado', title='Facturación y pagos por registro'))
    if disponible(df, 'Saldo_Cartera'):
        saldo = df[df.Saldo_Cartera > 0]
        st.write(f'Facturas con saldo: {facturas(saldo)}. Saldo positivo promedio por registro: {formato(saldo.Saldo_Cartera.mean(), True)}; mediana: {formato(saldo.Saldo_Cartera.median(), True)}.')
        tabla(saldo.sort_values('Saldo_Cartera', ascending=False))
        for grupo in ['Entidad_Pagadora', 'Servicio', 'Sede']:
            if grupo in df:
                r = agrupar(df, grupo, st.session_state.sesion).sort_values('Saldo_Cartera', ascending=False)
                grafico(px.bar(r.head(15), x=grupo, y='Saldo_Cartera', title=f'Cartera por {grupo}'))
                if grupo == 'Entidad_Pagadora' and 'Recaudo_%' in r:
                    st.write('Alto volumen (≥ mediana) y recaudo inferior a la mediana del grupo')
                    tabla(r[(r.Valor_Neto >= r.Valor_Neto.median()) & (r['Recaudo_%'] < r['Recaudo_%'].median())])
        if 'Fecha_Radicacion' in df and pd.api.types.is_datetime64_any_dtype(df.Fecha_Radicacion):
            corte = st.date_input('Fecha de corte para antigüedad aproximada', value=date.today())
            limites = st.text_input('Límites de antigüedad (días, separados por coma)', '30,60,90,180')
            try:
                limites = [int(x.strip()) for x in limites.split(',')]
                if not limites or limites != sorted(set(limites)) or limites[0] < 1: raise ValueError()
                edad = (pd.Timestamp(corte) - saldo.Fecha_Radicacion).dt.days
                etiquetas = [f'0–{limites[0]}'] + [f'{a+1}–{b}' for a,b in zip(limites, limites[1:])] + [f'>{limites[-1]}']
                antiguedad = pd.cut(edad, [-1] + limites + [np.inf], labels=etiquetas).astype('string').fillna('Sin fecha / posterior al corte')
                r = saldo.assign(Antiguedad=antiguedad).groupby('Antiguedad').Saldo_Cartera.sum().reindex(etiquetas + ['Sin fecha / posterior al corte'])
                grafico(px.bar(r, title='Antigüedad del saldo informado desde radicación'))
                st.caption('Rangos analíticos editables. No miden mora contractual ni reconstruyen el saldo histórico a la fecha de corte.')
            except ValueError:
                st.warning('Ingrese límites enteros positivos y crecientes, por ejemplo 30,60,90,180.')


def estadisticos(s):
    s = s.dropna()
    if s.empty: return pd.DataFrame()
    q1, q3 = s.quantile([.25, .75])
    moda = s.mode()
    valores = {'Conteo': len(s), 'Media': s.mean(), 'Mediana': s.median(), 'Moda (primera si hay empate)': moda.iloc[0] if len(moda) < len(s) else np.nan, 'Desviación estándar': s.std(), 'Varianza': s.var(), 'Mínimo': s.min(), 'Máximo': s.max(), 'Rango': s.max()-s.min(), 'Cuartil 1': q1, 'Cuartil 3': q3, 'IQR': q3-q1, 'Percentil 75': q3, 'Percentil 90': s.quantile(.9), 'Percentil 95': s.quantile(.95), 'Coeficiente de variación %': dividir(s.std(), abs(s.mean()))*100, 'Asimetría': s.skew(), 'Curtosis (exceso)': s.kurt()}
    return pd.DataFrame(valores.items(), columns=['Estadístico', 'Valor'])


def mostrar_tiempos(df):
    c = 'Dias_Emision_Radicacion'
    if not disponible(df, c): st.info('No hay días de radicación ni fechas suficientes para calcularlos.'); return
    tabla(estadisticos(df[c]))
    grafico(px.histogram(muestra(df), x=c, title='Distribución de tiempos de radicación'))
    grafico(px.box(muestra(df), y=c, title='Tiempos de radicación', points=False))
    if 'Rango_Tiempo_Radicacion' in df:
        r = df.Rango_Tiempo_Radicacion.value_counts(sort=False).rename_axis('Rango').reset_index(name='Registros')
        grafico(px.bar(r, x='Rango', y='Registros', title='Rangos analíticos de radicación'))
    for grupo in ['Entidad_Pagadora', 'Sede', 'Canal_Radicacion', 'Servicio', 'Estado_Factura', 'Soporte_Completo', 'Periodo_Cierre']:
        if grupo in df:
            r = df.groupby(grupo, observed=True)[c].agg(['count', 'mean', 'median', 'max']).reset_index()
            tabla(r)
            grafico(px.line(r.sort_values(grupo), x=grupo, y='mean', title='Tiempo promedio por periodo') if grupo == 'Periodo_Cierre' else px.bar(r.nlargest(15, 'mean'), x=grupo, y='mean', title=f'Tiempo promedio por {grupo}'))
    tabla(df.sort_values(c, ascending=False))
    st.caption('Los días negativos se conservan y se alertan. Si existe el dato reportado se utiliza; el calculado desde fechas se conserva para comparación.')


def columnas_numericas(df):
    return [c for c in df.select_dtypes(include='number') if df[c].notna().sum() >= 1]


def ejecutar_prueba(df):
    st.caption('Análisis opcional: requiere observaciones independientes. No ajuste los filtros para buscar significancia; pruebas múltiples requieren corrección adicional.')
    prueba = st.selectbox('Prueba', ['Mann-Whitney U', 't de Student', 'Kruskal-Wallis', 'ANOVA', 'Chi-cuadrado', 'Shapiro-Wilk'])
    alpha = st.number_input('Nivel de significancia', .001, .2, .05, .01)
    numericas = columnas_numericas(df)
    categorias = [c for c in CATEGORIAS if c in df]
    variable = st.selectbox('Variable para prueba', numericas) if prueba != 'Chi-cuadrado' else None
    grupo = st.selectbox('Agrupación', categorias) if prueba != 'Shapiro-Wilk' and categorias else None
    segundo = st.selectbox('Segunda variable categórica', [c for c in categorias if c != grupo]) if prueba == 'Chi-cuadrado' and len(categorias) > 1 else None
    elegidos = None
    if grupo and prueba in ['Mann-Whitney U', 't de Student', 'Kruskal-Wallis', 'ANOVA']:
        opciones = sorted(df[grupo].dropna().astype(str).unique())
        elegidos = st.multiselect('Grupos a comparar', opciones, default=opciones[:2] if prueba in ['Mann-Whitney U', 't de Student'] else opciones[:5])
    independientes = st.checkbox('Confirmo que las observaciones pueden tratarse como independientes para esta prueba')
    if not st.button('Ejecutar prueba', type='primary'): return
    if not independientes:
        st.warning('Valide la independencia antes de ejecutar. Filas de una misma factura pueden estar relacionadas.'); return
    try:
        with warnings.catch_warnings():
            warnings.simplefilter('ignore')
            if prueba == 'Shapiro-Wilk':
                s = df[variable].dropna()
                if len(s) < 3 or s.nunique() < 2:
                    st.warning('Se requieren al menos tres observaciones y variabilidad.'); return
                original_n = len(s)
                s = s.sample(min(len(s), 5000), random_state=42)
                resultado = stats.shapiro(s)
                hipotesis = 'H0: la variable sigue una distribución normal. H1: se aparta de la normalidad.'
                st.write(f'Tamaño utilizado: {len(s)} de {original_n}; muestra reproducible si supera 5.000.')
            elif prueba == 'Chi-cuadrado':
                if not grupo or not segundo:
                    st.warning('Se requieren dos variables categóricas.'); return
                t = pd.crosstab(df[grupo], df[segundo])
                if min(t.shape) < 2:
                    st.warning('No existen suficientes grupos para realizar la prueba.'); return
                estadistico, p, _, esperadas = stats.chi2_contingency(t)
                st.write(f'Muestra: {int(t.to_numpy().sum())} registros.'); tabla(t.reset_index())
                if (esperadas < 1).any() or (esperadas < 5).mean() > .2:
                    st.warning('Frecuencias esperadas insuficientes: no se interpreta la aproximación chi-cuadrado. Agrupe categorías con criterio sustantivo.'); return
                resultado = (estadistico, p)
                hipotesis = 'H0: las variables son independientes. H1: existe asociación.'
            else:
                if not grupo or not elegidos:
                    st.warning('Seleccione suficientes grupos.'); return
                grupos = [df.loc[df[grupo].astype(str) == x, variable].dropna().to_numpy() for x in elegidos]
                dos = prueba in ['t de Student', 'Mann-Whitney U']
                if (dos and len(grupos) != 2) or (not dos and len(grupos) < 3) or any(len(x) < 3 for x in grupos):
                    st.warning('Se requieren dos grupos para t/Mann-Whitney o al menos tres para ANOVA/Kruskal, con tres datos válidos por grupo.'); return
                if any(np.std(x) == 0 for x in grupos):
                    st.warning('Hay grupos constantes; la prueba no se ejecuta.'); return
                st.write('Tamaños:', dict(zip(elegidos, map(len, grupos))))
                if prueba in ['t de Student', 'ANOVA']:
                    rng = np.random.default_rng(42)
                    normalidad = [stats.shapiro(rng.choice(x, min(5000, len(x)), replace=False)).pvalue for x in grupos]
                    homogeneidad = stats.levene(*grupos).pvalue
                    st.write('p Shapiro por grupo:', dict(zip(elegidos, normalidad)), '· p Levene:', float(homogeneidad))
                    st.caption('Shapiro usa hasta 5.000 observaciones por grupo. No rechazar no prueba normalidad; revise también los gráficos y el diseño de los datos.')
                    if min(normalidad) < alpha or homogeneidad < alpha or not np.isfinite(homogeneidad):
                        st.warning('Los diagnósticos no respaldan los supuestos paramétricos. Seleccione Mann-Whitney o Kruskal-Wallis.'); return
                funciones = {'Mann-Whitney U': stats.mannwhitneyu, 't de Student': stats.ttest_ind, 'Kruskal-Wallis': stats.kruskal, 'ANOVA': stats.f_oneway}
                resultado = funciones[prueba](*grupos)
                hipotesis = 'H0: las medias poblacionales son iguales. H1: al menos una difiere.' if prueba in ['t de Student', 'ANOVA'] else 'H0: las distribuciones de los grupos son iguales. H1: al menos una difiere. Una diferencia de medianas requiere supuestos adicionales de forma.'
        estadistico, p = float(resultado[0]), float(resultado[1])
        if not np.isfinite(p):
            st.warning('No fue posible obtener un p-valor válido.'); return
        st.write(hipotesis)
        st.write(f'Estadístico: {estadistico:.6g} · p-valor: {p:.6g} · α: {alpha}')
        st.info('Se rechaza H0: se observa evidencia estadística contra la hipótesis nula.' if p < alpha else 'No se rechaza H0: los datos no aportan evidencia suficiente contra ella. Esto no demuestra que H0 sea verdadera.')
        st.caption('El resultado no establece causalidad ni mide por sí solo relevancia financiera.')
    except Exception:
        st.warning('No fue posible ejecutar la prueba con esta selección. Revise tamaños, variabilidad y categorías.')


def mostrar_estadistica(df):
    numericas = columnas_numericas(df)
    if not numericas:
        st.info('No hay variables numéricas válidas.'); return
    tab = st.radio('Análisis', ['Descriptivos', 'Correlaciones', 'Pruebas estadísticas'], horizontal=True)
    if tab == 'Pruebas estadísticas': ejecutar_prueba(df); return
    if tab == 'Descriptivos':
        variable = st.selectbox('Variable numérica', numericas, index=numericas.index('Valor_Neto') if 'Valor_Neto' in numericas else 0)
        s = df[variable].dropna()
        tabla(estadisticos(s))
        grafico(px.histogram(muestra(df), x=variable, title='Distribución observada'))
        grafico(px.box(muestra(df), y=variable, title='Distribución y cuartiles', points=False))
        observado = np.sort(s.to_numpy())
        indices = np.unique(np.linspace(0, len(s)-1, min(5000, len(s))).astype(int))
        grafico(px.line(x=observado[indices], y=(indices+1)/len(s), title='Distribución acumulada empírica', labels={'x': variable, 'y': 'Proporción acumulada'}))
        st.write(f'Los datos muestran {len(s):,} registros válidos, media {formato(s.mean())} y mediana {formato(s.median())}. El 50 % central está entre {formato(s.quantile(.25))} y {formato(s.quantile(.75))}. No se presume normalidad.')
        return
    opciones = [c for c in FIN + ['Cantidad', 'Dias_Emision_Radicacion'] if c in numericas]
    variables = st.multiselect('Variables de correlación', opciones, default=opciones)
    metodo = st.radio('Método', ['Pearson', 'Spearman'], horizontal=True)
    fuerte = st.slider('Umbral de asociación fuerte (|r|)', .3, 1., .7)
    debil = st.slider('Umbral de asociación débil (|r|)', 0., fuerte, min(.3, fuerte))
    st.warning('La correlación no implica causalidad. Algunas relaciones pueden surgir directamente de las fórmulas contables entre las variables.')
    if len(variables) < 2:
        st.info('Seleccione al menos dos variables.'); return
    c = df[variables].corr(method=metodo.lower(), min_periods=3)
    grafico(go.Figure(go.Heatmap(z=c.values, x=c.columns, y=c.index, zmin=-1, zmax=1, colorscale='RdBu', reversescale=True)).update_layout(title=f'Correlaciones · {metodo}'))
    tabla(c.reset_index(names='Variable'))
    pares = []
    for i, a in enumerate(variables):
        for b in variables[i+1:]:
            r = c.loc[a, b]
            if pd.notna(r):
                clasificacion = ('Fuerte positiva' if r > 0 else 'Fuerte negativa') if abs(r) >= fuerte else ('Débil' if abs(r) <= debil else 'Intermedia')
                pares.append([a, b, r, int(df[[a, b]].notna().all(axis=1).sum()), clasificacion])
    tabla(pd.DataFrame(pares, columns=['Variable A', 'Variable B', 'Correlación', 'Pares válidos', 'Asociación']))
    a = st.selectbox('Eje X', variables)
    b = st.selectbox('Eje Y', [x for x in variables if x != a])
    puntos = df[[a, b]].dropna()
    fig = px.scatter(muestra(puntos), x=a, y=b, title='Asociación entre variables')
    if len(puntos) >= 3 and puntos[a].nunique() > 1 and metodo == 'Pearson':
        pendiente, intercepto, *_ = stats.linregress(puntos[a], puntos[b])
        x = np.array([puntos[a].min(), puntos[a].max()])
        fig.add_trace(go.Scatter(x=x, y=intercepto+pendiente*x, mode='lines', name='Tendencia lineal'))
    grafico(fig)


def detectar_atipicos(df, variable='Valor_Neto', metodo='IQR', factor=1.5, percentiles=(.01, .99), grupo=None):
    """Devuelve límites por registro. Grupos con menos de cuatro datos no se evalúan."""
    s = df[variable]
    if grupo:
        g = df.groupby(grupo, dropna=False, observed=True)[variable]
        cuenta = g.transform('count')
        q1, q3 = g.transform(lambda x: x.quantile(.25)), g.transform(lambda x: x.quantile(.75))
        media, desviacion = g.transform('mean'), g.transform('std')
    else:
        cuenta = pd.Series(s.count(), index=df.index)
        q1, q3, media, desviacion = s.quantile(.25), s.quantile(.75), s.mean(), s.std()
    if metodo == 'IQR':
        inferior, superior = q1-factor*(q3-q1), q3+factor*(q3-q1)
    elif metodo == 'Z':
        inferior, superior = media-factor*desviacion, media+factor*desviacion
    else:
        if grupo:
            inferior, superior = g.transform(lambda x: x.quantile(percentiles[0])), g.transform(lambda x: x.quantile(percentiles[1]))
        else:
            inferior, superior = s.quantile(percentiles[0]), s.quantile(percentiles[1])
    evaluable = s.notna() & (cuenta >= 4)
    if metodo == 'Z': evaluable &= desviacion > 0
    salida = pd.DataFrame({'Limite_Inferior': inferior, 'Limite_Superior': superior, 'Evaluable': evaluable}, index=df.index)
    salida['Atipico'] = evaluable & ((s < inferior) | (s > superior))
    return salida


def mostrar_atipicos(df):
    numericas = columnas_numericas(df)
    variable = st.selectbox('Variable a evaluar', numericas, index=numericas.index('Valor_Neto'), key='out_var')
    metodo = st.radio('Método de detección', ['IQR', 'Z', 'Percentiles'], horizontal=True, key='out_met')
    grupo = st.selectbox('Ámbito de detección', ['Global'] + [c for c in ['Servicio', 'Linea_Servicio', 'Entidad_Pagadora'] if c in df], key='out_grupo')
    factor = st.number_input('Multiplicador IQR / umbral Z', .1, 20., 3. if metodo == 'Z' else 1.5, .1, key='out_factor')
    percentiles = st.slider('Percentiles inferior y superior', 0., 100., (1., 99.), key='out_percentiles')
    if percentiles[0] >= percentiles[1]:
        st.warning('Los percentiles deben ser diferentes.'); return
    resultado = detectar_atipicos(df, variable, metodo, factor, tuple(x/100 for x in percentiles), None if grupo == 'Global' else grupo)
    st.session_state.config_atipicos = (variable, metodo, factor, tuple(x/100 for x in percentiles), None if grupo == 'Global' else grupo)
    st.write(f'Atípicos: {int(resultado.Atipico.sum())} / {int(resultado.Evaluable.sum())} evaluables ({formato(dividir(resultado.Atipico.sum()*100, resultado.Evaluable.sum()), porcentaje=True)}).')
    tabla(resultado[['Limite_Inferior', 'Limite_Superior', 'Evaluable']].drop_duplicates())
    tabla(df.loc[resultado.Atipico].join(resultado[['Limite_Inferior', 'Limite_Superior']]))
    visual = muestra(df.assign(Atipico=resultado.Atipico.astype(str)))
    grafico(px.box(visual, y=variable, color='Atipico', title='Atípicos y distribución', points=False))
    grafico(px.histogram(visual, x=variable, color='Atipico', title='Distribución de atípicos'))
    if 'Periodo_Cierre' in visual:
        grafico(px.scatter(visual, x='Periodo_Cierre', y=variable, color='Atipico', title='Atípicos por periodo'))
    st.info('Un valor atípico no es necesariamente un error. Puede representar una factura legítima de mayor complejidad o valor. IQR cero marca desviaciones del valor central; Z es sensible a distribuciones asimétricas.')


def atipicos_actuales(df):
    configuracion = st.session_state.get('config_atipicos', ('Valor_Neto', 'IQR', 1.5, (.01, .99), None))
    if configuracion[0] not in df or (configuracion[4] and configuracion[4] not in df):
        configuracion = ('Valor_Neto', 'IQR', 1.5, (.01, .99), None)
    return detectar_atipicos(df, *configuracion)


def generar_insights(df, informe):
    """Reglas calculadas sin inferencias causales ni servicios externos."""
    filas = []
    def agregar(hallazgo, valor, contexto, interpretacion='El resultado podría justificar una revisión focalizada.', accion='Validar el detalle y los soportes antes de concluir.'):
        filas.append({'Hallazgo calculado': hallazgo, 'Valor observado': str(valor), 'Contexto': contexto, 'Interpretación': interpretacion, 'Recomendación para investigar': accion})
    cierre = cierre_mensual(df)
    if not cierre.empty and cierre.Valor_Neto.notna().any():
        r = cierre.loc[cierre.Valor_Neto.idxmax()]
        agregar('Periodo con mayor facturación', formato(r.Valor_Neto, True), r.Periodo)
        cambios = cierre.dropna(subset=['Crecimiento_%'])
        if not cambios.empty:
            ultimo = cierre.iloc[-1]
            if pd.notna(ultimo['Crecimiento_%']): agregar('Variación del último periodo', formato(ultimo['Crecimiento_%'], porcentaje=True), ultimo.Periodo)
            caidas = cambios[cambios['Crecimiento_%'] < 0]
            if not caidas.empty:
                r = caidas.loc[caidas['Crecimiento_%'].idxmin()]
                agregar('Mayor caída entre periodos consecutivos', formato(r['Crecimiento_%'], porcentaje=True), r.Periodo, accion='Comprobar cobertura y si los periodos están completos.')
        if len(cierre) >= 3 and cierre.Valor_Neto.tail(3).notna().all():
            diferencias = cierre.Valor_Neto.tail(3).diff().dropna()
            if (diferencias > 0).all() or (diferencias < 0).all():
                agregar('Tendencia de tres meses', 'Creciente' if (diferencias > 0).all() else 'Decreciente', 'Tres meses consecutivos disponibles', 'Tendencia observada, sin extrapolación.')
    for grupo in ['Entidad_Pagadora', 'Servicio', 'Sede', 'Canal_Radicacion']:
        if grupo not in df: continue
        r = agrupar(df, grupo, st.session_state.sesion)
        campos = ['Valor_Neto', 'Saldo_Cartera', 'Glosa_%', 'Recaudo_%'] if grupo == 'Entidad_Pagadora' else (['Valor_Neto', 'Saldo_Cartera'] if grupo == 'Servicio' else ['Dias_Emision_Radicacion'])
        for c in campos:
            if c in r and r[c].notna().any():
                fila = r.loc[r[c].idxmin() if c == 'Recaudo_%' else r[c].idxmax()]
                agregar(f'{grupo}: {"menor" if c == "Recaudo_%" else "mayor"} {c}', formato(fila[c], c in FIN, '%' in c), str(fila[grupo]))
        if grupo == 'Entidad_Pagadora':
            for c in ['Valor_Neto', 'Saldo_Cartera']:
                if c in r and r[c].notna().any() and r[c].min() >= 0 and r[c].sum() > 0:
                    agregar('Concentración top 5 · ' + c, formato(r[c].nlargest(5).sum()/r[c].sum()*100, porcentaje=True), f'{len(r)} entidades', 'Mayor concentración implica exposición a un número reducido de pagadores.')
    if disponible(df, 'Glosa_Inicial'):
        agregar('Proporción de facturas glosadas', formato(dividir(facturas(df[df.Glosa_Inicial > 0])*100, facturas(df)), porcentaje=True), 'ID únicos con glosa positiva / ID únicos filtrados')
    if disponible(df, 'Soporte_Incompleto'):
        agregar('Soportes incompletos', formato(df.Soporte_Incompleto.mean()*100, porcentaje=True), f'{df.Soporte_Incompleto.notna().sum()} registros con soporte reconocido')
        if 'Estado_Factura' in df:
            r = df.groupby('Estado_Factura').Soporte_Incompleto.mean().dropna()
            if len(r): agregar('Estado con mayor proporción de soporte incompleto', formato(r.max()*100, porcentaje=True), str(r.idxmax()), 'Asociación descriptiva; no demuestra causa.')
    identidades = informe[informe.Regla.str.startswith('Identidad')]
    for _, r in identidades.iterrows(): agregar(r.Regla, int(r.Casos), f'{int(r.Evaluables)} registros evaluables', 'Diferencias frente a la identidad analítica seleccionada.', 'Conciliar glosas aceptadas, ajustes y pagos; una glosa inicial no siempre reduce cartera.')
    atipicos = atipicos_actuales(df)
    agregar('Registros atípicos', int(atipicos.Atipico.sum()), f'{int(atipicos.Evaluable.sum())} evaluables con configuración actual', 'Un atípico puede ser legítimo.')
    return pd.DataFrame(filas)


def mostrar_alertas(df):
    st.caption('Umbrales demostrativos y editables; no son estándares oficiales ni establecen incumplimientos. Verde: bajo el primer umbral; amarillo: entre umbrales; rojo: alcanza el segundo. En recaudo se invierte el sentido.')
    k = calcular_kpis(df)
    cierre = cierre_mensual(df)
    crecimiento = cierre.iloc[-1]['Crecimiento_%'] if len(cierre) else np.nan
    parametros = [('Glosa_%', 5., 10., False), ('Cartera_%', 30., 50., False), ('Recaudo_%', 80., 60., True), ('Dias_Emision_Radicacion', 7., 15., False), ('Soporte_Incompleto', 10., 20., False), ('Caída_mensual_%', 5., 15., False)]
    k['Caída_mensual_%'] = max(0., -crecimiento) if pd.notna(crecimiento) else np.nan
    resultados = []
    with st.expander('Configurar umbrales de alertas'):
        for nombre, amarillo, rojo, inverso in parametros:
            c1, c2 = st.columns(2)
            a = c1.number_input(nombre + ' · umbral amarillo', 0., value=amarillo, key='alerta_a_' + nombre)
            r = c2.number_input(nombre + ' · umbral rojo', 0., value=rojo, key='alerta_r_' + nombre)
            valor = k[nombre]
            if (inverso and r >= a) or (not inverso and a >= r):
                estado = 'Configuración inválida'
            elif pd.isna(valor): estado = 'No disponible'
            elif inverso: estado = '🔴 Rojo' if valor <= r else '🟡 Amarillo' if valor <= a else '🟢 Verde'
            else: estado = '🔴 Rojo' if valor >= r else '🟡 Amarillo' if valor >= a else '🟢 Verde'
            resultados.append([nombre, valor, estado, a, r])
    tabla(pd.DataFrame(resultados, columns=['Indicador', 'Valor', 'Estado', 'Umbral amarillo', 'Umbral rojo']))


def mostrar_explorador(df, mascaras):
    tipo = st.selectbox('Registros a mostrar', ['Todos', 'Con saldo', 'Con glosa', 'Soportes incompletos', 'Posibles inconsistencias', 'Atípicos'])
    vista = df
    if tipo in ['Con saldo', 'Con glosa', 'Soportes incompletos']:
        campo = {'Con saldo': 'Saldo_Cartera', 'Con glosa': 'Glosa_Inicial', 'Soportes incompletos': 'Soporte_Incompleto'}[tipo]
        if not disponible(df, campo): st.info('No se encontró la columna necesaria.'); return
        vista = df[df[campo] > 0]
    elif tipo == 'Posibles inconsistencias': vista = df[mascaras.any(axis=1)]
    elif tipo == 'Atípicos': vista = df[atipicos_actuales(df).Atipico]
    busqueda = st.text_input('Buscar texto (coincidencia literal)')
    if busqueda:
        mascara = pd.Series(False, index=vista.index)
        for c in vista:
            mascara |= vista[c].astype('string').str.contains(busqueda, case=False, regex=False, na=False)
        vista = vista[mascara]
    columnas = st.multiselect('Columnas visibles', list(df), default=list(df)[:15])
    ordenar = st.selectbox('Ordenar por', list(df))
    ascendente = st.checkbox('Orden ascendente', True)
    n = st.selectbox('Filas visibles', [25, 100, 300, 1000], index=1)
    try: vista = vista.sort_values(ordenar, ascending=ascendente, na_position='last')
    except TypeError: vista = vista.assign(_orden=vista[ordenar].astype(str)).sort_values('_orden', ascending=ascendente).drop(columns='_orden')
    tabla(vista[columnas], n)
    st.caption(f'{len(vista):,} registros encontrados. Los filtros globales siguen activos.')


def neutralizar_formulas(df):
    """Evita fórmulas activas al exportar texto a hojas de cálculo."""
    salida = df.copy()
    salida.columns = ["'" + str(c) if re.match(r'^\s*[=+@-]', str(c)) else str(c) for c in salida.columns]
    for c in salida:
        if not pd.api.types.is_numeric_dtype(salida[c]) and not pd.api.types.is_datetime64_any_dtype(salida[c]):
            s = salida[c].astype('string')
            peligrosa = s.str.match(r'^\s*[=+@-]', na=False)
            s.loc[peligrosa] = "'" + s.loc[peligrosa]
            salida[c] = s
    return salida


@st.cache_data(ttl=900, max_entries=12, show_spinner=False)
def exportar(df, extension, sesion):
    seguro = neutralizar_formulas(df)
    if extension == 'csv':
        return seguro.to_csv(index=False, sep=';', decimal=',').encode('utf-8-sig')
    memoria = BytesIO()
    with pd.ExcelWriter(memoria, engine='openpyxl') as writer:
        # Excel admite 1.048.576 filas incluyendo el encabezado.
        for i, inicio in enumerate(range(0, max(1, len(seguro)), 1_048_575)):
            seguro.iloc[inicio:inicio+1_048_575].to_excel(writer, sheet_name=f'Datos_{i+1}', index=False)
    return memoria.getvalue()


def mostrar_descargas(df, informe, nulos, conversiones, componentes):
    opciones = ['Datos filtrados', 'Cierre mensual', 'Análisis entidades', 'Análisis servicios', 'Análisis sedes', 'Registros con glosa', 'Registros con cartera', 'Registros atípicos', 'Informe de calidad', 'Resumen de insights']
    eleccion = st.selectbox('Contenido a descargar', opciones)
    extension = st.radio('Formato', ['csv', 'xlsx'], horizontal=True)
    nombre = normalizar_nombre_columna(eleccion)
    if st.button('Preparar descarga', type='primary'):
        try:
            with st.spinner('Preparando archivo en memoria…'):
                resultado = df
                if eleccion == 'Cierre mensual': resultado = cierre_mensual(df)
                elif eleccion.startswith('Análisis'):
                    campo = {'Análisis entidades': 'Entidad_Pagadora', 'Análisis servicios': 'Servicio', 'Análisis sedes': 'Sede'}[eleccion]
                    resultado = agrupar(df, campo, st.session_state.sesion)
                elif eleccion in ['Registros con glosa', 'Registros con cartera']:
                    campo = 'Glosa_Inicial' if eleccion.endswith('glosa') else 'Saldo_Cartera'
                    resultado = df[df[campo] > 0] if disponible(df, campo) else pd.DataFrame()
                elif eleccion == 'Registros atípicos': resultado = df[atipicos_actuales(df).Atipico]
                elif eleccion == 'Resumen de insights': resultado = generar_insights(df, informe)
                elif eleccion == 'Informe de calidad':
                    if extension == 'xlsx':
                        memoria = BytesIO()
                        with pd.ExcelWriter(memoria, engine='openpyxl') as writer:
                            for hoja, contenido in [('Reglas', informe), ('Nulos', nulos), ('Conversiones_archivo', conversiones), ('Puntaje', componentes.rename('Puntaje').rename_axis('Componente').reset_index())]:
                                neutralizar_formulas(contenido).to_excel(writer, sheet_name=hoja, index=False)
                        st.download_button('Descargar informe_calidad.xlsx', memoria.getvalue(), 'informe_calidad.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); return
                    resultado = informe
                if resultado.empty:
                    st.info('No hay registros o columnas para esta descarga.'); return
                st.download_button('Descargar ' + eleccion.lower(), exportar(resultado, extension, st.session_state.sesion), f'{nombre}.{extension}', 'text/csv' if extension == 'csv' else 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                if eleccion == 'Resumen de insights':
                    st.download_button('Descargar insights en texto', resultado.to_string(index=False), 'insights_facturacion.txt', 'text/plain')
        except Exception:
            st.error('No fue posible preparar la descarga. Reduzca el conjunto filtrado o pruebe CSV.')
    st.caption('Exportaciones en memoria. Texto que podría ejecutarse como fórmula recibe un apóstrofo de protección. CSV usa punto y coma, coma decimal y UTF-8 con BOM.')


def mostrar_calidad(df, original, conversiones, exclusiones, tolerancia):
    informe, mascaras, nulos, componentes = calidad(df, original, conversiones, tolerancia)
    st.metric('Puntaje orientativo de calidad', f'{componentes.mean():.1f} / 100')
    st.caption('El puntaje es un indicador exploratorio de calidad y no reemplaza una auditoría de datos. Es el promedio de los cuatro componentes; unicidad se basa en filas exactas, no en ID de facturas con detalle.')
    tabla(componentes.rename('Puntaje').rename_axis('Componente').reset_index())
    st.write(f'Registros iniciales: {len(original):,} · Tras limpieza: {len(st.session_state.limpios):,} · Visibles: {len(df):,}. Columnas visibles: {len(df.columns)}.')
    st.write('Columnas encontradas:', ', '.join(c for c in CAMPOS if c in df))
    st.write('Columnas ausentes:', ', '.join(c for c in CAMPOS if c not in df) or 'Ninguna')
    st.subheader('Reglas sobre registros filtrados'); tabla(informe)
    st.subheader('Valores nulos sobre registros filtrados'); tabla(nulos)
    st.subheader('Conversiones sobre el archivo antes de filtros'); tabla(conversiones)
    st.subheader('Exclusiones de limpieza (sin doble conteo)'); tabla(exclusiones)
    st.caption('La identidad de saldo neto − glosa inicial − pagos es una hipótesis de conciliación solicitada. La glosa inicial puede estar en disputa y no reducir el saldo contable. Las diferencias requieren revisión.')
    st.subheader('Registros por revisar'); tabla(df[mascaras.any(axis=1)])


def mostrar_metodologia():
    st.markdown('''
### Unidad de análisis y cobertura
Los importes se **suman por registro**; las facturas se cuentan con ID distintos y no nulos.
Si una factura tiene varias líneas, los importes deben ser aditivos por línea. Un total de factura
repetido en cada línea sobreestima los indicadores; la aplicación no adivina ni elimina esas repeticiones.
Un ID presente en varios grupos se cuenta en cada grupo: sus conteos no siempre son aditivos.
Los ID vacíos se excluyen del conteo, pero sus importes válidos permanecen en las sumas.
Los valores faltantes no se rellenan con cero; las sumas omiten nulos y se debe revisar la cobertura.

### Indicadores
- Facturación neta: suma de Valor_Neto informado; no se reconstruye silenciosamente.
- Identidad de control: Valor_Neto ≈ Valor_Bruto − Descuento − Copago_Cuota.
- Porcentaje de glosa: suma Glosa_Inicial / suma Valor_Neto × 100.
- Porcentaje de recaudo: suma Valor_Pagado / suma Valor_Neto × 100.
- Porcentaje de cartera: suma Saldo_Cartera / suma Valor_Neto × 100.
- Descuento y copago por registro: importe / Valor_Bruto × 100.
- Ticket promedio: suma Valor_Neto / número de facturas únicas.
- Frecuencia de glosa: facturas únicas con glosa positiva / facturas únicas × 100.
- Severidad de glosa: importe glosado / neto × 100.
- Días de radicación: Fecha_Radicacion − Fecha_Emision; se conserva el valor reportado si existe.
- Soportes incompletos: proporción sobre registros con valores sí/no reconocidos.
- Crecimiento: (actual − anterior) / anterior × 100, solo entre periodos consecutivos.
- Acumulado anual: suma de periodos observados del año, dentro de los filtros.
- Promedio móvil: tres periodos consecutivos disponibles (meses, trimestres o años según selección).
- Participación y concentración: importe del grupo / total; top 5 suma las cinco participaciones mayores.

Los denominadores cero producen un valor no disponible, nunca infinito.
Las tasas agregadas se calculan como cocientes de sumas, no como promedios de tasas individuales.
Los cierres, pagos, glosas y saldos se atribuyen al periodo de la factura: **no son un flujo de caja
por fecha de pago ni una reconstrucción de cartera histórica**. Periodo_Cierre tiene prioridad para
los cierres; los filtros de fechas usan radicación si existe y, de lo contrario, el periodo.
La etiqueta COP/USD/EUR es visual: no convierte divisas.

### Estadística exploratoria
La correlación Pearson mide asociación lineal y Spearman asociación monótona por rangos.
Se emplean pares completos y al menos tres observaciones; variables constantes no tienen correlación.
Una asociación no implica causalidad, especialmente entre variables relacionadas por identidades contables.
El p-valor expresa cuán incompatibles resultan los datos con H0 bajo los supuestos de la prueba;
no es la probabilidad de que H0 sea verdadera. No rechazar H0 no la demuestra.
Las pruebas paramétricas revisan normalidad y homogeneidad, pero requieren también independencia
sustantiva. No se aplica corrección automática por comparaciones múltiples.

Un atípico es una observación alejada según el criterio seleccionado, no necesariamente un error.
IQR = Q3 − Q1; límites = Q1 − k × IQR y Q3 + k × IQR, con k inicial 1,5.
Z usa media ± k desviaciones estándar; es sensible a asimetría y no evalúa grupos constantes.
Percentiles usa límites configurados. Se requieren cuatro datos por grupo.
Los atípicos no se excluyen automáticamente.

### Calidad, privacidad y límites
La calidad promedia completitud de celdas, unicidad de filas, éxito de conversión y proporción de
registros sin alertas. Es una aproximación dependiente de las reglas disponibles, no una certificación.
Los periodos de texto deben ser AAAA-MM. Se admiten fechas Excel tipadas y seriales razonables.
La conversión automática deja importes textuales ambiguos como nulos hasta elegir formato regional.
Las transformaciones se hacen sobre copias y se informan. Restaurar conserva el mapeo pero desactiva la limpieza.

Los datos permanecen en la memoria del servidor de Streamlit durante la sesión; no se envían a APIs
ni se escriben en archivos persistentes. La caché en memoria está separada por token de sesión,
tiene TTL de 15 minutos y entradas limitadas; puede existir hasta que expire aunque cierre la pestaña.
Reiniciar limpia el estado de esta sesión; no afecta a otros usuarios.
Solo se aceptan XLSX y CSV. El límite de carga lo determina la configuración de Streamlit;
Excel puede necesitar varias veces su tamaño en memoria. Los gráficos de detalle se limitan a
muestras reproducibles de 5.000 registros; tablas y descargas informan su cobertura.
''')


def reiniciar():
    for k in list(st.session_state): del st.session_state[k]


def main():
    configurar_pagina()
    if 'sesion' not in st.session_state: st.session_state.sesion = uuid.uuid4().hex
    st.sidebar.button('Reiniciar aplicación', on_click=reiniciar)
    original = cargar_archivo()
    pagina = st.sidebar.selectbox('Sección', PAGINAS, key='navegacion')
    st.sidebar.selectbox('Etiqueta de moneda (sin conversión)', ['COP', 'USD', 'EUR', 'Sin símbolo'], key='moneda')
    st.subheader(pagina)
    if original is None:
        st.warning('Evite cargar archivos que contengan nombres de pacientes, documentos de identidad, historias clínicas, diagnósticos u otra información personal o sensible. Utilice archivos anonimizados y siga las políticas de tratamiento de datos de su organización.')
        st.write('Cargue XLSX o CSV desde la barra lateral. Seleccione la hoja, revise la correspondencia y luego explore los indicadores.')
        st.write('**Mínimas:** ID_Factura, Valor_Neto y Fecha_Radicacion o Periodo_Cierre.')
        st.write('**Recomendadas:** ' + ', '.join(RECOMENDADAS))
        if pagina == 'Metodología': mostrar_metodologia()
        return
    if pagina == 'Inicio y carga de datos':
        st.warning('Utilice únicamente archivos anonimizados, sin nombres de pacientes, documentos, diagnósticos ni historias clínicas.')
        st.write(st.session_state.metadatos)
    mapeado = mapear_columnas(original)
    if mapeado is None: return
    opciones = configurar_limpieza()
    try:
        with st.spinner('Validando y preparando los datos…'):
            limpio, conversiones, exclusiones = transformar(mapeado, opciones, st.session_state.sesion)
    except Exception:
        st.error('No fue posible convertir los datos. Revise fechas, formato monetario y correspondencia de columnas.'); return
    st.session_state.limpios = limpio
    if len(conversiones):
        fallos = int(conversiones['No convertidos'].sum())
        if fallos: st.warning(f'{fallos:,} celdas no pudieron convertirse. Consulte Calidad de datos; los importes ambiguos requieren seleccionar formato regional.')
    if len(exclusiones) and exclusiones['Registros excluidos'].sum():
        st.info(f'Se excluyeron {int(exclusiones["Registros excluidos"].sum()):,} registros por las opciones elegidas. Consulte el detalle en Calidad de datos.')
    faltantes = [c for c in ['ID_Factura', 'Valor_Neto'] if c not in limpio]
    if not any(c in limpio for c in ['Fecha_Radicacion', 'Periodo_Cierre']): faltantes.append('Fecha_Radicacion o Periodo_Cierre')
    if faltantes:
        st.error('Faltan columnas mínimas: ' + ', '.join(faltantes))
        st.write('Encontradas:', list(limpio.columns))
        st.info('Corrija los encabezados o asigne las columnas mediante el mapeo. Los análisis principales están desactivados.')
        tabla(original); return
    ausentes = [c for c in RECOMENDADAS if c not in limpio]
    if ausentes: st.warning('Análisis parcialmente disponibles. Faltan: ' + ', '.join(ausentes))
    if limpio.empty: st.warning('No quedan registros después de la limpieza. Restaure los datos o revise las exclusiones.'); return
    df = aplicar_filtros(limpio)
    if df.empty: st.warning('Los filtros actuales no tienen registros. Use «Limpiar filtros».'); return
    informe, mascaras, nulos, componentes = calidad(df, mapeado, conversiones, opciones['tolerancia'])
    if pagina == 'Metodología': mostrar_metodologia(); return
    if pagina == 'Calidad de datos': mostrar_calidad(df, mapeado, conversiones, exclusiones, opciones['tolerancia']); return
    if pagina == 'Inicio y carga de datos':
        st.write(f'{len(original):,} filas originales · {len(df):,} visibles · {len(original.columns)} columnas originales')
        st.write({c: int(df[c].nunique()) for c in ['Entidad_Pagadora', 'Sede', 'Servicio'] if c in df})
        if 'Periodo_Cierre' in df and df.Periodo_Cierre.notna().any():
            st.write('Periodo identificado:', str(df.Periodo_Cierre.dropna().min()), 'a', str(df.Periodo_Cierre.dropna().max()))
        aleatoria = st.checkbox('Vista previa aleatoria')
        tabla(df.sample(min(25, len(df)), random_state=42) if aleatoria else df.head(25))
        st.info('Los importes se suman por fila. Si cada línea repite el total de una factura, corrija la granularidad antes de analizar. Los ID repetidos se señalan, pero no se eliminan automáticamente.')
        return
    if not disponible(df, 'Valor_Neto') or not df.Valor_Neto.notna().any():
        st.warning('Valor_Neto no contiene datos numéricos utilizables. Active las conversiones o seleccione el formato correcto.'); return
    if not opciones['aplicar']:
        st.info('Datos originales restaurados: las variables derivadas y algunos análisis requieren activar las transformaciones.')
    st.caption(f'Base filtrada: {len(df):,} registros, {facturas(df):,} facturas únicas · Neto filtrado {formato(sumar(df, "Valor_Neto"), True)} / completo tras limpieza {formato(sumar(limpio, "Valor_Neto"), True)}.')
    if df.ID_Factura.duplicated().any(): st.caption('Hay ID repetidos: los conteos usan ID únicos; los importes se suman por registro.')
    if pagina == 'Resumen ejecutivo':
        mostrar_kpis(df); evolucion(df)
        evolucion(df, ['Glosa_Inicial', 'Valor_Pagado'], 'Glosas y recaudo por periodo')
        for grupo in ['Estado_Factura', 'Entidad_Pagadora', 'Linea_Servicio']:
            if grupo in df:
                r = agrupar(df, grupo, st.session_state.sesion).sort_values('Valor_Neto', ascending=False).head(15)
                grafico(px.bar(r, x=grupo, y='Valor_Neto', title=f'Facturación por {grupo}'))
    elif pagina == 'Cierre mensual': mostrar_cierre(df)
    elif pagina == 'Entidades pagadoras': mostrar_grupo(df, 'Entidad_Pagadora')
    elif pagina in ['Servicios y líneas', 'Sedes y regímenes']:
        candidatos = ['Servicio', 'Linea_Servicio'] if pagina == 'Servicios y líneas' else ['Sede', 'Regimen']
        disponibles = [c for c in candidatos if c in df]
        if disponibles: mostrar_grupo(df, st.radio('Dimensión', disponibles, horizontal=True))
        else: st.info('No se encontraron columnas para este análisis: ' + ', '.join(candidatos))
    elif pagina == 'Estados de facturación': mostrar_estados(df)
    elif pagina == 'Glosas': mostrar_glosas(df)
    elif pagina == 'Recaudo y cartera': mostrar_cartera(df)
    elif pagina == 'Tiempos de radicación': mostrar_tiempos(df)
    elif pagina == 'Análisis estadístico': mostrar_estadistica(df)
    elif pagina == 'Valores atípicos': mostrar_atipicos(df)
    elif pagina == 'Insights y alertas':
        mostrar_alertas(df)
        insights = generar_insights(df, informe)
        for _, fila in insights.iterrows():
            with st.expander(f'{fila["Hallazgo calculado"]}: {fila["Valor observado"]}'):
                for c in ['Contexto', 'Interpretación', 'Recomendación para investigar']: st.write(f'**{c}:** {fila[c]}')
    elif pagina == 'Explorador de datos': mostrar_explorador(df, mascaras)
    elif pagina == 'Descargas': mostrar_descargas(df, informe, nulos, conversiones, componentes)


if __name__ == '__main__':
    try:
        main()
    except Exception:
        st.error('No fue posible completar este análisis. Revise las columnas y los filtros, o cambie de sección. Sus datos originales permanecen disponibles durante la sesión.')
