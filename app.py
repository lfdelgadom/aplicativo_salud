"""Analitica de facturacion en salud. Ejecutar: streamlit run app.py.
Todos los archivos del usuario se procesan en memoria; no hay APIs ni datasets.
"""
from io import BytesIO, StringIO
from datetime import date
import csv
import hashlib
import inspect
import re
import unicodedata
import uuid

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from scipy import stats
import streamlit as st

TITULO = 'ANAL\u00cdTICA DE FACTURACI\u00d3N EN SALUD'
CAMPOS = ['ID_Factura','Numero_Radicado','Fecha_Servicio','Fecha_Emision',
          'Fecha_Radicacion','Fecha_Aceptacion','Fecha_Pago','Periodo_Cierre',
          'Codigo_Entidad','Entidad_Pagadora','Regimen','Sede','Codigo_Servicio',
          'Servicio','Linea_Servicio','Cantidad','Canal_Radicacion','Estado_Factura',
          'Dias_Emision_Radicacion','Soporte_Completo','Valor_Bruto','Descuento',
          'Copago_Cuota','Valor_Neto','Glosa_Inicial','Valor_Pagado','Saldo_Cartera']
FECHAS = [c for c in CAMPOS if c.startswith('Fecha_')]
DINERO = ['Valor_Bruto','Descuento','Copago_Cuota','Valor_Neto','Glosa_Inicial','Valor_Pagado','Saldo_Cartera']
NUMEROS = DINERO + ['Cantidad','Dias_Emision_Radicacion']
CATEGORIAS = ['Entidad_Pagadora','Regimen','Sede','Servicio','Linea_Servicio','Estado_Factura','Canal_Radicacion','Soporte_Completo']
RECOMENDADAS = ['Entidad_Pagadora','Estado_Factura','Valor_Pagado','Saldo_Cartera','Glosa_Inicial']
PAGINAS = ['Inicio y carga de datos','Calidad de datos','Resumen ejecutivo','Cierre mensual',
           'Entidades pagadoras','Servicios y lineas','Sedes y regimenes','Estados de facturacion',
           'Glosas','Recaudo y cartera','Tiempos de radicacion','Analisis estadistico',
           'Valores atipicos','Insights y alertas','Explorador de datos','Descargas','Metodologia']
MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
DIAS = ['Lunes','Martes','Miercoles','Jueves','Viernes','Sabado','Domingo']
PRIVACIDAD = ('Evite cargar archivos que contengan nombres de pacientes, documentos de identidad, '
             'historias clinicas, diagnosticos u otra informacion personal o sensible. '
             'Utilice archivos anonimizados y siga las politicas de tratamiento de datos de su organizacion.')
AVISO = ('Los resultados dependen de la calidad y estructura del archivo cargado. Esta herramienta '
         'facilita el analisis exploratorio y no reemplaza una auditoria financiera, contable, medica o contractual.')


def cache_memoria(func):
    """Cache sin persistencia. En versiones recientes se limita a la sesion.
    En versiones anteriores se aisla con un token de sesion y TTL corto.
    """
    opts = dict(show_spinner=False, ttl=600, max_entries=16, persist=False)
    if 'scope' in inspect.signature(st.cache_data).parameters:
        opts['scope'] = 'session'
    return st.cache_data(**opts)(func)


def normalizar_nombre_columna(nombre):
    texto = unicodedata.normalize('NFKD', str(nombre).strip())
    texto = ''.join(c for c in texto if not unicodedata.combining(c)).lower()
    return re.sub(r'[^a-z0-9]+', '_', texto).strip('_')


ALIASES = {normalizar_nombre_columna(c): c for c in CAMPOS}
for alias, campo in {'factura':'ID_Factura','numero_factura':'ID_Factura',
                     'nro_factura':'ID_Factura','radicado':'Numero_Radicado',
                     'pagador':'Entidad_Pagadora','entidad':'Entidad_Pagadora',
                     'linea_de_servicio':'Linea_Servicio','copago':'Copago_Cuota',
                     'estado':'Estado_Factura','glosa':'Glosa_Inicial',
                     'pago':'Valor_Pagado','cartera':'Saldo_Cartera',
                     'periodo':'Periodo_Cierre','dias_radicacion':'Dias_Emision_Radicacion'}.items():
    ALIASES[alias] = campo


def detectar_columnas(df):
    """Sugiere solo equivalencias univocas; no renombra silenciosamente."""
    candidatos = {c: [] for c in CAMPOS}
    for col in df.columns:
        campo = ALIASES.get(normalizar_nombre_columna(col))
        if campo:
            candidatos[campo].append(col)
    return {c: v[0] if len(v) == 1 else None for c, v in candidatos.items()}, candidatos


@cache_memoria
def hojas_excel(contenido, token):
    with pd.ExcelFile(BytesIO(contenido), engine='openpyxl') as libro:
        return libro.sheet_names


@cache_memoria
def cargar_archivo(contenido, extension, hoja, separador, codificacion, token):
    """Lectura en memoria. CSV conserva cadenas para validar conversiones."""
    if extension == 'xlsx':
        df = pd.read_excel(BytesIO(contenido), sheet_name=hoja, engine='openpyxl', dtype=object)
        info = f'Excel / hoja {hoja}'
    else:
        encs = ['utf-8-sig','utf-16','cp1252','latin-1'] if codificacion == 'Automatica' else [codificacion]
        texto = None
        for enc in encs:
            try:
                texto = contenido.decode(enc)
                if '\x00' in texto and enc != 'utf-16':
                    continue
                break
            except UnicodeError:
                texto = None
        if texto is None:
            raise ValueError('No se pudo detectar la codificacion; seleccione otra.')
        sep = {'Coma':',','Punto y coma':';','Tabulacion':'\t'}.get(separador)
        if sep is None:
            try:
                sep = csv.Sniffer().sniff(texto[:20000], delimiters=',;\t|').delimiter
            except csv.Error:
                sep = ';' if texto.splitlines() and ';' in texto.splitlines()[0] else ','
        df = pd.read_csv(StringIO(texto), sep=sep, dtype=object)
        info = f'CSV / codificacion {enc} / separador {repr(sep)}'
    df.columns = [str(c) for c in df.columns]
    return df, info


def convertir_monetaria(serie, formato='Automatico'):
    """Conversion vectorizada. Un separador con tres digitos es ambiguo.
    En modo automatico se interpreta como miles y se registra advertencia.
    """
    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(serie, errors='coerce').replace([np.inf,-np.inf], np.nan), 0
    texto = serie.astype('string').str.strip()
    negativo = texto.str.match(r'^\(.*\)$', na=False)
    texto = texto.str.replace(r'[^0-9,\.\-+]', '', regex=True)
    ambos = texto.str.contains(',', na=False) & texto.str.contains(r'\.', na=False)
    ambiguos = texto.str.match(r'^[+-]?\d+[,.]\d{3}$', na=False)
    if formato == 'Colombiano (1.234,56)':
        limpio = texto.str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
    elif formato == 'Internacional (1,234.56)':
        limpio = texto.str.replace(',', '', regex=False)
    else:
        limpio = texto.copy()
        coma_decimal = ambos & (texto.str.rfind(',') > texto.str.rfind('.'))
        punto_decimal = ambos & ~coma_decimal
        limpio.loc[coma_decimal] = texto.loc[coma_decimal].str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
        limpio.loc[punto_decimal] = texto.loc[punto_decimal].str.replace(',', '', regex=False)
        miles = ~ambos & (ambiguos | texto.str.match(r'^[+-]?\d{1,3}([,.]\d{3}){2,}$', na=False))
        limpio.loc[miles] = texto.loc[miles].str.replace(r'[,.]', '', regex=True)
        solo_coma = ~ambos & ~miles & texto.str.contains(',', na=False)
        limpio.loc[solo_coma] = texto.loc[solo_coma].str.replace(',', '.', regex=False)
    salida = pd.to_numeric(limpio, errors='coerce').astype(float)
    salida.loc[negativo] = -salida.loc[negativo].abs()
    # Los numeros reales almacenados en Excel no deben perder sus decimales.
    reales = serie.map(lambda x: isinstance(x, (int,float,np.number)) and not isinstance(x,bool))
    salida.loc[reales] = pd.to_numeric(serie.loc[reales], errors='coerce')
    return salida.replace([np.inf,-np.inf], np.nan), int((ambiguos & ~reales).sum())


def convertir_fecha(serie, orden='Dia primero'):
    """Reconoce fechas Excel y cadenas, con errores convertidos a NaT."""
    if pd.api.types.is_datetime64_any_dtype(serie):
        return pd.to_datetime(serie, errors='coerce')
    salida = pd.to_datetime(serie.astype('string'), errors='coerce', format='mixed', dayfirst=orden == 'Dia primero')
    reales = serie.map(lambda x: isinstance(x,(int,float,np.number)) and not isinstance(x,bool))
    if reales.any():
        numeros = pd.to_numeric(serie.loc[reales], errors='coerce')
        salida.loc[reales] = pd.to_datetime(numeros, unit='D', origin='1899-12-30', errors='coerce')
    return salida


def periodo_valido(serie):
    texto = serie.astype('string').str.strip()
    exacto = texto.str.match(r'^\d{4}-(0[1-9]|1[0-2])$', na=False)
    fecha = pd.to_datetime(texto.where(exacto), format='%Y-%m', errors='coerce')
    # Las celdas Excel de fecha tambien son aceptables, pero no meses ambiguos.
    es_fecha = serie.map(lambda x: isinstance(x, (pd.Timestamp,date)))
    fecha.loc[es_fecha] = pd.to_datetime(serie.loc[es_fecha], errors='coerce')
    return fecha.dt.strftime('%Y-%m').astype('string')


@cache_memoria
def transformar(original, mapeo, opciones, formato, orden, limites, token):
    """Copia controlada, registro de conversiones y exclusiones."""
    renombres = {fuente: destino for destino, fuente in mapeo.items() if fuente is not None}
    df = original.rename(columns=renombres).copy()
    conversiones, exclusiones = [], []
    if opciones['aplicar']:
        for opcion, mascara, razon in [
            ('vacias', df.isna().all(axis=1), 'Filas completamente vacias'),
            ('duplicados', df.duplicated(), 'Duplicados exactos')]:
            if opciones[opcion]:
                n = int(mascara.reindex(df.index, fill_value=False).sum())
                df = df.loc[~mascara.reindex(df.index, fill_value=False)].copy()
                exclusiones.append({'Razon':razon,'Excluidos':n})
        for col in FECHAS + NUMEROS:
            if col not in df:
                continue
            habilitada = opciones['fechas'] if col in FECHAS else opciones['moneda']
            if not habilitada:
                continue
            anterior = df[col]
            convertido, ambiguos = (convertir_fecha(anterior, orden), 0) if col in FECHAS else convertir_monetaria(anterior, formato)
            presentes = anterior.notna() & anterior.astype('string').str.strip().ne('').fillna(False)
            conversiones.append({'Columna':col,'Fallos':int((presentes & convertido.isna()).sum()),
                                 'Nulos finales':int(convertido.isna().sum()),'Formatos ambiguos':ambiguos})
            df[col] = convertido
        if opciones['categorias']:
            for col in CATEGORIAS:
                if col in df:
                    df[col] = df[col].astype('string').str.strip().replace('', pd.NA).fillna('Sin informacion')
        if opciones['neto'] and 'Valor_Neto' in df:
            valido = convertir_monetaria(df['Valor_Neto'], formato)[0].notna()
            exclusiones.append({'Razon':'Valor neto invalido','Excluidos':int((~valido).sum())})
            df = df.loc[valido].copy()
        if opciones['fecha_invalida']:
            valido = pd.Series(True,index=df.index)
            for col in FECHAS:
                if col in df:
                    presente = df[col].notna() & df[col].astype('string').str.strip().ne('').fillna(False)
                    valido &= ~(presente & convertir_fecha(df[col], orden).isna())
            # Incluye las conversiones que ya quedaron nulas, contrastando originales.
            for col in FECHAS:
                fuente = mapeo.get(col)
                if fuente and fuente in original:
                    raw = original.loc[df.index, fuente]
                    presente = raw.notna() & raw.astype('string').str.strip().ne('').fillna(False)
                    valido &= ~(presente & convertir_fecha(raw, orden).isna())
            exclusiones.append({'Razon':'Fechas invalidas no vacias','Excluidos':int((~valido).sum())})
            df = df.loc[valido].copy()
    df = crear_variables_derivadas(df, limites)
    return df, pd.DataFrame(conversiones), pd.DataFrame(exclusiones)


def numerica(df, col):
    return col in df and pd.api.types.is_numeric_dtype(df[col])


def crear_variables_derivadas(df, limites=(3,7,15)):
    df = df.copy()
    if 'Periodo_Cierre' in df:
        df['Periodo_Cierre'] = periodo_valido(df['Periodo_Cierre'])
    if 'Fecha_Radicacion' in df and pd.api.types.is_datetime64_any_dtype(df['Fecha_Radicacion']):
        f = df['Fecha_Radicacion']
        if 'Periodo_Cierre' not in df:
            df['Periodo_Cierre'] = f.dt.strftime('%Y-%m').astype('string')
        df['Ano_Radicacion'] = f.dt.year
        df['Numero_Mes'] = f.dt.month
        df['Mes_Radicacion'] = f.dt.month.map(dict(enumerate(MESES,1)))
        df['Trimestre_Radicacion'] = f.dt.quarter
        df['Ano_Mes'] = f.dt.strftime('%Y-%m')
        df['Dia_Semana'] = f.dt.dayofweek.map(dict(enumerate(DIAS)))
        for origen, destino in [('Fecha_Emision','Dias_Emision_Radicacion'),('Fecha_Pago','Dias_Radicacion_Pago')]:
            if origen in df and pd.api.types.is_datetime64_any_dtype(df[origen]) and destino not in df:
                df[destino] = ((f-df[origen]) if origen == 'Fecha_Emision' else (df[origen]-f)).dt.days
    for num,den,nombre in [('Descuento','Valor_Bruto','Descuento'),('Copago_Cuota','Valor_Bruto','Copago'),
                            ('Glosa_Inicial','Valor_Neto','Glosa'),('Valor_Pagado','Valor_Neto','Recaudo'),
                            ('Saldo_Cartera','Valor_Neto','Cartera')]:
        if numerica(df,num) and numerica(df,den):
            df['Porcentaje_'+nombre] = (df[num]/df[den].replace(0,np.nan)*100).replace([np.inf,-np.inf],np.nan)
    for col,nombre in [('Glosa_Inicial','Tiene_Glosa'),('Saldo_Cartera','Tiene_Saldo')]:
        if numerica(df,col):
            df[nombre] = df[col].gt(0)
    if 'Estado_Factura' in df:
        df['Es_Pagada'] = df['Estado_Factura'].astype('string').str.strip().str.lower().eq('pagada')
    if numerica(df,'Dias_Emision_Radicacion'):
        a,b,c = limites
        df['Rango_Tiempo_Radicacion'] = pd.cut(df['Dias_Emision_Radicacion'],
            [-np.inf,-0.00001,a,b,c,np.inf], labels=['Negativo',f'0 a {a} dias',f'Mas de {a} a {b}',f'Mas de {b} a {c}',f'Mas de {c} dias'])
    return df


@cache_memoria
def validar_datos(df, tolerancia=1., formato='Automatico', orden='Dia primero', token=''):
    """Diagnostico sobre datos originales mapeados; no elimina ni corrige."""
    n = len(df)
    nulos = df.isna() | df.astype('string').apply(lambda s: s.str.strip().eq('').fillna(False))
    calidad = pd.DataFrame({'Columna':df.columns,'Nulos':nulos.sum().values,
                            'Porcentaje nulos':(nulos.mean()*100).values})
    reglas, mascaras = [], {}
    def agregar(nombre, mask):
        mask = mask.fillna(False).astype(bool)
        reglas.append({'Regla':nombre,'Registros':int(mask.sum()),'Porcentaje':100*mask.mean() if n else np.nan})
        mascaras[nombre] = mask
    agregar('Filas duplicadas exactas (excedentes)',df.duplicated())
    for col in ['ID_Factura','Numero_Radicado']:
        if col in df:
            agregar(col+' duplicado (excedentes)',df[col].notna() & df[col].duplicated())
    fechas, nums = {}, {}
    for col in FECHAS:
        if col in df:
            fechas[col] = convertir_fecha(df[col], orden)
            agregar(col+' invalida', ~nulos[col] & fechas[col].isna())
    for col in NUMEROS:
        if col in df:
            nums[col], _ = convertir_monetaria(df[col],formato)
            agregar(col+' no numerico',~nulos[col] & nums[col].isna())
            agregar(col+' negativo',nums[col].lt(0))
    if 'Valor_Neto' in nums:
        agregar('Valor neto igual a cero',nums['Valor_Neto'].eq(0))
        for col in ['Saldo_Cartera','Valor_Pagado','Glosa_Inicial']:
            if col in nums:
                agregar(col+' superior al valor neto',nums[col].gt(nums['Valor_Neto']+tolerancia))
    if 'Entidad_Pagadora' in df:
        agregar('Sin entidad pagadora',nulos['Entidad_Pagadora'])
    if 'Periodo_Cierre' in df:
        agregar('Sin periodo o periodo invalido',periodo_valido(df['Periodo_Cierre']).isna())
    elif 'Fecha_Radicacion' in fechas:
        agregar('Sin periodo',fechas['Fecha_Radicacion'].isna())
    for a,b,nombre in [('Fecha_Radicacion','Fecha_Emision','Radicacion anterior a emision'),
                       ('Fecha_Pago','Fecha_Radicacion','Pago anterior a radicacion')]:
        if a in fechas and b in fechas:
            agregar(nombre,fechas[a].lt(fechas[b]))
    for columnas,nombre in [(['Valor_Neto','Valor_Bruto','Descuento','Copago_Cuota'],'Identidad de valor neto inconsistente'),
                            (['Saldo_Cartera','Valor_Neto','Glosa_Inicial','Valor_Pagado'],'Identidad de saldo inconsistente')]:
        if all(c in nums for c in columnas):
            a,b,c,d = [nums[x] for x in columnas]
            agregar(nombre,(a-(b-c-d)).abs().gt(tolerancia))
    reporte = pd.DataFrame(reglas)
    invalidez = pd.Series(False,index=df.index)
    inconsistencia = pd.Series(False,index=df.index)
    for nombre,mask in mascaras.items():
        if any(x in nombre for x in ['invalida','no numerico','negativo','Sin periodo']):
            invalidez |= mask
        if any(x in nombre for x in ['inconsistente','anterior','superior']):
            inconsistencia |= mask
    completitud = 100*(1-nulos.to_numpy().mean()) if n and len(df.columns) else 0
    unicidad = 100*(1-df['ID_Factura'].duplicated().mean()) if n and 'ID_Factura' in df else 0
    componentes = {'Completitud':completitud,'Unicidad':unicidad,
                  'Validez':100*(1-invalidez.mean()) if n else 0,
                  'Consistencia':100*(1-inconsistencia.mean()) if n else 0}
    sospechosos = pd.Series(False,index=df.index)
    for mask in mascaras.values():
        sospechosos |= mask
    return calidad,reporte,componentes,sospechosos


def ratio(num,den):
    return float(num/den*100) if pd.notna(num) and pd.notna(den) and den != 0 else np.nan


def total(df,col):
    return df[col].sum(min_count=1) if numerica(df,col) else np.nan


def facturas(df):
    return int(df['ID_Factura'].nunique(dropna=True)) if 'ID_Factura' in df else 0


def calcular_kpis(df):
    k = {c:total(df,c) for c in NUMEROS}
    k['Facturas'] = facturas(df)
    k['Ticket promedio'] = k['Valor_Neto']/k['Facturas'] if k['Facturas'] else np.nan
    for col,nombre in [('Glosa_Inicial','Porcentaje de glosa'),('Valor_Pagado','Porcentaje de recaudo'),('Saldo_Cartera','Porcentaje de cartera')]:
        k[nombre] = ratio(k[col],k['Valor_Neto'])
    k['Tiempo promedio'] = df['Dias_Emision_Radicacion'].mean() if numerica(df,'Dias_Emision_Radicacion') else np.nan
    return k


@cache_memoria
def agrupar(df, grupo, token):
    """Montos se suman por registro; facturas se cuentan con nunique."""
    if grupo not in df or df.empty:
        return pd.DataFrame()
    agregados = {'ID_Factura':pd.Series.nunique}
    agregados.update({c:lambda s:s.sum(min_count=1) for c in NUMEROS if numerica(df,c) and c != 'Dias_Emision_Radicacion'})
    if numerica(df,'Dias_Emision_Radicacion'):
        agregados['Dias_Emision_Radicacion'] = 'mean'
    tabla = df.groupby(grupo,dropna=False,observed=True).agg(agregados).rename(columns={'ID_Factura':'Facturas'}).reset_index()
    if 'Valor_Neto' in tabla:
        tabla['Ticket promedio'] = tabla['Valor_Neto']/tabla['Facturas'].replace(0,np.nan)
        for col,nombre in [('Glosa_Inicial','Glosa'),('Valor_Pagado','Recaudo'),('Saldo_Cartera','Cartera')]:
            if col in tabla:
                tabla['Porcentaje '+nombre] = tabla[col]/tabla['Valor_Neto'].replace(0,np.nan)*100
        tabla['Participacion'] = tabla['Valor_Neto']/tabla['Valor_Neto'].sum()*100 if tabla['Valor_Neto'].sum() else np.nan
    if 'Soporte_Completo' in df:
        aux = df.assign(_incompleto=soporte_incompleto(df['Soporte_Completo']))
        tmp = aux.groupby(grupo,dropna=False,observed=True)['_incompleto'].mean().mul(100).rename('Porcentaje soportes incompletos').reset_index()
        tabla = tabla.merge(tmp,on=grupo,how='left')
    if 'Estado_Factura' in df:
        aux = df.assign(_devuelta=df['Estado_Factura'].astype('string').str.lower().str.contains('devuelt',na=False))
        tmp = aux.groupby(grupo,dropna=False,observed=True)['_devuelta'].sum().rename('Registros devueltos').reset_index()
        tabla = tabla.merge(tmp,on=grupo,how='left')
    return tabla.replace([np.inf,-np.inf],np.nan)


@cache_memoria
def construir_cierre_mensual(df, frecuencia, token):
    if 'Periodo_Cierre' not in df:
        return pd.DataFrame()
    fechas = pd.to_datetime(df['Periodo_Cierre'],format='%Y-%m',errors='coerce')
    copia = df.loc[fechas.notna()].copy()
    copia['Periodo'] = fechas.loc[fechas.notna()].dt.to_period({'Mensual':'M','Trimestral':'Q','Anual':'Y'}[frecuencia]).astype(str)
    if copia.empty:
        return pd.DataFrame()
    tabla = agrupar(copia,'Periodo',token).sort_values('Periodo').reset_index(drop=True)
    if tabla.empty or 'Valor_Neto' not in tabla:
        return tabla
    # Completar periodos faltantes con NaN: no se inventa facturacion cero.
    freq = {'Mensual':'M','Trimestral':'Q','Anual':'Y'}[frecuencia]
    rango = pd.period_range(tabla['Periodo'].min(),tabla['Periodo'].max(),freq=freq).astype(str)
    tabla = tabla.set_index('Periodo').reindex(rango).rename_axis('Periodo').reset_index()
    tabla['Crecimiento porcentual'] = tabla['Valor_Neto'].pct_change(fill_method=None)*100
    tabla['Acumulado anual'] = tabla.groupby(tabla['Periodo'].str[:4])['Valor_Neto'].cumsum()
    tabla['Promedio movil 3 periodos'] = tabla['Valor_Neto'].rolling(3,min_periods=3).mean()
    return tabla.replace([np.inf,-np.inf],np.nan)


def soporte_incompleto(serie):
    return serie.astype('string').map(normalizar_nombre_columna).isin(['no','false','0','incompleto','incompleta','n'])


def formatear_moneda(valor):
    if pd.isna(valor):
        return 'No disponible'
    moneda = st.session_state.get('moneda','COP')
    prefijo = {'COP':'$ ','USD':'USD ','EUR':'EUR ','Sin simbolo':''}[moneda]
    return prefijo + f'{valor:,.0f}'.replace(',','.')


def formatear_porcentaje(valor):
    return f'{valor:.1f} %' if pd.notna(valor) else 'No disponible'


def mostrar_tabla(df, limite=500):
    if df.empty:
        st.info('No hay registros para mostrar.')
        return
    configs = {c:st.column_config.NumberColumn(c,format='%.2f') for c in df.select_dtypes(include='number').columns}
    for c in DINERO:
        if numerica(df,c):
            etiqueta = st.session_state.get('moneda','COP')
            configs[c] = st.column_config.NumberColumn(c,format=('%.2f' if etiqueta=='Sin simbolo' else etiqueta+' %.2f'))
    configs.update({c:st.column_config.DatetimeColumn(c,format='DD/MM/YYYY') for c in df.select_dtypes(include='datetime').columns})
    st.dataframe(df.head(limite),**({'width':'stretch'} if 'width' in inspect.signature(st.dataframe).parameters else {'use_container_width':True}),hide_index=True,column_config=configs,height=min(520,70+35*min(len(df),12)))
    if len(df)>limite:
        st.caption(f'Vista limitada a {limite:,} registros. Descargue el conjunto completo en Descargas.')


def grafico(fig):
    try:
        fig.update_layout(template='plotly_white',font=dict(size=13),colorway=['#126782','#008b8b','#ed9b40','#a23b3b'],margin=dict(l=20,r=20,t=60,b=30))
        st.plotly_chart(fig,**({'width':'stretch'} if 'width' in inspect.signature(st.plotly_chart).parameters else {'use_container_width':True}))
    except Exception:
        st.warning('No fue posible generar este grafico. Los demas analisis siguen disponibles.')


def barras(tabla,grupo,valor,titulo,top=15):
    if tabla.empty or valor not in tabla or not pd.api.types.is_numeric_dtype(tabla[valor]):
        return
    datos = tabla.nlargest(top,valor).copy()
    datos[grupo] = datos[grupo].astype('string').fillna('Sin informacion')
    grafico(px.bar(datos.sort_values(valor),x=valor,y=grupo,orientation='h',title=titulo,hover_data=[c for c in ['Facturas','Ticket promedio'] if c in datos]))


def mapa_calor(df,filas,columnas,valor):
    if all(c in df for c in [filas,columnas,valor]) and numerica(df,valor):
        principales = df.groupby(filas,dropna=False)[valor].sum().nlargest(20).index
        tabla = df[df[filas].isin(principales)].pivot_table(index=filas,columns=columnas,values=valor,aggfunc='sum',observed=True)
        if not tabla.empty:
            grafico(go.Figure(go.Heatmap(z=tabla.to_numpy(),x=tabla.columns.astype(str),y=tabla.index.astype(str),colorscale='Blues')).update_layout(title=f'{valor}: {filas} y {columnas}'))


def calcular_pareto(tabla,grupo,valor):
    datos = tabla[[grupo,valor]].dropna().sort_values(valor,ascending=False).copy()
    datos = datos[datos[valor]>=0]
    suma = datos[valor].sum()
    datos['Participacion'] = datos[valor]/suma*100 if suma else np.nan
    datos['Acumulado'] = datos['Participacion'].cumsum()
    return datos


def mostrar_pareto(tabla,grupo,valor):
    if valor not in tabla:
        return
    p = calcular_pareto(tabla,grupo,valor)
    if p.empty or p['Acumulado'].isna().all():
        return
    n80 = int(p['Acumulado'].lt(80).sum()+1)
    st.caption(f'Pareto de {valor}: {n80} categorias alcanzan al menos el 80 %. Top 3: {p.Participacion.head(3).sum():.1f} %. Se excluyen montos negativos del Pareto.')
    fig = make_subplots(specs=[[{'secondary_y':True}]])
    fig.add_trace(go.Bar(x=p[grupo].astype(str),y=p[valor],name=valor),secondary_y=False)
    fig.add_trace(go.Scatter(x=p[grupo].astype(str),y=p['Acumulado'],name='% acumulado'),secondary_y=True)
    fig.add_hline(y=80,line_dash='dash',secondary_y=True)
    fig.update_yaxes(title_text='Participacion acumulada (%)',range=[0,105],secondary_y=True)
    fig.update_layout(title='Pareto: referencia analitica del 80 %, no regla obligatoria')
    grafico(fig)


def calcular_estadisticos(serie):
    x = pd.to_numeric(serie,errors='coerce').replace([np.inf,-np.inf],np.nan).dropna()
    if x.empty:
        return pd.DataFrame()
    q1,q3 = x.quantile([.25,.75])
    moda = x.mode()
    valores = {'Conteo':len(x),'Media':x.mean(),'Mediana':x.median(),'Moda':moda.iloc[0] if len(moda) else np.nan,
               'Desviacion estandar':x.std(),'Varianza':x.var(),'Minimo':x.min(),'Maximo':x.max(),
               'Rango':x.max()-x.min(),'Cuartil 1':q1,'Cuartil 3':q3,'Rango intercuartilico':q3-q1,
               'Percentil 75':x.quantile(.75),'Percentil 90':x.quantile(.90),'Percentil 95':x.quantile(.95),
               'Coeficiente de variacion (%)':x.std()/abs(x.mean())*100 if x.mean()!=0 else np.nan,
               'Asimetria':x.skew(),'Curtosis (exceso)':x.kurt()}
    return pd.DataFrame({'Estadistico':valores.keys(),'Valor':valores.values()})


def detectar_atipicos(df,col,metodo='IQR',factor=1.5,z=3.,percentiles=(.01,.99),grupo=None):
    """Limites globales o por categoria. No elimina observaciones."""
    if not numerica(df,col):
        return pd.Series(False,index=df.index),pd.DataFrame()
    if grupo and grupo in df:
        gb = df.groupby(grupo,dropna=False,observed=True)[col]
        if metodo == 'IQR':
            q1,q3 = gb.transform(lambda s:s.quantile(.25)),gb.transform(lambda s:s.quantile(.75))
            bajo,alto = q1-factor*(q3-q1),q3+factor*(q3-q1)
        elif metodo == 'Puntaje Z':
            media,desv = gb.transform('mean'),gb.transform('std')
            bajo,alto = media-z*desv,media+z*desv
        else:
            bajo,alto = gb.transform(lambda s:s.quantile(percentiles[0])),gb.transform(lambda s:s.quantile(percentiles[1]))
        limites = pd.DataFrame({grupo:df[grupo],'Inferior':bajo,'Superior':alto}).drop_duplicates()
    else:
        s = df[col].dropna()
        if s.empty:
            return pd.Series(False,index=df.index),pd.DataFrame()
        if metodo == 'IQR':
            q1,q3 = s.quantile([.25,.75]); bajo,alto = q1-factor*(q3-q1),q3+factor*(q3-q1)
        elif metodo == 'Puntaje Z':
            bajo,alto = s.mean()-z*s.std(),s.mean()+z*s.std()
        else:
            bajo,alto = s.quantile(percentiles[0]),s.quantile(percentiles[1])
        limites = pd.DataFrame({'Inferior':[bajo],'Superior':[alto]})
    return (df[col].lt(bajo)|df[col].gt(alto)).fillna(False),limites


def ejecutar_prueba_estadistica(df,variable,grupo,prueba,alpha=.05,otra=None):
    """Pruebas exploratorias, con guardas de tamano y supuestos."""
    detalle = []
    if prueba == 'Chi-cuadrado':
        tabla = pd.crosstab(df[grupo],df[otra])
        if min(tabla.shape)<2:
            raise ValueError('Se requieren dos categorias en cada variable.')
        estad,p,_,esperados = stats.chi2_contingency(tabla)
        if (esperados<5).any():
            detalle.append('Hay frecuencias esperadas menores a 5; la aproximacion chi-cuadrado puede no ser adecuada.')
        n = int(tabla.to_numpy().sum()); h0='Las variables categoricas son independientes.'; h1='Existe asociacion entre las variables.'
    elif prueba == 'Shapiro-Wilk':
        x = df[variable].dropna()
        if len(x)<3 or x.nunique()<2:
            raise ValueError('Se requieren al menos tres valores no constantes.')
        n = len(x)
        if n>5000:
            x = x.sample(5000,random_state=42)
            detalle.append(f'Muestra reproducible de 5000 observaciones, de {n} disponibles (semilla 42).')
        estad,p = stats.shapiro(x)
        h0='La distribucion es compatible con normalidad.'; h1='La distribucion difiere de la normal.'
        n = len(x)
    else:
        grupos = [s[variable].dropna().to_numpy() for _,s in df.groupby(grupo,observed=True,dropna=True)]
        if len(grupos)<2 or any(len(g)<3 for g in grupos):
            raise ValueError('Se requieren al menos dos grupos con tres observaciones cada uno.')
        n = [len(g) for g in grupos]
        if prueba in ['t de Student','Mann-Whitney U'] and len(grupos)!=2:
            raise ValueError('Esta prueba requiere exactamente dos grupos; filtre las categorias.')
        if prueba in ['t de Student','ANOVA']:
            normales = all(np.ptp(g)>0 and stats.shapiro(g if len(g)<=5000 else np.random.default_rng(42).choice(g,5000,replace=False)).pvalue>alpha for g in grupos)
            homogeneos = stats.levene(*grupos).pvalue>alpha
            if not normales or not homogeneos:
                raise ValueError('Los chequeos de normalidad u homogeneidad no son razonables. Use Mann-Whitney U o Kruskal-Wallis. La independencia debe verificarse externamente.')
            detalle.append('Chequeos de Shapiro y Levene no rechazados; esto no demuestra los supuestos. La independencia no puede comprobarse con estas columnas.')
        if prueba == 't de Student':
            estad,p = stats.ttest_ind(*grupos,equal_var=True)
        elif prueba == 'Mann-Whitney U':
            estad,p = stats.mannwhitneyu(*grupos,alternative='two-sided')
        elif prueba == 'ANOVA':
            estad,p = stats.f_oneway(*grupos)
        else:
            estad,p = stats.kruskal(*grupos)
        h0 = 'Las medias de los grupos son iguales.' if prueba in ['t de Student','ANOVA'] else 'Las distribuciones de los grupos son iguales.'
        h1 = 'Al menos un grupo difiere.'
    if not np.isfinite(p):
        raise ValueError('La prueba no produjo un p-valor valido; revise variabilidad y muestras.')
    return {'Prueba':prueba,'H0':h0,'H1':h1,'Estadistico':float(estad),'p-valor':float(p),
            'Alpha':alpha,'Tamano de muestras':str(n),'Decision':'Rechazar H0' if p<alpha else 'No rechazar H0',
            'Advertencias':' '.join(detalle)}


def generar_insights(df,token):
    resultados = []
    def agregar(hallazgo,valor,contexto,interpretacion,accion):
        resultados.append({'Hallazgo calculado':hallazgo,'Valor observado':str(valor),'Contexto':contexto,
                           'Interpretacion':interpretacion,'Revision sugerida':accion})
    cierre = construir_cierre_mensual(df,'Mensual',token)
    if not cierre.empty:
        validos = cierre.dropna(subset=['Valor_Neto'])
        if not validos.empty:
            r = validos.loc[validos.Valor_Neto.idxmax()]
            agregar('Periodo de mayor facturacion',f'{r.Periodo}: {r.Valor_Neto:,.2f}','Conjunto filtrado','Concentra el mayor monto mensual observado.','Revisar mezcla de servicios y volumen; no se establece la causa.')
        if len(cierre)>1 and pd.notna(cierre.iloc[-1]['Crecimiento porcentual']):
            v = cierre.iloc[-1]['Crecimiento porcentual']
            agregar('Variacion del ultimo periodo',f'{v:.1f} %','Comparacion con periodo inmediatamente anterior','Aumento observado.' if v>0 else 'Disminucion o estabilidad observada.','Comprobar cobertura y completitud de ambos periodos.')
        caidas = cierre[cierre['Crecimiento porcentual']<0]
        if not caidas.empty:
            r = caidas.loc[caidas['Crecimiento porcentual'].idxmin()]
            agregar('Mayor caida mensual',f'{r.Periodo}: {r["Crecimiento porcentual"]:.1f} %','Periodos consecutivos disponibles','Cambio descriptivo, no causal.','Revisar si el periodo esta completo.')
    for grupo in ['Entidad_Pagadora','Servicio','Sede','Canal_Radicacion']:
        if grupo not in df:
            continue
        tabla = agrupar(df,grupo,token)
        for col in ['Valor_Neto','Saldo_Cartera','Porcentaje Glosa','Dias_Emision_Radicacion']:
            if col in tabla and tabla[col].notna().any():
                r = tabla.loc[tabla[col].idxmax()]
                agregar(f'Mayor {col} por {grupo}',f'{r[grupo]}: {r[col]:,.2f}','Agregacion del conjunto filtrado','Resultado descriptivo; la informacion no establece causas.','Revisar registros, contratos y soportes aplicables.')
        if grupo=='Entidad_Pagadora':
            if 'Porcentaje Recaudo' in tabla and tabla['Porcentaje Recaudo'].notna().any():
                r = tabla.loc[tabla['Porcentaje Recaudo'].idxmin()]
                agregar('Menor porcentaje de recaudo',f'{r[grupo]}: {r["Porcentaje Recaudo"]:.1f} %','Pagos sobre facturacion neta','Puede justificar una revision del saldo observado.','Verificar fechas, acuerdos y corte; no atribuir conducta al pagador.')
            for col in ['Valor_Neto','Saldo_Cartera']:
                if col in tabla:
                    p = calcular_pareto(tabla,grupo,col)
                    if not p.empty and p.Participacion.notna().any():
                        agregar(f'Concentracion de {col}',f'Top 3: {p.Participacion.head(3).sum():.1f} %','Pareto sobre montos no negativos','Dependencia descriptiva de pocos pagadores.','Revisar exposicion y diversificacion sin asumir riesgo contractual.')
    if numerica(df,'Glosa_Inicial'):
        ids = df.loc[df.Glosa_Inicial>0,'ID_Factura'].nunique()
        agregar('Facturas con glosa',f'{ratio(ids,facturas(df)):.1f} %','Identificadores unicos con algun registro glosado','Frecuencia, diferente de severidad monetaria.','Revisar motivos y soportes de las facturas glosadas.')
    if 'Soporte_Completo' in df:
        agregar('Soportes incompletos',f'{soporte_incompleto(df.Soporte_Completo).mean()*100:.1f} %','Porcentaje de registros; valores no reconocidos no se clasifican como incompletos.','Asociacion no implica causalidad.','Revisar codificacion y documentacion.')
    if numerica(df,'Valor_Neto'):
        mask,_ = detectar_atipicos(df,'Valor_Neto')
        agregar('Valores netos atipicos (IQR)',int(mask.sum()),'Registros, multiplicador 1.5','Un valor atipico no es necesariamente un error.','Verificar complejidad y legitimidad del servicio.')
    return pd.DataFrame(resultados)


@cache_memoria
def convertir_dataframe_excel(df,token):
    buffer = BytesIO()
    copia = df.copy()
    for col in copia.select_dtypes(include=['category']).columns:
        copia[col] = copia[col].astype('string')
    # Neutralizar formulas al exportar cadenas de archivos no confiables.
    for col in copia.select_dtypes(include=['object','string']).columns:
        s = copia[col].astype('string')
        peligro = s.str.match(r'^\s*[=+@-]',na=False)
        copia.loc[peligro,col] = "'"+s.loc[peligro]
    if len(copia)>1048575:
        raise ValueError('El conjunto supera el limite de filas de Excel. Descargue CSV.')
    with pd.ExcelWriter(buffer,engine='openpyxl') as writer:
        copia.to_excel(writer,index=False,sheet_name='Analisis')
        hoja = writer.sheets['Analisis']
        hoja.freeze_panes='A2'; hoja.auto_filter.ref=hoja.dimensions
        for celda in hoja[1]:
            from openpyxl.styles import Font, PatternFill
            celda.font=Font(bold=True,color='FFFFFF'); celda.fill=PatternFill('solid',fgColor='126782')
    return buffer.getvalue()


@cache_memoria
def convertir_dataframe_csv(df,token):
    copia = df.copy()
    for col in copia.select_dtypes(include=['object','string']).columns:
        s = copia[col].astype('string'); peligro = s.str.match(r'^\s*[=+@-]',na=False)
        copia.loc[peligro,col] = "'"+s.loc[peligro]
    return copia.to_csv(index=False).encode('utf-8-sig')


def descargar(df,nombre,token):
    if df.empty:
        st.caption(f'{nombre}: sin registros disponibles.')
        return
    a,b = st.columns(2)
    try:
        a.download_button(f'{nombre} - CSV',convertir_dataframe_csv(df,token),nombre+'.csv','text/csv',key='csv_'+nombre)
        # Preparar Excel solamente a peticion para no duplicar trabajo pesado.
        if b.checkbox('Preparar Excel',key='preparar_'+nombre):
            b.download_button(f'{nombre} - Excel',convertir_dataframe_excel(df,token),nombre+'.xlsx',
                              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',key='xlsx_'+nombre)
    except Exception as exc:
        st.warning(f'No fue posible preparar la descarga: {exc}')


def aplicar_filtros(df):
    """Todos los modulos reciben este mismo DataFrame filtrado."""
    salida = df
    with st.sidebar.expander('Filtros globales',expanded=True):
        a,b = st.columns(2)
        if a.button('Seleccionar todos') or b.button('Limpiar filtros'):
            for k in list(st.session_state):
                if k.startswith('filtro_'):
                    del st.session_state[k]
            st.rerun()
        if 'Periodo_Cierre' in df:
            periodos = pd.to_datetime(df.Periodo_Cierre,format='%Y-%m',errors='coerce')
            if periodos.notna().any():
                anios = sorted(periodos.dt.year.dropna().astype(int).unique())
                elegidos = st.multiselect('Ano',anios,default=anios,key='filtro_ano')
                meses = sorted(periodos.dt.month.dropna().astype(int).unique())
                elegidos_m = st.multiselect('Mes',meses,default=meses,format_func=lambda x:MESES[x-1],key='filtro_mes')
                salida = salida.loc[periodos.dt.year.isin(elegidos)&periodos.dt.month.isin(elegidos_m)]
        if 'Fecha_Radicacion' in df and pd.api.types.is_datetime64_any_dtype(df.Fecha_Radicacion):
            f = df.Fecha_Radicacion.dropna()
            if not f.empty:
                rango = st.date_input('Rango de fechas de radicacion',(f.min().date(),f.max().date()),key='filtro_fechas')
                if len(rango)==2:
                    salida = salida.loc[salida.Fecha_Radicacion.between(pd.Timestamp(rango[0]),pd.Timestamp(rango[1])+pd.Timedelta(days=1)-pd.Timedelta(nanoseconds=1))]
        for col in CATEGORIAS:
            if col in df:
                opciones = sorted(df[col].dropna().astype(str).unique())
                sel = st.multiselect(col.replace('_',' '),opciones,default=opciones,key='filtro_'+col)
                salida = salida.loc[salida[col].astype(str).isin(sel)]
        for col,label in [('Glosa_Inicial','Solo facturas con glosa'),('Saldo_Cartera','Solo facturas con saldo')]:
            if numerica(df,col) and st.checkbox(label,key='filtro_'+col):
                salida = salida.loc[salida[col]>0]
        if numerica(df,'Valor_Neto') and df.Valor_Neto.notna().any():
            bajo,alto = float(df.Valor_Neto.min()),float(df.Valor_Neto.max())
            if bajo<alto:
                rango = st.slider('Rango de valor neto',bajo,alto,(bajo,alto),key='filtro_neto')
                salida = salida.loc[salida.Valor_Neto.between(*rango)]
        st.caption(f'{len(salida):,} de {len(df):,} registros ({ratio(len(salida),len(df)):.1f} %)')
    return salida.copy()


def mostrar_inicio(df,archivo,info):
    st.subheader('Archivo y alcance')
    st.write(f'Archivo: {archivo.name} | Tamano: {archivo.size/1024:.1f} KB | {info}')
    st.write(f'{len(df):,} registros y {len(df.columns)} columnas. Facturas unicas: {facturas(df):,}.')
    for col in ['Entidad_Pagadora','Sede','Servicio']:
        if col in df:
            st.caption(f'{col}: {df[col].nunique()} categorias')
    if 'Periodo_Cierre' in df and df.Periodo_Cierre.notna().any():
        st.caption(f'Periodo identificado: {df.Periodo_Cierre.min()} a {df.Periodo_Cierre.max()}')
    aleatoria = st.checkbox('Visualizar muestra aleatoria')
    mostrar_tabla(df.sample(min(20,len(df)),random_state=42) if aleatoria else df.head(20))
    st.info('Si una factura ocupa varias filas, los montos deben corresponder a cada linea, no repetirse como total de factura. La aplicacion suma montos por registro y cuenta identificadores unicos.')


def mostrar_calidad(original,limpio,conversiones,exclusiones,diagnostico):
    nulos,reglas,componentes,_ = diagnostico
    st.subheader('Diagnostico de calidad de datos originales mapeados')
    st.metric('Puntaje orientativo',f'{np.mean(list(componentes.values())):.1f} / 100')
    st.caption('El puntaje es un indicador exploratorio de calidad y no reemplaza una auditoria de datos. Promedio simple de completitud, unicidad, validez y consistencia; reglas no evaluables no acreditan cumplimiento.')
    mostrar_tabla(pd.DataFrame({'Componente':componentes.keys(),'Puntaje':componentes.values()}))
    st.write(f'Originales: {len(original):,} | Posteriores: {len(limpio):,} | Excluidos: {len(original)-len(limpio):,}')
    st.write('Columnas encontradas:',list(original.columns))
    st.write('Columnas esperadas ausentes:',[c for c in CAMPOS if c not in original])
    for titulo,tabla in [('Reglas de validacion',reglas),('Nulos por columna',nulos),('Conversiones aplicadas',conversiones),('Exclusiones',exclusiones)]:
        st.markdown('#### '+titulo); mostrar_tabla(tabla)
    st.caption('Un ID repetido puede corresponder a varias lineas legitimas. Una glosa inicial no siempre es una deduccion definitiva: confirme la identidad de saldo segun su modelo contable.')


def mostrar_resumen(df,completo,token):
    k = calcular_kpis(df)
    anterior = None
    if 'Periodo_Cierre' in df and df.Periodo_Cierre.notna().any():
        ultimo = pd.Period(df.Periodo_Cierre.max(),freq='M')
        actual = df[df.Periodo_Cierre==str(ultimo)]
        previo = df[df.Periodo_Cierre==str(ultimo-1)]
        if not actual.empty and not previo.empty:
            anterior = (calcular_kpis(actual),calcular_kpis(previo))
            st.caption(f'Deltas: {ultimo} frente a {ultimo-1}, ambos dentro del filtro. Las tarjetas muestran el total filtrado.')
    indicadores = [('Valor_Bruto','Facturacion bruta'),('Valor_Neto','Facturacion neta'),('Facturas','Facturas unicas'),
                   ('Cantidad','Cantidad de servicios'),('Valor_Pagado','Valor pagado'),('Saldo_Cartera','Saldo de cartera'),
                   ('Glosa_Inicial','Valor glosado'),('Porcentaje de recaudo','Porcentaje de recaudo'),
                   ('Porcentaje de glosa','Porcentaje de glosa'),('Ticket promedio','Ticket promedio'),('Tiempo promedio','Tiempo promedio de radicacion')]
    for i in range(0,len(indicadores),4):
        cols = st.columns(4)
        for cont,(clave,label) in zip(cols,indicadores[i:i+4]):
            valor = k[clave]
            texto = formatear_porcentaje(valor) if clave.startswith('Porcentaje') else (f'{valor:,.1f}' if clave in ['Facturas','Cantidad','Tiempo promedio'] and pd.notna(valor) else formatear_moneda(valor))
            delta = None
            if anterior:
                a,b = anterior[0][clave],anterior[1][clave]
                if pd.notna(a) and pd.notna(b) and b!=0:
                    delta=f'{(a/b-1)*100:+.1f} % ultimo mes'
            cont.metric(label,texto,delta)
    st.caption(f'Comparacion de alcance: neto completo {formatear_moneda(total(completo,"Valor_Neto"))}; neto filtrado {formatear_moneda(k["Valor_Neto"])}.')
    cierre = construir_cierre_mensual(df,'Mensual',token)
    if not cierre.empty:
        columnas = [c for c in ['Valor_Neto','Valor_Pagado','Saldo_Cartera','Glosa_Inicial'] if c in cierre]
        grafico(px.line(cierre,x='Periodo',y=columnas,markers=True,title='Facturacion, pagos, cartera y glosas por periodo de cierre'))
    for grupo in ['Estado_Factura','Entidad_Pagadora','Linea_Servicio']:
        if grupo in df:
            barras(agrupar(df,grupo,token),grupo,'Valor_Neto','Facturacion por '+grupo)


def mostrar_cierre(df,token):
    frecuencia = st.selectbox('Periodicidad',['Mensual','Trimestral','Anual'])
    tabla = construir_cierre_mensual(df,frecuencia,token)
    if tabla.empty:
        st.info('No hay periodos validos para construir el cierre.'); return
    mostrar_tabla(tabla)
    columnas = [c for c in ['Valor_Neto','Valor_Pagado','Saldo_Cartera','Promedio movil 3 periodos'] if c in tabla]
    grafico(px.line(tabla,x='Periodo',y=columnas,markers=True,title='Cierre y promedio movil'))
    fig = make_subplots(specs=[[{'secondary_y':True}]])
    fig.add_trace(go.Bar(x=tabla.Periodo,y=tabla.Valor_Neto,name='Valor neto'))
    fig.add_trace(go.Scatter(x=tabla.Periodo,y=tabla['Crecimiento porcentual'],name='Crecimiento (%)'),secondary_y=True)
    grafico(fig.update_layout(title='Facturacion y crecimiento del periodo'))
    for col in ['Valor_Neto','Valor_Pagado','Saldo_Cartera','Glosa_Inicial']:
        if col in tabla and tabla[col].notna().any():
            r = tabla.loc[tabla[col].idxmax()]
            st.write(f'Mayor {col}: {r.Periodo}, {formatear_moneda(r[col])}.')
    if not tabla.Valor_Neto.notna().any():
        st.warning('No hay valores netos validos en los periodos.'); return
    r = tabla.loc[tabla.Valor_Neto.idxmin()]; st.write(f'Menor facturacion: {r.Periodo}, {formatear_moneda(r.Valor_Neto)}.')
    if frecuencia=='Mensual':
        comp = tabla.assign(Ano=tabla.Periodo.str[:4],Mes=tabla.Periodo.str[5:])
        grafico(px.line(comp,x='Mes',y='Valor_Neto',color='Ano',markers=True,title='Comparacion entre anos'))
    st.caption('Pagos y saldos se agrupan por periodo de cierre de la factura, no por fecha efectiva del movimiento. Periodos sin registros se muestran como faltantes, no como cero. Promedio movil: tres periodos consecutivos.')


def mostrar_dimension(df,grupos,token):
    disponibles = [c for c in grupos if c in df]
    if not disponibles:
        st.info('No se encontro la columna necesaria para este analisis.'); return
    grupo = st.selectbox('Agrupacion',disponibles)
    tabla = agrupar(df,grupo,token)
    top_texto = st.selectbox('Categorias visibles',['Top 5','Top 10','Top 15','Todos'],index=2)
    top = len(tabla) if top_texto=='Todos' else int(top_texto.split()[1])
    mostrar_tabla(tabla.sort_values('Valor_Neto',ascending=False))
    valor = st.selectbox('Indicador del ranking',[c for c in ['Valor_Neto','Saldo_Cartera','Glosa_Inicial','Valor_Pagado','Porcentaje Recaudo'] if c in tabla])
    barras(tabla,grupo,valor,'Ranking por '+valor,top)
    if st.checkbox('Mostrar Pareto',value=True):
        mostrar_pareto(tabla,grupo,valor if valor in DINERO else 'Valor_Neto')
    positivos = tabla.nlargest(top,'Valor_Neto').query('Valor_Neto > 0').copy()
    positivos[grupo] = positivos[grupo].astype(str)
    if not positivos.empty:
        grafico(px.treemap(positivos,path=[grupo],values='Valor_Neto',title='Participacion de facturacion'))
    if 'Porcentaje Recaudo' in tabla:
        aux = tabla.copy(); aux[grupo]=aux[grupo].astype(str)
        tamano = 'Saldo_Cartera' if 'Saldo_Cartera' in aux else 'Valor_Neto'
        aux['_tamano'] = aux[tamano].clip(lower=0).fillna(0)
        if aux['_tamano'].sum()>0:
            grafico(px.scatter(aux,x='Valor_Neto',y='Porcentaje Recaudo',size='_tamano',color='Porcentaje Glosa' if 'Porcentaje Glosa' in aux else None,
                              hover_name=grupo,hover_data=['Facturas','Ticket promedio'],title='Facturacion y recaudo; tamano = '+tamano))
    categorias = tabla.nlargest(top,'Valor_Neto')[grupo].dropna().astype(str).tolist()
    muestra = df[df[grupo].astype(str).isin(categorias)]
    if len(muestra)>10000:
        muestra = muestra.sample(10000,random_state=42); st.caption('Boxplot: muestra reproducible de 10000 registros.')
    grafico(px.box(muestra,x=grupo,y='Valor_Neto',title='Distribucion de valor neto',points=False))
    mapa_calor(df,grupo,'Periodo_Cierre','Valor_Neto')
    if grupo in ['Servicio','Linea_Servicio']:
        mapa_calor(df,grupo,'Entidad_Pagadora','Valor_Neto')
    if grupo in ['Sede','Regimen'] and 'Estado_Factura' in df:
        cruz = pd.crosstab(df[grupo],df.Estado_Factura,normalize='index')*100
        grafico(px.bar(cruz.reset_index().melt(id_vars=grupo),x=grupo,y='value',color='Estado_Factura',title='Estados por categoria (100 %)',barmode='stack'))
    if categorias:
        seleccion = st.selectbox('Detalle de categoria',categorias)
        detalle = df[df[grupo].astype(str)==seleccion]
        cierre = construir_cierre_mensual(detalle,'Mensual',token)
        if not cierre.empty:
            grafico(px.line(cierre,x='Periodo',y='Valor_Neto',markers=True,title='Evolucion de '+seleccion))
        mostrar_tabla(detalle,100)


def mostrar_estados(df,token):
    if 'Estado_Factura' not in df:
        st.info('No se encontro Estado_Factura.'); return
    tabla = agrupar(df,'Estado_Factura',token)
    tabla['Porcentaje facturas'] = tabla.Facturas/facturas(df)*100 if facturas(df) else np.nan
    mostrar_tabla(tabla)
    barras(tabla,'Estado_Factura','Facturas','Facturas por estado')
    st.caption('Una factura con registros en distintos estados puede contarse en mas de un grupo; los porcentajes por estado no son una particion si el ID se repite entre estados.')
    for grupo in ['Periodo_Cierre','Entidad_Pagadora','Sede']:
        if grupo in df:
            tabla = df.groupby([grupo,'Estado_Factura'],observed=True,dropna=False).agg(Facturas=('ID_Factura','nunique')).reset_index()
            grafico(px.bar(tabla,x=grupo,y='Facturas',color='Estado_Factura',title='Estados por '+grupo,barmode='stack'))
            tabla['Porcentaje'] = tabla.Facturas/tabla.groupby(grupo,dropna=False).Facturas.transform('sum').replace(0,np.nan)*100
            grafico(px.bar(tabla,x=grupo,y='Porcentaje',color='Estado_Factura',title='Distribucion de estados al 100 %',barmode='stack'))
            mostrar_tabla(pd.crosstab(df[grupo],df.Estado_Factura).reset_index())
            mapa_calor(df,grupo,'Estado_Factura','Valor_Neto')


def mostrar_glosas(df,token):
    if not numerica(df,'Glosa_Inicial'):
        st.info('No hay Glosa_Inicial numerica.'); return
    g = df[df.Glosa_Inicial>0]
    k = calcular_kpis(df)
    st.metric('Glosa total',formatear_moneda(k['Glosa_Inicial']))
    st.write(f'Facturas unicas con glosa: {facturas(g)}; frecuencia: {formatear_porcentaje(ratio(facturas(g),facturas(df)))}; severidad: {formatear_porcentaje(k["Porcentaje de glosa"])}.')
    if g.empty:
        st.info('No hay registros con glosa positiva.'); return
    st.write(f'Promedio por registro glosado: {formatear_moneda(g.Glosa_Inicial.mean())}; mediana: {formatear_moneda(g.Glosa_Inicial.median())}; maximo: {formatear_moneda(g.Glosa_Inicial.max())}.')
    for grupo in ['Entidad_Pagadora','Servicio','Sede','Periodo_Cierre','Soporte_Completo']:
        if grupo in df:
            barras(agrupar(df,grupo,token),grupo,'Glosa_Inicial','Glosa por '+grupo)
    grafico(px.histogram(g,x='Glosa_Inicial',nbins=40,title='Distribucion de glosas positivas'))
    grafico(px.box(g,y='Glosa_Inicial',points=False,title='Glosas: boxplot'))
    mostrar_tabla(g.nlargest(100,'Glosa_Inicial'))
    st.caption('Una asociacion entre soportes y glosas no demuestra causalidad.')


def mostrar_cartera(df,token):
    k = calcular_kpis(df)
    for col in ['Valor_Neto','Valor_Pagado','Saldo_Cartera']:
        st.metric(col.replace('_',' '),formatear_moneda(k[col]))
    st.write(f'Recaudo: {formatear_porcentaje(k["Porcentaje de recaudo"])}; cartera: {formatear_porcentaje(k["Porcentaje de cartera"])}.')
    if numerica(df,'Valor_Pagado'):
        pagadas = df[df.Valor_Pagado.ge(df.Valor_Neto)&df.Valor_Neto.gt(0)]
        parciales = df[df.Valor_Pagado.gt(0)&df.Valor_Pagado.lt(df.Valor_Neto)]
        st.caption(f'Facturas con algun registro totalmente pagado: {facturas(pagadas)}; con algun registro parcialmente pagado: {facturas(parciales)}. Clasificacion por registro; un ID puede pertenecer a ambos conjuntos.')
        grafico(px.scatter(df,x='Valor_Neto',y='Valor_Pagado',title='Relacion facturacion y recaudo'))
    if not numerica(df,'Saldo_Cartera'):
        st.info('No hay saldo de cartera numerico.'); return
    saldo = df[df.Saldo_Cartera>0]
    st.write(f'Facturas con saldo: {facturas(saldo)}; saldo promedio por registro: {formatear_moneda(saldo.Saldo_Cartera.mean())}; mediana: {formatear_moneda(saldo.Saldo_Cartera.median())}.')
    for grupo in ['Entidad_Pagadora','Servicio','Sede','Periodo_Cierre']:
        if grupo in df:
            barras(agrupar(df,grupo,token),grupo,'Saldo_Cartera','Cartera por '+grupo)
    mostrar_tabla(saldo.nlargest(100,'Saldo_Cartera'))
    if 'Fecha_Radicacion' in df and pd.api.types.is_datetime64_any_dtype(df.Fecha_Radicacion):
        corte = st.date_input('Fecha de corte para antiguedad aproximada',date.today())
        limites = st.text_input('Limites de antiguedad (dias)', '30,60,90,180')
        try:
            puntos = [int(x.strip()) for x in limites.split(',')]
            if not puntos or min(puntos)<=0 or puntos!=sorted(set(puntos)):
                raise ValueError()
            antiguedad = (pd.Timestamp(corte)-saldo.Fecha_Radicacion).dt.days
            futuro = int(antiguedad.lt(0).sum())
            if futuro:
                st.warning(f'{futuro} registros tienen radicacion posterior al corte y se excluyen de antiguedad.')
            labels = [f'0 a {puntos[0]}']+[f'{a+1} a {b}' for a,b in zip(puntos,puntos[1:])]+[f'Mas de {puntos[-1]}']
            aux = saldo.loc[antiguedad.ge(0)].copy()
            aux['Antiguedad'] = pd.cut(antiguedad.loc[aux.index],[-.1]+puntos+[np.inf],labels=labels)
            tabla = aux.groupby('Antiguedad',observed=False).Saldo_Cartera.sum().reset_index()
            grafico(px.bar(tabla,x='Antiguedad',y='Saldo_Cartera',title='Antiguedad aproximada de saldo observado'))
        except ValueError:
            st.warning('Use limites enteros positivos, crecientes y separados por comas.')
        st.caption('Rangos analiticos configurables, no politica contable universal. El archivo no reconstruye saldos historicos al corte: clasifica el saldo observado por edad desde radicacion.')


def mostrar_tiempos(df,token):
    col='Dias_Emision_Radicacion'
    if not numerica(df,col):
        st.info('No se dispone de dias de radicacion numericos ni de fechas convertidas suficientes.'); return
    mostrar_tabla(calcular_estadisticos(df[col]))
    grafico(px.histogram(df,x=col,nbins=40,title='Dias entre emision y radicacion'))
    grafico(px.box(df,y=col,points=False,title='Distribucion de tiempos'))
    if 'Rango_Tiempo_Radicacion' in df:
        tabla=df.groupby('Rango_Tiempo_Radicacion',observed=False).size().reset_index(name='Registros')
        grafico(px.bar(tabla,x='Rango_Tiempo_Radicacion',y='Registros',title='Rangos analiticos configurables'))
    for grupo in ['Entidad_Pagadora','Sede','Canal_Radicacion','Servicio','Estado_Factura','Soporte_Completo','Periodo_Cierre']:
        if grupo in df:
            barras(agrupar(df,grupo,token),grupo,col,'Tiempo promedio por '+grupo)
    mostrar_tabla(df.nlargest(100,col))


def mostrar_estadistica(df):
    numericas = df.select_dtypes(include='number').columns.tolist()
    if not numericas:
        st.info('No hay variables numericas convertidas.'); return
    modo = st.radio('Modulo',['Descriptivos','Correlaciones','Pruebas estadisticas'],horizontal=True)
    if modo=='Descriptivos':
        variable = st.selectbox('Variable numerica',numericas)
        x=df[variable].dropna()
        if x.empty:
            st.info('La variable no contiene datos numericos validos.'); return
        mostrar_tabla(calcular_estadisticos(x))
        grafico(px.histogram(df,x=variable,nbins=40,title='Distribucion observada'))
        grafico(px.box(df,y=variable,points=False,title='Boxplot'))
        # Curva empirica acumulada, no una distribucion teorica asumida.
        ordenado=np.sort(x.to_numpy()); grafico(go.Figure(go.Scatter(x=ordenado,y=np.arange(1,len(x)+1)/len(x),mode='lines')).update_layout(title='Distribucion acumulada empirica',xaxis_title=variable,yaxis_title='Proporcion acumulada'))
        st.write(f'Los datos muestran media {x.mean():,.2f}, mediana {x.median():,.2f} y dispersion estandar {x.std():,.2f}. No se establece normalidad sin evaluarla.')
    elif modo=='Correlaciones':
        seleccion = st.multiselect('Variables',numericas,default=[c for c in NUMEROS if c in numericas])
        metodo=st.selectbox('Metodo',['Pearson','Spearman'])
        umbral=st.slider('Umbral de asociacion fuerte',.1,1.,.7,.05)
        if len(seleccion)<2:
            st.info('Seleccione al menos dos variables.'); return
        corr=df[seleccion].corr(method=metodo.lower(),min_periods=3)
        grafico(go.Figure(go.Heatmap(z=corr.to_numpy(),x=corr.columns,y=corr.index,zmin=-1,zmax=1,colorscale='RdBu')).update_layout(title='Matriz de correlacion'))
        mostrar_tabla(corr.reset_index())
        pares=[]
        for i,a in enumerate(seleccion):
            for b in seleccion[i+1:]:
                r=corr.loc[a,b]
                if pd.notna(r):
                    pares.append({'Variable A':a,'Variable B':b,'Correlacion':r,'Clasificacion':('Positiva fuerte' if r>0 else 'Negativa fuerte') if abs(r)>=umbral else 'Debil o moderada'})
        mostrar_tabla(pd.DataFrame(pares))
        a=st.selectbox('Eje X',seleccion); b=st.selectbox('Eje Y',seleccion,index=1)
        muestra=df[[a,b]].dropna() if a!=b else df[[a]].dropna()
        fig=px.scatter(df,x=a,y=b,title='Asociacion entre variables')
        if a!=b and len(muestra)>=3 and muestra[a].nunique()>1:
            pendiente,intercepto,*_=stats.linregress(muestra[a],muestra[b])
            xs=np.array([muestra[a].min(),muestra[a].max()]); fig.add_trace(go.Scatter(x=xs,y=pendiente*xs+intercepto,mode='lines',name='Tendencia lineal'))
        grafico(fig)
        st.warning('La correlacion no implica causalidad. Algunas relaciones pueden surgir directamente de las formulas contables entre las variables.')
    else:
        prueba=st.selectbox('Prueba',['Shapiro-Wilk','t de Student','Mann-Whitney U','ANOVA','Kruskal-Wallis','Chi-cuadrado'])
        alpha=st.number_input('Nivel de significancia',.001,.2,.05,.01)
        categorias=[c for c in CATEGORIAS if c in df]
        variable=st.selectbox('Variable numerica',numericas)
        grupo=st.selectbox('Agrupacion',categorias) if categorias else None
        otra=st.selectbox('Segunda variable categorica',categorias,index=min(1,len(categorias)-1)) if categorias and prueba=='Chi-cuadrado' else None
        st.caption('Analisis exploratorio: independencia y diseno de muestreo deben verificarse externamente. Multiples pruebas aumentan falsos positivos; no se aplica correccion automatica.')
        if st.button('Ejecutar prueba'):
            try:
                if prueba!='Shapiro-Wilk' and not grupo:
                    raise ValueError('No existen categorias para agrupar.')
                resultado=ejecutar_prueba_estadistica(df,variable,grupo,prueba,alpha,otra)
                mostrar_tabla(pd.DataFrame({'Elemento':resultado.keys(),'Resultado':[str(v) for v in resultado.values()]}))
                st.info('No rechazar H0 no demuestra que sea verdadera. Rechazar H0 tampoco establece causalidad ni relevancia practica.')
            except Exception as exc:
                st.warning(f'No fue posible ejecutar la prueba: {exc}')


def mostrar_atipicos(df):
    variables=df.select_dtypes(include='number').columns.tolist()
    if not variables:
        st.info('No hay variables numericas.'); return
    col=st.selectbox('Variable',variables,index=variables.index('Valor_Neto') if 'Valor_Neto' in variables else 0)
    metodo=st.selectbox('Metodo',['IQR','Puntaje Z','Percentiles'])
    grupo=st.selectbox('Alcance',['Global']+[c for c in ['Servicio','Linea_Servicio','Entidad_Pagadora'] if c in df])
    factor=st.number_input('Multiplicador IQR',.1,10.,1.5,.1)
    z=st.number_input('Limite de puntaje Z',.1,10.,3.,.1)
    percentiles=st.slider('Percentiles limites',0.,100.,(1.,99.))
    mask,limites=detectar_atipicos(df,col,metodo,factor,z,tuple(p/100 for p in percentiles),None if grupo=='Global' else grupo)
    st.session_state['atipicos_actuales']=df.loc[mask].copy()
    st.metric('Registros atipicos',int(mask.sum()),f'{ratio(mask.sum(),len(df)):.1f} % del filtro')
    mostrar_tabla(limites); mostrar_tabla(df.loc[mask])
    grafico(px.box(df,y=col,points=False,title='Boxplot'))
    grafico(px.histogram(df,x=col,nbins=40,title='Distribucion'))
    if 'Periodo_Cierre' in df:
        aux=df.assign(Atipico=mask.astype(str)); grafico(px.scatter(aux,x='Periodo_Cierre',y=col,color='Atipico',title='Atipicos por periodo'))
    st.caption('Un valor atipico no es necesariamente un error. Puede representar una factura legitima de mayor complejidad o valor. Puntaje Z es sensible a asimetria y dispersion nula; grupos pequenos producen limites inestables.')


def configurar_alertas():
    valores={}
    with st.sidebar.expander('Semaforos: parametros demostrativos'):
        st.caption('Editables; no son normas, contratos ni politicas oficiales. Verde/amarillo/rojo se acompanan de texto.')
        for nombre,base,alto in [('Glosa (%)',10.,20.),('Cartera (%)',30.,50.),('Recaudo (%)',70.,90.),
                                ('Dias de radicacion',7.,15.),('Soportes incompletos (%)',10.,20.),('Caida mensual (%)',5.,15.)]:
            a=st.number_input(nombre+' - limite 1',min_value=0.,value=base,key='alerta_1_'+nombre)
            b=st.number_input(nombre+' - limite 2',min_value=0.,value=alto,key='alerta_2_'+nombre)
            valores[nombre]=(min(a,b),max(a,b))
    return valores


def mostrar_insights(df,token,umbrales,diagnostico):
    insights=generar_insights(df,token)
    for _,r in insights.iterrows():
        with st.expander(r['Hallazgo calculado']+' | '+r['Valor observado']):
            st.write('Contexto: '+r.Contexto); st.write('Interpretacion: '+r.Interpretacion); st.write('Revision sugerida: '+r['Revision sugerida'])
    k=calcular_kpis(df)
    indicadores={'Glosa (%)':k['Porcentaje de glosa'],'Cartera (%)':k['Porcentaje de cartera'],
                 'Recaudo (%)':k['Porcentaje de recaudo'],'Dias de radicacion':k['Tiempo promedio'],
                 'Soportes incompletos (%)':soporte_incompleto(df.Soporte_Completo).mean()*100 if 'Soporte_Completo' in df else np.nan}
    cierre=construir_cierre_mensual(df,'Mensual',token)
    indicadores['Caida mensual (%)']=max(0,-cierre.iloc[-1]['Crecimiento porcentual']) if not cierre.empty and pd.notna(cierre.iloc[-1]['Crecimiento porcentual']) else np.nan
    for nombre,valor in indicadores.items():
        if pd.isna(valor):
            st.caption(nombre+': no disponible'); continue
        a,b=umbrales[nombre]
        if nombre=='Recaudo (%)':
            estado='Rojo' if valor<a else 'Amarillo' if valor<b else 'Verde'
        else:
            estado='Rojo' if valor>b else 'Amarillo' if valor>a else 'Verde'
        texto=f'{estado} | {nombre}: {valor:.1f}. Limites editables: {a:g}, {b:g}.'
        {'Rojo':st.error,'Amarillo':st.warning,'Verde':st.success}[estado](texto)
    reglas=diagnostico[1]
    if not reglas.empty:
        inconsistencias=reglas[reglas.Regla.str.contains('inconsistente')]
        st.caption(f'Inconsistencias financieras en originales: {int(inconsistencias.Registros.sum())} incidencias (una fila puede tener mas de una).')


def mostrar_explorador(df,sospechosos):
    modo=st.selectbox('Registros',['Todos','Con saldo','Con glosa','Soportes incompletos','Posibles inconsistencias','Atipicos IQR valor neto'])
    aux=df.copy()
    if modo=='Con saldo' and numerica(aux,'Saldo_Cartera'): aux=aux[aux.Saldo_Cartera>0]
    elif modo=='Con glosa' and numerica(aux,'Glosa_Inicial'): aux=aux[aux.Glosa_Inicial>0]
    elif modo=='Soportes incompletos' and 'Soporte_Completo' in aux: aux=aux[soporte_incompleto(aux.Soporte_Completo)]
    elif modo=='Posibles inconsistencias': aux=aux.loc[sospechosos.reindex(aux.index,fill_value=False)]
    elif modo=='Atipicos IQR valor neto': aux=aux.loc[detectar_atipicos(aux,'Valor_Neto')[0]]
    elif modo!='Todos':
        st.info('No se dispone de la columna necesaria para ese filtro.'); return
    texto=st.text_input('Buscar texto (literal, sin expresiones regulares)')
    if texto:
        mask=pd.Series(False,index=aux.index)
        for col in aux.columns:
            mask |= aux[col].astype('string').str.contains(texto,case=False,regex=False,na=False)
        aux=aux.loc[mask]
    cols=st.multiselect('Columnas visibles',aux.columns.tolist(),default=aux.columns.tolist())
    ordenar=st.selectbox('Ordenar por',aux.columns.tolist())
    descendente=st.checkbox('Orden descendente',value=True)
    limite=st.number_input('Filas a mostrar',10,10000,100,10)
    try:
        aux=aux.sort_values(ordenar,ascending=not descendente)
    except (TypeError,ValueError):
        st.warning('La columna contiene tipos mezclados; se mantiene el orden original.')
    mostrar_tabla(aux[cols],int(limite))


def mostrar_descargas(df,token,diagnostico):
    descargar(df,'facturacion_filtrada',token)
    descargar(construir_cierre_mensual(df,'Mensual',token),'cierre_mensual',token)
    for col,nombre in [('Entidad_Pagadora','analisis_entidades'),('Servicio','analisis_servicios'),('Sede','analisis_sedes')]:
        if col in df: descargar(agrupar(df,col,token),nombre,token)
    for col,nombre in [('Glosa_Inicial','registros_glosados'),('Saldo_Cartera','registros_cartera')]:
        if numerica(df,col): descargar(df[df[col]>0],nombre,token)
    mask,_=detectar_atipicos(df,'Valor_Neto')
    descargar(df.loc[mask],'registros_atipicos_IQR',token)
    if 'atipicos_actuales' in st.session_state:
        actuales=st.session_state['atipicos_actuales']
        actuales=actuales.loc[actuales.index.intersection(df.index)]
        descargar(actuales,'atipicos_configuracion_seleccionada',token)
    descargar(diagnostico[0],'informe_calidad_nulos',token)
    descargar(diagnostico[1],'informe_calidad_reglas',token)
    insights=generar_insights(df,token)
    descargar(insights,'insights_facturacion',token)
    st.download_button('Insights - texto',insights.to_string(index=False).encode('utf-8'),'insights_facturacion.txt','text/plain')
    st.caption('Calidad se refiere al archivo original mapeado; las demas descargas corresponden al conjunto filtrado. Las cadenas con posibles formulas se neutralizan en exportaciones.')


def mostrar_metodologia():
    st.markdown('''### Indicadores y alcance
- **Facturacion neta:** se suma Valor_Neto del archivo. Se valida contra Valor_Bruto - Descuento - Copago_Cuota; no se reemplaza automaticamente.
- **Porcentaje de glosa:** suma Glosa_Inicial / suma Valor_Neto x 100.
- **Porcentaje de recaudo:** suma Valor_Pagado / suma Valor_Neto x 100.
- **Porcentaje de cartera:** suma Saldo_Cartera / suma Valor_Neto x 100.
- **Ticket promedio:** suma Valor_Neto / numero de ID_Factura unicos no vacios.
- **Identidad de saldo exploratoria:** Valor_Neto - Glosa_Inicial - Valor_Pagado. Una glosa inicial no siempre es una deduccion definitiva; confirme el modelo del archivo.
- **Division por cero:** resultado no disponible, nunca infinito.
- **Conteos:** los montos y servicios se suman por fila. Las facturas se cuentan por identificador unico. Si los totales de factura se repiten en varias lineas, deben corregirse en la fuente; no se deduplican montos silenciosamente.
- **Periodos:** se usa Periodo_Cierre; si no existe, se deriva de Fecha_Radicacion. Pagos agrupados por cierre no equivalen a flujo de caja por Fecha_Pago. Periodos sin datos no se suponen cero. Deltas usan periodos consecutivos del filtro; no prueban que los meses esten completos.
- **IQR:** Q1 y Q3 son cuartiles; IQR = Q3 - Q1. Limites = Q1 - factor x IQR y Q3 + factor x IQR. Factor inicial 1.5, editable. Atipico no significa error.
- **Puntaje Z:** distancia a la media en desviaciones estandar; sensible a distribuciones asimetricas.
- **Correlacion:** asociacion lineal (Pearson) o de rangos (Spearman). No implica causalidad; identidades financieras pueden generar asociaciones mecanicas.
- **p-valor:** bajo H0 y los supuestos, probabilidad de un resultado al menos tan extremo como el observado. No es la probabilidad de que H0 sea verdadera. No rechazar H0 no la demuestra.
- **Pruebas:** normalidad y homogeneidad se revisan para pruebas parametricas; independencia, representatividad y diseno deben verificarse externamente. No hay ajuste automatico por multiples comparaciones.
- **Calidad:** promedio de completitud de celdas, unicidad de ID, proporcion de filas sin fallos de validez y sin inconsistencias evaluables. Reglas sin columnas no pueden evaluarse; el puntaje no certifica calidad.
- **Moneda:** COP, USD y EUR cambian solo la etiqueta. No se convierten tasas de cambio. El modo automatico considera miles un separador unico seguido de tres digitos y lo advierte; seleccione el formato de la fuente para resolver ambiguedad.
- **Antiguedad:** edad desde radicacion al corte, aplicada al saldo observado, no reconstruccion historica ni vencimiento contractual.
- **Soportes:** incompletos reconocidos: no, false, 0, incompleto, incompleta, n. Valores desconocidos deben revisarse.
- **Privacidad:** procesamiento en memoria del servidor donde se despliegue Streamlit; no significa procesamiento exclusivo en su computador. Sin escritura permanente ni APIs externas. La cache tiene TTL y no persiste en disco; en versiones recientes se limita a la sesion.
- **Limitaciones:** datos ausentes, negativos, periodos incompletos, estados cambiantes y facturas multilinea pueden afectar los resultados. Semaforos y rangos son parametros analiticos, no normas oficiales.
''')
    st.warning(PRIVACIDAD); st.info(AVISO)


def restaurar_originales():
    st.session_state['limpieza_aplicar'] = False


def reiniciar():
    """Reinicia widgets y cache de esta sesion sin borrar datos de otros usuarios."""
    token=st.session_state.get('token')
    for k in list(st.session_state):
        del st.session_state[k]
    st.session_state['uploader_version']=str(uuid.uuid4())


def main():
    st.set_page_config(page_title='Analitica de Facturacion en Salud',page_icon='\U0001f4ca',layout='wide',initial_sidebar_state='expanded')
    st.markdown('''<style>
    .stApp {background:#f5f8fb;} h1,h2,h3 {color:#123c55;}
    [data-testid="stMetric"] {background:white;padding:16px;border-radius:12px;border:1px solid #d6e3ec;}
    [data-testid="stSidebar"] {background:#eaf2f7;}
    </style>''',unsafe_allow_html=True)
    st.title(TITULO); st.caption('Facturacion, radicacion, glosas, recaudo y cartera')
    st.info(AVISO); st.warning(PRIVACIDAD)
    st.sidebar.button('Reiniciar aplicacion',on_click=reiniciar)
    if 'token' not in st.session_state: st.session_state['token']=str(uuid.uuid4())
    token=st.session_state['token']
    pagina=st.sidebar.radio('Navegacion',PAGINAS,key='navegacion')
    st.sidebar.selectbox('Etiqueta de moneda',['COP','USD','EUR','Sin simbolo'],key='moneda',help='Solo etiqueta visual; no convierte tasas de cambio.')
    archivo=st.sidebar.file_uploader('Cargue el archivo de facturacion que desea analizar',type=['xlsx','csv'],key='archivo_'+st.session_state.get('uploader_version','0'))
    if archivo is None:
        if pagina=='Metodologia': mostrar_metodologia(); return
        st.subheader('Comience cargando un archivo anonimizado')
        st.write('Analice facturacion, glosas, recaudo y cartera con indicadores y graficos interactivos. Se aceptan Excel .xlsx y CSV. Excel .xls no esta habilitado.')
        st.markdown('**Minimas:** ID_Factura, Valor_Neto y Fecha_Radicacion o Periodo_Cierre.\n\n**Recomendadas:** '+', '.join(RECOMENDADAS))
        st.write('1. Cargue el archivo en la barra lateral. 2. Seleccione hoja o formato CSV. 3. Revise y aplique el mapeo. 4. Revise calidad y limpieza. 5. Filtre y navegue por los analisis.')
        st.caption('No se incluye dataset ni se solicitan datos de pacientes. En una sesion nueva debe cargar nuevamente el archivo. El limite del cargador depende de la configuracion de Streamlit; esta app no impone otro limite oculto.')
        return
    contenido=archivo.getvalue(); extension=archivo.name.rsplit('.',1)[-1].lower()
    identidad=hashlib.sha256(contenido).hexdigest()+archivo.name
    if st.session_state.get('identidad')!=identidad:
        for k in list(st.session_state):
            if k.startswith(('filtro_','map_')) or k in ['mapeo','datos_originales','datos_limpios','atipicos_actuales','hoja_previa']:
                del st.session_state[k]
        st.session_state['identidad']=identidad
    hoja=None; separador='Automatico'; codificacion='Automatica'
    try:
        if extension=='xlsx':
            hojas=hojas_excel(contenido,token)
            indice=next((i for i,h in enumerate(hojas) if normalizar_nombre_columna(h)=='facturacion'),0)
            hoja=st.sidebar.selectbox('Hoja de Excel',hojas,index=indice,key='hoja_'+identidad[:12])
        else:
            separador=st.sidebar.selectbox('Separador CSV',['Automatico','Coma','Punto y coma','Tabulacion'])
            codificacion=st.sidebar.selectbox('Codificacion CSV',['Automatica','utf-8-sig','utf-16','cp1252','latin-1'])
        with st.spinner('Leyendo archivo en memoria...'):
            original,info=cargar_archivo(contenido,extension,hoja,separador,codificacion,token)
    except Exception as exc:
        st.error(f'No fue posible leer el archivo. {exc}'); return
    if original.empty:
        st.warning('La hoja seleccionada o el archivo no contiene datos.'); return
    if original.columns.duplicated().any():
        st.error('Existen encabezados duplicados; corrijalos antes de continuar.'); return
    fuente_id=identidad+str(hoja)+separador+codificacion
    if st.session_state.get('hoja_previa')!=fuente_id:
        for k in list(st.session_state):
            if k.startswith(('map_','filtro_')) or k in ['mapeo','atipicos_actuales']:
                del st.session_state[k]
        st.session_state['hoja_previa']=fuente_id
    st.session_state['datos_originales']=original
    sugeridos,candidatos=detectar_columnas(original)
    with st.expander('Configurar correspondencia de columnas',expanded='mapeo' not in st.session_state):
        st.caption('Equivalencias sugeridas, aun no aplicadas. Cada columna fuente debe asignarse una sola vez. Campos ambiguos requieren seleccion manual.')
        mostrar_tabla(pd.DataFrame([{'Campo esperado':c,'Sugerencia':sugeridos[c] or 'Columna no disponible','Candidatos':', '.join(candidatos[c])} for c in CAMPOS]))
        with st.form('formulario_mapeo'):
            seleccion={}; columnas=st.columns(3); opciones=['Columna no disponible']+original.columns.tolist()
            for i,campo in enumerate(CAMPOS):
                valor=st.session_state.get('mapeo',sugeridos).get(campo)
                seleccionado=columnas[i%3].selectbox(campo,opciones,index=opciones.index(valor) if valor in opciones else 0,key='map_'+campo)
                seleccion[campo]=None if seleccionado=='Columna no disponible' else seleccionado
            aplicar=st.form_submit_button('Aplicar correspondencia')
        if aplicar:
            fuentes=[v for v in seleccion.values() if v]
            colisiones=[c for c,v in seleccion.items() if v and c in original and c!=v and c not in fuentes]
            if len(fuentes)!=len(set(fuentes)) or colisiones:
                st.error('Asignacion duplicada o colision con encabezados existentes. Asigne cada fuente una sola vez y revise los campos: '+', '.join(colisiones))
            else:
                st.session_state['mapeo']=seleccion; st.rerun()
    if 'mapeo' not in st.session_state:
        st.info('Revise las sugerencias y pulse Aplicar correspondencia para continuar.'); mostrar_tabla(original.head(10)); return
    mapeo=st.session_state['mapeo']
    mapeado=original.rename(columns={v:c for c,v in mapeo.items() if v}).copy()
    if mapeado.columns.duplicated().any():
        st.error('El mapeo genera encabezados duplicados. Revise correspondencias.'); return
    with st.sidebar.expander('Limpieza controlada',expanded=False):
        st.checkbox('Aplicar transformaciones',value=True,key='limpieza_aplicar')
        st.checkbox('Eliminar filas completamente vacias',value=True,key='limpieza_vacias')
        st.checkbox('Eliminar duplicados exactos',value=False,key='limpieza_duplicados')
        st.checkbox('Convertir fechas',value=True,key='limpieza_fechas')
        st.checkbox('Convertir columnas numericas y monetarias',value=True,key='limpieza_moneda')
        st.checkbox('Categorias vacias: Sin informacion',value=True,key='limpieza_categorias')
        st.checkbox('Excluir Valor_Neto invalido',value=False,key='limpieza_neto')
        st.checkbox('Excluir fechas invalidas no vacias',value=False,key='limpieza_fecha_invalida')
        formato=st.selectbox('Formato numerico',['Automatico','Colombiano (1.234,56)','Internacional (1,234.56)'])
        orden=st.selectbox('Orden de fechas',['Dia primero','Mes primero'])
        tolerancia=st.number_input('Tolerancia identidades financieras',min_value=0.,value=1.,step=1.)
        st.button('Restaurar datos originales',on_click=restaurar_originales)
        st.caption('Sin conversion numerica se desactivan analisis financieros; los originales se conservan. Variables derivadas se construyen en la copia analitica.')
    with st.sidebar.expander('Rangos de radicacion (analiticos)'):
        a=st.number_input('Primer limite',0,3650,3)
        b=st.number_input('Segundo limite',a+1,3651,max(7,a+1))
        c=st.number_input('Tercer limite',b+1,3652,max(15,b+1))
    umbrales=configurar_alertas()
    opciones={k:st.session_state['limpieza_'+k] for k in ['aplicar','vacias','duplicados','fechas','moneda','categorias','neto','fecha_invalida']}
    with st.spinner('Validando y preparando analisis...'):
        limpio,conversiones,exclusiones=transformar(original,mapeo,opciones,formato,orden,(a,b,c),token)
        diagnostico=validar_datos(mapeado,tolerancia,formato,orden,token)
    st.session_state['datos_limpios']=limpio
    if not conversiones.empty:
        fallos=int(conversiones.Fallos.sum()); ambiguos=int(conversiones['Formatos ambiguos'].sum())
        st.caption(f'Conversiones: {len(conversiones)} columnas; {fallos} valores no convertibles. Detalle en Calidad de datos.')
        if ambiguos and formato=='Automatico':
            st.warning(f'{ambiguos} valores con formato numerico ambiguo se interpretaron como miles. Confirme el formato en Limpieza controlada antes de usar los totales.')
    faltantes=[c for c in ['ID_Factura','Valor_Neto'] if c not in limpio]
    if not any(c in limpio for c in ['Fecha_Radicacion','Periodo_Cierre']): faltantes.append('Fecha_Radicacion o Periodo_Cierre')
    if faltantes:
        st.error('Faltan columnas minimas: '+', '.join(faltantes))
        st.write('Encontradas:',list(limpio.columns)); st.info('Corrija los encabezados en el archivo o use el asistente de correspondencia. Los analisis financieros no se ejecutan.')
        mostrar_calidad(mapeado,limpio,conversiones,exclusiones,diagnostico); return
    ausentes=[c for c in RECOMENDADAS if c not in limpio]
    if ausentes: st.warning('Columnas recomendadas ausentes: '+', '.join(ausentes)+'. Se desactivan solo los analisis dependientes.')
    if pagina=='Calidad de datos':
        mostrar_calidad(mapeado,limpio,conversiones,exclusiones,diagnostico); return
    if pagina=='Metodologia': mostrar_metodologia(); return
    if not numerica(limpio,'Valor_Neto'):
        st.warning('Active conversion numerica en Limpieza controlada para ejecutar analisis financieros.'); mostrar_inicio(limpio,archivo,info); return
    filtrado=aplicar_filtros(limpio)
    if filtrado.empty:
        st.warning('Los filtros actuales no tienen registros. Use Limpiar filtros.'); return
    st.subheader(pagina)
    if pagina=='Inicio y carga de datos': mostrar_inicio(filtrado,archivo,info)
    elif pagina=='Resumen ejecutivo': mostrar_resumen(filtrado,limpio,token)
    elif pagina=='Cierre mensual': mostrar_cierre(filtrado,token)
    elif pagina=='Entidades pagadoras': mostrar_dimension(filtrado,['Entidad_Pagadora'],token)
    elif pagina=='Servicios y lineas': mostrar_dimension(filtrado,['Servicio','Linea_Servicio'],token)
    elif pagina=='Sedes y regimenes': mostrar_dimension(filtrado,['Sede','Regimen'],token)
    elif pagina=='Estados de facturacion': mostrar_estados(filtrado,token)
    elif pagina=='Glosas': mostrar_glosas(filtrado,token)
    elif pagina=='Recaudo y cartera': mostrar_cartera(filtrado,token)
    elif pagina=='Tiempos de radicacion': mostrar_tiempos(filtrado,token)
    elif pagina=='Analisis estadistico': mostrar_estadistica(filtrado)
    elif pagina=='Valores atipicos': mostrar_atipicos(filtrado)
    elif pagina=='Insights y alertas': mostrar_insights(filtrado,token,umbrales,diagnostico)
    elif pagina=='Explorador de datos': mostrar_explorador(filtrado,diagnostico[3])
    elif pagina=='Descargas': mostrar_descargas(filtrado,token,diagnostico)


if __name__=='__main__':
    try:
        main()
    except Exception:
        st.error('No fue posible completar este modulo. Revise columnas, tipos de datos y filtros, o reinicie la aplicacion. Los datos originales no se modificaron.')
