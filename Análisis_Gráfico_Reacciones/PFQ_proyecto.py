# -*- coding: utf-8 -*-
"""
====================================================================
 PROPIEDADES FISICOQUÍMICAS (PFQ) — ANÁLISIS AUTOMÁTICO
====================================================================
 Lee "TABLA proyecto org.xlsx" y genera:
   · correlaciones Pearson / Spearman + mapa de calor completo
   · mapa de color de registros y compuestos
   · densidades (lineal y log10) por rol y por medio
   · boxplots, violines, pairplot, dispersiones, PCA, dendrograma,
     radar, barras y compuestos frecuentes
   · Excel con las tablas limpias + ZIP con todas las figuras

 USO
   1. Deja "TABLA proyecto org.xlsx" junto a este archivo
      (o en Descargas: el script lo busca solo).
   2. Abre PowerShell en esta carpeta y ejecuta:
          python PFQ_proyecto.py

 SI FALTA ALGUNA LIBRERÍA:
          python -m pip install pandas numpy matplotlib seaborn scikit-learn scipy openpyxl
====================================================================
"""
import os, sys

# trabajar siempre en la carpeta del script (ahí quedan las salidas)
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# 'display' solo existe en Jupyter/Colab; aquí lo sustituimos por print
def display(obj):
    try:
        print(obj.to_string() if hasattr(obj, 'to_string') else obj)
    except Exception:
        print(obj)

# ============================================================
# 1) LIBRERÍAS Y CONFIGURACIÓN
# ============================================================
import os, re, zipfile, warnings
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
from scipy.cluster import hierarchy
from sklearn.preprocessing import StandardScaler, RobustScaler
from sklearn.decomposition import PCA
from sklearn.impute import SimpleImputer

warnings.filterwarnings('ignore')
sns.set_theme(style='whitegrid')
plt.rcParams['figure.dpi'] = 110
plt.rcParams['font.size'] = 10
plt.rcParams['axes.titlesize'] = 12
plt.rcParams['axes.titleweight'] = 'bold'
plt.rcParams['savefig.bbox'] = 'tight'

ARCHIVO = 'TABLA proyecto org.xlsx'   # nombre del Excel
HOJA    = 0                            # 0 = primera hoja
CARPETA = 'graficas'                   # carpeta de salida
GUARDAR = True                         # False si solo quieres verlas
os.makedirs(CARPETA, exist_ok=True)

print('✅ Librerías cargadas | pandas', pd.__version__, '| seaborn', sns.__version__)


# ============================================================
# 2) LOCALIZAR EL ARCHIVO EXCEL
#    Busca en la carpeta del script y en Descargas.
# ============================================================
if not os.path.exists(ARCHIVO):
    candidatos = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), 'TABLA proyecto org.xlsx'),
        os.path.join(os.path.expanduser('~'), 'Downloads', 'TABLA proyecto org.xlsx'),
        os.path.join(os.path.expanduser('~'), 'Descargas', 'TABLA proyecto org.xlsx'),
    ]
    for c in candidatos:
        if os.path.exists(c):
            ARCHIVO = c
            break

if not os.path.exists(ARCHIVO):
    sys.exit('❌ No encontré "TABLA proyecto org.xlsx".\n'
             '   Ponlo en la misma carpeta que este script y vuelve a ejecutar.')

print('📄 Archivo:', ARCHIVO)


# ============================================================
# 3) LIMPIEZA NUMÉRICA
#    Convierte el texto sucio de Excel en números reales:
#    '6,26 x 10⁻⁴ (gas)' -> 0.000626   |  '40 a 50' -> 45
#    '452,279±45,00' -> 452.279        |  '/' -> NaN
# ============================================================
SUP = {'⁻': '-', '⁺': '+', '⁰': '0', '¹': '1', '²': '2', '³': '3',
       '⁴': '4', '⁵': '5', '⁶': '6', '⁷': '7', '⁸': '8', '⁹': '9'}
NULOS = {'', '/', '-', '--', '---', 'na', 'n/a', 'nd', 'n/d', 'none', 'nan', '?', '.', 's/d'}

def _potencias(texto):
    # '10⁻⁴' -> '10^-4'  (si no, quedaría '104')
    def rep(m):
        return '^' + ''.join(SUP.get(c, c) for c in m.group(0))
    return re.sub(r'[⁻⁺⁰¹²³⁴⁵⁶⁷⁸⁹]+', rep, texto)

def _a_float(t):
    t = t.strip().replace(' ', '')
    if ',' in t and '.' in t:
        if t.rfind(',') > t.rfind('.'):      # coma decimal, punto de miles
            t = t.replace('.', '').replace(',', '.')
        else:                                 # punto decimal, coma de miles
            t = t.replace(',', '')
    elif ',' in t:
        t = t.replace(',', '.')
    t = t.rstrip('.')
    try:
        return float(t)
    except ValueError:
        return np.nan

def to_num(x):
    '''Convierte cualquier celda de la tabla en float (o NaN).'''
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return np.nan
    if isinstance(x, (int, float, np.integer, np.floating)):
        v = float(x)
        return v if np.isfinite(v) else np.nan
    s = str(x).replace('−', '-').replace('–', '-').replace('—', '-').replace('×', 'x')
    s = _potencias(s).strip()
    s = re.sub(r'\(.*?\)', ' ', s).strip()          # quita notas: (gas), (D)
    if s.lower() in NULOS:
        return np.nan
    # rangos: '1,8 - 2,0' | '40 a 50' | '-3,5 a -3,0'  -> punto medio
    m = re.search(r'^(-?[\d.,]+)\s*(?:a|hasta|to|-)\s*(-?[\d.,]+)$', s, flags=re.IGNORECASE)
    if m:
        a, b = _a_float(m.group(1)), _a_float(m.group(2))
        if np.isfinite(a) and np.isfinite(b):
            return (a + b) / 2
    # incertidumbre: '452,279±45,00' -> valor central
    if '±' in s:
        s = s.split('±')[0]
    # notación científica: '6,26x10^-4' | '3,138 x 10^7'
    m = re.search(r'(-?\d[\d.,]*)\s*x\s*10\s*\^?\s*(-?\d+)', s)
    if m:
        base, expo = _a_float(m.group(1)), _a_float(m.group(2))
        if np.isfinite(base) and np.isfinite(expo):
            return base * (10 ** int(expo))
    return _a_float(s)

print('Prueba del limpiador:')
for p in ['6,26 x 10⁻⁴  (gas)', '3,138 x 10⁷', '40 a 50', '1,8 - 2,0',
          '452,279±45,00', '/', '1.283', '-182.6', '5,19 x 10⁵', '0.026']:
    print(f'  {p!r:26s} -> {to_num(p)}')


# ============================================================
# 4) ESTRUCTURA DE LA TABLA
#    Bloque A (cols 0-31)  : SUSTRATO + PRODUCTO + condiciones
#    Bloque B (cols 33-54) : REACTANTE + condiciones
# ============================================================
PROPS = ['Peso molecular', 'Punto de fusión', 'Punto de ebullición', 'Densidad',
         'pKa', 'Momento dipolar', 'Solubilidad', 'Presión de vapor']
UNIDAD = {'Peso molecular': 'g/mol', 'Punto de fusión': '°C', 'Punto de ebullición': '°C',
          'Densidad': 'g/cm³', 'pKa': '(adim.)', 'Momento dipolar': 'D',
          'Solubilidad': 'g/L', 'Presión de vapor': 'Pa'}
# propiedades que abarcan varios órdenes de magnitud -> se grafican en log10
PROPS_LOG = ['Peso molecular', 'Densidad', 'Solubilidad', 'Presión de vapor']

# rangos físicamente posibles (sirven para detectar errores de digitación)
RANGOS = {'Peso molecular': (1, 2000), 'Punto de fusión': (-273, 1200),
          'Punto de ebullición': (-273, 1200), 'Densidad': (1e-4, 25),
          'pKa': (-25, 70), 'Momento dipolar': (0, 20),
          'Solubilidad': (0, 5000), 'Presión de vapor': (0, 1e8)}

crudo = pd.read_excel(ARCHIVO, sheet_name=HOJA, header=None)
encabezados = [str(c).strip() if pd.notna(c) else f'COL_{i}' for i, c in enumerate(crudo.iloc[0])]

# nombres únicos (la tabla repite 'REACTIVO', 'TEMPERATURA REACCIÓN (°C)', etc.)
vistos, unicos = {}, []
for h in encabezados:
    vistos[h] = vistos.get(h, 0) + 1
    unicos.append(h if vistos[h] == 1 else f'{h} #{vistos[h]}')
encabezados = unicos

datos = crudo.iloc[1:].reset_index(drop=True)
datos.columns = encabezados
datos = datos.dropna(how='all').reset_index(drop=True)
print('Filas de datos:', len(datos), '| columnas:', len(datos.columns))

def construir_bloque(df, i_nombre, i_props, i_temp, i_medio, i_tipo, rol):
    '''Extrae un bloque (rol) por POSICIÓN, porque hay encabezados repetidos.'''
    out = pd.DataFrame()
    out['compuesto'] = df.iloc[:, i_nombre].astype(str).str.strip()
    for k, p in enumerate(PROPS):
        out[p] = df.iloc[:, i_props[0] + k].map(to_num)
    out['temperatura_C'] = df.iloc[:, i_temp].map(to_num)
    out['medio'] = df.iloc[:, i_medio].astype(str).str.strip()
    out['tipo_reaccion'] = df.iloc[:, i_tipo].astype(str).str.strip()
    out['rol'] = rol
    out['id'] = np.arange(len(out))
    return out

sustrato  = construir_bloque(datos,  0, list(range(1, 9)),  25, 26, 27, 'Sustrato')
producto  = construir_bloque(datos, 10, list(range(11, 19)), 25, 26, 27, 'Producto')
reactante = construir_bloque(datos, 33, list(range(34, 42)), 48, 49, 50, 'Reactante')

# ---------------- NORMALIZAR EL 'MEDIO' ----------------
# La columna trae ~20 variantes ('Neutro', 'neutro', 'Neutro-polar', 'basico'...).
# Se agrupan en 6 categorías para que las gráficas sean legibles.
import unicodedata

def normalizar_medio(s):
    t = unicodedata.normalize('NFKD', str(s).lower()).encode('ascii', 'ignore').decode()
    t = re.sub(r'[^a-z ]', ' ', t)
    t = re.sub(r'\s+', ' ', t).strip()
    if not t or t == 'nan':
        return np.nan
    if 'acido' in t:     return 'Ácido'
    if 'basico' in t:    return 'Básico'
    if 'anhidrido' in t: return 'Anhídrido'
    if 'neutro' in t:    return 'Neutro'
    if 'apolar' in t:    return 'Apolar'
    if 'polar' in t:     return 'Polar'
    return 'Otro'

for _d in (sustrato, producto, reactante):
    _d['medio'] = _d['medio'].map(normalizar_medio)

# ---------------- CONTROL DE CALIDAD ----------------
atipicos = []
for d in (sustrato, producto, reactante):
    for prop, (lo, hi) in RANGOS.items():
        mal = d[(d[prop] < lo) | (d[prop] > hi)]
        for _, r in mal.iterrows():
            atipicos.append({'propiedad': prop, 'rol': r['rol'], 'compuesto': r['compuesto'][:40],
                             'valor_leido': round(float(r[prop]), 4),
                             'rango_esperado': f'{lo} a {hi}'})
tabla_atipicos = pd.DataFrame(atipicos)
print(f'\n⚠️  {len(tabla_atipicos)} valores fuera de rango físico (revisar en el Excel):')
display(tabla_atipicos)

FILTRAR_ATIPICOS = True
if FILTRAR_ATIPICOS:
    for d in (sustrato, producto, reactante):
        for prop, (lo, hi) in RANGOS.items():
            d.loc[(d[prop] < lo) | (d[prop] > hi), prop] = np.nan
    print('→ Valores imposibles convertidos en NaN para no distorsionar las gráficas.')

# ---------- tabla ANCHA: una fila por reacción, 24 columnas numéricas ----------
def prefijar(d, rol):
    return d[PROPS].add_suffix(' · ' + rol)

numericas = pd.concat([prefijar(sustrato, 'Sustrato'),
                       prefijar(producto, 'Producto'),
                       prefijar(reactante, 'Reactante')], axis=1)

# ---------- tabla LARGA: compuesto / rol / propiedad / valor ----------
largo = pd.concat([
    d.melt(id_vars=['id', 'compuesto', 'rol', 'temperatura_C', 'medio', 'tipo_reaccion'],
           value_vars=PROPS, var_name='propiedad', value_name='valor')
    for d in (sustrato, producto, reactante)
], ignore_index=True)

# ---------- tabla de COMPUESTOS (promedio por compuesto) ----------
compuestos = (largo.pivot_table(index=['rol', 'compuesto'], columns='propiedad',
                                values='valor', aggfunc='mean').reindex(columns=PROPS))

print('\nValores válidos por propiedad y rol:')
display(largo.groupby(['rol', 'propiedad'])['valor'].agg(['count', 'median']).round(3).unstack(0))
print('\nTablas: numericas', numericas.shape, '| largo', largo.shape, '| compuestos', compuestos.shape)


# ============================================================
# 5) UTILIDADES DE GRAFICACIÓN
# ============================================================
def guardar(fig, nombre):
    if GUARDAR:
        fig.savefig(os.path.join(CARPETA, nombre), dpi=150, bbox_inches='tight')

def limites_tukey(v):
    '''Rango intercuartílico: sirve para recortar valores extremos al graficar.'''
    q1, q3 = v.quantile(.25), v.quantile(.75)
    iqr = q3 - q1
    return q1 - 1.5 * iqr, q3 + 1.5 * iqr

def abarca_ordenes(v, factor=100, tolerancia=0.02):
    '''¿La variable abarca muchos órdenes de magnitud? (decide si usar log10)
    Exige que casi todos los valores sean positivos: log10 no admite ≤ 0.'''
    v = v.dropna()
    if len(v) < 5:
        return False
    if (v <= 0).mean() > tolerancia:
        return False
    pos = v[v > 0]
    return pos.max() / max(pos.quantile(.05), 1e-12) > factor

def preparar_matriz(df, min_datos=20, min_filas=2):
    '''Selecciona columnas con datos, aplica log10 a las de cola pesada
    y devuelve una estandarización robusta (mediana / rango intercuartílico).'''
    m = df.dropna(axis=1, thresh=min_datos).copy()
    m = m.dropna(thresh=max(min_filas, int(m.shape[1] * 0.30)))
    for c in m.columns:
        if abarca_ordenes(m[c]):
            m[c] = np.log10(m[c].where(m[c] > 0))
    q1, q3 = m.quantile(.25), m.quantile(.75)
    iqr = (q3 - q1).replace(0, np.nan)
    z = ((m - m.median()) / iqr).clip(-3, 3)
    return m, z

# tabla ancha por propiedad (para dispersiones)
ancho_props = (largo.pivot_table(index=['id', 'rol'], columns='propiedad',
                                 values='valor', aggfunc='mean').reset_index())

# abreviaturas para etiquetas cortas (biplot del PCA)
ABREV = {'Peso molecular': 'PM', 'Punto de fusión': 'PF', 'Punto de ebullición': 'PE',
         'Densidad': 'ρ', 'pKa': 'pKa', 'Momento dipolar': 'μ',
         'Solubilidad': 'Sol', 'Presión de vapor': 'PV'}
ROL_ABREV = {'Sustrato': 'S', 'Producto': 'P', 'Reactante': 'R'}

def etiqueta_corta(nombre):
    partes = str(nombre).split(' · ')
    if len(partes) == 2:
        return ABREV.get(partes[0], partes[0][:4]) + '·' + ROL_ABREV.get(partes[1], partes[1][:1])
    return nombre

print('✅ Utilidades listas')


# ============================================================
# 6) MAPA DE CALOR COMPLETO DE CORRELACIONES (Pearson y Spearman)
#    Relación entre TODAS las PFQ de sustrato, producto y reactante.
# ============================================================
MIN_PARES = 8          # mínimo de pares para calcular una correlación
pearson  = numericas.corr(method='pearson',  min_periods=MIN_PARES)
spearman = numericas.corr(method='spearman', min_periods=MIN_PARES)

def mapa_correlacion(matriz, titulo, archivo, cmap='coolwarm'):
    fig, ax = plt.subplots(figsize=(15, 12))
    sns.heatmap(matriz, annot=True, fmt='.2f', cmap=cmap, center=0, vmin=-1, vmax=1,
                square=True, linewidths=.4, linecolor='white',
                annot_kws={'size': 5.5},
                cbar_kws={'shrink': .7, 'label': 'coeficiente'}, ax=ax)
    ax.set_title(titulo + '\n(matriz completa: las 24 propiedades)', pad=14)
    ax.set_xticklabels([t.get_text().replace(' · ', '\n') for t in ax.get_xticklabels()],
                       rotation=90, fontsize=6.5)
    ax.set_yticklabels(ax.get_yticklabels(), fontsize=6.5)
    guardar(fig, archivo)
    plt.close('all')

mapa_correlacion(pearson,  'Correlación de PEARSON entre propiedades fisicoquímicas',
                 '01_correlacion_pearson.png')
mapa_correlacion(spearman, 'Correlación de SPEARMAN (robusta a valores extremos)',
                 '02_correlacion_spearman.png')

print('Celdas con correlación válida (Pearson):',
      int(pearson.notna().sum().sum()), 'de', pearson.size)


# ============================================================
# 7) TOP CORRELACIONES (leer la tabla anterior sin perderse)
# ============================================================
def top_correlaciones(matriz, n=20):
    m = matriz.where(np.triu(np.ones(matriz.shape), k=1).astype(bool)).stack()
    return m.dropna().sort_values(key=abs, ascending=False).head(n)

tabla_top = top_correlaciones(pearson, 20).rename('Pearson').to_frame()
tabla_top['Spearman'] = [spearman.loc[a, b] for a, b in tabla_top.index]
display(tabla_top.round(3))

fig, ax = plt.subplots(figsize=(10, 7))
etiquetas = [f'{a}  ↔  {b}' for a, b in tabla_top.index]
colores = ['#2a9d8f' if v > 0 else '#e76f51' for v in tabla_top['Pearson']]
ax.barh(range(len(etiquetas)), tabla_top['Pearson'], color=colores, edgecolor='white')
ax.set_yticks(range(len(etiquetas)))
ax.set_yticklabels(etiquetas, fontsize=7.5)
ax.invert_yaxis(); ax.axvline(0, color='k', lw=.8)
ax.set_xlabel('Coeficiente de Pearson')
ax.set_title('Top 20 correlaciones más fuertes entre PFQ')
for i, v in enumerate(tabla_top['Pearson']):
    ax.text(v + (0.02 if v >= 0 else -0.02), i, f'{v:.2f}', va='center',
            ha='left' if v >= 0 else 'right', fontsize=7)
guardar(fig, '03_top_correlaciones.png')
plt.close('all')


# ============================================================
# 8) MAPA DE COLOR COMPLETO
#    Filas = registros o compuestos | Columnas = propiedades
#    Escala robusta (z): azul = bajo, rojo = alto.
# ============================================================
# --- 8a) por REGISTRO (todas las reacciones) ---
m_reg, z_reg = preparar_matriz(numericas, min_datos=20)
filas = hierarchy.linkage(z_reg.fillna(0).values, method='average')
cols  = hierarchy.linkage(z_reg.fillna(0).values.T, method='average')

fig, ax = plt.subplots(figsize=(14, 10))
sns.heatmap(z_reg.iloc[hierarchy.leaves_list(filas), hierarchy.leaves_list(cols)],
            cmap='RdBu_r', center=0, vmin=-3, vmax=3, yticklabels=False,
            linewidths=0, cbar_kws={'label': 'z-score robusto', 'shrink': .7}, ax=ax)
ax.set_title('Mapa de color completo: registros (filas) × propiedades (columnas)\n'
             'z-score robusto, filas y columnas agrupadas por similitud')
ax.set_xticklabels([t.get_text().replace(' · ', '\n') for t in ax.get_xticklabels()],
                   rotation=90, fontsize=7)
guardar(fig, '04_mapa_calor_completo.png')
plt.close('all')

# --- 8b) por COMPUESTO (promedio de sus propiedades) ---
m_comp, z_comp = preparar_matriz(compuestos, min_datos=20)
# quedarse con compuestos que tengan al menos 4 propiedades medidas
suficientes = m_comp.notna().sum(axis=1) >= 4
m_comp, z_comp = m_comp[suficientes], z_comp[suficientes]
if len(m_comp) > 90:                      # si hay muchísimos, los 90 más completos
    idx = m_comp.notna().sum(axis=1).sort_values(ascending=False).head(90).index
    m_comp, z_comp = m_comp.loc[idx], z_comp.loc[idx]
print('Compuestos en el mapa:', len(m_comp))

if len(m_comp) > 3:
    filas = hierarchy.linkage(z_comp.fillna(0).values, method='average')
    cols  = hierarchy.linkage(z_comp.fillna(0).values.T, method='average')
    etiquetas = [f'{r[:3]} | {c[:38]}' for r, c in z_comp.index]
    fig, ax = plt.subplots(figsize=(11, max(6, 0.26 * len(m_comp))))
    sns.heatmap(z_comp.iloc[hierarchy.leaves_list(filas), hierarchy.leaves_list(cols)],
                cmap='RdBu_r', center=0, vmin=-3, vmax=3,
                yticklabels=[etiquetas[i] for i in hierarchy.leaves_list(filas)],
                linewidths=.2, linecolor='white',
                cbar_kws={'label': 'z-score robusto', 'shrink': .6}, ax=ax)
    ax.set_title('Mapa de color por compuesto (promedio de sus propiedades)\n'
                 'Sust = sustrato · Pro = producto · Rea = reactante')
    ax.tick_params(axis='y', labelsize=6)
    ax.set_xticklabels(ax.get_xticklabels(), rotation=45, ha='right', fontsize=8)
    guardar(fig, '05_mapa_calor_compuestos.png')
    plt.close('all')


# ============================================================
# 9) DENSIDADES (KDE)
#    'largo' tiene columnas: rol | propiedad | valor
# ============================================================
def panel_densidades(df, titulo, archivo, props, por_rol=True, ncols=4,
                     recortar=True, log=False):
    props = [p for p in props if df.loc[df['propiedad'] == p, 'valor'].notna().sum() >= 5]
    if not props:
        print('Sin datos suficientes'); return
    nfilas = int(np.ceil(len(props) / ncols))
    fig, axes = plt.subplots(nfilas, ncols, figsize=(4.2 * ncols, 3.4 * nfilas))
    axes = np.atleast_1d(axes).ravel()
    for ax, p in zip(axes, props):
        d = df.loc[df['propiedad'] == p, ['rol', 'valor']].dropna()
        if log:
            descartados = int((d['valor'] <= 0).sum())
            if descartados:
                print(f'  · {p}: {descartados} valores ≤ 0 no se pueden pasar a log10')
            d = d[d['valor'] > 0].copy()
            d['valor'] = np.log10(d['valor'])
            unidad = 'log10(' + UNIDAD[p] + ')'
        else:
            unidad = UNIDAD[p]
            if recortar and len(d) > 8:
                lo, hi = limites_tukey(d['valor'])
                d = d[(d['valor'] >= lo) & (d['valor'] <= hi)]
        if por_rol:
            for rol, g in d.groupby('rol'):
                if len(g) >= 5:
                    sns.kdeplot(x=g['valor'], ax=ax, fill=True, alpha=.25,
                                label=rol, warn_singular=False)
        sns.histplot(d['valor'], ax=ax, bins=16, stat='density', alpha=.15,
                     color='gray', edgecolor='white')
        ax.set_title(f'{p}\n({unidad})')
        ax.set_xlabel(''); ax.set_ylabel('densidad')
        if por_rol:
            ax.legend(fontsize=6.5, frameon=False)
    for ax in axes[len(props):]:
        ax.axis('off')
    fig.suptitle(titulo, fontsize=14, fontweight='bold', y=1.005)
    fig.tight_layout()
    guardar(fig, archivo)
    plt.close('all')

# 9a) escala lineal, con recorte de valores extremos
panel_densidades(largo, 'Densidad de cada PFQ comparando roles (escala lineal, sin extremos)',
                 '06_densidades_por_rol.png', PROPS)

# 9b) escala log10 para las propiedades que abarcan varios órdenes de magnitud
panel_densidades(largo, 'Densidad de las PFQ en escala log10 (incluye gases y valores muy bajos)',
                 '06b_densidades_log.png', PROPS_LOG, log=True)

# 9c) densidad por MEDIO de reacción
fig, axes = plt.subplots(2, 4, figsize=(17, 7.5))
axes = axes.ravel()
for ax, p in zip(axes, PROPS):
    d = largo.loc[largo['propiedad'] == p, ['medio', 'valor']].dropna()
    d = d[d['medio'].notna()]
    unidad = UNIDAD[p]
    if abarca_ordenes(d['valor']):          # propiedades de varios órdenes -> log10
        d = d[d['valor'] > 0].copy()
        d['valor'] = np.log10(d['valor'])
        unidad = 'log10 ' + unidad
    else:
        lo, hi = limites_tukey(d['valor'])
        d = d[(d['valor'] >= lo) & (d['valor'] <= hi)]
    for medio, g in d.groupby('medio'):
        if len(g) >= 6:
            sns.kdeplot(x=g['valor'], ax=ax, lw=1.6, label=medio[:22], warn_singular=False)
    ax.set_title(p)
    ax.set_xlabel(unidad); ax.set_ylabel('densidad')
    ax.legend(fontsize=6, frameon=False)
fig.suptitle('Densidad de las PFQ según el MEDIO de reacción', fontsize=14, fontweight='bold')
fig.tight_layout()
guardar(fig, '07_densidades_por_medio.png')
plt.close('all')


# ============================================================
# 10) BOXPLOTS Y VIOLINES
# ============================================================
fig, axes = plt.subplots(2, 4, figsize=(17, 7.5))
axes = axes.ravel()
for ax, p in zip(axes, PROPS):
    d = largo.loc[largo['propiedad'] == p, ['rol', 'valor']].dropna()
    sns.boxplot(data=d, x='rol', y='valor', ax=ax, width=.55,
                palette='Set2', fliersize=2, linewidth=.8)
    sns.stripplot(data=d, x='rol', y='valor', ax=ax, color='k', size=2, alpha=.35)
    lo, hi = limites_tukey(d['valor'])
    margen = (hi - lo) * .08
    ax.set_ylim(lo - margen, hi + margen)
    ax.set_title(p); ax.set_xlabel(''); ax.set_ylabel(UNIDAD[p] + ' (eje recortado)')
    ax.tick_params(axis='x', labelrotation=15, labelsize=8)
fig.suptitle('Distribución de las PFQ por rol del compuesto', fontsize=14, fontweight='bold')
fig.tight_layout()
guardar(fig, '08_boxplots_por_rol.png')
plt.close('all')

# violines por medio (4 propiedades con más datos)
elegidas = (largo.groupby('propiedad')['valor'].count().sort_values(ascending=False)
            .head(4).index.tolist())
fig, axes = plt.subplots(1, 4, figsize=(18, 4.8))
for ax, p in zip(axes, elegidas):
    d = largo[largo['propiedad'] == p].dropna(subset=['valor'])
    d = d[d['medio'].notna()]
    lo, hi = limites_tukey(d['valor'])
    d = d[(d['valor'] >= lo) & (d['valor'] <= hi)]
    sns.violinplot(data=d, x='medio', y='valor', ax=ax, palette='mako',
                   cut=0, inner='quartile', linewidth=.7)
    ax.set_title(p); ax.set_xlabel(''); ax.set_ylabel(UNIDAD[p])
    ax.tick_params(axis='x', rotation=30, labelsize=7)
fig.suptitle('Forma de la distribución de cada PFQ según el medio', fontsize=14, fontweight='bold')
fig.tight_layout()
guardar(fig, '09_violines_por_medio.png')
plt.close('all')


# ============================================================
# 11) MATRIZ DE DISPERSIÓN (pairplot)
#     Se usan las 6 propiedades más completas; las que abarcan
#     varios órdenes de magnitud se pasan a log10.
# ============================================================
top6 = (largo.groupby('propiedad')['valor'].count().sort_values(ascending=False)
        .head(6).index.tolist())
d6 = ancho_props[['rol'] + top6].copy()
renombres = {}
for c in top6:
    if abarca_ordenes(d6[c]):
        d6[c] = np.log10(d6[c].where(d6[c] > 0))
        renombres[c] = c + ' (log10)'
d6 = d6.rename(columns=renombres).dropna(thresh=3)
print('Propiedades usadas:', list(d6.columns[1:]), '| filas:', len(d6))

g = sns.pairplot(d6, hue='rol', corner=False, diag_kind='kde',
                 plot_kws={'s': 16, 'alpha': .65, 'edgecolor': 'none'},
                 height=2.0, palette='Set2')
g.fig.suptitle('Matriz de dispersión: relaciones entre las PFQ más completas',
               y=1.02, fontsize=14, fontweight='bold')
if GUARDAR:
    g.fig.savefig(os.path.join(CARPETA, '10_pairplot.png'), dpi=150, bbox_inches='tight')
plt.close('all')


# ============================================================
# 12) DISPERSIONES CLAVE con línea de tendencia
# ============================================================
pares = [('Peso molecular', 'Punto de ebullición'),
         ('Peso molecular', 'Densidad'),
         ('Peso molecular', 'Punto de fusión'),
         ('Punto de fusión', 'Punto de ebullición')]

fig, axes = plt.subplots(2, 2, figsize=(13, 9))
for ax, (x, y) in zip(axes.ravel(), pares):
    d = ancho_props[['rol', x, y]].dropna().copy()
    xl, yl = f'{x} ({UNIDAD[x]})', f'{y} ({UNIDAD[y]})'
    if abarca_ordenes(d[x]):
        d[x] = np.log10(d[x].where(d[x] > 0)); xl = f'log10 · {x} ({UNIDAD[x]})'
    if abarca_ordenes(d[y]):
        d[y] = np.log10(d[y].where(d[y] > 0)); yl = f'log10 · {y} ({UNIDAD[y]})'
    d = d.dropna()
    if len(d) < 3:
        ax.set_title(f'{x} vs {y} (sin datos)'); continue
    sns.scatterplot(data=d, x=x, y=y, hue='rol', s=45, alpha=.8,
                    palette='Set2', edgecolor='white', ax=ax)
    b, a = np.polyfit(d[x], d[y], 1)
    xs = np.linspace(d[x].min(), d[x].max(), 50)
    r2 = np.corrcoef(d[x], d[y])[0, 1] ** 2
    ax.plot(xs, a + b * xs, 'k--', lw=1.2, label=f'ajuste  R² = {r2:.2f}')
    ax.legend(fontsize=7, frameon=False)
    ax.set_title(f'{x}  vs  {y}')
    ax.set_xlabel(xl); ax.set_ylabel(yl)
fig.suptitle('Relaciones entre propiedades (con tendencia lineal)', fontsize=14, fontweight='bold')
fig.tight_layout()
guardar(fig, '11_dispersion_propiedades.png')
plt.close('all')


# ============================================================
# 13) ANÁLISIS DE COMPONENTES PRINCIPALES (PCA)
#     Resume las 24 propiedades en 2 ejes para ver agrupamientos.
# ============================================================
X = numericas.dropna(axis=1, thresh=20)
m_pca, z_pca = preparar_matriz(X, min_datos=20)
Xz = z_pca.fillna(0).values          # z-score robusto ya recortado a ±3

pca = PCA(n_components=min(4, Xz.shape[1]))
scores = pca.fit_transform(Xz)
var = pca.explained_variance_ratio_ * 100
print('Varianza explicada (%):', np.round(var, 1), '| acumulada 2D:', round(var[:2].sum(), 1), '%')

meta = sustrato[['id', 'medio']].reset_index(drop=True)
d = pd.DataFrame(scores[:, :2], columns=['PC1', 'PC2'])
d['medio'] = meta['medio'].values[:len(d)]
d['medio'] = d['medio'].fillna('sin dato')

# límites comunes para que puntos y flechas se vean bien
lim = float(np.percentile(np.abs(scores[:, :2]), 98) * 1.25)

fig, axes = plt.subplots(1, 2, figsize=(15, 6.4))
ax = axes[0]
sns.scatterplot(data=d, x='PC1', y='PC2', hue='medio', s=55, palette='tab10',
                alpha=.85, ax=ax, edgecolor='white')
ax.set_title(f'PCA — registros (PC1 {var[0]:.1f}% · PC2 {var[1]:.1f}%)')
ax.set_xlabel(f'PC1 ({var[0]:.1f}%)'); ax.set_ylabel(f'PC2 ({var[1]:.1f}%)')
ax.set_xlim(-lim, lim); ax.set_ylim(-lim, lim)
ax.legend(fontsize=7.5, frameon=False, title='medio', title_fontsize=8,
          loc='upper right', bbox_to_anchor=(1.32, 1))

ax = axes[1]
cargas = pd.DataFrame(pca.components_[:2].T, index=m_pca.columns, columns=['PC1', 'PC2'])
factor = 0.85 * lim / max(np.abs(cargas.values).max(), 1e-9)
for nombre, fila in cargas.iterrows():
    x, y = fila.PC1 * factor, fila.PC2 * factor
    ax.arrow(0, 0, x, y, color='#333', width=.002, head_width=.05,
             alpha=.75, length_includes_head=True)
    ax.text(x * 1.12, y * 1.12, etiqueta_corta(nombre), fontsize=7.5,
            ha='center', va='center', fontweight='bold')
ax.set_xlim(-lim, lim); ax.set_ylim(-lim, lim)
ax.axhline(0, color='gray', lw=.6); ax.axvline(0, color='gray', lw=.6)
ax.set_title('PCA — aporte de cada propiedad (biplot)')
ax.set_xlabel(f'PC1 ({var[0]:.1f}%)'); ax.set_ylabel(f'PC2 ({var[1]:.1f}%)')
fig.suptitle('Análisis de componentes principales de las PFQ', fontsize=14, fontweight='bold')
fig.tight_layout()
guardar(fig, '12_pca.png')
plt.close('all')


# ============================================================
# 14) DENDROGRAMA: agrupamiento de compuestos parecidos
# ============================================================
Z = z_comp.dropna(axis=1, thresh=20)
Z = Z.fillna(Z.median())
etiquetas = [f'{r[:3]}·{c[:26]}' for r, c in Z.index]
enlace = hierarchy.linkage(StandardScaler().fit_transform(Z), method='average')

fig, ax = plt.subplots(figsize=(11, max(6, 0.20 * len(Z))))
hierarchy.dendrogram(enlace, labels=etiquetas, orientation='right',
                     leaf_font_size=5.5, color_threshold=0.7 * max(enlace[:, 2]), ax=ax)
ax.set_title('Dendrograma: similitud entre compuestos según sus PFQ\n'
             'Sust = sustrato · Pro = producto · Rea = reactante')
ax.set_xlabel('distancia (propiedades estandarizadas)')
guardar(fig, '13_dendrograma.png')
plt.close('all')


# ============================================================
# 15) VARIABLES CATEGÓRICAS Y TEMPERATURA DE REACCIÓN
# ============================================================
fig, axes = plt.subplots(1, 3, figsize=(18, 5))

ax = axes[0]
v = sustrato['medio'].replace('', np.nan).dropna().value_counts().head(10)
sns.barplot(x=v.values, y=v.index, ax=ax, palette='crest')
ax.set_title('Medio de reacción más frecuente')
ax.set_xlabel('nº de reacciones'); ax.set_ylabel('')

ax = axes[1]
v = (sustrato['tipo_reaccion'].replace('', np.nan).dropna()
     .str.slice(0, 45).value_counts().head(10))
sns.barplot(x=v.values, y=v.index, ax=ax, palette='flare')
ax.set_title('Tipos de reacción más frecuentes')
ax.set_xlabel('nº de reacciones'); ax.set_ylabel('')
ax.tick_params(axis='y', labelsize=7)

ax = axes[2]
t = sustrato['temperatura_C'].dropna()
sns.histplot(t, bins=20, kde=True, ax=ax, color='#457b9d', edgecolor='white')
ax.axvline(t.median(), color='crimson', ls='--', label=f'mediana = {t.median():.0f} °C')
ax.set_title('Temperatura de reacción'); ax.set_xlabel('°C'); ax.set_ylabel('frecuencia')
ax.legend(fontsize=8, frameon=False)
fig.suptitle('Contexto de las reacciones', fontsize=14, fontweight='bold')
fig.tight_layout()
guardar(fig, '14_categoricas_temperatura.png')
plt.close('all')

print('Temperatura: n =', len(t), '| mín', t.min(), '| mediana', t.median(), '| máx', t.max())


# ============================================================
# 16) PERFIL COMPARADO (gráfico radial) de cada rol
# ============================================================
perfil = z_comp.reset_index().groupby('rol')[list(z_comp.columns)].mean()
perfil = perfil.dropna(axis=1, how='all')
angulos = np.linspace(0, 2 * np.pi, len(perfil.columns), endpoint=False).tolist()
angulos += angulos[:1]

fig = plt.figure(figsize=(7.5, 7))
ax = fig.add_subplot(111, polar=True)
for rol, fila in perfil.iterrows():
    vals = fila.fillna(0).tolist(); vals += vals[:1]
    ax.plot(angulos, vals, lw=2, label=rol)
    ax.fill(angulos, vals, alpha=.12)
ax.set_xticks(angulos[:-1])
ax.set_xticklabels([c.replace(' ', '\n') for c in perfil.columns], fontsize=8)
ax.set_rlabel_position(200)
ax.tick_params(axis='y', labelsize=7)
ax.set_title('Perfil promedio de cada rol (z-score)\n0 = promedio general', pad=25,
             fontweight='bold')
ax.legend(loc='upper right', bbox_to_anchor=(1.25, 1.12), fontsize=9, frameon=False)
guardar(fig, '15_radar_perfiles.png')
plt.close('all')


# ============================================================
# 17) COMPUESTOS MÁS REPETIDOS EN LA TABLA
# ============================================================
fig, axes = plt.subplots(1, 3, figsize=(18, 5.5))
for ax, (rol, d) in zip(axes, [('Sustrato', sustrato), ('Producto', producto),
                               ('Reactante', reactante)]):
    v = (d['compuesto'].replace('nan', np.nan).dropna()
         .str.slice(0, 40).value_counts().head(12))
    sns.barplot(x=v.values, y=v.index, ax=ax, palette='viridis')
    ax.set_title(f'{rol}: compuestos más frecuentes')
    ax.set_xlabel('apariciones'); ax.set_ylabel('')
    ax.tick_params(axis='y', labelsize=7)
fig.suptitle('Compuestos con más registros', fontsize=14, fontweight='bold')
fig.tight_layout()
guardar(fig, '16_compuestos_frecuentes.png')
plt.close('all')


# ============================================================
# 18) GUARDAR TABLAS Y EMPAQUETAR TODO
# ============================================================
salida_xlsx = 'PFQ_tablas_limpias.xlsx'
with pd.ExcelWriter(salida_xlsx, engine='openpyxl') as xw:
    numericas.to_excel(xw, sheet_name='Numericas')
    largo.to_excel(xw, sheet_name='Largo', index=False)
    compuestos.to_excel(xw, sheet_name='Compuestos')
    pearson.to_excel(xw, sheet_name='Pearson')
    spearman.to_excel(xw, sheet_name='Spearman')
    tabla_top.to_excel(xw, sheet_name='Top correlaciones')
    tabla_atipicos.to_excel(xw, sheet_name='Valores atipicos', index=False)
print('📗 Guardado:', salida_xlsx)

archivos = sorted(os.listdir(CARPETA))
with zipfile.ZipFile('graficas_PFQ.zip', 'w') as z:
    for f in archivos:
        z.write(os.path.join(CARPETA, f))
print('🗜️  graficas_PFQ.zip ->', len(archivos), 'figuras')
for f in archivos:
    print('   •', f)

# abrir la carpeta de resultados en el Explorador de Windows
try:
    os.startfile(os.path.abspath(CARPETA))
except Exception:
    print('Abre la carpeta:', os.path.abspath(CARPETA))

print('\n✅ LISTO. Revisa la carpeta:', os.path.abspath(CARPETA))

print('\n✅ Análisis terminado.')
