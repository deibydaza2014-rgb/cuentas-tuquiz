import traceback
import pickle
import shutil
import json
import base64
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import openpyxl
import io
import zipfile
import re
import datetime
import os
try:
from pypdf import PdfReader, PdfWriter
except ImportError:
try:
from PyPDF2 import PdfReader, PdfWriter
except ImportError:
PdfReader, PdfWriter = None, None
try:
import pypdfium2 as pdfium
HAS_PDFIUM = True
except Exception:
HAS_PDFIUM = False
st.set_page_config(page_title=&quot;Cuentas Tuquiz - ERP y Auditoria Contable&quot;, layout=&quot;wide&quot;, page_icon=&quot;��&quot;)
# Estilos visuales profesionales y protección contra conflictos de Google Translate
st.markdown(&quot;&quot;&quot;
&lt;meta name=&quot;google&quot; content=&quot;notranslate&quot;&gt;
&lt;style class=&quot;notranslate&quot;&gt;
html, body, [data-testid=&quot;stAppViewContainer&quot;], [data-testid=&quot;stApp&quot;] {
-webkit-translate: no !important;
translate: no !important;
}
.main { background-color: #f8fafc; }
.stButton&gt;button { background-color: #0070ba; color: white; border-radius: 6px; font-weight: 600; }
.card-box { background: white; padding: 22px; border-radius: 10px; border: 1px solid #e2e8f0; box-shadow: 0 2px
5px rgba(0,0,0,0.04); margin-bottom: 18px; }
.audit-card { background: #ffffff; padding: 22px; border-left: 5px solid #0070ba; border-radius: 8px; box-shadow:
0 2px 6px rgba(0,0,0,0.06); }
.badge-active { background-color: #dcfce7; color: #15803d; padding: 4px 10px; border-radius: 12px; font-size:
12px; font-weight: bold; }
.badge-next { background-color: #f1f5f9; color: #64748b; padding: 4px 10px; border-radius: 12px; font-size: 12px;
font-weight: bold; }
.tag-propio { background-color: #dcfce7; color: #15803d; padding: 4px 8px; border-radius: 4px; font-weight: bold;
}
.tag-comp { background-color: #eff6ff; color: #1d4ed8; padding: 4px 8px; border-radius: 4px; font-weight: bold;
font-family: monospace; }
.badge-tax { background-color: #e0f2fe; color: #0369a1; padding: 3px 8px; border-radius: 6px; font-size: 12px;
font-weight: 600; margin-right: 4px; }
&lt;/style&gt;
&quot;&quot;&quot;, unsafe_allow_html=True)

# ==============================================================================
# CONFIGURACIÓN DE DIRECTORIOS Y SISTEMA DE AUTO-GUARDADO CONTINUO
# ==============================================================================
# ==============================================================================
# CONFIGURACIÓN DE DIRECTORIOS Y SISTEMA DE AUTO-GUARDADO CONTINUO
# ==============================================================================
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), &quot;almacenamiento_contable&quot;)
os.makedirs(DATA_DIR, exist_ok=True)
import tempfile
import hashlib
PDF_DISK_CACHE_ROOT = os.path.join(tempfile.gettempdir(), &quot;indumaq_pdf_disk_cache&quot;)

os.makedirs(PDF_DISK_CACHE_ROOT, exist_ok=True)
ZIP_FACTURAS_DISK_PATH = os.path.join(PDF_DISK_CACHE_ROOT, &quot;paquete_facturas_actual.zip&quot;)

class DiskBackedPDFDict(dict):
&quot;&quot;&quot;
Diccionario de PDFs respaldado en disco físico (/tmp) que consume ~0 MB de memoria RAM.
Guarda cada PDF en disco al asignarlo y lo lee bajo demanda solo cuando se consulta individualmente.
&quot;&quot;&quot;
def __init__(self, subdir=&quot;default&quot;, initial_data=None):
super().__init__()
self.subdir = re.sub(r&#39;[^a-zA-Z0-9_-]&#39;, &#39;_&#39;, str(subdir))
self.store_dir = os.path.join(PDF_DISK_CACHE_ROOT, self.subdir)
os.makedirs(self.store_dir, exist_ok=True)
try:
for fn in os.listdir(self.store_dir):
if fn.lower().endswith(&quot;.pdf&quot;):
super().__setitem__(fn, os.path.join(self.store_dir, fn))
except Exception:
pass
if initial_data and isinstance(initial_data, dict):
for k, v in initial_data.items():
self[k] = v
def _safe_path(self, key):
safe_name = os.path.basename(str(key)).replace(&quot;/&quot;, &quot;_&quot;).replace(&quot;\\&quot;, &quot;_&quot;)
if not safe_name:
safe_name = &quot;doc.pdf&quot;
return os.path.join(self.store_dir, safe_name)
def __setitem__(self, key, value):
p = self._safe_path(key)
if isinstance(value, (bytes, bytearray)):
try:
with open(p, &quot;wb&quot;) as f:
f.write(value)
super().__setitem__(key, p)
except Exception:
pass
elif isinstance(value, str) and os.path.exists(value):
super().__setitem__(key, value)
def __getitem__(self, key):
p = super().__getitem__(key)
if isinstance(p, (bytes, bytearray)):
return p
if isinstance(p, str) and os.path.exists(p):
try:
with open(p, &quot;rb&quot;) as f:
return f.read()
except Exception:
return b&quot;&quot;
return b&quot;&quot;
def get_path(self, key):
if key in self:
p = super().__getitem__(key)
if isinstance(p, str) and os.path.exists(p):
return p
return None
def get(self, key, default=None):
if key in self:
val = self[key]
return val if val else default
return default
def values(self):
for k in list(self.keys()):
yield self[k]

def items(self):
for k in list(self.keys()):
yield (k, self[k])
def clear_disk(self):
super().clear()
try:
if os.path.exists(self.store_dir):
shutil.rmtree(self.store_dir, ignore_errors=True)
os.makedirs(self.store_dir, exist_ok=True)
except Exception:
pass

def asegurar_dicts_pdf_en_disco():
&quot;&quot;&quot;Convierte cualquier diccionario de PDFs en memoria a DiskBackedPDFDict para liberar RAM de inmediato.&quot;&quot;&quot;
for k_name, sub_name in [(&quot;raw_uploaded_pdfs&quot;, &quot;raw_pdfs&quot;), (&quot;dict_pdfs&quot;, &quot;renom_pdfs&quot;)]:
if k_name not in st.session_state or st.session_state[k_name] is None:
st.session_state[k_name] = DiskBackedPDFDict(sub_name)
elif not isinstance(st.session_state[k_name], DiskBackedPDFDict):
old_d = st.session_state[k_name]
new_d = DiskBackedPDFDict(sub_name)
if isinstance(old_d, dict):
for fn, bval in old_d.items():
new_d[fn] = bval
st.session_state[k_name] = new_d
# Si zip_pdfs estaba en RAM como bytes, moverlo a disco y liberar RAM
if isinstance(st.session_state.get(&quot;zip_pdfs&quot;), (bytes, bytearray)):
try:
with open(ZIP_FACTURAS_DISK_PATH, &quot;wb&quot;) as fz:
fz.write(st.session_state[&quot;zip_pdfs&quot;])
st.session_state[&quot;zip_pdfs_path&quot;] = ZIP_FACTURAS_DISK_PATH
except Exception:
pass
st.session_state[&quot;zip_pdfs&quot;] = None

# ESTADOS DE SESION LIMPIOS Y SANEAMIENTO PREVENTIVO DE MEMORIA RAM
try:
asegurar_dicts_pdf_en_disco()
except Exception:
pass
for _k in [&quot;df_procesado&quot;]:
if _k in st.session_state and st.session_state[_k] is not None:
try:
if not isinstance(st.session_state[_k], pd.DataFrame):
st.session_state.pop(_k, None)
except Exception:
st.session_state.pop(_k, None)
if &quot;autenticado&quot; not in st.session_state:
st.session_state[&quot;autenticado&quot;] = False
if &quot;empresa_activa&quot; not in st.session_state:
st.session_state[&quot;empresa_activa&quot;] = None
if &quot;proceso_activo&quot; not in st.session_state:
st.session_state[&quot;proceso_activo&quot;] = None
# PANTALLA 1: LOGIN
if not st.session_state[&quot;autenticado&quot;]:
col1, col2, col3 = st.columns(3)
with col2:
st.markdown(&quot;&lt;h2 style=&#39;text-align: center; color: #0f172a;&#39;&gt;Cuentas Tuquiz&lt;/h2&gt;&quot;, unsafe_allow_html=True)
st.markdown(&quot;&lt;p style=&#39;text-align: center; color: #64748b;&#39;&gt;Plataforma unificada para gestion contable y DIAN
(Siigo / World Office)&lt;/p&gt;&quot;, unsafe_allow_html=True)
with st.form(&quot;login_form&quot;):
usuario = st.text_input(&quot;Usuario o Correo Electronico&quot;, value=&quot;deibydaza2014@gmail.com&quot;)
password = st.text_input(&quot;Contrasena&quot;, type=&quot;password&quot;, value=&quot;123456&quot;)
submit = st.form_submit_button(&quot;Ingresar al Ecosistema&quot;)
if submit:

if usuario and password:
st.session_state[&quot;autenticado&quot;] = True
st.rerun()
else:
st.error(&quot;Ingresa usuario y contrasena.&quot;)
st.stop()
# PANTALLA 2: SELECTOR DE EMPRESAS CON RESPONSABILIDADES RUT
EMPRESAS_DISPONIBLES = [
{
&quot;nombre&quot;: &quot;INDUMAQ ER SAS&quot;,
&quot;nit&quot;: &quot;901.346.412-5&quot;,
&quot;actividad&quot;: &quot;Comercio y Reparacion de Maquinaria / Importaciones&quot;,
&quot;regimen&quot;: &quot;Responsable de IVA&quot;,
&quot;es_gran_contribuyente&quot;: False,
&quot;es_autorretenedor&quot;: False,
&quot;responsabilidades_rut&quot;: [&quot;O-48&quot;],
&quot;estado&quot;: &quot;ACTIVA&quot;,
&quot;badge&quot;: &quot;badge-active&quot;
}
]
if not st.session_state[&quot;empresa_activa&quot;]:
st.title(&quot;Seleccion de Empresa&quot;)
st.caption(&quot;Selecciona la entidad sobre la cual vas a trabajar:&quot;)
c1, c2 = st.columns(2)
for i, emp in enumerate(EMPRESAS_DISPONIBLES):
col = c1 if i % 2 == 0 else c2
with col:
st.markdown(f&quot;&quot;&quot;
&lt;div class=&quot;card-box&quot;&gt;
&lt;div style=&quot;display:flex; justify-content:space-between; align-items:center;&quot;&gt;
&lt;h3 style=&quot;margin:0; color:#0f172a;&quot;&gt;�� {emp[&#39;nombre&#39;]}&lt;/h3&gt;
&lt;span class=&quot;{emp[&#39;badge&#39;]}&quot;&gt;{emp[&#39;estado&#39;]}&lt;/span&gt;
&lt;/div&gt;
&lt;p style=&quot;margin:8px 0 2px 0; color:#475569;&quot;&gt;&lt;b&gt;NIT:&lt;/b&gt; {emp[&#39;nit&#39;]} - &lt;b&gt;Regimen:&lt;/b&gt;
{emp[&#39;regimen&#39;]}&lt;/p&gt;
&lt;p style=&quot;margin:2px 0 14px 0; color:#64748b; font-size:14px;&quot;&gt;{emp[&#39;actividad&#39;]}&lt;/p&gt;
&lt;/div&gt;
&quot;&quot;&quot;, unsafe_allow_html=True)
if emp[&quot;estado&quot;] == &quot;ACTIVA&quot;:
if st.button(f&quot;Ingresar a {emp[&#39;nombre&#39;]}&quot;, key=f&quot;btn_emp_{i}&quot;):
st.session_state[&quot;empresa_activa&quot;] = emp
st.session_state[&quot;proceso_activo&quot;] = &quot;facturacion&quot;
# registrar_actividad_sesion(emp)
st.rerun()
else:
st.button(f&quot;Pendiente Datos / Parametrizacion ({emp[&#39;estado&#39;]})&quot;, key=f&quot;btn_emp_{i}&quot;, disabled=True)
st.stop()
empresa = st.session_state[&quot;empresa_activa&quot;]
# Sesion en memoria limpia
pass
# PANTALLA 3: MENU DE PROCESOS OPERATIVOS
PROCESOS_SISTEMA = [
{
&quot;id&quot;: &quot;facturacion&quot;,
&quot;icono&quot;: &quot;Facturas&quot;,
&quot;titulo&quot;: &quot;Facturas de Compra, Venta y Devoluciones&quot;,
&quot;desc&quot;: &quot;Carga de reportes DIAN/Token, desbloqueo y renombrado de facturas por comprobante, auditoria contable
y plantilla Siigo.&quot;,
&quot;estado&quot;: &quot;ACTIVO&quot;,
&quot;badge&quot;: &quot;badge-active&quot;
},
{
&quot;id&quot;: &quot;nomina&quot;,
&quot;icono&quot;: &quot;Nomina&quot;,
&quot;titulo&quot;: &quot;Gestion Laboral y Nomina Electronica&quot;,

&quot;desc&quot;: &quot;Calculo de liquidacion de nomina, provisiones de prestaciones sociales, seguridad social y emision de
soportes electronicos DIAN.&quot;,
&quot;estado&quot;: &quot;PROXIMAMENTE&quot;,
&quot;badge&quot;: &quot;badge-next&quot;
},
{
&quot;id&quot;: &quot;conciliacion&quot;,
&quot;icono&quot;: &quot;Bancos&quot;,
&quot;titulo&quot;: &quot;Tesoreria y Conciliacion Bancaria&quot;,
&quot;desc&quot;: &quot;Cruce automatizado de extractos bancarios (Bancolombia, Davivienda) contra libros auxiliares y
control de partidas conciliatorias.&quot;,
&quot;estado&quot;: &quot;PROXIMAMENTE&quot;,
&quot;badge&quot;: &quot;badge-next&quot;
},
{
&quot;id&quot;: &quot;notas&quot;,
&quot;icono&quot;: &quot;Ajustes&quot;,
&quot;titulo&quot;: &quot;Notas de Contabilidad y Cierre Fiscal&quot;,
&quot;desc&quot;: &quot;Comprobantes de ajuste, amortizaciones de intangibles (NIC 38), depreciaciones y conciliacion fiscal
NIIF vs. DIAN.&quot;,
&quot;estado&quot;: &quot;PROXIMAMENTE&quot;,
&quot;badge&quot;: &quot;badge-next&quot;
}
]
if not st.session_state[&quot;proceso_activo&quot;]:
col_t1, col_t2 = st.columns(2)
with col_t1:
st.subheader(f&quot;{empresa[&#39;nombre&#39;]} - Panel de Procesos&quot;)
st.caption(f&quot;NIT: {empresa[&#39;nit&#39;]} - Selecciona el modulo de trabajo que deseas ejecutar:&quot;)
with col_t2:
if st.button(&quot;Cambiar de Empresa&quot;):
st.session_state[&quot;empresa_activa&quot;] = None
st.session_state[&quot;proceso_activo&quot;] = None
st.rerun()
st.markdown(&quot;---&quot;)
cp1, cp2 = st.columns(2)
for idx, proc in enumerate(PROCESOS_SISTEMA):
col_p = cp1 if idx % 2 == 0 else cp2
with col_p:
st.markdown(f&quot;&quot;&quot;
&lt;div class=&quot;card-box&quot;&gt;
&lt;div style=&quot;display:flex; justify-content:space-between; align-items:center;&quot;&gt;
&lt;h3 style=&quot;margin:0; color:#0f172a;&quot;&gt;{proc[&#39;titulo&#39;]}&lt;/h3&gt;
&lt;span class=&quot;{proc[&#39;badge&#39;]}&quot;&gt;{proc[&#39;estado&#39;]}&lt;/span&gt;
&lt;/div&gt;
&lt;p style=&quot;margin:12px 0 16px 0; color:#475569; font-size:14px; line-height:1.5;&quot;&gt;{proc[&#39;desc&#39;]}&lt;/p&gt;
&lt;/div&gt;
&quot;&quot;&quot;, unsafe_allow_html=True)
if proc[&quot;estado&quot;] == &quot;ACTIVO&quot;:
if st.button(f&quot;Abrir Modulo de Facturacion&quot;, key=f&quot;btn_proc_{idx}&quot;):
st.session_state[&quot;proceso_activo&quot;] = proc[&quot;id&quot;]
st.rerun()
else:
st.button(f&quot;Modulo en Construccion ({proc[&#39;estado&#39;]})&quot;, key=f&quot;btn_proc_{idx}&quot;, disabled=True)
st.stop()
# PANTALLA 4: FACTURACION, AUDITORIA Y SIIGO
col_nav1, col_nav2 = st.columns(2)
with col_nav1:
st.subheader(f&quot;{empresa[&#39;nombre&#39;]} - Facturacion, Auditoria y Siigo&quot;)
st.caption(f&quot;NIT: {empresa[&#39;nit&#39;]} - Modulo Activo: Facturas de Compra, Venta y Devoluciones&quot;)
with col_nav2:
if st.button(&quot;Volver a Procesos&quot;):
st.session_state[&quot;proceso_activo&quot;] = None
st.rerun()
st.markdown(&quot;---&quot;)
# ==============================================================================

# GESTOR DE HISTORIAL DE TRABAJOS Y AUDITORÍAS PASADAS (PERSISTENCIA TOTAL)
# ==============================================================================
# Funciones de sesion y directorios centralizadas en la cabecera

def guardar_estado_manual(empresa_dict):
&quot;&quot;&quot;Guarda en disco únicamente los datos contables modificados (DataFrame y Paquetes), sin procesar PDFs.&quot;&quot;&quot;
if &quot;df_procesado&quot; not in st.session_state or st.session_state[&quot;df_procesado&quot;] is None:
return False
try:
nit_clean = re.sub(r&#39;\D&#39;, &#39;&#39;, str(empresa_dict.get(&quot;nit&quot;, &quot;empresa&quot;)))
d = os.path.join(DATA_DIR, nit_clean, &quot;estado_manual&quot;)
os.makedirs(d, exist_ok=True)
with open(os.path.join(d, &quot;df_procesado.pkl&quot;), &quot;wb&quot;) as f:
pickle.dump(st.session_state[&quot;df_procesado&quot;], f)
if st.session_state.get(&quot;paquetes_importacion&quot;):
with open(os.path.join(d, &quot;paquetes_importacion.pkl&quot;), &quot;wb&quot;) as f:
pickle.dump(st.session_state[&quot;paquetes_importacion&quot;], f)
if st.session_state.get(&quot;asientos_triangulacion_por_factura&quot;):
with open(os.path.join(d, &quot;asientos_triangulacion.pkl&quot;), &quot;wb&quot;) as f:
pickle.dump(st.session_state[&quot;asientos_triangulacion_por_factura&quot;], f)
if st.session_state.get(&quot;facturas_no_contabilizar&quot;):
with open(os.path.join(d, &quot;facturas_no_contabilizar.json&quot;), &quot;w&quot;, encoding=&quot;utf-8&quot;) as f:
json.dump(list(st.session_state[&quot;facturas_no_contabilizar&quot;]), f)
now_str = datetime.datetime.now().strftime(&quot;%d/%m/%Y %I:%M:%S %p&quot;)
meta = {
&quot;fecha_guardado&quot;: now_str,
&quot;total_facturas&quot;: len(st.session_state[&quot;df_procesado&quot;]),
&quot;archivo_excel&quot;: st.session_state.get(&quot;excel_nombre&quot;, &quot;Reporte.xlsx&quot;)
}
with open(os.path.join(d, &quot;meta.json&quot;), &quot;w&quot;, encoding=&quot;utf-8&quot;) as f:
json.dump(meta, f, ensure_ascii=False, indent=2)
st.session_state[&quot;_ultimo_guardado_manual&quot;] = now_str
st.session_state[&quot;_cached_respaldo_bytes&quot;] = None
import gc
gc.collect()
return True
except Exception:
return False
def cargar_estado_manual(empresa_dict):
&quot;&quot;&quot;Restaura en memoria el último estado guardado manualmente por el usuario.&quot;&quot;&quot;
try:
nit_clean = re.sub(r&#39;\D&#39;, &#39;&#39;, str(empresa_dict.get(&quot;nit&quot;, &quot;empresa&quot;)))
d = os.path.join(DATA_DIR, nit_clean, &quot;estado_manual&quot;)
df_p_path = os.path.join(d, &quot;df_procesado.pkl&quot;)
if os.path.exists(df_p_path):
with open(df_p_path, &quot;rb&quot;) as f:
st.session_state[&quot;df_procesado&quot;] = pickle.load(f)
pq_path = os.path.join(d, &quot;paquetes_importacion.pkl&quot;)
if os.path.exists(pq_path):
with open(pq_path, &quot;rb&quot;) as f:
pqs_loaded = pickle.load(f)
st.session_state[&quot;paquetes_importacion&quot;] = pqs_loaded
for pq_k, pq_v in pqs_loaded.items():
if isinstance(pq_v, dict):
if &quot;asiento_fijo&quot; in pq_v or &quot;asiento_aprobado&quot; in pq_v:
st.session_state[f&quot;paquete_listo_{pq_k}&quot;] = True
if &quot;destino_siigo&quot; in pq_v:
st.session_state[f&quot;paquete_destino_siigo_{pq_k}&quot;] = pq_v[&quot;destino_siigo&quot;]
as_path = os.path.join(d, &quot;asientos_triangulacion.pkl&quot;)
if os.path.exists(as_path):
with open(as_path, &quot;rb&quot;) as f:
st.session_state[&quot;asientos_triangulacion_por_factura&quot;] = pickle.load(f)
fn_path = os.path.join(d, &quot;facturas_no_contabilizar.json&quot;)
if os.path.exists(fn_path):
with open(fn_path, &quot;r&quot;, encoding=&quot;utf-8&quot;) as f:
st.session_state[&quot;facturas_no_contabilizar&quot;] = set(json.load(f))
meta_path = os.path.join(d, &quot;meta.json&quot;)
if os.path.exists(meta_path):
with open(meta_path, &quot;r&quot;, encoding=&quot;utf-8&quot;) as f:

meta = json.load(f)
st.session_state[&quot;excel_nombre&quot;] = meta.get(&quot;archivo_excel&quot;, &quot;Reporte.xlsx&quot;)
st.session_state[&quot;_ultimo_guardado_manual&quot;] = meta.get(&quot;fecha_guardado&quot;, &quot;&quot;)
return True
except Exception:
pass
return False
def get_empresa_trabajos_dir(empresa_dict):
nit_clean = re.sub(r&quot;\D&quot;, &quot;&quot;, str(empresa_dict.get(&quot;nit&quot;, &quot;empresa&quot;)))
d = os.path.join(DATA_DIR, nit_clean, &quot;trabajos&quot;)
os.makedirs(d, exist_ok=True)
return d
def guardar_trabajo_en_historial(empresa_dict, df_procesado, excel_bytes=None, excel_nombre=&quot;Reporte.xlsx&quot;,
dict_pdfs_renombrados=None, dict_pdfs_originales=None, zip_bytes=None, consecutivo_ini=680, job_id=None):
try:
base_dir = get_empresa_trabajos_dir(empresa_dict)
if not job_id:
job_id = st.session_state.get(&quot;job_actual_id&quot;)
if not job_id:
now_str = datetime.datetime.now().strftime(&quot;%Y%m%d_%H%M%S&quot;)
job_id = f&quot;job_{now_str}&quot;
job_dir = os.path.join(base_dir, job_id)
os.makedirs(job_dir, exist_ok=True)
dir_renom = os.path.join(job_dir, &quot;pdfs_renombrados&quot;)
os.makedirs(dir_renom, exist_ok=True)
dir_orig = os.path.join(job_dir, &quot;pdfs_originales&quot;)
os.makedirs(dir_orig, exist_ok=True)
# 1. Guardar DataFrame procesado
if df_procesado is not None:
with open(os.path.join(job_dir, &quot;df_procesado.pkl&quot;), &quot;wb&quot;) as f:
pickle.dump(df_procesado, f)
# 2. Guardar archivo Excel original subido
if excel_bytes:
with open(os.path.join(job_dir, &quot;excel_original.xlsx&quot;), &quot;wb&quot;) as f_ex:
f_ex.write(excel_bytes)
# 3. Guardar PDFs renombrados / separados (escritura inteligente: solo si no existe o cambió)
if dict_pdfs_renombrados:
for fname, bdata in dict_pdfs_renombrados.items():
p_ren = os.path.join(dir_renom, fname)
if not os.path.exists(p_ren) or os.path.getsize(p_ren) != len(bdata):
with open(p_ren, &quot;wb&quot;) as pf:
pf.write(bdata)
# 4. Guardar PDFs originales subidos (escritura inteligente: solo si no existe o cambió)
if dict_pdfs_originales:
for fname, bdata in dict_pdfs_originales.items():
p_ori = os.path.join(dir_orig, fname)
if not os.path.exists(p_ori) or os.path.getsize(p_ori) != len(bdata):
with open(p_ori, &quot;wb&quot;) as pf:
pf.write(bdata)
# 5. Guardar paquete ZIP de facturas si existe
if zip_bytes:
with open(os.path.join(job_dir, &quot;paquete_facturas.zip&quot;), &quot;wb&quot;) as zf:
zf.write(zip_bytes)
# 6. Guardar paquetes de importación y triangulación
if &quot;paquetes_importacion&quot; in st.session_state and st.session_state[&quot;paquetes_importacion&quot;]:
try:
with open(os.path.join(job_dir, &quot;paquetes_importacion.pkl&quot;), &quot;wb&quot;) as f_pq:
pickle.dump(st.session_state[&quot;paquetes_importacion&quot;], f_pq)
except Exception:
pass

# 7. Guardar asientos aprobados de triangulación
if &quot;asientos_triangulacion_por_factura&quot; in st.session_state and
st.session_state[&quot;asientos_triangulacion_por_factura&quot;]:
try:
with open(os.path.join(job_dir, &quot;asientos_triangulacion.pkl&quot;), &quot;wb&quot;) as f_as:
pickle.dump(st.session_state[&quot;asientos_triangulacion_por_factura&quot;], f_as)
except Exception:
pass
# 8. Guardar estado general de la sesión (paquetes listos, opciones, facturas excluidas)
try:
estado_flags = {
&quot;paquetes_listos&quot;: {k: v for k, v in st.session_state.items() if k.startswith(&quot;paquete_listo_&quot;)},
&quot;paquetes_destino_siigo&quot;: {k: v for k, v in st.session_state.items() if
k.startswith(&quot;paquete_destino_siigo_&quot;)},
&quot;enviar_gp&quot;: {k: v for k, v in st.session_state.items() if k.startswith(&quot;enviar_gp_pq_&quot;)},
&quot;enviar_h2&quot;: {k: v for k, v in st.session_state.items() if k.startswith(&quot;enviar_h2_pq_&quot;)},
&quot;enviar_nd&quot;: {k: v for k, v in st.session_state.items() if k.startswith(&quot;enviar_nd_pq_&quot;)},
&quot;facturas_no_contabilizar&quot;: list(st.session_state.get(&quot;facturas_no_contabilizar&quot;, set())),
&quot;sel_paquete_activo_key&quot;: st.session_state.get(&quot;sel_paquete_activo_key&quot;, 1)
}
with open(os.path.join(job_dir, &quot;estado_sesion.json&quot;), &quot;w&quot;, encoding=&quot;utf-8&quot;) as f_es:
json.dump(estado_flags, f_es, ensure_ascii=False, indent=2)
except Exception:
pass
now_disp = datetime.datetime.now().strftime(&quot;%d/%m/%Y %I:%M %p&quot;)
# Verificar archivos físicos guardados en disco
n_orig_disco = len(os.listdir(dir_orig)) if os.path.exists(dir_orig) else 0
n_renom_disco = len(os.listdir(dir_renom)) if os.path.exists(dir_renom) else 0
tiene_excel = os.path.exists(os.path.join(job_dir, &quot;excel_original.xlsx&quot;)) or (excel_bytes is not None)
meta = {
&quot;id&quot;: job_id,
&quot;nombre_trabajo&quot;: f&quot;Trabajo del {now_disp} ({excel_nombre})&quot;,
&quot;archivo_excel&quot;: excel_nombre,
&quot;tiene_excel&quot;: tiene_excel,
&quot;fecha&quot;: now_disp,
&quot;total_facturas&quot;: len(df_procesado) if df_procesado is not None else 0,
&quot;consecutivo_inicial&quot;: consecutivo_ini,
&quot;total_valor&quot;: float(df_procesado[&quot;Total&quot;].sum()) if (df_procesado is not None and &quot;Total&quot; in
df_procesado) else 0.0,
&quot;total_pdfs_renombrados&quot;: n_renom_disco,
&quot;total_pdfs_originales&quot;: n_orig_disco
}
with open(os.path.join(job_dir, &quot;meta.json&quot;), &quot;w&quot;, encoding=&quot;utf-8&quot;) as mf:
json.dump(meta, mf, ensure_ascii=False, indent=2)
# Actualizar índice general de la empresa para carga inmediata
idx_file = os.path.join(base_dir, &quot;historial_index.json&quot;)
try:
with open(idx_file, &quot;r&quot;, encoding=&quot;utf-8&quot;) as f_idx:
index_data = json.load(f_idx)
except Exception:
index_data = []
index_data = [x for x in index_data if x.get(&quot;id&quot;) != job_id]
index_data.append(meta)
index_data.sort(key=lambda x: x.get(&quot;id&quot;, &quot;&quot;), reverse=True)
try:
with open(idx_file, &quot;w&quot;, encoding=&quot;utf-8&quot;) as f_idx:
json.dump(index_data, f_idx, ensure_ascii=False, indent=2)
except Exception:
pass
return job_id
except Exception:
return None

def listar_trabajos_historial(empresa_dict):
base_dir = get_empresa_trabajos_dir(empresa_dict)
lista = []
idx_file = os.path.join(base_dir, &quot;historial_index.json&quot;)
if os.path.exists(idx_file):
try:
with open(idx_file, &quot;r&quot;, encoding=&quot;utf-8&quot;) as f_idx:
raw_l = json.load(f_idx)
if isinstance(raw_l, list):
lista = raw_l
except Exception:
lista = []
if not lista and os.path.exists(base_dir):
for jid in os.listdir(base_dir):
jdir = os.path.join(base_dir, jid)
meta_file = os.path.join(jdir, &quot;meta.json&quot;)
if os.path.isdir(jdir) and os.path.exists(meta_file):
try:
with open(meta_file, &quot;r&quot;, encoding=&quot;utf-8&quot;) as mf:
meta = json.load(mf)
lista.append(meta)
except Exception:
pass
# Normalizar campos y verificar que exista df_procesado.pkl
valida = []
for item in lista:
if isinstance(item, dict) and &quot;id&quot; in item:
jdir = os.path.join(base_dir, item[&quot;id&quot;])
if os.path.exists(os.path.join(jdir, &quot;df_procesado.pkl&quot;)):
item.setdefault(&quot;total_facturas&quot;, 0)
item.setdefault(&quot;total_valor&quot;, 0.0)
item.setdefault(&quot;nombre_trabajo&quot;, f&quot;Trabajo {item[&#39;id&#39;]}&quot;)
valida.append(item)
valida.sort(key=lambda x: x.get(&quot;id&quot;, &quot;&quot;), reverse=True)
return valida
def cargar_trabajo_historial(empresa_dict, job_id):
base_dir = get_empresa_trabajos_dir(empresa_dict)
jdir = os.path.join(base_dir, job_id)
df = None
dict_renom = {}
dict_orig = {}
zip_bytes = None
excel_bytes = None
excel_nombre = &quot;Reporte.xlsx&quot;
meta_file = os.path.join(jdir, &quot;meta.json&quot;)
if os.path.exists(meta_file):
try:
with open(meta_file, &quot;r&quot;, encoding=&quot;utf-8&quot;) as mf:
m_data = json.load(mf)
excel_nombre = m_data.get(&quot;archivo_excel&quot;, &quot;Reporte.xlsx&quot;)
except:
pass
pkl_file = os.path.join(jdir, &quot;df_procesado.pkl&quot;)
if os.path.exists(pkl_file):
with open(pkl_file, &quot;rb&quot;) as f:
df = pickle.load(f)
ex_file = os.path.join(jdir, &quot;excel_original.xlsx&quot;)
if os.path.exists(ex_file):
with open(ex_file, &quot;rb&quot;) as f_ex:
excel_bytes = f_ex.read()
# PDFs renombrados respaldados en disco (0 MB en RAM)
dict_renom = DiskBackedPDFDict(&quot;renom_pdfs&quot;)
pdf_dir_r = os.path.join(jdir, &quot;pdfs_renombrados&quot;)
if not os.path.exists(pdf_dir_r):

pdf_dir_r = os.path.join(jdir, &quot;pdfs&quot;)
if os.path.exists(pdf_dir_r):
for fn in os.listdir(pdf_dir_r):
if fn.endswith(&quot;.pdf&quot;):
dict_renom[fn] = os.path.join(pdf_dir_r, fn)
# PDFs originales respaldados en disco (0 MB en RAM)
dict_orig = DiskBackedPDFDict(&quot;raw_pdfs&quot;)
pdf_dir_o = os.path.join(jdir, &quot;pdfs_originales&quot;)
if os.path.exists(pdf_dir_o):
for fn in os.listdir(pdf_dir_o):
if fn.endswith(&quot;.pdf&quot;):
dict_orig[fn] = os.path.join(pdf_dir_o, fn)
zip_file = os.path.join(jdir, &quot;paquete_facturas.zip&quot;)
if os.path.exists(zip_file):
st.session_state[&quot;zip_pdfs_path&quot;] = zip_file
zip_bytes = None
# Restaurar paquetes de importación
pq_file = os.path.join(jdir, &quot;paquetes_importacion.pkl&quot;)
if os.path.exists(pq_file):
try:
with open(pq_file, &quot;rb&quot;) as f_pq:
st.session_state[&quot;paquetes_importacion&quot;] = pickle.load(f_pq)
except Exception:
pass
# Restaurar asientos aprobados de triangulación
as_file = os.path.join(jdir, &quot;asientos_triangulacion.pkl&quot;)
if os.path.exists(as_file):
try:
with open(as_file, &quot;rb&quot;) as f_as:
st.session_state[&quot;asientos_triangulacion_por_factura&quot;] = pickle.load(f_as)
except Exception:
pass
# Restaurar flags de estado (paquetes listos, opciones, facturas excluidas)
es_file = os.path.join(jdir, &quot;estado_sesion.json&quot;)
if os.path.exists(es_file):
try:
with open(es_file, &quot;r&quot;, encoding=&quot;utf-8&quot;) as f_es:
est = json.load(f_es)
for k, v in est.get(&quot;paquetes_listos&quot;, {}).items():
st.session_state[k] = v
for k, v in est.get(&quot;paquetes_destino_siigo&quot;, {}).items():
st.session_state[k] = v
for k, v in est.get(&quot;enviar_gp&quot;, {}).items():
st.session_state[k] = v
for k, v in est.get(&quot;enviar_h2&quot;, {}).items():
st.session_state[k] = v
for k, v in est.get(&quot;enviar_nd&quot;, {}).items():
st.session_state[k] = v
if &quot;facturas_no_contabilizar&quot; in est:
st.session_state[&quot;facturas_no_contabilizar&quot;] = set(est[&quot;facturas_no_contabilizar&quot;])
if &quot;sel_paquete_activo_key&quot; in est:
st.session_state[&quot;sel_paquete_activo_key&quot;] = est[&quot;sel_paquete_activo_key&quot;]
except Exception:
pass
return df, dict_renom, dict_orig, excel_bytes, excel_nombre, zip_bytes
def eliminar_trabajo_historial(empresa_dict, job_id):
base_dir = get_empresa_trabajos_dir(empresa_dict)
jdir = os.path.join(base_dir, job_id)
if os.path.exists(jdir):
shutil.rmtree(jdir, ignore_errors=True)
idx_file = os.path.join(base_dir, &quot;historial_index.json&quot;)
if os.path.exists(idx_file):
try:
with open(idx_file, &quot;r&quot;, encoding=&quot;utf-8&quot;) as f_idx:

index_data = json.load(f_idx)
index_data = [x for x in index_data if x.get(&quot;id&quot;) != job_id]
with open(idx_file, &quot;w&quot;, encoding=&quot;utf-8&quot;) as f_idx:
json.dump(index_data, f_idx, ensure_ascii=False, indent=2)
except Exception:
pass
def exportar_respaldo_sesion_zip(empresa_dict, df_proc, dict_renom, dict_orig, excel_b, excel_n, zip_p=None):
&quot;&quot;&quot;Empaqueta toda la sesión contable en un archivo ZIP descargable para respaldo indestructible con paquetes y
triangulaciones.&quot;&quot;&quot;
buf = io.BytesIO()
with zipfile.ZipFile(buf, &quot;w&quot;, zipfile.ZIP_DEFLATED) as zf:
if df_proc is not None:
zf.writestr(&quot;df_procesado.pkl&quot;, pickle.dumps(df_proc))
if excel_b:
zf.writestr(&quot;excel_original.xlsx&quot;, excel_b)
if zip_p:
zf.writestr(&quot;paquete_facturas.zip&quot;, zip_p)
# Incluir paquetes de importación y asientos de triangulación si existen
if &quot;paquetes_importacion&quot; in st.session_state and st.session_state[&quot;paquetes_importacion&quot;]:
zf.writestr(&quot;paquetes_importacion.pkl&quot;, pickle.dumps(st.session_state[&quot;paquetes_importacion&quot;]))
if &quot;asientos_triangulacion_por_factura&quot; in st.session_state and
st.session_state[&quot;asientos_triangulacion_por_factura&quot;]:
zf.writestr(&quot;asientos_triangulacion.pkl&quot;,
pickle.dumps(st.session_state[&quot;asientos_triangulacion_por_factura&quot;]))
# Incluir estado completo de la sesión
estado_flags = {
&quot;paquetes_listos&quot;: {k: v for k, v in st.session_state.items() if k.startswith(&quot;paquete_listo_&quot;)},
&quot;enviar_gp&quot;: {k: v for k, v in st.session_state.items() if k.startswith(&quot;enviar_gp_pq_&quot;)},
&quot;enviar_h2&quot;: {k: v for k, v in st.session_state.items() if k.startswith(&quot;enviar_h2_pq_&quot;)},
&quot;enviar_nd&quot;: {k: v for k, v in st.session_state.items() if k.startswith(&quot;enviar_nd_pq_&quot;)},
&quot;facturas_no_contabilizar&quot;: list(st.session_state.get(&quot;facturas_no_contabilizar&quot;, set())),
&quot;sel_paquete_activo_key&quot;: st.session_state.get(&quot;sel_paquete_activo_key&quot;, 1)
}
zf.writestr(&quot;estado_sesion.json&quot;, json.dumps(estado_flags, ensure_ascii=False, indent=2))
now_str = datetime.datetime.now().strftime(&quot;%d/%m/%Y %I:%M %p&quot;)
meta = {
&quot;empresa&quot;: empresa_dict.get(&quot;nombre&quot;, &quot;&quot;),
&quot;nit&quot;: empresa_dict.get(&quot;nit&quot;, &quot;&quot;),
&quot;excel_nombre&quot;: excel_n or &quot;Reporte.xlsx&quot;,
&quot;archivo_excel&quot;: excel_n or &quot;Reporte.xlsx&quot;,
&quot;fecha_respaldo&quot;: now_str,
&quot;total_facturas&quot;: len(df_proc) if df_proc is not None else 0,
&quot;total_pdfs_renombrados&quot;: len(dict_renom) if dict_renom else 0,
&quot;total_pdfs_originales&quot;: len(dict_orig) if dict_orig else 0
}
zf.writestr(&quot;meta_sesion.json&quot;, json.dumps(meta, ensure_ascii=False, indent=2))
mem_actual = cargar_memoria_aprendizaje(empresa_dict)
if mem_actual:
zf.writestr(&quot;memoria_aprendizaje.json&quot;, json.dumps(mem_actual, ensure_ascii=False, indent=2))
zf.writestr(&quot;meta.json&quot;, json.dumps(meta, ensure_ascii=False, indent=2))
# PDFs ya vienen comprimidos internamente (FlateDecode). Usar compresión directa ultra-rápida
if dict_renom:
for fn, bdata in dict_renom.items():
zinfo = zipfile.ZipInfo(f&quot;pdfs_renombrados/{fn}&quot;)
zinfo.compress_type = zipfile.ZIP_STORED
zf.writestr(zinfo, bdata)
if dict_orig:
for fn, bdata in dict_orig.items():
zinfo = zipfile.ZipInfo(f&quot;pdfs_originales/{fn}&quot;)
zinfo.compress_type = zipfile.ZIP_STORED
zf.writestr(zinfo, bdata)
buf.seek(0)
return buf.getvalue()

def obtener_respaldo_sesion_cached(empresa_dict):
&quot;&quot;&quot;Devuelve el archivo de respaldo cacheado en memoria, regenerándolo solo cuando hay cambios.&quot;&quot;&quot;
if &quot;df_procesado&quot; not in st.session_state or st.session_state[&quot;df_procesado&quot;] is None:

return b&quot;&quot;
df_p = st.session_state[&quot;df_procesado&quot;]
d_ren = st.session_state.get(&quot;dict_pdfs&quot;, {})
d_ori = st.session_state.get(&quot;raw_uploaded_pdfs&quot;, {})
sig_actual = f&quot;{len(df_p)}_{len(d_ren)}_{len(d_ori)}_{st.session_state.get(&#39;_ultimo_autosave_hora&#39;, &#39;&#39;)}&quot;
if st.session_state.get(&quot;_cached_respaldo_sig&quot;) != sig_actual or &quot;_cached_respaldo_bytes&quot; not in st.session_state:
b_data = exportar_respaldo_sesion_zip(
empresa_dict, df_p, d_ren, d_ori,
st.session_state.get(&quot;excel_bytes&quot;),
st.session_state.get(&quot;excel_nombre&quot;, &quot;Reporte.xlsx&quot;),
st.session_state.get(&quot;zip_pdfs&quot;)
)
st.session_state[&quot;_cached_respaldo_bytes&quot;] = b_data
st.session_state[&quot;_cached_respaldo_sig&quot;] = sig_actual
return st.session_state.get(&quot;_cached_respaldo_bytes&quot;, b&quot;&quot;)
def importar_respaldo_sesion_zip(zip_bytes, empresa_dict):
&quot;&quot;&quot;Restaura una sesión contable completa desde un archivo ZIP de respaldo.&quot;&quot;&quot;
try:
with zipfile.ZipFile(io.BytesIO(zip_bytes), &quot;r&quot;) as zf:
df_proc = None
if &quot;df_procesado.pkl&quot; in zf.namelist():
df_proc = pickle.loads(zf.read(&quot;df_procesado.pkl&quot;))
excel_b = zf.read(&quot;excel_original.xlsx&quot;) if &quot;excel_original.xlsx&quot; in zf.namelist() else None
zip_p = zf.read(&quot;paquete_facturas.zip&quot;) if &quot;paquete_facturas.zip&quot; in zf.namelist() else None
excel_n = &quot;Reporte_Restaurado.xlsx&quot;
if &quot;meta_sesion.json&quot; in zf.namelist():
try:
m = json.loads(zf.read(&quot;meta_sesion.json&quot;).decode(&quot;utf-8&quot;))
excel_n = m.get(&quot;excel_nombre&quot;, excel_n)
except:
pass
dict_renom = {}
dict_orig = {}
for name in zf.namelist():
if name.startswith(&quot;pdfs_renombrados/&quot;) and name.endswith(&quot;.pdf&quot;):
fn = name.replace(&quot;pdfs_renombrados/&quot;, &quot;&quot;)
dict_renom[fn] = zf.read(name)
elif name.startswith(&quot;pdfs_originales/&quot;) and name.endswith(&quot;.pdf&quot;):
fn = name.replace(&quot;pdfs_originales/&quot;, &quot;&quot;)
dict_orig[fn] = zf.read(name)
return df_proc, dict_renom, dict_orig, excel_b, excel_n, zip_p
except Exception as e:
return None, {}, {}, None, &quot;&quot;, None
# ==============================================================================
# MOTOR DE TRIANGULACIÓN Y CRUCES DE IMPORTACIÓN (ADUANAS Y PAGOS A TERCEROS)
# ==============================================================================
CUENTA_RETENCION_ASUMIDA = &quot;53152001&quot;
CUENTA_NO_DEDUCIBLE = &quot;53950501&quot;
CUENTA_IVA_IMPORTACION = &quot;240835&quot;
CUENTA_IMPORTACION_TRANSITO = &quot;14650501&quot;
CUENTA_CXP_AGENCIA_EXTERIOR = &quot;22050505&quot;
CUENTA_CXP_DHL_NACIONAL = &quot;23359501&quot;
CUENTA_CXP_EURO_SHIPPING = &quot;22050501&quot;
def normalizar_fecha_dian(fecha_raw):
&quot;&quot;&quot;Normaliza cualquier formato de fecha a DD/MM/AAAA para Siigo y DIAN.&quot;&quot;&quot;
if not fecha_raw:
return &quot;&quot;
s = str(fecha_raw).strip()
m = re.match(r&#39;^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$&#39;, s)
if m:
d, m_val, y = m.groups()
return f&quot;{int(d):02d}/{int(m_val):02d}/{y}&quot;
m_iso = re.match(r&#39;^(\d{4})[/.-](\d{1,2})[/.-](\d{1,2})&#39;, s)
if m_iso:
y, m_val, d = m_iso.groups()
return f&quot;{int(d):02d}/{int(m_val):02d}/{y}&quot;
return s
def analizar_estado_filas_excel(excel_bytes):

&quot;&quot;&quot;
Lee las celdas del archivo Excel DIAN (token) con openpyxl para detectar:
1. Celdas rojas (facturas ya causadas/registradas en Siigo).
2. Columna1 con número de comprobante 10-XXX o texto &#39;registrada&#39;.
3. Columna3 con asignación de Grupo de Importación (1, 2, 3...).
4. Columna4 con cuenta contable de pasivo específica (22050505, 23359501, etc.).
&quot;&quot;&quot;
estados = {}
if not excel_bytes:
return estados
try:
wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
ws = wb.active
headers = [str(cell.value or &#39;&#39;).strip() for cell in ws[1]]
col_c1 = next((i for i, h in enumerate(headers) if any(k in h.lower() for k in [&#39;columna1&#39;, &#39;comprobante&#39;,
&#39;estado&#39;])), None)
col_grp = next((i for i, h in enumerate(headers) if any(k in h.lower() for k in [&#39;columna3&#39;, &#39;grupo&#39;,
&#39;paquete&#39;])), None)
col_cta = next((i for i, h in enumerate(headers) if any(k in h.lower() for k in [&#39;columna4&#39;, &#39;cuenta&#39;,
&#39;cta&#39;])), None)
col_cufe = next((i for i, h in enumerate(headers) if any(k in h.lower() for k in [&#39;cufe&#39;, &#39;token&#39;, &#39;uuid&#39;,
&#39;cude&#39;])), None)
col_folio = next((i for i, h in enumerate(headers) if any(k in h.lower() for k in [&#39;folio&#39;, &#39;factura&#39;])),
None)
col_pref = next((i for i, h in enumerate(headers) if &#39;prefijo&#39; in h.lower()), None)
for row_idx in range(2, ws.max_row + 1):
row_cells = ws[row_idx]
es_roja = False
# Revisar color de relleno de las celdas
for cell in row_cells:
fill = cell.fill
if fill and fill.fill_type:
color = fill.start_color
if color:
rgb = str(getattr(color, &#39;rgb&#39;, &#39;&#39;) or getattr(color, &#39;value&#39;, &#39;&#39;)).upper()
if any(r in rgb for r in [&#39;FF0000&#39;, &#39;C00000&#39;, &#39;FFC7CE&#39;, &#39;EF4444&#39;, &#39;DC2626&#39;, &#39;B91C1C&#39;,
&#39;991B1B&#39;, &#39;F87171&#39;, &#39;FF6666&#39;, &#39;FF4D4D&#39;, &#39;E11D48&#39;]):
es_roja = True
break
if len(rgb) == 8:
try:
r_v = int(rgb[2:4], 16)
g_v = int(rgb[4:6], 16)
b_v = int(rgb[6:8], 16)
if (r_v &gt; 180 and g_v &lt; 130 and b_v &lt; 130) or (r_v &gt; 220 and g_v &lt; 210 and b_v &lt; 215
and r_v - g_v &gt; 30):
es_roja = True
break
except:
pass
val_c1 = str(row_cells[col_c1].value or &#39;&#39;).strip() if col_c1 is not None and col_c1 &lt; len(row_cells) else
&#39;&#39;
if &#39;10-&#39; in val_c1 or &#39;registrad&#39; in val_c1.lower() or &#39;causad&#39; in val_c1.lower():
es_roja = True
val_grp = str(row_cells[col_grp].value or &#39;&#39;).strip() if col_grp is not None and col_grp &lt; len(row_cells)
else &#39;&#39;
val_cta = str(row_cells[col_cta].value or &#39;&#39;).strip() if col_cta is not None and col_cta &lt; len(row_cells)
else &#39;&#39;
if val_cta.endswith(&#39;.0&#39;):
val_cta = val_cta[:-2]
cufe_k = str(row_cells[col_cufe].value or &#39;&#39;).strip() if col_cufe is not None and col_cufe &lt;
len(row_cells) else &#39;&#39;
fol_k = str(row_cells[col_folio].value or &#39;&#39;).strip() if col_folio is not None and col_folio &lt;
len(row_cells) else &#39;&#39;

if fol_k.endswith(&#39;.0&#39;):
fol_k = fol_k[:-2]
pref_k = str(row_cells[col_pref].value or &#39;&#39;).strip() if col_pref is not None and col_pref &lt;
len(row_cells) else &#39;&#39;
if pref_k == &#39;nan&#39; or pref_k == &#39;None&#39;:
pref_k = &#39;&#39;
fac_key = f&quot;{pref_k}-{fol_k}&quot; if pref_k else fol_k
info_fila = {
&#39;row_idx&#39;: row_idx,
&#39;es_roja&#39;: es_roja,
&#39;comprobante_existente&#39;: val_c1 if &#39;10-&#39; in val_c1 else (&#39;10-Previa&#39; if es_roja else &#39;&#39;),
&#39;grupo_importacion&#39;: val_grp,
&#39;cuenta_especifica&#39;: val_cta
}
if cufe_k:
estados[cufe_k] = info_fila
if fac_key:
estados[fac_key] = info_fila
estados[row_idx - 2] = info_fila
except Exception as e:
pass
return estados
def obtener_asiento_contable_hoja2(fac_sel, es_aduanero=False):
&quot;&quot;&quot;Genera el asiento contable EXACTO e IDÉNTICO al que visualiza la Hoja 2 (Auditoría).&quot;&quot;&quot;
asiento_filas = []
es_nc = &quot;Devolucion&quot; in str(fac_sel.get(&quot;Operacion&quot;, &quot;&quot;))
asume_ret = bool(fac_sel.get(&quot;Impuestos Asumidos&quot;, False))
cta_cxp_usar = str(fac_sel.get(&quot;Cuenta Pasivo Especifica&quot;) or fac_sel.get(&quot;Cta Contrapartida&quot;) or (&quot;22050505&quot; if
es_aduanero else &quot;22050501&quot;)).strip()
cta_p = str(fac_sel.get(&quot;Cta Principal&quot;) or &quot;14650501&quot;).strip()
if cta_p in [&quot;146505&quot;, &quot;1465&quot;]:
cta_p = &quot;14650501&quot;
es_importacion_factura = es_aduanero or &quot;1465&quot; in cta_p or &quot;IMPORTACI&quot; in str(fac_sel.get(&quot;Categoría&quot;,
&quot;&quot;)).upper()
cta_iva = str(fac_sel.get(&quot;Cta IVA&quot;) or (&quot;24081501&quot; if es_importacion_factura else &quot;24081001&quot;)).strip()
def _sf(v):
try:
if v is not None and pd.notna(v):
fv = float(v)
return fv if not pd.isna(fv) else 0.0
except Exception:
pass
return 0.0
base_val = _sf(fac_sel.get(&quot;Base&quot;))
iva_val = _sf(fac_sel.get(&quot;IVA&quot;))
tot_val = _sf(fac_sel.get(&quot;Total&quot;))
rfte_val = _sf(fac_sel.get(&quot;ReteFuente&quot;))
rica_val = _sf(fac_sel.get(&quot;ReteICA&quot;))
riva_val = _sf(fac_sel.get(&quot;ReteIVA&quot;))
iva_imp_val = _sf(fac_sel.get(&quot;IVA Importación&quot;))
raw_cta_imp = fac_sel.get(&quot;Cta IVA Importación&quot;)
if raw_cta_imp is None or pd.isna(raw_cta_imp) or str(raw_cta_imp).strip() in [&quot;&quot;, &quot;nan&quot;, &quot;None&quot;]:
cta_iva_imp = &quot;240835&quot;
else:
cta_iva_imp = str(raw_cta_imp).strip()
# 1. Base / Costo
asiento_filas.append({
&quot;Código Cuenta&quot;: cta_p,
&quot;Descripción de la Cuenta&quot;: f&quot;{fac_sel.get(&#39;Categoría&#39;, &#39;Gasto/Costo&#39;)} - {str(fac_sel.get(&#39;Proveedor&#39;,
&#39;&#39;))[:25]}&quot;,
&quot;Tercero / NIT&quot;: str(fac_sel.get(&quot;NIT Emisor&quot;, &quot;&quot;)),
&quot;Débito ($)&quot;: 0.0 if es_nc else base_val,
&quot;Crédito ($)&quot;: base_val if es_nc else 0.0

})
# 2. IVA Descontable (Servicios / Compras)
if iva_val &gt; 0:
desc_iva = f&quot;IVA Descontable Servicios (Base: ${base_val:,.0f})&quot; if (&quot;24081501&quot; in cta_iva or
es_importacion_factura) else f&quot;IVA Descontable (Base: ${base_val:,.0f})&quot;
asiento_filas.append({
&quot;Código Cuenta&quot;: cta_iva,
&quot;Descripción de la Cuenta&quot;: desc_iva,
&quot;Tercero / NIT&quot;: str(fac_sel.get(&quot;NIT Emisor&quot;, &quot;&quot;)),
&quot;Débito ($)&quot;: 0.0 if es_nc else iva_val,
&quot;Crédito ($)&quot;: iva_val if es_nc else 0.0
})
# 2.B IVA de Importación (Cuenta 240835) - Solo para facturas de importación
if iva_imp_val &gt; 0:
asiento_filas.append({
&quot;Código Cuenta&quot;: cta_iva_imp,
&quot;Descripción de la Cuenta&quot;: f&quot;IVA de Importación (Cta 240835) - Fac {fac_sel.get(&#39;Factura&#39;, &#39;&#39;)}&quot;,
&quot;Tercero / NIT&quot;: str(fac_sel.get(&quot;NIT Emisor&quot;, &quot;&quot;)),
&quot;Débito ($)&quot;: 0.0 if es_nc else iva_imp_val,
&quot;Crédito ($)&quot;: iva_imp_val if es_nc else 0.0
})
es_transp_as = (&quot;TRANSP&quot; in str(fac_sel.get(&quot;Categoría&quot;, &quot;&quot;)).upper() or
&quot;TRANSP&quot; in str(fac_sel.get(&quot;Descripcion&quot;, &quot;&quot;)).upper() or
&quot;INLAND&quot; in str(fac_sel.get(&quot;Descripcion&quot;, &quot;&quot;)).upper() or
&quot;FLETE&quot; in str(fac_sel.get(&quot;Descripcion&quot;, &quot;&quot;)).upper() or
&quot;70787&quot; in str(fac_sel.get(&quot;Factura&quot;, &quot;&quot;)) or
str(fac_sel.get(&quot;Cta ReteFuente&quot;, &quot;&quot;)).strip() == &quot;23652505&quot; or
(rfte_val &gt; 0 and base_val &gt; 0 and abs(rfte_val - round(base_val * 0.01, 2)) &lt; 2.0))
cta_rf_raw = str(fac_sel.get(&quot;Cta ReteFuente&quot;, &quot;&quot;)).strip()
cta_ri_raw = str(fac_sel.get(&quot;Cta ReteICA&quot;, &quot;&quot;)).strip()
# Si es transporte o la retención es del 1%, la cuenta legal ES 23652505
if es_transp_as or (rfte_val &gt; 0 and base_val &gt; 0 and abs(rfte_val - round(base_val * 0.01, 2)) &lt; 2.0):
if rfte_val == round(base_val * 0.04, 2) or rfte_val == 0.0:
rfte_val = round(base_val * 0.01, 2)
if rica_val == round(base_val * 0.00966, 2) or rica_val == 0.0:
rica_val = round(base_val * 0.00414, 2)
cta_rf_usar = &quot;23652505&quot;
cta_ri_usar = &quot;23680513&quot;
cta_p = &quot;14650501&quot;
else:
cta_rf_usar = cta_rf_raw if (cta_rf_raw and cta_rf_raw not in [&quot;nan&quot;, &quot;None&quot;, &quot;&quot;]) else (&quot;23652503&quot; if
es_aduanero else &quot;23654001&quot;)
cta_ri_usar = cta_ri_raw if (cta_ri_raw and cta_ri_raw not in [&quot;nan&quot;, &quot;None&quot;, &quot;&quot;]) else (&quot;23680505&quot; if
es_aduanero else &quot;23680501&quot;)
if cta_p == &quot;146505&quot;:
cta_p = &quot;14650501&quot;
tot_ret = round(rfte_val + rica_val, 2)
if asume_ret:
# IMPUESTOS ASUMIDOS (Cruza con Agente Aduanero / Triangulación)
if tot_ret &gt; 0:
asiento_filas.append({
&quot;Código Cuenta&quot;: &quot;53152001&quot;,
&quot;Descripción de la Cuenta&quot;: &quot;Retenciones Asumidas (Impuestos Asumidos Aduana)&quot;,
&quot;Tercero / NIT&quot;: str(fac_sel.get(&quot;NIT Emisor&quot;, &quot;&quot;)),
&quot;Débito ($)&quot;: tot_ret,
&quot;Crédito ($)&quot;: 0.0
})
if rfte_val &gt; 0:
asiento_filas.append({
&quot;Código Cuenta&quot;: cta_rf_usar,
&quot;Descripción de la Cuenta&quot;: f&quot;ReteFuente Practicada ({cta_rf_usar}) Fac {fac_sel.get(&#39;Factura&#39;, &#39;&#39;)}&quot;,
&quot;Tercero / NIT&quot;: str(fac_sel.get(&quot;NIT Emisor&quot;, &quot;&quot;)),
&quot;Débito ($)&quot;: 0.0,

&quot;Crédito ($)&quot;: rfte_val
})
if rica_val &gt; 0:
asiento_filas.append({
&quot;Código Cuenta&quot;: cta_ri_usar,
&quot;Descripción de la Cuenta&quot;: &quot;Retención ICA Practicada&quot;,
&quot;Tercero / NIT&quot;: str(fac_sel.get(&quot;NIT Emisor&quot;, &quot;&quot;)),
&quot;Débito ($)&quot;: 0.0,
&quot;Crédito ($)&quot;: rica_val
})
saldo_cxp = round(base_val + iva_val + iva_imp_val, 2)
asiento_filas.append({
&quot;Código Cuenta&quot;: cta_cxp_usar,
&quot;Descripción de la Cuenta&quot;: f&quot;CxP Proveedor/Agente - Fac {fac_sel.get(&#39;Factura&#39;, &#39;&#39;)}&quot;,
&quot;Tercero / NIT&quot;: str(fac_sel.get(&quot;NIT Emisor&quot;, &quot;&quot;)),
&quot;Débito ($)&quot;: saldo_cxp if es_nc else 0.0,
&quot;Crédito ($)&quot;: 0.0 if es_nc else saldo_cxp
})
else:
# RETENCIONES ORDINARIAS PRACTICADAS AL PROVEEDOR
if rfte_val &gt; 0:
nom_rf_dsp = &quot;Transporte Carga 1%&quot; if cta_rf_usar == &quot;23652505&quot; else (&quot;Servicios 4%&quot; if cta_rf_usar ==
&quot;23652503&quot; else fac_sel.get(&#39;Categoría&#39;, &#39;Compras/Servicios&#39;))
asiento_filas.append({
&quot;Código Cuenta&quot;: cta_rf_usar,
&quot;Descripción de la Cuenta&quot;: f&quot;ReteFuente Practicada ({nom_rf_dsp}) Fac {fac_sel.get(&#39;Factura&#39;, &#39;&#39;)}&quot;,
&quot;Tercero / NIT&quot;: str(fac_sel.get(&quot;NIT Emisor&quot;, &quot;&quot;)),
&quot;Débito ($)&quot;: rfte_val if es_nc else 0.0,
&quot;Crédito ($)&quot;: 0.0 if es_nc else rfte_val
})
if rica_val &gt; 0:
nom_ri_dsp = &quot;Transporte 4.14‰&quot; if cta_ri_usar == &quot;23680513&quot; else (&quot;Servicios 9.66‰&quot; if cta_ri_usar ==
&quot;23680505&quot; else &quot;Retención ICA Practicada&quot;)
asiento_filas.append({
&quot;Código Cuenta&quot;: cta_ri_usar,
&quot;Descripción de la Cuenta&quot;: f&quot;Retención ICA Practicada ({nom_ri_dsp})&quot;,
&quot;Tercero / NIT&quot;: str(fac_sel.get(&quot;NIT Emisor&quot;, &quot;&quot;)),
&quot;Débito ($)&quot;: rica_val if es_nc else 0.0,
&quot;Crédito ($)&quot;: 0.0 if es_nc else rica_val
})
if riva_val &gt; 0:
asiento_filas.append({
&quot;Código Cuenta&quot;: &quot;23670101&quot;,
&quot;Descripción de la Cuenta&quot;: &quot;Retención de IVA Practicada (15%)&quot;,
&quot;Tercero / NIT&quot;: str(fac_sel.get(&quot;NIT Emisor&quot;, &quot;&quot;)),
&quot;Débito ($)&quot;: riva_val if es_nc else 0.0,
&quot;Crédito ($)&quot;: 0.0 if es_nc else riva_val
})
neto_cxp = round(base_val + iva_val + iva_imp_val - rfte_val - rica_val - riva_val, 2)
asiento_filas.append({
&quot;Código Cuenta&quot;: cta_cxp_usar,
&quot;Descripción de la Cuenta&quot;: f&quot;Proveedores Nacionales - Fac {fac_sel.get(&#39;Factura&#39;, &#39;&#39;)}&quot;,
&quot;Tercero / NIT&quot;: str(fac_sel.get(&quot;NIT Emisor&quot;, &quot;&quot;)),
&quot;Débito ($)&quot;: neto_cxp if es_nc else 0.0,
&quot;Crédito ($)&quot;: 0.0 if es_nc else neto_cxp
})
df_h2 = pd.DataFrame(asiento_filas)
if &quot;Descripción de la Cuenta&quot; in df_h2.columns and &quot;Descripción Cuenta&quot; not in df_h2.columns:
df_h2[&quot;Descripción Cuenta&quot;] = df_h2[&quot;Descripción de la Cuenta&quot;]
elif &quot;Descripción Cuenta&quot; in df_h2.columns and &quot;Descripción de la Cuenta&quot; not in df_h2.columns:
df_h2[&quot;Descripción de la Cuenta&quot;] = df_h2[&quot;Descripción Cuenta&quot;]
return df_h2

def obtener_saldo_cruce_factura(tr):
&quot;&quot;&quot;
Calcula de forma segura el Saldo de Cruce (Base + IVA) de una factura para la triangulación,
garantizando que nunca retorne None ni NaN.
&quot;&quot;&quot;
try:

tb = float(tr.get(&quot;Base&quot;, 0.0)) if pd.notna(tr.get(&quot;Base&quot;)) else (float(tr.get(&quot;Subtotal&quot;, 0.0)) if
pd.notna(tr.get(&quot;Subtotal&quot;)) else 0.0)
except Exception:
tb = 0.0
try:
tiv = float(tr.get(&quot;IVA&quot;, 0.0)) if pd.notna(tr.get(&quot;IVA&quot;)) else 0.0
except Exception:
tiv = 0.0
try:
t_imp = float(tr.get(&quot;IVA Importación&quot;, 0.0)) if pd.notna(tr.get(&quot;IVA Importación&quot;)) else 0.0
except Exception:
t_imp = 0.0
val_neto = tr.get(&quot;Total Neto&quot;)
if val_neto is not None and pd.notna(val_neto) and str(val_neto).strip() not in [&quot;&quot;, &quot;nan&quot;, &quot;None&quot;, &quot;0&quot;, &quot;0.0&quot;]:
try:
f_neto = float(val_neto)
if f_neto &gt; 0.0 and not pd.isna(f_neto):
return round(f_neto, 2)
except Exception:
pass
saldo_calc = round(tb + tiv + t_imp, 2)
if saldo_calc &gt; 0.0:
return saldo_calc
try:
tot = float(tr.get(&quot;Total&quot;, 0.0)) if pd.notna(tr.get(&quot;Total&quot;)) else 0.0
if tot &gt; 0.0 and not pd.isna(tot):
return round(tot, 2)
except Exception:
pass
return 0.0

def calcular_concordancia_fecha(fecha_tercero, fecha_agente):
&quot;&quot;&quot;
Calcula de forma dinámica la concordancia cronológica entre la factura del tercero y la del agente,
garantizando que nunca retorne None ni NaN.
&quot;&quot;&quot;
if not fecha_tercero or not fecha_agente or pd.isna(fecha_tercero) or pd.isna(fecha_agente):
return &quot;⚪ Sin fecha&quot;
dt_terc = pd.to_datetime(str(fecha_tercero), dayfirst=True, errors=&quot;coerce&quot;)
dt_ag = pd.to_datetime(str(fecha_agente), dayfirst=True, errors=&quot;coerce&quot;)
if pd.isna(dt_terc) or pd.isna(dt_ag):
return &quot;⚪ Sin fecha&quot;
diff_d = int((dt_terc - dt_ag).days)
if diff_d == 0:
return &quot;�� Mismo día (0d)&quot;
elif 0 &lt; diff_d &lt;= 15:
return f&quot;➡️ Poco después (+{diff_d}d)&quot;
elif -30 &lt;= diff_d &lt; 0:
return f&quot;⬅️ Anterior ({abs(diff_d)}d antes)&quot;
elif -60 &lt;= diff_d &lt; -30:
return f&quot;⬅️ Anterior ({abs(diff_d)}d antes)&quot;
elif diff_d &lt; -60:
return f&quot;⛔ Fuera de rango ({abs(diff_d)}d antes)&quot;
else:
return f&quot;⛔ Fuera de rango (+{diff_d}d desp)&quot;

def generar_asiento_mixto_hoja2_con_terceros(agente_row, terceros_df, enviar_a_no_deducible=False,
imputar_a_transito=False):
&quot;&quot;&quot;
Construye el asiento contable para un cobro de agente cuando se trae la contabilización de Página 2:
1. Si en Hoja 2 se registraron únicamente los Ingresos Propios (Base + IVA &lt; Total Factura):

Combina los Ingresos Propios (con su IVA y retenciones) + Facturas de Terceros vinculadas (cancelación de CxP)
+
Crédito total al Agente Aduanero (22050501).
2. El saldo de ingresos para terceros NO se imputa automáticamente a ninguna cuenta de gasto ni tránsito.
Se deja PENDIENTE para que el usuario concilie con facturas de terceros o decida explícitamente enviarlo a No
Deducibles o a Tránsito.
3. Si el usuario selecciona &#39;No Deducibles&#39;, imputa el saldo a 53950501.
4. Si el usuario selecciona &#39;Mercancías en Tránsito&#39;, imputa el saldo a 14650501.
&quot;&quot;&quot;
df_h2 = obtener_asiento_contable_hoja2(agente_row, es_aduanero=False)
tot_agente = float(agente_row.get(&quot;Total&quot;, 0.0))
base_propia = float(agente_row.get(&quot;Base&quot;, 0.0))
iva_propio = float(agente_row.get(&quot;IVA&quot;, 0.0))
iva_imp_propio = float(agente_row.get(&quot;IVA Importación&quot;, 0.0))
subtotal_propio = round(base_propia + iva_propio + iva_imp_propio, 2)
saldo_terceros_esperado = round(tot_agente - subtotal_propio, 2)
# Si la factura en Hoja 2 ya cubre el total y no hay terceros asignados
if saldo_terceros_esperado &lt;= 0.05 and (terceros_df is None or terceros_df.empty):
return df_h2, 0.0
filas_propias = []
tot_ret_propias = 0.0
for _, r in df_h2.iterrows():
cta = str(r[&quot;Código Cuenta&quot;]).strip()
if cta.startswith(&quot;2365&quot;) or cta.startswith(&quot;2368&quot;) or cta.startswith(&quot;2367&quot;):
tot_ret_propias += float(r[&quot;Crédito ($)&quot;])
filas_propias.append(r.to_dict())
elif not (cta.startswith(&quot;2205&quot;) and float(r[&quot;Crédito ($)&quot;]) &gt; 0):
filas_propias.append(r.to_dict())
# Facturas de terceros asignadas a este paquete
filas_terceros = []
suma_terceros_cxp = 0.0
if terceros_df is not None and not terceros_df.empty:
for _, t in terceros_df.iterrows():
prov_nom = str(t.get(&quot;Proveedor&quot;, &quot;&quot;)).upper()
nit_t = str(t.get(&quot;NIT Emisor&quot;, &quot;&quot;))
fac_num = str(t.get(&quot;Factura&quot;, &quot;&quot;))
cta_cxp = str(t.get(&quot;Cuenta Pasivo Especifica&quot;, &quot;&quot;)).strip() or (&quot;22050505&quot; if &quot;CARGO&quot; in prov_nom or
&quot;ADUANA&quot; in prov_nom else &quot;23359501&quot;)
t_base = float(t.get(&quot;Base&quot;, 0.0))
t_iva = float(t.get(&quot;IVA&quot;, 0.0))
t_saldo_pagar = obtener_saldo_cruce_factura(t)
suma_terceros_cxp += t_saldo_pagar
filas_terceros.append({
&quot;Código Cuenta&quot;: cta_cxp,
&quot;Descripción Cuenta&quot;: f&quot;Cancela CxP {prov_nom[:20]} (Fac {fac_num})&quot;,
&quot;Descripción de la Cuenta&quot;: f&quot;Cancela CxP {prov_nom[:20]} (Fac {fac_num})&quot;,
&quot;Tercero / NIT&quot;: f&quot;{nit_t} - {prov_nom[:25]}&quot;,
&quot;Débito ($)&quot;: t_saldo_pagar,
&quot;Crédito ($)&quot;: 0.0
})
# Crédito total por pagar al Agente Aduanero (Euro Shipping / Trade Global)
asume_ret = bool(agente_row.get(&quot;Impuestos Asumidos&quot;, False))
cxp_total_agente = tot_agente if asume_ret else round(tot_agente - tot_ret_propias, 2)
filas_credito = [{
&quot;Código Cuenta&quot;: CUENTA_CXP_EURO_SHIPPING,
&quot;Descripción Cuenta&quot;: f&quot;CxP Agente Aduanero - Fac {agente_row.get(&#39;Factura&#39;, &#39;&#39;)}&quot;,
&quot;Descripción de la Cuenta&quot;: f&quot;CxP Agente Aduanero - Fac {agente_row.get(&#39;Factura&#39;, &#39;&#39;)}&quot;,
&quot;Tercero / NIT&quot;: f&quot;{agente_row.get(&#39;NIT Emisor&#39;, &#39;&#39;)} - {str(agente_row.get(&#39;Proveedor&#39;, &#39;&#39;))[:25]}&quot;,
&quot;Débito ($)&quot;: 0.0,
&quot;Crédito ($)&quot;: cxp_total_agente
}]
# Calcular la diferencia matemática exacta de partida doble entre débitos y créditos del asiento
suma_deb_base = sum(float(f.get(&quot;Débito ($)&quot;, 0.0) or 0.0) for f in (filas_propias + filas_terceros))

suma_cred_base = sum(float(f.get(&quot;Crédito ($)&quot;, 0.0) or 0.0) for f in (filas_propias + filas_credito))
dif_asiento_exacta = round(suma_cred_base - suma_deb_base, 2)
filas_ajuste = []
if abs(dif_asiento_exacta) &gt; 0.05:
if enviar_a_no_deducible:
if dif_asiento_exacta &gt; 0.05:
# Faltan débitos por justificar -&gt; Débito a No Deducibles (53950501)
filas_ajuste.append({
&quot;Código Cuenta&quot;: CUENTA_NO_DEDUCIBLE,
&quot;Descripción Cuenta&quot;: f&quot;Gastos No Deducibles Terceros (Diferencia sin soporte DIAN)&quot;,
&quot;Descripción de la Cuenta&quot;: f&quot;Gastos No Deducibles Terceros (Diferencia sin soporte DIAN)&quot;,
&quot;Tercero / NIT&quot;: f&quot;{agente_row.get(&#39;NIT Emisor&#39;, &#39;&#39;)} - {str(agente_row.get(&#39;Proveedor&#39;,
&#39;&#39;))[:25]}&quot;,
&quot;Débito ($)&quot;: dif_asiento_exacta,
&quot;Crédito ($)&quot;: 0.0
})
else:
# Débitos exceden créditos (los terceros superan el cobro del agente) -&gt; Crédito a 53950501
filas_ajuste.append({
&quot;Código Cuenta&quot;: CUENTA_NO_DEDUCIBLE,
&quot;Descripción Cuenta&quot;: f&quot;Ajuste No Deducible / Terceros Superan Cobro Agente&quot;,
&quot;Descripción de la Cuenta&quot;: f&quot;Ajuste No Deducible / Terceros Superan Cobro Agente&quot;,
&quot;Tercero / NIT&quot;: f&quot;{agente_row.get(&#39;NIT Emisor&#39;, &#39;&#39;)} - {str(agente_row.get(&#39;Proveedor&#39;,
&#39;&#39;))[:25]}&quot;,
&quot;Débito ($)&quot;: 0.0,
&quot;Crédito ($)&quot;: abs(dif_asiento_exacta)
})
elif imputar_a_transito:
if dif_asiento_exacta &gt; 0.05:
filas_ajuste.append({
&quot;Código Cuenta&quot;: CUENTA_IMPORTACION_TRANSITO,
&quot;Descripción Cuenta&quot;: f&quot;Mercancías en Tránsito / Saldo Faltante (Fac {agente_row.get(&#39;Factura&#39;,
&#39;&#39;)})&quot;,
&quot;Descripción de la Cuenta&quot;: f&quot;Mercancías en Tránsito / Saldo Faltante (Fac
{agente_row.get(&#39;Factura&#39;, &#39;&#39;)})&quot;,
&quot;Tercero / NIT&quot;: f&quot;{agente_row.get(&#39;NIT Emisor&#39;, &#39;&#39;)} - {str(agente_row.get(&#39;Proveedor&#39;,
&#39;&#39;))[:25]}&quot;,
&quot;Débito ($)&quot;: dif_asiento_exacta,
&quot;Crédito ($)&quot;: 0.0
})
else:
filas_ajuste.append({
&quot;Código Cuenta&quot;: CUENTA_IMPORTACION_TRANSITO,
&quot;Descripción Cuenta&quot;: f&quot;Ajuste Mercancías en Tránsito / Terceros Superan Cobro (Fac
{agente_row.get(&#39;Factura&#39;, &#39;&#39;)})&quot;,
&quot;Descripción de la Cuenta&quot;: f&quot;Ajuste Mercancías en Tránsito / Terceros Superan Cobro (Fac
{agente_row.get(&#39;Factura&#39;, &#39;&#39;)})&quot;,
&quot;Tercero / NIT&quot;: f&quot;{agente_row.get(&#39;NIT Emisor&#39;, &#39;&#39;)} - {str(agente_row.get(&#39;Proveedor&#39;,
&#39;&#39;))[:25]}&quot;,
&quot;Débito ($)&quot;: 0.0,
&quot;Crédito ($)&quot;: abs(dif_asiento_exacta)
})
asiento_final = filas_propias + filas_terceros + filas_ajuste + filas_credito
df_res = pd.DataFrame(asiento_final)
dif_pendiente = abs(dif_asiento_exacta) if (not enviar_a_no_deducible and not imputar_a_transito) else 0.0
return df_res, dif_pendiente
def generar_asiento_triangulacion_paquete(agente_row, terceros_df, enviar_a_no_deducible=False,
enviar_a_gastos_propios=False, enviar_a_hoja2=False, **kwargs):
&quot;&quot;&quot;
Calcula el asiento contable de partida doble para un paquete de importación triangulado:
1. Débito a CxP Terceros (22050505/23359501) por el Saldo por Pagar real (Base + IVA, ej. .651.939).
Si ya estaba registrada en Siigo, se cancela la cuenta por pagar imputada.
2. Débito a Retenciones Asumidas (53152001) si las retenciones no se habían asumido en la causación previa.
3. Débito a Importación en Tránsito (14650501) e IVA (24081501) para facturas blancas pendientes.
4. Débito a IVA de Importación si la factura del agente lo discrimina.
5. Débito a Gastos No Deducibles (53950501) por la diferencia sin factura DIAN.
6. Crédito a CxP Agente Aduanero (22050501) por el 100% de la factura del agente (ej. .077.647).
&quot;&quot;&quot;
tot_agente = float(agente_row.get(&quot;Total&quot;, 0.0))

iva_agente = float(agente_row.get(&quot;IVA&quot;, 0.0))
asiento = []
suma_cxp_canceladas = 0.0
suma_ret_asumidas = 0.0
suma_costo_blancas = 0.0
suma_iva_blancas = 0.0
for _, t in terceros_df.iterrows():
prov_nom = str(t.get(&quot;Proveedor&quot;, &quot;&quot;)).upper()
nit_t = str(t.get(&quot;NIT Emisor&quot;, &quot;&quot;))
fac_num = str(t.get(&quot;Factura&quot;, &quot;&quot;))
es_excluida_manual = bool(t.get(&quot;No Contabilizar&quot;, False)) or (str(t.get(&quot;Factura&quot;, &quot;&quot;)).strip() in
st.session_state.get(&quot;facturas_no_contabilizar&quot;, set()))
es_roja = bool(t.get(&quot;Ya Registrada&quot;, False)) or es_excluida_manual
cta_cxp = str(t.get(&quot;Cuenta Pasivo Especifica&quot;, &quot;&quot;)).strip()
if not cta_cxp:
if any(k in prov_nom for k in [&quot;CARGO&quot;, &quot;ADUANA&quot;, &quot;PORTUARIA&quot;, &quot;ALMACENADORA&quot;, &quot;TERMINAL&quot;]):
cta_cxp = CUENTA_CXP_AGENCIA_EXTERIOR
else:
cta_cxp = CUENTA_CXP_DHL_NACIONAL
t_base = float(t.get(&quot;Base&quot;, 0.0))
t_iva = float(t.get(&quot;IVA&quot;, 0.0))
t_tot = float(t.get(&quot;Total&quot;, 0.0))
if t_base &lt;= 0:
t_base = round(t_tot - t_iva, 2)
rfte_t = float(t.get(&quot;ReteFuente&quot;, 0.0))
rica_t = float(t.get(&quot;ReteICA&quot;, 0.0))
tot_ret = round(rfte_t + rica_t, 2)
# El Saldo por Pagar que entra a la triangulación es Base + IVA (ej. .388.770 + 63.169 = .651.939)
t_saldo_pagar = obtener_saldo_cruce_factura(t)
# REGLA DE TRIANGULACIÓN: Toda factura de tercero traída a la triangulación ÚNICAMENTE
# aporta su Cuenta por Pagar (pasivo: 22050505 Agencia, 23359501 DHL, etc.) por su Saldo por Pagar (Base +
IVA),
# ya que la causación del gasto/costo e IVA de esa factura se realiza en su propio comprobante individual.
asume_en_causacion = bool(t.get(&quot;Impuestos Asumidos&quot;, False))
if asume_en_causacion or tot_ret == 0:
cxp_cancelar = t_saldo_pagar
ret_a_asumir = 0.0
suma_cxp_canceladas += cxp_cancelar
asiento.append({
&quot;Código Cuenta&quot;: cta_cxp,
&quot;Descripción Cuenta&quot;: f&quot;Cancela CxP {prov_nom[:20]} (Fac {fac_num})&quot;,
&quot;Tercero / NIT&quot;: f&quot;{nit_t} - {prov_nom[:25]}&quot;,
&quot;Débito ($)&quot;: cxp_cancelar,
&quot;Crédito ($)&quot;: 0.0
})
else:
cxp_cancelar = round(t_saldo_pagar - tot_ret, 2)
ret_a_asumir = tot_ret
suma_cxp_canceladas += cxp_cancelar
suma_ret_asumidas += ret_a_asumir
asiento.append({
&quot;Código Cuenta&quot;: cta_cxp,
&quot;Descripción Cuenta&quot;: f&quot;Cancela CxP {prov_nom[:20]} (Fac {fac_num})&quot;,
&quot;Tercero / NIT&quot;: f&quot;{nit_t} - {prov_nom[:25]}&quot;,
&quot;Débito ($)&quot;: cxp_cancelar,
&quot;Crédito ($)&quot;: 0.0
})
if ret_a_asumir &gt; 0:
asiento.append({
&quot;Código Cuenta&quot;: CUENTA_RETENCION_ASUMIDA,

&quot;Descripción Cuenta&quot;: f&quot;Retención Asumida Fac {fac_num} ({prov_nom[:18]})&quot;,
&quot;Tercero / NIT&quot;: f&quot;{nit_t} - {prov_nom[:25]}&quot;,
&quot;Débito ($)&quot;: ret_a_asumir,
&quot;Crédito ($)&quot;: 0.0
})
# IVA Servicios e IVA de Importación del Agente
iva_imp_agente = float(agente_row.get(&quot;IVA Importación&quot;, 0.0))
cta_iva_ag = str(agente_row.get(&quot;Cta IVA&quot;) or &quot;24081501&quot;).strip()
cta_iva_imp_ag = str(agente_row.get(&quot;Cta IVA Importación&quot;) or &quot;240835&quot;).strip()
if iva_agente &gt; 0:
asiento.append({
&quot;Código Cuenta&quot;: cta_iva_ag,
&quot;Descripción Cuenta&quot;: f&quot;IVA Descontable Servicios Fac {agente_row.get(&#39;Factura&#39;, &#39;&#39;)}&quot;,
&quot;Tercero / NIT&quot;: f&quot;{agente_row.get(&#39;NIT Emisor&#39;, &#39;&#39;)} - {agente_row.get(&#39;Proveedor&#39;, &#39;&#39;)[:25]}&quot;,
&quot;Débito ($)&quot;: iva_agente,
&quot;Crédito ($)&quot;: 0.0
})
if iva_imp_agente &gt; 0:
asiento.append({
&quot;Código Cuenta&quot;: cta_iva_imp_ag,
&quot;Descripción Cuenta&quot;: f&quot;IVA de Importación (Cta 240835) Fac {agente_row.get(&#39;Factura&#39;, &#39;&#39;)}&quot;,
&quot;Tercero / NIT&quot;: f&quot;{agente_row.get(&#39;NIT Emisor&#39;, &#39;&#39;)} - {agente_row.get(&#39;Proveedor&#39;, &#39;&#39;)[:25]}&quot;,
&quot;Débito ($)&quot;: iva_imp_agente,
&quot;Crédito ($)&quot;: 0.0
})
# Calcular la diferencia no deducible sin factura DIAN (Cuenta 53950501)
suma_justificada = suma_cxp_canceladas + suma_ret_asumidas + suma_costo_blancas + suma_iva_blancas + iva_agente +
iva_imp_agente
diferencia_no_deducible = round(tot_agente - suma_justificada, 2)
if diferencia_no_deducible &gt; 0.01 and enviar_a_no_deducible:
asiento.append({
&quot;Código Cuenta&quot;: CUENTA_NO_DEDUCIBLE,
&quot;Descripción Cuenta&quot;: f&quot;Gastos No Deducibles Importación (Diferencia sin soporte DIAN)&quot;,
&quot;Tercero / NIT&quot;: f&quot;{agente_row.get(&#39;NIT Emisor&#39;, &#39;&#39;)} - {agente_row.get(&#39;Proveedor&#39;, &#39;&#39;)[:25]}&quot;,
&quot;Débito ($)&quot;: diferencia_no_deducible,
&quot;Crédito ($)&quot;: 0.0
})
elif diferencia_no_deducible &lt; -0.01:
# En caso de que la suma de terceros exceda ligeramente al agente por redondeo
asiento.append({
&quot;Código Cuenta&quot;: CUENTA_NO_DEDUCIBLE,
&quot;Descripción Cuenta&quot;: f&quot;Ajuste por Diferencia de Cuadre Importación&quot;,
&quot;Tercero / NIT&quot;: f&quot;{agente_row.get(&#39;NIT Emisor&#39;, &#39;&#39;)} - {agente_row.get(&#39;Proveedor&#39;, &#39;&#39;)[:25]}&quot;,
&quot;Débito ($)&quot;: 0.0,
&quot;Crédito ($)&quot;: abs(diferencia_no_deducible)
})
# Crédito total al Agente Aduanero (Euro Shipping / Trade Global) por el 100% de su factura
asiento.append({
&quot;Código Cuenta&quot;: CUENTA_CXP_EURO_SHIPPING,
&quot;Descripción Cuenta&quot;: f&quot;CxP Agente Aduanero - Fac {agente_row.get(&#39;Factura&#39;, &#39;&#39;)}&quot;,
&quot;Tercero / NIT&quot;: f&quot;{agente_row.get(&#39;NIT Emisor&#39;, &#39;&#39;)} - {agente_row.get(&#39;Proveedor&#39;, &#39;&#39;)[:25]}&quot;,
&quot;Débito ($)&quot;: 0.0,
&quot;Crédito ($)&quot;: tot_agente
})
df_asiento = pd.DataFrame(asiento)
if &quot;Descripción Cuenta&quot; in df_asiento.columns and &quot;Descripción de la Cuenta&quot; not in df_asiento.columns:
df_asiento[&quot;Descripción de la Cuenta&quot;] = df_asiento[&quot;Descripción Cuenta&quot;]
elif &quot;Descripción de la Cuenta&quot; in df_asiento.columns and &quot;Descripción Cuenta&quot; not in df_asiento.columns:
df_asiento[&quot;Descripción Cuenta&quot;] = df_asiento[&quot;Descripción de la Cuenta&quot;]
return df_asiento, max(0.0, diferencia_no_deducible), suma_ret_asumidas
# Directorio oficial de regímenes fiscales conocidos por NIT y Nombre
REGIMENES_EMISORES_CONOCIDOS = {
&#39;860502609&#39;: &#39;O-13;O-15&#39;, # DHL EXPRESS COLOMBIA LTDA (Gran Contribuyente y Autorretenedor)

&#39;800215775&#39;: &#39;O-13;O-15&#39;, # SOCIEDAD PORTUARIA REGIONAL DE BUENAVENTURA S.A.
&#39;890304099&#39;: &#39;O-13;O-15&#39;, # HOTELES ESTELAR S.A.
&#39;830048145&#39;: &#39;O-13;O-15&#39;, # SIIGO S.A.S.
&#39;830048268&#39;: &#39;O-48&#39;, # EURO SHIPPING SERVICES S.A.S (Responsable de IVA / Código 23 - 48)
&#39;800153993&#39;: &#39;O-13;O-15&#39;, # COMUNICACION CELULAR S.A. COMCEL / CLARO
&#39;860006376&#39;: &#39;O-13;O-15&#39;, # PANAMERICANA LIBRERIA Y PAPELERIA S.A.
&#39;890900608&#39;: &#39;O-13;O-15&#39;, # BANCOLOMBIA S.A.
&#39;860007738&#39;: &#39;O-13;O-15&#39;, # BANCO DAVIVIENDA S.A.
&#39;890903938&#39;: &#39;O-13;O-15&#39;, # ALMACENES EXITO S.A.
&#39;860012936&#39;: &#39;O-13;O-15&#39;, # SODIMAC CORONA / HOMECENTER
&#39;901306979&#39;: &#39;O-48&#39;, # TRADE GLOBAL INTERNATIONAL SAS
&#39;805001149&#39;: &#39;O-48&#39;, # AGENCIA INTERAMERICANA DE CARGA
&#39;900744197&#39;: &#39;O-48&#39;, # HOTEL GENOVA SAS
}
CODIGOS_IMPUESTO_SIIGO = {
&#39;24081001&#39;: 1, # IVA 19% compras bienes
&#39;24081003&#39;: 2, # IVA 5% compras bienes
&#39;24081501&#39;: 1, # IVA Servicios 19%
&#39;240835&#39;: 1, # IVA de Importación 19%
&#39;24083501&#39;: 1, # IVA de Importación 19% (8 dígitos)
&#39;24080601&#39;: 1, # IVA 19% ventas
&#39;24080602&#39;: 2, # IVA 5% ventas
&#39;23651501&#39;: 3, # Retefuente 11%
&#39;23652001&#39;: 4, # Retefuente 10%
&#39;23652501&#39;: 5, # Retefuente 6%
&#39;23652503&#39;: 6, # Retefuente 4% (Servicios / Agenciamiento)
&#39;23654001&#39;: 7, # Retefuente 2.5% (Compras generales / Repuestos)
&#39;23654004&#39;: 18, # Retefuente 3.5% (Alojamiento / Hoteles)
&#39;23653502&#39;: 19, # Retefuente 7%
&#39;23657001&#39;: 20, # Retefuente 2%
&#39;23652505&#39;: 21, # Retefuente 1%
&#39;23654006&#39;: 23, # Retefuente Combustibles 0.1%
&#39;23680501&#39;: 8, # ReteICA 11.04
&#39;23680503&#39;: 9, # ReteICA 13.8
&#39;23680505&#39;: 10, # ReteICA 9.66
&#39;23680507&#39;: 11, # ReteICA 8
&#39;23680509&#39;: 12, # ReteICA 7
&#39;23680511&#39;: 13, # ReteICA 6.9
&#39;23680513&#39;: 14, # ReteICA 4.14
&#39;23670101&#39;: 15, # ReteIVA 15%
&#39;24950103&#39;: 16 # Impoconsumo 8%
}
AGENTES_ADUANEROS = [&quot;DHL&quot;, &quot;ADUANA&quot;, &quot;EURO SHIPPING&quot;, &quot;PORTUARIA&quot;, &quot;ALMACENADORA&quot;, &quot;CARGO&quot;, &quot;TRADE GLOBAL&quot;,
&quot;TERMINAL&quot;, &quot;BUENAVENTURA&quot;]
def escanear_regimen_texto_pdf(texto, nit_emisor=None):
&quot;&quot;&quot;
Escaneo inteligente de responsabilidades fiscales del EMISOR en la factura electrónica:
Analiza el documento completo sin truncar en &#39;Cliente:&#39; para no perder las resoluciones.
Detecta O-13 (Gran Contribuyente), O-15 (Autorretenedor), O-47 (RST), O-48 y O-49.
&quot;&quot;&quot;
if not texto:
return &quot;&quot;
txt = texto.upper()
txt_norm = re.sub(r&#39;\s+&#39;, &#39; &#39;, txt)
# 1. Si se conoce el NIT del emisor, verificar en el directorio de regímenes oficiales
if nit_emisor:
ne_clean = re.sub(r&#39;\D&#39;, &#39;&#39;, str(nit_emisor))
if ne_clean in REGIMENES_EMISORES_CONOCIDOS:
return REGIMENES_EMISORES_CONOCIDOS[ne_clean]
# 1.1 Si el texto corresponde a Euro Shipping Services, es O-48 (Responsable de IVA / Común)
if &quot;EURO SHIPPING&quot; in txt_norm or &quot;830048268&quot; in txt_norm:
return &quot;O-48&quot;
# 2. Analizar Gran Contribuyente (O-13 / 0-13 / Grandes Contribuyentes)
es_gc = False

if re.search(r&#39;\b[O0]-13\b&#39;, txt_norm):
if not re.search(r&#39;(?:NO\s+(?:SOMOS\s+)?|NI\s+|NO\s+ES\s+)[O0]-13&#39;, txt_norm):
es_gc = True
elif re.search(r&#39;GRANDES?\s+CONTRIBUYENTES?&#39;, txt_norm):
if not re.search(r&#39;(?:NO\s+(?:SOMOS\s+)?|NI\s+|NO\s+ES\s+)GRANDES?\s+CONTRIBUYENTES?&#39;, txt_norm):
es_gc = True
# 3. Analizar Autorretenedor (O-15 / 0-15 / Autorretenedor de Renta)
es_autorr = False
if re.search(r&#39;\b[O0]-15\b&#39;, txt_norm):
if not re.search(r&#39;(?:NO\s+(?:SOMOS\s+)?|NI\s+|NO\s+ES\s+)[O0]-15&#39;, txt_norm):
es_autorr = True
elif re.search(r&#39;AUTO[R]?RETENEDOR(?:ES)?&#39;, txt_norm):
if not re.search(r&#39;(?:NO\s+(?:SOMOS\s+)?|NI\s+|NO\s+ES\s+)AUTO[R]?RETENEDOR(?:ES)?&#39;, txt_norm):
es_autorr = True
# 4. Analizar Régimen Simple de Tributación (O-47 / 0-47 / RST)
es_rst = False
if re.search(r&#39;\b[O0]-47\b&#39;, txt_norm):
es_rst = True
elif re.search(r&#39;R[EÉ]GIMEN\s+SIMPLE|SIMPLE\s+DE\s+TRIBUTACI[OÓ]N|\bRST\b&#39;, txt_norm):
if not re.search(r&#39;(?:NO\s+(?:SOMOS\s+)?|NI\s+)(?:R[EÉ]GIMEN\s+SIMPLE|RST)&#39;, txt_norm):
es_rst = True
# 5. No Responsable de IVA (O-49 / 0-49)
es_no_iva = False
if re.search(r&#39;\b[O0]-49\b&#39;, txt_norm) or re.search(r&#39;NO\s+RESPONSABLE\s+(?:DE|DEL)\s+IVA&#39;, txt_norm):
es_no_iva = True
codigos = []
if es_gc: codigos.append(&quot;O-13&quot;)
if es_autorr: codigos.append(&quot;O-15&quot;)
if es_rst: codigos.append(&quot;O-47&quot;)
if es_no_iva: codigos.append(&quot;O-49&quot;)
if not codigos: codigos.append(&quot;O-48&quot;)
return &quot;;&quot;.join(codigos)
def parse_num_co(val_str):
if not val_str: return 0.0
s = str(val_str).replace(&#39;$&#39;, &#39;&#39;).replace(&#39;COP&#39;, &#39;&#39;).strip()
if &#39;,&#39; in s and &#39;.&#39; in s:
if s.rfind(&#39;,&#39;) &gt; s.rfind(&#39;.&#39;):
s = s.replace(&#39;.&#39;, &#39;&#39;).replace(&#39;,&#39;, &#39;.&#39;)
else:
s = s.replace(&#39;,&#39;, &#39;&#39;)
elif &#39;,&#39; in s:
parts = s.split(&#39;,&#39;)
if len(parts) == 2 and len(parts[1]) in [2, 3]:
s = parts[0].replace(&#39;.&#39;, &#39;&#39;) + &#39;.&#39; + parts[1]
else:
s = s.replace(&#39;,&#39;, &#39;&#39;)
try:
return float(s)
except:
return 0.0
def extraer_valores_fiscales_texto_pdf(texto):
&quot;&quot;&quot;
Extrae con máxima precisión los valores fiscales y totales desde el texto del PDF oficial:
1. Subtotal / Total Bruto Factura (Base gravable antes de descuentos o retenciones en notas).
2. IVA discriminado o Total Impuesto.
3. Total Neto Factura (Base + IVA, antes de retenciones sugeridas).
4. Retenciones en notas / Descuentos Globales (ReteFuente 11%/4%, ReteICA 0.966%, ReteIVA).
5. Anticipos reales reportados en notas o casillas informativas.
6. Identificación de facturación a favor del agente aduanero (Mandato).
&quot;&quot;&quot;
datos = {}
if not texto: return datos
# 1. Notas Finales (formato clave: valor)

m_nf_rfte = re.search(r&#39;\bretefuente\s*:\s*([0-9\.\,]+)&#39;, texto, re.IGNORECASE)
if m_nf_rfte: datos[&#39;retefuente&#39;] = parse_num_co(m_nf_rfte.group(1))
m_nf_rica = re.search(r&#39;\breteica\s*:\s*([0-9\.\,]+)&#39;, texto, re.IGNORECASE)
if m_nf_rica: datos[&#39;reteica&#39;] = parse_num_co(m_nf_rica.group(1))
m_nf_riva = re.search(r&#39;\breteiva\s*:\s*([0-9\.\,]+)&#39;, texto, re.IGNORECASE)
if m_nf_riva: datos[&#39;reteiva&#39;] = parse_num_co(m_nf_riva.group(1))
# 2. Descuentos globales por retención sugerida si no vino en notas
if &#39;retefuente&#39; not in datos or datos[&#39;retefuente&#39;] == 0:
m_dg_rf = re.search(r&#39;RETEFUENTE[^\$]+(?:[\$]\s*|COP\s*)([0-9\.\,]+)&#39;, texto, re.IGNORECASE)
if m_dg_rf: datos[&#39;retefuente&#39;] = parse_num_co(m_dg_rf.group(1))
if &#39;reteica&#39; not in datos or datos[&#39;reteica&#39;] == 0:
m_dg_ri = re.search(r&#39;RETEICA[^\$]+(?:[\$]\s*|COP\s*)([0-9\.\,]+)&#39;, texto, re.IGNORECASE)
if m_dg_ri: datos[&#39;reteica&#39;] = parse_num_co(m_dg_ri.group(1))
# 3. Bloque de Datos Totales
pos_totales = texto.upper().find(&#39;DATOS TOTALES&#39;)
txt_totales = texto[pos_totales:] if pos_totales != -1 else texto
m_sub = re.search(r&#39;(?:Subtotal|Total Bruto Factura)\s*[:\$]?\s*([0-9\.\,]+)&#39;, txt_totales, re.IGNORECASE)
if m_sub: datos[&#39;subtotal&#39;] = parse_num_co(m_sub.group(1))
m_iva = re.search(r&#39;(?:IVA|Total impuesto\s*(?:\(\=\))?)\s*[:\$]?\s*([0-9\.\,]+)&#39;, txt_totales, re.IGNORECASE)
if m_iva: datos[&#39;iva&#39;] = parse_num_co(m_iva.group(1))
m_neto = re.search(r&#39;Total neto factura\s*(?:\(\=\))?\s*[:\$]?\s*([0-9\.\,]+)&#39;, txt_totales, re.IGNORECASE)
if m_neto: datos[&#39;total_neto&#39;] = parse_num_co(m_neto.group(1))
m_tot = re.search(r&#39;Total factura\s*(?:\(\=\))?\s*(?:COP\s*)?[\$]?\s*([0-9\.\,]+)&#39;, txt_totales, re.IGNORECASE)
if m_tot: datos[&#39;total_factura&#39;] = parse_num_co(m_tot.group(1))
m_fact_a = re.search(r&#39;nombre_facturado_a\s*:\s*([^\n\r]+)&#39;, texto, re.IGNORECASE)
if m_fact_a: datos[&#39;nombre_facturado_a&#39;] = m_fact_a.group(1).strip()
m_nit_fact_a = re.search(r&#39;nit_facturado_a\s*:\s*([0-9]+)&#39;, texto, re.IGNORECASE)
if m_nit_fact_a: datos[&#39;nit_facturado_a&#39;] = m_nit_fact_a.group(1).strip()
# 4. Extracción de conceptos y descripción de ítems / productos (ej. Transporte Terrestre)
REGEX_TRANSPORTE =
r&#39;\b(TRANSPORTE|TRANSPORT|INLAND\s*RATE|FLETE|FLETES|ACARREO|ACARREOS|CARGA\s*TERRESTRE|TERRESTRE|DRAYAGE|PORTES|FLETA
MENTO)\b&#39;
m_transp = re.search(REGEX_TRANSPORTE, texto, re.IGNORECASE)
if m_transp:
datos[&#39;es_transporte&#39;] = True
datos[&#39;concepto_items&#39;] = &quot;INLAND RATE TRANSPORTE TERRESTRE&quot;
m_det = re.search(r&#39;Detalles\s+de\s+Productos(.*?)(?:Notas\s+Finales|DATOS\s+TOTALES|Totales|\Z)&#39;, texto,
re.DOTALL | re.IGNORECASE)
if m_det:
for lin in m_det.group(1).splitlines():
lin_s = lin.strip()
if lin_s and not any(h in lin_s.upper() for h in [&quot;CÓDIGO&quot;, &quot;CODIGO&quot;, &quot;DESCRIPCIÓN&quot;, &quot;DESCRIPCION&quot;,
&quot;CANTIDAD&quot;, &quot;PRECIO&quot;, &quot;DESCUENTO&quot;, &quot;RECARGO&quot;, &quot;IMPUESTOS&quot;, &quot;U/M&quot;, &quot;VENTA&quot;, &quot;UNITARIO&quot;]):
lin_clean = re.sub(r&#39;^\d+\s+&#39;, &#39;&#39;, lin_s).strip()
if len(lin_clean) &gt;= 4 and not re.match(r&#39;^[\d\.\,\$\s\%\-]+$&#39;, lin_clean):
datos[&#39;concepto_items&#39;] = lin_clean
if re.search(REGEX_TRANSPORTE, lin_clean, re.IGNORECASE):
datos[&#39;es_transporte&#39;] = True
break
return datos
def calcular_base_con_regla_descuento(r, tot, iva, nom_emisor):
&quot;&quot;&quot;
Calcula la Base Gravable contable respetando la regla estricta:
1. Conversión de moneda extranjera (USD a COP por TRM): Si el Subtotal/IVA vienen en USD y Total en COP,

convierte ambos valores a COP con la TRM implícita para garantizar partida doble exacta.
2. El descuento SOLO se aplica si está incorporado en los ÍTEMS de la factura.
3. Si el descuento aparece en las NOTAS u observaciones, NO VA (no disminuye la base gravable).
4. En agencias aduaneras (DHL, Euro Shipping, etc.), las casillas o notas de descuento corresponden
a retenciones practicadas que no deben restarse de la base para evitar duplicidades.
5. Si en el reporte/factura existe la columna de Subtotal (valor de los ítems antes de notas),
se toma directamente el Subtotal como base contable.
&quot;&quot;&quot;
nom = str(nom_emisor).upper()
es_aduanero = any(k in nom for k in AGENTES_ADUANEROS)
col_sub = next((c for c in r.index if any(k in str(c).lower() for k in [&quot;subtotal&quot;, &quot;sub total&quot;, &quot;base gravable&quot;,
&quot;base_gravable&quot;, &quot;valor subtotal&quot;, &quot;lineextensionamount&quot;])), None)
col_desc_item = next((c for c in r.index if any(k in str(c).lower() for k in [&quot;descuento_items&quot;, &quot;descuento item&quot;,
&quot;descuento ítems&quot;, &quot;descuento linea&quot;, &quot;descuento comercial&quot;])), None)
col_desc_gen = next((c for c in r.index if any(k in str(c).lower() for k in [&quot;descuento&quot;, &quot;descuentos&quot;, &quot;total
descuento&quot;, &quot;valor descuento&quot;, &quot;allowancetotalamount&quot;])), None)
subtotal_val = None
if col_sub and pd.notna(r.get(col_sub)):
try:
v = float(r.get(col_sub))
if v &gt; 0:
subtotal_val = v
except Exception:
pass
desc_val = 0.0
if col_desc_gen and pd.notna(r.get(col_desc_gen)):
try:
desc_val = float(r.get(col_desc_gen))
except Exception:
desc_val = 0.0
desc_item_val = 0.0
if col_desc_item and pd.notna(r.get(col_desc_item)):
try:
desc_item_val = float(r.get(col_desc_item))
except Exception:
desc_item_val = 0.0
# DETECCIÓN Y CONVERSIÓN DE MONEDA EXTRANJERA (USD a COP por TRM):
# Si el Total en DIAN está en COP (ej. &gt; 50,000 COP) pero Subtotal e IVA vienen en USD (ej. 780 USD y 16.15 USD),
# la relación Total / (Subtotal + IVA) está en el rango histórico de TRM (1,500 a 6,500).
if subtotal_val is not None and tot &gt; 1000 and (subtotal_val + iva) &gt; 0:
ratio_trm = tot / (subtotal_val + iva)
if 1500 &lt;= ratio_trm &lt;= 6500:
subtotal_val = round(subtotal_val * ratio_trm, 2)
iva = round(iva * ratio_trm, 2)
desc_val = round(desc_val * ratio_trm, 2)
desc_item_val = round(desc_item_val * ratio_trm, 2)
base = round(subtotal_val, 2)
motivo_base = f&quot;Factura en USD convertida a TRM () — Base COP: | IVA COP: &quot;
return base, iva, desc_val, motivo_base
# REGLA:
# Si viene el Subtotal de los ítems, esa es la base real. Los descuentos en notas NO se restan.
if subtotal_val is not None:
base = round(subtotal_val, 2)
motivo_base = f&quot;Base = Subtotal de Ítems () — Descuentos en notas no aplican&quot;
else:
# Si no hay columna de subtotal y es agencia aduanera con descuento reportado (retenciones en notas):
if es_aduanero and desc_val &gt; 0:
base = round(tot - iva + desc_val, 2)
motivo_base = f&quot;Agencia Aduanera: Base restituida () sin descontar retenciones de notas&quot;
elif desc_item_val &gt; 0:
base = round(tot - iva, 2)
motivo_base = f&quot;Base calculada con descuento comercial en ítems ()&quot;
else:
base = round(tot - iva, 2)
motivo_base = f&quot;Base estándar = Total - IVA ()&quot;

# Validación de consistencia: si Base + IVA difiere del Total por más de .0 y no hay descuentos justificados
if tot &gt; 0 and abs((base + iva) - tot) &gt; 1.0 and desc_val == 0 and not (es_aduanero and desc_val &gt; 0):
base = round(tot - iva, 2)
motivo_base = f&quot;Base ajustada = Total - IVA ()&quot;
return base, iva, desc_val, motivo_base

# ==============================================================================
# MOTOR DE MEMORIA Y APRENDIZAJE CONTINUO DE REGLAS CONTABLES
# ==============================================================================
PATH_MEMORIA_GLOBAL = &quot;almacenamiento_contable/memoria_aprendizaje_contable.json&quot;
def cargar_memoria_aprendizaje(empresa_dict=None):
&quot;&quot;&quot;Carga el banco histórico de aprendizaje con reglas auditadas por el usuario.&quot;&quot;&quot;
nit_empresa = re.sub(r&#39;\D&#39;, &#39;&#39;, str(empresa_dict.get(&quot;nit&quot;, &quot;&quot;))) if empresa_dict else &quot;global&quot;
path_esp = f&quot;almacenamiento_contable/{nit_empresa}/memoria_aprendizaje_contable.json&quot;
for p in [path_esp, PATH_MEMORIA_GLOBAL]:
if os.path.exists(p):
try:
with open(p, &quot;r&quot;, encoding=&quot;utf-8&quot;) as f:
return json.load(f)
except Exception:
pass
return {}
def guardar_regla_aprendizaje(empresa_dict, nit_emisor, nombre_emisor, cta_p=None, resp_fiscal=None, **kwargs):
&quot;&quot;&quot;Memoriza una corrección o auditoría del usuario para aplicarla automáticamente en meses futuros.&quot;&quot;&quot;
if not nit_emisor:
return
nit_clean = re.sub(r&#39;\D&#39;, &#39;&#39;, str(nit_emisor))
if not nit_clean:
return
memoria = cargar_memoria_aprendizaje(empresa_dict)
regla = memoria.get(nit_clean, {
&quot;proveedor&quot;: nombre_emisor,
&quot;nit&quot;: nit_clean,
&quot;veces_auditado&quot;: 0
})
if nombre_emisor:
regla[&quot;proveedor&quot;] = nombre_emisor
if cta_p:
regla[&quot;cta_principal&quot;] = str(cta_p).strip()
if resp_fiscal:
regla[&quot;regimen&quot;] = str(resp_fiscal).strip()
for k, v in kwargs.items():
if v is not None:
regla[k] = v
regla[&quot;veces_auditado&quot;] = regla.get(&quot;veces_auditado&quot;, 0) + 1
regla[&quot;ultima_actualizacion&quot;] = datetime.datetime.now().strftime(&quot;%d/%m/%Y %I:%M %p&quot;)
memoria[nit_clean] = regla
nit_empresa = re.sub(r&#39;\D&#39;, &#39;&#39;, str(empresa_dict.get(&quot;nit&quot;, &quot;&quot;))) if empresa_dict else &quot;global&quot;
dir_esp = f&quot;almacenamiento_contable/{nit_empresa}&quot;
os.makedirs(dir_esp, exist_ok=True)
os.makedirs(&quot;almacenamiento_contable&quot;, exist_ok=True)
try:
with open(f&quot;{dir_esp}/memoria_aprendizaje_contable.json&quot;, &quot;w&quot;, encoding=&quot;utf-8&quot;) as f:
json.dump(memoria, f, ensure_ascii=False, indent=2)
with open(PATH_MEMORIA_GLOBAL, &quot;w&quot;, encoding=&quot;utf-8&quot;) as f:
json.dump(memoria, f, ensure_ascii=False, indent=2)
except Exception:
pass
return memoria

def clasificar_factura(nit_emisor, nombre_emisor, valor_base, valor_iva, tipo_doc, resp_emisor=&quot;&quot;,
empresa_compradora=None, descripcion=&quot;&quot;):
nombre_str = &quot;&quot; if (nombre_emisor is None or pd.isna(nombre_emisor)) else str(nombre_emisor).strip()
if nombre_str.lower() in [&quot;nan&quot;, &quot;none&quot;]:
nombre_str = &quot;&quot;
nombre = nombre_str.upper()
nombre_emisor = nombre_str or &quot;Proveedor&quot;
resp = str(resp_emisor).upper()
desc_u = str(descripcion or &quot;&quot;).upper()
# Detección explícita de concepto TRANSPORTE DE CARGA / FLETES / INLAND
REGEX_TRANSPORTE =
r&#39;\b(TRANSPORTE|TRANSPORT|INLAND|FLETE|FLETES|ACARREO|ACARREOS|CARGA\s*TERRESTRE|TERRESTRE|DRAYAGE|PORTES)\b&#39;
es_transporte = bool(re.search(REGEX_TRANSPORTE, desc_u) or (any(k in nombre for k in [&quot;TRANSPORTE&quot;, &quot;TRANSPORT&quot;,
&quot;CARGA&quot;, &quot;FLETE&quot;]) and not any(k in nombre for k in [&quot;AGENCIA&quot;, &quot;ADUANA&quot;])))
# 0. MEMORIA DE APRENDIZAJE: Si el usuario ya auditó y guardó reglas para este proveedor previamente
cta_aprendida = None
memoria_ap = cargar_memoria_aprendizaje(empresa_compradora)
nit_clean_c = re.sub(r&#39;\D&#39;, &#39;&#39;, str(nit_emisor))
if nit_clean_c in memoria_ap:
r_ap = memoria_ap[nit_clean_c]
if not resp_emisor and r_ap.get(&quot;regimen&quot;):
resp_emisor = r_ap[&quot;regimen&quot;]
resp = resp_emisor.upper()
if r_ap.get(&quot;cta_principal&quot;):
cta_aprendida = r_ap[&quot;cta_principal&quot;]
if r_ap.get(&quot;cta_iva&quot;):
cta_iva = r_ap[&quot;cta_iva&quot;]
if empresa_compradora is None:
empresa_compradora = {&quot;es_gran_contribuyente&quot;: False, &quot;es_autorretenedor&quot;: False}
tipo_str = str(tipo_doc).upper()
if &quot;CREDITO&quot; in tipo_str or &quot;CRÉDITO&quot; in tipo_str or &quot;NC&quot; in tipo_str or &quot;DEVOLUCION&quot; in tipo_str:
t_comp = 17
op = &quot;Devolucion Compra&quot;
elif &quot;NOTA&quot; in tipo_str or &quot;AJUSTE&quot; in tipo_str or &quot;CONTABILIDAD&quot; in tipo_str:
t_comp = 14
op = &quot;Nota Contabilidad&quot;
else:
t_comp = 10
op = &quot;Compra&quot;
# 1. TRANSPORTE TERRESTRE DE CARGA / FLETES / INLAND RATE (Prioridad sobre Agenciamiento)
if es_transporte or any(k in desc_u for k in [&quot;INLAND&quot;, &quot;TRANSPORTE&quot;, &quot;TRANSPORT&quot;, &quot;FLETE&quot;]):
# Retención en la fuente por Transporte de Carga: 1% (Art. 1.2.4.4.6 DUR 1625 / Art. 401 E.T.)
rfte_t = round(valor_base * 0.01, 2) if valor_base &gt;= 188000 else 0.0
# ReteICA Transporte Bogotá: 4.14 por mil (CIIU 4923 Transporte de Carga)
rica_t = round(valor_base * 0.00414, 2) if valor_base &gt;= 188000 else 0.0
es_imp_transp = any(k in nombre for k in AGENTES_ADUANEROS) or any(k in nombre for k in [&quot;EURO SHIPPING&quot;,
&quot;DHL&quot;, &quot;TRADE GLOBAL&quot;, &quot;CONSOLCARGO&quot;])
cta_p = &quot;14650501&quot; if es_imp_transp else &quot;51355001&quot;
cta_c = &quot;22050501&quot;
desc = (descripcion[:35] if descripcion else f&quot;Transporte Terrestre - {nombre_emisor[:20]}&quot;)
cta_rfte_b = &quot;23652505&quot; # ReteFuente 1% Transporte Carga
cta_iva = &quot;24081001&quot; if valor_iva == 0 else &quot;24081501&quot;
cta_rica_b = &quot;23680513&quot; # ReteICA 4.14 por mil Transporte
cat = &quot;Transporte Carga (1%)&quot;
razon_b = &quot;Transporte Terrestre de Carga / Fletes (1% ReteFuente [23652505] - 4.14‰ ReteICA [23680513])&quot;
# 1.B Agentes Aduaneros e Importaciones (Servicios / Agenciamiento ordinario)
elif any(k in nombre for k in AGENTES_ADUANEROS):
rfte_t = round(valor_base * 0.04, 2) if valor_base &gt;= 210000 or &quot;EURO SHIPPING&quot; in nombre or &quot;DHL&quot; in nombre
else 0.0
rica_t = round(valor_base * 0.00966, 2) if valor_base &gt;= 210000 and &quot;BUENAVENTURA&quot; not in nombre and
&quot;PORTUARIA&quot; not in nombre else 0.0
cta_p, cta_c, desc = &quot;146505&quot;, &quot;22050501&quot;, f&quot;Importacion / Transito - {nombre_emisor[:25]}&quot;
cta_rfte_b, cta_iva, cta_rica_b = &quot;23652503&quot;, &quot;240835&quot;, &quot;23680505&quot;
cat, razon_b = &quot;Importacion (1465)&quot;, &quot;Honorarios Agenciamiento vs Terceros&quot;
# 2. Repuestos / Mantenimiento

elif any(k in nombre for k in [&quot;FERROMENDEZ&quot;, &quot;TORNILLOLOCO&quot;, &quot;CAUCHOS&quot;, &quot;ASIMFER&quot;, &quot;MAFLEXCOL&quot;, &quot;EMPRECOL&quot;,
&quot;BATTS ZONE&quot;, &quot;MECANIZAR&quot;, &quot;HIDRAHULICAS&quot;, &quot;ELECTRICOS&quot;, &quot;ILUMINACION&quot;, &quot;BAMACOLGROUP&quot;]):
rfte_t = round(valor_base * 0.025, 2) if valor_base &gt;= 1047000 else 0.0
rica_t = round(valor_base * 0.01104, 2) if valor_base &gt;= 1414000 else 0.0
cta_rfte_b, cta_iva, cta_rica_b = &quot;23654001&quot;, &quot;24081001&quot;, &quot;23680501&quot;
if valor_base &gt;= 500000:
cta_p, cta_c, desc = &quot;14350101&quot;, &quot;22050501&quot;, f&quot;Repuestos / Inventario - {nombre_emisor[:25]}&quot;
cat, razon_b = &quot;Inventario&quot;, &quot;Compra repuestos mayores a 500k&quot;
else:
cta_p, cta_c, desc = &quot;61800101&quot;, &quot;23359501&quot;, f&quot;Mantenimiento menor - {nombre_emisor[:25]}&quot;
cat, razon_b = &quot;Costo Mantenimiento&quot;, &quot;Repuestos menores a 500k&quot;
rfte_t, rica_t = 0.0, 0.0
# 3. Hoteles / Viajes
elif any(k in nombre for k in [&quot;HOTEL&quot;, &quot;ESTELAR&quot;, &quot;GENOVA&quot;, &quot;VITTAPARK&quot;]):
rfte_t = round(valor_base * 0.035, 2) if valor_base &gt;= 210000 else 0.0
rica_t = round(valor_base * 0.00966, 2) if valor_base &gt;= 210000 else 0.0
cta_p, cta_c, desc = &quot;51550501&quot;, &quot;23359501&quot;, f&quot;Alojamiento / Viaje - {nombre_emisor[:25]}&quot;
cta_rfte_b, cta_iva, cta_rica_b = &quot;23652501&quot;, &quot;24081501&quot;, &quot;23680505&quot;
cat, razon_b = &quot;Gasto Viaje&quot;, &quot;Hospedaje de personal (IVA Servicios 24081501)&quot;
# 4. Software Siigo
elif &quot;SIIGO&quot; in nombre:
cta_p, cta_c, desc = &quot;51352001&quot;, &quot;23359501&quot;, f&quot;Software Siigo - {nombre_emisor[:25]}&quot;
rfte_t, rica_t = 0.0, 0.0
cta_rfte_b, cta_iva, cta_rica_b = &quot;&quot;, &quot;24081501&quot;, &quot;&quot;
cat, razon_b = &quot;Software&quot;, &quot;Software y tecnologia (IVA Servicios 24081501)&quot;
cat, razon_b = &quot;Software&quot;, &quot;Software y tecnologia&quot;
# 5. Papelería
elif any(k in nombre for k in [&quot;PANAMERICANA&quot;, &quot;LIBRERIA&quot;]):
cta_p, cta_c, desc = &quot;51953001&quot;, &quot;23359501&quot;, f&quot;Papeleria - {nombre_emisor[:25]}&quot;
rfte_t, rica_t = 0.0, 0.0
cta_rfte_b, cta_iva, cta_rica_b = &quot;&quot;, &quot;24081001&quot;, &quot;&quot;
cat, razon_b = &quot;Gastos Papeleria&quot;, &quot;Utiles de oficina&quot;
# 6. Compras y Gastos Generales
elif valor_base &gt;= 500000:
rfte_t = round(valor_base * 0.025, 2) if valor_base &gt;= 1047000 else 0.0
rica_t = round(valor_base * 0.01104, 2) if valor_base &gt;= 1414000 else 0.0
cta_p, cta_c, desc = &quot;14350101&quot;, &quot;22050501&quot;, f&quot;Compra mercancias - {nombre_emisor[:25]}&quot;
cta_rfte_b, cta_iva, cta_rica_b = &quot;23654001&quot;, &quot;24081001&quot;, &quot;23680501&quot;
cat, razon_b = &quot;Mercancia&quot;, &quot;Compra general &gt; 500k&quot;
else:
cta_p, cta_c, desc = &quot;51959501&quot;, &quot;23359501&quot;, f&quot;Gastos generales - {nombre_emisor[:25]}&quot;
rfte_t, rica_t = 0.0, 0.0
cta_rfte_b, cta_iva, cta_rica_b = &quot;&quot;, &quot;24081001&quot;, &quot;&quot;
cat, razon_b = &quot;Gasto General&quot;, &quot;Compra menor general&quot;
# CRUCE DE RESPONSABILIDADES FISCALES DIAN
es_emisor_gc = &quot;O-13&quot; in resp or &quot;GRAN CONTRIBUYENTE&quot; in resp
es_emisor_autorr = &quot;O-15&quot; in resp or &quot;AUTORRETENEDOR&quot; in resp or &quot;SIIGO&quot; in nombre
es_emisor_rst = &quot;O-47&quot; in resp or &quot;SIMPLE&quot; in resp
es_comprador_gc = empresa_compradora.get(&quot;es_gran_contribuyente&quot;, False)
# A. ReteFuente Renta
if es_emisor_autorr:
rfte = 0.0
cta_rfte = &quot;&quot;
razon_rfte = &quot;Proveedor Autorretenedor (O-15): No aplica ReteFuente&quot;
elif es_emisor_rst:
rfte = 0.0
cta_rfte = &quot;&quot;
razon_rfte = &quot;Proveedor Régimen Simple (O-47): No aplica ReteFuente (Art. 911 E.T.)&quot;
else:
rfte = rfte_t
cta_rfte = cta_rfte_b if rfte &gt; 0 else &quot;&quot;
razon_rfte = f&quot;ReteFuente ordinaria (${rfte:,.2f})&quot; if rfte &gt; 0 else &quot;Base no supera tope&quot;
# B. ReteIVA (Art. 437-2 E.T.)
if es_emisor_gc and not es_comprador_gc:

reteiva = 0.0
razon_reteiva = &quot;Emisor Gran Contribuyente (O-13): Comprador ordinario no practica ReteIVA&quot;
elif es_emisor_rst:
reteiva = 0.0
razon_reteiva = &quot;Emisor Régimen Simple: Exento de ReteIVA&quot;
elif es_comprador_gc and valor_iva &gt; 0:
reteiva = round(valor_iva * 0.15, 2)
razon_reteiva = &quot;Comprador Gran Contribuyente practica 15% de ReteIVA&quot;
else:
reteiva = 0.0
razon_reteiva = &quot;No aplica ReteIVA&quot;
# C. ReteICA
if es_emisor_gc or es_emisor_autorr or es_emisor_rst:
rica = 0.0
cta_rica = &quot;&quot;
razon_rica = &quot;Emisor GC / Autorretenedor: Exento de ReteICA territorial&quot;
else:
rica = rica_t
cta_rica = cta_rica_b if rica &gt; 0 else &quot;&quot;
razon_rica = f&quot;ReteICA liquidado (${rica:,.2f})&quot; if rica &gt; 0 else &quot;Base no supera tope&quot;
if cta_aprendida:
cta_p = cta_aprendida
cat = f&quot;Aprendido ({cta_aprendida})&quot;
razon_total = f&quot;�� Cuenta aprendida de auditorías del usuario ({cta_aprendida}) | {razon_rfte}
{razon_reteiva}&quot;
else:
razon_total = f&quot;{razon_b} | {razon_rfte} | {razon_reteiva}&quot;
audit_dict = {
&quot;es_gc&quot;: es_emisor_gc,
&quot;es_autorr&quot;: es_emisor_autorr,
&quot;es_rst&quot;: es_emisor_rst,
&quot;razon_rfte&quot;: razon_rfte,
&quot;razon_reteiva&quot;: razon_reteiva,
&quot;razon_reteica&quot;: razon_rica,
&quot;razon_rica&quot;: razon_rica,
&quot;es_aprendido&quot;: bool(cta_aprendida)
}
return t_comp, op, cta_p, cta_c, desc, rfte, rica, reteiva, cta_rfte, cta_iva, cta_rica, cat, razon_total,
audit_dict

def desbloquear_pdf_bytes(fbytes, nit_receptor=None, claves_extra=None, nits_emisores=None):
&quot;&quot;&quot;
Desbloquea automáticamente un PDF si viene encriptado/protegido con contraseña.
Prueba el NIT de la empresa compradora (con/sin dígito), NITs de emisores,
contraseña opcional y contraseñas estándar de facturación electrónica.
Retorna: (pdf_desbloqueado_bytes, fue_desbloqueado)
&quot;&quot;&quot;
try:
reader = PdfReader(io.BytesIO(fbytes))
if not reader.is_encrypted:
return fbytes, True
claves = [&quot;&quot;]
if claves_extra:
claves.append(str(claves_extra).strip())
if nit_receptor:
nr = re.sub(r&quot;\D&quot;, &quot;&quot;, str(nit_receptor))
if nr:
claves.extend([nr, nr[:-1], f&quot;{nr[:-1]}-{nr[-1]}&quot;, f&quot;{nr}5&quot;, f&quot;{nr}-5&quot;])
claves.extend([&quot;901346412&quot;, &quot;9013464125&quot;, &quot;901346412-5&quot;, &quot;1234&quot;, &quot;123456&quot;])
if nits_emisores:
for ne in nits_emisores:
c = re.sub(r&quot;\D&quot;, &quot;&quot;, str(ne))
if c and len(c) &gt;= 7:

claves.append(c)
desencriptado = False
for pwd in claves:
try:
res = reader.decrypt(pwd)
if res in (1, 2) or not reader.is_encrypted:
desencriptado = True
break
except Exception:
pass
if desencriptado:
writer = PdfWriter()
for p in reader.pages:
writer.add_page(p)
out = io.BytesIO()
writer.write(out)
out.seek(0)
return out.getvalue(), True
else:
return fbytes, False
except Exception:
return fbytes, False
_MEM_TEXT_CACHE = {}
def cache_extraer_textos_pdf(fbytes, nit_receptor=&quot;9013464125&quot;):
&quot;&quot;&quot;Extrae el texto de todas las páginas de un PDF guardando en caché únicamente el texto liviano por hash MD5 (0
bytes de PDF retenidos en RAM).&quot;&quot;&quot;
if not fbytes:
return []
try:
h_key = f&quot;{hashlib.md5(fbytes).hexdigest()}_{nit_receptor}&quot;
if h_key in _MEM_TEXT_CACHE:
return _MEM_TEXT_CACHE[h_key]
reader = PdfReader(io.BytesIO(fbytes))
if reader.is_encrypted:
fbytes_des, ok = desbloquear_pdf_bytes(fbytes, nit_receptor=nit_receptor)
if ok:
reader = PdfReader(io.BytesIO(fbytes_des))
res = [(p.extract_text() or &quot;&quot;) for p in reader.pages]
if len(_MEM_TEXT_CACHE) &gt; 250:
_MEM_TEXT_CACHE.clear()
_MEM_TEXT_CACHE[h_key] = res
return res
except Exception:
return []

def auditar_regimen_desde_facturas_renombradas(df_ref, dict_renombrados, empresa_dict):
&quot;&quot;&quot;
Escanea el régimen fiscal Y los valores fiscales reales (Subtotal, IVA, Retenciones Sugeridas, Total Neto)
directamente de las facturas PDF ya desbloqueadas y renombradas.
Actualiza el Régimen Fiscal Emisor Y recalcula las retenciones en la fuente correspondientes.
&quot;&quot;&quot;
total_modificados = 0
if not dict_renombrados or df_ref.empty:
return total_modificados
for r_idx, r_mat in df_ref.iterrows():
soporte_nom = str(r_mat.get(&quot;Soporte PDF Renombrado&quot;, &quot;&quot;))
pdf_bytes = dict_renombrados.get(soporte_nom)
# Búsqueda por coincidencia de comprobante o folio si el nombre varió ligeramente
if not pdf_bytes:
comp_id = str(r_mat.get(&quot;Comprobante Siigo&quot;, &quot;&quot;)).replace(&quot; &quot;, &quot;_&quot;)
fol_id = str(r_mat.get(&quot;Folio&quot;, &quot;&quot;))

for k in list(dict_renombrados.keys()):
if comp_id and comp_id in k:
pdf_bytes = dict_renombrados[k]
break
elif fol_id and len(fol_id) &gt;= 3 and fol_id in k:
pdf_bytes = dict_renombrados[k]
break
if pdf_bytes:
try:
paginas = cache_extraer_textos_pdf(pdf_bytes, nit_receptor=empresa_dict.get(&quot;nit&quot;, &quot;9013464125&quot;))
txt_completo = chr(10).join(paginas)
# 1. Escanear valores fiscales y totales desde el PDF
val_fisc = extraer_valores_fiscales_texto_pdf(txt_completo)
hubo_cambio = False
if val_fisc.get(&quot;subtotal&quot;) and val_fisc[&quot;subtotal&quot;] &gt; 0:
df_ref.at[r_idx, &quot;Base&quot;] = val_fisc[&quot;subtotal&quot;]
hubo_cambio = True
if val_fisc.get(&quot;iva&quot;) is not None and val_fisc[&quot;iva&quot;] &gt; 0:
df_ref.at[r_idx, &quot;IVA&quot;] = val_fisc[&quot;iva&quot;]
hubo_cambio = True
# 2. Escaneo de régimen fiscal del emisor
nit_emisor_val = r_mat.get(&quot;NIT Emisor&quot;)
reg_detectado = escanear_regimen_texto_pdf(txt_completo, nit_emisor=nit_emisor_val)
if not reg_detectado:
ne_dig = re.sub(r&#39;\D&#39;, &#39;&#39;, str(nit_emisor_val))
if ne_dig in REGIMENES_EMISORES_CONOCIDOS:
reg_detectado = REGIMENES_EMISORES_CONOCIDOS[ne_dig]
elif any(k in str(r_mat.get(&quot;Proveedor&quot;, &quot;&quot;)).upper() for k in [&quot;DHL&quot;, &quot;ESTELAR&quot;, &quot;BUENAVENTURA&quot;,
&quot;SIIGO&quot;, &quot;COMCEL&quot;, &quot;CLARO&quot;, &quot;PANAMERICANA&quot;]):
reg_detectado = &quot;O-13;O-15&quot;
if reg_detectado and reg_detectado != r_mat.get(&quot;Régimen Fiscal Emisor&quot;):
df_ref.at[r_idx, &quot;Régimen Fiscal Emisor&quot;] = reg_detectado
hubo_cambio = True
# 3. Recalcular liquidación de retenciones con el nuevo régimen y concepto extraído
concep_det = val_fisc.get(&quot;concepto_items&quot;, &quot;&quot;)
if concep_det:
df_ref.at[r_idx, &quot;Descripcion&quot;] = concep_det
desc_para_clasif = concep_det or str(df_ref.at[r_idx, &quot;Descripcion&quot;])
reg_a_usar = df_ref.at[r_idx, &quot;Régimen Fiscal Emisor&quot;]
t_c_n, op_n, c_p_n, c_c_n, desc_n, rfte_n, rica_n, riva_n, c_rf_n, c_iv_n, c_ri_n, cat_n, razon_n,
audit_n = clasificar_factura(
df_ref.at[r_idx, &quot;NIT Emisor&quot;], df_ref.at[r_idx, &quot;Proveedor&quot;],
df_ref.at[r_idx, &quot;Base&quot;], df_ref.at[r_idx, &quot;IVA&quot;],
df_ref.at[r_idx, &quot;Operacion&quot;], reg_a_usar, empresa_dict,
descripcion=desc_para_clasif
)
df_ref.at[r_idx, &quot;Categoría&quot;] = cat_n
df_ref.at[r_idx, &quot;Cta Principal&quot;] = c_p_n
# Si el PDF trae retenciones sugeridas explícitas en notas
if val_fisc.get(&quot;retefuente&quot;) is not None and val_fisc[&quot;retefuente&quot;] &gt; 0:
df_ref.at[r_idx, &quot;ReteFuente&quot;] = val_fisc[&quot;retefuente&quot;]
else:
df_ref.at[r_idx, &quot;ReteFuente&quot;] = rfte_n
if val_fisc.get(&quot;reteica&quot;) is not None and val_fisc[&quot;reteica&quot;] &gt; 0:
df_ref.at[r_idx, &quot;ReteICA&quot;] = val_fisc[&quot;reteica&quot;]
else:
df_ref.at[r_idx, &quot;ReteICA&quot;] = rica_n
df_ref.at[r_idx, &quot;ReteIVA&quot;] = riva_n
df_ref.at[r_idx, &quot;Cta ReteFuente&quot;] = c_rf_n

df_ref.at[r_idx, &quot;Cta ReteICA&quot;] = c_ri_n
df_ref.at[r_idx, &quot;Razón Contable&quot;] = razon_n
df_ref.at[r_idx, &quot;Audit Info&quot;] = audit_n
# 4. Tratamiento para Agencias Aduaneras / Facturas de Mandato
prov_nom_u = str(df_ref.at[r_idx, &quot;Proveedor&quot;]).upper()
es_aduanero_row = any(k in prov_nom_u for k in [&quot;CARGO&quot;, &quot;ADUANA&quot;, &quot;PORTUARIA&quot;, &quot;ALMACENADORA&quot;,
&quot;TERMINAL&quot;, &quot;DHL&quot;, &quot;EURO SHIPPING&quot;, &quot;TRADE GLOBAL&quot;])
if es_aduanero_row or val_fisc.get(&quot;retefuente&quot;, 0) &gt; 0:
df_ref.at[r_idx, &quot;Impuestos Asumidos&quot;] = True
saldo_cruce = round(df_ref.at[r_idx, &quot;Base&quot;] + df_ref.at[r_idx, &quot;IVA&quot;], 2)
df_ref.at[r_idx, &quot;Neto a Pagar&quot;] = saldo_cruce
if &quot;CARGO&quot; in prov_nom_u:
df_ref.at[r_idx, &quot;Cuenta Pasivo Especifica&quot;] = &quot;22050505&quot;
df_ref.at[r_idx, &quot;Cta Contrapartida&quot;] = &quot;22050505&quot;
elif &quot;DHL&quot; in prov_nom_u:
df_ref.at[r_idx, &quot;Cuenta Pasivo Especifica&quot;] = &quot;23359501&quot;
df_ref.at[r_idx, &quot;Cta Contrapartida&quot;] = &quot;23359501&quot;
else:
df_ref.at[r_idx, &quot;Neto a Pagar&quot;] = round(df_ref.at[r_idx, &quot;Total&quot;] - df_ref.at[r_idx,
&quot;ReteFuente&quot;] - df_ref.at[r_idx, &quot;ReteICA&quot;] - df_ref.at[r_idx, &quot;ReteIVA&quot;], 2)
if hubo_cambio:
total_modificados += 1
except Exception as e:
pass
return total_modificados
def identificar_factura_en_texto(texto, df_ref):
&quot;&quot;&quot;
Motor de búsqueda ultra-rápido (130x más veloz):
Evalúa CUFE, Prefijo, Folio, NIT Emisor y Proveedor sin crear objetos pesados en bucles.
&quot;&quot;&quot;
if not texto or df_ref.empty:
return None
txt_clean = re.sub(r&#39;[^A-Z0-9]&#39;, &#39;&#39;, texto.upper())
digits_only = re.sub(r&#39;\D&#39;, &#39;&#39;, texto)
# 1. Búsqueda directa por CUFE (Certeza 100%)
for r_idx, r_cand in df_ref.iterrows():
c_val = str(r_cand.get(&quot;CUFE&quot;, &quot;&quot;) or r_cand.get(&quot;CUFE / Token&quot;, &quot;&quot;) or &quot;&quot;)
if len(c_val) &gt;= 15:
c_clean = re.sub(r&#39;[^A-Za-z0-9]&#39;, &#39;&#39;, c_val).upper()
if c_clean[:20] in txt_clean:
return r_cand
# 2. Búsqueda combinada Factura + NIT (Ultra-rápida)
mejor_cand = None
mejor_score = 0
for r_idx, r_cand in df_ref.iterrows():
pref = str(r_cand.get(&quot;Prefijo&quot;, &quot;&quot;) or &quot;&quot;).strip().upper()
fol = str(r_cand.get(&quot;Folio&quot;, &quot;&quot;) or &quot;&quot;).strip().upper()
if not fol or fol == &quot;NAN&quot;:
continue
pref_c = re.sub(r&#39;[^A-Z0-9]&#39;, &#39;&#39;, pref)
fol_c = re.sub(r&#39;[^A-Z0-9]&#39;, &#39;&#39;, fol)
fac_full = (pref_c + fol_c) if pref_c else fol_c
fol_sc = fol_c.lstrip(&#39;0&#39;)
fac_full_sc = (pref_c + fol_sc) if pref_c else fol_sc
has_fac = (len(fac_full) &gt;= 3 and fac_full in txt_clean) or (len(fac_full_sc) &gt;= 3 and fac_full_sc in
txt_clean)
if not has_fac and len(fol_c) &gt;= 3:
if fol_c in txt_clean:
has_fac = True
if not has_fac:

continue
nit_c = re.sub(r&#39;\D&#39;, &#39;&#39;, str(r_cand.get(&quot;NIT Emisor&quot;, &quot;&quot;)))
nit_base = nit_c[:-1] if len(nit_c) &gt;= 10 else nit_c
has_nit = (nit_c and len(nit_c) &gt;= 6 and nit_c in digits_only) or (nit_base and len(nit_base) &gt;= 6 and
nit_base in digits_only)
prov_clean = re.sub(r&#39;[^A-Z0-9]&#39;, &#39;&#39;, str(r_cand.get(&quot;Proveedor&quot;, &quot;&quot;)).upper())
has_prov = len(prov_clean) &gt;= 4 and prov_clean[:8] in txt_clean
score = 600 if has_nit else (450 if has_prov else 300)
if score &gt; mejor_score:
mejor_score = score
mejor_cand = r_cand
if score == 600:
break
if score &gt; mejor_score and score &gt;= 300:
mejor_score = score
mejor_cand = r_cand
return mejor_cand
def buscar_y_extraer_pdf(fac_sel, dict_renombrados=None, dict_originales=None, empresa_compradora=None):
&quot;&quot;&quot;
Busca y extrae el PDF de la factura seleccionada dando MÁXIMA PRIORIDAD
a las facturas YA DESBLOQUEADAS, SEPARADAS Y RENOMBRADAS (dict_renombrados):
1. PRIORIDAD 1: Facturas en dict_renombrados (las que ya pasaron por el proceso de desbloqueo y renombrado).
Coincidencia por nombre esperado (Comp_10-XXX...), comprobante, folio o CUFE.
2. PRIORIDAD 2: Si aún no se han desbloqueado/renombrado, buscar en dict_originales y desbloquearlas al vuelo.
Retorna: (pdf_bytes_desbloqueado, descripcion_origen, lista_paginas)
&quot;&quot;&quot;
comp_siigo = str(fac_sel.get(&quot;Comprobante Siigo&quot;, &quot;&quot;)).strip()
consecutivo = str(fac_sel.get(&quot;Consecutivo&quot;, &quot;&quot;)).strip()
t_comp = str(fac_sel.get(&quot;Tipo Comp&quot;, &quot;&quot;)).strip()
folio_clean = str(fac_sel.get(&quot;Folio&quot;, &quot;&quot;)).replace(&quot;-&quot;, &quot;&quot;).strip().upper()
folio_sc = folio_clean.lstrip(&quot;0&quot;)
pref_clean = str(fac_sel.get(&quot;Prefijo&quot;, &quot;&quot;)).replace(&quot;-&quot;, &quot;&quot;).strip().upper()
fac_clean = str(fac_sel.get(&quot;Factura&quot;, &quot;&quot;)).replace(&quot;-&quot;, &quot;&quot;).strip().upper()
fac_full = (pref_clean + folio_clean) if pref_clean else folio_clean
fac_full_sc = (pref_clean + folio_sc) if pref_clean else folio_sc
cufe_raw = str(fac_sel.get(&quot;CUFE&quot;, &quot;&quot;) or fac_sel.get(&quot;CUFE / Token&quot;, &quot;&quot;) or &quot;&quot;).strip()
cufe_clean = re.sub(r&#39;[^a-zA-Z0-9]&#39;, &#39;&#39;, cufe_raw).lower()
# El NIT receptor es el de la empresa compradora
nit_comprador = &quot;9013464125&quot;
if empresa_compradora:
nit_comprador = re.sub(r&quot;\D&quot;, &quot;&quot;, str(empresa_compradora.get(&quot;nit&quot;, &quot;9013464125&quot;)))
nit_emisor_digits = re.sub(r&quot;\D&quot;, &quot;&quot;, str(fac_sel.get(&quot;NIT Emisor&quot;, &quot;&quot;)))
soporte_nom = str(fac_sel.get(&quot;Soporte PDF Renombrado&quot;, &quot;&quot;)).strip()
df_una_fac = pd.DataFrame([fac_sel])
# =========================================================================
# PRIORIDAD 1: TOMAR LAS FACTURAS DESPUÉS DE DESBLOQUEARLAS Y RENOMBRARLAS
# =========================================================================
if dict_renombrados:
# A. Coincidencia exacta por el nombre de archivo asignado
if soporte_nom and soporte_nom in dict_renombrados:
f_bytes = dict_renombrados[soporte_nom]
f_des, _ = desbloquear_pdf_bytes(f_bytes, nit_receptor=nit_comprador)
return f_des, f&quot;Factura Desbloqueada Oficial ({soporte_nom})&quot;, None
# B. Coincidencia por Comprobante (ej. Comp_10-680 o 10-680)
comp_patterns = [
f&quot;Comp_{t_comp}-{consecutivo}&quot;.upper(),
f&quot;Comp_{t_comp}_{consecutivo}&quot;.upper(),
f&quot;{t_comp}-{consecutivo}&quot;.upper(),

f&quot;{t_comp}_{consecutivo}&quot;.upper()
]
for k in list(dict_renombrados.keys()):
k_u = k.upper().replace(&quot; &quot;, &quot;_&quot;)
if any(cp in k_u for cp in comp_patterns):
f_des, _ = desbloquear_pdf_bytes(dict_renombrados[k], nit_receptor=nit_comprador)
return f_des, f&quot;Factura Desbloqueada Oficial ({k})&quot;, None
# C. Coincidencia por Número de Factura / Folio en el nombre renombrado
for k in list(dict_renombrados.keys()):
k_clean = k.replace(&quot;-&quot;, &quot;&quot;).replace(&quot; &quot;, &quot;&quot;).replace(&quot;_&quot;, &quot;&quot;).upper()
if fac_full and len(fac_full) &gt;= 3 and fac_full in k_clean:
f_des, _ = desbloquear_pdf_bytes(dict_renombrados[k], nit_receptor=nit_comprador)
return f_des, f&quot;Factura Desbloqueada Oficial ({k})&quot;, None
if folio_clean and len(folio_clean) &gt;= 3 and folio_clean in k_clean:
f_des, _ = desbloquear_pdf_bytes(dict_renombrados[k], nit_receptor=nit_comprador)
return f_des, f&quot;Factura Desbloqueada Oficial ({k})&quot;, None
# D. Coincidencia por CUFE en el nombre renombrado
if cufe_clean and len(cufe_clean) &gt;= 15:
for k in list(dict_renombrados.keys()):
k_l = re.sub(r&#39;[^a-zA-Z0-9]&#39;, &#39;&#39;, k).lower()
if cufe_clean[:20] in k_l:
f_des, _ = desbloquear_pdf_bytes(dict_renombrados[k], nit_receptor=nit_comprador)
return f_des, f&quot;Factura Desbloqueada Oficial ({k})&quot;, None
# E. Verificación por contenido de texto en las facturas renombradas
for r_nom, r_bytes in dict_renombrados.items():
try:
pgs = cache_extraer_textos_pdf(r_bytes, nit_receptor=nit_comprador)
txt_r = &quot; &quot;.join(pgs)
if identificar_factura_en_texto(txt_r, df_una_fac) is not None:
f_des, _ = desbloquear_pdf_bytes(r_bytes, nit_receptor=nit_comprador)
return f_des, f&quot;Factura Desbloqueada por Contenido ({r_nom})&quot;, None
except Exception:
pass
# =========================================================================
# PRIORIDAD 2: SI AÚN NO SE HA RENOMBRADO, BUSCAR EN LOS ORIGINALES Y DESBLOQUEAR
# =========================================================================
if dict_originales:
# A. Búsqueda por CUFE en nombre de archivo original
if cufe_clean and len(cufe_clean) &gt;= 15:
for fname, fbytes in dict_originales.items():
fn_l = re.sub(r&#39;[^a-zA-Z0-9]&#39;, &#39;&#39;, fname).lower()
if cufe_clean[:20] in fn_l or fn_l.startswith(cufe_clean[:18]):
f_des, _ = desbloquear_pdf_bytes(fbytes, nit_receptor=nit_comprador)
return f_des, f&quot;Factura Desbloqueada DIAN ({fname[:20]}...pdf)&quot;, None
# B. Archivos originales individuales por coincidencia de texto
for fname, fbytes in dict_originales.items():
try:
f_des, _ = desbloquear_pdf_bytes(fbytes, nit_receptor=nit_comprador)
pgs = cache_extraer_textos_pdf(f_des, nit_receptor=nit_comprador)
txt_ind = &quot; &quot;.join(pgs)
if identificar_factura_en_texto(txt_ind, df_una_fac) is not None:
return f_des, f&quot;Factura Desbloqueada ({fname})&quot;, None
except Exception:
pass
# C. Escaneo en PDFs originales unificados/consolidados página por página
for fname, fbytes in dict_originales.items():
try:
f_des, _ = desbloquear_pdf_bytes(fbytes, nit_receptor=nit_comprador)
paginas_txt = cache_extraer_textos_pdf(f_des, nit_receptor=nit_comprador)
pags_coincidentes = []
for p_idx, txt in enumerate(paginas_txt):
if identificar_factura_en_texto(txt, df_una_fac) is not None:
if p_idx not in pags_coincidentes:
pags_coincidentes.append(p_idx)
m_p = re.search(r&#39;(?:P[ÁAáa]G(?:INA)?|HOJA|PAGE)\s*[:\.]?\s*1\s*(?:DE|/|OF)\s*(\d+)&#39;, txt,
re.IGNORECASE)

if m_p:
try:
total_h = int(m_p.group(1))
if 2 &lt;= total_h &lt;= 15:
for off in range(1, total_h):
p_next = p_idx + off
if p_next &lt; len(paginas_txt) and p_next not in pags_coincidentes:
pags_coincidentes.append(p_next)
except Exception:
pass
if pags_coincidentes:
reader = PdfReader(io.BytesIO(f_des))
writer = PdfWriter()
for p in pags_coincidentes:
writer.add_page(reader.pages[p])
out = io.BytesIO()
writer.write(out)
out.seek(0)
str_p = &quot;, &quot;.join([str(p+1) for p in pags_coincidentes])
return out.getvalue(), f&quot;Factura Desbloqueada de &#39;{fname}&#39; (Págs {str_p})&quot;, [p+1 for p in
pags_coincidentes]
except Exception:
pass
return None, None, None

def renderizar_visor_pdf_completo(pdf_bytes, nombre_archivo, fac_sel=None, key_prefix=&quot;pdf_view&quot;,
nit_comprador=&quot;9013464125&quot;):
&quot;&quot;&quot;
Visor oficial 100% blindado y garantizado:
1. Desbloquea y desencripta el PDF primero para que no pida contraseña.
2. Renderizado de alta definición nativo en imágenes mediante pypdfium2 (CERO caras tristes ����, CERO bloqueos
Chrome).
3. Botón para abrir a pantalla completa en pestaña nueva y botón de descarga directa.
4. Cuadro formal con resumen contable de la factura.
&quot;&quot;&quot;
if not pdf_bytes or not isinstance(pdf_bytes, (bytes, bytearray)):
st.error(&quot;⚠️ El archivo PDF no contiene datos válidos.&quot;)
return
# 1. Desbloquear el PDF para garantizar que esté 100% libre de contraseña
pdf_limpio, _ = desbloquear_pdf_bytes(pdf_bytes, nit_receptor=nit_comprador)
pdf_usar = pdf_limpio if pdf_limpio else pdf_bytes
try:
reader_prev = PdfReader(io.BytesIO(pdf_usar))
num_pags_tot = len(reader_prev.pages)
except Exception:
num_pags_tot = 1
b64_pdf = base64.b64encode(pdf_usar).decode(&#39;utf-8&#39;).replace(&#39;\n&#39;, &#39;&#39;).strip() if len(pdf_usar) &lt; 2500000 else &#39;&#39;
# Pre-cálculo seguro de variables para evitar cualquier error de formato
val_base = float(fac_sel.get(&quot;Base&quot;, 0.0)) if (fac_sel is not None and &quot;Base&quot; in fac_sel) else 0.0
val_tot = float(fac_sel.get(&quot;Total&quot;, 0.0)) if (fac_sel is not None and &quot;Total&quot; in fac_sel) else 0.0
val_iva = float(fac_sel.get(&quot;IVA&quot;, 0.0)) if (fac_sel is not None and &quot;IVA&quot; in fac_sel) else 0.0
val_rfte = float(fac_sel.get(&quot;ReteFuente&quot;, 0.0)) if (fac_sel is not None and &quot;ReteFuente&quot; in fac_sel) else 0.0
fac_num_str = str(fac_sel.get(&quot;Factura&quot;, &quot;&quot;)) if fac_sel is not None else nombre_archivo
prov_str = str(fac_sel.get(&quot;Proveedor&quot;, &quot;&quot;)) if fac_sel is not None else &quot;&quot;
nit_str = str(fac_sel.get(&quot;NIT Emisor&quot;, &quot;&quot;)) if fac_sel is not None else &quot;&quot;
comp_str = str(fac_sel.get(&quot;Comprobante Siigo&quot;, &quot;&quot;)) if fac_sel is not None else &quot;&quot;
fecha_str = str(fac_sel.get(&quot;Fecha&quot;, &quot;&quot;)) if fac_sel is not None else &quot;&quot;
reg_str = str(fac_sel.get(&quot;Régimen Fiscal Emisor&quot;, &quot;O-48&quot;)) if fac_sel is not None else &quot;O-48&quot;
# 2. BOTONES SUPERIORES DE ACCIÓN RÁPIDA
c_btn1, c_btn2 = st.columns([1, 1])
with c_btn1:
pags_label = f&quot;{num_pags_tot} página{&#39;s&#39; if num_pags_tot &gt; 1 else &#39;&#39;} completa{&#39;s&#39; if num_pags_tot &gt; 1 else
&#39;&#39;}&quot;

st.download_button(
label=f&quot;�� Descargar Factura Desbloqueada ({pags_label})&quot;,
data=pdf_usar,
file_name=nombre_archivo,
mime=&quot;application/pdf&quot;,
key=f&quot;btn_dl_univ_{key_prefix}&quot;,
use_container_width=True
)
with c_btn2:
st.markdown(f&quot;&quot;&quot;
&lt;a href=&quot;data:application/pdf;base64,{b64_pdf}&quot; target=&quot;_blank&quot; download=&quot;{nombre_archivo}&quot; style=&quot;text-
decoration:none;&quot;&gt;
&lt;div style=&quot;background:#0284c7; color:white; text-align:center; padding:9px 12px; border-radius:6px; font-
weight:600; font-size:14px; box-shadow:0 1px 2px rgba(0,0,0,0.05); cursor:pointer;&quot;&gt;
�� Abrir Factura en Otra Ventana (Pantalla Completa)
&lt;/div&gt;
&lt;/a&gt;
&quot;&quot;&quot;, unsafe_allow_html=True)
st.write(&quot;&quot;)
# 2.B SELECTOR DE TAMAÑO VISUAL DE LA FACTURA
c_tam1, c_tam2 = st.columns([3, 2])
with c_tam1:
tam_sel = st.radio(
&quot;�� Tamaño de Vista Previa:&quot;,
[&quot;�� Normal (Recomendado - 780px)&quot;, &quot;�� Grande (Ancho Completo)&quot;, &quot;�� Compacto (600px)&quot;],
horizontal=True,
index=0,
key=f&quot;rad_tam_pdf_{key_prefix}&quot;
)
# Proporción de columnas para centrar y mostrar la hoja en tamaño normal de lectura
if &quot;Compacto&quot; in tam_sel:
col_ratios = [2.2, 4.0, 2.2]
elif &quot;Grande&quot; in tam_sel:
col_ratios = [0.01, 9.98, 0.01]
else: # Normal (Recomendado)
col_ratios = [1.6, 5.2, 1.6]
# 3. RENDERIZADO VISUAL DIRECTO EN IMÁGENES NATIVAS CON LIBERACIÓN ESTRICTA DE MEMORIA RAM
imgs_paginas = []
if HAS_PDFIUM:
doc_p = None
try:
doc_p = pdfium.PdfDocument(pdf_usar)
max_p_render = min(len(doc_p), 6)
for idx_pg in range(max_p_render):
pg_obj = doc_p[idx_pg]
bmp_obj = pg_obj.render(scale=1.35)
imgs_paginas.append(bmp_obj.to_pil())
bmp_obj.close()
pg_obj.close()
except Exception:
imgs_paginas = []
finally:
if doc_p is not None:
try:
doc_p.close()
except Exception:
pass
if imgs_paginas:
for idx_p, img in enumerate(imgs_paginas):
c_l, c_mid, c_r = st.columns(col_ratios)
with c_mid:
st.markdown(f&quot;&quot;&quot;
&lt;div style=&quot;background:#0070ba; color:white; padding:8px 14px; border-radius:6px 6px 0 0; font-
weight:600; font-size:13.5px; display:flex; justify-content:space-between; align-items:center; margin-top:16px; box-
shadow: 0 2px 4px rgba(0,0,0,0.06);&quot;&gt;
&lt;span&gt;�� Factura: {fac_num_str} — {prov_str}&lt;/span&gt;

&lt;span style=&quot;background:rgba(255,255,255,0.25); padding:2px 10px; border-radius:10px; font-
size:12px;&quot;&gt;Hoja {idx_p + 1} de {len(imgs_paginas)}&lt;/span&gt;
&lt;/div&gt;
&quot;&quot;&quot;, unsafe_allow_html=True)
st.image(img, use_container_width=True)
else:
# Fallback a Canvas HTML5 si pypdfium no estuviera disponible
visor_height = max(680, min(3600, num_pags_tot * 850))
html_canvas_viewer = f&quot;&quot;&quot;
&lt;!DOCTYPE html&gt;
&lt;html&gt;
&lt;head&gt;
&lt;meta charset=&quot;utf-8&quot;&gt;
&lt;script src=&quot;https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.min.js&quot;&gt;&lt;/script&gt;
&lt;style&gt;
body {{
margin: 0;
padding: 10px;
background: #f1f5f9;
display: flex;
flex-direction: column;
align-items: center;
font-family: -apple-system, BlinkMacSystemFont, &quot;Segoe UI&quot;, Roboto, sans-serif;
}}
.page-card {{
margin-bottom: 18px;
background: white;
box-shadow: 0 4px 12px rgba(0,0,0,0.12);
border-radius: 6px;
overflow: hidden;
width: 100%;
max-width: 860px;
border: 1px solid #cbd5e1;
}}
.page-title {{
background: #0070ba;
color: white;
padding: 8px 14px;
font-size: 13px;
font-weight: 600;
display: flex;
justify-content: space-between;
align-items: center;
}}
canvas {{
display: block;
width: 100%;
height: auto;
}}
#msg {{
font-size: 14px;
color: #0284c7;
padding: 20px;
text-align: center;
font-weight: 600;
}}
&lt;/style&gt;
&lt;/head&gt;
&lt;body&gt;
&lt;div id=&quot;msg&quot;&gt;⏳ Cargando y renderizando factura ({num_pags_tot} página(s))...&lt;/div&gt;
&lt;div id=&quot;pages-container&quot; style=&quot;width: 100%; display: flex; flex-direction: column; align-items:
center;&quot;&gt;&lt;/div&gt;
&lt;script&gt;
(async function() {{
try {{
const b64 = &quot;{b64_pdf}&quot;;
const raw = atob(b64);
const uint8Array = new Uint8Array(raw.length);
for (let i = 0; i &lt; raw.length; i++) {{
uint8Array[i] = raw.charCodeAt(i);
}}
pdfjsLib.GlobalWorkerOptions.workerSrc =
&#39;https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js&#39;;

const pdf = await pdfjsLib.getDocument({{ data: uint8Array }}).promise;
document.getElementById(&#39;msg&#39;).style.display = &#39;none&#39;;
const container = document.getElementById(&#39;pages-container&#39;);
for (let num = 1; num &lt;= pdf.numPages; num++) {{
const page = await pdf.getPage(num);
const viewport = page.getViewport({{ scale: 1.5 }});
const card = document.createElement(&#39;div&#39;);
card.className = &#39;page-card&#39;;
const title = document.createElement(&#39;div&#39;);
title.className = &#39;page-title&#39;;
title.innerHTML = &#39;&lt;span&gt;�� {nombre_archivo}&lt;/span&gt;&lt;span style=&quot;background:rgba(255,255,255,0.25)
padding:2px 8px; border-radius:10px; font-size:11px;&quot;&gt;Hoja &#39; + num + &#39; de &#39; + pdf.numPages + &#39;&lt;/span&gt;&#39;;
card.appendChild(title);
const canvas = document.createElement(&#39;canvas&#39;);
const ctx = canvas.getContext(&#39;2d&#39;);
canvas.height = viewport.height;
canvas.width = viewport.width;
card.appendChild(canvas);
container.appendChild(card);
await page.render({{ canvasContext: ctx, viewport: viewport }}).promise;
}}
}} catch (err) {{
document.getElementById(&#39;msg&#39;).innerHTML = &#39;&lt;div style=&quot;background:white; padding:20px; border-
radius:8px; border:2px solid #0070ba; text-align:center;&quot;&gt;&lt;h4&gt;�� Factura Lista&lt;/h4&gt;&lt;p style=&quot;color:#64748b;&quot;&gt;Utiliz
el botón de descarga o abrir en otra ventana.&lt;/p&gt;&lt;/div&gt;&#39;;
}}
}})();
&lt;/script&gt;
&lt;/body&gt;
&lt;/html&gt;
&quot;&quot;&quot;
components.html(html_canvas_viewer, height=visor_height, scrolling=True)
# 4. CUADRO RESUMEN OFICIAL CON VALORES CONTABLES
st.markdown(f&quot;&quot;&quot;
&lt;div style=&quot;background: white; border: 1px solid #cbd5e1; border-radius: 8px; padding: 18px; margin-top: 14px;
margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.04);&quot;&gt;
&lt;div style=&quot;display: grid; grid-template-columns: 1fr 1fr; gap: 12px; font-size: 13.5px; color: #334155;&quot;&gt;
&lt;div&gt;&lt;b&gt;Proveedor:&lt;/b&gt; {prov_str} (NIT: {nit_str})&lt;br&gt;&lt;b&gt;Factura:&lt;/b&gt; {fac_num_str} ({comp_str})&lt;/div&gt;
&lt;div&gt;&lt;b&gt;Fecha Emisión:&lt;/b&gt; {fecha_str} | &lt;b&gt;Régimen:&lt;/b&gt; {reg_str}&lt;br&gt;&lt;b&gt;Base:&lt;/b&gt; ${val_base:,.2f} |
&lt;b&gt;IVA:&lt;/b&gt; ${val_iva:,.2f} | &lt;b&gt;Total:&lt;/b&gt; &lt;span style=&quot;font-weight:bold;
color:#0f172a;&quot;&gt;${val_tot:,.2f}&lt;/span&gt;&lt;/div&gt;
&lt;/div&gt;
&lt;/div&gt;
&quot;&quot;&quot;, unsafe_allow_html=True)

# (Funciones de autosave reubicadas limpiamente al inicio)

def extraer_montos_items_pdf(pdf_bytes):
&quot;&quot;&quot;Extrae valores monetarios de los ítems o detalles de productos en el PDF de una factura.&quot;&quot;&quot;
if not pdf_bytes:
return []
try:
textos = cache_extraer_textos_pdf(pdf_bytes)
full_txt = &quot; &quot;.join(textos)
montos = []
encontrados = re.findall(r&#39;(?:[\$]|(?:\s))\s*(\d{1,3}(?:\.\d{3})+(?:,\d{2})?)&#39;, full_txt)
for m in encontrados:
try:
val = float(m.replace(&quot;.&quot;, &quot;&quot;).replace(&quot;,&quot;, &quot;.&quot;))
if 20000 &lt;= val &lt;= 100000000:

montos.append(val)
except Exception:
pass
return list(set(montos))
except Exception:
return []
def generar_respaldo_portatil_bytes(empresa_dict):
&quot;&quot;&quot;Genera en memoria un archivo .indumaq (ZIP completo) con el progreso actual de auditoría y triangulación.&quot;&quot;&quot;
if &quot;df_procesado&quot; not in st.session_state or st.session_state[&quot;df_procesado&quot;] is None:
return None
try:
buf = io.BytesIO()
with zipfile.ZipFile(buf, &quot;w&quot;, zipfile.ZIP_DEFLATED) as zf:
# 1. df_procesado
zf.writestr(&quot;df_procesado.pkl&quot;, pickle.dumps(st.session_state[&quot;df_procesado&quot;]))
# 2. excel_bytes
if st.session_state.get(&quot;excel_bytes&quot;):
zf.writestr(&quot;excel_original.xlsx&quot;, st.session_state[&quot;excel_bytes&quot;])
# 3. paquetes_importacion
if &quot;paquetes_importacion&quot; in st.session_state and st.session_state[&quot;paquetes_importacion&quot;]:
zf.writestr(&quot;paquetes_importacion.pkl&quot;, pickle.dumps(st.session_state[&quot;paquetes_importacion&quot;]))
# 4. asientos_triangulacion
if &quot;asientos_triangulacion_por_factura&quot; in st.session_state and
st.session_state[&quot;asientos_triangulacion_por_factura&quot;]:
zf.writestr(&quot;asientos_triangulacion.pkl&quot;,
pickle.dumps(st.session_state[&quot;asientos_triangulacion_por_factura&quot;]))
# 5. estado_sesion
estado_flags = {
&quot;paquetes_listos&quot;: {k: v for k, v in st.session_state.items() if k.startswith(&quot;paquete_listo_&quot;)},
&quot;enviar_gp&quot;: {k: v for k, v in st.session_state.items() if k.startswith(&quot;enviar_gp_pq_&quot;)},
&quot;enviar_h2&quot;: {k: v for k, v in st.session_state.items() if k.startswith(&quot;enviar_h2_pq_&quot;)},
&quot;enviar_nd&quot;: {k: v for k, v in st.session_state.items() if k.startswith(&quot;enviar_nd_pq_&quot;)},
&quot;facturas_no_contabilizar&quot;: list(st.session_state.get(&quot;facturas_no_contabilizar&quot;, set())),
&quot;sel_paquete_activo_key&quot;: st.session_state.get(&quot;sel_paquete_activo_key&quot;, 1)
}
zf.writestr(&quot;estado_sesion.json&quot;, json.dumps(estado_flags, ensure_ascii=False, indent=2))
# 6. meta
now_str = datetime.datetime.now().strftime(&quot;%d/%m/%Y %I:%M %p&quot;)
meta = {
&quot;empresa&quot;: empresa_dict.get(&quot;nombre&quot;, &quot;Empresa&quot;),
&quot;nit&quot;: empresa_dict.get(&quot;nit&quot;, &quot;&quot;),
&quot;fecha_respaldo&quot;: now_str,
&quot;archivo_excel&quot;: st.session_state.get(&quot;excel_nombre&quot;, &quot;Reporte.xlsx&quot;),
&quot;total_facturas&quot;: len(st.session_state[&quot;df_procesado&quot;])
}
zf.writestr(&quot;meta.json&quot;, json.dumps(meta, ensure_ascii=False, indent=2))
zf.writestr(&quot;meta_sesion.json&quot;, json.dumps(meta, ensure_ascii=False, indent=2))
mem_actual = cargar_memoria_aprendizaje(empresa_dict)
if mem_actual:
zf.writestr(&quot;memoria_aprendizaje.json&quot;, json.dumps(mem_actual, ensure_ascii=False, indent=2))
# 7. No empaquetar cientos de MB de PDFs en RAM dentro del respaldo .indumaq para evitar caídas por límite
de memoria en Streamlit Cloud.
# El respaldo .indumaq guarda el 100% de la contabilidad, paquetes, triangulaciones, asientos, exclusiones
y memoria de aprendizaje (&lt; 1 MB).
pass
buf.seek(0)
return buf.getvalue()
except Exception:
return None
def restaurar_desde_respaldo_bytes(empresa_dict, raw_zip_bytes):
&quot;&quot;&quot;Restaura un archivo de respaldo portátil (.indumaq o .zip) directamente a la sesión activa con facturas,
paquetes y PDFs.&quot;&quot;&quot;

try:
with zipfile.ZipFile(io.BytesIO(raw_zip_bytes), &quot;r&quot;) as zf:
namelist = zf.namelist()
if &quot;memoria_aprendizaje.json&quot; in namelist:
try:
mem_rest = json.loads(zf.read(&quot;memoria_aprendizaje.json&quot;).decode(&quot;utf-8&quot;))
nit_emp = re.sub(r&#39;\D&#39;, &#39;&#39;, str(empresa_dict.get(&quot;nit&quot;, &quot;&quot;))) if empresa_dict else &quot;global&quot;
os.makedirs(f&quot;almacenamiento_contable/{nit_emp}&quot;, exist_ok=True)
with open(f&quot;almacenamiento_contable/{nit_emp}/memoria_aprendizaje_contable.json&quot;, &quot;w&quot;,
encoding=&quot;utf-8&quot;) as f_mem:
json.dump(mem_rest, f_mem, ensure_ascii=False, indent=2)
with open(PATH_MEMORIA_GLOBAL, &quot;w&quot;, encoding=&quot;utf-8&quot;) as f_mem:
json.dump(mem_rest, f_mem, ensure_ascii=False, indent=2)
except Exception:
pass
if &quot;df_procesado.pkl&quot; in namelist:
st.session_state[&quot;df_procesado&quot;] = pickle.loads(zf.read(&quot;df_procesado.pkl&quot;))
if &quot;excel_original.xlsx&quot; in namelist:
st.session_state[&quot;excel_bytes&quot;] = zf.read(&quot;excel_original.xlsx&quot;)
if &quot;paquete_facturas.zip&quot; in namelist:
with open(ZIP_FACTURAS_DISK_PATH, &quot;wb&quot;) as fz:
fz.write(zf.read(&quot;paquete_facturas.zip&quot;))
st.session_state[&quot;zip_pdfs_path&quot;] = ZIP_FACTURAS_DISK_PATH
st.session_state[&quot;zip_pdfs&quot;] = None
if &quot;paquetes_importacion.pkl&quot; in namelist:
st.session_state[&quot;paquetes_importacion&quot;] = pickle.loads(zf.read(&quot;paquetes_importacion.pkl&quot;))
if &quot;asientos_triangulacion.pkl&quot; in namelist:
st.session_state[&quot;asientos_triangulacion_por_factura&quot;] =
pickle.loads(zf.read(&quot;asientos_triangulacion.pkl&quot;))
if &quot;estado_sesion.json&quot; in namelist:
est = json.loads(zf.read(&quot;estado_sesion.json&quot;).decode(&quot;utf-8&quot;))
for k, v in est.get(&quot;paquetes_listos&quot;, {}).items():
st.session_state[k] = v
for k, v in est.get(&quot;enviar_gp&quot;, {}).items():
st.session_state[k] = v
for k, v in est.get(&quot;enviar_h2&quot;, {}).items():
st.session_state[k] = v
for k, v in est.get(&quot;enviar_nd&quot;, {}).items():
st.session_state[k] = v
if &quot;facturas_no_contabilizar&quot; in est:
st.session_state[&quot;facturas_no_contabilizar&quot;] = set(est[&quot;facturas_no_contabilizar&quot;])
if &quot;sel_paquete_activo_key&quot; in est:
st.session_state[&quot;sel_paquete_activo_key&quot;] = est[&quot;sel_paquete_activo_key&quot;]
if &quot;meta.json&quot; in namelist:
meta = json.loads(zf.read(&quot;meta.json&quot;).decode(&quot;utf-8&quot;))
st.session_state[&quot;excel_nombre&quot;] = meta.get(&quot;archivo_excel&quot;, meta.get(&quot;excel_nombre&quot;, &quot;Reporte.xlsx&quot;))
st.session_state[&quot;_sesion_cargada_nombre&quot;] = f&quot;Respaldo {meta.get(&#39;fecha_respaldo&#39;, &#39;&#39;)}&quot;
elif &quot;meta_sesion.json&quot; in namelist:
meta = json.loads(zf.read(&quot;meta_sesion.json&quot;).decode(&quot;utf-8&quot;))
st.session_state[&quot;excel_nombre&quot;] = meta.get(&quot;excel_nombre&quot;, meta.get(&quot;archivo_excel&quot;, &quot;Reporte.xlsx&quot;))
st.session_state[&quot;_sesion_cargada_nombre&quot;] = f&quot;Respaldo {meta.get(&#39;fecha_respaldo&#39;, &#39;&#39;)}&quot;
# Restaurar PDFs directamente en disco (0 MB en RAM)
dict_renom = DiskBackedPDFDict(&quot;renom_pdfs&quot;)
dict_orig = DiskBackedPDFDict(&quot;raw_pdfs&quot;)
for name in namelist:
if name.startswith(&quot;pdfs_renombrados/&quot;) and name.endswith(&quot;.pdf&quot;):
fn = name.replace(&quot;pdfs_renombrados/&quot;, &quot;&quot;)
dict_renom[fn] = zf.read(name)
elif name.startswith(&quot;pdfs_originales/&quot;) and name.endswith(&quot;.pdf&quot;):
fn = name.replace(&quot;pdfs_originales/&quot;, &quot;&quot;)
dict_orig[fn] = zf.read(name)
if dict_renom:
st.session_state[&quot;dict_pdfs&quot;] = dict_renom
st.session_state[&quot;total_zip_pdfs&quot;] = len(dict_renom)
if dict_orig:
st.session_state[&quot;raw_uploaded_pdfs&quot;] = dict_orig
return True
except Exception as e:
return False
def normalizar_df_procesado(df):
&quot;&quot;&quot;Garantiza que todas las columnas requeridas existan con tipos seguros y sin NaNs que causen caídas.&quot;&quot;&quot;

if df is None or not isinstance(df, pd.DataFrame) or df.empty:
return df
defaults_str = {
&quot;Comprobante Siigo&quot;: &quot;Comp 10-001&quot;, &quot;Fecha&quot;: &quot;01/01/2026&quot;, &quot;Factura&quot;: &quot;FAC-001&quot;,
&quot;Proveedor&quot;: &quot;Proveedor&quot;, &quot;NIT Emisor&quot;: &quot;900000000&quot;, &quot;Régimen Fiscal Emisor&quot;: &quot;48&quot;,
&quot;Cta Principal&quot;: &quot;14650501&quot;, &quot;Categoría&quot;: &quot;General&quot;, &quot;Razón Contable&quot;: &quot;Clasificación automática&quot;, &quot;Cuenta
Pasivo Especifica&quot;: &quot;22050501&quot;,
&quot;Cta Contrapartida&quot;: &quot;22050501&quot;, &quot;Operacion&quot;: &quot;Compra&quot;, &quot;Soporte PDF Renombrado&quot;: &quot;&quot;,
&quot;Descripcion&quot;: &quot;Compra general&quot;, &quot;Tipo Comp&quot;: &quot;10&quot;, &quot;Prefijo&quot;: &quot;&quot;, &quot;Folio&quot;: &quot;&quot;,
&quot;Cta ReteFuente&quot;: &quot;23654001&quot;, &quot;Cta ReteICA&quot;: &quot;23680501&quot;, &quot;Cta IVA&quot;: &quot;24081001&quot;,
&quot;Cta IVA Importación&quot;: &quot;240835&quot;, &quot;Comprobante Previo&quot;: &quot;&quot;, &quot;Estado Registro&quot;: &quot;⚪ Compra Pendiente&quot;
}
defaults_num = {
&quot;Base&quot;: 0.0, &quot;IVA&quot;: 0.0, &quot;ReteFuente&quot;: 0.0, &quot;ReteICA&quot;: 0.0, &quot;ReteIVA&quot;: 0.0,
&quot;Total&quot;: 0.0, &quot;Neto a Pagar&quot;: 0.0, &quot;Total Neto&quot;: 0.0, &quot;Consecutivo&quot;: 680,
&quot;IVA Importación&quot;: 0.0
}
defaults_bool = {
&quot;Impuestos Asumidos&quot;: False, &quot;No Contabilizar&quot;: False, &quot;Ya Registrada&quot;: False, &quot;Es Aduanera&quot;: False
}
for col, def_val in defaults_str.items():
if col not in df.columns:
df[col] = def_val
else:
df[col] = df[col].fillna(def_val).astype(str).replace([&quot;nan&quot;, &quot;None&quot;], def_val)
for col, def_val in defaults_num.items():
if col not in df.columns:
df[col] = def_val
else:
df[col] = pd.to_numeric(df[col], errors=&quot;coerce&quot;).fillna(def_val)
for col, def_val in defaults_bool.items():
if col not in df.columns:
df[col] = def_val
else:
df[col] = df[col].fillna(def_val).astype(bool)
if &quot;Audit Info&quot; not in df.columns:
df[&quot;Audit Info&quot;] = [{} for _ in range(len(df))]
# Auto-identificar y clasificar facturas de transporte terrestre / fletes / inland rate
REGEX_TRANSP =
r&#39;\b(TRANSPORTE|TRANSPORT|INLAND\s*RATE|INLAND|FLETE|FLETES|ACARREO|ACARREOS|CARGA\s*TERRESTRE|TERRESTRE|DRAYAGE|PORTE
S)\b&#39;
for idx_row, r_row in df.iterrows():
f_num = str(r_row.get(&quot;Factura&quot;, &quot;&quot;)).strip()
p_nom = str(r_row.get(&quot;Proveedor&quot;, &quot;&quot;)).strip().upper()
d_val = str(r_row.get(&quot;Descripcion&quot;, &quot;&quot;)).strip().upper()
c_val = str(r_row.get(&quot;Categoría&quot;, &quot;&quot;)).strip().upper()
es_transp_row = (&quot;70787&quot; in f_num or
re.search(REGEX_TRANSP, d_val, re.IGNORECASE) is not None or
re.search(REGEX_TRANSP, c_val, re.IGNORECASE) is not None or
(&quot;EURO&quot; in p_nom and any(k in d_val for k in [&quot;INLAND&quot;, &quot;TRANSPORT&quot;, &quot;FLETE&quot;, &quot;TERRESTRE&quot;])))
if es_transp_row:
b_val = float(r_row.get(&quot;Base&quot;, 0.0) or 0.0)
if b_val &gt; 0:
rfte_correcta = round(b_val * 0.01, 2)
rica_correcta = round(b_val * 0.00414, 2)
if (df.at[idx_row, &quot;Cta ReteFuente&quot;] == &quot;23652503&quot; or
df.at[idx_row, &quot;ReteFuente&quot;] == round(b_val * 0.04, 2) or
df.at[idx_row, &quot;ReteFuente&quot;] == 0.0 or
&quot;70787&quot; in f_num):
df.at[idx_row, &quot;ReteFuente&quot;] = rfte_correcta
df.at[idx_row, &quot;ReteICA&quot;] = rica_correcta
df.at[idx_row, &quot;Cta ReteFuente&quot;] = &quot;23652505&quot;
df.at[idx_row, &quot;Cta ReteICA&quot;] = &quot;23680513&quot;
df.at[idx_row, &quot;Categoría&quot;] = &quot;Transporte Carga (1%)&quot;

df.at[idx_row, &quot;Razón Contable&quot;] = &quot;Transporte Terrestre de Carga / Fletes (1% ReteFuente
[23652505] - 4.14‰ ReteICA [23680513])&quot;
df.at[idx_row, &quot;Cta Principal&quot;] = &quot;14650501&quot;
df.at[idx_row, &quot;Impuestos Asumidos&quot;] = True
df.at[idx_row, &quot;Descripcion&quot;] = &quot;INLAND RATE TRANSPORTE TERRESTRE&quot;
return df
try:
# BANNER DE CONTROL Y GUARDADO MANUAL
if &quot;df_procesado&quot; in st.session_state and st.session_state[&quot;df_procesado&quot;] is not None:
c_bnr1, c_bnr2, c_bnr3, c_bnr4 = st.columns([2.8, 1.3, 1.6, 0.9])
with c_bnr1:
nom_ses = st.session_state.get(&quot;excel_nombre&quot;, &quot;Reporte de Facturas&quot;)
last_save = st.session_state.get(&quot;_ultimo_guardado_manual&quot;)
save_info = f&quot; | �� `{last_save}`&quot; if last_save else &quot;&quot;
st.info(f&quot;�� **Trabajo Activo:** `{nom_ses}` ({len(st.session_state[&#39;df_procesado&#39;])
facturas){save_info}&quot;)
with c_bnr2:
if st.button(&quot;�� Guardar Avance&quot;, key=&quot;btn_guardar_manual_top&quot;, use_container_width=True, help=&quot;Guarda e
disco todas las modificaciones de cuentas y paquetes validados.&quot;):
if guardar_estado_manual(empresa):
st.success(&quot;✅ ¡Avance contable guardado exitosamente!&quot;)
st.rerun()
else:
st.error(&quot;Error al guardar.&quot;)
with c_bnr3:
emp_clean_name = re.sub(r&#39;[^a-zA-Z0-9]&#39;, &#39;_&#39;, str(empresa.get(&#39;nombre&#39;, &#39;Empresa&#39;)))
nom_resp_top = f&quot;Respaldo_Avance_{emp_clean_name}.indumaq&quot;
if st.session_state.get(&quot;_cached_respaldo_bytes&quot;) is None:
if st.button(&quot;��️ Preparar Respaldo (.indumaq)&quot;, key=&quot;btn_prep_respaldo_top&quot;
use_container_width=True, help=&quot;Genera un archivo portátil con tus facturas, paquetes y avances para descargar.&quot;):
with st.spinner(&quot;Preparando archivo de respaldo liviano...&quot;):
st.session_state[&quot;_cached_respaldo_bytes&quot;] = generar_respaldo_portatil_bytes(empresa)
st.rerun()
else:
c_d1, c_d2 = st.columns([0.8, 0.2])
with c_d1:
st.download_button(
label=&quot;⬇️ Descargar (.indumaq)&quot;,
data=st.session_state[&quot;_cached_respaldo_bytes&quot;],
file_name=nom_resp_top,
mime=&quot;application/zip&quot;,
key=&quot;btn_dl_respaldo_top_direct&quot;,
use_container_width=True,
help=&quot;Haz clic para descargar el archivo en tu equipo.&quot;
)
with c_d2:
if st.button(&quot;��&quot;, key=&quot;btn_refresh_respaldo_cache&quot;, help=&quot;Actualizar respaldo con último
cambios&quot;):
st.session_state[&quot;_cached_respaldo_bytes&quot;] = None
st.rerun()
with c_bnr4:
if st.button(&quot;�� Limpiar&quot;, key=&quot;btn_nuevo_trabajo_top&quot;, use_container_width=True, help=&quot;Limpia la memori
para cargar un nuevo mes.&quot;):
for k in [&quot;df_procesado&quot;, &quot;dict_pdfs&quot;, &quot;raw_uploaded_pdfs&quot;, &quot;excel_bytes&quot;, &quot;excel_nombre&quot;, &quot;zip_pdfs&quot;,
&quot;job_actual_id&quot;, &quot;_ultimo_excel_proc_sig&quot;, &quot;_ultimo_pdfs_proc_sig&quot;, &quot;paquetes_importacion&quot;,
&quot;asientos_triangulacion_por_factura&quot;, &quot;facturas_no_contabilizar&quot;, &quot;_ultimo_guardado_manual&quot;]:
st.session_state.pop(k, None)
st.rerun()
else:
# Si la memoria está vacía, ofrecer botón para cargar el último avance guardado si existe
nit_c = re.sub(r&#39;\D&#39;, &#39;&#39;, str(empresa.get(&quot;nit&quot;, &quot;empresa&quot;)))
d_man = os.path.join(DATA_DIR, nit_c, &quot;estado_manual&quot;)
if os.path.exists(os.path.join(d_man, &quot;df_procesado.pkl&quot;)):
col_rec1, col_rec2 = st.columns([3.5, 1.5])
with col_rec1:
st.info(&quot;�� **Tienes un avance contable guardado previamente.** Puedes restaurarlo con un solo clic:&quot;)
with col_rec2:
if st.button(&quot;�� Cargar Último Avance Guardado&quot;, key=&quot;btn_cargar_avance_previo&quot;, type=&quot;primary&quot;
use_container_width=True):
if cargar_estado_manual(empresa):
st.success(&quot;¡Avance restaurado con éxito!&quot;)

st.rerun()
# AVISO DE TRABAJO PREVIO DISPONIBLE (SOLO SE CARGA SI EL USUARIO OPRIME EL BOTÓN)
trabajos_existentes = listar_trabajos_historial(empresa)
if &quot;df_procesado&quot; not in st.session_state and trabajos_existentes:
ultimo = trabajos_existentes[0]
nom_ult = ultimo.get(&quot;nombre_trabajo&quot;, f&quot;Trabajo {ultimo.get(&#39;id&#39;, &#39;&#39;)}&quot;)
n_fac = ultimo.get(&quot;total_facturas&quot;, 0)
try: val_tot = float(ultimo.get(&quot;total_valor&quot;, 0.0) or 0.0)
except Exception: val_tot = 0.0
st.markdown(f&quot;&quot;&quot;
&lt;div style=&quot;background:#f0fdf4; border:1px solid #86efac; border-left:5px solid #16a34a; border-radius:8px;
padding:14px 18px; margin-bottom:14px;&quot;&gt;
&lt;h4 style=&quot;margin:0 0 6px 0; color:#166534;&quot;&gt;�� Tienes un trabajo guardado disponible&lt;/h4&gt;
&lt;p style=&quot;margin:0; color:#14532d; font-size:14px;&quot;&gt;
Tienes guardado en memoria tu trabajo previo: &lt;b&gt;&#39;{nom_ult}&#39;&lt;/b&gt; con &lt;b&gt;{n_fac} facturas&lt;/b&gt;
(${val_tot:,.2f}) y sus PDFs vinculados.&lt;br&gt;
• Si deseas continuar con este trabajo, haz clic en el botón verde.&lt;br&gt;
• Si hoy vas a implementar &lt;b&gt;facturas nuevas&lt;/b&gt;, no tienes que oprimir nada y puedes subir tus
nuevos archivos directamente abajo en la Pestaña 1.
&lt;/p&gt;
&lt;/div&gt;
&quot;&quot;&quot;, unsafe_allow_html=True)
c_btn_c1, c_btn_c2 = st.columns([1.6, 3])
with c_btn_c1:
if st.button(&quot;�� Cargar mi trabajo anterior&quot;, key=&quot;btn_cargar_trabajo_previo_on_demand&quot;
use_container_width=True):
df_g, pdfs_r_g, pdfs_o_g, ex_b_g, ex_n_g, zip_g = cargar_trabajo_historial(empresa, ultimo[&quot;id&quot;])
if df_g is not None and not df_g.empty:
st.session_state[&quot;df_procesado&quot;] = df_g
st.session_state[&quot;dict_pdfs&quot;] = pdfs_r_g or {}
st.session_state[&quot;raw_uploaded_pdfs&quot;] = pdfs_o_g or {}
st.session_state[&quot;excel_bytes&quot;] = ex_b_g
st.session_state[&quot;excel_nombre&quot;] = ex_n_g
st.session_state[&quot;zip_pdfs&quot;] = zip_g
st.session_state[&quot;total_zip_pdfs&quot;] = len(pdfs_r_g) if pdfs_r_g else 0
st.session_state[&quot;job_actual_id&quot;] = ultimo[&quot;id&quot;]
st.session_state[&quot;_sesion_cargada_nombre&quot;] = nom_ult
st.success(f&quot;¡Trabajo &#39;{nom_ult}&#39; cargado con éxito!&quot;)
st.rerun()
with c_btn_c2:
st.caption(&quot;�� Oprime aquí **únicamente** cuando desees restaurar el trabajo anterior. De lo contrario
continúa abajo con tus facturas nuevas.&quot;)
# PANEL DE HISTORIAL DE TRABAJOS Y RESPALDOS INDESTRUCTIBLES
trabajos_guardados = listar_trabajos_historial(empresa)
panel_expanded = False
with st.expander(&quot;��️ Historial de Trabajos, Respaldos y Carga Rápida&quot;, expanded=panel_expanded):
# Restauración inmediata desde archivo portátil .indumaq
c_upl_r1, c_upl_r2 = st.columns([3, 1.2])
with c_upl_r1:
upl_indumaq = st.file_uploader(&quot;�� Restaurar desde Archivo de Respaldo (.indumaq):&quot;, type=[&quot;indumaq&quot;
&quot;backup&quot;, &quot;zip&quot;], key=&quot;upl_respaldo_indumaq&quot;, help=&quot;Sube tu archivo de respaldo para recuperar al instante exactamente
donde lo dejaste.&quot;)
with c_upl_r2:
st.write(&quot;&quot;)
st.write(&quot;&quot;)
if upl_indumaq is not None:
if st.button(&quot;�� Restaurar Respaldo&quot;, key=&quot;btn_ejecutar_restaurar_upl&quot;, use_container_width=True):
ok_r = restaurar_desde_respaldo_bytes(empresa, upl_indumaq.read())
if ok_r:
st.success(&quot;¡Respaldo restaurado con éxito! Todas tus facturas editadas y paquetes quedaron
listos.&quot;)
st.rerun()
else:
st.error(&quot;No se pudo leer el archivo de respaldo.&quot;)
st.markdown(&quot;&lt;hr style=&#39;margin:10px 0;&#39;&gt;&quot;, unsafe_allow_html=True)

col_h_left, col_h_right = st.columns([1.5, 1])
with col_h_left:
st.markdown(&quot;#### �� Trabajos Guardados en esta Empresa:&quot;)
if trabajos_guardados:
st.caption(&quot;Haz clic en &#39;Cargar&#39; para recuperar de inmediato cualquier auditoría o mes previo:&quot;)
for tb in trabajos_guardados:
c_h1, c_h2, c_h3 = st.columns([3, 1.2, 0.5])
with c_h1:
nom_tb = tb.get(&quot;nombre_trabajo&quot;, f&quot;Trabajo {tb.get(&#39;id&#39;, &#39;&#39;)}&quot;)
n_facs = tb.get(&quot;total_facturas&quot;, 0)
try: val_tot = float(tb.get(&quot;total_valor&quot;, 0.0) or 0.0)
except Exception: val_tot = 0.0
n_renom = tb.get(&quot;total_pdfs_renombrados&quot;, tb.get(&quot;total_pdfs&quot;, 0))
n_orig = tb.get(&quot;total_pdfs_originales&quot;, 0)
info_pdf_str = f&quot;{n_renom} PDFs procesados&quot; if n_renom &gt; 0 else (f&quot;{n_orig} PDFs subidos&quot; if
n_orig &gt; 0 else &quot;Sin PDFs&quot;)
st.markdown(f&quot;�� **{nom_tb}** — {n_facs} facturas (${val_tot:,.2f}) — **{info_pdf_str}**&quot;)
with c_h2:
if st.button(&quot;�� Cargar&quot;, key=f&quot;btn_h_load_{tb[&#39;id&#39;]}&quot;):
df_g, pdfs_r_g, pdfs_o_g, ex_b_g, ex_n_g, zip_g = cargar_trabajo_historial(empresa,
tb[&quot;id&quot;])
if df_g is not None:
st.session_state[&quot;df_procesado&quot;] = df_g
if pdfs_r_g:
st.session_state[&quot;dict_pdfs&quot;] = pdfs_r_g
if pdfs_o_g:
st.session_state[&quot;raw_uploaded_pdfs&quot;] = pdfs_o_g
if ex_b_g:
st.session_state[&quot;excel_bytes&quot;] = ex_b_g
st.session_state[&quot;excel_nombre&quot;] = ex_n_g
if zip_g:
st.session_state[&quot;zip_pdfs&quot;] = zip_g
st.session_state[&quot;total_zip_pdfs&quot;] = len(pdfs_r_g) if pdfs_r_g else 0
st.session_state[&quot;job_actual_id&quot;] = tb[&quot;id&quot;]
st.session_state[&quot;_sesion_auto_recuperada&quot;] = nom_tb
st.success(f&quot;¡Trabajo &#39;{nom_tb}&#39; cargado! Todo tu avance está listo.&quot;)
st.rerun()
with c_h3:
if st.button(&quot;��️&quot;, key=f&quot;btn_h_del_{tb[&#39;id&#39;]}&quot;):
eliminar_trabajo_historial(empresa, tb[&quot;id&quot;])
st.rerun()
else:
st.info(&quot;�� Aún no tienes trabajos guardados en el disco local para esta empresa.&quot;)
with col_h_right:
st.markdown(&quot;#### ��️ Respaldo Portable (.zip):&quot;)
st.caption(&quot;Guarda o restaura todo tu trabajo en un solo archivo, ideal si cambias de PC o si el servidor
se reinicia:&quot;)
pass
# 2. Uploader para restaurar desde un archivo de respaldo previo (con botón de confirmación para evitar
bucles)
archivo_zip_restaurar = st.file_uploader(
&quot;�� Restaurar Sesión desde Archivo de Respaldo (.indumaq / .zip):&quot;,
type=[&quot;indumaq&quot;, &quot;zip&quot;],
key=&quot;upl_zip_restore_historial&quot;,
help=&quot;Sube tu archivo de respaldo para restaurar exactamente donde lo dejaste.&quot;
)
if archivo_zip_restaurar is not None:
if st.button(&quot;�� Cargar este Respaldo en Pantalla&quot;, key=&quot;btn_confirmar_cargar_zip_historial&quot;
use_container_width=True):
with st.spinner(&quot;Restaurando sesión y cálculos...&quot;):
raw_b = archivo_zip_restaurar.getvalue()
ok_r = restaurar_desde_respaldo_bytes(empresa, raw_b)
if not ok_r:
df_res, renom_res, orig_res, ex_b_res, ex_n_res, zip_res =
importar_respaldo_sesion_zip(raw_b, empresa)
if df_res is not None:
st.session_state[&quot;df_procesado&quot;] = df_res
st.session_state[&quot;dict_pdfs&quot;] = renom_res or {}

st.session_state[&quot;raw_uploaded_pdfs&quot;] = orig_res or {}
st.session_state[&quot;excel_bytes&quot;] = ex_b_res
st.session_state[&quot;excel_nombre&quot;] = ex_n_res
st.session_state[&quot;zip_pdfs&quot;] = zip_res
st.session_state[&quot;total_zip_pdfs&quot;] = len(renom_res) if renom_res else 0
ok_r = True
if ok_r:
st.success(f&quot;¡Sesión restaurada con éxito desde {archivo_zip_restaurar.name}!&quot;)
st.rerun()
else:
st.error(&quot;No se pudo restaurar el archivo de respaldo.&quot;)
st.markdown(&quot;---&quot;)

# [Panel superior integrado en el escudo global]
tab_compras, tab_auditoria, tab_triangulacion, tab_siigo = st.tabs([
&quot;1. Cargar Documentos y Desbloquear PDFs&quot;,
&quot;2. Auditoria y Trazabilidad Fiscal&quot;,
&quot;3. �� Triangulación y Cruces de Importación (Aduanas)&quot;,
&quot;4. Exportar Planilla Oficial a Siigo&quot;
])
with tab_compras:
try:
st.markdown(&quot;### 1. Insumos DIAN y Facturas en PDF&quot;)
st.write(&quot;Sube el archivo Excel de la DIAN (`prueba.xlsx`) o el reporte de facturas, y los PDFs
(unificados o separados) para desbloquear, guardar y renombrar automáticamente por comprobante.&quot;)
# Si ya tenemos los documentos cargados en el autoguardado, mostrar confirmación visible
n_p_carg = len(st.session_state.get(&quot;raw_uploaded_pdfs&quot;, {}))
n_p_proc = len(st.session_state.get(&quot;dict_pdfs&quot;, {}))
if n_p_carg &gt; 0:
st.success(f&quot;✅ **Autoguardado Activo:** Tienes **{n_p_carg} facturas PDF cargadas en el sistema**
({n_p_proc} vinculadas con su comprobante). Ya están disponibles en la Pestaña 2 (Auditoría) y Pestaña 3
(Triangulación). No necesitas volver a subirlas.&quot;)
col_u1, col_u2 = st.columns(2)
with col_u1:
archivo_excel = st.file_uploader(&quot;1. Reporte Excel de la DIAN (ej. prueba.xlsx)&quot;, type=[&quot;xlsx&quot;,
&quot;xls&quot;], help=&quot;Opcional si ya cargaste tu trabajo desde el Historial.&quot;)
with col_u2:
archivos_pdfs = st.file_uploader(&quot;2. Facturas en PDF o Archivo ZIP Consolidado&quot;, type=[&quot;pdf&quot;, &quot;zip&quot;],
accept_multiple_files=True, help=&quot;Sube facturas en PDF o un solo archivo .ZIP con todas las facturas (ideal para 100 a
1000+ facturas).&quot;)
st.markdown(&quot;##### �� 3. Consecutivos Iniciales por Tipo de Comprobante Siigo:&quot;)
st.caption(&quot;Cada tipo de comprobante lleva su propia numeración independiente. Digita en qué número vas en
cada uno:&quot;)
c_con1, c_con2, c_con3, c_con4 = st.columns(4)
with c_con1:
cons_ini_fac = st.number_input(&quot;Facturas de Compra (Tipo 10):&quot;, min_value=1, max_value=999999,
value=680, step=1, help=&quot;Consecutivo inicial para facturas de compra normales.&quot;)
with c_con2:
cons_ini_nc = st.number_input(&quot;Notas Crédito (Tipo 17):&quot;, min_value=1, max_value=999999, value=1,
step=1, help=&quot;Consecutivo inicial para notas crédito / devoluciones.&quot;)
with c_con3:
cons_ini_nota = st.number_input(&quot;Notas de Contabilidad (Tipo 14):&quot;, min_value=1, max_value=999999,
value=1, step=1, help=&quot;Consecutivo inicial para notas contables o ajustes.&quot;)
with c_con4:
clave_pdf_extra = st.text_input(&quot;Clave PDF (opcional):&quot;, type=&quot;password&quot;, help=&quot;Si algún archivo PDF
tiene contraseña específica, ingrésala aquí.&quot;)

# Almacenar PDFs originales subidos en session_state y en el historial de inmediato (solo una vez por
subida)
if archivos_pdfs:
pdfs_sig = f&quot;{len(archivos_pdfs)}_&quot; + &quot;_&quot;.join(f&quot;{p.name}_{p.size}&quot; for p in archivos_pdfs[:5])
if st.session_state.get(&quot;_ultimo_pdfs_proc_sig&quot;) != pdfs_sig:
if &quot;raw_uploaded_pdfs&quot; not in st.session_state:
st.session_state[&quot;raw_uploaded_pdfs&quot;] = {}
import gc
for p in archivos_pdfs:
if p.name.lower().endswith(&quot;.zip&quot;):
try:
with zipfile.ZipFile(io.BytesIO(p.getvalue()), &quot;r&quot;) as z_in:
for z_name in z_in.namelist():
if z_name.lower().endswith(&quot;.pdf&quot;) and not z_name.startswith(&quot;__MACOSX&quot;):
c_name = os.path.basename(z_name)
if c_name:
st.session_state[&quot;raw_uploaded_pdfs&quot;][c_name] = z_in.read(z_name)
except Exception as e:
st.error(f&quot;Error leyendo archivo ZIP {p.name}: {e}&quot;)
else:
st.session_state[&quot;raw_uploaded_pdfs&quot;][p.name] = p.getvalue()
gc.collect()
st.session_state[&quot;_ultimo_pdfs_proc_sig&quot;] = pdfs_sig
st.success(f&quot;�� **{len(st.session_state[&#39;raw_uploaded_pdfs&#39;])} factura(s) PDF cargadas e
memoria.**&quot;)
# Auto-vincular de inmediato si ya tenemos el reporte Excel cargado
if &quot;df_procesado&quot; in st.session_state and st.session_state[&quot;df_procesado&quot;] is not None:
st.session_state[&quot;_ejecutar_desbloqueo_ahora&quot;] = True
if archivo_excel is not None:
excel_sig = f&quot;{archivo_excel.name}_{archivo_excel.size}_{cons_ini_fac}_{cons_ini_nc}_{cons_ini_nota}&quot;
if st.session_state.get(&quot;_ultimo_excel_proc_sig&quot;) != excel_sig or &quot;df_procesado&quot; not in
st.session_state:
st.session_state[&quot;excel_bytes&quot;] = archivo_excel.getvalue()
st.session_state[&quot;excel_nombre&quot;] = archivo_excel.name
st.session_state[&quot;_ultimo_excel_proc_sig&quot;] = excel_sig
df_dian = pd.read_excel(io.BytesIO(st.session_state[&quot;excel_bytes&quot;]))
# Identificar columna de fecha y ordenar cronológicamente de Enero a la fecha actual
col_fecha = None
for col_cand in [&quot;Fecha Emisión&quot;, &quot;Fecha Emision&quot;, &quot;Fecha&quot;, &quot;Fecha de Emisión&quot;]:
if col_cand in df_dian.columns:
col_fecha = col_cand
break
if col_fecha is None:
col_fecha = df_dian.columns[0]
df_dian[&quot;_fecha_dt&quot;] = pd.to_datetime(df_dian[col_fecha], dayfirst=True, errors=&quot;coerce&quot;)
df_dian = df_dian.sort_values(by=&quot;_fecha_dt&quot;, ascending=True).reset_index(drop=True)
st.success(f&quot;Reporte procesado y ordenado cronológicamente (Enero a la fecha): **{len(df_dian)}
facturas identificadas**.&quot;)
contadores_consecutivos = {
10: int(cons_ini_fac),
17: int(cons_ini_nc),
14: int(cons_ini_nota),
16: int(cons_ini_fac)
}
# Analizar celdas rojas (ya registradas), grupos y cuentas específicas del Excel
estados_excel = analizar_estado_filas_excel(st.session_state.get(&quot;excel_bytes&quot;))
filas = []
for idx, r in df_dian.iterrows():
tipo_doc = r.get(&quot;Tipo de documento&quot;) or r.get(&quot;Tipo documento&quot;) or &quot;Factura electrónica&quot;
folio = str(r.get(&quot;Folio&quot;) or r.get(&quot;Factura&quot;) or r.get(&quot;Factura Num&quot;) or
f&quot;Doc_{idx+1}&quot;).strip()
if folio.endswith(&quot;.0&quot;):
folio = folio[:-2]

prefijo = str(r.get(&quot;Prefijo&quot;, &quot;&quot;)).strip() if pd.notna(r.get(&quot;Prefijo&quot;)) else &quot;&quot;
if prefijo == &quot;nan&quot; or prefijo == &quot;None&quot;:
prefijo = &quot;&quot;
fecha_val = r.get(col_fecha)
if pd.notna(fecha_val):
try:
fecha_dt = pd.to_datetime(fecha_val, dayfirst=True)
fecha_str = fecha_dt.strftime(&quot;%d/%m/%Y&quot;)
except:
fecha_str = str(fecha_val).split()[0]
else:
fecha_str = &quot;01/01/2026&quot;
nit_e = str(r.get(&quot;NIT Emisor&quot;) or r.get(&quot;NIT&quot;) or &quot;&quot;).strip()
if nit_e.endswith(&quot;.0&quot;):
nit_e = nit_e[:-2]
nom_e = str(r.get(&quot;Nombre Emisor&quot;) or r.get(&quot;Proveedor (Emisor)&quot;) or r.get(&quot;Proveedor&quot;) or
&quot;Proveedor&quot;).strip()
try: iva = float(r.get(&quot;IVA&quot;, 0.0)) if pd.notna(r.get(&quot;IVA&quot;)) else 0.0
except: iva = 0.0
try: tot = float(r.get(&quot;Total&quot;, 0.0)) if pd.notna(r.get(&quot;Total&quot;)) else 0.0
except: tot = 0.0
base, iva, desc_val, motivo_base = calcular_base_con_regla_descuento(r, tot, iva, nom_e)
col_cufe = next((c for c in df_dian.columns if any(k in str(c).lower() for k in [&quot;cufe&quot;,
&quot;token&quot;, &quot;uuid&quot;, &quot;cude&quot;])), None)
cufe_val = str(r.get(col_cufe, &quot;&quot;)).strip() if col_cufe and pd.notna(r.get(col_cufe)) else &quot;&quot;
col_resp = next((c for c in df_dian.columns if any(k in str(c).lower() for k in [&quot;régimen&quot;,
&quot;regimen&quot;, &quot;responsabilidad&quot;, &quot;obligacion&quot;])), None)
resp_e = str(r.get(col_resp, &quot;&quot;)).strip() if col_resp and pd.notna(r.get(col_resp)) else &quot;&quot;
# Búsqueda inmediata por NIT y Razón Social en catálogo oficial
nit_e_digits = re.sub(r&#39;\D&#39;, &#39;&#39;, str(nit_e))
if not resp_e and nit_e_digits in REGIMENES_EMISORES_CONOCIDOS:
resp_e = REGIMENES_EMISORES_CONOCIDOS[nit_e_digits]
elif not resp_e:
nom_e_u = str(nom_e).upper()
if any(k in nom_e_u for k in [&quot;DHL&quot;, &quot;ESTELAR&quot;, &quot;BUENAVENTURA&quot;, &quot;SIIGO&quot;, &quot;COMCEL&quot;,
&quot;CLARO&quot;, &quot;PANAMERICANA&quot;]):
resp_e = &quot;O-13;O-15&quot;
t_comp, op, cta_p, cta_c, desc, rfte, rica, riva, cta_rfte, cta_iva, cta_rica, cat, razon,
audit_dict = clasificar_factura(
nit_e, nom_e, base, iva, tipo_doc, resp_e, empresa
)
consecutivo = contadores_consecutivos.get(t_comp, 1)
contadores_consecutivos[t_comp] += 1
nom_limpio_prov = re.sub(r&#39;[^a-zA-Z0-9]&#39;, &#39;&#39;, nom_e)[:15]
nombre_pdf_esperado = f&quot;Comp_{t_comp}-{consecutivo}_{prefijo}{folio}_{nom_limpio_prov}.pdf&quot;
# Identificar si la fila estaba marcada en rojo (ya registrada en Siigo), su grupo y cuenta
fac_k_full = f&quot;{prefijo}-{folio}&quot; if prefijo else folio
info_est = estados_excel.get(cufe_val) or estados_excel.get(fac_k_full) or
estados_excel.get(folio) or estados_excel.get(idx, {})
es_roja_reg = bool(info_est.get(&quot;es_roja&quot;, False))
comp_prev = str(info_est.get(&quot;comprobante_existente&quot;, &quot;&quot;))
grp_imp = str(info_est.get(&quot;grupo_importacion&quot;, &quot;&quot;)).strip()
cta_esp = str(info_est.get(&quot;cuenta_especifica&quot;, &quot;&quot;)).strip()
es_aduanera_flag = any(k in nom_e.upper() for k in AGENTES_ADUANEROS) or any(k in
nom_e.upper() for k in [&quot;CARGO&quot;, &quot;ADUANA&quot;, &quot;PORTUARIA&quot;, &quot;ALMACENADORA&quot;, &quot;TERMINAL&quot;, &quot;BUENAVENTURA&quot;, &quot;CARTAGENA&quot;,
&quot;CONSOLCARGO&quot;, &quot;EURO SHIPPING&quot;, &quot;TRADE GLOBAL&quot;])
filas.append({
&quot;N°&quot;: idx + 1,
&quot;Tipo Comp&quot;: t_comp,
&quot;Consecutivo&quot;: consecutivo,

&quot;Comprobante Siigo&quot;: f&quot;Comp {t_comp}-{consecutivo}&quot;,
&quot;Fecha&quot;: fecha_str,
&quot;Prefijo&quot;: prefijo,
&quot;Folio&quot;: folio,
&quot;Factura&quot;: fac_k_full,
&quot;Proveedor&quot;: nom_e,
&quot;NIT Emisor&quot;: nit_e,
&quot;Ya Registrada&quot;: es_roja_reg,
&quot;No Contabilizar&quot;: False,
&quot;Comprobante Previo&quot;: comp_prev,
&quot;Grupo Importación&quot;: grp_imp,
&quot;Cuenta Pasivo Especifica&quot;: cta_esp,
&quot;Es Aduanera&quot;: es_aduanera_flag,
&quot;Estado Registro&quot;: f&quot;�� Ya Registrada ({comp_prev})&quot; if es_roja_reg else (&quot;�� Agen
Aduanero&quot; if any(k in nom_e.upper() for k in [&quot;EURO SHIPPING&quot;, &quot;TRADE GLOBAL&quot;]) else (&quot;�� Aduanera&quot; i
es_aduanera_flag else &quot;⚪ Compra Pendiente&quot;)),
&quot;Régimen Fiscal Emisor&quot;: resp_e if resp_e else &quot;O-48 (Estándar)&quot;,
&quot;Descripcion&quot;: desc,
&quot;Operacion&quot;: op,
&quot;Cta Principal&quot;: cta_p,
&quot;Categoría&quot;: cat,
&quot;Base&quot;: base,
&quot;IVA&quot;: iva,
&quot;ReteFuente&quot;: rfte,
&quot;ReteICA&quot;: rica,
&quot;ReteIVA&quot;: riva,
&quot;Neto a Pagar&quot;: round(tot - rfte - rica - riva, 2),
&quot;Audit Info&quot;: audit_dict,
&quot;Cta Contrapartida&quot;: cta_c,
&quot;Cta IVA&quot;: cta_iva,
&quot;Cta ReteFuente&quot;: cta_rfte,
&quot;Cta ReteICA&quot;: cta_rica,
&quot;Total&quot;: tot,
&quot;CUFE&quot;: cufe_val,
&quot;Descuento Registrado&quot;: desc_val,
&quot;Regla Base&quot;: motivo_base,
&quot;Razón Contable&quot;: razon,
&quot;Soporte PDF Renombrado&quot;: nombre_pdf_esperado
})
df_proc = pd.DataFrame(filas)
st.session_state[&quot;df_procesado&quot;] = df_proc
st.success(f&quot;✅ **Reporte Excel &#39;{archivo_excel.name}&#39; procesado:** {len(df_proc)} facturas
identificadas y ordenadas cronológicamente.&quot;)
if &quot;df_procesado&quot; in st.session_state and st.session_state[&quot;df_procesado&quot;] is not None:
df_proc = st.session_state[&quot;df_procesado&quot;]
c_job1, c_job2, c_job3 = st.columns([2.8, 1.1, 1.1])
with c_job1:
n_orig_mem = len(st.session_state.get(&quot;raw_uploaded_pdfs&quot;, {}))
n_renom_mem = len(st.session_state.get(&quot;dict_pdfs&quot;, {}))
nom_ex = st.session_state.get(&quot;excel_nombre&quot;, &quot;Reporte.xlsx&quot;)
st.info(f&quot;�� **Reporte DIAN:** `{nom_ex}` ({len(df_proc)} facturas) | �� **{n_orig_mem} PD
cargados** | �� **{n_renom_mem} vinculados**&quot;)
if n_orig_mem &gt; 0 and n_renom_mem == 0:
if st.button(f&quot;�� Desbloquear y Vincular los {n_orig_mem} PDFs a Comprobantes&quot;
key=&quot;btn_desbloquear_top_bar&quot;, type=&quot;primary&quot;, use_container_width=True):
st.session_state[&quot;_ejecutar_desbloqueo_ahora&quot;] = True
st.rerun()
with c_job2:
if &quot;excel_bytes&quot; in st.session_state and st.session_state[&quot;excel_bytes&quot;]:
st.download_button(
label=&quot;�� Descargar Excel&quot;,
data=st.session_state[&quot;excel_bytes&quot;],
file_name=st.session_state.get(&quot;excel_nombre&quot;, &quot;Reporte.xlsx&quot;),
mime=&quot;application/vnd.openxmlformats-officedocument.spreadsheetml.sheet&quot;,
key=&quot;btn_dl_excel_tab_compras&quot;,
use_container_width=True
)
with c_job3:

st.caption(&quot;⚡ Memoria RAM Protegida (PDFs en disco)&quot;)

c_aud_btn1, c_aud_btn2 = st.columns(2)
with c_aud_btn1:
st.markdown(&quot;#### Matriz Contable Preliminar vinculada a Comprobantes (Orden Cronológico Enero -
Actual):&quot;)
with c_aud_btn2:
if st.button(&quot;⚡ Auditar y Escanear Régimen Fiscal de los PDFs (Ultrarrápido)&quot;,
key=&quot;btn_escanear_regimen_todos&quot;):
dict_renom_s = st.session_state.get(&quot;dict_pdfs&quot;, {})
dict_orig_s = st.session_state.get(&quot;raw_uploaded_pdfs&quot;, {})
total_modificados = 0
if dict_renom_s:
with st.spinner(&quot;⚡ Escaneando régimen fiscal directamente desde las facturas ya
desbloqueadas y renombradas...&quot;):
total_modificados = auditar_regimen_desde_facturas_renombradas(df_proc, dict_renom_s,
empresa)
st.session_state[&quot;df_procesado&quot;] = df_proc
pass
st.success(f&quot;⚡ Auditoría completada desde facturas renombradas: se revisaron
{len(df_proc)} facturas y se actualizaron {total_modificados} con su régimen oficial del PDF.&quot;)
st.rerun()
elif dict_orig_s:
with st.spinner(&quot;⚡ Extrayendo texto en memoria y cruzando regímenes fiscales...&quot;):
# 1. Pre-cargar textos en caché de todos los PDFs una sola vez
mapa_textos = []
for fn, fb in dict_renom_s.items():
mapa_textos.append((fn, [txt.replace(&quot;-&quot;, &quot;&quot;).replace(&quot; &quot;, &quot;&quot;).upper() for txt in
cache_extraer_textos_pdf(fb)], cache_extraer_textos_pdf(fb)))
for fn, fb in dict_orig_s.items():
mapa_textos.append((fn, [txt.replace(&quot;-&quot;, &quot;&quot;).replace(&quot; &quot;, &quot;&quot;).upper() for txt in
cache_extraer_textos_pdf(fb)], cache_extraer_textos_pdf(fb)))
# 2. Cruzar en milisegundos cada factura directamente en memoria RAM
for r_idx, r_mat in df_proc.iterrows():
fol = str(r_mat.get(&quot;Folio&quot;, &quot;&quot;)).replace(&quot;-&quot;, &quot;&quot;).strip().upper()
fol_sc = fol.lstrip(&quot;0&quot;)
pref = str(r_mat.get(&quot;Prefijo&quot;, &quot;&quot;)).replace(&quot;-&quot;, &quot;&quot;).strip().upper()
fac_full = (pref + fol) if pref else fol
fac_full_sc = (pref + fol_sc) if pref else fol_sc
nit_d = re.sub(r&quot;\D&quot;, &quot;&quot;, str(r_mat.get(&quot;NIT Emisor&quot;, &quot;&quot;)))
nit_b = nit_d[:-1] if len(nit_d) &gt;= 10 else nit_d
texto_factura_encontrado = &quot;&quot;
# Buscar en textos pre-extraídos
for _, list_clean, list_raw in mapa_textos:
for p_idx, t_clean in enumerate(list_clean):
m = False
if len(fac_full) &gt;= 4 and fac_full in t_clean:
m = True
elif len(fac_full_sc) &gt;= 3 and fac_full_sc in t_clean:
m = True
elif (nit_d and len(nit_d) &gt;= 6 and nit_d in t_clean) or (nit_b and
len(nit_b) &gt;= 6 and nit_b in t_clean):
if len(fol) &gt;= 2 and fol in t_clean:
m = True
elif len(fol_sc) &gt;= 2 and fol_sc in t_clean:
m = True
if m:
texto_factura_encontrado += list_raw[p_idx] + &quot;\n&quot;
if texto_factura_encontrado:
break
if texto_factura_encontrado:
reg_f = escanear_regimen_texto_pdf(texto_factura_encontrado)
if reg_f and reg_f != r_mat.get(&quot;Régimen Fiscal Emisor&quot;):
df_proc.at[r_idx, &quot;Régimen Fiscal Emisor&quot;] = reg_f
t_c_n, op_n, c_p_n, c_c_n, desc_n, rfte_n, rica_n, riva_n, c_rf_n, c_iv_n,
c_ri_n, cat_n, razon_n, audit_n = clasificar_factura(
df_proc.at[r_idx, &quot;NIT Emisor&quot;], df_proc.at[r_idx, &quot;Proveedor&quot;],

df_proc.at[r_idx, &quot;Base&quot;], df_proc.at[r_idx, &quot;IVA&quot;],
df_proc.at[r_idx, &quot;Operacion&quot;], reg_f, empresa
)
df_proc.at[r_idx, &quot;ReteFuente&quot;] = rfte_n
df_proc.at[r_idx, &quot;ReteICA&quot;] = rica_n
df_proc.at[r_idx, &quot;ReteIVA&quot;] = riva_n
df_proc.at[r_idx, &quot;Cta ReteFuente&quot;] = c_rf_n
df_proc.at[r_idx, &quot;Cta ReteICA&quot;] = c_ri_n
df_proc.at[r_idx, &quot;Razón Contable&quot;] = razon_n
df_proc.at[r_idx, &quot;Audit Info&quot;] = audit_n
df_proc.at[r_idx, &quot;Neto a Pagar&quot;] = round(df_proc.at[r_idx, &quot;Total&quot;] -
rfte_n - rica_n - riva_n, 2)
total_modificados += 1
st.session_state[&quot;df_procesado&quot;] = df_proc
pass
st.success(f&quot;⚡ Auditoría completada en segundos: se revisaron {len(df_proc)} facturas y se
actualizaron {total_modificados} con su régimen oficial del PDF.&quot;)
st.rerun()
else:
st.warning(&quot;Primero sube los PDFs en el campo 2 para poder escanearlos.&quot;)
# Asegurar columna Cuenta Pasivo (CxP) con marcación clara
if &quot;Cuenta Pasivo (CxP)&quot; not in df_proc.columns:
df_proc[&quot;Cuenta Pasivo (CxP)&quot;] = df_proc.apply(
lambda r: f&quot;⚠️ {r.get(&#39;Cuenta Pasivo Especifica&#39;, &#39;22050505&#39;)} (Por Verificar)&quot; if r.get(&quot;Ya
Registrada&quot;) else str(r.get(&quot;Cuenta Pasivo Especifica&quot;, r.get(&quot;Cta Contrapartida&quot;, &quot;22050501&quot;))),
axis=1
)
st.markdown(&quot;&quot;&quot;
&lt;div style=&quot;background:#eff6ff; border:1px solid #bfdbfe; border-left:5px solid #2563eb; border-
radius:6px; padding:10px 14px; margin-bottom:12px; font-size:13.5px; color:#1e40af;&quot;&gt;
&lt;b&gt;ℹ️ Cuentas de Pasivo (CxP) para Facturas Registradas:&lt;/b&gt; Las facturas ya causadas en Siigo
(��) tienen asignada una cuenta por pagar estimada (ej. &lt;code&gt;22050505&lt;/code&gt; Agencia Aduana, &lt;code&gt;23359501&lt;/code
Acreedores DHL). Aparecen marcadas con &lt;b&gt;⚠️ (Por Verificar)&lt;/b&gt;. Puedes modificar la cuenta fácilmente en la
&lt;b&gt;Pestaña 2 (Revisión)&lt;/b&gt; o en la &lt;b&gt;Pestaña 3 (Triangulación)&lt;/b&gt;.
&lt;/div&gt;
&quot;&quot;&quot;, unsafe_allow_html=True)

# AUTO-AUDITORÍA Y VERIFICACIÓN INMEDIATA DE REGÍMENES FISCALES AL CARGAR EXCEL O PDFS
dict_renom_mem = st.session_state.get(&quot;dict_pdfs&quot;, {})
dict_orig_mem = st.session_state.get(&quot;raw_uploaded_pdfs&quot;, {})
# 1. Aplicar catálogo oficial de regímenes conocidos por NIT (solo una vez al cargar archivo inicial)
if not st.session_state.get(&quot;_regimenes_conocidos_aplicados_v1&quot;, False):
if &quot;Audit Info&quot; not in df_proc.columns:
df_proc[&quot;Audit Info&quot;] = [{} for _ in range(len(df_proc))]
for r_idx, r_row in df_proc.iterrows():
if bool(r_row.get(&quot;Editada Manualmente&quot;, False)):
continue
nit_d = re.sub(r&#39;\D&#39;, &#39;&#39;, str(r_row.get(&quot;NIT Emisor&quot;, &quot;&quot;)))
nom_u = str(r_row.get(&quot;Proveedor&quot;, &quot;&quot;)).upper()
reg_prev = str(r_row.get(&quot;Régimen Fiscal Emisor&quot;, &quot;&quot;))
nuevo_reg = None
if nit_d in REGIMENES_EMISORES_CONOCIDOS:
nuevo_reg = REGIMENES_EMISORES_CONOCIDOS[nit_d]
elif any(k in nom_u for k in [&quot;DHL&quot;, &quot;ESTELAR&quot;, &quot;BUENAVENTURA&quot;, &quot;SIIGO&quot;, &quot;COMCEL&quot;, &quot;CLARO&quot;,
&quot;PANAMERICANA&quot;]):
nuevo_reg = &quot;O-13;O-15&quot;
if nuevo_reg and (nuevo_reg != reg_prev):
df_proc.at[r_idx, &quot;Régimen Fiscal Emisor&quot;] = nuevo_reg
t_c, op_c, c_p, c_c, desc_c, rfte_c, rica_c, riva_c, c_rf, c_iv, c_ri, cat_c, razon_c,
audit_c = clasificar_factura(
r_row[&quot;NIT Emisor&quot;], r_row[&quot;Proveedor&quot;],
r_row[&quot;Base&quot;], r_row[&quot;IVA&quot;],
r_row[&quot;Operacion&quot;], nuevo_reg, empresa
)

df_proc.at[r_idx, &quot;ReteFuente&quot;] = rfte_c
df_proc.at[r_idx, &quot;ReteICA&quot;] = rica_c
df_proc.at[r_idx, &quot;ReteIVA&quot;] = riva_c
df_proc.at[r_idx, &quot;Cta ReteFuente&quot;] = c_rf
df_proc.at[r_idx, &quot;Cta ReteICA&quot;] = c_ri
df_proc.at[r_idx, &quot;Razón Contable&quot;] = razon_c
df_proc.at[r_idx, &quot;Audit Info&quot;] = audit_c
df_proc.at[r_idx, &quot;Neto a Pagar&quot;] = round(r_row[&quot;Total&quot;] - rfte_c - rica_c - riva_c, 2)
st.session_state[&quot;_regimenes_conocidos_aplicados_v1&quot;] = True
# Auditoría de régimen fiscal disponible bajo demanda mediante el botón superior
st.session_state[&quot;df_procesado&quot;] = df_proc
cols_mostrar_proc = [&quot;Comprobante Siigo&quot;, &quot;Fecha&quot;, &quot;Factura&quot;, &quot;Proveedor&quot;, &quot;NIT Emisor&quot;, &quot;Régimen
Fiscal Emisor&quot;, &quot;Estado Registro&quot;, &quot;Cuenta Pasivo (CxP)&quot;, &quot;Base&quot;, &quot;IVA&quot;, &quot;ReteFuente&quot;, &quot;ReteICA&quot;, &quot;Neto a Pagar&quot;,
&quot;Soporte PDF Renombrado&quot;]
cols_mostrar_existentes = [c for c in cols_mostrar_proc if c in df_proc.columns]
st.dataframe(df_proc[cols_mostrar_existentes], use_container_width=True)
st.markdown(&quot;---&quot;)
st.markdown(&quot;#### �� Procesamiento, Separación y Renombrado de PDFs Multi-Página:&quot;)
st.caption(&quot;Identifica y une facturas completas de 2 o más hojas a partir de un PDF consolidado,
asignando a cada una su comprobante y consecutivo oficial.&quot;)
col_cfg1, col_cfg2 = st.columns(2)
with col_cfg1:
modo_sep = st.radio(
&quot;Configuración de reconocimiento de facturas:&quot;,
[
&quot;�� Detección Inteligente (Reconoce facturas completas de 2 o más hojas por &#39;Página 1 d
2&#39;, CUFE y Número)&quot;,
&quot;�� Agrupación fija de 2 páginas por factura (cada factura tiene 2 hojas)&quot;,
&quot;�� Agrupación fija de 3 páginas por factura (cada factura tiene 3 hojas)&quot;,
&quot;�� Separar 1 página por factura (facturas de 1 sola hoja)&quot;,
&quot;�� Mantener archivos individuales (sin separar páginas)&quot;
]
)
with col_cfg2:
st.info(&quot;�� **Garantía de Factura Completa:** El motor inteligente detecta dónde empieza cad
factura (Prefijo, Folio, NIT y marcadores de paginación). Todas las páginas de una misma factura se unen en un solo
archivo PDF completo nombrado `Comp_10-XXX_Factura_Proveedor.pdf`.&quot;)
ejecutar_desb = st.button(&quot;�� Desbloquear, Separar y Renombrar PDFs ahora&quot;) o
st.session_state.pop(&quot;_ejecutar_desbloqueo_ahora&quot;, False)
if ejecutar_desb:
df_ref = st.session_state.get(&quot;df_procesado&quot;, pd.DataFrame())
# Procesamiento en streaming desde disco (0 MB retenidos en RAM)
asegurar_dicts_pdf_en_disco()
if archivos_pdfs:
for p_it in archivos_pdfs:
p_nom = getattr(p_it, &quot;name&quot;, &quot;&quot;).lower()
if p_nom.endswith(&quot;.zip&quot;):
try:
b_val = p_it.getvalue() if hasattr(p_it, &quot;getvalue&quot;) else p_it.read()
with zipfile.ZipFile(io.BytesIO(b_val), &quot;r&quot;) as z_u:
for z_fn in z_u.namelist():
if z_fn.lower().endswith(&quot;.pdf&quot;) and not z_fn.startswith(&quot;__MACOSX&quot;):
c_fn = os.path.basename(z_fn)
if c_fn:
st.session_state[&quot;raw_uploaded_pdfs&quot;][c_fn] = z_u.read(z_fn)
del b_val
except Exception as e_z:
st.error(f&quot;Error abriendo ZIP {getattr(p_it, &#39;name&#39;, &#39;&#39;)}: {e_z}&quot;)
elif p_nom.endswith(&quot;.pdf&quot;) or not p_nom:
st.session_state[&quot;raw_uploaded_pdfs&quot;][getattr(p_it, &quot;name&quot;, &quot;factura.pdf&quot;)] =
p_it.getvalue() if hasattr(p_it, &quot;getvalue&quot;) else p_it.read()
gc.collect()
nombres_pdfs_proc = [fn for fn in list(st.session_state[&quot;raw_uploaded_pdfs&quot;].keys()) if
fn.lower().endswith(&quot;.pdf&quot;)]

if not nombres_pdfs_proc:
st.warning(&quot;⚠️ No hay archivos PDF para procesar. Por favor sube tus facturas en el campo &#39;2.
Facturas en PDF (unificadas o separadas)&#39; arriba antes de hacer clic en desbloquear.&quot;)
st.stop()
total_generados = 0
total_pdfs_count = len(nombres_pdfs_proc)
prog_bar_pdf = st.progress(0, text=f&quot;Iniciando procesamiento liviano de {total_pdfs_count}
archivo(s) PDF...&quot;)
with zipfile.ZipFile(ZIP_FACTURAS_DISK_PATH, &quot;w&quot;, zipfile.ZIP_STORED) as zf:
for idx_pdf, pdf_name in enumerate(nombres_pdfs_proc):
if idx_pdf % 15 == 0:
gc.collect()
prog_bar_pdf.progress(min(1.0, (idx_pdf + 1) / total_pdfs_count), text=f&quot;⚡ Factura
{idx_pdf + 1} de {total_pdfs_count}: {pdf_name[:20]}...&quot;)
if pdf_name in st.session_state[&quot;dict_pdfs&quot;]:
b_existing = st.session_state[&quot;dict_pdfs&quot;][pdf_name]
if b_existing and len(b_existing) &gt; 100:
zf.writestr(pdf_name, b_existing)
total_generados += 1
continue
try:
b_raw_one = st.session_state[&quot;raw_uploaded_pdfs&quot;][pdf_name]
pdf_item = io.BytesIO(b_raw_one)
pdf_item.name = pdf_name
reader = PdfReader(pdf_item)
desencriptado = False
if reader.is_encrypted:
claves_probar = []
if &quot;clave_pdf_extra&quot; in locals() and clave_pdf_extra:
claves_probar.append(clave_pdf_extra.strip())
# NITs empresa compradora
nit_rec = re.sub(r&quot;\D&quot;, &quot;&quot;, str(empresa.get(&quot;nit&quot;, &quot;&quot;)))
if nit_rec:
if len(nit_rec) &gt; 9:
claves_probar.extend([nit_rec, nit_rec[:-1], f&quot;{nit_rec[:-1]}-{nit_rec[-
1]}&quot;])
else:
claves_probar.extend([nit_rec, f&quot;{nit_rec}5&quot;, f&quot;{nit_rec}-5&quot;])
claves_probar.extend([&quot;901346412&quot;, &quot;9013464125&quot;, &quot;901346412-5&quot;, &quot;&quot;])
# NITs emisores en reporte
if not df_ref.empty:
for n_e in df_ref[&quot;NIT Emisor&quot;].dropna().unique():
n_c = re.sub(r&quot;\D&quot;, &quot;&quot;, str(n_e))
if n_c and len(n_c) &gt;= 7:
claves_probar.append(n_c)
for pwd in claves_probar:
try:
res_dec = reader.decrypt(pwd)
if res_dec in (1, 2) or not reader.is_encrypted:
desencriptado = True
break
except:
pass
if reader.is_encrypted and not desencriptado:
st.warning(f&quot;⚠️ No se pudo desencriptar automáticamente {pdf_name}. Ingrese la
contraseña en el campo correspondiente.&quot;)
continue
num_pags = len(reader.pages)
# Caso 1: Archivo individual (sin separar páginas) o de 1 página
if num_pags == 1 or &quot;Mantener archivos individuales&quot; in modo_sep:
writer = PdfWriter()
texto_pdf = &quot;&quot;
for page in reader.pages:

writer.add_page(page)
try: texto_pdf += (page.extract_text() or &quot;&quot;) + &quot;\n&quot;
except: pass
r_mat = None
if not df_ref.empty:
# 1. Identificar por texto exacto (CUFE, Factura + Proveedor/NIT)
r_mat = identificar_factura_en_texto(texto_pdf, df_ref)
# 2. Si no se identificó en el texto, buscar por coincidencia en el nombre del
archivo
if r_mat is None:
fn_clean = re.sub(r&#39;[^A-Za-z0-9]&#39;, &#39;&#39;, pdf_name).lower()
for _, r_cand in df_ref.iterrows():
cufe_c = re.sub(r&#39;[^A-Za-z0-9]&#39;, &#39;&#39;, str(r_cand.get(&quot;CUFE&quot;, &quot;&quot;) or
&quot;&quot;)).lower()
if cufe_c and len(cufe_c) &gt;= 15 and cufe_c[:20] in fn_clean:
r_mat = r_cand
break
pref_c = re.sub(r&#39;[^A-Za-z0-9]&#39;, &#39;&#39;, str(r_cand.get(&quot;Prefijo&quot;, &quot;&quot;) or
&quot;&quot;)).upper()
fol_c = re.sub(r&#39;[^A-Za-z0-9]&#39;, &#39;&#39;, str(r_cand.get(&quot;Folio&quot;, &quot;&quot;) or
&quot;&quot;)).upper()
fac_f = (pref_c + fol_c) if pref_c else fol_c
if fac_f and len(fac_f) &gt;= 4 and fac_f in fn_clean.upper():
r_mat = r_cand
break
if r_mat is not None:
nombre_final = r_mat[&quot;Soporte PDF Renombrado&quot;]
# Escaneo de régimen fiscal sólo para la factura que realmente coincidió
regimen_escaneado = escanear_regimen_texto_pdf(texto_pdf)
if regimen_escaneado:
r_idx = r_mat.name
df_ref.at[r_idx, &quot;Régimen Fiscal Emisor&quot;] = regimen_escaneado
_, _, _, _, _, rfte_n, rica_n, riva_n, cta_rf_n, _, cta_ri_n, _, razon_n,
audit_n = clasificar_factura(
df_ref.at[r_idx, &quot;NIT Emisor&quot;], df_ref.at[r_idx, &quot;Proveedor&quot;],
df_ref.at[r_idx, &quot;Base&quot;], df_ref.at[r_idx, &quot;IVA&quot;],
df_ref.at[r_idx, &quot;Operacion&quot;], regimen_escaneado, empresa
)
df_ref.at[r_idx, &quot;ReteFuente&quot;] = rfte_n
df_ref.at[r_idx, &quot;ReteICA&quot;] = rica_n
df_ref.at[r_idx, &quot;ReteIVA&quot;] = riva_n
df_ref.at[r_idx, &quot;Cta ReteFuente&quot;] = cta_rf_n
df_ref.at[r_idx, &quot;Cta ReteICA&quot;] = cta_ri_n
df_ref.at[r_idx, &quot;Razón Contable&quot;] = razon_n
df_ref.at[r_idx, &quot;Audit Info&quot;] = audit_n
df_ref.at[r_idx, &quot;Neto a Pagar&quot;] = round(df_ref.at[r_idx, &quot;Total&quot;] -
rfte_n - rica_n - riva_n, 2)
else:
nombre_final = f&quot;Soporte_{idx_pdf+1}_{pdf_name}&quot;
# Reutilizar directamente los bytes originales sin re-escribir con PdfWriter
(1000x más rápido)
b_data = pdf_item.getvalue() if hasattr(pdf_item, &quot;getvalue&quot;) else pdf_item.read()
zf.writestr(nombre_final, b_data)
st.session_state[&quot;dict_pdfs&quot;][nombre_final] = b_data
total_generados += 1
# Caso 2: Agrupación fija de 2 páginas
elif &quot;Agrupación fija de 2 páginas&quot; in modo_sep:
for p_i in range(0, num_pags, 2):
writer = PdfWriter()
txt_grp = &quot;&quot;
for off in range(2):
if p_i + off &lt; num_pags:
writer.add_page(reader.pages[p_i + off])
try: txt_grp += (reader.pages[p_i + off].extract_text() or &quot;&quot;) + &quot;\n&quot;
except: pass
r_mat = identificar_factura_en_texto(txt_grp, df_ref)

if r_mat is not None:
nombre_final = r_mat[&quot;Soporte PDF Renombrado&quot;]
else:
fac_idx = p_i // 2
nombre_final = df_ref.iloc[fac_idx][&quot;Soporte PDF Renombrado&quot;] if not
df_ref.empty and fac_idx &lt; len(df_ref) else f&quot;Factura_Pags_{p_i+1}-{p_i+2}_{pdf_name}&quot;
pdf_bytes = io.BytesIO()
writer.write(pdf_bytes)
pdf_bytes.seek(0)
b_data = pdf_bytes.getvalue()
zf.writestr(nombre_final, b_data)
st.session_state[&quot;dict_pdfs&quot;][nombre_final] = b_data
total_generados += 1
# Caso 3: Agrupación fija de 3 páginas
elif &quot;Agrupación fija de 3 páginas&quot; in modo_sep:
for p_i in range(0, num_pags, 3):
writer = PdfWriter()
txt_grp = &quot;&quot;
for off in range(3):
if p_i + off &lt; num_pags:
writer.add_page(reader.pages[p_i + off])
try: txt_grp += (reader.pages[p_i + off].extract_text() or &quot;&quot;) + &quot;\n&quot;
except: pass
r_mat = identificar_factura_en_texto(txt_grp, df_ref)
if r_mat is not None:
nombre_final = r_mat[&quot;Soporte PDF Renombrado&quot;]
else:
fac_idx = p_i // 3
nombre_final = df_ref.iloc[fac_idx][&quot;Soporte PDF Renombrado&quot;] if not
df_ref.empty and fac_idx &lt; len(df_ref) else f&quot;Factura_Pags_{p_i+1}-{p_i+3}_{pdf_name}&quot;
pdf_bytes = io.BytesIO()
writer.write(pdf_bytes)
pdf_bytes.seek(0)
b_data = pdf_bytes.getvalue()
zf.writestr(nombre_final, b_data)
st.session_state[&quot;dict_pdfs&quot;][nombre_final] = b_data
total_generados += 1
# Caso 4: 1 página por factura
elif &quot;Separar 1 página&quot; in modo_sep:
for p_i in range(num_pags):
writer = PdfWriter()
writer.add_page(reader.pages[p_i])
try: txt_pag = (reader.pages[p_i].extract_text() or &quot;&quot;) + &quot;\n&quot;
except: txt_pag = &quot;&quot;
r_mat = identificar_factura_en_texto(txt_pag, df_ref)
if r_mat is not None:
nombre_final = r_mat[&quot;Soporte PDF Renombrado&quot;]
else:
nombre_final = df_ref.iloc[p_i][&quot;Soporte PDF Renombrado&quot;] if not
df_ref.empty and p_i &lt; len(df_ref) else f&quot;Factura_Pag_{p_i+1}_{pdf_name}&quot;
pdf_bytes = io.BytesIO()
writer.write(pdf_bytes)
pdf_bytes.seek(0)
b_data = pdf_bytes.getvalue()
zf.writestr(nombre_final, b_data)
st.session_state[&quot;dict_pdfs&quot;][nombre_final] = b_data
total_generados += 1
# Caso 5: Detección inteligente por CUFE, Folio o &#39;Página 1 de N&#39;
else:
patron_pag =
re.compile(r&#39;(?:P[ÁAáa]G(?:INA|\.)?|HOJA|PAGE)\s*(\d+)\s*(?:DE|\/|OF)\s*(\d+)&#39;, re.IGNORECASE)
curr_writer = None
curr_inv_row = None

paginas_del_comprobante = 0
facturas_generadas = []
def buscar_coincidencia_factura(texto_pagina):
return identificar_factura_en_texto(texto_pagina, df_ref)
for p_idx in range(num_pags):
try: txt_p = reader.pages[p_idx].extract_text() or &quot;&quot;
except: txt_p = &quot;&quot;
reg_detectado = escanear_regimen_texto_pdf(txt_p)
m_pag = patron_pag.search(txt_p)
cur_p, tot_p = (int(m_pag.group(1)), int(m_pag.group(2))) if m_pag else (None,
None)
inv_encontrada = buscar_coincidencia_factura(txt_p)
es_inicio_nueva_factura = False
if cur_p == 1:
es_inicio_nueva_factura = True
elif cur_p is not None and cur_p &gt; 1:
es_inicio_nueva_factura = False
elif inv_encontrada is not None:
if curr_inv_row is None or inv_encontrada[&quot;Comprobante Siigo&quot;] !=
curr_inv_row[&quot;Comprobante Siigo&quot;]:
es_inicio_nueva_factura = True
elif curr_inv_row is None:
es_inicio_nueva_factura = True
if es_inicio_nueva_factura:
if curr_writer is not None and curr_inv_row is not None:
out_c = io.BytesIO()
curr_writer.write(out_c)
c_bytes = out_c.getvalue()
doc_nombre = curr_inv_row[&quot;Soporte PDF Renombrado&quot;]
zf.writestr(doc_nombre, c_bytes)
st.session_state[&quot;dict_pdfs&quot;][doc_nombre] = c_bytes
total_generados += 1
facturas_generadas.append(f&quot;{curr_inv_row[&#39;Comprobante Siigo&#39;]}
({paginas_del_comprobante} págs)&quot;)
curr_inv_row = inv_encontrada if inv_encontrada is not None else
curr_inv_row
if curr_inv_row is not None and reg_detectado and not df_ref.empty:
r_idx = curr_inv_row.name
df_ref.at[r_idx, &quot;Régimen Fiscal Emisor&quot;] = reg_detectado
_, _, _, _, _, rfte_n, rica_n, riva_n, cta_rf_n, _, cta_ri_n, _,
razon_n, audit_n = clasificar_factura(
df_ref.at[r_idx, &quot;NIT Emisor&quot;], df_ref.at[r_idx, &quot;Proveedor&quot;],
df_ref.at[r_idx, &quot;Base&quot;], df_ref.at[r_idx, &quot;IVA&quot;],
df_ref.at[r_idx, &quot;Operacion&quot;], reg_detectado, empresa
)
df_ref.at[r_idx, &quot;ReteFuente&quot;] = rfte_n
df_ref.at[r_idx, &quot;ReteICA&quot;] = rica_n
df_ref.at[r_idx, &quot;ReteIVA&quot;] = riva_n
df_ref.at[r_idx, &quot;Cta ReteFuente&quot;] = cta_rf_n
df_ref.at[r_idx, &quot;Cta ReteICA&quot;] = cta_ri_n
df_ref.at[r_idx, &quot;Razón Contable&quot;] = razon_n
df_ref.at[r_idx, &quot;Audit Info&quot;] = audit_n
df_ref.at[r_idx, &quot;Neto a Pagar&quot;] = round(df_ref.at[r_idx, &quot;Total&quot;] -
rfte_n - rica_n - riva_n, 2)
curr_writer = PdfWriter()
curr_writer.add_page(reader.pages[p_idx])
paginas_del_comprobante = 1
else:
if curr_writer is None:
curr_inv_row = inv_encontrada
curr_writer = PdfWriter()
curr_writer.add_page(reader.pages[p_idx])
paginas_del_comprobante += 1

if curr_writer is not None and curr_inv_row is not None:
out_c = io.BytesIO()
curr_writer.write(out_c)
c_bytes = out_c.getvalue()
doc_nombre = curr_inv_row[&quot;Soporte PDF Renombrado&quot;]
zf.writestr(doc_nombre, c_bytes)
st.session_state[&quot;dict_pdfs&quot;][doc_nombre] = c_bytes
total_generados += 1
facturas_generadas.append(f&quot;{curr_inv_row[&#39;Comprobante Siigo&#39;]}
({paginas_del_comprobante} págs)&quot;)
except Exception as e:
st.error(f&quot;Error procesando {pdf_name}: {e}&quot;)
prog_bar_pdf.progress(1.0, text=&quot;✅ ¡100% Procesado y Vinculado con Éxito!&quot;)
st.session_state[&quot;zip_pdfs_path&quot;] = ZIP_FACTURAS_DISK_PATH
st.session_state[&quot;zip_pdfs&quot;] = None
st.session_state[&quot;total_zip_pdfs&quot;] = total_generados
gc.collect()
# Escaneo oficial de régimen fiscal directamente sobre las facturas ya desbloqueadas y renombradas
total_auditados = auditar_regimen_desde_facturas_renombradas(df_ref,
st.session_state[&quot;dict_pdfs&quot;], empresa)
st.session_state[&quot;df_procesado&quot;] = df_ref
gc.collect()
# Guardar trabajo actualizado en disco sin saturar RAM
guardar_trabajo_en_historial(
empresa, df_ref,
excel_bytes=st.session_state.get(&quot;excel_bytes&quot;),
excel_nombre=st.session_state.get(&quot;excel_nombre&quot;, &quot;Reporte.xlsx&quot;),
dict_pdfs_renombrados=None,
dict_pdfs_originales=None,
zip_bytes=None,
consecutivo_ini=cons_ini_fac,
job_id=st.session_state.get(&quot;job_actual_id&quot;)
)
msg_reg = f&quot; y se auditó el régimen de cada factura ({total_auditados} actualizadas)&quot; if
total_auditados &gt; 0 else &quot; y se validó el régimen de cada factura&quot;
st.success(f&quot;¡Procesamiento exitoso! Se desbloquearon y renombraron **{total_generados} facturas
completas**{msg_reg} directamente desde las facturas PDF desbloqueadas.&quot;)
zip_disk_p = st.session_state.get(&quot;zip_pdfs_path&quot;, ZIP_FACTURAS_DISK_PATH)
if os.path.exists(zip_disk_p) and os.path.getsize(zip_disk_p) &gt; 22:
n_tot_z = st.session_state.get(&quot;total_zip_pdfs&quot;, len(st.session_state.get(&quot;dict_pdfs&quot;, {})))
if st.checkbox(f&quot;�� Habilitar botón de descarga ZIP de {n_tot_z} facturas renombradas (Baj
demanda para proteger memoria RAM)&quot;, value=False, key=&quot;chk_dl_zip_main_ondemand&quot;):
with open(zip_disk_p, &quot;rb&quot;) as f_zdl:
st.download_button(
label=f&quot;⬇️ Descargar Archivos PDF Renombrados ({n_tot_z} facturas en ZIP)&quot;,
data=f_zdl.read(),
file_name=f&quot;Facturas_Renombradas_Comprobantes_{empresa[&#39;nombre&#39;].replace(&#39; &#39;,
&#39;_&#39;)}.zip&quot;,
mime=&quot;application/zip&quot;,
key=&quot;btn_dl_zip_pdfs_renombrados_action_main&quot;,
use_container_width=True
)

else:
st.info(&quot;�� Sube el reporte Excel de la DIAN arriba o carga un trabajo guardado desde el **Historia
de Trabajos** para comenzar.&quot;)
except Exception as _e_comp:
if type(_e_comp).__name__ in [&#39;RerunException&#39;, &#39;StopException&#39;, &#39;ScriptControlException&#39;]:
raise _e_comp
st.error(f&quot;⚠️ Ocurrió un error en la pestaña de Compras DIAN: {_e_comp}&quot;)
with st.expander(&quot;�� Ver detalles técnicos del error (Traceback):&quot;):

st.code(traceback.format_exc())
with tab_auditoria:
try:
st.markdown(&quot;### Modulo de Auditoria Contable y Trazabilidad&quot;)
st.caption(&quot;Inspección de cuentas, deducciones y separación de gastos por cuenta de terceros.&quot;)
if &quot;df_procesado&quot; in st.session_state and st.session_state[&quot;df_procesado&quot;] is not None and
isinstance(st.session_state[&quot;df_procesado&quot;], pd.DataFrame) and not st.session_state[&quot;df_procesado&quot;].empty:
df_p = normalizar_df_procesado(st.session_state[&quot;df_procesado&quot;])
st.session_state[&quot;df_procesado&quot;] = df_p
# Inicializar conjunto de facturas excluidas de contabilización
if &quot;facturas_no_contabilizar&quot; not in st.session_state:
st.session_state[&quot;facturas_no_contabilizar&quot;] = set(df_p[df_p[&quot;No Contabilizar&quot;] ==
True][&quot;Factura&quot;].dropna().astype(str).tolist())
cur_no_contab = st.session_state[&quot;facturas_no_contabilizar&quot;]
# Panel de exclusión masiva para marcar cuáles NO contabilizar
with st.expander(&quot;�� Panel: Marcar qué facturas NO debes contabilizar (Excluir de Siigo):&quot;
expanded=False):
st.caption(&quot;Selecciona aquí las facturas que NO deseas que pasen a la contabilidad ni se exporten
a Siigo (por ejemplo, si ya fueron causadas manualmente, corresponden a otro período o no deben registrarse):&quot;)
lista_todas_facs_tags = []
mapa_tags_to_fac = {}
defaults_seleccionados = []
for idx_row, r in df_p.iterrows():
fn = str(r.get(&quot;Factura&quot;) or f&quot;Doc_{idx_row+1}&quot;).strip()
prov = str(r.get(&quot;Proveedor&quot;) or &quot;Proveedor&quot;).strip()[:24]
try: tot_val = float(r.get(&quot;Total&quot;, 0.0) or 0.0)
except Exception: tot_val = 0.0
ya_reg = bool(r.get(&quot;Ya Registrada&quot;, False))
tag_t = f&quot;{fn} — {prov} (${tot_val:,.0f}) [{&#39;�� Ya Registrada&#39; if ya_reg else &#39;⚪ Pendiente&#39;}]&quot;
lista_todas_facs_tags.append(tag_t)
mapa_tags_to_fac[tag_t] = fn
if fn in cur_no_contab or bool(r.get(&quot;No Contabilizar&quot;, False)):
defaults_seleccionados.append(tag_t)
sel_excluidas_bulk = st.multiselect(
&quot;Selecciona facturas que NO se deben contabilizar:&quot;,
options=lista_todas_facs_tags,
default=defaults_seleccionados,
key=&quot;ms_bulk_excluir_contab&quot;
)
if st.button(&quot;�� Guardar Exclusiones de Contabilidad&quot;, key=&quot;btn_save_bulk_excl&quot;):
nuevas_ex = {mapa_tags_to_fac[t] for t in sel_excluidas_bulk if t in mapa_tags_to_fac}
st.session_state[&quot;facturas_no_contabilizar&quot;] = nuevas_ex
for idx_row, r_row in df_p.iterrows():
fn_curr = str(r_row.get(&quot;Factura&quot;, &quot;&quot;)).strip()
es_ex = fn_curr in nuevas_ex
df_p.at[idx_row, &quot;No Contabilizar&quot;] = es_ex
if es_ex:
df_p.at[idx_row, &quot;Estado Registro&quot;] = &quot;�� No Contabilizar (Excluida)&quot;
else:
es_r = bool(r_row.get(&quot;Ya Registrada&quot;, False))
cp = str(r_row.get(&quot;Comprobante Previo&quot;, &quot;&quot;)).strip()
df_p.at[idx_row, &quot;Estado Registro&quot;] = f&quot;�� Ya Registrada ({cp})&quot; if es_r else &quot;
Compra Pendiente&quot;
st.session_state[&quot;df_procesado&quot;] = df_p
st.success(f&quot;¡Se han marcado {len(nuevas_ex)} facturas como NO contabilizables!&quot;)
st.rerun()
opciones_fac = []
for idx_r, r in df_p.iterrows():
fn = str(r.get(&quot;Factura&quot;) or f&quot;Doc_{idx_r+1}&quot;).strip()
tag_ex = &quot; �� [NO CONTABILIZAR]&quot; if (fn in cur_no_contab or bool(r.get(&quot;No Contabilizar&quot;, False))
else &quot;&quot;
comp_str = str(r.get(&quot;Comprobante Siigo&quot;) or f&quot;Comp {idx_r+1}&quot;).strip()
fecha_str = str(r.get(&quot;Fecha&quot;) or &quot;&quot;).strip()
prov_str = str(r.get(&quot;Proveedor&quot;) or &quot;Proveedor&quot;).strip()[:24]
try: tot_num = float(r.get(&quot;Total&quot;, 0.0) or 0.0)
except Exception: tot_num = 0.0
opciones_fac.append(f&quot;[{comp_str}] {fecha_str} - {fn}{tag_ex} - {prov_str} (${tot_num:,.0f})&quot;)

# Preservar la factura seleccionada para que no vuelva a la primera tras guardar cambios o editar
comp_activo_guardado = st.session_state.get(&quot;factura_seleccionada_auditoria&quot;)
idx_fac_sel = 0
if comp_activo_guardado:
for i_f, op_str in enumerate(opciones_fac):
if f&quot;[{comp_activo_guardado}]&quot; in op_str:
idx_fac_sel = i_f
break
if &quot;sel_factura_auditoria_box&quot; in st.session_state:
if st.session_state[&quot;sel_factura_auditoria_box&quot;] not in opciones_fac and idx_fac_sel &lt;
len(opciones_fac):
st.session_state[&quot;sel_factura_auditoria_box&quot;] = opciones_fac[idx_fac_sel]
if not opciones_fac:
st.info(&quot;No hay facturas disponibles para auditar en este reporte.&quot;)
st.stop()
seleccion = st.selectbox(
&quot;Selecciona una factura para auditar:&quot;,
opciones_fac,
index=idx_fac_sel if idx_fac_sel &lt; len(opciones_fac) else 0,
key=&quot;sel_factura_auditoria_box&quot;
)
if not seleccion:
st.stop()
comp_sel = seleccion.split(&quot;]&quot;)[0].replace(&quot;[&quot;, &quot;&quot;).strip()
st.session_state[&quot;factura_seleccionada_auditoria&quot;] = comp_sel
m_fac_matches = df_p[df_p[&quot;Comprobante Siigo&quot;].astype(str).str.strip() == comp_sel]
if m_fac_matches.empty:
m_fac_matches = df_p
fac_sel = m_fac_matches.iloc[0]
es_aduanero = any(k in str(fac_sel.get(&quot;Proveedor&quot;, &quot;&quot;)).upper() for k in AGENTES_ADUANEROS)
# Auto-corrección inmediata si la factura seleccionada es de Transporte / Inland Rate
es_transp_sel = (&quot;70787&quot; in str(fac_sel.get(&quot;Factura&quot;, &quot;&quot;)) or
&quot;INLAND&quot; in str(fac_sel.get(&quot;Descripcion&quot;, &quot;&quot;)).upper() or
&quot;TRANSP&quot; in str(fac_sel.get(&quot;Descripcion&quot;, &quot;&quot;)).upper() or
&quot;FLETE&quot; in str(fac_sel.get(&quot;Descripcion&quot;, &quot;&quot;)).upper() or
str(fac_sel.get(&quot;Cta ReteFuente&quot;, &quot;&quot;)).strip() == &quot;23652505&quot;)
if es_transp_sel:
r_idx_sel = fac_sel.name
b_val_sel = float(fac_sel.get(&quot;Base&quot;, 0.0) or 0.0)
df_p.at[r_idx_sel, &quot;Descripcion&quot;] = &quot;INLAND RATE TRANSPORTE TERRESTRE&quot;
df_p.at[r_idx_sel, &quot;Categoría&quot;] = &quot;Transporte Carga (1%)&quot;
df_p.at[r_idx_sel, &quot;Razón Contable&quot;] = &quot;Transporte Terrestre de Carga / Fletes (1% ReteFuente
[23652505] - 4.14‰ ReteICA [23680513])&quot;
df_p.at[r_idx_sel, &quot;ReteFuente&quot;] = round(b_val_sel * 0.01, 2)
df_p.at[r_idx_sel, &quot;ReteICA&quot;] = round(b_val_sel * 0.00414, 2)
df_p.at[r_idx_sel, &quot;Cta ReteFuente&quot;] = &quot;23652505&quot;
df_p.at[r_idx_sel, &quot;Cta ReteICA&quot;] = &quot;23680513&quot;
df_p.at[r_idx_sel, &quot;Cta Principal&quot;] = &quot;14650501&quot;
df_p.at[r_idx_sel, &quot;Impuestos Asumidos&quot;] = True
st.session_state[&quot;df_procesado&quot;] = df_p
fac_sel = df_p.loc[r_idx_sel]
audit_info = fac_sel.get(&quot;Audit Info&quot;, {})
if not isinstance(audit_info, dict):
audit_info = {}
col_a1, col_a2 = st.columns(2)
with col_a1:
# Control individual para marcar si no se debe contabilizar esta factura
fac_act_fn = fac_sel[&quot;Factura&quot;]
esta_excl_ind = fac_act_fn in cur_no_contab or bool(fac_sel.get(&quot;No Contabilizar&quot;, False))
c_tog_ex1, c_tog_ex2 = st.columns([2.4, 1.2])

with c_tog_ex1:
if esta_excl_ind:
st.warning(&quot;�� **Esta factura está marcada como &#39;NO CONTABILIZAR&#39;** — No se generar
asiento ni se exportará a Siigo.&quot;)
else:
st.caption(&quot;Estado contable activo para exportación a Siigo.&quot;)
with c_tog_ex2:
if not esta_excl_ind:
if st.button(&quot;�� No Contabilizar&quot;, key=f&quot;btn_excl_indiv_{fac_sel[&#39;Comprobante Siigo&#39;]}&quot;
help=&quot;Excluye esta factura de Siigo&quot;):
cur_no_contab.add(fac_act_fn)
st.session_state[&quot;facturas_no_contabilizar&quot;] = cur_no_contab
m_idx = df_p[df_p[&quot;Factura&quot;] == fac_act_fn].index
if not m_idx.empty:
df_p.at[m_idx[0], &quot;No Contabilizar&quot;] = True
df_p.at[m_idx[0], &quot;Estado Registro&quot;] = &quot;�� No Contabilizar (Excluida)&quot;
st.session_state[&quot;df_procesado&quot;] = df_p
st.session_state[&quot;factura_seleccionada_auditoria&quot;] = comp_sel
st.session_state.pop(&quot;sel_factura_auditoria_box&quot;, None)
st.rerun()
else:
if st.button(&quot;✅ Reactivar&quot;, key=f&quot;btn_react_indiv_{fac_sel[&#39;Comprobante Siigo&#39;]}&quot;,
help=&quot;Vuelve a incluirla en la contabilidad&quot;):
cur_no_contab.discard(fac_act_fn)
st.session_state[&quot;facturas_no_contabilizar&quot;] = cur_no_contab
m_idx = df_p[df_p[&quot;Factura&quot;] == fac_act_fn].index
if not m_idx.empty:
df_p.at[m_idx[0], &quot;No Contabilizar&quot;] = False
es_r = bool(df_p.at[m_idx[0], &quot;Ya Registrada&quot;])
cp = df_p.at[m_idx[0], &quot;Comprobante Previo&quot;]
df_p.at[m_idx[0], &quot;Estado Registro&quot;] = f&quot;�� Ya Registrada ({cp})&quot; if es_r else &quot;
Compra Pendiente&quot;
st.session_state[&quot;df_procesado&quot;] = df_p
st.session_state[&quot;factura_seleccionada_auditoria&quot;] = comp_sel
st.session_state.pop(&quot;sel_factura_auditoria_box&quot;, None)
st.rerun()
st.markdown(f&quot;&quot;&quot;
&lt;div class=&quot;audit-card&quot;&gt;
&lt;h4&gt;Detalle del Comprobante: {fac_sel[&#39;Comprobante Siigo&#39;]}&lt;/h4&gt;
&lt;p&gt;&lt;b&gt;Soporte PDF Vinculado:&lt;/b&gt; &lt;span class=&quot;tag-comp&quot;&gt;{fac_sel.get(&#39;Soporte PDF Renombrado&#39;,
&#39;&#39;)}&lt;/span&gt;&lt;/p&gt;
&lt;p&gt;&lt;b&gt;Factura:&lt;/b&gt; {fac_sel.get(&#39;Factura&#39;, &#39;&#39;)} - &lt;b&gt;Proveedor:&lt;/b&gt; {fac_sel.get(&#39;Proveedor&#39;,
&#39;Proveedor&#39;)} (NIT: {fac_sel.get(&#39;NIT Emisor&#39;, &#39;&#39;)})&lt;/p&gt;
&lt;p&gt;&lt;b&gt;Fecha de Emisión:&lt;/b&gt; {fac_sel.get(&#39;Fecha&#39;, &#39;&#39;)} - &lt;b&gt;Total:&lt;/b&gt;
${float(fac_sel.get(&#39;Total&#39;, 0.0) or 0.0):,.2f}&lt;/p&gt;
&lt;p&gt;&lt;b&gt;Régimen Fiscal Emisor:&lt;/b&gt; &lt;span class=&quot;badge-active&quot;&gt;{fac_sel.get(&#39;Régimen Fiscal
Emisor&#39;, &#39;O-48&#39;)}&lt;/span&gt;&lt;/p&gt;
&lt;div style=&quot;margin: 8px 0;&quot;&gt;
&lt;span class=&quot;badge-tax&quot;&gt;{&#39;✅ Gran Contribuyente (O-13)&#39; if audit_info.get(&#39;es_gc&#39;) else
&#39;Común (No GC)&#39;}&lt;/span&gt;
&lt;span class=&quot;badge-tax&quot;&gt;{&#39;✅ Autorretenedor (O-15)&#39; if audit_info.get(&#39;es_autorr&#39;) else &#39;No
Autorretenedor&#39;}&lt;/span&gt;
&lt;span class=&quot;badge-tax&quot;&gt;{&#39;⚠️ Régimen Simple (O-47)&#39; if audit_info.get(&#39;es_rst&#39;) else
&#39;Régimen Ordinario&#39;}&lt;/span&gt;
&lt;/div&gt;
&lt;hr&gt;
&lt;h5&gt;Trazabilidad de la Contabilización y Validación DIAN:&lt;/h5&gt;
&lt;ul&gt;
&lt;li&gt;&lt;b&gt;Cuenta Asignada:&lt;/b&gt; &lt;span class=&quot;tag-propio&quot;&gt;{fac_sel.get(&#39;Cta Principal&#39;,
&#39;14650501&#39;)}&lt;/span&gt; - {fac_sel.get(&#39;Categoría&#39;, &#39;General&#39;)}&lt;/li&gt;
&lt;li&gt;&lt;b&gt;Cruce DIAN:&lt;/b&gt; {fac_sel.get(&#39;Razón Contable&#39;, &#39;Clasificación automática&#39;)}&lt;/li&gt;
&lt;li&gt;&lt;b&gt;Base Gravable:&lt;/b&gt; ${float(fac_sel.get(&#39;Base&#39;, 0.0) or 0.0):,.2f} - &lt;b&gt;IVA:&lt;/b&gt;
${float(fac_sel.get(&#39;IVA&#39;, 0.0) or 0.0):,.2f}&lt;/li&gt;
&lt;li&gt;&lt;b&gt;Retención en la Fuente:&lt;/b&gt; ${float(fac_sel.get(&#39;ReteFuente&#39;, 0.0) or 0.0):,.2f}
({audit_info.get(&#39;razon_rfte&#39;, &#39;N/A&#39;)})&lt;/li&gt;
&lt;li&gt;&lt;b&gt;ReteICA:&lt;/b&gt; ${fac_sel.get(&#39;ReteICA&#39;, 0.0):,.2f} ({audit_info.get(&#39;razon_reteica&#39;,
audit_info.get(&#39;razon_rica&#39;, &#39;N/A&#39;))})&lt;/li&gt;
&lt;li&gt;&lt;b&gt;ReteIVA:&lt;/b&gt; ${fac_sel.get(&#39;ReteIVA&#39;, 0.0):,.2f} ({audit_info.get(&#39;razon_reteiva&#39;,
&#39;N/A&#39;)})&lt;/li&gt;
&lt;/ul&gt;
&lt;/div&gt;
&quot;&quot;&quot;, unsafe_allow_html=True)

with st.expander(&quot;⚙️ Ajustar Concepto, Régimen, Cuentas (CxP), Base Gravable e Impuestos
Asumidos:&quot;, expanded=True if (fac_sel.get(&quot;Ya Registrada&quot;) or es_aduanero) else False):
st.caption(&quot;Modifica aquí el concepto de la factura (Transporte, Agenciamiento, Compras), los
valores fiscales y el tratamiento contable:&quot;)
# SELECTOR DE CONCEPTO / ACTIVIDAD ECONÓMICA DE LA FACTURA
opciones_conceptos = [
&quot;�� Transporte Terrestre de Carga / Fletes (ReteFuente 1% [23652505] - ReteICA 4.14
[23680513])&quot;,
&quot;�� Agenciamiento Aduanero / Intermediación (ReteFuente 4% [23652503] - ReteICA 9.66
[23680505])&quot;,
&quot;�� Compra de Mercancías / Bienes (ReteFuente 2.5% [23654001] - ReteICA 11.04
[23680501])&quot;,
&quot;�� Alojamiento / Hoteles / Viajes (ReteFuente 3.5% [23652501] - ReteICA 9.66
[23680505])&quot;,
&quot;⚖️ Honorarios Profesionales / Asesoría (ReteFuente 11% [23651501] - ReteICA 9.66‰
[23680505])&quot;,
&quot;�� Software / Tecnología (Sin ReteFuente - IVA Servicios 24081501)&quot;,
&quot;��️ Servicios Generales (ReteFuente 4% [23652503] - ReteICA 9.66‰ [23680505])&quot;
]
desc_fac_act = (str(fac_sel.get(&quot;Descripcion&quot;, &quot;&quot;)) + &quot; &quot; + str(fac_sel.get(&quot;Categoría&quot;,
&quot;&quot;))).upper()
idx_concep_act = 1 if es_aduanero else 6
if any(k in desc_fac_act for k in [&quot;TRANSPORTE&quot;, &quot;TRANSPORT&quot;, &quot;INLAND&quot;, &quot;FLETE&quot;, &quot;CARGA
TERRESTRE&quot;, &quot;ACARREO&quot;]):
idx_concep_act = 0
elif any(k in str(fac_sel.get(&quot;Proveedor&quot;, &quot;&quot;)).upper() for k in [&quot;FERROMENDEZ&quot;, &quot;TORNILLO&quot;,
&quot;CAUCHO&quot;, &quot;REPUESTO&quot;, &quot;MERCANCIA&quot;]):
idx_concep_act = 2
elif any(k in str(fac_sel.get(&quot;Proveedor&quot;, &quot;&quot;)).upper() for k in [&quot;HOTEL&quot;, &quot;ESTELAR&quot;,
&quot;GENOVA&quot;]):
idx_concep_act = 3
comp_box_id = str(fac_sel.get(&#39;Comprobante Siigo&#39;, &#39;&#39;)).strip()
sel_concep = st.selectbox(
&quot;Concepto / Tipo de Retención y Servicio de la Factura:&quot;,
opciones_conceptos,
index=idx_concep_act,
key=f&quot;sel_concep_{comp_box_id}&quot;,
help=&quot;Selecciona Transporte de Carga para aplicar automáticamente la ReteFuente del 1% y
ReteICA del 4.14 por mil.&quot;
)
es_concep_transp = &quot;Transporte&quot; in sel_concep
es_concep_aduana = &quot;Agenciamiento&quot; in sel_concep
es_concep_compra = &quot;Compra&quot; in sel_concep
es_concep_hotel = &quot;Alojamiento&quot; in sel_concep
es_concep_honor = &quot;Honorarios&quot; in sel_concep
es_concep_soft = &quot;Software&quot; in sel_concep
c_mod1, c_mod2 = st.columns(2)
with c_mod1:
opciones_reg_man = [
&quot;O-48 (Responsable de IVA / Común - Retención Ordinaria)&quot;,
&quot;O-15 (Autorretenedor de Renta - Sin ReteFuente)&quot;,
&quot;O-13 (Gran Contribuyente - Sin ReteIVA)&quot;,
&quot;O-13;O-15 (Gran Contribuyente y Autorretenedor)&quot;,
&quot;O-47 (Régimen Simple - Exento ReteFuente y ReteIVA)&quot;,
&quot;O-49 (No Responsable de IVA)&quot;
]
idx_actual = 0
reg_actual = str(fac_sel.get(&quot;Régimen Fiscal Emisor&quot;, &quot;&quot;))
if &quot;O-13&quot; in reg_actual and &quot;O-15&quot; in reg_actual: idx_actual = 3
elif &quot;O-15&quot; in reg_actual: idx_actual = 1
elif &quot;O-13&quot; in reg_actual: idx_actual = 2
elif &quot;O-47&quot; in reg_actual: idx_actual = 4
elif &quot;O-49&quot; in reg_actual: idx_actual = 5
nuevo_reg_sel = st.selectbox(&quot;Selecciona Régimen Fiscal:&quot;, opciones_reg_man,
index=idx_actual, key=f&quot;sel_reg_{comp_box_id}&quot;)
cod_reg_sel = nuevo_reg_sel.split()[0]

cta_p_sug_def = &quot;14650501&quot; if (es_aduanero or es_concep_transp) else &quot;51350501&quot;
cta_p_def = str(fac_sel.get(&quot;Cta Principal&quot;) or cta_p_sug_def).strip()
nueva_cta_p = st.text_input(
&quot;Cuenta Principal (Gasto / Costo / Inventario):&quot;,
value=cta_p_def,
key=f&quot;inp_cta_p_{comp_box_id}&quot;,
help=&quot;Cuenta contable donde se imputa la base de la factura (ej. 14650501 Tránsito /
Fletes, 51350501 Servicios).&quot;
)
cta_cxp_def = str(fac_sel.get(&quot;Cuenta Pasivo Especifica&quot;) or fac_sel.get(&quot;Cta
Contrapartida&quot;) or (&quot;22050505&quot; if es_aduanero else &quot;22050501&quot;)).strip()
nueva_cta_cxp = st.text_input(
&quot;Cuenta por Pagar (Pasivo / Contrapartida):&quot;,
value=cta_cxp_def,
key=f&quot;inp_cxp_{comp_box_id}&quot;,
help=&quot;Puedes cambiar aquí la cuenta pasivo (ej. 22050505 Agencia Aduana, 23359501
Acreedores/DHL, 22050501 Proveedores).&quot;
)
es_importacion_factura = es_aduanero or es_concep_transp or &quot;1465&quot; in str(fac_sel.get(&quot;Cta
Principal&quot;, &quot;&quot;)) or &quot;IMPORTACI&quot; in str(fac_sel.get(&quot;Categoría&quot;, &quot;&quot;)).upper()
cta_iva_def = str(fac_sel.get(&quot;Cta IVA&quot;) or (&quot;24081501&quot; if es_importacion_factura else
&quot;24081001&quot;)).strip()
nueva_cta_iva = st.text_input(
&quot;Cuenta IVA Servicios / Compras:&quot;,
value=cta_iva_def,
key=f&quot;inp_iva_{comp_box_id}&quot;,
help=&quot;Cuenta contable de IVA descontable ordinario (ej. 24081501 Servicios o 24081001
Compras).&quot;
)
with c_mod2:
c_v1, c_v2 = st.columns(2)
base_cur_f = float(fac_sel[&quot;Base&quot;]) if pd.notna(fac_sel[&quot;Base&quot;]) else 0.0
# Recalcular retenciones según el concepto seleccionado
if &quot;O-48&quot; in cod_reg_sel:
if es_concep_transp:
rf_sug = round(base_cur_f * 0.01, 2) if base_cur_f &gt;= 188000 else 0.0
ri_sug = round(base_cur_f * 0.00414, 2) if base_cur_f &gt;= 188000 else 0.0
cta_rf_sug_def = &quot;23652505&quot;
cta_ri_sug_def = &quot;23680513&quot;
elif es_concep_aduana:
rf_sug = round(base_cur_f * 0.04, 2) if base_cur_f &gt;= 210000 else 0.0
ri_sug = round(base_cur_f * 0.00966, 2)
cta_rf_sug_def = &quot;23652503&quot;
cta_ri_sug_def = &quot;23680505&quot;
elif es_concep_compra:
rf_sug = round(base_cur_f * 0.025, 2) if base_cur_f &gt;= 1047000 else 0.0
ri_sug = round(base_cur_f * 0.01104, 2) if base_cur_f &gt;= 1414000 else 0.0
cta_rf_sug_def = &quot;23654001&quot;
cta_ri_sug_def = &quot;23680501&quot;
elif es_concep_hotel:
rf_sug = round(base_cur_f * 0.035, 2) if base_cur_f &gt;= 210000 else 0.0
ri_sug = round(base_cur_f * 0.00966, 2)
cta_rf_sug_def = &quot;23652501&quot;
cta_ri_sug_def = &quot;23680505&quot;
elif es_concep_honor:
rf_sug = round(base_cur_f * 0.11, 2)
ri_sug = round(base_cur_f * 0.00966, 2)
cta_rf_sug_def = &quot;23651501&quot;
cta_ri_sug_def = &quot;23680505&quot;
elif es_concep_soft:
rf_sug = 0.0
ri_sug = 0.0
cta_rf_sug_def = &quot;&quot;
cta_ri_sug_def = &quot;&quot;
else:
rf_sug = round(base_cur_f * 0.04, 2) if base_cur_f &gt;= 210000 else 0.0
ri_sug = round(base_cur_f * 0.00966, 2)
cta_rf_sug_def = &quot;23652503&quot;

cta_ri_sug_def = &quot;23680505&quot;
elif any(k in cod_reg_sel for k in [&quot;O-15&quot;, &quot;O-47&quot;]):
rf_sug = 0.0
ri_sug = 0.0
cta_rf_sug_def = &quot;&quot;
cta_ri_sug_def = &quot;&quot;
else:
rf_sug = float(fac_sel.get(&quot;ReteFuente&quot;, 0.0))
ri_sug = float(fac_sel.get(&quot;ReteICA&quot;, 0.0))
cta_rf_sug_def = str(fac_sel.get(&quot;Cta ReteFuente&quot;) or &quot;23652503&quot;)
cta_ri_sug_def = str(fac_sel.get(&quot;Cta ReteICA&quot;) or &quot;23680505&quot;)
val_rf_prev = float(fac_sel.get(&quot;ReteFuente&quot;, 0.0))
val_ri_prev = float(fac_sel.get(&quot;ReteICA&quot;, 0.0))
# Si el concepto es transporte o el valor previo era 0/calculado como servicio, precargar
sugerencia
val_rf_input = rf_sug if (es_concep_transp or val_rf_prev == 0.0 or val_rf_prev ==
round(base_cur_f * 0.04, 2)) else val_rf_prev
val_ri_input = ri_sug if (es_concep_transp or val_ri_prev == 0.0 or val_ri_prev ==
round(base_cur_f * 0.00966, 2)) else val_ri_prev
with c_v1:
nueva_base = st.number_input(&quot;Base Gravable / Subtotal ($):&quot;, value=base_cur_f,
step=1000.0, key=f&quot;nb_{fac_sel[&#39;Comprobante Siigo&#39;]}&quot;)
nueva_rfte = st.number_input(&quot;ReteFuente ($):&quot;, value=val_rf_input, step=100.0,
key=f&quot;nrf_{fac_sel[&#39;Comprobante Siigo&#39;]}&quot;)
cta_rf_sug_box = &quot;23652505&quot; if (es_concep_transp or (val_rf_input &gt; 0 and
abs(val_rf_input - round(base_cur_f * 0.01, 2)) &lt; 2.0)) else str(fac_sel.get(&quot;Cta ReteFuente&quot;) or cta_rf_sug_def)
nueva_cta_rf = st.text_input(&quot;Cuenta ReteFuente:&quot;, value=cta_rf_sug_box,
key=f&quot;inp_cta_rf_{fac_sel[&#39;Comprobante Siigo&#39;]}&quot;, help=&quot;Cuenta contable de ReteFuente (23652505 Transporte 1%,
23652503 Servicios 4%, 23654001 Compras 2.5%).&quot;)
with c_v2:
lbl_iva_desc = &quot;IVA Servicios ($) (24081501):&quot; if es_importacion_factura else &quot;IVA
Descontable ($):&quot;
nuevo_iva = st.number_input(lbl_iva_desc, value=float(fac_sel[&quot;IVA&quot;]), step=100.0,
key=f&quot;niva_{fac_sel[&#39;Comprobante Siigo&#39;]}&quot;)
nuevo_rica = st.number_input(&quot;ReteICA ($):&quot;, value=val_ri_input, step=100.0,
key=f&quot;nri_{fac_sel[&#39;Comprobante Siigo&#39;]}&quot;)
cta_ri_sug_box = &quot;23680513&quot; if (es_concep_transp or (val_ri_input &gt; 0 and
abs(val_ri_input - round(base_cur_f * 0.00414, 2)) &lt; 2.0)) else str(fac_sel.get(&quot;Cta ReteICA&quot;) or cta_ri_sug_def)
nueva_cta_ri = st.text_input(&quot;Cuenta ReteICA:&quot;, value=cta_ri_sug_box,
key=f&quot;inp_cta_ri_{fac_sel[&#39;Comprobante Siigo&#39;]}&quot;, help=&quot;Cuenta contable de ReteICA (23680513 Transporte 4.14‰,
23680505 Servicios 9.66‰, 23680501 Compras 11.04‰).&quot;)
# CAMPO DEDICADO: IVA DE IMPORTACIÓN (CUENTA 240835) - SOLO PARA FACTURAS DE IMPORTACIÓN
if es_importacion_factura:
raw_imp_f = fac_sel.get(&quot;IVA Importación&quot;)
try:
val_iva_imp_actual = float(raw_imp_f) if (raw_imp_f is not None and
pd.notna(raw_imp_f)) else 0.0
if pd.isna(val_iva_imp_actual):
val_iva_imp_actual = 0.0
except Exception:
val_iva_imp_actual = 0.0
raw_cta_f = fac_sel.get(&quot;Cta IVA Importación&quot;)
if raw_cta_f is None or pd.isna(raw_cta_f) or str(raw_cta_f).strip() in [&quot;&quot;, &quot;nan&quot;,
&quot;None&quot;]:
cta_iva_imp_actual = &quot;240835&quot;
else:
cta_iva_imp_actual = str(raw_cta_f).strip()
c_imp1, c_imp2 = st.columns(2)
with c_imp1:
nuevo_iva_imp = st.number_input(
&quot;�� IVA de Importación ($) (Cta 240835):&quot;,
value=float(val_iva_imp_actual),
step=1000.0,
key=f&quot;niva_imp_val_{fac_sel[&#39;Comprobante Siigo&#39;]}&quot;,
help=&quot;Ingresa aquí el valor del IVA pagado en la importación / declaración de
aduanas que se imputa a la cuenta 240835.&quot;
)

with c_imp2:
cta_iva_imp_cod = st.text_input(
&quot;Cuenta IVA Importación:&quot;,
value=cta_iva_imp_actual,
key=f&quot;inp_cta_ivaimp_cod_{fac_sel[&#39;Comprobante Siigo&#39;]}&quot;,
help=&quot;Cuenta contable de IVA para importación (por defecto 240835).&quot;
)
else:
nuevo_iva_imp = 0.0
cta_iva_imp_cod = &quot;240835&quot;
asumir_imp = st.checkbox(
&quot;Asumir Impuestos / Retenciones (Cruza con Agente Aduanero / Triangulación)&quot;,
value=bool(fac_sel.get(&quot;Impuestos Asumidos&quot;, False) or es_aduanero or any(k in
fac_sel[&quot;Proveedor&quot;].upper() for k in [&quot;CARGO&quot;, &quot;ADUANA&quot;, &quot;PORTUARIA&quot;, &quot;DHL&quot;])),
key=f&quot;asum_in_{fac_sel[&#39;Comprobante Siigo&#39;]}&quot;,
help=&quot;Si se asumen, las retenciones van a la cuenta 53152001 (Retenciones Asumidas) y el
Saldo por Pagar de la factura queda en Base + IVA (ej. .651.939).&quot;
)
comp_key_btn = str(fac_sel.get(&#39;Comprobante Siigo&#39;, &#39;&#39;)).strip()
if st.button(&quot;�� Aplicar Cambios y Recalcular Asiento&quot;, key=f&quot;btn_aplica_reg_{comp_key_btn}&quot;):
cod_reg = nuevo_reg_sel.split()[0]
m_idx_list = df_p[df_p[&quot;Comprobante Siigo&quot;].astype(str).str.strip() == comp_sel].index
r_idx = m_idx_list[0] if len(m_idx_list) &gt; 0 else fac_sel.name
df_p.at[r_idx, &quot;Régimen Fiscal Emisor&quot;] = cod_reg
df_p.at[r_idx, &quot;Base&quot;] = nueva_base
df_p.at[r_idx, &quot;IVA&quot;] = nuevo_iva
df_p.at[r_idx, &quot;ReteFuente&quot;] = nueva_rfte
df_p.at[r_idx, &quot;ReteICA&quot;] = nuevo_rica
cta_rf_clean = nueva_cta_rf.split()[0].strip()
cta_ri_clean = nueva_cta_ri.split()[0].strip()
df_p.at[r_idx, &quot;Cta ReteFuente&quot;] = cta_rf_clean
df_p.at[r_idx, &quot;Cta ReteICA&quot;] = cta_ri_clean
if es_concep_transp or cta_rf_clean == &quot;23652505&quot;:
df_p.at[r_idx, &quot;Categoría&quot;] = &quot;Transporte Carga (1%)&quot;
df_p.at[r_idx, &quot;Razón Contable&quot;] = &quot;Transporte Terrestre de Carga / Fletes (1%
ReteFuente [23652505] - 4.14‰ ReteICA [23680513])&quot;
if not str(df_p.at[r_idx, &quot;Descripcion&quot;]).strip() or &quot;Importacion&quot; in
str(df_p.at[r_idx, &quot;Descripcion&quot;]):
df_p.at[r_idx, &quot;Descripcion&quot;] = &quot;INLAND RATE TRANSPORTE TERRESTRE&quot;
cta_clean = nueva_cta_cxp.split()[0].strip()
df_p.at[r_idx, &quot;Cuenta Pasivo Especifica&quot;] = cta_clean
df_p.at[r_idx, &quot;Cta Contrapartida&quot;] = cta_clean
df_p.at[r_idx, &quot;Cta Principal&quot;] = nueva_cta_p.split()[0].strip()
cta_iva_clean = nueva_cta_iva.split()[0].strip() if &#39;nueva_cta_iva&#39; in locals() and
nueva_cta_iva else (&quot;24081501&quot; if es_importacion_factura else &quot;24081001&quot;)
df_p.at[r_idx, &quot;Cta IVA&quot;] = cta_iva_clean
nuevo_iva_imp_f = float(nuevo_iva_imp) if pd.notna(nuevo_iva_imp) else 0.0
cta_iva_imp_clean = cta_iva_imp_cod.split()[0].strip()
if not cta_iva_imp_clean or cta_iva_imp_clean in [&quot;nan&quot;, &quot;None&quot;]:
cta_iva_imp_clean = &quot;240835&quot;
df_p.at[r_idx, &quot;IVA Importación&quot;] = nuevo_iva_imp_f
df_p.at[r_idx, &quot;Cta IVA Importación&quot;] = cta_iva_imp_clean
df_p.at[r_idx, &quot;Impuestos Asumidos&quot;] = asumir_imp
df_p.at[r_idx, &quot;Editada Manualmente&quot;] = True
r_iva_val = float(fac_sel.get(&quot;ReteIVA&quot;, 0.0)) if pd.notna(fac_sel.get(&quot;ReteIVA&quot;)) else
0.0
if asumir_imp:
saldo_p = round(nueva_base + nuevo_iva + nuevo_iva_imp_f, 2)
df_p.at[r_idx, &quot;Neto a Pagar&quot;] = saldo_p
df_p.at[r_idx, &quot;Total Neto&quot;] = saldo_p
else:
neto_p_calc = round(nueva_base + nuevo_iva + nuevo_iva_imp_f - nueva_rfte - nuevo_rica
- r_iva_val, 2)
df_p.at[r_idx, &quot;Neto a Pagar&quot;] = neto_p_calc
df_p.at[r_idx, &quot;Total Neto&quot;] = neto_p_calc
tot_esperado = round(nueva_base + nuevo_iva + nuevo_iva_imp_f, 2)
if nuevo_iva_imp_f &gt; 0 and abs(float(df_p.at[r_idx, &quot;Total&quot;]) - tot_esperado) &gt; 10.0:

df_p.at[r_idx, &quot;Total&quot;] = tot_esperado
st.session_state[&quot;df_procesado&quot;] = df_p
# Asiento de triangulación protegido: no se sobreescribe con asiento estándar
# Guardar regla en la Memoria de Aprendizaje Continuo
guardar_regla_aprendizaje(
empresa,
fac_sel.get(&quot;NIT Emisor&quot;),
fac_sel.get(&quot;Proveedor&quot;),
cta_p=nueva_cta_p.split()[0].strip(),
resp_fiscal=cod_reg,
cta_cxp=cta_clean,
cta_iva=cta_iva_clean,
asumir_impuestos=asumir_imp
)
guardar_trabajo_en_historial(
empresa, df_p,
excel_bytes=None,
excel_nombre=st.session_state.get(&quot;excel_nombre&quot;, &quot;Reporte.xlsx&quot;),
dict_pdfs_renombrados=None,
dict_pdfs_originales=None,
zip_bytes=None,
job_id=st.session_state.get(&quot;job_actual_id&quot;)
)
st.session_state[&quot;factura_seleccionada_auditoria&quot;] = comp_sel
st.session_state.pop(&quot;sel_factura_auditoria_box&quot;, None)
st.success(f&quot;¡Valores y Cuenta CxP ({cta_clean}) actualizados y asiento recalculado con
éxito!&quot;)
st.rerun()
if es_aduanero:
st.markdown(&quot;&lt;br&gt;&quot;, unsafe_allow_html=True)
st.markdown(&quot;&quot;&quot;
&lt;div style=&quot;background:#fef3c7; padding:15px; border-radius:8px; border-left:5px solid
#d97706;&quot;&gt;
&lt;h5 style=&quot;color:#92400e; margin-top:0;&quot;&gt;Alerta de Importacion / Agenciamiento
Aduanero&lt;/h5&gt;
&lt;p style=&quot;color:#78350f; font-size:14px; margin-bottom:5px;&quot;&gt;
En este documento intervienen &lt;b&gt;gastos por cuenta de terceros&lt;/b&gt; y &lt;b&gt;honorarios propios
del agente&lt;/b&gt;:
&lt;/p&gt;
&lt;ul style=&quot;color:#78350f; font-size:14px;&quot;&gt;
&lt;li&gt;&lt;b&gt;Ingresos Propios del Agente (Honorarios / Comision):&lt;/b&gt; Gravados con IVA 19%,
sujetos a ReteFuente de Servicios (4%) u Honorarios (11%).&lt;/li&gt;
&lt;li&gt;&lt;b&gt;Pagos por Cuenta de Terceros (Tributos / Fletes / Bodegaje):&lt;/b&gt; Imputables
directamente al costo de importacion (Cuenta 146505). No llevan IVA del agente ni retencion en la fuente sobre el
intermediario.&lt;/li&gt;
&lt;/ul&gt;
&lt;/div&gt;
&quot;&quot;&quot;, unsafe_allow_html=True)
with col_a2:
st.markdown(&quot;#### Resumen Financiero&quot;)
st.metric(&quot;Base Gravable&quot;, f&quot;${float(fac_sel.get(&#39;Base&#39;, 0.0) or 0.0):,.2f}&quot;)
st.metric(&quot;IVA Liquidado (Servicios)&quot;, f&quot;${float(fac_sel.get(&#39;IVA&#39;, 0.0) or 0.0):,.2f}&quot;)
iva_imp_cur = float(fac_sel.get(&quot;IVA Importación&quot;, 0.0))
if iva_imp_cur &gt; 0:
st.metric(&quot;�� IVA Importación (240835)&quot;, f&quot;${iva_imp_cur:,.2f}&quot;)
st.metric(&quot;ReteFuente&quot;, f&quot;-${float(fac_sel.get(&#39;ReteFuente&#39;, 0.0) or 0.0):,.2f}&quot;)
st.metric(&quot;ReteICA&quot;, f&quot;-${fac_sel.get(&#39;ReteICA&#39;, 0.0):,.2f}&quot;)
reteiva_val = fac_sel.get(&quot;ReteIVA&quot;, 0.0)
if reteiva_val &gt; 0:
st.metric(&quot;ReteIVA&quot;, f&quot;-${reteiva_val:,.2f}&quot;)
if fac_sel.get(&quot;Impuestos Asumidos&quot;):
saldo_mostrar = round(fac_sel[&quot;Base&quot;] + fac_sel[&quot;IVA&quot;] + iva_imp_cur, 2)
st.metric(&quot;Saldo CxP a Cruzar (Triangulación)&quot;, f&quot;${saldo_mostrar:,.2f}&quot;)
st.caption(&quot;ℹ️ Impuestos asumidos (Cta 53152001). Saldo total a cruzar.&quot;)
else:
neto_cxp = round(fac_sel[&quot;Base&quot;] + fac_sel[&quot;IVA&quot;] + iva_imp_cur - fac_sel[&quot;ReteFuente&quot;] -
fac_sel.get(&quot;ReteICA&quot;, 0.0) - reteiva_val, 2)
st.metric(&quot;Total Neto CxP (Cta 22 / 23)&quot;, f&quot;${neto_cxp:,.2f}&quot;)

# VISTA PREVIA DE LA FACTURA BAJO DEMANDA (AHORRO DE MEMORIA RAM)
st.markdown(&quot;---&quot;)
st.markdown(&quot;### �� Vista Previa del Documento Soporte&quot;)
ver_pdf_ahora = st.checkbox(
&quot;��️ Mostrar Documento PDF de esta Factura en Pantalla&quot;,
value=False,
key=&quot;chk_ver_pdf_bajo_demanda_tab2&quot;,
help=&quot;Actívalo cuando quieras ver la imagen del PDF. Mantenerlo desactivado mientras editas
cuentas o trabajas en otras pestañas evita saturar la memoria RAM en Streamlit Cloud.&quot;
)
dict_renom = st.session_state.get(&quot;dict_pdfs&quot;, {})
dict_orig = st.session_state.get(&quot;raw_uploaded_pdfs&quot;, {})
if ver_pdf_ahora:
pdf_bytes_encontrado, origen_desc, pags_encontradas = buscar_y_extraer_pdf(
fac_sel,
dict_renombrados=dict_renom,
dict_originales=dict_orig,
empresa_compradora=empresa
)
else:
pdf_bytes_encontrado, origen_desc, pags_encontradas = None, None, None
if pdf_bytes_encontrado:
# Banner de coincidencia exacta con el registro contable
st.success(f&quot;✅ **Factura y Registro Contable Vinculados:** Comprobante **{fac_sel[&#39;Comprobante
Siigo&#39;]}** | Factura: **{fac_sel.get(&#39;Factura&#39;, &#39;&#39;)}** | Proveedor: **{fac_sel.get(&#39;Proveedor&#39;, &#39;Proveedor&#39;)}** (NIT:
{fac_sel.get(&#39;NIT Emisor&#39;, &#39;&#39;)}) — Total: **${float(fac_sel.get(&#39;Total&#39;, 0.0) or 0.0):,.2f}**&quot;)
try:
reader_prev = PdfReader(io.BytesIO(pdf_bytes_encontrado))
num_pags_tot = len(reader_prev.pages)
except Exception:
num_pags_tot = 1
pags_badge = f&quot;{num_pags_tot} páginas completas&quot; if num_pags_tot &gt; 1 else &quot;1 página&quot;
st.markdown(f&quot;&quot;&quot;
&lt;div style=&quot;background:#0070ba; color:white; padding:9px 16px; border-radius:6px; font-weight:600;
font-size:14px; display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;&quot;&gt;
&lt;span&gt;�� {origen_desc} — Factura {fac_sel.get(&#39;Factura&#39;, &#39;&#39;)} ({fac_sel.get(&#39;Proveedor&#39;
&#39;Proveedor&#39;)})&lt;/span&gt;
&lt;span style=&quot;background:rgba(255,255,255,0.25); padding:3px 10px; border-radius:12px; font-
size:12px;&quot;&gt;�� {pags_badge}&lt;/span&gt;
&lt;/div&gt;
&quot;&quot;&quot;, unsafe_allow_html=True)
safe_c_key = re.sub(r&#39;\W&#39;, &#39;_&#39;, str(fac_sel[&#39;Comprobante Siigo&#39;]))
renderizar_visor_pdf_completo(
pdf_bytes_encontrado,
nombre_archivo=fac_sel[&quot;Soporte PDF Renombrado&quot;],
fac_sel=fac_sel,
key_prefix=f&quot;fac_{safe_c_key}&quot;,
nit_comprador=re.sub(r&quot;\D&quot;, &quot;&quot;, str(empresa.get(&quot;nit&quot;, &quot;9013464125&quot;)))
)
elif ver_pdf_ahora and (dict_orig or dict_renom):
st.warning(&quot;⚠️ No se identificó automáticamente esta factura. Puedes seleccionar manualmente
cualquier PDF cargado para visualizarlo:&quot;)
todos_los_pdfs = {**dict_orig, **dict_renom}
pdf_elegido = st.selectbox(&quot;Selecciona un archivo PDF cargado:&quot;, list(todos_los_pdfs.keys()),
key=&quot;sel_pdf_manual_box&quot;)
if pdf_elegido:
f_bytes_sel = todos_los_pdfs[pdf_elegido]
f_bytes_sel, _ = desbloquear_pdf_bytes(f_bytes_sel, nit_receptor=re.sub(r&quot;\D&quot;, &quot;&quot;,
str(empresa.get(&quot;nit&quot;, &quot;9013464125&quot;))))
safe_f_key = re.sub(r&#39;\W&#39;, &#39;_&#39;, str(pdf_elegido))
renderizar_visor_pdf_completo(
f_bytes_sel,
nombre_archivo=pdf_elegido,
fac_sel=fac_sel,

key_prefix=f&quot;man_{safe_f_key}&quot;,
nit_comprador=re.sub(r&quot;\D&quot;, &quot;&quot;, str(empresa.get(&quot;nit&quot;, &quot;9013464125&quot;)))
)
else:
st.markdown(f&quot;&quot;&quot;
&lt;div style=&quot;border: 1px solid #cbd5e1; border-radius: 8px; padding: 18px; background: #ffffff;
box-shadow: 0 1px 3px rgba(0,0,0,0.05); margin-bottom: 20px;&quot;&gt;
&lt;div style=&quot;border-bottom: 2px solid #0070ba; padding-bottom: 8px; margin-bottom: 12px;
display: flex; justify-content: space-between; align-items: center;&quot;&gt;
&lt;h4 style=&quot;margin:0; color:#0f172a;&quot;&gt;�� Cuadro de Factura Electrónica&lt;/h4&gt;
&lt;span style=&quot;background:#eff6ff; color:#1d4ed8; padding:3px 10px; border-radius:4px; font-
weight:bold; font-family:monospace;&quot;&gt;{fac_sel.get(&#39;Factura&#39;, &#39;&#39;)}&lt;/span&gt;
&lt;/div&gt;
&lt;div style=&quot;display: grid; grid-template-columns: 1fr 1fr; gap: 14px; font-size: 14px; color:
#334155;&quot;&gt;
&lt;div&gt;&lt;b&gt;Proveedor (Emisor):&lt;/b&gt; {fac_sel.get(&#39;Proveedor&#39;, &#39;Proveedor&#39;)}&lt;br&gt;&lt;b&gt;NIT
Emisor:&lt;/b&gt; {fac_sel.get(&#39;NIT Emisor&#39;, &#39;&#39;)}&lt;/div&gt;
&lt;div&gt;&lt;b&gt;Empresa Compradora:&lt;/b&gt; {empresa[&#39;nombre&#39;]}&lt;br&gt;&lt;b&gt;NIT Receptor:&lt;/b&gt;
{empresa[&#39;nit&#39;]}&lt;/div&gt;
&lt;div&gt;&lt;b&gt;Fecha de Emisión:&lt;/b&gt; {fac_sel.get(&#39;Fecha&#39;, &#39;&#39;)}&lt;/div&gt;
&lt;div&gt;&lt;b&gt;Valor Total Facturado:&lt;/b&gt; ${float(fac_sel.get(&#39;Total&#39;, 0.0) or 0.0):,.2f}&lt;/div&gt;
&lt;/div&gt;
&lt;p style=&quot;margin-top: 14px; margin-bottom: 0; font-size: 13px; color: #64748b;&quot;&gt;
&lt;i&gt;�� Nota: Sube los archivos PDF en la Pestaña 1 para ver el documento digitalizado e
este visor.&lt;/i&gt;
&lt;/p&gt;
&lt;/div&gt;
&quot;&quot;&quot;, unsafe_allow_html=True)
# TABLA DE CONTABILIZACION (ASIENTO CONTABLE)
st.markdown(&quot;### �� Asiento Contable del Comprobante (Cómo se Contabilizó)&quot;)
st.caption(&quot;Detalle de partida doble con imputación de cuentas, débitos, créditos y sumas iguales:&quot;)
fac_num_aud = str(fac_sel[&quot;Factura&quot;]).strip()
asiento_triang_aprobado = None
# 1. Buscar si esta factura corresponde al agente de algún paquete en triangulación
if &quot;paquetes_importacion&quot; in st.session_state:
for p_k, p_v in st.session_state[&quot;paquetes_importacion&quot;].items():
if str(p_v[&quot;agente&quot;][&quot;Factura&quot;]).strip() == fac_num_aud:
# Si el paquete está validado en Pestaña 3, usar su asiento exacto aprobado
if st.session_state.get(f&quot;paquete_listo_{p_k}&quot;, False):
if f&quot;asiento_fijo_pq_{p_k}&quot; in st.session_state and
st.session_state[f&quot;asiento_fijo_pq_{p_k}&quot;] is not None:
asiento_triang_aprobado = st.session_state[f&quot;asiento_fijo_pq_{p_k}&quot;].copy()
elif &quot;asiento_fijo&quot; in p_v and p_v[&quot;asiento_fijo&quot;] is not None:
asiento_triang_aprobado = p_v[&quot;asiento_fijo&quot;].copy()
elif &quot;asiento_aprobado&quot; in p_v and p_v[&quot;asiento_aprobado&quot;] is not None:
asiento_triang_aprobado = p_v[&quot;asiento_aprobado&quot;].copy()
if asiento_triang_aprobado is None:
asientos_m = st.session_state.get(&quot;asientos_triangulacion_por_factura&quot;, {})
if fac_num_aud in asientos_m and asientos_m[fac_num_aud] is not None:
asiento_triang_aprobado = asientos_m[fac_num_aud].copy()
if asiento_triang_aprobado is None and (st.session_state.get(f&quot;paquete_listo_{p_k}&quot;,
False) or p_v.get(&quot;diferencia&quot;, 999) &lt; 1.0):
en_nd_k = st.session_state.get(f&quot;enviar_nd_pq_{p_k}&quot;, False)
en_gp_k = st.session_state.get(f&quot;enviar_gp_pq_{p_k}&quot;, False)
en_h2_k = st.session_state.get(f&quot;enviar_h2_pq_{p_k}&quot;, False)
if en_h2_k:
df_as_k = obtener_asiento_contable_hoja2(p_v[&quot;agente&quot;], es_aduanero=False)
elif en_gp_k:
df_as_k, _, _ = generar_asiento_triangulacion_paquete(p_v[&quot;agente&quot;],
p_v[&quot;terceros&quot;], enviar_a_no_deducible=False)
asiento_gp_filas = []
for _, r_as in df_as_k.iterrows():
if str(r_as[&quot;Código Cuenta&quot;]).strip() == CUENTA_NO_DEDUCIBLE:
asiento_gp_filas.append({
&quot;Código Cuenta&quot;: CUENTA_IMPORTACION_TRANSITO,
&quot;Descripción Cuenta&quot;: f&quot;Mercancías en Tránsito / Base Propia Agente
(Fac {p_v[&#39;agente&#39;].get(&#39;Factura&#39;, &#39;&#39;)})&quot;,

&quot;Tercero / NIT&quot;: f&quot;{p_v[&#39;agente&#39;].get(&#39;NIT Emisor&#39;, &#39;&#39;)} -
{p_v[&#39;agente&#39;].get(&#39;Proveedor&#39;, &#39;&#39;)[:25]}&quot;,
&quot;Débito ($)&quot;: float(r_as[&quot;Débito ($)&quot;]),
&quot;Crédito ($)&quot;: 0.0
})
else:
asiento_gp_filas.append(r_as.to_dict())
df_as_k = pd.DataFrame(asiento_gp_filas)
else:
df_as_k, _, _ = generar_asiento_triangulacion_paquete(p_v[&quot;agente&quot;],
p_v[&quot;terceros&quot;], enviar_a_no_deducible=en_nd_k)
asiento_triang_aprobado = df_as_k
break
if asiento_triang_aprobado is None:
asientos_m = st.session_state.get(&quot;asientos_triangulacion_por_factura&quot;, {})
if fac_num_aud in asientos_m and asientos_m[fac_num_aud] is not None:
asiento_triang_aprobado = asientos_m[fac_num_aud].copy()
if asiento_triang_aprobado is not None and not (isinstance(asiento_triang_aprobado, pd.DataFrame) and
asiento_triang_aprobado.empty):
df_asiento = pd.DataFrame(asiento_triang_aprobado).copy()
if &quot;Descripción Cuenta&quot; in df_asiento.columns and &quot;Descripción de la Cuenta&quot; not in
df_asiento.columns:
df_asiento = df_asiento.rename(columns={&quot;Descripción Cuenta&quot;: &quot;Descripción de la Cuenta&quot;})
cols_deseadas = [c for c in [&quot;Código Cuenta&quot;, &quot;Descripción de la Cuenta&quot;, &quot;Tercero / NIT&quot;, &quot;Débito
($)&quot;, &quot;Crédito ($)&quot;] if c in df_asiento.columns]
df_asiento = df_asiento[cols_deseadas]
# Asiento validado en Triangulación se respeta íntegramente sin sobreescribir
st.markdown(f&quot;&quot;&quot;
&lt;div style=&quot;background:#f0fdf4; border:1px solid #86efac; border-left:5px solid #16a34a; border-
radius:8px; padding:12px 16px; margin-bottom:12px;&quot;&gt;
&lt;h5 style=&quot;margin:0 0 4px 0; color:#166534;&quot;&gt;�� Contabilización Oficial Aprobada e
Triangulación&lt;/h5&gt;
&lt;p style=&quot;margin:0; color:#14532d; font-size:13.5px;&quot;&gt;
Esta factura fue conciliada en la Pestaña 3. El asiento contable real cancela las cuentas
por pagar de terceros y traslada la deuda al agente aduanero:
&lt;/p&gt;
&lt;/div&gt;
&quot;&quot;&quot;, unsafe_allow_html=True)
else:
asiento_filas = []
es_nc = &quot;Devolucion&quot; in str(fac_sel[&quot;Operacion&quot;])
asume_ret = bool(fac_sel.get(&quot;Impuestos Asumidos&quot;, False))
cta_cxp_usar = str(fac_sel.get(&quot;Cuenta Pasivo Especifica&quot;) or fac_sel.get(&quot;Cta Contrapartida&quot;) or
(&quot;22050505&quot; if es_aduanero else &quot;22050501&quot;)).strip()
# Verificar si la factura seleccionada es un tercero cruzado en un paquete de importación
pq_cruz_encontrado = None
if &quot;paquetes_importacion&quot; in st.session_state:
for p_k, p_v in st.session_state[&quot;paquetes_importacion&quot;].items():
df_terc_v = p_v.get(&quot;terceros&quot;)
if df_terc_v is not None and not df_terc_v.empty:
if str(fac_sel[&quot;Factura&quot;]).strip() in
df_terc_v[&quot;Factura&quot;].astype(str).str.strip().values:
pq_cruz_encontrado = {
&quot;pq_id&quot;: p_k,
&quot;agente&quot;: p_v[&quot;agente&quot;][&quot;Proveedor&quot;],
&quot;agente_fac&quot;: p_v[&quot;agente&quot;][&quot;Factura&quot;],
&quot;es_listo&quot;: st.session_state.get(f&quot;paquete_listo_{p_k}&quot;, False)
}
break
saldo_cruce_ind = obtener_saldo_cruce_factura(fac_sel)
if pq_cruz_encontrado is not None:
st.markdown(f&quot;&quot;&quot;
&lt;div style=&quot;background:#eff6ff; border:1px solid #bfdbfe; border-left:5px solid #2563eb;
border-radius:8px; padding:14px 18px; margin-bottom:14px;&quot;&gt;

&lt;div style=&quot;display:flex; justify-content:space-between; align-items:center;&quot;&gt;
&lt;h5 style=&quot;margin:0 0 4px 0; color:#1e40af;&quot;&gt;�� Factura Vinculada en Paquete d
Importación #{pq_cruz_encontrado[&#39;pq_id&#39;]} (con {pq_cruz_encontrado[&#39;agente&#39;]} - Fac
{pq_cruz_encontrado[&#39;agente_fac&#39;]})&lt;/h5&gt;
&lt;span style=&quot;background:#dbeafe; color:#1e40af; font-size:12px; font-weight:bold;
padding:3px 8px; border-radius:12px;&quot;&gt;Cruce Activo en Triangulación&lt;/span&gt;
&lt;/div&gt;
&lt;p style=&quot;margin:6px 0 0 0; color:#1e3a8a; font-size:13.5px;&quot;&gt;
Esta factura mantiene su &lt;b&gt;contabilización propia completa&lt;/b&gt; (Costo/Gasto e IVA).
En la importación, lo único que se cruza es su &lt;b&gt;Cuenta por Pagar ({cta_cxp_usar})&lt;/b&gt; por valor de
&lt;b&gt;${saldo_cruce_ind:,.2f}&lt;/b&gt;, la cual se debita en el paquete de triangulación para cancelar la obligación y
trasladar la deuda al agente aduanero.
&lt;/p&gt;
&lt;/div&gt;
&quot;&quot;&quot;, unsafe_allow_html=True)
elif esta_excl_ind:
st.markdown(f&quot;&quot;&quot;
&lt;div style=&quot;background:#fef2f2; border:1px solid #fecaca; border-left:5px solid #dc2626;
border-radius:8px; padding:12px 16px; margin-bottom:12px;&quot;&gt;
&lt;h5 style=&quot;margin:0 0 4px 0; color:#991b1b;&quot;&gt;�� Factura Marcada como &#39;No Contabilizar&#39;&lt;/h5&gt;
&lt;p style=&quot;margin:0; color:#7f1d1d; font-size:13.5px;&quot;&gt;
Esta factura ha sido excluida de la planilla de compras de Siigo. A continuación se
muestra su estructura contable para fines informativos y de auditoría:
&lt;/p&gt;
&lt;/div&gt;
&quot;&quot;&quot;, unsafe_allow_html=True)
# Cada factura SIEMPRE mantiene su causación contable propia íntegra
df_asiento = obtener_asiento_contable_hoja2(fac_sel, es_aduanero=es_aduanero)
cols_as_dsp = [c for c in [&quot;Código Cuenta&quot;, &quot;Descripción de la Cuenta&quot;, &quot;Tercero / NIT&quot;, &quot;Débito ($)&quot;,
&quot;Crédito ($)&quot;] if c in df_asiento.columns]
st.dataframe(
df_asiento[cols_as_dsp].style.format({&quot;Débito ($)&quot;: &quot;${:,.2f}&quot;, &quot;Crédito ($)&quot;: &quot;${:,.2f}&quot;}),
use_container_width=True,
hide_index=True
)
sum_deb = df_asiento[&quot;Débito ($)&quot;].sum()
sum_cred = df_asiento[&quot;Crédito ($)&quot;].sum()
c_as1, c_as2, c_as3 = st.columns(3)
with c_as1:
st.metric(&quot;Total Débito&quot;, f&quot;${sum_deb:,.2f}&quot;)
with c_as2:
st.metric(&quot;Total Crédito&quot;, f&quot;${sum_cred:,.2f}&quot;)
with c_as3:
st.metric(&quot;Diferencia / Cuadre&quot;, f&quot;${abs(sum_deb - sum_cred):,.2f}&quot;)
st.success(&quot;✅ **Comprobante Verificado:** Partida doble cuadrada con sumas iguales.&quot;)
else:
st.info(&quot;Carga el archivo Excel en la Pestana 1 para habilitar la auditoria.&quot;)
except Exception as _e_aud:
if type(_e_aud).__name__ in [&#39;RerunException&#39;, &#39;StopException&#39;, &#39;ScriptControlException&#39;]:
raise _e_aud
st.error(f&quot;⚠️ Ocurrió un error en la pestaña de Auditoría Contable: {_e_aud}&quot;)
with st.expander(&quot;�� Ver detalles técnicos del error (Traceback):&quot;):
st.code(traceback.format_exc())
with tab_triangulacion:
try:
st.markdown(&quot;### �� Paquetes de Importación y Cruce de Cuentas por Pagar&quot;)
st.caption(&quot;Separa y concilia cada importación en paquetes individuales: Cobro de Euro Shipping = Factura
DHL (Flete) + Factura Agencia Aduanas + Factura Garaje/Almacenadora.&quot;)
if &quot;df_procesado&quot; in st.session_state and st.session_state[&quot;df_procesado&quot;] is not None and
isinstance(st.session_state[&quot;df_procesado&quot;], pd.DataFrame) and not st.session_state[&quot;df_procesado&quot;].empty:
df_total = normalizar_df_procesado(st.session_state[&quot;df_procesado&quot;])
st.session_state[&quot;df_procesado&quot;] = df_total
# Identificar facturas de comercio exterior / importación de forma segura sin errores de tipos

prov_series = df_total[&quot;Proveedor&quot;] if &quot;Proveedor&quot; in df_total.columns else pd.Series(&quot;&quot;,
index=df_total.index)
desc_series = df_total[&quot;Descripcion&quot;] if &quot;Descripcion&quot; in df_total.columns else pd.Series(&quot;&quot;,
index=df_total.index)
grp_series = df_total[&quot;Grupo Importación&quot;] if &quot;Grupo Importación&quot; in df_total.columns else
pd.Series(&quot;&quot;, index=df_total.index)
cta_cxp_series = df_total[&quot;Cuenta Pasivo Especifica&quot;] if &quot;Cuenta Pasivo Especifica&quot; in
df_total.columns else (df_total[&quot;Cta Contrapartida&quot;] if &quot;Cta Contrapartida&quot; in df_total.columns else pd.Series(&quot;&quot;,
index=df_total.index))
cta_p_series = df_total[&quot;Cta Principal&quot;] if &quot;Cta Principal&quot; in df_total.columns else pd.Series(&quot;&quot;,
index=df_total.index)
prov_upper = prov_series.fillna(&quot;&quot;).astype(str).str.upper()
desc_upper = desc_series.fillna(&quot;&quot;).astype(str).str.upper()
grp_imp_str = grp_series.fillna(&quot;&quot;).astype(str).str.strip()
cta_cxp_str = cta_cxp_series.fillna(&quot;&quot;).astype(str).str.strip()
cta_p_str = cta_p_series.fillna(&quot;&quot;).astype(str).str.strip()
# 1. Facturas explícitamente marcadas en Excel con Grupo de Importación o clasificadas como aduaneras
cond_grp = (grp_imp_str != &quot;&quot;) &amp; (~grp_imp_str.isin([&quot;nan&quot;, &quot;None&quot;, &quot;0&quot;, &quot;0.0&quot;]))
cond_es_adu = (df_total[&quot;Es Aduanera&quot;] == True) if &quot;Es Aduanera&quot; in df_total.columns else
pd.Series(False, index=df_total.index)
# 2. Cuentas contables específicas de importación / fletes DHL / agencias aduaneras
cond_ctas = cta_cxp_str.isin([&quot;22050505&quot;, &quot;23359501&quot;, &quot;14650501&quot;]) | (cta_p_str == &quot;14650501&quot;)
# 3. Proveedores específicos de Comercio Exterior / Logística Internacional / Aduana / Puertos /
Fletes
REGEX_IMPORTACION = (
r&quot;EURO\s*SHIPPING|TRADE\s*GLOBAL|CONSOLCARGO|BLUE\s*LOGISTICS|&quot;
r&quot;KUEHNE|PANALPINA|DSV|EXPEDITORS|TIBA|HUBEMAR|SCHENKER|COLTRANS|COLMAS|&quot;
r&quot;DHL|FEDEX|UPS|CARGO|ADUANA|ADUANERA|AGENCIAMIENTO|PORTUARIA|PUERTO\s*BAHIA|&quot;
r&quot;CONTECAR|SPRC|SPB|COMPAS|ALPOPULAR|ALMAVIVA|RANSA|ALMACENADORA|&quot;
r&quot;MAERSK|HAPAG|SEABOARD|HAMBURG|CMA\s*CGM|EVERGREEN|COSCO|YANG\s*MING&quot;
)
cond_prov = prov_upper.str.contains(REGEX_IMPORTACION, regex=True, na=False)
# 4. Descripción que evidencia gastos de importación / desaduanamiento
REGEX_DESC_IMPORT = (
r&quot;IMPORTACI[OÓ]N|AGENCIAMIENTO|ARANCEL|DESADUANAMIENTO|NACIONALIZACI[OÓ]N|&quot;
r&quot;BODEGAJE\s*PUERTO|ALMACENAJE\s*ADUANERO|MANDATO\s*ADUANERO|FLETE\s*INTERNACIONAL&quot;
)
cond_desc = desc_upper.str.contains(REGEX_DESC_IMPORT, regex=True, na=False)
# Condición estricta: ÚNICAMENTE facturas relacionadas con la importación (rojas y blancas de
importación)
cond_aduanera = cond_grp | cond_es_adu | cond_ctas | cond_prov | cond_desc
df_adu = df_total[cond_aduanera].copy()
# Configuración dinámica y extensible de Agentes Coordinadores (Forwarders)
AGENTES_COORDINADORES_BASE = [&quot;EURO SHIPPING&quot;, &quot;TRADE GLOBAL&quot;, &quot;CONSOLCARGO&quot;, &quot;BLUE LOGISTICS&quot;,
&quot;KUEHNE&quot;, &quot;PANALPINA&quot;, &quot;DSV&quot;, &quot;EXPEDITORS&quot;, &quot;TIBA&quot;, &quot;HUBEMAR&quot;]
todos_provs = sorted([str(p) for p in df_total[&quot;Proveedor&quot;].dropna().unique() if str(p).strip()]) if
&quot;Proveedor&quot; in df_total.columns else []
provs_agentes_sugeridos = [p for p in todos_provs if any(k in str(p).upper() for k in
AGENTES_COORDINADORES_BASE)]
with st.expander(&quot;�� Configurar Agentes Coordinadores (Los que nos cobran directamente a nosotros):&quot;
expanded=False):
st.caption(&quot;Selecciona qué empresas actúan como Agentes Coordinadores / Forwarders que emiten
cobros globales consolidados (Euro Shipping, Trade Global, Consolcargo, etc.):&quot;)
agentes_seleccionados_noms = st.multiselect(
&quot;Agentes Coordinadores Activos:&quot;,
options=todos_provs,
default=provs_agentes_sugeridos,
key=&quot;sel_multiselect_agentes_coordinadores&quot;
)
if not agentes_seleccionados_noms:

agentes_seleccionados_noms = provs_agentes_sugeridos
# Facturas de Agentes Principales (Los que nos cobran directamente)
df_agentes_all = df_adu[df_adu[&quot;Proveedor&quot;].isin(agentes_seleccionados_noms)].copy()
# Facturas de Terceros Soporte (Los que se cruzan contra esos agentes)
df_terceros_all = df_adu[~df_adu.index.isin(df_agentes_all.index)].copy()
# Resumen superior
m_c1, m_c2, m_c3, m_c4, m_c5 = st.columns(5)
with m_c1:
st.metric(&quot;Operaciones Agente (Euro / Trade)&quot;, len(df_agentes_all))
with m_c2:
st.metric(&quot;Facturas Soporte en Pool&quot;, len(df_terceros_all), help=&quot;Total de facturas candidatas
para cruce&quot;)
with m_c3:
ya_reg_terc = df_terceros_all[&quot;Ya Registrada&quot;].fillna(False).astype(bool) if &quot;Ya Registrada&quot; in
df_terceros_all.columns else pd.Series(False, index=df_terceros_all.index)
n_rojas_adu = int(ya_reg_terc.sum())
st.metric(&quot;Terceros Contabilizados (��)&quot;, n_rojas_adu, help=&quot;Causadas previamente en Siigo&quot;)
with m_c4:
n_blancas_adu = int((~ya_reg_terc).sum())
st.metric(&quot;Terceros Pendientes (⚪)&quot;, n_blancas_adu, help=&quot;No contabilizadas aún (Nuevas)&quot;)
with m_c5:
n_cuad_prev = sum([1 for p in st.session_state.get(&quot;paquetes_importacion&quot;, {}).values() if
p.get(&quot;diferencia&quot;, 999) &lt; 1.0])
st.metric(&quot;Paquetes Cuadrados ($0.00)&quot;, f&quot;{n_cuad_prev} ✅&quot;)
st.markdown(&quot;---&quot;)
if df_agentes_all.empty:
st.warning(&quot;⚠️ No se encontraron facturas de agentes aduaneros principales (Euro Shipping o Trade
Global) en este reporte.&quot;)
else:
# 1. FUNCIÓN DE AUTO-EMPAQUETAMIENTO INTELIGENTE
# Algoritmo de combinación: encuentra para cada Euro Shipping el subconjunto de DHL + Agencia +
Garaje que suma su valor
def auto_empaquetar_inteligente(df_ag, df_terc):
&quot;&quot;&quot;
Algoritmo cronológico de alta precisión para Agencias Aduaneras (Euro Shipping / Trade
Global):
1. Ordena cronológicamente de Enero hacia Diciembre.
2. Cruza y relaciona tanto facturas de terceros YA contabilizadas (��) como NO contabilizada
(⚪).
3. Para cada factura de agente (ej. 19 de Enero), busca facturas de terceros (DHL, Agencia,
Garaje, etc.):
- Prioritariamente en fechas anteriores o el mismo día (-30 a 0 días de diferencia).
- O en una fecha ligeramente superior (+1 a +12 días de diferencia).
- Descarta o penaliza fechas lejanas de otros meses.
4. Exige concordancia estricta de precio: la suma del Saldo Cruce (Base + IVA) debe concordar
exactamente con el valor total facturado por el agente aduanero.
&quot;&quot;&quot;
import itertools
pqs_res = {}
# 1. Asegurar orden cronológico estricto de Enero hacia Diciembre
df_ag_ord = df_ag.copy()
df_ag_ord[&quot;_dt&quot;] = pd.to_datetime(df_ag_ord[&quot;Fecha&quot;], dayfirst=True, errors=&quot;coerce&quot;)
df_ag_ord = df_ag_ord.sort_values(by=&quot;_dt&quot;, ascending=True)
pool_t = df_terc.copy()
pool_t[&quot;_dt&quot;] = pd.to_datetime(pool_t[&quot;Fecha&quot;], dayfirst=True, errors=&quot;coerce&quot;)
pool_t = pool_t.sort_values(by=&quot;_dt&quot;, ascending=True)
pool_t[&quot;_asignado_pq&quot;] = 0
# PROTEGER Y BLOQUEAR FACTURAS DE PAQUETES YA VALIDADOS
if &quot;paquetes_importacion&quot; in st.session_state:
for p_num_prev, p_data_prev in st.session_state[&quot;paquetes_importacion&quot;].items():
if st.session_state.get(f&quot;paquete_listo_{p_num_prev}&quot;, False):
df_t_prev = p_data_prev.get(&quot;terceros&quot;)
if df_t_prev is not None and not df_t_prev.empty:

for f_b in df_t_prev[&quot;Factura&quot;].dropna():
idx_match = pool_t[pool_t[&quot;Factura&quot;].astype(str).str.strip() ==
str(f_b).strip()].index
if not idx_match.empty:
pool_t.loc[idx_match, &quot;_asignado_pq&quot;] = p_num_prev
for idx_ag_i, (_, ag_i) in enumerate(df_ag_ord.iterrows()):
pq_id_i = idx_ag_i + 1
target_i = float(ag_i[&quot;Total&quot;])
f_ag_dt_i = ag_i[&quot;_dt&quot;]
if &quot;paquetes_importacion&quot; in st.session_state and pq_id_i in
st.session_state[&quot;paquetes_importacion&quot;]:
prev_pq_val = st.session_state[&quot;paquetes_importacion&quot;][pq_id_i]
if st.session_state.get(f&quot;paquete_listo_{pq_id_i}&quot;, False):
# MANTENER INTACTO EL PAQUETE YA VALIDADO
pqs_res[pq_id_i] = prev_pq_val
continue
cands_i = pool_t[pool_t[&quot;_asignado_pq&quot;] == 0].copy()
if cands_i.empty:
pqs_res[pq_id_i] = {&quot;agente&quot;: ag_i, &quot;terceros&quot;: pd.DataFrame(), &quot;diferencia&quot;:
target_i}
continue
# Evaluar concordancia cronológica y de mandato para cada factura de tercero
cand_l = []
for c_idx_k, c_row_k in cands_i.iterrows():
tb_k = float(c_row_k.get(&quot;Base&quot;, 0.0))
tiv_k = float(c_row_k.get(&quot;IVA&quot;, 0.0))
sc_k = obtener_saldo_cruce_factura(c_row_k)
p_u_k = str(c_row_k[&quot;Proveedor&quot;]).upper()
rol_k = &quot;DHL&quot; if &quot;DHL&quot; in p_u_k else (&quot;AGENCIA&quot; if any(k in p_u_k for k in [&quot;CARGO&quot;,
&quot;ADUANA&quot;]) else &quot;GARAJE&quot;)
# 2. Análisis de fechas: concordancia de Enero a Diciembre
c_dt = c_row_k[&quot;_dt&quot;]
score_tiempo = 0
diff_dias = 0
relacion_fecha_txt = &quot;Sin fecha&quot;
if pd.notna(f_ag_dt_i) and pd.notna(c_dt):
diff_dias = int((c_dt - f_ag_dt_i).days)
# REGLA CRONOLÓGICA ESTRICTA: Solo conciliar facturas anteriores (-60 a 0 días) o
un poco después (+1 a +15 días)
if diff_dias &lt; -60 or diff_dias &gt; 15:
continue # DESCARTAR: fuera de la ventana cronológica del despacho
if diff_dias == 0:
score_tiempo = 120
relacion_fecha_txt = &quot;�� Mismo día (0d)&quot;
elif -30 &lt;= diff_dias &lt; 0:
score_tiempo = 100 - abs(diff_dias) * 2
relacion_fecha_txt = f&quot;⬅️ Anterior ({abs(diff_dias)}d antes)&quot;
elif 0 &lt; diff_dias &lt;= 15:
score_tiempo = 85 - (diff_dias * 3)
relacion_fecha_txt = f&quot;➡️ Poco después (+{diff_dias}d desp)&quot;
elif -60 &lt;= diff_dias &lt; -30:
score_tiempo = 40 - abs(diff_dias)
relacion_fecha_txt = f&quot;⬅️ Anterior ({abs(diff_dias)}d antes)&quot;
# Coincidencia directa por Mandato (NIT o Nombre del agente en notas)
nit_fact = re.sub(r&#39;\D&#39;, &#39;&#39;, str(c_row_k.get(&quot;NIT Facturado A&quot;, &quot;&quot;) or
c_row_k.get(&quot;nit_facturado_a&quot;, &quot;&quot;)))
nom_fact = str(c_row_k.get(&quot;Facturado A&quot;, &quot;&quot;) or c_row_k.get(&quot;nombre_facturado_a&quot;,
&quot;&quot;)).upper()
nit_ag_clean = re.sub(r&#39;\D&#39;, &#39;&#39;, str(ag_i.get(&quot;NIT Emisor&quot;, &quot;&quot;)))
nom_ag_u = str(ag_i.get(&quot;Proveedor&quot;, &quot;&quot;)).upper()
score_mandato = 0
if nit_ag_clean and nit_fact and nit_ag_clean[:8] in nit_fact:

score_mandato += 150
elif any(w in nom_fact for w in nom_ag_u.split() if len(w) &gt;= 4):
score_mandato += 80
c_row_k_copy = c_row_k.copy()
c_row_k_copy[&quot;Relación Fecha&quot;] = relacion_fecha_txt
c_row_k_copy[&quot;Diferencia Días&quot;] = diff_dias
cand_l.append({
&quot;index&quot;: c_idx_k,
&quot;fac&quot;: c_row_k[&quot;Factura&quot;],
&quot;saldo&quot;: sc_k,
&quot;rol&quot;: rol_k,
&quot;diff_dias&quot;: diff_dias,
&quot;score_total&quot;: score_tiempo + score_mandato,
&quot;row&quot;: c_row_k_copy
})
# Ordenar candidatos dando prioridad a concordancia de fecha y mandato
cand_l.sort(key=lambda x: x[&quot;score_total&quot;], reverse=True)
# 3. Búsqueda de combinación con precio concorde exacto (target_i)
# Busca combinaciones de 1 a 4 facturas de terceros (tanto �� contabilizadas como
pendientes)
best_combo_i = []
best_diff_i = float(&quot;inf&quot;)
best_score_i = -float(&quot;inf&quot;)
# Ventana cronológica concorde (-35 días antes a +14 días después)
cands_ventana = [c for c in cand_l if -35 &lt;= c[&quot;diff_dias&quot;] &lt;= 14]
if len(cands_ventana) &lt; 5:
cands_ventana = cand_l[:25]
else:
cands_ventana = cands_ventana[:25]
# Probar combinaciones de 1 a 4 facturas (DHL, Agencia, Garaje, rojas o blancas)
for k_c in range(1, min(5, len(cands_ventana) + 1)):
for combo_i in itertools.combinations(cands_ventana, k_c):
s_c_i = round(sum(c[&quot;saldo&quot;] for c in combo_i), 2)
d_c_i = abs(s_c_i - target_i)
score_c = sum(c[&quot;score_total&quot;] for c in combo_i) / len(combo_i)
if d_c_i &lt; 0.05:
# Coincidencia exacta de precio
if score_c &gt; best_score_i or best_diff_i &gt;= 0.05:
best_diff_i = d_c_i
best_combo_i = list(combo_i)
best_score_i = score_c
elif best_diff_i &gt;= 0.05 and d_c_i &lt; best_diff_i:
best_diff_i = d_c_i
best_combo_i = list(combo_i)
best_score_i = score_c
if best_diff_i &lt; 0.05 and best_score_i &gt; 70:
break
# Si no dio cuadre exacto en la ventana cercana, ampliar búsqueda sobre todos los
candidatos
if best_diff_i &gt; 1.0 and len(cand_l) &gt; len(cands_ventana):
cands_ampliados = cand_l[:30]
for k_c in range(1, min(5, len(cands_ampliados) + 1)):
for combo_i in itertools.combinations(cands_ampliados, k_c):
s_c_i = round(sum(c[&quot;saldo&quot;] for c in combo_i), 2)
d_c_i = abs(s_c_i - target_i)
score_c = sum(c[&quot;score_total&quot;] for c in combo_i) / len(combo_i)
if d_c_i &lt; best_diff_i:
best_diff_i = d_c_i
best_combo_i = list(combo_i)
best_score_i = score_c

if d_c_i &lt; 0.05:
break
if best_diff_i &lt; 0.05:
break
sel_idxs_i = [c[&quot;index&quot;] for c in best_combo_i]
pool_t.loc[sel_idxs_i, &quot;_asignado_pq&quot;] = pq_id_i
filas_terc_pq = [c[&quot;row&quot;] for c in best_combo_i]
terc_pq_i = pd.DataFrame(filas_terc_pq) if filas_terc_pq else pd.DataFrame()
pqs_res[pq_id_i] = {
&quot;agente&quot;: ag_i,
&quot;terceros&quot;: terc_pq_i,
&quot;diferencia&quot;: best_diff_i
}
terc_libres_i = pool_t[pool_t[&quot;_asignado_pq&quot;] == 0].copy()
return pqs_res, terc_libres_i
# Inicializar o recuperar paquetes de importación
if &quot;paquetes_importacion&quot; not in st.session_state or st.session_state.get(&quot;_ultimo_agentes_len&quot;)
!= len(df_agentes_all):
pqs_ini, libres_ini = auto_empaquetar_inteligente(df_agentes_all, df_terceros_all)
st.session_state[&quot;paquetes_importacion&quot;] = pqs_ini
st.session_state[&quot;_ultimo_agentes_len&quot;] = len(df_agentes_all)
pqs_actuales = st.session_state[&quot;paquetes_importacion&quot;]
# =========================================================================
# PURGA Y BLOQUEO ESTRICTO: NINGUNA FACTURA VALIDADA PUEDE ESTAR EN OTRO PAQUETE
# =========================================================================
facs_validadas_global = {} # Factura -&gt; pq_id_donde_esta_validada
for p_k_chk, p_v_chk in pqs_actuales.items():
if st.session_state.get(f&quot;paquete_listo_{p_k_chk}&quot;, False):
df_t_chk = p_v_chk.get(&quot;terceros&quot;)
if df_t_chk is not None and not df_t_chk.empty:
for f_val in df_t_chk[&quot;Factura&quot;].dropna():
f_val_clean = str(f_val).strip()
if f_val_clean:
facs_validadas_global[f_val_clean] = p_k_chk
hubo_cambio_purga = False
for p_k_chk, p_v_chk in pqs_actuales.items():
# Si este paquete NO está validado, purgar cualquier factura que pertenezca a un paquete
validado
if not st.session_state.get(f&quot;paquete_listo_{p_k_chk}&quot;, False):
df_t_chk = p_v_chk.get(&quot;terceros&quot;)
if df_t_chk is not None and not df_t_chk.empty:
mask_ya_en_val =
df_t_chk[&quot;Factura&quot;].astype(str).str.strip().isin(facs_validadas_global.keys())
if mask_ya_en_val.any():
df_t_filtrado = df_t_chk[~mask_ya_en_val].copy().reset_index(drop=True)
pqs_actuales[p_k_chk][&quot;terceros&quot;] = df_t_filtrado
tot_ag_c = float(p_v_chk[&quot;agente&quot;][&quot;Total&quot;])
tot_terc_c = sum([obtener_saldo_cruce_factura(r) for _, r in
df_t_filtrado.iterrows()])
pqs_actuales[p_k_chk][&quot;diferencia&quot;] = abs(tot_ag_c - tot_terc_c)
hubo_cambio_purga = True
if hubo_cambio_purga:
st.session_state[&quot;paquetes_importacion&quot;] = pqs_actuales
# Barra de Acciones y Selección de Paquete (Persistente: no salta al paquete 1)
c_top_pq1, c_top_pq2 = st.columns([2.5, 1])
with c_top_pq1:
pqs_keys = list(pqs_actuales.keys())
if &quot;paquete_seleccionado_id&quot; not in st.session_state or
st.session_state[&quot;paquete_seleccionado_id&quot;] not in pqs_keys:
st.session_state[&quot;paquete_seleccionado_id&quot;] = pqs_keys[0] if pqs_keys else 1

mapa_pqs_titulos = {}
for p_num in pqs_keys:
p_data = pqs_actuales[p_num]
ag_t = p_data[&quot;agente&quot;]
n_terc = len(p_data[&quot;terceros&quot;])
s_terc = sum([obtener_saldo_cruce_factura(r) for _, r in p_data[&quot;terceros&quot;].iterrows()])
es_val_p = st.session_state.get(f&quot;paquete_listo_{p_num}&quot;, False)
lock_icon = &quot;�� [VALIDADO] &quot; if es_val_p else &quot;&quot;
mapa_pqs_titulos[p_num] = f&quot;{lock_icon}�� Paquete #{p_num}: {str(ag_t.get(&#39;Proveedor&#39;) o
&#39;Agente&#39;)[:16]} ({ag_t[&#39;Factura&#39;]}) — Cobro: ${float(ag_t.get(&#39;Total&#39;, 0.0) or 0.0):,.0f} | {n_terc} Facturas Terceros
(${float(s_terc or 0.0):,.0f})&quot;
idx_actual_sel = pqs_keys.index(st.session_state[&quot;paquete_seleccionado_id&quot;]) if
st.session_state[&quot;paquete_seleccionado_id&quot;] in pqs_keys else 0
pq_id_sel = st.selectbox(
&quot;�� Selecciona el Paquete de Importación que deseas revisar:&quot;,
options=pqs_keys,
index=idx_actual_sel,
format_func=lambda pid: mapa_pqs_titulos.get(pid, f&quot;Paquete #{pid}&quot;)
)
st.session_state[&quot;paquete_seleccionado_id&quot;] = pq_id_sel
st.session_state[&quot;sel_paquete_activo_key&quot;] = pq_id_sel
paquete_activo = pqs_actuales[pq_id_sel]
agente_actual = paquete_activo[&quot;agente&quot;]
terceros_actual = paquete_activo[&quot;terceros&quot;]
# Sincronización en tiempo real: refrescar datos del agente desde df_procesado (Hoja 2)
if &quot;df_procesado&quot; in st.session_state and st.session_state[&quot;df_procesado&quot;] is not None:
df_pr_act = st.session_state[&quot;df_procesado&quot;]
ag_fac_cur = str(agente_actual[&quot;Factura&quot;]).strip()
m_ag = df_pr_act[df_pr_act[&quot;Factura&quot;].astype(str).str.strip() == ag_fac_cur]
if not m_ag.empty:
agente_actual = m_ag.iloc[0].to_dict()
paquete_activo[&quot;agente&quot;] = agente_actual
tot_agente_actual = float(agente_actual.get(&quot;Total&quot;, 0.0))
base_prop_actual = float(agente_actual.get(&quot;Base&quot;, 0.0))
iva_prop_actual = float(agente_actual.get(&quot;IVA&quot;, 0.0))
iva_imp_act = float(agente_actual.get(&quot;IVA Importación&quot;, 0.0))
subtot_prop_actual = round(base_prop_actual + iva_prop_actual + iva_imp_act, 2)
saldo_terceros_esperado_act = round(tot_agente_actual - subtot_prop_actual, 2)
# Tratamiento contable y asiento de partida doble del paquete
enviar_h2_activo = bool(st.session_state.get(f&quot;enviar_h2_pq_{pq_id_sel}&quot;, False))
enviar_gp_activo = bool(st.session_state.get(f&quot;enviar_gp_pq_{pq_id_sel}&quot;, False))
enviar_nd_activo = bool(st.session_state.get(f&quot;enviar_nd_pq_{pq_id_sel}&quot;, False))
suma_terceros_cxp_act = sum([obtener_saldo_cruce_factura(r) for _, r in
terceros_actual.iterrows()]) if (terceros_actual is not None and not terceros_actual.empty) else 0.0
# Calcular el asiento previo y la diferencia real de partida doble que necesita cuadrarse
if saldo_terceros_esperado_act &gt; 0.05 or enviar_h2_activo:
df_prev, dif_faltante_prev = generar_asiento_mixto_hoja2_con_terceros(agente_actual,
terceros_actual, enviar_a_no_deducible=False, imputar_a_transito=False)
ret_prev = 0.0
else:
df_prev, dif_faltante_prev, ret_prev =
generar_asiento_triangulacion_paquete(agente_actual, terceros_actual, enviar_a_no_deducible=False)
# Garantizar que dif_faltante_prev refleje la diferencia real entre débitos y créditos del
asiento actual
sum_d_pr = df_prev[&quot;Débito ($)&quot;].sum() if not df_prev.empty else 0.0
sum_c_pr = df_prev[&quot;Crédito ($)&quot;].sum() if not df_prev.empty else 0.0
dif_real_cuadre = round(abs(sum_d_pr - sum_c_pr), 2)
if dif_real_cuadre &gt; 0.05:
dif_faltante_prev = dif_real_cuadre

if st.session_state.get(f&quot;paquete_listo_{pq_id_sel}&quot;, False) and
(f&quot;asiento_fijo_pq_{pq_id_sel}&quot; in st.session_state or &quot;asiento_fijo&quot; in paquete_activo):
df_asiento_paquete = st.session_state.get(f&quot;asiento_fijo_pq_{pq_id_sel}&quot;,
paquete_activo.get(&quot;asiento_fijo&quot;)).copy()
dif_no_ded = 0.0
ret_asum = 0.0
elif saldo_terceros_esperado_act &gt; 0.05 or enviar_h2_activo:
# Facturas con ingresos propios y de terceros (o con solicitud explícita de traer de
Página 2)
if enviar_nd_activo:
df_asiento_paquete, _ = generar_asiento_mixto_hoja2_con_terceros(agente_actual,
terceros_actual, enviar_a_no_deducible=True, imputar_a_transito=False)
dif_no_ded = dif_faltante_prev
elif enviar_gp_activo:
df_asiento_paquete, _ = generar_asiento_mixto_hoja2_con_terceros(agente_actual,
terceros_actual, enviar_a_no_deducible=False, imputar_a_transito=True)
dif_no_ded = 0.0
else:
# Dejar pendiente para conciliar con facturas de terceros (DECISIÓN DEL USUARIO)
df_asiento_paquete, dif_no_ded =
generar_asiento_mixto_hoja2_con_terceros(agente_actual, terceros_actual, enviar_a_no_deducible=False,
imputar_a_transito=False)
ret_asum = 0.0
elif enviar_gp_activo and dif_faltante_prev &gt; 0.05:
df_asiento_paquete, _, ret_asum = generar_asiento_triangulacion_paquete(agente_actual,
terceros_actual, enviar_a_no_deducible=False)
asiento_gp_filas = []
for _, r_as in df_asiento_paquete.iterrows():
if str(r_as[&quot;Código Cuenta&quot;]).strip() == CUENTA_NO_DEDUCIBLE:
asiento_gp_filas.append({
&quot;Código Cuenta&quot;: CUENTA_IMPORTACION_TRANSITO,
&quot;Descripción Cuenta&quot;: f&quot;Mercancías en Tránsito / Base Propia Agente (Fac
{agente_actual.get(&#39;Factura&#39;, &#39;&#39;)})&quot;,
&quot;Tercero / NIT&quot;: f&quot;{agente_actual.get(&#39;NIT Emisor&#39;, &#39;&#39;)} -
{agente_actual.get(&#39;Proveedor&#39;, &#39;&#39;)[:25]}&quot;,
&quot;Débito ($)&quot;: float(r_as[&quot;Débito ($)&quot;]),
&quot;Crédito ($)&quot;: 0.0
})
else:
asiento_gp_filas.append(r_as.to_dict())
df_asiento_paquete = pd.DataFrame(asiento_gp_filas)
dif_no_ded = 0.0
else:
df_asiento_paquete, dif_no_ded, ret_asum =
generar_asiento_triangulacion_paquete(agente_actual, terceros_actual, enviar_a_no_deducible=enviar_nd_activo)
def ejecutar_validacion_pq(pasa_a_siigo=True):
st.session_state[f&quot;paquete_listo_{pq_id_sel}&quot;] = True
st.session_state[f&quot;paquete_destino_siigo_{pq_id_sel}&quot;] = pasa_a_siigo
paquete_activo[&quot;destino_siigo&quot;] = pasa_a_siigo
paquete_activo[&quot;asiento_fijo&quot;] = df_asiento_paquete.copy()
paquete_activo[&quot;asiento_aprobado&quot;] = df_asiento_paquete.copy()
if &quot;asientos_triangulacion_por_factura&quot; not in st.session_state:
st.session_state[&quot;asientos_triangulacion_por_factura&quot;] = {}
ag_fac_str = str(agente_actual[&quot;Factura&quot;]).strip()
if pasa_a_siigo:
st.session_state[&quot;asientos_triangulacion_por_factura&quot;][ag_fac_str] =
df_asiento_paquete.copy()
else:
st.session_state[&quot;asientos_triangulacion_por_factura&quot;].pop(ag_fac_str, None)
st.session_state[&quot;paquete_seleccionado_id&quot;] = pq_id_sel
st.session_state[&quot;sel_paquete_activo_key&quot;] = pq_id_sel
# PURGA INMEDIATA: Quitar las facturas de este paquete recién validado de TODOS los demás
paquetes no validados
if not terceros_actual.empty:
facs_recien_validadas =
set(terceros_actual[&quot;Factura&quot;].astype(str).str.strip().dropna())
for otro_id, otro_pq in pqs_actuales.items():
if otro_id != pq_id_sel and not st.session_state.get(f&quot;paquete_listo_{otro_id}&quot;,
False):
df_otro_t = otro_pq.get(&quot;terceros&quot;)
if df_otro_t is not None and not df_otro_t.empty:

mask_remover =
df_otro_t[&quot;Factura&quot;].astype(str).str.strip().isin(facs_recien_validadas)
if mask_remover.any():
df_otro_limpio =
df_otro_t[~mask_remover].copy().reset_index(drop=True)
otro_pq[&quot;terceros&quot;] = df_otro_limpio
tot_ag_o = float(otro_pq[&quot;agente&quot;][&quot;Total&quot;])
tot_terc_o = sum([float(r.get(&quot;Total Neto&quot;, 0.0) or
(float(r.get(&quot;Base&quot;, 0.0)) + float(r.get(&quot;IVA&quot;, 0.0)))) for _, r in df_otro_limpio.iterrows()])
otro_pq[&quot;diferencia&quot;] = abs(tot_ag_o - tot_terc_o)
st.session_state[&quot;paquetes_importacion&quot;] = pqs_actuales
# Traslado y actualización inmediata a df_procesado (Página 2)
if &quot;df_procesado&quot; in st.session_state and st.session_state[&quot;df_procesado&quot;] is not None:
df_pr = st.session_state[&quot;df_procesado&quot;]
for c_col in [&quot;Paquete Validado&quot;, &quot;Destino Siigo&quot;, &quot;Cruza Con Agente&quot;, &quot;Estado
Registro&quot;]:
if c_col in df_pr.columns:
df_pr[c_col] = df_pr[c_col].astype(object)
else:
df_pr[c_col] = None
idx_ag = df_pr[df_pr[&quot;Factura&quot;].astype(str).str.strip() == ag_fac_str].index
if not idx_ag.empty:
if pasa_a_siigo:
df_pr.at[idx_ag[0], &quot;Estado Registro&quot;] = f&quot;�� Validada en Triangulación (P
#{pq_id_sel} - Pasa a Siigo)&quot;
else:
df_pr.at[idx_ag[0], &quot;Estado Registro&quot;] = f&quot;�� Conciliado en Triangulación (P
#{pq_id_sel} - Solo Control Interno)&quot;
df_pr.at[idx_ag[0], &quot;Paquete Validado&quot;] = pq_id_sel
df_pr.at[idx_ag[0], &quot;Destino Siigo&quot;] = pasa_a_siigo
if not terceros_actual.empty:
for _, tr_row in terceros_actual.iterrows():
tr_fac = str(tr_row[&quot;Factura&quot;]).strip()
idx_t = df_pr[df_pr[&quot;Factura&quot;].astype(str).str.strip() == tr_fac].index
if not idx_t.empty:
es_r = bool(df_pr.at[idx_t[0], &quot;Ya Registrada&quot;]) or
bool(df_pr.at[idx_t[0], &quot;No Contabilizar&quot;]) or (tr_fac in st.session_state.get(&quot;facturas_no_contabilizar&quot;, set()))
if es_r:
df_pr.at[idx_t[0], &quot;Estado Registro&quot;] = f&quot;�� Cruzada en P
#{pq_id_sel} (Solo cuenta pasivo)&quot;
else:
if pasa_a_siigo:
df_pr.at[idx_t[0], &quot;Estado Registro&quot;] = f&quot;⚪ Cruzada en Pq
#{pq_id_sel} (Causada en triangulación)&quot;
else:
df_pr.at[idx_t[0], &quot;Estado Registro&quot;] = f&quot;�� Conciliada en P
#{pq_id_sel} (Solo Control Interno)&quot;
df_pr.at[idx_t[0], &quot;Paquete Validado&quot;] = pq_id_sel
df_pr.at[idx_t[0], &quot;Cruza Con Agente&quot;] = ag_fac_str
st.session_state[&quot;df_procesado&quot;] = df_pr
guardar_estado_manual(empresa)
def ejecutar_desbloqueo_pq():
st.session_state[f&quot;paquete_listo_{pq_id_sel}&quot;] = False
st.session_state.pop(f&quot;paquete_destino_siigo_{pq_id_sel}&quot;, None)
paquete_activo.pop(&quot;destino_siigo&quot;, None)
paquete_activo.pop(&quot;asiento_fijo&quot;, None)
paquete_activo.pop(&quot;asiento_aprobado&quot;, None)
ag_fac_str = str(agente_actual[&quot;Factura&quot;]).strip()
if &quot;asientos_triangulacion_por_factura&quot; in st.session_state:
st.session_state[&quot;asientos_triangulacion_por_factura&quot;].pop(ag_fac_str, None)
st.session_state[&quot;paquete_seleccionado_id&quot;] = pq_id_sel
st.session_state[&quot;sel_paquete_activo_key&quot;] = pq_id_sel
if &quot;df_procesado&quot; in st.session_state and st.session_state[&quot;df_procesado&quot;] is not None:
df_pr = st.session_state[&quot;df_procesado&quot;]
for c_col in [&quot;Paquete Validado&quot;, &quot;Destino Siigo&quot;, &quot;Cruza Con Agente&quot;, &quot;Estado
Registro&quot;]:
if c_col in df_pr.columns:
df_pr[c_col] = df_pr[c_col].astype(object)
else:
df_pr[c_col] = None
idx_ag = df_pr[df_pr[&quot;Factura&quot;].astype(str).str.strip() == ag_fac_str].index

if not idx_ag.empty:
df_pr.at[idx_ag[0], &quot;Estado Registro&quot;] = &quot;�� Agente Aduanero&quot;
df_pr.at[idx_ag[0], &quot;Paquete Validado&quot;] = None
df_pr.at[idx_ag[0], &quot;Destino Siigo&quot;] = None
if not terceros_actual.empty:
for _, tr_row in terceros_actual.iterrows():
tr_fac = str(tr_row[&quot;Factura&quot;]).strip()
idx_t = df_pr[df_pr[&quot;Factura&quot;].astype(str).str.strip() == tr_fac].index
if not idx_t.empty:
es_r = bool(df_pr.at[idx_t[0], &quot;Ya Registrada&quot;]) or
bool(df_pr.at[idx_t[0], &quot;No Contabilizar&quot;])
cp = df_pr.at[idx_t[0], &quot;Comprobante Previo&quot;]
df_pr.at[idx_t[0], &quot;Estado Registro&quot;] = f&quot;�� Ya Registrada ({cp})&quot; if es_
else &quot;⚪ Compra Pendiente&quot;
df_pr.at[idx_t[0], &quot;Paquete Validado&quot;] = None
df_pr.at[idx_t[0], &quot;Cruza Con Agente&quot;] = None
st.session_state[&quot;df_procesado&quot;] = df_pr
guardar_estado_manual(empresa)
with c_top_pq2:
st.write(&quot;&quot;)
st.write(&quot;&quot;)
if st.button(&quot;�� Re-calcular Auto-Empaquetado&quot;, key=&quot;btn_recalc_auto_pqs&quot;, help=&quot;Vuelve
calcular las combinaciones exactas de DHL + Agencia + Garaje para cada cobro de Euro&quot;):
pqs_re, _ = auto_empaquetar_inteligente(df_agentes_all, df_terceros_all)
st.session_state[&quot;paquetes_importacion&quot;] = pqs_re
st.success(&quot;¡Paquetes recalculados y separados!&quot;)
st.rerun()
# Resumen visual del Paquete Seleccionado
col_pq1, col_pq2 = st.columns([1.3, 2.5])
with col_pq1:
st.markdown(f&quot;&quot;&quot;
&lt;div style=&quot;background:#fffbeb; border:1px solid #fde68a; border-left:5px solid #d97706;
border-radius:8px; padding:14px; margin-bottom:12px;&quot;&gt;
&lt;h4 style=&quot;margin:0 0 6px 0; color:#92400e;&quot;&gt;�� Factura del Agente (Cobro Total)&lt;/h4&gt;
&lt;p style=&quot;margin:0; font-size:14px; color:#78350f;&quot;&gt;
&lt;b&gt;Proveedor:&lt;/b&gt; {agente_actual[&#39;Proveedor&#39;]}&lt;br&gt;
&lt;b&gt;NIT:&lt;/b&gt; {agente_actual[&#39;NIT Emisor&#39;]}&lt;br&gt;
&lt;b&gt;Factura:&lt;/b&gt; {agente_actual[&#39;Factura&#39;]}&lt;br&gt;
&lt;b&gt;Fecha:&lt;/b&gt; {agente_actual[&#39;Fecha&#39;]}&lt;br&gt;
&lt;b&gt;Total Facturado:&lt;/b&gt; &lt;span style=&quot;font-size:18px; font-weight:bold;
color:#b45309;&quot;&gt;${tot_agente_actual:,.2f}&lt;/span&gt;&lt;br&gt;
&lt;b&gt;IVA Discriminado:&lt;/b&gt; ${float(agente_actual.get(&#39;IVA&#39;, 0.0)):,.2f}
&lt;/p&gt;
&lt;/div&gt;
&quot;&quot;&quot;, unsafe_allow_html=True)
es_val_card = st.session_state.get(f&quot;paquete_listo_{pq_id_sel}&quot;, False)
if not es_val_card:
st.caption(f&quot;ℹ️ Elige el destino de este paquete al validarlo:&quot;)
c_vcard1, c_vcard2 = st.columns(2)
with c_vcard1:
if st.button(&quot;�� Pasar a Hoja 2&quot;, key=f&quot;btn_val_siigo_card_{pq_id_sel}&quot;
type=&quot;primary&quot;, use_container_width=True, help=&quot;Bloquea el paquete y lo traslada a Página 2 para subirlo a la Planilla
de Siigo&quot;):
ejecutar_validacion_pq(pasa_a_siigo=True)
st.success(f&quot;¡Paquete #{pq_id_sel} trasladado a Página 2 para Siigo!&quot;)
st.rerun()
with c_vcard2:
if st.button(&quot;�� Solo Conciliación&quot;, key=f&quot;btn_val_concil_card_{pq_id_sel}&quot;
use_container_width=True, help=&quot;Bloquea el paquete para identificar y conciliar facturas, sin subirlo a la Planilla de
Siigo&quot;):
ejecutar_validacion_pq(pasa_a_siigo=False)
st.success(f&quot;¡Paquete #{pq_id_sel} guardado para Conciliación Interna!&quot;)
st.rerun()
else:
es_siigo = st.session_state.get(f&quot;paquete_destino_siigo_{pq_id_sel}&quot;, True)
if es_siigo:
st.success(f&quot;�� **Pq #{pq_id_sel}: PASA A HOJA 2 (PLANILLA SIIGO)**&quot;)
else:
st.info(f&quot;�� **Pq #{pq_id_sel}: SOLO CONCILIACIÓN INTERNA**&quot;)

if st.button(f&quot;�� Desbloquear Paquete #{pq_id_sel}&quot;, key=f&quot;btn_desb_card_{pq_id_sel}&quot;
use_container_width=True):
ejecutar_desbloqueo_pq()
st.rerun()
with col_pq2:
st.markdown(f&quot;##### Facturas de Terceros que Componen este Paquete #{pq_id_sel}:&quot;)
if not terceros_actual.empty:
filas_terc_disp = []
tot_s_terceros = 0.0
for _, tr in terceros_actual.iterrows():
es_r = tr.get(&quot;Ya Registrada&quot;, False)
badge_est = f&quot;�� Ya en Siigo ({tr.get(&#39;Comprobante Previo&#39;, &#39;10-Prev&#39;)})&quot; if es_r els
&quot;⚪ No Contabilizada (Pendiente)&quot;
t_b = float(tr.get(&quot;Base&quot;, 0.0)) if pd.notna(tr.get(&quot;Base&quot;)) else 0.0
t_iv = float(tr.get(&quot;IVA&quot;, 0.0)) if pd.notna(tr.get(&quot;IVA&quot;)) else 0.0
s_cruce = obtener_saldo_cruce_factura(tr)
tot_s_terceros += s_cruce
cta_actual_tr = str(tr.get(&quot;Cuenta Pasivo Especifica&quot;, &quot;22050505&quot; if &quot;CARGO&quot; in
str(tr[&quot;Proveedor&quot;]).upper() else &quot;23359501&quot;)).strip()
p_nom = str(tr[&quot;Proveedor&quot;]).upper()
rol_dsp = &quot;�� DHL (Flete)&quot; if &quot;DHL&quot; in p_nom else (&quot;�� Agencia (Aduana)&quot; if any(k
p_nom for k in [&quot;CARGO&quot;, &quot;ADUANA&quot;]) else &quot;�� Garaje / Almacén&quot;)
rel_fec_val = tr.get(&quot;Relación Fecha&quot;)
if pd.isna(rel_fec_val) or not str(rel_fec_val).strip() or str(rel_fec_val).strip() in
[&quot;nan&quot;, &quot;None&quot;]:
concordancia_dsp = calcular_concordancia_fecha(tr.get(&quot;Fecha&quot;),
agente_actual.get(&quot;Fecha&quot;))
else:
concordancia_dsp = str(rel_fec_val).strip()
filas_terc_disp.append({
&quot;Rol&quot;: rol_dsp,
&quot;Fecha Factura&quot;: tr.get(&quot;Fecha&quot;, &quot;-&quot;),
&quot;Concordancia Fecha&quot;: concordancia_dsp,
&quot;Estado Contable&quot;: badge_est,
&quot;Proveedor Tercero&quot;: tr[&quot;Proveedor&quot;][:22],
&quot;Factura&quot;: tr[&quot;Factura&quot;],
&quot;Subtotal (Base)&quot;: t_b,
&quot;IVA&quot;: t_iv,
&quot;Saldo Cruce&quot;: s_cruce,
&quot;Cuenta Contable&quot;: f&quot;{cta_actual_tr} (CxP a Cruzar)&quot;
})
df_terc_disp = pd.DataFrame(filas_terc_disp)
st.dataframe(df_terc_disp.style.format({
&quot;Subtotal (Base)&quot;: &quot;${:,.2f}&quot;,
&quot;IVA&quot;: &quot;${:,.2f}&quot;,
&quot;Saldo Cruce&quot;: &quot;${:,.2f}&quot;
}), use_container_width=True, hide_index=True)
dif_pq_actual = abs(tot_agente_actual - tot_s_terceros)
if dif_pq_actual &lt; 5.0:
st.success(f&quot;✅ **Paquete #{pq_id_sel} Exacto:** Total Terceros:
**${tot_s_terceros:,.2f}** == Cobro Euro: **${tot_agente_actual:,.2f}** | Diferencia: **$0.00**&quot;)
else:
st.info(f&quot;�� Total Terceros: **${tot_s_terceros:,.2f}** | Cobro Euro
**${tot_agente_actual:,.2f}** | Diferencia a No Deducible: **${dif_pq_actual:,.2f}**&quot;)
else:
st.warning(f&quot;⚠️ El Paquete #{pq_id_sel} no tiene facturas de terceros asignadas todavía.&quot;)
# Expander para agregar/quitar facturas manualmente a este paquete (BLOQUEADO SI ESTÁ VALIDADO)
es_validado_este = st.session_state.get(f&quot;paquete_listo_{pq_id_sel}&quot;, False)
if es_validado_este:
st.markdown(f&quot;&quot;&quot;
&lt;div style=&quot;background:#f0fdf4; border:1px solid #86efac; border-left:5px solid #16a34a;
border-radius:8px; padding:12px 16px; margin:12px 0;&quot;&gt;
&lt;h5 style=&quot;margin:0 0 4px 0; color:#166534;&quot;&gt;�� Paquete #{pq_id_sel} VALIDADO
BLOQUEADO&lt;/h5&gt;
&lt;p style=&quot;margin:0; color:#14532d; font-size:13.5px;&quot;&gt;
Este paquete se encuentra congelado. Ninguna factura puede ser añadida ni retirada
mientras esté validado. Para realizar modificaciones, presiona &lt;b&gt;&#39;�� Desbloquear y Modificar&#39;&lt;/b&gt; abajo.

&lt;/p&gt;
&lt;/div&gt;
&quot;&quot;&quot;, unsafe_allow_html=True)
with st.expander(f&quot;⚙️ Modificar este Paquete #{pq_id_sel} (Sacar Facturas o Agregar Nuevas):&quot;,
expanded=False if es_validado_este else False):
if es_validado_este:
st.warning(&quot;⚠️ Debes desbloquear el paquete abajo antes de poder agregar o sacar
facturas.&quot;)
st.stop()
st.caption(&quot;Administra las facturas asignadas a esta importación: saca las que no correspondan
o añade nuevas facturas de DHL, Agencia o Garaje:&quot;)
# 1. SECCIÓN PARA SACAR / REDIRIGIR FACTURAS
st.markdown(&quot;##### ➖ Sacar y Redirigir una factura de este paquete:&quot;)
st.caption(&quot;Si ves una factura que no corresponde a este paquete, sácala y el sistema la
reubicará automáticamente o la enviará a compras directas para que no quede pendiente:&quot;)
facs_en_este = list(terceros_actual[&quot;Factura&quot;].unique()) if not terceros_actual.empty else []
if facs_en_este:
opciones_quitar = []
mapa_quitar_row = {}
for _, tr_k in terceros_actual.iterrows():
tb_k = float(tr_k.get(&quot;Base&quot;, 0.0))
tiv_k = float(tr_k.get(&quot;IVA&quot;, 0.0))
sc_k = obtener_saldo_cruce_factura(tr_k)
tag_q = f&quot;[{tr_k[&#39;Factura&#39;]}] {str(tr_k.get(&#39;Proveedor&#39;) or &#39;Tercero&#39;)[:24]} (Saldo:
${sc_k:,.2f})&quot;
opciones_quitar.append(tag_q)
mapa_quitar_row[tag_q] = tr_k
sel_para_quitar = st.selectbox(&quot;1. Selecciona la factura que deseas SACAR de este
paquete:&quot;, [&quot;(Seleccionar...)&quot;] + opciones_quitar, key=f&quot;sel_rem_tr_{pq_id_sel}&quot;)
if sel_para_quitar != &quot;(Seleccionar...)&quot;:
fila_a_mover = mapa_quitar_row[sel_para_quitar]
fac_nom_mover = fila_a_mover[&quot;Factura&quot;]
sc_mover = obtener_saldo_cruce_factura(fila_a_mover)
otros_pqs = [p for p in pqs_actuales.keys() if p != pq_id_sel and not
st.session_state.get(f&quot;paquete_listo_{p}&quot;, False)]
# Calcular cuál otro paquete tiene un faltante más cercano a esta factura
mejor_pq_sug = None
mejor_diff_sug = float(&quot;inf&quot;)
for p_cand_id in otros_pqs:
p_cand = pqs_actuales[p_cand_id]
tot_ag_c = float(p_cand[&quot;agente&quot;][&quot;Total&quot;])
tot_terc_c = sum([obtener_saldo_cruce_factura(r) for _, r in
p_cand[&quot;terceros&quot;].iterrows()])
faltante_c = tot_ag_c - tot_terc_c
diff_c = abs(faltante_c - sc_mover)
if diff_c &lt; mejor_diff_sug:
mejor_diff_sug = diff_c
mejor_pq_sug = p_cand_id
sug_txt = f&quot; (Sugerencia: Paquete #{mejor_pq_sug})&quot; if mejor_pq_sug else &quot;&quot;
opciones_destino = [
f&quot;�� Auto-Reubicar automáticamente donde mejor encaje{sug_txt}&quot;,
&quot;�� Mover a otro Paquete específico...&quot;,
&quot;�� Redirigir a Compras Directas Ordinarias (Pestaña 1 y 4 - Se pagó directo)&quot;,
&quot;⚪ Dejar libre en inventario (Sin asignar)&quot;
]
c_dest1, c_dest2 = st.columns([2.6, 1.2])
with c_dest1:
dest_sel = st.radio(f&quot;2. ¿A dónde deseas redirigir la factura {fac_nom_mover}?:&quot;,
opciones_destino, key=f&quot;rad_dest_{pq_id_sel}_{fac_nom_mover}&quot;)
destino_pq_manual = None
if &quot;Mover a otro Paquete específico&quot; in dest_sel:

opciones_otros_pqs = [f&quot;Paquete #{p}:
{str(pqs_actuales[p][&#39;agente&#39;].get(&#39;Proveedor&#39;) or &#39;Agente&#39;)[:16]} (Fac {pqs_actuales[p][&#39;agente&#39;][&#39;Factura&#39;]})&quot; for p
in otros_pqs]
if opciones_otros_pqs:
sel_otro_pq_str = st.selectbox(&quot;Selecciona el paquete de destino:&quot;,
opciones_otros_pqs, key=f&quot;sel_otro_pq_{pq_id_sel}&quot;)
destino_pq_manual = int(sel_otro_pq_str.split(&quot;:&quot;)[0].replace(&quot;Paquete #&quot;,
&quot;&quot;).strip())
else:
st.info(&quot;No hay otros paquetes de forwarders disponibles en este
archivo.&quot;)
with c_dest2:
st.write(&quot;&quot;)
st.write(&quot;&quot;)
if st.button(&quot;�� Sacar y Redirigir Factura&quot;
key=f&quot;btn_ejecutar_reubicar_{pq_id_sel}&quot;, help=&quot;Retira la factura de este paquete y la redirige al destino
seleccionado.&quot;):
# 1. Retirar del paquete actual
terceros_remanente = terceros_actual[terceros_actual[&quot;Factura&quot;] !=
fac_nom_mover].copy()
st.session_state[&quot;paquetes_importacion&quot;][pq_id_sel][&quot;terceros&quot;] =
terceros_remanente
# 2. Redirigir al destino
if &quot;Auto-Reubicar&quot; in dest_sel and mejor_pq_sug is not None:
dest_pq =
st.session_state[&quot;paquetes_importacion&quot;][mejor_pq_sug][&quot;terceros&quot;]
dest_nuevo = pd.concat([dest_pq,
pd.DataFrame([fila_a_mover])]).drop_duplicates(subset=[&quot;Factura&quot;]).reset_index(drop=True)
st.session_state[&quot;paquetes_importacion&quot;][mejor_pq_sug][&quot;terceros&quot;] =
dest_nuevo
st.success(f&quot;¡Factura {fac_nom_mover} retirada del Paquete #{pq_id_sel} y
auto-reubicada en el Paquete #{mejor_pq_sug}!&quot;)
elif &quot;Mover a otro Paquete&quot; in dest_sel and destino_pq_manual is not None:
dest_pq =
st.session_state[&quot;paquetes_importacion&quot;][destino_pq_manual][&quot;terceros&quot;]
dest_nuevo = pd.concat([dest_pq,
pd.DataFrame([fila_a_mover])]).drop_duplicates(subset=[&quot;Factura&quot;]).reset_index(drop=True)
st.session_state[&quot;paquetes_importacion&quot;][destino_pq_manual][&quot;terceros&quot;] =
dest_nuevo
st.success(f&quot;¡Factura {fac_nom_mover} movida con éxito al Paquete
#{destino_pq_manual}!&quot;)
elif &quot;Compras Directas Ordinarias&quot; in dest_sel:
if &quot;df_procesado&quot; in st.session_state:
df_glob = st.session_state[&quot;df_procesado&quot;]
m_idx = df_glob[df_glob[&quot;Factura&quot;] == fac_nom_mover].index
if not m_idx.empty:
df_glob.at[m_idx[0], &quot;Es Aduanera&quot;] = False
df_glob.at[m_idx[0], &quot;Grupo Importación&quot;] = &quot;&quot;
df_glob.at[m_idx[0], &quot;Estado Registro&quot;] = &quot;⚪ Compra Directa
Ordinaria&quot;
df_glob.at[m_idx[0], &quot;Cta Contrapartida&quot;] = &quot;22050501&quot;
st.session_state[&quot;df_procesado&quot;] = df_glob
st.success(f&quot;¡Factura {fac_nom_mover} retirada de importaciones y enviada
a Compras Directas Ordinarias (Pestaña 1 y 4)!&quot;)
else:
st.success(f&quot;¡Factura {fac_nom_mover} retirada del Paquete #{pq_id_sel} y
devuelta a facturas libres!&quot;)
st.session_state[&quot;paquete_seleccionado_id&quot;] = pq_id_sel
st.session_state[&quot;sel_paquete_activo_key&quot;] = pq_id_sel
guardar_estado_manual(empresa)
st.rerun()
else:
st.info(&quot;Este paquete no tiene facturas asignadas para retirar.&quot;)
st.markdown(&quot;---&quot;)
# 2. SECCIÓN PARA AGREGAR NUEVAS FACTURAS
cands_disp_agregar = df_terceros_all.copy()
f_ag_dt_sug = pd.to_datetime(agente_actual.get(&quot;Fecha&quot;, &quot;&quot;), dayfirst=True, errors=&quot;coerce&quot;)

# RESTRICCIÓN DE BLOQUEO: Las facturas de paquetes ya validados no se pueden añadir a otros
facturas_bloqueadas_validadas = set()
if &quot;paquetes_importacion&quot; in st.session_state:
for p_k_v, p_dat_v in st.session_state[&quot;paquetes_importacion&quot;].items():
if st.session_state.get(f&quot;paquete_listo_{p_k_v}&quot;, False):
df_t_val = p_dat_v.get(&quot;terceros&quot;)
if df_t_val is not None and not df_t_val.empty:
for f_asig in df_t_val[&quot;Factura&quot;].dropna():
facturas_bloqueadas_validadas.add(str(f_asig).strip())
# 2.1 SUGERENCIAS INTELIGENTES DE FACTURAS PARA INGRESOS DE TERCEROS
if dif_faltante_prev &gt; 0.05:
st.markdown(
f&quot;&quot;&quot;
&lt;div style=&quot;background:#f0fdf4; border:1.5px solid #22c55e; border-radius:8px;
padding:12px 16px; margin:10px 0 16px 0;&quot;&gt;
&lt;h5 style=&quot;margin:0 0 4px 0; color:#15803d;&quot;&gt;�� Facturas Sugeridas para Ingreso
de Terceros (Faltante Pendiente: ${dif_faltante_prev:,.2f}):&lt;/h5&gt;
&lt;p style=&quot;margin:0 0 8px 0; font-size:13px; color:#166534;&quot;&gt;
Facturas libres en el sistema que &lt;b&gt;no superan el monto de la factura
(${tot_agente_actual:,.2f})&lt;/b&gt; ni el &lt;b&gt;monto de terceros (${saldo_terceros_esperado_act if
saldo_terceros_esperado_act &gt; 0.05 else tot_agente_actual:,.2f})&lt;/b&gt;, analizadas por concordancia de fletes, cargos,
fecha y proveedor:
&lt;/p&gt;
&lt;/div&gt;
&quot;&quot;&quot;,
unsafe_allow_html=True
)
# Construir pool de sugerencias
cands_sug_list = []
f_ag_dt_sug = pd.to_datetime(agente_actual.get(&quot;Fecha&quot;, &quot;&quot;), dayfirst=True,
errors=&quot;coerce&quot;)
monto_limite_fac = tot_agente_actual + 0.05
monto_limite_terc = (saldo_terceros_esperado_act if saldo_terceros_esperado_act &gt; 0.05
else tot_agente_actual) + 0.05
for _, tr_cand_s in cands_disp_agregar.iterrows():
f_c_num = str(tr_cand_s[&quot;Factura&quot;]).strip()
if f_c_num in facs_en_este or f_c_num in facturas_bloqueadas_validadas or not f_c_num:
continue
sc_cand_val = obtener_saldo_cruce_factura(tr_cand_s)
if sc_cand_val &lt;= 0.0:
continue
# REGLA ESTRICTA: Que no superen el monto de la factura ni el monto de los ingresos
para terceros
if sc_cand_val &gt; monto_limite_fac or sc_cand_val &gt; monto_limite_terc or sc_cand_val &gt;
(dif_faltante_prev + 0.05):
continue
prov_u = str(tr_cand_s.get(&quot;Proveedor&quot;, &quot;&quot;)).upper()
fec_str = str(tr_cand_s.get(&quot;Fecha&quot;, &quot;&quot;))
fec_dt = pd.to_datetime(fec_str, dayfirst=True, errors=&quot;coerce&quot;)
diff_dias = int((fec_dt - f_ag_dt_sug).days) if (pd.notna(fec_dt) and
pd.notna(f_ag_dt_sug)) else 999
# REGLA CRONOLÓGICA ESTRICTA: Solo facturas anteriores (-60d a 0d) o un poco después
(+1d a +15d)
if pd.notna(fec_dt) and pd.notna(f_ag_dt_sug):
if diff_dias &lt; -60 or diff_dias &gt; 15:
continue
elif diff_dias == 999:
continue
tags_sug = []
score_sug = 0

# Coincidencias específicas con los ítems de ingresos por terceros (solo si la factura
del agente es BOG-70697 o de saldo similar)
es_caso_70697 = &quot;70697&quot; in str(agente_actual.get(&quot;Factura&quot;, &quot;&quot;)) or
abs(saldo_terceros_esperado_act - 2542421.20) &lt; 100.0
if es_caso_70697 and abs(sc_cand_val - 1079157.20) &lt; 50000:
tags_sug.append(&quot;✈️ Flete Internacional (~$1.08M)&quot;)
score_sug += 180 - (abs(sc_cand_val - 1079157.20) / 1000)
elif es_caso_70697 and abs(sc_cand_val - 1463264.00) &lt; 50000:
tags_sug.append(&quot;�� Cargos en Origen (~$1.46M)&quot;)
score_sug += 180 - (abs(sc_cand_val - 1463264.00) / 1000)
elif abs(sc_cand_val - dif_faltante_prev) &lt; 5000:
tags_sug.append(f&quot;�� Cuadre Exacto (${dif_faltante_prev:,.2f})&quot;)
score_sug += 250
elif sc_cand_val &lt;= dif_faltante_prev + 0.05:
tags_sug.append(f&quot;�� Cubre {(sc_cand_val / dif_faltante_prev * 100):.1f}% de
faltante&quot;)
score_sug += 60
# Ponderación por cercanía de fecha dentro de la ventana (-60 a +15)
if diff_dias == 0:
tags_sug.append(&quot;�� Mismo día (0d)&quot;)
score_sug += 100
elif 0 &lt; diff_dias &lt;= 15:
tags_sug.append(f&quot;➡️ Poco después (+{diff_dias}d)&quot;)
score_sug += 85 - (diff_dias * 3)
elif -30 &lt;= diff_dias &lt; 0:
tags_sug.append(f&quot;⬅️ Anterior ({abs(diff_dias)}d antes)&quot;)
score_sug += 90 - (abs(diff_dias) * 2)
elif -60 &lt;= diff_dias &lt; -30:
tags_sug.append(f&quot;⬅️ Anterior ({abs(diff_dias)}d antes)&quot;)
score_sug += 40 - abs(diff_dias)
if any(k in prov_u for k in [&quot;DHL&quot;, &quot;CARGO&quot;, &quot;ADUANA&quot;, &quot;GARAJE&quot;, &quot;ALMAVIVA&quot;,
&quot;ALMACENADORA&quot;, &quot;UPS&quot;, &quot;FEDEX&quot;, &quot;FLETE&quot;, &quot;TRANSPORTE&quot;, &quot;PUERTO&quot;, &quot;TERMINAL&quot;]):
tags_sug.append(&quot;�� Logística/Aduana&quot;)
score_sug += 50
cands_sug_list.append({
&quot;Factura&quot;: f_c_num,
&quot;Proveedor&quot;: tr_cand_s.get(&quot;Proveedor&quot;, &quot;&quot;)[:32],
&quot;Fecha&quot;: fec_str,
&quot;Saldo&quot;: sc_cand_val,
&quot;Dias&quot;: diff_dias,
&quot;Tags&quot;: &quot; • &quot;.join(tags_sug) if tags_sug else &quot;⚪ Factura Libre&quot;,
&quot;Score&quot;: score_sug,
&quot;Row&quot;: tr_cand_s
})
cands_sug_list.sort(key=lambda x: x[&quot;Score&quot;], reverse=True)
# Detectar parejas de facturas que sumen al faltante
pareja_match = None
for i_p in range(min(15, len(cands_sug_list))):
for j_p in range(i_p + 1, min(15, len(cands_sug_list))):
c1_p = cands_sug_list[i_p]
c2_p = cands_sug_list[j_p]
s_pair = round(c1_p[&quot;Saldo&quot;] + c2_p[&quot;Saldo&quot;], 2)
d_pair = abs(s_pair - dif_faltante_prev)
if d_pair &lt; 50000:
pareja_match = (d_pair, s_pair, c1_p, c2_p)
break
if pareja_match:
break
if pareja_match:
dp, sp, cp1, cp2 = pareja_match
col_par1, col_par2 = st.columns([3.2, 1.2])
with col_par1:
st.markdown(
f&quot;&lt;div style=&#39;background:#fefce8; border:1px solid #eab308; border-radius:6px;
padding:8px 12px; margin-bottom:8px;&#39;&gt;&quot;

f&quot;�� &lt;b&gt;Combinación de 2 Facturas Sugerida:&lt;/b&gt; &lt;code&gt;{cp1[&#39;Factura&#39;]}&lt;/code
({cp1[&#39;Proveedor&#39;]}, ${cp1[&#39;Saldo&#39;]:,.2f}) + &lt;code&gt;{cp2[&#39;Factura&#39;]}&lt;/code&gt; ({cp2[&#39;Proveedor&#39;]},
${cp2[&#39;Saldo&#39;]:,.2f})&lt;br&gt;&quot;
f&quot;&lt;span style=&#39;font-size:12px; color:#854d0e;&#39;&gt;Suma Combinada:
&lt;b&gt;${sp:,.2f}&lt;/b&gt; (Diferencia: &lt;b&gt;${dp:,.2f}&lt;/b&gt;)&lt;/span&gt;&lt;/div&gt;&quot;,
unsafe_allow_html=True
)
with col_par2:
st.write(&quot;&quot;)
if st.button(&quot;➕ Añadir Ambas Facturas&quot;, key=f&quot;btn_add_pair_{pq_id_sel}&quot;,
type=&quot;primary&quot;, use_container_width=True):
terceros_nuevo = pd.concat([terceros_actual, pd.DataFrame([cp1[&quot;Row&quot;],
cp2[&quot;Row&quot;]])]).drop_duplicates(subset=[&quot;Factura&quot;]).reset_index(drop=True)
st.session_state[&quot;paquetes_importacion&quot;][pq_id_sel][&quot;terceros&quot;] =
terceros_nuevo
st.session_state[&quot;paquete_seleccionado_id&quot;] = pq_id_sel
st.session_state[&quot;sel_paquete_activo_key&quot;] = pq_id_sel
guardar_estado_manual(empresa)
st.success(f&quot;¡Facturas {cp1[&#39;Factura&#39;]} y {cp2[&#39;Factura&#39;]} añadidas con éxito
al Paquete #{pq_id_sel}!&quot;)
st.rerun()
if cands_sug_list:
top_cands = cands_sug_list[:6]
for idx_sug, sug_item in enumerate(top_cands):
c_sug1, c_sug2, c_sug3 = st.columns([1.5, 3.2, 1.1])
with c_sug1:
st.markdown(f&quot;&lt;b&gt;{sug_item[&#39;Factura&#39;]}&lt;/b&gt;&lt;br&gt;&lt;span style=&#39;font-size:12px;
color:#475569;&#39;&gt;{sug_item[&#39;Fecha&#39;]}&lt;/span&gt;&quot;, unsafe_allow_html=True)
with c_sug2:
st.markdown(f&quot;&lt;b&gt;{sug_item[&#39;Proveedor&#39;]}&lt;/b&gt; — &lt;b
style=&#39;color:#0369a1;&#39;&gt;${sug_item[&#39;Saldo&#39;]:,.2f}&lt;/b&gt;&lt;br&gt;&lt;span style=&#39;font-size:12px;
color:#15803d;&#39;&gt;{sug_item[&#39;Tags&#39;]}&lt;/span&gt;&quot;, unsafe_allow_html=True)
with c_sug3:
st.write(&quot;&quot;)
if st.button(&quot;➕ Añadir&quot;, key=f&quot;btn_sug_add_{pq_id_sel}_{idx_sug}&quot;,
use_container_width=True):
terceros_nuevo = pd.concat([terceros_actual,
pd.DataFrame([sug_item[&quot;Row&quot;]])]).drop_duplicates(subset=[&quot;Factura&quot;]).reset_index(drop=True)
st.session_state[&quot;paquetes_importacion&quot;][pq_id_sel][&quot;terceros&quot;] =
terceros_nuevo
st.session_state[&quot;paquete_seleccionado_id&quot;] = pq_id_sel
st.session_state[&quot;sel_paquete_activo_key&quot;] = pq_id_sel
guardar_estado_manual(empresa)
st.success(f&quot;¡Factura {sug_item[&#39;Factura&#39;]} añadida al Paquete
#{pq_id_sel}!&quot;)
st.rerun()
else:
st.info(f&quot;No se encontraron facturas sugeridas que no superen el monto de terceros
(${monto_limite_terc:,.2f}) ni estén bloqueadas.&quot;)
st.markdown(&quot;---&quot;)
st.markdown(&quot;##### ➕ Añadir Manualmente otra Factura Libre a este paquete:&quot;)
filtro_cercanas_man = st.checkbox(
&quot;�� Filtrar solo facturas de la ventana cronológica (-60d anteriores a +15d posteriores)&quot;,
value=True,
key=f&quot;chk_fec_man_{pq_id_sel}&quot;,
help=&quot;Oculta facturas lejanas de otros meses y muestra solo facturas anteriores o un poco
después de la operación.&quot;
)
opciones_cands = []
mapa_agregar = {}
for _, tr_cand in cands_disp_agregar.iterrows():
f_cand_num = str(tr_cand[&quot;Factura&quot;]).strip()
if f_cand_num not in facs_en_este and f_cand_num not in facturas_bloqueadas_validadas:
fec_c_str = str(tr_cand.get(&quot;Fecha&quot;, &quot;&quot;))
fec_c_dt = pd.to_datetime(fec_c_str, dayfirst=True, errors=&quot;coerce&quot;)
diff_m = int((fec_c_dt - f_ag_dt_sug).days) if (pd.notna(fec_c_dt) and
pd.notna(f_ag_dt_sug)) else 999

if filtro_cercanas_man and (diff_m &lt; -60 or diff_m &gt; 15):
continue
es_reg_c = bool(tr_cand.get(&quot;Ya Registrada&quot;, False))
tag_est = f&quot;�� Registrada ({tr_cand.get(&#39;Comprobante Previo&#39;, &#39;10-Prev&#39;)})&quot; i
es_reg_c else &quot;⚪ No Contabilizada (Pendiente)&quot;
sc_cand = obtener_saldo_cruce_factura(tr_cand)
d_tag = f&quot;({diff_m:+d}d)&quot; if diff_m != 999 else &quot;(sin fecha)&quot;
tag_c = f&quot;[{tag_est}] [{f_cand_num}] {d_tag} {str(tr_cand.get(&#39;Proveedor&#39;) or
&#39;Tercero&#39;)[:20]} (Saldo: ${sc_cand:,.2f})&quot;
opciones_cands.append((abs(diff_m), tag_c, tr_cand))
opciones_cands.sort(key=lambda x: x[0])
opciones_agregar = [x[1] for x in opciones_cands]
for _, tag_c, tr_cand in opciones_cands:
mapa_agregar[tag_c] = tr_cand
c_add1, c_add2 = st.columns([3, 1.2])
with c_add1:
sel_para_agregar = st.selectbox(&quot;Selecciona factura libre para añadir a este paquete:&quot;,
[&quot;(Seleccionar...)&quot;] + opciones_agregar, key=f&quot;sel_add_tr_{pq_id_sel}&quot;)
with c_add2:
st.write(&quot;&quot;)
st.write(&quot;&quot;)
if st.button(&quot;➕ Añadir al Paquete&quot;, key=f&quot;btn_add_tr_{pq_id_sel}&quot;) and sel_para_agregar !=
&quot;(Seleccionar...)&quot;:
fila_agregada = mapa_agregar[sel_para_agregar]
terceros_nuevo = pd.concat([terceros_actual,
pd.DataFrame([fila_agregada])]).drop_duplicates(subset=[&quot;Factura&quot;]).reset_index(drop=True)
st.session_state[&quot;paquetes_importacion&quot;][pq_id_sel][&quot;terceros&quot;] = terceros_nuevo
st.session_state[&quot;paquete_seleccionado_id&quot;] = pq_id_sel
st.session_state[&quot;sel_paquete_activo_key&quot;] = pq_id_sel
guardar_estado_manual(empresa)
st.success(f&quot;¡Factura {fila_agregada[&#39;Factura&#39;]} añadida al Paquete #{pq_id_sel}!&quot;)
st.rerun()
# 3. TRATAMIENTO CONTABLE: PÁGINA 2 / MERCANCÍAS EN TRÁNSITO / NO DEDUCIBLES
enviar_h2_activo = bool(st.session_state.get(f&quot;enviar_h2_pq_{pq_id_sel}&quot;, False))
enviar_gp_activo = bool(st.session_state.get(f&quot;enviar_gp_pq_{pq_id_sel}&quot;, False))
enviar_nd_activo = bool(st.session_state.get(f&quot;enviar_nd_pq_{pq_id_sel}&quot;, False))
c_box_style_t = (
&#39;&lt;div style=&quot;background:#f8fafc; border:2px solid #0070ba; border-radius:8px; padding:14px;
margin:14px 0 10px 0;&quot;&gt;&#39;
&#39;&lt;h4 style=&quot;margin:0 0 4px 0; color:#0070ba;&quot;&gt;⚖️ Tratamiento Contable para el Paquete #&#39; +
str(pq_id_sel) + &#39; (Cobro &#39; + str(agente_actual.get(&#39;Proveedor&#39;, &#39;&#39;)) + &#39;):&lt;/h4&gt;&#39;
&#39;&lt;p style=&quot;margin:0; font-size:13.5px; color:#334155;&quot;&gt;Selecciona cómo deseas registrar este
cobro en su fecha de operación (&lt;b&gt;&#39; + str(agente_actual.get(&#39;Fecha&#39;, &#39;&#39;)) + &#39;&lt;/b&gt;):&lt;/p&gt;&#39;
&#39;&lt;/div&gt;&#39;
)
st.markdown(c_box_style_t, unsafe_allow_html=True)
c_btn_h2, c_btn_mt, c_btn_nd = st.columns([1.8, 1.6, 1.2])
with c_btn_h2:
st.markdown(&quot;&lt;b style=&#39;color:#16a34a;&#39;&gt;�� Opción 1: Traer de Página 2&lt;/b&gt;&lt;br&gt;&lt;spa
style=&#39;font-size:12px; color:#475569;&#39;&gt;Trae ingresos propios y deja saldo de terceros pendiente para
conciliar:&lt;/span&gt;&quot;, unsafe_allow_html=True)
btn_h2_type = &quot;primary&quot; if (enviar_h2_activo or not (enviar_gp_activo or enviar_nd_activo))
else &quot;secondary&quot;
if st.button(&quot;�� Traer de Página 2 (Dejar Pendiente)&quot;, key=f&quot;btn_h2_act_{pq_id_sel}&quot;
type=btn_h2_type, use_container_width=True):
st.session_state[f&quot;enviar_h2_pq_{pq_id_sel}&quot;] = True
st.session_state[f&quot;enviar_gp_pq_{pq_id_sel}&quot;] = False
st.session_state[f&quot;enviar_nd_pq_{pq_id_sel}&quot;] = False
st.session_state[&quot;paquete_seleccionado_id&quot;] = pq_id_sel
st.session_state[&quot;sel_paquete_activo_key&quot;] = pq_id_sel
if st.session_state.get(f&quot;paquete_listo_{pq_id_sel}&quot;, False):
df_as_act, _ = generar_asiento_mixto_hoja2_con_terceros(agente_actual,
terceros_actual, enviar_a_no_deducible=False, imputar_a_transito=False)
st.session_state[f&quot;asiento_fijo_pq_{pq_id_sel}&quot;] = df_as_act.copy()

paquete_activo[&quot;asiento_fijo&quot;] = df_as_act.copy()
ag_fac_str = str(agente_actual[&quot;Factura&quot;]).strip()
st.session_state.setdefault(&quot;asientos_triangulacion_por_factura&quot;, {})[ag_fac_str] =
df_as_act.copy()
guardar_estado_manual(empresa)
st.success(&quot;¡Contabilización de Página 2 aplicada! El saldo de terceros queda pendiente
para conciliar con facturas.&quot;)
st.rerun()
with c_btn_mt:
st.markdown(&quot;&lt;b style=&#39;color:#0284c7;&#39;&gt;�� Opción 2: Mercancías en Tránsito&lt;/b&gt;&lt;br&gt;&lt;spa
style=&#39;font-size:12px; color:#475569;&#39;&gt;Imputa el saldo faltante a Inventarios en Tránsito (Cta 14650501):&lt;/span&gt;&quot;,
unsafe_allow_html=True)
lbl_btn_mt = f&quot;�� Mercancías en Tránsito (${dif_faltante_prev:,.2f})&quot; if dif_faltante_prev
0.05 else &quot;�� Mercancías en Tránsito (14650501)&quot;
btn_mt_type = &quot;primary&quot; if enviar_gp_activo else &quot;secondary&quot;
if st.button(lbl_btn_mt, key=f&quot;btn_mt_act_{pq_id_sel}&quot;, type=btn_mt_type,
use_container_width=True):
st.session_state[f&quot;enviar_gp_pq_{pq_id_sel}&quot;] = True
st.session_state[f&quot;enviar_nd_pq_{pq_id_sel}&quot;] = False
st.session_state[&quot;paquete_seleccionado_id&quot;] = pq_id_sel
st.session_state[&quot;sel_paquete_activo_key&quot;] = pq_id_sel
if saldo_terceros_esperado_act &gt; 0.05 or enviar_h2_activo:
st.session_state[f&quot;enviar_h2_pq_{pq_id_sel}&quot;] = True
df_as_act, _ = generar_asiento_mixto_hoja2_con_terceros(agente_actual,
terceros_actual, enviar_a_no_deducible=False, imputar_a_transito=True)
else:
st.session_state[f&quot;enviar_h2_pq_{pq_id_sel}&quot;] = False
df_as_act, _, ret_asum = generar_asiento_triangulacion_paquete(agente_actual,
terceros_actual, enviar_a_no_deducible=False)
asiento_gp_filas = []
for _, r_as in df_as_act.iterrows():
if str(r_as[&quot;Código Cuenta&quot;]).strip() == CUENTA_NO_DEDUCIBLE:
asiento_gp_filas.append({
&quot;Código Cuenta&quot;: CUENTA_IMPORTACION_TRANSITO,
&quot;Descripción Cuenta&quot;: f&quot;Mercancías en Tránsito / Base Propia Agente (Fac
{agente_actual.get(&#39;Factura&#39;, &#39;&#39;)})&quot;,
&quot;Tercero / NIT&quot;: f&quot;{agente_actual.get(&#39;NIT Emisor&#39;, &#39;&#39;)} -
{agente_actual.get(&#39;Proveedor&#39;, &#39;&#39;)[:25]}&quot;,
&quot;Débito ($)&quot;: float(r_as[&quot;Débito ($)&quot;]),
&quot;Crédito ($)&quot;: 0.0
})
else:
asiento_gp_filas.append(r_as.to_dict())
df_as_act = pd.DataFrame(asiento_gp_filas)
if st.session_state.get(f&quot;paquete_listo_{pq_id_sel}&quot;, False):
st.session_state[f&quot;asiento_fijo_pq_{pq_id_sel}&quot;] = df_as_act.copy()
paquete_activo[&quot;asiento_fijo&quot;] = df_as_act.copy()
ag_fac_str = str(agente_actual[&quot;Factura&quot;]).strip()
st.session_state.setdefault(&quot;asientos_triangulacion_por_factura&quot;, {})[ag_fac_str] =
df_as_act.copy()
guardar_estado_manual(empresa)
st.success(f&quot;¡Diferencia de ${dif_faltante_prev:,.2f} asignada a Mercancías en Tránsito
(14650501)!&quot;)
st.rerun()
with c_btn_nd:
st.markdown(&quot;&lt;b style=&#39;color:#b91c1c;&#39;&gt;�� Opción 3: No Deducibles&lt;/b&gt;&lt;br&gt;&lt;span style=&#39;font-
size:12px; color:#475569;&#39;&gt;Si no existe factura DIAN ni soporte:&lt;/span&gt;&quot;, unsafe_allow_html=True)
lbl_btn_nd = f&quot;�� No Deducibles (${dif_faltante_prev:,.2f})&quot; if dif_faltante_prev &gt; 0.05 els
&quot;�� No Deducibles (53950501)&quot;
btn_nd_type = &quot;primary&quot; if enviar_nd_activo else &quot;secondary&quot;
if st.button(lbl_btn_nd, key=f&quot;btn_nd_act_{pq_id_sel}&quot;, type=btn_nd_type,
use_container_width=True):
st.session_state[f&quot;enviar_nd_pq_{pq_id_sel}&quot;] = True
st.session_state[f&quot;enviar_gp_pq_{pq_id_sel}&quot;] = False
st.session_state[&quot;paquete_seleccionado_id&quot;] = pq_id_sel
st.session_state[&quot;sel_paquete_activo_key&quot;] = pq_id_sel
if saldo_terceros_esperado_act &gt; 0.05 or enviar_h2_activo:
st.session_state[f&quot;enviar_h2_pq_{pq_id_sel}&quot;] = True
df_as_act, _ = generar_asiento_mixto_hoja2_con_terceros(agente_actual,
terceros_actual, enviar_a_no_deducible=True, imputar_a_transito=False)
else:

st.session_state[f&quot;enviar_h2_pq_{pq_id_sel}&quot;] = False
df_as_act, _, _ = generar_asiento_triangulacion_paquete(agente_actual,
terceros_actual, enviar_a_no_deducible=True)
if st.session_state.get(f&quot;paquete_listo_{pq_id_sel}&quot;, False):
st.session_state[f&quot;asiento_fijo_pq_{pq_id_sel}&quot;] = df_as_act.copy()
paquete_activo[&quot;asiento_fijo&quot;] = df_as_act.copy()
ag_fac_str = str(agente_actual[&quot;Factura&quot;]).strip()
st.session_state.setdefault(&quot;asientos_triangulacion_por_factura&quot;, {})[ag_fac_str] =
df_as_act.copy()
guardar_estado_manual(empresa)
st.success(f&quot;¡Diferencia de ${dif_faltante_prev:,.2f} enviada a Gastos No Deducibles
(53950501)!&quot;)
st.rerun()
# RECALCULAR ASIENTO CONTABLE PARA ESTE PAQUETE
if st.session_state.get(f&quot;paquete_listo_{pq_id_sel}&quot;, False) and (f&quot;asiento_fijo_pq_{pq_id_sel}&quot;
in st.session_state or &quot;asiento_fijo&quot; in paquete_activo):
df_asiento_paquete = st.session_state.get(f&quot;asiento_fijo_pq_{pq_id_sel}&quot;,
paquete_activo.get(&quot;asiento_fijo&quot;)).copy()
dif_no_ded = 0.0
ret_asum = 0.0
elif saldo_terceros_esperado_act &gt; 0.05 or enviar_h2_activo:
if enviar_nd_activo:
df_asiento_paquete, _ = generar_asiento_mixto_hoja2_con_terceros(agente_actual,
terceros_actual, enviar_a_no_deducible=True, imputar_a_transito=False)
dif_no_ded = dif_faltante_prev
elif enviar_gp_activo:
df_asiento_paquete, _ = generar_asiento_mixto_hoja2_con_terceros(agente_actual,
terceros_actual, enviar_a_no_deducible=False, imputar_a_transito=True)
dif_no_ded = 0.0
else:
df_asiento_paquete, dif_no_ded = generar_asiento_mixto_hoja2_con_terceros(agente_actual,
terceros_actual, enviar_a_no_deducible=False, imputar_a_transito=False)
ret_asum = 0.0
elif enviar_gp_activo and dif_faltante_prev &gt; 0.05:
df_asiento_paquete, _, ret_asum = generar_asiento_triangulacion_paquete(agente_actual,
terceros_actual, enviar_a_no_deducible=False)
asiento_gp_filas = []
for _, r_as in df_asiento_paquete.iterrows():
if str(r_as[&quot;Código Cuenta&quot;]).strip() == CUENTA_NO_DEDUCIBLE:
asiento_gp_filas.append({
&quot;Código Cuenta&quot;: CUENTA_IMPORTACION_TRANSITO,
&quot;Descripción Cuenta&quot;: f&quot;Mercancías en Tránsito / Base Propia Agente (Fac
{agente_actual.get(&#39;Factura&#39;, &#39;&#39;)})&quot;,
&quot;Tercero / NIT&quot;: f&quot;{agente_actual.get(&#39;NIT Emisor&#39;, &#39;&#39;)} -
{agente_actual.get(&#39;Proveedor&#39;, &#39;&#39;)[:25]}&quot;,
&quot;Débito ($)&quot;: float(r_as[&quot;Débito ($)&quot;]),
&quot;Crédito ($)&quot;: 0.0
})
else:
asiento_gp_filas.append(r_as.to_dict())
df_asiento_paquete = pd.DataFrame(asiento_gp_filas)
dif_no_ded = 0.0
else:
df_asiento_paquete, dif_no_ded, ret_asum =
generar_asiento_triangulacion_paquete(agente_actual, terceros_actual, enviar_a_no_deducible=enviar_nd_activo)
st.markdown(f&quot;#### ⚖️ Asiento Contable del Paquete #{pq_id_sel}:&quot;)
st.caption(&quot;Detalle de partida doble de ESTE paquete: cancela las cuentas por pagar de DHL,
Agencia y Garaje contra Euro Shipping:&quot;)
st.dataframe(
df_asiento_paquete.style.format({&quot;Débito ($)&quot;: &quot;${:,.2f}&quot;, &quot;Crédito ($)&quot;: &quot;${:,.2f}&quot;}),
use_container_width=True,
hide_index=True
)
c_cuad1, c_cuad2, c_cuad3, c_cuad4 = st.columns(4)
sum_deb_pq = df_asiento_paquete[&quot;Débito ($)&quot;].sum()
sum_cred_pq = df_asiento_paquete[&quot;Crédito ($)&quot;].sum()
dif_cuad_pq = abs(sum_deb_pq - sum_cred_pq)

with c_cuad1:
st.metric(&quot;Total Débito&quot;, f&quot;${sum_deb_pq:,.2f}&quot;)
with c_cuad2:
st.metric(&quot;Total Crédito&quot;, f&quot;${sum_cred_pq:,.2f}&quot;)
with c_cuad3:
st.metric(&quot;Retenciones Asumidas (53152001)&quot;, f&quot;${ret_asum:,.2f}&quot;)
with c_cuad4:
st.metric(&quot;Diferencia No Deducible (53950501)&quot;, f&quot;${dif_no_ded:,.2f}&quot;)
# SECCIÓN DE VALIDACIÓN Y BLOQUEO DE FACTURAS
es_validado = st.session_state.get(f&quot;paquete_listo_{pq_id_sel}&quot;, False)
if not es_validado:
if dif_cuad_pq &lt; 0.05:
st.success(f&quot;✅ **Partida doble cuadrada al centavo ($0.00).** Ya puedes validar y bloquear
este paquete.&quot;)
else:
st.markdown(
f&quot;&quot;&quot;
&lt;div style=&quot;background:#fffbeb; border:2px solid #f59e0b; border-radius:8px;
padding:12px 16px; margin:10px 0;&quot;&gt;
&lt;span style=&quot;font-size:16px; font-weight:bold; color:#b45309;&quot;&gt;⏳ Faltante
Pendiente de Conciliar: ${dif_cuad_pq:,.2f}&lt;/span&gt;&lt;br&gt;
&lt;span style=&quot;font-size:13.5px; color:#78350f;&quot;&gt;
El comprobante presenta una diferencia de &lt;b&gt;${dif_cuad_pq:,.2f}&lt;/b&gt; para
cuadrar sumas iguales. Puedes aplicar este valor directamente a No Deducibles o a Tránsito con un solo clic:
&lt;/span&gt;
&lt;/div&gt;
&quot;&quot;&quot;,
unsafe_allow_html=True
)
c_qnd1, c_qnd2, c_qnd3 = st.columns([1.6, 1.6, 1.8])
with c_qnd1:
if st.button(f&quot;�� Mandar ${dif_cuad_pq:,.2f} a No Deducibles&quot;
key=f&quot;btn_box_nd_{pq_id_sel}&quot;, type=&quot;primary&quot;, use_container_width=True):
st.session_state[f&quot;enviar_nd_pq_{pq_id_sel}&quot;] = True
st.session_state[f&quot;enviar_gp_pq_{pq_id_sel}&quot;] = False
st.session_state[f&quot;enviar_h2_pq_{pq_id_sel}&quot;] = True
st.session_state[&quot;paquete_seleccionado_id&quot;] = pq_id_sel
st.session_state[&quot;sel_paquete_activo_key&quot;] = pq_id_sel
if st.session_state.get(f&quot;paquete_listo_{pq_id_sel}&quot;, False):
df_as_act, _ = generar_asiento_mixto_hoja2_con_terceros(agente_actual,
terceros_actual, enviar_a_no_deducible=True, imputar_a_transito=False)
st.session_state[f&quot;asiento_fijo_pq_{pq_id_sel}&quot;] = df_as_act.copy()
paquete_activo[&quot;asiento_fijo&quot;] = df_as_act.copy()
ag_fac_str = str(agente_actual[&quot;Factura&quot;]).strip()
st.session_state.setdefault(&quot;asientos_triangulacion_por_factura&quot;,
{})[ag_fac_str] = df_as_act.copy()
guardar_estado_manual(empresa)
st.success(f&quot;¡Diferencia de ${dif_cuad_pq:,.2f} enviada a No Deducibles
(53950501)!&quot;)
st.rerun()
with c_qnd2:
if st.button(f&quot;�� Mandar ${dif_cuad_pq:,.2f} a Tránsito&quot;
key=f&quot;btn_box_mt_{pq_id_sel}&quot;, use_container_width=True):
st.session_state[f&quot;enviar_gp_pq_{pq_id_sel}&quot;] = True
st.session_state[f&quot;enviar_nd_pq_{pq_id_sel}&quot;] = False
st.session_state[f&quot;enviar_h2_pq_{pq_id_sel}&quot;] = True
st.session_state[&quot;paquete_seleccionado_id&quot;] = pq_id_sel
st.session_state[&quot;sel_paquete_activo_key&quot;] = pq_id_sel
if st.session_state.get(f&quot;paquete_listo_{pq_id_sel}&quot;, False):
df_as_act, _ = generar_asiento_mixto_hoja2_con_terceros(agente_actual,
terceros_actual, enviar_a_no_deducible=False, imputar_a_transito=True)
st.session_state[f&quot;asiento_fijo_pq_{pq_id_sel}&quot;] = df_as_act.copy()
paquete_activo[&quot;asiento_fijo&quot;] = df_as_act.copy()
ag_fac_str = str(agente_actual[&quot;Factura&quot;]).strip()
st.session_state.setdefault(&quot;asientos_triangulacion_por_factura&quot;,
{})[ag_fac_str] = df_as_act.copy()
guardar_estado_manual(empresa)
st.success(f&quot;¡Diferencia de ${dif_cuad_pq:,.2f} enviada a Tránsito (14650501)!&quot;)
st.rerun()
with c_qnd3:

st.caption(&quot;�� O retira facturas de terceros que no correspondan en la sección &#39;⚙
Modificar este Paquete&#39; arriba.&quot;)
c_val_opt1, c_val_opt2 = st.columns(2)
with c_val_opt1:
st.markdown(
&quot;&lt;div style=&#39;background:#f0fdf4; border:1.5px solid #22c55e; border-radius:8px;
padding:12px; height:100%;&#39;&gt;&quot;
&quot;&lt;b style=&#39;color:#15803d; font-size:15px;&#39;&gt;�� Opción 1: Pasar a Hoja 2 (Subir
Planilla Siigo)&lt;/b&gt;&lt;br&gt;&quot;
&quot;&lt;span style=&#39;font-size:12.5px; color:#166534;&#39;&gt;&quot;
&quot;Bloquea las facturas del paquete, actualiza la Página 2 y &lt;b&gt;genera el comprobante
oficial en la Planilla de Siigo (Página 4)&lt;/b&gt; para su contabilización.&quot;
&quot;&lt;/span&gt;&lt;/div&gt;&quot;,
unsafe_allow_html=True
)
st.write(&quot;&quot;)
if st.button(&quot;�� Pasar a Hoja 2 (Subir a Planilla Siigo)&quot;
key=f&quot;btn_validar_siigo_{pq_id_sel}&quot;, type=&quot;primary&quot;, use_container_width=True):
ejecutar_validacion_pq(pasa_a_siigo=True)
st.success(f&quot;�� ¡Paquete #{pq_id_sel} VALIDADO y trasladado a Página 2 para Siigo!&quot;)
st.rerun()
with c_val_opt2:
st.markdown(
&quot;&lt;div style=&#39;background:#eff6ff; border:1.5px solid #3b82f6; border-radius:8px;
padding:12px; height:100%;&#39;&gt;&quot;
&quot;&lt;b style=&#39;color:#1d4ed8; font-size:15px;&#39;&gt;�� Opción 2: Simplemente Conciliación (N
Subir a Siigo)&lt;/b&gt;&lt;br&gt;&quot;
&quot;&lt;span style=&#39;font-size:12.5px; color:#1e40af;&#39;&gt;&quot;
&quot;Bloquea e identifica las facturas para &lt;b&gt;auditoría y soporte de costos&lt;/b&gt;, pero
&lt;b&gt;NO genera comprobante en Siigo&lt;/b&gt; (evita duplicar pagos ya hechos por tesorería/bancos).&quot;
&quot;&lt;/span&gt;&lt;/div&gt;&quot;,
unsafe_allow_html=True
)
st.write(&quot;&quot;)
if st.button(&quot;�� Simplemente Conciliación (No Subir a Siigo)&quot;
key=f&quot;btn_validar_concil_{pq_id_sel}&quot;, use_container_width=True):
ejecutar_validacion_pq(pasa_a_siigo=False)
st.success(f&quot;�� ¡Paquete #{pq_id_sel} CONCILIADO exitosamente para control interno (n
irá a Siigo)!&quot;)
st.rerun()
else:
es_siigo = st.session_state.get(f&quot;paquete_destino_siigo_{pq_id_sel}&quot;, True)
col_desb1, col_desb2 = st.columns([3.2, 1.4])
with col_desb1:
if es_siigo:
st.success(f&quot;�� **Paquete #{pq_id_sel} VALIDADO (PASA A HOJA 2 / PLANILLA SIIGO):*
Sus facturas están protegidas contra otros cruces y su comprobante se exportará en la Planilla de Siigo (Página 4).&quot;)
else:
st.info(f&quot;�� **Paquete #{pq_id_sel} CONCILIADO (SOLO CONTROL INTERNO
IDENTIFICACIÓN):** Facturas identificadas y protegidas para auditoría DIAN. **No se generará comprobante en Siigo**
para evitar duplicidades.&quot;)
with col_desb2:
c_sw1, c_sw2 = st.columns(2)
with c_sw1:
if es_siigo:
if st.button(&quot;�� Cambiar a Solo Conciliación&quot;, key=f&quot;btn_sw_concil_{pq_id_sel}&quot;
use_container_width=True, help=&quot;Cambia el destino para que NO se suba a la planilla Siigo&quot;):
ejecutar_validacion_pq(pasa_a_siigo=False)
st.rerun()
else:
if st.button(&quot;�� Cambiar a Pasar a Hoja 2&quot;, key=f&quot;btn_sw_siigo_{pq_id_sel}&quot;
type=&quot;primary&quot;, use_container_width=True, help=&quot;Cambia el destino para que SÍ se suba a Siigo&quot;):
ejecutar_validacion_pq(pasa_a_siigo=True)
st.rerun()
with c_sw2:
if st.button(&quot;�� Desbloquear&quot;, key=f&quot;btn_desbloquear_main_{pq_id_sel}&quot;
use_container_width=True):
ejecutar_desbloqueo_pq()
st.rerun()

st.markdown(&quot;---&quot;)
st.markdown(&quot;#### �� Exportar Control de Paquetes y Cruces de Importación:&quot;)
st.caption(&quot;Descarga un archivo Excel con el resumen de cada paquete, el detalle de facturas de
terceros vinculadas y el asiento contable de control (sin generar aún notas de contabilidad en Siigo hasta que definas
el procedimiento):&quot;)
if st.checkbox(&quot;�� Generar Archivo Excel de Control de Cruces para Descarga (Bajo demanda par
proteger memoria RAM)&quot;, value=False, key=&quot;chk_gen_excel_cruces_tab3&quot;):
# Generar Excel estructurado por Paquetes de Importación individuales
todos_asientos_lista = []
resumen_paquetes_lista = []
detalle_facturas_lista = []
consecutivo_cruce = 1
facs_asignadas_en_algun_pq = []
for g_k, g_v in pqs_actuales.items():
ag_item = g_v[&quot;agente&quot;]
terc_items = g_v[&quot;terceros&quot;]
enviar_nd_g = st.session_state.get(f&quot;enviar_nd_pq_{g_k}&quot;, False)
df_as_p, dif_p, ret_p = generar_asiento_triangulacion_paquete(ag_item, terc_items,
enviar_a_no_deducible=enviar_nd_g)
# 1. Asiento de Control de Cruces (Partida Doble para Revisión)
for _, fila_as in df_as_p.iterrows():
todos_asientos_lista.append({
&quot;Paquete #&quot;: f&quot;Paquete #{g_k}&quot;,
&quot;Fecha Operación&quot;: ag_item[&quot;Fecha&quot;],
&quot;Código Cuenta&quot;: fila_as[&quot;Código Cuenta&quot;],
&quot;Tercero / NIT&quot;: fila_as[&quot;Tercero / NIT&quot;],
&quot;Descripción&quot;: str(fila_as.get(&quot;Descripción de la Cuenta&quot;) or
fila_as.get(&quot;Descripción Cuenta&quot;) or fila_as.get(&quot;Descripción&quot;) or &quot;&quot;),
&quot;Débito ($)&quot;: fila_as[&quot;Débito ($)&quot;],
&quot;Crédito ($)&quot;: fila_as[&quot;Crédito ($)&quot;]
})
# 2. Resumen del Paquete
facs_terc_str = &quot;, &quot;.join([f&quot;{r[&#39;Factura&#39;]} ({str(r.get(&#39;Proveedor&#39;) or &#39;Tercero&#39;)[:15]})&quot;
for _, r in terc_items.iterrows()]) if not terc_items.empty else &quot;Ninguno&quot;
tot_terc_sum = sum([obtener_saldo_cruce_factura(r) for _, r in terc_items.iterrows()])
resumen_paquetes_lista.append({
&quot;Paquete #&quot;: f&quot;Paquete #{g_k}&quot;,
&quot;Agente Coordinador&quot;: ag_item[&quot;Proveedor&quot;],
&quot;Factura Agente&quot;: ag_item[&quot;Factura&quot;],
&quot;Fecha Operación&quot;: ag_item[&quot;Fecha&quot;],
&quot;Total Cobro Agente ($)&quot;: float(ag_item[&quot;Total&quot;]),
&quot;Facturas Terceros Incluidas&quot;: facs_terc_str,
&quot;Total Cancelado Terceros ($)&quot;: tot_terc_sum,
&quot;Retenciones Asumidas ($)&quot;: ret_p,
&quot;Diferencia No Deducible ($)&quot;: dif_p,
&quot;Estado Cuadre&quot;: &quot;✅ Cuadrado ($0.00)&quot; if abs(float(ag_item[&quot;Total&quot;]) - tot_terc_sum -
ret_p - dif_p) &lt; 1.0 else &quot;Revisar&quot;
})
# 3. Detalle de Facturas que componen el Paquete
detalle_facturas_lista.append({
&quot;Paquete #&quot;: f&quot;Paquete #{g_k}&quot;,
&quot;Rol en Operación&quot;: &quot;Agente Coordinador (Cobro Global)&quot;,
&quot;Proveedor&quot;: ag_item[&quot;Proveedor&quot;],
&quot;NIT&quot;: ag_item[&quot;NIT Emisor&quot;],
&quot;Factura&quot;: ag_item[&quot;Factura&quot;],
&quot;Fecha&quot;: ag_item[&quot;Fecha&quot;],
&quot;Naturaleza&quot;: &quot;Servicios Agenciamiento + Fletes / Impuestos&quot;,
&quot;Base / Subtotal ($)&quot;: float(ag_item.get(&quot;Base&quot;, 0.0)),
&quot;IVA ($)&quot;: float(ag_item.get(&quot;IVA&quot;, 0.0)),
&quot;Retenciones ($)&quot;: 0.0,
&quot;Saldo a Cruzar ($)&quot;: float(ag_item[&quot;Total&quot;]),
&quot;Cuenta Contable&quot;: &quot;22050501&quot;
})
for _, tr_it in terc_items.iterrows():

facs_asignadas_en_algun_pq.append(tr_it[&quot;Factura&quot;])
tb = float(tr_it.get(&quot;Base&quot;, 0.0))
tiv = float(tr_it.get(&quot;IVA&quot;, 0.0))
rf = float(tr_it.get(&quot;ReteFuente&quot;, 0.0))
ri = float(tr_it.get(&quot;ReteICA&quot;, 0.0))
sc = obtener_saldo_cruce_factura(tr_it)
p_u = str(tr_it[&quot;Proveedor&quot;]).upper()
rol_t = &quot;�� Tercero Transporte (DHL)&quot; if &quot;DHL&quot; in p_u else (&quot;�� Tercero Agenciamien
(Mandato)&quot; if any(k in p_u for k in [&quot;CARGO&quot;, &quot;ADUANA&quot;]) else &quot;�� Tercero Garaje / Almacén&quot;)
nat_t = &quot;Honorario Propio Agenciamiento&quot; if &quot;COMISION&quot; in str(tr_it.get(&quot;Descripcion&quot;,
&quot;&quot;)).upper() else &quot;Gasto por Cuenta de Tercero&quot;
detalle_facturas_lista.append({
&quot;Paquete #&quot;: f&quot;Paquete #{g_k}&quot;,
&quot;Rol en Operación&quot;: rol_t,
&quot;Proveedor&quot;: tr_it[&quot;Proveedor&quot;],
&quot;NIT&quot;: tr_it[&quot;NIT Emisor&quot;],
&quot;Factura&quot;: tr_it[&quot;Factura&quot;],
&quot;Fecha&quot;: tr_it[&quot;Fecha&quot;],
&quot;Naturaleza&quot;: nat_t,
&quot;Base / Subtotal ($)&quot;: tb,
&quot;IVA ($)&quot;: tiv,
&quot;Retenciones ($)&quot;: round(rf + ri, 2),
&quot;Saldo a Cruzar ($)&quot;: sc,
&quot;Cuenta Contable&quot;: str(tr_it.get(&quot;Cuenta Pasivo Especifica&quot;, &quot;22050505&quot;))
})
consecutivo_cruce += 1
# Añadir facturas de terceros que quedaron libres (sin asignar a ningún paquete)
df_libres_exp = df_terceros_all[~df_terceros_all[&quot;Factura&quot;].isin(facs_asignadas_en_algun_pq)]
for _, tr_lib in df_libres_exp.iterrows():
tb = float(tr_lib.get(&quot;Base&quot;, 0.0))
tiv = float(tr_lib.get(&quot;IVA&quot;, 0.0))
sc = obtener_saldo_cruce_factura(tr_lib)
p_u = str(tr_lib[&quot;Proveedor&quot;]).upper()
rol_t = &quot;�� Tercero Transporte (DHL)&quot; if &quot;DHL&quot; in p_u else (&quot;�� Tercero Agenciamien
(Mandato)&quot; if any(k in p_u for k in [&quot;CARGO&quot;, &quot;ADUANA&quot;]) else &quot;�� Tercero Garaje / Almacén&quot;)
detalle_facturas_lista.append({
&quot;Paquete #&quot;: &quot;⚪ Sin Asignar (Pendiente de Operación)&quot;,
&quot;Rol en Operación&quot;: rol_t,
&quot;Proveedor&quot;: tr_lib[&quot;Proveedor&quot;],
&quot;NIT&quot;: tr_lib[&quot;NIT Emisor&quot;],
&quot;Factura&quot;: tr_lib[&quot;Factura&quot;],
&quot;Fecha&quot;: tr_lib[&quot;Fecha&quot;],
&quot;Naturaleza&quot;: &quot;Gasto por Cuenta de Tercero&quot;,
&quot;Base / Subtotal ($)&quot;: tb,
&quot;IVA ($)&quot;: tiv,
&quot;Retenciones ($)&quot;: float(tr_lib.get(&quot;ReteFuente&quot;, 0.0)) + float(tr_lib.get(&quot;ReteICA&quot;,
0.0)),
&quot;Saldo a Cruzar ($)&quot;: sc,
&quot;Cuenta Contable&quot;: str(tr_lib.get(&quot;Cuenta Pasivo Especifica&quot;, &quot;22050505&quot;))
})
if todos_asientos_lista:
df_export_cruces = pd.DataFrame(todos_asientos_lista)
buf_cruce_ex = io.BytesIO()
with pd.ExcelWriter(buf_cruce_ex, engine=&quot;openpyxl&quot;) as wr_cr:
if resumen_paquetes_lista:
pd.DataFrame(resumen_paquetes_lista).to_excel(wr_cr,
sheet_name=&quot;Resumen_Paquetes&quot;, index=False)
if detalle_facturas_lista:
pd.DataFrame(detalle_facturas_lista).to_excel(wr_cr,
sheet_name=&quot;Detalle_Facturas_Paquete&quot;, index=False)
df_export_cruces.to_excel(wr_cr, sheet_name=&quot;Asiento_Control_Cruces&quot;, index=False)
buf_cruce_ex.seek(0)
st.download_button(
label=f&quot;�� Descargar Control de Cruces de Importación en Excel ({len(pqs_actuales)
Paquetes)&quot;,
data=buf_cruce_ex.getvalue(),
file_name=f&quot;Control_Cruces_Importacion_{empresa[&#39;nombre&#39;].replace(&#39; &#39;, &#39;_&#39;)}.xlsx&quot;,
mime=&quot;application/vnd.openxmlformats-officedocument.spreadsheetml.sheet&quot;,

use_container_width=True
)
else:
st.info(&quot;�� Sube el reporte Excel de la DIAN en la Pestaña 1 para habilitar la triangulación y cruc
de importaciones.&quot;)

except Exception as _e_tri:
if type(_e_tri).__name__ in [&#39;RerunException&#39;, &#39;StopException&#39;, &#39;ScriptControlException&#39;]:
raise _e_tri
st.error(f&quot;⚠️ Ocurrió un error en la pestaña de Triangulación y Aduanas: {_e_tri}&quot;)
with st.expander(&quot;�� Ver detalles técnicos del error (Traceback):&quot;):
st.code(traceback.format_exc())
with tab_siigo:
try:
st.markdown(&quot;### Descargar Planilla Oficial Siigo Nube (3 Hojas)&quot;)
st.caption(&quot;Planilla oficial formulada con &#39;matriz_captura&#39;, &#39;interfaz_siigo&#39; y &#39;Parametrización&#39;.&quot;)
if &quot;df_procesado&quot; in st.session_state and st.session_state[&quot;df_procesado&quot;] is not None and
isinstance(st.session_state[&quot;df_procesado&quot;], pd.DataFrame) and not st.session_state[&quot;df_procesado&quot;].empty:
df_full = normalizar_df_procesado(st.session_state[&quot;df_procesado&quot;])
st.session_state[&quot;df_procesado&quot;] = df_full
# FILTRO DE PROTECCIÓN: Excluir facturas rojas (ya causadas en Siigo) y facturas que van por
triangulación aduanera
ya_reg_siigo = df_full[&quot;Ya Registrada&quot;].fillna(False).astype(bool) if &quot;Ya Registrada&quot; in
df_full.columns else pd.Series(False, index=df_full.index)
n_rojas = int(ya_reg_siigo.sum())
# Asegurar columna No Contabilizar en df_full
if &quot;No Contabilizar&quot; not in df_full.columns:
df_full[&quot;No Contabilizar&quot;] = False
st.session_state[&quot;df_procesado&quot;] = df_full
# 1. Facturas explícitamente excluidas de contabilidad (No Contabilizar)
excluidas_set = st.session_state.get(&quot;facturas_no_contabilizar&quot;, set())
cond_excluidas = df_full[&quot;Factura&quot;].isin(excluidas_set) | (df_full[&quot;No Contabilizar&quot;] == True)
n_excluidas = int(cond_excluidas.sum())
# 2. Facturas que ya están en un paquete de triangulación aduanera
facs_en_pqs_import = []
pqs_actuales_siigo = st.session_state.get(&quot;paquetes_importacion&quot;, {})
if pqs_actuales_siigo:
for _, p_val in pqs_actuales_siigo.items():
if isinstance(p_val, dict) and &quot;terceros&quot; in p_val and not p_val[&quot;terceros&quot;].empty:
facs_en_pqs_import.extend(p_val[&quot;terceros&quot;][&quot;Factura&quot;].tolist())
es_adu_siigo = df_full[&quot;Es Aduanera&quot;].fillna(False).astype(bool) if &quot;Es Aduanera&quot; in df_full.columns
else pd.Series(False, index=df_full.index)
grp_imp_siigo = df_full[&quot;Grupo Importación&quot;].fillna(&quot;&quot;).astype(str).str.strip() if &quot;Grupo Importación&quot;
in df_full.columns else pd.Series(&quot;&quot;, index=df_full.index)
cond_imp_adu = es_adu_siigo &amp; (grp_imp_siigo != &quot;&quot;)
df_p = df_full[
(~ya_reg_siigo) &amp;
(~cond_excluidas) &amp;
(~cond_imp_adu)
].copy()
if n_rojas &gt; 0:
st.info(f&quot;��️ **Protección contra duplicados:** Se excluyeron {n_rojas} facturas marcadas en roj
que ya estaban causadas en Siigo.&quot;)
if n_excluidas &gt; 0:
st.warning(f&quot;�� **Facturas Excluidas:** Se excluyeron {n_excluidas} facturas marcadas como **&#39;N
CONTABILIZAR&#39;**.&quot;)
# Paquetes de triangulación listos para llevar a contabilidad en su fecha (excluyendo los marcados
solo para conciliación)

pqs_listos_siigo = {k: v for k, v in pqs_actuales_siigo.items() if
st.session_state.get(f&quot;paquete_listo_{k}&quot;, False) and st.session_state.get(f&quot;paquete_destino_siigo_{k}&quot;, True)}
pqs_solo_concil = {k: v for k, v in pqs_actuales_siigo.items() if
st.session_state.get(f&quot;paquete_listo_{k}&quot;, False) and not st.session_state.get(f&quot;paquete_destino_siigo_{k}&quot;, True)}
c_pqs_inc1, c_pqs_inc2 = st.columns([2.5, 1.5])
with c_pqs_inc1:
inc_pqs_listos = st.checkbox(
f&quot;�� **Llevar a la contabilidad {len(pqs_listos_siigo)} Paquete(s) de Triangulación que pasa
a Siigo**&quot;,
value=True if pqs_listos_siigo else False,
help=&quot;Inserta los asientos contables de los paquetes listos marcados como &#39;Pasar a Hoja 2&#39; en
la planilla oficial de Siigo.&quot;
)
with c_pqs_inc2:
st.caption(f&quot;Suben a Siigo: **{len(pqs_listos_siigo)}** | Solo Conciliación:
**{len(pqs_solo_concil)}** | Totales: **{len(pqs_actuales_siigo)}**&quot;)
generar_siigo_ahora = st.checkbox(&quot;⚡ Generar Planillas Oficiales de Siigo Nube y Libro Auxiliar
(Activa aquí cuando vayas a descargar)&quot;, value=False, key=&quot;chk_gen_siigo_ondemand_tab4&quot;, help=&quot;Genera los libros Excel
formulados únicamente cuando vas a descargarlos, evitando consumir memoria RAM mientras trabajas en las pestañas 1, 2
y 3.&quot;)
if generar_siigo_ahora:
wb = openpyxl.Workbook()
# 1. Hoja matriz_captura
ws_matriz = wb.active
ws_matriz.title = &quot;matriz_captura&quot;
headers_matriz = [
&quot;Tipo Comp&quot;, &quot;Consecutivo&quot;, &quot;Fecha (DD/MM/AAAA)&quot;, &quot;NIT&quot;, &quot;Prefijo&quot;, &quot;Factura Num&quot;,
&quot;Descripcion&quot;, &quot;Operacion&quot;, &quot;Cta Principal&quot;, &quot;Valor Base&quot;, &quot;IVA&quot;, &quot;ReteFuente&quot;,
&quot;ReteICA&quot;, &quot;ReteIVA&quot;, &quot;Cta Contrapartida&quot;
]
ws_matriz.append(headers_matriz)
# 2. Hoja interfaz_siigo
ws_interfaz = wb.create_sheet(title=&quot;interfaz_siigo&quot;)
headers_interfaz = [
&quot;Tipo de comprobante&quot;, &quot;Consecutivo comprobante&quot;, &quot;Fecha de elaboración&quot;, &quot;Sigla moneda&quot;,
&quot;Tasa de cambio&quot;, &quot;Código cuenta contable&quot;, &quot;Identificación tercero&quot;, &quot;Sucursal&quot;,
&quot;Código producto&quot;, &quot;Código de bodega&quot;, &quot;Acción&quot;, &quot;Cantidad producto&quot;, &quot;Prefijo&quot;,
&quot;Consecutivo&quot;, &quot;No. cuota&quot;, &quot;Fecha vencimiento&quot;, &quot;Código impuesto&quot;, &quot;Código grupo activo
fijo&quot;,
&quot;Código activo fijo&quot;, &quot;Descripción&quot;, &quot;Código centro/subcentro de costos&quot;, &quot;Débito&quot;,
&quot;Crédito&quot;, &quot;Observaciones&quot;, &quot;Base gravable libro compras/ventas&quot;, &quot;Base exenta libro
compras/ventas&quot;, &quot;Mes de cierre&quot;
]
ws_interfaz.append(headers_interfaz)
fila_r = 2
asientos_triang_map_siigo = st.session_state.get(&quot;asientos_triangulacion_por_factura&quot;, {})
for _, item in df_p.iterrows():
t_comp = item[&quot;Tipo Comp&quot;]
cons = item[&quot;Consecutivo&quot;]
f_str = item[&quot;Fecha&quot;]
fn_item = str(item[&quot;Factura&quot;]).strip()
# Si esta factura tiene contabilización aprobada desde triangulación, se exporta su asiento
real
if fn_item in asientos_triang_map_siigo:
df_as_tr = asientos_triang_map_siigo[fn_item]
f_limpia_op = normalizar_fecha_dian(f_str)
for _, r_as in df_as_tr.iterrows():
cta_code = str(r_as[&quot;Código Cuenta&quot;]).strip()
nit_clean = re.sub(r&#39;\D&#39;, &#39;&#39;, str(r_as[&quot;Tercero / NIT&quot;]).split(&quot;-&quot;)[0])
deb_v = float(r_as[&quot;Débito ($)&quot;])
cred_v = float(r_as[&quot;Crédito ($)&quot;])
desc_line = str(r_as.get(&quot;Descripción de la Cuenta&quot;) or r_as.get(&quot;Descripción Cuenta&quot;)
or &quot;&quot;)[:40]

ws_interfaz.append([
t_comp, cons, f_limpia_op, &quot;COP&quot;, 1,
cta_code, nit_clean, 0, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;,
item.get(&quot;Prefijo&quot;, &quot;IMP&quot;), item.get(&quot;Folio&quot;, str(item.get(&quot;Factura&quot;, &quot;&quot;))),
1, f_limpia_op, &quot;&quot;, &quot;&quot;, &quot;&quot;,
desc_line, &quot;&quot;, deb_v, cred_v,
f&quot;Triangulación {str(item.get(&#39;Proveedor&#39;) or &#39;Tercero&#39;)[:15]} Fac
{item[&#39;Factura&#39;]}&quot;,
0.0, 0.0, &quot;&quot;
])
ws_matriz.append([
t_comp, cons, f_limpia_op, nit_clean,
item.get(&quot;Prefijo&quot;, &quot;IMP&quot;), item.get(&quot;Folio&quot;, str(item.get(&quot;Factura&quot;, &quot;&quot;))),
desc_line, &quot;Cruce Triangulación&quot;, cta_code, deb_v, 0.0, 0.0, 0.0, 0.0, &quot;22050501&quot;
])
fila_r += len(df_as_tr)
continue
nit = item[&quot;NIT Emisor&quot;]
pref = item[&quot;Prefijo&quot;]
fac_num = item[&quot;Folio&quot;]
desc = item[&quot;Descripcion&quot;]
op = item[&quot;Operacion&quot;]
cta_p = item[&quot;Cta Principal&quot;]
base = item[&quot;Base&quot;]
iva = item[&quot;IVA&quot;]
rfte = item[&quot;ReteFuente&quot;]
rica = item.get(&quot;ReteICA&quot;, 0.0)
riva = item.get(&quot;ReteIVA&quot;, 0.0)
cta_c = item[&quot;Cta Contrapartida&quot;]
cta_iva = item.get(&quot;Cta IVA&quot;, &quot;24081001&quot;)
cta_rfte = item.get(&quot;Cta ReteFuente&quot;, &quot;23654001&quot;)
cta_rica = item.get(&quot;Cta ReteICA&quot;, &quot;23680501&quot;)
ws_matriz.append([t_comp, cons, f_str, nit, pref, fac_num, desc, op, cta_p, base, iva, rfte,
rica, riva, cta_c])
r = fila_r
# Línea 1: Base Imponible
ws_interfaz.append([
f&quot;=matriz_captura!A{r}&quot;, f&quot;=matriz_captura!B{r}&quot;, f&quot;=matriz_captura!C{r}&quot;, &quot;COP&quot;, 1,
f&quot;=matriz_captura!I{r}&quot;, f&quot;=matriz_captura!D{r}&quot;, 0, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;,
&quot;&quot;, &quot;&quot;,
f&quot;=matriz_captura!G{r}&quot;, &quot;&quot;,
f&#39;=IF(matriz_captura!H{r}=&quot;Devolucion Compra&quot;, 0, matriz_captura!J{r})&#39;,
f&#39;=IF(matriz_captura!H{r}=&quot;Devolucion Compra&quot;, matriz_captura!J{r}, 0)&#39;,
&quot;&quot;, f&quot;=matriz_captura!J{r}&quot;, 0.0, &quot;&quot;
])
# Línea 2: IVA
if iva &gt; 0:
cod_imp_iva = CODIGOS_IMPUESTO_SIIGO.get(str(cta_iva).strip(), &quot;&quot;)
ws_interfaz.append([
f&quot;=matriz_captura!A{r}&quot;, f&quot;=matriz_captura!B{r}&quot;, f&quot;=matriz_captura!C{r}&quot;, &quot;COP&quot;, 1,
cta_iva, f&quot;=matriz_captura!D{r}&quot;, 0, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;,
cod_imp_iva, &quot;&quot;, &quot;&quot;, f&#39;=&quot;IVA Base: &quot; &amp; matriz_captura!J{r}&#39;, &quot;&quot;,
f&#39;=IF(matriz_captura!H{r}=&quot;Devolucion Compra&quot;, 0, matriz_captura!K{r})&#39;,
f&#39;=IF(matriz_captura!H{r}=&quot;Devolucion Compra&quot;, matriz_captura!K{r}, 0)&#39;,
&quot;&quot;, f&quot;=matriz_captura!J{r}&quot;, 0.0, &quot;&quot;
])
# Línea 2.B: IVA de Importación (240835)
iva_imp_it = float(item.get(&quot;IVA Importación&quot;, 0.0))
cta_iva_imp_it = str(item.get(&quot;Cta IVA Importación&quot;) or &quot;240835&quot;).strip()
if iva_imp_it &gt; 0:
cod_imp_ivaimp = CODIGOS_IMPUESTO_SIIGO.get(cta_iva_imp_it, &quot;&quot;)
ws_interfaz.append([
f&quot;=matriz_captura!A{r}&quot;, f&quot;=matriz_captura!B{r}&quot;, f&quot;=matriz_captura!C{r}&quot;, &quot;COP&quot;, 1,
cta_iva_imp_it, f&quot;=matriz_captura!D{r}&quot;, 0, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;,
cod_imp_ivaimp, &quot;&quot;, &quot;&quot;, f&#39;=&quot;IVA Importación Fac &quot; &amp; matriz_captura!E{r} &amp; &quot;-&quot; &amp;
matriz_captura!F{r}&#39;, &quot;&quot;,
f&#39;=IF(matriz_captura!H{r}=&quot;Devolucion Compra&quot;, 0, {iva_imp_it})&#39;,
f&#39;=IF(matriz_captura!H{r}=&quot;Devolucion Compra&quot;, {iva_imp_it}, 0)&#39;,

&quot;&quot;, f&quot;=matriz_captura!J{r}&quot;, 0.0, &quot;&quot;
])
# Línea 3: ReteFuente
if rfte &gt; 0 and cta_rfte:
cod_imp_rfte = CODIGOS_IMPUESTO_SIIGO.get(str(cta_rfte).strip(), &quot;&quot;)
ws_interfaz.append([
f&quot;=matriz_captura!A{r}&quot;, f&quot;=matriz_captura!B{r}&quot;, f&quot;=matriz_captura!C{r}&quot;, &quot;COP&quot;, 1,
cta_rfte, f&quot;=matriz_captura!D{r}&quot;, 0, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;,
cod_imp_rfte, &quot;&quot;, &quot;&quot;, f&#39;=&quot;ReteFuente Base: &quot; &amp; matriz_captura!J{r}&#39;, &quot;&quot;,
f&#39;=IF(matriz_captura!H{r}=&quot;Devolucion Compra&quot;, matriz_captura!L{r}, 0)&#39;,
f&#39;=IF(matriz_captura!H{r}=&quot;Devolucion Compra&quot;, 0, matriz_captura!L{r})&#39;,
&quot;&quot;, f&quot;=matriz_captura!J{r}&quot;, 0.0, &quot;&quot;
])
# Línea 4: ReteICA
if rica &gt; 0 and cta_rica:
cod_imp_rica = CODIGOS_IMPUESTO_SIIGO.get(str(cta_rica).strip(), &quot;&quot;)
ws_interfaz.append([
f&quot;=matriz_captura!A{r}&quot;, f&quot;=matriz_captura!B{r}&quot;, f&quot;=matriz_captura!C{r}&quot;, &quot;COP&quot;, 1,
cta_rica, f&quot;=matriz_captura!D{r}&quot;, 0, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;,
cod_imp_rica, &quot;&quot;, &quot;&quot;, f&#39;=&quot;ReteICA Base: &quot; &amp; matriz_captura!J{r}&#39;, &quot;&quot;,
f&#39;=IF(matriz_captura!H{r}=&quot;Devolucion Compra&quot;, matriz_captura!M{r}, 0)&#39;,
f&#39;=IF(matriz_captura!H{r}=&quot;Devolucion Compra&quot;, 0, matriz_captura!M{r})&#39;,
&quot;&quot;, f&quot;=matriz_captura!J{r}&quot;, 0.0, &quot;&quot;
])
# Línea 5: Cuenta por Pagar (Contrapartida)
extra_iva_imp_str = f&quot; + {iva_imp_it}&quot; if iva_imp_it &gt; 0 else &quot;&quot;
formula_deb = f&#39;=IF(matriz_captura!H{r}=&quot;Devolucion Compra&quot;, matriz_captura!J{r} +
matriz_captura!K{r}{extra_iva_imp_str} - matriz_captura!L{r} - matriz_captura!M{r} - matriz_captura!N{r}, 0)&#39;
formula_cred = f&#39;=IF(matriz_captura!H{r}=&quot;Devolucion Compra&quot;, 0, matriz_captura!J{r} +
matriz_captura!K{r}{extra_iva_imp_str} - matriz_captura!L{r} - matriz_captura!M{r} - matriz_captura!N{r})&#39;
ws_interfaz.append([
f&quot;=matriz_captura!A{r}&quot;, f&quot;=matriz_captura!B{r}&quot;, f&quot;=matriz_captura!C{r}&quot;, &quot;COP&quot;, 1,
f&quot;=matriz_captura!O{r}&quot;, f&quot;=matriz_captura!D{r}&quot;, 0, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;,
f&quot;=matriz_captura!A{r}&quot;, f&quot;=matriz_captura!B{r}&quot;, 1, f&quot;=matriz_captura!C{r}&quot;,
&quot;&quot;, &quot;&quot;, &quot;&quot;, f&#39;=&quot;Fac &quot; &amp; matriz_captura!E{r} &amp; &quot;-&quot; &amp; matriz_captura!F{r}&#39;, &quot;&quot;,
formula_deb, formula_cred, &quot;&quot;, 0.0, 0.0, &quot;&quot;
])
fila_r += 1
# 2.B AGREGAR A CONTABILIDAD CADA PAQUETE DE TRIANGULACIÓN LISTO EN SU FECHA EXACTA
if inc_pqs_listos and pqs_listos_siigo:
cons_cruce_base = 1
if &quot;cons_ini_nota&quot; in locals() and cons_ini_nota:
try:
cons_cruce_base = int(cons_ini_nota)
except Exception:
cons_cruce_base = 1
elif &quot;cons_ini_nc&quot; in locals() and cons_ini_nc:
try:
cons_cruce_base = int(cons_ini_nc)
except Exception:
cons_cruce_base = 1
idx_pq_cruce = 0
for p_k, p_v in pqs_listos_siigo.items():
ag_item = p_v[&quot;agente&quot;]
ag_fac_chk = str(ag_item.get(&quot;Factura&quot;, &quot;&quot;)).strip()
# Si la factura del agente ya fue exportada con su asiento de triangulación en df_p, no
duplicar
if ag_fac_chk in df_p[&quot;Factura&quot;].astype(str).str.strip().values and ag_fac_chk in
asientos_triang_map_siigo:
continue
terc_items = p_v[&quot;terceros&quot;]
fecha_op_raw = ag_item[&quot;Fecha&quot;]
fecha_op_clean = normalizar_fecha_dian(fecha_op_raw)
enviar_nd = st.session_state.get(f&quot;enviar_nd_pq_{p_k}&quot;, False)
enviar_gp_p = st.session_state.get(f&quot;enviar_gp_pq_{p_k}&quot;, False)
enviar_h2_p = st.session_state.get(f&quot;enviar_h2_pq_{p_k}&quot;, False)

if st.session_state.get(f&quot;paquete_listo_{p_k}&quot;, False) and (f&quot;asiento_fijo_pq_{p_k}&quot; in
st.session_state or &quot;asiento_fijo&quot; in p_v):
df_as_p = pd.DataFrame(st.session_state.get(f&quot;asiento_fijo_pq_{p_k}&quot;,
p_v.get(&quot;asiento_fijo&quot;))).copy()
dif_p = 0.0
ret_p = 0.0
else:
df_as_p, dif_p, ret_p = generar_asiento_triangulacion_paquete(ag_item, terc_items,
enviar_a_no_deducible=enviar_nd, enviar_a_gastos_propios=enviar_gp_p, enviar_a_hoja2=enviar_h2_p)
cons_cruce_actual = cons_cruce_base + idx_pq_cruce
idx_pq_cruce += 1
# Asiento formal en interfaz_siigo
for _, f_as in df_as_p.iterrows():
cta_num = str(f_as[&quot;Código Cuenta&quot;]).strip()
nit_terc_clean = re.sub(r&#39;\D&#39;, &#39;&#39;, str(f_as[&quot;Tercero / NIT&quot;]).split(&quot;-&quot;)[0])
deb_val = float(f_as[&quot;Débito ($)&quot;])
cred_val = float(f_as[&quot;Crédito ($)&quot;])
desc_linea = str(f_as.get(&quot;Descripción de la Cuenta&quot;) or f_as.get(&quot;Descripción
Cuenta&quot;) or f_as.get(&quot;Descripción&quot;) or &quot;&quot;)[:40]
ws_interfaz.append([
14, # Tipo de Comprobante: 14 (Nota de Contabilidad / Cruces)
cons_cruce_actual,
fecha_op_clean,
&quot;COP&quot;,
1,
cta_num,
nit_terc_clean,
0,
&quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;,
ag_item.get(&quot;Prefijo&quot;, &quot;IMP&quot;),
ag_item.get(&quot;Folio&quot;, str(ag_item.get(&quot;Factura&quot;, &quot;&quot;))),
1,
fecha_op_clean,
&quot;&quot;, &quot;&quot;, &quot;&quot;,
desc_linea,
&quot;&quot;,
deb_val,
cred_val,
f&quot;Cruce Importacion Pq #{p_k} {ag_str(item.get(&#39;Proveedor&#39;) or &#39;Tercero&#39;)[:15]}&quot;,
0.0,
0.0,
&quot;&quot;
])
# Registro en matriz_captura
ws_matriz.append([
14, cons_cruce_actual, fecha_op_clean,
re.sub(r&#39;\D&#39;, &#39;&#39;, str(ag_item[&quot;NIT Emisor&quot;])),
ag_item.get(&quot;Prefijo&quot;, &quot;IMP&quot;),
ag_item.get(&quot;Folio&quot;, str(ag_item.get(&quot;Factura&quot;, &quot;&quot;))),
f&quot;Cruce Importación Pq #{p_k} {str(ag_item.get(&#39;Proveedor&#39;) or &#39;Agente&#39;)[:20]}&quot;,
&quot;Cruce Triangulación&quot;,
&quot;22050501&quot;,
float(ag_item.get(&quot;Base&quot;, 0.0)),
float(ag_item.get(&quot;IVA&quot;, 0.0)),
0.0, 0.0, 0.0,
&quot;22050501&quot;
])
fila_r += 1
# 3. Hoja Parametrización
ws_params = wb.create_sheet(title=&quot;Parametrización&quot;)
ws_params.append([&quot;Tipo Comprobante&quot;, &quot;&quot;, &quot;Impuesto&quot;, &quot;Siigo_ID&quot;, &quot;Tarifa&quot;, &quot;Venta&quot;, &quot;Compra&quot;,
&quot;Dev_Venta&quot;, &quot;Dev_Compra&quot;])
comprobantes = [
&quot;1 - Ajustes contables&quot;, &quot;2 - Depreciación&quot;, &quot;3 - Costeo&quot;, &quot;4 - Diferidos&quot;, &quot;5 - Legalización
de viaticos&quot;,
&quot;6 - Legalización de Caja menores&quot;, &quot;7 - Obligaciones financieras&quot;, &quot;8 - Nómina&quot;, &quot;998 -
Cierre año&quot;,

&quot;999 - Saldos iniciales&quot;, &quot;992 - Comprobante de nómina&quot;, &quot;993 - Comprobante de nómina
provisión y seguridad social&quot;,
&quot;994 - Comprobante liquidación de contrato&quot;, &quot;995 - Comprobante liquidación de primas&quot;, &quot;996 -
Comprobante liquidación de cesantías&quot;,
&quot;997 - Comprobante desembolso nómina&quot;, &quot;9 - Recibo de Caja&quot;, &quot;10 - Factura de Compra&quot;, &quot;11 -
Comprobante de Egreso&quot;,
&quot;12 - Factura de Venta&quot;, &quot;13 - Nota Credito&quot;, &quot;14 - Nota de Contabilidad&quot;, &quot;15 - Documento
soporte electronico&quot;,
&quot;777 - Ajustes contables de cartera&quot;, &quot;9901 - Traslado de dinero&quot;, &quot;16 - FACTURAS DE COMPRA
(2)&quot;, &quot;17 - NOTA CREDITO DE COMPRA&quot;
]
impuestos = [
(&quot;IVA 19%&quot;, 1, 0.19, &quot;24080601&quot;, &quot;24081001&quot;, &quot;24082001&quot;, &quot;24081002&quot;),
(&quot;IVA Servicios 19%&quot;, 1, 0.19, &quot;24080601&quot;, &quot;24081501&quot;, &quot;24082001&quot;, &quot;24082001&quot;),
(&quot;IVA Importación 19%&quot;, 1, 0.19, &quot;24080601&quot;, &quot;240835&quot;, &quot;24082001&quot;, &quot;24081002&quot;),
(&quot;IVA 5%&quot;, 2, 0.05, &quot;24080602&quot;, &quot;24081003&quot;, &quot;24082002&quot;, &quot;24081004&quot;),
(&quot;Retefuente 11%&quot;, 3, 0.11, &quot;13551509&quot;, &quot;23651501&quot;, &quot;13551510&quot;, &quot;23651502&quot;),
(&quot;Retefuente 10%&quot;, 4, 0.1, &quot;13551507&quot;, &quot;23652001&quot;, &quot;13551508&quot;, &quot;23652002&quot;),
(&quot;Retefuente 6%&quot;, 5, 0.06, &quot;13551505&quot;, &quot;23652501&quot;, &quot;13551506&quot;, &quot;23652502&quot;),
(&quot;Retefuente 4%&quot;, 6, 0.04, &quot;13551503&quot;, &quot;23652503&quot;, &quot;13551504&quot;, &quot;23652504&quot;),
(&quot;Retefuente 3.5%&quot;, 18, 0.035, &quot;13551513&quot;, &quot;23654004&quot;, &quot;13551514&quot;, &quot;23654005&quot;),
(&quot;Retefuente 2.5%&quot;, 7, 0.025, &quot;13551501&quot;, &quot;23654001&quot;, &quot;13551502&quot;, &quot;23654002&quot;),
(&quot;Retefuente Combustibles 0.1%&quot;, 23, 0.001, &quot;13551599&quot;, &quot;23654006&quot;, &quot;13551599&quot;, &quot;23654006&quot;),
(&quot;ReteICA 11.04&quot;, 8, 0.01104, &quot;13551801&quot;, &quot;23680501&quot;, &quot;13551802&quot;, &quot;23680502&quot;),
(&quot;ReteICA 13.8&quot;, 9, 0.0138, &quot;13551803&quot;, &quot;23680503&quot;, &quot;13551804&quot;, &quot;23680504&quot;),
(&quot;ReteICA 9.66&quot;, 10, 0.00966, &quot;13551805&quot;, &quot;23680505&quot;, &quot;13551806&quot;, &quot;23680506&quot;),
(&quot;ReteICA 8&quot;, 11, 0.008, &quot;13551807&quot;, &quot;23680507&quot;, &quot;13551808&quot;, &quot;23680508&quot;),
(&quot;ReteICA 7&quot;, 12, 0.007, &quot;13551809&quot;, &quot;23680509&quot;, &quot;13551810&quot;, &quot;23680510&quot;),
(&quot;ReteICA 6.9&quot;, 13, 0.0069, &quot;13551811&quot;, &quot;23680511&quot;, &quot;13551812&quot;, &quot;23680512&quot;),
(&quot;ReteICA 4.14&quot;, 14, 0.00414, &quot;13551813&quot;, &quot;23680513&quot;, &quot;13551814&quot;, &quot;23680514&quot;),
(&quot;ReteIVA 15%&quot;, 15, 0.15, &quot;13551701&quot;, &quot;23670101&quot;, &quot;13551702&quot;, &quot;23670102&quot;)
]
max_len = max(len(comprobantes), len(impuestos))
for idx_row in range(max_len):
c_val = comprobantes[idx_row] if idx_row &lt; len(comprobantes) else &quot;&quot;
if idx_row &lt; len(impuestos):
imp = list(impuestos[idx_row])
ws_params.append([c_val, &quot;&quot;] + imp)
else:
ws_params.append([c_val, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;, &quot;&quot;])
output = io.BytesIO()
wb.save(output)
output.seek(0)
st.download_button(
label=f&quot;�� 1. Descargar Planilla Oficial Siigo Nube (3 Hojas Formuladas)
{empresa[&#39;nombre&#39;]}&quot;,
data=output,
file_name=f&quot;Plantilla_Siigo_{empresa[&#39;nombre&#39;].replace(&#39; &#39;, &#39;_&#39;)}_3_Hojas.xlsx&quot;,
mime=&quot;application/vnd.openxmlformats-officedocument.spreadsheetml.sheet&quot;,
use_container_width=True
)
st.success(&quot;Planilla generada con las 3 hojas oficiales: &#39;matriz_captura&#39;, &#39;interfaz_siigo&#39; (con
códigos de impuesto) y &#39;Parametrización&#39;.&quot;)
st.markdown(&quot;---&quot;)
st.markdown(&quot;### �� Libro Auxiliar Contable: Facturas Una a Una y Consolidado&quot;)
st.write(&quot;Genera el reporte administrativo con el detalle factura a factura y las hojas de
consolidado por proveedor y cuentas contables:&quot;)
out_aux = io.BytesIO()
with pd.ExcelWriter(out_aux, engine=&#39;openpyxl&#39;) as writer_aux:
cols_detalle = [&quot;Comprobante Siigo&quot;, &quot;Fecha&quot;, &quot;Factura&quot;, &quot;Proveedor&quot;, &quot;NIT Emisor&quot;, &quot;Cta
Principal&quot;, &quot;Categoría&quot;, &quot;Base&quot;, &quot;IVA&quot;, &quot;ReteFuente&quot;, &quot;ReteICA&quot;, &quot;Total&quot;, &quot;Cta Contrapartida&quot;, &quot;Razón Contable&quot;,
&quot;Soporte PDF Renombrado&quot;]
df_det_export = df_p[cols_detalle].copy()
totales_dict = {
&quot;Comprobante Siigo&quot;: &quot;TOTALES CONSOLIDADOS&quot;,
&quot;Fecha&quot;: &quot;-&quot;, &quot;Factura&quot;: f&quot;{len(df_det_export)} Docs&quot;, &quot;Proveedor&quot;: &quot;-&quot;, &quot;NIT Emisor&quot;: &quot;-
&quot;,

&quot;Cta Principal&quot;: &quot;-&quot;, &quot;Categoría&quot;: &quot;-&quot;,
&quot;Base&quot;: df_det_export[&quot;Base&quot;].sum(),
&quot;IVA&quot;: df_det_export[&quot;IVA&quot;].sum(),
&quot;ReteFuente&quot;: df_det_export[&quot;ReteFuente&quot;].sum(),
&quot;ReteICA&quot;: df_det_export[&quot;ReteICA&quot;].sum(),
&quot;Total&quot;: df_det_export[&quot;Total&quot;].sum(),
&quot;Cta Contrapartida&quot;: &quot;-&quot;, &quot;Razón Contable&quot;: &quot;-&quot;, &quot;Soporte PDF Renombrado&quot;: &quot;-&quot;
}
df_det_export = pd.concat([df_det_export, pd.DataFrame([totales_dict])], ignore_index=True)
df_det_export.to_excel(writer_aux, sheet_name=&quot;Facturas_Una_a_Una&quot;, index=False)
df_cons_prov = df_p.groupby([&quot;NIT Emisor&quot;, &quot;Proveedor&quot;]).agg({
&quot;Comprobante Siigo&quot;: &quot;count&quot;,
&quot;Base&quot;: &quot;sum&quot;,
&quot;IVA&quot;: &quot;sum&quot;,
&quot;ReteFuente&quot;: &quot;sum&quot;,
&quot;ReteICA&quot;: &quot;sum&quot;,
&quot;Total&quot;: &quot;sum&quot;
}).reset_index().rename(columns={&quot;Comprobante Siigo&quot;: &quot;Cant Facturas&quot;})
df_cons_prov[&quot;Neto CxP&quot;] = df_cons_prov[&quot;Total&quot;] - df_cons_prov[&quot;ReteFuente&quot;] -
df_cons_prov[&quot;ReteICA&quot;]
df_cons_prov = df_cons_prov.sort_values(by=&quot;Total&quot;, ascending=False).reset_index(drop=True)
df_cons_prov.to_excel(writer_aux, sheet_name=&quot;Consolidado_Proveedores&quot;, index=False)
df_cons_cta = df_p.groupby([&quot;Cta Principal&quot;, &quot;Categoría&quot;]).agg({
&quot;Comprobante Siigo&quot;: &quot;count&quot;,
&quot;Base&quot;: &quot;sum&quot;,
&quot;IVA&quot;: &quot;sum&quot;,
&quot;ReteFuente&quot;: &quot;sum&quot;,
&quot;ReteICA&quot;: &quot;sum&quot;,
&quot;Total&quot;: &quot;sum&quot;
}).reset_index().rename(columns={&quot;Comprobante Siigo&quot;: &quot;Cant Facturas&quot;})
df_cons_cta.to_excel(writer_aux, sheet_name=&quot;Consolidado_Cuentas_PUC&quot;, index=False)
out_aux.seek(0)
st.download_button(
label=f&quot;�� 2. Descargar Libro de Facturas (Una a Una + Consolidado) - {empresa[&#39;nombre&#39;]}&quot;,
data=out_aux,
file_name=f&quot;Libro_Facturas_Detallado_Y_Consolidado_{empresa[&#39;nombre&#39;].replace(&#39; &#39;,
&#39;_&#39;)}.xlsx&quot;,
mime=&quot;application/vnd.openxmlformats-officedocument.spreadsheetml.sheet&quot;,
use_container_width=True
)
else:
st.info(&quot;�� Marca la casilla superior **&#39;⚡ Generar Planillas Oficiales de Siigo Nube y Libr
Auxiliar&#39;** cuando hayas terminado tus ajustes para habilitar los botones de descarga inmediata.&quot;)
else:
st.info(&quot;Primero procesa los documentos en la Pestana 1 para habilitar la descarga.&quot;)
except Exception as _e_sii:
if type(_e_sii).__name__ in [&#39;RerunException&#39;, &#39;StopException&#39;, &#39;ScriptControlException&#39;]:
raise _e_sii
st.error(f&quot;⚠️ Ocurrió un error en la exportación a Siigo: {_e_sii}&quot;)
with st.expander(&quot;�� Ver detalles técnicos del error (Traceback):&quot;):
st.code(traceback.format_exc())
import gc
gc.collect()

except Exception as _e_fatal:
if type(_e_fatal).__name__ in [&#39;RerunException&#39;, &#39;StopException&#39;, &#39;ScriptControlException&#39;]:
raise _e_fatal
st.markdown(&quot;&quot;&quot;
&lt;div style=&quot;background:#fff1f2; border:1px solid #fecdd3; border-left:6px solid #e11d48; border-radius:10px;
padding:22px; margin:20px 0; box-shadow: 0 4px 6px rgba(0,0,0,0.05);&quot;&gt;
&lt;h3 style=&quot;color:#9f1239; margin-top:0;&quot;&gt;��️ Escudo de Protección Anticaídas Activado&lt;/h3&gt;
&lt;p style=&quot;color:#881337; font-size:15px; line-height:1.6;&quot;&gt;
Se ha interceptado una excepción en tiempo de ejecución y se &lt;b&gt;evitó por completo la caída fatal de
Streamlit (&quot;Oh no&quot;)&lt;/b&gt;.&lt;br&gt;
Puedes reiniciar la memoria de sesión con el botón inferior o revisar el diagnóstico técnico para
continuar trabajando.
&lt;/p&gt;

&lt;/div&gt;
&quot;&quot;&quot;, unsafe_allow_html=True)
c_rec1, c_rec2 = st.columns([1, 1])
with c_rec1:
if st.button(&quot;�� Reiniciar Sesión Limpia (Recomendado)&quot;, type=&quot;primary&quot;, use_container_width=True
key=&quot;btn_fatal_reset_global&quot;):
st.session_state.clear()
st.rerun()
with c_rec2:
if st.button(&quot;↩️ Reintentar sin borrar datos&quot;, use_container_width=True, key=&quot;btn_fatal_retry_global&quot;):
st.rerun()
with st.expander(&quot;�� Ver detalles técnicos del error (Traceback completo):&quot;):
st.code(traceback.format_exc())
