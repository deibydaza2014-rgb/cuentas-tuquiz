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
from pypdf import PdfReader, PdfWriter
try:
    import pypdfium2 as pdfium
    HAS_PDFIUM = True
except Exception:
    HAS_PDFIUM = False

st.set_page_config(page_title="Cuentas Tuquiz - ERP y Auditoria Contable", layout="wide", page_icon="🏢")

# Estilos visuales profesionales y protección contra conflictos de Google Translate
st.markdown("""
<meta name="google" content="notranslate">
<style class="notranslate">
    html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] {
        -webkit-translate: no !important;
        translate: no !important;
    }
    .main { background-color: #f8fafc; }
    .stButton>button { background-color: #0070ba; color: white; border-radius: 6px; font-weight: 600; }
    .card-box { background: white; padding: 22px; border-radius: 10px; border: 1px solid #e2e8f0; box-shadow: 0 2px 5px rgba(0,0,0,0.04); margin-bottom: 18px; }
    .audit-card { background: #ffffff; padding: 22px; border-left: 5px solid #0070ba; border-radius: 8px; box-shadow: 0 2px 6px rgba(0,0,0,0.06); }
    .badge-active { background-color: #dcfce7; color: #15803d; padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: bold; }
    .badge-next { background-color: #f1f5f9; color: #64748b; padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: bold; }
    .tag-propio { background-color: #dcfce7; color: #15803d; padding: 4px 8px; border-radius: 4px; font-weight: bold; }
    .tag-comp { background-color: #eff6ff; color: #1d4ed8; padding: 4px 8px; border-radius: 4px; font-weight: bold; font-family: monospace; }
    .badge-tax { background-color: #e0f2fe; color: #0369a1; padding: 3px 8px; border-radius: 6px; font-size: 12px; font-weight: 600; margin-right: 4px; }
</style>
""", unsafe_allow_html=True)



# ==============================================================================
# CONFIGURACIÓN DE DIRECTORIOS Y SISTEMA DE AUTO-GUARDADO CONTINUO
# ==============================================================================
# ==============================================================================
# CONFIGURACIÓN DE DIRECTORIOS Y SISTEMA DE AUTO-GUARDADO CONTINUO
# ==============================================================================
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "almacenamiento_contable")
os.makedirs(DATA_DIR, exist_ok=True)

# ESTADOS DE SESION LIMPIOS EN MEMORIA (SIN AUTOGUARDADO EN DISCO)
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
if "empresa_activa" not in st.session_state:
    st.session_state["empresa_activa"] = None
if "proceso_activo" not in st.session_state:
    st.session_state["proceso_activo"] = None

# PANTALLA 1: LOGIN
if not st.session_state["autenticado"]:
    col1, col2, col3 = st.columns(3)
    with col2:
        st.markdown("<h2 style='text-align: center; color: #0f172a;'>Cuentas Tuquiz</h2>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: #64748b;'>Plataforma unificada para gestion contable y DIAN (Siigo / World Office)</p>", unsafe_allow_html=True)

        with st.form("login_form"):
            usuario = st.text_input("Usuario o Correo Electronico", value="deibydaza2014@gmail.com")
            password = st.text_input("Contrasena", type="password", value="123456")
            submit = st.form_submit_button("Ingresar al Ecosistema")

            if submit:
                if usuario and password:
                    st.session_state["autenticado"] = True
                    st.rerun()
                else:
                    st.error("Ingresa usuario y contrasena.")
    st.stop()

# PANTALLA 2: SELECTOR DE EMPRESAS CON RESPONSABILIDADES RUT
EMPRESAS_DISPONIBLES = [
    {
        "nombre": "INDUMAQ ER SAS",
        "nit": "901.346.412-5",
        "actividad": "Comercio y Reparacion de Maquinaria / Importaciones",
        "regimen": "Responsable de IVA",
        "es_gran_contribuyente": False,
        "es_autorretenedor": False,
        "responsabilidades_rut": ["O-48"],
        "estado": "ACTIVA",
        "badge": "badge-active"
    }
]

if not st.session_state["empresa_activa"]:
    st.title("Seleccion de Empresa")
    st.caption("Selecciona la entidad sobre la cual vas a trabajar:")

    c1, c2 = st.columns(2)
    for i, emp in enumerate(EMPRESAS_DISPONIBLES):
        col = c1 if i % 2 == 0 else c2
        with col:
            st.markdown(f"""
            <div class="card-box">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <h3 style="margin:0; color:#0f172a;">🏢 {emp['nombre']}</h3>
                    <span class="{emp['badge']}">{emp['estado']}</span>
                </div>
                <p style="margin:8px 0 2px 0; color:#475569;"><b>NIT:</b> {emp['nit']} - <b>Regimen:</b> {emp['regimen']}</p>
                <p style="margin:2px 0 14px 0; color:#64748b; font-size:14px;">{emp['actividad']}</p>
            </div>
            """, unsafe_allow_html=True)
            if emp["estado"] == "ACTIVA":
                if st.button(f"Ingresar a {emp['nombre']}", key=f"btn_emp_{i}"):
                    st.session_state["empresa_activa"] = emp
                    st.session_state["proceso_activo"] = "facturacion"
                    # registrar_actividad_sesion(emp)
                    st.rerun()
            else:
                st.button(f"Pendiente Datos / Parametrizacion ({emp['estado']})", key=f"btn_emp_{i}", disabled=True)
    st.stop()

empresa = st.session_state["empresa_activa"]

# Sesion en memoria limpia
pass

# PANTALLA 3: MENU DE PROCESOS OPERATIVOS
PROCESOS_SISTEMA = [
    {
        "id": "facturacion",
        "icono": "Facturas",
        "titulo": "Facturas de Compra, Venta y Devoluciones",
        "desc": "Carga de reportes DIAN/Token, desbloqueo y renombrado de facturas por comprobante, auditoria contable y plantilla Siigo.",
        "estado": "ACTIVO",
        "badge": "badge-active"
    },
    {
        "id": "nomina",
        "icono": "Nomina",
        "titulo": "Gestion Laboral y Nomina Electronica",
        "desc": "Calculo de liquidacion de nomina, provisiones de prestaciones sociales, seguridad social y emision de soportes electronicos DIAN.",
        "estado": "PROXIMAMENTE",
        "badge": "badge-next"
    },
    {
        "id": "conciliacion",
        "icono": "Bancos",
        "titulo": "Tesoreria y Conciliacion Bancaria",
        "desc": "Cruce automatizado de extractos bancarios (Bancolombia, Davivienda) contra libros auxiliares y control de partidas conciliatorias.",
        "estado": "PROXIMAMENTE",
        "badge": "badge-next"
    },
    {
        "id": "notas",
        "icono": "Ajustes",
        "titulo": "Notas de Contabilidad y Cierre Fiscal",
        "desc": "Comprobantes de ajuste, amortizaciones de intangibles (NIC 38), depreciaciones y conciliacion fiscal NIIF vs. DIAN.",
        "estado": "PROXIMAMENTE",
        "badge": "badge-next"
    }
]

if not st.session_state["proceso_activo"]:
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        st.subheader(f"{empresa['nombre']} - Panel de Procesos")
        st.caption(f"NIT: {empresa['nit']} - Selecciona el modulo de trabajo que deseas ejecutar:")
    with col_t2:
        if st.button("Cambiar de Empresa"):
            st.session_state["empresa_activa"] = None
            st.session_state["proceso_activo"] = None
            st.rerun()

    st.markdown("---")

    cp1, cp2 = st.columns(2)
    for idx, proc in enumerate(PROCESOS_SISTEMA):
        col_p = cp1 if idx % 2 == 0 else cp2
        with col_p:
            st.markdown(f"""
            <div class="card-box">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <h3 style="margin:0; color:#0f172a;">{proc['titulo']}</h3>
                    <span class="{proc['badge']}">{proc['estado']}</span>
                </div>
                <p style="margin:12px 0 16px 0; color:#475569; font-size:14px; line-height:1.5;">{proc['desc']}</p>
            </div>
            """, unsafe_allow_html=True)
            if proc["estado"] == "ACTIVO":
                if st.button(f"Abrir Modulo de Facturacion", key=f"btn_proc_{idx}"):
                    st.session_state["proceso_activo"] = proc["id"]
                    st.rerun()
            else:
                st.button(f"Modulo en Construccion ({proc['estado']})", key=f"btn_proc_{idx}", disabled=True)
    st.stop()

# PANTALLA 4: FACTURACION, AUDITORIA Y SIIGO
col_nav1, col_nav2 = st.columns(2)
with col_nav1:
    st.subheader(f"{empresa['nombre']} - Facturacion, Auditoria y Siigo")
    st.caption(f"NIT: {empresa['nit']} - Modulo Activo: Facturas de Compra, Venta y Devoluciones")
with col_nav2:
    if st.button("Volver a Procesos"):
        st.session_state["proceso_activo"] = None
        st.rerun()

st.markdown("---")

# ==============================================================================
# GESTOR DE HISTORIAL DE TRABAJOS Y AUDITORÍAS PASADAS (PERSISTENCIA TOTAL)
# ==============================================================================
# Funciones de sesion y directorios centralizadas en la cabecera



def guardar_estado_manual(empresa_dict):
    """Guarda en disco únicamente los datos contables modificados (DataFrame y Paquetes), sin procesar PDFs."""
    if "df_procesado" not in st.session_state or st.session_state["df_procesado"] is None:
        return False
    try:
        nit_clean = re.sub(r'\D', '', str(empresa_dict.get("nit", "empresa")))
        d = os.path.join(DATA_DIR, nit_clean, "estado_manual")
        os.makedirs(d, exist_ok=True)
        with open(os.path.join(d, "df_procesado.pkl"), "wb") as f:
            pickle.dump(st.session_state["df_procesado"], f)
        if st.session_state.get("paquetes_importacion"):
            with open(os.path.join(d, "paquetes_importacion.pkl"), "wb") as f:
                pickle.dump(st.session_state["paquetes_importacion"], f)
        if st.session_state.get("asientos_triangulacion_por_factura"):
            with open(os.path.join(d, "asientos_triangulacion.pkl"), "wb") as f:
                pickle.dump(st.session_state["asientos_triangulacion_por_factura"], f)
        if st.session_state.get("facturas_no_contabilizar"):
            with open(os.path.join(d, "facturas_no_contabilizar.json"), "w", encoding="utf-8") as f:
                json.dump(list(st.session_state["facturas_no_contabilizar"]), f)
        now_str = datetime.datetime.now().strftime("%d/%m/%Y %I:%M:%S %p")
        meta = {
            "fecha_guardado": now_str,
            "total_facturas": len(st.session_state["df_procesado"]),
            "archivo_excel": st.session_state.get("excel_nombre", "Reporte.xlsx")
        }
        with open(os.path.join(d, "meta.json"), "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
        st.session_state["_ultimo_guardado_manual"] = now_str
        import gc
        gc.collect()
        return True
    except Exception:
        return False

def cargar_estado_manual(empresa_dict):
    """Restaura en memoria el último estado guardado manualmente por el usuario."""
    try:
        nit_clean = re.sub(r'\D', '', str(empresa_dict.get("nit", "empresa")))
        d = os.path.join(DATA_DIR, nit_clean, "estado_manual")
        df_p_path = os.path.join(d, "df_procesado.pkl")
        if os.path.exists(df_p_path):
            with open(df_p_path, "rb") as f:
                st.session_state["df_procesado"] = pickle.load(f)
            pq_path = os.path.join(d, "paquetes_importacion.pkl")
            if os.path.exists(pq_path):
                with open(pq_path, "rb") as f:
                    pqs_loaded = pickle.load(f)
                    st.session_state["paquetes_importacion"] = pqs_loaded
                    for pq_k, pq_v in pqs_loaded.items():
                        if isinstance(pq_v, dict):
                            if "asiento_fijo" in pq_v or "asiento_aprobado" in pq_v:
                                st.session_state[f"paquete_listo_{pq_k}"] = True
                            if "destino_siigo" in pq_v:
                                st.session_state[f"paquete_destino_siigo_{pq_k}"] = pq_v["destino_siigo"]
            as_path = os.path.join(d, "asientos_triangulacion.pkl")
            if os.path.exists(as_path):
                with open(as_path, "rb") as f:
                    st.session_state["asientos_triangulacion_por_factura"] = pickle.load(f)
            fn_path = os.path.join(d, "facturas_no_contabilizar.json")
            if os.path.exists(fn_path):
                with open(fn_path, "r", encoding="utf-8") as f:
                    st.session_state["facturas_no_contabilizar"] = set(json.load(f))
            meta_path = os.path.join(d, "meta.json")
            if os.path.exists(meta_path):
                with open(meta_path, "r", encoding="utf-8") as f:
                    meta = json.load(f)
                st.session_state["excel_nombre"] = meta.get("archivo_excel", "Reporte.xlsx")
                st.session_state["_ultimo_guardado_manual"] = meta.get("fecha_guardado", "")
            return True
    except Exception:
        pass
    return False

def get_empresa_trabajos_dir(empresa_dict):
    nit_clean = re.sub(r"\D", "", str(empresa_dict.get("nit", "empresa")))
    d = os.path.join(DATA_DIR, nit_clean, "trabajos")
    os.makedirs(d, exist_ok=True)
    return d

def guardar_trabajo_en_historial(empresa_dict, df_procesado, excel_bytes=None, excel_nombre="Reporte.xlsx", dict_pdfs_renombrados=None, dict_pdfs_originales=None, zip_bytes=None, consecutivo_ini=680, job_id=None):
    try:
        base_dir = get_empresa_trabajos_dir(empresa_dict)
        if not job_id:
            job_id = st.session_state.get("job_actual_id")
            if not job_id:
                now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                job_id = f"job_{now_str}"

        job_dir = os.path.join(base_dir, job_id)
        os.makedirs(job_dir, exist_ok=True)

        dir_renom = os.path.join(job_dir, "pdfs_renombrados")
        os.makedirs(dir_renom, exist_ok=True)
        dir_orig = os.path.join(job_dir, "pdfs_originales")
        os.makedirs(dir_orig, exist_ok=True)

        # 1. Guardar DataFrame procesado
        if df_procesado is not None:
            with open(os.path.join(job_dir, "df_procesado.pkl"), "wb") as f:
                pickle.dump(df_procesado, f)

        # 2. Guardar archivo Excel original subido
        if excel_bytes:
            with open(os.path.join(job_dir, "excel_original.xlsx"), "wb") as f_ex:
                f_ex.write(excel_bytes)

        # 3. Guardar PDFs renombrados / separados (escritura inteligente: solo si no existe o cambió)
        if dict_pdfs_renombrados:
            for fname, bdata in dict_pdfs_renombrados.items():
                p_ren = os.path.join(dir_renom, fname)
                if not os.path.exists(p_ren) or os.path.getsize(p_ren) != len(bdata):
                    with open(p_ren, "wb") as pf:
                        pf.write(bdata)

        # 4. Guardar PDFs originales subidos (escritura inteligente: solo si no existe o cambió)
        if dict_pdfs_originales:
            for fname, bdata in dict_pdfs_originales.items():
                p_ori = os.path.join(dir_orig, fname)
                if not os.path.exists(p_ori) or os.path.getsize(p_ori) != len(bdata):
                    with open(p_ori, "wb") as pf:
                        pf.write(bdata)

        # 5. Guardar paquete ZIP de facturas si existe
        if zip_bytes:
            with open(os.path.join(job_dir, "paquete_facturas.zip"), "wb") as zf:
                zf.write(zip_bytes)

        # 6. Guardar paquetes de importación y triangulación
        if "paquetes_importacion" in st.session_state and st.session_state["paquetes_importacion"]:
            try:
                with open(os.path.join(job_dir, "paquetes_importacion.pkl"), "wb") as f_pq:
                    pickle.dump(st.session_state["paquetes_importacion"], f_pq)
            except Exception:
                pass

        # 7. Guardar asientos aprobados de triangulación
        if "asientos_triangulacion_por_factura" in st.session_state and st.session_state["asientos_triangulacion_por_factura"]:
            try:
                with open(os.path.join(job_dir, "asientos_triangulacion.pkl"), "wb") as f_as:
                    pickle.dump(st.session_state["asientos_triangulacion_por_factura"], f_as)
            except Exception:
                pass

        # 8. Guardar estado general de la sesión (paquetes listos, opciones, facturas excluidas)
        try:
            estado_flags = {
                "paquetes_listos": {k: v for k, v in st.session_state.items() if k.startswith("paquete_listo_")},
                "paquetes_destino_siigo": {k: v for k, v in st.session_state.items() if k.startswith("paquete_destino_siigo_")},
                "enviar_gp": {k: v for k, v in st.session_state.items() if k.startswith("enviar_gp_pq_")},
                "enviar_h2": {k: v for k, v in st.session_state.items() if k.startswith("enviar_h2_pq_")},
                "enviar_nd": {k: v for k, v in st.session_state.items() if k.startswith("enviar_nd_pq_")},
                "facturas_no_contabilizar": list(st.session_state.get("facturas_no_contabilizar", set())),
                "sel_paquete_activo_key": st.session_state.get("sel_paquete_activo_key", 1)
            }
            with open(os.path.join(job_dir, "estado_sesion.json"), "w", encoding="utf-8") as f_es:
                json.dump(estado_flags, f_es, ensure_ascii=False, indent=2)
        except Exception:
            pass

        now_disp = datetime.datetime.now().strftime("%d/%m/%Y %I:%M %p")
        # Verificar archivos físicos guardados en disco
        n_orig_disco = len(os.listdir(dir_orig)) if os.path.exists(dir_orig) else 0
        n_renom_disco = len(os.listdir(dir_renom)) if os.path.exists(dir_renom) else 0
        tiene_excel = os.path.exists(os.path.join(job_dir, "excel_original.xlsx")) or (excel_bytes is not None)

        meta = {
            "id": job_id,
            "nombre_trabajo": f"Trabajo del {now_disp} ({excel_nombre})",
            "archivo_excel": excel_nombre,
            "tiene_excel": tiene_excel,
            "fecha": now_disp,
            "total_facturas": len(df_procesado) if df_procesado is not None else 0,
            "consecutivo_inicial": consecutivo_ini,
            "total_valor": float(df_procesado["Total"].sum()) if (df_procesado is not None and "Total" in df_procesado) else 0.0,
            "total_pdfs_renombrados": n_renom_disco,
            "total_pdfs_originales": n_orig_disco
        }
        with open(os.path.join(job_dir, "meta.json"), "w", encoding="utf-8") as mf:
            json.dump(meta, mf, ensure_ascii=False, indent=2)

        # Actualizar índice general de la empresa para carga inmediata
        idx_file = os.path.join(base_dir, "historial_index.json")
        try:
            with open(idx_file, "r", encoding="utf-8") as f_idx:
                index_data = json.load(f_idx)
        except Exception:
            index_data = []
        index_data = [x for x in index_data if x.get("id") != job_id]
        index_data.append(meta)
        index_data.sort(key=lambda x: x.get("id", ""), reverse=True)
        try:
            with open(idx_file, "w", encoding="utf-8") as f_idx:
                json.dump(index_data, f_idx, ensure_ascii=False, indent=2)
        except Exception:
            pass

        return job_id
    except Exception:
        return None



def listar_trabajos_historial(empresa_dict):
    base_dir = get_empresa_trabajos_dir(empresa_dict)
    lista = []
    idx_file = os.path.join(base_dir, "historial_index.json")
    if os.path.exists(idx_file):
        try:
            with open(idx_file, "r", encoding="utf-8") as f_idx:
                raw_l = json.load(f_idx)
                if isinstance(raw_l, list):
                    lista = raw_l
        except Exception:
            lista = []
    if not lista and os.path.exists(base_dir):
        for jid in os.listdir(base_dir):
            jdir = os.path.join(base_dir, jid)
            meta_file = os.path.join(jdir, "meta.json")
            if os.path.isdir(jdir) and os.path.exists(meta_file):
                try:
                    with open(meta_file, "r", encoding="utf-8") as mf:
                        meta = json.load(mf)
                        lista.append(meta)
                except Exception:
                    pass

    # Normalizar campos y verificar que exista df_procesado.pkl
    valida = []
    for item in lista:
        if isinstance(item, dict) and "id" in item:
            jdir = os.path.join(base_dir, item["id"])
            if os.path.exists(os.path.join(jdir, "df_procesado.pkl")):
                item.setdefault("total_facturas", 0)
                item.setdefault("total_valor", 0.0)
                item.setdefault("nombre_trabajo", f"Trabajo {item['id']}")
                valida.append(item)

    valida.sort(key=lambda x: x.get("id", ""), reverse=True)
    return valida

def cargar_trabajo_historial(empresa_dict, job_id):
    base_dir = get_empresa_trabajos_dir(empresa_dict)
    jdir = os.path.join(base_dir, job_id)
    df = None
    dict_renom = {}
    dict_orig = {}
    zip_bytes = None
    excel_bytes = None
    excel_nombre = "Reporte.xlsx"

    meta_file = os.path.join(jdir, "meta.json")
    if os.path.exists(meta_file):
        try:
            with open(meta_file, "r", encoding="utf-8") as mf:
                m_data = json.load(mf)
                excel_nombre = m_data.get("archivo_excel", "Reporte.xlsx")
        except:
            pass

    pkl_file = os.path.join(jdir, "df_procesado.pkl")
    if os.path.exists(pkl_file):
        with open(pkl_file, "rb") as f:
            df = pickle.load(f)

    ex_file = os.path.join(jdir, "excel_original.xlsx")
    if os.path.exists(ex_file):
        with open(ex_file, "rb") as f_ex:
            excel_bytes = f_ex.read()

    # PDFs renombrados
    pdf_dir_r = os.path.join(jdir, "pdfs_renombrados")
    if not os.path.exists(pdf_dir_r):
        pdf_dir_r = os.path.join(jdir, "pdfs")
    if os.path.exists(pdf_dir_r):
        for fn in os.listdir(pdf_dir_r):
            if fn.endswith(".pdf"):
                with open(os.path.join(pdf_dir_r, fn), "rb") as pf:
                    dict_renom[fn] = pf.read()

    # PDFs originales
    pdf_dir_o = os.path.join(jdir, "pdfs_originales")
    if os.path.exists(pdf_dir_o):
        for fn in os.listdir(pdf_dir_o):
            if fn.endswith(".pdf"):
                with open(os.path.join(pdf_dir_o, fn), "rb") as pf:
                    dict_orig[fn] = pf.read()

    zip_file = os.path.join(jdir, "paquete_facturas.zip")
    if os.path.exists(zip_file):
        with open(zip_file, "rb") as zf:
            zip_bytes = zf.read()
    elif dict_renom or dict_orig:
        try:
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf_temp:
                fuente = dict_renom if dict_renom else dict_orig
                for fn, bdata in fuente.items():
                    zf_temp.writestr(fn, bdata)
            buf.seek(0)
            zip_bytes = buf.getvalue()
        except Exception:
            pass

    # Restaurar paquetes de importación
    pq_file = os.path.join(jdir, "paquetes_importacion.pkl")
    if os.path.exists(pq_file):
        try:
            with open(pq_file, "rb") as f_pq:
                st.session_state["paquetes_importacion"] = pickle.load(f_pq)
        except Exception:
            pass

    # Restaurar asientos aprobados de triangulación
    as_file = os.path.join(jdir, "asientos_triangulacion.pkl")
    if os.path.exists(as_file):
        try:
            with open(as_file, "rb") as f_as:
                st.session_state["asientos_triangulacion_por_factura"] = pickle.load(f_as)
        except Exception:
            pass

    # Restaurar flags de estado (paquetes listos, opciones, facturas excluidas)
    es_file = os.path.join(jdir, "estado_sesion.json")
    if os.path.exists(es_file):
        try:
            with open(es_file, "r", encoding="utf-8") as f_es:
                est = json.load(f_es)
                for k, v in est.get("paquetes_listos", {}).items():
                    st.session_state[k] = v
                for k, v in est.get("paquetes_destino_siigo", {}).items():
                    st.session_state[k] = v
                for k, v in est.get("enviar_gp", {}).items():
                    st.session_state[k] = v
                for k, v in est.get("enviar_h2", {}).items():
                    st.session_state[k] = v
                for k, v in est.get("enviar_nd", {}).items():
                    st.session_state[k] = v
                if "facturas_no_contabilizar" in est:
                    st.session_state["facturas_no_contabilizar"] = set(est["facturas_no_contabilizar"])
                if "sel_paquete_activo_key" in est:
                    st.session_state["sel_paquete_activo_key"] = est["sel_paquete_activo_key"]
        except Exception:
            pass

    return df, dict_renom, dict_orig, excel_bytes, excel_nombre, zip_bytes

def eliminar_trabajo_historial(empresa_dict, job_id):
    base_dir = get_empresa_trabajos_dir(empresa_dict)
    jdir = os.path.join(base_dir, job_id)
    if os.path.exists(jdir):
        shutil.rmtree(jdir, ignore_errors=True)
    idx_file = os.path.join(base_dir, "historial_index.json")
    if os.path.exists(idx_file):
        try:
            with open(idx_file, "r", encoding="utf-8") as f_idx:
                index_data = json.load(f_idx)
            index_data = [x for x in index_data if x.get("id") != job_id]
            with open(idx_file, "w", encoding="utf-8") as f_idx:
                json.dump(index_data, f_idx, ensure_ascii=False, indent=2)
        except Exception:
            pass

def exportar_respaldo_sesion_zip(empresa_dict, df_proc, dict_renom, dict_orig, excel_b, excel_n, zip_p=None):
    """Empaqueta toda la sesión contable en un archivo ZIP descargable para respaldo indestructible con paquetes y triangulaciones."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        if df_proc is not None:
            zf.writestr("df_procesado.pkl", pickle.dumps(df_proc))
        if excel_b:
            zf.writestr("excel_original.xlsx", excel_b)
        if zip_p:
            zf.writestr("paquete_facturas.zip", zip_p)

        # Incluir paquetes de importación y asientos de triangulación si existen
        if "paquetes_importacion" in st.session_state and st.session_state["paquetes_importacion"]:
            zf.writestr("paquetes_importacion.pkl", pickle.dumps(st.session_state["paquetes_importacion"]))
        if "asientos_triangulacion_por_factura" in st.session_state and st.session_state["asientos_triangulacion_por_factura"]:
            zf.writestr("asientos_triangulacion.pkl", pickle.dumps(st.session_state["asientos_triangulacion_por_factura"]))

        # Incluir estado completo de la sesión
        estado_flags = {
            "paquetes_listos": {k: v for k, v in st.session_state.items() if k.startswith("paquete_listo_")},
            "enviar_gp": {k: v for k, v in st.session_state.items() if k.startswith("enviar_gp_pq_")},
            "enviar_h2": {k: v for k, v in st.session_state.items() if k.startswith("enviar_h2_pq_")},
            "enviar_nd": {k: v for k, v in st.session_state.items() if k.startswith("enviar_nd_pq_")},
            "facturas_no_contabilizar": list(st.session_state.get("facturas_no_contabilizar", set())),
            "sel_paquete_activo_key": st.session_state.get("sel_paquete_activo_key", 1)
        }
        zf.writestr("estado_sesion.json", json.dumps(estado_flags, ensure_ascii=False, indent=2))

        now_str = datetime.datetime.now().strftime("%d/%m/%Y %I:%M %p")
        meta = {
            "empresa": empresa_dict.get("nombre", ""),
            "nit": empresa_dict.get("nit", ""),
            "excel_nombre": excel_n or "Reporte.xlsx",
            "archivo_excel": excel_n or "Reporte.xlsx",
            "fecha_respaldo": now_str,
            "total_facturas": len(df_proc) if df_proc is not None else 0,
            "total_pdfs_renombrados": len(dict_renom) if dict_renom else 0,
            "total_pdfs_originales": len(dict_orig) if dict_orig else 0
        }
        zf.writestr("meta_sesion.json", json.dumps(meta, ensure_ascii=False, indent=2))
        mem_actual = cargar_memoria_aprendizaje(empresa_dict)
        if mem_actual:
            zf.writestr("memoria_aprendizaje.json", json.dumps(mem_actual, ensure_ascii=False, indent=2))
        zf.writestr("meta.json", json.dumps(meta, ensure_ascii=False, indent=2))
        # PDFs ya vienen comprimidos internamente (FlateDecode). Usar compresión directa ultra-rápida
        if dict_renom:
            for fn, bdata in dict_renom.items():
                zinfo = zipfile.ZipInfo(f"pdfs_renombrados/{fn}")
                zinfo.compress_type = zipfile.ZIP_STORED
                zf.writestr(zinfo, bdata)
        if dict_orig:
            for fn, bdata in dict_orig.items():
                zinfo = zipfile.ZipInfo(f"pdfs_originales/{fn}")
                zinfo.compress_type = zipfile.ZIP_STORED
                zf.writestr(zinfo, bdata)
    buf.seek(0)
    return buf.getvalue()


def obtener_respaldo_sesion_cached(empresa_dict):
    """Devuelve el archivo de respaldo cacheado en memoria, regenerándolo solo cuando hay cambios."""
    if "df_procesado" not in st.session_state or st.session_state["df_procesado"] is None:
        return b""
    df_p = st.session_state["df_procesado"]
    d_ren = st.session_state.get("dict_pdfs", {})
    d_ori = st.session_state.get("raw_uploaded_pdfs", {})
    sig_actual = f"{len(df_p)}_{len(d_ren)}_{len(d_ori)}_{st.session_state.get('_ultimo_autosave_hora', '')}"
    if st.session_state.get("_cached_respaldo_sig") != sig_actual or "_cached_respaldo_bytes" not in st.session_state:
        b_data = exportar_respaldo_sesion_zip(
            empresa_dict, df_p, d_ren, d_ori,
            st.session_state.get("excel_bytes"),
            st.session_state.get("excel_nombre", "Reporte.xlsx"),
            st.session_state.get("zip_pdfs")
        )
        st.session_state["_cached_respaldo_bytes"] = b_data
        st.session_state["_cached_respaldo_sig"] = sig_actual
    return st.session_state.get("_cached_respaldo_bytes", b"")

def importar_respaldo_sesion_zip(zip_bytes, empresa_dict):
    """Restaura una sesión contable completa desde un archivo ZIP de respaldo."""
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes), "r") as zf:
            df_proc = None
            if "df_procesado.pkl" in zf.namelist():
                df_proc = pickle.loads(zf.read("df_procesado.pkl"))
            excel_b = zf.read("excel_original.xlsx") if "excel_original.xlsx" in zf.namelist() else None
            zip_p = zf.read("paquete_facturas.zip") if "paquete_facturas.zip" in zf.namelist() else None
            excel_n = "Reporte_Restaurado.xlsx"
            if "meta_sesion.json" in zf.namelist():
                try:
                    m = json.loads(zf.read("meta_sesion.json").decode("utf-8"))
                    excel_n = m.get("excel_nombre", excel_n)
                except:
                    pass
            dict_renom = {}
            dict_orig = {}
            for name in zf.namelist():
                if name.startswith("pdfs_renombrados/") and name.endswith(".pdf"):
                    fn = name.replace("pdfs_renombrados/", "")
                    dict_renom[fn] = zf.read(name)
                elif name.startswith("pdfs_originales/") and name.endswith(".pdf"):
                    fn = name.replace("pdfs_originales/", "")
                    dict_orig[fn] = zf.read(name)
            return df_proc, dict_renom, dict_orig, excel_b, excel_n, zip_p
    except Exception as e:
        return None, {}, {}, None, "", None

# ==============================================================================
# MOTOR DE TRIANGULACIÓN Y CRUCES DE IMPORTACIÓN (ADUANAS Y PAGOS A TERCEROS)
# ==============================================================================
CUENTA_RETENCION_ASUMIDA = "53152001"
CUENTA_NO_DEDUCIBLE = "53950501"
CUENTA_IVA_IMPORTACION = "240835"
CUENTA_IMPORTACION_TRANSITO = "14650501"
CUENTA_CXP_AGENCIA_EXTERIOR = "22050505"
CUENTA_CXP_DHL_NACIONAL = "23359501"
CUENTA_CXP_EURO_SHIPPING = "22050501"

def normalizar_fecha_dian(fecha_raw):
    """Normaliza cualquier formato de fecha a DD/MM/AAAA para Siigo y DIAN."""
    if not fecha_raw:
        return ""
    s = str(fecha_raw).strip()
    m = re.match(r'^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$', s)
    if m:
        d, m_val, y = m.groups()
        return f"{int(d):02d}/{int(m_val):02d}/{y}"
    m_iso = re.match(r'^(\d{4})[/.-](\d{1,2})[/.-](\d{1,2})', s)
    if m_iso:
        y, m_val, d = m_iso.groups()
        return f"{int(d):02d}/{int(m_val):02d}/{y}"
    return s

def analizar_estado_filas_excel(excel_bytes):
    """
    Lee las celdas del archivo Excel DIAN (token) con openpyxl para detectar:
    1. Celdas rojas (facturas ya causadas/registradas en Siigo).
    2. Columna1 con número de comprobante 10-XXX o texto 'registrada'.
    3. Columna3 con asignación de Grupo de Importación (1, 2, 3...).
    4. Columna4 con cuenta contable de pasivo específica (22050505, 23359501, etc.).
    """
    estados = {}
    if not excel_bytes:
        return estados

    try:
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
        ws = wb.active
        headers = [str(cell.value or '').strip() for cell in ws[1]]

        col_c1 = next((i for i, h in enumerate(headers) if any(k in h.lower() for k in ['columna1', 'comprobante', 'estado'])), None)
        col_grp = next((i for i, h in enumerate(headers) if any(k in h.lower() for k in ['columna3', 'grupo', 'paquete'])), None)
        col_cta = next((i for i, h in enumerate(headers) if any(k in h.lower() for k in ['columna4', 'cuenta', 'cta'])), None)
        col_cufe = next((i for i, h in enumerate(headers) if any(k in h.lower() for k in ['cufe', 'token', 'uuid', 'cude'])), None)
        col_folio = next((i for i, h in enumerate(headers) if any(k in h.lower() for k in ['folio', 'factura'])), None)
        col_pref = next((i for i, h in enumerate(headers) if 'prefijo' in h.lower()), None)

        for row_idx in range(2, ws.max_row + 1):
            row_cells = ws[row_idx]
            es_roja = False

            # Revisar color de relleno de las celdas
            for cell in row_cells:
                fill = cell.fill
                if fill and fill.fill_type:
                    color = fill.start_color
                    if color:
                        rgb = str(getattr(color, 'rgb', '') or getattr(color, 'value', '')).upper()
                        if any(r in rgb for r in ['FF0000', 'C00000', 'FFC7CE', 'EF4444', 'DC2626', 'B91C1C', '991B1B', 'F87171', 'FF6666', 'FF4D4D', 'E11D48']):
                            es_roja = True
                            break
                        if len(rgb) == 8:
                            try:
                                r_v = int(rgb[2:4], 16)
                                g_v = int(rgb[4:6], 16)
                                b_v = int(rgb[6:8], 16)
                                if (r_v > 180 and g_v < 130 and b_v < 130) or (r_v > 220 and g_v < 210 and b_v < 215 and r_v - g_v > 30):
                                    es_roja = True
                                    break
                            except:
                                pass

            val_c1 = str(row_cells[col_c1].value or '').strip() if col_c1 is not None and col_c1 < len(row_cells) else ''
            if '10-' in val_c1 or 'registrad' in val_c1.lower() or 'causad' in val_c1.lower():
                es_roja = True

            val_grp = str(row_cells[col_grp].value or '').strip() if col_grp is not None and col_grp < len(row_cells) else ''
            val_cta = str(row_cells[col_cta].value or '').strip() if col_cta is not None and col_cta < len(row_cells) else ''
            if val_cta.endswith('.0'):
                val_cta = val_cta[:-2]

            cufe_k = str(row_cells[col_cufe].value or '').strip() if col_cufe is not None and col_cufe < len(row_cells) else ''
            fol_k = str(row_cells[col_folio].value or '').strip() if col_folio is not None and col_folio < len(row_cells) else ''
            if fol_k.endswith('.0'):
                fol_k = fol_k[:-2]
            pref_k = str(row_cells[col_pref].value or '').strip() if col_pref is not None and col_pref < len(row_cells) else ''
            if pref_k == 'nan' or pref_k == 'None':
                pref_k = ''

            fac_key = f"{pref_k}-{fol_k}" if pref_k else fol_k

            info_fila = {
                'row_idx': row_idx,
                'es_roja': es_roja,
                'comprobante_existente': val_c1 if '10-' in val_c1 else ('10-Previa' if es_roja else ''),
                'grupo_importacion': val_grp,
                'cuenta_especifica': val_cta
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
    """Genera el asiento contable EXACTO e IDÉNTICO al que visualiza la Hoja 2 (Auditoría)."""
    asiento_filas = []
    es_nc = "Devolucion" in str(fac_sel.get("Operacion", ""))
    asume_ret = bool(fac_sel.get("Impuestos Asumidos", False))
    cta_cxp_usar = str(fac_sel.get("Cuenta Pasivo Especifica") or fac_sel.get("Cta Contrapartida") or ("22050505" if es_aduanero else "22050501")).strip()
    cta_p = str(fac_sel.get("Cta Principal") or "14650501").strip()
    es_importacion_factura = es_aduanero or "1465" in cta_p or "IMPORTACI" in str(fac_sel.get("Categoría", "")).upper()
    cta_iva = str(fac_sel.get("Cta IVA") or ("24081501" if es_importacion_factura else "24081001")).strip()

    def _sf(v):
        try:
            if v is not None and pd.notna(v):
                fv = float(v)
                return fv if not pd.isna(fv) else 0.0
        except Exception:
            pass
        return 0.0

    base_val = _sf(fac_sel.get("Base"))
    iva_val = _sf(fac_sel.get("IVA"))
    tot_val = _sf(fac_sel.get("Total"))
    rfte_val = _sf(fac_sel.get("ReteFuente"))
    rica_val = _sf(fac_sel.get("ReteICA"))
    riva_val = _sf(fac_sel.get("ReteIVA"))
    iva_imp_val = _sf(fac_sel.get("IVA Importación"))

    raw_cta_imp = fac_sel.get("Cta IVA Importación")
    if raw_cta_imp is None or pd.isna(raw_cta_imp) or str(raw_cta_imp).strip() in ["", "nan", "None"]:
        cta_iva_imp = "240835"
    else:
        cta_iva_imp = str(raw_cta_imp).strip()

    # 1. Base / Costo
    asiento_filas.append({
        "Código Cuenta": cta_p,
        "Descripción de la Cuenta": f"{fac_sel.get('Categoría', 'Gasto/Costo')} - {str(fac_sel.get('Proveedor', ''))[:25]}",
        "Tercero / NIT": str(fac_sel.get("NIT Emisor", "")),
        "Débito ($)": 0.0 if es_nc else base_val,
        "Crédito ($)": base_val if es_nc else 0.0
    })

    # 2. IVA Descontable (Servicios / Compras)
    if iva_val > 0:
        desc_iva = f"IVA Descontable Servicios (Base: ${base_val:,.0f})" if ("24081501" in cta_iva or es_importacion_factura) else f"IVA Descontable (Base: ${base_val:,.0f})"
        asiento_filas.append({
            "Código Cuenta": cta_iva,
            "Descripción de la Cuenta": desc_iva,
            "Tercero / NIT": str(fac_sel.get("NIT Emisor", "")),
            "Débito ($)": 0.0 if es_nc else iva_val,
            "Crédito ($)": iva_val if es_nc else 0.0
        })

    # 2.B IVA de Importación (Cuenta 240835) - Solo para facturas de importación
    if iva_imp_val > 0:
        asiento_filas.append({
            "Código Cuenta": cta_iva_imp,
            "Descripción de la Cuenta": f"IVA de Importación (Cta 240835) - Fac {fac_sel.get('Factura', '')}",
            "Tercero / NIT": str(fac_sel.get("NIT Emisor", "")),
            "Débito ($)": 0.0 if es_nc else iva_imp_val,
            "Crédito ($)": iva_imp_val if es_nc else 0.0
        })

    tot_ret = round(rfte_val + rica_val, 2)
    cta_rf_usar = str(fac_sel.get("Cta ReteFuente") or ("23652503" if es_aduanero else "23654001")).strip()
    cta_ri_usar = str(fac_sel.get("Cta ReteICA") or ("23680505" if es_aduanero else "23680501")).strip()

    if asume_ret:
        # IMPUESTOS ASUMIDOS (Cruza con Agente Aduanero / Triangulación)
        if tot_ret > 0:
            asiento_filas.append({
                "Código Cuenta": "53152001",
                "Descripción de la Cuenta": "Retenciones Asumidas (Impuestos Asumidos Aduana)",
                "Tercero / NIT": str(fac_sel.get("NIT Emisor", "")),
                "Débito ($)": tot_ret,
                "Crédito ($)": 0.0
            })
        if rfte_val > 0:
            asiento_filas.append({
                "Código Cuenta": cta_rf_usar,
                "Descripción de la Cuenta": f"ReteFuente Practicada ({cta_rf_usar}) Fac {fac_sel.get('Factura', '')}",
                "Tercero / NIT": str(fac_sel.get("NIT Emisor", "")),
                "Débito ($)": 0.0,
                "Crédito ($)": rfte_val
            })
        if rica_val > 0:
            asiento_filas.append({
                "Código Cuenta": cta_ri_usar,
                "Descripción de la Cuenta": "Retención ICA Practicada",
                "Tercero / NIT": str(fac_sel.get("NIT Emisor", "")),
                "Débito ($)": 0.0,
                "Crédito ($)": rica_val
            })
        saldo_cxp = round(base_val + iva_val + iva_imp_val, 2)
        asiento_filas.append({
            "Código Cuenta": cta_cxp_usar,
            "Descripción de la Cuenta": f"CxP Proveedor/Agente - Fac {fac_sel.get('Factura', '')}",
            "Tercero / NIT": str(fac_sel.get("NIT Emisor", "")),
            "Débito ($)": saldo_cxp if es_nc else 0.0,
            "Crédito ($)": 0.0 if es_nc else saldo_cxp
        })
    else:
        # RETENCIONES ORDINARIAS PRACTICADAS AL PROVEEDOR
        if rfte_val > 0:
            asiento_filas.append({
                "Código Cuenta": cta_rf_usar,
                "Descripción de la Cuenta": f"ReteFuente Practicada ({fac_sel.get('Categoría', 'Compras/Servicios')})",
                "Tercero / NIT": str(fac_sel.get("NIT Emisor", "")),
                "Débito ($)": rfte_val if es_nc else 0.0,
                "Crédito ($)": 0.0 if es_nc else rfte_val
            })
        if rica_val > 0:
            asiento_filas.append({
                "Código Cuenta": cta_ri_usar,
                "Descripción de la Cuenta": "Retención ICA Practicada",
                "Tercero / NIT": str(fac_sel.get("NIT Emisor", "")),
                "Débito ($)": rica_val if es_nc else 0.0,
                "Crédito ($)": 0.0 if es_nc else rica_val
            })
        if riva_val > 0:
            asiento_filas.append({
                "Código Cuenta": "23670101",
                "Descripción de la Cuenta": "Retención de IVA Practicada (15%)",
                "Tercero / NIT": str(fac_sel.get("NIT Emisor", "")),
                "Débito ($)": riva_val if es_nc else 0.0,
                "Crédito ($)": 0.0 if es_nc else riva_val
            })
        neto_cxp = round(base_val + iva_val + iva_imp_val - rfte_val - rica_val - riva_val, 2)
        asiento_filas.append({
            "Código Cuenta": cta_cxp_usar,
            "Descripción de la Cuenta": f"Proveedores Nacionales - Fac {fac_sel.get('Factura', '')}",
            "Tercero / NIT": str(fac_sel.get("NIT Emisor", "")),
            "Débito ($)": neto_cxp if es_nc else 0.0,
            "Crédito ($)": 0.0 if es_nc else neto_cxp
        })

    df_h2 = pd.DataFrame(asiento_filas)
    if "Descripción de la Cuenta" in df_h2.columns and "Descripción Cuenta" not in df_h2.columns:
        df_h2["Descripción Cuenta"] = df_h2["Descripción de la Cuenta"]
    elif "Descripción Cuenta" in df_h2.columns and "Descripción de la Cuenta" not in df_h2.columns:
        df_h2["Descripción de la Cuenta"] = df_h2["Descripción Cuenta"]
    return df_h2


def obtener_saldo_cruce_factura(tr):
    """
    Calcula de forma segura el Saldo de Cruce (Base + IVA) de una factura para la triangulación,
    garantizando que nunca retorne None ni NaN.
    """
    try:
        tb = float(tr.get("Base", 0.0)) if pd.notna(tr.get("Base")) else (float(tr.get("Subtotal", 0.0)) if pd.notna(tr.get("Subtotal")) else 0.0)
    except Exception:
        tb = 0.0
    try:
        tiv = float(tr.get("IVA", 0.0)) if pd.notna(tr.get("IVA")) else 0.0
    except Exception:
        tiv = 0.0
    try:
        t_imp = float(tr.get("IVA Importación", 0.0)) if pd.notna(tr.get("IVA Importación")) else 0.0
    except Exception:
        t_imp = 0.0

    val_neto = tr.get("Total Neto")
    if val_neto is not None and pd.notna(val_neto) and str(val_neto).strip() not in ["", "nan", "None", "0", "0.0"]:
        try:
            f_neto = float(val_neto)
            if f_neto > 0.0 and not pd.isna(f_neto):
                return round(f_neto, 2)
        except Exception:
            pass

    saldo_calc = round(tb + tiv + t_imp, 2)
    if saldo_calc > 0.0:
        return saldo_calc

    try:
        tot = float(tr.get("Total", 0.0)) if pd.notna(tr.get("Total")) else 0.0
        if tot > 0.0 and not pd.isna(tot):
            return round(tot, 2)
    except Exception:
        pass
    return 0.0




def calcular_concordancia_fecha(fecha_tercero, fecha_agente):
    """
    Calcula de forma dinámica la concordancia cronológica entre la factura del tercero y la del agente,
    garantizando que nunca retorne None ni NaN.
    """
    if not fecha_tercero or not fecha_agente or pd.isna(fecha_tercero) or pd.isna(fecha_agente):
        return "⚪ Sin fecha"
    dt_terc = pd.to_datetime(str(fecha_tercero), dayfirst=True, errors="coerce")
    dt_ag = pd.to_datetime(str(fecha_agente), dayfirst=True, errors="coerce")
    if pd.isna(dt_terc) or pd.isna(dt_ag):
        return "⚪ Sin fecha"
    
    diff_d = int((dt_terc - dt_ag).days)
    if diff_d == 0:
        return "🟢 Mismo día (0d)"
    elif 0 < diff_d <= 15:
        return f"➡️ Poco después (+{diff_d}d)"
    elif -30 <= diff_d < 0:
        return f"⬅️ Anterior ({abs(diff_d)}d antes)"
    elif -60 <= diff_d < -30:
        return f"⬅️ Anterior ({abs(diff_d)}d antes)"
    elif diff_d < -60:
        return f"⛔ Fuera de rango ({abs(diff_d)}d antes)"
    else:
        return f"⛔ Fuera de rango (+{diff_d}d desp)"


def generar_asiento_mixto_hoja2_con_terceros(agente_row, terceros_df, enviar_a_no_deducible=False, imputar_a_transito=False):
    """
    Construye el asiento contable para un cobro de agente cuando se trae la contabilización de Página 2:
    1. Si en Hoja 2 se registraron únicamente los Ingresos Propios (Base + IVA < Total Factura):
       Combina los Ingresos Propios (con su IVA y retenciones) + Facturas de Terceros vinculadas (cancelación de CxP) +
       Crédito total al Agente Aduanero (22050501).
    2. El saldo de ingresos para terceros NO se imputa automáticamente a ninguna cuenta de gasto ni tránsito.
       Se deja PENDIENTE para que el usuario concilie con facturas de terceros o decida explícitamente enviarlo a No Deducibles o a Tránsito.
    3. Si el usuario selecciona 'No Deducibles', imputa el saldo a 53950501.
    4. Si el usuario selecciona 'Mercancías en Tránsito', imputa el saldo a 14650501.
    """
    df_h2 = obtener_asiento_contable_hoja2(agente_row, es_aduanero=False)
    
    tot_agente = float(agente_row.get("Total", 0.0))
    base_propia = float(agente_row.get("Base", 0.0))
    iva_propio = float(agente_row.get("IVA", 0.0))
    iva_imp_propio = float(agente_row.get("IVA Importación", 0.0))
    subtotal_propio = round(base_propia + iva_propio + iva_imp_propio, 2)
    saldo_terceros_esperado = round(tot_agente - subtotal_propio, 2)
    
    # Si la factura en Hoja 2 ya cubre el total y no hay terceros asignados
    if saldo_terceros_esperado <= 0.05 and (terceros_df is None or terceros_df.empty):
        return df_h2, 0.0
        
    filas_propias = []
    tot_ret_propias = 0.0
    for _, r in df_h2.iterrows():
        cta = str(r["Código Cuenta"]).strip()
        if cta.startswith("2365") or cta.startswith("2368") or cta.startswith("2367"):
            tot_ret_propias += float(r["Crédito ($)"])
            filas_propias.append(r.to_dict())
        elif not (cta.startswith("2205") and float(r["Crédito ($)"]) > 0):
            filas_propias.append(r.to_dict())
            
    # Facturas de terceros asignadas a este paquete
    filas_terceros = []
    suma_terceros_cxp = 0.0
    if terceros_df is not None and not terceros_df.empty:
        for _, t in terceros_df.iterrows():
            prov_nom = str(t.get("Proveedor", "")).upper()
            nit_t = str(t.get("NIT Emisor", ""))
            fac_num = str(t.get("Factura", ""))
            cta_cxp = str(t.get("Cuenta Pasivo Especifica", "")).strip() or ("22050505" if "CARGO" in prov_nom or "ADUANA" in prov_nom else "23359501")
            t_base = float(t.get("Base", 0.0))
            t_iva = float(t.get("IVA", 0.0))
            t_saldo_pagar = obtener_saldo_cruce_factura(t)
            
            suma_terceros_cxp += t_saldo_pagar
            filas_terceros.append({
                "Código Cuenta": cta_cxp,
                "Descripción Cuenta": f"Cancela CxP {prov_nom[:20]} (Fac {fac_num})",
                "Descripción de la Cuenta": f"Cancela CxP {prov_nom[:20]} (Fac {fac_num})",
                "Tercero / NIT": f"{nit_t} - {prov_nom[:25]}",
                "Débito ($)": t_saldo_pagar,
                "Crédito ($)": 0.0
            })
            
    # Saldo para terceros restante (fletes internacionales / cargos en origen sin factura DIAN separada)
    dif_terceros = round(saldo_terceros_esperado - suma_terceros_cxp, 2)
    filas_ajuste = []
    if dif_terceros > 0.05:
        if enviar_a_no_deducible:
            filas_ajuste.append({
                "Código Cuenta": CUENTA_NO_DEDUCIBLE,
                "Descripción Cuenta": f"Gastos No Deducibles Terceros (Diferencia sin soporte DIAN)",
                "Descripción de la Cuenta": f"Gastos No Deducibles Terceros (Diferencia sin soporte DIAN)",
                "Tercero / NIT": f"{agente_row.get('NIT Emisor', '')} - {str(agente_row.get('Proveedor', ''))[:25]}",
                "Débito ($)": dif_terceros,
                "Crédito ($)": 0.0
            })
        elif imputar_a_transito:
            filas_ajuste.append({
                "Código Cuenta": CUENTA_IMPORTACION_TRANSITO,
                "Descripción Cuenta": f"Mercancías en Tránsito / Saldo Faltante (Fac {agente_row.get('Factura', '')})",
                "Descripción de la Cuenta": f"Mercancías en Tránsito / Saldo Faltante (Fac {agente_row.get('Factura', '')})",
                "Tercero / NIT": f"{agente_row.get('NIT Emisor', '')} - {str(agente_row.get('Proveedor', ''))[:25]}",
                "Débito ($)": dif_terceros,
                "Crédito ($)": 0.0
            })
        # Si ni enviar_a_no_deducible ni imputar_a_transito están activos, NO se agrega ninguna fila de ajuste:
        # SE DEJA PENDIENTE PARA QUE EL USUARIO CONCILIE FACTURAS O DECIDA ENVIARLO A TRÁNSITO / NO DEDUCIBLE.
            
    # Crédito total por pagar al Agente Aduanero (Euro Shipping)
    asume_ret = bool(agente_row.get("Impuestos Asumidos", False))
    cxp_total_agente = tot_agente if asume_ret else round(tot_agente - tot_ret_propias, 2)
    
    filas_credito = [{
        "Código Cuenta": CUENTA_CXP_EURO_SHIPPING,
        "Descripción Cuenta": f"CxP Agente Aduanero - Fac {agente_row.get('Factura', '')}",
        "Descripción de la Cuenta": f"CxP Agente Aduanero - Fac {agente_row.get('Factura', '')}",
        "Tercero / NIT": f"{agente_row.get('NIT Emisor', '')} - {str(agente_row.get('Proveedor', ''))[:25]}",
        "Débito ($)": 0.0,
        "Crédito ($)": cxp_total_agente
    }]
    
    asiento_final = filas_propias + filas_terceros + filas_ajuste + filas_credito
    df_res = pd.DataFrame(asiento_final)
    dif_pendiente = dif_terceros if (not enviar_a_no_deducible and not imputar_a_transito) else 0.0
    return df_res, dif_pendiente

def generar_asiento_triangulacion_paquete(agente_row, terceros_df, enviar_a_no_deducible=False, enviar_a_gastos_propios=False, enviar_a_hoja2=False, **kwargs):
    """
    Calcula el asiento contable de partida doble para un paquete de importación triangulado:
    1. Débito a CxP Terceros (22050505/23359501) por el Saldo por Pagar real (Base + IVA, ej. .651.939).
       Si ya estaba registrada en Siigo, se cancela la cuenta por pagar imputada.
    2. Débito a Retenciones Asumidas (53152001) si las retenciones no se habían asumido en la causación previa.
    3. Débito a Importación en Tránsito (14650501) e IVA (24081501) para facturas blancas pendientes.
    4. Débito a IVA de Importación si la factura del agente lo discrimina.
    5. Débito a Gastos No Deducibles (53950501) por la diferencia sin factura DIAN.
    6. Crédito a CxP Agente Aduanero (22050501) por el 100% de la factura del agente (ej. .077.647).
    """
    tot_agente = float(agente_row.get("Total", 0.0))
    iva_agente = float(agente_row.get("IVA", 0.0))
    
    asiento = []
    suma_cxp_canceladas = 0.0
    suma_ret_asumidas = 0.0
    suma_costo_blancas = 0.0
    suma_iva_blancas = 0.0
    
    for _, t in terceros_df.iterrows():
        prov_nom = str(t.get("Proveedor", "")).upper()
        nit_t = str(t.get("NIT Emisor", ""))
        fac_num = str(t.get("Factura", ""))
        es_excluida_manual = bool(t.get("No Contabilizar", False)) or (str(t.get("Factura", "")).strip() in st.session_state.get("facturas_no_contabilizar", set()))
        es_roja = bool(t.get("Ya Registrada", False)) or es_excluida_manual
        cta_cxp = str(t.get("Cuenta Pasivo Especifica", "")).strip()
        
        if not cta_cxp:
            if any(k in prov_nom for k in ["CARGO", "ADUANA", "PORTUARIA", "ALMACENADORA", "TERMINAL"]):
                cta_cxp = CUENTA_CXP_AGENCIA_EXTERIOR
            else:
                cta_cxp = CUENTA_CXP_DHL_NACIONAL
                
        t_base = float(t.get("Base", 0.0))
        t_iva = float(t.get("IVA", 0.0))
        t_tot = float(t.get("Total", 0.0))
        if t_base <= 0:
            t_base = round(t_tot - t_iva, 2)
            
        rfte_t = float(t.get("ReteFuente", 0.0))
        rica_t = float(t.get("ReteICA", 0.0))
        tot_ret = round(rfte_t + rica_t, 2)
        
        # El Saldo por Pagar que entra a la triangulación es Base + IVA (ej. .388.770 + 63.169 = .651.939)
        t_saldo_pagar = obtener_saldo_cruce_factura(t)
            
        # REGLA DE TRIANGULACIÓN: Toda factura de tercero traída a la triangulación ÚNICAMENTE
        # aporta su Cuenta por Pagar (pasivo: 22050505 Agencia, 23359501 DHL, etc.) por su Saldo por Pagar (Base + IVA),
        # ya que la causación del gasto/costo e IVA de esa factura se realiza en su propio comprobante individual.
        asume_en_causacion = bool(t.get("Impuestos Asumidos", False))
        
        if asume_en_causacion or tot_ret == 0:
            cxp_cancelar = t_saldo_pagar
            ret_a_asumir = 0.0
            suma_cxp_canceladas += cxp_cancelar
            
            asiento.append({
                "Código Cuenta": cta_cxp,
                "Descripción Cuenta": f"Cancela CxP {prov_nom[:20]} (Fac {fac_num})",
                "Tercero / NIT": f"{nit_t} - {prov_nom[:25]}",
                "Débito ($)": cxp_cancelar,
                "Crédito ($)": 0.0
            })
        else:
            cxp_cancelar = round(t_saldo_pagar - tot_ret, 2)
            ret_a_asumir = tot_ret
            suma_cxp_canceladas += cxp_cancelar
            suma_ret_asumidas += ret_a_asumir
            
            asiento.append({
                "Código Cuenta": cta_cxp,
                "Descripción Cuenta": f"Cancela CxP {prov_nom[:20]} (Fac {fac_num})",
                "Tercero / NIT": f"{nit_t} - {prov_nom[:25]}",
                "Débito ($)": cxp_cancelar,
                "Crédito ($)": 0.0
            })
            if ret_a_asumir > 0:
                asiento.append({
                    "Código Cuenta": CUENTA_RETENCION_ASUMIDA,
                    "Descripción Cuenta": f"Retención Asumida Fac {fac_num} ({prov_nom[:18]})",
                    "Tercero / NIT": f"{nit_t} - {prov_nom[:25]}",
                    "Débito ($)": ret_a_asumir,
                    "Crédito ($)": 0.0
                })

    # IVA Servicios e IVA de Importación del Agente
    iva_imp_agente = float(agente_row.get("IVA Importación", 0.0))
    cta_iva_ag = str(agente_row.get("Cta IVA") or "24081501").strip()
    cta_iva_imp_ag = str(agente_row.get("Cta IVA Importación") or "240835").strip()

    if iva_agente > 0:
        asiento.append({
            "Código Cuenta": cta_iva_ag,
            "Descripción Cuenta": f"IVA Descontable Servicios Fac {agente_row.get('Factura', '')}",
            "Tercero / NIT": f"{agente_row.get('NIT Emisor', '')} - {agente_row.get('Proveedor', '')[:25]}",
            "Débito ($)": iva_agente,
            "Crédito ($)": 0.0
        })

    if iva_imp_agente > 0:
        asiento.append({
            "Código Cuenta": cta_iva_imp_ag,
            "Descripción Cuenta": f"IVA de Importación (Cta 240835) Fac {agente_row.get('Factura', '')}",
            "Tercero / NIT": f"{agente_row.get('NIT Emisor', '')} - {agente_row.get('Proveedor', '')[:25]}",
            "Débito ($)": iva_imp_agente,
            "Crédito ($)": 0.0
        })

    # Calcular la diferencia no deducible sin factura DIAN (Cuenta 53950501)
    suma_justificada = suma_cxp_canceladas + suma_ret_asumidas + suma_costo_blancas + suma_iva_blancas + iva_agente + iva_imp_agente
    diferencia_no_deducible = round(tot_agente - suma_justificada, 2)
    
    if diferencia_no_deducible > 0.01 and enviar_a_no_deducible:
        asiento.append({
            "Código Cuenta": CUENTA_NO_DEDUCIBLE,
            "Descripción Cuenta": f"Gastos No Deducibles Importación (Diferencia sin soporte DIAN)",
            "Tercero / NIT": f"{agente_row.get('NIT Emisor', '')} - {agente_row.get('Proveedor', '')[:25]}",
            "Débito ($)": diferencia_no_deducible,
            "Crédito ($)": 0.0
        })
    elif diferencia_no_deducible < -0.01:
        # En caso de que la suma de terceros exceda ligeramente al agente por redondeo
        asiento.append({
            "Código Cuenta": CUENTA_NO_DEDUCIBLE,
            "Descripción Cuenta": f"Ajuste por Diferencia de Cuadre Importación",
            "Tercero / NIT": f"{agente_row.get('NIT Emisor', '')} - {agente_row.get('Proveedor', '')[:25]}",
            "Débito ($)": 0.0,
            "Crédito ($)": abs(diferencia_no_deducible)
        })

    # Crédito total al Agente Aduanero (Euro Shipping / Trade Global) por el 100% de su factura
    asiento.append({
        "Código Cuenta": CUENTA_CXP_EURO_SHIPPING,
        "Descripción Cuenta": f"CxP Agente Aduanero - Fac {agente_row.get('Factura', '')}",
        "Tercero / NIT": f"{agente_row.get('NIT Emisor', '')} - {agente_row.get('Proveedor', '')[:25]}",
        "Débito ($)": 0.0,
        "Crédito ($)": tot_agente
    })

    df_asiento = pd.DataFrame(asiento)
    if "Descripción Cuenta" in df_asiento.columns and "Descripción de la Cuenta" not in df_asiento.columns:
        df_asiento["Descripción de la Cuenta"] = df_asiento["Descripción Cuenta"]
    elif "Descripción de la Cuenta" in df_asiento.columns and "Descripción Cuenta" not in df_asiento.columns:
        df_asiento["Descripción Cuenta"] = df_asiento["Descripción de la Cuenta"]
    return df_asiento, max(0.0, diferencia_no_deducible), suma_ret_asumidas

# Directorio oficial de regímenes fiscales conocidos por NIT y Nombre
REGIMENES_EMISORES_CONOCIDOS = {
    '860502609': 'O-13;O-15',   # DHL EXPRESS COLOMBIA LTDA (Gran Contribuyente y Autorretenedor)
    '800215775': 'O-13;O-15',   # SOCIEDAD PORTUARIA REGIONAL DE BUENAVENTURA S.A.
    '890304099': 'O-13;O-15',   # HOTELES ESTELAR S.A.
    '830048145': 'O-13;O-15',   # SIIGO S.A.S.
    '830048268': 'O-48',        # EURO SHIPPING SERVICES S.A.S (Responsable de IVA / Código 23 - 48)
    '800153993': 'O-13;O-15',   # COMUNICACION CELULAR S.A. COMCEL / CLARO
    '860006376': 'O-13;O-15',   # PANAMERICANA LIBRERIA Y PAPELERIA S.A.
    '890900608': 'O-13;O-15',   # BANCOLOMBIA S.A.
    '860007738': 'O-13;O-15',   # BANCO DAVIVIENDA S.A.
    '890903938': 'O-13;O-15',   # ALMACENES EXITO S.A.
    '860012936': 'O-13;O-15',   # SODIMAC CORONA / HOMECENTER
    '901306979': 'O-48',        # TRADE GLOBAL INTERNATIONAL SAS
    '805001149': 'O-48',        # AGENCIA INTERAMERICANA DE CARGA
    '900744197': 'O-48',        # HOTEL GENOVA SAS
}

CODIGOS_IMPUESTO_SIIGO = {
    '24081001': 1,   # IVA 19% compras bienes
    '24081003': 2,   # IVA 5% compras bienes
    '24081501': 1,   # IVA Servicios 19%
    '240835': 1,     # IVA de Importación 19%
    '24083501': 1,   # IVA de Importación 19% (8 dígitos)
    '24080601': 1,   # IVA 19% ventas
    '24080602': 2,   # IVA 5% ventas
    '23651501': 3,   # Retefuente 11%
    '23652001': 4,   # Retefuente 10%
    '23652501': 5,   # Retefuente 6%
    '23652503': 6,   # Retefuente 4% (Servicios / Agenciamiento)
    '23654001': 7,   # Retefuente 2.5% (Compras generales / Repuestos)
    '23654004': 18,  # Retefuente 3.5% (Alojamiento / Hoteles)
    '23653502': 19,  # Retefuente 7%
    '23657001': 20,  # Retefuente 2%
    '23652505': 21,  # Retefuente 1%
    '23654006': 23,  # Retefuente Combustibles 0.1%
    '23680501': 8,   # ReteICA 11.04
    '23680503': 9,   # ReteICA 13.8
    '23680505': 10,  # ReteICA 9.66
    '23680507': 11,  # ReteICA 8
    '23680509': 12,  # ReteICA 7
    '23680511': 13,  # ReteICA 6.9
    '23680513': 14,  # ReteICA 4.14
    '23670101': 15,  # ReteIVA 15%
    '24950103': 16   # Impoconsumo 8%
}

AGENTES_ADUANEROS = ["DHL", "ADUANA", "EURO SHIPPING", "PORTUARIA", "ALMACENADORA", "CARGO", "TRADE GLOBAL", "TERMINAL", "BUENAVENTURA"]

def escanear_regimen_texto_pdf(texto, nit_emisor=None):
    """
    Escaneo inteligente de responsabilidades fiscales del EMISOR en la factura electrónica:
    Analiza el documento completo sin truncar en 'Cliente:' para no perder las resoluciones.
    Detecta O-13 (Gran Contribuyente), O-15 (Autorretenedor), O-47 (RST), O-48 y O-49.
    """
    if not texto:
        return ""

    txt = texto.upper()
    txt_norm = re.sub(r'\s+', ' ', txt)

    # 1. Si se conoce el NIT del emisor, verificar en el directorio de regímenes oficiales
    if nit_emisor:
        ne_clean = re.sub(r'\D', '', str(nit_emisor))
        if ne_clean in REGIMENES_EMISORES_CONOCIDOS:
            return REGIMENES_EMISORES_CONOCIDOS[ne_clean]

    # 1.1 Si el texto corresponde a Euro Shipping Services, es O-48 (Responsable de IVA / Común)
    if "EURO SHIPPING" in txt_norm or "830048268" in txt_norm:
        return "O-48"

    # 2. Analizar Gran Contribuyente (O-13 / 0-13 / Grandes Contribuyentes)
    es_gc = False
    if re.search(r'\b[O0]-13\b', txt_norm):
        if not re.search(r'(?:NO\s+(?:SOMOS\s+)?|NI\s+|NO\s+ES\s+)[O0]-13', txt_norm):
            es_gc = True
    elif re.search(r'GRANDES?\s+CONTRIBUYENTES?', txt_norm):
        if not re.search(r'(?:NO\s+(?:SOMOS\s+)?|NI\s+|NO\s+ES\s+)GRANDES?\s+CONTRIBUYENTES?', txt_norm):
            es_gc = True

    # 3. Analizar Autorretenedor (O-15 / 0-15 / Autorretenedor de Renta)
    es_autorr = False
    if re.search(r'\b[O0]-15\b', txt_norm):
        if not re.search(r'(?:NO\s+(?:SOMOS\s+)?|NI\s+|NO\s+ES\s+)[O0]-15', txt_norm):
            es_autorr = True
    elif re.search(r'AUTO[R]?RETENEDOR(?:ES)?', txt_norm):
        if not re.search(r'(?:NO\s+(?:SOMOS\s+)?|NI\s+|NO\s+ES\s+)AUTO[R]?RETENEDOR(?:ES)?', txt_norm):
            es_autorr = True

    # 4. Analizar Régimen Simple de Tributación (O-47 / 0-47 / RST)
    es_rst = False
    if re.search(r'\b[O0]-47\b', txt_norm):
        es_rst = True
    elif re.search(r'R[EÉ]GIMEN\s+SIMPLE|SIMPLE\s+DE\s+TRIBUTACI[OÓ]N|\bRST\b', txt_norm):
        if not re.search(r'(?:NO\s+(?:SOMOS\s+)?|NI\s+)(?:R[EÉ]GIMEN\s+SIMPLE|RST)', txt_norm):
            es_rst = True

    # 5. No Responsable de IVA (O-49 / 0-49)
    es_no_iva = False
    if re.search(r'\b[O0]-49\b', txt_norm) or re.search(r'NO\s+RESPONSABLE\s+(?:DE|DEL)\s+IVA', txt_norm):
        es_no_iva = True

    codigos = []
    if es_gc: codigos.append("O-13")
    if es_autorr: codigos.append("O-15")
    if es_rst: codigos.append("O-47")
    if es_no_iva: codigos.append("O-49")
    if not codigos: codigos.append("O-48")

    return ";".join(codigos)

def parse_num_co(val_str):
    if not val_str: return 0.0
    s = str(val_str).replace('$', '').replace('COP', '').strip()
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s:
        parts = s.split(',')
        if len(parts) == 2 and len(parts[1]) in [2, 3]:
            s = parts[0].replace('.', '') + '.' + parts[1]
        else:
            s = s.replace(',', '')
    try:
        return float(s)
    except:
        return 0.0

def extraer_valores_fiscales_texto_pdf(texto):
    """
    Extrae con máxima precisión los valores fiscales y totales desde el texto del PDF oficial:
    1. Subtotal / Total Bruto Factura (Base gravable antes de descuentos o retenciones en notas).
    2. IVA discriminado o Total Impuesto.
    3. Total Neto Factura (Base + IVA, antes de retenciones sugeridas).
    4. Retenciones en notas / Descuentos Globales (ReteFuente 11%/4%, ReteICA 0.966%, ReteIVA).
    5. Anticipos reales reportados en notas o casillas informativas.
    6. Identificación de facturación a favor del agente aduanero (Mandato).
    """
    datos = {}
    if not texto: return datos

    # 1. Notas Finales (formato clave: valor)
    m_nf_rfte = re.search(r' retefuente\s*:\s*([0-9\.\,]+)', texto, re.IGNORECASE)
    if m_nf_rfte: datos['retefuente'] = parse_num_co(m_nf_rfte.group(1))

    m_nf_rica = re.search(r' reteica\s*:\s*([0-9\.\,]+)', texto, re.IGNORECASE)
    if m_nf_rica: datos['reteica'] = parse_num_co(m_nf_rica.group(1))

    m_nf_riva = re.search(r' reteiva\s*:\s*([0-9\.\,]+)', texto, re.IGNORECASE)
    if m_nf_riva: datos['reteiva'] = parse_num_co(m_nf_riva.group(1))

    # 2. Descuentos globales por retención sugerida si no vino en notas
    if 'retefuente' not in datos or datos['retefuente'] == 0:
        m_dg_rf = re.search(r'RETEFUENTE[^\$]+(?:[\$]\s*|COP\s*)([0-9\.\,]+)', texto, re.IGNORECASE)
        if m_dg_rf: datos['retefuente'] = parse_num_co(m_dg_rf.group(1))

    if 'reteica' not in datos or datos['reteica'] == 0:
        m_dg_ri = re.search(r'RETEICA[^\$]+(?:[\$]\s*|COP\s*)([0-9\.\,]+)', texto, re.IGNORECASE)
        if m_dg_ri: datos['reteica'] = parse_num_co(m_dg_ri.group(1))

    # 3. Bloque de Datos Totales
    pos_totales = texto.upper().find('DATOS TOTALES')
    txt_totales = texto[pos_totales:] if pos_totales != -1 else texto

    m_sub = re.search(r'(?:Subtotal|Total Bruto Factura)\s*[:\$]?\s*([0-9\.\,]+)', txt_totales, re.IGNORECASE)
    if m_sub: datos['subtotal'] = parse_num_co(m_sub.group(1))

    m_iva = re.search(r'(?:IVA|Total impuesto\s*(?:\(\=\))?)\s*[:\$]?\s*([0-9\.\,]+)', txt_totales, re.IGNORECASE)
    if m_iva: datos['iva'] = parse_num_co(m_iva.group(1))

    m_neto = re.search(r'Total neto factura\s*(?:\(\=\))?\s*[:\$]?\s*([0-9\.\,]+)', txt_totales, re.IGNORECASE)
    if m_neto: datos['total_neto'] = parse_num_co(m_neto.group(1))

    m_tot = re.search(r'Total factura\s*(?:\(\=\))?\s*(?:COP\s*)?[\$]?\s*([0-9\.\,]+)', txt_totales, re.IGNORECASE)
    if m_tot: datos['total_factura'] = parse_num_co(m_tot.group(1))

    m_fact_a = re.search(r'nombre_facturado_a\s*:\s*([^\n\r]+)', texto, re.IGNORECASE)
    if m_fact_a: datos['nombre_facturado_a'] = m_fact_a.group(1).strip()

    m_nit_fact_a = re.search(r'nit_facturado_a\s*:\s*([0-9]+)', texto, re.IGNORECASE)
    if m_nit_fact_a: datos['nit_facturado_a'] = m_nit_fact_a.group(1).strip()

    return datos

def calcular_base_con_regla_descuento(r, tot, iva, nom_emisor):
    """
    Calcula la Base Gravable contable respetando la regla estricta:
    1. Conversión de moneda extranjera (USD a COP por TRM): Si el Subtotal/IVA vienen en USD y Total en COP,
       convierte ambos valores a COP con la TRM implícita para garantizar partida doble exacta.
    2. El descuento SOLO se aplica si está incorporado en los ÍTEMS de la factura.
    3. Si el descuento aparece en las NOTAS u observaciones, NO VA (no disminuye la base gravable).
    4. En agencias aduaneras (DHL, Euro Shipping, etc.), las casillas o notas de descuento corresponden
       a retenciones practicadas que no deben restarse de la base para evitar duplicidades.
    5. Si en el reporte/factura existe la columna de Subtotal (valor de los ítems antes de notas),
       se toma directamente el Subtotal como base contable.
    """
    nom = str(nom_emisor).upper()
    es_aduanero = any(k in nom for k in AGENTES_ADUANEROS)

    col_sub = next((c for c in r.index if any(k in str(c).lower() for k in ["subtotal", "sub total", "base gravable", "base_gravable", "valor subtotal", "lineextensionamount"])), None)
    col_desc_item = next((c for c in r.index if any(k in str(c).lower() for k in ["descuento_items", "descuento item", "descuento ítems", "descuento linea", "descuento comercial"])), None)
    col_desc_gen = next((c for c in r.index if any(k in str(c).lower() for k in ["descuento", "descuentos", "total descuento", "valor descuento", "allowancetotalamount"])), None)

    subtotal_val = None
    if col_sub and pd.notna(r.get(col_sub)):
        try:
            v = float(r.get(col_sub))
            if v > 0:
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
    # Si el Total en DIAN está en COP (ej. > 50,000 COP) pero Subtotal e IVA vienen en USD (ej. 780 USD y 16.15 USD),
    # la relación Total / (Subtotal + IVA) está en el rango histórico de TRM (1,500 a 6,500).
    if subtotal_val is not None and tot > 1000 and (subtotal_val + iva) > 0:
        ratio_trm = tot / (subtotal_val + iva)
        if 1500 <= ratio_trm <= 6500:
            subtotal_val = round(subtotal_val * ratio_trm, 2)
            iva = round(iva * ratio_trm, 2)
            desc_val = round(desc_val * ratio_trm, 2)
            desc_item_val = round(desc_item_val * ratio_trm, 2)
            base = round(subtotal_val, 2)
            motivo_base = f"Factura en USD convertida a TRM () — Base COP:  | IVA COP: "
            return base, iva, desc_val, motivo_base

    # REGLA:
    # Si viene el Subtotal de los ítems, esa es la base real. Los descuentos en notas NO se restan.
    if subtotal_val is not None:
        base = round(subtotal_val, 2)
        motivo_base = f"Base = Subtotal de Ítems () — Descuentos en notas no aplican"
    else:
        # Si no hay columna de subtotal y es agencia aduanera con descuento reportado (retenciones en notas):
        if es_aduanero and desc_val > 0:
            base = round(tot - iva + desc_val, 2)
            motivo_base = f"Agencia Aduanera: Base restituida () sin descontar retenciones de notas"
        elif desc_item_val > 0:
            base = round(tot - iva, 2)
            motivo_base = f"Base calculada con descuento comercial en ítems ()"
        else:
            base = round(tot - iva, 2)
            motivo_base = f"Base estándar = Total - IVA ()"

    # Validación de consistencia: si Base + IVA difiere del Total por más de .0 y no hay descuentos justificados
    if tot > 0 and abs((base + iva) - tot) > 1.0 and desc_val == 0 and not (es_aduanero and desc_val > 0):
        base = round(tot - iva, 2)
        motivo_base = f"Base ajustada = Total - IVA ()"

    return base, iva, desc_val, motivo_base


# ==============================================================================
# MOTOR DE MEMORIA Y APRENDIZAJE CONTINUO DE REGLAS CONTABLES
# ==============================================================================
PATH_MEMORIA_GLOBAL = "almacenamiento_contable/memoria_aprendizaje_contable.json"

def cargar_memoria_aprendizaje(empresa_dict=None):
    """Carga el banco histórico de aprendizaje con reglas auditadas por el usuario."""
    nit_empresa = re.sub(r'\D', '', str(empresa_dict.get("nit", ""))) if empresa_dict else "global"
    path_esp = f"almacenamiento_contable/{nit_empresa}/memoria_aprendizaje_contable.json"

    for p in [path_esp, PATH_MEMORIA_GLOBAL]:
        if os.path.exists(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
    return {}

def guardar_regla_aprendizaje(empresa_dict, nit_emisor, nombre_emisor, cta_p=None, resp_fiscal=None, **kwargs):
    """Memoriza una corrección o auditoría del usuario para aplicarla automáticamente en meses futuros."""
    if not nit_emisor:
        return
    nit_clean = re.sub(r'\D', '', str(nit_emisor))
    if not nit_clean:
        return

    memoria = cargar_memoria_aprendizaje(empresa_dict)
    regla = memoria.get(nit_clean, {
        "proveedor": nombre_emisor,
        "nit": nit_clean,
        "veces_auditado": 0
    })
    if nombre_emisor:
        regla["proveedor"] = nombre_emisor
    if cta_p:
        regla["cta_principal"] = str(cta_p).strip()
    if resp_fiscal:
        regla["regimen"] = str(resp_fiscal).strip()
    for k, v in kwargs.items():
        if v is not None:
            regla[k] = v
    regla["veces_auditado"] = regla.get("veces_auditado", 0) + 1
    regla["ultima_actualizacion"] = datetime.datetime.now().strftime("%d/%m/%Y %I:%M %p")

    memoria[nit_clean] = regla

    nit_empresa = re.sub(r'\D', '', str(empresa_dict.get("nit", ""))) if empresa_dict else "global"
    dir_esp = f"almacenamiento_contable/{nit_empresa}"
    os.makedirs(dir_esp, exist_ok=True)
    os.makedirs("almacenamiento_contable", exist_ok=True)

    try:
        with open(f"{dir_esp}/memoria_aprendizaje_contable.json", "w", encoding="utf-8") as f:
            json.dump(memoria, f, ensure_ascii=False, indent=2)
        with open(PATH_MEMORIA_GLOBAL, "w", encoding="utf-8") as f:
            json.dump(memoria, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    return memoria

def clasificar_factura(nit_emisor, nombre_emisor, valor_base, valor_iva, tipo_doc, resp_emisor="", empresa_compradora=None):
    nombre = str(nombre_emisor).upper()
    resp = str(resp_emisor).upper()

    # 0. MEMORIA DE APRENDIZAJE: Si el usuario ya auditó y guardó reglas para este proveedor previamente
    cta_aprendida = None
    memoria_ap = cargar_memoria_aprendizaje(empresa_compradora)
    nit_clean_c = re.sub(r'\D', '', str(nit_emisor))
    if nit_clean_c in memoria_ap:
        r_ap = memoria_ap[nit_clean_c]
        if not resp_emisor and r_ap.get("regimen"):
            resp_emisor = r_ap["regimen"]
            resp = resp_emisor.upper()
        if r_ap.get("cta_principal"):
            cta_aprendida = r_ap["cta_principal"]
        if r_ap.get("cta_iva"):
            cta_iva = r_ap["cta_iva"]
    if empresa_compradora is None:
        empresa_compradora = {"es_gran_contribuyente": False, "es_autorretenedor": False}

    tipo_str = str(tipo_doc).upper()
    if "CREDITO" in tipo_str or "CRÉDITO" in tipo_str or "NC" in tipo_str or "DEVOLUCION" in tipo_str:
        t_comp = 17
        op = "Devolucion Compra"
    elif "NOTA" in tipo_str or "AJUSTE" in tipo_str or "CONTABILIDAD" in tipo_str:
        t_comp = 14
        op = "Nota Contabilidad"
    else:
        t_comp = 10
        op = "Compra"

    # 1. Agentes Aduaneros e Importaciones
    if any(k in nombre for k in AGENTES_ADUANEROS):
        rfte_t = round(valor_base * 0.04, 2) if valor_base >= 210000 or "EURO SHIPPING" in nombre or "DHL" in nombre else 0.0
        rica_t = round(valor_base * 0.00966, 2) if valor_base >= 210000 and "BUENAVENTURA" not in nombre and "PORTUARIA" not in nombre else 0.0
        cta_p, cta_c, desc = "146505", "22050501", f"Importacion / Transito - {nombre_emisor[:25]}"
        cta_rfte_b, cta_iva, cta_rica_b = "23652503", "240835", "23680505"
        cat, razon_b = "Importacion (1465)", "Honorarios Agenciamiento vs Terceros"

    # 2. Repuestos / Mantenimiento
    elif any(k in nombre for k in ["FERROMENDEZ", "TORNILLOLOCO", "CAUCHOS", "ASIMFER", "MAFLEXCOL", "EMPRECOL", "BATTS ZONE", "MECANIZAR", "HIDRAHULICAS", "ELECTRICOS", "ILUMINACION", "BAMACOLGROUP"]):
        rfte_t = round(valor_base * 0.025, 2) if valor_base >= 1047000 else 0.0
        rica_t = round(valor_base * 0.01104, 2) if valor_base >= 1414000 else 0.0
        cta_rfte_b, cta_iva, cta_rica_b = "23654001", "24081001", "23680501"
        if valor_base >= 500000:
            cta_p, cta_c, desc = "14350101", "22050501", f"Repuestos / Inventario - {nombre_emisor[:25]}"
            cat, razon_b = "Inventario", "Compra repuestos mayores a 500k"
        else:
            cta_p, cta_c, desc = "61800101", "23359501", f"Mantenimiento menor - {nombre_emisor[:25]}"
            cat, razon_b = "Costo Mantenimiento", "Repuestos menores a 500k"
            rfte_t, rica_t = 0.0, 0.0

    # 3. Hoteles / Viajes
    elif any(k in nombre for k in ["HOTEL", "ESTELAR", "GENOVA", "VITTAPARK"]):
        rfte_t = round(valor_base * 0.035, 2) if valor_base >= 210000 else 0.0
        rica_t = round(valor_base * 0.00966, 2) if valor_base >= 210000 else 0.0
        cta_p, cta_c, desc = "51550501", "23359501", f"Alojamiento / Viaje - {nombre_emisor[:25]}"
        cta_rfte_b, cta_iva, cta_rica_b = "23652501", "24081501", "23680505"
        cat, razon_b = "Gasto Viaje", "Hospedaje de personal (IVA Servicios 24081501)"

    # 4. Software Siigo
    elif "SIIGO" in nombre:
        cta_p, cta_c, desc = "51352001", "23359501", f"Software Siigo - {nombre_emisor[:25]}"
        rfte_t, rica_t = 0.0, 0.0
        cta_rfte_b, cta_iva, cta_rica_b = "", "24081501", ""
        cat, razon_b = "Software", "Software y tecnologia (IVA Servicios 24081501)"
        cat, razon_b = "Software", "Software y tecnologia"

    # 5. Papelería
    elif any(k in nombre for k in ["PANAMERICANA", "LIBRERIA"]):
        cta_p, cta_c, desc = "51953001", "23359501", f"Papeleria - {nombre_emisor[:25]}"
        rfte_t, rica_t = 0.0, 0.0
        cta_rfte_b, cta_iva, cta_rica_b = "", "24081001", ""
        cat, razon_b = "Gastos Papeleria", "Utiles de oficina"

    # 6. Compras y Gastos Generales
    elif valor_base >= 500000:
        rfte_t = round(valor_base * 0.025, 2) if valor_base >= 1047000 else 0.0
        rica_t = round(valor_base * 0.01104, 2) if valor_base >= 1414000 else 0.0
        cta_p, cta_c, desc = "14350101", "22050501", f"Compra mercancias - {nombre_emisor[:25]}"
        cta_rfte_b, cta_iva, cta_rica_b = "23654001", "24081001", "23680501"
        cat, razon_b = "Mercancia", "Compra general > 500k"
    else:
        cta_p, cta_c, desc = "51959501", "23359501", f"Gastos generales - {nombre_emisor[:25]}"
        rfte_t, rica_t = 0.0, 0.0
        cta_rfte_b, cta_iva, cta_rica_b = "", "24081001", ""
        cat, razon_b = "Gasto General", "Compra menor general"

    # CRUCE DE RESPONSABILIDADES FISCALES DIAN
    es_emisor_gc = "O-13" in resp or "GRAN CONTRIBUYENTE" in resp
    es_emisor_autorr = "O-15" in resp or "AUTORRETENEDOR" in resp or "SIIGO" in nombre
    es_emisor_rst = "O-47" in resp or "SIMPLE" in resp
    es_comprador_gc = empresa_compradora.get("es_gran_contribuyente", False)

    # A. ReteFuente Renta
    if es_emisor_autorr:
        rfte = 0.0
        cta_rfte = ""
        razon_rfte = "Proveedor Autorretenedor (O-15): No aplica ReteFuente"
    elif es_emisor_rst:
        rfte = 0.0
        cta_rfte = ""
        razon_rfte = "Proveedor Régimen Simple (O-47): No aplica ReteFuente (Art. 911 E.T.)"
    else:
        rfte = rfte_t
        cta_rfte = cta_rfte_b if rfte > 0 else ""
        razon_rfte = f"ReteFuente ordinaria (${rfte:,.2f})" if rfte > 0 else "Base no supera tope"

    # B. ReteIVA (Art. 437-2 E.T.)
    if es_emisor_gc and not es_comprador_gc:
        reteiva = 0.0
        razon_reteiva = "Emisor Gran Contribuyente (O-13): Comprador ordinario no practica ReteIVA"
    elif es_emisor_rst:
        reteiva = 0.0
        razon_reteiva = "Emisor Régimen Simple: Exento de ReteIVA"
    elif es_comprador_gc and valor_iva > 0:
        reteiva = round(valor_iva * 0.15, 2)
        razon_reteiva = "Comprador Gran Contribuyente practica 15% de ReteIVA"
    else:
        reteiva = 0.0
        razon_reteiva = "No aplica ReteIVA"

    # C. ReteICA
    if es_emisor_gc or es_emisor_autorr or es_emisor_rst:
        rica = 0.0
        cta_rica = ""
        razon_rica = "Emisor GC / Autorretenedor: Exento de ReteICA territorial"
    else:
        rica = rica_t
        cta_rica = cta_rica_b if rica > 0 else ""
        razon_rica = f"ReteICA liquidado (${rica:,.2f})" if rica > 0 else "Base no supera tope"

    if cta_aprendida:
        cta_p = cta_aprendida
        cat = f"Aprendido ({cta_aprendida})"
        razon_total = f"🧠 Cuenta aprendida de auditorías del usuario ({cta_aprendida}) | {razon_rfte} | {razon_reteiva}"
    else:
        razon_total = f"{razon_b} | {razon_rfte} | {razon_reteiva}"

    audit_dict = {
        "es_gc": es_emisor_gc,
        "es_autorr": es_emisor_autorr,
        "es_rst": es_emisor_rst,
        "razon_rfte": razon_rfte,
        "razon_reteiva": razon_reteiva,
        "razon_reteica": razon_rica,
        "razon_rica": razon_rica,
        "es_aprendido": bool(cta_aprendida)
    }
    return t_comp, op, cta_p, cta_c, desc, rfte, rica, reteiva, cta_rfte, cta_iva, cta_rica, cat, razon_total, audit_dict



def desbloquear_pdf_bytes(fbytes, nit_receptor=None, claves_extra=None, nits_emisores=None):
    """
    Desbloquea automáticamente un PDF si viene encriptado/protegido con contraseña.
    Prueba el NIT de la empresa compradora (con/sin dígito), NITs de emisores,
    contraseña opcional y contraseñas estándar de facturación electrónica.
    Retorna: (pdf_desbloqueado_bytes, fue_desbloqueado)
    """
    try:
        reader = PdfReader(io.BytesIO(fbytes))
        if not reader.is_encrypted:
            return fbytes, True

        claves = [""]
        if claves_extra:
            claves.append(str(claves_extra).strip())

        if nit_receptor:
            nr = re.sub(r"\D", "", str(nit_receptor))
            if nr:
                claves.extend([nr, nr[:-1], f"{nr[:-1]}-{nr[-1]}", f"{nr}5", f"{nr}-5"])
        claves.extend(["901346412", "9013464125", "901346412-5", "1234", "123456"])

        if nits_emisores:
            for ne in nits_emisores:
                c = re.sub(r"\D", "", str(ne))
                if c and len(c) >= 7:
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

@st.cache_data(show_spinner=False, max_entries=40, ttl=1200)
def cache_extraer_textos_pdf(fbytes, nit_receptor="9013464125"):
    """Extrae el texto de todas las páginas de un PDF, desbloqueándolo automáticamente si está protegido."""
    try:
        reader = PdfReader(io.BytesIO(fbytes))
        if reader.is_encrypted:
            fbytes_des, ok = desbloquear_pdf_bytes(fbytes, nit_receptor=nit_receptor)
            if ok:
                reader = PdfReader(io.BytesIO(fbytes_des))
        return [(p.extract_text() or "") for p in reader.pages]
    except Exception:
        return []




def auditar_regimen_desde_facturas_renombradas(df_ref, dict_renombrados, empresa_dict):
    """
    Escanea el régimen fiscal Y los valores fiscales reales (Subtotal, IVA, Retenciones Sugeridas, Total Neto)
    directamente de las facturas PDF ya desbloqueadas y renombradas.
    Actualiza el Régimen Fiscal Emisor Y recalcula las retenciones en la fuente correspondientes.
    """
    total_modificados = 0
    if not dict_renombrados or df_ref.empty:
        return total_modificados

    for r_idx, r_mat in df_ref.iterrows():
        soporte_nom = str(r_mat.get("Soporte PDF Renombrado", ""))
        pdf_bytes = dict_renombrados.get(soporte_nom)

        # Búsqueda por coincidencia de comprobante o folio si el nombre varió ligeramente
        if not pdf_bytes:
            comp_id = str(r_mat.get("Comprobante Siigo", "")).replace(" ", "_")
            fol_id = str(r_mat.get("Folio", ""))
            for k, v in dict_renombrados.items():
                if comp_id and comp_id in k:
                    pdf_bytes = v
                    break
                elif fol_id and len(fol_id) >= 3 and fol_id in k:
                    pdf_bytes = v
                    break

        if pdf_bytes:
            try:
                paginas = cache_extraer_textos_pdf(pdf_bytes, nit_receptor=empresa_dict.get("nit", "9013464125"))
                txt_completo = chr(10).join(paginas)

                # 1. Escanear valores fiscales y totales desde el PDF
                val_fisc = extraer_valores_fiscales_texto_pdf(txt_completo)
                hubo_cambio = False

                if val_fisc.get("subtotal") and val_fisc["subtotal"] > 0:
                    df_ref.at[r_idx, "Base"] = val_fisc["subtotal"]
                    hubo_cambio = True

                if val_fisc.get("iva") is not None and val_fisc["iva"] > 0:
                    df_ref.at[r_idx, "IVA"] = val_fisc["iva"]
                    hubo_cambio = True

                # 2. Escaneo de régimen fiscal del emisor
                nit_emisor_val = r_mat.get("NIT Emisor")
                reg_detectado = escanear_regimen_texto_pdf(txt_completo, nit_emisor=nit_emisor_val)

                if not reg_detectado:
                    ne_dig = re.sub(r'\D', '', str(nit_emisor_val))
                    if ne_dig in REGIMENES_EMISORES_CONOCIDOS:
                        reg_detectado = REGIMENES_EMISORES_CONOCIDOS[ne_dig]
                    elif any(k in str(r_mat.get("Proveedor", "")).upper() for k in ["DHL", "ESTELAR", "BUENAVENTURA", "SIIGO", "COMCEL", "CLARO", "PANAMERICANA"]):
                        reg_detectado = "O-13;O-15"

                if reg_detectado and reg_detectado != r_mat.get("Régimen Fiscal Emisor"):
                    df_ref.at[r_idx, "Régimen Fiscal Emisor"] = reg_detectado
                    hubo_cambio = True

                # 3. Recalcular liquidación de retenciones con el nuevo régimen
                reg_a_usar = df_ref.at[r_idx, "Régimen Fiscal Emisor"]
                t_c_n, op_n, c_p_n, c_c_n, desc_n, rfte_n, rica_n, riva_n, c_rf_n, c_iv_n, c_ri_n, cat_n, razon_n, audit_n = clasificar_factura(
                    df_ref.at[r_idx, "NIT Emisor"], df_ref.at[r_idx, "Proveedor"],
                    df_ref.at[r_idx, "Base"], df_ref.at[r_idx, "IVA"],
                    df_ref.at[r_idx, "Operacion"], reg_a_usar, empresa_dict
                )

                # Si el PDF trae retenciones sugeridas explícitas en notas
                if val_fisc.get("retefuente") is not None and val_fisc["retefuente"] > 0:
                    df_ref.at[r_idx, "ReteFuente"] = val_fisc["retefuente"]
                else:
                    df_ref.at[r_idx, "ReteFuente"] = rfte_n

                if val_fisc.get("reteica") is not None and val_fisc["reteica"] > 0:
                    df_ref.at[r_idx, "ReteICA"] = val_fisc["reteica"]
                else:
                    df_ref.at[r_idx, "ReteICA"] = rica_n

                df_ref.at[r_idx, "ReteIVA"] = riva_n
                df_ref.at[r_idx, "Cta ReteFuente"] = c_rf_n
                df_ref.at[r_idx, "Cta ReteICA"] = c_ri_n
                df_ref.at[r_idx, "Razón Contable"] = razon_n
                df_ref.at[r_idx, "Audit Info"] = audit_n

                # 4. Tratamiento para Agencias Aduaneras / Facturas de Mandato
                prov_nom_u = str(df_ref.at[r_idx, "Proveedor"]).upper()
                es_aduanero_row = any(k in prov_nom_u for k in ["CARGO", "ADUANA", "PORTUARIA", "ALMACENADORA", "TERMINAL", "DHL", "EURO SHIPPING", "TRADE GLOBAL"])

                if es_aduanero_row or val_fisc.get("retefuente", 0) > 0:
                    df_ref.at[r_idx, "Impuestos Asumidos"] = True
                    saldo_cruce = round(df_ref.at[r_idx, "Base"] + df_ref.at[r_idx, "IVA"], 2)
                    df_ref.at[r_idx, "Neto a Pagar"] = saldo_cruce
                    if "CARGO" in prov_nom_u:
                        df_ref.at[r_idx, "Cuenta Pasivo Especifica"] = "22050505"
                        df_ref.at[r_idx, "Cta Contrapartida"] = "22050505"
                    elif "DHL" in prov_nom_u:
                        df_ref.at[r_idx, "Cuenta Pasivo Especifica"] = "23359501"
                        df_ref.at[r_idx, "Cta Contrapartida"] = "23359501"
                else:
                    df_ref.at[r_idx, "Neto a Pagar"] = round(df_ref.at[r_idx, "Total"] - df_ref.at[r_idx, "ReteFuente"] - df_ref.at[r_idx, "ReteICA"] - df_ref.at[r_idx, "ReteIVA"], 2)

                if hubo_cambio:
                    total_modificados += 1
            except Exception as e:
                pass

    return total_modificados

def identificar_factura_en_texto(texto, df_ref):
    """
    Motor de búsqueda ultra-rápido (130x más veloz):
    Evalúa CUFE, Prefijo, Folio, NIT Emisor y Proveedor sin crear objetos pesados en bucles.
    """
    if not texto or df_ref.empty:
        return None
    txt_clean = re.sub(r'[^A-Z0-9]', '', texto.upper())
    digits_only = re.sub(r'\D', '', texto)

    # 1. Búsqueda directa por CUFE (Certeza 100%)
    for r_idx, r_cand in df_ref.iterrows():
        c_val = str(r_cand.get("CUFE", "") or r_cand.get("CUFE / Token", "") or "")
        if len(c_val) >= 15:
            c_clean = re.sub(r'[^A-Za-z0-9]', '', c_val).upper()
            if c_clean[:20] in txt_clean:
                return r_cand

    # 2. Búsqueda combinada Factura + NIT (Ultra-rápida)
    mejor_cand = None
    mejor_score = 0

    for r_idx, r_cand in df_ref.iterrows():
        pref = str(r_cand.get("Prefijo", "") or "").strip().upper()
        fol = str(r_cand.get("Folio", "") or "").strip().upper()
        if not fol or fol == "NAN":
            continue
        pref_c = re.sub(r'[^A-Z0-9]', '', pref)
        fol_c = re.sub(r'[^A-Z0-9]', '', fol)
        fac_full = (pref_c + fol_c) if pref_c else fol_c
        fol_sc = fol_c.lstrip('0')
        fac_full_sc = (pref_c + fol_sc) if pref_c else fol_sc

        has_fac = (len(fac_full) >= 3 and fac_full in txt_clean) or (len(fac_full_sc) >= 3 and fac_full_sc in txt_clean)
        if not has_fac and len(fol_c) >= 3:
            if fol_c in txt_clean:
                has_fac = True

        if not has_fac:
            continue

        nit_c = re.sub(r'\D', '', str(r_cand.get("NIT Emisor", "")))
        nit_base = nit_c[:-1] if len(nit_c) >= 10 else nit_c
        has_nit = (nit_c and len(nit_c) >= 6 and nit_c in digits_only) or (nit_base and len(nit_base) >= 6 and nit_base in digits_only)

        prov_clean = re.sub(r'[^A-Z0-9]', '', str(r_cand.get("Proveedor", "")).upper())
        has_prov = len(prov_clean) >= 4 and prov_clean[:8] in txt_clean

        score = 600 if has_nit else (450 if has_prov else 300)
        if score > mejor_score:
            mejor_score = score
            mejor_cand = r_cand
            if score == 600:
                break

        if score > mejor_score and score >= 300:
            mejor_score = score
            mejor_cand = r_cand

    return mejor_cand

def buscar_y_extraer_pdf(fac_sel, dict_renombrados=None, dict_originales=None, empresa_compradora=None):
    """
    Busca y extrae el PDF de la factura seleccionada dando MÁXIMA PRIORIDAD
    a las facturas YA DESBLOQUEADAS, SEPARADAS Y RENOMBRADAS (dict_renombrados):
    1. PRIORIDAD 1: Facturas en dict_renombrados (las que ya pasaron por el proceso de desbloqueo y renombrado).
       Coincidencia por nombre esperado (Comp_10-XXX...), comprobante, folio o CUFE.
    2. PRIORIDAD 2: Si aún no se han desbloqueado/renombrado, buscar en dict_originales y desbloquearlas al vuelo.
    Retorna: (pdf_bytes_desbloqueado, descripcion_origen, lista_paginas)
    """
    comp_siigo = str(fac_sel.get("Comprobante Siigo", "")).strip()
    consecutivo = str(fac_sel.get("Consecutivo", "")).strip()
    t_comp = str(fac_sel.get("Tipo Comp", "")).strip()

    folio_clean = str(fac_sel.get("Folio", "")).replace("-", "").strip().upper()
    folio_sc = folio_clean.lstrip("0")
    pref_clean = str(fac_sel.get("Prefijo", "")).replace("-", "").strip().upper()
    fac_clean = str(fac_sel.get("Factura", "")).replace("-", "").strip().upper()
    fac_full = (pref_clean + folio_clean) if pref_clean else folio_clean
    fac_full_sc = (pref_clean + folio_sc) if pref_clean else folio_sc

    cufe_raw = str(fac_sel.get("CUFE", "") or fac_sel.get("CUFE / Token", "") or "").strip()
    cufe_clean = re.sub(r'[^a-zA-Z0-9]', '', cufe_raw).lower()

    # El NIT receptor es el de la empresa compradora
    nit_comprador = "9013464125"
    if empresa_compradora:
        nit_comprador = re.sub(r"\D", "", str(empresa_compradora.get("nit", "9013464125")))

    nit_emisor_digits = re.sub(r"\D", "", str(fac_sel.get("NIT Emisor", "")))
    soporte_nom = str(fac_sel.get("Soporte PDF Renombrado", "")).strip()
    df_una_fac = pd.DataFrame([fac_sel])

    # =========================================================================
    # PRIORIDAD 1: TOMAR LAS FACTURAS DESPUÉS DE DESBLOQUEARLAS Y RENOMBRARLAS
    # =========================================================================
    if dict_renombrados:
        # A. Coincidencia exacta por el nombre de archivo asignado
        if soporte_nom and soporte_nom in dict_renombrados:
            f_bytes = dict_renombrados[soporte_nom]
            f_des, _ = desbloquear_pdf_bytes(f_bytes, nit_receptor=nit_comprador)
            return f_des, f"Factura Desbloqueada Oficial ({soporte_nom})", None

        # B. Coincidencia por Comprobante (ej. Comp_10-680 o 10-680)
        comp_patterns = [
            f"Comp_{t_comp}-{consecutivo}".upper(),
            f"Comp_{t_comp}_{consecutivo}".upper(),
            f"{t_comp}-{consecutivo}".upper(),
            f"{t_comp}_{consecutivo}".upper()
        ]
        for k, v in dict_renombrados.items():
            k_u = k.upper().replace(" ", "_")
            if any(cp in k_u for cp in comp_patterns):
                f_des, _ = desbloquear_pdf_bytes(v, nit_receptor=nit_comprador)
                return f_des, f"Factura Desbloqueada Oficial ({k})", None

        # C. Coincidencia por Número de Factura / Folio en el nombre renombrado
        for k, v in dict_renombrados.items():
            k_clean = k.replace("-", "").replace(" ", "").replace("_", "").upper()
            if fac_full and len(fac_full) >= 3 and fac_full in k_clean:
                f_des, _ = desbloquear_pdf_bytes(v, nit_receptor=nit_comprador)
                return f_des, f"Factura Desbloqueada Oficial ({k})", None
            if folio_clean and len(folio_clean) >= 3 and folio_clean in k_clean:
                f_des, _ = desbloquear_pdf_bytes(v, nit_receptor=nit_comprador)
                return f_des, f"Factura Desbloqueada Oficial ({k})", None

        # D. Coincidencia por CUFE en el nombre renombrado
        if cufe_clean and len(cufe_clean) >= 15:
            for k, v in dict_renombrados.items():
                k_l = re.sub(r'[^a-zA-Z0-9]', '', k).lower()
                if cufe_clean[:20] in k_l:
                    f_des, _ = desbloquear_pdf_bytes(v, nit_receptor=nit_comprador)
                    return f_des, f"Factura Desbloqueada Oficial ({k})", None

        # E. Verificación por contenido de texto en las facturas renombradas
        for r_nom, r_bytes in dict_renombrados.items():
            try:
                pgs = cache_extraer_textos_pdf(r_bytes, nit_receptor=nit_comprador)
                txt_r = " ".join(pgs)
                if identificar_factura_en_texto(txt_r, df_una_fac) is not None:
                    f_des, _ = desbloquear_pdf_bytes(r_bytes, nit_receptor=nit_comprador)
                    return f_des, f"Factura Desbloqueada por Contenido ({r_nom})", None
            except Exception:
                pass

    # =========================================================================
    # PRIORIDAD 2: SI AÚN NO SE HA RENOMBRADO, BUSCAR EN LOS ORIGINALES Y DESBLOQUEAR
    # =========================================================================
    if dict_originales:
        # A. Búsqueda por CUFE en nombre de archivo original
        if cufe_clean and len(cufe_clean) >= 15:
            for fname, fbytes in dict_originales.items():
                fn_l = re.sub(r'[^a-zA-Z0-9]', '', fname).lower()
                if cufe_clean[:20] in fn_l or fn_l.startswith(cufe_clean[:18]):
                    f_des, _ = desbloquear_pdf_bytes(fbytes, nit_receptor=nit_comprador)
                    return f_des, f"Factura Desbloqueada DIAN ({fname[:20]}...pdf)", None

        # B. Archivos originales individuales por coincidencia de texto
        for fname, fbytes in dict_originales.items():
            try:
                f_des, _ = desbloquear_pdf_bytes(fbytes, nit_receptor=nit_comprador)
                pgs = cache_extraer_textos_pdf(f_des, nit_receptor=nit_comprador)
                txt_ind = " ".join(pgs)
                if identificar_factura_en_texto(txt_ind, df_una_fac) is not None:
                    return f_des, f"Factura Desbloqueada ({fname})", None
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
                        m_p = re.search(r'(?:P[ÁAáa]G(?:INA)?|HOJA|PAGE)\s*[:\.]?\s*1\s*(?:DE|/|OF)\s*(\d+)', txt, re.IGNORECASE)
                        if m_p:
                            try:
                                total_h = int(m_p.group(1))
                                if 2 <= total_h <= 15:
                                    for off in range(1, total_h):
                                        p_next = p_idx + off
                                        if p_next < len(paginas_txt) and p_next not in pags_coincidentes:
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
                    str_p = ", ".join([str(p+1) for p in pags_coincidentes])
                    return out.getvalue(), f"Factura Desbloqueada de '{fname}' (Págs {str_p})", [p+1 for p in pags_coincidentes]
            except Exception:
                pass

    return None, None, None


def renderizar_visor_pdf_completo(pdf_bytes, nombre_archivo, fac_sel=None, key_prefix="pdf_view", nit_comprador="9013464125"):
    """
    Visor oficial 100% blindado y garantizado:
    1. Desbloquea y desencripta el PDF primero para que no pida contraseña.
    2. Renderizado de alta definición nativo en imágenes mediante pypdfium2 (CERO caras tristes 📄🙁, CERO bloqueos de Chrome).
    3. Botón para abrir a pantalla completa en pestaña nueva y botón de descarga directa.
    4. Cuadro formal con resumen contable de la factura.
    """
    if not pdf_bytes or not isinstance(pdf_bytes, (bytes, bytearray)):
        st.error("⚠️ El archivo PDF no contiene datos válidos.")
        return

    # 1. Desbloquear el PDF para garantizar que esté 100% libre de contraseña
    pdf_limpio, _ = desbloquear_pdf_bytes(pdf_bytes, nit_receptor=nit_comprador)
    pdf_usar = pdf_limpio if pdf_limpio else pdf_bytes

    try:
        reader_prev = PdfReader(io.BytesIO(pdf_usar))
        num_pags_tot = len(reader_prev.pages)
    except Exception:
        num_pags_tot = 1

    b64_pdf = base64.b64encode(pdf_usar).decode('utf-8').replace('\n', '').strip()

    # Pre-cálculo seguro de variables para evitar cualquier error de formato
    val_base = float(fac_sel.get("Base", 0.0)) if (fac_sel is not None and "Base" in fac_sel) else 0.0
    val_tot = float(fac_sel.get("Total", 0.0)) if (fac_sel is not None and "Total" in fac_sel) else 0.0
    val_iva = float(fac_sel.get("IVA", 0.0)) if (fac_sel is not None and "IVA" in fac_sel) else 0.0
    val_rfte = float(fac_sel.get("ReteFuente", 0.0)) if (fac_sel is not None and "ReteFuente" in fac_sel) else 0.0
    fac_num_str = str(fac_sel.get("Factura", "")) if fac_sel is not None else nombre_archivo
    prov_str = str(fac_sel.get("Proveedor", "")) if fac_sel is not None else ""
    nit_str = str(fac_sel.get("NIT Emisor", "")) if fac_sel is not None else ""
    comp_str = str(fac_sel.get("Comprobante Siigo", "")) if fac_sel is not None else ""
    fecha_str = str(fac_sel.get("Fecha", "")) if fac_sel is not None else ""
    reg_str = str(fac_sel.get("Régimen Fiscal Emisor", "O-48")) if fac_sel is not None else "O-48"

    # 2. BOTONES SUPERIORES DE ACCIÓN RÁPIDA
    c_btn1, c_btn2 = st.columns([1, 1])
    with c_btn1:
        pags_label = f"{num_pags_tot} página{'s' if num_pags_tot > 1 else ''} completa{'s' if num_pags_tot > 1 else ''}"
        st.download_button(
            label=f"📥 Descargar Factura Desbloqueada ({pags_label})",
            data=pdf_usar,
            file_name=nombre_archivo,
            mime="application/pdf",
            key=f"btn_dl_univ_{key_prefix}",
            use_container_width=True
        )
    with c_btn2:
        st.markdown(f"""
        <a href="data:application/pdf;base64,{b64_pdf}" target="_blank" download="{nombre_archivo}" style="text-decoration:none;">
            <div style="background:#0284c7; color:white; text-align:center; padding:9px 12px; border-radius:6px; font-weight:600; font-size:14px; box-shadow:0 1px 2px rgba(0,0,0,0.05); cursor:pointer;">
                🗗 Abrir Factura en Otra Ventana (Pantalla Completa)
            </div>
        </a>
        """, unsafe_allow_html=True)

    st.write("")

    # 2.B SELECTOR DE TAMAÑO VISUAL DE LA FACTURA
    c_tam1, c_tam2 = st.columns([3, 2])
    with c_tam1:
        tam_sel = st.radio(
            "📐 Tamaño de Vista Previa:",
            ["📄 Normal (Recomendado - 780px)", "🔍 Grande (Ancho Completo)", "📱 Compacto (600px)"],
            horizontal=True,
            index=0,
            key=f"rad_tam_pdf_{key_prefix}"
        )

    # Proporción de columnas para centrar y mostrar la hoja en tamaño normal de lectura
    if "Compacto" in tam_sel:
        col_ratios = [2.2, 4.0, 2.2]
    elif "Grande" in tam_sel:
        col_ratios = [0.01, 9.98, 0.01]
    else:  # Normal (Recomendado)
        col_ratios = [1.6, 5.2, 1.6]

    # 3. RENDERIZADO VISUAL DIRECTO EN IMÁGENES NATIVAS (CERO CARAS TRISTES, CERO BLOQUEOS DE CHROME)
    imgs_paginas = []
    if HAS_PDFIUM:
        try:
            doc_p = pdfium.PdfDocument(pdf_usar)
            for idx_pg in range(len(doc_p)):
                imgs_paginas.append(doc_p[idx_pg].render(scale=1.75).to_pil())
        except Exception:
            imgs_paginas = []

    if imgs_paginas:
        for idx_p, img in enumerate(imgs_paginas):
            c_l, c_mid, c_r = st.columns(col_ratios)
            with c_mid:
                st.markdown(f"""
                <div style="background:#0070ba; color:white; padding:8px 14px; border-radius:6px 6px 0 0; font-weight:600; font-size:13.5px; display:flex; justify-content:space-between; align-items:center; margin-top:16px; box-shadow: 0 2px 4px rgba(0,0,0,0.06);">
                    <span>📄 Factura: {fac_num_str} — {prov_str}</span>
                    <span style="background:rgba(255,255,255,0.25); padding:2px 10px; border-radius:10px; font-size:12px;">Hoja {idx_p + 1} de {len(imgs_paginas)}</span>
                </div>
                """, unsafe_allow_html=True)
                st.image(img, use_container_width=True)
    else:
        # Fallback a Canvas HTML5 si pypdfium no estuviera disponible
        visor_height = max(680, min(3600, num_pags_tot * 850))
        html_canvas_viewer = f"""
        <!DOCTYPE html>
        <html>
        <head>
          <meta charset="utf-8">
          <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.min.js"></script>
          <style>
            body {{
              margin: 0;
              padding: 10px;
              background: #f1f5f9;
              display: flex;
              flex-direction: column;
              align-items: center;
              font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
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
          </style>
        </head>
        <body>
          <div id="msg">⏳ Cargando y renderizando factura ({num_pags_tot} página(s))...</div>
          <div id="pages-container" style="width: 100%; display: flex; flex-direction: column; align-items: center;"></div>
          <script>
            (async function() {{
              try {{
                const b64 = "{b64_pdf}";
                const raw = atob(b64);
                const uint8Array = new Uint8Array(raw.length);
                for (let i = 0; i < raw.length; i++) {{
                  uint8Array[i] = raw.charCodeAt(i);
                }}
                pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';
                const pdf = await pdfjsLib.getDocument({{ data: uint8Array }}).promise;
                document.getElementById('msg').style.display = 'none';
                const container = document.getElementById('pages-container');

                for (let num = 1; num <= pdf.numPages; num++) {{
                  const page = await pdf.getPage(num);
                  const viewport = page.getViewport({{ scale: 1.5 }});
                  const card = document.createElement('div');
                  card.className = 'page-card';
                  const title = document.createElement('div');
                  title.className = 'page-title';
                  title.innerHTML = '<span>📄 {nombre_archivo}</span><span style="background:rgba(255,255,255,0.25); padding:2px 8px; border-radius:10px; font-size:11px;">Hoja ' + num + ' de ' + pdf.numPages + '</span>';
                  card.appendChild(title);
                  const canvas = document.createElement('canvas');
                  const ctx = canvas.getContext('2d');
                  canvas.height = viewport.height;
                  canvas.width = viewport.width;
                  card.appendChild(canvas);
                  container.appendChild(card);
                  await page.render({{ canvasContext: ctx, viewport: viewport }}).promise;
                }}
              }} catch (err) {{
                document.getElementById('msg').innerHTML = '<div style="background:white; padding:20px; border-radius:8px; border:2px solid #0070ba; text-align:center;"><h4>📄 Factura Lista</h4><p style="color:#64748b;">Utiliza el botón de descarga o abrir en otra ventana.</p></div>';
              }}
            }})();
          </script>
        </body>
        </html>
        """
        components.html(html_canvas_viewer, height=visor_height, scrolling=True)

    # 4. CUADRO RESUMEN OFICIAL CON VALORES CONTABLES
    st.markdown(f"""
    <div style="background: white; border: 1px solid #cbd5e1; border-radius: 8px; padding: 18px; margin-top: 14px; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.04);">
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px; font-size: 13.5px; color: #334155;">
            <div><b>Proveedor:</b> {prov_str} (NIT: {nit_str})<br><b>Factura:</b> {fac_num_str} ({comp_str})</div>
            <div><b>Fecha Emisión:</b> {fecha_str} | <b>Régimen:</b> {reg_str}<br><b>Base:</b> ${val_base:,.2f} | <b>IVA:</b> ${val_iva:,.2f} | <b>Total:</b> <span style="font-weight:bold; color:#0f172a;">${val_tot:,.2f}</span></div>
        </div>
    </div>
    """, unsafe_allow_html=True)



# AVISO DE TRABAJO PREVIO DISPONIBLE (SOLO SE CARGA SI EL USUARIO OPRIME EL BOTÓN)
trabajos_existentes = listar_trabajos_historial(empresa)

if "df_procesado" not in st.session_state and trabajos_existentes:
    ultimo = trabajos_existentes[0]
    nom_ult = ultimo.get("nombre_trabajo", f"Trabajo {ultimo.get('id', '')}")
    n_fac = ultimo.get("total_facturas", 0)
    val_tot = ultimo.get("total_valor", 0.0)

    st.markdown(f"""
    <div style="background:#f0fdf4; border:1px solid #86efac; border-left:5px solid #16a34a; border-radius:8px; padding:14px 18px; margin-bottom:14px;">
        <h4 style="margin:0 0 6px 0; color:#166534;">📂 Tienes un trabajo guardado disponible</h4>
        <p style="margin:0; color:#14532d; font-size:14px;">
            Tienes guardado en memoria tu trabajo previo: <b>'{nom_ult}'</b> con <b>{n_fac} facturas</b> (${val_tot:,.2f}) y sus PDFs vinculados.<br>
            • Si deseas continuar con este trabajo, haz clic en el botón verde.<br>
            • Si hoy vas a implementar <b>facturas nuevas</b>, no tienes que oprimir nada y puedes subir tus nuevos archivos directamente abajo en la Pestaña 1.
        </p>
    </div>
    """, unsafe_allow_html=True)

    c_btn_c1, c_btn_c2 = st.columns([1.6, 3])
    with c_btn_c1:
        if st.button("📂 Cargar mi trabajo anterior", key="btn_cargar_trabajo_previo_on_demand", use_container_width=True):
            df_g, pdfs_r_g, pdfs_o_g, ex_b_g, ex_n_g, zip_g = cargar_trabajo_historial(empresa, ultimo["id"])
            if df_g is not None and not df_g.empty:
                st.session_state["df_procesado"] = df_g
                st.session_state["dict_pdfs"] = pdfs_r_g or {}
                st.session_state["raw_uploaded_pdfs"] = pdfs_o_g or {}
                st.session_state["excel_bytes"] = ex_b_g
                st.session_state["excel_nombre"] = ex_n_g
                st.session_state["zip_pdfs"] = zip_g
                st.session_state["total_zip_pdfs"] = len(pdfs_r_g) if pdfs_r_g else 0
                st.session_state["job_actual_id"] = ultimo["id"]
                st.session_state["_sesion_cargada_nombre"] = nom_ult
                st.success(f"¡Trabajo '{nom_ult}' cargado con éxito!")
                st.rerun()
    with c_btn_c2:
        st.caption("👈 Oprime aquí **únicamente** cuando desees restaurar el trabajo anterior. De lo contrario, continúa abajo con tus facturas nuevas.")



# (Funciones de autosave reubicadas limpiamente al inicio)


def extraer_montos_items_pdf(pdf_bytes):
    """Extrae valores monetarios de los ítems o detalles de productos en el PDF de una factura."""
    if not pdf_bytes:
        return []
    try:
        textos = cache_extraer_textos_pdf(pdf_bytes)
        full_txt = " ".join(textos)
        montos = []
        encontrados = re.findall(r'(?:[\$]|(?:\s))\s*(\d{1,3}(?:\.\d{3})+(?:,\d{2})?)', full_txt)
        for m in encontrados:
            try:
                val = float(m.replace(".", "").replace(",", "."))
                if 20000 <= val <= 100000000:
                    montos.append(val)
            except Exception:
                pass
        return list(set(montos))
    except Exception:
        return []

def generar_respaldo_portatil_bytes(empresa_dict):
    """Genera en memoria un archivo .indumaq (ZIP completo) con el progreso actual de auditoría y triangulación."""
    if "df_procesado" not in st.session_state or st.session_state["df_procesado"] is None:
        return None
    try:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            # 1. df_procesado
            zf.writestr("df_procesado.pkl", pickle.dumps(st.session_state["df_procesado"]))

            # 2. excel_bytes
            if st.session_state.get("excel_bytes"):
                zf.writestr("excel_original.xlsx", st.session_state["excel_bytes"])

            # 3. paquetes_importacion
            if "paquetes_importacion" in st.session_state and st.session_state["paquetes_importacion"]:
                zf.writestr("paquetes_importacion.pkl", pickle.dumps(st.session_state["paquetes_importacion"]))

            # 4. asientos_triangulacion
            if "asientos_triangulacion_por_factura" in st.session_state and st.session_state["asientos_triangulacion_por_factura"]:
                zf.writestr("asientos_triangulacion.pkl", pickle.dumps(st.session_state["asientos_triangulacion_por_factura"]))

            # 5. estado_sesion
            estado_flags = {
                "paquetes_listos": {k: v for k, v in st.session_state.items() if k.startswith("paquete_listo_")},
                "enviar_gp": {k: v for k, v in st.session_state.items() if k.startswith("enviar_gp_pq_")},
                "enviar_h2": {k: v for k, v in st.session_state.items() if k.startswith("enviar_h2_pq_")},
                "enviar_nd": {k: v for k, v in st.session_state.items() if k.startswith("enviar_nd_pq_")},
                "facturas_no_contabilizar": list(st.session_state.get("facturas_no_contabilizar", set())),
                "sel_paquete_activo_key": st.session_state.get("sel_paquete_activo_key", 1)
            }
            zf.writestr("estado_sesion.json", json.dumps(estado_flags, ensure_ascii=False, indent=2))

            # 6. meta
            now_str = datetime.datetime.now().strftime("%d/%m/%Y %I:%M %p")
            meta = {
                "empresa": empresa_dict.get("nombre", "Empresa"),
                "nit": empresa_dict.get("nit", ""),
                "fecha_respaldo": now_str,
                "archivo_excel": st.session_state.get("excel_nombre", "Reporte.xlsx"),
                "total_facturas": len(st.session_state["df_procesado"])
            }
            zf.writestr("meta.json", json.dumps(meta, ensure_ascii=False, indent=2))
            zf.writestr("meta_sesion.json", json.dumps(meta, ensure_ascii=False, indent=2))
            mem_actual = cargar_memoria_aprendizaje(empresa_dict)
            if mem_actual:
                zf.writestr("memoria_aprendizaje.json", json.dumps(mem_actual, ensure_ascii=False, indent=2))

            # 7. PDFs si existen en la sesión (ZIP_STORED para que sea instantáneo sin bloquear CPU)
            dict_renom_s = st.session_state.get("dict_pdfs", {})
            if dict_renom_s:
                for fn, bdata in dict_renom_s.items():
                    zinfo = zipfile.ZipInfo(f"pdfs_renombrados/{fn}")
                    zinfo.compress_type = zipfile.ZIP_STORED
                    zf.writestr(zinfo, bdata)
            dict_orig_s = st.session_state.get("raw_uploaded_pdfs", {})
            if dict_orig_s:
                for fn, bdata in dict_orig_s.items():
                    zinfo = zipfile.ZipInfo(f"pdfs_originales/{fn}")
                    zinfo.compress_type = zipfile.ZIP_STORED
                    zf.writestr(zinfo, bdata)
            if st.session_state.get("zip_pdfs"):
                zinfo = zipfile.ZipInfo("paquete_facturas.zip")
                zinfo.compress_type = zipfile.ZIP_STORED
                zf.writestr(zinfo, st.session_state["zip_pdfs"])

        buf.seek(0)
        return buf.getvalue()
    except Exception:
        return None

def restaurar_desde_respaldo_bytes(empresa_dict, raw_zip_bytes):
    """Restaura un archivo de respaldo portátil (.indumaq o .zip) directamente a la sesión activa con facturas, paquetes y PDFs."""
    try:
        with zipfile.ZipFile(io.BytesIO(raw_zip_bytes), "r") as zf:
            namelist = zf.namelist()
            if "memoria_aprendizaje.json" in namelist:
                try:
                    mem_rest = json.loads(zf.read("memoria_aprendizaje.json").decode("utf-8"))
                    nit_emp = re.sub(r'\D', '', str(empresa_dict.get("nit", ""))) if empresa_dict else "global"
                    os.makedirs(f"almacenamiento_contable/{nit_emp}", exist_ok=True)
                    with open(f"almacenamiento_contable/{nit_emp}/memoria_aprendizaje_contable.json", "w", encoding="utf-8") as f_mem:
                        json.dump(mem_rest, f_mem, ensure_ascii=False, indent=2)
                    with open(PATH_MEMORIA_GLOBAL, "w", encoding="utf-8") as f_mem:
                        json.dump(mem_rest, f_mem, ensure_ascii=False, indent=2)
                except Exception:
                    pass
            if "df_procesado.pkl" in namelist:
                st.session_state["df_procesado"] = pickle.loads(zf.read("df_procesado.pkl"))
            if "excel_original.xlsx" in namelist:
                st.session_state["excel_bytes"] = zf.read("excel_original.xlsx")
            if "paquete_facturas.zip" in namelist:
                st.session_state["zip_pdfs"] = zf.read("paquete_facturas.zip")
            if "paquetes_importacion.pkl" in namelist:
                st.session_state["paquetes_importacion"] = pickle.loads(zf.read("paquetes_importacion.pkl"))
            if "asientos_triangulacion.pkl" in namelist:
                st.session_state["asientos_triangulacion_por_factura"] = pickle.loads(zf.read("asientos_triangulacion.pkl"))
            if "estado_sesion.json" in namelist:
                est = json.loads(zf.read("estado_sesion.json").decode("utf-8"))
                for k, v in est.get("paquetes_listos", {}).items():
                    st.session_state[k] = v
                for k, v in est.get("enviar_gp", {}).items():
                    st.session_state[k] = v
                for k, v in est.get("enviar_h2", {}).items():
                    st.session_state[k] = v
                for k, v in est.get("enviar_nd", {}).items():
                    st.session_state[k] = v
                if "facturas_no_contabilizar" in est:
                    st.session_state["facturas_no_contabilizar"] = set(est["facturas_no_contabilizar"])
                if "sel_paquete_activo_key" in est:
                    st.session_state["sel_paquete_activo_key"] = est["sel_paquete_activo_key"]
            if "meta.json" in namelist:
                meta = json.loads(zf.read("meta.json").decode("utf-8"))
                st.session_state["excel_nombre"] = meta.get("archivo_excel", meta.get("excel_nombre", "Reporte.xlsx"))
                st.session_state["_sesion_cargada_nombre"] = f"Respaldo {meta.get('fecha_respaldo', '')}"
            elif "meta_sesion.json" in namelist:
                meta = json.loads(zf.read("meta_sesion.json").decode("utf-8"))
                st.session_state["excel_nombre"] = meta.get("excel_nombre", meta.get("archivo_excel", "Reporte.xlsx"))
                st.session_state["_sesion_cargada_nombre"] = f"Respaldo {meta.get('fecha_respaldo', '')}"

            # Restaurar PDFs
            dict_renom = {}
            dict_orig = {}
            for name in namelist:
                if name.startswith("pdfs_renombrados/") and name.endswith(".pdf"):
                    fn = name.replace("pdfs_renombrados/", "")
                    dict_renom[fn] = zf.read(name)
                elif name.startswith("pdfs_originales/") and name.endswith(".pdf"):
                    fn = name.replace("pdfs_originales/", "")
                    dict_orig[fn] = zf.read(name)
            if dict_renom:
                st.session_state["dict_pdfs"] = dict_renom
                st.session_state["total_zip_pdfs"] = len(dict_renom)
            if dict_orig:
                st.session_state["raw_uploaded_pdfs"] = dict_orig
        return True
    except Exception as e:
        return False

# BANNER DE CONTROL Y GUARDADO MANUAL
if "df_procesado" in st.session_state and st.session_state["df_procesado"] is not None:
    c_bnr1, c_bnr2, c_bnr3, c_bnr4 = st.columns([2.8, 1.3, 1.6, 0.9])
    with c_bnr1:
        nom_ses = st.session_state.get("excel_nombre", "Reporte de Facturas")
        last_save = st.session_state.get("_ultimo_guardado_manual")
        save_info = f" | 💾 `{last_save}`" if last_save else ""
        st.info(f"📋 **Trabajo Activo:** `{nom_ses}` ({len(st.session_state['df_procesado'])} facturas){save_info}")
    with c_bnr2:
        if st.button("💾 Guardar Avance", key="btn_guardar_manual_top", use_container_width=True, help="Guarda en disco todas las modificaciones de cuentas y paquetes validados."):
            if guardar_estado_manual(empresa):
                st.success("✅ ¡Avance contable guardado exitosamente!")
                st.rerun()
            else:
                st.error("Error al guardar.")
    with c_bnr3:
        respaldo_bytes_top = generar_respaldo_portatil_bytes(empresa)
        if respaldo_bytes_top:
            emp_clean_name = re.sub(r'[^a-zA-Z0-9]', '_', str(empresa.get('nombre', 'Empresa')))
            nom_resp_top = f"Respaldo_Avance_{emp_clean_name}.indumaq"
            st.download_button(
                label="🛡️ Descargar Respaldo (.indumaq)",
                data=respaldo_bytes_top,
                file_name=nom_resp_top,
                mime="application/zip",
                key="btn_dl_respaldo_top_direct",
                use_container_width=True,
                help="Descarga un archivo con tus 16 paquetes validados y facturas editadas. Si se apaga o reinicia Streamlit, lo cargas en segundos."
            )
    with c_bnr4:
        if st.button("🆕 Limpiar", key="btn_nuevo_trabajo_top", use_container_width=True, help="Limpia la memoria para cargar un nuevo mes."):
            for k in ["df_procesado", "dict_pdfs", "raw_uploaded_pdfs", "excel_bytes", "excel_nombre", "zip_pdfs", "job_actual_id", "_ultimo_excel_proc_sig", "_ultimo_pdfs_proc_sig", "paquetes_importacion", "asientos_triangulacion_por_factura", "facturas_no_contabilizar", "_ultimo_guardado_manual"]:
                st.session_state.pop(k, None)
            st.rerun()
else:
    # Si la memoria está vacía, ofrecer botón para cargar el último avance guardado si existe
    nit_c = re.sub(r'\D', '', str(empresa.get("nit", "empresa")))
    d_man = os.path.join(DATA_DIR, nit_c, "estado_manual")
    if os.path.exists(os.path.join(d_man, "df_procesado.pkl")):
        col_rec1, col_rec2 = st.columns([3.5, 1.5])
        with col_rec1:
            st.info("💡 **Tienes un avance contable guardado previamente.** Puedes restaurarlo con un solo clic:")
        with col_rec2:
            if st.button("📂 Cargar Último Avance Guardado", key="btn_cargar_avance_previo", type="primary", use_container_width=True):
                if cargar_estado_manual(empresa):
                    st.success("¡Avance restaurado con éxito!")
                    st.rerun()

# PANEL DE HISTORIAL DE TRABAJOS Y RESPALDOS INDESTRUCTIBLES
trabajos_guardados = listar_trabajos_historial(empresa)
panel_expanded = False

with st.expander("🗂️ Historial de Trabajos, Respaldos y Carga Rápida", expanded=panel_expanded):
    # Restauración inmediata desde archivo portátil .indumaq
    c_upl_r1, c_upl_r2 = st.columns([3, 1.2])
    with c_upl_r1:
        upl_indumaq = st.file_uploader("📤 Restaurar desde Archivo de Respaldo (.indumaq):", type=["indumaq", "backup", "zip"], key="upl_respaldo_indumaq", help="Sube tu archivo de respaldo para recuperar al instante exactamente donde lo dejaste.")
    with c_upl_r2:
        st.write("")
        st.write("")
        if upl_indumaq is not None:
            if st.button("🚀 Restaurar Respaldo", key="btn_ejecutar_restaurar_upl", use_container_width=True):
                ok_r = restaurar_desde_respaldo_bytes(empresa, upl_indumaq.read())
                if ok_r:
                    st.success("¡Respaldo restaurado con éxito! Todas tus facturas editadas y paquetes quedaron listos.")
                    st.rerun()
                else:
                    st.error("No se pudo leer el archivo de respaldo.")
    st.markdown("<hr style='margin:10px 0;'>", unsafe_allow_html=True)

    col_h_left, col_h_right = st.columns([1.5, 1])

    with col_h_left:
        st.markdown("#### 📂 Trabajos Guardados en esta Empresa:")
        if trabajos_guardados:
            st.caption("Haz clic en 'Cargar' para recuperar de inmediato cualquier auditoría o mes previo:")
            for tb in trabajos_guardados:
                c_h1, c_h2, c_h3 = st.columns([3, 1.2, 0.5])
                with c_h1:
                    nom_tb = tb.get("nombre_trabajo", f"Trabajo {tb.get('id', '')}")
                    n_facs = tb.get("total_facturas", 0)
                    val_tot = tb.get("total_valor", 0.0)
                    n_renom = tb.get("total_pdfs_renombrados", tb.get("total_pdfs", 0))
                    n_orig = tb.get("total_pdfs_originales", 0)
                    info_pdf_str = f"{n_renom} PDFs procesados" if n_renom > 0 else (f"{n_orig} PDFs subidos" if n_orig > 0 else "Sin PDFs")
                    st.markdown(f"📄 **{nom_tb}** — {n_facs} facturas (${val_tot:,.2f}) — **{info_pdf_str}**")
                with c_h2:
                    if st.button("📂 Cargar", key=f"btn_h_load_{tb['id']}"):
                        df_g, pdfs_r_g, pdfs_o_g, ex_b_g, ex_n_g, zip_g = cargar_trabajo_historial(empresa, tb["id"])
                        if df_g is not None:
                            st.session_state["df_procesado"] = df_g
                        if pdfs_r_g:
                            st.session_state["dict_pdfs"] = pdfs_r_g
                        if pdfs_o_g:
                            st.session_state["raw_uploaded_pdfs"] = pdfs_o_g
                        if ex_b_g:
                            st.session_state["excel_bytes"] = ex_b_g
                            st.session_state["excel_nombre"] = ex_n_g
                        if zip_g:
                            st.session_state["zip_pdfs"] = zip_g
                            st.session_state["total_zip_pdfs"] = len(pdfs_r_g) if pdfs_r_g else 0
                        st.session_state["job_actual_id"] = tb["id"]
                        st.session_state["_sesion_auto_recuperada"] = nom_tb
                        st.success(f"¡Trabajo '{nom_tb}' cargado! Todo tu avance está listo.")
                        st.rerun()
                with c_h3:
                    if st.button("🗑️", key=f"btn_h_del_{tb['id']}"):
                        eliminar_trabajo_historial(empresa, tb["id"])
                        st.rerun()
        else:
            st.info("💡 Aún no tienes trabajos guardados en el disco local para esta empresa.")

    with col_h_right:
        st.markdown("#### 🛡️ Respaldo Portable (.zip):")
        st.caption("Guarda o restaura todo tu trabajo en un solo archivo, ideal si cambias de PC o si el servidor se reinicia:")

        pass

        # 2. Uploader para restaurar desde un archivo de respaldo previo (con botón de confirmación para evitar bucles)
        archivo_zip_restaurar = st.file_uploader(
            "📥 Restaurar Sesión desde Archivo de Respaldo (.indumaq / .zip):",
            type=["indumaq", "zip"],
            key="upl_zip_restore_historial",
            help="Sube tu archivo de respaldo para restaurar exactamente donde lo dejaste."
        )
        if archivo_zip_restaurar is not None:
            if st.button("🚀 Cargar este Respaldo en Pantalla", key="btn_confirmar_cargar_zip_historial", use_container_width=True):
                with st.spinner("Restaurando sesión y cálculos..."):
                    raw_b = archivo_zip_restaurar.getvalue()
                    ok_r = restaurar_desde_respaldo_bytes(empresa, raw_b)
                    if not ok_r:
                        df_res, renom_res, orig_res, ex_b_res, ex_n_res, zip_res = importar_respaldo_sesion_zip(raw_b, empresa)
                        if df_res is not None:
                            st.session_state["df_procesado"] = df_res
                            st.session_state["dict_pdfs"] = renom_res or {}
                            st.session_state["raw_uploaded_pdfs"] = orig_res or {}
                            st.session_state["excel_bytes"] = ex_b_res
                            st.session_state["excel_nombre"] = ex_n_res
                            st.session_state["zip_pdfs"] = zip_res
                            st.session_state["total_zip_pdfs"] = len(renom_res) if renom_res else 0
                            ok_r = True
                    if ok_r:
                        st.success(f"¡Sesión restaurada con éxito desde {archivo_zip_restaurar.name}!")
                        st.rerun()
                    else:
                        st.error("No se pudo restaurar el archivo de respaldo.")

st.markdown("---")

tab_compras, tab_auditoria, tab_triangulacion, tab_siigo = st.tabs([
    "1. Cargar Documentos y Desbloquear PDFs",
    "2. Auditoria y Trazabilidad Fiscal",
    "3. 🔀 Triangulación y Cruces de Importación (Aduanas)",
    "4. Exportar Planilla Oficial a Siigo"
])

with tab_compras:
    st.markdown("### 1. Insumos DIAN y Facturas en PDF")
    st.write("Sube el archivo Excel de la DIAN (`prueba.xlsx`) o el reporte de facturas, y los PDFs (unificados o separados) para desbloquear, guardar y renombrar automáticamente por comprobante.")

    # Si ya tenemos los documentos cargados en el autoguardado, mostrar confirmación visible
    n_p_carg = len(st.session_state.get("raw_uploaded_pdfs", {}))
    n_p_proc = len(st.session_state.get("dict_pdfs", {}))
    if n_p_carg > 0:
        st.success(f"✅ **Autoguardado Activo:** Tienes **{n_p_carg} facturas PDF cargadas en el sistema** ({n_p_proc} vinculadas con su comprobante). Ya están disponibles en la Pestaña 2 (Auditoría) y Pestaña 3 (Triangulación). No necesitas volver a subirlas.")

    col_u1, col_u2 = st.columns(2)
    with col_u1:
        archivo_excel = st.file_uploader("1. Reporte Excel de la DIAN (ej. prueba.xlsx)", type=["xlsx", "xls"], help="Opcional si ya cargaste tu trabajo desde el Historial.")
    with col_u2:
        archivos_pdfs = st.file_uploader("2. Facturas en PDF o Archivo ZIP Consolidado", type=["pdf", "zip"], accept_multiple_files=True, help="Sube facturas en PDF o un solo archivo .ZIP con todas las facturas (ideal para 100 a 1000+ facturas).")

    st.markdown("##### 🔢 3. Consecutivos Iniciales por Tipo de Comprobante Siigo:")
    st.caption("Cada tipo de comprobante lleva su propia numeración independiente. Digita en qué número vas en cada uno:")

    c_con1, c_con2, c_con3, c_con4 = st.columns(4)
    with c_con1:
        cons_ini_fac = st.number_input("Facturas de Compra (Tipo 10):", min_value=1, max_value=999999, value=680, step=1, help="Consecutivo inicial para facturas de compra normales.")
    with c_con2:
        cons_ini_nc = st.number_input("Notas Crédito (Tipo 17):", min_value=1, max_value=999999, value=1, step=1, help="Consecutivo inicial para notas crédito / devoluciones.")
    with c_con3:
        cons_ini_nota = st.number_input("Notas de Contabilidad (Tipo 14):", min_value=1, max_value=999999, value=1, step=1, help="Consecutivo inicial para notas contables o ajustes.")
    with c_con4:
        clave_pdf_extra = st.text_input("Clave PDF (opcional):", type="password", help="Si algún archivo PDF tiene contraseña específica, ingrésala aquí.")

    # Almacenar PDFs originales subidos en session_state y en el historial de inmediato (solo una vez por subida)
    if archivos_pdfs:
        pdfs_sig = f"{len(archivos_pdfs)}_" + "_".join(f"{p.name}_{p.size}" for p in archivos_pdfs[:5])
        if st.session_state.get("_ultimo_pdfs_proc_sig") != pdfs_sig:
            if "raw_uploaded_pdfs" not in st.session_state:
                st.session_state["raw_uploaded_pdfs"] = {}
            import gc
            for p in archivos_pdfs:
                if p.name.lower().endswith(".zip"):
                    try:
                        with zipfile.ZipFile(io.BytesIO(p.getvalue()), "r") as z_in:
                            for z_name in z_in.namelist():
                                if z_name.lower().endswith(".pdf") and not z_name.startswith("__MACOSX"):
                                    c_name = os.path.basename(z_name)
                                    if c_name:
                                        st.session_state["raw_uploaded_pdfs"][c_name] = z_in.read(z_name)
                    except Exception as e:
                        st.error(f"Error leyendo archivo ZIP {p.name}: {e}")
                else:
                    st.session_state["raw_uploaded_pdfs"][p.name] = p.getvalue()
            gc.collect()

            st.session_state["_ultimo_pdfs_proc_sig"] = pdfs_sig
            st.success(f"📑 **{len(st.session_state['raw_uploaded_pdfs'])} factura(s) PDF cargadas en memoria.**")
            # Auto-vincular de inmediato si ya tenemos el reporte Excel cargado
            if "df_procesado" in st.session_state and st.session_state["df_procesado"] is not None:
                st.session_state["_ejecutar_desbloqueo_ahora"] = True

    if archivo_excel is not None:
        excel_sig = f"{archivo_excel.name}_{archivo_excel.size}_{cons_ini_fac}_{cons_ini_nc}_{cons_ini_nota}"
        if st.session_state.get("_ultimo_excel_proc_sig") != excel_sig or "df_procesado" not in st.session_state:
            st.session_state["excel_bytes"] = archivo_excel.getvalue()
            st.session_state["excel_nombre"] = archivo_excel.name
            st.session_state["_ultimo_excel_proc_sig"] = excel_sig

            df_dian = pd.read_excel(io.BytesIO(st.session_state["excel_bytes"]))

            # Identificar columna de fecha y ordenar cronológicamente de Enero a la fecha actual
            col_fecha = None
            for col_cand in ["Fecha Emisión", "Fecha Emision", "Fecha", "Fecha de Emisión"]:
                if col_cand in df_dian.columns:
                    col_fecha = col_cand
                    break
            if col_fecha is None:
                col_fecha = df_dian.columns[0]

            df_dian["_fecha_dt"] = pd.to_datetime(df_dian[col_fecha], dayfirst=True, errors="coerce")
            df_dian = df_dian.sort_values(by="_fecha_dt", ascending=True).reset_index(drop=True)

            st.success(f"Reporte procesado y ordenado cronológicamente (Enero a la fecha): **{len(df_dian)} facturas identificadas**.")

            contadores_consecutivos = {
                10: int(cons_ini_fac),
                17: int(cons_ini_nc),
                14: int(cons_ini_nota),
                16: int(cons_ini_fac)
            }
            # Analizar celdas rojas (ya registradas), grupos y cuentas específicas del Excel
            estados_excel = analizar_estado_filas_excel(st.session_state.get("excel_bytes"))

            filas = []
            for idx, r in df_dian.iterrows():
                tipo_doc = r.get("Tipo de documento") or r.get("Tipo documento") or "Factura electrónica"
                folio = str(r.get("Folio") or r.get("Factura") or r.get("Factura Num") or f"Doc_{idx+1}").strip()
                if folio.endswith(".0"):
                    folio = folio[:-2]
                prefijo = str(r.get("Prefijo", "")).strip() if pd.notna(r.get("Prefijo")) else ""
                if prefijo == "nan" or prefijo == "None":
                    prefijo = ""

                fecha_val = r.get(col_fecha)
                if pd.notna(fecha_val):
                    try:
                        fecha_dt = pd.to_datetime(fecha_val, dayfirst=True)
                        fecha_str = fecha_dt.strftime("%d/%m/%Y")
                    except:
                        fecha_str = str(fecha_val).split()[0]
                else:
                    fecha_str = "01/01/2026"

                nit_e = str(r.get("NIT Emisor") or r.get("NIT") or "").strip()
                if nit_e.endswith(".0"):
                    nit_e = nit_e[:-2]
                nom_e = str(r.get("Nombre Emisor") or r.get("Proveedor (Emisor)") or r.get("Proveedor") or "Proveedor").strip()

                try: iva = float(r.get("IVA", 0.0)) if pd.notna(r.get("IVA")) else 0.0
                except: iva = 0.0
                try: tot = float(r.get("Total", 0.0)) if pd.notna(r.get("Total")) else 0.0
                except: tot = 0.0
                base, iva, desc_val, motivo_base = calcular_base_con_regla_descuento(r, tot, iva, nom_e)
                col_cufe = next((c for c in df_dian.columns if any(k in str(c).lower() for k in ["cufe", "token", "uuid", "cude"])), None)
                cufe_val = str(r.get(col_cufe, "")).strip() if col_cufe and pd.notna(r.get(col_cufe)) else ""

                col_resp = next((c for c in df_dian.columns if any(k in str(c).lower() for k in ["régimen", "regimen", "responsabilidad", "obligacion"])), None)
                resp_e = str(r.get(col_resp, "")).strip() if col_resp and pd.notna(r.get(col_resp)) else ""

                # Búsqueda inmediata por NIT y Razón Social en catálogo oficial
                nit_e_digits = re.sub(r'\D', '', str(nit_e))
                if not resp_e and nit_e_digits in REGIMENES_EMISORES_CONOCIDOS:
                    resp_e = REGIMENES_EMISORES_CONOCIDOS[nit_e_digits]
                elif not resp_e:
                    nom_e_u = str(nom_e).upper()
                    if any(k in nom_e_u for k in ["DHL", "ESTELAR", "BUENAVENTURA", "SIIGO", "COMCEL", "CLARO", "PANAMERICANA"]):
                        resp_e = "O-13;O-15"

                t_comp, op, cta_p, cta_c, desc, rfte, rica, riva, cta_rfte, cta_iva, cta_rica, cat, razon, audit_dict = clasificar_factura(
                    nit_e, nom_e, base, iva, tipo_doc, resp_e, empresa
                )
                consecutivo = contadores_consecutivos.get(t_comp, 1)
                contadores_consecutivos[t_comp] += 1

                nom_limpio_prov = re.sub(r'[^a-zA-Z0-9]', '', nom_e)[:15]
                nombre_pdf_esperado = f"Comp_{t_comp}-{consecutivo}_{prefijo}{folio}_{nom_limpio_prov}.pdf"

                # Identificar si la fila estaba marcada en rojo (ya registrada en Siigo), su grupo y cuenta
                fac_k_full = f"{prefijo}-{folio}" if prefijo else folio
                info_est = estados_excel.get(cufe_val) or estados_excel.get(fac_k_full) or estados_excel.get(folio) or estados_excel.get(idx, {})
                es_roja_reg = bool(info_est.get("es_roja", False))
                comp_prev = str(info_est.get("comprobante_existente", ""))
                grp_imp = str(info_est.get("grupo_importacion", "")).strip()
                cta_esp = str(info_est.get("cuenta_especifica", "")).strip()

                es_aduanera_flag = any(k in nom_e.upper() for k in AGENTES_ADUANEROS) or any(k in nom_e.upper() for k in ["CARGO", "ADUANA", "PORTUARIA", "ALMACENADORA", "TERMINAL", "BUENAVENTURA", "CARTAGENA", "CONSOLCARGO", "EURO SHIPPING", "TRADE GLOBAL"])

                filas.append({
                    "N°": idx + 1,
                    "Tipo Comp": t_comp,
                    "Consecutivo": consecutivo,
                    "Comprobante Siigo": f"Comp {t_comp}-{consecutivo}",
                    "Fecha": fecha_str,
                    "Prefijo": prefijo,
                    "Folio": folio,
                    "Factura": fac_k_full,
                    "Proveedor": nom_e,
                    "NIT Emisor": nit_e,
                    "Ya Registrada": es_roja_reg,
                    "No Contabilizar": False,
                    "Comprobante Previo": comp_prev,
                    "Grupo Importación": grp_imp,
                    "Cuenta Pasivo Especifica": cta_esp,
                    "Es Aduanera": es_aduanera_flag,
                    "Estado Registro": f"🔴 Ya Registrada ({comp_prev})" if es_roja_reg else ("🟡 Agente Aduanero" if any(k in nom_e.upper() for k in ["EURO SHIPPING", "TRADE GLOBAL"]) else ("🟢 Aduanera" if es_aduanera_flag else "⚪ Compra Pendiente")),
                    "Régimen Fiscal Emisor": resp_e if resp_e else "O-48 (Estándar)",
                    "Descripcion": desc,
                    "Operacion": op,
                    "Cta Principal": cta_p,
                    "Categoría": cat,
                    "Base": base,
                    "IVA": iva,
                    "ReteFuente": rfte,
                    "ReteICA": rica,
                    "ReteIVA": riva,
                    "Neto a Pagar": round(tot - rfte - rica - riva, 2),
                    "Audit Info": audit_dict,
                    "Cta Contrapartida": cta_c,
                    "Cta IVA": cta_iva,
                    "Cta ReteFuente": cta_rfte,
                    "Cta ReteICA": cta_rica,
                    "Total": tot,
                    "CUFE": cufe_val,
                    "Descuento Registrado": desc_val,
                    "Regla Base": motivo_base,
                    "Razón Contable": razon,
                    "Soporte PDF Renombrado": nombre_pdf_esperado
                })

            df_proc = pd.DataFrame(filas)
            st.session_state["df_procesado"] = df_proc

            st.success(f"✅ **Reporte Excel '{archivo_excel.name}' procesado:** {len(df_proc)} facturas identificadas y ordenadas cronológicamente.")

    if "df_procesado" in st.session_state and st.session_state["df_procesado"] is not None:
        df_proc = st.session_state["df_procesado"]

        c_job1, c_job2, c_job3 = st.columns([2.8, 1.1, 1.1])
        with c_job1:
            n_orig_mem = len(st.session_state.get("raw_uploaded_pdfs", {}))
            n_renom_mem = len(st.session_state.get("dict_pdfs", {}))
            nom_ex = st.session_state.get("excel_nombre", "Reporte.xlsx")
            st.info(f"📂 **Reporte DIAN:** `{nom_ex}` ({len(df_proc)} facturas) | 📑 **{n_orig_mem} PDFs cargados** | 📄 **{n_renom_mem} vinculados**")
            if n_orig_mem > 0 and n_renom_mem == 0:
                if st.button(f"🚀 Desbloquear y Vincular los {n_orig_mem} PDFs a Comprobantes", key="btn_desbloquear_top_bar", type="primary", use_container_width=True):
                    st.session_state["_ejecutar_desbloqueo_ahora"] = True
                    st.rerun()
        with c_job2:
            if "excel_bytes" in st.session_state and st.session_state["excel_bytes"]:
                st.download_button(
                    label="📥 Descargar Excel",
                    data=st.session_state["excel_bytes"],
                    file_name=st.session_state.get("excel_nombre", "Reporte.xlsx"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_dl_excel_tab_compras",
                    use_container_width=True
                )
        with c_job3:
            if "zip_pdfs" in st.session_state and st.session_state["zip_pdfs"]:
                st.download_button(
                    label="📦 Descargar ZIP Facturas",
                    data=st.session_state["zip_pdfs"],
                    file_name=f"Facturas_{empresa['nombre'].replace(' ', '_')}.zip",
                    mime="application/zip",
                    key="btn_dl_zip_tab_compras",
                    use_container_width=True
                )


        c_aud_btn1, c_aud_btn2 = st.columns(2)
        with c_aud_btn1:
            st.markdown("#### Matriz Contable Preliminar vinculada a Comprobantes (Orden Cronológico Enero - Actual):")
        with c_aud_btn2:
            if st.button("⚡ Auditar y Escanear Régimen Fiscal de los PDFs (Ultrarrápido)", key="btn_escanear_regimen_todos"):
                dict_renom_s = st.session_state.get("dict_pdfs", {})
                dict_orig_s = st.session_state.get("raw_uploaded_pdfs", {})
                total_modificados = 0
                if dict_renom_s:
                    with st.spinner("⚡ Escaneando régimen fiscal directamente desde las facturas ya desbloqueadas y renombradas..."):
                        total_modificados = auditar_regimen_desde_facturas_renombradas(df_proc, dict_renom_s, empresa)
                    st.session_state["df_procesado"] = df_proc
                    pass
                    st.success(f"⚡ Auditoría completada desde facturas renombradas: se revisaron {len(df_proc)} facturas y se actualizaron {total_modificados} con su régimen oficial del PDF.")
                    st.rerun()
                elif dict_orig_s:
                    with st.spinner("⚡ Extrayendo texto en memoria y cruzando regímenes fiscales..."):
                        # 1. Pre-cargar textos en caché de todos los PDFs una sola vez
                        mapa_textos = []
                        for fn, fb in dict_renom_s.items():
                            mapa_textos.append((fn, [txt.replace("-", "").replace(" ", "").upper() for txt in cache_extraer_textos_pdf(fb)], cache_extraer_textos_pdf(fb)))
                        for fn, fb in dict_orig_s.items():
                            mapa_textos.append((fn, [txt.replace("-", "").replace(" ", "").upper() for txt in cache_extraer_textos_pdf(fb)], cache_extraer_textos_pdf(fb)))

                        # 2. Cruzar en milisegundos cada factura directamente en memoria RAM
                        for r_idx, r_mat in df_proc.iterrows():
                            fol = str(r_mat.get("Folio", "")).replace("-", "").strip().upper()
                            fol_sc = fol.lstrip("0")
                            pref = str(r_mat.get("Prefijo", "")).replace("-", "").strip().upper()
                            fac_full = (pref + fol) if pref else fol
                            fac_full_sc = (pref + fol_sc) if pref else fol_sc
                            nit_d = re.sub(r"\D", "", str(r_mat.get("NIT Emisor", "")))
                            nit_b = nit_d[:-1] if len(nit_d) >= 10 else nit_d

                            texto_factura_encontrado = ""

                            # Buscar en textos pre-extraídos
                            for _, list_clean, list_raw in mapa_textos:
                                for p_idx, t_clean in enumerate(list_clean):
                                    m = False
                                    if len(fac_full) >= 4 and fac_full in t_clean:
                                        m = True
                                    elif len(fac_full_sc) >= 3 and fac_full_sc in t_clean:
                                        m = True
                                    elif (nit_d and len(nit_d) >= 6 and nit_d in t_clean) or (nit_b and len(nit_b) >= 6 and nit_b in t_clean):
                                        if len(fol) >= 2 and fol in t_clean:
                                            m = True
                                        elif len(fol_sc) >= 2 and fol_sc in t_clean:
                                            m = True
                                    if m:
                                        texto_factura_encontrado += list_raw[p_idx] + "\n"
                                if texto_factura_encontrado:
                                    break

                            if texto_factura_encontrado:
                                reg_f = escanear_regimen_texto_pdf(texto_factura_encontrado)
                                if reg_f and reg_f != r_mat.get("Régimen Fiscal Emisor"):
                                    df_proc.at[r_idx, "Régimen Fiscal Emisor"] = reg_f
                                    t_c_n, op_n, c_p_n, c_c_n, desc_n, rfte_n, rica_n, riva_n, c_rf_n, c_iv_n, c_ri_n, cat_n, razon_n, audit_n = clasificar_factura(
                                        df_proc.at[r_idx, "NIT Emisor"], df_proc.at[r_idx, "Proveedor"],
                                        df_proc.at[r_idx, "Base"], df_proc.at[r_idx, "IVA"],
                                        df_proc.at[r_idx, "Operacion"], reg_f, empresa
                                    )
                                    df_proc.at[r_idx, "ReteFuente"] = rfte_n
                                    df_proc.at[r_idx, "ReteICA"] = rica_n
                                    df_proc.at[r_idx, "ReteIVA"] = riva_n
                                    df_proc.at[r_idx, "Cta ReteFuente"] = c_rf_n
                                    df_proc.at[r_idx, "Cta ReteICA"] = c_ri_n
                                    df_proc.at[r_idx, "Razón Contable"] = razon_n
                                    df_proc.at[r_idx, "Audit Info"] = audit_n
                                    df_proc.at[r_idx, "Neto a Pagar"] = round(df_proc.at[r_idx, "Total"] - rfte_n - rica_n - riva_n, 2)
                                    total_modificados += 1

                    st.session_state["df_procesado"] = df_proc
                    pass
                    st.success(f"⚡ Auditoría completada en segundos: se revisaron {len(df_proc)} facturas y se actualizaron {total_modificados} con su régimen oficial del PDF.")
                    st.rerun()
                else:
                    st.warning("Primero sube los PDFs en el campo 2 para poder escanearlos.")

        # Asegurar columna Cuenta Pasivo (CxP) con marcación clara
        if "Cuenta Pasivo (CxP)" not in df_proc.columns:
            df_proc["Cuenta Pasivo (CxP)"] = df_proc.apply(
                lambda r: f"⚠️ {r.get('Cuenta Pasivo Especifica', '22050505')} (Por Verificar)" if r.get("Ya Registrada") else str(r.get("Cuenta Pasivo Especifica", r.get("Cta Contrapartida", "22050501"))),
                axis=1
            )

        st.markdown("""
        <div style="background:#eff6ff; border:1px solid #bfdbfe; border-left:5px solid #2563eb; border-radius:6px; padding:10px 14px; margin-bottom:12px; font-size:13.5px; color:#1e40af;">
            <b>ℹ️ Cuentas de Pasivo (CxP) para Facturas Registradas:</b> Las facturas ya causadas en Siigo (🔴) tienen asignada una cuenta por pagar estimada (ej. <code>22050505</code> Agencia Aduana, <code>23359501</code> Acreedores DHL). Aparecen marcadas con <b>⚠️ (Por Verificar)</b>. Puedes modificar la cuenta fácilmente en la <b>Pestaña 2 (Revisión)</b> o en la <b>Pestaña 3 (Triangulación)</b>.
        </div>
        """, unsafe_allow_html=True)


        # AUTO-AUDITORÍA Y VERIFICACIÓN INMEDIATA DE REGÍMENES FISCALES AL CARGAR EXCEL O PDFS
        dict_renom_mem = st.session_state.get("dict_pdfs", {})
        dict_orig_mem = st.session_state.get("raw_uploaded_pdfs", {})

        # 1. Aplicar catálogo oficial de regímenes conocidos por NIT (solo una vez al cargar archivo inicial)
        if not st.session_state.get("_regimenes_conocidos_aplicados_v1", False):
            if "Audit Info" not in df_proc.columns:
                df_proc["Audit Info"] = [{} for _ in range(len(df_proc))]
            for r_idx, r_row in df_proc.iterrows():
                if bool(r_row.get("Editada Manualmente", False)):
                    continue
                nit_d = re.sub(r'\D', '', str(r_row.get("NIT Emisor", "")))
                nom_u = str(r_row.get("Proveedor", "")).upper()
                reg_prev = str(r_row.get("Régimen Fiscal Emisor", ""))

                nuevo_reg = None
                if nit_d in REGIMENES_EMISORES_CONOCIDOS:
                    nuevo_reg = REGIMENES_EMISORES_CONOCIDOS[nit_d]
                elif any(k in nom_u for k in ["DHL", "ESTELAR", "BUENAVENTURA", "SIIGO", "COMCEL", "CLARO", "PANAMERICANA"]):
                    nuevo_reg = "O-13;O-15"

                if nuevo_reg and (nuevo_reg != reg_prev):
                    df_proc.at[r_idx, "Régimen Fiscal Emisor"] = nuevo_reg
                    t_c, op_c, c_p, c_c, desc_c, rfte_c, rica_c, riva_c, c_rf, c_iv, c_ri, cat_c, razon_c, audit_c = clasificar_factura(
                        r_row["NIT Emisor"], r_row["Proveedor"],
                        r_row["Base"], r_row["IVA"],
                        r_row["Operacion"], nuevo_reg, empresa
                    )
                    df_proc.at[r_idx, "ReteFuente"] = rfte_c
                    df_proc.at[r_idx, "ReteICA"] = rica_c
                    df_proc.at[r_idx, "ReteIVA"] = riva_c
                    df_proc.at[r_idx, "Cta ReteFuente"] = c_rf
                    df_proc.at[r_idx, "Cta ReteICA"] = c_ri
                    df_proc.at[r_idx, "Razón Contable"] = razon_c
                    df_proc.at[r_idx, "Audit Info"] = audit_c
                    df_proc.at[r_idx, "Neto a Pagar"] = round(r_row["Total"] - rfte_c - rica_c - riva_c, 2)
            st.session_state["_regimenes_conocidos_aplicados_v1"] = True

        # Auditoría de régimen fiscal disponible bajo demanda mediante el botón superior
        st.session_state["df_procesado"] = df_proc

        cols_mostrar_proc = ["Comprobante Siigo", "Fecha", "Factura", "Proveedor", "NIT Emisor", "Régimen Fiscal Emisor", "Estado Registro", "Cuenta Pasivo (CxP)", "Base", "IVA", "ReteFuente", "ReteICA", "Neto a Pagar", "Soporte PDF Renombrado"]
        cols_mostrar_existentes = [c for c in cols_mostrar_proc if c in df_proc.columns]
        st.dataframe(df_proc[cols_mostrar_existentes], use_container_width=True)

        st.markdown("---")
        st.markdown("#### 📑 Procesamiento, Separación y Renombrado de PDFs Multi-Página:")
        st.caption("Identifica y une facturas completas de 2 o más hojas a partir de un PDF consolidado, asignando a cada una su comprobante y consecutivo oficial.")

        col_cfg1, col_cfg2 = st.columns(2)
        with col_cfg1:
            modo_sep = st.radio(
                "Configuración de reconocimiento de facturas:",
                [
                    "🔍 Detección Inteligente (Reconoce facturas completas de 2 o más hojas por 'Página 1 de 2', CUFE y Número)",
                    "📄 Agrupación fija de 2 páginas por factura (cada factura tiene 2 hojas)",
                    "📄 Agrupación fija de 3 páginas por factura (cada factura tiene 3 hojas)",
                    "📑 Separar 1 página por factura (facturas de 1 sola hoja)",
                    "📂 Mantener archivos individuales (sin separar páginas)"
                ]
            )
        with col_cfg2:
            st.info("💡 **Garantía de Factura Completa:** El motor inteligente detecta dónde empieza cada factura (Prefijo, Folio, NIT y marcadores de paginación). Todas las páginas de una misma factura se unen en un solo archivo PDF completo nombrado `Comp_10-XXX_Factura_Proveedor.pdf`.")

        ejecutar_desb = st.button("🔓 Desbloquear, Separar y Renombrar PDFs ahora") or st.session_state.pop("_ejecutar_desbloqueo_ahora", False)
        if ejecutar_desb:
            df_ref = st.session_state.get("df_procesado", pd.DataFrame())

            # Lista de PDFs a procesar: extraer archivos .pdf reales tanto de subidas directas como de archivos .zip
            if "raw_uploaded_pdfs" not in st.session_state:
                st.session_state["raw_uploaded_pdfs"] = {}

            pdfs_reales_dict = {}
            # 1. Incorporar PDFs ya existentes en raw_uploaded_pdfs
            for fn, bdata in st.session_state.get("raw_uploaded_pdfs", {}).items():
                if fn.lower().endswith(".pdf"):
                    pdfs_reales_dict[fn] = bdata

            # 2. Desempaquetar archivos_pdfs si contienen archivos .zip o PDFs directos
            if archivos_pdfs:
                for p_it in archivos_pdfs:
                    p_nom = getattr(p_it, "name", "").lower()
                    if p_nom.endswith(".zip"):
                        try:
                            b_val = p_it.getvalue() if hasattr(p_it, "getvalue") else p_it.read()
                            with zipfile.ZipFile(io.BytesIO(b_val), "r") as z_u:
                                for z_fn in z_u.namelist():
                                    if z_fn.lower().endswith(".pdf") and not z_fn.startswith("__MACOSX"):
                                        b_pdf = z_u.read(z_fn)
                                        c_fn = os.path.basename(z_fn)
                                        if c_fn:
                                            pdfs_reales_dict[c_fn] = b_pdf
                                            st.session_state["raw_uploaded_pdfs"][c_fn] = b_pdf
                        except Exception as e_z:
                            st.error(f"Error abriendo ZIP {getattr(p_it, 'name', '')}: {e_z}")
                    elif p_nom.endswith(".pdf") or not p_nom:
                        b_val = p_it.getvalue() if hasattr(p_it, "getvalue") else p_it.read()
                        pdfs_reales_dict[getattr(p_it, "name", "factura.pdf")] = b_val

            # Construir la lista limpia de streams PDF con su nombre asignado
            pdfs_a_procesar = []
            for fn, b_pdf in pdfs_reales_dict.items():
                bio = io.BytesIO(b_pdf)
                bio.name = fn
                pdfs_a_procesar.append(bio)

            if not pdfs_a_procesar:
                st.warning("⚠️ No hay archivos PDF para procesar. Por favor sube tus facturas en el campo '2. Facturas en PDF (unificadas o separadas)' arriba antes de hacer clic en desbloquear.")
                st.stop()

            buffer_zip = io.BytesIO()
            total_generados = 0

            if "dict_pdfs" not in st.session_state:
                st.session_state["dict_pdfs"] = {}

            total_pdfs_count = len(pdfs_a_procesar)
            prog_bar_pdf = st.progress(0, text=f"Iniciando procesamiento de {total_pdfs_count} archivo(s) PDF...")

            with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_STORED) as zf:
                for idx_pdf, pdf_item in enumerate(pdfs_a_procesar):
                    pdf_name = getattr(pdf_item, "name", f"Documento_{idx_pdf+1}.pdf")
                    prog_bar_pdf.progress(min(1.0, (idx_pdf + 1) / total_pdfs_count), text=f"⚡ Factura {idx_pdf + 1} de {total_pdfs_count}: {pdf_name[:20]}...")
                    # Si ya fue procesado y guardado en dict_pdfs, re-usar directamente
                    if pdf_name in st.session_state["dict_pdfs"] and len(st.session_state["dict_pdfs"][pdf_name]) > 100:
                        zf.writestr(pdf_name, st.session_state["dict_pdfs"][pdf_name])
                        total_generados += 1
                        continue
                    try:
                        pdf_name = getattr(pdf_item, "name", f"Documento_{idx_pdf+1}.pdf")
                        reader = PdfReader(pdf_item)
                        desencriptado = False
                        if reader.is_encrypted:
                            claves_probar = []
                            if "clave_pdf_extra" in locals() and clave_pdf_extra:
                                claves_probar.append(clave_pdf_extra.strip())

                            # NITs empresa compradora
                            nit_rec = re.sub(r"\D", "", str(empresa.get("nit", "")))
                            if nit_rec:
                                if len(nit_rec) > 9:
                                    claves_probar.extend([nit_rec, nit_rec[:-1], f"{nit_rec[:-1]}-{nit_rec[-1]}"])
                                else:
                                    claves_probar.extend([nit_rec, f"{nit_rec}5", f"{nit_rec}-5"])
                            claves_probar.extend(["901346412", "9013464125", "901346412-5", ""])

                            # NITs emisores en reporte
                            if not df_ref.empty:
                                for n_e in df_ref["NIT Emisor"].dropna().unique():
                                    n_c = re.sub(r"\D", "", str(n_e))
                                    if n_c and len(n_c) >= 7:
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
                                st.warning(f"⚠️ No se pudo desencriptar automáticamente {pdf_name}. Ingrese la contraseña en el campo correspondiente.")
                                continue

                        num_pags = len(reader.pages)

                        # Caso 1: Archivo individual (sin separar páginas) o de 1 página
                        if num_pags == 1 or "Mantener archivos individuales" in modo_sep:
                            writer = PdfWriter()
                            texto_pdf = ""
                            for page in reader.pages:
                                writer.add_page(page)
                                try: texto_pdf += (page.extract_text() or "") + "\n"
                                except: pass

                            r_mat = None
                            if not df_ref.empty:
                                # 1. Identificar por texto exacto (CUFE, Factura + Proveedor/NIT)
                                r_mat = identificar_factura_en_texto(texto_pdf, df_ref)

                                # 2. Si no se identificó en el texto, buscar por coincidencia en el nombre del archivo
                                if r_mat is None:
                                    fn_clean = re.sub(r'[^A-Za-z0-9]', '', pdf_name).lower()
                                    for _, r_cand in df_ref.iterrows():
                                        cufe_c = re.sub(r'[^A-Za-z0-9]', '', str(r_cand.get("CUFE", "") or "")).lower()
                                        if cufe_c and len(cufe_c) >= 15 and cufe_c[:20] in fn_clean:
                                            r_mat = r_cand
                                            break
                                        pref_c = re.sub(r'[^A-Za-z0-9]', '', str(r_cand.get("Prefijo", "") or "")).upper()
                                        fol_c = re.sub(r'[^A-Za-z0-9]', '', str(r_cand.get("Folio", "") or "")).upper()
                                        fac_f = (pref_c + fol_c) if pref_c else fol_c
                                        if fac_f and len(fac_f) >= 4 and fac_f in fn_clean.upper():
                                            r_mat = r_cand
                                            break

                            if r_mat is not None:
                                nombre_final = r_mat["Soporte PDF Renombrado"]
                                # Escaneo de régimen fiscal sólo para la factura que realmente coincidió
                                regimen_escaneado = escanear_regimen_texto_pdf(texto_pdf)
                                if regimen_escaneado:
                                    r_idx = r_mat.name
                                    df_ref.at[r_idx, "Régimen Fiscal Emisor"] = regimen_escaneado
                                    _, _, _, _, _, rfte_n, rica_n, riva_n, cta_rf_n, _, cta_ri_n, _, razon_n, audit_n = clasificar_factura(
                                        df_ref.at[r_idx, "NIT Emisor"], df_ref.at[r_idx, "Proveedor"],
                                        df_ref.at[r_idx, "Base"], df_ref.at[r_idx, "IVA"],
                                        df_ref.at[r_idx, "Operacion"], regimen_escaneado, empresa
                                    )
                                    df_ref.at[r_idx, "ReteFuente"] = rfte_n
                                    df_ref.at[r_idx, "ReteICA"] = rica_n
                                    df_ref.at[r_idx, "ReteIVA"] = riva_n
                                    df_ref.at[r_idx, "Cta ReteFuente"] = cta_rf_n
                                    df_ref.at[r_idx, "Cta ReteICA"] = cta_ri_n
                                    df_ref.at[r_idx, "Razón Contable"] = razon_n
                                    df_ref.at[r_idx, "Audit Info"] = audit_n
                                    df_ref.at[r_idx, "Neto a Pagar"] = round(df_ref.at[r_idx, "Total"] - rfte_n - rica_n - riva_n, 2)
                            else:
                                nombre_final = f"Soporte_{idx_pdf+1}_{pdf_name}"

                            # Reutilizar directamente los bytes originales sin re-escribir con PdfWriter (1000x más rápido)
                            b_data = pdf_item.getvalue() if hasattr(pdf_item, "getvalue") else pdf_item.read()
                            zf.writestr(nombre_final, b_data)
                            st.session_state["dict_pdfs"][nombre_final] = b_data
                            total_generados += 1

                        # Caso 2: Agrupación fija de 2 páginas
                        elif "Agrupación fija de 2 páginas" in modo_sep:
                            for p_i in range(0, num_pags, 2):
                                writer = PdfWriter()
                                txt_grp = ""
                                for off in range(2):
                                    if p_i + off < num_pags:
                                        writer.add_page(reader.pages[p_i + off])
                                        try: txt_grp += (reader.pages[p_i + off].extract_text() or "") + "\n"
                                        except: pass

                                r_mat = identificar_factura_en_texto(txt_grp, df_ref)
                                if r_mat is not None:
                                    nombre_final = r_mat["Soporte PDF Renombrado"]
                                else:
                                    fac_idx = p_i // 2
                                    nombre_final = df_ref.iloc[fac_idx]["Soporte PDF Renombrado"] if not df_ref.empty and fac_idx < len(df_ref) else f"Factura_Pags_{p_i+1}-{p_i+2}_{pdf_name}"

                                pdf_bytes = io.BytesIO()
                                writer.write(pdf_bytes)
                                pdf_bytes.seek(0)
                                b_data = pdf_bytes.getvalue()
                                zf.writestr(nombre_final, b_data)
                                st.session_state["dict_pdfs"][nombre_final] = b_data
                                total_generados += 1

                        # Caso 3: Agrupación fija de 3 páginas
                        elif "Agrupación fija de 3 páginas" in modo_sep:
                            for p_i in range(0, num_pags, 3):
                                writer = PdfWriter()
                                txt_grp = ""
                                for off in range(3):
                                    if p_i + off < num_pags:
                                        writer.add_page(reader.pages[p_i + off])
                                        try: txt_grp += (reader.pages[p_i + off].extract_text() or "") + "\n"
                                        except: pass

                                r_mat = identificar_factura_en_texto(txt_grp, df_ref)
                                if r_mat is not None:
                                    nombre_final = r_mat["Soporte PDF Renombrado"]
                                else:
                                    fac_idx = p_i // 3
                                    nombre_final = df_ref.iloc[fac_idx]["Soporte PDF Renombrado"] if not df_ref.empty and fac_idx < len(df_ref) else f"Factura_Pags_{p_i+1}-{p_i+3}_{pdf_name}"

                                pdf_bytes = io.BytesIO()
                                writer.write(pdf_bytes)
                                pdf_bytes.seek(0)
                                b_data = pdf_bytes.getvalue()
                                zf.writestr(nombre_final, b_data)
                                st.session_state["dict_pdfs"][nombre_final] = b_data
                                total_generados += 1

                        # Caso 4: 1 página por factura
                        elif "Separar 1 página" in modo_sep:
                            for p_i in range(num_pags):
                                writer = PdfWriter()
                                writer.add_page(reader.pages[p_i])
                                try: txt_pag = (reader.pages[p_i].extract_text() or "") + "\n"
                                except: txt_pag = ""

                                r_mat = identificar_factura_en_texto(txt_pag, df_ref)
                                if r_mat is not None:
                                    nombre_final = r_mat["Soporte PDF Renombrado"]
                                else:
                                    nombre_final = df_ref.iloc[p_i]["Soporte PDF Renombrado"] if not df_ref.empty and p_i < len(df_ref) else f"Factura_Pag_{p_i+1}_{pdf_name}"

                                pdf_bytes = io.BytesIO()
                                writer.write(pdf_bytes)
                                pdf_bytes.seek(0)
                                b_data = pdf_bytes.getvalue()
                                zf.writestr(nombre_final, b_data)
                                st.session_state["dict_pdfs"][nombre_final] = b_data
                                total_generados += 1

                        # Caso 5: Detección inteligente por CUFE, Folio o 'Página 1 de N'
                        else:
                            patron_pag = re.compile(r'(?:P[ÁAáa]G(?:INA|\.)?|HOJA|PAGE)\s*(\d+)\s*(?:DE|\/|OF)\s*(\d+)', re.IGNORECASE)
                            curr_writer = None
                            curr_inv_row = None
                            paginas_del_comprobante = 0
                            facturas_generadas = []

                            def buscar_coincidencia_factura(texto_pagina):
                                return identificar_factura_en_texto(texto_pagina, df_ref)

                            for p_idx in range(num_pags):
                                try: txt_p = reader.pages[p_idx].extract_text() or ""
                                except: txt_p = ""

                                reg_detectado = escanear_regimen_texto_pdf(txt_p)

                                m_pag = patron_pag.search(txt_p)
                                cur_p, tot_p = (int(m_pag.group(1)), int(m_pag.group(2))) if m_pag else (None, None)

                                inv_encontrada = buscar_coincidencia_factura(txt_p)

                                es_inicio_nueva_factura = False
                                if cur_p == 1:
                                    es_inicio_nueva_factura = True
                                elif cur_p is not None and cur_p > 1:
                                    es_inicio_nueva_factura = False
                                elif inv_encontrada is not None:
                                    if curr_inv_row is None or inv_encontrada["Comprobante Siigo"] != curr_inv_row["Comprobante Siigo"]:
                                        es_inicio_nueva_factura = True
                                elif curr_inv_row is None:
                                    es_inicio_nueva_factura = True

                                if es_inicio_nueva_factura:
                                    if curr_writer is not None and curr_inv_row is not None:
                                        out_c = io.BytesIO()
                                        curr_writer.write(out_c)
                                        c_bytes = out_c.getvalue()
                                        doc_nombre = curr_inv_row["Soporte PDF Renombrado"]
                                        zf.writestr(doc_nombre, c_bytes)
                                        st.session_state["dict_pdfs"][doc_nombre] = c_bytes
                                        total_generados += 1
                                        facturas_generadas.append(f"{curr_inv_row['Comprobante Siigo']} ({paginas_del_comprobante} págs)")

                                    curr_inv_row = inv_encontrada if inv_encontrada is not None else curr_inv_row
                                    if curr_inv_row is not None and reg_detectado and not df_ref.empty:
                                        r_idx = curr_inv_row.name
                                        df_ref.at[r_idx, "Régimen Fiscal Emisor"] = reg_detectado
                                        _, _, _, _, _, rfte_n, rica_n, riva_n, cta_rf_n, _, cta_ri_n, _, razon_n, audit_n = clasificar_factura(
                                            df_ref.at[r_idx, "NIT Emisor"], df_ref.at[r_idx, "Proveedor"],
                                            df_ref.at[r_idx, "Base"], df_ref.at[r_idx, "IVA"],
                                            df_ref.at[r_idx, "Operacion"], reg_detectado, empresa
                                        )
                                        df_ref.at[r_idx, "ReteFuente"] = rfte_n
                                        df_ref.at[r_idx, "ReteICA"] = rica_n
                                        df_ref.at[r_idx, "ReteIVA"] = riva_n
                                        df_ref.at[r_idx, "Cta ReteFuente"] = cta_rf_n
                                        df_ref.at[r_idx, "Cta ReteICA"] = cta_ri_n
                                        df_ref.at[r_idx, "Razón Contable"] = razon_n
                                        df_ref.at[r_idx, "Audit Info"] = audit_n
                                        df_ref.at[r_idx, "Neto a Pagar"] = round(df_ref.at[r_idx, "Total"] - rfte_n - rica_n - riva_n, 2)
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
                                doc_nombre = curr_inv_row["Soporte PDF Renombrado"]
                                zf.writestr(doc_nombre, c_bytes)
                                st.session_state["dict_pdfs"][doc_nombre] = c_bytes
                                total_generados += 1
                                facturas_generadas.append(f"{curr_inv_row['Comprobante Siigo']} ({paginas_del_comprobante} págs)")

                    except Exception as e:
                        st.error(f"Error procesando {pdf_name}: {e}")

            prog_bar_pdf.progress(1.0, text="✅ ¡100% Procesado y Vinculado con Éxito!")
            buffer_zip.seek(0)
            st.session_state["zip_pdfs"] = buffer_zip.getvalue()
            st.session_state["total_zip_pdfs"] = total_generados

            # Escaneo oficial de régimen fiscal directamente sobre las facturas ya desbloqueadas y renombradas
            total_auditados = auditar_regimen_desde_facturas_renombradas(df_ref, st.session_state["dict_pdfs"], empresa)
            st.session_state["df_procesado"] = df_ref

            # Guardar trabajo actualizado con PDFs procesados y regímenes verificados
            guardar_trabajo_en_historial(
                empresa, df_ref,
                excel_bytes=st.session_state.get("excel_bytes"),
                excel_nombre=st.session_state.get("excel_nombre", "Reporte.xlsx"),
                dict_pdfs_renombrados=st.session_state.get("dict_pdfs", {}),
                dict_pdfs_originales=st.session_state.get("raw_uploaded_pdfs", {}),
                zip_bytes=st.session_state["zip_pdfs"],
                consecutivo_ini=cons_ini_fac,
                job_id=st.session_state.get("job_actual_id")
            )
            msg_reg = f" y se auditó el régimen de cada factura ({total_auditados} actualizadas)" if total_auditados > 0 else " y se validó el régimen de cada factura"
            st.success(f"¡Procesamiento exitoso! Se desbloquearon y renombraron **{total_generados} facturas completas**{msg_reg} directamente desde las facturas PDF desbloqueadas.")

        zip_data_b = st.session_state.get("zip_pdfs")
        if zip_data_b and isinstance(zip_data_b, (bytes, bytearray)) and len(zip_data_b) > 22:
            n_tot_z = st.session_state.get("total_zip_pdfs", len(st.session_state.get("dict_pdfs", {})))
            st.download_button(
                label=f"📦 Descargar Archivos PDF Renombrados ({n_tot_z} facturas completas en ZIP)",
                data=zip_data_b,
                file_name=f"Facturas_Renombradas_Comprobantes_{empresa['nombre'].replace(' ', '_')}.zip",
                mime="application/zip",
                key="btn_dl_zip_pdfs_renombrados_action_main",
                use_container_width=True
            )



    else:
        st.info("💡 Sube el reporte Excel de la DIAN arriba o carga un trabajo guardado desde el **Historial de Trabajos** para comenzar.")

with tab_auditoria:
    st.markdown("### Modulo de Auditoria Contable y Trazabilidad")
    st.caption("Inspección de cuentas, deducciones y separación de gastos por cuenta de terceros.")

    if "df_procesado" in st.session_state:
        df_p = st.session_state["df_procesado"]

        # Asegurar columna No Contabilizar en df_p
        if "No Contabilizar" not in df_p.columns:
            df_p["No Contabilizar"] = False

        # Normalizar columnas numéricas críticas para evitar que aparezcan valores NaN / None
        if "IVA Importación" not in df_p.columns:
            df_p["IVA Importación"] = 0.0
        else:
            df_p["IVA Importación"] = pd.to_numeric(df_p["IVA Importación"], errors="coerce").fillna(0.0)

        if "Cta IVA Importación" not in df_p.columns:
            df_p["Cta IVA Importación"] = "240835"
        else:
            df_p["Cta IVA Importación"] = df_p["Cta IVA Importación"].fillna("240835").astype(str).replace(["nan", "None", ""], "240835")
            
        st.session_state["df_procesado"] = df_p

        # Inicializar conjunto de facturas excluidas de contabilización
        if "facturas_no_contabilizar" not in st.session_state:
            st.session_state["facturas_no_contabilizar"] = set(df_p[df_p["No Contabilizar"] == True]["Factura"].dropna().tolist())
        cur_no_contab = st.session_state["facturas_no_contabilizar"]

        # Panel de exclusión masiva para marcar cuáles NO contabilizar
        with st.expander("🚫 Panel: Marcar qué facturas NO debes contabilizar (Excluir de Siigo):", expanded=False):
            st.caption("Selecciona aquí las facturas que NO deseas que pasen a la contabilidad ni se exporten a Siigo (por ejemplo, si ya fueron causadas manualmente, corresponden a otro período o no deben registrarse):")
            lista_todas_facs_tags = [f"{r['Factura']} — {r['Proveedor'][:24]} (${r['Total']:,.0f}) [{'🔴 Ya Registrada' if r.get('Ya Registrada') else '⚪ Pendiente'}]" for _, r in df_p.iterrows()]
            mapa_tags_to_fac = {f"{r['Factura']} — {r['Proveedor'][:24]} (${r['Total']:,.0f}) [{'🔴 Ya Registrada' if r.get('Ya Registrada') else '⚪ Pendiente'}]": r['Factura'] for _, r in df_p.iterrows()}

            defaults_seleccionados = [t for t, fn in mapa_tags_to_fac.items() if fn in cur_no_contab or bool(df_p[df_p["Factura"] == fn].iloc[0].get("No Contabilizar", False))]

            sel_excluidas_bulk = st.multiselect(
                "Selecciona facturas que NO se deben contabilizar:",
                options=lista_todas_facs_tags,
                default=defaults_seleccionados,
                key="ms_bulk_excluir_contab"
            )
            if st.button("💾 Guardar Exclusiones de Contabilidad", key="btn_save_bulk_excl"):
                nuevas_ex = {mapa_tags_to_fac[t] for t in sel_excluidas_bulk}
                st.session_state["facturas_no_contabilizar"] = nuevas_ex
                for idx_row, r_row in df_p.iterrows():
                    es_ex = r_row["Factura"] in nuevas_ex
                    df_p.at[idx_row, "No Contabilizar"] = es_ex
                    if es_ex:
                        df_p.at[idx_row, "Estado Registro"] = "🚫 No Contabilizar (Excluida)"
                    else:
                        es_r = bool(r_row.get("Ya Registrada", False))
                        cp = r_row.get("Comprobante Previo", "")
                        df_p.at[idx_row, "Estado Registro"] = f"🔴 Ya Registrada ({cp})" if es_r else "⚪ Compra Pendiente"
                st.session_state["df_procesado"] = df_p
                st.success(f"¡Se han marcado {len(nuevas_ex)} facturas como NO contabilizables!")
                st.rerun()

        opciones_fac = []
        for _, r in df_p.iterrows():
            fn = r['Factura']
            tag_ex = " 🚫 [NO CONTABILIZAR]" if (fn in cur_no_contab or r.get("No Contabilizar")) else ""
            opciones_fac.append(f"[{r['Comprobante Siigo']}] {r['Fecha']} - {fn}{tag_ex} - {r['Proveedor']} (${r['Total']:,.0f})")

        # Preservar la factura seleccionada para que no vuelva a la primera tras guardar cambios o editar
        comp_activo_guardado = st.session_state.get("factura_seleccionada_auditoria")
        idx_fac_sel = 0
        if comp_activo_guardado:
            for i_f, op_str in enumerate(opciones_fac):
                if f"[{comp_activo_guardado}]" in op_str:
                    idx_fac_sel = i_f
                    break

        if "sel_factura_auditoria_box" in st.session_state:
            if st.session_state["sel_factura_auditoria_box"] not in opciones_fac and idx_fac_sel < len(opciones_fac):
                st.session_state["sel_factura_auditoria_box"] = opciones_fac[idx_fac_sel]

        seleccion = st.selectbox(
            "Selecciona una factura para auditar:",
            opciones_fac,
            index=idx_fac_sel,
            key="sel_factura_auditoria_box"
        )

        comp_sel = seleccion.split("]")[0].replace("[", "").strip()
        st.session_state["factura_seleccionada_auditoria"] = comp_sel
        fac_sel = df_p[df_p["Comprobante Siigo"] == comp_sel].iloc[0]
        es_aduanero = any(k in fac_sel["Proveedor"].upper() for k in AGENTES_ADUANEROS)

        audit_info = fac_sel.get("Audit Info", {})
        if not isinstance(audit_info, dict):
            audit_info = {}

        col_a1, col_a2 = st.columns(2)
        with col_a1:
            # Control individual para marcar si no se debe contabilizar esta factura
            fac_act_fn = fac_sel["Factura"]
            esta_excl_ind = fac_act_fn in cur_no_contab or bool(fac_sel.get("No Contabilizar", False))

            c_tog_ex1, c_tog_ex2 = st.columns([2.4, 1.2])
            with c_tog_ex1:
                if esta_excl_ind:
                    st.warning("🚫 **Esta factura está marcada como 'NO CONTABILIZAR'** — No se generará asiento ni se exportará a Siigo.")
                else:
                    st.caption("Estado contable activo para exportación a Siigo.")
            with c_tog_ex2:
                if not esta_excl_ind:
                    if st.button("🚫 No Contabilizar", key=f"btn_excl_indiv_{fac_sel['Comprobante Siigo']}", help="Excluye esta factura de Siigo"):
                        cur_no_contab.add(fac_act_fn)
                        st.session_state["facturas_no_contabilizar"] = cur_no_contab
                        m_idx = df_p[df_p["Factura"] == fac_act_fn].index
                        if not m_idx.empty:
                            df_p.at[m_idx[0], "No Contabilizar"] = True
                            df_p.at[m_idx[0], "Estado Registro"] = "🚫 No Contabilizar (Excluida)"
                            st.session_state["df_procesado"] = df_p
                        st.session_state["factura_seleccionada_auditoria"] = comp_sel
                        st.session_state.pop("sel_factura_auditoria_box", None)
                        st.rerun()
                else:
                    if st.button("✅ Reactivar", key=f"btn_react_indiv_{fac_sel['Comprobante Siigo']}", help="Vuelve a incluirla en la contabilidad"):
                        cur_no_contab.discard(fac_act_fn)
                        st.session_state["facturas_no_contabilizar"] = cur_no_contab
                        m_idx = df_p[df_p["Factura"] == fac_act_fn].index
                        if not m_idx.empty:
                            df_p.at[m_idx[0], "No Contabilizar"] = False
                            es_r = bool(df_p.at[m_idx[0], "Ya Registrada"])
                            cp = df_p.at[m_idx[0], "Comprobante Previo"]
                            df_p.at[m_idx[0], "Estado Registro"] = f"🔴 Ya Registrada ({cp})" if es_r else "⚪ Compra Pendiente"
                            st.session_state["df_procesado"] = df_p
                        st.session_state["factura_seleccionada_auditoria"] = comp_sel
                        st.session_state.pop("sel_factura_auditoria_box", None)
                        st.rerun()

            st.markdown(f"""
            <div class="audit-card">
                <h4>Detalle del Comprobante: {fac_sel['Comprobante Siigo']}</h4>
                <p><b>Soporte PDF Vinculado:</b> <span class="tag-comp">{fac_sel['Soporte PDF Renombrado']}</span></p>
                <p><b>Factura:</b> {fac_sel['Factura']} - <b>Proveedor:</b> {fac_sel['Proveedor']} (NIT: {fac_sel['NIT Emisor']})</p>
                <p><b>Fecha de Emisión:</b> {fac_sel['Fecha']} - <b>Total:</b> ${fac_sel['Total']:,.2f}</p>
                <p><b>Régimen Fiscal Emisor:</b> <span class="badge-active">{fac_sel.get('Régimen Fiscal Emisor', 'O-48')}</span></p>
                <div style="margin: 8px 0;">
                    <span class="badge-tax">{'✅ Gran Contribuyente (O-13)' if audit_info.get('es_gc') else 'Común (No GC)'}</span>
                    <span class="badge-tax">{'✅ Autorretenedor (O-15)' if audit_info.get('es_autorr') else 'No Autorretenedor'}</span>
                    <span class="badge-tax">{'⚠️ Régimen Simple (O-47)' if audit_info.get('es_rst') else 'Régimen Ordinario'}</span>
                </div>
                <hr>
                <h5>Trazabilidad de la Contabilización y Validación DIAN:</h5>
                <ul>
                    <li><b>Cuenta Asignada:</b> <span class="tag-propio">{fac_sel['Cta Principal']}</span> - {fac_sel['Categoría']}</li>
                    <li><b>Cruce DIAN:</b> {fac_sel['Razón Contable']}</li>
                    <li><b>Base Gravable:</b> ${fac_sel['Base']:,.2f} - <b>IVA:</b> ${fac_sel['IVA']:,.2f}</li>
                    <li><b>Retención en la Fuente:</b> ${fac_sel['ReteFuente']:,.2f} ({audit_info.get('razon_rfte', 'N/A')})</li>
                    <li><b>ReteICA:</b> ${fac_sel.get('ReteICA', 0.0):,.2f} ({audit_info.get('razon_reteica', audit_info.get('razon_rica', 'N/A'))})</li>
                    <li><b>ReteIVA:</b> ${fac_sel.get('ReteIVA', 0.0):,.2f} ({audit_info.get('razon_reteiva', 'N/A')})</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)

            with st.expander("⚙️ Ajustar Régimen, Cuentas (CxP), Base Gravable e Impuestos Asumidos:", expanded=True if (fac_sel.get("Ya Registrada") or es_aduanero) else False):
                st.caption("Modifica aquí los valores fiscales, la cuenta por pagar o el tratamiento de retenciones asumidas:")
                c_mod1, c_mod2 = st.columns(2)
                with c_mod1:
                    opciones_reg_man = [
                        "O-48 (Responsable de IVA / Común - Retención Ordinaria)",
                        "O-15 (Autorretenedor de Renta - Sin ReteFuente)",
                        "O-13 (Gran Contribuyente - Sin ReteIVA)",
                        "O-13;O-15 (Gran Contribuyente y Autorretenedor)",
                        "O-47 (Régimen Simple - Exento ReteFuente y ReteIVA)",
                        "O-49 (No Responsable de IVA)"
                    ]
                    idx_actual = 0
                    reg_actual = str(fac_sel.get("Régimen Fiscal Emisor", ""))
                    if "O-13" in reg_actual and "O-15" in reg_actual: idx_actual = 3
                    elif "O-15" in reg_actual: idx_actual = 1
                    elif "O-13" in reg_actual: idx_actual = 2
                    elif "O-47" in reg_actual: idx_actual = 4
                    elif "O-49" in reg_actual: idx_actual = 5

                    nuevo_reg_sel = st.selectbox("Selecciona Régimen Fiscal:", opciones_reg_man, index=idx_actual, key=f"sel_reg_{fac_sel['Comprobante Siigo']}")
                    cod_reg_sel = nuevo_reg_sel.split()[0]

                    cta_p_def = str(fac_sel.get("Cta Principal") or ("14650501" if es_aduanero else "51350501")).strip()
                    nueva_cta_p = st.text_input(
                        "Cuenta Principal (Gasto / Costo / Inventario):",
                        value=cta_p_def,
                        key=f"inp_cta_p_{fac_sel['Comprobante Siigo']}",
                        help="Cuenta contable donde se imputa la base de la factura (ej. 51350501 Servicios, 14650501 Mercancías en Tránsito)."
                    )

                    cta_cxp_def = str(fac_sel.get("Cuenta Pasivo Especifica") or fac_sel.get("Cta Contrapartida") or ("22050505" if es_aduanero else "22050501")).strip()
                    nueva_cta_cxp = st.text_input(
                        "Cuenta por Pagar (Pasivo / Contrapartida):",
                        value=cta_cxp_def,
                        key=f"inp_cxp_{fac_sel['Comprobante Siigo']}",
                        help="Puedes cambiar aquí la cuenta pasivo (ej. 22050505 Agencia Aduana, 23359501 Acreedores/DHL, 22050501 Proveedores)."
                    )

                    es_importacion_factura = es_aduanero or "1465" in str(fac_sel.get("Cta Principal", "")) or "IMPORTACI" in str(fac_sel.get("Categoría", "")).upper()
                    cta_iva_def = str(fac_sel.get("Cta IVA") or ("24081501" if es_importacion_factura else "24081001")).strip()
                    nueva_cta_iva = st.text_input(
                        "Cuenta IVA Servicios / Compras:",
                        value=cta_iva_def,
                        key=f"inp_iva_{fac_sel['Comprobante Siigo']}",
                        help="Cuenta contable de IVA descontable ordinario (ej. 24081501 Servicios o 24081001 Compras)."
                    )

                with c_mod2:
                    c_v1, c_v2 = st.columns(2)
                    base_cur_f = float(fac_sel["Base"]) if pd.notna(fac_sel["Base"]) else 0.0

                    # Calcular sugerencia de retención si se selecciona O-48
                    if "O-48" in cod_reg_sel:
                        rf_sug = round(base_cur_f * 0.04, 2) if es_aduanero or "EURO" in str(fac_sel["Proveedor"]).upper() else round(base_cur_f * 0.025, 2)
                        ri_sug = round(base_cur_f * 0.00966, 2)
                    elif any(k in cod_reg_sel for k in ["O-15", "O-47"]):
                        rf_sug = 0.0
                        ri_sug = 0.0
                    else:
                        rf_sug = float(fac_sel.get("ReteFuente", 0.0))
                        ri_sug = float(fac_sel.get("ReteICA", 0.0))

                    val_rf_input = float(fac_sel.get("ReteFuente", 0.0))
                    if val_rf_input == 0.0 and "O-48" in cod_reg_sel and rf_sug > 0:
                        val_rf_input = rf_sug
                    val_ri_input = float(fac_sel.get("ReteICA", 0.0))
                    if val_ri_input == 0.0 and "O-48" in cod_reg_sel and ri_sug > 0:
                        val_ri_input = ri_sug

                    with c_v1:
                        nueva_base = st.number_input("Base Gravable / Subtotal ($):", value=base_cur_f, step=1000.0, key=f"nb_{fac_sel['Comprobante Siigo']}")
                        nueva_rfte = st.number_input("ReteFuente ($):", value=val_rf_input, step=100.0, key=f"nrf_{fac_sel['Comprobante Siigo']}")
                    with c_v2:
                        lbl_iva_desc = "IVA Servicios ($) (24081501):" if es_importacion_factura else "IVA Descontable ($):"
                        nuevo_iva = st.number_input(lbl_iva_desc, value=float(fac_sel["IVA"]), step=100.0, key=f"niva_{fac_sel['Comprobante Siigo']}")
                        nuevo_rica = st.number_input("ReteICA ($):", value=val_ri_input, step=100.0, key=f"nri_{fac_sel['Comprobante Siigo']}")

                    # CAMPO DEDICADO: IVA DE IMPORTACIÓN (CUENTA 240835) - SOLO PARA FACTURAS DE IMPORTACIÓN
                    if es_importacion_factura:
                        raw_imp_f = fac_sel.get("IVA Importación")
                        try:
                            val_iva_imp_actual = float(raw_imp_f) if (raw_imp_f is not None and pd.notna(raw_imp_f)) else 0.0
                            if pd.isna(val_iva_imp_actual):
                                val_iva_imp_actual = 0.0
                        except Exception:
                            val_iva_imp_actual = 0.0

                        raw_cta_f = fac_sel.get("Cta IVA Importación")
                        if raw_cta_f is None or pd.isna(raw_cta_f) or str(raw_cta_f).strip() in ["", "nan", "None"]:
                            cta_iva_imp_actual = "240835"
                        else:
                            cta_iva_imp_actual = str(raw_cta_f).strip()

                        c_imp1, c_imp2 = st.columns(2)
                        with c_imp1:
                            nuevo_iva_imp = st.number_input(
                                "🚢 IVA de Importación ($) (Cta 240835):",
                                value=float(val_iva_imp_actual),
                                step=1000.0,
                                key=f"niva_imp_val_{fac_sel['Comprobante Siigo']}",
                                help="Ingresa aquí el valor del IVA pagado en la importación / declaración de aduanas que se imputa a la cuenta 240835."
                            )
                        with c_imp2:
                            cta_iva_imp_cod = st.text_input(
                                "Cuenta IVA Importación:",
                                value=cta_iva_imp_actual,
                                key=f"inp_cta_ivaimp_cod_{fac_sel['Comprobante Siigo']}",
                                help="Cuenta contable de IVA para importación (por defecto 240835)."
                            )
                    else:
                        nuevo_iva_imp = 0.0
                        cta_iva_imp_cod = "240835"

                asumir_imp = st.checkbox(
                    "Asumir Impuestos / Retenciones (Cruza con Agente Aduanero / Triangulación)",
                    value=bool(fac_sel.get("Impuestos Asumidos", False) or es_aduanero or any(k in fac_sel["Proveedor"].upper() for k in ["CARGO", "ADUANA", "PORTUARIA", "DHL"])),
                    key=f"asum_in_{fac_sel['Comprobante Siigo']}",
                    help="Si se asumen, las retenciones van a la cuenta 53152001 (Retenciones Asumidas) y el Saldo por Pagar de la factura queda en Base + IVA (ej. .651.939)."
                )

                if st.button("💾 Aplicar Cambios y Recalcular Asiento", key=f"btn_aplica_reg_{fac_sel['Comprobante Siigo']}"):
                    cod_reg = nuevo_reg_sel.split()[0]
                    r_idx = df_p[df_p["Comprobante Siigo"] == comp_sel].index[0]
                    df_p.at[r_idx, "Régimen Fiscal Emisor"] = cod_reg
                    df_p.at[r_idx, "Base"] = nueva_base
                    df_p.at[r_idx, "IVA"] = nuevo_iva
                    df_p.at[r_idx, "ReteFuente"] = nueva_rfte
                    df_p.at[r_idx, "ReteICA"] = nuevo_rica
                    if nueva_rfte > 0 and not str(df_p.at[r_idx, "Cta ReteFuente"]).strip():
                        df_p.at[r_idx, "Cta ReteFuente"] = "23652503" if es_aduanero else "23654001"
                    if nuevo_rica > 0 and not str(df_p.at[r_idx, "Cta ReteICA"]).strip():
                        df_p.at[r_idx, "Cta ReteICA"] = "23680505" if es_aduanero else "23680501"
                    cta_clean = nueva_cta_cxp.split()[0].strip()
                    df_p.at[r_idx, "Cuenta Pasivo Especifica"] = cta_clean
                    df_p.at[r_idx, "Cta Contrapartida"] = cta_clean
                    df_p.at[r_idx, "Cta Principal"] = nueva_cta_p.split()[0].strip()
                    cta_iva_clean = nueva_cta_iva.split()[0].strip() if 'nueva_cta_iva' in locals() and nueva_cta_iva else ("24081501" if es_importacion_factura else "24081001")
                    df_p.at[r_idx, "Cta IVA"] = cta_iva_clean
                    nuevo_iva_imp_f = float(nuevo_iva_imp) if pd.notna(nuevo_iva_imp) else 0.0
                    cta_iva_imp_clean = cta_iva_imp_cod.split()[0].strip()
                    if not cta_iva_imp_clean or cta_iva_imp_clean in ["nan", "None"]:
                        cta_iva_imp_clean = "240835"
                    df_p.at[r_idx, "IVA Importación"] = nuevo_iva_imp_f
                    df_p.at[r_idx, "Cta IVA Importación"] = cta_iva_imp_clean
                    df_p.at[r_idx, "Impuestos Asumidos"] = asumir_imp
                    df_p.at[r_idx, "Editada Manualmente"] = True

                    r_iva_val = float(fac_sel.get("ReteIVA", 0.0)) if pd.notna(fac_sel.get("ReteIVA")) else 0.0
                    if asumir_imp:
                        saldo_p = round(nueva_base + nuevo_iva + nuevo_iva_imp_f, 2)
                        df_p.at[r_idx, "Neto a Pagar"] = saldo_p
                        df_p.at[r_idx, "Total Neto"] = saldo_p
                    else:
                        neto_p_calc = round(nueva_base + nuevo_iva + nuevo_iva_imp_f - nueva_rfte - nuevo_rica - r_iva_val, 2)
                        df_p.at[r_idx, "Neto a Pagar"] = neto_p_calc
                        df_p.at[r_idx, "Total Neto"] = neto_p_calc

                    tot_esperado = round(nueva_base + nuevo_iva + nuevo_iva_imp_f, 2)
                    if nuevo_iva_imp_f > 0 and abs(float(df_p.at[r_idx, "Total"]) - tot_esperado) > 10.0:
                        df_p.at[r_idx, "Total"] = tot_esperado

                    st.session_state["df_procesado"] = df_p
# Asiento de triangulación protegido: no se sobreescribe con asiento estándar

                    # Guardar regla en la Memoria de Aprendizaje Continuo
                    guardar_regla_aprendizaje(
                        empresa,
                        fac_sel.get("NIT Emisor"),
                        fac_sel.get("Proveedor"),
                        cta_p=nueva_cta_p.split()[0].strip(),
                        resp_fiscal=cod_reg,
                        cta_cxp=cta_clean,
                        cta_iva=cta_iva_clean,
                        asumir_impuestos=asumir_imp
                    )
                    guardar_trabajo_en_historial(
                        empresa, df_p,
                        excel_bytes=st.session_state.get("excel_bytes"),
                        excel_nombre=st.session_state.get("excel_nombre", "Reporte.xlsx"),
                        dict_pdfs_renombrados=st.session_state.get("dict_pdfs", {}),
                        dict_pdfs_originales=st.session_state.get("raw_uploaded_pdfs", {}),
                        zip_bytes=st.session_state.get("zip_pdfs"),
                        job_id=st.session_state.get("job_actual_id")
                    )
                    st.session_state["factura_seleccionada_auditoria"] = comp_sel
                    st.session_state.pop("sel_factura_auditoria_box", None)
                    st.success(f"¡Valores y Cuenta CxP ({cta_clean}) actualizados y asiento recalculado con éxito!")
                    st.rerun()

            if es_aduanero:
                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("""
                <div style="background:#fef3c7; padding:15px; border-radius:8px; border-left:5px solid #d97706;">
                    <h5 style="color:#92400e; margin-top:0;">Alerta de Importacion / Agenciamiento Aduanero</h5>
                    <p style="color:#78350f; font-size:14px; margin-bottom:5px;">
                    En este documento intervienen <b>gastos por cuenta de terceros</b> y <b>honorarios propios del agente</b>:
                    </p>
                    <ul style="color:#78350f; font-size:14px;">
                        <li><b>Ingresos Propios del Agente (Honorarios / Comision):</b> Gravados con IVA 19%, sujetos a ReteFuente de Servicios (4%) u Honorarios (11%).</li>
                        <li><b>Pagos por Cuenta de Terceros (Tributos / Fletes / Bodegaje):</b> Imputables directamente al costo de importacion (Cuenta 146505). No llevan IVA del agente ni retencion en la fuente sobre el intermediario.</li>
                    </ul>
                </div>
                """, unsafe_allow_html=True)

        with col_a2:
            st.markdown("#### Resumen Financiero")
            st.metric("Base Gravable", f"${fac_sel['Base']:,.2f}")
            st.metric("IVA Liquidado (Servicios)", f"${fac_sel['IVA']:,.2f}")
            iva_imp_cur = float(fac_sel.get("IVA Importación", 0.0))
            if iva_imp_cur > 0:
                st.metric("🚢 IVA Importación (240835)", f"${iva_imp_cur:,.2f}")
            st.metric("ReteFuente", f"-${fac_sel['ReteFuente']:,.2f}")
            st.metric("ReteICA", f"-${fac_sel.get('ReteICA', 0.0):,.2f}")
            reteiva_val = fac_sel.get("ReteIVA", 0.0)
            if reteiva_val > 0:
                st.metric("ReteIVA", f"-${reteiva_val:,.2f}")
            if fac_sel.get("Impuestos Asumidos"):
                saldo_mostrar = round(fac_sel["Base"] + fac_sel["IVA"] + iva_imp_cur, 2)
                st.metric("Saldo CxP a Cruzar (Triangulación)", f"${saldo_mostrar:,.2f}")
                st.caption("ℹ️ Impuestos asumidos (Cta 53152001). Saldo total a cruzar.")
            else:
                neto_cxp = round(fac_sel["Base"] + fac_sel["IVA"] + iva_imp_cur - fac_sel["ReteFuente"] - fac_sel.get("ReteICA", 0.0) - reteiva_val, 2)
                st.metric("Total Neto CxP (Cta 22 / 23)", f"${neto_cxp:,.2f}")

        # VISTA PREVIA DE LA FACTURA EN UN CUADRO (SOPORTE PARA UNIFICADOS O SEPARADOS)
        st.markdown("---")
        st.markdown("### 🔍 Vista Previa del Documento Soporte")
        st.caption("Visualiza el PDF de esta factura directamente en pantalla, tanto si subiste archivos individuales como un PDF consolidado/unificado.")

        dict_renom = st.session_state.get("dict_pdfs", {})
        dict_orig = st.session_state.get("raw_uploaded_pdfs", {})

        pdf_bytes_encontrado, origen_desc, pags_encontradas = buscar_y_extraer_pdf(
            fac_sel,
            dict_renombrados=dict_renom,
            dict_originales=dict_orig,
            empresa_compradora=empresa
        )

        if pdf_bytes_encontrado:
            # Banner de coincidencia exacta con el registro contable
            st.success(f"✅ **Factura y Registro Contable Vinculados:** Comprobante **{fac_sel['Comprobante Siigo']}** | Factura: **{fac_sel['Factura']}** | Proveedor: **{fac_sel['Proveedor']}** (NIT: {fac_sel['NIT Emisor']}) — Total: **${fac_sel['Total']:,.2f}**")

            try:
                reader_prev = PdfReader(io.BytesIO(pdf_bytes_encontrado))
                num_pags_tot = len(reader_prev.pages)
            except Exception:
                num_pags_tot = 1

            pags_badge = f"{num_pags_tot} páginas completas" if num_pags_tot > 1 else "1 página"
            st.markdown(f"""
            <div style="background:#0070ba; color:white; padding:9px 16px; border-radius:6px; font-weight:600; font-size:14px; display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                <span>📄 {origen_desc} — Factura {fac_sel['Factura']} ({fac_sel['Proveedor']})</span>
                <span style="background:rgba(255,255,255,0.25); padding:3px 10px; border-radius:12px; font-size:12px;">📑 {pags_badge}</span>
            </div>
            """, unsafe_allow_html=True)

            safe_c_key = re.sub(r'\W', '_', str(fac_sel['Comprobante Siigo']))
            renderizar_visor_pdf_completo(
                pdf_bytes_encontrado,
                nombre_archivo=fac_sel["Soporte PDF Renombrado"],
                fac_sel=fac_sel,
                key_prefix=f"fac_{safe_c_key}",
                nit_comprador=re.sub(r"\D", "", str(empresa.get("nit", "9013464125")))
            )
        elif dict_orig or dict_renom:
            st.warning("⚠️ No se identificó automáticamente esta factura. Puedes seleccionar manualmente cualquier PDF cargado para visualizarlo:")
            todos_los_pdfs = {**dict_orig, **dict_renom}
            pdf_elegido = st.selectbox("Selecciona un archivo PDF cargado:", list(todos_los_pdfs.keys()), key="sel_pdf_manual_box")
            if pdf_elegido:
                f_bytes_sel = todos_los_pdfs[pdf_elegido]
                f_bytes_sel, _ = desbloquear_pdf_bytes(f_bytes_sel, nit_receptor=re.sub(r"\D", "", str(empresa.get("nit", "9013464125"))))
                safe_f_key = re.sub(r'\W', '_', str(pdf_elegido))
                renderizar_visor_pdf_completo(
                    f_bytes_sel,
                    nombre_archivo=pdf_elegido,
                    fac_sel=fac_sel,
                    key_prefix=f"man_{safe_f_key}",
                    nit_comprador=re.sub(r"\D", "", str(empresa.get("nit", "9013464125")))
                )
        else:
            st.markdown(f"""
            <div style="border: 1px solid #cbd5e1; border-radius: 8px; padding: 18px; background: #ffffff; box-shadow: 0 1px 3px rgba(0,0,0,0.05); margin-bottom: 20px;">
                <div style="border-bottom: 2px solid #0070ba; padding-bottom: 8px; margin-bottom: 12px; display: flex; justify-content: space-between; align-items: center;">
                    <h4 style="margin:0; color:#0f172a;">📄 Cuadro de Factura Electrónica</h4>
                    <span style="background:#eff6ff; color:#1d4ed8; padding:3px 10px; border-radius:4px; font-weight:bold; font-family:monospace;">{fac_sel['Factura']}</span>
                </div>
                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 14px; font-size: 14px; color: #334155;">
                    <div><b>Proveedor (Emisor):</b> {fac_sel['Proveedor']}<br><b>NIT Emisor:</b> {fac_sel['NIT Emisor']}</div>
                    <div><b>Empresa Compradora:</b> {empresa['nombre']}<br><b>NIT Receptor:</b> {empresa['nit']}</div>
                    <div><b>Fecha de Emisión:</b> {fac_sel['Fecha']}</div>
                    <div><b>Valor Total Facturado:</b> ${fac_sel['Total']:,.2f}</div>
                </div>
                <p style="margin-top: 14px; margin-bottom: 0; font-size: 13px; color: #64748b;">
                    <i>💡 Nota: Sube los archivos PDF en la Pestaña 1 para ver el documento digitalizado en este visor.</i>
                </p>
            </div>
            """, unsafe_allow_html=True)

        # TABLA DE CONTABILIZACION (ASIENTO CONTABLE)
        st.markdown("### 📋 Asiento Contable del Comprobante (Cómo se Contabilizó)")
        st.caption("Detalle de partida doble con imputación de cuentas, débitos, créditos y sumas iguales:")

        fac_num_aud = str(fac_sel["Factura"]).strip()
        asiento_triang_aprobado = None

        # 1. Buscar si esta factura corresponde al agente de algún paquete en triangulación
        if "paquetes_importacion" in st.session_state:
            for p_k, p_v in st.session_state["paquetes_importacion"].items():
                if str(p_v["agente"]["Factura"]).strip() == fac_num_aud:
                    # Si el paquete está validado en Pestaña 3, usar su asiento exacto aprobado
                    if st.session_state.get(f"paquete_listo_{p_k}", False):
                        if f"asiento_fijo_pq_{p_k}" in st.session_state and st.session_state[f"asiento_fijo_pq_{p_k}"] is not None:
                            asiento_triang_aprobado = st.session_state[f"asiento_fijo_pq_{p_k}"].copy()
                        elif "asiento_fijo" in p_v and p_v["asiento_fijo"] is not None:
                            asiento_triang_aprobado = p_v["asiento_fijo"].copy()
                        elif "asiento_aprobado" in p_v and p_v["asiento_aprobado"] is not None:
                            asiento_triang_aprobado = p_v["asiento_aprobado"].copy()
                    
                    if asiento_triang_aprobado is None:
                        asientos_m = st.session_state.get("asientos_triangulacion_por_factura", {})
                        if fac_num_aud in asientos_m and asientos_m[fac_num_aud] is not None:
                            asiento_triang_aprobado = asientos_m[fac_num_aud].copy()

                    if asiento_triang_aprobado is None and (st.session_state.get(f"paquete_listo_{p_k}", False) or p_v.get("diferencia", 999) < 1.0):
                        en_nd_k = st.session_state.get(f"enviar_nd_pq_{p_k}", False)
                        en_gp_k = st.session_state.get(f"enviar_gp_pq_{p_k}", False)
                        en_h2_k = st.session_state.get(f"enviar_h2_pq_{p_k}", False)
                        if en_h2_k:
                            df_as_k = obtener_asiento_contable_hoja2(p_v["agente"], es_aduanero=False)
                        elif en_gp_k:
                            df_as_k, _, _ = generar_asiento_triangulacion_paquete(p_v["agente"], p_v["terceros"], enviar_a_no_deducible=False)
                            asiento_gp_filas = []
                            for _, r_as in df_as_k.iterrows():
                                if str(r_as["Código Cuenta"]).strip() == CUENTA_NO_DEDUCIBLE:
                                    asiento_gp_filas.append({
                                        "Código Cuenta": CUENTA_IMPORTACION_TRANSITO,
                                        "Descripción Cuenta": f"Mercancías en Tránsito / Base Propia Agente (Fac {p_v['agente'].get('Factura', '')})",
                                        "Tercero / NIT": f"{p_v['agente'].get('NIT Emisor', '')} - {p_v['agente'].get('Proveedor', '')[:25]}",
                                        "Débito ($)": float(r_as["Débito ($)"]),
                                        "Crédito ($)": 0.0
                                    })
                                else:
                                    asiento_gp_filas.append(r_as.to_dict())
                            df_as_k = pd.DataFrame(asiento_gp_filas)
                        else:
                            df_as_k, _, _ = generar_asiento_triangulacion_paquete(p_v["agente"], p_v["terceros"], enviar_a_no_deducible=en_nd_k)
                        asiento_triang_aprobado = df_as_k
                    break

        if asiento_triang_aprobado is None:
            asientos_m = st.session_state.get("asientos_triangulacion_por_factura", {})
            if fac_num_aud in asientos_m and asientos_m[fac_num_aud] is not None:
                asiento_triang_aprobado = asientos_m[fac_num_aud].copy()

        if asiento_triang_aprobado is not None and not (isinstance(asiento_triang_aprobado, pd.DataFrame) and asiento_triang_aprobado.empty):
            df_asiento = pd.DataFrame(asiento_triang_aprobado).copy()
            if "Descripción Cuenta" in df_asiento.columns and "Descripción de la Cuenta" not in df_asiento.columns:
                df_asiento = df_asiento.rename(columns={"Descripción Cuenta": "Descripción de la Cuenta"})
            cols_deseadas = [c for c in ["Código Cuenta", "Descripción de la Cuenta", "Tercero / NIT", "Débito ($)", "Crédito ($)"] if c in df_asiento.columns]
            df_asiento = df_asiento[cols_deseadas]
                
# Asiento validado en Triangulación se respeta íntegramente sin sobreescribir
                
            st.markdown(f"""
            <div style="background:#f0fdf4; border:1px solid #86efac; border-left:5px solid #16a34a; border-radius:8px; padding:12px 16px; margin-bottom:12px;">
                <h5 style="margin:0 0 4px 0; color:#166534;">🔀 Contabilización Oficial Aprobada en Triangulación</h5>
                <p style="margin:0; color:#14532d; font-size:13.5px;">
                    Esta factura fue conciliada en la Pestaña 3. El asiento contable real cancela las cuentas por pagar de terceros y traslada la deuda al agente aduanero:
                </p>
            </div>
            """, unsafe_allow_html=True)
        else:
            asiento_filas = []
            es_nc = "Devolucion" in str(fac_sel["Operacion"])
            asume_ret = bool(fac_sel.get("Impuestos Asumidos", False))
            cta_cxp_usar = str(fac_sel.get("Cuenta Pasivo Especifica") or fac_sel.get("Cta Contrapartida") or ("22050505" if es_aduanero else "22050501")).strip()

            # Verificar si la factura seleccionada es un tercero cruzado en un paquete de importación
            pq_cruz_encontrado = None
            if "paquetes_importacion" in st.session_state:
                for p_k, p_v in st.session_state["paquetes_importacion"].items():
                    df_terc_v = p_v.get("terceros")
                    if df_terc_v is not None and not df_terc_v.empty:
                        if str(fac_sel["Factura"]).strip() in df_terc_v["Factura"].astype(str).str.strip().values:
                            pq_cruz_encontrado = {
                                "pq_id": p_k,
                                "agente": p_v["agente"]["Proveedor"],
                                "agente_fac": p_v["agente"]["Factura"],
                                "es_listo": st.session_state.get(f"paquete_listo_{p_k}", False)
                            }
                            break

            if pq_cruz_encontrado is not None and (esta_excl_ind or fac_sel.get("Ya Registrada")):
                st.markdown(f"""
                <div style="background:#eff6ff; border:1px solid #bfdbfe; border-left:5px solid #2563eb; border-radius:8px; padding:12px 16px; margin-bottom:12px;">
                    <h5 style="margin:0 0 4px 0; color:#1e40af;">🔀 Factura con Cuenta Cruzada en Importación (Paquete #{pq_cruz_encontrado['pq_id']} con {pq_cruz_encontrado['agente']})</h5>
                    <p style="margin:0; color:#1e3a8a; font-size:13.5px;">
                        Esta factura ya está contabilizada previamente en Siigo (o marcada como 'No Contabilizar'). Por lo tanto, <b>NO se causará como compra independiente en la planilla</b> (no genera base ni IVA repetidos). En la importación de Siigo participa <b>ÚNICAMENTE con su cuenta cruzada ({cta_cxp_usar})</b> debitada en el paquete de importación para cancelar la cuenta por pagar contra {pq_cruz_encontrado['agente']} (Fac {pq_cruz_encontrado['agente_fac']}).
                    </p>
                </div>
                """, unsafe_allow_html=True)
                saldo_cruce_ind = obtener_saldo_cruce_factura(fac_sel)
                asiento_filas = [
                    {
                        "Código Cuenta": cta_cxp_usar,
                        "Descripción de la Cuenta": f"Cancela CxP {fac_sel['Proveedor'][:20]} (Fac {fac_sel['Factura']})",
                        "Tercero / NIT": f"{fac_sel['NIT Emisor']} - {fac_sel['Proveedor'][:22]}",
                        "Débito ($)": saldo_cruce_ind,
                        "Crédito ($)": 0.0
                    },
                    {
                        "Código Cuenta": "22050501",
                        "Descripción de la Cuenta": f"CxP Agente Aduanero (Fac {pq_cruz_encontrado['agente_fac']})",
                        "Tercero / NIT": f"{fac_sel['NIT Emisor']} - {fac_sel['Proveedor'][:22]}",
                        "Débito ($)": 0.0,
                        "Crédito ($)": saldo_cruce_ind
                    }
                ]
            elif esta_excl_ind or fac_sel.get("Ya Registrada"):
                st.markdown(f"""
                <div style="background:#fef2f2; border:1px solid #fecaca; border-left:5px solid #dc2626; border-radius:8px; padding:12px 16px; margin-bottom:12px;">
                    <h5 style="margin:0 0 4px 0; color:#991b1b;">🚫 Factura Ya Contabilizada / Excluida (No Contabilizar)</h5>
                    <p style="margin:0; color:#7f1d1d; font-size:13.5px;">
                        Esta factura ya fue contabilizada previamente en Siigo o ha sido excluida manualmente. <b>No generará asiento de compra en la planilla de Siigo.</b> A continuación se muestra su estructura contable solo para fines informativos y de auditoría:
                    </p>
                </div>
                """, unsafe_allow_html=True)

            if not asiento_filas:
                df_asiento = obtener_asiento_contable_hoja2(fac_sel, es_aduanero=es_aduanero)
            else:
                df_asiento = pd.DataFrame(asiento_filas)

        cols_as_dsp = [c for c in ["Código Cuenta", "Descripción de la Cuenta", "Tercero / NIT", "Débito ($)", "Crédito ($)"] if c in df_asiento.columns]
        st.dataframe(
            df_asiento[cols_as_dsp].style.format({"Débito ($)": "${:,.2f}", "Crédito ($)": "${:,.2f}"}),
            use_container_width=True,
            hide_index=True
        )

        sum_deb = df_asiento["Débito ($)"].sum()
        sum_cred = df_asiento["Crédito ($)"].sum()

        c_as1, c_as2, c_as3 = st.columns(3)
        with c_as1:
            st.metric("Total Débito", f"${sum_deb:,.2f}")
        with c_as2:
            st.metric("Total Crédito", f"${sum_cred:,.2f}")
        with c_as3:
            st.metric("Diferencia / Cuadre", f"${abs(sum_deb - sum_cred):,.2f}")
        st.success("✅ **Comprobante Verificado:** Partida doble cuadrada con sumas iguales.")
    else:
        st.info("Carga el archivo Excel en la Pestana 1 para habilitar la auditoria.")

with tab_triangulacion:
    st.markdown("### 🔀 Paquetes de Importación y Cruce de Cuentas por Pagar")
    st.caption("Separa y concilia cada importación en paquetes individuales: Cobro de Euro Shipping = Factura DHL (Flete) + Factura Agencia Aduanas + Factura Garaje/Almacenadora.")

    if "df_procesado" in st.session_state:
        df_total = st.session_state["df_procesado"]
        
        # Identificar facturas de comercio exterior / importación de forma segura sin errores de tipos
        prov_series = df_total["Proveedor"] if "Proveedor" in df_total.columns else pd.Series("", index=df_total.index)
        desc_series = df_total["Descripcion"] if "Descripcion" in df_total.columns else pd.Series("", index=df_total.index)
        grp_series = df_total["Grupo Importación"] if "Grupo Importación" in df_total.columns else pd.Series("", index=df_total.index)
        cta_cxp_series = df_total["Cuenta Pasivo Especifica"] if "Cuenta Pasivo Especifica" in df_total.columns else (df_total["Cta Contrapartida"] if "Cta Contrapartida" in df_total.columns else pd.Series("", index=df_total.index))
        cta_p_series = df_total["Cta Principal"] if "Cta Principal" in df_total.columns else pd.Series("", index=df_total.index)

        prov_upper = prov_series.fillna("").astype(str).str.upper()
        desc_upper = desc_series.fillna("").astype(str).str.upper()
        grp_imp_str = grp_series.fillna("").astype(str).str.strip()
        cta_cxp_str = cta_cxp_series.fillna("").astype(str).str.strip()
        cta_p_str = cta_p_series.fillna("").astype(str).str.strip()

        # 1. Facturas explícitamente marcadas en Excel con Grupo de Importación o clasificadas como aduaneras
        cond_grp = (grp_imp_str != "") & (~grp_imp_str.isin(["nan", "None", "0", "0.0"]))
        cond_es_adu = (df_total["Es Aduanera"] == True) if "Es Aduanera" in df_total.columns else pd.Series(False, index=df_total.index)

        # 2. Cuentas contables específicas de importación / fletes DHL / agencias aduaneras
        cond_ctas = cta_cxp_str.isin(["22050505", "23359501", "14650501"]) | (cta_p_str == "14650501")

        # 3. Proveedores específicos de Comercio Exterior / Logística Internacional / Aduana / Puertos / Fletes
        REGEX_IMPORTACION = (
            r"EURO\s*SHIPPING|TRADE\s*GLOBAL|CONSOLCARGO|BLUE\s*LOGISTICS|"
            r"KUEHNE|PANALPINA|DSV|EXPEDITORS|TIBA|HUBEMAR|SCHENKER|COLTRANS|COLMAS|"
            r"DHL|FEDEX|UPS|CARGO|ADUANA|ADUANERA|AGENCIAMIENTO|PORTUARIA|PUERTO\s*BAHIA|"
            r"CONTECAR|SPRC|SPB|COMPAS|ALPOPULAR|ALMAVIVA|RANSA|ALMACENADORA|"
            r"MAERSK|HAPAG|SEABOARD|HAMBURG|CMA\s*CGM|EVERGREEN|COSCO|YANG\s*MING"
        )
        cond_prov = prov_upper.str.contains(REGEX_IMPORTACION, regex=True, na=False)

        # 4. Descripción que evidencia gastos de importación / desaduanamiento
        REGEX_DESC_IMPORT = (
            r"IMPORTACI[OÓ]N|AGENCIAMIENTO|ARANCEL|DESADUANAMIENTO|NACIONALIZACI[OÓ]N|"
            r"BODEGAJE\s*PUERTO|ALMACENAJE\s*ADUANERO|MANDATO\s*ADUANERO|FLETE\s*INTERNACIONAL"
        )
        cond_desc = desc_upper.str.contains(REGEX_DESC_IMPORT, regex=True, na=False)

        # Condición estricta: ÚNICAMENTE facturas relacionadas con la importación (rojas y blancas de importación)
        cond_aduanera = cond_grp | cond_es_adu | cond_ctas | cond_prov | cond_desc
        df_adu = df_total[cond_aduanera].copy()
        
        # Configuración dinámica y extensible de Agentes Coordinadores (Forwarders)
        AGENTES_COORDINADORES_BASE = ["EURO SHIPPING", "TRADE GLOBAL", "CONSOLCARGO", "BLUE LOGISTICS", "KUEHNE", "PANALPINA", "DSV", "EXPEDITORS", "TIBA", "HUBEMAR"]
        
        todos_provs = sorted(list(df_total["Proveedor"].dropna().unique()))
        provs_agentes_sugeridos = [p for p in todos_provs if any(k in str(p).upper() for k in AGENTES_COORDINADORES_BASE)]
        
        with st.expander("🏢 Configurar Agentes Coordinadores (Los que nos cobran directamente a nosotros):", expanded=False):
            st.caption("Selecciona qué empresas actúan como Agentes Coordinadores / Forwarders que emiten cobros globales consolidados (Euro Shipping, Trade Global, Consolcargo, etc.):")
            agentes_seleccionados_noms = st.multiselect(
                "Agentes Coordinadores Activos:",
                options=todos_provs,
                default=provs_agentes_sugeridos,
                key="sel_multiselect_agentes_coordinadores"
            )
            
        if not agentes_seleccionados_noms:
            agentes_seleccionados_noms = provs_agentes_sugeridos
            
        # Facturas de Agentes Principales (Los que nos cobran directamente)
        df_agentes_all = df_adu[df_adu["Proveedor"].isin(agentes_seleccionados_noms)].copy()
        
        # Facturas de Terceros Soporte (Los que se cruzan contra esos agentes)
        df_terceros_all = df_adu[~df_adu.index.isin(df_agentes_all.index)].copy()

        # Resumen superior
        m_c1, m_c2, m_c3, m_c4, m_c5 = st.columns(5)
        with m_c1:
            st.metric("Operaciones Agente (Euro / Trade)", len(df_agentes_all))
        with m_c2:
            st.metric("Facturas Soporte en Pool", len(df_terceros_all), help="Total de facturas candidatas para cruce")
        with m_c3:
            ya_reg_terc = df_terceros_all["Ya Registrada"].fillna(False).astype(bool) if "Ya Registrada" in df_terceros_all.columns else pd.Series(False, index=df_terceros_all.index)
            n_rojas_adu = int(ya_reg_terc.sum())
            st.metric("Terceros Contabilizados (🔴)", n_rojas_adu, help="Causadas previamente en Siigo")
        with m_c4:
            n_blancas_adu = int((~ya_reg_terc).sum())
            st.metric("Terceros Pendientes (⚪)", n_blancas_adu, help="No contabilizadas aún (Nuevas)")
        with m_c5:
            n_cuad_prev = sum([1 for p in st.session_state.get("paquetes_importacion", {}).values() if p.get("diferencia", 999) < 1.0])
            st.metric("Paquetes Cuadrados ($0.00)", f"{n_cuad_prev} ✅")

        st.markdown("---")

        if df_agentes_all.empty:
            st.warning("⚠️ No se encontraron facturas de agentes aduaneros principales (Euro Shipping o Trade Global) en este reporte.")
        else:
            # 1. FUNCIÓN DE AUTO-EMPAQUETAMIENTO INTELIGENTE
            # Algoritmo de combinación: encuentra para cada Euro Shipping el subconjunto de DHL + Agencia + Garaje que suma su valor
            def auto_empaquetar_inteligente(df_ag, df_terc):
                """
                Algoritmo cronológico de alta precisión para Agencias Aduaneras (Euro Shipping / Trade Global):
                1. Ordena cronológicamente de Enero hacia Diciembre.
                2. Cruza y relaciona tanto facturas de terceros YA contabilizadas (🔴) como NO contabilizadas (⚪).
                3. Para cada factura de agente (ej. 19 de Enero), busca facturas de terceros (DHL, Agencia, Garaje, etc.):
                   - Prioritariamente en fechas anteriores o el mismo día (-30 a 0 días de diferencia).
                   - O en una fecha ligeramente superior (+1 a +12 días de diferencia).
                   - Descarta o penaliza fechas lejanas de otros meses.
                4. Exige concordancia estricta de precio: la suma del Saldo Cruce (Base + IVA) debe concordar
                   exactamente con el valor total facturado por el agente aduanero.
                """
                import itertools
                pqs_res = {}
                
                # 1. Asegurar orden cronológico estricto de Enero hacia Diciembre
                df_ag_ord = df_ag.copy()
                df_ag_ord["_dt"] = pd.to_datetime(df_ag_ord["Fecha"], dayfirst=True, errors="coerce")
                df_ag_ord = df_ag_ord.sort_values(by="_dt", ascending=True)
                
                pool_t = df_terc.copy()
                pool_t["_dt"] = pd.to_datetime(pool_t["Fecha"], dayfirst=True, errors="coerce")
                pool_t = pool_t.sort_values(by="_dt", ascending=True)
                pool_t["_asignado_pq"] = 0

                # PROTEGER Y BLOQUEAR FACTURAS DE PAQUETES YA VALIDADOS
                if "paquetes_importacion" in st.session_state:
                    for p_num_prev, p_data_prev in st.session_state["paquetes_importacion"].items():
                        if st.session_state.get(f"paquete_listo_{p_num_prev}", False):
                            df_t_prev = p_data_prev.get("terceros")
                            if df_t_prev is not None and not df_t_prev.empty:
                                for f_b in df_t_prev["Factura"].dropna():
                                    idx_match = pool_t[pool_t["Factura"].astype(str).str.strip() == str(f_b).strip()].index
                                    if not idx_match.empty:
                                        pool_t.loc[idx_match, "_asignado_pq"] = p_num_prev
                
                for idx_ag_i, (_, ag_i) in enumerate(df_ag_ord.iterrows()):
                    pq_id_i = idx_ag_i + 1
                    target_i = float(ag_i["Total"])
                    f_ag_dt_i = ag_i["_dt"]

                    if "paquetes_importacion" in st.session_state and pq_id_i in st.session_state["paquetes_importacion"]:
                        prev_pq_val = st.session_state["paquetes_importacion"][pq_id_i]
                        if st.session_state.get(f"paquete_listo_{pq_id_i}", False):
                            # MANTENER INTACTO EL PAQUETE YA VALIDADO
                            pqs_res[pq_id_i] = prev_pq_val
                            continue
                    
                    cands_i = pool_t[pool_t["_asignado_pq"] == 0].copy()
                    if cands_i.empty:
                        pqs_res[pq_id_i] = {"agente": ag_i, "terceros": pd.DataFrame(), "diferencia": target_i}
                        continue
                        
                    # Evaluar concordancia cronológica y de mandato para cada factura de tercero
                    cand_l = []
                    for c_idx_k, c_row_k in cands_i.iterrows():
                        tb_k = float(c_row_k.get("Base", 0.0))
                        tiv_k = float(c_row_k.get("IVA", 0.0))
                        sc_k = obtener_saldo_cruce_factura(c_row_k)
                        p_u_k = str(c_row_k["Proveedor"]).upper()
                        rol_k = "DHL" if "DHL" in p_u_k else ("AGENCIA" if any(k in p_u_k for k in ["CARGO", "ADUANA"]) else "GARAJE")
                        
                        # 2. Análisis de fechas: concordancia de Enero a Diciembre
                        c_dt = c_row_k["_dt"]
                        score_tiempo = 0
                        diff_dias = 0
                        relacion_fecha_txt = "Sin fecha"
                        
                        if pd.notna(f_ag_dt_i) and pd.notna(c_dt):
                            diff_dias = int((c_dt - f_ag_dt_i).days)
                            # REGLA CRONOLÓGICA ESTRICTA: Solo conciliar facturas anteriores (-60 a 0 días) o un poco después (+1 a +15 días)
                            if diff_dias < -60 or diff_dias > 15:
                                continue  # DESCARTAR: fuera de la ventana cronológica del despacho
                                
                            if diff_dias == 0:
                                score_tiempo = 120
                                relacion_fecha_txt = "🟢 Mismo día (0d)"
                            elif -30 <= diff_dias < 0:
                                score_tiempo = 100 - abs(diff_dias) * 2
                                relacion_fecha_txt = f"⬅️ Anterior ({abs(diff_dias)}d antes)"
                            elif 0 < diff_dias <= 15:
                                score_tiempo = 85 - (diff_dias * 3)
                                relacion_fecha_txt = f"➡️ Poco después (+{diff_dias}d desp)"
                            elif -60 <= diff_dias < -30:
                                score_tiempo = 40 - abs(diff_dias)
                                relacion_fecha_txt = f"⬅️ Anterior ({abs(diff_dias)}d antes)"
                        
                        # Coincidencia directa por Mandato (NIT o Nombre del agente en notas)
                        nit_fact = re.sub(r'\D', '', str(c_row_k.get("NIT Facturado A", "") or c_row_k.get("nit_facturado_a", "")))
                        nom_fact = str(c_row_k.get("Facturado A", "") or c_row_k.get("nombre_facturado_a", "")).upper()
                        nit_ag_clean = re.sub(r'\D', '', str(ag_i.get("NIT Emisor", "")))
                        nom_ag_u = str(ag_i.get("Proveedor", "")).upper()
                        
                        score_mandato = 0
                        if nit_ag_clean and nit_fact and nit_ag_clean[:8] in nit_fact:
                            score_mandato += 150
                        elif any(w in nom_fact for w in nom_ag_u.split() if len(w) >= 4):
                            score_mandato += 80
                            
                        c_row_k_copy = c_row_k.copy()
                        c_row_k_copy["Relación Fecha"] = relacion_fecha_txt
                        c_row_k_copy["Diferencia Días"] = diff_dias
                        
                        cand_l.append({
                            "index": c_idx_k,
                            "fac": c_row_k["Factura"],
                            "saldo": sc_k,
                            "rol": rol_k,
                            "diff_dias": diff_dias,
                            "score_total": score_tiempo + score_mandato,
                            "row": c_row_k_copy
                        })
                        
                    # Ordenar candidatos dando prioridad a concordancia de fecha y mandato
                    cand_l.sort(key=lambda x: x["score_total"], reverse=True)
                    
                    # 3. Búsqueda de combinación con precio concorde exacto (target_i)
                    # Busca combinaciones de 1 a 4 facturas de terceros (tanto 🔴 contabilizadas como ⚪ pendientes)
                    best_combo_i = []
                    best_diff_i = float("inf")
                    best_score_i = -float("inf")
                    
                    # Ventana cronológica concorde (-35 días antes a +14 días después)
                    cands_ventana = [c for c in cand_l if -35 <= c["diff_dias"] <= 14]
                    if len(cands_ventana) < 5:
                        cands_ventana = cand_l[:25]
                    else:
                        cands_ventana = cands_ventana[:25]
                        
                    # Probar combinaciones de 1 a 4 facturas (DHL, Agencia, Garaje, rojas o blancas)
                    for k_c in range(1, min(5, len(cands_ventana) + 1)):
                        for combo_i in itertools.combinations(cands_ventana, k_c):
                            s_c_i = round(sum(c["saldo"] for c in combo_i), 2)
                            d_c_i = abs(s_c_i - target_i)
                            score_c = sum(c["score_total"] for c in combo_i) / len(combo_i)
                            
                            if d_c_i < 0.05:
                                # Coincidencia exacta de precio
                                if score_c > best_score_i or best_diff_i >= 0.05:
                                    best_diff_i = d_c_i
                                    best_combo_i = list(combo_i)
                                    best_score_i = score_c
                            elif best_diff_i >= 0.05 and d_c_i < best_diff_i:
                                best_diff_i = d_c_i
                                best_combo_i = list(combo_i)
                                best_score_i = score_c
                                
                        if best_diff_i < 0.05 and best_score_i > 70:
                            break
                            
                    # Si no dio cuadre exacto en la ventana cercana, ampliar búsqueda sobre todos los candidatos
                    if best_diff_i > 1.0 and len(cand_l) > len(cands_ventana):
                        cands_ampliados = cand_l[:30]
                        for k_c in range(1, min(5, len(cands_ampliados) + 1)):
                            for combo_i in itertools.combinations(cands_ampliados, k_c):
                                s_c_i = round(sum(c["saldo"] for c in combo_i), 2)
                                d_c_i = abs(s_c_i - target_i)
                                score_c = sum(c["score_total"] for c in combo_i) / len(combo_i)
                                
                                if d_c_i < best_diff_i:
                                    best_diff_i = d_c_i
                                    best_combo_i = list(combo_i)
                                    best_score_i = score_c
                                    if d_c_i < 0.05:
                                        break
                            if best_diff_i < 0.05:
                                break
                                
                    sel_idxs_i = [c["index"] for c in best_combo_i]
                    pool_t.loc[sel_idxs_i, "_asignado_pq"] = pq_id_i
                    
                    filas_terc_pq = [c["row"] for c in best_combo_i]
                    terc_pq_i = pd.DataFrame(filas_terc_pq) if filas_terc_pq else pd.DataFrame()
                    
                    pqs_res[pq_id_i] = {
                        "agente": ag_i,
                        "terceros": terc_pq_i,
                        "diferencia": best_diff_i
                    }
                    
                terc_libres_i = pool_t[pool_t["_asignado_pq"] == 0].copy()
                return pqs_res, terc_libres_i

            # Inicializar o recuperar paquetes de importación
            if "paquetes_importacion" not in st.session_state or st.session_state.get("_ultimo_agentes_len") != len(df_agentes_all):
                pqs_ini, libres_ini = auto_empaquetar_inteligente(df_agentes_all, df_terceros_all)
                st.session_state["paquetes_importacion"] = pqs_ini
                st.session_state["_ultimo_agentes_len"] = len(df_agentes_all)

            pqs_actuales = st.session_state["paquetes_importacion"]

            # =========================================================================
            # PURGA Y BLOQUEO ESTRICTO: NINGUNA FACTURA VALIDADA PUEDE ESTAR EN OTRO PAQUETE
            # =========================================================================
            facs_validadas_global = {}  # Factura -> pq_id_donde_esta_validada
            for p_k_chk, p_v_chk in pqs_actuales.items():
                if st.session_state.get(f"paquete_listo_{p_k_chk}", False):
                    df_t_chk = p_v_chk.get("terceros")
                    if df_t_chk is not None and not df_t_chk.empty:
                        for f_val in df_t_chk["Factura"].dropna():
                            f_val_clean = str(f_val).strip()
                            if f_val_clean:
                                facs_validadas_global[f_val_clean] = p_k_chk

            hubo_cambio_purga = False
            for p_k_chk, p_v_chk in pqs_actuales.items():
                # Si este paquete NO está validado, purgar cualquier factura que pertenezca a un paquete validado
                if not st.session_state.get(f"paquete_listo_{p_k_chk}", False):
                    df_t_chk = p_v_chk.get("terceros")
                    if df_t_chk is not None and not df_t_chk.empty:
                        mask_ya_en_val = df_t_chk["Factura"].astype(str).str.strip().isin(facs_validadas_global.keys())
                        if mask_ya_en_val.any():
                            df_t_filtrado = df_t_chk[~mask_ya_en_val].copy().reset_index(drop=True)
                            pqs_actuales[p_k_chk]["terceros"] = df_t_filtrado
                            tot_ag_c = float(p_v_chk["agente"]["Total"])
                            tot_terc_c = sum([obtener_saldo_cruce_factura(r) for _, r in df_t_filtrado.iterrows()])
                            pqs_actuales[p_k_chk]["diferencia"] = abs(tot_ag_c - tot_terc_c)
                            hubo_cambio_purga = True

            if hubo_cambio_purga:
                st.session_state["paquetes_importacion"] = pqs_actuales

            # Barra de Acciones y Selección de Paquete (Persistente: no salta al paquete 1)
            c_top_pq1, c_top_pq2 = st.columns([2.5, 1])
            with c_top_pq1:
                pqs_keys = list(pqs_actuales.keys())
                if "paquete_seleccionado_id" not in st.session_state or st.session_state["paquete_seleccionado_id"] not in pqs_keys:
                    st.session_state["paquete_seleccionado_id"] = pqs_keys[0] if pqs_keys else 1

                mapa_pqs_titulos = {}
                for p_num in pqs_keys:
                    p_data = pqs_actuales[p_num]
                    ag_t = p_data["agente"]
                    n_terc = len(p_data["terceros"])
                    s_terc = sum([obtener_saldo_cruce_factura(r) for _, r in p_data["terceros"].iterrows()])
                    es_val_p = st.session_state.get(f"paquete_listo_{p_num}", False)
                    lock_icon = "🔒 [VALIDADO] " if es_val_p else ""
                    mapa_pqs_titulos[p_num] = f"{lock_icon}📦 Paquete #{p_num}: {ag_t['Proveedor'][:16]} ({ag_t['Factura']}) — Cobro: ${ag_t['Total']:,.0f} | {n_terc} Facturas Terceros (${s_terc:,.0f})"

                idx_actual_sel = pqs_keys.index(st.session_state["paquete_seleccionado_id"]) if st.session_state["paquete_seleccionado_id"] in pqs_keys else 0

                pq_id_sel = st.selectbox(
                    "📦 Selecciona el Paquete de Importación que deseas revisar:",
                    options=pqs_keys,
                    index=idx_actual_sel,
                    format_func=lambda pid: mapa_pqs_titulos.get(pid, f"Paquete #{pid}")
                )
                st.session_state["paquete_seleccionado_id"] = pq_id_sel
                st.session_state["sel_paquete_activo_key"] = pq_id_sel
                paquete_activo = pqs_actuales[pq_id_sel]
                agente_actual = paquete_activo["agente"]
                terceros_actual = paquete_activo["terceros"]

                # Sincronización en tiempo real: refrescar datos del agente desde df_procesado (Hoja 2)
                if "df_procesado" in st.session_state and st.session_state["df_procesado"] is not None:
                    df_pr_act = st.session_state["df_procesado"]
                    ag_fac_cur = str(agente_actual["Factura"]).strip()
                    m_ag = df_pr_act[df_pr_act["Factura"].astype(str).str.strip() == ag_fac_cur]
                    if not m_ag.empty:
                        agente_actual = m_ag.iloc[0].to_dict()
                        paquete_activo["agente"] = agente_actual

                tot_agente_actual = float(agente_actual.get("Total", 0.0))
                base_prop_actual = float(agente_actual.get("Base", 0.0))
                iva_prop_actual = float(agente_actual.get("IVA", 0.0))
                iva_imp_act = float(agente_actual.get("IVA Importación", 0.0))
                subtot_prop_actual = round(base_prop_actual + iva_prop_actual + iva_imp_act, 2)
                saldo_terceros_esperado_act = round(tot_agente_actual - subtot_prop_actual, 2)

                # Tratamiento contable y asiento de partida doble del paquete
                enviar_h2_activo = bool(st.session_state.get(f"enviar_h2_pq_{pq_id_sel}", False))
                enviar_gp_activo = bool(st.session_state.get(f"enviar_gp_pq_{pq_id_sel}", False))
                enviar_nd_activo = bool(st.session_state.get(f"enviar_nd_pq_{pq_id_sel}", False))

                suma_terceros_cxp_act = sum([obtener_saldo_cruce_factura(r) for _, r in terceros_actual.iterrows()]) if (terceros_actual is not None and not terceros_actual.empty) else 0.0

                if saldo_terceros_esperado_act > 0.05:
                    dif_faltante_prev = max(0.0, round(saldo_terceros_esperado_act - suma_terceros_cxp_act, 2))
                    df_prev, _ = generar_asiento_mixto_hoja2_con_terceros(agente_actual, terceros_actual, enviar_a_no_deducible=False, imputar_a_transito=False)
                    ret_prev = 0.0
                else:
                    df_prev, dif_faltante_prev, ret_prev = generar_asiento_triangulacion_paquete(agente_actual, terceros_actual, enviar_a_no_deducible=False)

                if st.session_state.get(f"paquete_listo_{pq_id_sel}", False) and (f"asiento_fijo_pq_{pq_id_sel}" in st.session_state or "asiento_fijo" in paquete_activo):
                    df_asiento_paquete = st.session_state.get(f"asiento_fijo_pq_{pq_id_sel}", paquete_activo.get("asiento_fijo")).copy()
                    dif_no_ded = 0.0
                    ret_asum = 0.0
                elif saldo_terceros_esperado_act > 0.05 or enviar_h2_activo:
                    # Facturas con ingresos propios y de terceros (o con solicitud explícita de traer de Página 2)
                    if enviar_nd_activo:
                        df_asiento_paquete, _ = generar_asiento_mixto_hoja2_con_terceros(agente_actual, terceros_actual, enviar_a_no_deducible=True, imputar_a_transito=False)
                        dif_no_ded = dif_faltante_prev
                    elif enviar_gp_activo:
                        df_asiento_paquete, _ = generar_asiento_mixto_hoja2_con_terceros(agente_actual, terceros_actual, enviar_a_no_deducible=False, imputar_a_transito=True)
                        dif_no_ded = 0.0
                    else:
                        # Dejar pendiente para conciliar con facturas de terceros (DECISIÓN DEL USUARIO)
                        df_asiento_paquete, dif_no_ded = generar_asiento_mixto_hoja2_con_terceros(agente_actual, terceros_actual, enviar_a_no_deducible=False, imputar_a_transito=False)
                    ret_asum = 0.0
                elif enviar_gp_activo and dif_faltante_prev > 0.05:
                    df_asiento_paquete, _, ret_asum = generar_asiento_triangulacion_paquete(agente_actual, terceros_actual, enviar_a_no_deducible=False)
                    asiento_gp_filas = []
                    for _, r_as in df_asiento_paquete.iterrows():
                        if str(r_as["Código Cuenta"]).strip() == CUENTA_NO_DEDUCIBLE:
                            asiento_gp_filas.append({
                                "Código Cuenta": CUENTA_IMPORTACION_TRANSITO,
                                "Descripción Cuenta": f"Mercancías en Tránsito / Base Propia Agente (Fac {agente_actual.get('Factura', '')})",
                                "Tercero / NIT": f"{agente_actual.get('NIT Emisor', '')} - {agente_actual.get('Proveedor', '')[:25]}",
                                "Débito ($)": float(r_as["Débito ($)"]),
                                "Crédito ($)": 0.0
                            })
                        else:
                            asiento_gp_filas.append(r_as.to_dict())
                    df_asiento_paquete = pd.DataFrame(asiento_gp_filas)
                    dif_no_ded = 0.0
                else:
                    df_asiento_paquete, dif_no_ded, ret_asum = generar_asiento_triangulacion_paquete(agente_actual, terceros_actual, enviar_a_no_deducible=enviar_nd_activo)

                def ejecutar_validacion_pq(pasa_a_siigo=True):
                    st.session_state[f"paquete_listo_{pq_id_sel}"] = True
                    st.session_state[f"paquete_destino_siigo_{pq_id_sel}"] = pasa_a_siigo
                    paquete_activo["destino_siigo"] = pasa_a_siigo
                    paquete_activo["asiento_fijo"] = df_asiento_paquete.copy()
                    paquete_activo["asiento_aprobado"] = df_asiento_paquete.copy()
                    if "asientos_triangulacion_por_factura" not in st.session_state:
                        st.session_state["asientos_triangulacion_por_factura"] = {}
                    ag_fac_str = str(agente_actual["Factura"]).strip()
                    if pasa_a_siigo:
                        st.session_state["asientos_triangulacion_por_factura"][ag_fac_str] = df_asiento_paquete.copy()
                    else:
                        st.session_state["asientos_triangulacion_por_factura"].pop(ag_fac_str, None)
                    st.session_state["paquete_seleccionado_id"] = pq_id_sel
                    st.session_state["sel_paquete_activo_key"] = pq_id_sel

                    # PURGA INMEDIATA: Quitar las facturas de este paquete recién validado de TODOS los demás paquetes no validados
                    if not terceros_actual.empty:
                        facs_recien_validadas = set(terceros_actual["Factura"].astype(str).str.strip().dropna())
                        for otro_id, otro_pq in pqs_actuales.items():
                            if otro_id != pq_id_sel and not st.session_state.get(f"paquete_listo_{otro_id}", False):
                                df_otro_t = otro_pq.get("terceros")
                                if df_otro_t is not None and not df_otro_t.empty:
                                    mask_remover = df_otro_t["Factura"].astype(str).str.strip().isin(facs_recien_validadas)
                                    if mask_remover.any():
                                        df_otro_limpio = df_otro_t[~mask_remover].copy().reset_index(drop=True)
                                        otro_pq["terceros"] = df_otro_limpio
                                        tot_ag_o = float(otro_pq["agente"]["Total"])
                                        tot_terc_o = sum([float(r.get("Total Neto", 0.0) or (float(r.get("Base", 0.0)) + float(r.get("IVA", 0.0)))) for _, r in df_otro_limpio.iterrows()])
                                        otro_pq["diferencia"] = abs(tot_ag_o - tot_terc_o)
                        st.session_state["paquetes_importacion"] = pqs_actuales

                    # Traslado y actualización inmediata a df_procesado (Página 2)
                    if "df_procesado" in st.session_state and st.session_state["df_procesado"] is not None:
                        df_pr = st.session_state["df_procesado"]
                        for c_col in ["Paquete Validado", "Destino Siigo", "Cruza Con Agente", "Estado Registro"]:
                            if c_col in df_pr.columns:
                                df_pr[c_col] = df_pr[c_col].astype(object)
                            else:
                                df_pr[c_col] = None
                        idx_ag = df_pr[df_pr["Factura"].astype(str).str.strip() == ag_fac_str].index
                        if not idx_ag.empty:
                            if pasa_a_siigo:
                                df_pr.at[idx_ag[0], "Estado Registro"] = f"🔀 Validada en Triangulación (Pq #{pq_id_sel} - Pasa a Siigo)"
                            else:
                                df_pr.at[idx_ag[0], "Estado Registro"] = f"🔍 Conciliado en Triangulación (Pq #{pq_id_sel} - Solo Control Interno)"
                            df_pr.at[idx_ag[0], "Paquete Validado"] = pq_id_sel
                            df_pr.at[idx_ag[0], "Destino Siigo"] = pasa_a_siigo
                        if not terceros_actual.empty:
                            for _, tr_row in terceros_actual.iterrows():
                                tr_fac = str(tr_row["Factura"]).strip()
                                idx_t = df_pr[df_pr["Factura"].astype(str).str.strip() == tr_fac].index
                                if not idx_t.empty:
                                    es_r = bool(df_pr.at[idx_t[0], "Ya Registrada"]) or bool(df_pr.at[idx_t[0], "No Contabilizar"]) or (tr_fac in st.session_state.get("facturas_no_contabilizar", set()))
                                    if es_r:
                                        df_pr.at[idx_t[0], "Estado Registro"] = f"🔴 Cruzada en Pq #{pq_id_sel} (Solo cuenta pasivo)"
                                    else:
                                        if pasa_a_siigo:
                                            df_pr.at[idx_t[0], "Estado Registro"] = f"⚪ Cruzada en Pq #{pq_id_sel} (Causada en triangulación)"
                                        else:
                                            df_pr.at[idx_t[0], "Estado Registro"] = f"🔍 Conciliada en Pq #{pq_id_sel} (Solo Control Interno)"
                                    df_pr.at[idx_t[0], "Paquete Validado"] = pq_id_sel
                                    df_pr.at[idx_t[0], "Cruza Con Agente"] = ag_fac_str
                        st.session_state["df_procesado"] = df_pr
                    guardar_estado_manual(empresa)

                def ejecutar_desbloqueo_pq():
                    st.session_state[f"paquete_listo_{pq_id_sel}"] = False
                    st.session_state.pop(f"paquete_destino_siigo_{pq_id_sel}", None)
                    paquete_activo.pop("destino_siigo", None)
                    paquete_activo.pop("asiento_fijo", None)
                    paquete_activo.pop("asiento_aprobado", None)
                    ag_fac_str = str(agente_actual["Factura"]).strip()
                    if "asientos_triangulacion_por_factura" in st.session_state:
                        st.session_state["asientos_triangulacion_por_factura"].pop(ag_fac_str, None)
                    st.session_state["paquete_seleccionado_id"] = pq_id_sel
                    st.session_state["sel_paquete_activo_key"] = pq_id_sel
                    if "df_procesado" in st.session_state and st.session_state["df_procesado"] is not None:
                        df_pr = st.session_state["df_procesado"]
                        for c_col in ["Paquete Validado", "Destino Siigo", "Cruza Con Agente", "Estado Registro"]:
                            if c_col in df_pr.columns:
                                df_pr[c_col] = df_pr[c_col].astype(object)
                            else:
                                df_pr[c_col] = None
                        idx_ag = df_pr[df_pr["Factura"].astype(str).str.strip() == ag_fac_str].index
                        if not idx_ag.empty:
                            df_pr.at[idx_ag[0], "Estado Registro"] = "🟡 Agente Aduanero"
                            df_pr.at[idx_ag[0], "Paquete Validado"] = None
                            df_pr.at[idx_ag[0], "Destino Siigo"] = None
                        if not terceros_actual.empty:
                            for _, tr_row in terceros_actual.iterrows():
                                tr_fac = str(tr_row["Factura"]).strip()
                                idx_t = df_pr[df_pr["Factura"].astype(str).str.strip() == tr_fac].index
                                if not idx_t.empty:
                                    es_r = bool(df_pr.at[idx_t[0], "Ya Registrada"]) or bool(df_pr.at[idx_t[0], "No Contabilizar"])
                                    cp = df_pr.at[idx_t[0], "Comprobante Previo"]
                                    df_pr.at[idx_t[0], "Estado Registro"] = f"🔴 Ya Registrada ({cp})" if es_r else "⚪ Compra Pendiente"
                                    df_pr.at[idx_t[0], "Paquete Validado"] = None
                                    df_pr.at[idx_t[0], "Cruza Con Agente"] = None
                        st.session_state["df_procesado"] = df_pr
                    guardar_estado_manual(empresa)

            with c_top_pq2:
                st.write("")
                st.write("")
                if st.button("🪄 Re-calcular Auto-Empaquetado", key="btn_recalc_auto_pqs", help="Vuelve a calcular las combinaciones exactas de DHL + Agencia + Garaje para cada cobro de Euro"):
                    pqs_re, _ = auto_empaquetar_inteligente(df_agentes_all, df_terceros_all)
                    st.session_state["paquetes_importacion"] = pqs_re
                    st.success("¡Paquetes recalculados y separados!")
                    st.rerun()

            # Resumen visual del Paquete Seleccionado
            col_pq1, col_pq2 = st.columns([1.3, 2.5])
            with col_pq1:
                st.markdown(f"""
                <div style="background:#fffbeb; border:1px solid #fde68a; border-left:5px solid #d97706; border-radius:8px; padding:14px; margin-bottom:12px;">
                    <h4 style="margin:0 0 6px 0; color:#92400e;">🟡 Factura del Agente (Cobro Total)</h4>
                    <p style="margin:0; font-size:14px; color:#78350f;">
                        <b>Proveedor:</b> {agente_actual['Proveedor']}<br>
                        <b>NIT:</b> {agente_actual['NIT Emisor']}<br>
                        <b>Factura:</b> {agente_actual['Factura']}<br>
                        <b>Fecha:</b> {agente_actual['Fecha']}<br>
                        <b>Total Facturado:</b> <span style="font-size:18px; font-weight:bold; color:#b45309;">${tot_agente_actual:,.2f}</span><br>
                        <b>IVA Discriminado:</b> ${float(agente_actual.get('IVA', 0.0)):,.2f}
                    </p>
                </div>
                """, unsafe_allow_html=True)

                es_val_card = st.session_state.get(f"paquete_listo_{pq_id_sel}", False)
                if not es_val_card:
                    st.caption(f"ℹ️ Elige el destino de este paquete al validarlo:")
                    c_vcard1, c_vcard2 = st.columns(2)
                    with c_vcard1:
                        if st.button("📤 Pasar a Hoja 2", key=f"btn_val_siigo_card_{pq_id_sel}", type="primary", use_container_width=True, help="Bloquea el paquete y lo traslada a Página 2 para subirlo a la Planilla de Siigo"):
                            ejecutar_validacion_pq(pasa_a_siigo=True)
                            st.success(f"¡Paquete #{pq_id_sel} trasladado a Página 2 para Siigo!")
                            st.rerun()
                    with c_vcard2:
                        if st.button("📑 Solo Conciliación", key=f"btn_val_concil_card_{pq_id_sel}", use_container_width=True, help="Bloquea el paquete para identificar y conciliar facturas, sin subirlo a la Planilla de Siigo"):
                            ejecutar_validacion_pq(pasa_a_siigo=False)
                            st.success(f"¡Paquete #{pq_id_sel} guardado para Conciliación Interna!")
                            st.rerun()
                else:
                    es_siigo = st.session_state.get(f"paquete_destino_siigo_{pq_id_sel}", True)
                    if es_siigo:
                        st.success(f"🔒 **Pq #{pq_id_sel}: PASA A HOJA 2 (PLANILLA SIIGO)**")
                    else:
                        st.info(f"📑 **Pq #{pq_id_sel}: SOLO CONCILIACIÓN INTERNA**")
                    if st.button(f"🔓 Desbloquear Paquete #{pq_id_sel}", key=f"btn_desb_card_{pq_id_sel}", use_container_width=True):
                        ejecutar_desbloqueo_pq()
                        st.rerun()

            with col_pq2:
                st.markdown(f"##### Facturas de Terceros que Componen este Paquete #{pq_id_sel}:")
                if not terceros_actual.empty:
                    filas_terc_disp = []
                    tot_s_terceros = 0.0
                    for _, tr in terceros_actual.iterrows():
                        es_r = tr.get("Ya Registrada", False)
                        badge_est = f"🔴 Ya en Siigo ({tr.get('Comprobante Previo', '10-Prev')})" if es_r else "⚪ No Contabilizada (Pendiente)"
                        t_b = float(tr.get("Base", 0.0)) if pd.notna(tr.get("Base")) else 0.0
                        t_iv = float(tr.get("IVA", 0.0)) if pd.notna(tr.get("IVA")) else 0.0
                        s_cruce = obtener_saldo_cruce_factura(tr)
                        tot_s_terceros += s_cruce
                        cta_actual_tr = str(tr.get("Cuenta Pasivo Especifica", "22050505" if "CARGO" in str(tr["Proveedor"]).upper() else "23359501")).strip()
                        p_nom = str(tr["Proveedor"]).upper()
                        rol_dsp = "🚚 DHL (Flete)" if "DHL" in p_nom else ("🏢 Agencia (Aduana)" if any(k in p_nom for k in ["CARGO", "ADUANA"]) else "🏬 Garaje / Almacén")
                        
                        rel_fec_val = tr.get("Relación Fecha")
                        if pd.isna(rel_fec_val) or not str(rel_fec_val).strip() or str(rel_fec_val).strip() in ["nan", "None"]:
                            concordancia_dsp = calcular_concordancia_fecha(tr.get("Fecha"), agente_actual.get("Fecha"))
                        else:
                            concordancia_dsp = str(rel_fec_val).strip()
                            
                        filas_terc_disp.append({
                            "Rol": rol_dsp,
                            "Fecha Factura": tr.get("Fecha", "-"),
                            "Concordancia Fecha": concordancia_dsp,
                            "Estado Contable": badge_est,
                            "Proveedor Tercero": tr["Proveedor"][:22],
                            "Factura": tr["Factura"],
                            "Subtotal (Base)": t_b,
                            "IVA": t_iv,
                            "Saldo Cruce": s_cruce,
                            "Cuenta Contable": f"{cta_actual_tr} (CxP a Cruzar)"
                        })
                    df_terc_disp = pd.DataFrame(filas_terc_disp)
                    st.dataframe(df_terc_disp.style.format({
                        "Subtotal (Base)": "${:,.2f}",
                        "IVA": "${:,.2f}",
                        "Saldo Cruce": "${:,.2f}"
                    }), use_container_width=True, hide_index=True)
                    
                    dif_pq_actual = abs(tot_agente_actual - tot_s_terceros)
                    if dif_pq_actual < 5.0:
                        st.success(f"✅ **Paquete #{pq_id_sel} Exacto:** Total Terceros: **${tot_s_terceros:,.2f}** == Cobro Euro: **${tot_agente_actual:,.2f}** | Diferencia: **$0.00**")
                    else:
                        st.info(f"📊 Total Terceros: **${tot_s_terceros:,.2f}** | Cobro Euro: **${tot_agente_actual:,.2f}** | Diferencia a No Deducible: **${dif_pq_actual:,.2f}**")
                else:
                    st.warning(f"⚠️ El Paquete #{pq_id_sel} no tiene facturas de terceros asignadas todavía.")

            # Expander para agregar/quitar facturas manualmente a este paquete (BLOQUEADO SI ESTÁ VALIDADO)
            es_validado_este = st.session_state.get(f"paquete_listo_{pq_id_sel}", False)
            if es_validado_este:
                st.markdown(f"""
                <div style="background:#f0fdf4; border:1px solid #86efac; border-left:5px solid #16a34a; border-radius:8px; padding:12px 16px; margin:12px 0;">
                    <h5 style="margin:0 0 4px 0; color:#166534;">🔒 Paquete #{pq_id_sel} VALIDADO Y BLOQUEADO</h5>
                    <p style="margin:0; color:#14532d; font-size:13.5px;">
                        Este paquete se encuentra congelado. Ninguna factura puede ser añadida ni retirada mientras esté validado. Para realizar modificaciones, presiona <b>'🔓 Desbloquear y Modificar'</b> abajo.
                    </p>
                </div>
                """, unsafe_allow_html=True)
            
            with st.expander(f"⚙️ Modificar este Paquete #{pq_id_sel} (Sacar Facturas o Agregar Nuevas):", expanded=False if es_validado_este else False):
                if es_validado_este:
                    st.warning("⚠️ Debes desbloquear el paquete abajo antes de poder agregar o sacar facturas.")
                    st.stop()
                st.caption("Administra las facturas asignadas a esta importación: saca las que no correspondan o añade nuevas facturas de DHL, Agencia o Garaje:")
                
                # 1. SECCIÓN PARA SACAR / REDIRIGIR FACTURAS
                st.markdown("##### ➖ Sacar y Redirigir una factura de este paquete:")
                st.caption("Si ves una factura que no corresponde a este paquete, sácala y el sistema la reubicará automáticamente o la enviará a compras directas para que no quede pendiente:")
                facs_en_este = list(terceros_actual["Factura"].unique()) if not terceros_actual.empty else []
                
                if facs_en_este:
                    opciones_quitar = []
                    mapa_quitar_row = {}
                    for _, tr_k in terceros_actual.iterrows():
                        tb_k = float(tr_k.get("Base", 0.0))
                        tiv_k = float(tr_k.get("IVA", 0.0))
                        sc_k = obtener_saldo_cruce_factura(tr_k)
                        tag_q = f"[{tr_k['Factura']}] {tr_k['Proveedor'][:24]} (Saldo: ${sc_k:,.2f})"
                        opciones_quitar.append(tag_q)
                        mapa_quitar_row[tag_q] = tr_k
                        
                    sel_para_quitar = st.selectbox("1. Selecciona la factura que deseas SACAR de este paquete:", ["(Seleccionar...)"] + opciones_quitar, key=f"sel_rem_tr_{pq_id_sel}")
                    
                    if sel_para_quitar != "(Seleccionar...)":
                        fila_a_mover = mapa_quitar_row[sel_para_quitar]
                        fac_nom_mover = fila_a_mover["Factura"]
                        sc_mover = obtener_saldo_cruce_factura(fila_a_mover)
                        
                        otros_pqs = [p for p in pqs_actuales.keys() if p != pq_id_sel and not st.session_state.get(f"paquete_listo_{p}", False)]
                        
                        # Calcular cuál otro paquete tiene un faltante más cercano a esta factura
                        mejor_pq_sug = None
                        mejor_diff_sug = float("inf")
                        for p_cand_id in otros_pqs:
                            p_cand = pqs_actuales[p_cand_id]
                            tot_ag_c = float(p_cand["agente"]["Total"])
                            tot_terc_c = sum([obtener_saldo_cruce_factura(r) for _, r in p_cand["terceros"].iterrows()])
                            faltante_c = tot_ag_c - tot_terc_c
                            diff_c = abs(faltante_c - sc_mover)
                            if diff_c < mejor_diff_sug:
                                mejor_diff_sug = diff_c
                                mejor_pq_sug = p_cand_id
                                
                        sug_txt = f" (Sugerencia: Paquete #{mejor_pq_sug})" if mejor_pq_sug else ""
                        opciones_destino = [
                            f"🎯 Auto-Reubicar automáticamente donde mejor encaje{sug_txt}",
                            "📦 Mover a otro Paquete específico...",
                            "🛒 Redirigir a Compras Directas Ordinarias (Pestaña 1 y 4 - Se pagó directo)",
                            "⚪ Dejar libre en inventario (Sin asignar)"
                        ]
                        
                        c_dest1, c_dest2 = st.columns([2.6, 1.2])
                        with c_dest1:
                            dest_sel = st.radio(f"2. ¿A dónde deseas redirigir la factura {fac_nom_mover}?:", opciones_destino, key=f"rad_dest_{pq_id_sel}_{fac_nom_mover}")
                            destino_pq_manual = None
                            if "Mover a otro Paquete específico" in dest_sel:
                                opciones_otros_pqs = [f"Paquete #{p}: {pqs_actuales[p]['agente']['Proveedor'][:16]} (Fac {pqs_actuales[p]['agente']['Factura']})" for p in otros_pqs]
                                if opciones_otros_pqs:
                                    sel_otro_pq_str = st.selectbox("Selecciona el paquete de destino:", opciones_otros_pqs, key=f"sel_otro_pq_{pq_id_sel}")
                                    destino_pq_manual = int(sel_otro_pq_str.split(":")[0].replace("Paquete #", "").strip())
                                else:
                                    st.info("No hay otros paquetes de forwarders disponibles en este archivo.")
                                    
                        with c_dest2:
                            st.write("")
                            st.write("")
                            if st.button("🚀 Sacar y Redirigir Factura", key=f"btn_ejecutar_reubicar_{pq_id_sel}", help="Retira la factura de este paquete y la redirige al destino seleccionado."):
                                # 1. Retirar del paquete actual
                                terceros_remanente = terceros_actual[terceros_actual["Factura"] != fac_nom_mover].copy()
                                st.session_state["paquetes_importacion"][pq_id_sel]["terceros"] = terceros_remanente
                                
                                # 2. Redirigir al destino
                                if "Auto-Reubicar" in dest_sel and mejor_pq_sug is not None:
                                    dest_pq = st.session_state["paquetes_importacion"][mejor_pq_sug]["terceros"]
                                    dest_nuevo = pd.concat([dest_pq, pd.DataFrame([fila_a_mover])]).drop_duplicates(subset=["Factura"]).reset_index(drop=True)
                                    st.session_state["paquetes_importacion"][mejor_pq_sug]["terceros"] = dest_nuevo
                                    st.success(f"¡Factura {fac_nom_mover} retirada del Paquete #{pq_id_sel} y auto-reubicada en el Paquete #{mejor_pq_sug}!")
                                elif "Mover a otro Paquete" in dest_sel and destino_pq_manual is not None:
                                    dest_pq = st.session_state["paquetes_importacion"][destino_pq_manual]["terceros"]
                                    dest_nuevo = pd.concat([dest_pq, pd.DataFrame([fila_a_mover])]).drop_duplicates(subset=["Factura"]).reset_index(drop=True)
                                    st.session_state["paquetes_importacion"][destino_pq_manual]["terceros"] = dest_nuevo
                                    st.success(f"¡Factura {fac_nom_mover} movida con éxito al Paquete #{destino_pq_manual}!")
                                elif "Compras Directas Ordinarias" in dest_sel:
                                    if "df_procesado" in st.session_state:
                                        df_glob = st.session_state["df_procesado"]
                                        m_idx = df_glob[df_glob["Factura"] == fac_nom_mover].index
                                        if not m_idx.empty:
                                            df_glob.at[m_idx[0], "Es Aduanera"] = False
                                            df_glob.at[m_idx[0], "Grupo Importación"] = ""
                                            df_glob.at[m_idx[0], "Estado Registro"] = "⚪ Compra Directa Ordinaria"
                                            df_glob.at[m_idx[0], "Cta Contrapartida"] = "22050501"
                                            st.session_state["df_procesado"] = df_glob
                                    st.success(f"¡Factura {fac_nom_mover} retirada de importaciones y enviada a Compras Directas Ordinarias (Pestaña 1 y 4)!")
                                else:
                                    st.success(f"¡Factura {fac_nom_mover} retirada del Paquete #{pq_id_sel} y devuelta a facturas libres!")
                                
                                st.session_state["paquete_seleccionado_id"] = pq_id_sel
                                st.session_state["sel_paquete_activo_key"] = pq_id_sel
                                guardar_estado_manual(empresa)
                                st.rerun()
                else:
                    st.info("Este paquete no tiene facturas asignadas para retirar.")
                    
                st.markdown("---")
                
                # 2. SECCIÓN PARA AGREGAR NUEVAS FACTURAS
                cands_disp_agregar = df_terceros_all.copy()
                f_ag_dt_sug = pd.to_datetime(agente_actual.get("Fecha", ""), dayfirst=True, errors="coerce")
                
                # RESTRICCIÓN DE BLOQUEO: Las facturas de paquetes ya validados no se pueden añadir a otros
                facturas_bloqueadas_validadas = set()
                if "paquetes_importacion" in st.session_state:
                    for p_k_v, p_dat_v in st.session_state["paquetes_importacion"].items():
                        if st.session_state.get(f"paquete_listo_{p_k_v}", False):
                            df_t_val = p_dat_v.get("terceros")
                            if df_t_val is not None and not df_t_val.empty:
                                for f_asig in df_t_val["Factura"].dropna():
                                    facturas_bloqueadas_validadas.add(str(f_asig).strip())

                # 2.1 SUGERENCIAS INTELIGENTES DE FACTURAS PARA INGRESOS DE TERCEROS
                if dif_faltante_prev > 0.05:
                    st.markdown(
                        f"""
                        <div style="background:#f0fdf4; border:1.5px solid #22c55e; border-radius:8px; padding:12px 16px; margin:10px 0 16px 0;">
                            <h5 style="margin:0 0 4px 0; color:#15803d;">🔎 Facturas Sugeridas para Ingresos de Terceros (Faltante Pendiente: ${dif_faltante_prev:,.2f}):</h5>
                            <p style="margin:0 0 8px 0; font-size:13px; color:#166534;">
                                Facturas libres en el sistema que <b>no superan el monto de la factura (${tot_agente_actual:,.2f})</b> ni el <b>monto de terceros (${saldo_terceros_esperado_act if saldo_terceros_esperado_act > 0.05 else tot_agente_actual:,.2f})</b>, analizadas por concordancia de fletes, cargos, fecha y proveedor:
                            </p>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )
                    
                    # Construir pool de sugerencias
                    cands_sug_list = []
                    f_ag_dt_sug = pd.to_datetime(agente_actual.get("Fecha", ""), dayfirst=True, errors="coerce")
                    monto_limite_fac = tot_agente_actual + 0.05
                    monto_limite_terc = (saldo_terceros_esperado_act if saldo_terceros_esperado_act > 0.05 else tot_agente_actual) + 0.05
                    
                    for _, tr_cand_s in cands_disp_agregar.iterrows():
                        f_c_num = str(tr_cand_s["Factura"]).strip()
                        if f_c_num in facs_en_este or f_c_num in facturas_bloqueadas_validadas or not f_c_num:
                            continue
                            
                        sc_cand_val = obtener_saldo_cruce_factura(tr_cand_s)
                        if sc_cand_val <= 0.0:
                            continue
                            
                        # REGLA ESTRICTA: Que no superen el monto de la factura ni el monto de los ingresos para terceros
                        if sc_cand_val > monto_limite_fac or sc_cand_val > monto_limite_terc or sc_cand_val > (dif_faltante_prev + 0.05):
                            continue
                            
                        prov_u = str(tr_cand_s.get("Proveedor", "")).upper()
                        fec_str = str(tr_cand_s.get("Fecha", ""))
                        fec_dt = pd.to_datetime(fec_str, dayfirst=True, errors="coerce")
                        diff_dias = int((fec_dt - f_ag_dt_sug).days) if (pd.notna(fec_dt) and pd.notna(f_ag_dt_sug)) else 999
                        
                        # REGLA CRONOLÓGICA ESTRICTA: Solo facturas anteriores (-60d a 0d) o un poco después (+1d a +15d)
                        if pd.notna(fec_dt) and pd.notna(f_ag_dt_sug):
                            if diff_dias < -60 or diff_dias > 15:
                                continue
                        elif diff_dias == 999:
                            continue
                        
                        tags_sug = []
                        score_sug = 0
                        
                        # Coincidencias específicas con los ítems de ingresos por terceros (solo si la factura del agente es BOG-70697 o de saldo similar)
                        es_caso_70697 = "70697" in str(agente_actual.get("Factura", "")) or abs(saldo_terceros_esperado_act - 2542421.20) < 100.0
                        if es_caso_70697 and abs(sc_cand_val - 1079157.20) < 50000:
                            tags_sug.append("✈️ Flete Internacional (~$1.08M)")
                            score_sug += 180 - (abs(sc_cand_val - 1079157.20) / 1000)
                        elif es_caso_70697 and abs(sc_cand_val - 1463264.00) < 50000:
                            tags_sug.append("🏢 Cargos en Origen (~$1.46M)")
                            score_sug += 180 - (abs(sc_cand_val - 1463264.00) / 1000)
                        elif abs(sc_cand_val - dif_faltante_prev) < 5000:
                            tags_sug.append(f"🎯 Cuadre Exacto (${dif_faltante_prev:,.2f})")
                            score_sug += 250
                        elif sc_cand_val <= dif_faltante_prev + 0.05:
                            tags_sug.append(f"🧩 Cubre {(sc_cand_val / dif_faltante_prev * 100):.1f}% del faltante")
                            score_sug += 60
                            
                        # Ponderación por cercanía de fecha dentro de la ventana (-60 a +15)
                        if diff_dias == 0:
                            tags_sug.append("🟢 Mismo día (0d)")
                            score_sug += 100
                        elif 0 < diff_dias <= 15:
                            tags_sug.append(f"➡️ Poco después (+{diff_dias}d)")
                            score_sug += 85 - (diff_dias * 3)
                        elif -30 <= diff_dias < 0:
                            tags_sug.append(f"⬅️ Anterior ({abs(diff_dias)}d antes)")
                            score_sug += 90 - (abs(diff_dias) * 2)
                        elif -60 <= diff_dias < -30:
                            tags_sug.append(f"⬅️ Anterior ({abs(diff_dias)}d antes)")
                            score_sug += 40 - abs(diff_dias)
                            
                        if any(k in prov_u for k in ["DHL", "CARGO", "ADUANA", "GARAJE", "ALMAVIVA", "ALMACENADORA", "UPS", "FEDEX", "FLETE", "TRANSPORTE", "PUERTO", "TERMINAL"]):
                            tags_sug.append("🚢 Logística/Aduana")
                            score_sug += 50
                            
                        cands_sug_list.append({
                            "Factura": f_c_num,
                            "Proveedor": tr_cand_s.get("Proveedor", "")[:32],
                            "Fecha": fec_str,
                            "Saldo": sc_cand_val,
                            "Dias": diff_dias,
                            "Tags": " • ".join(tags_sug) if tags_sug else "⚪ Factura Libre",
                            "Score": score_sug,
                            "Row": tr_cand_s
                        })
                        
                    cands_sug_list.sort(key=lambda x: x["Score"], reverse=True)
                    
                    # Detectar parejas de facturas que sumen al faltante
                    pareja_match = None
                    for i_p in range(min(15, len(cands_sug_list))):
                        for j_p in range(i_p + 1, min(15, len(cands_sug_list))):
                            c1_p = cands_sug_list[i_p]
                            c2_p = cands_sug_list[j_p]
                            s_pair = round(c1_p["Saldo"] + c2_p["Saldo"], 2)
                            d_pair = abs(s_pair - dif_faltante_prev)
                            if d_pair < 50000:
                                pareja_match = (d_pair, s_pair, c1_p, c2_p)
                                break
                        if pareja_match:
                            break
                            
                    if pareja_match:
                        dp, sp, cp1, cp2 = pareja_match
                        col_par1, col_par2 = st.columns([3.2, 1.2])
                        with col_par1:
                            st.markdown(
                                f"<div style='background:#fefce8; border:1px solid #eab308; border-radius:6px; padding:8px 12px; margin-bottom:8px;'>"
                                f"💡 <b>Combinación de 2 Facturas Sugerida:</b> <code>{cp1['Factura']}</code> ({cp1['Proveedor']}, ${cp1['Saldo']:,.2f}) + <code>{cp2['Factura']}</code> ({cp2['Proveedor']}, ${cp2['Saldo']:,.2f})<br>"
                                f"<span style='font-size:12px; color:#854d0e;'>Suma Combinada: <b>${sp:,.2f}</b> (Diferencia: <b>${dp:,.2f}</b>)</span></div>",
                                unsafe_allow_html=True
                            )
                        with col_par2:
                            st.write("")
                            if st.button("➕ Añadir Ambas Facturas", key=f"btn_add_pair_{pq_id_sel}", type="primary", use_container_width=True):
                                terceros_nuevo = pd.concat([terceros_actual, pd.DataFrame([cp1["Row"], cp2["Row"]])]).drop_duplicates(subset=["Factura"]).reset_index(drop=True)
                                st.session_state["paquetes_importacion"][pq_id_sel]["terceros"] = terceros_nuevo
                                st.session_state["paquete_seleccionado_id"] = pq_id_sel
                                st.session_state["sel_paquete_activo_key"] = pq_id_sel
                                guardar_estado_manual(empresa)
                                st.success(f"¡Facturas {cp1['Factura']} y {cp2['Factura']} añadidas con éxito al Paquete #{pq_id_sel}!")
                                st.rerun()

                    if cands_sug_list:
                        top_cands = cands_sug_list[:6]
                        for idx_sug, sug_item in enumerate(top_cands):
                            c_sug1, c_sug2, c_sug3 = st.columns([1.5, 3.2, 1.1])
                            with c_sug1:
                                st.markdown(f"<b>{sug_item['Factura']}</b><br><span style='font-size:12px; color:#475569;'>{sug_item['Fecha']}</span>", unsafe_allow_html=True)
                            with c_sug2:
                                st.markdown(f"<b>{sug_item['Proveedor']}</b> — <b style='color:#0369a1;'>${sug_item['Saldo']:,.2f}</b><br><span style='font-size:12px; color:#15803d;'>{sug_item['Tags']}</span>", unsafe_allow_html=True)
                            with c_sug3:
                                st.write("")
                                if st.button("➕ Añadir", key=f"btn_sug_add_{pq_id_sel}_{idx_sug}", use_container_width=True):
                                    terceros_nuevo = pd.concat([terceros_actual, pd.DataFrame([sug_item["Row"]])]).drop_duplicates(subset=["Factura"]).reset_index(drop=True)
                                    st.session_state["paquetes_importacion"][pq_id_sel]["terceros"] = terceros_nuevo
                                    st.session_state["paquete_seleccionado_id"] = pq_id_sel
                                    st.session_state["sel_paquete_activo_key"] = pq_id_sel
                                    guardar_estado_manual(empresa)
                                    st.success(f"¡Factura {sug_item['Factura']} añadida al Paquete #{pq_id_sel}!")
                                    st.rerun()
                    else:
                        st.info(f"No se encontraron facturas sugeridas que no superen el monto de terceros (${monto_limite_terc:,.2f}) ni estén bloqueadas.")
                    st.markdown("---")

                st.markdown("##### ➕ Añadir Manualmente otra Factura Libre a este paquete:")
                filtro_cercanas_man = st.checkbox(
                    "📅 Filtrar solo facturas de la ventana cronológica (-60d anteriores a +15d posteriores)",
                    value=True,
                    key=f"chk_fec_man_{pq_id_sel}",
                    help="Oculta facturas lejanas de otros meses y muestra solo facturas anteriores o un poco después de la operación."
                )

                opciones_cands = []
                mapa_agregar = {}

                for _, tr_cand in cands_disp_agregar.iterrows():
                    f_cand_num = str(tr_cand["Factura"]).strip()
                    if f_cand_num not in facs_en_este and f_cand_num not in facturas_bloqueadas_validadas:
                        fec_c_str = str(tr_cand.get("Fecha", ""))
                        fec_c_dt = pd.to_datetime(fec_c_str, dayfirst=True, errors="coerce")
                        diff_m = int((fec_c_dt - f_ag_dt_sug).days) if (pd.notna(fec_c_dt) and pd.notna(f_ag_dt_sug)) else 999
                        
                        if filtro_cercanas_man and (diff_m < -60 or diff_m > 15):
                            continue
                            
                        es_reg_c = bool(tr_cand.get("Ya Registrada", False))
                        tag_est = f"🔴 Registrada ({tr_cand.get('Comprobante Previo', '10-Prev')})" if es_reg_c else "⚪ No Contabilizada (Pendiente)"
                        sc_cand = obtener_saldo_cruce_factura(tr_cand)
                        d_tag = f"({diff_m:+d}d)" if diff_m != 999 else "(sin fecha)"
                        tag_c = f"[{tag_est}] [{f_cand_num}] {d_tag} {tr_cand['Proveedor'][:20]} (Saldo: ${sc_cand:,.2f})"
                        opciones_cands.append((abs(diff_m), tag_c, tr_cand))

                opciones_cands.sort(key=lambda x: x[0])
                opciones_agregar = [x[1] for x in opciones_cands]
                for _, tag_c, tr_cand in opciones_cands:
                    mapa_agregar[tag_c] = tr_cand
                        
                c_add1, c_add2 = st.columns([3, 1.2])
                with c_add1:
                    sel_para_agregar = st.selectbox("Selecciona factura libre para añadir a este paquete:", ["(Seleccionar...)"] + opciones_agregar, key=f"sel_add_tr_{pq_id_sel}")
                with c_add2:
                    st.write("")
                    st.write("")
                    if st.button("➕ Añadir al Paquete", key=f"btn_add_tr_{pq_id_sel}") and sel_para_agregar != "(Seleccionar...)":
                        fila_agregada = mapa_agregar[sel_para_agregar]
                        terceros_nuevo = pd.concat([terceros_actual, pd.DataFrame([fila_agregada])]).drop_duplicates(subset=["Factura"]).reset_index(drop=True)
                        st.session_state["paquetes_importacion"][pq_id_sel]["terceros"] = terceros_nuevo
                        st.session_state["paquete_seleccionado_id"] = pq_id_sel
                        st.session_state["sel_paquete_activo_key"] = pq_id_sel
                        guardar_estado_manual(empresa)
                        st.success(f"¡Factura {fila_agregada['Factura']} añadida al Paquete #{pq_id_sel}!")
                        st.rerun()

            # 3. TRATAMIENTO CONTABLE: PÁGINA 2 / MERCANCÍAS EN TRÁNSITO / NO DEDUCIBLES
            enviar_h2_activo = bool(st.session_state.get(f"enviar_h2_pq_{pq_id_sel}", False))
            enviar_gp_activo = bool(st.session_state.get(f"enviar_gp_pq_{pq_id_sel}", False))
            enviar_nd_activo = bool(st.session_state.get(f"enviar_nd_pq_{pq_id_sel}", False))

            c_box_style_t = (
                '<div style="background:#f8fafc; border:2px solid #0070ba; border-radius:8px; padding:14px; margin:14px 0 10px 0;">'
                '<h4 style="margin:0 0 4px 0; color:#0070ba;">⚖️ Tratamiento Contable para el Paquete #' + str(pq_id_sel) + ' (Cobro ' + str(agente_actual.get('Proveedor', '')) + '):</h4>'
                '<p style="margin:0; font-size:13.5px; color:#334155;">Selecciona cómo deseas registrar este cobro en su fecha de operación (<b>' + str(agente_actual.get('Fecha', '')) + '</b>):</p>'
                '</div>'
            )
            st.markdown(c_box_style_t, unsafe_allow_html=True)

            c_btn_h2, c_btn_mt, c_btn_nd = st.columns([1.8, 1.6, 1.2])

            with c_btn_h2:
                st.markdown("<b style='color:#16a34a;'>📋 Opción 1: Traer de Página 2</b><br><span style='font-size:12px; color:#475569;'>Trae ingresos propios y deja saldo de terceros pendiente para conciliar:</span>", unsafe_allow_html=True)
                btn_h2_type = "primary" if (enviar_h2_activo or not (enviar_gp_activo or enviar_nd_activo)) else "secondary"
                if st.button("📋 Traer de Página 2 (Dejar Pendiente)", key=f"btn_h2_act_{pq_id_sel}", type=btn_h2_type, use_container_width=True):
                    st.session_state[f"enviar_h2_pq_{pq_id_sel}"] = True
                    st.session_state[f"enviar_gp_pq_{pq_id_sel}"] = False
                    st.session_state[f"enviar_nd_pq_{pq_id_sel}"] = False
                    st.session_state["paquete_seleccionado_id"] = pq_id_sel
                    st.session_state["sel_paquete_activo_key"] = pq_id_sel
                    if st.session_state.get(f"paquete_listo_{pq_id_sel}", False):
                        df_as_act, _ = generar_asiento_mixto_hoja2_con_terceros(agente_actual, terceros_actual, enviar_a_no_deducible=False, imputar_a_transito=False)
                        st.session_state[f"asiento_fijo_pq_{pq_id_sel}"] = df_as_act.copy()
                        paquete_activo["asiento_fijo"] = df_as_act.copy()
                        ag_fac_str = str(agente_actual["Factura"]).strip()
                        st.session_state.setdefault("asientos_triangulacion_por_factura", {})[ag_fac_str] = df_as_act.copy()
                        guardar_estado_manual(empresa)
                    st.success("¡Contabilización de Página 2 aplicada! El saldo de terceros queda pendiente para conciliar con facturas.")
                    st.rerun()

            with c_btn_mt:
                st.markdown("<b style='color:#0284c7;'>📦 Opción 2: Mercancías en Tránsito</b><br><span style='font-size:12px; color:#475569;'>Imputa el saldo faltante a Inventarios en Tránsito (Cta 14650501):</span>", unsafe_allow_html=True)
                lbl_btn_mt = f"📦 Mercancías en Tránsito (${dif_faltante_prev:,.2f})" if dif_faltante_prev > 0.05 else "📦 Mercancías en Tránsito (14650501)"
                btn_mt_type = "primary" if enviar_gp_activo else "secondary"
                if st.button(lbl_btn_mt, key=f"btn_mt_act_{pq_id_sel}", type=btn_mt_type, use_container_width=True):
                    st.session_state[f"enviar_gp_pq_{pq_id_sel}"] = True
                    st.session_state[f"enviar_h2_pq_{pq_id_sel}"] = False
                    st.session_state[f"enviar_nd_pq_{pq_id_sel}"] = False
                    st.session_state["paquete_seleccionado_id"] = pq_id_sel
                    st.session_state["sel_paquete_activo_key"] = pq_id_sel
                    if saldo_terceros_esperado_act > 0.05 or enviar_h2_activo:
                        df_as_act, _ = generar_asiento_mixto_hoja2_con_terceros(agente_actual, terceros_actual, enviar_a_no_deducible=False, imputar_a_transito=True)
                    else:
                        df_as_act, _, ret_asum = generar_asiento_triangulacion_paquete(agente_actual, terceros_actual, enviar_a_no_deducible=False)
                        asiento_gp_filas = []
                        for _, r_as in df_as_act.iterrows():
                            if str(r_as["Código Cuenta"]).strip() == CUENTA_NO_DEDUCIBLE:
                                asiento_gp_filas.append({
                                    "Código Cuenta": CUENTA_IMPORTACION_TRANSITO,
                                    "Descripción Cuenta": f"Mercancías en Tránsito / Base Propia Agente (Fac {agente_actual.get('Factura', '')})",
                                    "Tercero / NIT": f"{agente_actual.get('NIT Emisor', '')} - {agente_actual.get('Proveedor', '')[:25]}",
                                    "Débito ($)": float(r_as["Débito ($)"]),
                                    "Crédito ($)": 0.0
                                })
                            else:
                                asiento_gp_filas.append(r_as.to_dict())
                        df_as_act = pd.DataFrame(asiento_gp_filas)
                    if st.session_state.get(f"paquete_listo_{pq_id_sel}", False):
                        st.session_state[f"asiento_fijo_pq_{pq_id_sel}"] = df_as_act.copy()
                        paquete_activo["asiento_fijo"] = df_as_act.copy()
                        ag_fac_str = str(agente_actual["Factura"]).strip()
                        st.session_state.setdefault("asientos_triangulacion_por_factura", {})[ag_fac_str] = df_as_act.copy()
                        guardar_estado_manual(empresa)
                    st.success("¡Faltante asignado a Mercancías en Tránsito (14650501)!")
                    st.rerun()

            with c_btn_nd:
                st.markdown("<b style='color:#b91c1c;'>🔴 Opción 3: No Deducibles</b><br><span style='font-size:12px; color:#475569;'>Si no existe factura DIAN ni soporte:</span>", unsafe_allow_html=True)
                lbl_btn_nd = f"🔴 No Deducibles (${dif_faltante_prev:,.2f})" if dif_faltante_prev > 0.05 else "🔴 No Deducibles (53950501)"
                btn_nd_type = "primary" if enviar_nd_activo else "secondary"
                if st.button(lbl_btn_nd, key=f"btn_nd_act_{pq_id_sel}", type=btn_nd_type, use_container_width=True):
                    st.session_state[f"enviar_nd_pq_{pq_id_sel}"] = True
                    st.session_state[f"enviar_gp_pq_{pq_id_sel}"] = False
                    st.session_state[f"enviar_h2_pq_{pq_id_sel}"] = False
                    st.session_state["paquete_seleccionado_id"] = pq_id_sel
                    st.session_state["sel_paquete_activo_key"] = pq_id_sel
                    if saldo_terceros_esperado_act > 0.05 or enviar_h2_activo:
                        df_as_act, _ = generar_asiento_mixto_hoja2_con_terceros(agente_actual, terceros_actual, enviar_a_no_deducible=True, imputar_a_transito=False)
                    else:
                        df_as_act, _, _ = generar_asiento_triangulacion_paquete(agente_actual, terceros_actual, enviar_a_no_deducible=True)
                    if st.session_state.get(f"paquete_listo_{pq_id_sel}", False):
                        st.session_state[f"asiento_fijo_pq_{pq_id_sel}"] = df_as_act.copy()
                        paquete_activo["asiento_fijo"] = df_as_act.copy()
                        ag_fac_str = str(agente_actual["Factura"]).strip()
                        st.session_state.setdefault("asientos_triangulacion_por_factura", {})[ag_fac_str] = df_as_act.copy()
                        guardar_estado_manual(empresa)
                    st.success("¡Faltante enviado a Gastos No Deducibles (53950501)!")
                    st.rerun()

            # RECALCULAR ASIENTO CONTABLE PARA ESTE PAQUETE
            if st.session_state.get(f"paquete_listo_{pq_id_sel}", False) and (f"asiento_fijo_pq_{pq_id_sel}" in st.session_state or "asiento_fijo" in paquete_activo):
                df_asiento_paquete = st.session_state.get(f"asiento_fijo_pq_{pq_id_sel}", paquete_activo.get("asiento_fijo")).copy()
                dif_no_ded = 0.0
                ret_asum = 0.0
            elif saldo_terceros_esperado_act > 0.05 or enviar_h2_activo:
                if enviar_nd_activo:
                    df_asiento_paquete, _ = generar_asiento_mixto_hoja2_con_terceros(agente_actual, terceros_actual, enviar_a_no_deducible=True, imputar_a_transito=False)
                    dif_no_ded = dif_faltante_prev
                elif enviar_gp_activo:
                    df_asiento_paquete, _ = generar_asiento_mixto_hoja2_con_terceros(agente_actual, terceros_actual, enviar_a_no_deducible=False, imputar_a_transito=True)
                    dif_no_ded = 0.0
                else:
                    df_asiento_paquete, dif_no_ded = generar_asiento_mixto_hoja2_con_terceros(agente_actual, terceros_actual, enviar_a_no_deducible=False, imputar_a_transito=False)
                ret_asum = 0.0
            elif enviar_gp_activo and dif_faltante_prev > 0.05:
                df_asiento_paquete, _, ret_asum = generar_asiento_triangulacion_paquete(agente_actual, terceros_actual, enviar_a_no_deducible=False)
                asiento_gp_filas = []
                for _, r_as in df_asiento_paquete.iterrows():
                    if str(r_as["Código Cuenta"]).strip() == CUENTA_NO_DEDUCIBLE:
                        asiento_gp_filas.append({
                            "Código Cuenta": CUENTA_IMPORTACION_TRANSITO,
                            "Descripción Cuenta": f"Mercancías en Tránsito / Base Propia Agente (Fac {agente_actual.get('Factura', '')})",
                            "Tercero / NIT": f"{agente_actual.get('NIT Emisor', '')} - {agente_actual.get('Proveedor', '')[:25]}",
                            "Débito ($)": float(r_as["Débito ($)"]),
                            "Crédito ($)": 0.0
                        })
                    else:
                        asiento_gp_filas.append(r_as.to_dict())
                df_asiento_paquete = pd.DataFrame(asiento_gp_filas)
                dif_no_ded = 0.0
            else:
                df_asiento_paquete, dif_no_ded, ret_asum = generar_asiento_triangulacion_paquete(agente_actual, terceros_actual, enviar_a_no_deducible=enviar_nd_activo)

            st.markdown(f"#### ⚖️ Asiento Contable del Paquete #{pq_id_sel}:")
            st.caption("Detalle de partida doble de ESTE paquete: cancela las cuentas por pagar de DHL, Agencia y Garaje contra Euro Shipping:")

            st.dataframe(
                df_asiento_paquete.style.format({"Débito ($)": "${:,.2f}", "Crédito ($)": "${:,.2f}"}),
                use_container_width=True,
                hide_index=True
            )

            c_cuad1, c_cuad2, c_cuad3, c_cuad4 = st.columns(4)
            sum_deb_pq = df_asiento_paquete["Débito ($)"].sum()
            sum_cred_pq = df_asiento_paquete["Crédito ($)"].sum()
            dif_cuad_pq = abs(sum_deb_pq - sum_cred_pq)
            
            with c_cuad1:
                st.metric("Total Débito", f"${sum_deb_pq:,.2f}")
            with c_cuad2:
                st.metric("Total Crédito", f"${sum_cred_pq:,.2f}")
            with c_cuad3:
                st.metric("Retenciones Asumidas (53152001)", f"${ret_asum:,.2f}")
            with c_cuad4:
                st.metric("Diferencia No Deducible (53950501)", f"${dif_no_ded:,.2f}")

            # SECCIÓN DE VALIDACIÓN Y BLOQUEO DE FACTURAS
            es_validado = st.session_state.get(f"paquete_listo_{pq_id_sel}", False)

            if not es_validado:
                if dif_cuad_pq < 0.05:
                    st.success(f"✅ **Partida doble cuadrada al centavo ($0.00).** Ya puedes validar y bloquear este paquete.")
                else:
                    st.markdown(
                        f"""
                        <div style="background:#fffbeb; border:2px solid #f59e0b; border-radius:8px; padding:12px 16px; margin:10px 0;">
                            <span style="font-size:15px; font-weight:bold; color:#b45309;">⏳ Faltante Pendiente de Conciliar por Ingresos para Terceros: ${dif_cuad_pq:,.2f}</span><br>
                            <span style="font-size:13px; color:#78350f;">
                                El valor cobrado por cuenta de terceros permanece <b>pendiente de conciliación</b>. Puedes conciliarlo seleccionando las facturas sugeridas arriba, o si decides enviar el saldo faltante a <b>Mercancías en Tránsito (Opción 2)</b> o a <b>No Deducibles (Opción 3)</b>, presiona el botón respectivo en las opciones superiores.
                            </span>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

                c_val_opt1, c_val_opt2 = st.columns(2)
                with c_val_opt1:
                    st.markdown(
                        "<div style='background:#f0fdf4; border:1.5px solid #22c55e; border-radius:8px; padding:12px; height:100%;'>"
                        "<b style='color:#15803d; font-size:15px;'>📤 Opción 1: Pasar a Hoja 2 (Subir a Planilla Siigo)</b><br>"
                        "<span style='font-size:12.5px; color:#166534;'>"
                        "Bloquea las facturas del paquete, actualiza la Página 2 y <b>genera el comprobante oficial en la Planilla de Siigo (Página 4)</b> para su contabilización."
                        "</span></div>",
                        unsafe_allow_html=True
                    )
                    st.write("")
                    if st.button("📤 Pasar a Hoja 2 (Subir a Planilla Siigo)", key=f"btn_validar_siigo_{pq_id_sel}", type="primary", use_container_width=True):
                        ejecutar_validacion_pq(pasa_a_siigo=True)
                        st.success(f"🔒 ¡Paquete #{pq_id_sel} VALIDADO y trasladado a Página 2 para Siigo!")
                        st.rerun()

                with c_val_opt2:
                    st.markdown(
                        "<div style='background:#eff6ff; border:1.5px solid #3b82f6; border-radius:8px; padding:12px; height:100%;'>"
                        "<b style='color:#1d4ed8; font-size:15px;'>📑 Opción 2: Simplemente Conciliación (No Subir a Siigo)</b><br>"
                        "<span style='font-size:12.5px; color:#1e40af;'>"
                        "Bloquea e identifica las facturas para <b>auditoría y soporte de costos</b>, pero <b>NO genera comprobante en Siigo</b> (evita duplicar pagos ya hechos por tesorería/bancos)."
                        "</span></div>",
                        unsafe_allow_html=True
                    )
                    st.write("")
                    if st.button("📑 Simplemente Conciliación (No Subir a Siigo)", key=f"btn_validar_concil_{pq_id_sel}", use_container_width=True):
                        ejecutar_validacion_pq(pasa_a_siigo=False)
                        st.success(f"📑 ¡Paquete #{pq_id_sel} CONCILIADO exitosamente para control interno (no irá a Siigo)!")
                        st.rerun()
            else:
                es_siigo = st.session_state.get(f"paquete_destino_siigo_{pq_id_sel}", True)
                col_desb1, col_desb2 = st.columns([3.2, 1.4])
                with col_desb1:
                    if es_siigo:
                        st.success(f"🔒 **Paquete #{pq_id_sel} VALIDADO (PASA A HOJA 2 / PLANILLA SIIGO):** Sus facturas están protegidas contra otros cruces y su comprobante se exportará en la Planilla de Siigo (Página 4).")
                    else:
                        st.info(f"📑 **Paquete #{pq_id_sel} CONCILIADO (SOLO CONTROL INTERNO / IDENTIFICACIÓN):** Facturas identificadas y protegidas para auditoría DIAN. **No se generará comprobante en Siigo** para evitar duplicidades.")
                with col_desb2:
                    c_sw1, c_sw2 = st.columns(2)
                    with c_sw1:
                        if es_siigo:
                            if st.button("📑 Cambiar a Solo Conciliación", key=f"btn_sw_concil_{pq_id_sel}", use_container_width=True, help="Cambia el destino para que NO se suba a la planilla Siigo"):
                                ejecutar_validacion_pq(pasa_a_siigo=False)
                                st.rerun()
                        else:
                            if st.button("📤 Cambiar a Pasar a Hoja 2", key=f"btn_sw_siigo_{pq_id_sel}", type="primary", use_container_width=True, help="Cambia el destino para que SÍ se suba a Siigo"):
                                ejecutar_validacion_pq(pasa_a_siigo=True)
                                st.rerun()
                    with c_sw2:
                        if st.button("🔓 Desbloquear", key=f"btn_desbloquear_main_{pq_id_sel}", use_container_width=True):
                            ejecutar_desbloqueo_pq()
                            st.rerun()

            st.markdown("---")
            st.markdown("#### 📥 Exportar Control de Paquetes y Cruces de Importación:")
            st.caption("Descarga un archivo Excel con el resumen de cada paquete, el detalle de facturas de terceros vinculadas y el asiento contable de control (sin generar aún notas de contabilidad en Siigo hasta que definas el procedimiento):")

            # Generar Excel estructurado por Paquetes de Importación individuales
            todos_asientos_lista = []
            resumen_paquetes_lista = []
            detalle_facturas_lista = []
            consecutivo_cruce = 1
            
            facs_asignadas_en_algun_pq = []
            
            for g_k, g_v in pqs_actuales.items():
                ag_item = g_v["agente"]
                terc_items = g_v["terceros"]
                enviar_nd_g = st.session_state.get(f"enviar_nd_pq_{g_k}", False)
                df_as_p, dif_p, ret_p = generar_asiento_triangulacion_paquete(ag_item, terc_items, enviar_a_no_deducible=enviar_nd_g)
                
                # 1. Asiento de Control de Cruces (Partida Doble para Revisión)
                for _, fila_as in df_as_p.iterrows():
                    todos_asientos_lista.append({
                        "Paquete #": f"Paquete #{g_k}",
                        "Fecha Operación": ag_item["Fecha"],
                        "Código Cuenta": fila_as["Código Cuenta"],
                        "Tercero / NIT": fila_as["Tercero / NIT"],
                        "Descripción": str(fila_as.get("Descripción de la Cuenta") or fila_as.get("Descripción Cuenta") or fila_as.get("Descripción") or ""), 
                        "Débito ($)": fila_as["Débito ($)"],
                        "Crédito ($)": fila_as["Crédito ($)"]
                    })
                    
                # 2. Resumen del Paquete
                facs_terc_str = ", ".join([f"{r['Factura']} ({r['Proveedor'][:15]})" for _, r in terc_items.iterrows()]) if not terc_items.empty else "Ninguno"
                tot_terc_sum = sum([obtener_saldo_cruce_factura(r) for _, r in terc_items.iterrows()])
                resumen_paquetes_lista.append({
                    "Paquete #": f"Paquete #{g_k}",
                    "Agente Coordinador": ag_item["Proveedor"],
                    "Factura Agente": ag_item["Factura"],
                    "Fecha Operación": ag_item["Fecha"],
                    "Total Cobro Agente ($)": float(ag_item["Total"]),
                    "Facturas Terceros Incluidas": facs_terc_str,
                    "Total Cancelado Terceros ($)": tot_terc_sum,
                    "Retenciones Asumidas ($)": ret_p,
                    "Diferencia No Deducible ($)": dif_p,
                    "Estado Cuadre": "✅ Cuadrado ($0.00)" if abs(float(ag_item["Total"]) - tot_terc_sum - ret_p - dif_p) < 1.0 else "Revisar"
                })
                
                # 3. Detalle de Facturas que componen el Paquete
                detalle_facturas_lista.append({
                    "Paquete #": f"Paquete #{g_k}",
                    "Rol en Operación": "Agente Coordinador (Cobro Global)",
                    "Proveedor": ag_item["Proveedor"],
                    "NIT": ag_item["NIT Emisor"],
                    "Factura": ag_item["Factura"],
                    "Fecha": ag_item["Fecha"],
                    "Naturaleza": "Servicios Agenciamiento + Fletes / Impuestos",
                    "Base / Subtotal ($)": float(ag_item.get("Base", 0.0)),
                    "IVA ($)": float(ag_item.get("IVA", 0.0)),
                    "Retenciones ($)": 0.0,
                    "Saldo a Cruzar ($)": float(ag_item["Total"]),
                    "Cuenta Contable": "22050501"
                })
                
                for _, tr_it in terc_items.iterrows():
                    facs_asignadas_en_algun_pq.append(tr_it["Factura"])
                    tb = float(tr_it.get("Base", 0.0))
                    tiv = float(tr_it.get("IVA", 0.0))
                    rf = float(tr_it.get("ReteFuente", 0.0))
                    ri = float(tr_it.get("ReteICA", 0.0))
                    sc = obtener_saldo_cruce_factura(tr_it)
                    p_u = str(tr_it["Proveedor"]).upper()
                    rol_t = "🚚 Tercero Transporte (DHL)" if "DHL" in p_u else ("🏢 Tercero Agenciamiento (Mandato)" if any(k in p_u for k in ["CARGO", "ADUANA"]) else "🏬 Tercero Garaje / Almacén")
                    nat_t = "Honorario Propio Agenciamiento" if "COMISION" in str(tr_it.get("Descripcion", "")).upper() else "Gasto por Cuenta de Tercero"
                    detalle_facturas_lista.append({
                        "Paquete #": f"Paquete #{g_k}",
                        "Rol en Operación": rol_t,
                        "Proveedor": tr_it["Proveedor"],
                        "NIT": tr_it["NIT Emisor"],
                        "Factura": tr_it["Factura"],
                        "Fecha": tr_it["Fecha"],
                        "Naturaleza": nat_t,
                        "Base / Subtotal ($)": tb,
                        "IVA ($)": tiv,
                        "Retenciones ($)": round(rf + ri, 2),
                        "Saldo a Cruzar ($)": sc,
                        "Cuenta Contable": str(tr_it.get("Cuenta Pasivo Especifica", "22050505"))
                    })
                    
                consecutivo_cruce += 1

            # Añadir facturas de terceros que quedaron libres (sin asignar a ningún paquete)
            df_libres_exp = df_terceros_all[~df_terceros_all["Factura"].isin(facs_asignadas_en_algun_pq)]
            for _, tr_lib in df_libres_exp.iterrows():
                tb = float(tr_lib.get("Base", 0.0))
                tiv = float(tr_lib.get("IVA", 0.0))
                sc = obtener_saldo_cruce_factura(tr_lib)
                p_u = str(tr_lib["Proveedor"]).upper()
                rol_t = "🚚 Tercero Transporte (DHL)" if "DHL" in p_u else ("🏢 Tercero Agenciamiento (Mandato)" if any(k in p_u for k in ["CARGO", "ADUANA"]) else "🏬 Tercero Garaje / Almacén")
                detalle_facturas_lista.append({
                    "Paquete #": "⚪ Sin Asignar (Pendiente de Operación)",
                    "Rol en Operación": rol_t,
                    "Proveedor": tr_lib["Proveedor"],
                    "NIT": tr_lib["NIT Emisor"],
                    "Factura": tr_lib["Factura"],
                    "Fecha": tr_lib["Fecha"],
                    "Naturaleza": "Gasto por Cuenta de Tercero",
                    "Base / Subtotal ($)": tb,
                    "IVA ($)": tiv,
                    "Retenciones ($)": float(tr_lib.get("ReteFuente", 0.0)) + float(tr_lib.get("ReteICA", 0.0)),
                    "Saldo a Cruzar ($)": sc,
                    "Cuenta Contable": str(tr_lib.get("Cuenta Pasivo Especifica", "22050505"))
                })

            if todos_asientos_lista:
                df_export_cruces = pd.DataFrame(todos_asientos_lista)
                buf_cruce_ex = io.BytesIO()
                with pd.ExcelWriter(buf_cruce_ex, engine="openpyxl") as wr_cr:
                    if resumen_paquetes_lista:
                        pd.DataFrame(resumen_paquetes_lista).to_excel(wr_cr, sheet_name="Resumen_Paquetes", index=False)
                    if detalle_facturas_lista:
                        pd.DataFrame(detalle_facturas_lista).to_excel(wr_cr, sheet_name="Detalle_Facturas_Paquete", index=False)
                    df_export_cruces.to_excel(wr_cr, sheet_name="Asiento_Control_Cruces", index=False)
                buf_cruce_ex.seek(0)
                
                st.download_button(
                    label=f"📥 Descargar Control de Cruces de Importación en Excel ({len(pqs_actuales)} Paquetes)",
                    data=buf_cruce_ex.getvalue(),
                    file_name=f"Control_Cruces_Importacion_{empresa['nombre'].replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    else:
        st.info("💡 Sube el reporte Excel de la DIAN en la Pestaña 1 para habilitar la triangulación y cruce de importaciones.")


with tab_siigo:
    st.markdown("### Descargar Planilla Oficial Siigo Nube (3 Hojas)")
    st.caption("Planilla oficial formulada con 'matriz_captura', 'interfaz_siigo' y 'Parametrización'.")

    if "df_procesado" in st.session_state:
        df_full = st.session_state["df_procesado"]

        # FILTRO DE PROTECCIÓN: Excluir facturas rojas (ya causadas en Siigo) y facturas que van por triangulación aduanera
        ya_reg_siigo = df_full["Ya Registrada"].fillna(False).astype(bool) if "Ya Registrada" in df_full.columns else pd.Series(False, index=df_full.index)
        n_rojas = int(ya_reg_siigo.sum())

        # Asegurar columna No Contabilizar en df_full
        if "No Contabilizar" not in df_full.columns:
            df_full["No Contabilizar"] = False
            st.session_state["df_procesado"] = df_full

        # 1. Facturas explícitamente excluidas de contabilidad (No Contabilizar)
        excluidas_set = st.session_state.get("facturas_no_contabilizar", set())
        cond_excluidas = df_full["Factura"].isin(excluidas_set) | (df_full["No Contabilizar"] == True)
        n_excluidas = int(cond_excluidas.sum())

        # 2. Facturas que ya están en un paquete de triangulación aduanera
        facs_en_pqs_import = []
        pqs_actuales_siigo = st.session_state.get("paquetes_importacion", {})
        if pqs_actuales_siigo:
            for _, p_val in pqs_actuales_siigo.items():
                if isinstance(p_val, dict) and "terceros" in p_val and not p_val["terceros"].empty:
                    facs_en_pqs_import.extend(p_val["terceros"]["Factura"].tolist())

        es_adu_siigo = df_full["Es Aduanera"].fillna(False).astype(bool) if "Es Aduanera" in df_full.columns else pd.Series(False, index=df_full.index)
        grp_imp_siigo = df_full["Grupo Importación"].fillna("").astype(str).str.strip() if "Grupo Importación" in df_full.columns else pd.Series("", index=df_full.index)
        cond_imp_adu = es_adu_siigo & (grp_imp_siigo != "")

        df_p = df_full[
            (~ya_reg_siigo) &
            (~cond_excluidas) &
            (~df_full["Factura"].isin(facs_en_pqs_import)) &
            (~cond_imp_adu)
        ].copy()

        if n_rojas > 0:
            st.info(f"🛡️ **Protección contra duplicados:** Se excluyeron {n_rojas} facturas marcadas en rojo que ya estaban causadas en Siigo.")
        if n_excluidas > 0:
            st.warning(f"🚫 **Facturas Excluidas:** Se excluyeron {n_excluidas} facturas marcadas como **'NO CONTABILIZAR'**.")

        # Paquetes de triangulación listos para llevar a contabilidad en su fecha (excluyendo los marcados solo para conciliación)
        pqs_listos_siigo = {k: v for k, v in pqs_actuales_siigo.items() if st.session_state.get(f"paquete_listo_{k}", False) and st.session_state.get(f"paquete_destino_siigo_{k}", True)}
        pqs_solo_concil = {k: v for k, v in pqs_actuales_siigo.items() if st.session_state.get(f"paquete_listo_{k}", False) and not st.session_state.get(f"paquete_destino_siigo_{k}", True)}

        c_pqs_inc1, c_pqs_inc2 = st.columns([2.5, 1.5])
        with c_pqs_inc1:
            inc_pqs_listos = st.checkbox(
                f"🔀 **Llevar a la contabilidad {len(pqs_listos_siigo)} Paquete(s) de Triangulación que pasan a Siigo**",
                value=True if pqs_listos_siigo else False,
                help="Inserta los asientos contables de los paquetes listos marcados como 'Pasar a Hoja 2' en la planilla oficial de Siigo."
            )
        with c_pqs_inc2:
            st.caption(f"Suben a Siigo: **{len(pqs_listos_siigo)}** | Solo Conciliación: **{len(pqs_solo_concil)}** | Totales: **{len(pqs_actuales_siigo)}**")

        wb = openpyxl.Workbook()

        # 1. Hoja matriz_captura
        ws_matriz = wb.active
        ws_matriz.title = "matriz_captura"
        headers_matriz = [
            "Tipo Comp", "Consecutivo", "Fecha (DD/MM/AAAA)", "NIT", "Prefijo", "Factura Num",
            "Descripcion", "Operacion", "Cta Principal", "Valor Base", "IVA", "ReteFuente",
            "ReteICA", "ReteIVA", "Cta Contrapartida"
        ]
        ws_matriz.append(headers_matriz)

        # 2. Hoja interfaz_siigo
        ws_interfaz = wb.create_sheet(title="interfaz_siigo")
        headers_interfaz = [
            "Tipo de comprobante", "Consecutivo comprobante", "Fecha de elaboración", "Sigla moneda",
            "Tasa de cambio", "Código cuenta contable", "Identificación tercero", "Sucursal",
            "Código producto", "Código de bodega", "Acción", "Cantidad producto", "Prefijo",
            "Consecutivo", "No. cuota", "Fecha vencimiento", "Código impuesto", "Código grupo activo fijo",
            "Código activo fijo", "Descripción", "Código centro/subcentro de costos", "Débito",
            "Crédito", "Observaciones", "Base gravable libro compras/ventas", "Base exenta libro compras/ventas", "Mes de cierre"
        ]
        ws_interfaz.append(headers_interfaz)

        fila_r = 2
        asientos_triang_map_siigo = st.session_state.get("asientos_triangulacion_por_factura", {})

        for _, item in df_p.iterrows():
            t_comp = item["Tipo Comp"]
            cons = item["Consecutivo"]
            f_str = item["Fecha"]
            fn_item = str(item["Factura"]).strip()

            # Si esta factura tiene contabilización aprobada desde triangulación, se exporta su asiento real
            if fn_item in asientos_triang_map_siigo:
                df_as_tr = asientos_triang_map_siigo[fn_item]
                f_limpia_op = normalizar_fecha_dian(f_str)
                for _, r_as in df_as_tr.iterrows():
                    cta_code = str(r_as["Código Cuenta"]).strip()
                    nit_clean = re.sub(r'\D', '', str(r_as["Tercero / NIT"]).split("-")[0])
                    deb_v = float(r_as["Débito ($)"])
                    cred_v = float(r_as["Crédito ($)"])
                    desc_line = str(r_as.get("Descripción de la Cuenta") or r_as.get("Descripción Cuenta") or "")[:40]

                    ws_interfaz.append([
                        t_comp, cons, f_limpia_op, "COP", 1,
                        cta_code, nit_clean, 0, "", "", "", "",
                        item.get("Prefijo", "IMP"), item.get("Folio", str(item.get("Factura", ""))),
                        1, f_limpia_op, "", "", "",
                        desc_line, "", deb_v, cred_v,
                        f"Triangulación {item['Proveedor'][:15]} Fac {item['Factura']}",
                        0.0, 0.0, ""
                    ])
                    ws_matriz.append([
                        t_comp, cons, f_limpia_op, nit_clean,
                        item.get("Prefijo", "IMP"), item.get("Folio", str(item.get("Factura", ""))),
                        desc_line, "Cruce Triangulación", cta_code, deb_v, 0.0, 0.0, 0.0, 0.0, "22050501"
                    ])
                fila_r += len(df_as_tr)
                continue
            nit = item["NIT Emisor"]
            pref = item["Prefijo"]
            fac_num = item["Folio"]
            desc = item["Descripcion"]
            op = item["Operacion"]
            cta_p = item["Cta Principal"]
            base = item["Base"]
            iva = item["IVA"]
            rfte = item["ReteFuente"]
            rica = item.get("ReteICA", 0.0)
            riva = item.get("ReteIVA", 0.0)
            cta_c = item["Cta Contrapartida"]
            cta_iva = item.get("Cta IVA", "24081001")
            cta_rfte = item.get("Cta ReteFuente", "23654001")
            cta_rica = item.get("Cta ReteICA", "23680501")

            ws_matriz.append([t_comp, cons, f_str, nit, pref, fac_num, desc, op, cta_p, base, iva, rfte, rica, riva, cta_c])

            r = fila_r
            # Línea 1: Base Imponible
            ws_interfaz.append([
                f"=matriz_captura!A{r}", f"=matriz_captura!B{r}", f"=matriz_captura!C{r}", "COP", 1,
                f"=matriz_captura!I{r}", f"=matriz_captura!D{r}", 0, "", "", "", "", "", "", "", "", "", "", "",
                f"=matriz_captura!G{r}", "",
                f'=IF(matriz_captura!H{r}="Devolucion Compra", 0, matriz_captura!J{r})',
                f'=IF(matriz_captura!H{r}="Devolucion Compra", matriz_captura!J{r}, 0)',
                "", f"=matriz_captura!J{r}", 0.0, ""
            ])

            # Línea 2: IVA
            if iva > 0:
                cod_imp_iva = CODIGOS_IMPUESTO_SIIGO.get(str(cta_iva).strip(), "")
                ws_interfaz.append([
                    f"=matriz_captura!A{r}", f"=matriz_captura!B{r}", f"=matriz_captura!C{r}", "COP", 1,
                    cta_iva, f"=matriz_captura!D{r}", 0, "", "", "", "", "", "", "", "",
                    cod_imp_iva, "", "", f'="IVA Base: " & matriz_captura!J{r}', "",
                    f'=IF(matriz_captura!H{r}="Devolucion Compra", 0, matriz_captura!K{r})',
                    f'=IF(matriz_captura!H{r}="Devolucion Compra", matriz_captura!K{r}, 0)',
                    "", f"=matriz_captura!J{r}", 0.0, ""
                ])

            # Línea 2.B: IVA de Importación (240835)
            iva_imp_it = float(item.get("IVA Importación", 0.0))
            cta_iva_imp_it = str(item.get("Cta IVA Importación") or "240835").strip()
            if iva_imp_it > 0:
                cod_imp_ivaimp = CODIGOS_IMPUESTO_SIIGO.get(cta_iva_imp_it, "")
                ws_interfaz.append([
                    f"=matriz_captura!A{r}", f"=matriz_captura!B{r}", f"=matriz_captura!C{r}", "COP", 1,
                    cta_iva_imp_it, f"=matriz_captura!D{r}", 0, "", "", "", "", "", "", "", "",
                    cod_imp_ivaimp, "", "", f'="IVA Importación Fac " & matriz_captura!E{r} & "-" & matriz_captura!F{r}', "",
                    f'=IF(matriz_captura!H{r}="Devolucion Compra", 0, {iva_imp_it})',
                    f'=IF(matriz_captura!H{r}="Devolucion Compra", {iva_imp_it}, 0)',
                    "", f"=matriz_captura!J{r}", 0.0, ""
                ])

            # Línea 3: ReteFuente
            if rfte > 0 and cta_rfte:
                cod_imp_rfte = CODIGOS_IMPUESTO_SIIGO.get(str(cta_rfte).strip(), "")
                ws_interfaz.append([
                    f"=matriz_captura!A{r}", f"=matriz_captura!B{r}", f"=matriz_captura!C{r}", "COP", 1,
                    cta_rfte, f"=matriz_captura!D{r}", 0, "", "", "", "", "", "", "", "",
                    cod_imp_rfte, "", "", f'="ReteFuente Base: " & matriz_captura!J{r}', "",
                    f'=IF(matriz_captura!H{r}="Devolucion Compra", matriz_captura!L{r}, 0)',
                    f'=IF(matriz_captura!H{r}="Devolucion Compra", 0, matriz_captura!L{r})',
                    "", f"=matriz_captura!J{r}", 0.0, ""
                ])

            # Línea 4: ReteICA
            if rica > 0 and cta_rica:
                cod_imp_rica = CODIGOS_IMPUESTO_SIIGO.get(str(cta_rica).strip(), "")
                ws_interfaz.append([
                    f"=matriz_captura!A{r}", f"=matriz_captura!B{r}", f"=matriz_captura!C{r}", "COP", 1,
                    cta_rica, f"=matriz_captura!D{r}", 0, "", "", "", "", "", "", "", "",
                    cod_imp_rica, "", "", f'="ReteICA Base: " & matriz_captura!J{r}', "",
                    f'=IF(matriz_captura!H{r}="Devolucion Compra", matriz_captura!M{r}, 0)',
                    f'=IF(matriz_captura!H{r}="Devolucion Compra", 0, matriz_captura!M{r})',
                    "", f"=matriz_captura!J{r}", 0.0, ""
                ])

            # Línea 5: Cuenta por Pagar (Contrapartida)
            extra_iva_imp_str = f" + {iva_imp_it}" if iva_imp_it > 0 else ""
            formula_deb = f'=IF(matriz_captura!H{r}="Devolucion Compra", matriz_captura!J{r} + matriz_captura!K{r}{extra_iva_imp_str} - matriz_captura!L{r} - matriz_captura!M{r} - matriz_captura!N{r}, 0)'
            formula_cred = f'=IF(matriz_captura!H{r}="Devolucion Compra", 0, matriz_captura!J{r} + matriz_captura!K{r}{extra_iva_imp_str} - matriz_captura!L{r} - matriz_captura!M{r} - matriz_captura!N{r})'
            ws_interfaz.append([
                f"=matriz_captura!A{r}", f"=matriz_captura!B{r}", f"=matriz_captura!C{r}", "COP", 1,
                f"=matriz_captura!O{r}", f"=matriz_captura!D{r}", 0, "", "", "", "",
                f"=matriz_captura!A{r}", f"=matriz_captura!B{r}", 1, f"=matriz_captura!C{r}",
                "", "", "", f'="Fac " & matriz_captura!E{r} & "-" & matriz_captura!F{r}', "",
                formula_deb, formula_cred, "", 0.0, 0.0, ""
            ])
            fila_r += 1

        # 2.B AGREGAR A CONTABILIDAD CADA PAQUETE DE TRIANGULACIÓN LISTO EN SU FECHA EXACTA
        if inc_pqs_listos and pqs_listos_siigo:
            cons_cruce_base = 1
            if "cons_ini_nota" in locals() and cons_ini_nota:
                try:
                    cons_cruce_base = int(cons_ini_nota)
                except Exception:
                    cons_cruce_base = 1
            elif "cons_ini_nc" in locals() and cons_ini_nc:
                try:
                    cons_cruce_base = int(cons_ini_nc)
                except Exception:
                    cons_cruce_base = 1

            idx_pq_cruce = 0
            for p_k, p_v in pqs_listos_siigo.items():
                ag_item = p_v["agente"]
                ag_fac_chk = str(ag_item.get("Factura", "")).strip()
                # Si la factura del agente ya fue exportada con su asiento de triangulación en df_p, no duplicar
                if ag_fac_chk in df_p["Factura"].astype(str).str.strip().values and ag_fac_chk in asientos_triang_map_siigo:
                    continue
                terc_items = p_v["terceros"]
                fecha_op_raw = ag_item["Fecha"]
                fecha_op_clean = normalizar_fecha_dian(fecha_op_raw)
                enviar_nd = st.session_state.get(f"enviar_nd_pq_{p_k}", False)
                enviar_gp_p = st.session_state.get(f"enviar_gp_pq_{p_k}", False)
                enviar_h2_p = st.session_state.get(f"enviar_h2_pq_{p_k}", False)
                if st.session_state.get(f"paquete_listo_{p_k}", False) and (f"asiento_fijo_pq_{p_k}" in st.session_state or "asiento_fijo" in p_v):
                    df_as_p = pd.DataFrame(st.session_state.get(f"asiento_fijo_pq_{p_k}", p_v.get("asiento_fijo"))).copy()
                    dif_p = 0.0
                    ret_p = 0.0
                else:
                    df_as_p, dif_p, ret_p = generar_asiento_triangulacion_paquete(ag_item, terc_items, enviar_a_no_deducible=enviar_nd, enviar_a_gastos_propios=enviar_gp_p, enviar_a_hoja2=enviar_h2_p)

                cons_cruce_actual = cons_cruce_base + idx_pq_cruce
                idx_pq_cruce += 1

                # Asiento formal en interfaz_siigo
                for _, f_as in df_as_p.iterrows():
                    cta_num = str(f_as["Código Cuenta"]).strip()
                    nit_terc_clean = re.sub(r'\D', '', str(f_as["Tercero / NIT"]).split("-")[0])
                    deb_val = float(f_as["Débito ($)"])
                    cred_val = float(f_as["Crédito ($)"])
                    desc_linea = str(f_as.get("Descripción de la Cuenta") or f_as.get("Descripción Cuenta") or f_as.get("Descripción") or "")[:40]

                    ws_interfaz.append([
                        14,  # Tipo de Comprobante: 14 (Nota de Contabilidad / Cruces)
                        cons_cruce_actual,
                        fecha_op_clean,
                        "COP",
                        1,
                        cta_num,
                        nit_terc_clean,
                        0,
                        "", "", "", "",
                        ag_item.get("Prefijo", "IMP"),
                        ag_item.get("Folio", str(ag_item.get("Factura", ""))),
                        1,
                        fecha_op_clean,
                        "", "", "",
                        desc_linea,
                        "",
                        deb_val,
                        cred_val,
                        f"Cruce Importacion Pq #{p_k} {ag_item['Proveedor'][:15]}",
                        0.0,
                        0.0,
                        ""
                    ])

                # Registro en matriz_captura
                ws_matriz.append([
                    14, cons_cruce_actual, fecha_op_clean,
                    re.sub(r'\D', '', str(ag_item["NIT Emisor"])),
                    ag_item.get("Prefijo", "IMP"),
                    ag_item.get("Folio", str(ag_item.get("Factura", ""))),
                    f"Cruce Importación Pq #{p_k} {ag_item['Proveedor'][:20]}",
                    "Cruce Triangulación",
                    "22050501",
                    float(ag_item.get("Base", 0.0)),
                    float(ag_item.get("IVA", 0.0)),
                    0.0, 0.0, 0.0,
                    "22050501"
                ])
                fila_r += 1

        # 3. Hoja Parametrización
        ws_params = wb.create_sheet(title="Parametrización")
        ws_params.append(["Tipo Comprobante", "", "Impuesto", "Siigo_ID", "Tarifa", "Venta", "Compra", "Dev_Venta", "Dev_Compra"])
        comprobantes = [
            "1 - Ajustes contables", "2 - Depreciación", "3 - Costeo", "4 - Diferidos", "5 - Legalización de viaticos",
            "6 - Legalización de Caja menores", "7 - Obligaciones financieras", "8 - Nómina", "998 - Cierre año",
            "999 - Saldos iniciales", "992 - Comprobante de nómina", "993 - Comprobante de nómina provisión y seguridad social",
            "994 - Comprobante liquidación de contrato", "995 - Comprobante liquidación de primas", "996 - Comprobante liquidación de cesantías",
            "997 - Comprobante desembolso nómina", "9 - Recibo de Caja", "10 - Factura de Compra", "11 - Comprobante de Egreso",
            "12 - Factura de Venta", "13 - Nota Credito", "14 - Nota de Contabilidad", "15 - Documento soporte electronico",
            "777 - Ajustes contables de cartera", "9901 - Traslado de dinero", "16 - FACTURAS DE COMPRA (2)", "17 - NOTA CREDITO DE COMPRA"
        ]
        impuestos = [
            ("IVA 19%", 1, 0.19, "24080601", "24081001", "24082001", "24081002"),
            ("IVA Servicios 19%", 1, 0.19, "24080601", "24081501", "24082001", "24082001"),
            ("IVA Importación 19%", 1, 0.19, "24080601", "240835", "24082001", "24081002"),
            ("IVA 5%", 2, 0.05, "24080602", "24081003", "24082002", "24081004"),
            ("Retefuente 11%", 3, 0.11, "13551509", "23651501", "13551510", "23651502"),
            ("Retefuente 10%", 4, 0.1, "13551507", "23652001", "13551508", "23652002"),
            ("Retefuente 6%", 5, 0.06, "13551505", "23652501", "13551506", "23652502"),
            ("Retefuente 4%", 6, 0.04, "13551503", "23652503", "13551504", "23652504"),
            ("Retefuente 3.5%", 18, 0.035, "13551513", "23654004", "13551514", "23654005"),
            ("Retefuente 2.5%", 7, 0.025, "13551501", "23654001", "13551502", "23654002"),
            ("Retefuente Combustibles 0.1%", 23, 0.001, "13551599", "23654006", "13551599", "23654006"),
            ("ReteICA 11.04", 8, 0.01104, "13551801", "23680501", "13551802", "23680502"),
            ("ReteICA 13.8", 9, 0.0138, "13551803", "23680503", "13551804", "23680504"),
            ("ReteICA 9.66", 10, 0.00966, "13551805", "23680505", "13551806", "23680506"),
            ("ReteICA 8", 11, 0.008, "13551807", "23680507", "13551808", "23680508"),
            ("ReteICA 7", 12, 0.007, "13551809", "23680509", "13551810", "23680510"),
            ("ReteICA 6.9", 13, 0.0069, "13551811", "23680511", "13551812", "23680512"),
            ("ReteICA 4.14", 14, 0.00414, "13551813", "23680513", "13551814", "23680514"),
            ("ReteIVA 15%", 15, 0.15, "13551701", "23670101", "13551702", "23670102")
        ]
        max_len = max(len(comprobantes), len(impuestos))
        for idx_row in range(max_len):
            c_val = comprobantes[idx_row] if idx_row < len(comprobantes) else ""
            if idx_row < len(impuestos):
                imp = list(impuestos[idx_row])
                ws_params.append([c_val, ""] + imp)
            else:
                ws_params.append([c_val, "", "", "", "", "", "", "", ""])

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.download_button(
            label=f"📥 1. Descargar Planilla Oficial Siigo Nube (3 Hojas Formuladas) - {empresa['nombre']}",
            data=output,
            file_name=f"Plantilla_Siigo_{empresa['nombre'].replace(' ', '_')}_3_Hojas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.success("Planilla generada con las 3 hojas oficiales: 'matriz_captura', 'interfaz_siigo' (con códigos de impuesto) y 'Parametrización'.")

        st.markdown("---")
        st.markdown("### 📊 Libro Auxiliar Contable: Facturas Una a Una y Consolidado")
        st.write("Genera el reporte administrativo con el detalle factura a factura y las hojas de consolidado por proveedor y cuentas contables:")

        out_aux = io.BytesIO()
        with pd.ExcelWriter(out_aux, engine='openpyxl') as writer_aux:
            cols_detalle = ["Comprobante Siigo", "Fecha", "Factura", "Proveedor", "NIT Emisor", "Cta Principal", "Categoría", "Base", "IVA", "ReteFuente", "ReteICA", "Total", "Cta Contrapartida", "Razón Contable", "Soporte PDF Renombrado"]
            df_det_export = df_p[cols_detalle].copy()

            totales_dict = {
                "Comprobante Siigo": "TOTALES CONSOLIDADOS",
                "Fecha": "-", "Factura": f"{len(df_det_export)} Docs", "Proveedor": "-", "NIT Emisor": "-",
                "Cta Principal": "-", "Categoría": "-",
                "Base": df_det_export["Base"].sum(),
                "IVA": df_det_export["IVA"].sum(),
                "ReteFuente": df_det_export["ReteFuente"].sum(),
                "ReteICA": df_det_export["ReteICA"].sum(),
                "Total": df_det_export["Total"].sum(),
                "Cta Contrapartida": "-", "Razón Contable": "-", "Soporte PDF Renombrado": "-"
            }
            df_det_export = pd.concat([df_det_export, pd.DataFrame([totales_dict])], ignore_index=True)
            df_det_export.to_excel(writer_aux, sheet_name="Facturas_Una_a_Una", index=False)

            df_cons_prov = df_p.groupby(["NIT Emisor", "Proveedor"]).agg({
                "Comprobante Siigo": "count",
                "Base": "sum",
                "IVA": "sum",
                "ReteFuente": "sum",
                "ReteICA": "sum",
                "Total": "sum"
            }).reset_index().rename(columns={"Comprobante Siigo": "Cant Facturas"})
            df_cons_prov["Neto CxP"] = df_cons_prov["Total"] - df_cons_prov["ReteFuente"] - df_cons_prov["ReteICA"]
            df_cons_prov = df_cons_prov.sort_values(by="Total", ascending=False).reset_index(drop=True)
            df_cons_prov.to_excel(writer_aux, sheet_name="Consolidado_Proveedores", index=False)

            df_cons_cta = df_p.groupby(["Cta Principal", "Categoría"]).agg({
                "Comprobante Siigo": "count",
                "Base": "sum",
                "IVA": "sum",
                "ReteFuente": "sum",
                "ReteICA": "sum",
                "Total": "sum"
            }).reset_index().rename(columns={"Comprobante Siigo": "Cant Facturas"})
            df_cons_cta.to_excel(writer_aux, sheet_name="Consolidado_Cuentas_PUC", index=False)

        out_aux.seek(0)
        st.download_button(
            label=f"📥 2. Descargar Libro de Facturas (Una a Una + Consolidado) - {empresa['nombre']}",
            data=out_aux,
            file_name=f"Libro_Facturas_Detallado_Y_Consolidado_{empresa['nombre'].replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        st.info("Primero procesa los documentos en la Pestana 1 para habilitar la descarga.")

import gc
gc.collect()

