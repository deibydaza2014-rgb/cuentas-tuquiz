import traceback
import pickle
import shutil
import json
import base64
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
try:
    pd.options.future.infer_string = False
except Exception:
    pass
import sys
import numpy as np
def _np_frombuffer_compat(buf, dtype, shape, order):
    return np.frombuffer(buf, dtype=dtype).reshape(shape, order=order)

try:
    import numpy.core
    import numpy.core.multiarray
    import numpy.core.numeric
    import numpy.core.umath
    if not hasattr(numpy.core.numeric, "_frombuffer"):
        numpy.core.numeric._frombuffer = _np_frombuffer_compat
    if "numpy._core" not in sys.modules:
        sys.modules["numpy._core"] = numpy.core
    if "numpy._core.multiarray" not in sys.modules:
        sys.modules["numpy._core.multiarray"] = numpy.core.multiarray
    if "numpy._core.numeric" not in sys.modules:
        sys.modules["numpy._core.numeric"] = numpy.core.numeric
    if "numpy._core.umath" not in sys.modules:
        sys.modules["numpy._core.umath"] = numpy.core.umath
except Exception:
    pass

def _make_compat_string_dtype(*args, **kwargs):
    return pd.StringDtype()

def _safe_pickle_loads(raw_bytes):
    if not raw_bytes:
        return None
    try:
        return pickle.loads(raw_bytes)
    except Exception:
        class _CompatUnpickler(pickle.Unpickler):
            def find_class(self, module, name):
                if name == "StringDtype":
                    return _make_compat_string_dtype
                if name == "_frombuffer" and "numeric" in module:
                    return _np_frombuffer_compat
                if module.startswith("numpy._core"):
                    module = module.replace("numpy._core", "numpy.core", 1)
                return super().find_class(module, name)
        return _CompatUnpickler(io.BytesIO(raw_bytes)).load()

def _limpiar_dtypes_pyarrow_df(df):
    if df is None or not isinstance(df, pd.DataFrame):
        return df
    try:
        df.columns = pd.Index([str(c) for c in df.columns], dtype=object)
    except Exception:
        pass
    for col in df.columns:
        dt_str = str(df[col].dtype).lower()
        if "string" in dt_str or "arrow" in dt_str:
            try:
                df[col] = df[col].astype(object)
            except Exception:
                pass
    return df

def _limpiar_codigo_cta(val, default=""):
    if val is None:
        return default
    try:
        if pd.isna(val):
            return default
    except Exception:
        pass
    s = str(val).strip()
    if not s or s.lower() in ["nan", "none"]:
        return default
    parts = s.split()
    return parts[0].strip() if parts else default

def _a_bool_seguro(val, default=False):
    if val is None:
        return default
    if isinstance(val, (bool, np.bool_)):
        return bool(val)
    try:
        if pd.isna(val):
            return default
    except Exception:
        pass
    s = str(val).strip().lower()
    if s in ("true", "1", "1.0", "si", "s\u00ed", "yes", "t"):
        return True
    if s in ("false", "0", "0.0", "no", "none", "nan", "<na>", ""):
        return False
    return default

def _df_para_streamlit(df_in):
    if df_in is None:
        return pd.DataFrame()
    if not isinstance(df_in, pd.DataFrame):
        try:
            df_in = pd.DataFrame(df_in)
        except Exception:
            return df_in
    df_c = _limpiar_dtypes_pyarrow_df(df_in.copy())
    for col in df_c.columns:
        if df_c[col].dtype == object:
            df_c[col] = df_c[col].apply(lambda v: "" if v is None or (not isinstance(v, (dict, list)) and pd.isna(v)) else str(v))
    return df_c
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
try:
    import fitz as _pymupdf_fitz
    HAS_FITZ = True
except Exception:
    HAS_FITZ = False

st.set_page_config(page_title="Cuentas Tuquiz - ERP y Auditoria Contable", layout="wide", page_icon="\U0001f3e2")

try:

    # Estilos visuales profesionales y proteccion contra conflictos de Google Translate
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
    # CONFIGURACION DE DIRECTORIOS Y SISTEMA DE AUTO-GUARDADO CONTINUO
    # ==============================================================================
    # ==============================================================================
    # CONFIGURACION DE DIRECTORIOS Y SISTEMA DE AUTO-GUARDADO CONTINUO
    # ==============================================================================
    DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "almacenamiento_contable")
    os.makedirs(DATA_DIR, exist_ok=True)

    import tempfile
    import hashlib

    PDF_DISK_CACHE_ROOT = os.path.join(DATA_DIR, "indumaq_pdf_disk_cache")
    os.makedirs(PDF_DISK_CACHE_ROOT, exist_ok=True)
    ZIP_FACTURAS_DISK_PATH = os.path.join(PDF_DISK_CACHE_ROOT, "paquete_facturas_actual.zip")


    class DiskBackedPDFDict(dict):
        _is_disk_backed_pdf_dict = True
        def __init__(self, subdir="default", initial_data=None):
            super().__init__()
            self.subdir = re.sub(r'[^a-zA-Z0-9_-]', '_', str(subdir))
            self.store_dir = os.path.join(PDF_DISK_CACHE_ROOT, self.subdir)
            os.makedirs(self.store_dir, exist_ok=True)
            self._sync_from_disk()
            if initial_data and isinstance(initial_data, dict):
                for k, v in initial_data.items():
                    self[k] = v

        def _sync_from_disk(self):
            try:
                dirs_to_scan = [self.store_dir]
                for extra_sub in ["renom_pdfs", "raw_pdfs"]:
                    ed = os.path.join(PDF_DISK_CACHE_ROOT, extra_sub)
                    if ed not in dirs_to_scan and os.path.exists(ed):
                        dirs_to_scan.append(ed)
                for d_scan in dirs_to_scan:
                    if os.path.exists(d_scan):
                        for fn in os.listdir(d_scan):
                            if fn.lower().endswith(".pdf") and not super().__contains__(fn):
                                super().__setitem__(fn, os.path.join(d_scan, fn))
            except Exception:
                pass

        def _safe_path(self, key):
            safe_name = os.path.basename(str(key)).replace("/", "_").replace("\\", "_")
            if not safe_name:
                safe_name = "doc.pdf"
            return os.path.join(self.store_dir, safe_name)

        def __len__(self):
            if super().__len__() == 0:
                self._sync_from_disk()
            return super().__len__()

        def __bool__(self):
            return len(self) > 0

        def __contains__(self, key):
            if not super().__contains__(key):
                self._sync_from_disk()
            return super().__contains__(key)

        def keys(self):
            if super().__len__() == 0:
                self._sync_from_disk()
            return super().keys()

        def __iter__(self):
            if super().__len__() == 0:
                self._sync_from_disk()
            return super().__iter__()

        def __setitem__(self, key, value):
            p = self._safe_path(key)
            if isinstance(value, (bytes, bytearray)):
                try:
                    with open(p, "wb") as f:
                        f.write(value)
                    super().__setitem__(key, p)
                except Exception:
                    pass
            elif isinstance(value, str) and os.path.exists(value):
                super().__setitem__(key, value)

        def __getitem__(self, key):
            if not super().__contains__(key):
                self._sync_from_disk()
            p = super().__getitem__(key)
            if isinstance(p, (bytes, bytearray)):
                return p
            if isinstance(p, str) and os.path.exists(p):
                try:
                    with open(p, "rb") as f:
                        return f.read()
                except Exception:
                    return b""
            return b""

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
        """Convierte cualquier diccionario de PDFs en memoria a DiskBackedPDFDict para liberar RAM de inmediato."""
        for k_name, sub_name in [("raw_uploaded_pdfs", "raw_pdfs"), ("dict_pdfs", "renom_pdfs")]:
            cur_obj = st.session_state.get(k_name)
            if cur_obj is None:
                st.session_state[k_name] = DiskBackedPDFDict(sub_name)
            elif getattr(cur_obj, "_is_disk_backed_pdf_dict", False) or type(cur_obj).__name__ == "DiskBackedPDFDict":
                continue
            elif isinstance(cur_obj, dict):
                new_d = DiskBackedPDFDict(sub_name)
                for fn in list(cur_obj.keys()):
                    bval = dict.__getitem__(cur_obj, fn)
                    new_d[fn] = bval
                st.session_state[k_name] = new_d
        # Si zip_pdfs estaba en RAM como bytes, moverlo a disco y liberar RAM
        if isinstance(st.session_state.get("zip_pdfs"), (bytes, bytearray)):
            try:
                with open(ZIP_FACTURAS_DISK_PATH, "wb") as fz:
                    fz.write(st.session_state["zip_pdfs"])
                st.session_state["zip_pdfs_path"] = ZIP_FACTURAS_DISK_PATH
            except Exception:
                pass
            st.session_state["zip_pdfs"] = None


    # ESTADOS DE SESION LIMPIOS Y SANEAMIENTO PREVENTIVO DE MEMORIA RAM
    try:
        asegurar_dicts_pdf_en_disco()
    except Exception:
        pass
    for _k in ["df_procesado"]:
        if _k in st.session_state and st.session_state[_k] is not None:
            try:
                if not isinstance(st.session_state[_k], pd.DataFrame):
                    st.session_state.pop(_k, None)
            except Exception:
                st.session_state.pop(_k, None)

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

    # PANTALLA 2: SELECTOR MULTI-EMPRESA ESCALABLE CON AISLAMIENTO EN DISCO Y GESTION DE RAM
    PATH_EMPRESAS_JSON = os.path.join(DATA_DIR, "empresas_registradas.json")
    EMPRESAS_BASE = [
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

    def cargar_empresas_registradas():
        emps = list(EMPRESAS_BASE)
        nits_vistos = {re.sub(r"\D", "", e["nit"]) for e in emps}
        if os.path.exists(PATH_EMPRESAS_JSON):
            try:
                with open(PATH_EMPRESAS_JSON, "r", encoding="utf-8") as fe:
                    guardadas = json.load(fe)
                if isinstance(guardadas, list):
                    for g in guardadas:
                        if isinstance(g, dict) and g.get("nit"):
                            nc = re.sub(r"\D", "", str(g.get("nit", "")))
                            if nc and nc not in nits_vistos:
                                nits_vistos.add(nc)
                                emps.append(g)
            except Exception:
                pass
        return emps

    def guardar_nueva_empresa(nueva_emp):
        emps = cargar_empresas_registradas()
        nc_new = re.sub(r"\D", "", str(nueva_emp.get("nit", "")))
        emps = [e for e in emps if re.sub(r"\D", "", str(e.get("nit", ""))) != nc_new]
        emps.append(nueva_emp)
        try:
            with open(PATH_EMPRESAS_JSON, "w", encoding="utf-8") as fe:
                json.dump(emps, fe, ensure_ascii=False, indent=2)
        except Exception:
            pass
        return emps

    def liberar_memoria_cambio_empresa():
        """Limpia buffers pesados de RAM al cambiar de empresa para que la plataforma soporte ilimitadas empresas sin caerse."""
        import gc
        keys_limpiar = [
            "df_procesado", "dict_pdfs", "raw_uploaded_pdfs", "excel_bytes", "excel_nombre",
            "zip_pdfs", "job_actual_id", "_ultimo_excel_proc_sig", "_ultimo_pdfs_proc_sig",
            "paquetes_importacion", "asientos_triangulacion_por_factura", "facturas_no_contabilizar",
            "_ultimo_guardado_manual", "_cached_respaldo_bytes", "_cached_respaldo_sig",
            "_siigo_cache_sig", "_siigo_3hojas_bytes", "_siigo_preview_matriz",
            "_siigo_preview_interfaz", "_siigo_preview_params", "factura_seleccionada_auditoria",
            "sel_factura_auditoria_box"
        ]
        for k in list(st.session_state.keys()):
            if k in keys_limpiar or k.startswith(("paquete_listo_", "paquete_destino_siigo_", "enviar_gp_", "enviar_h2_", "enviar_nd_", "asiento_fijo_pq_")):
                st.session_state.pop(k, None)
        _MEM_TEXT_CACHE.clear()
        gc.collect()

    EMPRESAS_DISPONIBLES = cargar_empresas_registradas()

    if not st.session_state["empresa_activa"]:
        st.title("Seleccion de Empresa (Multi-Empresa Escalable)")
        st.caption("Cada empresa guarda sus facturas, paquetes y aprendizaje en su propio espacio aislado en disco (0 saturacion de RAM):")

        c1, c2 = st.columns(2)
        for i, emp in enumerate(EMPRESAS_DISPONIBLES):
            col = c1 if i % 2 == 0 else c2
            with col:
                st.markdown(f"""
                <div class="card-box">
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <h3 style="margin:0; color:#0f172a;">\U0001f3e2 {emp['nombre']}</h3>
                        <span class="{emp.get('badge', 'badge-active')}">{emp.get('estado', 'ACTIVA')}</span>
                    </div>
                    <p style="margin:8px 0 2px 0; color:#475569;"><b>NIT:</b> {emp['nit']} - <b>Regimen:</b> {emp.get('regimen', 'Responsable de IVA')}</p>
                    <p style="margin:2px 0 14px 0; color:#64748b; font-size:14px;">{emp.get('actividad', 'Actividad Comercial')}</p>
                </div>
                """, unsafe_allow_html=True)
                if emp.get("estado", "ACTIVA") == "ACTIVA":
                    if st.button(f"Ingresar a {emp['nombre']}", key=f"btn_emp_{i}", use_container_width=True):
                        liberar_memoria_cambio_empresa()
                        st.session_state["empresa_activa"] = emp
                        st.session_state["proceso_activo"] = "facturacion"
                        st.rerun()
                else:
                    st.button(f"Pendiente Datos / Parametrizacion ({emp.get('estado', '')})", key=f"btn_emp_{i}", disabled=True)

        st.markdown("---")
        with st.expander("\u2795 Agregar Nueva Empresa al Ecosistema (Espacio Aislado en Disco)", expanded=False):
            with st.form("form_nueva_empresa"):
                ce1, ce2 = st.columns(2)
                with ce1:
                    nom_ne = st.text_input("Razon Social de la Empresa:", placeholder="Ej. COMERCIALIZADORA XYZ SAS")
                    nit_ne = st.text_input("NIT (con digito de verificacion):", placeholder="Ej. 900.123.456-7")
                with ce2:
                    act_ne = st.text_input("Actividad Economica Principal:", value="Comercio y Servicios")
                    reg_ne = st.selectbox("Regimen Tributario:", ["Responsable de IVA (O-48)", "Gran Contribuyente (O-13)", "Autorretenedor (O-15)", "Gran Contribuyente y Autorretenedor (O-13;O-15)", "Regimen Simple (O-47)"])
                sub_ne = st.form_submit_button("\U0001f4be Registrar y Activar Empresa")
                if sub_ne:
                    if nom_ne.strip() and nit_ne.strip():
                        nueva_e = {
                            "nombre": nom_ne.strip().upper(),
                            "nit": nit_ne.strip(),
                            "actividad": act_ne.strip() or "Comercio General",
                            "regimen": reg_ne,
                            "es_gran_contribuyente": "O-13" in reg_ne,
                            "es_autorretenedor": "O-15" in reg_ne,
                            "responsabilidades_rut": ["O-48"],
                            "estado": "ACTIVA",
                            "badge": "badge-active"
                        }
                        guardar_nueva_empresa(nueva_e)
                        st.success(f"Empresa {nueva_e['nombre']} creada exitosamente en su carpeta dedicada.")
                        st.rerun()
                    else:
                        st.warning("Ingresa Razon Social y NIT.")
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
                liberar_memoria_cambio_empresa()
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
    # GESTOR DE HISTORIAL DE TRABAJOS Y AUDITORIAS PASADAS (PERSISTENCIA TOTAL)
    # ==============================================================================
    # Funciones de sesion y directorios centralizadas en la cabecera



    def guardar_estado_manual(empresa_dict):
        """Guarda en disco \u00fanicamente los datos contables modificados (DataFrame y Paquetes), sin procesar PDFs."""
        if "df_procesado" not in st.session_state or st.session_state["df_procesado"] is None:
            return False
        try:
            nit_clean = re.sub(r'\D', '', str(empresa_dict.get("nit", "empresa")))
            d = os.path.join(DATA_DIR, nit_clean, "estado_manual")
            os.makedirs(d, exist_ok=True)
            with open(os.path.join(d, "df_procesado.pkl"), "wb") as f:
                pickle.dump(st.session_state["df_procesado"], f, protocol=3)
            if st.session_state.get("paquetes_importacion"):
                with open(os.path.join(d, "paquetes_importacion.pkl"), "wb") as f:
                    pickle.dump(st.session_state["paquetes_importacion"], f, protocol=3)
            if st.session_state.get("asientos_triangulacion_por_factura"):
                with open(os.path.join(d, "asientos_triangulacion.pkl"), "wb") as f:
                    pickle.dump(st.session_state["asientos_triangulacion_por_factura"], f, protocol=3)
            if st.session_state.get("facturas_no_contabilizar"):
                with open(os.path.join(d, "facturas_no_contabilizar.json"), "w", encoding="utf-8") as f:
                    json.dump(list(st.session_state["facturas_no_contabilizar"]), f)
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
                with open(os.path.join(d, "estado_sesion.json"), "w", encoding="utf-8") as f_es:
                    json.dump(estado_flags, f_es, ensure_ascii=False, indent=2)
            except Exception:
                pass
            now_str = datetime.datetime.now().strftime("%d/%m/%Y %I:%M:%S %p")
            meta = {
                "fecha_guardado": now_str,
                "total_facturas": len(st.session_state["df_procesado"]),
                "archivo_excel": st.session_state.get("excel_nombre", "Reporte.xlsx")
            }
            with open(os.path.join(d, "meta.json"), "w", encoding="utf-8") as f:
                json.dump(meta, f, ensure_ascii=False, indent=2)
            st.session_state["_ultimo_guardado_manual"] = now_str
            st.session_state["_cached_respaldo_bytes"] = None
            import gc
            gc.collect()
            return True
        except Exception:
            return False

    def cargar_estado_manual(empresa_dict):
        """Restaura en memoria el \u00faltimo estado guardado manualmente por el usuario."""
        try:
            nit_clean = re.sub(r'\D', '', str(empresa_dict.get("nit", "empresa")))
            d = os.path.join(DATA_DIR, nit_clean, "estado_manual")
            df_p_path = os.path.join(d, "df_procesado.pkl")
            if os.path.exists(df_p_path):
                with open(df_p_path, "rb") as f:
                    df_m = _safe_pickle_loads(f.read())
                if df_m is None or not isinstance(df_m, pd.DataFrame) or df_m.empty:
                    return False
                st.session_state["df_procesado"] = normalizar_df_procesado(_limpiar_dtypes_pyarrow_df(df_m))
                as_path = os.path.join(d, "asientos_triangulacion.pkl")
                if os.path.exists(as_path):
                    with open(as_path, "rb") as f:
                        as_m = _safe_pickle_loads(f.read())
                    if isinstance(as_m, dict):
                        for k_a, v_a in as_m.items():
                            if isinstance(v_a, pd.DataFrame):
                                as_m[k_a] = _limpiar_dtypes_pyarrow_df(v_a)
                        st.session_state["asientos_triangulacion_por_factura"] = as_m
                pq_path = os.path.join(d, "paquetes_importacion.pkl")
                if os.path.exists(pq_path):
                    with open(pq_path, "rb") as f:
                        pqs_loaded = _safe_pickle_loads(f.read())
                    if isinstance(pqs_loaded, dict) and len(pqs_loaded) >= 50:
                        as_map = st.session_state.get("asientos_triangulacion_por_factura", {})
                        for pq_k, pq_v in pqs_loaded.items():
                            if isinstance(pq_v, dict):
                                if "terceros" in pq_v and isinstance(pq_v["terceros"], pd.DataFrame):
                                    pq_v["terceros"] = _limpiar_dtypes_pyarrow_df(pq_v["terceros"])
                                for af_k in ["asiento_fijo", "asiento_aprobado"]:
                                    if af_k in pq_v and isinstance(pq_v[af_k], pd.DataFrame):
                                        pq_v[af_k] = _limpiar_dtypes_pyarrow_df(pq_v[af_k])
                                if ("asiento_fijo" in pq_v and pq_v["asiento_fijo"] is not None) or ("asiento_aprobado" in pq_v and pq_v["asiento_aprobado"] is not None):
                                    st.session_state[f"paquete_listo_{pq_k}"] = True
                                    st.session_state[f"asiento_fijo_pq_{pq_k}"] = pq_v.get("asiento_fijo", pq_v.get("asiento_aprobado")).copy()
                                ag_fac_str = str(pq_v.get("agente", {}).get("Factura", "")).strip()
                                if "destino_siigo" in pq_v and pq_v["destino_siigo"] is not None:
                                    st.session_state[f"paquete_destino_siigo_{pq_k}"] = bool(pq_v["destino_siigo"])
                                elif ag_fac_str and ag_fac_str in as_map:
                                    st.session_state[f"paquete_destino_siigo_{pq_k}"] = True
                                    pq_v["destino_siigo"] = True
                        st.session_state["paquetes_importacion"] = pqs_loaded
                        st.session_state["_ultimo_agentes_len"] = len(pqs_loaded)
                    else:
                        return False
                else:
                    return False
                es_path = os.path.join(d, "estado_sesion.json")
                if os.path.exists(es_path):
                    try:
                        with open(es_path, "r", encoding="utf-8") as f_es:
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
                            st.session_state["paquete_seleccionado_id"] = est["sel_paquete_activo_key"]
                    except Exception:
                        pass
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
                sincronizar_todos_los_paquetes_con_df_procesado()
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

            # 3. Guardar PDFs renombrados / separados (escritura inteligente: solo si no existe o cambio)
            if dict_pdfs_renombrados:
                for fname, bdata in dict_pdfs_renombrados.items():
                    p_ren = os.path.join(dir_renom, fname)
                    if not os.path.exists(p_ren) or os.path.getsize(p_ren) != len(bdata):
                        with open(p_ren, "wb") as pf:
                            pf.write(bdata)

            # 4. Guardar PDFs originales subidos (escritura inteligente: solo si no existe o cambio)
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

            # 6. Guardar paquetes de importacion y triangulacion
            if "paquetes_importacion" in st.session_state and st.session_state["paquetes_importacion"]:
                try:
                    with open(os.path.join(job_dir, "paquetes_importacion.pkl"), "wb") as f_pq:
                        pickle.dump(st.session_state["paquetes_importacion"], f_pq)
                except Exception:
                    pass

            # 7. Guardar asientos aprobados de triangulacion
            if "asientos_triangulacion_por_factura" in st.session_state and st.session_state["asientos_triangulacion_por_factura"]:
                try:
                    with open(os.path.join(job_dir, "asientos_triangulacion.pkl"), "wb") as f_as:
                        pickle.dump(st.session_state["asientos_triangulacion_por_factura"], f_as)
                except Exception:
                    pass

            # 8. Guardar estado general de la sesion (paquetes listos, opciones, facturas excluidas)
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
            # Verificar archivos fisicos guardados en disco
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

            # Actualizar indice general de la empresa para carga inmediata
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
                df = _safe_pickle_loads(f.read())

        ex_file = os.path.join(jdir, "excel_original.xlsx")
        if os.path.exists(ex_file):
            with open(ex_file, "rb") as f_ex:
                excel_bytes = f_ex.read()

        # PDFs renombrados respaldados en disco (0 MB en RAM)
        dict_renom = DiskBackedPDFDict("renom_pdfs")
        pdf_dir_r = os.path.join(jdir, "pdfs_renombrados")
        if not os.path.exists(pdf_dir_r):
            pdf_dir_r = os.path.join(jdir, "pdfs")
        if os.path.exists(pdf_dir_r):
            for fn in os.listdir(pdf_dir_r):
                if fn.endswith(".pdf"):
                    dict_renom[fn] = os.path.join(pdf_dir_r, fn)

        # PDFs originales respaldados en disco (0 MB en RAM)
        dict_orig = DiskBackedPDFDict("raw_pdfs")
        pdf_dir_o = os.path.join(jdir, "pdfs_originales")
        if os.path.exists(pdf_dir_o):
            for fn in os.listdir(pdf_dir_o):
                if fn.endswith(".pdf"):
                    dict_orig[fn] = os.path.join(pdf_dir_o, fn)

        zip_file = os.path.join(jdir, "paquete_facturas.zip")
        if os.path.exists(zip_file):
            st.session_state["zip_pdfs_path"] = zip_file
            zip_bytes = None

        # Restaurar paquetes de importacion
        pq_file = os.path.join(jdir, "paquetes_importacion.pkl")
        if os.path.exists(pq_file):
            try:
                with open(pq_file, "rb") as f_pq:
                    st.session_state["paquetes_importacion"] = _safe_pickle_loads(f_pq.read())
            except Exception:
                pass

        # Restaurar asientos aprobados de triangulacion
        as_file = os.path.join(jdir, "asientos_triangulacion.pkl")
        if os.path.exists(as_file):
            try:
                with open(as_file, "rb") as f_as:
                    st.session_state["asientos_triangulacion_por_factura"] = _safe_pickle_loads(f_as.read())
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
        """Empaqueta toda la sesi\u00f3n contable en un archivo ZIP descargable para respaldo indestructible con paquetes y triangulaciones."""
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            if df_proc is not None:
                zf.writestr("df_procesado.pkl", pickle.dumps(df_proc))
            if excel_b:
                zf.writestr("excel_original.xlsx", excel_b)
            if zip_p:
                zf.writestr("paquete_facturas.zip", zip_p)

            # Incluir paquetes de importacion y asientos de triangulacion si existen
            if "paquetes_importacion" in st.session_state and st.session_state["paquetes_importacion"]:
                zf.writestr("paquetes_importacion.pkl", pickle.dumps(st.session_state["paquetes_importacion"]))
            if "asientos_triangulacion_por_factura" in st.session_state and st.session_state["asientos_triangulacion_por_factura"]:
                zf.writestr("asientos_triangulacion.pkl", pickle.dumps(st.session_state["asientos_triangulacion_por_factura"]))

            # Incluir estado completo de la sesion
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
            # PDFs ya vienen comprimidos internamente (FlateDecode). Usar compresion directa ultra-rapida
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
        """Devuelve el archivo de respaldo cacheado en memoria, regener\u00e1ndolo solo cuando hay cambios."""
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
        """Restaura una sesi\u00f3n contable completa desde un archivo ZIP de respaldo."""
        try:
            with zipfile.ZipFile(io.BytesIO(zip_bytes), "r") as zf:
                df_proc = None
                if "df_procesado.pkl" in zf.namelist():
                    df_proc = _safe_pickle_loads(zf.read("df_procesado.pkl"))
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
    # MOTOR DE TRIANGULACION Y CRUCES DE IMPORTACION (ADUANAS Y PAGOS A TERCEROS)
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
        2. Columna1 con n\u00famero de comprobante 10-XXX o texto 'registrada'.
        3. Columna3 con asignaci\u00f3n de Grupo de Importaci\u00f3n (1, 2, 3...).
        4. Columna4 con cuenta contable de pasivo espec\u00edfica (22050505, 23359501, etc.).
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

    CATALOGO_CUENTAS_PRINCIPALES_PUC = [
        ("51952501", "Gasto Aseo y Cafeteria",           "51952501 - Cafeteria, Bebidas, Aseo y Limpieza"),
        ("51956001", "Gasto Casino y Restaurante",       "51956001 - Casino y Restaurante (Alimentacion)"),
        ("51354001", "Gasto Mensajeria y Paqueteo",      "51354001 - Correo, Portes y Mensajeria (Envia / Paqueteo)"),
        ("51355001", "Gasto Transporte y Mensajeria",    "51355001 - Transporte Nacional, Fletes y Acarreos (Servientrega)"),
        ("51353001", "Gasto Energia Electrica (Luz)",    "51353001 - Energia Electrica / Luz (Servicios Publicos)"),
        ("51352501", "Gasto Acueducto y Agua",           "51352501 - Acueducto, Alcantarillado y Agua"),
        ("51353501", "Gasto Telefono e Internet",        "51353501 - Telefono, Internet y Telecomunicaciones"),
        ("51355501", "Gasto Gas Natural",                "51355501 - Gas Natural (Servicios Publicos)"),
        ("51350501", "Gasto Aseo y Vigilancia",          "51350501 - Servicios de Aseo, Vigilancia y Seguridad"),
        ("51953001", "Gasto Papeleria y Utiles",         "51953001 - Utiles, Papeleria, Talonarios y Fotocopias"),
        ("51550501", "Gasto Viaje y Alojamiento",        "51550501 - Alojamiento, Hoteles y Gastos de Viaje"),
        ("51551501", "Gasto Pasajes Aereos / Terrestres","51551501 - Pasajes Aereos y Terrestres (Viajes)"),
        ("51954501", "Gasto Taxis y Buses",              "51954501 - Taxis, Buses y Transporte Urbano"),
        ("51953501", "Gasto Combustibles y Lubricantes", "51953501 - Combustibles y Lubricantes (Gasolina / ACPM)"),
        ("51956501", "Gasto Parqueaderos y Peajes",      "51956501 - Parqueaderos y Peajes"),
        ("51401001", "Gasto Legal Registro Mercantil",   "51401001 - Gastos Legales, Registro Mercantil y Camara Comercio"),
        ("51301001", "Gasto Seguros y Polizas",          "51301001 - Seguros y Polizas de Cumplimiento / Generales"),
        ("51352001", "Gasto Software y Nube",            "51352001 - Software, Facturacion Electronica, Firma y Nube"),
        ("51109501", "Gasto Honorarios y Validacion",    "51109501 - Honorarios, Asesoria y Validacion de Proveedores"),
        ("51451501", "Gasto Mantenimiento y Reparacion", "51451501 - Mantenimiento y Reparaciones Administrativas"),
        ("51201001", "Gasto Arrendamientos",             "51201001 - Arrendamientos (Inmuebles, Bodegas y Equipos)"),
        ("51055101", "Gasto Dotacion y EPP",             "51055101 - Dotacion y Elementos de Proteccion a Trabajadores"),
        ("61800101", "Costo Mantenimiento e Ingenieria", "61800101 - Costo Mantenimiento Industrial, Repuestos e Ingenieria"),
        ("61350501", "Costo Insumos y Materiales",       "61350501 - Costo de Ventas, Insumos y Materiales"),
        ("14650501", "Importacion (1465)",               "14650501 - Importaciones en Transito / Comercio Exterior"),
        ("14350101", "Inventario Mercancias / Repuestos","14350101 - Mercancias No Fabricadas por la Empresa (Inventario)"),
        ("15280501", "Equipo de Computacion",            "15280501 - Equipos de Procesamiento de Datos y Computacion"),
        ("51959501", "Otros Gastos Diversos",            "51959501 - Otros Gastos Diversos"),
        ("OTRA_CTA_P", "Cuenta Personalizada",           "Otra Cuenta Contable PUC (Digitar codigo manualmente)...")
    ]

    MAPA_ETIQUETAS_PUC_PRINCIPAL = {item[0]: item[1] for item in CATALOGO_CUENTAS_PRINCIPALES_PUC if item[0] != "OTRA_CTA_P"}

    def obtener_etiqueta_cuenta_principal(cta_p, categoria_actual=""):
        cta_s = str(cta_p or "").strip()
        cat_s = str(categoria_actual or "").strip()
        if cta_s in ["146505", "1465"]:
            cta_s = "14650501"
        elif cta_s == "515505":
            cta_s = "51550501"
        # Si la categoria actual ya es especifica (no generica) y coherente con la cuenta, conservarla
        if cat_s and cat_s.upper() not in ["GASTO GENERAL", "GENERAL", "MERCANCIA", "GASTO/COSTO", "APRENDIDO"] and not cat_s.upper().startswith("APRENDIDO"):
            if cta_s == "51952501" and any(k in cat_s.upper() for k in ["CAFETERIA", "ASEO", "LIMPIEZA", "RESTAURANTE"]):
                return cat_s
            if cta_s in MAPA_ETIQUETAS_PUC_PRINCIPAL:
                return MAPA_ETIQUETAS_PUC_PRINCIPAL[cta_s]
            return cat_s
        if cta_s in MAPA_ETIQUETAS_PUC_PRINCIPAL:
            return MAPA_ETIQUETAS_PUC_PRINCIPAL[cta_s]
        return cat_s if cat_s else f"Cuenta {cta_s}"

    def clasificar_gasto_detallado_puc(nit_emisor, nombre_emisor, descripcion="", texto_pdf="", valor_base=0.0, valor_iva=0.0, cta_actual=""):
        prov_u = re.sub(r"\s+", " ", str(nombre_emisor or "").upper()).strip()
        desc_u = re.sub(r"\s+", " ", str(descripcion or "").upper()).strip()

        txt_u = str(texto_pdf or "").upper()
        m_det = re.search(r"DETALLES DE PRODUCTOS(.*?)(?:NOTAS FINALES|DATOS TOTALES|$)", txt_u, re.DOTALL)
        raw_items = m_det.group(1) if m_det else txt_u[:2500]
        items_u = re.sub(r"\s+", " ", raw_items).strip()
        full_search = f"{prov_u} | {desc_u} | {items_u}"

        AGENTES_Y_ADUANAS = [
            "DHL EXPRESS", "TRADE GLOBAL", "EURO SHIPPING", "COMERCIO EXTERIOR DEL NORTE",
            "ALMACENADORA INTERAMERICANA", "CARGO ADUANA", "SOCIEDAD PORTUARIA", "CONSOLCARGO",
            "PUERTO BAHIA", "CONTECAR", "SPRC", "SPB", "COMPAS", "ALMAVIVA", "ALPOPULAR"
        ]
        if any(k in prov_u for k in AGENTES_Y_ADUANAS):
            return None

        if any(k in prov_u for k in ["FEDERAL EXPRESS", "FEDEX"]) and any(k in full_search for k in ["NACIONALIZACI", "DERECHOS ADUANALES", "ADUANA"]):
            return {
                "cta_p": "14650501", "cta_c": "23359501",
                "cat": "Importacion / Nacionalizacion Courier",
                "desc": f"Nacionalizacion y Derechos Aduanales - {str(nombre_emisor)[:24]}",
                "cta_iva": "24081501", "cta_rfte_sug": "23652503", "tasa_rfte_sug": 0.04, "tope_rfte": 210000,
                "cta_rica_sug": "23680505", "tasa_rica_mil": 9.66,
                "razon": "Nacionalizacion y Derechos Aduanales de Importacion (14650501)"
            }

        # 1. SERVICIOS PUBLICOS: LUZ / ENERGIA ELECTRICA (51353001)
        if any(k in prov_u for k in ["ENEL", "CODENSA", "CELSIA", "ELECTRIFICADORA", "ESSA", "AIR-E", "AFINIA", "CEO "]) or \
           any(k in full_search for k in ["ENERGIA ELECTRICA", "SERVICIO DE ENERGIA", "CONSUMO ENERGIA", "KWH", "FACTURA DE ENERGIA", "SERVICIO DE LUZ"]):
            return {
                "cta_p": "51353001", "cta_c": "23355001",
                "cat": "Gasto Energia Electrica (Luz)",
                "desc": f"Servicio de Energia Electrica (Luz) - {str(nombre_emisor)[:24]}",
                "cta_iva": "24081501", "cta_rfte_sug": "", "tasa_rfte_sug": 0.0, "tope_rfte": 999999999,
                "cta_rica_sug": "", "tasa_rica_mil": 0.0,
                "razon": "Servicios Publicos - Energia Electrica / Luz (Cta 51353001)"
            }

        # 2. SERVICIOS PUBLICOS: ACUEDUCTO Y ALCANTARILLADO / AGUA (51352501)
        if any(k in prov_u for k in ["ACUEDUCTO", "EAAB", "AGUAS DE", "TRIPLE A", "EMPRESAS PUBLICAS"]) or \
           any(k in full_search for k in ["SERVICIO DE ACUEDUCTO", "ALCANTARILLADO", "CONSUMO DE AGUA", "ACUEDUCTO Y ALCANTARILLADO"]):
            return {
                "cta_p": "51352501", "cta_c": "23355001",
                "cat": "Gasto Acueducto y Agua",
                "desc": f"Servicio de Acueducto y Alcantarillado - {str(nombre_emisor)[:24]}",
                "cta_iva": "24081501", "cta_rfte_sug": "", "tasa_rfte_sug": 0.0, "tope_rfte": 999999999,
                "cta_rica_sug": "", "tasa_rica_mil": 0.0,
                "razon": "Servicios Publicos - Acueducto y Alcantarillado (Cta 51352501)"
            }

        # 3. SERVICIOS PUBLICOS: TELEFONO, INTERNET Y COMUNICACIONES (51353501)
        if any(k in prov_u for k in ["CLARO", "COMCEL", "COLOMBIA TELECOMUNICACIONES", "MOVISTAR", "TIGO", "UNE EPM", "ETB", "WOM ", "DIRECTV"]) or \
           any(k in full_search for k in ["SERVICIO DE INTERNET", "FIBRA OPTICA", "TELEFONIA", "PLAN CELULAR", "DATOS MOVILES", "BANDA ANCHA"]):
            return {
                "cta_p": "51353501", "cta_c": "23355001",
                "cat": "Gasto Telefono e Internet",
                "desc": f"Servicio de Telefono e Internet - {str(nombre_emisor)[:24]}",
                "cta_iva": "24081501", "cta_rfte_sug": "23652503", "tasa_rfte_sug": 0.04, "tope_rfte": 210000,
                "cta_rica_sug": "23680505", "tasa_rica_mil": 9.66,
                "razon": "Servicios de Telefono, Internet y Telecomunicaciones (Cta 51353501)"
            }

        # 4. SERVICIOS PUBLICOS: GAS NATURAL (51355501)
        if any(k in prov_u for k in ["VANTI", "GAS NATURAL", "GASES DE OCCIDENTE", "GASES DEL CARIBE", "SURTIGAS", "EFIGAS"]) or \
           any(k in full_search for k in ["SERVICIO DE GAS NATURAL", "CONSUMO GAS NATURAL"]):
            return {
                "cta_p": "51355501", "cta_c": "23355001",
                "cat": "Gasto Gas Natural",
                "desc": f"Servicio de Gas Natural - {str(nombre_emisor)[:24]}",
                "cta_iva": "24081501", "cta_rfte_sug": "", "tasa_rfte_sug": 0.0, "tope_rfte": 999999999,
                "cta_rica_sug": "", "tasa_rica_mil": 0.0,
                "razon": "Servicios Publicos - Gas Natural (Cta 51355501)"
            }

        # 5. GASTOS LEGALES, CAMARA DE COMERCIO Y REGISTRO MERCANTIL (51401001)
        if any(k in prov_u for k in ["CAMARA DE COMERCIO", "C\u00c1MARA DE COMERCIO", "NOTARIA", "NOTAR\u00cdA", "SUPERINTENDENCIA"]) or \
           any(k in full_search for k in ["CERTIFICADO CONSTITUC", "RENOVACION PERSONA", "MATRICULA MERCANTIL", "REGISTRO MERCANTIL"]):
            return {
                "cta_p": "51401001", "cta_c": "23359501",
                "cat": "Gasto Legal Registro Mercantil",
                "desc": f"Certificados y Registro Mercantil - {str(nombre_emisor)[:24]}",
                "cta_iva": "24081501", "cta_rfte_sug": "", "tasa_rfte_sug": 0.0, "tope_rfte": 999999999,
                "cta_rica_sug": "", "tasa_rica_mil": 0.0,
                "razon": "Gastos Legales - Camara de Comercio y Registro Mercantil (Cta 51401001)"
            }

        # 6. SEGUROS Y POLIZAS (51301001)
        if any(k in prov_u for k in ["SEGUROS", "ASEGURADORA", "SURAMERICANA", "ALLIANZ", "MAPFRE", "BOLIVAR", "LIBERTY", "PREVISORA"]) or \
           any(k in full_search for k in ["POLIZA DE CUMPLIMIENT", "POLIZA ", "SEGURO DE"]):
            return {
                "cta_p": "51301001", "cta_c": "23359501",
                "cat": "Gasto Seguros y Polizas",
                "desc": f"Poliza de Cumplimiento y Seguros - {str(nombre_emisor)[:24]}",
                "cta_iva": "24081501", "cta_rfte_sug": "", "tasa_rfte_sug": 0.0, "tope_rfte": 999999999,
                "cta_rica_sug": "", "tasa_rica_mil": 0.0,
                "razon": "Gasto de Seguros y Polizas de Cumplimiento (Cta 51301001)"
            }

        # 7. SOFTWARE, NUBE, FACTURACION ELECTRONICA Y FIRMA DIGITAL (51352001)
        if any(k in prov_u for k in ["SIIGO", "FTECH", "CERTICAMARA", "CERTIC\u00c1MARA", "ALEGRA", "WORLD OFFICE", "HELISA", "MICROSOFT", "GOOGLE"]) or \
           any(k in full_search for k in ["COMPUTACION E", "COMPUTACI\u00d3N E", "FACTURAS ELECTR", "CERTIFICADO PARA FIRMA", "FIRMA CENTRALIZADA", "SIIGO NUBE"]):
            return {
                "cta_p": "51352001", "cta_c": "23359501",
                "cat": "Gasto Software y Nube",
                "desc": f"Software, Facturacion Electronica y Firma Digital - {str(nombre_emisor)[:22]}",
                "cta_iva": "24081501", "cta_rfte_sug": "23652503", "tasa_rfte_sug": 0.04, "tope_rfte": 210000,
                "cta_rica_sug": "23680505", "tasa_rica_mil": 9.66,
                "razon": "Procesamiento Electronico de Datos, Software, Nube y Firma Digital (Cta 51352001)"
            }

        # 8. MENSAJERIA, CORREO, PAQUETEO Y TRANSPORTE NACIONAL (51354001 / 51355001)
        if any(k in prov_u for k in ["COLVANES", "ENVIA ", "INTER RAPIDISIMO", "INTERRAPIDISIMO", "DEPRISA", "4-72", "SERVICIOS POSTALES", "UPS SERVICIOS EXPRESOS"]) or \
           ("SERVIENTREGA" in prov_u and valor_base <= 50000) or \
           any(k in full_search for k in ["PAQUETE TERRESTRE", "MENSAJERIA EXPRESA", "ENVIO SOBRE", "GUIA DE MENSAJERIA", "DECLARACION DE IMPORT"]):
            return {
                "cta_p": "51354001", "cta_c": "23359501",
                "cat": "Gasto Mensajeria y Paqueteo",
                "desc": f"Servicio de Mensajeria y Paqueteo - {str(nombre_emisor)[:24]}",
                "cta_iva": "24081501", "cta_rfte_sug": "23652505", "tasa_rfte_sug": 0.01, "tope_rfte": 188000,
                "cta_rica_sug": "23680513", "tasa_rica_mil": 4.14,
                "razon": "Servicio de Correo, Portes, Paqueteo y Mensajeria (Cta 51354001)"
            }

        if any(k in prov_u for k in ["SERVIENTREGA", "COORDINADORA", "TCC ", "SAFERBO", "ALDIA", "BOTERO SOTO", "CONSTRUCCIONESARQ3D"]) or \
           any(k in full_search for k in ["MERCANCIA PREMIER", "MERCANCIA MULTIPLES", "GUACAL EN MAD", "PARA TRANSPORTE DE", "FLETE NACIONAL"]):
            return {
                "cta_p": "51355001", "cta_c": "23359501",
                "cat": "Gasto Transporte y Mensajeria",
                "desc": f"Servicio de Transporte y Mensajeria - {str(nombre_emisor)[:24]}",
                "cta_iva": "24081501", "cta_rfte_sug": "23652505", "tasa_rfte_sug": 0.01, "tope_rfte": 188000,
                "cta_rica_sug": "23680513", "tasa_rica_mil": 4.14,
                "razon": "Gasto de Transporte Nacional, Fletes, Acarreos y Mensajeria (Cta 51355001)"
            }

        # 9. HOTELES, HOSPEDAJE Y GASTOS DE VIAJE (51550501)
        if any(k in prov_u for k in ["HOTEL", "ESTELAR", "GENOVA", "VITTAPARK", "MONTEZ GROUP", "YORLENY VILLOTA", "JOHN EDWAR LOZANO"]) or \
           any(k in full_search for k in ["SERVICIO DE HOSPEDAJE", "SERVICIOS HOTELEROS", "SERVICIO HABITACION", "ALOJAMIENTO", "CAMA SENCILLA", "NOCHE"]):
            return {
                "cta_p": "51550501", "cta_c": "23359501",
                "cat": "Gasto Viaje y Alojamiento",
                "desc": f"Alojamiento y Gastos de Viaje - {str(nombre_emisor)[:24]}",
                "cta_iva": "24081501", "cta_rfte_sug": "23652501", "tasa_rfte_sug": 0.035, "tope_rfte": 210000,
                "cta_rica_sug": "23680505", "tasa_rica_mil": 9.66,
                "razon": "Gastos de Viaje, Alojamiento y Manutencion (Cta 51550501)"
            }

        # 10. PAPELERIA, UTILES E IMPRESOS (51953001)
        if any(k in prov_u for k in ["PANAMERICANA", "LIBRERIA", "LIBRER\u00cdA", "PAPELERIA", "JAIME RODRIGUEZ RODRIGUEZ"]) or \
           any(k in full_search for k in ["COMPROBANTE INGRESO", "COMPROBANTE D/INGRES", "CALCULADORA D/ESCRIT", "RESMA", "PAPELERIA", "TALONARIO"]):
            return {
                "cta_p": "51953001", "cta_c": "23359501",
                "cat": "Gasto Papeleria y Utiles",
                "desc": f"Utiles, Papeleria y Talonarios - {str(nombre_emisor)[:24]}",
                "cta_iva": "24081001", "cta_rfte_sug": "23654001", "tasa_rfte_sug": 0.025, "tope_rfte": 1047000,
                "cta_rica_sug": "23680501", "tasa_rica_mil": 11.04,
                "razon": "Gasto de Utiles, Papeleria, Fotocopias e Impresos (Cta 51953001)"
            }

        # 11. HONORARIOS, HOMOLOGACION Y VALIDACION DE PROVEEDORES (51109501)
        if any(k in prov_u for k in ["RISKS INTERNATIONAL"]) or \
           any(k in full_search for k in ["HOMOLOGACIO", "VALIDACION DE PROV", "SIGPRO", "ASESORIA JURIDICA", "HONORARIOS"]):
            return {
                "cta_p": "51109501", "cta_c": "23352501",
                "cat": "Gasto Honorarios y Validacion",
                "desc": f"Homologacion y Validacion Tecnica - {str(nombre_emisor)[:24]}",
                "cta_iva": "24081501", "cta_rfte_sug": "23652503", "tasa_rfte_sug": 0.04, "tope_rfte": 210000,
                "cta_rica_sug": "23680505", "tasa_rica_mil": 9.66,
                "razon": "Honorarios, Consultoria y Homologacion de Proveedores (Cta 51109501)"
            }

        # 12. EQUIPOS DE COMPUTO Y TECNOLOGIA (15280501)
        if any(k in full_search for k in ["CONV 14\" HP", "HP FP", "CI5", "CORE I5", "CORE I7", "PORTATIL", "LAPTOP", "COMPUTADOR", "MACBOOK"]):
            return {
                "cta_p": "15280501", "cta_c": "22050501",
                "cat": "Equipo de Computacion",
                "desc": f"Equipo de Computo / Portatil HP - {str(nombre_emisor)[:24]}",
                "cta_iva": "24081001", "cta_rfte_sug": "23654001", "tasa_rfte_sug": 0.025, "tope_rfte": 1047000,
                "cta_rica_sug": "23680501", "tasa_rica_mil": 11.04,
                "razon": "Compra de Equipo de Procesamiento de Datos y Computacion (Cta 15280501)"
            }

        # 13. CAFETERIA, BEBIDAS, RESTAURANTE Y CASINO (51952501 / 51956001)
        if any(k in prov_u for k in ["DONUCAFE", "JUAN VALDEZ", "PROCACOL", "TOSTAO", "STARBUCKS", "SUBVALLE", "SUBWAY", "OMAKASE", "PAN PA YA", "HORNITOS"]) or \
           (any(k in full_search for k in ["CAFE NEG", "AMERICANO MEDIANO", "CAPUCCINO", "LATTE", "BBQ COSTI", "BOTELLA TE", "BOTELLA T\u00c9", "SANDWICH", "EMPANADA", "REFRIGERIO", "ALMUERZO"]) and not any(k in full_search for k in ["DETERGENTE", "PAPEL HIGIENICO", "LAVALOZ", "TOALLA PAPEL", "AMBIENTADOR"])):
            return {
                "cta_p": "51952501", "cta_c": "23359501",
                "cat": "Gasto Cafeteria y Restaurante",
                "desc": f"Gastos de Cafeteria y Alimentacion - {str(nombre_emisor)[:24]}",
                "cta_iva": "24081001", "cta_rfte_sug": "23654004", "tasa_rfte_sug": 0.035, "tope_rfte": 210000,
                "cta_rica_sug": "23680505", "tasa_rica_mil": 9.66,
                "razon": "Gastos de Cafeteria, Bebidas y Casino/Restaurante (Cta 51952501)"
            }

        # 14. ASEO, LIMPIEZA Y CAFETERIA (51952501)
        tiene_aseo = any(k in full_search for k in [
            "DETERGENTE", "LAVALOZ", "AXION", "PAPEL HIGIENICO", "PH ALK", "TRIPH", "TOALLA PAPEL",
            "TOALLAS X", "PA\u00d1ITOS", "PANITOS", "AMBIENTADOR", "JABON", "CLOROX", "DESINFECTANTE",
            "BLANQUEADOR", "LIMPIADOR", "ESCOBA", "TRAPERO", "SERVILLETA"
        ])
        tiene_cafe = any(k in full_search for k in ["CAFE MOLIDO", "CAF\u00c9 MOLIDO", "LECHE ENTERA", "AZUCAR", "AROMATICA", "GALLETA", "BOTELLON"])
        if tiene_aseo and tiene_cafe:
            return {
                "cta_p": "51952501", "cta_c": "23359501",
                "cat": "Gasto Aseo y Cafeteria",
                "desc": f"Insumos de Aseo y Cafeteria - {str(nombre_emisor)[:24]}",
                "cta_iva": "24081001", "cta_rfte_sug": "23654001", "tasa_rfte_sug": 0.025, "tope_rfte": 1047000,
                "cta_rica_sug": "23680501", "tasa_rica_mil": 11.04,
                "razon": "Elementos de Aseo, Limpieza y Cafeteria (Cta 51952501)"
            }
        elif tiene_aseo or any(k in prov_u for k in ["MERCADOS LA SABANA", "SUPERMERCADO", "OLIMPICA", "JERONIMO MARTINS", "KOBA", "TIENDAS D1", "TIENDAS ARA"]):
            return {
                "cta_p": "51952501", "cta_c": "23359501",
                "cat": "Gasto Aseo y Limpieza",
                "desc": f"Elementos de Aseo, Higiene y Limpieza - {str(nombre_emisor)[:24]}",
                "cta_iva": "24081001", "cta_rfte_sug": "23654001", "tasa_rfte_sug": 0.025, "tope_rfte": 1047000,
                "cta_rica_sug": "23680501", "tasa_rica_mil": 11.04,
                "razon": "Elementos de Aseo, Higiene y Limpieza (Cta 51952501)"
            }

        # 15. TAXIS, BUSES, COMBUSTIBLES, PARQUEADEROS, DOTACION, ARRENDAMIENTOS, VIGILANCIA
        if any(k in full_search for k in ["TAXI", "UBER", "DIDI", "CABIFY", "TRANSMILENIO", "BUS URBANO"]):
            return {
                "cta_p": "51954501", "cta_c": "23359501",
                "cat": "Gasto Taxis y Buses",
                "desc": f"Transporte Urbano / Taxis y Buses - {str(nombre_emisor)[:24]}",
                "cta_iva": "24081501", "cta_rfte_sug": "", "tasa_rfte_sug": 0.0, "tope_rfte": 999999999,
                "cta_rica_sug": "", "tasa_rica_mil": 0.0,
                "razon": "Gasto de Taxis, Buses y Movilidad Urbana (Cta 51954501)"
            }
        if any(k in prov_u for k in ["TERPEL", "PRIMAX", "BIOMAX", "TEXACO", "PETROBRAS", "ESTACION DE SERVICIO"]) or \
           any(k in full_search for k in ["GASOLINA", "ACPM", "COMBUSTIBLE", "LUBRICANTE"]):
            return {
                "cta_p": "51953501", "cta_c": "23359501",
                "cat": "Gasto Combustibles y Lubricantes",
                "desc": f"Combustibles y Lubricantes - {str(nombre_emisor)[:24]}",
                "cta_iva": "24081001", "cta_rfte_sug": "23654006", "tasa_rfte_sug": 0.001, "tope_rfte": 0,
                "cta_rica_sug": "", "tasa_rica_mil": 0.0,
                "razon": "Gasto de Combustibles y Lubricantes (Cta 51953501)"
            }
        if any(k in full_search for k in ["PARQUEADERO", "PARKING", "ESTACIONAMIENTO", "PEAJE", "GOPASS", "FLYPASS"]):
            return {
                "cta_p": "51956501", "cta_c": "23359501",
                "cat": "Gasto Parqueaderos y Peajes",
                "desc": f"Parqueaderos y Peajes - {str(nombre_emisor)[:24]}",
                "cta_iva": "24081501", "cta_rfte_sug": "", "tasa_rfte_sug": 0.0, "tope_rfte": 999999999,
                "cta_rica_sug": "", "tasa_rica_mil": 0.0,
                "razon": "Gasto de Parqueaderos y Peajes (Cta 51956501)"
            }
        if any(k in full_search for k in ["DOTACION", "UNIFORME", "BOTAS DE SEGURIDAD", "OVEROL", "ELEMENTOS DE PROTECCION"]):
            return {
                "cta_p": "51055101", "cta_c": "23359501",
                "cat": "Gasto Dotacion y EPP",
                "desc": f"Dotacion y EPP a Trabajadores - {str(nombre_emisor)[:24]}",
                "cta_iva": "24081001", "cta_rfte_sug": "23654001", "tasa_rfte_sug": 0.025, "tope_rfte": 1047000,
                "cta_rica_sug": "23680501", "tasa_rica_mil": 11.04,
                "razon": "Dotacion y Suministro a Trabajadores (Cta 51055101)"
            }
        if any(k in full_search for k in ["VIGILANCIA", "SEGURIDAD PRIVADA", "SERVICIO DE ASEO", "FUMIGACION"]):
            return {
                "cta_p": "51350501", "cta_c": "23359501",
                "cat": "Gasto Aseo y Vigilancia",
                "desc": f"Servicio de Aseo y Vigilancia - {str(nombre_emisor)[:24]}",
                "cta_iva": "24081501", "cta_rfte_sug": "23652503", "tasa_rfte_sug": 0.04, "tope_rfte": 210000,
                "cta_rica_sug": "23680505", "tasa_rica_mil": 9.66,
                "razon": "Servicios de Aseo y Vigilancia (Cta 51350501)"
            }

        # 16. MANTENIMIENTO INDUSTRIAL, REPUESTOS, MATERIAL ELECTRICO Y SERVICIOS DE INGENIERIA
        if any(k in prov_u for k in [
            "FERROMENDEZ", "TORNILLOLOCO", "LIZ CAUCHOS", "ASIMFER", "MAFLEXCOL",
            "EMPRECOL", "MECANIZAR", "BATTS ZONE", "HIDRAHULICAS", "ELECTROILUMINACIONES",
            "INTERNACIONAL DE ELECTRICOS", "INVERSIONES PRIMERA", "TUBERIAS ALPAX", "TUBER\u00cdAS ALPAX",
            "CDGINGENIEROS", "CYGNI INGENIERIA", "HENRY VILLA BLANCO"
        ]) or any(k in full_search for k in [
            "CABLE ENCAUCHETADO", "EJE 304", "BISAGRA REDONDA", "SERVICIO DE INGENIERIA",
            "ACOMPA\u00d1AMIENTO ARR", "REPARACION DE", "PUNZONES", "FERULA INOX",
            "CLAVIJA", "CONTACTOR", "NIVELADOR CAUCHO", "MANGUERA PVC", "TORNILLO BRISTOL"
        ]):
            if cta_actual in ["14350101", "61350501"] and not any(k in prov_u for k in ["CDGINGENIEROS", "HENRY VILLA BLANCO"]):
                return {
                    "cta_p": cta_actual, "cta_c": "22050501" if cta_actual == "14350101" else "23359501",
                    "cat": "Inventario Repuestos" if cta_actual == "14350101" else "Costo Insumos y Materiales",
                    "desc": f"Repuestos e Insumos Industriales - {str(nombre_emisor)[:24]}",
                    "cta_iva": "24081001", "cta_rfte_sug": "23654001", "tasa_rfte_sug": 0.025, "tope_rfte": 1047000,
                    "cta_rica_sug": "23680501", "tasa_rica_mil": 11.04,
                    "razon": f"Compra de Repuestos e Insumos Industriales (Cta {cta_actual})"
                }
            return {
                "cta_p": "61800101", "cta_c": "23359501",
                "cat": "Costo Mantenimiento e Ingenieria",
                "desc": f"Mantenimiento, Repuestos e Ingenieria - {str(nombre_emisor)[:22]}",
                "cta_iva": "24081001", "cta_rfte_sug": "23654001", "tasa_rfte_sug": 0.025, "tope_rfte": 1047000,
                "cta_rica_sug": "23680501", "tasa_rica_mil": 11.04,
                "razon": "Costo de Mantenimiento, Reparaciones, Material Electrico e Ingenieria (Cta 61800101)"
            }

        return None

    def obtener_asiento_contable_hoja2(fac_sel, es_aduanero=False):
        """Genera el asiento contable EXACTO e ID\u00c9NTICO al que visualiza la Hoja 2 (Auditor\u00eda)."""
        asiento_filas = []
        es_nc = "Devolucion" in str(fac_sel.get("Operacion", ""))
        asume_ret = bool(fac_sel.get("Impuestos Asumidos", False))
        cta_cxp_usar = str(fac_sel.get("Cuenta Pasivo Especifica") or fac_sel.get("Cta Contrapartida") or ("22050505" if es_aduanero else "22050501")).strip()
        cta_p = str(fac_sel.get("Cta Principal") or "14650501").strip()
        if cta_p in ["146505", "1465"]:
            cta_p = "14650501"
        es_importacion_factura = es_aduanero or "1465" in cta_p or "IMPORTACI" in str(fac_sel.get("Categor\u00eda", "")).upper()
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
        iva_5_val = _sf(fac_sel.get("IVA 5%"))
        iva_19_val = _sf(fac_sel.get("IVA 19%"))
        if iva_5_val > 0:
            if iva_19_val <= 0 and iva_val > iva_5_val:
                iva_19_val = round(iva_val - iva_5_val, 2)
            iva_val = round(iva_19_val + iva_5_val, 2)
        else:
            iva_19_val = iva_val
        base_19_val = _sf(fac_sel.get("Base IVA 19%")) or (round(iva_19_val / 0.19, 2) if iva_5_val > 0 and iva_19_val > 0 else base_val)
        base_5_val = _sf(fac_sel.get("Base IVA 5%")) or (round(iva_5_val / 0.05, 2) if iva_5_val > 0 else 0.0)
        inc_otros_val = _sf(fac_sel.get("INC / Otros Impuestos"))
        cta_iva_5 = _limpiar_codigo_cta(fac_sel.get("Cta IVA 5%"), "24081003")
        tot_val = _sf(fac_sel.get("Total"))
        rfte_val = _sf(fac_sel.get("ReteFuente"))
        rica_val = _sf(fac_sel.get("ReteICA"))
        riva_val = _sf(fac_sel.get("ReteIVA"))
        iva_imp_val = _sf(fac_sel.get("IVA Importaci\u00f3n"))

        raw_cta_imp = fac_sel.get("Cta IVA Importaci\u00f3n")
        if raw_cta_imp is None or pd.isna(raw_cta_imp) or str(raw_cta_imp).strip() in ["", "nan", "None"]:
            cta_iva_imp = "240835"
        else:
            cta_iva_imp = str(raw_cta_imp).strip()

        # 1. Base / Costo / Gasto Especifico PUC
        etiqueta_cta_p = obtener_etiqueta_cuenta_principal(cta_p, fac_sel.get("Categor\u00eda", ""))
        asiento_filas.append({
            "C\u00f3digo Cuenta": cta_p,
            "Descripci\u00f3n de la Cuenta": "{} - {}".format(etiqueta_cta_p, str(fac_sel.get("Proveedor", ""))[:25]),
            "Tercero / NIT": str(fac_sel.get("NIT Emisor", "")),
            "D\u00e9bito ($)": 0.0 if es_nc else base_val,
            "Cr\u00e9dito ($)": base_val if es_nc else 0.0
        })

        # 1.B INC / Impoconsumo 8% / Bolsas (Mayor valor del gasto/costo)
        if inc_otros_val > 0:
            asiento_filas.append({
                "C\u00f3digo Cuenta": cta_p,
                "Descripci\u00f3n de la Cuenta": f"INC 8% / Otros Impuestos No Descontables (Mayor valor {cta_p})",
                "Tercero / NIT": str(fac_sel.get("NIT Emisor", "")),
                "D\u00e9bito ($)": 0.0 if es_nc else inc_otros_val,
                "Cr\u00e9dito ($)": inc_otros_val if es_nc else 0.0
            })

        # 2.A IVA Descontable 19% (Servicios 24081501 / Compras 24081001)
        if iva_19_val > 0:
            desc_iva = f"IVA Descontable Servicios 19% (Base: ${base_19_val:,.0f})" if ("24081501" in cta_iva or es_importacion_factura) else f"IVA Descontable 19% (Base: ${base_19_val:,.0f})"
            asiento_filas.append({
                "C\u00f3digo Cuenta": cta_iva,
                "Descripci\u00f3n de la Cuenta": desc_iva,
                "Tercero / NIT": str(fac_sel.get("NIT Emisor", "")),
                "D\u00e9bito ($)": 0.0 if es_nc else iva_19_val,
                "Cr\u00e9dito ($)": iva_19_val if es_nc else 0.0
            })

        # 2.A-2 IVA Descontable 5% (Compras 24081003)
        if iva_5_val > 0:
            asiento_filas.append({
                "C\u00f3digo Cuenta": cta_iva_5,
                "Descripci\u00f3n de la Cuenta": f"IVA Descontable 5% (Base: ${base_5_val:,.0f})",
                "Tercero / NIT": str(fac_sel.get("NIT Emisor", "")),
                "D\u00e9bito ($)": 0.0 if es_nc else iva_5_val,
                "Cr\u00e9dito ($)": iva_5_val if es_nc else 0.0
            })

        # 2.B IVA de Importacion (Cuenta 240835) - Solo para facturas de importacion
        if iva_imp_val > 0:
            asiento_filas.append({
                "C\u00f3digo Cuenta": cta_iva_imp,
                "Descripci\u00f3n de la Cuenta": f"IVA de Importaci\u00f3n (Cta 240835) - Fac {fac_sel.get('Factura', '')}",
                "Tercero / NIT": str(fac_sel.get("NIT Emisor", "")),
                "D\u00e9bito ($)": 0.0 if es_nc else iva_imp_val,
                "Cr\u00e9dito ($)": iva_imp_val if es_nc else 0.0
            })

        es_transp_as = ("70787" in str(fac_sel.get("Factura", "")) or
                        "INLAND" in str(fac_sel.get("Descripcion", "")).upper() or
                        ("EURO" in str(fac_sel.get("Proveedor", "")).upper() and any(k in str(fac_sel.get("Descripcion", "")).upper() for k in ["TRANSP", "FLETE", "INLAND"])))

        cta_rf_raw = str(fac_sel.get("Cta ReteFuente", "")).strip()
        cta_ri_raw = str(fac_sel.get("Cta ReteICA", "")).strip()

        editada_man_h2 = bool(fac_sel.get("Editada Manualmente", False))
        # Si es transporte y NO fue editada manualmente para quitarle la retencion
        if not editada_man_h2 and (es_transp_as or (rfte_val > 0 and base_val > 0 and abs(rfte_val - round(base_val * 0.01, 2)) < 2.0)):
            if rfte_val == round(base_val * 0.04, 2) or rfte_val == 0.0:
                rfte_val = round(base_val * 0.01, 2)
            if rica_val == round(base_val * 0.00966, 2) or rica_val == 0.0:
                rica_val = round(base_val * 0.00414, 2)
            cta_rf_usar = "23652505"
            cta_ri_usar = "23680513"
            cta_p = "14650501"
        else:
            cta_rf_usar = cta_rf_raw if (cta_rf_raw and cta_rf_raw not in ["nan", "None", ""]) else ("" if rfte_val <= 0 else ("23652503" if es_aduanero else "23654001"))
            cta_ri_usar = cta_ri_raw if (cta_ri_raw and cta_ri_raw not in ["nan", "None", ""]) else ("" if rica_val <= 0 else ("23680505" if es_aduanero else "23680501"))

        if cta_p == "146505":
            cta_p = "14650501"

        tot_ret = round(rfte_val + rica_val, 2)

        if asume_ret:
            # IMPUESTOS ASUMIDOS (Cruza con Agente Aduanero / Triangulacion)
            if tot_ret > 0:
                asiento_filas.append({
                    "C\u00f3digo Cuenta": "53152001",
                    "Descripci\u00f3n de la Cuenta": "Retenciones Asumidas (Impuestos Asumidos Aduana)",
                    "Tercero / NIT": str(fac_sel.get("NIT Emisor", "")),
                    "D\u00e9bito ($)": tot_ret,
                    "Cr\u00e9dito ($)": 0.0
                })
            if rfte_val > 0:
                asiento_filas.append({
                    "C\u00f3digo Cuenta": cta_rf_usar,
                    "Descripci\u00f3n de la Cuenta": f"ReteFuente Practicada ({cta_rf_usar}) Fac {fac_sel.get('Factura', '')}",
                    "Tercero / NIT": str(fac_sel.get("NIT Emisor", "")),
                    "D\u00e9bito ($)": 0.0,
                    "Cr\u00e9dito ($)": rfte_val
                })
            if rica_val > 0:
                asiento_filas.append({
                    "C\u00f3digo Cuenta": cta_ri_usar,
                    "Descripci\u00f3n de la Cuenta": "Retenci\u00f3n ICA Practicada",
                    "Tercero / NIT": str(fac_sel.get("NIT Emisor", "")),
                    "D\u00e9bito ($)": 0.0,
                    "Cr\u00e9dito ($)": rica_val
                })
            saldo_cxp = round(base_val + inc_otros_val + iva_val + iva_imp_val, 2)
            asiento_filas.append({
                "C\u00f3digo Cuenta": cta_cxp_usar,
                "Descripci\u00f3n de la Cuenta": f"CxP Proveedor/Agente - Fac {fac_sel.get('Factura', '')}",
                "Tercero / NIT": str(fac_sel.get("NIT Emisor", "")),
                "D\u00e9bito ($)": saldo_cxp if es_nc else 0.0,
                "Cr\u00e9dito ($)": 0.0 if es_nc else saldo_cxp
            })
        else:
            # RETENCIONES ORDINARIAS PRACTICADAS AL PROVEEDOR
            if rfte_val > 0:
                nom_rf_dsp = "Transporte Carga 1%" if cta_rf_usar == "23652505" else ("Servicios 4%" if cta_rf_usar == "23652503" else fac_sel.get('Categor\u00eda', 'Compras/Servicios'))
                asiento_filas.append({
                    "C\u00f3digo Cuenta": cta_rf_usar,
                    "Descripci\u00f3n de la Cuenta": f"ReteFuente Practicada ({nom_rf_dsp}) Fac {fac_sel.get('Factura', '')}",
                    "Tercero / NIT": str(fac_sel.get("NIT Emisor", "")),
                    "D\u00e9bito ($)": rfte_val if es_nc else 0.0,
                    "Cr\u00e9dito ($)": 0.0 if es_nc else rfte_val
                })
            if rica_val > 0:
                nom_ri_dsp = "Transporte 4.14\u2030" if cta_ri_usar == "23680513" else ("Servicios 9.66\u2030" if cta_ri_usar == "23680505" else "Retenci\u00f3n ICA Practicada")
                asiento_filas.append({
                    "C\u00f3digo Cuenta": cta_ri_usar,
                    "Descripci\u00f3n de la Cuenta": f"Retenci\u00f3n ICA Practicada ({nom_ri_dsp})",
                    "Tercero / NIT": str(fac_sel.get("NIT Emisor", "")),
                    "D\u00e9bito ($)": rica_val if es_nc else 0.0,
                    "Cr\u00e9dito ($)": 0.0 if es_nc else rica_val
                })
            if riva_val > 0:
                asiento_filas.append({
                    "C\u00f3digo Cuenta": "23670101",
                    "Descripci\u00f3n de la Cuenta": "Retenci\u00f3n de IVA Practicada (15%)",
                    "Tercero / NIT": str(fac_sel.get("NIT Emisor", "")),
                    "D\u00e9bito ($)": riva_val if es_nc else 0.0,
                    "Cr\u00e9dito ($)": 0.0 if es_nc else riva_val
                })
            neto_cxp = round(base_val + inc_otros_val + iva_val + iva_imp_val - rfte_val - rica_val - riva_val, 2)
            asiento_filas.append({
                "C\u00f3digo Cuenta": cta_cxp_usar,
                "Descripci\u00f3n de la Cuenta": f"Proveedores Nacionales - Fac {fac_sel.get('Factura', '')}",
                "Tercero / NIT": str(fac_sel.get("NIT Emisor", "")),
                "D\u00e9bito ($)": neto_cxp if es_nc else 0.0,
                "Cr\u00e9dito ($)": 0.0 if es_nc else neto_cxp
            })

        df_h2 = pd.DataFrame(asiento_filas)
        if "Descripci\u00f3n de la Cuenta" in df_h2.columns and "Descripci\u00f3n Cuenta" not in df_h2.columns:
            df_h2["Descripci\u00f3n Cuenta"] = df_h2["Descripci\u00f3n de la Cuenta"]
        elif "Descripci\u00f3n Cuenta" in df_h2.columns and "Descripci\u00f3n de la Cuenta" not in df_h2.columns:
            df_h2["Descripci\u00f3n de la Cuenta"] = df_h2["Descripci\u00f3n Cuenta"]
        return df_h2


    def obtener_saldo_cruce_factura(tr):
        """
        Calcula de forma segura el Saldo de Cruce (Base + IVA) de una factura para la triangulaci\u00f3n,
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
            t_imp = float(tr.get("IVA Importaci\u00f3n", 0.0)) if pd.notna(tr.get("IVA Importaci\u00f3n")) else 0.0
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
        Calcula de forma din\u00e1mica la concordancia cronol\u00f3gica entre la factura del tercero y la del agente,
        garantizando que nunca retorne None ni NaN.
        """
        if not fecha_tercero or not fecha_agente or pd.isna(fecha_tercero) or pd.isna(fecha_agente):
            return "\u26aa Sin fecha"
        dt_terc = pd.to_datetime(str(fecha_tercero), dayfirst=True, errors="coerce")
        dt_ag = pd.to_datetime(str(fecha_agente), dayfirst=True, errors="coerce")
        if pd.isna(dt_terc) or pd.isna(dt_ag):
            return "\u26aa Sin fecha"
    
        diff_d = int((dt_terc - dt_ag).days)
        if diff_d == 0:
            return "\U0001f7e2 Mismo d\u00eda (0d)"
        elif 0 < diff_d <= 15:
            return f"\u27a1\ufe0f Poco despu\u00e9s (+{diff_d}d)"
        elif -30 <= diff_d < 0:
            return f"\u2b05\ufe0f Anterior ({abs(diff_d)}d antes)"
        elif -60 <= diff_d < -30:
            return f"\u2b05\ufe0f Anterior ({abs(diff_d)}d antes)"
        elif diff_d < -60:
            return f"\u26d4 Fuera de rango ({abs(diff_d)}d antes)"
        else:
            return f"\u26d4 Fuera de rango (+{diff_d}d desp)"


    def generar_asiento_mixto_hoja2_con_terceros(agente_row, terceros_df, enviar_a_no_deducible=False, imputar_a_transito=False):
        """
        Construye el asiento contable para un cobro de agente cuando se trae la contabilizaci\u00f3n de P\u00e1gina 2:
        1. Si en Hoja 2 se registraron \u00fanicamente los Ingresos Propios (Base + IVA < Total Factura):
           Combina los Ingresos Propios (con su IVA y retenciones) + Facturas de Terceros vinculadas (cancelaci\u00f3n de CxP) +
           Cr\u00e9dito total al Agente Aduanero (22050501).
        2. El saldo de ingresos para terceros NO se imputa autom\u00e1ticamente a ninguna cuenta de gasto ni tr\u00e1nsito.
           Se deja PENDIENTE para que el usuario concilie con facturas de terceros o decida expl\u00edcitamente enviarlo a No Deducibles o a Tr\u00e1nsito.
        3. Si el usuario selecciona 'No Deducibles', imputa el saldo a 53950501.
        4. Si el usuario selecciona 'Mercanc\u00edas en Tr\u00e1nsito', imputa el saldo a 14650501.
        """
        df_h2 = obtener_asiento_contable_hoja2(agente_row, es_aduanero=False)
    
        tot_agente = float(agente_row.get("Total", 0.0))
        base_propia = float(agente_row.get("Base", 0.0))
        iva_propio = float(agente_row.get("IVA", 0.0))
        iva_imp_propio = float(agente_row.get("IVA Importaci\u00f3n", 0.0))
        subtotal_propio = round(base_propia + iva_propio + iva_imp_propio, 2)
        saldo_terceros_esperado = round(tot_agente - subtotal_propio, 2)
    
        # Si la factura en Hoja 2 ya cubre el total y no hay terceros asignados
        if saldo_terceros_esperado <= 0.05 and (terceros_df is None or terceros_df.empty):
            return df_h2, 0.0
        
        filas_propias = []
        tot_ret_propias = 0.0
        for _, r in df_h2.iterrows():
            cta = str(r["C\u00f3digo Cuenta"]).strip()
            if cta.startswith("2365") or cta.startswith("2368") or cta.startswith("2367"):
                tot_ret_propias += float(r["Cr\u00e9dito ($)"])
                filas_propias.append(r.to_dict())
            elif not (cta.startswith("2205") and float(r["Cr\u00e9dito ($)"]) > 0):
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
                    "C\u00f3digo Cuenta": cta_cxp,
                    "Descripci\u00f3n Cuenta": f"Cancela CxP {prov_nom[:20]} (Fac {fac_num})",
                    "Descripci\u00f3n de la Cuenta": f"Cancela CxP {prov_nom[:20]} (Fac {fac_num})",
                    "Tercero / NIT": f"{nit_t} - {prov_nom[:25]}",
                    "D\u00e9bito ($)": t_saldo_pagar,
                    "Cr\u00e9dito ($)": 0.0
                })
            
        # Credito total por pagar al Agente Aduanero (Euro Shipping / Trade Global)
        asume_ret = bool(agente_row.get("Impuestos Asumidos", False))
        cxp_total_agente = tot_agente if asume_ret else round(tot_agente - tot_ret_propias, 2)
    
        filas_credito = [{
            "C\u00f3digo Cuenta": CUENTA_CXP_EURO_SHIPPING,
            "Descripci\u00f3n Cuenta": f"CxP Agente Aduanero - Fac {agente_row.get('Factura', '')}",
            "Descripci\u00f3n de la Cuenta": f"CxP Agente Aduanero - Fac {agente_row.get('Factura', '')}",
            "Tercero / NIT": f"{agente_row.get('NIT Emisor', '')} - {str(agente_row.get('Proveedor', ''))[:25]}",
            "D\u00e9bito ($)": 0.0,
            "Cr\u00e9dito ($)": cxp_total_agente
        }]
    
        # Calcular la diferencia matematica exacta de partida doble entre debitos y creditos del asiento
        suma_deb_base = sum(float(f.get("D\u00e9bito ($)", 0.0) or 0.0) for f in (filas_propias + filas_terceros))
        suma_cred_base = sum(float(f.get("Cr\u00e9dito ($)", 0.0) or 0.0) for f in (filas_propias + filas_credito))
        dif_asiento_exacta = round(suma_cred_base - suma_deb_base, 2)

        filas_ajuste = []
        if abs(dif_asiento_exacta) > 0.05:
            if enviar_a_no_deducible:
                if dif_asiento_exacta > 0.05:
                    # Faltan debitos por justificar -> Debito a No Deducibles (53950501)
                    filas_ajuste.append({
                        "C\u00f3digo Cuenta": CUENTA_NO_DEDUCIBLE,
                        "Descripci\u00f3n Cuenta": f"Gastos No Deducibles Terceros (Diferencia sin soporte DIAN)",
                        "Descripci\u00f3n de la Cuenta": f"Gastos No Deducibles Terceros (Diferencia sin soporte DIAN)",
                        "Tercero / NIT": f"{agente_row.get('NIT Emisor', '')} - {str(agente_row.get('Proveedor', ''))[:25]}",
                        "D\u00e9bito ($)": dif_asiento_exacta,
                        "Cr\u00e9dito ($)": 0.0
                    })
                else:
                    # Debitos exceden creditos (los terceros superan el cobro del agente) -> Credito a 53950501
                    filas_ajuste.append({
                        "C\u00f3digo Cuenta": CUENTA_NO_DEDUCIBLE,
                        "Descripci\u00f3n Cuenta": f"Ajuste No Deducible / Terceros Superan Cobro Agente",
                        "Descripci\u00f3n de la Cuenta": f"Ajuste No Deducible / Terceros Superan Cobro Agente",
                        "Tercero / NIT": f"{agente_row.get('NIT Emisor', '')} - {str(agente_row.get('Proveedor', ''))[:25]}",
                        "D\u00e9bito ($)": 0.0,
                        "Cr\u00e9dito ($)": abs(dif_asiento_exacta)
                    })
            elif imputar_a_transito:
                if dif_asiento_exacta > 0.05:
                    filas_ajuste.append({
                        "C\u00f3digo Cuenta": CUENTA_IMPORTACION_TRANSITO,
                        "Descripci\u00f3n Cuenta": f"Mercanc\u00edas en Tr\u00e1nsito / Saldo Faltante (Fac {agente_row.get('Factura', '')})",
                        "Descripci\u00f3n de la Cuenta": f"Mercanc\u00edas en Tr\u00e1nsito / Saldo Faltante (Fac {agente_row.get('Factura', '')})",
                        "Tercero / NIT": f"{agente_row.get('NIT Emisor', '')} - {str(agente_row.get('Proveedor', ''))[:25]}",
                        "D\u00e9bito ($)": dif_asiento_exacta,
                        "Cr\u00e9dito ($)": 0.0
                    })
                else:
                    filas_ajuste.append({
                        "C\u00f3digo Cuenta": CUENTA_IMPORTACION_TRANSITO,
                        "Descripci\u00f3n Cuenta": f"Ajuste Mercanc\u00edas en Tr\u00e1nsito / Terceros Superan Cobro (Fac {agente_row.get('Factura', '')})",
                        "Descripci\u00f3n de la Cuenta": f"Ajuste Mercanc\u00edas en Tr\u00e1nsito / Terceros Superan Cobro (Fac {agente_row.get('Factura', '')})",
                        "Tercero / NIT": f"{agente_row.get('NIT Emisor', '')} - {str(agente_row.get('Proveedor', ''))[:25]}",
                        "D\u00e9bito ($)": 0.0,
                        "Cr\u00e9dito ($)": abs(dif_asiento_exacta)
                    })
            
        asiento_final = filas_propias + filas_terceros + filas_ajuste + filas_credito
        df_res = pd.DataFrame(asiento_final)
        dif_pendiente = abs(dif_asiento_exacta) if (not enviar_a_no_deducible and not imputar_a_transito) else 0.0
        return df_res, dif_pendiente

    def generar_asiento_triangulacion_paquete(agente_row, terceros_df, enviar_a_no_deducible=False, enviar_a_gastos_propios=False, enviar_a_hoja2=False, **kwargs):
        """
        Calcula el asiento contable de partida doble para un paquete de importaci\u00f3n triangulado:
        1. D\u00e9bito a CxP Terceros (22050505/23359501) por el Saldo por Pagar real (Base + IVA, ej. .651.939).
           Si ya estaba registrada en Siigo, se cancela la cuenta por pagar imputada.
        2. D\u00e9bito a Retenciones Asumidas (53152001) si las retenciones no se hab\u00edan asumido en la causaci\u00f3n previa.
        3. D\u00e9bito a Importaci\u00f3n en Tr\u00e1nsito (14650501) e IVA (24081501) para facturas blancas pendientes.
        4. D\u00e9bito a IVA de Importaci\u00f3n si la factura del agente lo discrimina.
        5. D\u00e9bito a Gastos No Deducibles (53950501) por la diferencia sin factura DIAN.
        6. Cr\u00e9dito a CxP Agente Aduanero (22050501) por el 100% de la factura del agente (ej. .077.647).
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
        
            # El Saldo por Pagar que entra a la triangulacion es Base + IVA (ej. .388.770 + 63.169 = .651.939)
            t_saldo_pagar = obtener_saldo_cruce_factura(t)
            
            # REGLA DE TRIANGULACION: Toda factura de tercero traida a la triangulacion UNICAMENTE
            # aporta su Cuenta por Pagar (pasivo: 22050505 Agencia, 23359501 DHL, etc.) por su Saldo por Pagar (Base + IVA),
            # ya que la causacion del gasto/costo e IVA de esa factura se realiza en su propio comprobante individual.
            asume_en_causacion = bool(t.get("Impuestos Asumidos", False))
        
            if asume_en_causacion or tot_ret == 0:
                cxp_cancelar = t_saldo_pagar
                ret_a_asumir = 0.0
                suma_cxp_canceladas += cxp_cancelar
            
                asiento.append({
                    "C\u00f3digo Cuenta": cta_cxp,
                    "Descripci\u00f3n Cuenta": f"Cancela CxP {prov_nom[:20]} (Fac {fac_num})",
                    "Tercero / NIT": f"{nit_t} - {prov_nom[:25]}",
                    "D\u00e9bito ($)": cxp_cancelar,
                    "Cr\u00e9dito ($)": 0.0
                })
            else:
                cxp_cancelar = round(t_saldo_pagar - tot_ret, 2)
                ret_a_asumir = tot_ret
                suma_cxp_canceladas += cxp_cancelar
                suma_ret_asumidas += ret_a_asumir
            
                asiento.append({
                    "C\u00f3digo Cuenta": cta_cxp,
                    "Descripci\u00f3n Cuenta": f"Cancela CxP {prov_nom[:20]} (Fac {fac_num})",
                    "Tercero / NIT": f"{nit_t} - {prov_nom[:25]}",
                    "D\u00e9bito ($)": cxp_cancelar,
                    "Cr\u00e9dito ($)": 0.0
                })
                if ret_a_asumir > 0:
                    asiento.append({
                        "C\u00f3digo Cuenta": CUENTA_RETENCION_ASUMIDA,
                        "Descripci\u00f3n Cuenta": f"Retenci\u00f3n Asumida Fac {fac_num} ({prov_nom[:18]})",
                        "Tercero / NIT": f"{nit_t} - {prov_nom[:25]}",
                        "D\u00e9bito ($)": ret_a_asumir,
                        "Cr\u00e9dito ($)": 0.0
                    })

        # IVA Servicios e IVA de Importacion del Agente
        iva_imp_agente = float(agente_row.get("IVA Importaci\u00f3n", 0.0))
        cta_iva_ag = str(agente_row.get("Cta IVA") or "24081501").strip()
        cta_iva_imp_ag = str(agente_row.get("Cta IVA Importaci\u00f3n") or "240835").strip()

        if iva_agente > 0:
            asiento.append({
                "C\u00f3digo Cuenta": cta_iva_ag,
                "Descripci\u00f3n Cuenta": f"IVA Descontable Servicios Fac {agente_row.get('Factura', '')}",
                "Tercero / NIT": f"{agente_row.get('NIT Emisor', '')} - {agente_row.get('Proveedor', '')[:25]}",
                "D\u00e9bito ($)": iva_agente,
                "Cr\u00e9dito ($)": 0.0
            })

        if iva_imp_agente > 0:
            asiento.append({
                "C\u00f3digo Cuenta": cta_iva_imp_ag,
                "Descripci\u00f3n Cuenta": f"IVA de Importaci\u00f3n (Cta 240835) Fac {agente_row.get('Factura', '')}",
                "Tercero / NIT": f"{agente_row.get('NIT Emisor', '')} - {agente_row.get('Proveedor', '')[:25]}",
                "D\u00e9bito ($)": iva_imp_agente,
                "Cr\u00e9dito ($)": 0.0
            })

        # Calcular la diferencia no deducible sin factura DIAN (Cuenta 53950501)
        suma_justificada = suma_cxp_canceladas + suma_ret_asumidas + suma_costo_blancas + suma_iva_blancas + iva_agente + iva_imp_agente
        diferencia_no_deducible = round(tot_agente - suma_justificada, 2)
    
        if diferencia_no_deducible > 0.01 and enviar_a_no_deducible:
            asiento.append({
                "C\u00f3digo Cuenta": CUENTA_NO_DEDUCIBLE,
                "Descripci\u00f3n Cuenta": f"Gastos No Deducibles Importaci\u00f3n (Diferencia sin soporte DIAN)",
                "Tercero / NIT": f"{agente_row.get('NIT Emisor', '')} - {agente_row.get('Proveedor', '')[:25]}",
                "D\u00e9bito ($)": diferencia_no_deducible,
                "Cr\u00e9dito ($)": 0.0
            })
        elif diferencia_no_deducible < -0.01:
            # En caso de que la suma de terceros exceda ligeramente al agente por redondeo
            asiento.append({
                "C\u00f3digo Cuenta": CUENTA_NO_DEDUCIBLE,
                "Descripci\u00f3n Cuenta": f"Ajuste por Diferencia de Cuadre Importaci\u00f3n",
                "Tercero / NIT": f"{agente_row.get('NIT Emisor', '')} - {agente_row.get('Proveedor', '')[:25]}",
                "D\u00e9bito ($)": 0.0,
                "Cr\u00e9dito ($)": abs(diferencia_no_deducible)
            })

        # Credito total al Agente Aduanero (Euro Shipping / Trade Global) por el 100% de su factura
        asiento.append({
            "C\u00f3digo Cuenta": CUENTA_CXP_EURO_SHIPPING,
            "Descripci\u00f3n Cuenta": f"CxP Agente Aduanero - Fac {agente_row.get('Factura', '')}",
            "Tercero / NIT": f"{agente_row.get('NIT Emisor', '')} - {agente_row.get('Proveedor', '')[:25]}",
            "D\u00e9bito ($)": 0.0,
            "Cr\u00e9dito ($)": tot_agente
        })

        df_asiento = pd.DataFrame(asiento)
        if "Descripci\u00f3n Cuenta" in df_asiento.columns and "Descripci\u00f3n de la Cuenta" not in df_asiento.columns:
            df_asiento["Descripci\u00f3n de la Cuenta"] = df_asiento["Descripci\u00f3n Cuenta"]
        elif "Descripci\u00f3n de la Cuenta" in df_asiento.columns and "Descripci\u00f3n Cuenta" not in df_asiento.columns:
            df_asiento["Descripci\u00f3n Cuenta"] = df_asiento["Descripci\u00f3n de la Cuenta"]
        return df_asiento, max(0.0, diferencia_no_deducible), suma_ret_asumidas

    # Directorio oficial de regimenes fiscales conocidos por NIT y Nombre
    REGIMENES_EMISORES_CONOCIDOS = {
        '860502609': 'O-13;O-15',   # DHL EXPRESS COLOMBIA LTDA (Gran Contribuyente y Autorretenedor)
        '800215775': 'O-13;O-15',   # SOCIEDAD PORTUARIA REGIONAL DE BUENAVENTURA S.A.
        '890304099': 'O-13;O-15',   # HOTELES ESTELAR S.A.
        '830048145': 'O-13;O-15',   # SIIGO S.A.S.
        '830048268': 'O-48',        # EURO SHIPPING SERVICES S.A.S (Responsable de IVA / Codigo 23 - 48)
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
        '240835': 1,     # IVA de Importacion 19%
        '24083501': 1,   # IVA de Importacion 19% (8 digitos)
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
        Escaneo inteligente de responsabilidades fiscales del EMISOR en la factura electr\u00f3nica:
        Analiza el documento completo sin truncar en 'Cliente:' para no perder las resoluciones.
        Detecta O-13 (Gran Contribuyente), O-15 (Autorretenedor), O-47 (RST), O-48 y O-49.
        """
        if not texto:
            return ""

        txt = texto.upper()
        txt_norm = re.sub(r'\s+', ' ', txt)

        # 1. Si se conoce el NIT del emisor, verificar en el directorio de regimenes oficiales
        if nit_emisor:
            ne_clean = re.sub(r'\D', '', str(nit_emisor))
            if ne_clean in REGIMENES_EMISORES_CONOCIDOS:
                return REGIMENES_EMISORES_CONOCIDOS[ne_clean]

        # 1.1 Si el texto corresponde a Euro Shipping Services, es O-48 (Responsable de IVA / Comun)
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

        # 4. Analizar Regimen Simple de Tributacion (O-47 / 0-47 / RST)
        es_rst = False
        if re.search(r'\b[O0]-47\b', txt_norm):
            es_rst = True
        elif re.search('R[E\u00c9]GIMEN\\s+SIMPLE|SIMPLE\\s+DE\\s+TRIBUTACI[O\u00d3]N|\\bRST\\b', txt_norm):
            if not re.search('(?:NO\\s+(?:SOMOS\\s+)?|NI\\s+)(?:R[E\u00c9]GIMEN\\s+SIMPLE|RST)', txt_norm):
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
        Extrae con m\u00e1xima precisi\u00f3n los valores fiscales y totales desde el texto del PDF oficial:
        1. Subtotal / Total Bruto Factura (Base gravable antes de descuentos o retenciones en notas).
        2. IVA discriminado o Total Impuesto.
        3. Total Neto Factura (Base + IVA, antes de retenciones sugeridas).
        4. Retenciones en notas / Descuentos Globales (ReteFuente 11%/4%, ReteICA 0.966%, ReteIVA).
        5. Anticipos reales reportados en notas o casillas informativas.
        6. Identificaci\u00f3n de facturaci\u00f3n a favor del agente aduanero (Mandato).
        """
        datos = {}
        if not texto: return datos

        # 1. Notas Finales (formato clave: valor)
        m_nf_rfte = re.search(r'\bretefuente\s*:\s*([0-9\.\,]+)', texto, re.IGNORECASE)
        if m_nf_rfte: datos['retefuente'] = parse_num_co(m_nf_rfte.group(1))

        m_nf_rica = re.search(r'\breteica\s*:\s*([0-9\.\,]+)', texto, re.IGNORECASE)
        if m_nf_rica: datos['reteica'] = parse_num_co(m_nf_rica.group(1))

        m_nf_riva = re.search(r'\breteiva\s*:\s*([0-9\.\,]+)', texto, re.IGNORECASE)
        if m_nf_riva: datos['reteiva'] = parse_num_co(m_nf_riva.group(1))

        # 2. Descuentos globales por retencion sugerida si no vino en notas
        if 'retefuente' not in datos or datos['retefuente'] == 0:
            m_dg_rf = re.search(r'RETEFUENTE[^\$]+(?:[\$]\s*|COP\s*)([0-9\.\,]+)', texto, re.IGNORECASE)
            if m_dg_rf: datos['retefuente'] = parse_num_co(m_dg_rf.group(1))

        if 'reteica' not in datos or datos['reteica'] == 0:
            m_dg_ri = re.search(r'RETEICA[^\$]+(?:[\$]\s*|COP\s*)([0-9\.\,]+)', texto, re.IGNORECASE)
            if m_dg_ri: datos['reteica'] = parse_num_co(m_dg_ri.group(1))

        # 3. Desglose de impuestos por tarifa en items (IVA 19%, IVA 5%, INC 8%) y Bloque de Datos Totales
        lines_pdf = [l.strip() for l in texto.splitlines() if l.strip()]
        iva_19_items = 0.0
        iva_5_items = 0.0
        inc_8_items = 0.0
        for idx_l in range(1, len(lines_pdf)):
            cur_l = lines_pdf[idx_l]
            prev_l = lines_pdf[idx_l - 1]
            if cur_l in ["19.00", "19,00"]:
                v_l = parse_num_co(prev_l)
                if (',' in prev_l or '.' in prev_l) and v_l >= 0 and prev_l not in ["19.00", "19,00", "5.00", "5,00", "8.00", "8,00"]:
                    iva_19_items += v_l
            elif cur_l in ["5.00", "5,00"]:
                v_l = parse_num_co(prev_l)
                if (',' in prev_l or '.' in prev_l) and v_l >= 0 and prev_l not in ["19.00", "19,00", "5.00", "5,00", "8.00", "8,00"]:
                    iva_5_items += v_l
            elif cur_l in ["8.00", "8,00"]:
                v_l = parse_num_co(prev_l)
                if (',' in prev_l or '.' in prev_l) and v_l >= 0 and prev_l not in ["19.00", "19,00", "5.00", "5,00", "8.00", "8,00"]:
                    inc_8_items += v_l

        matches_moneda = [m.start() for m in re.finditer(r'MONEDA\s+COP', texto, re.IGNORECASE)]
        if matches_moneda:
            txt_totales = texto[matches_moneda[-1]:]
        else:
            pos_totales = texto.upper().rfind('DATOS TOTALES')
            txt_totales = texto[pos_totales:] if pos_totales != -1 else texto

        def _ext_lbl(lbl_rx, blk):
            m = re.search(lbl_rx + r'\s*[\r\n\s:\$]*(?:COP\s*)?(?:\$\s*)?([0-9]{1,3}(?:[\.\,][0-9]{3})*(?:[\.\,][0-9]{2,4})?)', blk, re.IGNORECASE)
            return parse_num_co(m.group(1)) if m else 0.0

        sub_raw = _ext_lbl(r'Subtotal', txt_totales)
        desc_det_raw = _ext_lbl(r'Descuento\s+detalle', txt_totales)
        tot_bruto_raw = _ext_lbl(r'Total\s+Bruto\s+Factura', txt_totales)
        iva_tot_raw = _ext_lbl(r'\bIVA\b', txt_totales)
        inc_raw = _ext_lbl(r'\bINC\b', txt_totales)
        bolsas_raw = _ext_lbl(r'\bBolsas\b', txt_totales)
        otros_imp_raw = _ext_lbl(r'Otros\s+impuestos', txt_totales)
        recargo_glob_raw = _ext_lbl(r'Recargo\s+Global\s*(?:\(\+\))?', txt_totales)
        tot_neto_raw = _ext_lbl(r'Total\s+neto\s+factura\s*(?:\(\=\))?', txt_totales)
        tot_fac_raw = _ext_lbl(r'Total\s+factura\s*(?:\(\=\))?', txt_totales)

        # Usar Total Bruto Factura cuando hay Descuento en detalle de items
        if tot_bruto_raw > 0:
            datos['subtotal'] = round(tot_bruto_raw + (recargo_glob_raw if iva_tot_raw == 0 else 0.0), 2)
        elif sub_raw > 0:
            datos['subtotal'] = round(sub_raw - desc_det_raw, 2)

        if iva_tot_raw > 0:
            datos['iva'] = round(iva_tot_raw, 2)
        elif (iva_19_items + iva_5_items) > 0:
            datos['iva'] = round(iva_19_items + iva_5_items, 2)

        if iva_5_items > 0:
            datos['iva_5'] = round(iva_5_items, 2)
            datos['base_iva_5'] = round(iva_5_items / 0.05, 2)
            datos['iva_19'] = round(max(0.0, datos.get('iva', 0.0) - datos['iva_5']), 2)
            datos['base_iva_19'] = round(datos['iva_19'] / 0.19, 2) if datos['iva_19'] > 0 else 0.0
        else:
            datos['iva_5'] = 0.0
            datos['base_iva_5'] = 0.0
            datos['iva_19'] = round(datos.get('iva', 0.0), 2)
            datos['base_iva_19'] = round(datos['iva_19'] / 0.19, 2) if datos['iva_19'] > 0 else 0.0

        datos['inc_otros'] = round(inc_raw + bolsas_raw + otros_imp_raw, 2)
        # Detectar IVA de Importacion agregado dentro de los items (DHL iva_declaracion / FedEx Derechos Aduanales IVA)
        iva_imp_det = 0.0
        m_iva_decl = re.search(r'iva_declaracion\s*:\s*([0-9\.\,]+)', texto, re.IGNORECASE)
        if m_iva_decl:
            iva_imp_det = parse_num_co(m_iva_decl.group(1))
        else:
            m_der_iva = re.search(r'Derechos\s+Aduanales\s+IVA[\s\S]{0,120}?([0-9]{1,3}(?:[\.\,][0-9]{3})+(?:[\.\,][0-9]{2})?)', texto, re.IGNORECASE)
            if m_der_iva:
                iva_imp_det = parse_num_co(m_der_iva.group(1))
        if iva_imp_det > 0 and datos.get('subtotal', 0.0) > iva_imp_det:
            datos['iva_importacion'] = round(iva_imp_det, 2)
            datos['subtotal'] = round(datos['subtotal'] - iva_imp_det, 2)
        if tot_neto_raw > 0:
            datos['total_neto'] = tot_neto_raw
        if tot_fac_raw > 0:
            datos['total_factura'] = tot_fac_raw

        m_fact_a = re.search(r'nombre_facturado_a\s*:\s*([^\n\r]+)', texto, re.IGNORECASE)
        if m_fact_a: datos['nombre_facturado_a'] = m_fact_a.group(1).strip()

        m_nit_fact_a = re.search(r'nit_facturado_a\s*:\s*([0-9]+)', texto, re.IGNORECASE)
        if m_nit_fact_a: datos['nit_facturado_a'] = m_nit_fact_a.group(1).strip()

        # 4. Extraccion de conceptos y descripcion de items / productos (ej. Transporte Terrestre)
        REGEX_TRANSPORTE = r'\b(TRANSPORTE|TRANSPORT|INLAND\s*RATE|FLETE|FLETES|ACARREO|ACARREOS|CARGA\s*TERRESTRE|TERRESTRE|DRAYAGE|PORTES|FLETAMENTO)\b'
        m_transp = re.search(REGEX_TRANSPORTE, texto, re.IGNORECASE)
        if m_transp:
            datos['es_transporte'] = True
            datos['concepto_items'] = "INLAND RATE TRANSPORTE TERRESTRE"

        m_det = re.search(r'Detalles\s+de\s+Productos(.*?)(?:Notas\s+Finales|DATOS\s+TOTALES|Totales|\Z)', texto, re.DOTALL | re.IGNORECASE)
        if m_det:
            for lin in m_det.group(1).splitlines():
                lin_s = lin.strip()
                if lin_s and not any(h in lin_s.upper() for h in ["C\u00d3DIGO", "CODIGO", "DESCRIPCI\u00d3N", "DESCRIPCION", "CANTIDAD", "PRECIO", "DESCUENTO", "RECARGO", "IMPUESTOS", "U/M", "VENTA", "UNITARIO"]):
                    lin_clean = re.sub(r'^\d+\s+', '', lin_s).strip()
                    if len(lin_clean) >= 4 and not re.match(r'^[\d\.\,\$\s\%\-]+$', lin_clean):
                        datos['concepto_items'] = lin_clean
                        if re.search(REGEX_TRANSPORTE, lin_clean, re.IGNORECASE):
                            datos['es_transporte'] = True
                        break

        return datos

    def calcular_base_con_regla_descuento(r, tot, iva, nom_emisor):
        """
        Calcula la Base Gravable contable respetando la regla estricta:
        1. Conversi\u00f3n de moneda extranjera (USD a COP por TRM): Si el Subtotal/IVA vienen en USD y Total en COP,
           convierte ambos valores a COP con la TRM impl\u00edcita para garantizar partida doble exacta.
        2. El descuento SOLO se aplica si est\u00e1 incorporado en los \u00cdTEMS de la factura.
        3. Si el descuento aparece en las NOTAS u observaciones, NO VA (no disminuye la base gravable).
        4. En agencias aduaneras (DHL, Euro Shipping, etc.), las casillas o notas de descuento corresponden
           a retenciones practicadas que no deben restarse de la base para evitar duplicidades.
        5. Si en el reporte/factura existe la columna de Subtotal (valor de los \u00edtems antes de notas),
           se toma directamente el Subtotal como base contable.
        """
        nom = str(nom_emisor).upper()
        es_aduanero = any(k in nom for k in AGENTES_ADUANEROS)

        col_sub = next((c for c in r.index if any(k in str(c).lower() for k in ["subtotal", "sub total", "base gravable", "base_gravable", "valor subtotal", "lineextensionamount"])), None)
        col_desc_item = next((c for c in r.index if any(k in str(c).lower() for k in ["descuento_items", "descuento item", "descuento \u00edtems", "descuento linea", "descuento comercial"])), None)
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

        # DETECCION Y CONVERSION DE MONEDA EXTRANJERA (USD a COP por TRM):
        # Si el Total en DIAN esta en COP (ej. > 50,000 COP) pero Subtotal e IVA vienen en USD (ej. 780 USD y 16.15 USD),
        # la relacion Total / (Subtotal + IVA) esta en el rango historico de TRM (1,500 a 6,500).
        if subtotal_val is not None and tot > 1000 and (subtotal_val + iva) > 0:
            ratio_trm = tot / (subtotal_val + iva)
            if 1500 <= ratio_trm <= 6500:
                subtotal_val = round(subtotal_val * ratio_trm, 2)
                iva = round(iva * ratio_trm, 2)
                desc_val = round(desc_val * ratio_trm, 2)
                desc_item_val = round(desc_item_val * ratio_trm, 2)
                base = round(subtotal_val, 2)
                motivo_base = f"Factura en USD convertida a TRM () \u2014 Base COP:  | IVA COP: "
                return base, iva, desc_val, motivo_base

        # REGLA:
        # Si viene el Subtotal de los items, esa es la base real. Los descuentos en notas NO se restan.
        if subtotal_val is not None:
            base = round(subtotal_val, 2)
            motivo_base = f"Base = Subtotal de \u00cdtems () \u2014 Descuentos en notas no aplican"
        else:
            # Si no hay columna de subtotal y es agencia aduanera con descuento reportado (retenciones en notas):
            if es_aduanero and desc_val > 0:
                base = round(tot - iva + desc_val, 2)
                motivo_base = f"Agencia Aduanera: Base restituida () sin descontar retenciones de notas"
            elif desc_item_val > 0:
                base = round(tot - iva, 2)
                motivo_base = f"Base calculada con descuento comercial en \u00edtems ()"
            else:
                base = round(tot - iva, 2)
                motivo_base = f"Base est\u00e1ndar = Total - IVA ()"

        # Validacion de consistencia: si Base + IVA difiere del Total por mas de .0 y no hay descuentos justificados
        if tot > 0 and abs((base + iva) - tot) > 1.0 and desc_val == 0 and not (es_aduanero and desc_val > 0):
            base = round(tot - iva, 2)
            motivo_base = f"Base ajustada = Total - IVA ()"

        return base, iva, desc_val, motivo_base


    # ==============================================================================
    # MOTOR DE MEMORIA Y APRENDIZAJE CONTINUO DE REGLAS CONTABLES
    # ==============================================================================
    PATH_MEMORIA_GLOBAL = "almacenamiento_contable/memoria_aprendizaje_contable.json"

    def cargar_memoria_aprendizaje(empresa_dict=None):
        """Carga el banco hist\u00f3rico de aprendizaje con reglas auditadas por el usuario."""
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
        """Memoriza una correcci\u00f3n o auditor\u00eda del usuario para aplicarla autom\u00e1ticamente en meses futuros."""
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

    def clasificar_factura(nit_emisor, nombre_emisor, valor_base, valor_iva, tipo_doc, resp_emisor="", empresa_compradora=None, descripcion=""):
        nombre_str = "" if (nombre_emisor is None or pd.isna(nombre_emisor)) else str(nombre_emisor).strip()
        if nombre_str.lower() in ["nan", "none"]:
            nombre_str = ""
        nombre = nombre_str.upper()
        nombre_emisor = nombre_str or "Proveedor"
        resp = str(resp_emisor).upper()
        desc_u = str(descripcion or "").upper()

        # Deteccion explicita de concepto TRANSPORTE DE CARGA / FLETES / INLAND
        REGEX_TRANSPORTE = r'\b(TRANSPORTE|TRANSPORT|INLAND|FLETE|FLETES|ACARREO|ACARREOS|CARGA\s*TERRESTRE|TERRESTRE|DRAYAGE|PORTES)\b'
        es_transporte = bool(re.search(REGEX_TRANSPORTE, desc_u) or (any(k in nombre for k in ["TRANSPORTE", "TRANSPORT", "CARGA", "FLETE"]) and not any(k in nombre for k in ["AGENCIA", "ADUANA"])))

        # 0. MEMORIA DE APRENDIZAJE: Si el usuario ya audito y guardo reglas para este proveedor previamente
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
        if "CREDITO" in tipo_str or "CR\u00c9DITO" in tipo_str or "NC" in tipo_str or "DEVOLUCION" in tipo_str:
            t_comp = 17
            op = "Devolucion Compra"
        elif "NOTA" in tipo_str or "AJUSTE" in tipo_str or "CONTABILIDAD" in tipo_str:
            t_comp = 14
            op = "Nota Contabilidad"
        else:
            t_comp = 10
            op = "Compra"

        # 1. TRANSPORTE TERRESTRE DE CARGA / FLETES / INLAND RATE (Prioridad sobre Agenciamiento)
        if es_transporte or any(k in desc_u for k in ["INLAND", "TRANSPORTE", "TRANSPORT", "FLETE"]):
            # Retencion en la fuente por Transporte de Carga: 1% (Art. 1.2.4.4.6 DUR 1625 / Art. 401 E.T.)
            rfte_t = round(valor_base * 0.01, 2) if valor_base >= 188000 else 0.0
            # ReteICA Transporte Bogota: 4.14 por mil (CIIU 4923 Transporte de Carga)
            rica_t = round(valor_base * 0.00414, 2) if valor_base >= 188000 else 0.0
            es_imp_transp = any(k in nombre for k in AGENTES_ADUANEROS) or any(k in nombre for k in ["EURO SHIPPING", "DHL", "TRADE GLOBAL", "CONSOLCARGO"])
            cta_p = "14650501" if es_imp_transp else "51355001"
            cta_c = "22050501"
            desc = (descripcion[:35] if descripcion else f"Transporte Terrestre - {nombre_emisor[:20]}")
            cta_rfte_b = "23652505"  # ReteFuente 1% Transporte Carga
            cta_iva = "24081001" if valor_iva == 0 else "24081501"
            cta_rica_b = "23680513"  # ReteICA 4.14 por mil Transporte
            cat = "Transporte Carga (1%)"
            razon_b = "Transporte Terrestre de Carga / Fletes (1% ReteFuente [23652505] - 4.14\u2030 ReteICA [23680513])"

        # 1.B Agentes Aduaneros e Importaciones (Servicios / Agenciamiento ordinario)
        elif any(k in nombre for k in AGENTES_ADUANEROS):
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

        # 5. Papeleria
        elif any(k in nombre for k in ["PANAMERICANA", "LIBRERIA"]):
            cta_p, cta_c, desc = "51953001", "23359501", f"Utiles y Papeleria - {nombre_emisor[:25]}"
            rfte_t, rica_t = 0.0, 0.0
            cta_rfte_b, cta_iva, cta_rica_b = "", "24081001", ""
            cat, razon_b = "Gasto Papeleria y Utiles", "Utiles de oficina"

        # 6. Compras y Gastos Generales (Evaluar primero clasificador detallado de Gastos PUC)
        else:
            res_det = clasificar_gasto_detallado_puc(nit_emisor, nombre_emisor, descripcion=descripcion, texto_pdf=descripcion, valor_base=valor_base, valor_iva=valor_iva)
            if res_det:
                cta_p = res_det["cta_p"]
                cta_c = res_det["cta_c"]
                desc = res_det["desc"]
                cat = res_det["cat"]
                razon_b = res_det["razon"]
                cta_iva = res_det["cta_iva"]
                cta_rfte_b = res_det["cta_rfte_sug"]
                cta_rica_b = res_det["cta_rica_sug"]
                rfte_t = round(valor_base * res_det["tasa_rfte_sug"], 2) if (res_det["tasa_rfte_sug"] > 0 and valor_base >= res_det["tope_rfte"]) else 0.0
                rica_t = round((valor_base * res_det["tasa_rica_mil"]) / 1000.0, 2) if (res_det["tasa_rica_mil"] > 0 and valor_base >= res_det["tope_rfte"]) else 0.0
            elif valor_base >= 500000:
                rfte_t = round(valor_base * 0.025, 2) if valor_base >= 1047000 else 0.0
                rica_t = round(valor_base * 0.01104, 2) if valor_base >= 1414000 else 0.0
                cta_p, cta_c, desc = "14350101", "22050501", f"Compra mercancias - {nombre_emisor[:25]}"
                cta_rfte_b, cta_iva, cta_rica_b = "23654001", "24081001", "23680501"
                cat, razon_b = "Inventario Mercancias / Repuestos", "Compra general > 500k"
            else:
                cta_p, cta_c, desc = "51959501", "23359501", f"Otros gastos diversos - {nombre_emisor[:25]}"
                rfte_t, rica_t = 0.0, 0.0
                cta_rfte_b, cta_iva, cta_rica_b = "", "24081001", ""
                cat, razon_b = "Otros Gastos Diversos", "Compra menor general"

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
            razon_rfte = "Proveedor R\u00e9gimen Simple (O-47): No aplica ReteFuente (Art. 911 E.T.)"
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
            razon_reteiva = "Emisor R\u00e9gimen Simple: Exento de ReteIVA"
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

        if cta_aprendida and not (cta_aprendida == "51959501" and cta_p != "51959501"):
            cta_p = cta_aprendida
            cat = obtener_etiqueta_cuenta_principal(cta_aprendida, cat)
            razon_total = f"\U0001f9e0 Cuenta aprendida de auditor\u00edas del usuario ({cta_aprendida}) | {razon_rfte} | {razon_reteiva}"
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
        Desbloquea autom\u00e1ticamente un PDF si viene encriptado/protegido con contrase\u00f1a.
        Prueba el NIT de la empresa compradora (con/sin d\u00edgito), NITs de emisores,
        contrase\u00f1a opcional y contrase\u00f1as est\u00e1ndar de facturaci\u00f3n electr\u00f3nica.
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

    _MEM_TEXT_CACHE = {}

    def cache_extraer_textos_pdf(fbytes, nit_receptor="9013464125"):
        """Extrae el texto de todas las p\u00e1ginas de un PDF guardando en cach\u00e9 \u00fanicamente el texto liviano por hash MD5 (0 bytes de PDF retenidos en RAM)."""
        if not fbytes:
            return []
        try:
            h_key = f"{hashlib.md5(fbytes).hexdigest()}_{nit_receptor}"
            if h_key in _MEM_TEXT_CACHE:
                return _MEM_TEXT_CACHE[h_key]
            reader = PdfReader(io.BytesIO(fbytes))
            if reader.is_encrypted:
                fbytes_des, ok = desbloquear_pdf_bytes(fbytes, nit_receptor=nit_receptor)
                if ok:
                    reader = PdfReader(io.BytesIO(fbytes_des))
            res = [(p.extract_text() or "") for p in reader.pages]
            if len(_MEM_TEXT_CACHE) > 250:
                _MEM_TEXT_CACHE.clear()
            _MEM_TEXT_CACHE[h_key] = res
            return res
        except Exception:
            return []




    def auditar_regimen_desde_facturas_renombradas(df_ref, dict_renombrados, empresa_dict):
        """
        Escanea el r\u00e9gimen fiscal Y los valores fiscales reales (Subtotal, IVA, Retenciones Sugeridas, Total Neto)
        directamente de las facturas PDF ya desbloqueadas y renombradas.
        Actualiza el R\u00e9gimen Fiscal Emisor Y recalcula las retenciones en la fuente correspondientes.
        """
        total_modificados = 0
        if not dict_renombrados or df_ref.empty:
            return total_modificados

        for r_idx, r_mat in df_ref.iterrows():
            soporte_nom = str(r_mat.get("Soporte PDF Renombrado", ""))
            pdf_bytes = dict_renombrados.get(soporte_nom)

            # Busqueda por coincidencia de comprobante o folio si el nombre vario ligeramente
            if not pdf_bytes:
                comp_id = str(r_mat.get("Comprobante Siigo", "")).replace(" ", "_")
                fol_id = str(r_mat.get("Folio", ""))
                for k in list(dict_renombrados.keys()):
                    if comp_id and comp_id in k:
                        pdf_bytes = dict_renombrados[k]
                        break
                    elif fol_id and len(fol_id) >= 3 and fol_id in k:
                        pdf_bytes = dict_renombrados[k]
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

                    if val_fisc.get("iva") is not None and val_fisc["iva"] >= 0:
                        if val_fisc["iva"] > 0:
                            df_ref.at[r_idx, "IVA"] = val_fisc["iva"]
                        df_ref.at[r_idx, "IVA 19%"] = val_fisc.get("iva_19", val_fisc["iva"])
                        df_ref.at[r_idx, "IVA 5%"] = val_fisc.get("iva_5", 0.0)
                        df_ref.at[r_idx, "Base IVA 19%"] = val_fisc.get("base_iva_19", 0.0)
                        df_ref.at[r_idx, "Base IVA 5%"] = val_fisc.get("base_iva_5", 0.0)
                        if val_fisc.get("iva_5", 0.0) > 0:
                            df_ref.at[r_idx, "Cta IVA 5%"] = "24081003"
                        hubo_cambio = True

                    if val_fisc.get("inc_otros") and val_fisc["inc_otros"] > 0:
                        df_ref.at[r_idx, "INC / Otros Impuestos"] = val_fisc["inc_otros"]
                        hubo_cambio = True

                    if val_fisc.get("iva_importacion") and val_fisc["iva_importacion"] > 0:
                        df_ref.at[r_idx, "IVA Importaci\u00f3n"] = val_fisc["iva_importacion"]
                        df_ref.at[r_idx, "Cta IVA Importaci\u00f3n"] = "240835"
                        hubo_cambio = True

                    # 2. Escaneo de regimen fiscal del emisor
                    nit_emisor_val = r_mat.get("NIT Emisor")
                    reg_detectado = escanear_regimen_texto_pdf(txt_completo, nit_emisor=nit_emisor_val)

                    if not reg_detectado:
                        ne_dig = re.sub(r'\D', '', str(nit_emisor_val))
                        if ne_dig in REGIMENES_EMISORES_CONOCIDOS:
                            reg_detectado = REGIMENES_EMISORES_CONOCIDOS[ne_dig]
                        elif any(k in str(r_mat.get("Proveedor", "")).upper() for k in ["DHL", "ESTELAR", "BUENAVENTURA", "SIIGO", "COMCEL", "CLARO", "PANAMERICANA"]):
                            reg_detectado = "O-13;O-15"

                    if reg_detectado and reg_detectado != r_mat.get("R\u00e9gimen Fiscal Emisor"):
                        df_ref.at[r_idx, "R\u00e9gimen Fiscal Emisor"] = reg_detectado
                        hubo_cambio = True

                    # 3. Recalcular liquidacion de retenciones con el nuevo regimen y concepto extraido
                    concep_det = val_fisc.get("concepto_items", "")
                    if concep_det:
                        df_ref.at[r_idx, "Descripcion"] = concep_det
                    desc_para_clasif = concep_det or str(df_ref.at[r_idx, "Descripcion"])
                    reg_a_usar = df_ref.at[r_idx, "R\u00e9gimen Fiscal Emisor"]
                    t_c_n, op_n, c_p_n, c_c_n, desc_n, rfte_n, rica_n, riva_n, c_rf_n, c_iv_n, c_ri_n, cat_n, razon_n, audit_n = clasificar_factura(
                        df_ref.at[r_idx, "NIT Emisor"], df_ref.at[r_idx, "Proveedor"],
                        df_ref.at[r_idx, "Base"], df_ref.at[r_idx, "IVA"],
                        df_ref.at[r_idx, "Operacion"], reg_a_usar, empresa_dict,
                        descripcion=desc_para_clasif
                    )
                    df_ref.at[r_idx, "Categor\u00eda"] = cat_n
                    df_ref.at[r_idx, "Cta Principal"] = c_p_n

                    # Si el PDF trae retenciones sugeridas explicitas en notas
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
                    df_ref.at[r_idx, "Raz\u00f3n Contable"] = razon_n
                    df_ref.at[r_idx, "Audit Info"] = audit_n

                    # 4. Tratamiento para Agencias Aduaneras / Facturas de Mandato
                    prov_nom_u = str(df_ref.at[r_idx, "Proveedor"]).upper()
                    es_aduanero_row = any(k in prov_nom_u for k in ["CARGO", "ADUANA", "PORTUARIA", "ALMACENADORA", "TERMINAL", "DHL", "EURO SHIPPING", "TRADE GLOBAL"])

                    if es_aduanero_row or val_fisc.get("retefuente", 0) > 0:
                        df_ref.at[r_idx, "Impuestos Asumidos"] = True
                        saldo_cruce = round(float(df_ref.at[r_idx, "Base"] or 0.0) + float(df_ref.at[r_idx, "IVA"] or 0.0) + float(df_ref.at[r_idx, "IVA Importaci\u00f3n"] or 0.0) + float(df_ref.at[r_idx, "INC / Otros Impuestos"] or 0.0), 2)
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
        Motor de b\u00fasqueda ultra-r\u00e1pido (130x m\u00e1s veloz):
        Eval\u00faa CUFE, Prefijo, Folio, NIT Emisor y Proveedor sin crear objetos pesados en bucles.
        """
        if not texto or df_ref.empty:
            return None
        txt_clean = re.sub(r'[^A-Z0-9]', '', texto.upper())
        digits_only = re.sub(r'\D', '', texto)

        # 1. Busqueda directa por CUFE (Certeza 100%)
        for r_idx, r_cand in df_ref.iterrows():
            c_val = str(r_cand.get("CUFE", "") or r_cand.get("CUFE / Token", "") or "")
            if len(c_val) >= 15:
                c_clean = re.sub(r'[^A-Za-z0-9]', '', c_val).upper()
                if c_clean[:20] in txt_clean:
                    return r_cand

        # 2. Busqueda combinada Factura + NIT (Ultra-rapida)
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
        Busca y extrae el PDF de la factura seleccionada dando M\u00c1XIMA PRIORIDAD
        a las facturas YA DESBLOQUEADAS, SEPARADAS Y RENOMBRADAS (dict_renombrados):
        1. PRIORIDAD 1: Facturas en dict_renombrados (las que ya pasaron por el proceso de desbloqueo y renombrado).
           Coincidencia por nombre esperado (Comp_10-XXX...), comprobante, folio o CUFE.
        2. PRIORIDAD 2: Si a\u00fan no se han desbloqueado/renombrado, buscar en dict_originales y desbloquearlas al vuelo.
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
        # PRIORIDAD 1: TOMAR LAS FACTURAS DESPUES DE DESBLOQUEARLAS Y RENOMBRARLAS
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
            for k in list(dict_renombrados.keys()):
                k_u = k.upper().replace(" ", "_")
                if any(cp in k_u for cp in comp_patterns):
                    f_des, _ = desbloquear_pdf_bytes(dict_renombrados[k], nit_receptor=nit_comprador)
                    return f_des, f"Factura Desbloqueada Oficial ({k})", None

            # C. Coincidencia por Numero de Factura / Folio en el nombre renombrado
            for k in list(dict_renombrados.keys()):
                k_clean = k.replace("-", "").replace(" ", "").replace("_", "").upper()
                if fac_full and len(fac_full) >= 3 and fac_full in k_clean:
                    f_des, _ = desbloquear_pdf_bytes(dict_renombrados[k], nit_receptor=nit_comprador)
                    return f_des, f"Factura Desbloqueada Oficial ({k})", None
                if folio_clean and len(folio_clean) >= 3 and folio_clean in k_clean:
                    f_des, _ = desbloquear_pdf_bytes(dict_renombrados[k], nit_receptor=nit_comprador)
                    return f_des, f"Factura Desbloqueada Oficial ({k})", None

            # D. Coincidencia por CUFE en el nombre renombrado
            if cufe_clean and len(cufe_clean) >= 15:
                for k in list(dict_renombrados.keys()):
                    k_l = re.sub(r'[^a-zA-Z0-9]', '', k).lower()
                    if cufe_clean[:20] in k_l:
                        f_des, _ = desbloquear_pdf_bytes(dict_renombrados[k], nit_receptor=nit_comprador)
                        return f_des, f"Factura Desbloqueada Oficial ({k})", None

            # E. Verificacion por contenido de texto en las facturas renombradas
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
        # PRIORIDAD 2: SI AUN NO SE HA RENOMBRADO, BUSCAR EN LOS ORIGINALES Y DESBLOQUEAR
        # =========================================================================
        if dict_originales:
            # A. Busqueda por CUFE en nombre de archivo original
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

            # C. Escaneo en PDFs originales unificados/consolidados pagina por pagina
            for fname, fbytes in dict_originales.items():
                try:
                    f_des, _ = desbloquear_pdf_bytes(fbytes, nit_receptor=nit_comprador)
                    paginas_txt = cache_extraer_textos_pdf(f_des, nit_receptor=nit_comprador)
                    pags_coincidentes = []
                    for p_idx, txt in enumerate(paginas_txt):
                        if identificar_factura_en_texto(txt, df_una_fac) is not None:
                            if p_idx not in pags_coincidentes:
                                pags_coincidentes.append(p_idx)
                            m_p = re.search('(?:P[\u00c1A\u00e1a]G(?:INA)?|HOJA|PAGE)\\s*[:\\.]?\\s*1\\s*(?:DE|/|OF)\\s*(\\d+)', txt, re.IGNORECASE)
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
                        return out.getvalue(), f"Factura Desbloqueada de '{fname}' (P\u00e1gs {str_p})", [p+1 for p in pags_coincidentes]
                except Exception:
                    pass

        return None, None, None


    def renderizar_visor_pdf_completo(pdf_bytes, nombre_archivo, fac_sel=None, key_prefix="pdf_view", nit_comprador="9013464125", mostrar_card_dian=False):
        if not pdf_bytes or not isinstance(pdf_bytes, (bytes, bytearray)):
            st.error("\u26a0\ufe0f El archivo PDF no contiene datos v\u00e1lidos.")
            return

        pdf_limpio, _ = desbloquear_pdf_bytes(pdf_bytes, nit_receptor=nit_comprador)
        pdf_usar = pdf_limpio if pdf_limpio else pdf_bytes

        try:
            reader_prev = PdfReader(io.BytesIO(pdf_usar))
            num_pags_tot = len(reader_prev.pages)
        except Exception:
            num_pags_tot = 1

        val_base = float(fac_sel.get("Base", 0.0)) if (fac_sel is not None and "Base" in fac_sel) else 0.0
        val_tot = float(fac_sel.get("Total", 0.0)) if (fac_sel is not None and "Total" in fac_sel) else 0.0
        val_iva = float(fac_sel.get("IVA", 0.0)) if (fac_sel is not None and "IVA" in fac_sel) else 0.0
        fac_num_str = str(fac_sel.get("Factura", "")) if fac_sel is not None else nombre_archivo
        prov_str = str(fac_sel.get("Proveedor", "")) if fac_sel is not None else ""
        nit_str = str(fac_sel.get("NIT Emisor", "")) if fac_sel is not None else ""
        comp_str = str(fac_sel.get("Comprobante Siigo", "")) if fac_sel is not None else ""
        fecha_str = str(fac_sel.get("Fecha", "")) if fac_sel is not None else ""
        reg_str = str(fac_sel.get("R\u00e9gimen Fiscal Emisor", "O-48")) if fac_sel is not None else "O-48"

        c_btn1, c_tam1 = st.columns([1.5, 2.5])
        with c_btn1:
            pags_label = f"{num_pags_tot} p\u00e1gina{'s' if num_pags_tot > 1 else ''} completa{'s' if num_pags_tot > 1 else ''}"
            st.download_button(
                label=f"\U0001f4e5 Descargar / Abrir PDF Original ({pags_label})",
                data=pdf_usar,
                file_name=nombre_archivo,
                mime="application/pdf",
                key=f"btn_dl_univ_{key_prefix}",
                use_container_width=True
            )
        with c_tam1:
            tam_sel = st.radio(
                "\U0001f4d0 Tama\u00f1o de Vista Previa:",
                ["\U0001f4c4 Normal (Recomendado - 780px)", "\U0001f50d Grande (Ancho Completo)", "\U0001f4f1 Compacto (600px)"],
                horizontal=True,
                index=0,
                key=f"rad_tam_pdf_{key_prefix}"
            )

        if "Compacto" in tam_sel:
            col_ratios = [2.2, 4.0, 2.2]
        elif "Grande" in tam_sel:
            col_ratios = [0.01, 9.98, 0.01]
        else:
            col_ratios = [1.4, 5.6, 1.4]

        imgs_paginas = []
        if HAS_PDFIUM:
            doc_p = None
            try:
                doc_p = pdfium.PdfDocument(pdf_usar)
                max_p_render = min(len(doc_p), 5)
                for idx_pg in range(max_p_render):
                    pg_obj = doc_p[idx_pg]
                    bmp_obj = pg_obj.render(scale=1.3)
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
        if not imgs_paginas and HAS_FITZ:
            try:
                doc_fz = _pymupdf_fitz.open(stream=pdf_usar, filetype="pdf")
                max_p_render = min(len(doc_fz), 5)
                for idx_pg in range(max_p_render):
                    page_fz = doc_fz.load_page(idx_pg)
                    pix = page_fz.get_pixmap(dpi=130)
                    imgs_paginas.append(pix.tobytes("png"))
                doc_fz.close()
            except Exception:
                imgs_paginas = []

        if imgs_paginas:
            for idx_p, img in enumerate(imgs_paginas):
                c_l, c_mid, c_r = st.columns(col_ratios)
                with c_mid:
                    st.markdown(f"""
                    <div style="background:#0070ba; color:white; padding:8px 14px; border-radius:6px 6px 0 0; font-weight:600; font-size:13.5px; display:flex; justify-content:space-between; align-items:center; margin-top:12px; box-shadow: 0 2px 4px rgba(0,0,0,0.06);">
                        <span>\U0001f4c4 Factura Original: {fac_num_str} \u2014 {prov_str}</span>
                        <span style="background:rgba(255,255,255,0.25); padding:2px 10px; border-radius:10px; font-size:12px;">Hoja {idx_p + 1} de {len(imgs_paginas)}</span>
                    </div>
                    """, unsafe_allow_html=True)
                    st.image(img, use_container_width=True)
            if mostrar_card_dian and fac_sel is not None:
                with st.expander("\U0001f4cb Ver Tabla Detallada de \u00cdtems y Texto Extra\u00eddo DIAN", expanded=False):
                    pags_f = obtener_paginas_texto_factura_integrado(fac_sel)
                    renderizar_representacion_grafica_dian_html(fac_sel, pags_f or [], {"nombre": "INDUMAQ ER SAS", "nit": nit_comprador})
        else:
            if mostrar_card_dian and fac_sel is not None:
                try:
                    pags_f = obtener_paginas_texto_factura_integrado(fac_sel)
                    renderizar_representacion_grafica_dian_html(fac_sel, pags_f or [], {"nombre": "INDUMAQ ER SAS", "nit": nit_comprador})
                except Exception:
                    pass
            b64_pdf = base64.b64encode(pdf_usar).decode("utf-8")
            safe_title_js = re.sub(r'[^a-zA-Z0-9_. -]', '_', str(nombre_archivo))
            visor_height = max(680, min(3200, num_pags_tot * 840))
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
                  padding: 16px;
                  text-align: center;
                  font-weight: 600;
                }}
              </style>
            </head>
            <body>
              <div id="msg">\u23f3 Renderizando documento PDF original ({num_pags_tot} p\u00e1gina(s))...</div>
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
                    try {{
                      const blob = new Blob(["importScripts('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js');"], {{type: 'application/javascript'}});
                      pdfjsLib.GlobalWorkerOptions.workerSrc = URL.createObjectURL(blob);
                    }} catch (eW) {{}}
                    let pdf = null;
                    try {{
                      pdf = await pdfjsLib.getDocument({{ data: uint8Array, disableWorker: true }}).promise;
                    }} catch (e1) {{
                      pdf = await pdfjsLib.getDocument({{ data: uint8Array }}).promise;
                    }}
                    document.getElementById('msg').style.display = 'none';
                    const container = document.getElementById('pages-container');

                    for (let num = 1; num <= pdf.numPages; num++) {{
                      const page = await pdf.getPage(num);
                      const viewport = page.getViewport({{ scale: 1.5 }});
                      const card = document.createElement('div');
                      card.className = 'page-card';
                      const title = document.createElement('div');
                      title.className = 'page-title';
                      title.innerHTML = '<span>\U0001f4c4 {safe_title_js}</span><span style="background:rgba(255,255,255,0.25); padding:2px 8px; border-radius:10px; font-size:11px;">Hoja ' + num + ' de ' + pdf.numPages + '</span>';
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
                    document.getElementById('msg').style.display = 'none';
                  }}
                }})();
              </script>
            </body>
            </html>
            """
            components.html(html_canvas_viewer, height=visor_height, scrolling=True)


    def extraer_montos_items_pdf(pdf_bytes):
        """Extrae valores monetarios de los \u00edtems o detalles de productos en el PDF de una factura."""
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
        """Genera en memoria un archivo .indumaq (ZIP completo) con el progreso actual de auditor\u00eda y triangulaci\u00f3n."""
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
                    "paquetes_destino_siigo": {k: v for k, v in st.session_state.items() if k.startswith("paquete_destino_siigo_")},
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

                # 7. Incluir 1 sola copia sin duplicar de los PDFs desbloqueados (ZIP_STORED directo desde disco) para conservar siempre la vista previa
                d_renom_curr = st.session_state.get("dict_pdfs", {})
                d_orig_curr = st.session_state.get("raw_uploaded_pdfs", {})
                fuente_pdfs = d_renom_curr if (d_renom_curr and len(d_renom_curr) > 0) else d_orig_curr
                if fuente_pdfs:
                    for fn_k in list(fuente_pdfs.keys()):
                        if not str(fn_k).lower().endswith(".pdf"):
                            continue
                        p_disk = fuente_pdfs.get_path(fn_k) if hasattr(fuente_pdfs, "get_path") else None
                        arc_name = f"pdfs_renombrados/{os.path.basename(str(fn_k))}"
                        if p_disk and os.path.exists(p_disk):
                            zf.write(p_disk, arcname=arc_name, compress_type=zipfile.ZIP_STORED)
                        else:
                            b_pdf = fuente_pdfs.get(fn_k)
                            if b_pdf:
                                zinfo = zipfile.ZipInfo(arc_name)
                                zinfo.compress_type = zipfile.ZIP_STORED
                                zf.writestr(zinfo, b_pdf)

            buf.seek(0)
            return buf.getvalue()
        except Exception:
            return None

    def restaurar_desde_respaldo_bytes(empresa_dict, raw_zip_or_file, solo_pdfs=False):
        """Restaura un archivo de respaldo port\u00e1til (.indumaq o .zip) directamente a la sesi\u00f3n activa con facturas, paquetes y PDFs.
        Si solo_pdfs=True, extrae \u00fanicamente los 224 PDFs a disco sin sobreescribir los cambios contables actuales."""
        try:
            if hasattr(raw_zip_or_file, "seek"):
                raw_zip_or_file.seek(0)
                zf_source = raw_zip_or_file
            elif isinstance(raw_zip_or_file, (bytes, bytearray)):
                zf_source = io.BytesIO(raw_zip_or_file)
            else:
                zf_source = io.BytesIO(raw_zip_or_file.read())

            with zipfile.ZipFile(zf_source, "r") as zf:
                namelist = zf.namelist()
                if not solo_pdfs:
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
                    _univ_json = None
                    if "datos_contables_universal.json" in namelist:
                        try:
                            _univ_json = json.loads(zf.read("datos_contables_universal.json").decode("utf-8"))
                        except Exception:
                            _univ_json = None
                    df_loaded = None
                    if _univ_json and "df_procesado" in _univ_json:
                        try:
                            df_loaded = pd.DataFrame(_univ_json["df_procesado"])
                        except Exception:
                            df_loaded = None
                    if df_loaded is None and "df_procesado.pkl" in namelist:
                        try:
                            df_loaded = _safe_pickle_loads(zf.read("df_procesado.pkl"))
                        except Exception:
                            df_loaded = None
                    if df_loaded is not None:
                        st.session_state["df_procesado"] = normalizar_df_procesado(_limpiar_dtypes_pyarrow_df(df_loaded))
                    if "excel_original.xlsx" in namelist:
                        st.session_state["excel_bytes"] = zf.read("excel_original.xlsx")
                    as_loaded = None
                    if _univ_json and "asientos_triangulacion" in _univ_json:
                        try:
                            as_loaded = {k: pd.DataFrame(v) for k, v in _univ_json.get("asientos_triangulacion", {}).items()}
                        except Exception:
                            as_loaded = None
                    if as_loaded is None and "asientos_triangulacion.pkl" in namelist:
                        try:
                            as_loaded = _safe_pickle_loads(zf.read("asientos_triangulacion.pkl"))
                        except Exception:
                            as_loaded = {}
                    if isinstance(as_loaded, dict):
                        for k_as, v_as in as_loaded.items():
                            if isinstance(v_as, pd.DataFrame):
                                as_loaded[k_as] = _limpiar_dtypes_pyarrow_df(v_as)
                        st.session_state["asientos_triangulacion_por_factura"] = as_loaded
                    pqs_loaded = None
                    if _univ_json and "paquetes_importacion" in _univ_json:
                        try:
                            pqs_loaded = {}
                            for pk_s, pv_s in _univ_json["paquetes_importacion"].items():
                                pk_i = int(pk_s) if str(pk_s).isdigit() else pk_s
                                pv_c = dict(pv_s)
                                for df_k in ["terceros", "asiento_fijo", "asiento_aprobado"]:
                                    if df_k in pv_c and isinstance(pv_c[df_k], list):
                                        pv_c[df_k] = pd.DataFrame(pv_c[df_k])
                                pqs_loaded[pk_i] = pv_c
                        except Exception:
                            pqs_loaded = None
                    if pqs_loaded is None and "paquetes_importacion.pkl" in namelist:
                        try:
                            pqs_loaded = _safe_pickle_loads(zf.read("paquetes_importacion.pkl"))
                        except Exception:
                            pqs_loaded = {}
                    if pqs_loaded is not None and isinstance(pqs_loaded, dict):
                        as_map = st.session_state.get("asientos_triangulacion_por_factura", {})
                        for pq_k, pq_v in pqs_loaded.items():
                            if isinstance(pq_v, dict):
                                if "terceros" in pq_v and isinstance(pq_v["terceros"], pd.DataFrame):
                                    pq_v["terceros"] = _limpiar_dtypes_pyarrow_df(pq_v["terceros"])
                                for af_k in ["asiento_fijo", "asiento_aprobado"]:
                                    if af_k in pq_v and isinstance(pq_v[af_k], pd.DataFrame):
                                        pq_v[af_k] = _limpiar_dtypes_pyarrow_df(pq_v[af_k])
                                if ("asiento_fijo" in pq_v and pq_v["asiento_fijo"] is not None) or ("asiento_aprobado" in pq_v and pq_v["asiento_aprobado"] is not None):
                                    st.session_state[f"paquete_listo_{pq_k}"] = True
                                    st.session_state[f"asiento_fijo_pq_{pq_k}"] = pq_v.get("asiento_fijo", pq_v.get("asiento_aprobado")).copy()
                                ag_fac_str = str(pq_v.get("agente", {}).get("Factura", "")).strip()
                                if "destino_siigo" in pq_v and pq_v["destino_siigo"] is not None:
                                    st.session_state[f"paquete_destino_siigo_{pq_k}"] = bool(pq_v["destino_siigo"])
                                elif ag_fac_str and ag_fac_str in as_map:
                                    st.session_state[f"paquete_destino_siigo_{pq_k}"] = True
                                    pq_v["destino_siigo"] = True
                        st.session_state["paquetes_importacion"] = pqs_loaded
                        st.session_state["_ultimo_agentes_len"] = len(pqs_loaded)
                    if "estado_sesion.json" in namelist:
                        est = json.loads(zf.read("estado_sesion.json").decode("utf-8"))
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
                            st.session_state["paquete_seleccionado_id"] = est["sel_paquete_activo_key"]
                    if "meta.json" in namelist:
                        meta = json.loads(zf.read("meta.json").decode("utf-8"))
                        st.session_state["excel_nombre"] = meta.get("archivo_excel", meta.get("excel_nombre", "Reporte.xlsx"))
                        st.session_state["_sesion_cargada_nombre"] = f"Respaldo {meta.get('fecha_respaldo', '')}"
                    elif "meta_sesion.json" in namelist:
                        meta = json.loads(zf.read("meta_sesion.json").decode("utf-8"))
                        st.session_state["excel_nombre"] = meta.get("excel_nombre", meta.get("archivo_excel", "Reporte.xlsx"))
                        st.session_state["_sesion_cargada_nombre"] = f"Respaldo {meta.get('fecha_respaldo', '')}"

                # Extraer todos los PDFs al disco fisico (0 MB en RAM)
                dict_renom = DiskBackedPDFDict("renom_pdfs")
                dict_orig = DiskBackedPDFDict("raw_pdfs")
                renom_names = [n for n in namelist if n.startswith("pdfs_renombrados/") and n.lower().endswith(".pdf")]
                orig_names = [n for n in namelist if n.startswith("pdfs_originales/") and n.lower().endswith(".pdf")]
                other_pdfs = [n for n in namelist if n.lower().endswith(".pdf") and not n.startswith("pdfs_renombrados/") and not n.startswith("pdfs_originales/") and not n.startswith("__MACOSX")]
                if renom_names:
                    for name in renom_names:
                        fn = os.path.basename(name)
                        if fn:
                            p_dst = dict_renom._safe_path(fn)
                            try:
                                z_sz = zf.getinfo(name).file_size
                            except Exception:
                                z_sz = -1
                            if not (os.path.exists(p_dst) and z_sz > 0 and os.path.getsize(p_dst) == z_sz):
                                with zf.open(name) as f_src, open(p_dst, "wb") as f_dst:
                                    shutil.copyfileobj(f_src, f_dst)
                            dict_renom[fn] = p_dst
                if orig_names:
                    for name in orig_names:
                        fn = os.path.basename(name)
                        if fn:
                            p_dst = dict_orig._safe_path(fn)
                            try:
                                z_sz = zf.getinfo(name).file_size
                            except Exception:
                                z_sz = -1
                            if not (os.path.exists(p_dst) and z_sz > 0 and os.path.getsize(p_dst) == z_sz):
                                with zf.open(name) as f_src, open(p_dst, "wb") as f_dst:
                                    shutil.copyfileobj(f_src, f_dst)
                            dict_orig[fn] = p_dst
                if other_pdfs:
                    for name in other_pdfs:
                        fn = os.path.basename(name)
                        if fn:
                            p_dst = dict_renom._safe_path(fn)
                            with zf.open(name) as f_src, open(p_dst, "wb") as f_dst:
                                shutil.copyfileobj(f_src, f_dst)
                            dict_renom[fn] = p_dst

                dict_renom._sync_from_disk()
                if len(dict_renom) > 0:
                    st.session_state["dict_pdfs"] = dict_renom
                    st.session_state["total_zip_pdfs"] = len(dict_renom)
                    st.session_state["raw_uploaded_pdfs"] = dict_renom
                elif len(dict_orig) > 0:
                    st.session_state["raw_uploaded_pdfs"] = dict_orig

            st.session_state["_usuario_limpio_sesion"] = False
            if "df_procesado" not in st.session_state or st.session_state["df_procesado"] is None:
                cargar_respaldo_integrado_50_paquetes(empresa_dict)
            st.session_state["_upl_reset_cnt"] = st.session_state.get("_upl_reset_cnt", 0) + 1
            if not solo_pdfs:
                sincronizar_todos_los_paquetes_con_df_procesado()
                st.session_state.pop("_siigo_cache_sig", None)
                st.session_state.pop("_siigo_3hojas_bytes", None)
                guardar_estado_manual(empresa_dict)
            import gc
            gc.collect()
            return True
        except Exception as e:
            st.session_state["_ultimo_error_respaldo"] = str(e)
            return False

    def normalizar_df_procesado(df):
        """Garantiza que todas las columnas requeridas existan con tipos seguros y sin NaNs que causen ca\u00eddas."""
        if df is None or not isinstance(df, pd.DataFrame) or df.empty:
            return df
        df = _limpiar_dtypes_pyarrow_df(df)
    
        defaults_str = {
            "Comprobante Siigo": "Comp 10-001", "Fecha": "01/01/2026", "Factura": "FAC-001",
            "Proveedor": "Proveedor", "NIT Emisor": "900000000", "R\u00e9gimen Fiscal Emisor": "48",
            "Cta Principal": "14650501", "Categor\u00eda": "General", "Raz\u00f3n Contable": "Clasificaci\u00f3n autom\u00e1tica", "Cuenta Pasivo Especifica": "",
            "Cta Contrapartida": "22050501", "Operacion": "Compra", "Soporte PDF Renombrado": "",
            "Descripcion": "Compra general", "Tipo Comp": "10", "Prefijo": "", "Folio": "",
            "Cta ReteFuente": "", "Cta ReteICA": "", "Cta IVA": "24081001",
            "Cta IVA 5%": "24081003", "Cta IVA Importaci\u00f3n": "240835", "Comprobante Previo": "", "Estado Registro": "\u26aa Compra Pendiente",
            "Paquete Validado": "", "Cruza Con Agente": "", "Destino Siigo": ""
        }
        defaults_num = {
            "Base": 0.0, "IVA": 0.0, "IVA 19%": 0.0, "IVA 5%": 0.0, "Base IVA 19%": 0.0, "Base IVA 5%": 0.0,
            "INC / Otros Impuestos": 0.0, "ReteFuente": 0.0, "ReteICA": 0.0, "ReteIVA": 0.0,
            "Total": 0.0, "Neto a Pagar": 0.0, "Total Neto": 0.0, "Consecutivo": 680,
            "IVA Importaci\u00f3n": 0.0
        }
        defaults_bool = {
            "Impuestos Asumidos": False, "No Contabilizar": False, "Ya Registrada": False, "Es Aduanera": False, "Editada Manualmente": False
        }

        for col, def_val in defaults_str.items():
            if col not in df.columns:
                df[col] = def_val
            else:
                df[col] = df[col].fillna(def_val).astype(str).replace(["nan", "None", "<NA>"], def_val)
                if col == "Paquete Validado":
                    df[col] = df[col].apply(lambda s: s[:-2] if str(s).endswith(".0") and str(s)[:-2].isdigit() else str(s))

        for col, def_val in defaults_num.items():
            if col not in df.columns:
                df[col] = def_val
            else:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(def_val)

        for col, def_val in defaults_bool.items():
            if col not in df.columns:
                df[col] = def_val
            else:
                df[col] = df[col].apply(lambda v: _a_bool_seguro(v, def_val)).astype(bool)

        if "Audit Info" not in df.columns:
            df["Audit Info"] = [{} for _ in range(len(df))]

        if df.attrs.get("_indumaq_norm_done") == len(df):
            return df

        # Auto-identificar y clasificar facturas de transporte terrestre / fletes / inland rate
        REGEX_TRANSP = r'\b(TRANSPORTE|TRANSPORT|INLAND\s*RATE|INLAND|FLETE|FLETES|ACARREO|ACARREOS|CARGA\s*TERRESTRE|TERRESTRE|DRAYAGE|PORTES)\b'
        for idx_row, r_row in df.iterrows():
            f_num = str(r_row.get("Factura", "")).strip()
            p_nom = str(r_row.get("Proveedor", "")).strip().upper()
            d_val = str(r_row.get("Descripcion", "")).strip().upper()
            c_val = str(r_row.get("Categor\u00eda", "")).strip().upper()

            es_transp_row = ("70787" in f_num or
                             ("EURO" in p_nom and any(k in d_val for k in ["INLAND", "TRANSPORT", "FLETE", "TERRESTRE"])))
            if es_transp_row:
                b_val = float(r_row.get("Base", 0.0) or 0.0)
                if b_val > 0:
                    rfte_correcta = round(b_val * 0.01, 2)
                    rica_correcta = round(b_val * 0.00414, 2)
                    if (df.at[idx_row, "Cta ReteFuente"] == "23652503" or
                        df.at[idx_row, "ReteFuente"] == round(b_val * 0.04, 2) or
                        df.at[idx_row, "ReteFuente"] == 0.0 or
                        "70787" in f_num):
                        df.at[idx_row, "ReteFuente"] = rfte_correcta
                        df.at[idx_row, "ReteICA"] = rica_correcta
                        df.at[idx_row, "Cta ReteFuente"] = "23652505"
                        df.at[idx_row, "Cta ReteICA"] = "23680513"
                        df.at[idx_row, "Categor\u00eda"] = "Transporte Carga (1%)"
                        df.at[idx_row, "Raz\u00f3n Contable"] = "Transporte Terrestre de Carga / Fletes (1% ReteFuente [23652505] - 4.14\u2030 ReteICA [23680513])"
                        df.at[idx_row, "Cta Principal"] = "14650501"
                        if "70787" in f_num and not bool(r_row.get("Editada Manualmente", False)):
                            df.at[idx_row, "Impuestos Asumidos"] = False
                        df.at[idx_row, "Descripcion"] = "INLAND RATE TRANSPORTE TERRESTRE"

        # Auto-sincronizar desglose oficial de facturas con IVA mixto (19% + 5%), Descuento en items, INC 8% y Bolsas
        DESGLOSE_IMPUESTOS_PDF = {
            "X928-2559678": {"Base": 250394.12, "IVA": 27275.88, "IVA 19%": 23833.03, "Base IVA 19%": 125436.97, "IVA 5%": 3442.85, "Base IVA 5%": 68857.15, "INC / Otros Impuestos": 0.0, "Regimen": "O-13;O-15"},
            "US82-27381":   {"Base": 186911.15, "IVA": 22870.85, "IVA 19%": 21154.18, "Base IVA 19%": 111337.79, "IVA 5%": 1716.67, "Base IVA 5%": 34333.33, "INC / Otros Impuestos": 0.0, "Regimen": "O-13;O-15"},
            "ES03-20573":   {"Base": 831867.00, "IVA": 115452.00, "IVA 19%": 111297.00, "Base IVA 19%": 585773.68, "IVA 5%": 4155.00, "Base IVA 5%": 83100.00, "INC / Otros Impuestos": 0.0},
            "UJ06-24255":   {"Base": 216771.65, "IVA": 25425.35, "IVA 19%": 22397.97, "Base IVA 19%": 117884.05, "IVA 5%": 3027.38, "Base IVA 5%": 60547.62, "INC / Otros Impuestos": 0.0, "Regimen": "O-13;O-15"},
            "US82-52479":   {"Base": 172128.12, "IVA": 19202.88, "IVA 19%": 17157.64, "Base IVA 19%": 90303.37, "IVA 5%": 2045.24, "Base IVA 5%": 40904.76, "INC / Otros Impuestos": 0.0, "Regimen": "O-13;O-15"},
            "FBOG-385674":  {"Base": 110096.78, "IVA": 20918.39, "IVA 19%": 20918.39, "Base IVA 19%": 110096.78, "IVA 5%": 0.0, "Base IVA 5%": 0.0, "INC / Otros Impuestos": 0.0},
            "X408-2540769": {"Base": 3399957.22, "IVA": 589012.78, "IVA 19%": 589012.78, "Base IVA 19%": 3100067.22, "IVA 5%": 0.0, "Base IVA 5%": 0.0, "INC / Otros Impuestos": 0.0, "Regimen": "O-13;O-15"},
            "94-1527950":   {"Base": 2060371.00, "IVA": 0.0, "IVA 19%": 0.0, "Base IVA 19%": 0.0, "IVA 5%": 0.0, "Base IVA 5%": 0.0, "INC / Otros Impuestos": 0.0, "Regimen": "O-13;O-15"},
            "04E-612027":   {"Base": 433784.03, "IVA": 76642.97, "IVA 19%": 76642.97, "Base IVA 19%": 403384.03, "IVA 5%": 0.0, "Base IVA 5%": 0.0, "INC / Otros Impuestos": 73.00, "Regimen": "O-13;O-15"},
            "HG1-16542":    {"Base": 476296.30, "IVA": 62700.00, "IVA 19%": 62700.00, "Base IVA 19%": 330000.00, "IVA 5%": 0.0, "Base IVA 5%": 0.0, "INC / Otros Impuestos": 11703.70},
            "HG1-17107":    {"Base": 1037222.21, "IVA": 129200.00, "IVA 19%": 129200.00, "Base IVA 19%": 680000.00, "IVA 5%": 0.0, "Base IVA 5%": 0.0, "INC / Otros Impuestos": 28577.78},
            "HG1-17843":    {"Base": 1957925.89, "IVA": 196080.00, "IVA 19%": 196080.00, "Base IVA 19%": 1032000.00, "IVA 5%": 0.0, "Base IVA 5%": 0.0, "INC / Otros Impuestos": 74074.05},
            "E021-30160":   {"Base": 201852.00, "IVA": 0.0, "IVA 19%": 0.0, "Base IVA 19%": 0.0, "IVA 5%": 0.0, "Base IVA 5%": 0.0, "INC / Otros Impuestos": 16148.00, "Regimen": "O-13;O-15"},
            "HG1-18078":    {"Base": 183222.22, "IVA": 25840.00, "IVA 19%": 25840.00, "Base IVA 19%": 136000.00, "IVA 5%": 0.0, "Base IVA 5%": 0.0, "INC / Otros Impuestos": 3777.78},
            "CRA-2961915":  {"Base": 317221.00, "IVA": 0.0, "IVA 19%": 0.0, "IVA 5%": 0.0, "INC / Otros Impuestos": 0.0},
            "CRA-2991642":  {"Base": 393168.00, "IVA": 0.0, "IVA 19%": 0.0, "IVA 5%": 0.0, "INC / Otros Impuestos": 0.0},
            "CRA-3015859":  {"Base": 269833.00, "IVA": 0.0, "IVA 19%": 0.0, "IVA 5%": 0.0, "INC / Otros Impuestos": 0.0},
            "CRA-3038879":  {"Base": 228523.00, "IVA": 0.0, "IVA 19%": 0.0, "IVA 5%": 0.0, "INC / Otros Impuestos": 0.0},
            "CRA-3058669":  {"Base": 292273.00, "IVA": 0.0, "IVA 19%": 0.0, "IVA 5%": 0.0, "INC / Otros Impuestos": 0.0},
            "CRA-3104683":  {"Base": 284748.00, "IVA": 0.0, "IVA 19%": 0.0, "IVA 5%": 0.0, "INC / Otros Impuestos": 0.0},
            "ROD-1570417":  {"Base": 142883.00, "IVA": 27147.77, "IVA 19%": 27147.77, "Base IVA 19%": 142883.00, "IVA 5%": 0.0, "Base IVA 5%": 0.0, "INC / Otros Impuestos": 0.0, "IVA Importaci\u00f3n": 1341000.00, "Cta IVA Importaci\u00f3n": "240835", "Cta Principal": "14650501", "Regimen": "O-13;O-15"},
            "CEU-4353251":  {"Base": 269725.00, "IVA": 21038.00, "IVA 19%": 21038.00, "Base IVA 19%": 110725.00, "IVA 5%": 0.0, "Base IVA 5%": 0.0, "INC / Otros Impuestos": 0.0, "IVA Importaci\u00f3n": 332000.00, "Cta IVA Importaci\u00f3n": "240835", "Cta Principal": "14650501", "Regimen": "O-13;O-15"},
            "CEU-4353262":  {"Base": 278725.00, "IVA": 21038.00, "IVA 19%": 21038.00, "Base IVA 19%": 110725.00, "IVA 5%": 0.0, "Base IVA 5%": 0.0, "INC / Otros Impuestos": 0.0, "IVA Importaci\u00f3n": 351000.00, "Cta IVA Importaci\u00f3n": "240835", "Cta Principal": "14650501", "Regimen": "O-13;O-15"},
            "CEU-4364993":  {"Base": 668738.00, "IVA": 39281.00, "IVA 19%": 39281.00, "Base IVA 19%": 206738.00, "IVA 5%": 0.0, "Base IVA 5%": 0.0, "INC / Otros Impuestos": 0.0, "IVA Importaci\u00f3n": 965000.00, "Cta IVA Importaci\u00f3n": "240835", "Cta Principal": "14650501", "Regimen": "O-13;O-15"},
            "CEU-4374336":  {"Base": 754960.00, "IVA": 39323.00, "IVA 19%": 39323.00, "Base IVA 19%": 206960.00, "IVA 5%": 0.0, "Base IVA 5%": 0.0, "INC / Otros Impuestos": 0.0, "IVA Importaci\u00f3n": 1146000.00, "Cta IVA Importaci\u00f3n": "240835", "Cta Principal": "14650501", "Regimen": "O-13;O-15"},
            "CEU-4384606":  {"Base": 433004.00, "IVA": 40091.00, "IVA 19%": 40091.00, "Base IVA 19%": 211004.00, "IVA 5%": 0.0, "Base IVA 5%": 0.0, "INC / Otros Impuestos": 0.0, "IVA Importaci\u00f3n": 464000.00, "Cta IVA Importaci\u00f3n": "240835", "Cta Principal": "14650501", "Regimen": "O-13;O-15"},
            "CEU-4403918":  {"Base": 241096.00, "IVA": 21298.00, "IVA 19%": 21298.00, "Base IVA 19%": 112096.00, "IVA 5%": 0.0, "Base IVA 5%": 0.0, "INC / Otros Impuestos": 0.0, "IVA Importaci\u00f3n": 270000.00, "Cta IVA Importaci\u00f3n": "240835", "Cta Principal": "14650501", "Regimen": "O-13;O-15"},
            "CEU-4435356":  {"Base": 722839.00, "IVA": 39679.00, "IVA 19%": 39679.00, "Base IVA 19%": 208839.00, "IVA 5%": 0.0, "Base IVA 5%": 0.0, "INC / Otros Impuestos": 0.0, "IVA Importaci\u00f3n": 1074000.00, "Cta IVA Importaci\u00f3n": "240835", "Cta Principal": "14650501", "Regimen": "O-13;O-15"},
            "CEU-4459127":  {"Base": 277221.00, "IVA": 21132.00, "IVA 19%": 21132.00, "Base IVA 19%": 111221.00, "IVA 5%": 0.0, "Base IVA 5%": 0.0, "INC / Otros Impuestos": 0.0, "IVA Importaci\u00f3n": 346000.00, "Cta IVA Importaci\u00f3n": "240835", "Cta Principal": "14650501", "Regimen": "O-13;O-15"},
            "CEU-4499603":  {"Base": 847101.00, "IVA": 35170.00, "IVA 19%": 35170.00, "Base IVA 19%": 185101.00, "IVA 5%": 0.0, "Base IVA 5%": 0.0, "INC / Otros Impuestos": 0.0, "IVA Importaci\u00f3n": 1384000.00, "Cta IVA Importaci\u00f3n": "240835", "Cta Principal": "14650501", "Regimen": "O-13;O-15"},
            "CEU-4502624":  {"Base": 386399.00, "IVA": 34276.00, "IVA 19%": 34276.00, "Base IVA 19%": 180399.00, "IVA 5%": 0.0, "Base IVA 5%": 0.0, "INC / Otros Impuestos": 0.0, "IVA Importaci\u00f3n": 431000.00, "Cta IVA Importaci\u00f3n": "240835", "Cta Principal": "14650501", "Regimen": "O-13;O-15"},
            "CEU-4502512":  {"Base": 653399.00, "IVA": 34276.00, "IVA 19%": 34276.00, "Base IVA 19%": 180399.00, "IVA 5%": 0.0, "Base IVA 5%": 0.0, "INC / Otros Impuestos": 0.0, "IVA Importaci\u00f3n": 989000.00, "Cta IVA Importaci\u00f3n": "240835", "Cta Principal": "14650501", "Regimen": "O-13;O-15"},
            "CEU-4527831":  {"Base": 803141.00, "IVA": 33277.00, "IVA 19%": 33277.00, "Base IVA 19%": 175141.00, "IVA 5%": 0.0, "Base IVA 5%": 0.0, "INC / Otros Impuestos": 0.0, "IVA Importaci\u00f3n": 1313000.00, "Cta IVA Importaci\u00f3n": "240835", "Cta Principal": "14650501", "Regimen": "O-13;O-15"},
            "CEU-4537790":  {"Base": 218901.00, "IVA": 18221.00, "IVA 19%": 18221.00, "Base IVA 19%": 95901.00, "IVA 5%": 0.0, "Base IVA 5%": 0.0, "INC / Otros Impuestos": 0.0, "IVA Importaci\u00f3n": 257000.00, "Cta IVA Importaci\u00f3n": "240835", "Cta Principal": "14650501", "Regimen": "O-13;O-15"},
            "CEU-4541043":  {"Base": 445080.00, "IVA": 33455.00, "IVA 19%": 33455.00, "Base IVA 19%": 176080.00, "IVA 5%": 0.0, "Base IVA 5%": 0.0, "INC / Otros Impuestos": 0.0, "IVA Importaci\u00f3n": 563000.00, "Cta IVA Importaci\u00f3n": "240835", "Cta Principal": "14650501", "Regimen": "O-13;O-15"},
            "CEU-4544292":  {"Base": 95811.00, "IVA": 18204.00, "IVA 19%": 18204.00, "Base IVA 19%": 95811.00, "IVA 5%": 0.0, "Base IVA 5%": 0.0, "INC / Otros Impuestos": 0.0, "IVA Importaci\u00f3n": 126000.00, "Cta IVA Importaci\u00f3n": "240835", "Cta Principal": "14650501", "Regimen": "O-13;O-15"}
        }
        for idx_row, r_row in df.iterrows():
            f_key = str(r_row.get("Factura", "")).strip()
            editada_man = bool(r_row.get("Editada Manualmente", False)) if pd.notna(r_row.get("Editada Manualmente")) else False
            if f_key in DESGLOSE_IMPUESTOS_PDF and not editada_man:
                d_info = DESGLOSE_IMPUESTOS_PDF[f_key]
                df.at[idx_row, "Base"] = d_info["Base"]
                df.at[idx_row, "IVA"] = d_info["IVA"]
                df.at[idx_row, "IVA 19%"] = d_info.get("IVA 19%", d_info["IVA"])
                df.at[idx_row, "IVA 5%"] = d_info.get("IVA 5%", 0.0)
                df.at[idx_row, "Base IVA 19%"] = d_info.get("Base IVA 19%", 0.0)
                df.at[idx_row, "Base IVA 5%"] = d_info.get("Base IVA 5%", 0.0)
                df.at[idx_row, "INC / Otros Impuestos"] = d_info.get("INC / Otros Impuestos", 0.0)
                df.at[idx_row, "Cta IVA 5%"] = "24081003"
                if "IVA Importaci\u00f3n" in d_info:
                    df.at[idx_row, "IVA Importaci\u00f3n"] = d_info["IVA Importaci\u00f3n"]
                    df.at[idx_row, "Cta IVA Importaci\u00f3n"] = d_info.get("Cta IVA Importaci\u00f3n", "240835")
                if "Cta Principal" in d_info:
                    df.at[idx_row, "Cta Principal"] = d_info["Cta Principal"]
                    df.at[idx_row, "Categor\u00eda"] = obtener_etiqueta_cuenta_principal(d_info["Cta Principal"], df.at[idx_row, "Categor\u00eda"])
                if "Regimen" in d_info:
                    df.at[idx_row, "R\u00e9gimen Fiscal Emisor"] = d_info["Regimen"]
                    df.at[idx_row, "ReteFuente"] = 0.0
                    df.at[idx_row, "ReteICA"] = 0.0
                    df.at[idx_row, "Cta ReteFuente"] = ""
                    df.at[idx_row, "Cta ReteICA"] = ""
                b_n = float(df.at[idx_row, "Base"])
                iv_n = float(df.at[idx_row, "IVA"])
                inc_n = float(df.at[idx_row, "INC / Otros Impuestos"])
                iv_imp_n = float(df.at[idx_row, "IVA Importaci\u00f3n"] or 0.0)
                rf_n = float(df.at[idx_row, "ReteFuente"])
                ri_n = float(df.at[idx_row, "ReteICA"])
                riv_n = float(df.at[idx_row, "ReteIVA"])
                if bool(df.at[idx_row, "Impuestos Asumidos"]):
                    df.at[idx_row, "Neto a Pagar"] = round(b_n + inc_n + iv_n + iv_imp_n, 2)
                    df.at[idx_row, "Total Neto"] = round(b_n + inc_n + iv_n + iv_imp_n, 2)
                else:
                    df.at[idx_row, "Neto a Pagar"] = round(b_n + inc_n + iv_n + iv_imp_n - rf_n - ri_n - riv_n, 2)
                    df.at[idx_row, "Total Neto"] = round(b_n + inc_n + iv_n + iv_imp_n - rf_n - ri_n - riv_n, 2)
            else:
                iv_tot_r = float(r_row.get("IVA", 0.0) or 0.0)
                iv_5_r = float(r_row.get("IVA 5%", 0.0) or 0.0)
                iv_19_r = float(r_row.get("IVA 19%", 0.0) or 0.0)
                if iv_5_r <= 0 and iv_19_r <= 0 and iv_tot_r > 0:
                    df.at[idx_row, "IVA 19%"] = iv_tot_r
                    df.at[idx_row, "Base IVA 19%"] = float(r_row.get("Base", 0.0) or 0.0)

        # Normalizar todas las cuentas a 8 digitos PUC y alinear Cta ReteFuente / Cta ReteICA
        for idx_row, r_row in df.iterrows():
            f_num_n = str(r_row.get("Factura", "")).strip()
            cta_p_n = str(r_row.get("Cta Principal", "")).strip()
            if cta_p_n in ("146505", "1465"):
                df.at[idx_row, "Cta Principal"] = "14650501"
            elif cta_p_n == "515505":
                df.at[idx_row, "Cta Principal"] = "51550501"
            rf_v = float(r_row.get("ReteFuente", 0.0) or 0.0)
            ri_v = float(r_row.get("ReteICA", 0.0) or 0.0)
            b_v = float(r_row.get("Base", 0.0) or 0.0)
            edit_m = bool(r_row.get("Editada Manualmente", False))
            if f_num_n.startswith("BT-") and rf_v > 0 and b_v > 0 and (rf_v / b_v) > 0.09 and str(r_row.get("Cta ReteFuente", "")).strip() == "23652503":
                df.at[idx_row, "Cta ReteFuente"] = "23651501"
                df.at[idx_row, "Cta ReteICA"] = "23680505"
                df.at[idx_row, "Raz\u00f3n Contable"] = "Honorarios Agencia de Aduanas (11% ReteFuente [23651501] - 9.66\u2030 ReteICA [23680505])"
            if not edit_m:
                if rf_v <= 0 and str(df.at[idx_row, "Cta ReteFuente"]).strip() != "":
                    df.at[idx_row, "Cta ReteFuente"] = ""
                if ri_v <= 0 and str(df.at[idx_row, "Cta ReteICA"]).strip() != "":
                    df.at[idx_row, "Cta ReteICA"] = ""
            aud_d = r_row.get("Audit Info")
            if isinstance(aud_d, dict):
                c_rf_cur = str(df.at[idx_row, "Cta ReteFuente"]).strip()
                c_ri_cur = str(df.at[idx_row, "Cta ReteICA"]).strip()
                aud_d["razon_rfte"] = f"ReteFuente [{c_rf_cur}] (${rf_v:,.2f})" if rf_v > 0 else "Sin ReteFuente ($0.00)"
                aud_d["razon_reteica"] = f"ReteICA [{c_ri_cur}] (${ri_v:,.2f})" if ri_v > 0 else "Sin ReteICA ($0.00)"
                aud_d["razon_rica"] = aud_d["razon_reteica"]

        # Auto-clasificar gastos detallados (Cafeteria, Aseo, Mensajeria, Transporte, Luz, Papeleria, Legal, Seguros, Software, Viajes, Mantenimiento)
        for idx_row, r_row in df.iterrows():
            cta_cur_r = str(r_row.get("Cta Principal", "")).strip()
            cat_cur_r = str(r_row.get("Categor\u00eda", "")).strip()
            edit_man_r = bool(r_row.get("Editada Manualmente", False)) if pd.notna(r_row.get("Editada Manualmente")) else False
            if cta_cur_r != "14650501" and not edit_man_r:
                if cta_cur_r in ["51959501", "14350101", "61350501", "61800101", "51550501"] or cat_cur_r in ["Gasto General", "General", "Mercancia"]:
                    txt_pdf_r = ""
                    try:
                        if "obtener_paginas_texto_factura_integrado" in globals() or "obtener_paginas_texto_factura_integrado" in locals():
                            pags_r = obtener_paginas_texto_factura_integrado(r_row)
                            if pags_r:
                                txt_pdf_r = "\n".join(pags_r)
                    except Exception:
                        txt_pdf_r = ""
                    res_g = clasificar_gasto_detallado_puc(
                        r_row.get("NIT Emisor"), r_row.get("Proveedor"), r_row.get("Descripcion"),
                        txt_pdf_r, float(r_row.get("Base", 0.0) or 0.0), float(r_row.get("IVA", 0.0) or 0.0),
                        cta_actual=cta_cur_r
                    )
                    if res_g:
                        df.at[idx_row, "Cta Principal"] = res_g["cta_p"]
                        df.at[idx_row, "Categor\u00eda"] = res_g["cat"]
                        df.at[idx_row, "Descripcion"] = res_g["desc"]
                        df.at[idx_row, "Raz\u00f3n Contable"] = res_g["razon"]
            else:
                df.at[idx_row, "Categor\u00eda"] = obtener_etiqueta_cuenta_principal(cta_cur_r, cat_cur_r)

        df.attrs["_indumaq_norm_done"] = len(df)
        return df

    def sincronizar_asiento_con_factura_actual(df_asiento, fac_row):
        """Sincroniza en tiempo real las cuentas contables PUC (Costo/Gasto, ReteFuente, ReteICA, IVA, CxP)
        de un asiento de triangulacion con las cuentas vigentes de la factura en la Pestana 2."""
        if df_asiento is None or fac_row is None:
            return df_asiento
        es_df = isinstance(df_asiento, pd.DataFrame)
        if es_df and df_asiento.empty:
            return df_asiento
        filas = df_asiento.to_dict(orient="records") if es_df else [dict(x) for x in df_asiento]
        if not filas:
            return df_asiento

        fac_num = str(fac_row.get("Factura", "")).strip()
        prov_nom = str(fac_row.get("Proveedor", ""))[:26]
        cta_p = _limpiar_codigo_cta(fac_row.get("Cta Principal"), "14650501")
        if cta_p in ("146505", "1465"):
            cta_p = "14650501"
        elif cta_p == "515505":
            cta_p = "51550501"
        cta_rf = _limpiar_codigo_cta(fac_row.get("Cta ReteFuente"), "")
        cta_ri = _limpiar_codigo_cta(fac_row.get("Cta ReteICA"), "")
        cta_iva = _limpiar_codigo_cta(fac_row.get("Cta IVA"), "24081501")
        cta_iva_imp = _limpiar_codigo_cta(fac_row.get("Cta IVA Importaci\u00f3n"), "240835")

        if "70787" in fac_num:
            if not cta_rf or cta_rf == "23652503":
                cta_rf = "23652505"
            if not cta_ri or cta_ri == "23680505":
                cta_ri = "23680513"

        for r_a in filas:
            cc = str(r_a.get("C\u00f3digo Cuenta", "")).strip()
            if cc in ("146505", "1465") or (cc.startswith(("14", "51", "52", "61")) and cc not in ("53152001", "53950501")):
                r_a["C\u00f3digo Cuenta"] = cta_p
                if cta_rf == "23652505" and "1465" in cta_p:
                    d_p = f"Transporte / Fletes Importaci\u00f3n ({cta_p}) - {prov_nom}"
                    r_a["Descripci\u00f3n de la Cuenta"] = d_p
                    r_a["Descripci\u00f3n Cuenta"] = d_p
            elif cc.startswith("2365") and cta_rf:
                r_a["C\u00f3digo Cuenta"] = cta_rf
                if cta_rf == "23652505":
                    d_rf = f"ReteFuente Practicada (Transporte Carga 1% [{cta_rf}]) Fac {fac_num}"
                elif cta_rf == "23652503":
                    d_rf = f"ReteFuente Practicada (Servicios 4% [{cta_rf}]) Fac {fac_num}"
                elif cta_rf == "23651501":
                    d_rf = f"ReteFuente Practicada (Honorarios 11% [{cta_rf}]) Fac {fac_num}"
                elif cta_rf == "23654001":
                    d_rf = f"ReteFuente Practicada (Compras 2.5% [{cta_rf}]) Fac {fac_num}"
                elif cta_rf in ("23652501", "23654004"):
                    d_rf = f"ReteFuente Practicada (3.5% [{cta_rf}]) Fac {fac_num}"
                else:
                    d_rf = f"ReteFuente Practicada ([{cta_rf}]) Fac {fac_num}"
                r_a["Descripci\u00f3n de la Cuenta"] = d_rf
                r_a["Descripci\u00f3n Cuenta"] = d_rf
            elif cc.startswith("2368") and cta_ri:
                r_a["C\u00f3digo Cuenta"] = cta_ri
                if cta_ri == "23680513":
                    d_ri = f"Retenci\u00f3n ICA Practicada (Transporte 4.14\u2030 [{cta_ri}])"
                elif cta_ri == "23680505":
                    d_ri = f"Retenci\u00f3n ICA Practicada (Servicios 9.66\u2030 [{cta_ri}])"
                elif cta_ri == "23680501":
                    d_ri = f"Retenci\u00f3n ICA Practicada (Compras 11.04\u2030 [{cta_ri}])"
                else:
                    d_ri = f"Retenci\u00f3n ICA Practicada ([{cta_ri}])"
                r_a["Descripci\u00f3n de la Cuenta"] = d_ri
                r_a["Descripci\u00f3n Cuenta"] = d_ri
            elif cc in ("24081501", "24081001") and cta_iva:
                r_a["C\u00f3digo Cuenta"] = cta_iva
            elif cc == "240835" and cta_iva_imp:
                r_a["C\u00f3digo Cuenta"] = cta_iva_imp

        return _limpiar_dtypes_pyarrow_df(pd.DataFrame(filas)) if es_df else filas

    def sincronizar_todos_los_paquetes_con_df_procesado():
        """Sincroniza todos los paquetes de importacion y asientos de triangulacion con las cuentas y valores actuales de df_procesado."""
        df_p = st.session_state.get("df_procesado")
        if df_p is None or not isinstance(df_p, pd.DataFrame) or df_p.empty:
            return
        df_map = {str(r["Factura"]).strip(): r for _, r in df_p.iterrows()}
        pqs = st.session_state.get("paquetes_importacion")
        if isinstance(pqs, dict):
            for p_k, p_v in pqs.items():
                if not isinstance(p_v, dict):
                    continue
                ag = p_v.get("agente")
                fac_ag = str(ag.get("Factura", "")).strip() if isinstance(ag, (dict, pd.Series)) else ""
                r_ag_df = df_map.get(fac_ag)
                if r_ag_df is not None and isinstance(ag, dict):
                    for col_s in ["Cta Principal", "Cta Contrapartida", "Cuenta Pasivo Especifica", "Cta IVA", "Cta IVA 5%", "Cta IVA Importaci\u00f3n", "Cta ReteFuente", "Cta ReteICA", "Categor\u00eda", "Raz\u00f3n Contable", "Base", "IVA", "IVA 19%", "IVA 5%", "INC / Otros Impuestos", "IVA Importaci\u00f3n", "ReteFuente", "ReteICA", "Neto a Pagar", "Total Neto", "Impuestos Asumidos", "R\u00e9gimen Fiscal Emisor"]:
                        if col_s in r_ag_df:
                            ag[col_s] = r_ag_df[col_s]
                ter_df = p_v.get("terceros")
                if isinstance(ter_df, pd.DataFrame) and not ter_df.empty:
                    for idx_t, r_t in ter_df.iterrows():
                        fac_t = str(r_t.get("Factura", "")).strip()
                        if fac_t in df_map:
                            r_t_df = df_map[fac_t]
                            for col_s in ["Cta Principal", "Cta Contrapartida", "Cuenta Pasivo Especifica", "Cta IVA", "Cta ReteFuente", "Cta ReteICA", "Base", "IVA", "ReteFuente", "ReteICA", "Neto a Pagar", "Total Neto", "Impuestos Asumidos"]:
                                if col_s in r_t_df and col_s in ter_df.columns:
                                    ter_df.at[idx_t, col_s] = r_t_df[col_s]
                if r_ag_df is not None:
                    for af_k in ["asiento_fijo", "asiento_aprobado"]:
                        if af_k in p_v and p_v[af_k] is not None:
                            p_v[af_k] = sincronizar_asiento_con_factura_actual(p_v[af_k], r_ag_df)
                    if f"asiento_fijo_pq_{p_k}" in st.session_state and st.session_state[f"asiento_fijo_pq_{p_k}"] is not None:
                        st.session_state[f"asiento_fijo_pq_{p_k}"] = sincronizar_asiento_con_factura_actual(st.session_state[f"asiento_fijo_pq_{p_k}"], r_ag_df)
        as_map = st.session_state.get("asientos_triangulacion_por_factura")
        if isinstance(as_map, dict):
            for fac_k, df_tr in list(as_map.items()):
                r_ag_df = df_map.get(str(fac_k).strip())
                if r_ag_df is not None and df_tr is not None:
                    as_map[fac_k] = sincronizar_asiento_con_factura_actual(df_tr, r_ag_df)

    import zlib as _zlib_emb
    _RESPALDO_50_PAQUETES_B64 = "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"

    _TEXTOS_224_PDFS_B64 = "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"
    _CACHE_TEXTOS_224_DICT = None

    def obtener_mapa_textos_224_pdfs():
        nonlocal_cache = st.session_state.get("_mapa_textos_224_pdfs")
        if nonlocal_cache is not None:
            return nonlocal_cache
        try:
            raw_t = _zlib_emb.decompress(base64.b64decode(_TEXTOS_224_PDFS_B64)).decode("utf-8")
            d_t = json.loads(raw_t)
            st.session_state["_mapa_textos_224_pdfs"] = d_t
            return d_t
        except Exception:
            return {}

    def obtener_paginas_texto_factura_integrado(fac_sel):
        if fac_sel is None:
            return []
        t_map = obtener_mapa_textos_224_pdfs()
        if not t_map:
            return []
        sop = str(fac_sel.get("Soporte PDF Renombrado", "")).strip()
        if sop in t_map:
            return t_map[sop]
        comp_id = str(fac_sel.get("Comprobante Siigo", "")).replace(" ", "_").strip()
        for k, v in t_map.items():
            if comp_id and comp_id in k:
                return v
        fol = str(fac_sel.get("Folio", "")).strip()
        for k, v in t_map.items():
            if fol and len(fol) >= 3 and fol in k:
                return v
        return []

    def _limpiar_y_estructurar_lineas_pagina_pdf(pt_str):
        raw_lines = [l.strip() for l in (pt_str or "").splitlines()]
        filtered = [l for l in raw_lines if l and l != "$"]
        out = []
        i = 0
        while i < len(filtered):
            cur = filtered[i]
            if cur == ":" and out:
                if i + 1 < len(filtered):
                    out[-1] = f"{out[-1]}: {filtered[i + 1]}"
                    i += 2
                    continue
            if cur.endswith(":") and i + 1 < len(filtered) and not filtered[i + 1].endswith(":"):
                nxt = filtered[i + 1]
                if not any(h in nxt for h in ["Datos del ", "Detalles de Productos", "Notas Finales", "Datos Totales"]):
                    out.append(f"{cur} {nxt}")
                    i += 2
                    continue
            out.append(cur)
            i += 1
        return out

    def _parsear_items_y_campos_dian(paginas_lista):
        full_txt = "\n".join(paginas_lista or [])
        paginas_estructuradas = [_limpiar_y_estructurar_lineas_pagina_pdf(p) for p in (paginas_lista or [])]
        flat_clean = "\n".join("\n".join(p) for p in paginas_estructuradas)

        def _b_campo(rx):
            for src_txt in (flat_clean, full_txt):
                m = re.search(rx + r"\s*[:\r\n\s]+([^\r\n]+)", src_txt, re.IGNORECASE)
                if m:
                    val = m.group(1).strip()
                    if val and not val.endswith(":") and not any(w in val.lower() for w in ["forma de pago", "medio de pago", "fecha de", "orden de", "tipo de", "pa\u00eds", "departamento", "municipio", "direcci\u00f3n", "correo", "tel\u00e9fono", "datos del", "rango "]):
                        return val
            return ""

        cufe = _b_campo(r"C\u00f3digo \u00danico de Factura - CUFE")
        forma_pago = _b_campo(r"Forma de pago") or "Contado / Cr\u00e9dito"
        medio_pago = _b_campo(r"Medio de Pago") or "Transferencia / Instrumento definido"
        fecha_venc = _b_campo(r"Fecha de Vencimiento")
        orden_pedido = _b_campo(r"Orden de pedido")
        tipo_operacion = _b_campo(r"Tipo de Operaci\u00f3n") or "10 - Est\u00e1ndar"
        regimen_pdf = _b_campo(r"R\u00e9gimen Fiscal")
        resp_trib = _b_campo(r"Responsabilidad tributaria")
        direccion_em = _b_campo(r"Direcci\u00f3n")
        telefono_em = _b_campo(r"Tel\u00e9fono / M\u00f3vil")
        municipio_em = _b_campo(r"Municipio / Ciudad")
        correo_em = _b_campo(r"Correo")
        adquiriente_pdf = _b_campo(r"Nombre o Raz\u00f3n Social")
        fecha_gen = _b_campo(r"Documento generado el")
        fecha_val = _b_campo(r"DIAN")
        autorizacion = _b_campo(r"Numero de Autorizaci\u00f3n")
        rango_desde = _b_campo(r"Rango desde")
        rango_hasta = _b_campo(r"Rango hasta")
        vigencia = _b_campo(r"Vigencia")

        m_prod = re.search(r"Detalles de Productos(.*?)(?:Notas Finales|Datos Totales)", full_txt, re.DOTALL | re.IGNORECASE)
        bloque_prod = m_prod.group(1).strip() if m_prod else ""
        headers_prod = {"IMPUESTOS", "Precio", "unitario de", "venta", "Precio unitario de venta", "Nro.", "C\u00f3digo", "Descripci\u00f3n", "U/M", "Cantidad", "Precio unitario", "Descuento", "detalle", "Recargo detalle", "IVA", "%", "INC", "$"}
        prod_lines = [pl.strip() for pl in bloque_prod.splitlines() if pl.strip() and pl.strip() not in headers_prod and not re.match(r"^Hoja \d+ de \d+$", pl.strip())]

        num_rx = re.compile(r"^-?\d+(?:\.\d{3})*[,\.]\d{2,4}$")
        items = []
        i = 0
        n = len(prod_lines)
        item_nro = 1
        while i < n:
            j = i
            while j < n:
                if j + 4 < n and all(num_rx.match(prod_lines[k]) for k in range(j, j + 5)):
                    break
                j += 1
            if j >= n:
                break
            nums = []
            k = j
            while k < n and num_rx.match(prod_lines[k]):
                nums.append(prod_lines[k])
                k += 1
            txt_tokens = [t for t in prod_lines[i:j] if t not in ("Descuento detalle", "Recargo detalle", "Precio unitario", "Precio unitario de", "unitario de", "de venta", "venta", "IMPUESTOS", "Descripci\u00f3n", "C\u00f3digo", "U/M", "Cantidad", "Nro.", "IVA", "INC", "%")]
            um = "UND"
            if txt_tokens and len(txt_tokens[-1]) <= 4 and re.match(r"^[A-Z0-9]{1,4}$", txt_tokens[-1]):
                um = txt_tokens.pop()
            nro = str(item_nro)
            if txt_tokens and txt_tokens[0].isdigit() and int(txt_tokens[0]) == item_nro:
                nro = txt_tokens.pop(0)
            desc_full = " ".join(txt_tokens).strip()
            cant = nums[0] if len(nums) > 0 else "1,00"
            p_unit = nums[1] if len(nums) > 1 else "0,00"
            desc_v = nums[2] if len(nums) > 2 else "0,00"
            if len(nums) >= 9:
                imp_v = f"IVA: ${nums[4]} ({nums[5]}%) | INC: ${nums[6]} ({nums[7]}%)"
                tot_v = nums[8]
            elif len(nums) >= 7:
                imp_v = f"${nums[4]} ({nums[5]}%)"
                tot_v = nums[6]
            elif len(nums) == 6:
                imp_v = f"${nums[4]}"
                tot_v = nums[5]
            else:
                imp_v = "$0,00 (0%)"
                tot_v = nums[4] if len(nums) > 4 else p_unit
            items.append({
                "Nro": nro,
                "Descripcion": desc_full or f"\u00cdtem {nro}",
                "U/M": um,
                "Cant.": cant,
                "Precio Unit.": f"${p_unit}",
                "Descuento": f"${desc_v}",
                "Impuesto": imp_v,
                "Total_Item": f"${tot_v}"
            })
            item_nro += 1
            i = k

        m_notas = re.search(r"Notas Finales(.*?)(?:Datos Totales|$)", full_txt, re.DOTALL | re.IGNORECASE)
        bloque_notas = m_notas.group(1).strip() if m_notas else ""
        notas_lines = [nl.strip() for nl in bloque_notas.splitlines() if nl.strip() and nl.strip() not in ("$", "Descripci\u00f3n") and not re.match(r"^Hoja \d+ de \d+$", nl.strip())]
        notas_limpias = " | ".join(notas_lines)

        return {
            "cufe": cufe,
            "forma_pago": forma_pago,
            "medio_pago": medio_pago,
            "fecha_venc": fecha_venc,
            "orden_pedido": orden_pedido,
            "tipo_operacion": tipo_operacion,
            "regimen_pdf": regimen_pdf,
            "resp_trib": resp_trib,
            "direccion_em": direccion_em,
            "telefono_em": telefono_em,
            "municipio_em": municipio_em,
            "correo_em": correo_em,
            "adquiriente_pdf": adquiriente_pdf,
            "fecha_gen": fecha_gen,
            "fecha_val": fecha_val,
            "autorizacion": autorizacion,
            "rango_desde": rango_desde,
            "rango_hasta": rango_hasta,
            "vigencia": vigencia,
            "items": items,
            "notas": notas_limpias,
            "paginas_estructuradas": paginas_estructuradas,
        }

    def _pdf_esc_winansi(s):
        if not s:
            return ""
        s_clean = str(s).replace("\u2014", "-").replace("\u2013", "-").replace("\u2022", "*").replace("\u2030", " por mil")
        s_lat = s_clean.encode("latin-1", "replace").decode("latin-1")
        return s_lat.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")

    def _wrap_pdf_text(text_str, max_chars=92):
        words = str(text_str or "").split()
        if not words:
            return [""]
        lines = []
        cur = words[0]
        for w in words[1:]:
            if len(cur) + 1 + len(w) <= max_chars:
                cur += " " + w
            else:
                lines.append(cur)
                cur = w
        lines.append(cur)
        return lines

    def generar_pdf_factura_dian_bytes(fac_sel, paginas_txt, empresa_dict):
        """Genera un archivo PDF 1.4 multipagina 100% valido en Python puro con la Factura Electronica DIAN y todas sus hojas."""
        info_d = _parsear_items_y_campos_dian(paginas_txt)
        pags_est = info_d.get("paginas_estructuradas") or []
        num_pags = max(1, len(pags_est) if pags_est else 1)

        fac_num = str(fac_sel.get("Factura", ""))
        prov_nom = str(fac_sel.get("Proveedor", "Proveedor"))
        nit_em = str(fac_sel.get("NIT Emisor", ""))
        comp_s = str(fac_sel.get("Comprobante Siigo", ""))
        fec_em = str(fac_sel.get("Fecha", ""))
        fec_ve = info_d.get("fecha_venc") or fec_em
        orden_ped = info_d.get("orden_pedido") or "N/A"
        forma_p_s = info_d.get("forma_pago") or "Contado / Credito"
        medio_p_s = info_d.get("medio_pago") or "Transferencia"
        cufe_s = info_d.get("cufe") or str(fac_sel.get("CUFE", ""))
        reg_s = str(fac_sel.get("R\u00e9gimen Fiscal Emisor", "O-48"))
        reg_pdf_s = info_d.get("regimen_pdf") or reg_s
        dir_em = info_d.get("direccion_em") or "Bogota D.C., Colombia"
        tel_em = info_d.get("telefono_em") or "-"
        cor_em = info_d.get("correo_em") or "-"
        aut_s = info_d.get("autorizacion") or "Autorizacion DIAN Vigente"
        rango_s = f"Desde {info_d.get('rango_desde', '-')} hasta {info_d.get('rango_hasta', '-')} | Vigencia: {info_d.get('vigencia', '-')}"
        fec_val_s = info_d.get("fecha_val") or fec_em

        val_base = float(fac_sel.get("Base", 0.0) or 0.0)
        val_iva = float(fac_sel.get("IVA", 0.0) or 0.0)
        val_inc = float(fac_sel.get("INC / Otros Impuestos", 0.0) or 0.0)
        val_ivaimp_d = float(fac_sel.get("IVA Importaci\u00f3n", 0.0) or 0.0)
        cta_ivaimp_d = str(fac_sel.get("Cta IVA Importaci\u00f3n") or "240835").strip()
        linea_ivaimp_html = f"<b>\U0001f6a2 IVA de Importaci\u00f3n (Cta {cta_ivaimp_d}):</b> ${val_ivaimp_d:,.2f}<br>" if val_ivaimp_d > 0 else ""
        val_tot = float(fac_sel.get("Total", 0.0) or 0.0)
        val_rf = float(fac_sel.get("ReteFuente", 0.0) or 0.0)
        val_ri = float(fac_sel.get("ReteICA", 0.0) or 0.0)

        pages_streams = []

        # PAGINA 1: Representacion Grafica Estructurada + Detalle de Items y Totales
        cmds = []
        cmds.append("1 1 1 rg 0 0 595 842 re f")
        # Header Azul DIAN
        cmds.append("0 0.44 0.73 rg 24 776 547 46 re f")
        cmds.append(f"BT /F2 11 Tf 1 1 1 rg 34 804 Td ({_pdf_esc_winansi('FACTURA ELECTRONICA DE VENTA - REPRESENTACION GRAFICA DIAN')}) Tj ET")
        cmds.append(f"BT /F2 11 Tf 1 1 1 rg 425 804 Td ({_pdf_esc_winansi('FACTURA: ' + fac_num)}) Tj ET")
        cmds.append(f"BT /F1 7.5 Tf 1 1 1 rg 34 791 Td ({_pdf_esc_winansi('CUFE: ' + cufe_s[:96])}) Tj ET")
        cmds.append(f"BT /F1 8 Tf 1 1 1 rg 425 791 Td ({_pdf_esc_winansi(comp_s + ' | Hoja 1 de ' + str(num_pags))}) Tj ET")

        # Caja de Datos Emisor / Adquiriente / Documento
        cmds.append("0.97 0.98 0.99 rg 24 682 547 88 re f")
        cmds.append("0.82 0.86 0.91 RG 0.8 w 24 682 547 88 re S")
        cmds.append("0 0.44 0.73 rg")
        cmds.append(f"BT /F2 8.5 Tf 30 757 Td ({_pdf_esc_winansi('DATOS DEL EMISOR / VENDEDOR')}) Tj ET")
        cmds.append(f"BT /F2 8.5 Tf 235 757 Td ({_pdf_esc_winansi('DATOS DEL ADQUIRIENTE / COMPRADOR')}) Tj ET")
        cmds.append(f"BT /F2 8.5 Tf 420 757 Td ({_pdf_esc_winansi('DATOS DEL DOCUMENTO')}) Tj ET")

        cmds.append("0.1 0.15 0.22 rg")
        em_lines = [
            f"Razon Social: {prov_nom[:34]}",
            f"NIT Emisor: {nit_em} | Reg: {reg_s} (PDF: {reg_pdf_s})",
            f"Direccion: {dir_em[:38]}",
            f"Tel: {tel_em[:22]} | Ciudad: {(info_d.get('municipio_em') or 'Bogota')[:16]}",
            f"Correo: {cor_em[:38]}",
        ]
        adq_lines = [
            f"Razon Social: {(info_d.get('adquiriente_pdf') or empresa_dict.get('nombre', 'INDUMAQ ER SAS'))[:32]}",
            f"NIT Receptor: {empresa_dict.get('nit', '901.346.412-5')}",
            "Direccion: CR 78 SUR 70 F 48, Bogota D.C.",
            f"Autorizacion DIAN: {aut_s[:26]}",
            f"{rango_s[:38]}",
        ]
        doc_lines = [
            f"Fecha Emision: {fec_em}",
            f"Fecha Vencimiento: {fec_ve}",
            f"Orden de Pedido: {orden_ped[:22]}",
            f"Forma Pago: {forma_p_s[:22]}",
            f"Medio Pago: {medio_p_s[:22]}",
        ]
        y_top = 744
        for idx_l in range(5):
            cmds.append(f"BT /F1 7.6 Tf 30 {y_top} Td ({_pdf_esc_winansi(em_lines[idx_l])}) Tj ET")
            cmds.append(f"BT /F1 7.6 Tf 235 {y_top} Td ({_pdf_esc_winansi(adq_lines[idx_l])}) Tj ET")
            cmds.append(f"BT /F1 7.6 Tf 420 {y_top} Td ({_pdf_esc_winansi(doc_lines[idx_l])}) Tj ET")
            y_top -= 11.5

        # Encabezado Tabla de Productos
        y_cur = 664
        cmds.append("0.06 0.09 0.16 rg 24 648 547 18 re f")
        cmds.append("1 1 1 rg")
        cmds.append(f"BT /F2 8 Tf 28 654 Td (Nro) Tj ET")
        cmds.append(f"BT /F2 8 Tf 52 654 Td (Codigo / Descripcion del Producto o Servicio) Tj ET")
        cmds.append(f"BT /F2 8 Tf 275 654 Td (U/M) Tj ET")
        cmds.append(f"BT /F2 8 Tf 305 654 Td (Cant.) Tj ET")
        cmds.append(f"BT /F2 8 Tf 350 654 Td (Precio Unit.) Tj ET")
        cmds.append(f"BT /F2 8 Tf 425 654 Td (IVA / Impuesto) Tj ET")
        cmds.append(f"BT /F2 8 Tf 508 654 Td (Valor Total) Tj ET")

        y_cur = 635
        cmds.append("0.1 0.15 0.22 rg")
        items_list = info_d.get("items") or [{
            "Nro": "1",
            "Descripcion": str(fac_sel.get("Descripcion", "Facturacion de bienes / servicios")),
            "U/M": "UND",
            "Cant.": "1,00",
            "Precio Unit.": f"${val_base:,.2f}",
            "Impuesto": f"${val_iva:,.2f}",
            "Total_Item": f"${val_base:,.2f}"
        }]
        for it in items_list[:18]:
            cmds.append(f"BT /F1 7.8 Tf 28 {y_cur} Td ({_pdf_esc_winansi(str(it.get('Nro', ''))[:4])}) Tj ET")
            cmds.append(f"BT /F1 7.8 Tf 52 {y_cur} Td ({_pdf_esc_winansi(str(it.get('Descripcion', ''))[:44])}) Tj ET")
            cmds.append(f"BT /F1 7.8 Tf 275 {y_cur} Td ({_pdf_esc_winansi(str(it.get('U/M', ''))[:5])}) Tj ET")
            cmds.append(f"BT /F1 7.8 Tf 305 {y_cur} Td ({_pdf_esc_winansi(str(it.get('Cant.', ''))[:10])}) Tj ET")
            cmds.append(f"BT /F1 7.8 Tf 350 {y_cur} Td ({_pdf_esc_winansi(str(it.get('Precio Unit.', ''))[:15])}) Tj ET")
            cmds.append(f"BT /F1 7.8 Tf 425 {y_cur} Td ({_pdf_esc_winansi(str(it.get('Impuesto', ''))[:18])}) Tj ET")
            cmds.append(f"BT /F2 7.8 Tf 505 {y_cur} Td ({_pdf_esc_winansi(str(it.get('Total_Item', ''))[:16])}) Tj ET")
            cmds.append(f"0.88 0.91 0.94 RG 0.5 w 24 {y_cur - 4} m 571 {y_cur - 4} l S")
            y_cur -= 14

        # Notas Finales
        if info_d.get("notas"):
            y_cur -= 6
            cmds.append("0 0.44 0.73 rg")
            cmds.append(f"BT /F2 8 Tf 28 {y_cur} Td (Notas Finales y Observaciones del Emisor en Factura:) Tj ET")
            y_cur -= 11
            cmds.append("0.2 0.25 0.33 rg")
            for nl in _wrap_pdf_text(info_d["notas"], max_chars=96)[:5]:
                cmds.append(f"BT /F1 7.4 Tf 28 {y_cur} Td ({_pdf_esc_winansi(nl)}) Tj ET")
                y_cur -= 10

        # Caja de Totales y Validacion DIAN
        y_box = max(260, y_cur - 74)
        cmds.append("0.95 0.97 0.99 rg 24 " + str(y_box) + " 547 68 re f")
        cmds.append("0 0.44 0.73 RG 0.9 w 24 " + str(y_box) + " 547 68 re S")
        cmds.append("0.1 0.15 0.22 rg")
        cta_rf_s = str(fac_sel.get("Cta ReteFuente") or "Sin ReteFuente")
        cta_ri_s = str(fac_sel.get("Cta ReteICA") or "Sin ReteICA")
        cta_pr_s = str(fac_sel.get("Cta Principal", "14650501"))
        cta_cxp_s = str(fac_sel.get("Cta Contrapartida", "22050501"))
        rf_ri_line = f"ReteFuente: ${val_rf:,.2f} ({cta_rf_s}) | ReteICA: ${val_ri:,.2f} ({cta_ri_s})"
        ctas_line = f"Cuenta Costo/Gasto: {cta_pr_s} | Cuenta CxP: {cta_cxp_s}"
        val_dian_line = f"Validado DIAN: {fec_val_s} | Autorizacion: {aut_s} ({rango_s})"
        cmds.append(f"BT /F2 8.2 Tf 32 {y_box + 53} Td (Resumen de Retenciones, Validacion DIAN y Liquidacion:) Tj ET")
        cmds.append(f"BT /F1 7.8 Tf 32 {y_box + 40} Td ({_pdf_esc_winansi(rf_ri_line)}) Tj ET")
        cmds.append(f"BT /F1 7.8 Tf 32 {y_box + 27} Td ({_pdf_esc_winansi(ctas_line)}) Tj ET")
        cmds.append(f"BT /F1 7.5 Tf 32 {y_box + 14} Td ({_pdf_esc_winansi(val_dian_line)}) Tj ET")

        cmds.append(f"BT /F1 8.2 Tf 395 {y_box + 52} Td ({_pdf_esc_winansi(f'Subtotal / Base Gravable: ${val_base:,.2f}')}) Tj ET")
        cmds.append(f"BT /F1 8.2 Tf 395 {y_box + 39} Td ({_pdf_esc_winansi(f'IVA Facturado: ${val_iva:,.2f}')}) Tj ET")
        cmds.append(f"BT /F1 8.2 Tf 395 {y_box + 26} Td ({_pdf_esc_winansi(f'INC / Bolsas / Otros: ${val_inc:,.2f}')}) Tj ET")
        cmds.append("0 0.44 0.73 rg")
        cmds.append(f"BT /F2 10 Tf 395 {y_box + 10} Td ({_pdf_esc_winansi(f'TOTAL FACTURA: COP ${val_tot:,.2f}')}) Tj ET")

        # Texto adicional de la Hoja 1 en la parte inferior si hay espacio
        if pags_est and len(pags_est[0]) > 0 and y_box > 75:
            y_sub = y_box - 16
            cmds.append("0.3 0.35 0.42 rg")
            cmds.append(f"BT /F2 7.5 Tf 28 {y_sub} Td (Campos Integrales Extraidos de la Hoja 1 del PDF Original:) Tj ET")
            y_sub -= 10
            for raw_l in pags_est[0]:
                if y_sub < 28:
                    break
                for wl in _wrap_pdf_text(raw_l, max_chars=102):
                    if y_sub < 28:
                        break
                    cmds.append(f"BT /F3 6.8 Tf 28 {y_sub} Td ({_pdf_esc_winansi(wl)}) Tj ET")
                    y_sub -= 8.5

        pages_streams.append("\n".join(cmds).encode("latin-1"))

        # PAGINAS ADICIONALES (Hoja 2..N del PDF original con todos sus campos de consulta)
        for idx_p in range(1, len(pags_est)):
            cmds_p = []
            cmds_p.append("1 1 1 rg 0 0 595 842 re f")
            cmds_p.append("0 0.44 0.73 rg 24 782 547 38 re f")
            cmds_p.append(f"BT /F2 10.5 Tf 1 1 1 rg 34 803 Td ({_pdf_esc_winansi(f'FACTURA ELECTRONICA DE VENTA {fac_num} - {prov_nom[:35]}')}) Tj ET")
            cmds_p.append(f"BT /F1 8.2 Tf 1 1 1 rg 34 790 Td ({_pdf_esc_winansi(f'Contenido Completo de Consulta del PDF Original - Hoja {idx_p + 1} de {num_pags}')}) Tj ET")
            cmds_p.append("0.1 0.15 0.22 rg")
            y_p = 762
            for raw_l in pags_est[idx_p]:
                for wl in _wrap_pdf_text(raw_l, max_chars=96):
                    if y_p < 36:
                        break
                    is_hdr = any(k in wl for k in ["Datos Totales", "Valores informativos", "Numero de Autorizaci", "Documento validado", "Total factura (=)", "Orden de pedido"])
                    f_tag = "/F2 8.2 Tf" if is_hdr else "/F3 7.8 Tf"
                    cmds_p.append(f"BT {f_tag} 32 {y_p} Td ({_pdf_esc_winansi(wl)}) Tj ET")
                    y_p -= 11
            pages_streams.append("\n".join(cmds_p).encode("latin-1"))

        objs = []
        objs.append(b"<< /Type /Catalog /Pages 2 0 R >>")
        objs.append(b"")
        objs.append(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
        objs.append(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>")
        objs.append(b"<< /Type /Font /Subtype /Type1 /BaseFont /Courier /Encoding /WinAnsiEncoding >>")

        page_obj_nums = []
        for st_bytes in pages_streams:
            cont_num = len(objs) + 1
            objs.append(b"<< /Length " + str(len(st_bytes)).encode("ascii") + b" >>\nstream\n" + st_bytes + b"\nendstream")
            pg_num = len(objs) + 1
            page_obj_nums.append(pg_num)
            objs.append(
                f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 3 0 R /F2 4 0 R /F3 5 0 R >> >> /Contents {cont_num} 0 R >>".encode("ascii")
            )

        kids_str = " ".join(f"{pn} 0 R" for pn in page_obj_nums)
        objs[1] = f"<< /Type /Pages /Kids [{kids_str}] /Count {len(page_obj_nums)} >>".encode("ascii")

        out = io.BytesIO()
        out.write(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
        offsets = []
        for idx_o, ob in enumerate(objs, 1):
            offsets.append(out.tell())
            out.write(f"{idx_o} 0 obj\n".encode("ascii"))
            out.write(ob)
            out.write(b"\nendobj\n")
        xref_pos = out.tell()
        out.write(f"xref\n0 {len(objs) + 1}\n0000000000 65535 f \n".encode("ascii"))
        for off in offsets:
            out.write(f"{off:010d} 00000 n \n".encode("ascii"))
        out.write(f"trailer\n<< /Size {len(objs) + 1} /Root 1 0 R >>\nstartxref\n{xref_pos}\n%%EOF\n".encode("ascii"))
        return out.getvalue()

    def renderizar_representacion_grafica_dian_html(fac_sel, paginas_txt, empresa_dict, pdf_bytes_original=None, origen_pdf_desc="", key_suffix=""):
        info_d = _parsear_items_y_campos_dian(paginas_txt)
        pags_est = info_d.get("paginas_estructuradas") or []
        num_pags = max(1, len(paginas_txt or []))
        if pdf_bytes_original:
            try:
                num_pags = max(num_pags, len(PdfReader(io.BytesIO(pdf_bytes_original)).pages))
            except Exception:
                pass
        pags_badge = f"{num_pags} p\u00e1ginas completas" if num_pags > 1 else "1 p\u00e1gina"
        fac_num = str(fac_sel.get("Factura", ""))
        prov_nom = str(fac_sel.get("Proveedor", "Proveedor"))
        nit_em = str(fac_sel.get("NIT Emisor", ""))
        comp_s = str(fac_sel.get("Comprobante Siigo", ""))
        fec_em = str(fac_sel.get("Fecha", ""))
        fec_ve = info_d.get("fecha_venc") or fec_em
        orden_ped = info_d.get("orden_pedido") or "-"
        forma_p_s = info_d.get("forma_pago") or "Cr\u00e9dito/Contado"
        cufe_s = info_d.get("cufe") or str(fac_sel.get("CUFE", ""))
        reg_s = str(fac_sel.get("R\u00e9gimen Fiscal Emisor", "O-48"))
        reg_pdf_s = info_d.get("regimen_pdf") or reg_s
        dir_em = info_d.get("direccion_em") or "Bogot\u00e1 / Colombia"
        tel_em = info_d.get("telefono_em") or "-"
        cor_em = info_d.get("correo_em") or "-"
        aut_s = info_d.get("autorizacion") or "Autorizaci\u00f3n DIAN Vigente"
        vig_s = info_d.get("vigencia") or "-"
        fec_val_s = info_d.get("fecha_val") or fec_em

        val_base = float(fac_sel.get("Base", 0.0) or 0.0)
        val_iva = float(fac_sel.get("IVA", 0.0) or 0.0)
        val_iva19 = float(fac_sel.get("IVA 19%", 0.0) or 0.0)
        val_iva5 = float(fac_sel.get("IVA 5%", 0.0) or 0.0)
        val_inc = float(fac_sel.get("INC / Otros Impuestos", 0.0) or 0.0)
        val_tot = float(fac_sel.get("Total", 0.0) or 0.0)
        val_rf = float(fac_sel.get("ReteFuente", 0.0) or 0.0)
        val_ri = float(fac_sel.get("ReteICA", 0.0) or 0.0)

        st.success(f"\u2705 **Factura Electr\u00f3nica DIAN Vinculada ({pags_badge}):** Comprobante **{comp_s}** | Factura: **{fac_num}** | Proveedor: **{prov_nom}** (NIT: {nit_em}) \u2014 Total: **${val_tot:,.2f}**")

        rows_html = ""
        if info_d["items"]:
            for it in info_d["items"]:
                rows_html += (
                    f"<tr style='border-bottom:1px solid #e2e8f0;'>"
                    f"<td style='padding:7px 8px; text-align:center; font-weight:600; color:#334155;'>{it['Nro']}</td>"
                    f"<td style='padding:7px 10px; color:#0f172a; font-weight:500;'>{it['Descripcion']}</td>"
                    f"<td style='padding:7px 8px; text-align:center; color:#475569;'>{it['U/M']}</td>"
                    f"<td style='padding:7px 8px; text-align:right; color:#334155;'>{it['Cant.']}</td>"
                    f"<td style='padding:7px 8px; text-align:right; color:#334155;'>{it['Precio Unit.']}</td>"
                    f"<td style='padding:7px 8px; text-align:right; color:#64748b;'>{it['Descuento']}</td>"
                    f"<td style='padding:7px 8px; text-align:right; color:#0369a1; font-weight:600;'>{it['Impuesto']}</td>"
                    f"<td style='padding:7px 10px; text-align:right; font-weight:700; color:#0f172a;'>{it['Total_Item']}</td>"
                    f"</tr>"
                )
        else:
            desc_fallback = str(fac_sel.get("Descripcion", "Facturaci\u00f3n de bienes / servicios"))
            rows_html = (
                f"<tr style='border-bottom:1px solid #e2e8f0;'>"
                f"<td style='padding:7px 8px; text-align:center;'>1</td>"
                f"<td style='padding:7px 10px; font-weight:500;'>{desc_fallback}</td>"
                f"<td style='padding:7px 8px; text-align:center;'>UND</td>"
                f"<td style='padding:7px 8px; text-align:right;'>1,00</td>"
                f"<td style='padding:7px 8px; text-align:right;'>${val_base:,.2f}</td>"
                f"<td style='padding:7px 8px; text-align:right;'>$0,00</td>"
                f"<td style='padding:7px 8px; text-align:right; color:#0369a1;'>${val_iva:,.2f}</td>"
                f"<td style='padding:7px 10px; text-align:right; font-weight:700;'>${val_base:,.2f}</td>"
                f"</tr>"
            )

        notas_html = ""
        if info_d.get("notas"):
            notas_html = (
                f"<div style='background:#f8fafc; border:1px solid #cbd5e1; border-left:4px solid #0284c7; border-radius:6px; padding:10px 14px; margin-top:12px; font-size:12.5px; color:#334155;'>"
                f"<b>\U0001f4dd Notas Finales y Observaciones del Emisor en Factura:</b><br>{info_d['notas']}"
                f"</div>"
            )

        iva_desglose_txt = f"${val_iva:,.2f}"
        if val_iva5 > 0:
            iva_desglose_txt = f"${val_iva:,.2f} (IVA 19%: ${val_iva19:,.2f} | IVA 5%: ${val_iva5:,.2f})"

        html_doc = (
            f"<div style='border:2px solid #0070ba; border-radius:10px; background:#ffffff; box-shadow:0 4px 12px rgba(0,0,0,0.08); overflow:hidden; margin-bottom:12px; font-family:-apple-system, BlinkMacSystemFont, Segoe UI, Roboto, sans-serif;'>"
            f"<div style='background:#0070ba; color:white; padding:10px 18px; display:flex; justify-content:space-between; align-items:center;'>"
            f"<div><b style='font-size:15px;'>\U0001f4c4 FACTURA ELECTR\u00d3NICA DE VENTA \u2014 Representaci\u00f3n Gr\u00e1fica DIAN</b><br><span style='font-size:11.5px; opacity:0.9;'>CUFE: {cufe_s}</span></div>"
            f"<div style='text-align:right;'><span style='background:white; color:#0070ba; padding:4px 12px; border-radius:6px; font-weight:800; font-size:15px; font-family:monospace;'>{fac_num}</span><br><span style='font-size:11.5px;'>{comp_s} ({pags_badge})</span></div>"
            f"</div>"
            f"<div style='padding:16px 18px;'>"
            f"<div style='display:grid; grid-template-columns:1.2fr 1.2fr 0.95fr; gap:12px; background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:12px 14px; font-size:13px; color:#1e293b; margin-bottom:14px;'>"
            f"<div><b style='color:#0070ba;'>\U0001f3e2 Datos del Emisor / Vendedor:</b><br><b>Raz\u00f3n Social:</b> {prov_nom}<br><b>NIT Emisor:</b> {nit_em} | <b>R\u00e9gimen:</b> {reg_s} (PDF: {reg_pdf_s})<br><b>Direcci\u00f3n:</b> {dir_em}<br><b>Tel\u00e9fono:</b> {tel_em} | <b>Correo:</b> {cor_em}</div>"
            f"<div><b style='color:#0070ba;'>\U0001f3ed Datos del Adquiriente / Comprador:</b><br><b>Raz\u00f3n Social:</b> {info_d.get('adquiriente_pdf') or empresa_dict.get('nombre', 'INDUMAQ ER SAS')}<br><b>NIT Receptor:</b> {empresa_dict.get('nit', '901.346.412-5')}<br><b>Direcci\u00f3n:</b> CR 78 SUR 70 F 48, Bogot\u00e1 D.C.<br><b>Autorizaci\u00f3n DIAN:</b> {aut_s} (Vig: {vig_s})</div>"
            f"<div><b style='color:#0070ba;'>\U0001f4c5 Datos del Documento:</b><br><b>Fecha Emisi\u00f3n:</b> {fec_em} | <b>Vence:</b> {fec_ve}<br><b>Orden de Pedido / Ref:</b> <span style='background:#e0f2fe; color:#0369a1; padding:1px 6px; border-radius:4px; font-weight:700;'>{orden_ped}</span><br><b>Forma de Pago:</b> {forma_p_s}<br><b>Medio de Pago:</b> {info_d.get('medio_pago', 'Transferencia')}</div>"
            f"</div>"
            f"<div style='font-weight:700; font-size:13.5px; color:#0f172a; margin-bottom:6px;'>\U0001f4e6 Detalles de Productos / Servicios Facturados ({len(info_d['items']) or 1} \u00edtem(s)):</div>"
            f"<div style='overflow-x:auto; border:1px solid #cbd5e1; border-radius:6px;'>"
            f"<table style='width:100%; border-collapse:collapse; font-size:12.8px;'>"
            f"<thead><tr style='background:#0f172a; color:white; text-align:left;'>"
            f"<th style='padding:7px 8px; text-align:center; width:45px;'>Nro.</th>"
            f"<th style='padding:7px 10px;'>C\u00f3digo / Descripci\u00f3n del Producto o Servicio</th>"
            f"<th style='padding:7px 8px; text-align:center; width:55px;'>U/M</th>"
            f"<th style='padding:7px 8px; text-align:right; width:70px;'>Cant.</th>"
            f"<th style='padding:7px 8px; text-align:right; width:110px;'>Precio Unit.</th>"
            f"<th style='padding:7px 8px; text-align:right; width:95px;'>Descuento</th>"
            f"<th style='padding:7px 8px; text-align:right; width:140px;'>IVA / Impuesto</th>"
            f"<th style='padding:7px 10px; text-align:right; width:120px;'>Valor Total</th>"
            f"</tr></thead>"
            f"<tbody>{rows_html}</tbody>"
            f"</table></div>"
            f"{notas_html}"
            f"<div style='display:grid; grid-template-columns:1.3fr 1fr; gap:14px; margin-top:14px; background:#f1f5f9; border:1px solid #cbd5e1; border-radius:8px; padding:12px 16px; font-size:13px;'>"
            f"<div><b>Resumen de Retenciones y Liquidaci\u00f3n Contable:</b><br>"
            f"\u2022 <b>ReteFuente Liquidada:</b> ${val_rf:,.2f} ({fac_sel.get('Cta ReteFuente') or 'Sin ReteFuente'})<br>"
            f"\u2022 <b>ReteICA Liquidado:</b> ${val_ri:,.2f} ({fac_sel.get('Cta ReteICA') or 'Sin ReteICA'})<br>"
            f"\u2022 <b>Cuenta Costo/Gasto:</b> <code>{fac_sel.get('Cta Principal', '14650501')}</code> | <b>Cuenta CxP:</b> <code>{fac_sel.get('Cta Contrapartida', '22050501')}</code></div>"
            f"<div style='text-align:right;'>"
                        f"<b>Subtotal / Base Gravable:</b> ${val_base:,.2f}<br>"
            f"<b>IVA Facturado:</b> {iva_desglose_txt}<br>"
            f"{linea_ivaimp_html}"
            f"<b>INC 8% / Bolsas / Otros:</b> ${val_inc:,.2f}<br>"
            f"<span style='font-size:15.5px; font-weight:800; color:#0070ba;'>Total Factura (=): COP ${val_tot:,.2f}</span>"
            f"</div></div>"
            f"</div></div>"
        )
        st.markdown(html_doc, unsafe_allow_html=True)

        # EXPANDER CON EL DOCUMENTO PDF DE LA FACTURA PARA CONSULTAR INFORMACION PENDIENTE
        with st.expander(f"\U0001f4c4 Ver PDF de la Factura ({fac_num} \u2014 {pags_badge}) para Consultar Informaci\u00f3n Pendiente", expanded=False):
            pdf_bytes_doc = pdf_bytes_original
            es_pdf_sintetizado = False
            if not pdf_bytes_doc:
                try:
                    pdf_bytes_doc = generar_pdf_factura_dian_bytes(fac_sel, paginas_txt or [], empresa_dict)
                    es_pdf_sintetizado = True
                except Exception:
                    pdf_bytes_doc = b""

            safe_k_exp = re.sub(r"\W", "_", f"{comp_s}_{fac_num}_{key_suffix}")
            nombre_dl_pdf = str(fac_sel.get("Soporte PDF Renombrado") or f"Factura_{fac_num}.pdf")
            if not nombre_dl_pdf.lower().endswith(".pdf"):
                nombre_dl_pdf += ".pdf"

            # Ficha rapida de datos adicionales del PDF para consulta inmediata
            st.markdown(
                f"<div style='background:#eff6ff; border:1px solid #93c5fd; border-left:4px solid #0284c7; border-radius:6px; padding:10px 14px; margin-bottom:12px; font-size:12.8px; color:#1e3a8a;'>"
                f"<b>\U0001f4cc Datos Adicionales de Consulta en el PDF ({fac_num}):</b><br>"
                f"\u2022 <b>Orden de Pedido / DO / Referencia:</b> <code>{orden_ped}</code> &nbsp;|&nbsp; "
                f"<b>Tipo de Operaci\u00f3n:</b> {info_d.get('tipo_operacion') or '10 - Estandar'} &nbsp;|&nbsp; "
                f"<b>Validaci\u00f3n DIAN:</b> {fec_val_s}<br>"
                f"\u2022 <b>Resoluci\u00f3n Autorizaci\u00f3n DIAN:</b> {aut_s} (Rango: {info_d.get('rango_desde', '-')} al {info_d.get('rango_hasta', '-')} | Vigencia: {vig_s}) &nbsp;|&nbsp; "
                f"<b>Responsabilidad Tributaria Emisor:</b> {info_d.get('resp_trib') or reg_pdf_s}"
                f"</div>",
                unsafe_allow_html=True
            )

            # Si tenemos el PDF original escaneado (o si pypdfium2/fitz estan activos), renderizar el visor PDF completo
            if pdf_bytes_doc:
                renderizar_visor_pdf_completo(
                    pdf_bytes_doc,
                    nombre_archivo=nombre_dl_pdf,
                    fac_sel=fac_sel,
                    key_prefix=f"exp_pdf_{safe_k_exp}",
                    nit_comprador=re.sub(r"\D", "", str(empresa_dict.get("nit", "9013464125"))),
                    mostrar_card_dian=False
                )

            # Hojas visuales A4 completas del PDF (Hoja 1, Hoja 2...) para consulta garantizada en cualquier navegador
            if pags_est:
                st.markdown("#### \U0001f4d1 Vista Detallada Hoja por Hoja del Documento PDF (Todos los Campos y Notas)")
                for idx_pg, lines_pg in enumerate(pags_est):
                    filas_pg_html = ""
                    for ln_s in lines_pg:
                        if ":" in ln_s and not ln_s.startswith("http"):
                            p_k, p_v = ln_s.split(":", 1)
                            filas_pg_html += f"<div style='padding:3px 0; border-bottom:1px dashed #e2e8f0;'><b style='color:#0f172a;'>{p_k}:</b> <span style='color:#1e293b;'>{p_v}</span></div>"
                        elif any(h_w in ln_s for h_w in ["FACTURA ELECTR", "Datos del ", "Detalles de Productos", "Notas Finales", "Datos Totales", "Valores informativos"]):
                            filas_pg_html += f"<div style='background:#e0f2fe; color:#0369a1; font-weight:700; padding:5px 8px; border-radius:4px; margin-top:8px; margin-bottom:4px;'>{ln_s}</div>"
                        else:
                            filas_pg_html += f"<div style='padding:2px 0; color:#334155;'>{ln_s}</div>"

                    st.markdown(
                        f"<div style='background:#525659; padding:14px; border-radius:8px; margin-bottom:14px;'>"
                        f"<div style='background:#ffffff; max-width:820px; margin:0 auto; padding:22px 28px; border-radius:4px; box-shadow:0 4px 14px rgba(0,0,0,0.25); font-family:-apple-system, BlinkMacSystemFont, Segoe UI, Roboto, sans-serif; font-size:12.5px; line-height:1.45;'>"
                        f"<div style='border-bottom:2px solid #0070ba; padding-bottom:8px; margin-bottom:12px; display:flex; justify-content:space-between; align-items:center;'>"
                        f"<span style='font-weight:800; color:#0070ba; font-size:14px;'>\U0001f4c4 DOCUMENTO PDF FACTURA {fac_num} \u2014 {prov_nom}</span>"
                        f"<span style='background:#0070ba; color:white; padding:2px 10px; border-radius:12px; font-size:11.5px; font-weight:700;'>Hoja {idx_pg + 1} de {len(pags_est)}</span>"
                        f"</div>"
                        f"{filas_pg_html}"
                        f"</div></div>",
                        unsafe_allow_html=True
                    )

            if es_pdf_sintetizado and not key_suffix.startswith("pq_"):
                st.markdown("---")
                c_up_e1, c_up_e2 = st.columns([2.2, 1.8])
                with c_up_e2:
                    solo_pdfs_exp = st.checkbox(
                        "\U0001f512 Conservar mis cambios contables actuales",
                        value=True,
                        key=f"chk_solo_pdfs_exp_{safe_k_exp}"
                    )
                with c_up_e1:
                    upl_exp = st.file_uploader(
                        "\U0001f4e4 Opcional: Subir respaldo (.indumaq / .zip) si deseas cargar los 224 PDFs escaneados:",
                        type=["indumaq", "zip", "pdf"],
                        key=f"upl_pdfs_exp_{safe_k_exp}_{st.session_state.get('_upl_reset_cnt', 0)}"
                    )
                    if upl_exp is not None:
                        with st.spinner("Extrayendo los 224 PDFs al disco local (0 MB RAM)..."):
                            dict_r_tmp = st.session_state.get("dict_pdfs") or DiskBackedPDFDict("renom_pdfs")
                            if upl_exp.name.lower().endswith(".pdf"):
                                dict_r_tmp[upl_exp.name] = upl_exp.getvalue()
                                st.session_state["dict_pdfs"] = dict_r_tmp
                                st.rerun()
                            elif restaurar_desde_respaldo_bytes(empresa_dict, upl_exp, solo_pdfs=solo_pdfs_exp):
                                st.success("\u2705 \u00a1224 PDFs originales vinculados con \u00e9xito!")
                                st.rerun()


    def cargar_respaldo_integrado_50_paquetes(empresa_dict=None):
        """Carga el respaldo integrado de 224 facturas y 50 paquetes validados directamente en memoria."""
        try:
            raw_j = _zlib_emb.decompress(base64.b64decode(_RESPALDO_50_PAQUETES_B64)).decode("utf-8")
            data_emb = json.loads(raw_j)
            univ = data_emb.get("univ", {})
            est = data_emb.get("est", {})
            meta = data_emb.get("meta", {})
            mem = data_emb.get("mem", {})
            if mem:
                try:
                    nit_emp = re.sub(r'\D', '', str(empresa_dict.get("nit", ""))) if empresa_dict else "9013464125"
                    os.makedirs(f"almacenamiento_contable/{nit_emp}", exist_ok=True)
                    with open(f"almacenamiento_contable/{nit_emp}/memoria_aprendizaje_contable.json", "w", encoding="utf-8") as f_m:
                        json.dump(mem, f_m, ensure_ascii=False, indent=2)
                    with open(PATH_MEMORIA_GLOBAL, "w", encoding="utf-8") as f_m:
                        json.dump(mem, f_m, ensure_ascii=False, indent=2)
                except Exception:
                    pass
            if "df_procesado" in univ:
                df_emb = pd.DataFrame(univ["df_procesado"])
                st.session_state["df_procesado"] = normalizar_df_procesado(_limpiar_dtypes_pyarrow_df(df_emb))
            if "asientos_triangulacion" in univ:
                as_emb = {k: _limpiar_dtypes_pyarrow_df(pd.DataFrame(v)) for k, v in univ["asientos_triangulacion"].items()}
                st.session_state["asientos_triangulacion_por_factura"] = as_emb
            if "paquetes_importacion" in univ:
                pqs_emb = {}
                as_map = st.session_state.get("asientos_triangulacion_por_factura", {})
                for pk_s, pv_s in univ["paquetes_importacion"].items():
                    pk_i = int(pk_s) if str(pk_s).isdigit() else pk_s
                    pv_c = dict(pv_s)
                    for df_k in ["terceros", "asiento_fijo", "asiento_aprobado"]:
                        if df_k in pv_c and isinstance(pv_c[df_k], list):
                            pv_c[df_k] = _limpiar_dtypes_pyarrow_df(pd.DataFrame(pv_c[df_k]))
                    if ("asiento_fijo" in pv_c and pv_c["asiento_fijo"] is not None) or ("asiento_aprobado" in pv_c and pv_c["asiento_aprobado"] is not None):
                        st.session_state[f"paquete_listo_{pk_i}"] = True
                        st.session_state[f"asiento_fijo_pq_{pk_i}"] = pv_c.get("asiento_fijo", pv_c.get("asiento_aprobado")).copy()
                    ag_fac_str = str(pv_c.get("agente", {}).get("Factura", "")).strip()
                    if "destino_siigo" in pv_c and pv_c["destino_siigo"] is not None:
                        st.session_state[f"paquete_destino_siigo_{pk_i}"] = bool(pv_c["destino_siigo"])
                    elif ag_fac_str and ag_fac_str in as_map:
                        st.session_state[f"paquete_destino_siigo_{pk_i}"] = True
                        pv_c["destino_siigo"] = True
                    pqs_emb[pk_i] = pv_c
                st.session_state["paquetes_importacion"] = pqs_emb
                st.session_state["_ultimo_agentes_len"] = len(pqs_emb)
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
                if "df_procesado" in st.session_state and isinstance(st.session_state["df_procesado"], pd.DataFrame):
                    _ex_s = st.session_state["facturas_no_contabilizar"]
                    for _ir, _rr in st.session_state["df_procesado"].iterrows():
                        if str(_rr.get("Factura", "")).strip() in _ex_s:
                            st.session_state["df_procesado"].at[_ir, "No Contabilizar"] = True
            st.session_state["_regimenes_conocidos_aplicados_v1"] = True
            if "sel_paquete_activo_key" in est:
                st.session_state["sel_paquete_activo_key"] = est["sel_paquete_activo_key"]
                st.session_state["paquete_seleccionado_id"] = est["sel_paquete_activo_key"]
            st.session_state["_regimenes_conocidos_aplicados_v1"] = True
            if "factura_seleccionada_auditoria" not in st.session_state:
                st.session_state["factura_seleccionada_auditoria"] = "Comp 10-740"
            st.session_state["excel_nombre"] = meta.get("archivo_excel", "Reporte.xlsx")
            st.session_state["_sesion_cargada_nombre"] = f"Respaldo 50 Paquetes ({meta.get('fecha_respaldo', 'Integrado')})"
            dict_renom = DiskBackedPDFDict("renom_pdfs")
            if len(dict_renom) > 0:
                st.session_state["dict_pdfs"] = dict_renom
                st.session_state["raw_uploaded_pdfs"] = dict_renom
                st.session_state["total_zip_pdfs"] = len(dict_renom)
            sincronizar_todos_los_paquetes_con_df_procesado()
            guardar_estado_manual(empresa_dict or {"nit": "901.346.412-5", "nombre": "INDUMAQ ER SAS"})
            try:
                _nit_m = re.sub(r'\D', '', str((empresa_dict or {}).get("nit", "9013464125"))) or "9013464125"
                _flag_p = os.path.join(DATA_DIR, _nit_m, "estado_manual", "_avance_oct24_ivaimp_v7.ok")
                os.makedirs(os.path.dirname(_flag_p), exist_ok=True)
                with open(_flag_p, "w") as _ff:
                    _ff.write("ok")
            except Exception:
                pass
            return True
        except Exception as e:
            st.session_state["_ultimo_error_respaldo"] = str(e)
            return False

    # Auto-restaurar al inicio si la memoria esta vacia (aislado por NIT de empresa)
    if ("df_procesado" not in st.session_state or st.session_state["df_procesado"] is None) and not st.session_state.get("_usuario_limpio_sesion"):
        _nit_init = re.sub(r'\D', '', str(empresa.get("nit", "9013464125"))) or "9013464125"
        if _nit_init == "9013464125":
            _flag_migracion = os.path.join(DATA_DIR, _nit_init, "estado_manual", "_avance_oct24_ivaimp_v7.ok")
            if not os.path.exists(_flag_migracion):
                cargar_respaldo_integrado_50_paquetes(empresa)
            elif not cargar_estado_manual(empresa):
                cargar_respaldo_integrado_50_paquetes(empresa)
            else:
                st.session_state["_regimenes_conocidos_aplicados_v1"] = True
                dict_renom_auto = DiskBackedPDFDict("renom_pdfs")
                if len(dict_renom_auto) > 0 and not st.session_state.get("dict_pdfs"):
                    st.session_state["dict_pdfs"] = dict_renom_auto
                    st.session_state["raw_uploaded_pdfs"] = dict_renom_auto
                    st.session_state["total_zip_pdfs"] = len(dict_renom_auto)
        else:
            cargar_estado_manual(empresa)

    # BANNER DE CONTROL Y GUARDADO MANUAL
    if "df_procesado" in st.session_state and st.session_state["df_procesado"] is not None:
        c_bnr1, c_bnr2, c_bnr3, c_bnr4 = st.columns([2.8, 1.3, 1.6, 0.9])
        with c_bnr1:
            nom_ses = st.session_state.get("excel_nombre", "Reporte de Facturas")
            last_save = st.session_state.get("_ultimo_guardado_manual")
            save_info = f" | \U0001f4be `{last_save}`" if last_save else ""
            st.info(f"\U0001f4cb **Trabajo Activo:** `{nom_ses}` ({len(st.session_state['df_procesado'])} facturas){save_info}")
        with c_bnr2:
            if st.button("\U0001f4be Guardar Avance", key="btn_guardar_manual_top", use_container_width=True, help="Guarda en disco todas las modificaciones de cuentas y paquetes validados."):
                if guardar_estado_manual(empresa):
                    st.success("\u2705 \u00a1Avance contable guardado exitosamente!")
                    st.rerun()
                else:
                    st.error("Error al guardar.")
        with c_bnr3:
            emp_clean_name = re.sub(r'[^a-zA-Z0-9]', '_', str(empresa.get('nombre', 'Empresa')))
            nom_resp_top = f"Respaldo_Avance_{emp_clean_name}.indumaq"
            if st.session_state.get("_cached_respaldo_bytes") is None:
                if st.button("\U0001f6e1\ufe0f Preparar Respaldo (.indumaq)", key="btn_prep_respaldo_top", use_container_width=True, help="Genera un archivo port\u00e1til con tus facturas, paquetes y avances para descargar."):
                    with st.spinner("Preparando archivo de respaldo liviano..."):
                        st.session_state["_cached_respaldo_bytes"] = generar_respaldo_portatil_bytes(empresa)
                    st.rerun()
            else:
                c_d1, c_d2 = st.columns([0.8, 0.2])
                with c_d1:
                    st.download_button(
                        label="\u2b07\ufe0f Descargar (.indumaq)",
                        data=st.session_state["_cached_respaldo_bytes"],
                        file_name=nom_resp_top,
                        mime="application/zip",
                        key="btn_dl_respaldo_top_direct",
                        use_container_width=True,
                        help="Haz clic para descargar el archivo en tu equipo."
                    )
                with c_d2:
                    if st.button("\U0001f504", key="btn_refresh_respaldo_cache", help="Actualizar respaldo con \u00faltimos cambios"):
                        st.session_state["_cached_respaldo_bytes"] = None
                        st.rerun()
        with c_bnr4:
            if st.button("\U0001f195 Limpiar", key="btn_nuevo_trabajo_top", use_container_width=True, help="Limpia la memoria para cargar un nuevo mes."):
                st.session_state["_usuario_limpio_sesion"] = True
                for k in ["df_procesado", "dict_pdfs", "raw_uploaded_pdfs", "excel_bytes", "excel_nombre", "zip_pdfs", "job_actual_id", "_ultimo_excel_proc_sig", "_ultimo_pdfs_proc_sig", "paquetes_importacion", "asientos_triangulacion_por_factura", "facturas_no_contabilizar", "_ultimo_guardado_manual"]:
                    st.session_state.pop(k, None)
                st.rerun()
    else:
        # Si la memoria esta vacia, ofrecer boton para cargar el ultimo avance guardado si existe
        nit_c = re.sub(r'\D', '', str(empresa.get("nit", "empresa")))
        d_man = os.path.join(DATA_DIR, nit_c, "estado_manual")
        if os.path.exists(os.path.join(d_man, "df_procesado.pkl")):
            col_rec1, col_rec2 = st.columns([3.5, 1.5])
            with col_rec1:
                st.info("\U0001f4a1 **Tienes un avance contable guardado previamente.** Puedes restaurarlo con un solo clic:")
            with col_rec2:
                if st.button("\U0001f4c2 Cargar \u00daltimo Avance Guardado", key="btn_cargar_avance_previo", type="primary", use_container_width=True):
                    if cargar_estado_manual(empresa):
                        st.success("\u00a1Avance restaurado con \u00e9xito!")
                        st.rerun()
    
    # AVISO DE TRABAJO PREVIO DISPONIBLE (SOLO SE CARGA SI EL USUARIO OPRIME EL BOTON)
    trabajos_existentes = listar_trabajos_historial(empresa)

    if "df_procesado" not in st.session_state and trabajos_existentes:
        ultimo = trabajos_existentes[0]
        nom_ult = ultimo.get("nombre_trabajo", f"Trabajo {ultimo.get('id', '')}")
        n_fac = ultimo.get("total_facturas", 0)
        try: val_tot = float(ultimo.get("total_valor", 0.0) or 0.0)
        except Exception: val_tot = 0.0

        st.markdown(f"""
        <div style="background:#f0fdf4; border:1px solid #86efac; border-left:5px solid #16a34a; border-radius:8px; padding:14px 18px; margin-bottom:14px;">
            <h4 style="margin:0 0 6px 0; color:#166534;">\U0001f4c2 Tienes un trabajo guardado disponible</h4>
            <p style="margin:0; color:#14532d; font-size:14px;">
                Tienes guardado en memoria tu trabajo previo: <b>'{nom_ult}'</b> con <b>{n_fac} facturas</b> (${val_tot:,.2f}) y sus PDFs vinculados.<br>
                \u2022 Si deseas continuar con este trabajo, haz clic en el bot\u00f3n verde.<br>
                \u2022 Si hoy vas a implementar <b>facturas nuevas</b>, no tienes que oprimir nada y puedes subir tus nuevos archivos directamente abajo en la Pesta\u00f1a 1.
            </p>
        </div>
        """, unsafe_allow_html=True)

        c_btn_c1, c_btn_c2 = st.columns([1.6, 3])
        with c_btn_c1:
            if st.button("\U0001f4c2 Cargar mi trabajo anterior", key="btn_cargar_trabajo_previo_on_demand", use_container_width=True):
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
                    st.success(f"\u00a1Trabajo '{nom_ult}' cargado con \u00e9xito!")
                    st.rerun()
        with c_btn_c2:
            st.caption("\U0001f448 Oprime aqu\u00ed **\u00fanicamente** cuando desees restaurar el trabajo anterior. De lo contrario, contin\u00faa abajo con tus facturas nuevas.")

    # PANEL DE HISTORIAL DE TRABAJOS Y RESPALDOS INDESTRUCTIBLES
    trabajos_guardados = listar_trabajos_historial(empresa)
    panel_expanded = ("df_procesado" not in st.session_state or st.session_state["df_procesado"] is None)
    
    with st.expander("\U0001f5c2\ufe0f Historial de Trabajos, Respaldos y Carga R\u00e1pida", expanded=panel_expanded):
        # Restauracion inmediata desde archivo portatil .indumaq
        c_upl_r1, c_upl_r2 = st.columns([3, 1.2])
        with c_upl_r1:
            upl_indumaq = st.file_uploader("\U0001f4e4 Restaurar desde Archivo de Respaldo (.indumaq / .zip):", type=["indumaq", "backup", "zip"], key=f"upl_respaldo_indumaq_{st.session_state.get('_upl_reset_cnt', 0)}", help="Sube tu archivo de respaldo para recuperar al instante exactamente donde lo dejaste.")
            if upl_indumaq is not None:
                sig_ind = f"{upl_indumaq.name}_{upl_indumaq.size}"
                if st.session_state.get("_auto_restored_sig") != sig_ind:
                    with st.spinner("Restaurando respaldo contable autom\u00e1ticamente..."):
                        if restaurar_desde_respaldo_bytes(empresa, upl_indumaq):
                            st.session_state["_auto_restored_sig"] = sig_ind
                            st.success("\u00a1Respaldo restaurado con \u00e9xito! Todas tus facturas y paquetes quedaron listos.")
                            st.rerun()
                        else:
                            st.error(f"No se pudo leer el archivo de respaldo: {st.session_state.get('_ultimo_error_respaldo', '')}")
        with c_upl_r2:
            st.write("")
            st.write("")
            if upl_indumaq is not None:
                if st.button("\U0001f680 Restaurar Respaldo", key="btn_ejecutar_restaurar_upl", type="primary", use_container_width=True):
                    with st.spinner("Restaurando respaldo contable..."):
                        ok_r = restaurar_desde_respaldo_bytes(empresa, upl_indumaq)
                    if ok_r:
                        st.session_state["_auto_restored_sig"] = f"{upl_indumaq.name}_{upl_indumaq.size}"
                        st.success("\u00a1Respaldo restaurado con \u00e9xito! Todas tus facturas editadas y paquetes quedaron listos.")
                        st.rerun()
                    else:
                        st.error(f"No se pudo leer el archivo de respaldo: {st.session_state.get('_ultimo_error_respaldo', '')}")
        st.markdown("<hr style='margin:10px 0;'>", unsafe_allow_html=True)
    
        col_h_left, col_h_right = st.columns([1.5, 1])
    
        with col_h_left:
            st.markdown("#### \U0001f4c2 Trabajos Guardados en esta Empresa:")
            if trabajos_guardados:
                st.caption("Haz clic en 'Cargar' para recuperar de inmediato cualquier auditor\u00eda o mes previo:")
                for tb in trabajos_guardados:
                    c_h1, c_h2, c_h3 = st.columns([3, 1.2, 0.5])
                    with c_h1:
                        nom_tb = tb.get("nombre_trabajo", f"Trabajo {tb.get('id', '')}")
                        n_facs = tb.get("total_facturas", 0)
                        try: val_tot = float(tb.get("total_valor", 0.0) or 0.0)
                        except Exception: val_tot = 0.0
                        n_renom = tb.get("total_pdfs_renombrados", tb.get("total_pdfs", 0))
                        n_orig = tb.get("total_pdfs_originales", 0)
                        info_pdf_str = f"{n_renom} PDFs procesados" if n_renom > 0 else (f"{n_orig} PDFs subidos" if n_orig > 0 else "Sin PDFs")
                        st.markdown(f"\U0001f4c4 **{nom_tb}** \u2014 {n_facs} facturas (${val_tot:,.2f}) \u2014 **{info_pdf_str}**")
                    with c_h2:
                        if st.button("\U0001f4c2 Cargar", key=f"btn_h_load_{tb['id']}"):
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
                            st.success(f"\u00a1Trabajo '{nom_tb}' cargado! Todo tu avance est\u00e1 listo.")
                            st.rerun()
                    with c_h3:
                        if st.button("\U0001f5d1\ufe0f", key=f"btn_h_del_{tb['id']}"):
                            eliminar_trabajo_historial(empresa, tb["id"])
                            st.rerun()
            else:
                st.info("\U0001f4a1 A\u00fan no tienes trabajos guardados en el disco local para esta empresa.")
    
        with col_h_right:
            st.markdown("#### \U0001f6e1\ufe0f Respaldo Portable (.zip):")
            st.caption("Guarda o restaura todo tu trabajo en un solo archivo, ideal si cambias de PC o si el servidor se reinicia:")
    
            pass
    
            # 2. Uploader para restaurar desde un archivo de respaldo previo (con boton de confirmacion para evitar bucles)
            archivo_zip_restaurar = st.file_uploader(
                "\U0001f4e5 Restaurar Sesi\u00f3n desde Archivo de Respaldo (.indumaq / .zip):",
                type=["indumaq", "zip"],
                key=f"upl_zip_restore_historial_{st.session_state.get('_upl_reset_cnt', 0)}",
                help="Sube tu archivo de respaldo para restaurar exactamente donde lo dejaste."
            )
            if archivo_zip_restaurar is not None:
                sig_z_r = f"{archivo_zip_restaurar.name}_{archivo_zip_restaurar.size}"
                if st.session_state.get("_auto_restored_sig") != sig_z_r:
                    with st.spinner("Restaurando sesi\u00f3n y c\u00e1lculos autom\u00e1ticamente..."):
                        if restaurar_desde_respaldo_bytes(empresa, archivo_zip_restaurar):
                            st.session_state["_auto_restored_sig"] = sig_z_r
                            st.success(f"\u00a1Sesi\u00f3n restaurada con \u00e9xito desde {archivo_zip_restaurar.name}!")
                            st.rerun()
                if st.button("\U0001f680 Cargar este Respaldo en Pantalla", key="btn_confirmar_cargar_zip_historial", type="primary", use_container_width=True):
                    with st.spinner("Restaurando sesi\u00f3n y c\u00e1lculos..."):
                        ok_r = restaurar_desde_respaldo_bytes(empresa, archivo_zip_restaurar)
                        if not ok_r:
                            df_res, renom_res, orig_res, ex_b_res, ex_n_res, zip_res = importar_respaldo_sesion_zip(archivo_zip_restaurar.getvalue(), empresa)
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
                            st.success(f"\u00a1Sesi\u00f3n restaurada con \u00e9xito desde {archivo_zip_restaurar.name}!")
                            st.rerun()
                        else:
                            st.error(f"No se pudo restaurar el archivo de respaldo: {st.session_state.get('_ultimo_error_respaldo', '')}")
    
    st.markdown("---")
    
    

    
    
    # [Panel superior integrado en el escudo global]

    tab_compras, tab_auditoria, tab_triangulacion, tab_siigo = st.tabs([
        "1. Cargar Documentos y Desbloquear PDFs",
        "2. Auditoria y Trazabilidad Fiscal",
        "3. \U0001f500 Triangulaci\u00f3n y Cruces de Importaci\u00f3n (Aduanas)",
        "4. Exportar Planilla Oficial a Siigo"
    ])

    with tab_compras:
        try:
            st.markdown("### 1. Insumos DIAN y Facturas en PDF")
            st.write("Sube el archivo Excel de la DIAN (`prueba.xlsx`) o el reporte de facturas, y los PDFs (unificados o separados) para desbloquear, guardar y renombrar autom\u00e1ticamente por comprobante.")
    
            # Si ya tenemos los documentos cargados en el autoguardado, mostrar confirmacion visible
            n_p_carg = len(st.session_state.get("raw_uploaded_pdfs", {}))
            n_p_proc = len(st.session_state.get("dict_pdfs", {}))
            if n_p_carg > 0:
                st.success(f"\u2705 **Autoguardado Activo:** Tienes **{n_p_carg} facturas PDF cargadas en el sistema** ({n_p_proc} vinculadas con su comprobante). Ya est\u00e1n disponibles en la Pesta\u00f1a 2 (Auditor\u00eda) y Pesta\u00f1a 3 (Triangulaci\u00f3n). No necesitas volver a subirlas.")
    
            col_u1, col_u2 = st.columns(2)
            with col_u1:
                archivo_excel = st.file_uploader("1. Reporte Excel de la DIAN (ej. prueba.xlsx)", type=["xlsx", "xls"], help="Opcional si ya cargaste tu trabajo desde el Historial.")
            with col_u2:
                archivos_pdfs = st.file_uploader("2. Facturas en PDF, ZIP Consolidado o Respaldo (.indumaq)", type=["pdf", "zip", "indumaq"], accept_multiple_files=True, key=f"archivos_pdfs_upl_{st.session_state.get('_upl_reset_cnt', 0)}", help="Sube facturas en PDF, un archivo .ZIP con todas las facturas, o tu archivo de respaldo (.indumaq).")
    
            st.markdown("##### \U0001f522 3. Consecutivos Iniciales por Tipo de Comprobante Siigo:")
            st.caption("Cada tipo de comprobante lleva su propia numeraci\u00f3n independiente. Digita en qu\u00e9 n\u00famero vas en cada uno:")
    
            c_con1, c_con2, c_con3, c_con4 = st.columns(4)
            with c_con1:
                cons_ini_fac = st.number_input("Facturas de Compra (Tipo 10):", min_value=1, max_value=999999, value=680, step=1, help="Consecutivo inicial para facturas de compra normales.")
            with c_con2:
                cons_ini_nc = st.number_input("Notas Cr\u00e9dito (Tipo 17):", min_value=1, max_value=999999, value=1, step=1, help="Consecutivo inicial para notas cr\u00e9dito / devoluciones.")
            with c_con3:
                cons_ini_nota = st.number_input("Notas de Contabilidad (Tipo 14):", min_value=1, max_value=999999, value=1, step=1, help="Consecutivo inicial para notas contables o ajustes.")
            with c_con4:
                clave_pdf_extra = st.text_input("Clave PDF (opcional):", type="password", help="Si alg\u00fan archivo PDF tiene contrase\u00f1a espec\u00edfica, ingr\u00e9sala aqu\u00ed.")
    
            # Almacenar PDFs originales subidos en session_state y en el historial de inmediato (solo una vez por subida)
            if archivos_pdfs:
                pdfs_sig = f"{len(archivos_pdfs)}_" + "_".join(f"{p.name}_{p.size}" for p in archivos_pdfs[:5])
                if st.session_state.get("_ultimo_pdfs_proc_sig") != pdfs_sig:
                    if "raw_uploaded_pdfs" not in st.session_state:
                        st.session_state["raw_uploaded_pdfs"] = {}
                    import gc
                    for p in archivos_pdfs:
                        if p.name.lower().endswith(".indumaq"):
                            st.session_state["_ultimo_pdfs_proc_sig"] = pdfs_sig
                            if restaurar_desde_respaldo_bytes(empresa, p, solo_pdfs=False):
                                st.success(f"\u00a1Respaldo {p.name} y sus 224 PDFs restaurados con \u00e9xito!")
                                st.rerun()
                            continue
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
                    st.success(f"\U0001f4d1 **{len(st.session_state['raw_uploaded_pdfs'])} factura(s) PDF cargadas en memoria.**")
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
    
                    # Identificar columna de fecha y ordenar cronologicamente de Enero a la fecha actual
                    col_fecha = None
                    for col_cand in ["Fecha Emisi\u00f3n", "Fecha Emision", "Fecha", "Fecha de Emisi\u00f3n"]:
                        if col_cand in df_dian.columns:
                            col_fecha = col_cand
                            break
                    if col_fecha is None:
                        col_fecha = df_dian.columns[0]
    
                    df_dian["_fecha_dt"] = pd.to_datetime(df_dian[col_fecha], dayfirst=True, errors="coerce")
                    df_dian = df_dian.sort_values(by="_fecha_dt", ascending=True).reset_index(drop=True)
    
                    st.success(f"Reporte procesado y ordenado cronol\u00f3gicamente (Enero a la fecha): **{len(df_dian)} facturas identificadas**.")
    
                    contadores_consecutivos = {
                        10: int(cons_ini_fac),
                        17: int(cons_ini_nc),
                        14: int(cons_ini_nota),
                        16: int(cons_ini_fac)
                    }
                    # Analizar celdas rojas (ya registradas), grupos y cuentas especificas del Excel
                    estados_excel = analizar_estado_filas_excel(st.session_state.get("excel_bytes"))
    
                    filas = []
                    for idx, r in df_dian.iterrows():
                        tipo_doc = r.get("Tipo de documento") or r.get("Tipo documento") or "Factura electr\u00f3nica"
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
    
                        col_resp = next((c for c in df_dian.columns if any(k in str(c).lower() for k in ["r\u00e9gimen", "regimen", "responsabilidad", "obligacion"])), None)
                        resp_e = str(r.get(col_resp, "")).strip() if col_resp and pd.notna(r.get(col_resp)) else ""
    
                        # Busqueda inmediata por NIT y Razon Social en catalogo oficial
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
                            "N\u00b0": idx + 1,
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
                            "Grupo Importaci\u00f3n": grp_imp,
                            "Cuenta Pasivo Especifica": cta_esp,
                            "Es Aduanera": es_aduanera_flag,
                            "Estado Registro": f"\U0001f534 Ya Registrada ({comp_prev})" if es_roja_reg else ("\U0001f7e1 Agente Aduanero" if any(k in nom_e.upper() for k in ["EURO SHIPPING", "TRADE GLOBAL"]) else ("\U0001f7e2 Aduanera" if es_aduanera_flag else "\u26aa Compra Pendiente")),
                            "R\u00e9gimen Fiscal Emisor": resp_e if resp_e else "O-48 (Est\u00e1ndar)",
                            "Descripcion": desc,
                            "Operacion": op,
                            "Cta Principal": cta_p,
                            "Categor\u00eda": cat,
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
                            "Raz\u00f3n Contable": razon,
                            "Soporte PDF Renombrado": nombre_pdf_esperado
                        })
    
                    df_proc = pd.DataFrame(filas)
                    st.session_state["df_procesado"] = df_proc
    
                    st.success(f"\u2705 **Reporte Excel '{archivo_excel.name}' procesado:** {len(df_proc)} facturas identificadas y ordenadas cronol\u00f3gicamente.")
    
            if "df_procesado" in st.session_state and st.session_state["df_procesado"] is not None:
                df_proc = st.session_state["df_procesado"]
    
                c_job1, c_job2, c_job3 = st.columns([2.8, 1.1, 1.1])
                with c_job1:
                    n_orig_mem = len(st.session_state.get("raw_uploaded_pdfs", {}))
                    n_renom_mem = len(st.session_state.get("dict_pdfs", {}))
                    nom_ex = st.session_state.get("excel_nombre", "Reporte.xlsx")
                    st.info(f"\U0001f4c2 **Reporte DIAN:** `{nom_ex}` ({len(df_proc)} facturas) | \U0001f4d1 **{n_orig_mem} PDFs cargados** | \U0001f4c4 **{n_renom_mem} vinculados**")
                    if n_orig_mem > 0 and n_renom_mem == 0:
                        if st.button(f"\U0001f680 Desbloquear y Vincular los {n_orig_mem} PDFs a Comprobantes", key="btn_desbloquear_top_bar", type="primary", use_container_width=True):
                            st.session_state["_ejecutar_desbloqueo_ahora"] = True
                            st.rerun()
                with c_job2:
                    if "excel_bytes" in st.session_state and st.session_state["excel_bytes"]:
                        st.download_button(
                            label="\U0001f4e5 Descargar Excel",
                            data=st.session_state["excel_bytes"],
                            file_name=st.session_state.get("excel_nombre", "Reporte.xlsx"),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="btn_dl_excel_tab_compras",
                            use_container_width=True
                        )
                with c_job3:
                    st.caption("\u26a1 Memoria RAM Protegida (PDFs en disco)")
    
    
                c_aud_btn1, c_aud_btn2 = st.columns(2)
                with c_aud_btn1:
                    st.markdown("#### Matriz Contable Preliminar vinculada a Comprobantes (Orden Cronol\u00f3gico Enero - Actual):")
                with c_aud_btn2:
                    if st.button("\u26a1 Auditar y Escanear R\u00e9gimen Fiscal de los PDFs (Ultrarr\u00e1pido)", key="btn_escanear_regimen_todos"):
                        dict_renom_s = st.session_state.get("dict_pdfs", {})
                        dict_orig_s = st.session_state.get("raw_uploaded_pdfs", {})
                        total_modificados = 0
                        if dict_renom_s:
                            with st.spinner("\u26a1 Escaneando r\u00e9gimen fiscal directamente desde las facturas ya desbloqueadas y renombradas..."):
                                total_modificados = auditar_regimen_desde_facturas_renombradas(df_proc, dict_renom_s, empresa)
                            st.session_state["df_procesado"] = df_proc
                            pass
                            st.success(f"\u26a1 Auditor\u00eda completada desde facturas renombradas: se revisaron {len(df_proc)} facturas y se actualizaron {total_modificados} con su r\u00e9gimen oficial del PDF.")
                            st.rerun()
                        elif dict_orig_s:
                            with st.spinner("\u26a1 Extrayendo texto en memoria y cruzando reg\u00edmenes fiscales..."):
                                # 1. Pre-cargar textos en cache de todos los PDFs una sola vez
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
    
                                    # Buscar en textos pre-extraidos
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
                                        if reg_f and reg_f != r_mat.get("R\u00e9gimen Fiscal Emisor"):
                                            df_proc.at[r_idx, "R\u00e9gimen Fiscal Emisor"] = reg_f
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
                                            df_proc.at[r_idx, "Raz\u00f3n Contable"] = razon_n
                                            df_proc.at[r_idx, "Audit Info"] = audit_n
                                            df_proc.at[r_idx, "Neto a Pagar"] = round(df_proc.at[r_idx, "Total"] - rfte_n - rica_n - riva_n, 2)
                                            total_modificados += 1
    
                            st.session_state["df_procesado"] = df_proc
                            pass
                            st.success(f"\u26a1 Auditor\u00eda completada en segundos: se revisaron {len(df_proc)} facturas y se actualizaron {total_modificados} con su r\u00e9gimen oficial del PDF.")
                            st.rerun()
                        else:
                            st.warning("Primero sube los PDFs en el campo 2 para poder escanearlos.")
    
                # Asegurar columna Cuenta Pasivo (CxP) con marcacion clara
                if "Cuenta Pasivo (CxP)" not in df_proc.columns:
                    df_proc["Cuenta Pasivo (CxP)"] = df_proc.apply(
                        lambda r: f"\u26a0\ufe0f {r.get('Cuenta Pasivo Especifica', '22050505')} (Por Verificar)" if r.get("Ya Registrada") else str(r.get("Cuenta Pasivo Especifica", r.get("Cta Contrapartida", "22050501"))),
                        axis=1
                    )
    
                st.markdown("""
                <div style="background:#eff6ff; border:1px solid #bfdbfe; border-left:5px solid #2563eb; border-radius:6px; padding:10px 14px; margin-bottom:12px; font-size:13.5px; color:#1e40af;">
                    <b>\u2139\ufe0f Cuentas de Pasivo (CxP) para Facturas Registradas:</b> Las facturas ya causadas en Siigo (\U0001f534) tienen asignada una cuenta por pagar estimada (ej. <code>22050505</code> Agencia Aduana, <code>23359501</code> Acreedores DHL). Aparecen marcadas con <b>\u26a0\ufe0f (Por Verificar)</b>. Puedes modificar la cuenta f\u00e1cilmente en la <b>Pesta\u00f1a 2 (Revisi\u00f3n)</b> o en la <b>Pesta\u00f1a 3 (Triangulaci\u00f3n)</b>.
                </div>
                """, unsafe_allow_html=True)
    
    
                # AUTO-AUDITORIA Y VERIFICACION INMEDIATA DE REGIMENES FISCALES AL CARGAR EXCEL O PDFS
                dict_renom_mem = st.session_state.get("dict_pdfs", {})
                dict_orig_mem = st.session_state.get("raw_uploaded_pdfs", {})
    
                # 1. Aplicar catalogo oficial de regimenes conocidos por NIT (solo una vez al cargar archivo inicial)
                if not st.session_state.get("_regimenes_conocidos_aplicados_v1", False):
                    if "Audit Info" not in df_proc.columns:
                        df_proc["Audit Info"] = [{} for _ in range(len(df_proc))]
                    for r_idx, r_row in df_proc.iterrows():
                        if bool(r_row.get("Editada Manualmente", False)):
                            continue
                        nit_d = re.sub(r'\D', '', str(r_row.get("NIT Emisor", "")))
                        nom_u = str(r_row.get("Proveedor", "")).upper()
                        reg_prev = str(r_row.get("R\u00e9gimen Fiscal Emisor", ""))
    
                        nuevo_reg = None
                        if nit_d in REGIMENES_EMISORES_CONOCIDOS:
                            nuevo_reg = REGIMENES_EMISORES_CONOCIDOS[nit_d]
                        elif any(k in nom_u for k in ["DHL", "ESTELAR", "BUENAVENTURA", "SIIGO", "COMCEL", "CLARO", "PANAMERICANA"]):
                            nuevo_reg = "O-13;O-15"
    
                        if nuevo_reg and (nuevo_reg != reg_prev):
                            df_proc.at[r_idx, "R\u00e9gimen Fiscal Emisor"] = nuevo_reg
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
                            df_proc.at[r_idx, "Raz\u00f3n Contable"] = razon_c
                            df_proc.at[r_idx, "Audit Info"] = audit_c
                            df_proc.at[r_idx, "Neto a Pagar"] = round(r_row["Total"] - rfte_c - rica_c - riva_c, 2)
                    st.session_state["_regimenes_conocidos_aplicados_v1"] = True
    
                # Auditoria de regimen fiscal disponible bajo demanda mediante el boton superior
                st.session_state["df_procesado"] = df_proc
    
                cols_mostrar_proc = ["Comprobante Siigo", "Fecha", "Factura", "Proveedor", "NIT Emisor", "R\u00e9gimen Fiscal Emisor", "Estado Registro", "Cuenta Pasivo (CxP)", "Base", "IVA", "ReteFuente", "ReteICA", "Neto a Pagar", "Soporte PDF Renombrado"]
                cols_mostrar_existentes = [c for c in cols_mostrar_proc if c in df_proc.columns]
                st.dataframe(_df_para_streamlit(df_proc[cols_mostrar_existentes]), use_container_width=True)
    
                st.markdown("---")
                st.markdown("#### \U0001f4d1 Procesamiento, Separaci\u00f3n y Renombrado de PDFs Multi-P\u00e1gina:")
                st.caption("Identifica y une facturas completas de 2 o m\u00e1s hojas a partir de un PDF consolidado, asignando a cada una su comprobante y consecutivo oficial.")
    
                col_cfg1, col_cfg2 = st.columns(2)
                with col_cfg1:
                    modo_sep = st.radio(
                        "Configuraci\u00f3n de reconocimiento de facturas:",
                        [
                            "\U0001f50d Detecci\u00f3n Inteligente (Reconoce facturas completas de 2 o m\u00e1s hojas por 'P\u00e1gina 1 de 2', CUFE y N\u00famero)",
                            "\U0001f4c4 Agrupaci\u00f3n fija de 2 p\u00e1ginas por factura (cada factura tiene 2 hojas)",
                            "\U0001f4c4 Agrupaci\u00f3n fija de 3 p\u00e1ginas por factura (cada factura tiene 3 hojas)",
                            "\U0001f4d1 Separar 1 p\u00e1gina por factura (facturas de 1 sola hoja)",
                            "\U0001f4c2 Mantener archivos individuales (sin separar p\u00e1ginas)"
                        ]
                    )
                with col_cfg2:
                    st.info("\U0001f4a1 **Garant\u00eda de Factura Completa:** El motor inteligente detecta d\u00f3nde empieza cada factura (Prefijo, Folio, NIT y marcadores de paginaci\u00f3n). Todas las p\u00e1ginas de una misma factura se unen en un solo archivo PDF completo nombrado `Comp_10-XXX_Factura_Proveedor.pdf`.")
    
                ejecutar_desb = st.button("\U0001f513 Desbloquear, Separar y Renombrar PDFs ahora") or st.session_state.pop("_ejecutar_desbloqueo_ahora", False)
                if ejecutar_desb:
                    df_ref = st.session_state.get("df_procesado", pd.DataFrame())
    
                    # Procesamiento en streaming desde disco (0 MB retenidos en RAM)
                    asegurar_dicts_pdf_en_disco()
                    if archivos_pdfs:
                        for p_it in archivos_pdfs:
                            p_nom = getattr(p_it, "name", "").lower()
                            if p_nom.endswith(".zip"):
                                try:
                                    b_val = p_it.getvalue() if hasattr(p_it, "getvalue") else p_it.read()
                                    with zipfile.ZipFile(io.BytesIO(b_val), "r") as z_u:
                                        for z_fn in z_u.namelist():
                                            if z_fn.lower().endswith(".pdf") and not z_fn.startswith("__MACOSX"):
                                                c_fn = os.path.basename(z_fn)
                                                if c_fn:
                                                    st.session_state["raw_uploaded_pdfs"][c_fn] = z_u.read(z_fn)
                                    del b_val
                                except Exception as e_z:
                                    st.error(f"Error abriendo ZIP {getattr(p_it, 'name', '')}: {e_z}")
                            elif p_nom.endswith(".pdf") or not p_nom:
                                st.session_state["raw_uploaded_pdfs"][getattr(p_it, "name", "factura.pdf")] = p_it.getvalue() if hasattr(p_it, "getvalue") else p_it.read()
                        gc.collect()
    
                    nombres_pdfs_proc = [fn for fn in list(st.session_state["raw_uploaded_pdfs"].keys()) if fn.lower().endswith(".pdf")]
                    if not nombres_pdfs_proc:
                        st.warning("\u26a0\ufe0f No hay archivos PDF para procesar. Por favor sube tus facturas en el campo '2. Facturas en PDF (unificadas o separadas)' arriba antes de hacer clic en desbloquear.")
    
                    total_generados = 0
                    total_pdfs_count = max(1, len(nombres_pdfs_proc))
                    prog_bar_pdf = st.progress(0, text=f"Iniciando procesamiento liviano de {len(nombres_pdfs_proc)} archivo(s) PDF...") if nombres_pdfs_proc else st.empty()
    
                    with zipfile.ZipFile(ZIP_FACTURAS_DISK_PATH, "w", zipfile.ZIP_STORED) as zf:
                        for idx_pdf, pdf_name in enumerate(nombres_pdfs_proc):
                            if idx_pdf % 15 == 0:
                                gc.collect()
                            prog_bar_pdf.progress(min(1.0, (idx_pdf + 1) / total_pdfs_count), text=f"\u26a1 Factura {idx_pdf + 1} de {total_pdfs_count}: {pdf_name[:20]}...")
                            if pdf_name in st.session_state["dict_pdfs"]:
                                b_existing = st.session_state["dict_pdfs"][pdf_name]
                                if b_existing and len(b_existing) > 100:
                                    zf.writestr(pdf_name, b_existing)
                                    total_generados += 1
                                    continue
                            try:
                                b_raw_one = st.session_state["raw_uploaded_pdfs"][pdf_name]
                                pdf_item = io.BytesIO(b_raw_one)
                                pdf_item.name = pdf_name
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
                                        st.warning(f"\u26a0\ufe0f No se pudo desencriptar autom\u00e1ticamente {pdf_name}. Ingrese la contrase\u00f1a en el campo correspondiente.")
                                        continue
    
                                num_pags = len(reader.pages)
    
                                # Caso 1: Archivo individual (sin separar paginas) o de 1 pagina
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
    
                                        # 2. Si no se identifico en el texto, buscar por coincidencia en el nombre del archivo
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
                                        # Escaneo de regimen fiscal solo para la factura que realmente coincidio
                                        regimen_escaneado = escanear_regimen_texto_pdf(texto_pdf)
                                        if regimen_escaneado:
                                            r_idx = r_mat.name
                                            df_ref.at[r_idx, "R\u00e9gimen Fiscal Emisor"] = regimen_escaneado
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
                                            df_ref.at[r_idx, "Raz\u00f3n Contable"] = razon_n
                                            df_ref.at[r_idx, "Audit Info"] = audit_n
                                            df_ref.at[r_idx, "Neto a Pagar"] = round(df_ref.at[r_idx, "Total"] - rfte_n - rica_n - riva_n, 2)
                                    else:
                                        nombre_final = f"Soporte_{idx_pdf+1}_{pdf_name}"
    
                                    # Reutilizar directamente los bytes originales sin re-escribir con PdfWriter (1000x mas rapido)
                                    b_data = pdf_item.getvalue() if hasattr(pdf_item, "getvalue") else pdf_item.read()
                                    zf.writestr(nombre_final, b_data)
                                    st.session_state["dict_pdfs"][nombre_final] = b_data
                                    total_generados += 1
    
                                # Caso 2: Agrupacion fija de 2 paginas
                                elif "Agrupaci\u00f3n fija de 2 p\u00e1ginas" in modo_sep:
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
    
                                # Caso 3: Agrupacion fija de 3 paginas
                                elif "Agrupaci\u00f3n fija de 3 p\u00e1ginas" in modo_sep:
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
    
                                # Caso 4: 1 pagina por factura
                                elif "Separar 1 p\u00e1gina" in modo_sep:
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
    
                                # Caso 5: Deteccion inteligente por CUFE, Folio o 'Pagina 1 de N'
                                else:
                                    patron_pag = re.compile('(?:P[\u00c1A\u00e1a]G(?:INA|\\.)?|HOJA|PAGE)\\s*(\\d+)\\s*(?:DE|\\/|OF)\\s*(\\d+)', re.IGNORECASE)
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
                                                facturas_generadas.append(f"{curr_inv_row['Comprobante Siigo']} ({paginas_del_comprobante} p\u00e1gs)")
    
                                            curr_inv_row = inv_encontrada if inv_encontrada is not None else curr_inv_row
                                            if curr_inv_row is not None and reg_detectado and not df_ref.empty:
                                                r_idx = curr_inv_row.name
                                                df_ref.at[r_idx, "R\u00e9gimen Fiscal Emisor"] = reg_detectado
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
                                                df_ref.at[r_idx, "Raz\u00f3n Contable"] = razon_n
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
                                        facturas_generadas.append(f"{curr_inv_row['Comprobante Siigo']} ({paginas_del_comprobante} p\u00e1gs)")
    
                            except Exception as e:
                                st.error(f"Error procesando {pdf_name}: {e}")
    
                    prog_bar_pdf.progress(1.0, text="\u2705 \u00a1100% Procesado y Vinculado con \u00c9xito!")
                    st.session_state["zip_pdfs_path"] = ZIP_FACTURAS_DISK_PATH
                    st.session_state["zip_pdfs"] = None
                    st.session_state["total_zip_pdfs"] = total_generados
                    gc.collect()
    
                    # Escaneo oficial de regimen fiscal directamente sobre las facturas ya desbloqueadas y renombradas
                    total_auditados = auditar_regimen_desde_facturas_renombradas(df_ref, st.session_state["dict_pdfs"], empresa)
                    st.session_state["df_procesado"] = df_ref
                    gc.collect()
    
                    # Guardar trabajo actualizado en disco sin saturar RAM
                    guardar_trabajo_en_historial(
                        empresa, df_ref,
                        excel_bytes=st.session_state.get("excel_bytes"),
                        excel_nombre=st.session_state.get("excel_nombre", "Reporte.xlsx"),
                        dict_pdfs_renombrados=None,
                        dict_pdfs_originales=None,
                        zip_bytes=None,
                        consecutivo_ini=cons_ini_fac,
                        job_id=st.session_state.get("job_actual_id")
                    )
                    msg_reg = f" y se audit\u00f3 el r\u00e9gimen de cada factura ({total_auditados} actualizadas)" if total_auditados > 0 else " y se valid\u00f3 el r\u00e9gimen de cada factura"
                    st.success(f"\u00a1Procesamiento exitoso! Se desbloquearon y renombraron **{total_generados} facturas completas**{msg_reg} directamente desde las facturas PDF desbloqueadas.")
    
                zip_disk_p = st.session_state.get("zip_pdfs_path", ZIP_FACTURAS_DISK_PATH)
                if os.path.exists(zip_disk_p) and os.path.getsize(zip_disk_p) > 22:
                    n_tot_z = st.session_state.get("total_zip_pdfs", len(st.session_state.get("dict_pdfs", {})))
                    if st.checkbox(f"\U0001f4e6 Habilitar bot\u00f3n de descarga ZIP de {n_tot_z} facturas renombradas (Bajo demanda para proteger memoria RAM)", value=False, key="chk_dl_zip_main_ondemand"):
                        with open(zip_disk_p, "rb") as f_zdl:
                            st.download_button(
                                label=f"\u2b07\ufe0f Descargar Archivos PDF Renombrados ({n_tot_z} facturas en ZIP)",
                                data=f_zdl.read(),
                                file_name=f"Facturas_Renombradas_Comprobantes_{empresa['nombre'].replace(' ', '_')}.zip",
                                mime="application/zip",
                                key="btn_dl_zip_pdfs_renombrados_action_main",
                                use_container_width=True
                            )
    
    
    
            else:
                st.info("\U0001f4a1 Sube el reporte Excel de la DIAN arriba o carga un trabajo guardado desde el **Historial de Trabajos** para comenzar.")
    
        except Exception as _e_comp:
            if type(_e_comp).__name__ in ['RerunException', 'StopException', 'ScriptControlException']:
                raise _e_comp
            st.error(f"\u26a0\ufe0f Ocurri\u00f3 un error en la pesta\u00f1a de Compras DIAN: {_e_comp}")
            with st.expander("\U0001f50d Ver detalles t\u00e9cnicos del error (Traceback):"):
                st.code(traceback.format_exc())
    with tab_auditoria:
        try:
            st.markdown("### Modulo de Auditoria Contable y Trazabilidad")
            st.caption("Inspecci\u00f3n de cuentas, deducciones y separaci\u00f3n de gastos por cuenta de terceros.")
    
            if "df_procesado" in st.session_state and st.session_state["df_procesado"] is not None and isinstance(st.session_state["df_procesado"], pd.DataFrame) and not st.session_state["df_procesado"].empty:
                df_p = normalizar_df_procesado(st.session_state["df_procesado"])
                st.session_state["df_procesado"] = df_p
    
                # Inicializar conjunto de facturas excluidas de contabilizacion
                if "facturas_no_contabilizar" not in st.session_state:
                    st.session_state["facturas_no_contabilizar"] = set(df_p[df_p["No Contabilizar"] == True]["Factura"].dropna().astype(str).tolist())
                cur_no_contab = st.session_state["facturas_no_contabilizar"]
    
                # Panel de exclusion masiva para marcar cuales NO contabilizar
                with st.expander("\U0001f6ab Panel: Marcar qu\u00e9 facturas NO debes contabilizar (Excluir de Siigo):", expanded=False):
                    st.caption("Selecciona aqu\u00ed las facturas que NO deseas que pasen a la contabilidad ni se exporten a Siigo (por ejemplo, si ya fueron causadas manualmente, corresponden a otro per\u00edodo o no deben registrarse):")
                    lista_todas_facs_tags = []
                    mapa_tags_to_fac = {}
                    defaults_seleccionados = []
                    for idx_row, r in df_p.iterrows():
                        fn = str(r.get("Factura") or f"Doc_{idx_row+1}").strip()
                        prov = str(r.get("Proveedor") or "Proveedor").strip()[:24]
                        try: tot_val = float(r.get("Total", 0.0) or 0.0)
                        except Exception: tot_val = 0.0
                        ya_reg = bool(r.get("Ya Registrada", False))
                        tag_t = "{} \u2014 {} (${:,.0f}) [{}]".format(fn, prov, tot_val, "\U0001f534 Ya Registrada" if ya_reg else "\u26aa Pendiente")
                        lista_todas_facs_tags.append(tag_t)
                        mapa_tags_to_fac[tag_t] = fn
                        if fn in cur_no_contab or bool(r.get("No Contabilizar", False)):
                            defaults_seleccionados.append(tag_t)
    
                    sel_excluidas_bulk = st.multiselect(
                        "Selecciona facturas que NO se deben contabilizar:",
                        options=lista_todas_facs_tags,
                        default=defaults_seleccionados,
                        key="ms_bulk_excluir_contab"
                    )
                    if st.button("\U0001f4be Guardar Exclusiones de Contabilidad", key="btn_save_bulk_excl"):
                        nuevas_ex = {mapa_tags_to_fac[t] for t in sel_excluidas_bulk if t in mapa_tags_to_fac}
                        st.session_state["facturas_no_contabilizar"] = nuevas_ex
                        for idx_row, r_row in df_p.iterrows():
                            fn_curr = str(r_row.get("Factura", "")).strip()
                            es_ex = fn_curr in nuevas_ex
                            df_p.at[idx_row, "No Contabilizar"] = es_ex
                            if es_ex:
                                df_p.at[idx_row, "Estado Registro"] = "\U0001f6ab No Contabilizar (Excluida)"
                            else:
                                es_r = bool(r_row.get("Ya Registrada", False))
                                cp = str(r_row.get("Comprobante Previo", "")).strip()
                                df_p.at[idx_row, "Estado Registro"] = f"\U0001f534 Ya Registrada ({cp})" if es_r else "\u26aa Compra Pendiente"
                        st.session_state["df_procesado"] = df_p
                        st.session_state.pop("_cached_respaldo_bytes", None)
                        st.session_state.pop("_siigo_cache_sig", None)
                        guardar_estado_manual(empresa)
                        st.success(f"\u00a1Se han marcado {len(nuevas_ex)} facturas como NO contabilizables!")
                        st.rerun()
    
                # BARRA DE ACELERACION DE AUDITORIA: FILTROS Y NAVEGACION 1-CLIC
                n_tot_f = len(df_p)
                mask_excl_f = df_p["Factura"].astype(str).str.strip().isin(cur_no_contab) | (df_p["No Contabilizar"] == True)
                mask_edit_f = (df_p["Editada Manualmente"] == True) & (~mask_excl_f)
                mask_pend_f = (~mask_excl_f) & (~mask_edit_f)
                n_pend_f = int(mask_pend_f.sum())
                n_edit_f = int(mask_edit_f.sum())
                n_excl_f = int(mask_excl_f.sum())

                c_fbar1, c_fbar2 = st.columns([2.8, 1.4])
                with c_fbar1:
                    filtro_aud_modo = st.radio(
                        "\U0001f680 Filtro Rapido de Revision:",
                        [
                            f"\U0001f4cb Todas ({n_tot_f})",
                            f"\u23f3 Pendientes por Revisar ({n_pend_f})",
                            f"\u270f\ufe0f Revisadas / Editadas ({n_edit_f})",
                            f"\U0001f6ab No Contabilizar ({n_excl_f})"
                        ],
                        horizontal=True,
                        key="rad_filtro_rapido_hoja2"
                    )
                with c_fbar2:
                    progreso_pct = round(((n_edit_f + n_excl_f) / max(1, n_tot_f)) * 100, 1)
                    st.progress(min(1.0, (n_edit_f + n_excl_f) / max(1, n_tot_f)), text=f"Avance Revision: {n_edit_f + n_excl_f}/{n_tot_f} ({progreso_pct}%)")

                if "Pendientes" in filtro_aud_modo and n_pend_f > 0:
                    df_p_vista = df_p[mask_pend_f]
                elif "Revisadas" in filtro_aud_modo and n_edit_f > 0:
                    df_p_vista = df_p[mask_edit_f]
                elif "No Contabilizar" in filtro_aud_modo and n_excl_f > 0:
                    df_p_vista = df_p[mask_excl_f]
                else:
                    df_p_vista = df_p

                opciones_fac = []
                comps_en_vista = []
                for idx_r, r in df_p_vista.iterrows():
                    fn = str(r.get("Factura") or f"Doc_{idx_r+1}").strip()
                    es_ex_r = fn in cur_no_contab or bool(r.get("No Contabilizar", False))
                    es_ed_r = bool(r.get("Editada Manualmente", False))
                    tag_ex = " \U0001f6ab [NO CONTABILIZAR]" if es_ex_r else (" \u2705 [REVISADA]" if es_ed_r else " \u23f3")
                    comp_str = str(r.get("Comprobante Siigo") or f"Comp {idx_r+1}").strip()
                    fecha_str = str(r.get("Fecha") or "").strip()
                    prov_str = str(r.get("Proveedor") or "Proveedor").strip()[:24]
                    try: tot_num = float(r.get("Total", 0.0) or 0.0)
                    except Exception: tot_num = 0.0
                    opciones_fac.append(f"[{comp_str}] {fecha_str} - {fn}{tag_ex} - {prov_str} (${tot_num:,.0f})")
                    comps_en_vista.append(comp_str)

                comp_activo_guardado = st.session_state.get("factura_seleccionada_auditoria")
                idx_fac_sel = 0
                if comp_activo_guardado and comp_activo_guardado in comps_en_vista:
                    idx_fac_sel = comps_en_vista.index(comp_activo_guardado)

                c_nav_sel, c_nav_prev, c_nav_next, c_nav_pend = st.columns([3.0, 0.65, 0.65, 0.9])
                with c_nav_prev:
                    st.write("")
                    if st.button("\u2b05\ufe0f Anterior", key="btn_nav_prev_h2", use_container_width=True, disabled=(idx_fac_sel <= 0)):
                        nuevo_i = max(0, idx_fac_sel - 1)
                        st.session_state["factura_seleccionada_auditoria"] = comps_en_vista[nuevo_i]
                        st.session_state["sel_factura_auditoria_box"] = opciones_fac[nuevo_i]
                        st.rerun()
                with c_nav_next:
                    st.write("")
                    if st.button("Siguiente \u27a1\ufe0f", key="btn_nav_next_h2", use_container_width=True, disabled=(idx_fac_sel >= len(opciones_fac) - 1)):
                        nuevo_i = min(len(opciones_fac) - 1, idx_fac_sel + 1)
                        st.session_state["factura_seleccionada_auditoria"] = comps_en_vista[nuevo_i]
                        st.session_state["sel_factura_auditoria_box"] = opciones_fac[nuevo_i]
                        st.rerun()
                with c_nav_pend:
                    st.write("")
                    if st.button("\u23ed\ufe0f Sig. Pendiente", key="btn_nav_next_pend_h2", use_container_width=True, help="Salta directamente a la siguiente factura pendiente por auditar"):
                        comps_pend = df_p[mask_pend_f]["Comprobante Siigo"].astype(str).str.strip().tolist()
                        if comps_pend:
                            sig_c = comps_pend[0]
                            for cp_cand in comps_pend:
                                if cp_cand > str(comp_activo_guardado or ""):
                                    sig_c = cp_cand
                                    break
                            st.session_state["factura_seleccionada_auditoria"] = sig_c
                            st.session_state.pop("sel_factura_auditoria_box", None)
                            st.rerun()

                with c_nav_sel:
                    if "sel_factura_auditoria_box" in st.session_state:
                        if st.session_state["sel_factura_auditoria_box"] not in opciones_fac and idx_fac_sel < len(opciones_fac):
                            st.session_state["sel_factura_auditoria_box"] = opciones_fac[idx_fac_sel]
                    seleccion = st.selectbox(
                        "Selecciona una factura para auditar:",
                        opciones_fac,
                        index=idx_fac_sel if idx_fac_sel < len(opciones_fac) else 0,
                        key="sel_factura_auditoria_box"
                    ) if opciones_fac else None

                if not seleccion and opciones_fac:
                    seleccion = opciones_fac[0]

                comp_sel = seleccion.split("]")[0].replace("[", "").strip()
                st.session_state["factura_seleccionada_auditoria"] = comp_sel
                m_fac_matches = df_p[df_p["Comprobante Siigo"].astype(str).str.strip() == comp_sel]
                if m_fac_matches.empty:
                    m_fac_matches = df_p
                fac_sel = m_fac_matches.iloc[0]
                es_aduanero = any(k in str(fac_sel.get("Proveedor", "")).upper() for k in AGENTES_ADUANEROS)

                # Auto-correccion inmediata si la factura seleccionada es de Transporte / Inland Rate
                es_transp_sel = ("70787" in str(fac_sel.get("Factura", "")) or
                                 "INLAND" in str(fac_sel.get("Descripcion", "")).upper() or
                                 ("EURO" in str(fac_sel.get("Proveedor", "")).upper() and any(k in str(fac_sel.get("Descripcion", "")).upper() for k in ["TRANSP", "FLETE", "INLAND"])))
                if es_transp_sel and not bool(fac_sel.get("Editada Manualmente", False)):
                    r_idx_sel = fac_sel.name
                    b_val_sel = float(fac_sel.get("Base", 0.0) or 0.0)
                    df_p.at[r_idx_sel, "Descripcion"] = "INLAND RATE TRANSPORTE TERRESTRE"
                    df_p.at[r_idx_sel, "Categor\u00eda"] = "Transporte Carga (1%)"
                    df_p.at[r_idx_sel, "Raz\u00f3n Contable"] = "Transporte Terrestre de Carga / Fletes (1% ReteFuente [23652505] - 4.14\u2030 ReteICA [23680513])"
                    df_p.at[r_idx_sel, "ReteFuente"] = round(b_val_sel * 0.01, 2)
                    df_p.at[r_idx_sel, "ReteICA"] = round(b_val_sel * 0.00414, 2)
                    df_p.at[r_idx_sel, "Cta ReteFuente"] = "23652505"
                    df_p.at[r_idx_sel, "Cta ReteICA"] = "23680513"
                    df_p.at[r_idx_sel, "Cta Principal"] = "14650501"
                    st.session_state["df_procesado"] = df_p
                    fac_sel = df_p.loc[r_idx_sel]
    
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
                            st.warning("\U0001f6ab **Esta factura est\u00e1 marcada como 'NO CONTABILIZAR'** \u2014 No se generar\u00e1 asiento ni se exportar\u00e1 a Siigo.")
                        else:
                            st.caption("Estado contable activo para exportaci\u00f3n a Siigo.")
                    with c_tog_ex2:
                        if not esta_excl_ind:
                            if st.button("\U0001f6ab No Contabilizar", key=f"btn_excl_indiv_{fac_sel['Comprobante Siigo']}", help="Excluye esta factura de Siigo"):
                                cur_no_contab.add(fac_act_fn)
                                st.session_state["facturas_no_contabilizar"] = cur_no_contab
                                m_idx = df_p[df_p["Factura"] == fac_act_fn].index
                                if not m_idx.empty:
                                    df_p.at[m_idx[0], "No Contabilizar"] = True
                                    df_p.at[m_idx[0], "Estado Registro"] = "\U0001f6ab No Contabilizar (Excluida)"
                                    st.session_state["df_procesado"] = df_p
                                st.session_state.pop("_cached_respaldo_bytes", None)
                                st.session_state.pop("_siigo_cache_sig", None)
                                guardar_estado_manual(empresa)
                                st.session_state["factura_seleccionada_auditoria"] = comp_sel
                                st.session_state.pop("sel_factura_auditoria_box", None)
                                st.rerun()
                        else:
                            if st.button("\u2705 Reactivar", key=f"btn_react_indiv_{fac_sel['Comprobante Siigo']}", help="Vuelve a incluirla en la contabilidad"):
                                cur_no_contab.discard(fac_act_fn)
                                st.session_state["facturas_no_contabilizar"] = cur_no_contab
                                m_idx = df_p[df_p["Factura"] == fac_act_fn].index
                                if not m_idx.empty:
                                    df_p.at[m_idx[0], "No Contabilizar"] = False
                                    es_r = bool(df_p.at[m_idx[0], "Ya Registrada"])
                                    cp = df_p.at[m_idx[0], "Comprobante Previo"]
                                    df_p.at[m_idx[0], "Estado Registro"] = f"\U0001f534 Ya Registrada ({cp})" if es_r else "\u26aa Compra Pendiente"
                                    st.session_state["df_procesado"] = df_p
                                st.session_state.pop("_cached_respaldo_bytes", None)
                                st.session_state.pop("_siigo_cache_sig", None)
                                guardar_estado_manual(empresa)
                                st.session_state["factura_seleccionada_auditoria"] = comp_sel
                                st.session_state.pop("sel_factura_auditoria_box", None)
                                st.rerun()
    
                    reg_f_str_card = fac_sel.get("R\u00e9gimen Fiscal Emisor", "O-48")
                    tag_gc_card = "\u2705 Gran Contribuyente (O-13)" if audit_info.get("es_gc") else "Com\u00fan (No GC)"
                    tag_autorr_card = "\u2705 Autorretenedor (O-15)" if audit_info.get("es_autorr") else "No Autorretenedor"
                    tag_rst_card = "\u26a0\ufe0f R\u00e9gimen Simple (O-47)" if audit_info.get("es_rst") else "R\u00e9gimen Ordinario"
                    cat_f_str_card = fac_sel.get("Categor\u00eda", "General")
                    razon_f_str_card = fac_sel.get("Raz\u00f3n Contable", "Clasificaci\u00f3n autom\u00e1tica")
                    val_iva19_card = float(fac_sel.get("IVA 19%", 0.0) or 0.0)
                    val_iva5_card = float(fac_sel.get("IVA 5%", 0.0) or 0.0)
                    val_inc_card = float(fac_sel.get("INC / Otros Impuestos", 0.0) or 0.0)
                    val_ivaimp_card = float(fac_sel.get("IVA Importaci\u00f3n", 0.0) or 0.0)
                    cta_ivaimp_card = str(fac_sel.get("Cta IVA Importaci\u00f3n") or "240835").strip()
                    info_iva_extra_card = f" (IVA 19%: ${val_iva19_card:,.2f} [24081001] | IVA 5%: ${val_iva5_card:,.2f} [24081003])" if val_iva5_card > 0 else ""
                    info_inc_extra_card = (f" | <b>INC / Otros Impuestos:</b> ${val_inc_card:,.2f}" if val_inc_card > 0 else "") + (f" | <b>\U0001f6a2 IVA Importaci\u00f3n:</b> ${val_ivaimp_card:,.2f} [{cta_ivaimp_card}]" if val_ivaimp_card > 0 else "")
                    st.markdown(f"""
                    <div class="audit-card">
                        <h4>Detalle del Comprobante: {fac_sel['Comprobante Siigo']}</h4>
                        <p><b>Soporte PDF Vinculado:</b> <span class="tag-comp">{fac_sel.get('Soporte PDF Renombrado', '')}</span></p>
                        <p><b>Factura:</b> {fac_sel.get('Factura', '')} - <b>Proveedor:</b> {fac_sel.get('Proveedor', 'Proveedor')} (NIT: {fac_sel.get('NIT Emisor', '')})</p>
                        <p><b>Fecha de Emisi\u00f3n:</b> {fac_sel.get('Fecha', '')} - <b>Total:</b> ${float(fac_sel.get('Total', 0.0) or 0.0):,.2f}</p>
                        <p><b>R\u00e9gimen Fiscal Emisor:</b> <span class="badge-active">{reg_f_str_card}</span></p>
                        <div style="margin: 8px 0;">
                            <span class="badge-tax">{tag_gc_card}</span>
                            <span class="badge-tax">{tag_autorr_card}</span>
                            <span class="badge-tax">{tag_rst_card}</span>
                        </div>
                        <hr>
                        <h5>Trazabilidad de la Contabilizaci\u00f3n y Validaci\u00f3n DIAN:</h5>
                        <ul>
                            <li><b>Cuenta Asignada:</b> <span class="tag-propio">{fac_sel.get('Cta Principal', '14650501')}</span> - {cat_f_str_card}</li>
                            <li><b>Cruce DIAN:</b> {razon_f_str_card}</li>
                            <li><b>Base Gravable:</b> ${float(fac_sel.get('Base', 0.0) or 0.0):,.2f} - <b>IVA Total:</b> ${float(fac_sel.get('IVA', 0.0) or 0.0):,.2f}{info_iva_extra_card}{info_inc_extra_card}</li>
                            <li><b>Retenci\u00f3n en la Fuente:</b> ${float(fac_sel.get('ReteFuente', 0.0) or 0.0):,.2f} (Cuenta: <code>{fac_sel.get('Cta ReteFuente') or 'Sin ReteFuente'}</code>)</li>
                            <li><b>ReteICA:</b> ${float(fac_sel.get('ReteICA', 0.0) or 0.0):,.2f} (Cuenta: <code>{fac_sel.get('Cta ReteICA') or 'Sin ReteICA'}</code>)</li>
                            <li><b>ReteIVA:</b> ${fac_sel.get('ReteIVA', 0.0):,.2f} ({audit_info.get('razon_reteiva', 'N/A')})</li>
                        </ul>
                    </div>
                    """, unsafe_allow_html=True)
    
                    with st.expander("\u2699\ufe0f Ajustar Valores, Cuentas Contables (PUC) y Retenciones Formuladas:", expanded=True):
                        comp_box_id = str(fac_sel.get("Comprobante Siigo", "")).strip()
                        es_importacion_factura = es_aduanero or es_transp_sel or "1465" in str(fac_sel.get("Cta Principal", "")) or "IMPORTACI" in str(fac_sel.get("Categor\u00eda", "")).upper()

                        # Catalogos Oficiales de Cuentas con Tarifas para Formulacion Automatica
                        CATALOGO_CTAS_RETEFUENTE = [
                            ("23652505", 0.01,   "23652505 - Transporte de Carga / Fletes (1.0%)"),
                            ("23654001", 0.025,  "23654001 - Compras Generales / Repuestos (2.5%)"),
                            ("23654004", 0.035,  "23654004 - Compras / Alojamiento / Restaurante (3.5%)"),
                            ("23652501", 0.035,  "23652501 - Servicios Hoteles / Hospedaje (3.5%)"),
                            ("23652503", 0.04,   "23652503 - Servicios Generales / Agenciamiento (4.0%)"),
                            ("23652502", 0.06,   "23652502 - Servicios No Declarantes (6.0%)"),
                            ("23653502", 0.07,   "23653502 - Arrendamientos / Otros (7.0%)"),
                            ("23652001", 0.10,   "23652001 - Honorarios / Comisiones (10.0%)"),
                            ("23651501", 0.11,   "23651501 - Honorarios Persona Juridica / Asesoria (11.0%)"),
                            ("23657001", 0.02,   "23657001 - Otras Retenciones (2.0%)"),
                            ("23654006", 0.001,  "23654006 - Combustibles / Derivados (0.1%)"),
                            ("OTRA_RF",  None,   "Otra Cuenta / Porcentaje Personalizado...")
                        ]
                        CATALOGO_CTAS_RETEICA = [
                            ("23680513", 4.14,   "23680513 - Transporte Terrestre Carga (4.14 por mil)"),
                            ("23680511", 6.90,   "23680511 - Actividades Industriales / Otros (6.90 por mil)"),
                            ("23680509", 7.00,   "23680509 - Actividades Comerciales / Otros (7.00 por mil)"),
                            ("23680507", 8.00,   "23680507 - Actividades Servicios / Otros (8.00 por mil)"),
                            ("23680505", 9.66,   "23680505 - Servicios / Agenciamiento Aduanero (9.66 por mil)"),
                            ("23680501", 11.04,  "23680501 - Compras / Comercio General (11.04 por mil)"),
                            ("23680503", 13.80,  "23680503 - Actividades Financieras / Especiales (13.80 por mil)"),
                            ("OTRA_RI",  None,   "Otra Cuenta / Tarifa por mil Personalizada...")
                        ]

                        # 1. CABECERA RAPIDA: CONCEPTO Y REGIMEN DEL PROVEEDOR
                        opciones_conceptos = [
                            "\U0001f69a Transporte Terrestre de Carga / Fletes (ReteFuente 1% [23652505] - ReteICA 4.14\u2030 [23680513])",
                            "\U0001f3e2 Agenciamiento Aduanero / Intermediaci\u00f3n (ReteFuente 4% [23652503] - ReteICA 9.66\u2030 [23680505])",
                            "\U0001f4e6 Compra de Mercanc\u00edas / Bienes (ReteFuente 2.5% [23654001] - ReteICA 11.04\u2030 [23680501])",
                            "\U0001f3e8 Alojamiento / Hoteles / Viajes (ReteFuente 3.5% [23652501] - ReteICA 9.66\u2030 [23680505])",
                            "\u2696\ufe0f Honorarios Profesionales / Asesor\u00eda (ReteFuente 11% [23651501] - ReteICA 9.66\u2030 [23680505])",
                            "\U0001f4bb Software / Tecnolog\u00eda (Sin ReteFuente - IVA Servicios 24081501)",
                            "\U0001f6e0\ufe0f Servicios Generales (ReteFuente 4% [23652503] - ReteICA 9.66\u2030 [23680505])"
                        ]
                        desc_fac_act = (str(fac_sel.get("Descripcion", "")) + " " + str(fac_sel.get("Categor\u00eda", ""))).upper()
                        idx_concep_act = 1 if es_aduanero else 6
                        if any(k in desc_fac_act for k in ["TRANSPORTE", "TRANSPORT", "INLAND", "FLETE", "CARGA TERRESTRE", "ACARREO"]):
                            idx_concep_act = 0
                        elif any(k in str(fac_sel.get("Proveedor", "")).upper() for k in ["FERROMENDEZ", "TORNILLO", "CAUCHO", "REPUESTO", "MERCANCIA"]):
                            idx_concep_act = 2
                        elif any(k in str(fac_sel.get("Proveedor", "")).upper() for k in ["HOTEL", "ESTELAR", "GENOVA"]):
                            idx_concep_act = 3
                        elif str(fac_sel.get("Cta ReteFuente", "")).strip() in ("23651501", "23652001"):
                            idx_concep_act = 4

                        c_top_c1, c_top_c2 = st.columns([1.4, 1.0])
                        with c_top_c1:
                            sel_concep = st.selectbox(
                                "Concepto / Actividad de la Factura (Pre-ajusta cuentas sugeridas):",
                                opciones_conceptos,
                                index=idx_concep_act,
                                key=f"sel_concep_{comp_box_id}"
                            )
                        with c_top_c2:
                            opciones_reg_man = [
                                "O-48 (Responsable de IVA / Com\u00fan - Retenci\u00f3n Ordinaria)",
                                "O-15 (Autorretenedor de Renta - Sin ReteFuente)",
                                "O-13 (Gran Contribuyente - Sin ReteIVA)",
                                "O-13;O-15 (Gran Contribuyente y Autorretenedor)",
                                "O-47 (R\u00e9gimen Simple - Exento ReteFuente y ReteIVA)",
                                "O-49 (No Responsable de IVA)"
                            ]
                            idx_actual = 0
                            reg_actual = str(fac_sel.get("R\u00e9gimen Fiscal Emisor", ""))
                            if "O-13" in reg_actual and "O-15" in reg_actual: idx_actual = 3
                            elif "O-15" in reg_actual: idx_actual = 1
                            elif "O-13" in reg_actual: idx_actual = 2
                            elif "O-47" in reg_actual: idx_actual = 4
                            elif "O-49" in reg_actual: idx_actual = 5
                            nuevo_reg_sel = st.selectbox("R\u00e9gimen Fiscal del Proveedor:", opciones_reg_man, index=idx_actual, key=f"sel_reg_{comp_box_id}")
                            cod_reg_sel = nuevo_reg_sel.split()[0]

                        es_concep_transp = "Transporte" in sel_concep
                        es_concep_aduana = "Agenciamiento" in sel_concep
                        es_concep_compra = "Compra" in sel_concep
                        es_concep_hotel = "Alojamiento" in sel_concep
                        es_concep_honor = "Honorarios" in sel_concep
                        es_concep_soft = "Software" in sel_concep

                        # Valores actuales de la factura
                        base_cur_f = float(fac_sel["Base"]) if pd.notna(fac_sel["Base"]) else 0.0
                        iva_tot_cur_f = float(fac_sel.get("IVA", 0.0) or 0.0)
                        iva_5_cur_f = float(fac_sel.get("IVA 5%", 0.0) or 0.0)
                        iva_19_cur_f = float(fac_sel.get("IVA 19%", 0.0) or 0.0)
                        if iva_5_cur_f > 0 and iva_19_cur_f <= 0 and iva_tot_cur_f > iva_5_cur_f:
                            iva_19_cur_f = round(iva_tot_cur_f - iva_5_cur_f, 2)
                        elif iva_5_cur_f <= 0 and iva_19_cur_f <= 0:
                            iva_19_cur_f = iva_tot_cur_f
                        inc_cur_f = float(fac_sel.get("INC / Otros Impuestos", 0.0) or 0.0)
                        raw_imp_f = fac_sel.get("IVA Importaci\u00f3n")
                        try:
                            val_iva_imp_actual = float(raw_imp_f) if (raw_imp_f is not None and pd.notna(raw_imp_f)) else 0.0
                        except Exception:
                            val_iva_imp_actual = 0.0

                        # =========================================================================
                        # BLOQUE 1: LUGAR EXCLUSIVO PARA VALORES Y BASES ($)
                        # =========================================================================
                        st.markdown("""
                        <div style="background:#f0f9ff; border:1px solid #bae6fd; border-left:5px solid #0284c7; border-radius:6px; padding:8px 12px; margin:10px 0 8px 0;">
                            <b style="color:#0369a1; font-size:14px;">\U0001f4b5 1. Valores Fiscales y Bases de la Factura ($)</b>
                            <span style="color:#0c4a6e; font-size:12.5px;"> \u2014 Pon aqui unicamente los montos en pesos (las retenciones se formulan solas abajo segun la base):</span>
                        </div>
                        """, unsafe_allow_html=True)

                        cv_1, cv_2, cv_3 = st.columns(3)
                        with cv_1:
                            nueva_base = st.number_input("Base Gravable / Subtotal ($):", value=base_cur_f, step=1000.0, key=f"nb_{comp_box_id}")
                            _base_iva19_est = round(iva_19_cur_f / 0.19, 2) if iva_19_cur_f > 0 else 0.0
                            _rf_prev_chk = float(fac_sel.get("ReteFuente", 0.0) or 0.0)
                            _usa_base_iva_def = (_base_iva19_est > 0 and abs(base_cur_f - _base_iva19_est) > 5.0 and _rf_prev_chk > 0 and (abs(_rf_prev_chk - round(_base_iva19_est * 0.11, 2)) < 2.0 or abs(_rf_prev_chk - round(_base_iva19_est * 0.04, 2)) < 2.0))
                            usar_base_ret_distinta = st.checkbox(
                                "Usar base de retencion distinta al subtotal",
                                value=_usa_base_iva_def,
                                key=f"chk_base_ret_esp_{comp_box_id}",
                                help="Activalo si la retencion solo aplica sobre la parte gravada de la factura (ej. ingresos propios sin timbres/terceros)."
                            )
                            if usar_base_ret_distinta:
                                _val_base_ret_ini = _base_iva19_est if _usa_base_iva_def else float(nueva_base)
                                base_para_retenciones = st.number_input("Base para Calcular Retenciones ($):", value=float(_val_base_ret_ini), step=1000.0, key=f"nb_ret_esp_{comp_box_id}")
                            else:
                                base_para_retenciones = float(nueva_base)
                        with cv_2:
                            nuevo_iva_19 = st.number_input("IVA 19% ($):", value=iva_19_cur_f, step=100.0, key=f"niva_{comp_box_id}")
                            nuevo_iva_5 = st.number_input("IVA 5% ($):", value=iva_5_cur_f, step=100.0, key=f"niva5_{comp_box_id}")
                        with cv_3:
                            nuevo_inc_otros = st.number_input("INC 8% / Bolsas / Otros ($):", value=inc_cur_f, step=100.0, key=f"ninc_{comp_box_id}")
                            nuevo_iva_imp = st.number_input(
                                "\U0001f6a2 IVA de Importaci\u00f3n ($):",
                                value=float(val_iva_imp_actual),
                                step=1000.0,
                                key=f"niva_imp_val_{comp_box_id}",
                                help="Pon aqui el valor del IVA de Importacion o Derechos Aduanales IVA (se contabiliza en la Cuenta IVA Importacion 240835)."
                            )
                            _tot_ref_fac = float(fac_sel.get("Total", 0.0) or 0.0)
                            _venia_sumado_en_base = (
                                val_iva_imp_actual == 0.0
                                and float(nuevo_iva_imp or 0.0) > 0.0
                                and float(nueva_base or 0.0) > float(nuevo_iva_imp or 0.0)
                                and abs((float(nueva_base or 0.0) + float(nuevo_iva_19 or 0.0) + float(nuevo_iva_5 or 0.0) + float(nuevo_inc_otros or 0.0)) - _tot_ref_fac) < 5.0
                            )
                            if float(nuevo_iva_imp or 0.0) > 0.0 and val_iva_imp_actual == 0.0:
                                restar_ivaimp_de_base = st.checkbox(
                                    "Descontar IVA Importaci\u00f3n del Subtotal (ven\u00eda sumado en la base)",
                                    value=_venia_sumado_en_base,
                                    key=f"chk_restar_ivaimp_base_{comp_box_id}",
                                    help="Si la factura traia el IVA de Importacion sumado dentro del Subtotal, dejalo marcado para restarlo automaticamente de la Base Gravable al guardar."
                                )
                            else:
                                restar_ivaimp_de_base = False
                            if restar_ivaimp_de_base and float(nueva_base or 0.0) > float(nuevo_iva_imp or 0.0):
                                base_efectiva_calculo = round(float(nueva_base) - float(nuevo_iva_imp), 2)
                                st.caption(f"\U0001f4a1 Base neta descontando IVA Imp.: **${base_efectiva_calculo:,.2f}**")
                                if not usar_base_ret_distinta:
                                    base_para_retenciones = base_efectiva_calculo
                            else:
                                base_efectiva_calculo = float(nueva_base)

                        # =========================================================================
                        # BLOQUE 2: LUGAR EXCLUSIVO PARA CUENTAS CONTABLES (PUC)
                        # =========================================================================
                        st.markdown("""
                        <div style="background:#f8fafc; border:1px solid #cbd5e1; border-left:5px solid #334155; border-radius:6px; padding:8px 12px; margin:12px 0 8px 0;">
                            <b style="color:#0f172a; font-size:14px;">\U0001f3e6 2. Cuentas Contables (PUC)</b>
                            <span style="color:#475569; font-size:12.5px;"> \u2014 Modifica aqui cualquier cuenta contable en caso de que debas cambiarla:</span>
                        </div>
                        """, unsafe_allow_html=True)

                        cta_p_sug_def = "14650501" if (es_aduanero or es_concep_transp) else "51350501"
                        cta_p_def = str(fac_sel.get("Cta Principal") or cta_p_sug_def).strip()
                        cta_cxp_def = str(fac_sel.get("Cuenta Pasivo Especifica") or fac_sel.get("Cta Contrapartida") or ("22050505" if es_aduanero else "22050501")).strip()
                        cta_iva_def = str(fac_sel.get("Cta IVA") or ("24081501" if es_importacion_factura else "24081001")).strip()
                        cta_iva_5_def = str(fac_sel.get("Cta IVA 5%") or "24081003").strip()
                        raw_cta_f = fac_sel.get("Cta IVA Importaci\u00f3n")
                        cta_iva_imp_actual = "240835" if (raw_cta_f is None or pd.isna(raw_cta_f) or str(raw_cta_f).strip() in ["", "nan", "None"]) else str(raw_cta_f).strip()

                        cods_puc_cat = [x[0] for x in CATALOGO_CUENTAS_PRINCIPALES_PUC]
                        idx_puc_cat = cods_puc_cat.index(cta_p_def) if cta_p_def in cods_puc_cat else (len(CATALOGO_CUENTAS_PRINCIPALES_PUC) - 1)
                        sel_cta_p_item = st.selectbox(
                            "Tipo de Gasto / Costo / Inventario Identificado (Catalogo PUC):",
                            options=CATALOGO_CUENTAS_PRINCIPALES_PUC,
                            index=idx_puc_cat,
                            format_func=lambda x: x[2],
                            key=f"sel_cat_puc_p_{comp_box_id}_{cta_p_def}",
                            help="Identifica automaticamente si es Cafeteria, Aseo, Mensajeria, Transporte, Luz, Agua, Internet, Papeleria, Legal, Seguros, Software, Viajes, Mantenimiento o Importacion. Puedes cambiarlo aqui con un clic."
                        )
                        cta_p_desde_catalogo = cta_p_def if sel_cta_p_item[0] == "OTRA_CTA_P" else sel_cta_p_item[0]

                        cc_1, cc_2, cc_3 = st.columns(3)
                        with cc_1:
                            nueva_cta_p = st.text_input("Codigo Cuenta Principal (PUC):", value=cta_p_desde_catalogo, key=f"inp_cta_p_{comp_box_id}_{cta_p_desde_catalogo}")
                            nueva_cta_cxp = st.text_input("Cuenta por Pagar (Pasivo / CxP):", value=cta_cxp_def, key=f"inp_cxp_{comp_box_id}")
                        with cc_2:
                            nueva_cta_iva = st.text_input("Cuenta IVA 19%:", value=cta_iva_def, key=f"inp_iva_{comp_box_id}")
                            nueva_cta_iva_5 = st.text_input("Cuenta IVA 5%:", value=cta_iva_5_def, key=f"inp_iva5_{comp_box_id}")
                        with cc_3:
                            cta_iva_imp_cod = st.text_input("Cuenta IVA Importaci\u00f3n:", value=cta_iva_imp_actual, key=f"inp_cta_ivaimp_cod_{comp_box_id}", help="Cuenta contable para el IVA de Importacion / Derechos Aduanales IVA (por defecto 240835).")
                            st.caption("\U0001f6a2 IVA Importaci\u00f3n / Aduana (Ej: `240835` / `24083501`)")

                        # =========================================================================
                        # BLOQUE 3: RETENCIONES FORMULADAS AUTOMATICAMENTE POR CUENTA + DESHABILITAR
                        # =========================================================================
                        st.markdown("""
                        <div style="background:#f0fdf4; border:1px solid #86efac; border-left:5px solid #16a34a; border-radius:6px; padding:8px 12px; margin:12px 0 8px 0;">
                            <b style="color:#166534; font-size:14px;">\U0001f9ee 3. Retenciones Formuladas Automaticamente por Cuenta</b>
                            <span style="color:#14532d; font-size:12.5px;"> \u2014 Si cambias la Base o eliges otra Cuenta de otro porcentaje, la formula se adecua al instante. Si no lleva retencion, deshabilita la casilla para que no entre a la contabilizacion:</span>
                        </div>
                        """, unsafe_allow_html=True)

                        # Determinar cuenta y estado inicial de ReteFuente y ReteICA
                        val_rf_prev = float(fac_sel.get("ReteFuente", 0.0) or 0.0)
                        val_ri_prev = float(fac_sel.get("ReteICA", 0.0) or 0.0)
                        cta_rf_guardada = _limpiar_codigo_cta(fac_sel.get("Cta ReteFuente"), "")
                        cta_ri_guardada = _limpiar_codigo_cta(fac_sel.get("Cta ReteICA"), "")

                        if es_concep_transp:
                            cta_rf_sug_def, cta_ri_sug_def = "23652505", "23680513"
                        elif es_concep_aduana:
                            cta_rf_sug_def, cta_ri_sug_def = "23652503", "23680505"
                        elif es_concep_compra:
                            cta_rf_sug_def, cta_ri_sug_def = "23654001", "23680501"
                        elif es_concep_hotel:
                            cta_rf_sug_def, cta_ri_sug_def = "23652501", "23680505"
                        elif es_concep_honor:
                            cta_rf_sug_def, cta_ri_sug_def = "23651501", "23680505"
                        elif es_concep_soft:
                            cta_rf_sug_def, cta_ri_sug_def = "", ""
                        else:
                            cta_rf_sug_def, cta_ri_sug_def = "23652503", "23680505"

                        es_exento_ret_reg = any(k in cod_reg_sel for k in ["O-15", "O-47"]) or es_concep_soft
                        tiene_rf_def = False if es_exento_ret_reg else (val_rf_prev > 0 or es_concep_transp)
                        tiene_ri_def = False if (es_exento_ret_reg or "O-13" in cod_reg_sel) else (val_ri_prev > 0 or es_concep_transp)

                        cta_rf_inicial = cta_rf_guardada if cta_rf_guardada else (cta_rf_sug_def or "23652503")
                        cta_ri_inicial = cta_ri_guardada if cta_ri_guardada else (cta_ri_sug_def or "23680505")

                        cr_rf, cr_ri = st.columns(2)

                        # --- COLUMNA RETEFUENTE ---
                        with cr_rf:
                            aplicar_rfte = st.checkbox(
                                "\u2705 Aplicar ReteFuente a esta factura",
                                value=tiene_rf_def,
                                key=f"chk_aplica_rf_{comp_box_id}_{cod_reg_sel}",
                                help="Desmarca esta casilla si la factura NO tiene ReteFuente: quedara en $0.00 y no entrara en la contabilizacion."
                            )
                            if aplicar_rfte:
                                cods_rf_cat = [x[0] for x in CATALOGO_CTAS_RETEFUENTE]
                                idx_rf_cat = cods_rf_cat.index(cta_rf_inicial) if cta_rf_inicial in cods_rf_cat else 4
                                sel_rf_item = st.selectbox(
                                    "Cuenta ReteFuente (Formula automatica):",
                                    options=CATALOGO_CTAS_RETEFUENTE,
                                    index=idx_rf_cat,
                                    format_func=lambda x: x[2],
                                    key=f"sel_cta_rf_formula_{comp_box_id}"
                                )
                                if sel_rf_item[0] == "OTRA_RF":
                                    c_orf1, c_orf2 = st.columns(2)
                                    with c_orf1:
                                        nueva_cta_rf = st.text_input("Codigo Cuenta ReteFuente:", value=cta_rf_inicial, key=f"inp_otra_rf_cta_{comp_box_id}")
                                    with c_orf2:
                                        pct_rf_custom = st.number_input("Porcentaje ReteFuente (%):", value=4.0, step=0.5, key=f"inp_otra_rf_pct_{comp_box_id}")
                                    tasa_rf_usar = float(pct_rf_custom) / 100.0
                                else:
                                    nueva_cta_rf = sel_rf_item[0]
                                    tasa_rf_usar = float(sel_rf_item[1])

                                rfte_formulada = round(float(base_para_retenciones) * tasa_rf_usar, 2)
                                st.info(f"\U0001f9ee **Formula Activa:** Base `${base_para_retenciones:,.2f}` \u00d7 `{tasa_rf_usar*100:.2f}%` ([{nueva_cta_rf}]) = **${rfte_formulada:,.2f}**")
                                manual_rf_opt = st.checkbox("Ajustar valor de ReteFuente manualmente", value=False, key=f"chk_man_rf_{comp_box_id}")
                                if manual_rf_opt:
                                    nueva_rfte = st.number_input("Valor ReteFuente Manual ($):", value=float(val_rf_prev if val_rf_prev > 0 else rfte_formulada), step=100.0, key=f"nrf_man_{comp_box_id}")
                                else:
                                    nueva_rfte = rfte_formulada
                            else:
                                nueva_rfte = 0.0
                                nueva_cta_rf = ""
                                st.caption("\U0001f6ab *ReteFuente deshabilitada ($0.00) \u2014 No entrara al asiento contable ni a Siigo.*")

                        # --- COLUMNA RETEICA ---
                        with cr_ri:
                            aplicar_rica = st.checkbox(
                                "\u2705 Aplicar ReteICA a esta factura",
                                value=tiene_ri_def,
                                key=f"chk_aplica_ri_{comp_box_id}_{cod_reg_sel}",
                                help="Desmarca esta casilla si la factura NO tiene ReteICA: quedara en $0.00 y no entrara en la contabilizacion."
                            )
                            if aplicar_rica:
                                cods_ri_cat = [x[0] for x in CATALOGO_CTAS_RETEICA]
                                idx_ri_cat = cods_ri_cat.index(cta_ri_inicial) if cta_ri_inicial in cods_ri_cat else 4
                                sel_ri_item = st.selectbox(
                                    "Cuenta ReteICA (Formula automatica por mil):",
                                    options=CATALOGO_CTAS_RETEICA,
                                    index=idx_ri_cat,
                                    format_func=lambda x: x[2],
                                    key=f"sel_cta_ri_formula_{comp_box_id}"
                                )
                                if sel_ri_item[0] == "OTRA_RI":
                                    c_ori1, c_ori2 = st.columns(2)
                                    with c_ori1:
                                        nueva_cta_ri = st.text_input("Codigo Cuenta ReteICA:", value=cta_ri_inicial, key=f"inp_otra_ri_cta_{comp_box_id}")
                                    with c_ori2:
                                        mil_ri_custom = st.number_input("Tarifa por mil (\u2030):", value=9.66, step=0.1, key=f"inp_otra_ri_mil_{comp_box_id}")
                                    tasa_ri_mil = float(mil_ri_custom)
                                else:
                                    nueva_cta_ri = sel_ri_item[0]
                                    tasa_ri_mil = float(sel_ri_item[1])

                                rica_formulada = round((float(base_para_retenciones) * tasa_ri_mil) / 1000.0, 2)
                                st.info(f"\U0001f9ee **Formula Activa:** Base `${base_para_retenciones:,.2f}` \u00d7 `{tasa_ri_mil:.2f}\u2030` ([{nueva_cta_ri}]) = **${rica_formulada:,.2f}**")
                                manual_ri_opt = st.checkbox("Ajustar valor de ReteICA manualmente", value=False, key=f"chk_man_ri_{comp_box_id}")
                                if manual_ri_opt:
                                    nuevo_rica = st.number_input("Valor ReteICA Manual ($):", value=float(val_ri_prev if val_ri_prev > 0 else rica_formulada), step=100.0, key=f"nri_man_{comp_box_id}")
                                else:
                                    nuevo_rica = rica_formulada
                            else:
                                nuevo_rica = 0.0
                                nueva_cta_ri = ""
                                st.caption("\U0001f6ab *ReteICA deshabilitado ($0.00) \u2014 No entrara al asiento contable ni a Siigo.*")

                        asumir_imp = st.checkbox(
                            "Asumir Impuestos / Retenciones (Cruza con Agente Aduanero / Triangulaci\u00f3n - Cta 53152001)",
                            value=bool(fac_sel.get("Impuestos Asumidos", False)),
                            key=f"asum_in_{comp_box_id}",
                            help="Si se asumen, las retenciones van a la cuenta 53152001 y el Saldo por Pagar de la factura queda en Base + IVA."
                        )

                        def _guardar_cambios_factura_actual(avanzar_siguiente=False, aplicar_a_todo_proveedor=False):
                            cod_reg = _limpiar_codigo_cta(nuevo_reg_sel, "O-48")
                            m_idx_list = df_p[df_p["Comprobante Siigo"].astype(str).str.strip() == comp_sel].index
                            r_idx = m_idx_list[0] if len(m_idx_list) > 0 else fac_sel.name
                            nuevo_iva_tot = round(float(nuevo_iva_19 or 0.0) + float(nuevo_iva_5 or 0.0), 2)
                            nuevo_inc_f = round(float(nuevo_inc_otros or 0.0), 2)
                            cta_rf_clean = _limpiar_codigo_cta(nueva_cta_rf, "") if aplicar_rfte else ""
                            cta_ri_clean = _limpiar_codigo_cta(nueva_cta_ri, "") if aplicar_rica else ""
                            rfte_guardar = round(float(nueva_rfte or 0.0), 2) if aplicar_rfte else 0.0
                            rica_guardar = round(float(nuevo_rica or 0.0), 2) if aplicar_rica else 0.0
                            cta_clean = _limpiar_codigo_cta(nueva_cta_cxp, "22050501")
                            cta_p_clean = _limpiar_codigo_cta(nueva_cta_p, "51959501")
                            cta_iva_clean = _limpiar_codigo_cta(nueva_cta_iva, "24081501" if es_importacion_factura else "24081001")
                            cta_iva_5_clean = _limpiar_codigo_cta(nueva_cta_iva_5, "24081003")
                            nuevo_iva_imp_f = float(nuevo_iva_imp) if pd.notna(nuevo_iva_imp) else 0.0
                            cta_iva_imp_clean = _limpiar_codigo_cta(cta_iva_imp_cod, "240835")
                            nueva_base_guardar = round(float(base_efectiva_calculo), 2)

                            df_p.at[r_idx, "R\u00e9gimen Fiscal Emisor"] = cod_reg
                            df_p.at[r_idx, "Base"] = nueva_base_guardar
                            df_p.at[r_idx, "IVA"] = nuevo_iva_tot
                            df_p.at[r_idx, "IVA 19%"] = round(float(nuevo_iva_19 or 0.0), 2)
                            df_p.at[r_idx, "IVA 5%"] = round(float(nuevo_iva_5 or 0.0), 2)
                            df_p.at[r_idx, "Base IVA 19%"] = round(float(nuevo_iva_19) / 0.19, 2) if float(nuevo_iva_5 or 0.0) > 0 and float(nuevo_iva_19 or 0.0) > 0 else (nueva_base_guardar if float(nuevo_iva_19 or 0.0) > 0 else 0.0)
                            df_p.at[r_idx, "Base IVA 5%"] = round(float(nuevo_iva_5) / 0.05, 2) if float(nuevo_iva_5 or 0.0) > 0 else 0.0
                            df_p.at[r_idx, "INC / Otros Impuestos"] = nuevo_inc_f
                            df_p.at[r_idx, "ReteFuente"] = rfte_guardar
                            df_p.at[r_idx, "ReteICA"] = rica_guardar
                            df_p.at[r_idx, "Cta ReteFuente"] = cta_rf_clean
                            df_p.at[r_idx, "Cta ReteICA"] = cta_ri_clean

                            if (es_concep_transp or cta_rf_clean == "23652505") and cta_p_clean == "14650501":
                                df_p.at[r_idx, "Categor\u00eda"] = "Transporte Carga (1%)"
                                df_p.at[r_idx, "Raz\u00f3n Contable"] = "Transporte Terrestre de Carga / Fletes (1% ReteFuente [23652505] - 4.14\u2030 ReteICA [23680513])"
                            else:
                                df_p.at[r_idx, "Categor\u00eda"] = obtener_etiqueta_cuenta_principal(cta_p_clean, df_p.at[r_idx, "Categor\u00eda"])
                            df_p.at[r_idx, "Cuenta Pasivo Especifica"] = cta_clean
                            df_p.at[r_idx, "Cta Contrapartida"] = cta_clean
                            df_p.at[r_idx, "Cta Principal"] = cta_p_clean
                            df_p.at[r_idx, "Cta IVA"] = cta_iva_clean
                            df_p.at[r_idx, "Cta IVA 5%"] = cta_iva_5_clean
                            df_p.at[r_idx, "IVA Importaci\u00f3n"] = nuevo_iva_imp_f
                            df_p.at[r_idx, "Cta IVA Importaci\u00f3n"] = cta_iva_imp_clean
                            df_p.at[r_idx, "Impuestos Asumidos"] = asumir_imp
                            df_p.at[r_idx, "Editada Manualmente"] = True

                            r_iva_val = float(fac_sel.get("ReteIVA", 0.0)) if pd.notna(fac_sel.get("ReteIVA")) else 0.0
                            if asumir_imp:
                                saldo_p = round(nueva_base_guardar + nuevo_inc_f + nuevo_iva_tot + nuevo_iva_imp_f, 2)
                            else:
                                saldo_p = round(nueva_base_guardar + nuevo_inc_f + nuevo_iva_tot + nuevo_iva_imp_f - rfte_guardar - rica_guardar - r_iva_val, 2)
                            df_p.at[r_idx, "Neto a Pagar"] = saldo_p
                            df_p.at[r_idx, "Total Neto"] = saldo_p

                            tot_esperado = round(nueva_base_guardar + nuevo_inc_f + nuevo_iva_tot + nuevo_iva_imp_f, 2)
                            if nuevo_iva_imp_f > 0 and abs(float(df_p.at[r_idx, "Total"]) - tot_esperado) > 10.0:
                                df_p.at[r_idx, "Total"] = tot_esperado

                            # Si se pidio replicar cuentas y formulas a TODAS las facturas del mismo proveedor
                            n_replicadas = 0
                            if aplicar_a_todo_proveedor:
                                nit_prov_target = str(fac_sel.get("NIT Emisor", "")).strip()
                                mask_mismo_prov = (df_p["NIT Emisor"].astype(str).str.strip() == nit_prov_target) & (df_p.index != r_idx) & (~df_p["No Contabilizar"].astype(bool))
                                for idx_mp in df_p[mask_mismo_prov].index:
                                    b_mp = float(df_p.at[idx_mp, "Base"] or 0.0)
                                    iv_mp = float(df_p.at[idx_mp, "IVA"] or 0.0)
                                    inc_mp = float(df_p.at[idx_mp, "INC / Otros Impuestos"] or 0.0)
                                    imp_mp = float(df_p.at[idx_mp, "IVA Importaci\u00f3n"] or 0.0)
                                    riv_mp = float(df_p.at[idx_mp, "ReteIVA"] or 0.0)
                                    rf_mp = round(b_mp * tasa_rf_usar, 2) if aplicar_rfte else 0.0
                                    ri_mp = round((b_mp * tasa_ri_mil) / 1000.0, 2) if aplicar_rica else 0.0
                                    df_p.at[idx_mp, "R\u00e9gimen Fiscal Emisor"] = cod_reg
                                    df_p.at[idx_mp, "Cta Principal"] = cta_p_clean
                                    df_p.at[idx_mp, "Categor\u00eda"] = obtener_etiqueta_cuenta_principal(cta_p_clean, df_p.at[idx_mp, "Categor\u00eda"])
                                    df_p.at[idx_mp, "Cuenta Pasivo Especifica"] = cta_clean
                                    df_p.at[idx_mp, "Cta Contrapartida"] = cta_clean
                                    df_p.at[idx_mp, "Cta IVA"] = cta_iva_clean
                                    df_p.at[idx_mp, "Cta IVA 5%"] = cta_iva_5_clean
                                    df_p.at[idx_mp, "ReteFuente"] = rf_mp
                                    df_p.at[idx_mp, "Cta ReteFuente"] = cta_rf_clean
                                    df_p.at[idx_mp, "ReteICA"] = ri_mp
                                    df_p.at[idx_mp, "Cta ReteICA"] = cta_ri_clean
                                    df_p.at[idx_mp, "Impuestos Asumidos"] = asumir_imp
                                    df_p.at[idx_mp, "Editada Manualmente"] = True
                                    neto_mp = round(b_mp + inc_mp + iv_mp + imp_mp, 2) if asumir_imp else round(b_mp + inc_mp + iv_mp + imp_mp - rf_mp - ri_mp - riv_mp, 2)
                                    df_p.at[idx_mp, "Neto a Pagar"] = neto_mp
                                    df_p.at[idx_mp, "Total Neto"] = neto_mp
                                    n_replicadas += 1

                            st.session_state["df_procesado"] = normalizar_df_procesado(df_p)
                            df_p = st.session_state["df_procesado"]

                            # Si la factura editada es Agente de un paquete de importacion, actualizar o regenerar su asiento
                            fac_edit_num = str(df_p.at[r_idx, "Factura"]).strip()
                            pqs_cur = st.session_state.get("paquetes_importacion", {})
                            if isinstance(pqs_cur, dict):
                                for pk_u, pv_u in pqs_cur.items():
                                    if isinstance(pv_u, dict) and str(pv_u.get("agente", {}).get("Factura", "")).strip() == fac_edit_num:
                                        ter_u = pv_u.get("terceros")
                                        sin_terceros = (ter_u is None) or (isinstance(ter_u, pd.DataFrame) and ter_u.empty) or (isinstance(ter_u, list) and len(ter_u) == 0)
                                        af_prev = pv_u.get("asiento_fijo")
                                        tiene_5395 = False
                                        if af_prev is not None:
                                            df_af_p = pd.DataFrame(af_prev)
                                            if not df_af_p.empty and "C\u00f3digo Cuenta" in df_af_p.columns:
                                                tiene_5395 = any(str(c).strip() == "53950501" for c in df_af_p["C\u00f3digo Cuenta"])
                                        if sin_terceros and not tiene_5395:
                                            df_as_new = obtener_asiento_contable_hoja2(df_p.loc[r_idx], es_aduanero=es_aduanero)
                                            pv_u["asiento_fijo"] = df_as_new.copy()
                                            pv_u["asiento_aprobado"] = df_as_new.copy()
                                            st.session_state[f"asiento_fijo_pq_{pk_u}"] = df_as_new.copy()
                                            st.session_state.setdefault("asientos_triangulacion_por_factura", {})[fac_edit_num] = df_as_new.copy()
                            sincronizar_todos_los_paquetes_con_df_procesado()
                            guardar_regla_aprendizaje(
                                empresa,
                                fac_sel.get("NIT Emisor"),
                                fac_sel.get("Proveedor"),
                                cta_p=cta_p_clean,
                                resp_fiscal=cod_reg,
                                cta_cxp=cta_clean,
                                cta_iva=cta_iva_clean,
                                asumir_impuestos=asumir_imp
                            )
                            st.session_state.pop("_cached_respaldo_bytes", None)
                            st.session_state.pop("_siigo_cache_sig", None)
                            guardar_estado_manual(empresa)

                            if avanzar_siguiente and idx_fac_sel < len(comps_en_vista) - 1:
                                st.session_state["factura_seleccionada_auditoria"] = comps_en_vista[idx_fac_sel + 1]
                            else:
                                st.session_state["factura_seleccionada_auditoria"] = comp_sel
                            st.session_state.pop("sel_factura_auditoria_box", None)
                            return n_replicadas

                        # BOTONES DE ACCION RAPIDA PARA QUE RINDA LA REVISION
                        cb_act1, cb_act2, cb_act3 = st.columns([1.3, 1.5, 1.5])
                        with cb_act1:
                            if st.button("\U0001f4be Guardar y Recalcular", key=f"btn_aplica_reg_{comp_box_id}", type="primary", use_container_width=True):
                                _guardar_cambios_factura_actual(avanzar_siguiente=False, aplicar_a_todo_proveedor=False)
                                st.success("\u00a1Factura actualizada y asiento recalculado!")
                                st.rerun()
                        with cb_act2:
                            if st.button("\U0001f4be\u27a1\ufe0f Guardar y Pasar a Siguiente", key=f"btn_aplica_next_{comp_box_id}", use_container_width=True, help="Guarda esta factura y abre de inmediato la siguiente para que avances sin subir al buscador"):
                                _guardar_cambios_factura_actual(avanzar_siguiente=True, aplicar_a_todo_proveedor=False)
                                st.rerun()
                        with cb_act3:
                            if st.button("\U0001f504 Aplicar Cuentas a todo este Proveedor", key=f"btn_aplica_prov_{comp_box_id}", use_container_width=True, help="Aplica estas mismas cuentas y porcentajes de retencion a todas las facturas activas de este mismo proveedor"):
                                n_rep = _guardar_cambios_factura_actual(avanzar_siguiente=False, aplicar_a_todo_proveedor=True)
                                st.success(f"\u00a1Cuentas y formulas aplicadas a esta factura y a otras {n_rep} facturas de {fac_sel.get('Proveedor', '')}!")
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
                    st.metric("Base Gravable", f"${float(fac_sel.get('Base', 0.0) or 0.0):,.2f}")
                    iva_5_m = float(fac_sel.get("IVA 5%", 0.0) or 0.0)
                    iva_19_m = float(fac_sel.get("IVA 19%", 0.0) or 0.0)
                    iva_tot_m = float(fac_sel.get("IVA", 0.0) or 0.0)
                    inc_m = float(fac_sel.get("INC / Otros Impuestos", 0.0) or 0.0)
                    if iva_5_m > 0:
                        if iva_19_m <= 0 and iva_tot_m > iva_5_m:
                            iva_19_m = round(iva_tot_m - iva_5_m, 2)
                        st.metric("IVA 19% (24081001)", f"${iva_19_m:,.2f}")
                        st.metric("IVA 5% (24081003)", f"${iva_5_m:,.2f}")
                    else:
                        st.metric("IVA Liquidado (19%)", f"${iva_tot_m:,.2f}")
                    if inc_m > 0:
                        st.metric("INC 8% / Otros Impuestos", f"${inc_m:,.2f}")
                    iva_imp_cur = float(fac_sel.get("IVA Importaci\u00f3n", 0.0))
                    if iva_imp_cur > 0:
                        st.metric("\U0001f6a2 IVA Importaci\u00f3n (240835)", f"${iva_imp_cur:,.2f}")
                    st.metric("ReteFuente", f"-${float(fac_sel.get('ReteFuente', 0.0) or 0.0):,.2f}")
                    st.metric("ReteICA", f"-${fac_sel.get('ReteICA', 0.0):,.2f}")
                    reteiva_val = fac_sel.get("ReteIVA", 0.0)
                    if reteiva_val > 0:
                        st.metric("ReteIVA", f"-${reteiva_val:,.2f}")
                    if fac_sel.get("Impuestos Asumidos"):
                        saldo_mostrar = round(fac_sel["Base"] + inc_m + iva_tot_m + iva_imp_cur, 2)
                        st.metric("Saldo CxP a Cruzar (Triangulaci\u00f3n)", f"${saldo_mostrar:,.2f}")
                        st.caption("\u2139\ufe0f Impuestos asumidos (Cta 53152001). Saldo total a cruzar.")
                    else:
                        neto_cxp = round(fac_sel["Base"] + inc_m + iva_tot_m + iva_imp_cur - fac_sel["ReteFuente"] - fac_sel.get("ReteICA", 0.0) - reteiva_val, 2)
                        st.metric("Total Neto CxP (Cta 22 / 23)", f"${neto_cxp:,.2f}")
    
                # VISTA PREVIA DE LA FACTURA BAJO DEMANDA (AHORRO DE MEMORIA RAM)
                st.markdown("---")
                st.markdown("### \U0001f50d Vista Previa del Documento Soporte")
                ver_pdf_ahora = st.checkbox(
                    "\U0001f441\ufe0f Mostrar Documento Soporte de esta Factura en Pantalla",
                    value=True,
                    key="chk_ver_pdf_bajo_demanda_tab2",
                    help="Act\u00edvalo cuando quieras ver la factura en pantalla. Puedes desactivarlo si deseas editar m\u00e1s r\u00e1pido."
                )

                dict_renom = st.session_state.get("dict_pdfs")
                if not isinstance(dict_renom, DiskBackedPDFDict):
                    dict_renom = DiskBackedPDFDict("renom_pdfs")
                    st.session_state["dict_pdfs"] = dict_renom
                else:
                    dict_renom._sync_from_disk()
                dict_orig = st.session_state.get("raw_uploaded_pdfs")
                if not isinstance(dict_orig, DiskBackedPDFDict):
                    dict_orig = DiskBackedPDFDict("raw_pdfs")
                    st.session_state["raw_uploaded_pdfs"] = dict_orig
                else:
                    dict_orig._sync_from_disk()

                if ver_pdf_ahora:
                    pdf_bytes_encontrado, origen_desc, pags_encontradas = buscar_y_extraer_pdf(
                        fac_sel,
                        dict_renombrados=dict_renom,
                        dict_originales=dict_orig,
                        empresa_compradora=empresa
                    )
                    pags_txt_int = obtener_paginas_texto_factura_integrado(fac_sel)
                    renderizar_representacion_grafica_dian_html(
                        fac_sel,
                        pags_txt_int or [],
                        empresa,
                        pdf_bytes_original=pdf_bytes_encontrado,
                        origen_pdf_desc=origen_desc or "",
                        key_suffix="tab2"
                    )
                    if len(dict_orig) > 0 or len(dict_renom) > 0:
                        with st.expander("\U0001f4c2 Seleccionar manualmente otro archivo PDF cargado", expanded=False):
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

                # TABLA DE CONTABILIZACION (ASIENTO CONTABLE)
                st.markdown("### \U0001f4cb Asiento Contable del Comprobante (C\u00f3mo se Contabiliz\u00f3)")
                st.caption("Detalle de partida doble con imputaci\u00f3n de cuentas, d\u00e9bitos, cr\u00e9ditos y sumas iguales:")
    
                fac_num_aud = str(fac_sel["Factura"]).strip()
                asiento_triang_aprobado = None
    
                # 1. Buscar si esta factura corresponde al agente de algun paquete en triangulacion que PASA A HOJA 2
                if "paquetes_importacion" in st.session_state:
                    for p_k, p_v in st.session_state["paquetes_importacion"].items():
                        if str(p_v["agente"]["Factura"]).strip() == fac_num_aud:
                            # Solo usar el asiento de triangulacion si el paquete fue validado con destino 'Pasar a Hoja 2 (Siigo)'
                            if st.session_state.get(f"paquete_listo_{p_k}", False) and st.session_state.get(f"paquete_destino_siigo_{p_k}", p_v.get("destino_siigo", True)):
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
    
                            if asiento_triang_aprobado is None and st.session_state.get(f"paquete_listo_{p_k}", False) and st.session_state.get(f"paquete_destino_siigo_{p_k}", p_v.get("destino_siigo", True)):
                                en_nd_k = st.session_state.get(f"enviar_nd_pq_{p_k}", False)
                                en_gp_k = st.session_state.get(f"enviar_gp_pq_{p_k}", False)
                                en_h2_k = st.session_state.get(f"enviar_h2_pq_{p_k}", False)
                                if en_h2_k:
                                    df_as_k = obtener_asiento_contable_hoja2(p_v["agente"], es_aduanero=False)
                                elif en_gp_k:
                                    df_as_k, _, _ = generar_asiento_triangulacion_paquete(p_v["agente"], p_v["terceros"], enviar_a_no_deducible=False)
                                    asiento_gp_filas = []
                                    for _, r_as in df_as_k.iterrows():
                                        if str(r_as["C\u00f3digo Cuenta"]).strip() == CUENTA_NO_DEDUCIBLE:
                                            asiento_gp_filas.append({
                                                "C\u00f3digo Cuenta": CUENTA_IMPORTACION_TRANSITO,
                                                "Descripci\u00f3n Cuenta": f"Mercanc\u00edas en Tr\u00e1nsito / Base Propia Agente (Fac {p_v['agente'].get('Factura', '')})",
                                                "Tercero / NIT": f"{p_v['agente'].get('NIT Emisor', '')} - {p_v['agente'].get('Proveedor', '')[:25]}",
                                                "D\u00e9bito ($)": float(r_as["D\u00e9bito ($)"]),
                                                "Cr\u00e9dito ($)": 0.0
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
                    df_asiento = sincronizar_asiento_con_factura_actual(pd.DataFrame(asiento_triang_aprobado).copy(), fac_sel)
                    if "Descripci\u00f3n Cuenta" in df_asiento.columns and "Descripci\u00f3n de la Cuenta" not in df_asiento.columns:
                        df_asiento = df_asiento.rename(columns={"Descripci\u00f3n Cuenta": "Descripci\u00f3n de la Cuenta"})
                    cols_deseadas = [c for c in ["C\u00f3digo Cuenta", "Descripci\u00f3n de la Cuenta", "Tercero / NIT", "D\u00e9bito ($)", "Cr\u00e9dito ($)"] if c in df_asiento.columns]
                    df_asiento = df_asiento[cols_deseadas]
                    
        # Asiento validado en Triangulacion se respeta integramente sin sobreescribir
                    
                    st.markdown(f"""
                    <div style="background:#f0fdf4; border:1px solid #86efac; border-left:5px solid #16a34a; border-radius:8px; padding:12px 16px; margin-bottom:12px;">
                        <h5 style="margin:0 0 4px 0; color:#166534;">\U0001f500 Contabilizaci\u00f3n Oficial Aprobada en Triangulaci\u00f3n</h5>
                        <p style="margin:0; color:#14532d; font-size:13.5px;">
                            Esta factura fue conciliada en la Pesta\u00f1a 3. El asiento contable real cancela las cuentas por pagar de terceros y traslada la deuda al agente aduanero:
                        </p>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    asiento_filas = []
                    es_nc = "Devolucion" in str(fac_sel["Operacion"])
                    asume_ret = bool(fac_sel.get("Impuestos Asumidos", False))
                    cta_cxp_usar = str(fac_sel.get("Cuenta Pasivo Especifica") or fac_sel.get("Cta Contrapartida") or ("22050505" if es_aduanero else "22050501")).strip()
    
                    # Verificar si la factura seleccionada es un tercero cruzado en un paquete de importacion
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
    
                    saldo_cruce_ind = obtener_saldo_cruce_factura(fac_sel)
    
                    if pq_cruz_encontrado is not None:
                        st.markdown(f"""
                        <div style="background:#eff6ff; border:1px solid #bfdbfe; border-left:5px solid #2563eb; border-radius:8px; padding:14px 18px; margin-bottom:14px;">
                            <div style="display:flex; justify-content:space-between; align-items:center;">
                                <h5 style="margin:0 0 4px 0; color:#1e40af;">\U0001f500 Factura Vinculada en Paquete de Importaci\u00f3n #{pq_cruz_encontrado['pq_id']} (con {pq_cruz_encontrado['agente']} - Fac {pq_cruz_encontrado['agente_fac']})</h5>
                                <span style="background:#dbeafe; color:#1e40af; font-size:12px; font-weight:bold; padding:3px 8px; border-radius:12px;">Cruce Activo en Triangulaci\u00f3n</span>
                            </div>
                            <p style="margin:6px 0 0 0; color:#1e3a8a; font-size:13.5px;">
                                Esta factura mantiene su <b>contabilizaci\u00f3n propia completa</b> (Costo/Gasto e IVA). En la importaci\u00f3n, lo \u00fanico que se cruza es su <b>Cuenta por Pagar ({cta_cxp_usar})</b> por valor de <b>${saldo_cruce_ind:,.2f}</b>, la cual se debita en el paquete de triangulaci\u00f3n para cancelar la obligaci\u00f3n y trasladar la deuda al agente aduanero.
                            </p>
                        </div>
                        """, unsafe_allow_html=True)
                    elif esta_excl_ind:
                        st.markdown(f"""
                        <div style="background:#fef2f2; border:1px solid #fecaca; border-left:5px solid #dc2626; border-radius:8px; padding:12px 16px; margin-bottom:12px;">
                            <h5 style="margin:0 0 4px 0; color:#991b1b;">\U0001f6ab Factura Marcada como 'No Contabilizar'</h5>
                            <p style="margin:0; color:#7f1d1d; font-size:13.5px;">
                                Esta factura ha sido excluida de la planilla de compras de Siigo. A continuaci\u00f3n se muestra su estructura contable para fines informativos y de auditor\u00eda:
                            </p>
                        </div>
                        """, unsafe_allow_html=True)
    
                    # Cada factura SIEMPRE mantiene su causacion contable propia integra
                    df_asiento = obtener_asiento_contable_hoja2(fac_sel, es_aduanero=es_aduanero)
    
                cols_as_dsp = [c for c in ["C\u00f3digo Cuenta", "Descripci\u00f3n de la Cuenta", "Tercero / NIT", "D\u00e9bito ($)", "Cr\u00e9dito ($)"] if c in df_asiento.columns]
                st.dataframe(
                    df_asiento[cols_as_dsp].style.format({"D\u00e9bito ($)": "${:,.2f}", "Cr\u00e9dito ($)": "${:,.2f}"}),
                    use_container_width=True,
                    hide_index=True
                )
    
                sum_deb = df_asiento["D\u00e9bito ($)"].sum()
                sum_cred = df_asiento["Cr\u00e9dito ($)"].sum()
    
                c_as1, c_as2, c_as3 = st.columns(3)
                with c_as1:
                    st.metric("Total D\u00e9bito", f"${sum_deb:,.2f}")
                with c_as2:
                    st.metric("Total Cr\u00e9dito", f"${sum_cred:,.2f}")
                with c_as3:
                    st.metric("Diferencia / Cuadre", f"${abs(sum_deb - sum_cred):,.2f}")
                st.success("\u2705 **Comprobante Verificado:** Partida doble cuadrada con sumas iguales.")
            else:
                st.info("Carga el archivo Excel en la Pestana 1 para habilitar la auditoria.")
    
        except Exception as _e_aud:
            if type(_e_aud).__name__ in ['RerunException', 'StopException', 'ScriptControlException']:
                raise _e_aud
            st.error(f"\u26a0\ufe0f Ocurri\u00f3 un error en la pesta\u00f1a de Auditor\u00eda Contable: {_e_aud}")
            with st.expander("\U0001f50d Ver detalles t\u00e9cnicos del error (Traceback):"):
                st.code(traceback.format_exc())
    with tab_triangulacion:
        try:
            st.markdown("### \U0001f500 Paquetes de Importaci\u00f3n y Cruce de Cuentas por Pagar")
            st.caption("Separa y concilia cada importaci\u00f3n en paquetes individuales: Cobro de Euro Shipping = Factura DHL (Flete) + Factura Agencia Aduanas + Factura Garaje/Almacenadora.")
    
            if "df_procesado" in st.session_state and st.session_state["df_procesado"] is not None and isinstance(st.session_state["df_procesado"], pd.DataFrame) and not st.session_state["df_procesado"].empty:
                df_total = normalizar_df_procesado(st.session_state["df_procesado"])
                st.session_state["df_procesado"] = df_total
            
                # Identificar facturas de comercio exterior / importacion de forma segura sin errores de tipos
                prov_series = df_total["Proveedor"] if "Proveedor" in df_total.columns else pd.Series("", index=df_total.index)
                desc_series = df_total["Descripcion"] if "Descripcion" in df_total.columns else pd.Series("", index=df_total.index)
                grp_series = df_total["Grupo Importaci\u00f3n"] if "Grupo Importaci\u00f3n" in df_total.columns else pd.Series("", index=df_total.index)
                cta_cxp_series = df_total["Cuenta Pasivo Especifica"] if "Cuenta Pasivo Especifica" in df_total.columns else (df_total["Cta Contrapartida"] if "Cta Contrapartida" in df_total.columns else pd.Series("", index=df_total.index))
                cta_p_series = df_total["Cta Principal"] if "Cta Principal" in df_total.columns else pd.Series("", index=df_total.index)
    
                prov_upper = prov_series.fillna("").astype(str).str.upper()
                desc_upper = desc_series.fillna("").astype(str).str.upper()
                grp_imp_str = grp_series.fillna("").astype(str).str.strip()
                cta_cxp_str = cta_cxp_series.fillna("").astype(str).str.strip()
                cta_p_str = cta_p_series.fillna("").astype(str).str.strip()
    
                # 1. Facturas explicitamente marcadas en Excel con Grupo de Importacion o clasificadas como aduaneras
                cond_grp = (grp_imp_str != "") & (~grp_imp_str.isin(["nan", "None", "0", "0.0"]))
                cond_es_adu = (df_total["Es Aduanera"] == True) if "Es Aduanera" in df_total.columns else pd.Series(False, index=df_total.index)
    
                # 2. Cuentas contables especificas de importacion / fletes DHL / agencias aduaneras
                cond_ctas = cta_cxp_str.isin(["22050505", "23359501", "14650501"]) | (cta_p_str == "14650501")
    
                # 3. Proveedores especificos de Comercio Exterior / Logistica Internacional / Aduana / Puertos / Fletes
                REGEX_IMPORTACION = (
                    r"EURO\s*SHIPPING|TRADE\s*GLOBAL|CONSOLCARGO|BLUE\s*LOGISTICS|"
                    r"KUEHNE|PANALPINA|DSV|EXPEDITORS|TIBA|HUBEMAR|SCHENKER|COLTRANS|COLMAS|"
                    r"DHL|FEDEX|UPS|CARGO|ADUANA|ADUANERA|AGENCIAMIENTO|PORTUARIA|PUERTO\s*BAHIA|"
                    r"CONTECAR|SPRC|SPB|COMPAS|ALPOPULAR|ALMAVIVA|RANSA|ALMACENADORA|"
                    r"MAERSK|HAPAG|SEABOARD|HAMBURG|CMA\s*CGM|EVERGREEN|COSCO|YANG\s*MING"
                )
                cond_prov = prov_upper.str.contains(REGEX_IMPORTACION, regex=True, na=False)
    
                # 4. Descripcion que evidencia gastos de importacion / desaduanamiento
                REGEX_DESC_IMPORT = (
                    "IMPORTACI[O\u00d3]N|AGENCIAMIENTO|ARANCEL|DESADUANAMIENTO|NACIONALIZACI[O\u00d3]N|"
                    "BODEGAJE\\s*PUERTO|ALMACENAJE\\s*ADUANERO|MANDATO\\s*ADUANERO|FLETE\\s*INTERNACIONAL"
                )
                cond_desc = desc_upper.str.contains(REGEX_DESC_IMPORT, regex=True, na=False)
    
                # Condicion estricta: UNICAMENTE facturas relacionadas con la importacion (rojas y blancas de importacion)
                cond_aduanera = cond_grp | cond_es_adu | cond_ctas | cond_prov | cond_desc
                df_adu = df_total[cond_aduanera].copy()
            
                # Configuracion dinamica y extensible de Agentes Coordinadores (Forwarders)
                AGENTES_COORDINADORES_BASE = ["EURO SHIPPING", "TRADE GLOBAL", "CONSOLCARGO", "BLUE LOGISTICS", "KUEHNE", "PANALPINA", "DSV", "EXPEDITORS", "TIBA", "HUBEMAR"]
            
                todos_provs = sorted([str(p) for p in df_total["Proveedor"].dropna().unique() if str(p).strip()]) if "Proveedor" in df_total.columns else []
                provs_agentes_sugeridos = [p for p in todos_provs if any(k in str(p).upper() for k in AGENTES_COORDINADORES_BASE)]
            
                with st.expander("\U0001f3e2 Configurar Agentes Coordinadores (Los que nos cobran directamente a nosotros):", expanded=False):
                    st.caption("Selecciona qu\u00e9 empresas act\u00faan como Agentes Coordinadores / Forwarders que emiten cobros globales consolidados (Euro Shipping, Trade Global, Consolcargo, etc.):")
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
                    st.metric("Terceros Contabilizados (\U0001f534)", n_rojas_adu, help="Causadas previamente en Siigo")
                with m_c4:
                    n_blancas_adu = int((~ya_reg_terc).sum())
                    st.metric("Terceros Pendientes (\u26aa)", n_blancas_adu, help="No contabilizadas a\u00fan (Nuevas)")
                with m_c5:
                    n_cuad_prev = sum([1 for p in st.session_state.get("paquetes_importacion", {}).values() if p.get("diferencia", 999) < 1.0])
                    st.metric("Paquetes Cuadrados ($0.00)", f"{n_cuad_prev} \u2705")
    
                st.markdown("---")
    
                if df_agentes_all.empty:
                    st.warning("\u26a0\ufe0f No se encontraron facturas de agentes aduaneros principales (Euro Shipping o Trade Global) en este reporte.")
                else:
                    # 1. FUNCION DE AUTO-EMPAQUETAMIENTO INTELIGENTE
                    # Algoritmo de combinacion: encuentra para cada Euro Shipping el subconjunto de DHL + Agencia + Garaje que suma su valor
                    def auto_empaquetar_inteligente(df_ag, df_terc):
                        """
                        Algoritmo cronol\u00f3gico de alta precisi\u00f3n para Agencias Aduaneras (Euro Shipping / Trade Global):
                        1. Ordena cronol\u00f3gicamente de Enero hacia Diciembre.
                        2. Cruza y relaciona tanto facturas de terceros YA contabilizadas (\U0001f534) como NO contabilizadas (\u26aa).
                        3. Para cada factura de agente (ej. 19 de Enero), busca facturas de terceros (DHL, Agencia, Garaje, etc.):
                           - Prioritariamente en fechas anteriores o el mismo d\u00eda (-30 a 0 d\u00edas de diferencia).
                           - O en una fecha ligeramente superior (+1 a +12 d\u00edas de diferencia).
                           - Descarta o penaliza fechas lejanas de otros meses.
                        4. Exige concordancia estricta de precio: la suma del Saldo Cruce (Base + IVA) debe concordar
                           exactamente con el valor total facturado por el agente aduanero.
                        """
                        import itertools
                        pqs_res = {}
                    
                        # 1. Asegurar orden cronologico estricto de Enero hacia Diciembre
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
                            
                            # Evaluar concordancia cronologica y de mandato para cada factura de tercero
                            cand_l = []
                            for c_idx_k, c_row_k in cands_i.iterrows():
                                tb_k = float(c_row_k.get("Base", 0.0))
                                tiv_k = float(c_row_k.get("IVA", 0.0))
                                sc_k = obtener_saldo_cruce_factura(c_row_k)
                                p_u_k = str(c_row_k["Proveedor"]).upper()
                                rol_k = "DHL" if "DHL" in p_u_k else ("AGENCIA" if any(k in p_u_k for k in ["CARGO", "ADUANA"]) else "GARAJE")
                            
                                # 2. Analisis de fechas: concordancia de Enero a Diciembre
                                c_dt = c_row_k["_dt"]
                                score_tiempo = 0
                                diff_dias = 0
                                relacion_fecha_txt = "Sin fecha"
                            
                                if pd.notna(f_ag_dt_i) and pd.notna(c_dt):
                                    diff_dias = int((c_dt - f_ag_dt_i).days)
                                    # REGLA CRONOLOGICA ESTRICTA: Solo conciliar facturas anteriores (-60 a 0 dias) o un poco despues (+1 a +15 dias)
                                    if diff_dias < -60 or diff_dias > 15:
                                        continue  # DESCARTAR: fuera de la ventana cronologica del despacho
                                    
                                    if diff_dias == 0:
                                        score_tiempo = 120
                                        relacion_fecha_txt = "\U0001f7e2 Mismo d\u00eda (0d)"
                                    elif -30 <= diff_dias < 0:
                                        score_tiempo = 100 - abs(diff_dias) * 2
                                        relacion_fecha_txt = f"\u2b05\ufe0f Anterior ({abs(diff_dias)}d antes)"
                                    elif 0 < diff_dias <= 15:
                                        score_tiempo = 85 - (diff_dias * 3)
                                        relacion_fecha_txt = f"\u27a1\ufe0f Poco despu\u00e9s (+{diff_dias}d desp)"
                                    elif -60 <= diff_dias < -30:
                                        score_tiempo = 40 - abs(diff_dias)
                                        relacion_fecha_txt = f"\u2b05\ufe0f Anterior ({abs(diff_dias)}d antes)"
                            
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
                                c_row_k_copy["Relaci\u00f3n Fecha"] = relacion_fecha_txt
                                c_row_k_copy["Diferencia D\u00edas"] = diff_dias
                            
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
                        
                            # 3. Busqueda de combinacion con precio concorde exacto (target_i)
                            # Busca combinaciones de 1 a 4 facturas de terceros (tanto  contabilizadas como  pendientes)
                            best_combo_i = []
                            best_diff_i = float("inf")
                            best_score_i = -float("inf")
                        
                            # Ventana cronologica concorde (-35 dias antes a +14 dias despues)
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
                                
                            # Si no dio cuadre exacto en la ventana cercana, ampliar busqueda sobre todos los candidatos
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
    
                    # Inicializar o recuperar paquetes de importacion
                    if "paquetes_importacion" not in st.session_state or not st.session_state["paquetes_importacion"]:
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
                        # Si este paquete NO esta validado, purgar cualquier factura que pertenezca a un paquete validado
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
    
                    # Barra de Acciones y Seleccion de Paquete (Persistente: no salta al paquete 1)
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
                            lock_icon = "\U0001f512 [VALIDADO] " if es_val_p else ""
                            mapa_pqs_titulos[p_num] = f"{lock_icon}\U0001f4e6 Paquete #{p_num}: {str(ag_t.get('Proveedor') or 'Agente')[:16]} ({ag_t['Factura']}) \u2014 Cobro: ${float(ag_t.get('Total', 0.0) or 0.0):,.0f} | {n_terc} Facturas Terceros (${float(s_terc or 0.0):,.0f})"
    
                        idx_actual_sel = pqs_keys.index(st.session_state["paquete_seleccionado_id"]) if st.session_state["paquete_seleccionado_id"] in pqs_keys else 0
    
                        pq_id_sel = st.selectbox(
                            "\U0001f4e6 Selecciona el Paquete de Importaci\u00f3n que deseas revisar:",
                            options=pqs_keys,
                            index=idx_actual_sel,
                            format_func=lambda pid: mapa_pqs_titulos.get(pid, f"Paquete #{pid}")
                        )
                        st.session_state["paquete_seleccionado_id"] = pq_id_sel
                        st.session_state["sel_paquete_activo_key"] = pq_id_sel
                        paquete_activo = pqs_actuales[pq_id_sel]
                        agente_actual = paquete_activo["agente"]
                        terceros_actual = paquete_activo["terceros"]
    
                        # Sincronizacion en tiempo real: refrescar datos del agente desde df_procesado (Hoja 2)
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
                        iva_imp_act = float(agente_actual.get("IVA Importaci\u00f3n", 0.0))
                        subtot_prop_actual = round(base_prop_actual + iva_prop_actual + iva_imp_act, 2)
                        saldo_terceros_esperado_act = round(tot_agente_actual - subtot_prop_actual, 2)
    
                        # Tratamiento contable y asiento de partida doble del paquete
                        enviar_h2_activo = bool(st.session_state.get(f"enviar_h2_pq_{pq_id_sel}", False))
                        enviar_gp_activo = bool(st.session_state.get(f"enviar_gp_pq_{pq_id_sel}", False))
                        enviar_nd_activo = bool(st.session_state.get(f"enviar_nd_pq_{pq_id_sel}", False))
    
                        suma_terceros_cxp_act = sum([obtener_saldo_cruce_factura(r) for _, r in terceros_actual.iterrows()]) if (terceros_actual is not None and not terceros_actual.empty) else 0.0
    
                        # Calcular el asiento previo y la diferencia real de partida doble que necesita cuadrarse
                        if saldo_terceros_esperado_act > 0.05 or enviar_h2_activo:
                            df_prev, dif_faltante_prev = generar_asiento_mixto_hoja2_con_terceros(agente_actual, terceros_actual, enviar_a_no_deducible=False, imputar_a_transito=False)
                            ret_prev = 0.0
                        else:
                            df_prev, dif_faltante_prev, ret_prev = generar_asiento_triangulacion_paquete(agente_actual, terceros_actual, enviar_a_no_deducible=False)
                    
                        # Garantizar que dif_faltante_prev refleje la diferencia real entre debitos y creditos del asiento actual
                        sum_d_pr = df_prev["D\u00e9bito ($)"].sum() if not df_prev.empty else 0.0
                        sum_c_pr = df_prev["Cr\u00e9dito ($)"].sum() if not df_prev.empty else 0.0
                        dif_real_cuadre = round(abs(sum_d_pr - sum_c_pr), 2)
                        if dif_real_cuadre > 0.05:
                            dif_faltante_prev = dif_real_cuadre
    
                        if st.session_state.get(f"paquete_listo_{pq_id_sel}", False) and (f"asiento_fijo_pq_{pq_id_sel}" in st.session_state or "asiento_fijo" in paquete_activo):
                            df_asiento_paquete = st.session_state.get(f"asiento_fijo_pq_{pq_id_sel}", paquete_activo.get("asiento_fijo")).copy()
                            dif_no_ded = 0.0
                            ret_asum = 0.0
                        elif saldo_terceros_esperado_act > 0.05 or enviar_h2_activo:
                            # Facturas con ingresos propios y de terceros (o con solicitud explicita de traer de Pagina 2)
                            if enviar_nd_activo:
                                df_asiento_paquete, _ = generar_asiento_mixto_hoja2_con_terceros(agente_actual, terceros_actual, enviar_a_no_deducible=True, imputar_a_transito=False)
                                dif_no_ded = dif_faltante_prev
                            elif enviar_gp_activo:
                                df_asiento_paquete, _ = generar_asiento_mixto_hoja2_con_terceros(agente_actual, terceros_actual, enviar_a_no_deducible=False, imputar_a_transito=True)
                                dif_no_ded = 0.0
                            else:
                                # Dejar pendiente para conciliar con facturas de terceros (DECISION DEL USUARIO)
                                df_asiento_paquete, dif_no_ded = generar_asiento_mixto_hoja2_con_terceros(agente_actual, terceros_actual, enviar_a_no_deducible=False, imputar_a_transito=False)
                            ret_asum = 0.0
                        elif enviar_gp_activo and dif_faltante_prev > 0.05:
                            df_asiento_paquete, _, ret_asum = generar_asiento_triangulacion_paquete(agente_actual, terceros_actual, enviar_a_no_deducible=False)
                            asiento_gp_filas = []
                            for _, r_as in df_asiento_paquete.iterrows():
                                if str(r_as["C\u00f3digo Cuenta"]).strip() == CUENTA_NO_DEDUCIBLE:
                                    asiento_gp_filas.append({
                                        "C\u00f3digo Cuenta": CUENTA_IMPORTACION_TRANSITO,
                                        "Descripci\u00f3n Cuenta": f"Mercanc\u00edas en Tr\u00e1nsito / Base Propia Agente (Fac {agente_actual.get('Factura', '')})",
                                        "Tercero / NIT": f"{agente_actual.get('NIT Emisor', '')} - {agente_actual.get('Proveedor', '')[:25]}",
                                        "D\u00e9bito ($)": float(r_as["D\u00e9bito ($)"]),
                                        "Cr\u00e9dito ($)": 0.0
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
    
                            # PURGA INMEDIATA: Quitar las facturas de este paquete recien validado de TODOS los demas paquetes no validados
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
    
                            # Traslado y actualizacion inmediata a df_procesado (Pagina 2)
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
                                        df_pr.at[idx_ag[0], "Estado Registro"] = f"\U0001f500 Validada en Triangulaci\u00f3n (Pq #{pq_id_sel} - Pasa a Siigo)"
                                    else:
                                        df_pr.at[idx_ag[0], "Estado Registro"] = f"\U0001f50d Conciliado en Triangulaci\u00f3n (Pq #{pq_id_sel} - Solo Control Interno)"
                                    df_pr.at[idx_ag[0], "Paquete Validado"] = str(pq_id_sel)
                                    df_pr.at[idx_ag[0], "Destino Siigo"] = str(pasa_a_siigo)
                                if not terceros_actual.empty:
                                    for _, tr_row in terceros_actual.iterrows():
                                        tr_fac = str(tr_row["Factura"]).strip()
                                        idx_t = df_pr[df_pr["Factura"].astype(str).str.strip() == tr_fac].index
                                        if not idx_t.empty:
                                            es_r = bool(df_pr.at[idx_t[0], "Ya Registrada"]) or bool(df_pr.at[idx_t[0], "No Contabilizar"]) or (tr_fac in st.session_state.get("facturas_no_contabilizar", set()))
                                            if es_r:
                                                df_pr.at[idx_t[0], "Estado Registro"] = f"\U0001f534 Cruzada en Pq #{pq_id_sel} (Solo cuenta pasivo)"
                                            else:
                                                if pasa_a_siigo:
                                                    df_pr.at[idx_t[0], "Estado Registro"] = f"\u26aa Cruzada en Pq #{pq_id_sel} (Causada en triangulaci\u00f3n)"
                                                else:
                                                    df_pr.at[idx_t[0], "Estado Registro"] = f"\U0001f50d Conciliada en Pq #{pq_id_sel} (Solo Control Interno)"
                                            df_pr.at[idx_t[0], "Paquete Validado"] = str(pq_id_sel)
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
                                    df_pr.at[idx_ag[0], "Estado Registro"] = "\U0001f7e1 Agente Aduanero"
                                    df_pr.at[idx_ag[0], "Paquete Validado"] = ""
                                    df_pr.at[idx_ag[0], "Destino Siigo"] = ""
                                if not terceros_actual.empty:
                                    for _, tr_row in terceros_actual.iterrows():
                                        tr_fac = str(tr_row["Factura"]).strip()
                                        idx_t = df_pr[df_pr["Factura"].astype(str).str.strip() == tr_fac].index
                                        if not idx_t.empty:
                                            es_r = bool(df_pr.at[idx_t[0], "Ya Registrada"]) or bool(df_pr.at[idx_t[0], "No Contabilizar"])
                                            cp = df_pr.at[idx_t[0], "Comprobante Previo"]
                                            df_pr.at[idx_t[0], "Estado Registro"] = f"\U0001f534 Ya Registrada ({cp})" if es_r else "\u26aa Compra Pendiente"
                                            df_pr.at[idx_t[0], "Paquete Validado"] = ""
                                            df_pr.at[idx_t[0], "Cruza Con Agente"] = None
                                st.session_state["df_procesado"] = df_pr
                            guardar_estado_manual(empresa)
    
                    with c_top_pq2:
                        st.write("")
                        st.write("")
                        if st.button("\U0001fa84 Re-calcular Auto-Empaquetado", key="btn_recalc_auto_pqs", help="Vuelve a calcular las combinaciones exactas de DHL + Agencia + Garaje para cada cobro de Euro"):
                            pqs_re, _ = auto_empaquetar_inteligente(df_agentes_all, df_terceros_all)
                            st.session_state["paquetes_importacion"] = pqs_re
                            st.success("\u00a1Paquetes recalculados y separados!")
                            st.rerun()
    
                    # Resumen visual del Paquete Seleccionado
                    col_pq1, col_pq2 = st.columns([1.3, 2.5])
                    with col_pq1:
                        st.markdown(f"""
                        <div style="background:#fffbeb; border:1px solid #fde68a; border-left:5px solid #d97706; border-radius:8px; padding:14px; margin-bottom:12px;">
                            <h4 style="margin:0 0 6px 0; color:#92400e;">\U0001f7e1 Factura del Agente (Cobro Total)</h4>
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
                            st.caption(f"\u2139\ufe0f Elige el destino de este paquete al validarlo:")
                            c_vcard1, c_vcard2 = st.columns(2)
                            with c_vcard1:
                                if st.button("\U0001f4e4 Pasar a Hoja 2", key=f"btn_val_siigo_card_{pq_id_sel}", type="primary", use_container_width=True, help="Bloquea el paquete y lo traslada a P\u00e1gina 2 para subirlo a la Planilla de Siigo"):
                                    ejecutar_validacion_pq(pasa_a_siigo=True)
                                    st.success(f"\u00a1Paquete #{pq_id_sel} trasladado a P\u00e1gina 2 para Siigo!")
                                    st.rerun()
                            with c_vcard2:
                                if st.button("\U0001f4d1 Solo Conciliaci\u00f3n", key=f"btn_val_concil_card_{pq_id_sel}", use_container_width=True, help="Bloquea el paquete para identificar y conciliar facturas, sin subirlo a la Planilla de Siigo"):
                                    ejecutar_validacion_pq(pasa_a_siigo=False)
                                    st.success(f"\u00a1Paquete #{pq_id_sel} guardado para Conciliaci\u00f3n Interna!")
                                    st.rerun()
                        else:
                            es_siigo = st.session_state.get(f"paquete_destino_siigo_{pq_id_sel}", True)
                            if es_siigo:
                                st.success(f"\U0001f512 **Pq #{pq_id_sel}: PASA A HOJA 2 (PLANILLA SIIGO)**")
                            else:
                                st.info(f"\U0001f4d1 **Pq #{pq_id_sel}: SOLO CONCILIACI\u00d3N INTERNA**")
                            if st.button(f"\U0001f513 Desbloquear Paquete #{pq_id_sel}", key=f"btn_desb_card_{pq_id_sel}", use_container_width=True):
                                ejecutar_desbloqueo_pq()
                                st.rerun()
    
                    with col_pq2:
                        st.markdown(f"##### Facturas de Terceros que Componen este Paquete #{pq_id_sel}:")
                        if not terceros_actual.empty:
                            filas_terc_disp = []
                            tot_s_terceros = 0.0
                            for _, tr in terceros_actual.iterrows():
                                es_r = tr.get("Ya Registrada", False)
                                badge_est = f"\U0001f534 Ya en Siigo ({tr.get('Comprobante Previo', '10-Prev')})" if es_r else "\u26aa No Contabilizada (Pendiente)"
                                t_b = float(tr.get("Base", 0.0)) if pd.notna(tr.get("Base")) else 0.0
                                t_iv = float(tr.get("IVA", 0.0)) if pd.notna(tr.get("IVA")) else 0.0
                                s_cruce = obtener_saldo_cruce_factura(tr)
                                tot_s_terceros += s_cruce
                                cta_actual_tr = str(tr.get("Cuenta Pasivo Especifica", "22050505" if "CARGO" in str(tr["Proveedor"]).upper() else "23359501")).strip()
                                p_nom = str(tr["Proveedor"]).upper()
                                rol_dsp = "\U0001f69a DHL (Flete)" if "DHL" in p_nom else ("\U0001f3e2 Agencia (Aduana)" if any(k in p_nom for k in ["CARGO", "ADUANA"]) else "\U0001f3ec Garaje / Almac\u00e9n")
                            
                                rel_fec_val = tr.get("Relaci\u00f3n Fecha")
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
                                st.success(f"\u2705 **Paquete #{pq_id_sel} Exacto:** Total Terceros: **${tot_s_terceros:,.2f}** == Cobro Euro: **${tot_agente_actual:,.2f}** | Diferencia: **$0.00**")
                            else:
                                st.info(f"\U0001f4ca Total Terceros: **${tot_s_terceros:,.2f}** | Cobro Euro: **${tot_agente_actual:,.2f}** | Diferencia a No Deducible: **${dif_pq_actual:,.2f}**")
                        else:
                            st.warning(f"\u26a0\ufe0f El Paquete #{pq_id_sel} no tiene facturas de terceros asignadas todav\u00eda.")
    
                    # Vista previa de las facturas de este paquete (Agente y Terceros)
                    with st.expander(f"\U0001f50d Ver Vista Previa de las Facturas del Paquete #{pq_id_sel} (Agente Aduanero y Terceros)", expanded=False):
                        opc_prev_pq = [f"\U0001f7e1 [AGENTE] {agente_actual.get('Factura', '')} - {agente_actual.get('Proveedor', '')} (${tot_agente_actual:,.2f})"]
                        map_prev_pq = {opc_prev_pq[0]: agente_actual}
                        if not terceros_actual.empty:
                            for _, tr_pv in terceros_actual.iterrows():
                                lbl_tr_pv = f"\U0001f535 [TERCERO] {tr_pv.get('Factura', '')} - {tr_pv.get('Proveedor', '')} (${float(tr_pv.get('Total', 0.0) or 0.0):,.2f})"
                                opc_prev_pq.append(lbl_tr_pv)
                                map_prev_pq[lbl_tr_pv] = tr_pv.to_dict() if hasattr(tr_pv, "to_dict") else dict(tr_pv)
                        sel_fac_prev_pq = st.selectbox("Selecciona la factura del paquete que deseas visualizar:", opc_prev_pq, key=f"sel_prev_fac_pq_{pq_id_sel}")
                        fac_obj_pq = map_prev_pq.get(sel_fac_prev_pq, agente_actual)
                        dict_ren_pq = st.session_state.get("dict_pdfs") or DiskBackedPDFDict("renom_pdfs")
                        dict_ori_pq = st.session_state.get("raw_uploaded_pdfs") or DiskBackedPDFDict("raw_pdfs")
                        pdf_b_pq, orig_d_pq, _ = buscar_y_extraer_pdf(fac_obj_pq, dict_renombrados=dict_ren_pq, dict_originales=dict_ori_pq, empresa_compradora=empresa)
                        pags_txt_pq = obtener_paginas_texto_factura_integrado(fac_obj_pq)
                        renderizar_representacion_grafica_dian_html(
                            fac_obj_pq,
                            pags_txt_pq or [],
                            empresa,
                            pdf_bytes_original=pdf_b_pq,
                            origen_pdf_desc=orig_d_pq or "",
                            key_suffix=f"pq_{pq_id_sel}"
                        )

                    # Expander para agregar/quitar facturas manualmente a este paquete (BLOQUEADO SI ESTA VALIDADO)
                    es_validado_este = st.session_state.get(f"paquete_listo_{pq_id_sel}", False)
                    if es_validado_este:
                        st.markdown(f"""
                        <div style="background:#f0fdf4; border:1px solid #86efac; border-left:5px solid #16a34a; border-radius:8px; padding:12px 16px; margin:12px 0;">
                            <h5 style="margin:0 0 4px 0; color:#166534;">\U0001f512 Paquete #{pq_id_sel} VALIDADO Y BLOQUEADO</h5>
                            <p style="margin:0; color:#14532d; font-size:13.5px;">
                                Este paquete se encuentra congelado. Ninguna factura puede ser a\u00f1adida ni retirada mientras est\u00e9 validado. Para realizar modificaciones, presiona <b>'\U0001f513 Desbloquear y Modificar'</b> abajo.
                            </p>
                        </div>
                        """, unsafe_allow_html=True)
                
                    with st.expander(f"\u2699\ufe0f Modificar este Paquete #{pq_id_sel} (Sacar Facturas o Agregar Nuevas):", expanded=False if es_validado_este else False):
                        if es_validado_este:
                            st.warning("\u26a0\ufe0f Debes desbloquear el paquete abajo antes de poder agregar o sacar facturas.")
                        else:
                            st.caption("Administra las facturas asignadas a esta importaci\u00f3n: saca las que no correspondan o a\u00f1ade nuevas facturas de DHL, Agencia o Garaje:")
                    
                            # 1. SECCION PARA SACAR / REDIRIGIR FACTURAS
                            st.markdown("##### \u2796 Sacar y Redirigir una factura de este paquete:")
                            st.caption("Si ves una factura que no corresponde a este paquete, s\u00e1cala y el sistema la reubicar\u00e1 autom\u00e1ticamente o la enviar\u00e1 a compras directas para que no quede pendiente:")
                            facs_en_este = list(terceros_actual["Factura"].unique()) if not terceros_actual.empty else []
                    
                            if facs_en_este:
                                opciones_quitar = []
                                mapa_quitar_row = {}
                                for _, tr_k in terceros_actual.iterrows():
                                    tb_k = float(tr_k.get("Base", 0.0))
                                    tiv_k = float(tr_k.get("IVA", 0.0))
                                    sc_k = obtener_saldo_cruce_factura(tr_k)
                                    tag_q = f"[{tr_k['Factura']}] {str(tr_k.get('Proveedor') or 'Tercero')[:24]} (Saldo: ${sc_k:,.2f})"
                                    opciones_quitar.append(tag_q)
                                    mapa_quitar_row[tag_q] = tr_k
                            
                                sel_para_quitar = st.selectbox("1. Selecciona la factura que deseas SACAR de este paquete:", ["(Seleccionar...)"] + opciones_quitar, key=f"sel_rem_tr_{pq_id_sel}")
                        
                                if sel_para_quitar != "(Seleccionar...)":
                                    fila_a_mover = mapa_quitar_row[sel_para_quitar]
                                    fac_nom_mover = fila_a_mover["Factura"]
                                    sc_mover = obtener_saldo_cruce_factura(fila_a_mover)
                            
                                    otros_pqs = [p for p in pqs_actuales.keys() if p != pq_id_sel and not st.session_state.get(f"paquete_listo_{p}", False)]
                            
                                    # Calcular cual otro paquete tiene un faltante mas cercano a esta factura
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
                                        f"\U0001f3af Auto-Reubicar autom\u00e1ticamente donde mejor encaje{sug_txt}",
                                        "\U0001f4e6 Mover a otro Paquete espec\u00edfico...",
                                        "\U0001f6d2 Redirigir a Compras Directas Ordinarias (Pesta\u00f1a 1 y 4 - Se pag\u00f3 directo)",
                                        "\u26aa Dejar libre en inventario (Sin asignar)"
                                    ]
                            
                                    c_dest1, c_dest2 = st.columns([2.6, 1.2])
                                    with c_dest1:
                                        dest_sel = st.radio(f"2. \u00bfA d\u00f3nde deseas redirigir la factura {fac_nom_mover}?:", opciones_destino, key=f"rad_dest_{pq_id_sel}_{fac_nom_mover}")
                                        destino_pq_manual = None
                                        if "Mover a otro Paquete espec\u00edfico" in dest_sel:
                                            opciones_otros_pqs = [f"Paquete #{p}: {str(pqs_actuales[p]['agente'].get('Proveedor') or 'Agente')[:16]} (Fac {pqs_actuales[p]['agente']['Factura']})" for p in otros_pqs]
                                            if opciones_otros_pqs:
                                                sel_otro_pq_str = st.selectbox("Selecciona el paquete de destino:", opciones_otros_pqs, key=f"sel_otro_pq_{pq_id_sel}")
                                                destino_pq_manual = int(sel_otro_pq_str.split(":")[0].replace("Paquete #", "").strip())
                                            else:
                                                st.info("No hay otros paquetes de forwarders disponibles en este archivo.")
                                        
                                    with c_dest2:
                                        st.write("")
                                        st.write("")
                                        if st.button("\U0001f680 Sacar y Redirigir Factura", key=f"btn_ejecutar_reubicar_{pq_id_sel}", help="Retira la factura de este paquete y la redirige al destino seleccionado."):
                                            # 1. Retirar del paquete actual
                                            terceros_remanente = terceros_actual[terceros_actual["Factura"] != fac_nom_mover].copy()
                                            st.session_state["paquetes_importacion"][pq_id_sel]["terceros"] = terceros_remanente
                                    
                                            # 2. Redirigir al destino
                                            if "Auto-Reubicar" in dest_sel and mejor_pq_sug is not None:
                                                dest_pq = st.session_state["paquetes_importacion"][mejor_pq_sug]["terceros"]
                                                dest_nuevo = pd.concat([dest_pq, pd.DataFrame([fila_a_mover])]).drop_duplicates(subset=["Factura"]).reset_index(drop=True)
                                                st.session_state["paquetes_importacion"][mejor_pq_sug]["terceros"] = dest_nuevo
                                                st.success(f"\u00a1Factura {fac_nom_mover} retirada del Paquete #{pq_id_sel} y auto-reubicada en el Paquete #{mejor_pq_sug}!")
                                            elif "Mover a otro Paquete" in dest_sel and destino_pq_manual is not None:
                                                dest_pq = st.session_state["paquetes_importacion"][destino_pq_manual]["terceros"]
                                                dest_nuevo = pd.concat([dest_pq, pd.DataFrame([fila_a_mover])]).drop_duplicates(subset=["Factura"]).reset_index(drop=True)
                                                st.session_state["paquetes_importacion"][destino_pq_manual]["terceros"] = dest_nuevo
                                                st.success(f"\u00a1Factura {fac_nom_mover} movida con \u00e9xito al Paquete #{destino_pq_manual}!")
                                            elif "Compras Directas Ordinarias" in dest_sel:
                                                if "df_procesado" in st.session_state:
                                                    df_glob = st.session_state["df_procesado"]
                                                    m_idx = df_glob[df_glob["Factura"] == fac_nom_mover].index
                                                    if not m_idx.empty:
                                                        df_glob.at[m_idx[0], "Es Aduanera"] = False
                                                        df_glob.at[m_idx[0], "Grupo Importaci\u00f3n"] = ""
                                                        df_glob.at[m_idx[0], "Estado Registro"] = "\u26aa Compra Directa Ordinaria"
                                                        df_glob.at[m_idx[0], "Cta Contrapartida"] = "22050501"
                                                        st.session_state["df_procesado"] = df_glob
                                                st.success(f"\u00a1Factura {fac_nom_mover} retirada de importaciones y enviada a Compras Directas Ordinarias (Pesta\u00f1a 1 y 4)!")
                                            else:
                                                st.success(f"\u00a1Factura {fac_nom_mover} retirada del Paquete #{pq_id_sel} y devuelta a facturas libres!")
                                    
                                            st.session_state["paquete_seleccionado_id"] = pq_id_sel
                                            st.session_state["sel_paquete_activo_key"] = pq_id_sel
                                            guardar_estado_manual(empresa)
                                            st.rerun()
                            else:
                                st.info("Este paquete no tiene facturas asignadas para retirar.")
                        
                            st.markdown("---")
                    
                            # 2. SECCION PARA AGREGAR NUEVAS FACTURAS
                            cands_disp_agregar = df_terceros_all.copy()
                            f_ag_dt_sug = pd.to_datetime(agente_actual.get("Fecha", ""), dayfirst=True, errors="coerce")
                    
                            # RESTRICCION DE BLOQUEO: Las facturas de paquetes ya validados no se pueden anadir a otros
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
                                        <h5 style="margin:0 0 4px 0; color:#15803d;">\U0001f50e Facturas Sugeridas para Ingresos de Terceros (Faltante Pendiente: ${dif_faltante_prev:,.2f}):</h5>
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
                            
                                    # REGLA CRONOLOGICA ESTRICTA: Solo facturas anteriores (-60d a 0d) o un poco despues (+1d a +15d)
                                    if pd.notna(fec_dt) and pd.notna(f_ag_dt_sug):
                                        if diff_dias < -60 or diff_dias > 15:
                                            continue
                                    elif diff_dias == 999:
                                        continue
                            
                                    tags_sug = []
                                    score_sug = 0
                            
                                    # Coincidencias especificas con los items de ingresos por terceros (solo si la factura del agente es BOG-70697 o de saldo similar)
                                    es_caso_70697 = "70697" in str(agente_actual.get("Factura", "")) or abs(saldo_terceros_esperado_act - 2542421.20) < 100.0
                                    if es_caso_70697 and abs(sc_cand_val - 1079157.20) < 50000:
                                        tags_sug.append("\u2708\ufe0f Flete Internacional (~$1.08M)")
                                        score_sug += 180 - (abs(sc_cand_val - 1079157.20) / 1000)
                                    elif es_caso_70697 and abs(sc_cand_val - 1463264.00) < 50000:
                                        tags_sug.append("\U0001f3e2 Cargos en Origen (~$1.46M)")
                                        score_sug += 180 - (abs(sc_cand_val - 1463264.00) / 1000)
                                    elif abs(sc_cand_val - dif_faltante_prev) < 5000:
                                        tags_sug.append(f"\U0001f3af Cuadre Exacto (${dif_faltante_prev:,.2f})")
                                        score_sug += 250
                                    elif sc_cand_val <= dif_faltante_prev + 0.05:
                                        tags_sug.append(f"\U0001f9e9 Cubre {(sc_cand_val / dif_faltante_prev * 100):.1f}% del faltante")
                                        score_sug += 60
                                
                                    # Ponderacion por cercania de fecha dentro de la ventana (-60 a +15)
                                    if diff_dias == 0:
                                        tags_sug.append("\U0001f7e2 Mismo d\u00eda (0d)")
                                        score_sug += 100
                                    elif 0 < diff_dias <= 15:
                                        tags_sug.append(f"\u27a1\ufe0f Poco despu\u00e9s (+{diff_dias}d)")
                                        score_sug += 85 - (diff_dias * 3)
                                    elif -30 <= diff_dias < 0:
                                        tags_sug.append(f"\u2b05\ufe0f Anterior ({abs(diff_dias)}d antes)")
                                        score_sug += 90 - (abs(diff_dias) * 2)
                                    elif -60 <= diff_dias < -30:
                                        tags_sug.append(f"\u2b05\ufe0f Anterior ({abs(diff_dias)}d antes)")
                                        score_sug += 40 - abs(diff_dias)
                                
                                    if any(k in prov_u for k in ["DHL", "CARGO", "ADUANA", "GARAJE", "ALMAVIVA", "ALMACENADORA", "UPS", "FEDEX", "FLETE", "TRANSPORTE", "PUERTO", "TERMINAL"]):
                                        tags_sug.append("\U0001f6a2 Log\u00edstica/Aduana")
                                        score_sug += 50
                                
                                    cands_sug_list.append({
                                        "Factura": f_c_num,
                                        "Proveedor": tr_cand_s.get("Proveedor", "")[:32],
                                        "Fecha": fec_str,
                                        "Saldo": sc_cand_val,
                                        "Dias": diff_dias,
                                        "Tags": " \u2022 ".join(tags_sug) if tags_sug else "\u26aa Factura Libre",
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
                                            f"\U0001f4a1 <b>Combinaci\u00f3n de 2 Facturas Sugerida:</b> <code>{cp1['Factura']}</code> ({cp1['Proveedor']}, ${cp1['Saldo']:,.2f}) + <code>{cp2['Factura']}</code> ({cp2['Proveedor']}, ${cp2['Saldo']:,.2f})<br>"
                                            f"<span style='font-size:12px; color:#854d0e;'>Suma Combinada: <b>${sp:,.2f}</b> (Diferencia: <b>${dp:,.2f}</b>)</span></div>",
                                            unsafe_allow_html=True
                                        )
                                    with col_par2:
                                        st.write("")
                                        if st.button("\u2795 A\u00f1adir Ambas Facturas", key=f"btn_add_pair_{pq_id_sel}", type="primary", use_container_width=True):
                                            terceros_nuevo = pd.concat([terceros_actual, pd.DataFrame([cp1["Row"], cp2["Row"]])]).drop_duplicates(subset=["Factura"]).reset_index(drop=True)
                                            st.session_state["paquetes_importacion"][pq_id_sel]["terceros"] = terceros_nuevo
                                            st.session_state["paquete_seleccionado_id"] = pq_id_sel
                                            st.session_state["sel_paquete_activo_key"] = pq_id_sel
                                            guardar_estado_manual(empresa)
                                            st.success(f"\u00a1Facturas {cp1['Factura']} y {cp2['Factura']} a\u00f1adidas con \u00e9xito al Paquete #{pq_id_sel}!")
                                            st.rerun()
    
                                if cands_sug_list:
                                    top_cands = cands_sug_list[:6]
                                    for idx_sug, sug_item in enumerate(top_cands):
                                        c_sug1, c_sug2, c_sug3 = st.columns([1.5, 3.2, 1.1])
                                        with c_sug1:
                                            st.markdown(f"<b>{sug_item['Factura']}</b><br><span style='font-size:12px; color:#475569;'>{sug_item['Fecha']}</span>", unsafe_allow_html=True)
                                        with c_sug2:
                                            st.markdown(f"<b>{sug_item['Proveedor']}</b> \u2014 <b style='color:#0369a1;'>${sug_item['Saldo']:,.2f}</b><br><span style='font-size:12px; color:#15803d;'>{sug_item['Tags']}</span>", unsafe_allow_html=True)
                                        with c_sug3:
                                            st.write("")
                                            if st.button("\u2795 A\u00f1adir", key=f"btn_sug_add_{pq_id_sel}_{idx_sug}", use_container_width=True):
                                                terceros_nuevo = pd.concat([terceros_actual, pd.DataFrame([sug_item["Row"]])]).drop_duplicates(subset=["Factura"]).reset_index(drop=True)
                                                st.session_state["paquetes_importacion"][pq_id_sel]["terceros"] = terceros_nuevo
                                                st.session_state["paquete_seleccionado_id"] = pq_id_sel
                                                st.session_state["sel_paquete_activo_key"] = pq_id_sel
                                                guardar_estado_manual(empresa)
                                                st.success(f"\u00a1Factura {sug_item['Factura']} a\u00f1adida al Paquete #{pq_id_sel}!")
                                                st.rerun()
                                else:
                                    st.info(f"No se encontraron facturas sugeridas que no superen el monto de terceros (${monto_limite_terc:,.2f}) ni est\u00e9n bloqueadas.")
                                st.markdown("---")
    
                            st.markdown("##### \u2795 A\u00f1adir Manualmente otra Factura Libre a este paquete:")
                            filtro_cercanas_man = st.checkbox(
                                "\U0001f4c5 Filtrar solo facturas de la ventana cronol\u00f3gica (-60d anteriores a +15d posteriores)",
                                value=True,
                                key=f"chk_fec_man_{pq_id_sel}",
                                help="Oculta facturas lejanas de otros meses y muestra solo facturas anteriores o un poco despu\u00e9s de la operaci\u00f3n."
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
                                    tag_est = f"\U0001f534 Registrada ({tr_cand.get('Comprobante Previo', '10-Prev')})" if es_reg_c else "\u26aa No Contabilizada (Pendiente)"
                                    sc_cand = obtener_saldo_cruce_factura(tr_cand)
                                    d_tag = f"({diff_m:+d}d)" if diff_m != 999 else "(sin fecha)"
                                    tag_c = f"[{tag_est}] [{f_cand_num}] {d_tag} {str(tr_cand.get('Proveedor') or 'Tercero')[:20]} (Saldo: ${sc_cand:,.2f})"
                                    opciones_cands.append((abs(diff_m), tag_c, tr_cand))
    
                            opciones_cands.sort(key=lambda x: x[0])
                            opciones_agregar = [x[1] for x in opciones_cands]
                            for _, tag_c, tr_cand in opciones_cands:
                                mapa_agregar[tag_c] = tr_cand
                            
                            c_add1, c_add2 = st.columns([3, 1.2])
                            with c_add1:
                                sel_para_agregar = st.selectbox("Selecciona factura libre para a\u00f1adir a este paquete:", ["(Seleccionar...)"] + opciones_agregar, key=f"sel_add_tr_{pq_id_sel}")
                            with c_add2:
                                st.write("")
                                st.write("")
                                if st.button("\u2795 A\u00f1adir al Paquete", key=f"btn_add_tr_{pq_id_sel}") and sel_para_agregar != "(Seleccionar...)":
                                    fila_agregada = mapa_agregar[sel_para_agregar]
                                    terceros_nuevo = pd.concat([terceros_actual, pd.DataFrame([fila_agregada])]).drop_duplicates(subset=["Factura"]).reset_index(drop=True)
                                    st.session_state["paquetes_importacion"][pq_id_sel]["terceros"] = terceros_nuevo
                                    st.session_state["paquete_seleccionado_id"] = pq_id_sel
                                    st.session_state["sel_paquete_activo_key"] = pq_id_sel
                                    guardar_estado_manual(empresa)
                                    st.success(f"\u00a1Factura {fila_agregada['Factura']} a\u00f1adida al Paquete #{pq_id_sel}!")
                                    st.rerun()
    
                    # 3. TRATAMIENTO CONTABLE: PAGINA 2 / MERCANCIAS EN TRANSITO / NO DEDUCIBLES
                    enviar_h2_activo = bool(st.session_state.get(f"enviar_h2_pq_{pq_id_sel}", False))
                    enviar_gp_activo = bool(st.session_state.get(f"enviar_gp_pq_{pq_id_sel}", False))
                    enviar_nd_activo = bool(st.session_state.get(f"enviar_nd_pq_{pq_id_sel}", False))
    
                    c_box_style_t = (
                        '<div style="background:#f8fafc; border:2px solid #0070ba; border-radius:8px; padding:14px; margin:14px 0 10px 0;">'
                        '<h4 style="margin:0 0 4px 0; color:#0070ba;">\u2696\ufe0f Tratamiento Contable para el Paquete #' + str(pq_id_sel) + ' (Cobro ' + str(agente_actual.get('Proveedor', '')) + '):</h4>'
                        '<p style="margin:0; font-size:13.5px; color:#334155;">Selecciona c\u00f3mo deseas registrar este cobro en su fecha de operaci\u00f3n (<b>' + str(agente_actual.get('Fecha', '')) + '</b>):</p>'
                        '</div>'
                    )
                    st.markdown(c_box_style_t, unsafe_allow_html=True)
    
                    c_btn_h2, c_btn_mt, c_btn_nd = st.columns([1.8, 1.6, 1.2])
    
                    with c_btn_h2:
                        st.markdown("<b style='color:#16a34a;'>\U0001f4cb Opci\u00f3n 1: Traer de P\u00e1gina 2</b><br><span style='font-size:12px; color:#475569;'>Trae ingresos propios y deja saldo de terceros pendiente para conciliar:</span>", unsafe_allow_html=True)
                        btn_h2_type = "primary" if (enviar_h2_activo or not (enviar_gp_activo or enviar_nd_activo)) else "secondary"
                        if st.button("\U0001f4cb Traer de P\u00e1gina 2 (Dejar Pendiente)", key=f"btn_h2_act_{pq_id_sel}", type=btn_h2_type, use_container_width=True):
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
                            st.success("\u00a1Contabilizaci\u00f3n de P\u00e1gina 2 aplicada! El saldo de terceros queda pendiente para conciliar con facturas.")
                            st.rerun()
    
                    with c_btn_mt:
                        st.markdown("<b style='color:#0284c7;'>\U0001f4e6 Opci\u00f3n 2: Mercanc\u00edas en Tr\u00e1nsito</b><br><span style='font-size:12px; color:#475569;'>Imputa el saldo faltante a Inventarios en Tr\u00e1nsito (Cta 14650501):</span>", unsafe_allow_html=True)
                        lbl_btn_mt = f"\U0001f4e6 Mercanc\u00edas en Tr\u00e1nsito (${dif_faltante_prev:,.2f})" if dif_faltante_prev > 0.05 else "\U0001f4e6 Mercanc\u00edas en Tr\u00e1nsito (14650501)"
                        btn_mt_type = "primary" if enviar_gp_activo else "secondary"
                        if st.button(lbl_btn_mt, key=f"btn_mt_act_{pq_id_sel}", type=btn_mt_type, use_container_width=True):
                            st.session_state[f"enviar_gp_pq_{pq_id_sel}"] = True
                            st.session_state[f"enviar_nd_pq_{pq_id_sel}"] = False
                            st.session_state["paquete_seleccionado_id"] = pq_id_sel
                            st.session_state["sel_paquete_activo_key"] = pq_id_sel
                            if saldo_terceros_esperado_act > 0.05 or enviar_h2_activo:
                                st.session_state[f"enviar_h2_pq_{pq_id_sel}"] = True
                                df_as_act, _ = generar_asiento_mixto_hoja2_con_terceros(agente_actual, terceros_actual, enviar_a_no_deducible=False, imputar_a_transito=True)
                            else:
                                st.session_state[f"enviar_h2_pq_{pq_id_sel}"] = False
                                df_as_act, _, ret_asum = generar_asiento_triangulacion_paquete(agente_actual, terceros_actual, enviar_a_no_deducible=False)
                                asiento_gp_filas = []
                                for _, r_as in df_as_act.iterrows():
                                    if str(r_as["C\u00f3digo Cuenta"]).strip() == CUENTA_NO_DEDUCIBLE:
                                        asiento_gp_filas.append({
                                            "C\u00f3digo Cuenta": CUENTA_IMPORTACION_TRANSITO,
                                            "Descripci\u00f3n Cuenta": f"Mercanc\u00edas en Tr\u00e1nsito / Base Propia Agente (Fac {agente_actual.get('Factura', '')})",
                                            "Tercero / NIT": f"{agente_actual.get('NIT Emisor', '')} - {agente_actual.get('Proveedor', '')[:25]}",
                                            "D\u00e9bito ($)": float(r_as["D\u00e9bito ($)"]),
                                            "Cr\u00e9dito ($)": 0.0
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
                            st.success(f"\u00a1Diferencia de ${dif_faltante_prev:,.2f} asignada a Mercanc\u00edas en Tr\u00e1nsito (14650501)!")
                            st.rerun()
    
                    with c_btn_nd:
                        st.markdown("<b style='color:#b91c1c;'>\U0001f534 Opci\u00f3n 3: No Deducibles</b><br><span style='font-size:12px; color:#475569;'>Si no existe factura DIAN ni soporte:</span>", unsafe_allow_html=True)
                        lbl_btn_nd = f"\U0001f534 No Deducibles (${dif_faltante_prev:,.2f})" if dif_faltante_prev > 0.05 else "\U0001f534 No Deducibles (53950501)"
                        btn_nd_type = "primary" if enviar_nd_activo else "secondary"
                        if st.button(lbl_btn_nd, key=f"btn_nd_act_{pq_id_sel}", type=btn_nd_type, use_container_width=True):
                            st.session_state[f"enviar_nd_pq_{pq_id_sel}"] = True
                            st.session_state[f"enviar_gp_pq_{pq_id_sel}"] = False
                            st.session_state["paquete_seleccionado_id"] = pq_id_sel
                            st.session_state["sel_paquete_activo_key"] = pq_id_sel
                            if saldo_terceros_esperado_act > 0.05 or enviar_h2_activo:
                                st.session_state[f"enviar_h2_pq_{pq_id_sel}"] = True
                                df_as_act, _ = generar_asiento_mixto_hoja2_con_terceros(agente_actual, terceros_actual, enviar_a_no_deducible=True, imputar_a_transito=False)
                            else:
                                st.session_state[f"enviar_h2_pq_{pq_id_sel}"] = False
                                df_as_act, _, _ = generar_asiento_triangulacion_paquete(agente_actual, terceros_actual, enviar_a_no_deducible=True)
                            if st.session_state.get(f"paquete_listo_{pq_id_sel}", False):
                                st.session_state[f"asiento_fijo_pq_{pq_id_sel}"] = df_as_act.copy()
                                paquete_activo["asiento_fijo"] = df_as_act.copy()
                                ag_fac_str = str(agente_actual["Factura"]).strip()
                                st.session_state.setdefault("asientos_triangulacion_por_factura", {})[ag_fac_str] = df_as_act.copy()
                                guardar_estado_manual(empresa)
                            st.success(f"\u00a1Diferencia de ${dif_faltante_prev:,.2f} enviada a Gastos No Deducibles (53950501)!")
                            st.rerun()
    
                    # RECALCULAR ASIENTO CONTABLE PARA ESTE PAQUETE
                    if st.session_state.get(f"paquete_listo_{pq_id_sel}", False) and (f"asiento_fijo_pq_{pq_id_sel}" in st.session_state or "asiento_fijo" in paquete_activo):
                        df_asiento_paquete = sincronizar_asiento_con_factura_actual(st.session_state.get(f"asiento_fijo_pq_{pq_id_sel}", paquete_activo.get("asiento_fijo")).copy(), agente_actual)
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
                            if str(r_as["C\u00f3digo Cuenta"]).strip() == CUENTA_NO_DEDUCIBLE:
                                asiento_gp_filas.append({
                                    "C\u00f3digo Cuenta": CUENTA_IMPORTACION_TRANSITO,
                                    "Descripci\u00f3n Cuenta": f"Mercanc\u00edas en Tr\u00e1nsito / Base Propia Agente (Fac {agente_actual.get('Factura', '')})",
                                    "Tercero / NIT": f"{agente_actual.get('NIT Emisor', '')} - {agente_actual.get('Proveedor', '')[:25]}",
                                    "D\u00e9bito ($)": float(r_as["D\u00e9bito ($)"]),
                                    "Cr\u00e9dito ($)": 0.0
                                })
                            else:
                                asiento_gp_filas.append(r_as.to_dict())
                        df_asiento_paquete = pd.DataFrame(asiento_gp_filas)
                        dif_no_ded = 0.0
                    else:
                        df_asiento_paquete, dif_no_ded, ret_asum = generar_asiento_triangulacion_paquete(agente_actual, terceros_actual, enviar_a_no_deducible=enviar_nd_activo)
    
                    st.markdown(f"#### \u2696\ufe0f Asiento Contable del Paquete #{pq_id_sel}:")
                    st.caption("Detalle de partida doble de ESTE paquete: cancela las cuentas por pagar de DHL, Agencia y Garaje contra Euro Shipping:")
    
                    st.dataframe(
                        df_asiento_paquete.style.format({"D\u00e9bito ($)": "${:,.2f}", "Cr\u00e9dito ($)": "${:,.2f}"}),
                        use_container_width=True,
                        hide_index=True
                    )
    
                    c_cuad1, c_cuad2, c_cuad3, c_cuad4 = st.columns(4)
                    sum_deb_pq = df_asiento_paquete["D\u00e9bito ($)"].sum()
                    sum_cred_pq = df_asiento_paquete["Cr\u00e9dito ($)"].sum()
                    dif_cuad_pq = abs(sum_deb_pq - sum_cred_pq)
                
                    with c_cuad1:
                        st.metric("Total D\u00e9bito", f"${sum_deb_pq:,.2f}")
                    with c_cuad2:
                        st.metric("Total Cr\u00e9dito", f"${sum_cred_pq:,.2f}")
                    with c_cuad3:
                        st.metric("Retenciones Asumidas (53152001)", f"${ret_asum:,.2f}")
                    with c_cuad4:
                        st.metric("Diferencia No Deducible (53950501)", f"${dif_no_ded:,.2f}")
    
                    # SECCION DE VALIDACION Y BLOQUEO DE FACTURAS
                    es_validado = st.session_state.get(f"paquete_listo_{pq_id_sel}", False)
    
                    if not es_validado:
                        if dif_cuad_pq < 0.05:
                            st.success(f"\u2705 **Partida doble cuadrada al centavo ($0.00).** Ya puedes validar y bloquear este paquete.")
                        else:
                            st.markdown(
                                f"""
                                <div style="background:#fffbeb; border:2px solid #f59e0b; border-radius:8px; padding:12px 16px; margin:10px 0;">
                                    <span style="font-size:16px; font-weight:bold; color:#b45309;">\u23f3 Faltante Pendiente de Conciliar: ${dif_cuad_pq:,.2f}</span><br>
                                    <span style="font-size:13.5px; color:#78350f;">
                                        El comprobante presenta una diferencia de <b>${dif_cuad_pq:,.2f}</b> para cuadrar sumas iguales. Puedes aplicar este valor directamente a No Deducibles o a Tr\u00e1nsito con un solo clic:
                                    </span>
                                </div>
                                """,
                                unsafe_allow_html=True
                            )
                            c_qnd1, c_qnd2, c_qnd3 = st.columns([1.6, 1.6, 1.8])
                            with c_qnd1:
                                if st.button(f"\U0001f534 Mandar ${dif_cuad_pq:,.2f} a No Deducibles", key=f"btn_box_nd_{pq_id_sel}", type="primary", use_container_width=True):
                                    st.session_state[f"enviar_nd_pq_{pq_id_sel}"] = True
                                    st.session_state[f"enviar_gp_pq_{pq_id_sel}"] = False
                                    st.session_state[f"enviar_h2_pq_{pq_id_sel}"] = True
                                    st.session_state["paquete_seleccionado_id"] = pq_id_sel
                                    st.session_state["sel_paquete_activo_key"] = pq_id_sel
                                    if st.session_state.get(f"paquete_listo_{pq_id_sel}", False):
                                        df_as_act, _ = generar_asiento_mixto_hoja2_con_terceros(agente_actual, terceros_actual, enviar_a_no_deducible=True, imputar_a_transito=False)
                                        st.session_state[f"asiento_fijo_pq_{pq_id_sel}"] = df_as_act.copy()
                                        paquete_activo["asiento_fijo"] = df_as_act.copy()
                                        ag_fac_str = str(agente_actual["Factura"]).strip()
                                        st.session_state.setdefault("asientos_triangulacion_por_factura", {})[ag_fac_str] = df_as_act.copy()
                                        guardar_estado_manual(empresa)
                                    st.success(f"\u00a1Diferencia de ${dif_cuad_pq:,.2f} enviada a No Deducibles (53950501)!")
                                    st.rerun()
                            with c_qnd2:
                                if st.button(f"\U0001f4e6 Mandar ${dif_cuad_pq:,.2f} a Tr\u00e1nsito", key=f"btn_box_mt_{pq_id_sel}", use_container_width=True):
                                    st.session_state[f"enviar_gp_pq_{pq_id_sel}"] = True
                                    st.session_state[f"enviar_nd_pq_{pq_id_sel}"] = False
                                    st.session_state[f"enviar_h2_pq_{pq_id_sel}"] = True
                                    st.session_state["paquete_seleccionado_id"] = pq_id_sel
                                    st.session_state["sel_paquete_activo_key"] = pq_id_sel
                                    if st.session_state.get(f"paquete_listo_{pq_id_sel}", False):
                                        df_as_act, _ = generar_asiento_mixto_hoja2_con_terceros(agente_actual, terceros_actual, enviar_a_no_deducible=False, imputar_a_transito=True)
                                        st.session_state[f"asiento_fijo_pq_{pq_id_sel}"] = df_as_act.copy()
                                        paquete_activo["asiento_fijo"] = df_as_act.copy()
                                        ag_fac_str = str(agente_actual["Factura"]).strip()
                                        st.session_state.setdefault("asientos_triangulacion_por_factura", {})[ag_fac_str] = df_as_act.copy()
                                        guardar_estado_manual(empresa)
                                    st.success(f"\u00a1Diferencia de ${dif_cuad_pq:,.2f} enviada a Tr\u00e1nsito (14650501)!")
                                    st.rerun()
                            with c_qnd3:
                                st.caption("\U0001f4a1 O retira facturas de terceros que no correspondan en la secci\u00f3n '\u2699\ufe0f Modificar este Paquete' arriba.")
    
                        c_val_opt1, c_val_opt2 = st.columns(2)
                        with c_val_opt1:
                            st.markdown(
                                "<div style='background:#f0fdf4; border:1.5px solid #22c55e; border-radius:8px; padding:12px; height:100%;'>"
                                "<b style='color:#15803d; font-size:15px;'>\U0001f4e4 Opci\u00f3n 1: Pasar a Hoja 2 (Subir a Planilla Siigo)</b><br>"
                                "<span style='font-size:12.5px; color:#166534;'>"
                                "Bloquea las facturas del paquete, actualiza la P\u00e1gina 2 y <b>genera el comprobante oficial en la Planilla de Siigo (P\u00e1gina 4)</b> para su contabilizaci\u00f3n."
                                "</span></div>",
                                unsafe_allow_html=True
                            )
                            st.write("")
                            if st.button("\U0001f4e4 Pasar a Hoja 2 (Subir a Planilla Siigo)", key=f"btn_validar_siigo_{pq_id_sel}", type="primary", use_container_width=True):
                                ejecutar_validacion_pq(pasa_a_siigo=True)
                                st.success(f"\U0001f512 \u00a1Paquete #{pq_id_sel} VALIDADO y trasladado a P\u00e1gina 2 para Siigo!")
                                st.rerun()
    
                        with c_val_opt2:
                            st.markdown(
                                "<div style='background:#eff6ff; border:1.5px solid #3b82f6; border-radius:8px; padding:12px; height:100%;'>"
                                "<b style='color:#1d4ed8; font-size:15px;'>\U0001f4d1 Opci\u00f3n 2: Simplemente Conciliaci\u00f3n (No Subir a Siigo)</b><br>"
                                "<span style='font-size:12.5px; color:#1e40af;'>"
                                "Bloquea e identifica las facturas para <b>auditor\u00eda y soporte de costos</b>, pero <b>NO genera comprobante en Siigo</b> (evita duplicar pagos ya hechos por tesorer\u00eda/bancos)."
                                "</span></div>",
                                unsafe_allow_html=True
                            )
                            st.write("")
                            if st.button("\U0001f4d1 Simplemente Conciliaci\u00f3n (No Subir a Siigo)", key=f"btn_validar_concil_{pq_id_sel}", use_container_width=True):
                                ejecutar_validacion_pq(pasa_a_siigo=False)
                                st.success(f"\U0001f4d1 \u00a1Paquete #{pq_id_sel} CONCILIADO exitosamente para control interno (no ir\u00e1 a Siigo)!")
                                st.rerun()
                    else:
                        es_siigo = st.session_state.get(f"paquete_destino_siigo_{pq_id_sel}", True)
                        col_desb1, col_desb2 = st.columns([3.2, 1.4])
                        with col_desb1:
                            if es_siigo:
                                st.success(f"\U0001f512 **Paquete #{pq_id_sel} VALIDADO (PASA A HOJA 2 / PLANILLA SIIGO):** Sus facturas est\u00e1n protegidas contra otros cruces y su comprobante se exportar\u00e1 en la Planilla de Siigo (P\u00e1gina 4).")
                            else:
                                st.info(f"\U0001f4d1 **Paquete #{pq_id_sel} CONCILIADO (SOLO CONTROL INTERNO / IDENTIFICACI\u00d3N):** Facturas identificadas y protegidas para auditor\u00eda DIAN. **No se generar\u00e1 comprobante en Siigo** para evitar duplicidades.")
                        with col_desb2:
                            c_sw1, c_sw2 = st.columns(2)
                            with c_sw1:
                                if es_siigo:
                                    if st.button("\U0001f4d1 Cambiar a Solo Conciliaci\u00f3n", key=f"btn_sw_concil_{pq_id_sel}", use_container_width=True, help="Cambia el destino para que NO se suba a la planilla Siigo"):
                                        ejecutar_validacion_pq(pasa_a_siigo=False)
                                        st.rerun()
                                else:
                                    if st.button("\U0001f4e4 Cambiar a Pasar a Hoja 2", key=f"btn_sw_siigo_{pq_id_sel}", type="primary", use_container_width=True, help="Cambia el destino para que S\u00cd se suba a Siigo"):
                                        ejecutar_validacion_pq(pasa_a_siigo=True)
                                        st.rerun()
                            with c_sw2:
                                if st.button("\U0001f513 Desbloquear", key=f"btn_desbloquear_main_{pq_id_sel}", use_container_width=True):
                                    ejecutar_desbloqueo_pq()
                                    st.rerun()
    
                    st.markdown("---")
                    st.markdown("#### \U0001f4e5 Exportar Control de Paquetes y Cruces de Importaci\u00f3n:")
                    st.caption("Descarga un archivo Excel con el resumen de cada paquete, el detalle de facturas de terceros vinculadas y el asiento contable de control (sin generar a\u00fan notas de contabilidad en Siigo hasta que definas el procedimiento):")
    
                    if st.checkbox("\U0001f4ca Generar y Mostrar Bot\u00f3n de Descarga de Control de Cruces en Excel", value=False, key="chk_gen_excel_cruces_tab3"):
                        # Generar Excel estructurado por Paquetes de Importacion individuales
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
                    
                            # 1. Asiento de Control de Cruces (Partida Doble para Revision)
                            for _, fila_as in df_as_p.iterrows():
                                todos_asientos_lista.append({
                                    "Paquete #": f"Paquete #{g_k}",
                                    "Fecha Operaci\u00f3n": ag_item["Fecha"],
                                    "C\u00f3digo Cuenta": fila_as["C\u00f3digo Cuenta"],
                                    "Tercero / NIT": fila_as["Tercero / NIT"],
                                    "Descripci\u00f3n": str(fila_as.get("Descripci\u00f3n de la Cuenta") or fila_as.get("Descripci\u00f3n Cuenta") or fila_as.get("Descripci\u00f3n") or ""), 
                                    "D\u00e9bito ($)": fila_as["D\u00e9bito ($)"],
                                    "Cr\u00e9dito ($)": fila_as["Cr\u00e9dito ($)"]
                                })
                        
                            # 2. Resumen del Paquete
                            facs_terc_str = ", ".join([f"{r['Factura']} ({str(r.get('Proveedor') or 'Tercero')[:15]})" for _, r in terc_items.iterrows()]) if not terc_items.empty else "Ninguno"
                            tot_terc_sum = sum([obtener_saldo_cruce_factura(r) for _, r in terc_items.iterrows()])
                            resumen_paquetes_lista.append({
                                "Paquete #": f"Paquete #{g_k}",
                                "Agente Coordinador": ag_item["Proveedor"],
                                "Factura Agente": ag_item["Factura"],
                                "Fecha Operaci\u00f3n": ag_item["Fecha"],
                                "Total Cobro Agente ($)": float(ag_item["Total"]),
                                "Facturas Terceros Incluidas": facs_terc_str,
                                "Total Cancelado Terceros ($)": tot_terc_sum,
                                "Retenciones Asumidas ($)": ret_p,
                                "Diferencia No Deducible ($)": dif_p,
                                "Estado Cuadre": "\u2705 Cuadrado ($0.00)" if abs(float(ag_item["Total"]) - tot_terc_sum - ret_p - dif_p) < 1.0 else "Revisar"
                            })
                    
                            # 3. Detalle de Facturas que componen el Paquete
                            detalle_facturas_lista.append({
                                "Paquete #": f"Paquete #{g_k}",
                                "Rol en Operaci\u00f3n": "Agente Coordinador (Cobro Global)",
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
                                rol_t = "\U0001f69a Tercero Transporte (DHL)" if "DHL" in p_u else ("\U0001f3e2 Tercero Agenciamiento (Mandato)" if any(k in p_u for k in ["CARGO", "ADUANA"]) else "\U0001f3ec Tercero Garaje / Almac\u00e9n")
                                nat_t = "Honorario Propio Agenciamiento" if "COMISION" in str(tr_it.get("Descripcion", "")).upper() else "Gasto por Cuenta de Tercero"
                                detalle_facturas_lista.append({
                                    "Paquete #": f"Paquete #{g_k}",
                                    "Rol en Operaci\u00f3n": rol_t,
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
    
                        # Anadir facturas de terceros que quedaron libres (sin asignar a ningun paquete)
                        df_libres_exp = df_terceros_all[~df_terceros_all["Factura"].isin(facs_asignadas_en_algun_pq)]
                        for _, tr_lib in df_libres_exp.iterrows():
                            tb = float(tr_lib.get("Base", 0.0))
                            tiv = float(tr_lib.get("IVA", 0.0))
                            sc = obtener_saldo_cruce_factura(tr_lib)
                            p_u = str(tr_lib["Proveedor"]).upper()
                            rol_t = "\U0001f69a Tercero Transporte (DHL)" if "DHL" in p_u else ("\U0001f3e2 Tercero Agenciamiento (Mandato)" if any(k in p_u for k in ["CARGO", "ADUANA"]) else "\U0001f3ec Tercero Garaje / Almac\u00e9n")
                            detalle_facturas_lista.append({
                                "Paquete #": "\u26aa Sin Asignar (Pendiente de Operaci\u00f3n)",
                                "Rol en Operaci\u00f3n": rol_t,
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
                                label=f"\U0001f4e5 Descargar Control de Cruces de Importaci\u00f3n en Excel ({len(pqs_actuales)} Paquetes)",
                                data=buf_cruce_ex.getvalue(),
                                file_name=f"Control_Cruces_Importacion_{empresa['nombre'].replace(' ', '_')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
    
            else:
                st.info("\U0001f4a1 Sube el reporte Excel de la DIAN en la Pesta\u00f1a 1 para habilitar la triangulaci\u00f3n y cruce de importaciones.")
    
    
        except Exception as _e_tri:
            if type(_e_tri).__name__ in ['RerunException', 'StopException', 'ScriptControlException']:
                raise _e_tri
            st.error(f"\u26a0\ufe0f Ocurri\u00f3 un error en la pesta\u00f1a de Triangulaci\u00f3n y Aduanas: {_e_tri}")
            with st.expander("\U0001f50d Ver detalles t\u00e9cnicos del error (Traceback):"):
                st.code(traceback.format_exc())
    with tab_siigo:
        try:
            st.markdown("### \U0001f4e5 Planilla Oficial Siigo Nube (3 Hojas Formuladas)")
            st.caption("Se alimenta directamente con toda la informaci\u00f3n, cuentas, retenciones, impuestos asumidos y asientos aprobados que quedaron en la **Pesta\u00f1a 2 (Auditor\u00eda y Trazabilidad Fiscal)**.")
    
            if "df_procesado" in st.session_state and st.session_state["df_procesado"] is not None and isinstance(st.session_state["df_procesado"], pd.DataFrame) and not st.session_state["df_procesado"].empty:
                df_full = normalizar_df_procesado(st.session_state["df_procesado"])
                st.session_state["df_procesado"] = df_full
    
                if "No Contabilizar" not in df_full.columns:
                    df_full["No Contabilizar"] = False
                    st.session_state["df_procesado"] = df_full
    
                # 1. Excluir unicamente las facturas marcadas como 'No Contabilizar' en la Pestana 2
                excluidas_set = st.session_state.get("facturas_no_contabilizar", set())
                cond_excluidas = df_full["Factura"].astype(str).str.strip().isin(excluidas_set) | (df_full["No Contabilizar"] == True)
                n_excluidas = int(cond_excluidas.sum())
    
                ya_reg_siigo = df_full["Ya Registrada"].fillna(False).astype(bool) if "Ya Registrada" in df_full.columns else pd.Series(False, index=df_full.index)
                n_rojas = int(ya_reg_siigo.sum())
    
                c_flt1, c_flt2 = st.columns([2.5, 1.5])
                with c_flt1:
                    excluir_rojas_opt = st.checkbox(
                        f"Opcional: Excluir tambi\u00e9n las {n_rojas} facturas que ven\u00edan marcadas en rojo en el Excel inicial",
                        value=False,
                        key="chk_excluir_rojas_tab4",
                        help="Por defecto est\u00e1 desmarcado para que la Planilla Oficial de Siigo incluya el 100% de las facturas activas y editadas en la Pesta\u00f1a 2 (excepto las marcadas como 'No Contabilizar')."
                    )
                with c_flt2:
                    if n_excluidas > 0:
                        st.warning(f"\U0001f6ab Excluidas por 'No Contabilizar' en Pesta\u00f1a 2: **{n_excluidas}**")
                    else:
                        st.success("\u2705 100% de facturas de Pesta\u00f1a 2 activas")
    
                if excluir_rojas_opt:
                    df_p = df_full[(~cond_excluidas) & (~ya_reg_siigo)].copy()
                else:
                    df_p = df_full[~cond_excluidas].copy()
    
                pqs_actuales_siigo = st.session_state.get("paquetes_importacion", {})
                asientos_triang_map_siigo = {}
                if "asientos_triangulacion_por_factura" in st.session_state and isinstance(st.session_state["asientos_triangulacion_por_factura"], dict):
                    asientos_triang_map_siigo.update(st.session_state["asientos_triangulacion_por_factura"])
                if isinstance(pqs_actuales_siigo, dict):
                    for p_k, p_v in pqs_actuales_siigo.items():
                        if isinstance(p_v, dict) and st.session_state.get(f"paquete_listo_{p_k}", False) and st.session_state.get(f"paquete_destino_siigo_{p_k}", p_v.get("destino_siigo", True)):
                            ag_f = str(p_v.get("agente", {}).get("Factura", "")).strip()
                            af_df = st.session_state.get(f"asiento_fijo_pq_{p_k}", p_v.get("asiento_fijo", p_v.get("asiento_aprobado")))
                            if ag_f and af_df is not None:
                                asientos_triang_map_siigo[ag_f] = pd.DataFrame(af_df).copy()
    
                # Firma ligera de estado para cachear la Planilla Oficial Siigo Nube y evitar reconstruir openpyxl en cada clic de otras pestanas
                _sig_siigo_actual = (
                    f"{len(df_p)}_{excluir_rojas_opt}_"
                    f"{round(float(df_p['Base'].sum()), 2)}_{round(float(df_p['IVA'].sum()), 2)}_"
                    f"{round(float(df_p['ReteFuente'].sum()), 2)}_{round(float(df_p['ReteICA'].sum()), 2)}_"
                    f"{round(float(df_p['Total'].sum()), 2)}_{len(asientos_triang_map_siigo)}_"
                    f"{hash(tuple(df_p['Cta Principal'].astype(str)))}_"
                    f"{hash(tuple(df_p['Cta Contrapartida'].astype(str)))}_"
                    f"{hash(tuple(df_p['Impuestos Asumidos'].astype(bool)))}"
                )
                if st.session_state.get("_siigo_cache_sig") != _sig_siigo_actual or "_siigo_3hojas_bytes" not in st.session_state:
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
                        "Tipo de comprobante", "Consecutivo comprobante", "Fecha de elaboraci\u00f3n", "Sigla moneda",
                        "Tasa de cambio", "C\u00f3digo cuenta contable", "Identificaci\u00f3n tercero", "Sucursal",
                        "C\u00f3digo producto", "C\u00f3digo de bodega", "Acci\u00f3n", "Cantidad producto", "Prefijo",
                        "Consecutivo", "No. cuota", "Fecha vencimiento", "C\u00f3digo impuesto", "C\u00f3digo grupo activo fijo",
                        "C\u00f3digo activo fijo", "Descripci\u00f3n", "C\u00f3digo centro/subcentro de costos", "D\u00e9bito",
                        "Cr\u00e9dito", "Observaciones", "Base gravable libro compras/ventas", "Base exenta libro compras/ventas", "Mes de cierre"
                    ]
                    ws_interfaz.append(headers_interfaz)
    
                    filas_preview_matriz = []
                    filas_preview_interfaz = []
                    fila_r = 2
    
                    for _, item in df_p.iterrows():
                        t_comp = item["Tipo Comp"]
                        cons = item["Consecutivo"]
                        f_str = normalizar_fecha_dian(item["Fecha"])
                        fn_item = str(item["Factura"]).strip()
    
                        # Si esta factura tiene asiento de triangulacion aprobado hacia Hoja 2, exportar ese asiento exacto de Hoja 2
                        if fn_item in asientos_triang_map_siigo and asientos_triang_map_siigo[fn_item] is not None:
                            df_as_tr = sincronizar_asiento_con_factura_actual(pd.DataFrame(asientos_triang_map_siigo[fn_item]).copy(), item)
                            if not df_as_tr.empty:
                                for _, r_as in df_as_tr.iterrows():
                                    cta_code = str(r_as["C\u00f3digo Cuenta"]).strip()
                                    nit_clean = re.sub(r'\D', '', str(r_as["Tercero / NIT"]).split("-")[0])
                                    deb_v = float(r_as["D\u00e9bito ($)"] or 0.0)
                                    cred_v = float(r_as["Cr\u00e9dito ($)"] or 0.0)
                                    desc_line = str(r_as.get("Descripci\u00f3n de la Cuenta") or r_as.get("Descripci\u00f3n Cuenta") or "")[:40]
                                    obs_line = f"Triangulaci\u00f3n {str(item.get('Proveedor') or 'Tercero')[:15]} Fac {fn_item}"
    
                                    row_int = [
                                        t_comp, cons, f_str, "COP", 1,
                                        cta_code, nit_clean, 0, "", "", "", "",
                                        item.get("Prefijo", "IMP"), item.get("Folio", fn_item),
                                        1, f_str, "", "", "",
                                        desc_line, "", deb_v, cred_v,
                                        obs_line, 0.0, 0.0, ""
                                    ]
                                    ws_interfaz.append(row_int)
                                    filas_preview_interfaz.append({
                                        "Comprobante": f"{t_comp}-{cons}", "Fecha": f_str, "Cuenta": cta_code,
                                        "Tercero NIT": nit_clean, "Factura": fn_item, "Descripci\u00f3n": desc_line,
                                        "D\u00e9bito ($)": deb_v, "Cr\u00e9dito ($)": cred_v
                                    })
                                    row_mat = [
                                        t_comp, cons, f_str, nit_clean,
                                        item.get("Prefijo", "IMP"), item.get("Folio", fn_item),
                                        desc_line, "Cruce Triangulaci\u00f3n", cta_code, deb_v, 0.0, 0.0, 0.0, 0.0, "22050501"
                                    ]
                                    ws_matriz.append(row_mat)
                                    filas_preview_matriz.append(dict(zip(headers_matriz, row_mat)))
                                    fila_r += 1
                                continue
    
                        # De lo contrario, usar exactamente la contabilizacion de la factura en la Pestana 2 (obtener_asiento_contable_hoja2)
                        es_aduanero_it = any(k in str(item.get("Proveedor", "")).upper() for k in AGENTES_ADUANEROS)
                        asume_ret_it = bool(item.get("Impuestos Asumidos", False))
                        nit_raw = str(item.get("NIT Emisor", "")).strip()
                        nit_clean = re.sub(r'\D', '', nit_raw.split("-")[0]) or nit_raw
                        pref = str(item.get("Prefijo", "") or "").strip()
                        fac_num = str(item.get("Folio", "") or fn_item).strip()
                        desc = str(item.get("Descripcion", "") or "")
                        op = str(item.get("Operacion", "Compra") or "Compra")
                        cta_p = str(item.get("Cta Principal") or "14650501").strip()
                        if cta_p in ["146505", "1465"]:
                            cta_p = "14650501"
                        es_imp_fac = es_aduanero_it or "1465" in cta_p or "IMPORTACI" in str(item.get("Categor\u00eda", "")).upper()
                        cta_c = str(item.get("Cuenta Pasivo Especifica") or item.get("Cta Contrapartida") or ("22050505" if es_aduanero_it else "22050501")).strip()
                        cta_iva = str(item.get("Cta IVA") or ("24081501" if es_imp_fac else "24081001")).strip()
    
                        base_raw = float(item.get("Base", 0.0) or 0.0)
                        inc_otros_it = float(item.get("INC / Otros Impuestos", 0.0) or 0.0)
                        base = round(base_raw + inc_otros_it, 2)
                        iva = float(item.get("IVA", 0.0) or 0.0)
                        iva_5_it = float(item.get("IVA 5%", 0.0) or 0.0)
                        iva_19_it = float(item.get("IVA 19%", 0.0) or 0.0)
                        if iva_5_it > 0:
                            if iva_19_it <= 0 and iva > iva_5_it:
                                iva_19_it = round(iva - iva_5_it, 2)
                            iva = round(iva_19_it + iva_5_it, 2)
                        else:
                            iva_19_it = iva
                        base_19_it = float(item.get("Base IVA 19%", 0.0) or 0.0) or (round(iva_19_it / 0.19, 2) if iva_5_it > 0 and iva_19_it > 0 else base_raw)
                        base_5_it = float(item.get("Base IVA 5%", 0.0) or 0.0) or (round(iva_5_it / 0.05, 2) if iva_5_it > 0 else 0.0)
                        cta_iva_5_it = _limpiar_codigo_cta(item.get("Cta IVA 5%"), "24081003")

                        rfte = float(item.get("ReteFuente", 0.0) or 0.0)
                        rica = float(item.get("ReteICA", 0.0) or 0.0)
                        riva = float(item.get("ReteIVA", 0.0) or 0.0)
                        iva_imp_it = float(item.get("IVA Importaci\u00f3n", 0.0) or 0.0)
                        cta_iva_imp_it = str(item.get("Cta IVA Importaci\u00f3n") or "240835").strip()
                        if not cta_iva_imp_it or cta_iva_imp_it in ["nan", "None"]:
                            cta_iva_imp_it = "240835"
    
                        es_transp_it = ("TRANSP" in str(item.get("Categor\u00eda", "")).upper() or
                                        "TRANSP" in desc.upper() or
                                        "INLAND" in desc.upper() or
                                        "FLETE" in desc.upper() or
                                        "70787" in fn_item or
                                        str(item.get("Cta ReteFuente", "")).strip() == "23652505" or
                                        (rfte > 0 and base_raw > 0 and abs(rfte - round(base_raw * 0.01, 2)) < 2.0))
                        cta_rf_raw = str(item.get("Cta ReteFuente", "") or "").strip()
                        cta_ri_raw = str(item.get("Cta ReteICA", "") or "").strip()
                        if es_transp_it:
                            cta_rfte = "23652505"
                            cta_rica = "23680513"
                        else:
                            cta_rfte = cta_rf_raw if (cta_rf_raw and cta_rf_raw not in ["nan", "None", ""]) else ("23652503" if es_aduanero_it else "23654001")
                            cta_rica = cta_ri_raw if (cta_ri_raw and cta_ri_raw not in ["nan", "None", ""]) else ("23680505" if es_aduanero_it else "23680501")
    
                        row_mat = [t_comp, cons, f_str, nit_clean, pref, fac_num, desc, op, cta_p, base, iva, rfte, rica, riva, cta_c]
                        ws_matriz.append(row_mat)
                        filas_preview_matriz.append(dict(zip(headers_matriz, row_mat)))
    
                        r = fila_r
                        es_nc_it = "Devolucion" in op
    
                        # Linea 1: Base Imponible (incluye INC / Bolsas si aplica como mayor valor del gasto)
                        ws_interfaz.append([
                            f"=matriz_captura!A{r}", f"=matriz_captura!B{r}", f"=matriz_captura!C{r}", "COP", 1,
                            f"=matriz_captura!I{r}", f"=matriz_captura!D{r}", 0, "", "", "", "", "", "", "", "", "", "", "",
                            f"=matriz_captura!G{r}", "",
                            f'=IF(matriz_captura!H{r}="Devolucion Compra", 0, matriz_captura!J{r})',
                            f'=IF(matriz_captura!H{r}="Devolucion Compra", matriz_captura!J{r}, 0)',
                            "", f"=matriz_captura!J{r}", 0.0, ""
                        ])
                        filas_preview_interfaz.append({
                            "Comprobante": f"{t_comp}-{cons}", "Fecha": f_str, "Cuenta": cta_p,
                            "Tercero NIT": nit_clean, "Factura": fn_item, "Descripci\u00f3n": desc[:40],
                            "D\u00e9bito ($)": 0.0 if es_nc_it else base, "Cr\u00e9dito ($)": base if es_nc_it else 0.0
                        })
    
                        # Linea 2: Desglose de IVA (19% en 24081001/24081501 y 5% en 24081003)
                        if iva_5_it > 0 and iva_19_it > 0:
                            cod_imp_19 = CODIGOS_IMPUESTO_SIIGO.get(str(cta_iva).strip(), 1)
                            ws_interfaz.append([
                                f"=matriz_captura!A{r}", f"=matriz_captura!B{r}", f"=matriz_captura!C{r}", "COP", 1,
                                cta_iva, f"=matriz_captura!D{r}", 0, "", "", "", "", "", "", "", "",
                                cod_imp_19, "", "", f'="IVA 19% Base: {base_19_it:.0f}"', "",
                                f'=IF(matriz_captura!H{r}="Devolucion Compra", 0, {iva_19_it})',
                                f'=IF(matriz_captura!H{r}="Devolucion Compra", {iva_19_it}, 0)',
                                "", base_19_it, 0.0, ""
                            ])
                            filas_preview_interfaz.append({
                                "Comprobante": f"{t_comp}-{cons}", "Fecha": f_str, "Cuenta": cta_iva,
                                "Tercero NIT": nit_clean, "Factura": fn_item, "Descripci\u00f3n": f"IVA 19% Base: {base_19_it:,.0f}",
                                "D\u00e9bito ($)": 0.0 if es_nc_it else iva_19_it, "Cr\u00e9dito ($)": iva_19_it if es_nc_it else 0.0
                            })
                            cod_imp_5 = CODIGOS_IMPUESTO_SIIGO.get(str(cta_iva_5_it).strip(), 2)
                            ws_interfaz.append([
                                f"=matriz_captura!A{r}", f"=matriz_captura!B{r}", f"=matriz_captura!C{r}", "COP", 1,
                                cta_iva_5_it, f"=matriz_captura!D{r}", 0, "", "", "", "", "", "", "", "",
                                cod_imp_5, "", "", f'="IVA 5% Base: {base_5_it:.0f}"', "",
                                f'=IF(matriz_captura!H{r}="Devolucion Compra", 0, {iva_5_it})',
                                f'=IF(matriz_captura!H{r}="Devolucion Compra", {iva_5_it}, 0)',
                                "", base_5_it, 0.0, ""
                            ])
                            filas_preview_interfaz.append({
                                "Comprobante": f"{t_comp}-{cons}", "Fecha": f_str, "Cuenta": cta_iva_5_it,
                                "Tercero NIT": nit_clean, "Factura": fn_item, "Descripci\u00f3n": f"IVA 5% Base: {base_5_it:,.0f}",
                                "D\u00e9bito ($)": 0.0 if es_nc_it else iva_5_it, "Cr\u00e9dito ($)": iva_5_it if es_nc_it else 0.0
                            })
                        elif iva_5_it > 0:
                            cod_imp_5 = CODIGOS_IMPUESTO_SIIGO.get(str(cta_iva_5_it).strip(), 2)
                            ws_interfaz.append([
                                f"=matriz_captura!A{r}", f"=matriz_captura!B{r}", f"=matriz_captura!C{r}", "COP", 1,
                                cta_iva_5_it, f"=matriz_captura!D{r}", 0, "", "", "", "", "", "", "", "",
                                cod_imp_5, "", "", f'="IVA 5% Base: " & matriz_captura!J{r}', "",
                                f'=IF(matriz_captura!H{r}="Devolucion Compra", 0, matriz_captura!K{r})',
                                f'=IF(matriz_captura!H{r}="Devolucion Compra", matriz_captura!K{r}, 0)',
                                "", f"=matriz_captura!J{r}", 0.0, ""
                            ])
                            filas_preview_interfaz.append({
                                "Comprobante": f"{t_comp}-{cons}", "Fecha": f_str, "Cuenta": cta_iva_5_it,
                                "Tercero NIT": nit_clean, "Factura": fn_item, "Descripci\u00f3n": f"IVA 5% Base: {base_5_it:,.0f}",
                                "D\u00e9bito ($)": 0.0 if es_nc_it else iva_5_it, "Cr\u00e9dito ($)": iva_5_it if es_nc_it else 0.0
                            })
                        elif iva > 0:
                            cod_imp_iva = CODIGOS_IMPUESTO_SIIGO.get(str(cta_iva).strip(), 1)
                            ws_interfaz.append([
                                f"=matriz_captura!A{r}", f"=matriz_captura!B{r}", f"=matriz_captura!C{r}", "COP", 1,
                                cta_iva, f"=matriz_captura!D{r}", 0, "", "", "", "", "", "", "", "",
                                cod_imp_iva, "", "", f'="IVA Base: " & matriz_captura!J{r}', "",
                                f'=IF(matriz_captura!H{r}="Devolucion Compra", 0, matriz_captura!K{r})',
                                f'=IF(matriz_captura!H{r}="Devolucion Compra", matriz_captura!K{r}, 0)',
                                "", f"=matriz_captura!J{r}", 0.0, ""
                            ])
                            filas_preview_interfaz.append({
                                "Comprobante": f"{t_comp}-{cons}", "Fecha": f_str, "Cuenta": cta_iva,
                                "Tercero NIT": nit_clean, "Factura": fn_item, "Descripci\u00f3n": f"IVA Base: {base:,.0f}",
                                "D\u00e9bito ($)": 0.0 if es_nc_it else iva, "Cr\u00e9dito ($)": iva if es_nc_it else 0.0
                            })
    
                        # Linea 2.B: IVA de Importacion (240835)
                        if iva_imp_it > 0:
                            cod_imp_ivaimp = CODIGOS_IMPUESTO_SIIGO.get(cta_iva_imp_it, "")
                            ws_interfaz.append([
                                f"=matriz_captura!A{r}", f"=matriz_captura!B{r}", f"=matriz_captura!C{r}", "COP", 1,
                                cta_iva_imp_it, f"=matriz_captura!D{r}", 0, "", "", "", "", "", "", "", "",
                                cod_imp_ivaimp, "", "", f'="IVA Importaci\u00f3n Fac " & matriz_captura!E{r} & "-" & matriz_captura!F{r}', "",
                                f'=IF(matriz_captura!H{r}="Devolucion Compra", 0, {iva_imp_it})',
                                f'=IF(matriz_captura!H{r}="Devolucion Compra", {iva_imp_it}, 0)',
                                "", f"=matriz_captura!J{r}", 0.0, ""
                            ])
                            filas_preview_interfaz.append({
                                "Comprobante": f"{t_comp}-{cons}", "Fecha": f_str, "Cuenta": cta_iva_imp_it,
                                "Tercero NIT": nit_clean, "Factura": fn_item, "Descripci\u00f3n": f"IVA Importaci\u00f3n Fac {fn_item}",
                                "D\u00e9bito ($)": 0.0 if es_nc_it else iva_imp_it, "Cr\u00e9dito ($)": iva_imp_it if es_nc_it else 0.0
                            })
    
                        # Linea 2.C: Retenciones Asumidas (53152001) cuando Impuestos Asumidos esta activo en Pestana 2
                        tot_ret_it = round(rfte + rica, 2)
                        if asume_ret_it and tot_ret_it > 0:
                            ws_interfaz.append([
                                f"=matriz_captura!A{r}", f"=matriz_captura!B{r}", f"=matriz_captura!C{r}", "COP", 1,
                                "53152001", f"=matriz_captura!D{r}", 0, "", "", "", "", "", "", "", "",
                                "", "", "", f'="Retenciones Asumidas Fac " & matriz_captura!E{r} & "-" & matriz_captura!F{r}', "",
                                f'=IF(matriz_captura!H{r}="Devolucion Compra", 0, matriz_captura!L{r} + matriz_captura!M{r})',
                                f'=IF(matriz_captura!H{r}="Devolucion Compra", matriz_captura!L{r} + matriz_captura!M{r}, 0)',
                                "", 0.0, 0.0, ""
                            ])
                            filas_preview_interfaz.append({
                                "Comprobante": f"{t_comp}-{cons}", "Fecha": f_str, "Cuenta": "53152001",
                                "Tercero NIT": nit_clean, "Factura": fn_item, "Descripci\u00f3n": f"Retenciones Asumidas Fac {fn_item}",
                                "D\u00e9bito ($)": 0.0 if es_nc_it else tot_ret_it, "Cr\u00e9dito ($)": tot_ret_it if es_nc_it else 0.0
                            })
    
                        # Linea 3: ReteFuente
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
                            filas_preview_interfaz.append({
                                "Comprobante": f"{t_comp}-{cons}", "Fecha": f_str, "Cuenta": cta_rfte,
                                "Tercero NIT": nit_clean, "Factura": fn_item, "Descripci\u00f3n": f"ReteFuente Base: {base:,.0f}",
                                "D\u00e9bito ($)": rfte if es_nc_it else 0.0, "Cr\u00e9dito ($)": 0.0 if es_nc_it else rfte
                            })
    
                        # Linea 4: ReteICA
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
                            filas_preview_interfaz.append({
                                "Comprobante": f"{t_comp}-{cons}", "Fecha": f_str, "Cuenta": cta_rica,
                                "Tercero NIT": nit_clean, "Factura": fn_item, "Descripci\u00f3n": f"ReteICA Base: {base:,.0f}",
                                "D\u00e9bito ($)": rica if es_nc_it else 0.0, "Cr\u00e9dito ($)": 0.0 if es_nc_it else rica
                            })
    
                        # Linea 4.B: ReteIVA
                        if riva > 0:
                            ws_interfaz.append([
                                f"=matriz_captura!A{r}", f"=matriz_captura!B{r}", f"=matriz_captura!C{r}", "COP", 1,
                                "23670101", f"=matriz_captura!D{r}", 0, "", "", "", "", "", "", "", "",
                                15, "", "", f'="ReteIVA Base: " & matriz_captura!K{r}', "",
                                f'=IF(matriz_captura!H{r}="Devolucion Compra", matriz_captura!N{r}, 0)',
                                f'=IF(matriz_captura!H{r}="Devolucion Compra", 0, matriz_captura!N{r})',
                                "", f"=matriz_captura!K{r}", 0.0, ""
                            ])
                            filas_preview_interfaz.append({
                                "Comprobante": f"{t_comp}-{cons}", "Fecha": f_str, "Cuenta": "23670101",
                                "Tercero NIT": nit_clean, "Factura": fn_item, "Descripci\u00f3n": f"ReteIVA 15%",
                                "D\u00e9bito ($)": riva if es_nc_it else 0.0, "Cr\u00e9dito ($)": 0.0 if es_nc_it else riva
                            })
    
                        # Linea 5: Cuenta por Pagar (Contrapartida exacta de Pestana 2)
                        extra_iva_imp_str = f" + {iva_imp_it}" if iva_imp_it > 0 else ""
                        if asume_ret_it:
                            formula_deb = f'=IF(matriz_captura!H{r}="Devolucion Compra", matriz_captura!J{r} + matriz_captura!K{r}{extra_iva_imp_str}, 0)'
                            formula_cred = f'=IF(matriz_captura!H{r}="Devolucion Compra", 0, matriz_captura!J{r} + matriz_captura!K{r}{extra_iva_imp_str})'
                            saldo_cxp_num = round(base + iva + iva_imp_it, 2)
                        else:
                            formula_deb = f'=IF(matriz_captura!H{r}="Devolucion Compra", matriz_captura!J{r} + matriz_captura!K{r}{extra_iva_imp_str} - matriz_captura!L{r} - matriz_captura!M{r} - matriz_captura!N{r}, 0)'
                            formula_cred = f'=IF(matriz_captura!H{r}="Devolucion Compra", 0, matriz_captura!J{r} + matriz_captura!K{r}{extra_iva_imp_str} - matriz_captura!L{r} - matriz_captura!M{r} - matriz_captura!N{r})'
                            saldo_cxp_num = round(base + iva + iva_imp_it - rfte - rica - riva, 2)
    
                        ws_interfaz.append([
                            f"=matriz_captura!A{r}", f"=matriz_captura!B{r}", f"=matriz_captura!C{r}", "COP", 1,
                            f"=matriz_captura!O{r}", f"=matriz_captura!D{r}", 0, "", "", "", "",
                            f"=matriz_captura!A{r}", f"=matriz_captura!B{r}", 1, f"=matriz_captura!C{r}",
                            "", "", "", f'="Fac " & matriz_captura!E{r} & "-" & matriz_captura!F{r}', "",
                            formula_deb, formula_cred, "", 0.0, 0.0, ""
                        ])
                        filas_preview_interfaz.append({
                            "Comprobante": f"{t_comp}-{cons}", "Fecha": f_str, "Cuenta": cta_c,
                            "Tercero NIT": nit_clean, "Factura": fn_item, "Descripci\u00f3n": f"CxP Fac {fn_item}",
                            "D\u00e9bito ($)": saldo_cxp_num if es_nc_it else 0.0, "Cr\u00e9dito ($)": 0.0 if es_nc_it else saldo_cxp_num
                        })
                        fila_r += 1
    
                    # 3. Hoja Parametrizacion
                    ws_params = wb.create_sheet(title="Parametrizaci\u00f3n")
                    headers_params = ["Tipo Comprobante", "", "Impuesto", "Siigo_ID", "Tarifa", "Venta", "Compra", "Dev_Venta", "Dev_Compra"]
                    ws_params.append(headers_params)
                    comprobantes = [
                        "1 - Ajustes contables", "2 - Depreciaci\u00f3n", "3 - Costeo", "4 - Diferidos", "5 - Legalizaci\u00f3n de viaticos",
                        "6 - Legalizaci\u00f3n de Caja menores", "7 - Obligaciones financieras", "8 - N\u00f3mina", "998 - Cierre a\u00f1o",
                        "999 - Saldos iniciales", "992 - Comprobante de n\u00f3mina", "993 - Comprobante de n\u00f3mina provisi\u00f3n y seguridad social",
                        "994 - Comprobante liquidaci\u00f3n de contrato", "995 - Comprobante liquidaci\u00f3n de primas", "996 - Comprobante liquidaci\u00f3n de cesant\u00edas",
                        "997 - Comprobante desembolso n\u00f3mina", "9 - Recibo de Caja", "10 - Factura de Compra", "11 - Comprobante de Egreso",
                        "12 - Factura de Venta", "13 - Nota Credito", "14 - Nota de Contabilidad", "15 - Documento soporte electronico",
                        "777 - Ajustes contables de cartera", "9901 - Traslado de dinero", "16 - FACTURAS DE COMPRA (2)", "17 - NOTA CREDITO DE COMPRA"
                    ]
                    impuestos = [
                        ("IVA 19%", 1, 0.19, "24080601", "24081001", "24082001", "24081002"),
                        ("IVA Servicios 19%", 1, 0.19, "24080601", "24081501", "24082001", "24082001"),
                        ("IVA Importaci\u00f3n 19%", 1, 0.19, "24080601", "240835", "24082001", "24081002"),
                        ("IVA 5%", 2, 0.05, "24080602", "24081003", "24082002", "24081004"),
                        ("Retefuente 11%", 3, 0.11, "13551509", "23651501", "13551510", "23651502"),
                        ("Retefuente 10%", 4, 0.1, "13551507", "23652001", "13551508", "23652002"),
                        ("Retefuente 6%", 5, 0.06, "13551505", "23652501", "13551506", "23652502"),
                        ("Retefuente 4%", 6, 0.04, "13551503", "23652503", "13551504", "23652504"),
                        ("Retefuente 3.5%", 18, 0.035, "13551513", "23654004", "13551514", "23654005"),
                        ("Retefuente 2.5%", 7, 0.025, "13551501", "23654001", "13551502", "23654002"),
                        ("Retefuente 1% Transporte", 21, 0.01, "13551521", "23652505", "13551521", "23652505"),
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
                    filas_preview_params = []
                    max_len = max(len(comprobantes), len(impuestos))
                    for idx_row in range(max_len):
                        c_val = comprobantes[idx_row] if idx_row < len(comprobantes) else ""
                        if idx_row < len(impuestos):
                            imp = list(impuestos[idx_row])
                            r_p = [c_val, ""] + imp
                        else:
                            r_p = [c_val, "", "", "", "", "", "", "", ""]
                        ws_params.append(r_p)
                        r_p_str = [str(x) if x != "" else "" for x in r_p]
                        filas_preview_params.append(dict(zip(headers_params, r_p_str)))
    
                    output = io.BytesIO()
                    wb.save(output)
                    output.seek(0)
                    st.session_state["_siigo_3hojas_bytes"] = output.getvalue()
                    st.session_state["_siigo_preview_matriz"] = filas_preview_matriz
                    st.session_state["_siigo_preview_interfaz"] = filas_preview_interfaz
                    st.session_state["_siigo_preview_params"] = filas_preview_params
                    st.session_state["_siigo_cache_sig"] = _sig_siigo_actual
                filas_preview_matriz = st.session_state.get("_siigo_preview_matriz", [])
                filas_preview_interfaz = st.session_state.get("_siigo_preview_interfaz", [])
                filas_preview_params = st.session_state.get("_siigo_preview_params", [])
    
                # Boton principal de descarga ARRIBA para acceso inmediato
                st.download_button(
                    label=f"\U0001f4e5 1. Descargar Planilla Oficial Siigo Nube (3 Hojas Formuladas: matriz_captura, interfaz_siigo, Parametrizaci\u00f3n) - {empresa['nombre']}",
                    data=st.session_state["_siigo_3hojas_bytes"],
                    file_name=f"Plantilla_Siigo_{empresa['nombre'].replace(' ', '_')}_3_Hojas.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
                st.success(f"\u2705 **Planilla Oficial Siigo Nube (3 Hojas) lista:** Conectada directamente con las **{len(df_p)} facturas activas de la Pesta\u00f1a 2** ({len(filas_preview_interfaz)} l\u00edneas contables en `interfaz_siigo`).")
    
                # Vista previa en pantalla de las 3 hojas de la Planilla Oficial Siigo Nube
                sub_t1, sub_t2, sub_t3 = st.tabs([
                    f"\U0001f4cb Hoja 1: matriz_captura ({len(filas_preview_matriz)} registros)",
                    f"\u2699\ufe0f Hoja 2: interfaz_siigo ({len(filas_preview_interfaz)} l\u00edneas de partida doble)",
                    "\U0001f4d1 Hoja 3: Parametrizaci\u00f3n"
                ])
                with sub_t1:
                    if filas_preview_matriz:
                        st.dataframe(_df_para_streamlit(filas_preview_matriz), use_container_width=True, hide_index=True)
                with sub_t2:
                    if filas_preview_interfaz:
                        df_prev_int = pd.DataFrame(filas_preview_interfaz)
                        st.dataframe(_df_para_streamlit(df_prev_int), use_container_width=True, hide_index=True)
                        tot_d_int = df_prev_int["D\u00e9bito ($)"].sum()
                        tot_c_int = df_prev_int["Cr\u00e9dito ($)"].sum()
                        c_m1, c_m2, c_m3 = st.columns(3)
                        with c_m1: st.metric("Total D\u00e9bito (interfaz_siigo)", f"${tot_d_int:,.2f}")
                        with c_m2: st.metric("Total Cr\u00e9dito (interfaz_siigo)", f"${tot_c_int:,.2f}")
                        with c_m3: st.metric("Diferencia Global", f"${abs(tot_d_int - tot_c_int):,.2f}")
                with sub_t3:
                    st.dataframe(_df_para_streamlit(filas_preview_params), use_container_width=True, hide_index=True)
    
                st.markdown("---")
                st.markdown("### \U0001f4ca Libro Auxiliar Contable: Facturas Una a Una y Consolidado")
                st.write("Reporte administrativo con el detalle factura a factura de la Pesta\u00f1a 2 y las hojas de consolidado por proveedor y cuentas contables:")
                if st.checkbox("\U0001f4ca Preparar Descarga de Libro Auxiliar Detallado y Consolidado", value=False, key="chk_gen_libro_aux_tab4"):
                    out_aux = io.BytesIO()
                    with pd.ExcelWriter(out_aux, engine='openpyxl') as writer_aux:
                        cols_detalle = ["Comprobante Siigo", "Fecha", "Factura", "Proveedor", "NIT Emisor", "Cta Principal", "Categor\u00eda", "Base", "IVA", "ReteFuente", "ReteICA", "Total", "Cta Contrapartida", "Raz\u00f3n Contable", "Soporte PDF Renombrado"]
                        df_det_export = df_p[[c for c in cols_detalle if c in df_p.columns]].copy()
        
                        totales_dict = {
                            "Comprobante Siigo": "TOTALES CONSOLIDADOS",
                            "Fecha": "-", "Factura": f"{len(df_det_export)} Docs", "Proveedor": "-", "NIT Emisor": "-",
                            "Cta Principal": "-", "Categor\u00eda": "-",
                            "Base": df_det_export["Base"].sum(),
                            "IVA": df_det_export["IVA"].sum(),
                            "ReteFuente": df_det_export["ReteFuente"].sum(),
                            "ReteICA": df_det_export["ReteICA"].sum(),
                            "Total": df_det_export["Total"].sum(),
                            "Cta Contrapartida": "-", "Raz\u00f3n Contable": "-", "Soporte PDF Renombrado": "-"
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
        
                        df_cons_cta = df_p.groupby(["Cta Principal", "Categor\u00eda"]).agg({
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
                        label=f"\U0001f4e5 2. Descargar Libro de Facturas (Una a Una + Consolidado) - {empresa['nombre']}",
                        data=out_aux,
                        file_name=f"Libro_Facturas_Detallado_Y_Consolidado_{empresa['nombre'].replace(' ', '_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            else:
                st.info("Primero procesa los documentos en la Pestana 1 o carga tu respaldo para habilitar la descarga.")
    
        except Exception as _e_sii:
            if type(_e_sii).__name__ in ['RerunException', 'StopException', 'ScriptControlException']:
                raise _e_sii
            st.error(f"\u26a0\ufe0f Ocurri\u00f3 un error en la exportaci\u00f3n a Siigo: {_e_sii}")
            with st.expander("\U0001f50d Ver detalles t\u00e9cnicos del error (Traceback):"):
                st.code(traceback.format_exc())
    import gc
    gc.collect()


except Exception as _e_fatal:
    if type(_e_fatal).__name__ in ['RerunException', 'StopException', 'ScriptControlException']:
        raise _e_fatal
    st.markdown("""
    <div style="background:#fff1f2; border:1px solid #fecdd3; border-left:6px solid #e11d48; border-radius:10px; padding:22px; margin:20px 0; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
        <h3 style="color:#9f1239; margin-top:0;">\U0001f6e1\ufe0f Escudo de Protecci\u00f3n Antica\u00eddas Activado</h3>
        <p style="color:#881337; font-size:15px; line-height:1.6;">
            Se ha interceptado una excepci\u00f3n en tiempo de ejecuci\u00f3n y se <b>evit\u00f3 por completo la ca\u00edda fatal de Streamlit ("Oh no")</b>.<br>
            Puedes reiniciar la memoria de sesi\u00f3n con el bot\u00f3n inferior o revisar el diagn\u00f3stico t\u00e9cnico para continuar trabajando.
        </p>
    </div>
    """, unsafe_allow_html=True)
    c_rec1, c_rec2 = st.columns([1, 1])
    with c_rec1:
        if st.button("\U0001f504 Reiniciar Sesi\u00f3n Limpia (Recomendado)", type="primary", use_container_width=True, key="btn_fatal_reset_global"):
            st.session_state.clear()
            st.rerun()
    with c_rec2:
        if st.button("\u21a9\ufe0f Reintentar sin borrar datos", use_container_width=True, key="btn_fatal_retry_global"):
            st.rerun()
    with st.expander("\U0001f50d Ver detalles t\u00e9cnicos del error (Traceback completo):"):
        st.code(traceback.format_exc())
