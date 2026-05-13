import streamlit as st
import pandas as pd
import io
from datetime import datetime, date
import warnings
warnings.filterwarnings('ignore')

# ── Configuración de página ──────────────────────────────────────────────────
st.set_page_config(
    page_title="Gestión Farmacia",
    page_icon="💊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS personalizado ────────────────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=DM+Sans:wght@300;400;500;600&display=swap');

  html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

  h1, h2, h3 { font-family: 'DM Serif Display', serif; }

  /* Sidebar */
  [data-testid="stSidebar"] {
      background: linear-gradient(160deg, #0f2027, #203a43, #2c5364);
      color: white;
  }
  [data-testid="stSidebar"] * { color: white !important; }
  [data-testid="stSidebar"] .stRadio label { font-weight: 500; font-size: 1rem; }

  /* Métricas */
  [data-testid="metric-container"] {
      background: #f8f9fa;
      border: 1px solid #e9ecef;
      border-radius: 12px;
      padding: 1rem;
  }

  /* Botón de descarga */
  .stDownloadButton button {
      background: #2c5364 !important;
      color: white !important;
      border-radius: 8px !important;
      font-weight: 600 !important;
      border: none !important;
  }
  .stDownloadButton button:hover {
      background: #203a43 !important;
      transform: translateY(-1px);
      box-shadow: 0 4px 12px rgba(0,0,0,0.15) !important;
  }
  .st-emotion-cache-n5r31u{
      color:black ! important;
  }

  /* Badges de alerta */
  .badge-danger  { background:#dc3545; color:white; padding:2px 8px; border-radius:12px; font-size:.75rem; font-weight:600; }
  .badge-warning { background:#fd7e14; color:white; padding:2px 8px; border-radius:12px; font-size:.75rem; font-weight:600; }
  .badge-ok      { background:#198754; color:white; padding:2px 8px; border-radius:12px; font-size:.75rem; font-weight:600; }

  /* Tabla limpia */
  .stDataFrame { border-radius: 10px; overflow: hidden; }
</style>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
#  FUNCIONES AUXILIARES
# ════════════════════════════════════════════════════════════════════════════

COLUMNAS_INTERES = ['Código', 'Descripción', 'Familia', 'S.Actual', 'S.Minimo', 'S.Maximo', 'Rotación', 'Caducidad', 'Uds.Vendidas']

@st.cache_data(show_spinner=False)
def cargar_excel(uploaded_file):
    """
    Carga y normaliza el Excel del inventario.
    Devuelve dos DataFrames:
      - df_raw:   original completo (para la pestaña de duplicados)
      - df_clean: sin duplicados (para todos los análisis de stock y caducidad)
    """
    df = pd.read_excel(uploaded_file, dtype=str)
    df.columns = df.columns.str.strip()

    # Convertir columnas numéricas
    for col in ['S.Actual', 'S.Minimo', 'S.Maximo', 'Rotación', 'Uds.Vendidas']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].str.replace(',', '.', regex=False), errors='coerce')

    # Marcar duplicados en el df original (para mostrarlos en su pestaña)
    if 'Código' in df.columns:
        df['_duplicado'] = df.duplicated(subset='Código', keep=False)
        n_dup_filas   = int(df['_duplicado'].sum())
        n_dup_codigos = int(df[df['_duplicado']]['Código'].nunique())
        # df limpio: primera ocurrencia de cada código
        df_clean = df[~df.duplicated(subset='Código', keep='first')].copy()
    else:
        df['_duplicado'] = False
        n_dup_filas, n_dup_codigos = 0, 0
        df_clean = df.copy()

    df_clean = df_clean.drop(columns=['_duplicado'], errors='ignore')

    return df, df_clean, n_dup_filas, n_dup_codigos


def preparar_caducidad(df: pd.DataFrame) -> pd.DataFrame:
    """Añade columna Caducidad_dt (datetime) al df."""
    cols = [c for c in COLUMNAS_INTERES if c in df.columns]
    d = df[cols].copy()
    d['Caducidad'] = d['Caducidad'].fillna('Unknown')

    # Intentar múltiples formatos de fecha
    def parse_fecha(val):
        for fmt in ('%d/%m/%y', '%d/%m/%Y', '%Y-%m-%d', '%m/%d/%Y', '%m/%d/%y'):
            try:
                return pd.to_datetime(val, format=fmt)
            except Exception:
                pass
        return pd.NaT

    mask = d['Caducidad'] != 'Unknown'
    d.loc[mask, 'Caducidad_dt'] = d.loc[mask, 'Caducidad'].apply(parse_fecha)
    return d


def exportar_excel_multihoja(hojas: dict) -> bytes:
    """Recibe {nombre_hoja: dataframe} y devuelve bytes de Excel."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        for nombre, df_hoja in hojas.items():
            df_hoja.to_excel(writer, sheet_name=nombre[:31], index=False)
    return buf.getvalue()


def exportar_excel_simple(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    df.to_excel(buf, index=False, engine='openpyxl')
    return buf.getvalue()


# ════════════════════════════════════════════════════════════════════════════
#  SIDEBAR – NAVEGACIÓN Y CARGA DE ARCHIVO
# ════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("## 💊 Farmacia Manager")
    st.markdown("---")

    uploaded_file = st.file_uploader(
        "📂 Cargar inventario Excel",
        type=["xlsx", "xls"],
        help="Sube el archivo de inventario de la farmacia"
    )

    st.markdown("---")
    pagina = st.radio(
        "Navegar a:",
        ["🏠 Inicio / Resumen",
         "📅 Caducidades",
         "📦 Gestión de Stock",
         "🔍 Duplicados",
         "📊 Análisis Total"],
        label_visibility="collapsed"
    )

# ════════════════════════════════════════════════════════════════════════════
#  SIN ARCHIVO → PANTALLA DE BIENVENIDA
# ════════════════════════════════════════════════════════════════════════════

if uploaded_file is None:
    st.markdown("# 💊 Farmacia Manager")
    st.markdown("### Gestión inteligente de inventario farmacéutico")
    st.info("👈 **Empieza cargando tu archivo Excel** de inventario en el panel lateral.")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("#### 📅 Caducidades\nDetecta productos caducados o próximos a caducar en los meses que elijas.")
    with col2:
        st.markdown("#### 📦 Stock\nControla niveles mínimos, rotación cero y productos sin información de caducidad.")
    with col3:
        st.markdown("#### 📊 Análisis\nVisión global del inventario: productos con más rotación, sin caducidad, etc.")
    st.stop()


# ════════════════════════════════════════════════════════════════════════════
#  CARGA DE DATOS
# ════════════════════════════════════════════════════════════════════════════

with st.spinner("Cargando inventario..."):
    df_raw, df_clean, n_dup_filas, n_dup_codigos = cargar_excel(uploaded_file)

n_total = len(df_clean)   # total sin duplicados, la cifra real de productos únicos
n_dup   = n_dup_codigos

# ════════════════════════════════════════════════════════════════════════════
#  PÁGINA: INICIO / RESUMEN
# ════════════════════════════════════════════════════════════════════════════

if pagina == "🏠 Inicio / Resumen":
    st.markdown("# 🏠 Resumen del Inventario")
    st.markdown(f"Archivo cargado: **{uploaded_file.name}** · {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    st.markdown("---")

    df_cad = preparar_caducidad(df_clean)
    hoy = pd.Timestamp.today().normalize()

    caducados = df_cad[(df_cad['Caducidad_dt'] < hoy) & (df_cad['S.Actual'].fillna(0) > 0)]
    sin_caducidad = df_cad[(df_cad['Caducidad'] == 'Unknown') & (df_cad['S.Actual'].fillna(0) > 0)]
    sin_rotacion = df_cad[(df_cad['Uds.Vendidas'].fillna(0) == 0) & (df_cad['S.Actual'].fillna(0) != 0)]

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("📦 Productos únicos", n_total)
    c2.metric("⚠️ Ya caducados", len(caducados))
    c3.metric("❓ Sin fecha caducidad", len(sin_caducidad))
    c4.metric("🔄 Sin rotación (con stock)", len(sin_rotacion))
    c5.metric("🔁 Duplicados eliminados", n_dup)

    if n_dup > 0:
        st.info(f"ℹ️ Se han eliminado **{n_dup_filas} filas duplicadas** ({n_dup} códigos) antes del análisis. Puedes verlos en la pestaña 🔍 Duplicados.")

    st.markdown("---")
    st.markdown("#### Familias en el inventario")
    if 'Familia' in df_clean.columns:
        fam_count = df_clean['Familia'].value_counts().reset_index()
        fam_count.columns = ['Familia', 'Nº productos']
        st.dataframe(fam_count, use_container_width=True, height=300)


# ════════════════════════════════════════════════════════════════════════════
#  PÁGINA: CADUCIDADES
# ════════════════════════════════════════════════════════════════════════════

elif pagina == "📅 Caducidades":
    st.markdown("# 📅 Gestión de Caducidades")
    st.markdown("---")

    df_cad = preparar_caducidad(df_clean)
    hoy = pd.Timestamp.today().normalize()

    tab1, tab2, tab3 = st.tabs(["⛔ Ya caducados", "⏳ Próximos a caducar", "❓ Sin fecha"])

    # ── TAB 1: Ya caducados ──────────────────────────────────────────────
    with tab1:
        df_caducados = df_cad[
            (df_cad['Caducidad_dt'].notna()) &
            (df_cad['Caducidad_dt'] < hoy) &
            (df_cad['S.Actual'].fillna(0) > 0)
        ].copy()
        df_caducados['Días caducado'] = (hoy - df_caducados['Caducidad_dt']).dt.days

        st.markdown(f"**{len(df_caducados)} productos caducados con stock actual > 0**")

        if df_caducados.empty:
            st.success("✅ ¡No hay productos caducados con stock!")
        else:
            st.warning(f"⚠️ Hay **{len(df_caducados)}** productos caducados en stock.")
            cols_show = [c for c in COLUMNAS_INTERES if c in df_caducados.columns] + ['Días caducado']
            st.dataframe(df_caducados[cols_show].sort_values('Días caducado', ascending=False),
                         use_container_width=True)
            st.download_button(
                "⬇️ Exportar caducados a Excel",
                data=exportar_excel_simple(df_caducados[cols_show].sort_values('Días caducado', ascending=False)),
                file_name="productos_caducados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # ── TAB 2: Próximos a caducar ────────────────────────────────────────
    with tab2:
        st.markdown("#### Selecciona el horizonte de búsqueda")
        col_a, col_b = st.columns([1, 3])
        with col_a:
            meses = st.number_input(
                "Meses hacia adelante",
                min_value=1, max_value=36, value=6, step=1,
                help="Productos que caducan dentro de los próximos N meses"
            )

        fecha_limite = hoy + pd.DateOffset(months=int(meses))

        df_proximos = df_cad[
            (df_cad['Caducidad_dt'].notna()) &
            (df_cad['Caducidad_dt'] >= hoy) &
            (df_cad['Caducidad_dt'] <= fecha_limite) &
            (df_cad['S.Actual'].fillna(0) > 0)
        ].copy()
        df_proximos = df_proximos.sort_values('Caducidad_dt', ascending=True)
        df_proximos['Días restantes'] = (df_proximos['Caducidad_dt'] - hoy).dt.days

        st.markdown(f"**Productos que caducan antes del {fecha_limite.strftime('%d/%m/%Y')}:** {len(df_proximos)}")

        if df_proximos.empty:
            st.success(f"✅ No hay productos que caduquen en los próximos {meses} meses.")
        else:
            if meses <= 2:
                st.error(f"🔴 {len(df_proximos)} productos caducan en menos de {meses} mes(es).")
            elif meses <= 6:
                st.warning(f"🟠 {len(df_proximos)} productos caducan en los próximos {meses} meses.")
            else:
                st.info(f"🔵 {len(df_proximos)} productos caducan en los próximos {meses} meses.")

            cols_show = [c for c in COLUMNAS_INTERES if c in df_proximos.columns] + ['Días restantes']
            st.dataframe(df_proximos[cols_show], use_container_width=True)
            st.download_button(
                f"⬇️ Exportar próximos a caducar ({meses} meses) a Excel",
                data=exportar_excel_simple(df_proximos[cols_show]),
                file_name=f"caducan_proximos_{meses}_meses.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # ── TAB 3: Sin fecha de caducidad ────────────────────────────────────
    with tab3:
        df_sin = df_cad[(df_cad['Caducidad'] == 'Unknown') & (df_cad['S.Actual'].fillna(0) > 0)].copy()
        st.markdown(f"**{len(df_sin)} productos con stock pero sin fecha de caducidad registrada**")
        cols_show = [c for c in COLUMNAS_INTERES if c in df_sin.columns]
        st.dataframe(df_sin[cols_show], use_container_width=True)
        if not df_sin.empty:
            st.download_button(
                "⬇️ Exportar sin fecha de caducidad a Excel",
                data=exportar_excel_simple(df_sin[cols_show]),
                file_name="sin_fecha_caducidad.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # ── Exportar TODO en un solo Excel ──────────────────────────────────
    st.markdown("---")
    st.markdown("#### 📥 Exportar informe completo de caducidades")

    df_prox_export = df_cad[
        (df_cad['Caducidad_dt'].notna()) &
        (df_cad['Caducidad_dt'] >= hoy) &
        (df_cad['Caducidad_dt'] <= hoy + pd.DateOffset(months=int(meses))) &
        (df_cad['S.Actual'].fillna(0) > 0)
    ].copy()
    df_prox_export['Días restantes'] = (df_prox_export['Caducidad_dt'] - hoy).dt.days

    cols_show_all = [c for c in COLUMNAS_INTERES if c in df_cad.columns]
    hojas = {
        'Caducados': df_caducados[cols_show_all + ['Días caducado']] if not df_caducados.empty else pd.DataFrame(),
        f'Caducan {meses} meses': df_prox_export[cols_show_all + ['Días restantes']],
        'Sin fecha caducidad': df_sin[cols_show_all],
    }
    st.download_button(
        "⬇️ Descargar informe completo (multi-hoja)",
        data=exportar_excel_multihoja(hojas),
        file_name="informe_caducidades_completo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ════════════════════════════════════════════════════════════════════════════
#  PÁGINA: GESTIÓN DE STOCK
# ════════════════════════════════════════════════════════════════════════════

elif pagina == "📦 Gestión de Stock":
    st.markdown("# 📦 Gestión de Stock")
    st.markdown("---")

    FAMILIAS_PRESELECCIONADAS = ['DIETETICA Y HERBORISTERIA', 'SOLARES', 'COSMÉTICA', 'CORPORAL']

    todas_familias = sorted(df_raw['Familia'].dropna().unique().tolist()) if 'Familia' in df_raw.columns else []

    st.markdown("#### Selecciona las familias de interés")
    familias_sel = st.multiselect(
        "Familias",
        options=todas_familias,
        default=[f for f in FAMILIAS_PRESELECCIONADAS if f in todas_familias],
        help="Las familias preseleccionadas son las de interés habitual. Puedes añadir o quitar."
    )

    if not familias_sel:
        st.warning("Selecciona al menos una familia.")
        st.stop()

    df_cad = preparar_caducidad(df_raw)

    tab1, tab2, tab3, tab4 = st.tabs([
        "📉 Stock bajo mínimo",
        "🔄 Sin rotación (con stock)",
        "❓ Sin caducidad registrada",
        "📥 Exportar todo"
    ])

    # ── Filtro base por familias ─────────────────────────────────────────
    df_fam = df_cad[df_cad['Familia'].isin(familias_sel)].copy()

    # ── TAB 1: Stock bajo mínimo ─────────────────────────────────────────
    with tab1:
        if 'S.Actual' in df_fam.columns: #and 'S.Minimo' in df_fam.columns:
            df_bajo = df_fam[
                #df_fam['S.Actual'].notna() &
                #df_fam['S.Actual'].notna() &
                #(df_fam['S.Actual'] < df_fam['S.Minimo']) &
                (df_fam['S.Actual'] == 1 ) &
                (df_fam['Uds.Vendidas'].fillna(0) > 0)
            ].copy()
            #df_bajo['Diferencia'] = df_bajo['S.Minimo'] - df_bajo['S.Actual']
            st.markdown(f"**{len(df_bajo)} productos de familia seleccionados con stock = 1 y uds vendidas > 0**")
            if df_bajo.empty:
                st.success("✅ Todos los productos de las familias seleccionadas están sobre el mínimo.")
            else:
                st.error(f"🔴 {len(df_bajo)} productos necesitan reposición.")
                cols_show = [c for c in COLUMNAS_INTERES if c in df_bajo.columns] #+ ['Diferencia']
                st.dataframe(df_bajo[cols_show], use_container_width=True) #.sort_values('Diferencia', ascending=False)
                st.download_button(
                    "⬇️ Exportar stock bajo mínimo",
                    data=exportar_excel_simple(df_bajo[cols_show]),
                    file_name="stock_bajo_minimo.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.info("El archivo no contiene columnas S.Actual / S.Minimo.")

    # ── TAB 2: Sin rotación con stock ────────────────────────────────────
    with tab2:
        df_sin_rot = df_fam[
            (df_fam['S.Actual'].fillna(0) != 0) &
            (df_fam['Uds.Vendidas'].fillna(0) == 0)
        ].drop_duplicates().copy()
        st.markdown(f"**{len(df_sin_rot)} productos de las familias seleccionadas con stock y rotación = 0**")
        if df_sin_rot.empty:
            st.success("✅ No hay productos sin rotación.")
        else:
            st.warning(f"🟠 {len(df_sin_rot)} productos parados en stock.")
            cols_show = [c for c in COLUMNAS_INTERES if c in df_sin_rot.columns]
            st.dataframe(df_sin_rot[cols_show], use_container_width=True)
            st.download_button(
                "⬇️ Exportar sin rotación",
                data=exportar_excel_simple(df_sin_rot[cols_show]),
                file_name="sin_rotacion.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # ── TAB 3: Sin caducidad registrada ─────────────────────────────────
    with tab3:
        df_sin_cad = df_fam[
            (df_fam['Caducidad'] == 'Unknown') &
            (df_fam['Uds.Vendidas'].fillna(0) > 0)
        ].copy()
        st.markdown(f"**{len(df_sin_cad)} productos con rotación > 0 pero sin fecha de caducidad**")
        if df_sin_cad.empty:
            st.success("✅ Todos los productos con rotación tienen fecha de caducidad.")
        else:
            st.warning(f"🟠 {len(df_sin_cad)} productos sin registrar caducidad.")
            cols_show = [c for c in COLUMNAS_INTERES if c in df_sin_cad.columns]
            st.dataframe(df_sin_cad[cols_show], use_container_width=True)
            st.download_button(
                "⬇️ Exportar sin caducidad registrada",
                data=exportar_excel_simple(df_sin_cad[cols_show]),
                file_name="stock_sin_caducidad.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # ── TAB 4: Exportar todo ─────────────────────────────────────────────
    with tab4:
        st.markdown("Genera un Excel con todas las hojas de stock en un solo archivo.")
        cols_base = [c for c in COLUMNAS_INTERES if c in df_fam.columns]

        hojas_stock = {}

        if 'S.Actual' in df_fam.columns and 'S.Minimo' in df_fam.columns:
            df_bajo_exp = df_fam[
                df_fam['S.Actual'].notna() & df_fam['S.Minimo'].notna() &
                (df_fam['S.Actual'] < df_fam['S.Minimo']) & (df_fam['Uds.Vendidas'].fillna(0) > 0)
            ].copy()
            df_bajo_exp['Diferencia'] = df_bajo_exp['S.Minimo'] - df_bajo_exp['S.Actual']
            hojas_stock['Stock bajo mínimo'] = df_bajo_exp[cols_base + ['Diferencia']]

        hojas_stock['Sin rotación'] = df_fam[
            (df_fam['S.Actual'].fillna(0) != 0) & (df_fam['Uds.Vendidas'].fillna(0) == 0)
        ][cols_base]

        hojas_stock['Sin caducidad'] = df_fam[
            (df_fam['Caducidad'] == 'Unknown') & (df_fam['Uds.Vendidas'].fillna(0) > 0)
        ][cols_base]

        st.download_button(
            "⬇️ Descargar informe completo de stock (multi-hoja)",
            data=exportar_excel_multihoja(hojas_stock),
            file_name="informe_stock_completo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


# ════════════════════════════════════════════════════════════════════════════
#  PÁGINA: DUPLICADOS
# ════════════════════════════════════════════════════════════════════════════

elif pagina == "🔍 Duplicados":
    st.markdown("# 🔍 Revisión de Duplicados")
    st.markdown("---")

    if 'Código' not in df_raw.columns:
        st.warning("El archivo no tiene columna 'Código' para detectar duplicados.")
        st.stop()

    df_dup = df_raw[df_raw.duplicated(subset='Código', keep=False)].copy()
    df_dup = df_dup.sort_values('Código')

    if df_dup.empty:
        st.success("✅ No se han encontrado duplicados en el inventario. ¡Todo limpio!")
    else:
        st.error(f"⚠️ Se han encontrado **{len(df_dup)} filas** con código duplicado ({df_dup['Código'].nunique()} códigos distintos).")
        cols_show = [c for c in COLUMNAS_INTERES if c in df_dup.columns]
        st.dataframe(df_dup[cols_show], use_container_width=True)
        st.download_button(
            "⬇️ Exportar duplicados a Excel",
            data=exportar_excel_simple(df_dup[cols_show]),
            file_name="duplicados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


# ════════════════════════════════════════════════════════════════════════════
#  PÁGINA: ANÁLISIS TOTAL
# ════════════════════════════════════════════════════════════════════════════

elif pagina == "📊 Análisis Total":
    st.markdown("# 📊 Análisis Global del Inventario")
    st.markdown("---")

    df_cad = preparar_caducidad(df_raw)
    cols_base = [c for c in COLUMNAS_INTERES if c in df_cad.columns]
    df_total = df_cad[cols_base].copy()

    tab1, tab2, tab3, tab4 = st.tabs(["🏆 Más rotación", "🏆 Más margen rotación", "❌ Sin caducidad (con stock)", "📥 Exportar"])

    with tab1:
        n_top = st.slider("Mostrar top N productos por rotación", 10, 200, 50)
        df_top = df_total[df_total['Uds.Vendidas'].notna()].sort_values('Uds.Vendidas', ascending=False).head(n_top)
        st.dataframe(df_top, use_container_width=True)

    with tab2:
        n_top_margen = st.slider("Mostrar top N productos por margen", 10, 200, 50)
        df_top_margen = df_total[df_total['Rotación'].notna()].sort_values('Rotación', ascending=False).head(n_top_margen)
        st.dataframe(df_top_margen, use_container_width=True)

    with tab3:
        df_sin_cad_total = df_total[
            (df_total['S.Actual'].fillna(0) != 0) &
            (df_total['Caducidad'] == 'Unknown')
        ]
        st.markdown(f"**{len(df_sin_cad_total)} productos con stock y sin fecha de caducidad registrada**")
        st.dataframe(df_sin_cad_total, use_container_width=True)

    with tab4:
        st.markdown("Exporta el inventario completo procesado con todas las columnas normalizadas.")

        df_top_exp = df_total[df_total['Uds.Vendidas'].notna()].sort_values('Uds.Vendidas', ascending=False)
        df_top_margen_exp = df_total[df_total['Rotación'].notna()].sort_values('Rotación', ascending=False)
        df_sin_rot_exp = df_total[(df_total['Uds.Vendidas'].fillna(0) == 0) & (df_total['S.Actual'].fillna(0) != 0)]

        hojas_total = {
            'Más rotación': df_top_exp,
            'Mas margen':df_top_margen_exp,
            'Sin caducidad': df_sin_cad_total,
            'Sin rotación con stock': df_sin_rot_exp,
            'Inventario completo': df_total,
        }
        st.download_button(
            "⬇️ Descargar análisis total (multi-hoja)",
            data=exportar_excel_multihoja(hojas_total),
            file_name="analisis_total_inventario.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
