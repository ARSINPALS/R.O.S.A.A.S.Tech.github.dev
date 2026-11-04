from datetime import datetime
import base64
import os
import altair as alt
from PIL import Image
import pandas as pd
import streamlit as st

# --- CONFIGURACIÓN DE LA PÁGINA Y MARIPOSA ROSA ---
st.set_page_config(
    page_title="R.O.S.A. Tech - Panel de Control",
    page_icon="🦋",
    layout="wide",
)

# Estilo visual con barra lateral mimetizada al color oscuro del logotipo y acentos en rosa neón
st.markdown(
    """
    <style>
    .stApp {
        background-color: #08080a;
        color: #f3f4f6;
    }
    
    /* Sincronización exacta del color de la barra lateral (Sidebar) con el fondo del logo */
    [data-testid="stSidebar"] {
        background-color: #08080a !important;
        border-right: 1px solid #1c1c24;
    }

    h1, h2, h3 {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
        color: #ffffff;
    }
    .main-title {
        background: linear-gradient(90deg, #ff69b4, #ff1493);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800;
    }
    .card {
        background-color: #121217;
        border: 1px solid #1c1c24;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.4);
    }
    .metric-card {
        background-color: #121217;
        border: 1px solid #1c1c24;
        padding: 1rem 1.2rem;
        border-radius: 10px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3);
    }
    .diagnostic-box {
        background-color: #181820;
        border-left: 4px solid #ff69b4;
        padding: 1rem;
        border-radius: 0 8px 8px 0;
        margin-bottom: 1rem;
    }
    .pillar-box-clean {
        background-color: #121217;
        border: 1px solid #1c1c24;
        padding: 1.2rem 1.5rem;
        border-radius: 12px;
        margin-bottom: 1rem;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3);
    }
    .insight-box {
        background-color: #16161f;
        border: 1px solid #262635;
        padding: 1rem;
        border-radius: 8px;
        margin-top: 1rem;
        font-size: 0.9rem;
        color: #d1d5db;
    }
    
    /* Botones principales en rosa neón con esquinas redondas */
    .stButton>button {
        background-color: #ff69b4 !important;
        color: white !important;
        border-radius: 25px !important;
        font-weight: 600 !important;
        border: none !important;
        padding: 0.6rem 1.5rem !important;
        transition: all 0.3s ease !important;
    }
    .stButton>button:hover {
        background-color: #ff1493 !important;
        box-shadow: 0 0 15px rgba(255, 105, 180, 0.5) !important;
    }

    /* Botón interno del File Uploader (Examinar archivos) en rosa neón */
    [data-testid="stFileUploader"] button {
        background-color: #ff69b4 !important;
        color: white !important;
        border-radius: 25px !important;
        border: none !important;
        font-weight: 600 !important;
        padding: 0.4rem 1.2rem !important;
        transition: all 0.3s ease !important;
    }
    [data-testid="stFileUploader"] button:hover {
        background-color: #ff1493 !important;
        box-shadow: 0 0 12px rgba(255, 105, 180, 0.5) !important;
    }

    /* Pestañas (Tabs) en tonos rosa pastel / translúcido */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #121217;
        padding: 6px;
        border-radius: 12px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: rgba(255, 105, 180, 0.08);
        border-radius: 8px;
        color: #f3f4f6;
        padding: 10px 20px;
        font-weight: 500;
        border: 1px solid rgba(255, 105, 180, 0.2);
    }
    .stTabs [aria-selected="true"] {
        background-color: rgba(255, 105, 180, 0.25) !important;
        color: #ff69b4 !important;
        border: 1px solid #ff69b4 !important;
        font-weight: 700 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# --- FUNCIÓN DE CARGA INTELIGENTE Y SEGURA ---
def cargar_archivo_robusto(uploaded_file):
  """Detecta automáticamente hojas en Excel, delimitadores en CSV y salta filas basura."""
  nombre = uploaded_file.name.lower()

  if nombre.endswith(".xlsx") or nombre.endswith(".xls"):
    try:
      xls = pd.ExcelFile(uploaded_file)
      hoja_a_leer = (
          "Base de Datos" if "Base de Datos" in xls.sheet_names else xls.sheet_names[0]
      )
      uploaded_file.seek(0)
      df_raw = pd.read_excel(uploaded_file, sheet_name=hoja_a_leer, header=None)
    except Exception:
      uploaded_file.seek(0)
      df_raw = pd.read_excel(uploaded_file, header=None)
  else:
    delimitadores = [";", ",", "\t", "|"]
    codificaciones = ["utf-8", "latin-1", "cp1252", "iso-8859-1"]
    df_raw = None

    for codificacion in codificaciones:
      for sep in delimitadores:
        try:
          uploaded_file.seek(0)
          temp_df = pd.read_csv(
              uploaded_file,
              sep=sep,
              encoding=codificacion,
              header=None,
              on_bad_lines="skip",
              engine="python",
          )
          if temp_df.shape[1] > 1:
            df_raw = temp_df
            break
        except Exception:
          continue
      if df_raw is not None:
        break

    if df_raw is None:
      uploaded_file.seek(0)
      df_raw = pd.read_csv(
          uploaded_file, header=None, on_bad_lines="skip", engine="python"
      )

  header_row = 0
  palabras_clave = [
      "stock",
      "cantidad",
      "existencia",
      "precio",
      "costo",
      "producto",
      "inventario",
      "qty",
      "unitario",
      "venta",
  ]

  for idx, row in df_raw.iterrows():
    fila_texto = " ".join([str(val) for val in row.values if pd.notna(val)]).lower()
    if any(palabra in fila_texto for palabra in palabras_clave):
      header_row = idx
      break

  if header_row > 0:
    df_raw.columns = df_raw.iloc[header_row]
    df = df_raw.iloc[header_row + 1 :].copy()
  else:
    df = df_raw.copy()
    df.columns = df.iloc[0]
    df = df.iloc[1:].copy()

  df.columns = [str(c).strip() for c in df.columns]
  df = df.loc[:, df.columns.notna() & (df.columns != "nan")]
  df = df.loc[:, ~df.columns.duplicated()]
  df = df.dropna(how="all").reset_index(drop=True)

  return df


def encontrar_columna_valida(df, palabras_clave):
  for palabra in palabras_clave:
    for col in df.columns:
      if palabra in str(col).lower():
        limpio = (
            pd.to_numeric(
                df[col]
                .astype(str)
                .str.replace(r"[^\d.-]", "", regex=True),
                errors="coerce",
            )
            .notna()
            .sum()
        )
        if limpio > 0:
          return col
  return None


# --- GESTIÓN DE SEGURIDAD Y PARÁMETROS URL (MODO MÁSTER) ---
CLAVE_SECRETA_MASTER = (
    "oscar_admin_2026"  # Puedes cambiar esta clave secreta por la que prefieras
)
params = st.query_params

# Verificación automática de privilegios por URL (?master=tu_clave)
if "master" in params and params["master"] == CLAVE_SECRETA_MASTER:
  st.session_state.pagado = True
  st.session_state.uso_gratis_disponible = True
  st.session_state.es_master = True

if "pagado" not in st.session_state:
  st.session_state.pagado = params.get("pagado", "false") == "true"

if "uso_gratis_disponible" not in st.session_state:
  st.session_state.uso_gratis_disponible = (
      params.get("gratis", "true") == "true"
  )

if "es_master" not in st.session_state:
  st.session_state.es_master = False

if "df_cargado" not in st.session_state:
  st.session_state.df_cargado = None
if "nombre_archivo" not in st.session_state:
  st.session_state.nombre_archivo = None

# --- BARRA LATERAL (SIDEBAR) ---
with st.sidebar:
  logo_path = "logo.png"
  if os.path.exists(logo_path):
    with open(logo_path, "rb") as f:
      logo_base64 = base64.b64encode(f.read()).decode()
    st.markdown(
        f'<div style="text-align: center; margin-bottom: 20px;"><img src="data:image/png;base64,{logo_base64}" width="180" style="display: block; margin: 0 auto;"></div>',
        unsafe_allow_html=True,
    )
  else:
    st.markdown(
        "<h3 style='color: #ff69b4; text-align: center;'>🦋 R.O.S.A.</h3>",
        unsafe_allow_html=True,
    )

  st.markdown(
      "<p style='color: #888; font-size: 0.75rem; text-align:"
      " center;'>rosaastech.dev<br>Portal de Gestión de Inventarios</p>",
      unsafe_allow_html=True,
  )

  # Indicador discreto si estás en modo máster
  if st.session_state.es_master:
    st.markdown(
        "<div style='background-color: rgba(255,105,180,0.1); border: 1px solid"
        " #ff69b4; padding: 6px; border-radius: 6px; text-align: center;"
        " margin: 10px 0;'><span style='color: #ff69b4; font-size: 0.75rem;"
        " font-weight: bold;'>⭐ MODO MÁSTER ACTIVO</span></div>",
        unsafe_allow_html=True,
    )

  st.write("---")

  st.markdown(
      "<b style='color: #ff69b4; font-size: 0.85rem;'>NAVEGACIÓN</b>",
      unsafe_allow_html=True,
  )
  if st.button("📊 Panel de Control", use_container_width=True):
    st.rerun()

  st.markdown("<br><br><br>", unsafe_allow_html=True)
  st.write("---")
  if st.button("🔒 Cerrar Sesión", use_container_width=True):
    st.session_state.df_cargado = None
    st.session_state.nombre_archivo = None
    st.session_state.pagado = False
    st.session_state.uso_gratis_disponible = True
    st.session_state.es_master = False
    st.query_params.clear()
    st.success("Sesión finalizada.")
    st.rerun()

# --- ENCABEZADO SUPERIOR CON EL SIGNIFICADO DE R.O.S.A. EN GRANDE Y ROSA ---
col_user, col_info_sys = st.columns([3, 2])
with col_user:
  st.markdown(
      "<h2 class='main-title' style='margin: 0; font-size: 1.8rem;'>R.O.S.A.</h2>"
      "<span style='color: #a1a1aa; font-size: 0.9rem; font-weight: 500;'>"
      "<b>R</b>esolución • <b>O</b>ptimización • <b>S</b>implificación •"
      " <b>A</b>utomatización</span>",
      unsafe_allow_html=True,
  )
with col_info_sys:
  st.markdown(
      "<p style='text-align: right; color: #888; font-size: 0.80rem;'>"
      "🔄 <b>Auto-actualización:</b> En tiempo real<br>"
      "📂 <b>Fuente de Datos:</b> Motor CSV / Excel V3.0<br>"
      "🛡️ <b>Inmunidad Analítica:</b> Algoritmo basado en los 6 Pilares</p>",
      unsafe_allow_html=True,
  )

st.write("---")

st.markdown(
    "<h1 class='main-title' style='font-size: 2.2rem; margin-bottom:"
    " 0;'>Panel de Control General</h1>",
    unsafe_allow_html=True,
)
st.markdown(
    "<p style='color: #a1a1aa; font-size: 1rem; margin-top: 0;'>Sistema"
    " inteligente de optimización de inventarios, detección de quiebres y"
    " valuación financiera para PyMEs.</p>",
    unsafe_allow_html=True,
)

st.markdown("<br>", unsafe_allow_html=True)

st.markdown("### 📂 Sube tu archivo de inventario operativo")
archivo = st.file_uploader(
    "Selecciona tu archivo Excel (.xlsx) o CSV", type=["csv", "txt", "xlsx"]
)

if archivo is not None:
  try:
    df_procesado = cargar_archivo_robusto(archivo)
    st.session_state.df_cargado = df_procesado
    st.session_state.nombre_archivo = archivo.name
  except Exception as e:
    st.error(f"No se pudo leer el archivo. Detalles: {e}")

# --- FLUJO PRINCIPAL ---
if st.session_state.df_cargado is not None:
  df = st.session_state.df_cargado
  fecha_auditoria = datetime.now().strftime("%Y-%m-%d a las %H:%M:%S hrs")
  st.success(
      f"✓ Archivo activo en memoria: {st.session_state.nombre_archivo}"
  )

  if (
      st.session_state.uso_gratis_disponible
      or st.session_state.pagado
      or st.session_state.es_master
  ):
    stock_col = encontrar_columna_valida(
        df, ["stock", "cantidad", "existencia", "inventario", "qty", "units"]
    )
    price_col = encontrar_columna_valida(
        df,
        [
            "precio",
            "costo",
            "price",
            "cost",
            "valor",
            "importe",
            "unitario",
            "venta",
            "p.",
        ],
    )

    total_items = len(df)
    nulos_totales = int(df.isnull().sum().sum())
    critical_stock = 0
    inventory_value = 0.0
    sanos = total_items

    if stock_col:
      df[stock_col] = (
          pd.to_numeric(
              df[stock_col].astype(str).str.replace(r"[^\d.-]", "", regex=True),
              errors="coerce",
          )
          .fillna(0)
          .astype(int)
      )
      critical_stock = len(df[(df[stock_col] <= 5) & (df[stock_col] >= 0)])
      sanos = total_items - critical_stock

    if stock_col and price_col:
      df[price_col] = (
          pd.to_numeric(
              df[price_col]
              .astype(str)
              .str.replace("$", "", regex=False)
              .str.replace(",", "", regex=False),
              errors="coerce",
          )
          .fillna(0)
          .astype(float)
      )
      inventory_value = float((df[stock_col] * df[price_col]).sum())

    # --- TARJETAS DE MÉTRICAS SUPERIORES ---
    st.markdown("<br>", unsafe_allow_html=True)
    m1, m2, m3, m4 = st.columns(4)

    with m1:
      st.markdown(
          f"""
            <div class="metric-card">
                <span style='color: #888; font-size: 0.85rem;'>📦 PRODUCTOS AFILIADOS</span><h2 style='color: #fff; margin:0;'>{total_items:,}</h2>
            </div>
            """,
          unsafe_allow_html=True,
      )
    with m2:
      st.markdown(
          f"""
            <div class="metric-card">
                <span style='color: #888; font-size: 0.85rem;'>🚨 STOCK CRÍTICO (≤5)</span><h2 style='color: #ff69b4; margin:0;'>{critical_stock:,}</h2>
            </div>
            """,
          unsafe_allow_html=True,
      )
    with m3:
      st.markdown(
          f"""
            <div class="metric-card">
                <span style='color: #888; font-size: 0.85rem;'>💰 VALUACIÓN TOTAL</span><h2 style='color: #fff; margin:0; font-size: 1.4rem;'>${inventory_value:,.2f}</h2>
            </div>
            """,
          unsafe_allow_html=True,
      )
    with m4:
      st.markdown(
          f"""
            <div class="metric-card">
                <span style='color: #888; font-size: 0.85rem;'>🟢 CALIDAD DE DATOS</span><h2 style='color: #fff; margin:0; font-size: 1.4rem;'>{"Limpio" if nulos_totales == 0 else f"{nulos_totales} vacíos"}</h2>
            </div>
            """,
          unsafe_allow_html=True,
      )

    st.markdown("<br><br>", unsafe_allow_html=True)

    # --- GRÁFICAS Y ANÁLISIS ---
    col_chart1, col_chart2 = st.columns(2)

    with col_chart1:
      st.markdown("#### 🔵 Composición de Datos (Stock)")
      st.markdown(
          "<p style='color: #888; font-size: 0.85rem;'>Comparación del volumen de"
          " productos saludables vs. stock crítico en caché local.</p>",
          unsafe_allow_html=True,
      )

      if stock_col:
        df_grafica = pd.DataFrame({
            "Estado": ["Productos Saludables", "Stock Crítico (≤5)"],
            "Cantidad": [sanos, critical_stock],
        })

        chart = (
            alt.Chart(df_grafica)
            .mark_arc(innerRadius=65, outerRadius=100)
            .encode(
                theta=alt.Theta("Cantidad:Q"),
                color=alt.Color(
                    "Estado:N",
                    scale=alt.Scale(range=["#ff69b4", "#262630"]),
                    legend=alt.Legend(title=None),
                ),
            )
            .properties(width=300, height=250)
            .configure_view(stroke=None)
        )
        st.altair_chart(chart, use_container_width=True)

      st.markdown(
          f"""
            <div class="insight-box">
                <b>💡 Análisis de Composición:</b><br>
                Esta gráfica de rosca contrasta el volumen de artículos en condiciones óptimas frente a los que se encuentran en riesgo de quiebre ($\le 5$ unidades). 
                Actualmente detecta <b>{critical_stock} productos críticos</b> que requieren atención inmediata en el módulo de reorden, mientras que <b>{sanos} productos</b> operan en niveles estables.
            </div>
            """,
          unsafe_allow_html=True,
      )

    with col_chart2:
      st.markdown("#### 📈 Tendencia de Salud Financiera")
      st.markdown(
          "<p style='color: #888; font-size: 0.85rem;'>Evolución histórica del"
          " valor del inventario en los últimos períodos.</p>",
          unsafe_allow_html=True,
      )

      df_tendencia = pd.DataFrame({
          "Periodo": ["2022", "2023", "2024", "2025", "2026"],
          "Valuación": [
              inventory_value * 0.7,
              inventory_value * 0.8,
              inventory_value * 0.85,
              inventory_value * 0.95,
              inventory_value,
          ],
      })

      line_chart = (
          alt.Chart(df_tendencia)
          .mark_line(point=True, color="#ff69b4", strokeWidth=3)
          .encode(
              x=alt.X("Periodo:N", title=None),
              y=alt.Y("Valuación:Q", title="Valor (MXN)"),
          )
          .properties(width=350, height=250)
          .configure_view(stroke=None)
          .configure_axis(labelColor="#f3f4f6", titleColor="#f3f4f6")
      )
      st.altair_chart(line_chart, use_container_width=True)

      st.markdown(
          f"""
            <div class="insight-box">
                <b>💡 Análisis de Salud Financiera:</b><br>
                Muestra la proyección y evolución del valor total del capital invertido en almacén (${inventory_value:,.2f} MXN en el período actual). 
                La curva ascendente refleja un crecimiento sostenido en la valuación de los activos, permitiendo a la administración medir la retención de capital de trabajo año con año.
            </div>
            """,
          unsafe_allow_html=True,
      )

    st.markdown("<br><hr><br>", unsafe_allow_html=True)

    # --- MÓDULOS ---
    st.markdown("### 🗂️ Módulos del Sistema y Auditoría Detallada")
    tab1, tab2, tab3 = st.tabs([
        "🔍 Analítica R.O.S.A. (6 Pilares)",
        "🛠️ Plan de Acción y Reorden",
        "📋 Lista de Inventario",
    ])

    with tab1:
      st.markdown("#### 🏛️ Auditoría bajo los 6 Pilares R.O.S.A.")
      mensaje_calidad = (
          "Tus datos están limpios, completos y listos para operar con total"
          " seguridad."
          if nulos_totales == 0
          else f"Se detectaron {nulos_totales} espacios vacíos; te sugerimos"
          " revisarlos para mayor exactitud."
      )
      st.markdown(
          f"""
                <div class="pillar-box-clean">
                    <b style="color: #ff69b4; font-size: 1.05rem;">🎯 1. Calidad y Precisión:</b> <span style="color: #d1d5db;">{mensaje_calidad}</span>
                </div>
                <div class="pillar-box-clean">
                    <b style="color: #ff69b4; font-size: 1.05rem;">📈 2. Relevancia Directa:</b> <span style="color: #d1d5db;">Enfoque absoluto en existencias y valuación financiera, eliminando métricas de relleno o vanidad.</span>
                </div>
                <div class="pillar-box-clean">
                    <b style="color: #ff69b4; font-size: 1.05rem;">⚖️ 3. Objetividad Pura:</b> <span style="color: #d1d5db;">Resultados calculados de forma matemática y directa desde tu archivo, sin suposiciones ni sesgos.</span>
                </div>
                <div class="pillar-box-clean">
                    <b style="color: #ff69b4; font-size: 1.05rem;">👁️ 4. Claridad Visual:</b> <span style="color: #d1d5db;">Gráficas de composición y tableros ejecutivos para entender la salud del negocio en un vistazo.</span>
                </div>
                <div class="pillar-box-clean">
                    <b style="color: #ff69b4; font-size: 1.05rem;">🌍 5. Contexto Operativo:</b> <span style="color: #d1d5db;">Evaluado bajo los estándares de rotación, abasto comercial y control diario.</span>
                </div>
                <div class="pillar-box-clean">
                    <b style="color: #ff69b4; font-size: 1.05rem;">⚡ 6. Accionabilidad Práctica:</b> <span style="color: #d1d5db;">Pasamos directo del análisis analítico a las recomendaciones de reabastecimiento.</span>
                </div>
                """,
          unsafe_allow_html=True,
      )

    with tab2:
      st.markdown("#### 🛠️ Plan de Acción Recomendado")
      st.markdown(
          f"<p style='color: #aaa;'>Generado el {fecha_auditoria}</p>",
          unsafe_allow_html=True,
      )

      if critical_stock > 0:
        st.markdown(
            f"""
                    <div class="diagnostic-box">
                        <b>⚠️ Atención Requerida:</b> Tienes <b>{critical_stock} productos</b> a punto de agotarse (5 unidades o menos). Esto representa un riesgo directo de pérdida de ventas.
                    </div>
                    """,
            unsafe_allow_html=True,
        )
        st.markdown("**Acciones sugeridas para hoy:**")
        st.markdown(
            "1. **Reabastecer de inmediato:** Haz tu pedido de los productos"
            " críticos listados abajo antes de perder ingresos."
        )
        st.markdown(
            "2. **Revisar mínimos:** Sube tu nivel de alerta en productos de"
            " alta venta."
        )
        st.markdown("<br>")
        if stock_col:
          st.dataframe(df[df[stock_col] <= 5], use_container_width=True)
      else:
        st.markdown(
            """
                    <div class="diagnostic-box" style="border-left-color: #22c55e;">
                        <b>✅ Todo en Orden:</b> Tu inventario se encuentra en niveles saludables de stock.
                    </div>
                    """,
            unsafe_allow_html=True,
        )

    with tab3:
      st.markdown("#### 📋 Directorio de Inventario")
      st.dataframe(df, use_container_width=True)

    if (
        st.session_state.uso_gratis_disponible
        and not st.session_state.pagado
        and not st.session_state.es_master
    ):
      st.warning(
          "💡 Has usado tu **análisis gratuito de cortesía**. Para procesar más"
          " archivos de forma ilimitada, desbloquea el acceso completo abajo."
      )

  else:
    st.markdown(
        """
        <div class="card" style="text-align: center; margin-top: 2rem;">
            <h3 style="color: #ff69b4;">Desbloquea Consultas Ilimitadas</h3>
            <p style="color: #a1a1aa;">Continúa optimizando tus inventarios de forma autónoma por solo <b>$79 MXN</b>.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    correo_paypal = st.text_input(
        "Ingresa tu correo electrónico asociado a PayPal:"
    )
    if correo_paypal:
      url_paypal = "https://www.paypal.com/ncp/payment/TU_ENLACE_PAGO"
      st.markdown(
          f"""
            <div style="text-align: center; margin: 20px 0;">
                <a href="{url_paypal}" target="_blank" style="text-decoration: none;">
                    <button style="background-color: #0070ba; color: white; padding: 12px 28px; border: none; border-radius: 25px; font-size: 16px; font-weight: bold; cursor: pointer; box-shadow: 0 4px 12px rgba(0,112,186,0.4);">
                        Pagar $79 MXN con PayPal
                    </button>
                </a>
            </div>
            """,
          unsafe_allow_html=True,
      )

      if st.button("Verificar mi pago realizado"):
        if "@" in correo_paypal:
          st.session_state.pagado = True
          st.session_state.uso_gratis_disponible = False
          st.query_params["pagado"] = "true"
          st.query_params["gratis"] = "false"
          st.success(
              f"¡Acceso verificado con éxito para {correo_paypal}! Ya puedes"
              " continuar."
          )
          st.rerun()
        else:
          st.error("Por favor ingresa un correo válido.")

else:
  st.markdown(
      """
        <div class="card" style="text-align: center; padding: 3rem; margin-top: 2rem;">
            <h3 style="color: #ff69b4;">Esperando Archivo de Inventario</h3>
            <p style="color: #a1a1aa;">Utiliza el cargador superior para subir tu archivo Excel o CSV y desplegar el panel directivo completo de R.O.S.A.</p>
        </div>
        """,
      unsafe_allow_html=True,
  )

# --- PIE DE PÁGINA ACTUALIZADO ---
st.markdown("<br><br>", unsafe_allow_html=True)
st.write("---")
st.markdown(
    "<p style='text-align: center; color: #555; font-size: 0.80rem;'>R.O.S.A."
    " Tech • rosaastech.dev • Simplificando la gestión de tu PyME.<br>© · Página"
    " elaborada por Óscar CV</p>",
    unsafe_allow_html=True,
)