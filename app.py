import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# ==========================================
# CONFIGURACIÓN DE PÁGINA Y TEMA ESTÉTICO
# ==========================================
st.set_page_config(
    page_title="Control Financiero Familiar",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilo CSS personalizado para tarjetas KPI y detalles visuales
st.markdown("""
<style>
    .kpi-card {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 15px 20px;
        border-left: 5px solid #1f77b4;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        margin-bottom: 10px;
    }
    .kpi-title {
        font-size: 0.85rem;
        color: #6c757d;
        font-weight: 600;
        text-transform: uppercase;
        margin-bottom: 5px;
    }
    .kpi-value {
        font-size: 1.6rem;
        color: #111827;
        font-weight: 700;
    }
    .kpi-subtext {
        font-size: 0.78rem;
        color: #4b5563;
        margin-top: 3px;
    }
    .stMetric {
        background-color: #f8f9fa;
        border-radius: 8px;
        padding: 10px;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# FUNCIONES DE CARGA Y GENERACIÓN DE DATOS
# ==========================================
def get_demo_data():
    """Genera datos de demostración si no se sube un archivo Excel."""
    # Tabla resumen por categoría
    data_cat = [
        {"Categoría": "Arriendo", "Gasto Real": 758000, "Proyectado": 758000, "Tipo": "Gastos Fijos"},
        {"Categoría": "Supermercado", "Gasto Real": 466965, "Proyectado": 600000, "Tipo": "Gastos Variables"},
        {"Categoría": "Bencina y Otros Auto", "Gasto Real": 507685, "Proyectado": 580000, "Tipo": "Gastos Variables"},
        {"Categoría": "Fintual", "Gasto Real": 397000, "Proyectado": 397000, "Tipo": "Ahorro"},
        {"Categoría": "Transf. Katy", "Gasto Real": 332091, "Proyectado": 332091, "Tipo": "Gastos Variables"},
        {"Categoría": "Gastos Comun", "Gasto Real": 206427, "Proyectado": 206427, "Tipo": "Gastos Fijos"},
        {"Categoría": "Ropa", "Gasto Real": 205195, "Proyectado": 205195, "Tipo": "Gastos Variables"},
        {"Categoría": "Luz, agua, Internet", "Gasto Real": 184929, "Proyectado": 184929, "Tipo": "Gastos Fijos"},
        {"Categoría": "Gastos varios", "Gasto Real": 131977, "Proyectado": 180000, "Tipo": "Gastos Variables"},
        {"Categoría": "Entretenimiento", "Gasto Real": 135825, "Proyectado": 135825, "Tipo": "Gastos Variables"},
        {"Categoría": "Restaurant", "Gasto Real": 144857, "Proyectado": 144857, "Tipo": "Gastos Variables"},
        {"Categoría": "Gas", "Gasto Real": 92800, "Proyectado": 92800, "Tipo": "Gastos Fijos"},
        {"Categoría": "Alimentación", "Gasto Real": 51154, "Proyectado": 83000, "Tipo": "Gastos Variables"},
        {"Categoría": "Farmacia, Salud y Belleza", "Gasto Real": 43101, "Proyectado": 80000, "Tipo": "Gastos Variables"},
        {"Categoría": "Regalos", "Gasto Real": 58032, "Proyectado": 58032, "Tipo": "Gastos Variables"},
        {"Categoría": "Gimnasio", "Gasto Real": 27589, "Proyectado": 55178, "Tipo": "Gastos Fijos"},
    ]
    df_cat = pd.DataFrame(data_cat)
    df_cat["Disponible"] = df_cat["Proyectado"] - df_cat["Gasto Real"]
    df_cat["% Usado"] = (df_cat["Gasto Real"] / df_cat["Proyectado"]).fillna(0)

    # Parametros generales
    params = {
        "mes_control": "2026.09 SEPTIEMBRE",
        "sueldo_fijo": 3100000,
        "meta_uber": 750000,
        "meta_ahorro_pct": 0.05,
        "ahorro_objetivo": 155000
    }

    # Transacciones
    fechas = pd.date_range(start="2026-09-01", periods=15, freq="D")
    df_trans = pd.DataFrame({
        "MES": ["2026.09 SEPTIEMBRE"] * 15,
        "FECHA": fechas,
        "DESCRIPCIÓN": [
            "Supermercado Lider", "Cps Copec Pve", "Pago Arriendo", "Transferencia Fintual",
            "Meli Dólar", "Farmacia Ahumada", "McDonalds", "Gimnasio Sportlife",
            "Jumbo Bilbao", "Gas Lipigas", "Enel Luz", "Enel Agua", "Restaurante Italia",
            "Zara Costanera", "Estacionamiento Costanera"
        ],
        "MONTO ($)": [-45000, -25000, -758000, -397000, 38, -12000, -8500, -27589, -120000, -45000, -35000, -18000, -42000, -65000, -5200],
        "ABONO/CARGO": ["Cargo", "Cargo", "Cargo", "Cargo", "Abono", "Cargo", "Cargo", "Cargo", "Cargo", "Cargo", "Cargo", "Cargo", "Cargo", "Cargo", "Cargo"],
        "CATEGORIA": ["Supermercado", "Bencina y Otros Auto", "Arriendo", "Fintual", "Beneficio Meli", "Farmacia, Salud y Belleza", "Alimentación", "Gimnasio", "Supermercado", "Gas", "Luz, agua, Internet", "Luz, agua, Internet", "Restaurant", "Ropa", "Estacionamiento"],
        "FIJO/VAR.": ["Gastos Variables", "Gastos Variables", "Gastos Fijos", "Ahorro", "Abono", "Gastos Variables", "Gastos Variables", "Gastos Fijos", "Gastos Variables", "Gastos Fijos", "Gastos Fijos", "Gastos Fijos", "Gastos Variables", "Gastos Variables", "Gastos Variables"]
    })

    return params, df_cat, df_trans

def parse_uploaded_excel(uploaded_file):
    """Procesa el archivo Excel cargado según las hojas específicas."""
    try:
        xls = pd.ExcelFile(uploaded_file)
        
        # 1. Cargar Hoja4 (Categorías y Resumen)
        if "Hoja4" in xls.sheet_names:
            df_raw4 = pd.read_excel(xls, "Hoja4", header=None)
            # Buscar el inicio de la tabla
            header_row_idx = None
            for idx, row in df_raw4.iterrows():
                if "Descripción" in row.values:
                    header_row_idx = idx
                    break
            
            if header_row_idx is not None:
                df_cat = df_raw4.iloc[header_row_idx+1:].copy()
                df_cat.columns = df_raw4.iloc[header_row_idx].values
                df_cat = df_cat.dropna(subset=["Descripción"])
                df_cat = df_cat[df_cat["Descripción"] != "Total general"]
                
                # Renombrar columnas clave
                rename_map = {
                    "Descripción": "Categoría",
                    "Gasto Real": "Gasto Real",
                    "Proyectado": "Proyectado",
                    "Dif": "Disponible",
                    "Gastos Fijos/Var.": "Tipo"
                }
                df_cat = df_cat.rename(columns=rename_map)
                
                # Normalizar tipos de datos numericos
                for col in ["Gasto Real", "Proyectado", "Disponible"]:
                    if col in df_cat.columns:
                        df_cat[col] = pd.to_numeric(df_cat[col], errors='coerce').fillna(0)
                        
                if "Proyectado" in df_cat.columns and "Gasto Real" in df_cat.columns:
                    df_cat["% Usado"] = (df_cat["Gasto Real"] / df_cat["Proyectado"].replace(0, 1)).fillna(0)
                else:
                    df_cat["% Usado"] = 0.0
            else:
                _, df_cat, _ = get_demo_data()
        else:
            _, df_cat, _ = get_demo_data()

        # 2. Cargar Parámetros del Dashboard
        params = {
            "mes_control": "2026.09 SEPTIEMBRE",
            "sueldo_fijo": 3100000,
            "meta_uber": 750000,
            "meta_ahorro_pct": 0.05,
            "ahorro_objetivo": 155000
        }
        if "Dashboard Familiar" in xls.sheet_names:
            df_dash = pd.read_excel(xls, "Dashboard Familiar")
            # Extraer variables si están disponibles
            try:
                for _, row in df_dash.iterrows():
                    val0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
                    if val0 == "Mes de control":
                        params["mes_control"] = str(row.iloc[1])
                    elif val0 == "Sueldo líquido fijo ($)":
                        params["sueldo_fijo"] = float(row.iloc[1])
                    elif val0 == "Meta Uber mensual ($)":
                        params["meta_uber"] = float(row.iloc[1])
                    elif val0 == "Ahorro objetivo ($)":
                        params["ahorro_objetivo"] = float(row.iloc[1])
            except Exception:
                pass

        # 3. Cargar Transacciones (Hoja5 o Historico)
        sheet_trans = "Hoja5" if "Hoja5" in xls.sheet_names else ("Historico" if "Historico" in xls.sheet_names else None)
        if sheet_trans:
            df_trans = pd.read_excel(xls, sheet_trans)
            # Limpiar nombres de columnas
            df_trans.columns = [str(c).strip() for c in df_trans.columns]
            
            # Ajustar columnas si vienen de Historico sin encabezado estándar
            if "DESCRIPCIÓN" not in df_trans.columns:
                # Buscar fila con encabezados
                for idx, row in df_trans.iterrows():
                    if "DESCRIPCIÓN" in row.values:
                        df_trans.columns = row.values
                        df_trans = df_trans.iloc[idx+1:].copy()
                        break
            
            if "FECHA" in df_trans.columns:
                df_trans["FECHA"] = pd.to_datetime(df_trans["FECHA"], errors='coerce')
            if "MONTO ($)" in df_trans.columns:
                df_trans["MONTO ($)"] = pd.to_numeric(df_trans["MONTO ($)"], errors='coerce').fillna(0)
        else:
            _, _, df_trans = get_demo_data()

        return params, df_cat, df_trans

    except Exception as e:
        st.error(f"Error procesando el archivo: {e}. Cargando datos de prueba...")
        return get_demo_data()

# ==========================================
# BARRA LATERAL (FILTROS Y CARGA)
# ==========================================
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/3135/3135706.png", width=70)
st.sidebar.title("Control Financiero")
st.sidebar.markdown("---")

uploaded_file = st.sidebar.file_uploader(
    "Cargar archivo Excel (.xlsx)",
    type=["xlsx"],
    help="Sube el archivo '202610_Presupuesto familiar 2026.xlsx'"
)

if uploaded_file is not None:
    params, df_cat, df_trans = parse_uploaded_excel(uploaded_file)
    st.sidebar.success("✅ Archivo cargado correctamente")
else:
    st.sidebar.info("💡 Usando datos de prueba. Sube tu archivo Excel para ver tus datos reales.")
    params, df_cat, df_trans = get_demo_data()

st.sidebar.markdown("### 🎯 Filtros Dinámicos")

# Selector de Mes
meses_disponibles = df_trans["MES"].dropna().unique().tolist() if "MES" in df_trans.columns else [params["mes_control"]]
if not meses_disponibles:
    meses_disponibles = [params["mes_control"]]
mes_seleccionado = st.sidebar.selectbox("Mes de Control", meses_disponibles, index=0)

# Filtro por Tipo de Gasto
tipos_disponibles = df_cat["Tipo"].dropna().unique().tolist() if "Tipo" in df_cat.columns else ["Gastos Fijos", "Gastos Variables", "Ahorro"]
filtro_tipo = st.sidebar.multiselect("Tipo de Gastos/Ahorro", tipos_disponibles, default=tipos_disponibles)

# Buscador de texto
busqueda_texto = st.sidebar.text_input("🔍 Buscar transacción", placeholder="Ej: Supermercado, Copec...")

# Aplicar filtros a DataFrames
df_cat_filtrado = df_cat[df_cat["Tipo"].isin(filtro_tipo)] if "Tipo" in df_cat.columns and filtro_tipo else df_cat

df_trans_filtrado = df_trans.copy()
if "MES" in df_trans_filtrado.columns and mes_seleccionado:
    df_trans_filtrado = df_trans_filtrado[df_trans_filtrado["MES"] == mes_seleccionado]
if "FIJO/VAR." in df_trans_filtrado.columns and filtro_tipo:
    df_trans_filtrado = df_trans_filtrado[df_trans_filtrado["FIJO/VAR."].isin(filtro_tipo)]
if busqueda_texto and "DESCRIPCIÓN" in df_trans_filtrado.columns:
    df_trans_filtrado = df_trans_filtrado[df_trans_filtrado["DESCRIPCIÓN"].astype(str).str.contains(busqueda_texto, case=False, na=False)]

# ==========================================
# CABECERA Y KPI CARDS
# ==========================================
st.title("📊 Dashboard de Presupuesto y Finanzas Familiares")
st.markdown(f"**Periodo Activo:** `{mes_seleccionado}` | Monitoreo en tiempo real de presupuesto vs. ejecuciones reales.")
st.markdown("---")

# Cálculos de Indicadores Clave
ingreso_esperado = params["sueldo_fijo"] + params["meta_uber"]
gasto_real_total = df_cat["Gasto Real"].sum() if "Gasto Real" in df_cat.columns else 0
presupuesto_proyectado = df_cat["Proyectado"].sum() if "Proyectado" in df_cat.columns else 0
presupuesto_disponible = presupuesto_proyectado - gasto_real_total

# Ahorro
df_ahorro = df_cat[df_cat["Tipo"] == "Ahorro"] if "Tipo" in df_cat.columns else pd.DataFrame()
ahorro_real = df_ahorro["Gasto Real"].sum() if not df_ahorro.empty else 0
ahorro_objetivo = params["ahorro_objetivo"]
cumplimiento_ahorro = (ahorro_real / ahorro_objetivo * 100) if ahorro_objetivo > 0 else 0

# Fila de Métricas
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric(
        label="Ingreso Esperado Total",
        value=f"${ingreso_esperado:,.0f}".replace(",", "."),
        delta=f"Fijo: ${params['sueldo_fijo']:,.0f}".replace(",", "."),
        delta_color="normal"
    )

with col2:
    pct_ejecucion = (gasto_real_total / presupuesto_proyectado * 100) if presupuesto_proyectado > 0 else 0
    st.metric(
        label="Gasto Real Acumulado",
        value=f"${gasto_real_total:,.0f}".replace(",", "."),
        delta=f"{pct_ejecucion:.1f}% del presupuesto",
        delta_color="inverse"
    )

with col3:
    st.metric(
        label="Presupuesto Restante",
        value=f"${presupuesto_disponible:,.0f}".replace(",", "."),
        delta="Saldo Disponible" if presupuesto_disponible >= 0 else "Sobregiro",
        delta_color="normal" if presupuesto_disponible >= 0 else "inverse"
    )

with col4:
    brecha = ahorro_real - ahorro_objetivo
    st.metric(
        label="Ahorro Proyectado vs Objetivo",
        value=f"${ahorro_real:,.0f}".replace(",", "."),
        delta=f"{cumplimiento_ahorro:.1f}% de la meta (${ahorro_objetivo:,.0f})".replace(",", "."),
        delta_color="normal" if brecha >= 0 else "inverse"
    )

st.markdown("---")

# ==========================================
# VISUALIZACIONES PRINCIPALES
# ==========================================
tab1, tab2, tab3 = st.tabs(["📉 Análisis por Categoría", "🍩 Distribución y Tendencia", "📄 Registro de Transacciones"])

with tab1:
    col_chart1, col_chart2 = st.columns([1.2, 1])
    
    with col_chart1:
        st.subheader("Comparativo: Presupuestado vs. Real por Categoría")
        if not df_cat_filtrado.empty:
            df_chart = df_cat_filtrado.sort_values(by="Proyectado", ascending=True)
            
            fig_comp = go.Figure()
            fig_comp.add_trace(go.Bar(
                y=df_chart["Categoría"],
                x=df_chart["Proyectado"],
                name="Presupuesto Proyectado",
                orientation='h',
                marker_color='#93c5fd'
            ))
            fig_comp.add_trace(go.Bar(
                y=df_chart["Categoría"],
                x=df_chart["Gasto Real"],
                name="Gasto Real",
                orientation='h',
                marker_color='#1d4ed8'
            ))
            
            fig_comp.update_layout(
                barmode='group',
                height=520,
                margin=dict(l=20, r=20, t=30, b=20),
                xaxis=dict(title="Monto ($)"),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            st.plotly_chart(fig_comp, use_container_width=True)
        else:
            st.warning("No hay datos disponibles para los filtros seleccionados.")

    with col_chart2:
        st.subheader("Semáforo de Ejecución del Presupuesto")
        if not df_cat_filtrado.empty:
            df_semaforo = df_cat_filtrado.copy()
            df_semaforo["% Usado Pct"] = df_semaforo["% Usado"] * 100
            df_semaforo = df_semaforo.sort_values(by="% Usado Pct", ascending=False)
            
            # Asignar colores según nivel de ejecución
            def get_color(val):
                if val > 100:
                    return "#ef4444"  # Rojo
                elif val >= 85:
                    return "#f59e0b"  # Amarillo
                else:
                    return "#10b981"  # Verde

            colors = [get_color(v) for v in df_semaforo["% Usado Pct"]]

            fig_sem = go.Figure(go.Bar(
                x=df_semaforo["% Usado Pct"],
                y=df_semaforo["Categoría"],
                orientation='h',
                marker_color=colors,
                text=df_semaforo["% Usado Pct"].apply(lambda x: f"{x:.1f}%"),
                textposition='outside'
            ))
            fig_sem.add_vline(x=100, line_dash="dash", line_color="red", annotation_text="Límite 100%")
            fig_sem.add_vline(x=85, line_dash="dot", line_color="orange", annotation_text="Alerta 85%")
            
            fig_sem.update_layout(
                height=520,
                margin=dict(l=20, r=20, t=30, b=20),
                xaxis=dict(title="% de Presupuesto Consumido", range=[0, max(df_semaforo["% Usado Pct"].max() * 1.15, 110)])
            )
            st.plotly_chart(fig_sem, use_container_width=True)

with tab2:
    col_pie, col_burn = st.columns([1, 1.2])
    
    with col_pie:
        st.subheader("Distribución de Gastos por Tipo")
        if "Tipo" in df_cat.columns and not df_cat.empty:
            df_pie = df_cat.groupby("Tipo")["Gasto Real"].sum().reset_index()
            fig_pie = px.pie(
                df_pie,
                names="Tipo",
                values="Gasto Real",
                hole=0.45,
                color_discrete_sequence=px.colors.qualitative.Set2
            )
            fig_pie.update_traces(textinfo="percent+label", hovertemplate="%{label}: $% {value:,.0f}")
            fig_pie.update_layout(height=450, margin=dict(l=20, r=20, t=30, b=20))
            st.plotly_chart(fig_pie, use_container_width=True)

    with col_burn:
        st.subheader("Ritmo de Gasto Acumulado (Burn Rate)")
        if not df_trans_filtrado.empty and "FECHA" in df_trans_filtrado.columns and "MONTO ($)" in df_trans_filtrado.columns:
            # Filtrar solo cargos (gastos positivos/negativos según formato)
            df_cargos = df_trans_filtrado[df_trans_filtrado["MONTO ($)"] < 0].copy()
            df_cargos["Gasto"] = df_cargos["MONTO ($)"].abs()
            df_daily = df_cargos.groupby("FECHA")["Gasto"].sum().reset_index().sort_values("FECHA")
            df_daily["Acumulado"] = df_daily["Gasto"].cumsum()

            fig_burn = px.line(
                df_daily,
                x="FECHA",
                y="Acumulado",
                markers=True,
                line_shape="linear",
                title="Evolución Diaria de Gasto Acumulado"
            )
            fig_burn.update_traces(line_color="#2563eb", line_width=3)
            fig_burn.update_layout(
                height=450,
                margin=dict(l=20, r=20, t=40, b=20),
                xaxis_title="Fecha",
                yaxis_title="Gasto Acumulado ($)"
            )
            st.plotly_chart(fig_burn, use_container_width=True)
        else:
            st.info("No hay suficientes transacciones detalladas con fecha para trazar la línea de tendencia.")

with tab3:
    st.subheader("Detalle Transaccional")
    st.markdown("Visualiza y busca dentro del registro histórico de movimientos.")

    if not df_trans_filtrado.empty:
        # Formatear DataFrame para visualización amigable
        df_display = df_trans_filtrado.copy()
        
        if "FECHA" in df_display.columns:
            df_display["FECHA"] = df_display["FECHA"].dt.strftime("%Y-%m-%d")

        # Aplicar formato visual interactivo
        st.dataframe(
            df_display,
            use_container_width=True,
            hide_index=True,
            column_config={
                "MONTO ($)": st.column_config.NumberColumn(
                    "Monto ($)",
                    format="$ %d",
                    help="Monto en CLP (Abonos positivos/negativos segun cartola)"
                ),
                "FECHA": st.column_config.TextColumn("Fecha"),
                "DESCRIPCIÓN": st.column_config.TextColumn("Descripción"),
                "CATEGORIA": st.column_config.TextColumn("Categoría"),
                "ABONO/CARGO": st.column_config.TextColumn("Tipo Movimiento")
            }
        )
        st.caption(f"Mostrando {len(df_display)} transacciones registradas.")
    else:
        st.warning("No se encontraron transacciones que coincidan con los filtros seleccionados.")

# ==========================================
# PIE DE PÁGINA Y NOTAS
# ==========================================
st.markdown("---")
st.caption("🤖 App de Control Financiero Familiar | Creado con Streamlit & Plotly")