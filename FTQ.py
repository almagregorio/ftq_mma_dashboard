import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import os
from datetime import datetime, timedelta

# --- CONFIGURACIÓN KOSTAL ---
st.set_page_config(page_title="KOSTAL - MERCEDES BENZ", layout="wide")

COLOR_KOSTAL = "#004b87"
COLOR_TARGET = "#28a745"
COLOR_ACTION = "#dc3545"

st.title("📊 FTQ - MERCEDES BENZ")
st.markdown("*(Working model 1 Shift x 5 days)*")

# ttl=300 hace que vuelva a leer el Excel cada 5 minutos automáticamente
@st.cache_data(ttl=300)
def procesar_datos_completos(filepath):
    if filepath.endswith('.xlsx') or filepath.endswith('.xls'):
        df_raw = pd.read_excel(filepath, header=3, engine='openpyxl')
    else:
        df_raw = pd.read_csv(filepath, header=3)
    
    # --- DATOS DE PRODUCCIÓN ---
    df_prod = df_raw.iloc[:, 1:11].copy()
    df_prod.columns = ['Fecha', 'Semana', 'Linea', 'Maquina', 'Version', 'OK', 'NOK', 'Div', 'FTQ_Orig', 'Pct']
    
    df_prod = df_prod.dropna(subset=['Fecha', 'Version'])
    df_prod['Version'] = df_prod['Version'].astype(str).str.replace(r'\.0$', '', regex=True)
    df_prod['OK'] = pd.to_numeric(df_prod['OK'], errors='coerce').fillna(0)
    df_prod['NOK'] = pd.to_numeric(df_prod['NOK'], errors='coerce').fillna(0)
    df_prod['Fecha'] = pd.to_datetime(df_prod['Fecha'])
    df_prod['Semana'] = pd.to_numeric(df_prod['Semana'], errors='coerce').astype(int)
    df_prod['Maquina'] = df_prod['Maquina'].astype(str)
    
    # Factor y FTQ general
    df_prod['Factor'] = df_prod.apply(lambda r: (1 - (r['NOK']/r['OK'])) if r['OK'] > 0 else 1.0, axis=1)
    df_ftq = df_prod.groupby(['Fecha', 'Semana', 'Version'])['Factor'].prod().reset_index()
    df_ftq['FTQ'] = df_ftq['Factor'] * 100
    df_ftq['Mes_Num'] = df_ftq['Fecha'].dt.month

    # --- DATOS DE DEFECTOS ---
    df_def = df_raw.iloc[:, 12:19].copy()
    df_def.columns = ['Fecha', 'Semana', 'Linea', 'Maquina', 'Version', 'Defecto', 'Cantidad']
    df_def = df_def.dropna(subset=['Defecto', 'Cantidad'])
    df_def['Version'] = df_def['Version'].astype(str).str.replace(r'\.0$', '', regex=True)
    df_def['Cantidad'] = pd.to_numeric(df_def['Cantidad'], errors='coerce').fillna(0)
    df_def['Fecha'] = pd.to_datetime(df_def['Fecha'])
    df_def['Semana'] = pd.to_numeric(df_def['Semana'], errors='coerce').fillna(0).astype(int)
    df_def['Maquina'] = df_def['Maquina'].astype(str)

    # Retornamos df_prod también para poder analizar el FTQ por máquina
    return df_prod, df_ftq, df_def

def generar_grafica_fallas(df, top_n, titulo):
    if df.empty: return None
    data_fallas = df.groupby('Defecto')['Cantidad'].sum().reset_index()
    data_fallas = data_fallas.sort_values(by='Cantidad', ascending=False).head(top_n)

    fig = go.Figure(go.Bar(x=data_fallas['Defecto'], y=data_fallas['Cantidad'], text=data_fallas['Cantidad'], textposition='outside', marker_color=COLOR_KOSTAL))
    fig.update_layout(title=dict(text=titulo, font=dict(size=18, color=COLOR_KOSTAL)), yaxis=dict(title="Cantidad NOK"), xaxis=dict(tickangle=45), margin=dict(l=50, r=50, t=80, b=150), height=500)
    return fig

def generar_grafica_operaciones(df, top_n, titulo):
    if df.empty: return None
    data_maq = df.groupby('Maquina')['Cantidad'].sum().reset_index()
    data_maq = data_maq.sort_values(by='Cantidad', ascending=False).head(top_n)

    fig = go.Figure(go.Bar(x=data_maq['Maquina'], y=data_maq['Cantidad'], text=data_maq['Cantidad'], textposition='outside', marker_color=COLOR_KOSTAL))
    fig.update_layout(
        title=dict(text=titulo, font=dict(size=18, color=COLOR_KOSTAL)),
        yaxis=dict(title="Cantidad NOK"),
        xaxis=dict(title="Operación / Estación", tickangle=0, type='category', categoryorder='array', categoryarray=data_maq['Maquina']), 
        margin=dict(l=50, r=50, t=80, b=150), height=500
    )
    return fig

# --- RUTA DIRECTA DE LA BASE DE DATOS ---
archivo_bd = "BASE DE DATOS.xlsx"

if os.path.exists(archivo_bd):
    df_prod, df_ftq, df_def = procesar_datos_completos(archivo_bd)
    
    # --- MENÚ LATERAL: LÍNEA Y VERSIÓN ---
    st.sidebar.header("Filtros de Análisis")
    linea_sel = st.sidebar.radio("Selecciona la Línea:", ["SCR", "SCCM"])
    
    # Mapeo de versiones según la línea seleccionada
    if linea_sel == "SCR":
        versiones_permitidas = ["10532587"]
    else:
        #
        versiones_permitidas = ["12289497", "12289475"] 

    # Filtramos para mostrar solo las versiones que existen en la base de datos y que pertenecen a la línea elegida
    versiones_reales = df_ftq['Version'].dropna().unique()
    versiones_disponibles = [v for v in versiones_permitidas if v in versiones_reales]

    if len(versiones_disponibles) > 0:
        version_sel = st.sidebar.selectbox("Selecciona la Versión", versiones_disponibles)
        
        # Filtramos todas las tablas por la versión seleccionada
        df_ftq_l = df_ftq[df_ftq['Version'] == version_sel].sort_values('Fecha')
        df_def_l = df_def[df_def['Version'] == version_sel]
        df_prod_l = df_prod[df_prod['Version'] == version_sel]
        
        titulo_seccion = f"{linea_sel} - Versión {version_sel}"
        
        # ANÁLISIS GENERAL
        st.header(f"📈 Análisis General - {titulo_seccion}")
        
        resumen_gen = []
        meses_nombres = {1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun", 7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"}
        meses_presentes = sorted(df_ftq_l['Mes_Num'].unique())
        
        if len(meses_presentes) > 0:
            ultimo_mes = meses_presentes[-1]
            for num_mes in meses_presentes[:-1]:
                m_data = df_ftq_l[df_ftq_l['Mes_Num'] == num_mes]
                if not m_data.empty: resumen_gen.append({'P': meses_nombres.get(num_mes, str(num_mes)), 'V': m_data['FTQ'].mean()})
            
            sem_data = df_ftq_l[df_ftq_l['Mes_Num'] == ultimo_mes].groupby('Semana')['FTQ'].mean().reset_index()
            for _, r in sem_data.iterrows(): resumen_gen.append({'P': f"W{int(r['Semana'])}", 'V': r['FTQ']})
    
        df_g = pd.DataFrame(resumen_gen)
        
        if not df_g.empty:
            fig_g = go.Figure(go.Scatter(x=df_g['P'], y=df_g['V'], mode='lines+markers+text', text=[f"{v:.1f}%" for v in df_g['V']], textposition="top center", line=dict(color=COLOR_KOSTAL, width=3)))
            fig_g.add_hline(y=95, line_color=COLOR_TARGET, annotation_text="Target 95%")
            fig_g.add_hline(y=85, line_dash="dash", line_color=COLOR_ACTION, annotation_text="Action Limit 85%")
            fig_g.update_layout(yaxis_range=[min(df_g['V'].min()-5, 75), 115])
            st.plotly_chart(fig_g, use_container_width=True)
        else:
            st.warning("Aún no hay datos suficientes para mostrar la gráfica principal.")

        fig_p_gen = generar_grafica_fallas(df_def_l, 10, "Top 10 Defectos Históricos")
        if fig_p_gen: st.plotly_chart(fig_p_gen, use_container_width=True)

        st.divider()

        # DESGLOSE POR OPERACIÓN
        st.header("⚙️ FTQ por Operación (Comparativo)")
        st.markdown("Evolución semanal del FTQ simultánea para todas las estaciones.")
        
        if not df_prod_l.empty:
            # Calculamos el FTQ agrupando por Semana y Máquina a la vez
            df_op = df_prod_l.copy()
            df_op['FTQ_Estacion'] = df_op['Factor'] * 100
            tendencia_todas = df_op.groupby(['Semana', 'Maquina'])['FTQ_Estacion'].mean().reset_index()
            tendencia_todas['Semana_Str'] = "W" + tendencia_todas['Semana'].astype(str)

            if not tendencia_todas.empty:
                fig_op = go.Figure()
                
                # Iteramos por cada máquina para agregar su propia línea a la gráfica
                maquinas_unicas = sorted(tendencia_todas['Maquina'].unique())
                for maq in maquinas_unicas:
                    df_filtro = tendencia_todas[tendencia_todas['Maquina'] == maq]
                    fig_op.add_trace(go.Scatter(
                        x=df_filtro['Semana_Str'], 
                        y=df_filtro['FTQ_Estacion'], 
                        mode='lines+markers', 
                        name=f"{maq}"
                    ))
                
                fig_op.add_hline(y=95, line_color=COLOR_TARGET, annotation_text="Target 95%")
                fig_op.add_hline(y=85, line_dash="dash", line_color=COLOR_ACTION, annotation_text="Action Limit 85%")
                
                # Diseño
                fig_op.update_layout(
                    title=dict(text="Comparativo FTQ Semanal por Operación", font=dict(size=18, color=COLOR_KOSTAL)), 
                    yaxis_range=[min(tendencia_todas['FTQ_Estacion'].min()-5, 75), 115],
                    hovermode="x unified", # Agrupa todos los porcentajes al pasar el mouse sobre un punto
                    legend=dict(title="Estaciones")
                )
                st.plotly_chart(fig_op, use_container_width=True)

                with st.expander("🔎 Ver detalle en tabla"):
                    # 1. Crear tabla dinámica (Filas: Máquinas, Columnas: Semanas)
                    tabla_detalle = df_op.groupby(['Maquina', 'Semana'])['FTQ_Estacion'].mean().unstack()
                    
                    # 2. Formatear los encabezados de las columnas para que digan "W30", "W31", etc.
                    # Nota: Al estar agrupado por semana, omitimos "Fecha" diaria. La Versión ya está filtrada a nivel global.
                    tabla_detalle.columns = [f"W{int(c)}" for c in tabla_detalle.columns]
                    
                    # 3. Función para pintar de amarillo el valor más bajo de cada columna (semana)
                    def resaltar_minimo(s):
                        es_minimo = s == s.min()
                        return ['background-color: #ffe600; color: black; font-weight: bold' if v else '' for v in es_minimo]
                    
                    # 4. Aplicar los estilos: % con 1 decimal y la función de color
                    tabla_estilizada = tabla_detalle.style.apply(resaltar_minimo, axis=0).format("{:.1f}%", na_rep="-")
                    
                    st.dataframe(tabla_estilizada, use_container_width=True)

            else:
                st.info("No hay datos calculables para las operaciones.")
        
        st.divider()

        # DETALLE SEMANAL
        st.header(f"📅 Detalle Semanal")
        if not df_ftq_l.empty:
            sem_sel = st.selectbox("Seleccione Semana:", sorted(df_ftq_l['Semana'].unique()), index=len(df_ftq_l['Semana'].unique())-1)
            
            df_s = df_ftq_l[df_ftq_l['Semana'] == sem_sel]
            if not df_s.empty:
                inicio = df_s['Fecha'].min() - timedelta(days=df_s['Fecha'].min().weekday())
                dias = [inicio + timedelta(days=i) for i in range(5)]
                df_r = pd.merge(pd.DataFrame({'Fecha': dias}), df_s, on='Fecha', how='left').fillna({'FTQ': 100.0})
                df_r['F_Str'] = df_r['Fecha'].dt.strftime('%A %d-%b')
                
                fig_s = go.Figure(go.Scatter(x=df_r['F_Str'], y=df_r['FTQ'], mode='lines+markers+text', text=[f"{v:.1f}%" for v in df_r['FTQ']], textposition="top center", line=dict(color=COLOR_KOSTAL, width=3)))
                fig_s.add_hline(y=95, line_color=COLOR_TARGET)
                fig_s.add_hline(y=85, line_dash="dash", line_color=COLOR_ACTION)
                fig_s.update_layout(yaxis_range=[min(df_r['FTQ'].min()-5, 75), 115])
                st.plotly_chart(fig_s, use_container_width=True)
            else:
                st.info("No se registraron datos en esta semana.")

            df_def_s = df_def_l[df_def_l['Semana'] == sem_sel]

            fig_p_maq = generar_grafica_operaciones(df_def_s, 5, f"Mayores Ofensores por Operación - Semana {sem_sel}")
            if fig_p_maq: st.plotly_chart(fig_p_maq, use_container_width=True)

            fig_p_sem = generar_grafica_fallas(df_def_s, 5, f"Principales Fallas - Semana {sem_sel}")
            if fig_p_sem: st.plotly_chart(fig_p_sem, use_container_width=True)
            
        else:
            st.info("No hay datos semanales para la versión seleccionada.")
            
    else:
        st.warning(f"Aún no hay datos cargados para la línea {linea_sel}.")
else:
    st.error("No se encontró el archivo 'BASE DE DATOS.xlsx' en el directorio.")