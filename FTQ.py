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

# ttl=300 lee el Excel cada 5 minutos automáticamente
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

# Ruta a base de datos
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

    # Filtramos para mostrar solo las versiones que existen y pertenecen a la línea elegida
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

        # DESGLOSE POR OPERACIÓN - GRAFICA

        st.header("⚙️ FPY")
        #st.markdown("Rendimiento de calidad")
        
        df_prod_linea = df_prod[df_prod['Version'].isin(versiones_permitidas)].copy()
        
        if not df_prod_linea.empty:
            df_prod_linea['FTQ_Estacion'] = df_prod_linea['Factor'] * 100
            
            PALETA_AZULES = ["#004b87", "#3b82f6", "#87ceeb", "#1e3a8a", "#60a5fa", "#b0c4de"]
            
            # GRÁFICAS POR SEMANA
            datos_semanales = df_prod_linea.groupby(['Semana', 'Fecha', 'Maquina', 'Version'])['FTQ_Estacion'].mean().reset_index()
            semanas_presentes = sorted(datos_semanales['Semana'].unique())
            
            for sem in semanas_presentes:
                df_sem = datos_semanales[datos_semanales['Semana'] == sem].copy()
                
                def ordenar_maquinas(m):
                    try: return (0, int(m))
                    except: return (1, str(m))
                orden_maquinas = [str(x) for x in sorted(df_sem['Maquina'].unique(), key=ordenar_maquinas)]
                
                titulo_grafica = f"Semana W{int(sem)} - Rendimiento por Operación"
                fig_pareto = go.Figure()
                
                corridas = df_sem[['Fecha', 'Version']].drop_duplicates().sort_values(by=['Fecha', 'Version'])
                
                for i, (_, row) in enumerate(corridas.iterrows()):
                    f = row['Fecha']
                    v = row['Version']
                    df_corrida = df_sem[(df_sem['Fecha'] == f) & (df_sem['Version'] == v)]
                    
                    fecha_str = pd.to_datetime(f).strftime('%d/%m')
                    nombre_barra = f"{fecha_str} (Ver. {v})"
                    
                    # Asignacion de color
                    color_barra = PALETA_AZULES[i % len(PALETA_AZULES)]
                    
                    fig_pareto.add_trace(go.Bar(
                        x=df_corrida['Maquina'].astype(str), 
                        y=df_corrida['FTQ_Estacion'], 
                        name=nombre_barra,
                        text=[f"{val:.1f}%" for val in df_corrida['FTQ_Estacion']], 
                        textposition='auto',
                        marker_color=color_barra
                    ))
                    
                    # 3. Advertencia en FTQ bajo
                    if not df_corrida.empty:
                        min_ftq = df_corrida['FTQ_Estacion'].min()
                        peor_maquina = df_corrida[df_corrida['FTQ_Estacion'] == min_ftq]['Maquina'].iloc[0]
                        
                        fig_pareto.add_annotation(
                            x=str(peor_maquina),
                            y=min_ftq,
                            text="⚠️",
                            showarrow=True,
                            arrowhead=0,
                            yshift=5,
                            ax=0,
                            ay=-30,
                            font=dict(size=22),
                            hovertext=f"Mayor ofensor del {fecha_str}"
                        )
                
                fig_pareto.add_hline(y=95, line_color=COLOR_TARGET, annotation_text="Target 95%")
                fig_pareto.add_hline(y=85, line_dash="dash", line_color=COLOR_ACTION, annotation_text="Action Limit 85%")
                
                fig_pareto.update_layout(
                    title=dict(text=titulo_grafica, font=dict(size=16, color=COLOR_KOSTAL)), 
                    yaxis=dict(title="FTQ (%)", range=[min(df_sem['FTQ_Estacion'].min()-5, 75), 125]),
                    xaxis=dict(
                        title="Operación", 
                        type='category',
                        categoryorder='array',
                        categoryarray=orden_maquinas
                    ),
                barmode='group', 
                height=500,
                margin=dict(l=50, r=50, t=50, b=50),
                legend=dict(title="Día y Versión")
            )
                
            st.plotly_chart(fig_pareto, use_container_width=True)
                    
            # --- 2. TABLA DE DETALLES (MÚLTIPLES VERSIONES - DÍA POR DÍA CON SEMANA) ---
            with st.expander("🔎 Ver detalle en tabla (Todas las versiones - Día por Día)"):
                tabla_detalle = df_prod_linea.groupby(['Version', 'Maquina', 'Semana', 'Fecha'])['FTQ_Estacion'].mean().unstack(level=['Semana', 'Fecha'])
                tabla_detalle = tabla_detalle.sort_index(axis=1)
                
                tabla_detalle.columns = pd.MultiIndex.from_tuples(
                    [(f"W{int(sem)}", f.strftime('%d/%m/%Y')) for sem, f in tabla_detalle.columns],
                    names=["Semana", "Fecha"]
                )
                
                def aplicar_estilos_combinados(df_data):
                    estilos = pd.DataFrame('', index=df_data.index, columns=df_data.columns)
                    
                    for col in df_data.columns:
                        min_val = df_data[col].min()
                        
                        for idx in df_data.index:
                            val = df_data.at[idx, col]
                            version_fila = str(idx[0]) 
                            
                            is_min = pd.notna(val) and val == min_val
                            is_selected = (version_fila == str(version_sel))
                            
                            if is_min:
                                estilos.at[idx, col] = 'background-color: #ffe600; color: black; font-weight: bold;'
                            elif is_selected:
                                estilos.at[idx, col] = 'background-color: #e0e0e0; color: black;'
                                
                    return estilos
                    
                tabla_estilizada = tabla_detalle.style.apply(aplicar_estilos_combinados, axis=None).format("{:.1f}%", na_rep="-")
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