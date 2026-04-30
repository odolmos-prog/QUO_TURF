import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
import os
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="TURF Optimizer Pro", layout="wide")

st.title("🎯 TURF Optimizer Pro")
st.markdown("""
Esta aplicación realiza un análisis de **Alcance y Frecuencia Total No Duplicado** utilizando 
la lógica de **Relevancia Maestra**.
""")

# --- FUNCIONES DE APOYO ---
def calc_relevancia(df_act, op_eval, restantes):
    """Calcula la sinergia futura según la lógica de desempate conversada."""
    return sum((df_act[op_eval].astype(int) & df_act[r].astype(int)).sum() for r in restantes)

# --- CARGA DE ARCHIVO ---
uploaded_file = st.file_uploader("📂 Carga tu matriz binaria (Excel)", type=["xlsx"])

if uploaded_file:
    # 1. Carga y Limpieza
    df = pd.read_excel(uploaded_file)
    df.columns = [str(c).strip() for c in df.columns]
    
    opciones = []
    filtros_disponibles = []
    encontro_filtros = False
    
    for col in df.columns[1:]:
        valores_unicos = set(df[col].dropna().unique())
        es_binaria = valores_unicos.issubset({0, 1, 0.0, 1.0})
        if not encontro_filtros and es_binaria:
            opciones.append(col)
        else:
            encontro_filtros = True
            filtros_disponibles.append(col)
    
    st.sidebar.header("⚙️ Configuración")
    st.sidebar.write(f"**Estímulos:** {len(opciones)}")
    
    # 2. Selección de Filtros Interactiva
    df_trabajo = df.copy()
    if filtros_disponibles:
        st.sidebar.subheader("Filtrar Base de Datos")
        filtros_seleccionados = st.sidebar.multiselect("Elige hasta 3 variables", filtros_disponibles, max_selections=3)
        
        for filtro in filtros_seleccionados:
            valores_opc = sorted(df_trabajo[filtro].astype(str).unique())
            val_elegido = st.sidebar.selectbox(f"Valor para {filtro}", valores_opc)
            df_trabajo = df_trabajo[df_trabajo[filtro].astype(str).str.strip() == val_elegido.strip()]

    n_muestra = len(df_trabajo)
    st.info(f"📊 Muestra actual para análisis: **N = {n_muestra}**")

    if n_muestra > 0:
        # 3. Algoritmo TURF Interactivo[cite: 1]
        if 'seleccionadas' not in st.session_state:
            st.session_state.seleccionadas = []
            st.session_state.alcances_inc = []
            st.session_state.mask = pd.Series(False, index=df_trabajo.index)

        col1, col2 = st.columns([1, 2])

        with col1:
            st.subheader("🛠️ Optimización")
            disponibles = [o for o in opciones if o not in st.session_state.seleccionadas]
            
            if disponibles:
                datos_paso = []
                for op in disponibles:
                    incremental = (df_trabajo[op].astype(int) & (~st.session_state.mask)).sum()
                    futuras = [r for r in disponibles if r != op]
                    relevancia = calc_relevancia(df_trabajo, op, futuras)
                    datos_paso.append({'Opcion': op, 'Incremental': incremental, 'Relevancia': relevancia})
                
                ranking = pd.DataFrame(datos_paso).sort_values(by=['Incremental', 'Relevancia'], ascending=False).reset_index(drop=True)
                
                st.write("**Ranking de Sugerencias:**")
                st.dataframe(ranking.head(10), use_container_width=True)
                
                # Selección por nombre directo (más fácil en App que por código)
                op_elegida = st.selectbox("Selecciona el estímulo a agregar:", ranking['Opcion'])
                
                if st.button("➕ Agregar al Set"):
                    fila = ranking[ranking['Opcion'] == op_elegida].iloc[0]
                    st.session_state.seleccionadas.append(op_elegida)
                    st.session_state.mask |= df_trabajo[op_elegida].astype(bool)
                    st.session_state.alcances_inc.append((fila['Incremental'] / n_muestra) * 100)
                    st.rerun()
            
            if st.button("🔄 Reiniciar Análisis"):
                for key in ['seleccionadas', 'alcances_inc', 'mask']:
                    if key in st.session_state: del st.session_state[key]
                st.rerun()

        with col2:
            if st.session_state.seleccionadas:
                # 4. Resultados y Visualización[cite: 1]
                resumen_data = []
                acum = 0
                for nom, alc in zip(st.session_state.seleccionadas, st.session_state.alcances_inc):
                    acum += alc
                    resumen_data.append({'Jerarquía': len(resumen_data)+1, 'Estímulo': nom, '% Inc': round(alc, 2), '% Acum': round(acum, 2)})
                
                df_jerarquia = pd.DataFrame(resumen_data)
                st.subheader("📋 Tabla de Jerarquía")
                st.table(df_jerarquia)

                # Gráfico con tus parámetros exactos[cite: 1]
                fig, ax = plt.subplots(figsize=(10, 5))
                colores = ['#2E8B57'] + ['#90EE90'] * (len(st.session_state.seleccionadas) - 1)
                base = 0
                for i, (nombre, valor) in enumerate(zip(st.session_state.seleccionadas, st.session_state.alcances_inc)):
                    ax.bar(nombre, valor, bottom=base, color=colores[i], edgecolor='white', alpha=0.8)
                    ax.text(i, base + valor/2, f"{valor:.1f}%", ha='center', va='center', fontweight='bold', fontsize=8)
                    base += valor

                ax.set_title(f"Alcance Acumulado: {base:.2f}% (N = {n_muestra})", fontweight='bold')
                plt.xticks(rotation=90, fontsize=8)
                plt.grid(axis='y', linestyle=':', alpha=0.6)
                st.pyplot(fig)

                # 5. Exportación a Excel[cite: 1]
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_jerarquia.to_excel(writer, sheet_name='TURF', index=False)
                    
                    # Guardamos imagen temporal para el Excel
                    temp_img = "temp_app_plot.png"
                    fig.savefig(temp_img, dpi=100)
                    ws_grafico = writer.book.create_sheet('Gráfico')
                    img = OpenpyxlImage(temp_img)
                    ws_grafico.add_image(img, 'B2')
                
                st.download_button(
                    label="📥 Descargar Reporte Excel",
                    data=buffer.getvalue(),
                    file_name="Reporte_TURF_Final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                if os.path.exists(temp_img): os.remove(temp_img)
    else:
        st.warning("⚠️ No hay datos para los filtros seleccionados.")