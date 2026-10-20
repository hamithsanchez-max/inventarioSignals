import streamlit as st
import pandas as pd
import os
from datetime import datetime

ARCHIVO_EXCEL = "Inventario_Taller.xlsx"

st.set_page_config(page_title="Control de Inventario - Taller", layout="wide")

def inicializar_excel():
    """Crea la estructura base del libro Excel si no existe."""
    if not os.path.exists(ARCHIVO_EXCEL):
        with pd.ExcelWriter(ARCHIVO_EXCEL, engine='openpyxl') as writer:
            df_maestro = pd.DataFrame(columns=['SKU', 'Nombre_Item', 'Categoria', 'Ubicacion', 'Stock_Minimo', 'Unidad_Medida'])
            df_maestro.to_excel(writer, sheet_name='Maestro_Inventario', index=False)
            
            df_movs = pd.DataFrame(columns=['ID_Movimiento', 'Fecha', 'Tipo_Movimiento', 'SKU', 'Cantidad', 'Origen_Destino', 'Proyecto_OT', 'Registrado_Por'])
            df_movs.to_excel(writer, sheet_name='Movimientos', index=False)
            
            df_rem = pd.DataFrame(columns=['ID_Remanente', 'SKU', 'Proyecto_Origen', 'Estado', 'Cantidad', 'Ubicacion_Especial', 'Notas'])
            df_rem.to_excel(writer, sheet_name='Remanentes_Proyectos', index=False)

inicializar_excel()

st.title("📦 Sistema de Control de Inventario Taller")

# Menú lateral para navegar
opcion = st.sidebar.radio("Selecciona una Opción", [
    "Ver Inventario y Tablas",
    "Registrar Entrada / Salida / Retorno",
    "Registrar Remanente / Material de Proyecto"
])

if opcion == "Ver Inventario y Tablas":
    st.header("📋 Registros Actuales")
    
    tab1, tab2, tab3 = st.tabs(["Maestro Inventario", "Movimientos", "Remanentes y Proyectos"])
    
    with tab1:
        df_m = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Maestro_Inventario')
        st.dataframe(df_m, use_container_width=True)
        
    with tab2:
        df_mov = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Movimientos')
        st.dataframe(df_mov, use_container_width=True)
        
    with tab3:
        df_r = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Remanentes_Proyectos')
        st.dataframe(df_r, use_container_width=True)

elif opcion == "Registrar Entrada / Salida / Retorno":
    st.header("📥 / 📤 Registrar Movimiento")
    
    with st.form("form_movimiento", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            sku = st.text_input("SKU / Código del Material").strip().upper()
            tipo = st.selectbox("Tipo de Movimiento", ["Entrada", "Salida", "Retorno_Remanente"])
            cantidad = st.number_input("Cantidad", min_value=0.1, step=1.0)
        with col2:
            origen_destino = st.text_input("Origen / Destino (Ej. Proveedor, Campo, Taller)")
            proyecto_ot = st.text_input("Código de Proyecto / OT (Opcional)")
            usuario = st.text_input("Registrado Por")
            
        guardar = st.form_submit_button("Guardar Movimiento")
        
        if guardar:
            if not sku or not usuario:
                st.error("Por favor completa al menos el SKU y la persona que registra.")
            else:
                df_movs = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Movimientos')
                nuevo_id = len(df_movs) + 1
                fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                nuevo_registro = {
                    'ID_Movimiento': nuevo_id,
                    'Fecha': fecha_actual,
                    'Tipo_Movimiento': tipo,
                    'SKU': sku,
                    'Cantidad': cantidad,
                    'Origen_Destino': origen_destino,
                    'Proyecto_OT': proyecto_ot,
                    'Registrado_Por': usuario
                }
                
                df_movs = pd.concat([df_movs, pd.DataFrame([nuevo_registro])], ignore_index=True)
                
                with pd.ExcelWriter(ARCHIVO_EXCEL, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    df_movs.to_excel(writer, sheet_name='Movimientos', index=False)
                    
                st.success(f"✅ Movimiento #{nuevo_id} registrado exitosamente para {sku}.")

elif opcion == "Registrar Remanente / Material de Proyecto":
    st.header("🧱 Registrar Remanente o Material de Proyecto")
    
    with st.form("form_remanente", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            sku = st.text_input("SKU / Código del Material").strip().upper()
            proyecto = st.text_input("Proyecto de Origen / Asignado")
            cantidad = st.number_input("Cantidad Sobrante / Asignada", min_value=0.1, step=1.0)
        with col2:
            estado = st.selectbox("Estado del Material", ["Disponible", "Asignado", "En Uso", "Por Revisar"])
            ubicacion = st.text_input("Ubicación Física Especial en Taller")
            notas = st.text_area("Observaciones / Condición del Material")
            
        guardar_rem = st.form_submit_button("Guardar Remanente")
        
        if guardar_rem:
            if not sku or not proyecto:
                st.error("Por favor completa el SKU y el Proyecto.")
            else:
                df_rem = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Remanentes_Proyectos')
                nuevo_id = f"REM-{len(df_rem) + 1:04d}"
                
                nuevo_remanente = {
                    'ID_Remanente': nuevo_id,
                    'SKU': sku,
                    'Proyecto_Origen': proyecto,
                    'Estado': estado,
                    'Cantidad': cantidad,
                    'Ubicacion_Especial': ubicacion,
                    'Notas': notas
                }
                
                df_rem = pd.concat([df_rem, pd.DataFrame([nuevo_remanente])], ignore_index=True)
                
                with pd.ExcelWriter(ARCHIVO_EXCEL, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    df_rem.to_excel(writer, sheet_name='Remanentes_Proyectos', index=False)
                    
                st.success(f"✅ Remanente {nuevo_id} guardado exitosamente.")
