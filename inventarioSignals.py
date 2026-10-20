import streamlit as st
import pandas as pd
import os
from datetime import datetime

# Archivos de Excel Independientes
ARCHIVO_EXCEL_INVENTARIO = "Inventario_Taller.xlsx"
ARCHIVO_EXCEL_FUERA = "Inventario_Fuera_Almacen.xlsx"

st.set_page_config(page_title="Control de Inventario - Taller", layout="wide")

def inicializar_excels():
    """Crea la estructura base de ambos libros de Excel si no existen."""
    # 1. Excel Principal
    if not os.path.exists(ARCHIVO_EXCEL_INVENTARIO):
        with pd.ExcelWriter(ARCHIVO_EXCEL_INVENTARIO, engine='openpyxl') as writer:
            df_maestro = pd.DataFrame(columns=['SKU', 'Nombre_Item', 'Categoria', 'Ubicacion', 'Stock_Minimo', 'Unidad_Medida'])
            df_maestro.to_excel(writer, sheet_name='Maestro_Inventario', index=False)
            
            df_movs = pd.DataFrame(columns=['ID_Movimiento', 'Fecha', 'Tipo_Movimiento', 'SKU', 'Cantidad', 'Origen_Destino', 'Proyecto_OT', 'Registrado_Por'])
            df_movs.to_excel(writer, sheet_name='Movimientos', index=False)
            
            df_rem = pd.DataFrame(columns=['ID_Remanente', 'SKU', 'Proyecto_Origen', 'Estado', 'Cantidad', 'Ubicacion_Especial', 'Notas'])
            df_rem.to_excel(writer, sheet_name='Remanentes_Proyectos', index=False)

    # 2. Excel Exclusivo para Fuera de Stock / Fuera de Almacén
    if not os.path.exists(ARCHIVO_EXCEL_FUERA):
        with pd.ExcelWriter(ARCHIVO_EXCEL_FUERA, engine='openpyxl') as writer:
            df_fuera = pd.DataFrame(columns=['ID_Registro', 'SKU', 'Descripcion', 'Cantidad', 'Ubicacion_Actual', 'Motivo_Estado', 'Responsable', 'Fecha_Registro'])
            df_fuera.to_excel(writer, sheet_name='Fuera_de_Almacen', index=False)

inicializar_excels()

st.title("📦 Sistema de Control de Inventario Taller")

# --- BARRA LATERAL: Gestión de Base de Datos ---
st.sidebar.header("📁 Archivo Inventario Principal")

# Descargar e Importar Excel Principal
if os.path.exists(ARCHIVO_EXCEL_INVENTARIO):
    with open(ARCHIVO_EXCEL_INVENTARIO, "rb") as file:
        st.sidebar.download_button(
            label="⬇️ Descargar Excel Principal",
            data=file,
            file_name=f"Inventario_Taller_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

uploaded_main = st.sidebar.file_uploader("⬆️ Cargar Excel Principal", type=["xlsx"], key="main_excel")
if uploaded_main is not None:
    if st.sidebar.button("Aplicar Cambios (Principal)"):
        with open(ARCHIVO_EXCEL_INVENTARIO, "wb") as f:
            f.write(uploaded_main.getbuffer())
        st.sidebar.success("✅ Base de datos principal actualizada.")
        st.rerun()

st.sidebar.divider()
st.sidebar.header("🚨 Archivo Fuera de Almacén")

# Descargar e Importar Excel Fuera de Almacén
if os.path.exists(ARCHIVO_EXCEL_FUERA):
    with open(ARCHIVO_EXCEL_FUERA, "rb") as file_f:
        st.sidebar.download_button(
            label="⬇️ Descargar Excel Fuera de Almacén",
            data=file_f,
            file_name=f"Fuera_de_Almacen_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

uploaded_fuera = st.sidebar.file_uploader("⬆️ Cargar Excel Fuera de Almacén", type=["xlsx"], key="fuera_excel")
if uploaded_fuera is not None:
    if st.sidebar.button("Aplicar Cambios (Fuera de Almacén)"):
        with open(ARCHIVO_EXCEL_FUERA, "wb") as f:
            f.write(uploaded_fuera.getbuffer())
        st.sidebar.success("✅ Base de datos 'Fuera de Almacén' actualizada.")
        st.rerun()

st.sidebar.divider()

# Navegación
opcion = st.sidebar.radio("Selecciona una Opción", [
    "Ver Inventario y Tablas",
    "Ver Tablas Fuera de Stock / Almacén",
    "Registrar Entrada / Salida / Retorno",
    "Registrar Remanente / Material de Proyecto",
    "Registrar Fuera de Stock / Fuera de Almacén"
])

# --- PESTAÑA 1: VISUALIZACIÓN INVENTARIO PRINCIPAL ---
if opcion == "Ver Inventario y Tablas":
    st.header("📋 Registros de Inventario Principal")
    
    tab1, tab2, tab3 = st.tabs(["Maestro Inventario", "Movimientos", "Remanentes y Proyectos"])
    
    with tab1:
        df_m = pd.read_excel(ARCHIVO_EXCEL_INVENTARIO, sheet_name='Maestro_Inventario')
        st.dataframe(df_m, use_container_width=True)
        
    with tab2:
        df_mov = pd.read_excel(ARCHIVO_EXCEL_INVENTARIO, sheet_name='Movimientos')
        st.dataframe(df_mov, use_container_width=True)
        
    with tab3:
        df_r = pd.read_excel(ARCHIVO_EXCEL_INVENTARIO, sheet_name='Remanentes_Proyectos')
        st.dataframe(df_r, use_container_width=True)

# --- PESTAÑA 2: VISUALIZACIÓN FUERA DE STOCK / ALMACÉN ---
elif opcion == "Ver Tablas Fuera de Stock / Almacén":
    st.header("🚨 Registros Fuera de Stock / Fuera de Almacén")
    
    df_f = pd.read_excel(ARCHIVO_EXCEL_FUERA, sheet_name='Fuera_de_Almacen')
    st.dataframe(df_f, use_container_width=True)

# --- REGISTRO MOVIMIENTOS ---
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
                df_movs = pd.read_excel(ARCHIVO_EXCEL_INVENTARIO, sheet_name='Movimientos')
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
                
                with pd.ExcelWriter(ARCHIVO_EXCEL_INVENTARIO, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    df_movs.to_excel(writer, sheet_name='Movimientos', index=False)
                    
                st.success(f"✅ Movimiento #{nuevo_id} registrado exitosamente para {sku}.")
                st.rerun()

# --- REGISTRO REMANENTES ---
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
                df_rem = pd.read_excel(ARCHIVO_EXCEL_INVENTARIO, sheet_name='Remanentes_Proyectos')
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
                
                with pd.ExcelWriter(ARCHIVO_EXCEL_INVENTARIO, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    df_rem.to_excel(writer, sheet_name='Remanentes_Proyectos', index=False)
                    
                st.success(f"✅ Remanente {nuevo_id} guardado exitosamente.")
                st.rerun()

# --- REGISTRO FUERA DE STOCK / ALMACÉN ---
elif opcion == "Registrar Fuera de Stock / Fuera de Almacén":
    st.header("🚨 Registrar Equipo/Material Fuera de Stock o Almacén")
    
    with st.form("form_fuera_almacen", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            sku = st.text_input("SKU / Código del Equipo o Material").strip().upper()
            descripcion = st.text_input("Descripción breve del Equipo/Material")
            cantidad = st.number_input("Cantidad Fuera de Almacén", min_value=0.1, step=1.0)
            ubicacion_actual = st.text_input("Ubicación Actual (Ej. Sitio de Vía, Vehículo Taller, Cliente)")
        with col2:
            motivo = st.selectbox("Motivo / Clasificación", [
                "Fuera de Stock / Remanente de Campo",
                "Asignado a Proyecto en Ejecución",
                "Prestado / En Tránsito",
                "En Reparación / Calibración",
                "Dañado / Para Descarte"
            ])
            responsable = st.text_input("Responsable o Custodio")
            
        guardar_fuera = st.form_submit_button("Registrar en Fuera de Almacén")
        
        if guardar_fuera:
            if not sku or not responsable:
                st.error("Por favor ingresa el SKU y el Responsable.")
            else:
                df_fuera = pd.read_excel(ARCHIVO_EXCEL_FUERA, sheet_name='Fuera_de_Almacen')
                nuevo_id = f"FA-{len(df_fuera) + 1:04d}"
                fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                nuevo_registro_fuera = {
                    'ID_Registro': nuevo_id,
                    'SKU': sku,
                    'Descripcion': descripcion,
                    'Cantidad': cantidad,
                    'Ubicacion_Actual': ubicacion_actual,
                    'Motivo_Estado': motivo,
                    'Responsable': responsable,
                    'Fecha_Registro': fecha_actual
                }
                
                df_fuera = pd.concat([df_fuera, pd.DataFrame([nuevo_registro_fuera])], ignore_index=True)
                
                with pd.ExcelWriter(ARCHIVO_EXCEL_FUERA, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    df_fuera.to_excel(writer, sheet_name='Fuera_de_Almacen', index=False)
                    
                st.success(f"✅ Registro {nuevo_id} guardado correctamente en 'Fuera_de_Almacen.xlsx'.")
                st.rerun()