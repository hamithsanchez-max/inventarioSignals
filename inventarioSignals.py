import streamlit as st
import pandas as pd
import os
from datetime import datetime
import io

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

            df_fuera = pd.DataFrame(columns=['ID_Registro', 'SKU', 'Descripcion', 'Cantidad', 'Ubicacion_Actual', 'Motivo_Estado', 'Responsable', 'Fecha_Registro'])
            df_fuera.to_excel(writer, sheet_name='Fuera_de_Almacen', index=False)

inicializar_excel()

st.title("📦 Sistema de Control de Inventario Taller")

# --- BARRA LATERAL: Opciones de Gestión de Archivo ---
st.sidebar.header("📁 Gestión de Base de Datos Excel")

# 1. Descargar Excel Actualizado
if os.path.exists(ARCHIVO_EXCEL):
    with open(ARCHIVO_EXCEL, "rb") as file:
        st.sidebar.download_button(
            label="⬇️ Descargar Excel Actualizado",
            data=file,
            file_name=f"Inventario_Taller_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# 2. Importar / Subir Excel Editado
uploaded_file = st.sidebar.file_uploader("⬆️ Cargar Excel Modificado", type=["xlsx"])
if uploaded_file is not None:
    if st.sidebar.button("Aplicar Cambios del Excel Subido"):
        with open(ARCHIVO_EXCEL, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.sidebar.success("✅ Base de datos actualizada correctamente desde el archivo subido.")
        st.rerun()

st.sidebar.divider()

# Menú lateral para navegar
opcion = st.sidebar.radio("Selecciona una Opción", [
    "Ver Inventario y Tablas",
    "Registrar Entrada / Salida / Retorno",
    "Registrar Remanente / Material de Proyecto",
    "Registrar Fuera de Stock / Fuera de Almacén"
])

# --- PESTAÑA 1: VISUALIZACIÓN ---
if opcion == "Ver Inventario y Tablas":
    st.header("📋 Registros Actuales")
    
    tab1, tab2, tab3, tab4 = st.tabs(["Maestro Inventario", "Movimientos", "Remanentes y Proyectos", "🚫 Fuera de Almacén"])
    
    with tab1:
        df_m = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Maestro_Inventario')
        st.dataframe(df_m)
        
    with tab2:
        df_mov = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Movimientos')
        st.dataframe(df_mov)
        
    with tab3:
        df_r = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Remanentes_Proyectos')
        st.dataframe(df_r)

    with tab4:
        try:
            df_f = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Fuera_de_Almacen')
            st.dataframe(df_f)
        except Exception:
            st.info("No hay registros en la pestaña 'Fuera_de_Almacen'.")

# --- PESTAÑA 2: MOVIMIENTOS GENERALES ---
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
                st.rerun()

# --- PESTAÑA 3: REMANENTES Y PROYECTOS ---
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
                st.rerun()

# --- PESTAÑA 4: FUERA DE STOCK / FUERA DE ALMACÉN ---
elif opcion == "Registrar Fuera de Stock / Fuera de Almacén":
    st.header("🚨 Clasificación: Fuera de Stock / Fuera de Almacén")
    st.caption("Usa esta opción para catalogar equipos o materiales asignados en campo, prestados, fuera de stock de recepción o en tránsito.")
    
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
            
        guardar_fuera = st.form_submit_button("Registrar Fuera de Almacén")
        
        if guardar_fuera:
            if not sku or not responsable:
                st.error("Por favor ingresa el SKU y el Responsable.")
            else:
                try:
                    df_fuera = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Fuera_de_Almacen')
                except Exception:
                    df_fuera = pd.DataFrame(columns=['ID_Registro', 'SKU', 'Descripcion', 'Cantidad', 'Ubicacion_Actual', 'Motivo_Estado', 'Responsable', 'Fecha_Registro'])
                
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
                
                with pd.ExcelWriter(ARCHIVO_EXCEL, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    df_fuera.to_excel(writer, sheet_name='Fuera_de_Almacen', index=False)
                    
                st.success(f"✅ Registro {nuevo_id} guardado correctamente.")
                st.rerun()