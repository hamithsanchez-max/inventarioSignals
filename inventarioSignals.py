import streamlit as st
import pandas as pd
import os
from datetime import datetime

ARCHIVO_EXCEL = "Inventario_Taller_Unificado.xlsx"

st.set_page_config(page_title="Control Simplificado de Inventario", layout="wide")

def inicializar_excel():
    """Crea un solo libro de Excel estructurado."""
    if not os.path.exists(ARCHIVO_EXCEL):
        with pd.ExcelWriter(ARCHIVO_EXCEL, engine='openpyxl') as writer:
            # 1. Catálogo / Maestro con Stock Mínimo
            df_maestro = pd.DataFrame(columns=[
                'SKU', 'Nombre_Item', 'Stock_Actual', 'Stock_Minimo', 'Unidad_Medida', 'Ubicacion'
            ])
            df_maestro.to_excel(writer, sheet_name='Maestro_Items', index=False)
            
            # 2. Historial Único de Movimientos y Ubicación
            df_movs = pd.DataFrame(columns=[
                'ID', 'Fecha', 'SKU', 'Tipo_Movimiento', 'Cantidad', 
                'Estado_Destino', 'Proyecto_OT', 'Responsable', 'Notas'
            ])
            df_movs.to_excel(writer, sheet_name='Historial_Movimientos', index=False)

inicializar_excel()

st.title("📦 Gestor de Inventario - Taller")

# --- BARRA LATERAL: Gestión del Archivo Excel ---
st.sidebar.header("📁 Base de Datos Excel")

if os.path.exists(ARCHIVO_EXCEL):
    with open(ARCHIVO_EXCEL, "rb") as file:
        st.sidebar.download_button(
            label="⬇️ Descargar Excel Completo",
            data=file,
            file_name=f"Inventario_Taller_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

uploaded_file = st.sidebar.file_uploader("⬆️ Subir/Restaurar Excel", type=["xlsx"])
if uploaded_file is not None:
    if st.sidebar.button("Aplicar Cambios"):
        with open(ARCHIVO_EXCEL, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.sidebar.success("✅ Base de datos actualizada.")
        st.rerun()

st.sidebar.divider()

# Navegación Simplificada (Solo 3 Opciones)
opcion = st.sidebar.radio("Navegación", [
    "🚨 Panel de Control y Stock Mínimo",
    "📝 Registrar Movimiento (Entrada/Salida)",
    "➕ Crear / Editar Ítem (Catálogo)"
])

# --- OPCIÓN 1: PANEL Y ALERTAS ---
if opcion == "🚨 Panel de Control y Stock Mínimo":
    st.header("📊 Estado del Inventario")
    
    df_m = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Maestro_Items')
    
    # ALERTAS DE REORDEN AUTOMÁTICAS
    if not df_m.empty:
        # Filtrar ítems donde Stock Actual <= Stock Minimo
        alertas = df_m[df_m['Stock_Actual'] <= df_m['Stock_Minimo']]
        
        if not alertas.empty:
            st.error(f"⚠️ ¡ATENCIÓN! Hay {len(alertas)} ítem(s) con necesidad urgente de reabastecimiento:")
            st.dataframe(alertas[['SKU', 'Nombre_Item', 'Stock_Actual', 'Stock_Minimo', 'Unidad_Medida']], use_container_width=True)
        else:
            st.success("✅ Todo el stock operativo está sobre el nivel mínimo.")
    
    st.divider()
    
    tab1, tab2 = st.tabs(["📋 Catálogo de Ítems y Cantidades", "📜 Historial de Movimientos"])
    with tab1:
        st.dataframe(df_m, use_container_width=True)
    with tab2:
        df_hist = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Historial_Movimientos')
        st.dataframe(df_hist, use_container_width=True)

# --- OPCIÓN 2: REGISTRO ÚNICO Y FÁCIL ---
elif opcion == "📝 Registrar Movimiento (Entrada/Salida)":
    st.header("📝 Registrar Entrada, Salida o Asignación")
    
    df_m = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Maestro_Items')
    
    if df_m.empty:
        st.warning("Primero debes registrar al menos un ítem en el Catálogo (Opción 'Crear / Editar Ítem').")
    else:
        lista_skus = df_m['SKU'].tolist()
        
        with st.form("form_movimiento_unico", clear_on_submit=True):
            col1, col2 = st.columns(2)
            
            with col1:
                sku_sel = st.selectbox("Seleccionar SKU / Ítem", lista_skus)
                tipo_mov = st.selectbox("Tipo de Operación", ["Entrada (Aumentar Stock)", "Salida (Reducir Stock)"])
                cantidad = st.number_input("Cantidad", min_value=0.1, step=1.0)
                
            with col2:
                estado_destino = st.selectbox("Clasificación / Destino del Material", [
                    "Almacén Principal (Stock Operativo)",
                    "Asignado a Proyecto / OT en Cursado",
                    "Sobrante / Remanente de Campo",
                    "Fuera de Almacén (En Campo / Vehículo)",
                    "En Reparación / Calibración"
                ])
                proyecto_ot = st.text_input("Proyecto / Orden de Trabajo (Opcional)")
                responsable = st.text_input("Tu Nombre / Responsable")
                notas = st.text_input("Observaciones / Notas")

            guardar = st.form_submit_button("💾 Guardar Movimiento")

            if guardar:
                if not responsable:
                    st.error("Ingresa el nombre del responsable.")
                else:
                    # 1. Registrar en el Historial
                    df_hist = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Historial_Movimientos')
                    nuevo_id = len(df_hist) + 1
                    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    
                    nuevo_mov = {
                        'ID': nuevo_id,
                        'Fecha': fecha_actual,
                        'SKU': sku_sel,
                        'Tipo_Movimiento': tipo_mov,
                        'Cantidad': cantidad,
                        'Estado_Destino': estado_destino,
                        'Proyecto_OT': proyecto_ot,
                        'Responsable': responsable,
                        'Notas': notas
                    }
                    df_hist = pd.concat([df_hist, pd.DataFrame([nuevo_mov])], ignore_index=True)
                    
                    # 2. Actualizar el Stock_Actual en el Maestro (si afecta Almacén Principal)
                    idx = df_m[df_m['SKU'] == sku_sel].index[0]
                    if "Entrada" in tipo_mov:
                        df_m.at[idx, 'Stock_Actual'] += cantidad
                    elif "Salida" in tipo_mov and "Almacén Principal" in estado_destino:
                        df_m.at[idx, 'Stock_Actual'] -= cantidad
                        
                    # Guardar ambos cambios en Excel
                    with pd.ExcelWriter(ARCHIVO_EXCEL, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        df_m.to_excel(writer, sheet_name='Maestro_Items', index=False)
                        df_hist.to_excel(writer, sheet_name='Historial_Movimientos', index=False)
                        
                    st.success(f"✅ Movimiento guardado correctamente. Stock de {sku_sel} actualizado.")
                    st.rerun()

# --- OPCIÓN 3: CREAR ÍTEMS EN EL CATÁLOGO ---
elif opcion == "➕ Crear / Editar Ítem (Catálogo)":
    st.header("➕ Agregar Nuevo Material o Componente al Catálogo")
    st.caption("Define el stock inicial y el nivel mínimo de alerta para cada parte.")
    
    with st.form("form_nuevo_item", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        with col1:
            sku = st.text_input("SKU / Código de Parte (Ej. SENSOR-2300)").strip().upper()
            nombre = st.text_input("Nombre / Descripción del Repuesto")
            unidad = st.selectbox("Unidad de Medida", ["Unidades (Unid)", "Metros (m)", "Juegos (Set)", "Litros (L)", "Cajas"])
            
        with col2:
            stock_ini = st.number_input("Stock Inicial en Almacén", min_value=0.0, step=1.0)
            stock_min = st.number_input("Stock Mínimo (Alerta de Compra)", min_value=1.0, step=1.0)
            ubicacion = st.text_input("Ubicación Física (Estante/Caja)")

        crear = st.form_submit_button("✨ Guardar Nuevo Ítem")

        if crear:
            if not sku or not nombre:
                st.error("Por favor completa el SKU y el Nombre del ítem.")
            else:
                df_m = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Maestro_Items')
                
                if sku in df_m['SKU'].values:
                    st.error(f"El SKU {sku} ya existe en el catálogo.")
                else:
                    nuevo_item = {
                        'SKU': sku,
                        'Nombre_Item': nombre,
                        'Stock_Actual': stock_ini,
                        'Stock_Minimo': stock_min,
                        'Unidad_Medida': unidad,
                        'Ubicacion': ubicacion
                    }
                    df_m = pd.concat([df_m, pd.DataFrame([nuevo_item])], ignore_index=True)
                    
                    with pd.ExcelWriter(ARCHIVO_EXCEL, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        df_m.to_excel(writer, sheet_name='Maestro_Items', index=False)
                        
                    st.success(f"🎉 Ítem {sku} creado exitosamente con nivel mínimo de {stock_min}.")
                    st.rerun()