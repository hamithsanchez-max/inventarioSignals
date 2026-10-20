import pandas as pd
import os
from datetime import datetime

ARCHIVO_EXCEL = "Inventario_Taller.xlsx"

def inicializar_excel():
    """Crea la estructura base del libro Excel si no existe."""
    if not os.path.exists(ARCHIVO_EXCEL):
        with pd.ExcelWriter(ARCHIVO_EXCEL, engine='openpyxl') as writer:
            # Pestaña Maestro
            df_maestro = pd.DataFrame(columns=[
                'SKU', 'Nombre_Item', 'Categoria', 'Ubicacion', 'Stock_Minimo', 'Unidad_Medida'
            ])
            df_maestro.to_excel(writer, sheet_name='Maestro_Inventario', index=False)
            
            # Pestaña Movimientos
            df_movs = pd.DataFrame(columns=[
                'ID_Movimiento', 'Fecha', 'Tipo_Movimiento', 'SKU', 'Cantidad', 'Origen_Destino', 'Proyecto_OT', 'Registrado_Por'
            ])
            df_movs.to_excel(writer, sheet_name='Movimientos', index=False)
            
            # Pestaña Remanentes y Proyectos
            df_rem = pd.DataFrame(columns=[
                'ID_Remanente', 'SKU', 'Proyecto_Origen', 'Estado', 'Cantidad', 'Ubicacion_Especial', 'Notas'
            ])
            df_rem.to_excel(writer, sheet_name='Remanentes_Proyectos', index=False)
        print(f"Base de datos creada exitosamente: {ARCHIVO_EXCEL}")

def registrar_movimiento():
    """Registra una entrada, salida o retorno de material."""
    print("\n--- REGISTRO DE MOVIMIENTO DE INVENTARIO ---")
    sku = input("SKU / Código del material: ").strip().upper()
    tipo = input("Tipo de Movimiento (Entrada / Salida / Retorno_Remanente): ").strip().capitalize()
    cantidad = float(input("Cantidad: "))
    origen_destino = input("Origen / Destino (Proveedor, Área, Taller): ").strip()
    proyecto_ot = input("Código de Proyecto u Orden de Trabajo (opcional): ").strip()
    usuario = input("Nombre de quien registra: ").strip()

    # Cargar datos actuales
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

    df_movs = pd.concat([df_movs, pd.DataFrame([nuevo_registro])], ignore_index=False)

    # Guardar en Excel
    with pd.ExcelWriter(ARCHIVO_EXCEL, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_movs.to_excel(writer, sheet_name='Movimientos', index=False)

    print(f"✅ Movimiento #{nuevo_id} guardado correctamente.")

def registrar_remanente():
    """Registra materiales sobrantes o específicos de un proyecto en ejecución."""
    print("\n--- REGISTRO DE REMANENTE O MATERIAL DE PROYECTO ---")
    sku = input("SKU / Código del material: ").strip().upper()
    proyecto = input("Proyecto de Origen / Asignado: ").strip()
    estado = input("Estado del Material (Disponible / Asignado / En Uso / Por Revisar): ").strip().capitalize()
    cantidad = float(input("Cantidad sobrante/asignada: "))
    ubicacion = input("Ubicación física especial en taller: ").strip()
    notas = input("Observaciones/Condición del material: ").strip()

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

    df_rem = pd.concat([df_rem, pd.DataFrame([nuevo_remanente])], ignore_index=False)

    with pd.ExcelWriter(ARCHIVO_EXCEL, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_rem.to_excel(writer, sheet_name='Remanentes_Proyectos', index=False)

    print(f"✅ Remanente {nuevo_id} guardado correctamente.")

def menu():
    inicializar_excel()
    while True:
        print("\n==========================================")
        print("  SISTEMA DE CONTROL DE INVENTARIO TALLER ")
        print("==========================================")
        print("1. Registrar Entrada / Salida / Retorno")
        print("2. Registrar Remanente o Material de Proyecto")
        print("3. Salir")
        
        opcion = input("Seleccione una opción (1-3): ").strip()
        
        if opcion == '1':
            registrar_movimiento()
        elif opcion == '2':
            registrar_remanente()
        elif opcion == '3':
            print("Cerrando el programa...")
            break
        else:
            print("Opción inválida. Intente de nuevo.")

if __name__ == "__main__":
    menu()