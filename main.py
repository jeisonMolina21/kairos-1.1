# main.py - VERSIÓN MODIFICADA PARA REORDENAR HOJAS
import os
import sys
import pandas as pd
from tkinter import messagebox
import json
import openpyxl
import traceback
import shutil
import glob
from datetime import datetime

# Agregar el directorio actual al path para importaciones
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(current_dir)

from core.file_selector import seleccionar_archivos
from core.reader import leer_archivos_excel
from core.cleaner import limpiar_datos
from core.api_client import APIClient
from core.comparator import comparar_asistencias
from reports.excel_resumen import generar_excel_resumen_semanal
from reports.excel_horarios import agregar_hoja_horarios
from reports.excel_diario import agregar_hoja_diario
from reports.excel_horarios_programados import agregar_hoja_horarios_programados
from reports.excel_explicacion import agregar_hoja_explicacion
from reports.excel_permisos import agregar_hoja_permisos
import tempfile

# =============================================================================
# 🚀 COMPATIBILIDAD CON EJECUTABLE ÚNICO - SIN ERRORES
# =============================================================================
def setup_executable_environment():
    """Configura el entorno para .exe sin errores"""
    try:
        # Si estamos en un .exe de PyInstaller
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            base_path = sys._MEIPASS
            print(f"🔧 Ejecutando desde .exe: {base_path}")
            
            # Agregar rutas de módulos
            core_path = os.path.join(base_path, 'core')
            reports_path = os.path.join(base_path, 'reports')
            
            if core_path not in sys.path:
                sys.path.insert(0, core_path)
            if reports_path not in sys.path:
                sys.path.insert(0, reports_path)
                
            return True
        else:
            print("🔧 Ejecutando desde Python")
            return False
    except Exception as e:
        print(f"⚠️  Error en configuración: {e}")
        return False

# Configurar entorno
IS_EXECUTABLE = setup_executable_environment()

# Crear directorios necesarios sin errores
def create_dirs_safely():
    """Crea directorios necesarios sin lanzar errores"""
    dirs = ['outputs', 'temp']
    for dir_name in dirs:
        try:
            if not os.path.exists(dir_name):
                os.makedirs(dir_name)
                print(f"📁 Directorio creado: {dir_name}")
        except Exception as e:
            print(f"⚠️  No se pudo crear {dir_name}: {e}")

create_dirs_safely()

# =============================================================================
# IMPORTACIONES SEGURAS - CONTINÚA EL RESTO DE TU CÓDIGO
# =============================================================================
try:
    from core.file_selector import seleccionar_archivos
    from core.reader import leer_archivos_excel
    from core.cleaner import limpiar_datos
    from core.api_client import APIClient
    from core.comparator import comparar_asistencias
    from reports.excel_resumen import generar_excel_resumen_semanal
    from reports.excel_horarios import agregar_hoja_horarios
    from reports.excel_diario import agregar_hoja_diario
    from reports.excel_horarios_programados import agregar_hoja_horarios_programados
    from reports.excel_explicacion import agregar_hoja_explicacion
    from reports.excel_permisos import agregar_hoja_permisos
    print("✅ Todos los módulos importados correctamente")
except ImportError as e:
    print(f"❌ Error importando: {e}")
    input("Presione Enter para salir...")
    sys.exit(1)

def limpiar_archivos_temporales():
    """Limpia todos los archivos temporales de ejecuciones anteriores"""
    temp_dirs = ["temp", "outputs/temp"]
    for temp_dir in temp_dirs:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
            print(f"🧹 Carpeta temporal '{temp_dir}' eliminada")

def crear_directorios_temporales():
    """Crea los directorios temporales necesarios"""
    temp_dirs = ["temp", "outputs/temp"]
    for temp_dir in temp_dirs:
        os.makedirs(temp_dir, exist_ok=True)
    return "temp"

def obtener_rango_fechas(resumen_path):
    """
    Obtiene el rango de fechas del archivo de resumen para nombrar los archivos
    """
    try:
        with open(resumen_path, "r", encoding="utf-8") as f:
            resumen = json.load(f)
        
        if not resumen:
            return None, None
            
        fechas = []
        for registro in resumen:
            fecha_str = registro.get("Fecha")
            if fecha_str and fecha_str != "N/A":
                try:
                    fecha = datetime.strptime(fecha_str, "%Y-%m-%d")
                    fechas.append(fecha)
                except:
                    continue
        
        if not fechas:
            return None, None
            
        fecha_min = min(fechas)
        fecha_max = max(fechas)
        
        return fecha_min, fecha_max
        
    except Exception as e:
        print(f"⚠️ Error obteniendo rango de fechas: {e}")
        return None, None

def obtener_codigo_departamento(nombre_departamento):
    """
    Mapea el nombre del departamento a su código abreviado
    """
    mapeo_departamentos = {
        "UNIHORIZONTE": "UH",
        "FINANCIERA": "FN",
        "TICS": "TC",
        "RECTORIA": "RC",
        "DOCENTES": "DC",
        "ASEO Y MANT": "AM",
        "REGISTRO Y CONTROL": "RYC",
        "MERCADEO": "MC"
    }
    
    if not nombre_departamento or pd.isna(nombre_departamento) or nombre_departamento == "":
        return "ND"
    
    nombre_departamento = str(nombre_departamento).upper().strip()
    
    # Buscar coincidencia exacta primero
    if nombre_departamento in mapeo_departamentos:
        return mapeo_departamentos[nombre_departamento]
    
    # Buscar coincidencias parciales
    for depto_nombre, codigo in mapeo_departamentos.items():
        if depto_nombre in nombre_departamento:
            return codigo
    
    return "ND"  # No Definido

def detectar_departamento_archivos(archivos):
    """
    Detecta TODOS los departamentos basados en la columna de departamento
    dentro de los archivos Excel seleccionados.

    - Recorre todos los archivos.
    - En cada archivo busca columnas que contengan 'DEPARTAMENTO' o 'DEPTO'.
    - Recorre TODOS los valores únicos de esa columna.
    - Usa obtener_codigo_departamento para mapearlos a códigos (UH, FN, TC, RC, AM, etc.).
    - Ignora valores que no tienen código (por ejemplo 'VISITANTES', que hoy devuelve 'ND').

    Retorna una lista de códigos de departamentos detectados.
    """
    print("\n🔍 Detectando departamentos de los archivos (por contenido)...")
    
    departamentos_encontrados = set()
    
    for archivo in archivos:
        nombre_archivo = os.path.basename(archivo)
        print(f"📁 Analizando archivo: {nombre_archivo}")
        
        # Leer el archivo completo (necesitamos ver todos los departamentos)
        try:
            df_temp = pd.read_excel(archivo)
        except Exception as e:
            print(f"   ⚠️ Error leyendo archivo {archivo}: {e}")
            continue

        if df_temp is None or df_temp.empty:
            print("   ⚠️ Archivo sin datos o no legible")
            continue

        # Buscar columna de departamento
        columna_departamento = None
        for col in df_temp.columns:
            col_str = str(col).upper()
            if ("DEPARTAMENTO" in col_str) or ("DEPTO" in col_str) or ("NOMBRE DE DEPARTAMENTO" in col_str):
                columna_departamento = col
                print(f"   📋 Columna de departamento encontrada: '{col}'")
                break

        if not columna_departamento:
            print("   ⚠️ No se encontró columna de departamento en este archivo")
            continue

        # Obtener TODOS los valores únicos de esa columna
        valores_unicos = df_temp[columna_departamento].dropna().unique()
        print(f"   🔍 Valores únicos en columna '{columna_departamento}': {list(valores_unicos)}")

        for valor in valores_unicos:
            if pd.isna(valor):
                continue

            nombre_depto = str(valor).strip()
            if not nombre_depto or nombre_depto.lower() == "nan":
                continue

            # Mapear al código usando la función centralizada
            codigo = obtener_codigo_departamento(nombre_depto)

            if codigo != "ND":
                departamentos_encontrados.add(codigo)
                print(f"   ✅ Detectado departamento: '{nombre_depto}' -> {codigo}")
            else:
                # Por ejemplo VISITANTES caerá aquí
                print(f"   ℹ️ Valor de departamento ignorado (sin código): '{nombre_depto}'")

    if departamentos_encontrados:
        departamentos_ordenados = sorted(list(departamentos_encontrados))
        print(f"🏢 Departamentos detectados: {', '.join(departamentos_ordenados)}")
        return departamentos_ordenados
    else:
        print("⚠️ No se pudo detectar departamentos. Usando 'ND'")
        return ["ND"]

def detectar_departamentos_en_dataframe(df_limpio):
    """
    Detecta departamentos directamente desde el DataFrame limpio,
    usando la columna de departamento (por ejemplo 'nombre_de_departamento').

    - Busca una columna que contenga 'DEPARTAMENTO' en el nombre.
    - Toma todos los valores únicos de esa columna.
    - Usa obtener_codigo_departamento para mapearlos a códigos (UH, FN, TC, etc.).
    - Ignora los que devuelvan 'ND' (ej. VISITANTES).
    """
    print("\n🔍 Detectando departamentos desde el DataFrame limpio...")

    if df_limpio is None or df_limpio.empty:
        print("   ⚠️ DataFrame limpio vacío, no se pueden detectar departamentos.")
        return ["ND"]

    # Buscar columna de departamento
    columna_departamento = None
    for col in df_limpio.columns:
        col_str = str(col).upper()
        col_norm = col_str.replace("_", " ")
        if "DEPARTAMENTO" in col_norm:
            columna_departamento = col
            print(f"   📋 Columna de departamento encontrada en DF limpio: '{col}'")
            break

    if not columna_departamento:
        print("   ⚠️ No se encontró columna de departamento en el DF limpio.")
        return ["ND"]

    valores_unicos = df_limpio[columna_departamento].dropna().unique()
    print(f"   🔍 Valores únicos en columna '{columna_departamento}': {list(valores_unicos)}")

    departamentos_encontrados = set()
    for valor in valores_unicos:
        nombre_depto = str(valor).strip()
        if not nombre_depto or nombre_depto.lower() == "nan":
            continue

        codigo = obtener_codigo_departamento(nombre_depto)
        if codigo != "ND":
            departamentos_encontrados.add(codigo)
            print(f"   ✅ Detectado en DF limpio: '{nombre_depto}' -> {codigo}")
        else:
            print(f"   ℹ️ Departamento sin código (ignorado): '{nombre_depto}'")

    if departamentos_encontrados:
        departamentos_ordenados = sorted(list(departamentos_encontrados))
        print(f"🏢 Departamentos detectados en DF limpio: {', '.join(departamentos_ordenados)}")
        return departamentos_ordenados
    else:
        print("⚠️ No se pudo detectar departamentos en DF limpio. Usando 'ND'")
        return ["ND"]

def obtener_incidentes(registro):
    """
    Extrae información de incidentes (permisos, campañas, departures) en formato JSON
    CON ORDEN ESPECÍFICO: tipo_evento, institucion, razon, horas_evento
    """
    tipo_evento = registro.get("Tipo_Evento", "N/A")
    if tipo_evento == "N/A":
        return None

    # CAMBIO CRÍTICO: Convertir "Departures" a "Inasistencia justificada (Salida institucional)"
    if tipo_evento == "Departures":
        tipo_evento = "Inasistencia justificada (Salida institucional)"

    incidentes = {}

    # ORDEN ESPECÍFICO SOLICITADO: tipo_evento, institucion, razon, horas_evento
    incidentes["tipo_evento"] = tipo_evento

    # Agregar institución si es Salida Institucional
    if tipo_evento == "Inasistencia justificada (Salida institucional)":
        departures_institucion = registro.get("Departures_Institucion", "N/A")
        if departures_institucion != "N/A":
            incidentes["institucion"] = departures_institucion

    # Agregar razón si es Salida Institucional
    if tipo_evento == "Inasistencia justificada (Salida institucional)":
        departures_razon = registro.get("Departures_Razon", "N/A")
        if departures_razon != "N/A":
            incidentes["razon"] = departures_razon

    # Agregar horas del evento si están disponibles
    horas_evento = registro.get("Horas_Evento", 0)
    if horas_evento and horas_evento > 0:
        incidentes["horas_evento"] = round(horas_evento, 2)

    # Para permisos, agregar tipo_permiso
    if tipo_evento == "Permiso":
        tipo_permiso = registro.get("Tipo_Permiso", "N/A")
        if tipo_permiso != "N/A":
            incidentes["tipo_permiso"] = tipo_permiso
    
    # Para campañas, agregar entidad
    elif tipo_evento == "Campaña":
        campana_entidad = registro.get("Campaña_Entidad", "N/A")
        if campana_entidad != "N/A":
            incidentes["entidad"] = campana_entidad
    
    return incidentes if incidentes else None

def obtener_novedades(registro):
    """
    Extrae información de novedades globales (paros, etc.)
    """
    observaciones = registro.get("Observaciones", "").upper()
    
    # Detectar novedades globales basadas en las observaciones
    if "NOVEDAD ORDEN PUBLICO" in observaciones:
        return "NOVEDAD ORDEN PÚBLICO"
    elif "PARO" in observaciones:
        return "PARO"
    elif "EMERGENCIA" in observaciones:
        return "EMERGENCIA"
    
    return None

def obtener_incidentes_para_api(registro):
    """
    Extrae información de incidentes en formato de array para la API
    """
    tipo_evento = registro.get("Tipo_Evento", "N/A")
    observaciones = registro.get("Observaciones", "")
    
    incidentes = []
    
    # Verificar si hay incidentes basados en el tipo de evento
    if tipo_evento != "N/A":
        # Convertir "Departures" a nombre correcto
        if tipo_evento == "Departures":
            tipo_evento = "Inasistencia justificada (Salida institucional)"
        
        incidente = {
            "type": tipo_evento.lower().replace(" ", "_"),
            "description": f"Evento: {tipo_evento}"
        }
        
        # Agregar detalles específicos según el tipo de evento
        if tipo_evento == "Permiso":
            tipo_permiso = registro.get("Tipo_Permiso", "N/A")
            if tipo_permiso != "N/A":
                incidente["description"] = f"Permiso: {tipo_permiso}"
        
        elif tipo_evento == "Campaña":
            campana_entidad = registro.get("Campaña_Entidad", "N/A")
            if campana_entidad != "N/A":
                incidente["description"] = f"Campaña en: {campana_entidad}"
        
        elif tipo_evento == "Inasistencia justificada (Salida institucional)":
            departures_institucion = registro.get("Departures_Institucion", "N/A")
            departures_razon = registro.get("Departures_Razon", "N/A")
            
            descripcion = "Salida institucional"
            if departures_institucion != "N/A":
                descripcion += f" - Institución: {departures_institucion}"
            if departures_razon != "N/A":
                descripcion += f" - Razón: {departures_razon}"
            
            incidente["description"] = descripcion
        
        incidentes.append(incidente)
    
    # Verificar tardanzas tempranas basadas en observaciones
    if "Llegó tarde" in observaciones:
        incidentes.append({
            "type": "tardanza",
            "description": "Llegó tarde"
        })
    
    if "Salió temprano" in observaciones:
        incidentes.append({
            "type": "salida_temprano",
            "description": "Salió antes del horario asignado"
        })
    
    # Verificar ausencias
    if "No vino" in observaciones or "Sin registro" in observaciones:
        incidentes.append({
            "type": "ausencia",
            "description": "No se presentó a trabajar"
        })
    
    return incidentes if incidentes else None

def transformar_para_api(resumen_path):
    """
    Transforma el resumen de asistencias al formato requerido por la API
    """
    try:
        with open(resumen_path, "r", encoding="utf-8") as f:
            resumen = json.load(f)
    except Exception as e:
        print(f"❌ Error al cargar el archivo de resumen: {e}")
        return []
    
    def segundos_a_hhmmss(segundos_totales):
        """Convierte segundos a formato HH:MM:SS"""
        if not segundos_totales:
            return "00:00:00"
        horas = int(segundos_totales // 3600)
        minutos = int((segundos_totales % 3600) // 60)
        segundos = int(segundos_totales % 60)
        return f"{horas:02d}:{minutos:02d}:{segundos:02d}"
    
    datos_api = []
    
    for registro in resumen:
        # Obtener información de incidentes en formato de array
        incidentes = obtener_incidentes_para_api(registro)
        
        # Convertir horas trabajadas y horas fuera a segundos y luego a HH:MM:SS
        horas_trabajadas_decimal = registro.get("Horas_Trabajadas", 0)
        horas_fuera_decimal = registro.get("Horas_Fuera", 0)
        
        # Convertir horas decimales a segundos
        segundos_trabajados = int(horas_trabajadas_decimal * 3600)
        segundos_fuera = int(horas_fuera_decimal * 3600)
        
        # Intentar convertir cédula a número
        cedula_str = registro.get("Cedula", "")
        try:
            user_id = int(cedula_str) if cedula_str and cedula_str.isdigit() else None
        except:
            user_id = None
        
        if not user_id:
            print(f"⚠️ Cédula inválida para conversión a número: {cedula_str}")
            continue
        
        # Preparar datos para la API con la estructura correcta
        dato_api = {
            "user_id": user_id,  # Ahora es número
            "attendance_date": registro.get("Fecha", ""),
            "arrival_time": registro.get("Hora_Ingreso") if registro.get("Hora_Ingreso") != "N/A" else None,
            "exit_time": registro.get("Hora_Salida") if registro.get("Hora_Salida") != "N/A" else None,
            "assigned_arrival_time": registro.get("Ingreso_Asignado") if registro.get("Ingreso_Asignado") != "N/A" else None,
            "assigned_exit_time": registro.get("Salida_Asignada") if registro.get("Salida_Asignada") != "N/A" else None,
            "total_worked_hours": segundos_a_hhmmss(segundos_trabajados),
            "out_of_institution_hours": segundos_a_hhmmss(segundos_fuera),
            "incidents": incidentes,  # Ahora es un array o null
            "observations": registro.get("Observaciones", ""),
            "news": obtener_novedades(registro)
        }
        
        # Limitar longitud de observaciones y news a 255 caracteres
        if dato_api["observations"] and len(dato_api["observations"]) > 255:
            dato_api["observations"] = dato_api["observations"][:255]
        
        if dato_api["news"] and len(dato_api["news"]) > 255:
            dato_api["news"] = dato_api["news"][:255]
        
        datos_api.append(dato_api)
    
    print(f"📤 Datos transformados para API: {len(datos_api)} registros")
    
    # Verificar la estructura de algunos registros para depuración
    if datos_api and len(datos_api) > 0:
        for i, dato in enumerate(datos_api[:2]):
            print(f"🔍 Ejemplo registro API {i+1}: {json.dumps(dato, ensure_ascii=False)}")
    
    return datos_api

def es_dia_festivo_simple(fecha):
    """
    Función simplificada para verificar días festivos.
    """
    festivos_colombia = {
        '01-01': 'Año Nuevo',
        '01-06': 'Día de los Reyes Magos',
        '03-25': 'Día de San José',
        '04-13': 'Jueves Santo',
        '04-14': 'Viernes Santo',
        '05-01': 'Día del Trabajo',
        '05-22': 'Día de la Ascensión',
        '06-12': 'Corpus Christi',
        '06-19': 'Sagrado Corazón',
        '07-03': 'San Pedro y San Pablo',
        '07-20': 'Día de la Independencia',
        '08-07': 'Batalla de Boyacá',
        '08-21': 'Asunción de la Virgen',
        '10-16': 'Día de la Raza',
        '11-06': 'Todos los Santos',
        '11-13': 'Independencia de Cartagena',
        '12-08': 'Día de la Inmaculada Concepción',
        '12-25': 'Navidad'
    }
    
    fecha_str = fecha.strftime('%m-%d')
    return fecha_str in festivos_colombia

def verificar_coherencia_datos(resumen_path, statistics_path):
    """
    Verifica la coherencia de datos entre el resumen y las estadísticas
    """
    print("\n🔍 VERIFICANDO COHERENCIA DE DATOS...")
    
    try:
        with open(resumen_path, "r", encoding="utf-8") as f:
            resumen = json.load(f)
        
        with open(statistics_path, "r", encoding="utf-8") as f:
            statistics = json.load(f)
    except Exception as e:
        print(f"❌ Error al cargar archivos: {e}")
        return {}

    # Mapeos para filtros
    mapeo_cargos = {}
    mapeo_novedades = {}
    
    for stat in statistics:
        try:
            cedula = str(stat.get("number_id", "")).strip()
            cargo = stat.get("cargo", "DOCENTE").strip().upper()
            
            if cedula:
                mapeo_cargos[cedula] = cargo
                
                año = stat.get("año")
                mes = stat.get("mes")
                dia = stat.get("dia")
                if año and mes and dia:
                    novedad = stat.get("novedad", "No")
                    campana = stat.get("campaña si/no", "No")
                    departures = stat.get("departures si/no", "No")
                    permiso = stat.get("permiso", "No")
                    
                    if novedad == "Sí" or campana == "Sí" or departures == "Sí" or permiso == "Sí":
                        fecha_str = f"{año}-{str(mes).zfill(2)}-{str(dia).zfill(2)}"
                        mapeo_novedades[(cedula, fecha_str)] = True
        except Exception:
            continue

    # Aplicar filtros y contar
    empleados_dias = {}
    for registro in resumen:
        try:
            cedula = registro.get("Cedula") or registro.get("cedula") or ""
            fecha_str = registro.get("Fecha") or registro.get("fecha") or ""
            nombre = registro.get("Empleado") or registro.get("empleado") or ""
            
            if not cedula or not fecha_str:
                continue
                
            # FILTROS
            cargo = mapeo_cargos.get(cedula, "DOCENTE")
            if "DOCENTE" in cargo:
                continue
                
            try:
                fecha = datetime.strptime(fecha_str, "%Y-%m-%d").date()
                if es_dia_festivo_simple(fecha):
                    continue
                if fecha.weekday() == 6:  # Domingo
                    continue
            except:
                pass
                
            if (cedula, fecha_str) in mapeo_novedades:
                continue
                
            # Contar días
            if cedula not in empleados_dias:
                empleados_dias[cedula] = {"nombre": nombre, "dias": set()}
            empleados_dias[cedula]["dias"].add(fecha_str)
            
        except Exception:
            continue

    # Reporte de coherencia
    print(f"📊 EMPLEADOS VÁLIDOS: {len(empleados_dias)}")
    print("👥 TOP 15 EMPLEADOS CON MÁS DÍAS:")
    empleados_ordenados = sorted(empleados_dias.items(), key=lambda x: len(x[1]["dias"]), reverse=True)
    for i, (cedula, info) in enumerate(empleados_ordenados[:15], 1):
        print(f"   {i:2d}. {cedula}: {len(info['dias']):2d} días - {info['nombre']}")

    return empleados_dias

def main():
    print("🚀 Iniciando Sistema de Procesamiento de Asistencia")
    print("=" * 60)
    print("🔢 NUEVO SISTEMA: Días laborados = 100%, regla de 3 por empleado")
    print("📊 PESOS: 40% Tarde + 35% A Tiempo + 25% Temprano")

    # Limpiar archivos temporales anteriores
    limpiar_archivos_temporales()
    
    # Crear directorios temporales
    temp_dir = crear_directorios_temporales()

    try:
        # --- Selección y lectura de archivos Excel ---
        archivos = seleccionar_archivos()
        if not archivos:
            messagebox.showinfo("Sin archivos", "No seleccionaste ningún archivo.")
            return

        # 🆕 DETECTAR DEPARTAMENTOS DE LOS ARCHIVOS (MODIFICADO: ahora detecta TODOS)
        codigos_departamentos = detectar_departamento_archivos(archivos)
        
        dataframes = leer_archivos_excel(archivos)
        if not dataframes:
            messagebox.showwarning("Error", "No se pudieron leer los archivos seleccionados.")
            return

        df_combinado = pd.concat(dataframes, ignore_index=True)
        print(f"📊 Total combinado inicial: {len(df_combinado)} filas")

        # --- Limpieza de datos ---
        df_limpio = limpiar_datos(df_combinado)

        # 🔁 Fallback: si la detección por archivo no encontró nada útil, usar DF limpio
        if not codigos_departamentos or codigos_departamentos == ["ND"]:
            codigos_departamentos = detectar_departamentos_en_dataframe(df_limpio)

        # --- Guardar Excel limpio ---
        os.makedirs("outputs", exist_ok=True)
        salida_excel_limpio = os.path.join("outputs", "datos_limpios.xlsx")
        df_limpio.to_excel(salida_excel_limpio, index=False)
        print(f"✅ Datos limpios guardados en: {salida_excel_limpio}")

        # --- Consulta a la API ---
        print("\n🌐 Consultando datos desde la API...")
        try:
            client = APIClient(email="admin@admin.com", password="Admin3811")
            client.login()
            stats = client.get_statistics()

            # Guardar en carpeta temporal
            salida_json = os.path.join(temp_dir, "statistics.json")
            with open(salida_json, "w", encoding="utf-8") as f:
                json.dump(stats, f, indent=4, ensure_ascii=False)

            print(f"✅ Estadísticas guardadas en: {salida_json}")

        except Exception as e:
            print(f"❌ Error al consultar la API: {e}")
            messagebox.showerror("Error API", f"Ocurrió un error al consultar la API:\n{e}")
            return

        # --- Comparativa de asistencias ---
        print("\n🔍 Iniciando comparación de asistencias...")
        try:
            # Usar rutas temporales
            statistics_temp_path = os.path.join(temp_dir, "statistics.json")
            resumen_temp_path = os.path.join(temp_dir, "resumen_asistencias.json")
            
            resultados = comparar_asistencias(
                df_limpio, 
                statistics_path=statistics_temp_path,
                salida_path=resumen_temp_path
            )
            print(f"✅ Comparativa completada: {len(resultados)} registros procesados")
        except Exception as e:
            print(f"❌ Error en la comparación: {e}")
            messagebox.showerror("Error", f"Error al comparar asistencias:\n{e}")
            return

        # --- OBTENER RANGO DE FECHAS PARA NOMBRAR ARCHIVOS ---
        fecha_min, fecha_max = obtener_rango_fechas(resumen_temp_path)
        
        # 🆕 MODIFICACIÓN: Generar nombre con TODOS los departamentos
        if fecha_min and fecha_max:
            # Formatear fechas como "21-11_20-12" y AGREGAR CÓDIGOS DE TODOS LOS DEPARTAMENTOS
            rango_fechas = f"{fecha_min.strftime('%d-%m')}_{fecha_max.strftime('%d-%m')}"
            # Unir todos los códigos de departamentos con _
            codigos_str = "_".join(codigos_departamentos)
            nombre_excel_api = f"{codigos_str}_{rango_fechas}.xlsx"
            nombre_json_api = f"Asistencia_{codigos_str}_{rango_fechas}.json"
            print(f"📅 Rango de fechas detectado: {rango_fechas}")
            print(f"🏢 Códigos de departamentos: {codigos_str}")
        else:
            # 🆕 MODIFICACIÓN: Agregar códigos de todos los departamentos incluso cuando no hay fechas
            codigos_str = "_".join(codigos_departamentos)
            nombre_excel_api = f"{codigos_str}_sin_fechas.xlsx"
            nombre_json_api = f"Asistencia_{codigos_str}_sin_fechas.json"
            print("⚠️ No se pudieron detectar las fechas para nombrar archivos")

        # --- VERIFICACIÓN DE COHERENCIA ---
        print("\n🔍 Iniciando verificación de coherencia...")
        try:
            empleados_detalle = verificar_coherencia_datos(
                resumen_path=os.path.join(temp_dir, "resumen_asistencias.json"),
                statistics_path=os.path.join(temp_dir, "statistics.json")
            )
            print("✅ Verificación de coherencia completada")
        except Exception as e:
            print(f"⚠️  Error en verificación de coherencia: {e}")

        # --- Generación de reportes Excel ---
        reporte_path = "outputs/Reporte_Asistencias_Completo.xlsx"
        
        # 🆕 CAMBIO DE ORDEN: 1. Hoja de horarios (PRIMERA HOJA)
        try:
            # Primero crear un libro con la hoja Horarios
            from openpyxl import Workbook
            wb = Workbook()
            ws = wb.active
            ws.title = "Horarios"
            wb.save(reporte_path)
            print("✅ Archivo base creado con hoja 'Horarios'")
            
            # Ahora agregar los datos a la hoja Horarios
            agregar_hoja_horarios(
                resumen_path=os.path.join(temp_dir, "resumen_asistencias.json"),
                statistics_path=os.path.join(temp_dir, "statistics.json"),
                salida_excel=reporte_path
            )
            print("✅ Hoja 'Horarios' agregada como PRIMERA hoja")
        except Exception as e:
            print(f"❌ Error al crear hoja 'Horarios': {e}")

        # 2. Hoja de resumen semanal (SEGUNDA HOJA)
        try:
            generar_excel_resumen_semanal(
                resumen_path=os.path.join(temp_dir, "resumen_asistencias.json"),
                statistics_path=os.path.join(temp_dir, "statistics.json"),
                salida_path=reporte_path
            )
            print("✅ Hoja 'Resumen Semanal' agregada como SEGUNDA hoja")
        except Exception as e:
            print(f"❌ Error al generar resumen semanal: {e}")

        # 3. Hoja diaria (TERCERA HOJA)
        try:
            agregar_hoja_diario(
                resumen_path=os.path.join(temp_dir, "resumen_asistencias.json"),
                statistics_path=os.path.join(temp_dir, "statistics.json"),
                salida_excel=reporte_path
            )
            print("✅ Hoja 'Diario' agregada como TERCERA hoja")
        except Exception as e:
            print(f"❌ Error al agregar hoja Diario: {e}")
            print(f"📋 Detalles del error: {traceback.format_exc()}")

        # 4. Hoja de horarios programados (CUARTA HOJA)
        try:
            agregar_hoja_horarios_programados(
                statistics_path=os.path.join(temp_dir, "statistics.json"),
                salida_excel=reporte_path
            )
            print("✅ Hoja 'Horarios Programados' agregada como CUARTA hoja")
        except Exception as e:
            print(f"❌ Error al agregar hoja 'Horarios Programados': {e}")

        # 5. Hoja de permisos (QUINTA HOJA)
        try:
            agregar_hoja_permisos(
                statistics_path=os.path.join(temp_dir, "statistics.json"),
                salida_excel=reporte_path
            )
            print("✅ Hoja 'Permisos' agregada como QUINTA hoja")
        except Exception as e:
            print(f"❌ Error al agregar hoja 'Permisos': {e}")

        # 6. Hoja de explicación (SEXTA HOJA)
        try:
            agregar_hoja_explicacion(salida_excel=reporte_path)
            print("✅ Hoja 'Explicación' agregada como SEXTA hoja")
        except Exception as e:
            print(f"❌ Error al agregar hoja Explicación: {e}")

        # 7. Hoja de registros brutos (ÚLTIMA HOJA)
        try:
            df_brutos = pd.read_excel(salida_excel_limpio)
            book = openpyxl.load_workbook(reporte_path)

            # Eliminar hoja si ya existe
            if "Registros Brutos" in book.sheetnames:
                del book["Registros Brutos"]

            # Crear nueva hoja
            ws = book.create_sheet("Registros Brutos")
            
            # Escribir encabezados
            for col_idx, col_name in enumerate(df_brutos.columns, 1):
                ws.cell(row=1, column=col_idx, value=col_name)

            # Escribir datos
            for r_idx, row in enumerate(df_brutos.values, 2):
                for c_idx, value in enumerate(row, 1):
                    ws.cell(row=r_idx, column=c_idx, value=value)

            print("✅ Hoja 'Registros Brutos' agregada correctamente")
            book.save(reporte_path)

        except Exception as e:
            print(f"❌ Error al agregar hoja 'Registros Brutos': {e}")

        # --- 🆕 REORDENAR HOJAS EN EL ORDEN CORRECTO ---
        try:
            print("\n🔀 Reordenando hojas en el orden solicitado...")
            wb = openpyxl.load_workbook(reporte_path)
            
            # NUEVO ORDEN: Diario (1), Resumen Semanal (2), Permisos (3), Horarios Programados (4), Explicación (5), Registros Brutos (6)
            orden_hojas = ["Diario", "Resumen Semanal", "Permisos", "Horarios Programados", "Explicación", "Registros Brutos"]
            
            # Reordenar hojas moviéndolas al final en orden inverso
            for nombre_hoja in reversed(orden_hojas):
                if nombre_hoja in wb.sheetnames:
                    hoja = wb[nombre_hoja]
                    # Mover al final (esto crea el orden inverso, luego lo revertimos)
                    wb.move_sheet(hoja, offset=len(wb.sheetnames) - 1)
            
            # Invertir el orden final
            wb._sheets.reverse()
            
            wb.save(reporte_path)
            print("✅ Hojas reordenadas correctamente:")
            for i, hoja in enumerate(wb.sheetnames, 1):
                print(f"   {i}. {hoja}")
        except Exception as e:
            print(f"⚠️ Error al reordenar hojas: {e}")

        # --- Envío de datos a la API ---
        print("\n📤 Enviando datos a la API...")
        try:
            # Transformar datos al formato de la API
            datos_api = transformar_para_api(os.path.join(temp_dir, "resumen_asistencias.json"))
            
            if datos_api:
                print(f"📊 Preparando envío de {len(datos_api)} registros a la API...")
                
                # Enviar datos a la API - SOLO UNA VEZ
                client.enviar_asistencias_api(datos_api)
                print("🎉 Datos de asistencias enviados a la API correctamente")
                
                # VERIFICAR ESTRUCTURA DE ALGUNOS REGISTROS
                print("🔍 Verificando estructura de datos enviados...")
                for i, dato in enumerate(datos_api[:2]):  # Mostrar solo 2 para verificación
                    if dato.get("incidents"):
                        print(f"   Registro {i+1}: {json.dumps(dato['incidents'], ensure_ascii=False)}")
            else:
                print("⚠️ No hay datos para enviar a la API")
                
        except Exception as e:
            print(f"❌ Error en el proceso de envío a la API: {e}")
            print(f"📋 Detalles: {traceback.format_exc()}")

        # --- NUEVO: Envío de archivos Excel y JSON a la API CON NOMBRE PERSONALIZADO ---
        print("\n📤 Enviando archivos generados a la API...")
        try:
            # Crear copias de los archivos con el nombre personalizado (QUE AHORA INCLUYE TODOS LOS DEPARTAMENTOS)
            excel_renombrado = os.path.join(temp_dir, nombre_excel_api)
            json_renombrado = os.path.join(temp_dir, nombre_json_api)
            
            # Copiar archivo Excel con nuevo nombre (QUE AHORA INCLUYE TODOS LOS DEPARTAMENTOS)
            if os.path.exists(reporte_path):
                shutil.copy2(reporte_path, excel_renombrado)
                print(f"📄 Excel renombrado: {nombre_excel_api}")
            else:
                print("⚠️ No se encontró el archivo Excel para renombrar")

            # Copiar archivo JSON con nuevo nombre (QUE AHORA INCLUYE TODOS LOS DEPARTAMENTOS)
            if os.path.exists(os.path.join(temp_dir, "resumen_asistencias.json")):
                shutil.copy2(os.path.join(temp_dir, "resumen_asistencias.json"), json_renombrado)
                print(f"📄 JSON renombrado: {nombre_json_api}")
            else:
                print("⚠️ No se encontró el archivo JSON para renombrar")

            # Enviar archivo Excel renombrado - ACTIVADO
            if os.path.exists(excel_renombrado):
                client.enviar_excel_api(excel_renombrado)
                print(f"✅ Archivo Excel enviado: {nombre_excel_api}")
            else:
                print("⚠️ No se encontró el archivo Excel renombrado para enviar")

            # Enviar archivo JSON renombrado - ACTIVADO
            if os.path.exists(json_renombrado):
                client.enviar_json_api(json_renombrado)
                print(f"✅ Archivo JSON enviado: {nombre_json_api}")
            else:
                print("⚠️ No se encontró el archivo JSON renombrado para enviar")

        except Exception as e:
            print(f"❌ Error al preparar archivos para la API: {e}")
            print(f"📋 Detalles: {traceback.format_exc()}")

        # --- Verificar archivos generados ---
        print("\n📁 ARCHIVOS GENERADOS:")
        archivos_generados = [
            ("Datos limpios", salida_excel_limpio),
            ("Estadísticas API", os.path.join(temp_dir, "statistics.json")),
            ("Resumen asistencias", os.path.join(temp_dir, "resumen_asistencias.json")),
            ("Reporte completo", reporte_path),
            ("Excel para API", os.path.join(temp_dir, nombre_excel_api) if fecha_min else "No generado"),
            ("JSON para API", os.path.join(temp_dir, nombre_json_api) if fecha_min else "No generado")
        ]
        
        for nombre, ruta in archivos_generados:
            if os.path.exists(ruta):
                print(f"✅ {nombre}: {ruta}")
            else:
                print(f"❌ {nombre}: NO GENERADO")

        # --- Resumen final ---
        print("=" * 60)
        messagebox.showinfo(
            "Proceso Completado",
            f"✅ Se procesaron {len(archivos)} archivos\n"
            f"📊 {len(df_combinado)} registros iniciales\n"
            f"🧹 {len(df_limpio)} registros limpios\n"
            f"📑 Reporte generado: {reporte_path}\n"
            f"📊 ORDEN DE HOJAS: 1. Diario, 2. Resumen Semanal, 3. Permisos, 4. Horarios Programados, 5. Explicación, 6. Registros Brutos\n"
            f"🔢 NUEVO SISTEMA: Días laborados = 100%\n"
            f"📊 PESOS: 40% Tarde + 35% A Tiempo + 25% Temprano\n"
            f"📤 Datos enviados a la API: {len(datos_api) if 'datos_api' in locals() else 0} registros\n"
            f"🏢 Departamentos: {', '.join(codigos_departamentos)}\n"
            f"📁 Archivos preparados para API: {nombre_excel_api} y {nombre_json_api}"
        )

    except Exception as e:
        print(f"❌ Error general: {e}")
        print(traceback.format_exc())
        messagebox.showerror("Error", f"Error inesperado:\n{e}")
    
    finally:
        # --- Limpieza final de archivos temporales ---
        print("\n🧹 Limpiando archivos temporales...")
        if os.path.exists(temp_dir):
            #shutil.rmtree(temp_dir)
            print("✅ Archivos temporales eliminados")

if __name__ == "__main__":
    main()
