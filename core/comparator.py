# core/comparator.py (VERSIÓN CORREGIDA - SIN DUPLICADOS)
import os
import json
import pandas as pd
from datetime import datetime, timedelta

def safe_int(value):
    try:
        return int(value)
    except (ValueError, TypeError):
        return None

def _parse_date_from_stats(persona):
    año = safe_int(persona.get("año"))
    mes = safe_int(persona.get("mes"))
    dia = safe_int(persona.get("dia"))
    if año and mes and dia:
        try:
            return datetime(año, mes, dia).date()
        except Exception:
            return None
    return None

def _parse_time(value):
    """Convierte texto HH:MM:SS a datetime.time o None."""
    if not value or value in ["N/A", ""]:
        return None
    try:
        return datetime.strptime(value, "%H:%M:%S").time()
    except Exception:
        return None

def _duracion_horas(inicio, fin):
    """Calcula la duración en horas entre dos objetos time."""
    if not inicio or not fin:
        return 0.0
    inicio_dt = datetime.combine(datetime.today(), inicio)
    fin_dt = datetime.combine(datetime.today(), fin)
    return round((fin_dt - inicio_dt).total_seconds() / 3600, 2)

def _calcular_horas_fuera(eventos):
    """
    Calcula las horas fuera dentro de un mismo día
    """
    if len(eventos) < 4:
        return 0.0

    col_tiempo = 'tiempo'
    col_lector = 'nombre_de_lector'
    
    eventos = eventos.sort_values(by=col_tiempo)
    horas_fuera = 0.0

    for i in range(1, len(eventos) - 1):
        anterior = eventos.iloc[i - 1]
        actual = eventos.iloc[i]

        lector_anterior = str(anterior[col_lector]).upper() if pd.notna(anterior[col_lector]) else ""
        lector_actual = str(actual[col_lector]).upper() if pd.notna(actual[col_lector]) else ""
        
        if "SAL" in lector_anterior and "ENT" in lector_actual:
            if col_tiempo in anterior and col_tiempo in actual:
                try:
                    delta = (actual[col_tiempo] - anterior[col_tiempo]).total_seconds() / 3600
                    if 0 < delta < 6:
                        horas_fuera += round(delta, 2)
                except Exception:
                    continue

    return round(horas_fuera, 2)

def _detectar_novedad_orden_publico(empleados_data, fecha):
    """
    Detecta si el 85% del personal llegó tarde o salió temprano en un día específico
    """
    total_empleados = 0
    empleados_con_problema = 0
    
    for cedula, info in empleados_data.items():
        if fecha in info["fechas_analizadas"]:
            total_empleados += 1
            if (fecha in info["fechas_tarde"] or 
                fecha in info["fechas_salida_temprano"]):
                empleados_con_problema += 1
    
    if total_empleados == 0:
        return False
    
    porcentaje_problema = (empleados_con_problema / total_empleados) * 100
    print(f"📊 Análisis orden público {fecha}: {porcentaje_problema:.1f}% con problemas ({empleados_con_problema}/{total_empleados})")
    
    return porcentaje_problema >= 85

def calcular_puntualidad(hora_real, hora_asignada):
    """
    Calcula la puntualidad de una hora real vs una hora asignada.
    Retorna:
        - "tarde"
        - "a_tiempo"
        - "temprano"
        - "error" si no se puede calcular
    """
    try:
        if not hora_real or not hora_asignada:
            return "error"
            
        if hora_real == "N/A" or hora_asignada == "N/A":
            return "error"
            
        h_real = datetime.strptime(hora_real, "%H:%M:%S")
        h_asig = datetime.strptime(hora_asignada, "%H:%M:%S")
        
        diferencia = h_real - h_asig
        diferencia_minutos = diferencia.total_seconds() / 60
        
        # Entrada
        if diferencia_minutos > 0:
            # Llegó tarde
            if diferencia_minutos <= 10:
                return "a_tiempo"
            else:
                return "tarde"
        else:
            # Llegó temprano o exacto
            if abs(diferencia_minutos) <= 5:
                return "a_tiempo"
            else:
                return "temprano"
    except Exception as e:
        print(f"❌ Error calculando puntualidad: {hora_real} vs {hora_asignada} - {e}")
        return "error"

def calcular_puntualidad_salida(hora_real, hora_asignada):
    """
    Calcula la puntualidad de una HORA DE SALIDA real vs una hora asignada.
    Retorna:
        - "tarde"
        - "a_tiempo"
        - "temprano"
        - "error" si no se puede calcular
    """
    try:
        if not hora_real or not hora_asignada:
            return "error"
            
        if hora_real == "N/A" or hora_asignada == "N/A":
            return "error"
            
        h_real = datetime.strptime(hora_real, "%H:%M:%S")
        h_asig = datetime.strptime(hora_asignada, "%H:%M:%S")
        
        diferencia = h_real - h_asig
        diferencia_minutos = diferencia.total_seconds() / 60
        
        # Salida
        if diferencia_minutos > 0:
            # Salió tarde
            if diferencia_minutos <= 10:
                return "a_tiempo"
            else:
                return "tarde"
        else:
            # Salió temprano o exacto
            if -10 <= diferencia_minutos <= 5:
                return "a_tiempo"
            else:
                return "temprano"
                
    except Exception as e:
        print(f"❌ Error calculando puntualidad: {hora_real} vs {hora_asignada} - {e}")
        return "error"

def validar_registro_puntualidad(registro):
    """
    Valida que un registro tenga toda la información necesaria para análisis de puntualidad
    """
    try:
        fecha_str = registro.get("Fecha") or registro.get("fecha") or ""
        if not fecha_str or fecha_str == "N/A":
            return False, "Fecha inválida"
            
        ingreso_real = registro.get("Hora_Ingreso") or registro.get("Hora Ingreso") or "N/A"
        ingreso_asignado = registro.get("Ingreso_Asignado") or registro.get("Ingreso Asignado") or "N/A"
        
        if ingreso_real == "N/A" or ingreso_asignado == "N/A":
            return False, "Horarios incompletos"
            
        try:
            datetime.strptime(ingreso_real, "%H:%M:%S")
            datetime.strptime(ingreso_asignado, "%H:%M:%S")
        except ValueError:
            return False, "Formato de hora inválido"
            
        return True, "OK"
        
    except Exception as e:
        return False, f"Error validación: {e}"

def comparar_asistencias(
    df_limpio,
    statistics_path="temp/statistics.json",
    salida_path="temp/resumen_asistencias.json"
):
    print("\n🔍 Iniciando comparativa entre Excel limpio y statistics.json...")

    try:
        with open(statistics_path, "r", encoding="utf-8") as f:
            stats = json.load(f)
            
            # PARCHE: Si la API devuelve un diccionario en vez de una lista
            if isinstance(stats, dict):
                # Extraemos la lista real si viene encapsulada en 'data' u otra llave
                for key, value in stats.items():
                    if isinstance(value, list):
                        stats = value
                        break
                # Si sigue siendo un dict (error de api sin lista adentro)
                if isinstance(stats, dict):
                    print(f"❌ Error: La API devolvió un formato inesperado: {stats}")
                    return []
                    
    except FileNotFoundError:
        print("❌ No se encontró el archivo statistics.json.")
        return []
    except json.JSONDecodeError:
        print("❌ Error al leer el JSON de estadísticas.")
        return []

    resultados = []
    df = df_limpio.copy()
    
    # Normalizar columnas mínimas necesarias desde el Excel limpio
    # Se asume que ya viene con 'tiempo', 'id', 'nombre_de_lector', 'tarjeta', etc.
    df["tiempo"] = pd.to_datetime(df["tiempo"], errors="coerce")
    df["id"] = df["id"].astype(str).str.strip()

    fecha_min_excel = df["tiempo"].min().date()
    fecha_max_excel = df["tiempo"].max().date()
    print(f"📊 Rango de fechas EXCLUSIVO de archivos Excel: {fecha_min_excel} a {fecha_max_excel}")

    # Agrupar estadísticas por empleado (number_id)
    empleados = {}
    for persona in stats:
        number_id = str(persona.get("number_id", "")).strip()
        if number_id:
            empleados.setdefault(number_id, []).append(persona)

    # Mapeo de carnets por fecha
    carnet_por_fecha_map = {}
    carnet_mapping_count = 0

    for persona in stats:
        carnet = persona.get("carnet", "No")
        carnet_number = persona.get("carnet_number", "N/A")
        number_id = str(persona.get("number_id", "")).strip()
        
        if carnet == "Sí" and carnet_number != "N/A" and number_id:
            año = persona.get("año")
            mes = persona.get("mes")
            dia = persona.get("dia")
            
            if año and mes and dia:
                try:
                    fecha_str = f"{año}-{str(mes).zfill(2)}-{str(dia).zfill(2)}"
                    clave = (carnet_number, fecha_str)
                    carnet_por_fecha_map[clave] = {
                        'number_id': number_id,
                        'nombre': persona.get("nombre", ""),
                        'fecha': fecha_str
                    }
                    carnet_mapping_count += 1
                except Exception as e:
                    continue

    print(f"📋 Mapeo de carnets por fecha creado: {carnet_mapping_count} entradas")

    # Aplicar mapeo de carnets al DataFrame de registros
    if 'tarjeta' in df.columns:
        tarjeta_col = 'tarjeta'
        mapeos_aplicados = 0
        # Vectorized mapping for 'id' replacement based on 'tarjeta' and 'tiempo'
        df['temp_fecha_str'] = df["tiempo"].dt.strftime("%Y-%m-%d").fillna("")
        df['temp_tarjeta'] = df[tarjeta_col].fillna("").astype(str).str.strip()
        
        # Preparar un diccionario de mapeo plano
        mapping_dict = {k: v['number_id'] for k, v in carnet_por_fecha_map.items()}
        
        # Crear la clave temporal y mapear
        df['temp_key'] = list(zip(df['temp_tarjeta'], df['temp_fecha_str']))
        df['nuevo_id'] = df['temp_key'].map(mapping_dict)
        
        mask = df['nuevo_id'].notna() & (df['nuevo_id'] != df['id'])
        mapeos_aplicados = mask.sum()
        
        if mapeos_aplicados > 0:
            df.loc[mask, 'id'] = df.loc[mask, 'nuevo_id']
            
        df.drop(columns=['temp_fecha_str', 'temp_tarjeta', 'temp_key', 'nuevo_id'], inplace=True)

        print(f"✅ Mapeos de carnets aplicados: {mapeos_aplicados}")

    empleados_data = {}
    fechas_analizadas = set()

    # Filtrar solo empleados que tienen registros en Excel
    empleados_con_registros = {}
    for number_id, registros_persona in empleados.items():
        registros_empleado = df[df["id"] == number_id].copy()
        if not registros_empleado.empty:
            empleados_con_registros[number_id] = registros_persona

    print(f"👥 Empleados con registros en Excel: {len(empleados_con_registros)}")

    total_registros_procesados = 0
    registros_validos_puntualidad = 0
    registros_invalidos_puntualidad = 0

    # Diccionario para evitar duplicados por empleado/fecha
    registros_por_empleado_fecha = {}

    for number_id, registros_persona in empleados_con_registros.items():
        nombre = registros_persona[0].get("nombre", "").strip()
        registros_empleado = df[df["id"] == number_id].copy()

        fecha_min = fecha_min_excel
        fecha_max = fecha_max_excel

        todas_las_fechas = [
            fecha_min + timedelta(days=i)
            for i in range((fecha_max - fecha_min).days + 1)
        ]

        mapa_por_fecha = {}
        for persona in registros_persona:
            f = _parse_date_from_stats(persona)
            if f and fecha_min <= f <= fecha_max:
                mapa_por_fecha[f] = persona

        if number_id not in empleados_data:
            empleados_data[number_id] = {
                "nombre": nombre,
                "fechas_temprano": set(),
                "fechas_atiempo": set(),
                "fechas_tarde": set(),
                "fechas_salida_temprano": set(),
                "fechas_analizadas": set(),
                "horario_habitual_entrada": None,
                "horario_habitual_salida": None,
                "registros_totales": 0,
                "registros_validos": 0
            }

        for fecha in todas_las_fechas:
            fecha_str = fecha.strftime("%Y-%m-%d")
            
            # Clave única por empleado/fecha
            clave_unica = f"{number_id}_{fecha_str}"
            if clave_unica in registros_por_empleado_fecha:
                continue  # Ya procesado
            
            registros_por_empleado_fecha[clave_unica] = True

            persona = mapa_por_fecha.get(fecha, {})
            horario_flag = persona.get("horario", "No")
            horario_entrada = persona.get("horario_entrada", "N/A")
            horario_salida = persona.get("horario_salida", "N/A")
            permiso = persona.get("permiso", "No")
            tipo_permiso = persona.get("tipo_permiso", "N/A")
            campaña = persona.get("campaña si/no", "No")
            departures = persona.get("departures si/no", "No")
            url_permiso = persona.get("permiso_url", "N/A")

            # Información de almuerzo (si existe en statistics.json)
            almuerzo_inicio = _parse_time(persona.get("horario_almuerzo_inicio"))
            almuerzo_fin = _parse_time(persona.get("horario_almuerzo_fin"))
            tiempo_almuerzo = _duracion_horas(almuerzo_inicio, almuerzo_fin)

            evento_tipo = None
            horas_evento = 0.0

            if permiso == "Sí":
                evento_tipo = "Permiso"
                evento_inicio = _parse_time(persona.get("hora_inicio"))
                evento_fin = _parse_time(persona.get("hora_fin"))
                horas_evento = _duracion_horas(evento_inicio, evento_fin)
            elif campaña == "Sí":
                evento_tipo = "Campaña"
                evento_inicio = _parse_time(persona.get("campaña-hora-inicio"))
                evento_fin = _parse_time(persona.get("campaña-hora-fin"))
                horas_evento = _duracion_horas(evento_inicio, evento_fin)
            elif departures == "Sí":
                evento_tipo = "Departures"  # Mantenemos "Departures" para que main.py lo transforme
                evento_inicio = _parse_time(persona.get("horaInicio"))
                evento_fin = _parse_time(persona.get("horaFin"))
                horas_evento = _duracion_horas(evento_inicio, evento_fin)
            else:
                evento_inicio = None
                evento_fin = None

            # Registros de este empleado en esta fecha
            registros_dia = registros_empleado[
                registros_empleado["tiempo"].dt.date == fecha
            ].copy()

            # Sin horario y sin registros: no se analiza
            if horario_flag == "No" and registros_dia.empty:
                resultados.append({
                    "Empleado": nombre,
                    "Cedula": number_id,
                    "Fecha": fecha_str,
                    "Ingreso_Asignado": "N/A",
                    "Salida_Asignada": "N/A",
                    "Hora_Ingreso": "N/A",
                    "Hora_Salida": "N/A",
                    "Horas_Netas": 0,
                    "Horas_Fuera": 0,
                    "Horas_Trabajadas": 0,
                    "Observaciones": "Sin horario ni registros",
                    "Tipo_Evento": evento_tipo or "N/A",
                    "Tipo_Permiso": tipo_permiso if evento_tipo == "Permiso" else "N/A",
                    "Hora_Evento_Inicio": persona.get("hora_inicio", "N/A") if evento_tipo == "Permiso" else persona.get("campaña-hora-inicio", "N/A"),
                    "Hora_Evento_Fin": persona.get("hora_fin", "N/A") if evento_tipo == "Permiso" else persona.get("campaña-hora-fin", "N/A"),
                    "Horas_Evento": horas_evento,
                    "Vista_Previa_Permiso": url_permiso,
                    "Campaña_Entidad": persona.get("campaña(nombre entidad)", "N/A") if evento_tipo == "Campaña" else "N/A",
                    "Departures_Institucion": persona.get("institución", "N/A") if evento_tipo == "Departures" else "N/A",
                    "Departures_Razon": persona.get("razón", "N/A") if evento_tipo == "Departures" else "N/A"
                })
                continue

            # Tiene horario pero sin registro
            if horario_flag == "Sí" and registros_dia.empty:
                observacion = "No vino"
                if permiso == "Sí":
                    observacion = f"Asistencia justificada (permiso: {tipo_permiso})"
                elif campaña == "Sí":
                    observacion = "Asistencia justificada (campaña)"
                elif departures == "Sí":
                    observacion = "Asistencia justificada (Departures)"

                resultados.append({
                    "Empleado": nombre,
                    "Cedula": number_id,
                    "Fecha": fecha_str,
                    "Ingreso_Asignado": horario_entrada,
                    "Salida_Asignada": horario_salida,
                    "Hora_Ingreso": "N/A",
                    "Hora_Salida": "N/A",
                    "Horas_Netas": 0,
                    "Horas_Fuera": 0,
                    "Horas_Trabajadas": 0,
                    "Observaciones": observacion,
                    "Tipo_Evento": evento_tipo or "N/A",
                    "Tipo_Permiso": tipo_permiso if evento_tipo == "Permiso" else "N/A",
                    "Hora_Evento_Inicio": persona.get("hora_inicio", "N/A") if evento_tipo == "Permiso" else persona.get("campaña-hora-inicio", "N/A"),
                    "Hora_Evento_Fin": persona.get("hora_fin", "N/A") if evento_tipo == "Permiso" else persona.get("campaña-hora-fin", "N/A"),
                    "Horas_Evento": horas_evento,
                    "Vista_Previa_Permiso": url_permiso,
                    "Campaña_Entidad": persona.get("campaña(nombre entidad)", "N/A") if evento_tipo == "Campaña" else "N/A",
                    "Departures_Institucion": persona.get("institución", "N/A") if evento_tipo == "Departures" else "N/A",
                    "Departures_Razon": persona.get("razón", "N/A") if evento_tipo == "Departures" else "N/A"
                })
                continue

            # Hay registros en ese día
            if not registros_dia.empty:
                eventos = registros_dia.sort_values(by="tiempo")
                hora_ingreso = eventos["tiempo"].min()
                hora_salida = eventos["tiempo"].max()

                # Ajuste por permisos (si el permiso empieza antes de la entrada o termina después de la salida)
                if evento_tipo == "Permiso" and evento_inicio and evento_fin:
                    fecha_evento = hora_ingreso.date()
                    permiso_inicio = datetime.combine(fecha_evento, evento_inicio)
                    permiso_fin = datetime.combine(fecha_evento, evento_fin)
                    
                    # Ajustar entrada
                    if permiso_inicio < hora_ingreso:
                        hora_ingreso = permiso_inicio
                        print(f"🔧 Ajuste entrada por permiso: {number_id} - {fecha_str} - Nueva entrada: {hora_ingreso.time()}")
                    
                    # Ajustar salida
                    if permiso_fin > hora_salida:
                        hora_salida = permiso_fin
                        print(f"🔧 Ajuste salida por permiso: {number_id} - {fecha_str} - Nueva salida: {hora_salida.time()}")
                # --- FIN MODIFICACIÓN ---

                horas_netas = round((hora_salida - hora_ingreso).total_seconds() / 3600, 2)
                horas_fuera = _calcular_horas_fuera(eventos)

                # NUEVA LÓGICA DE HORAS TRABAJADAS / NETAS CON ALMUERZO
                # - horas_netas: diferencia entre salida e ingreso (jornada bruta)
                # - horas_fuera: tiempo total fuera de la institución (todas las salidas/entradas)
                # - tiempo_almuerzo: duración de la hora de almuerzo programada (si existe)
                #
                # Reglas:
                # 1) Si la persona salió y entró solo DURANTE el almuerzo:
                #       horas_trabajadas = horas_netas - horas_fuera
                # 2) Si la persona NO salió pero SÍ tiene almuerzo:
                #       horas_trabajadas = horas_netas - tiempo_almuerzo
                # 3) Si la persona tiene tiempo fuera ADICIONAL al almuerzo:
                #       horas_trabajadas = horas_netas - horas_fuera - tiempo_almuerzo
                # 4) Si no hay almuerzo configurado:
                #       horas_trabajadas = horas_netas - horas_fuera
                #
                # Aproximación:
                #   - Si horas_fuera <= tiempo_almuerzo → asumimos que es solo dentro del almuerzo.
                #   - Si horas_fuera == 0 y tiempo_almuerzo > 0 → no salió pero tuvo almuerzo.
                #   - Si horas_fuera  > tiempo_almuerzo → hay salidas adicionales al almuerzo.

                # Normalizar valores
                if horas_fuera is None:
                    horas_fuera = 0.0
                if tiempo_almuerzo is None:
                    tiempo_almuerzo = 0.0

                if horas_fuera < 0:
                    horas_fuera = 0.0
                if tiempo_almuerzo < 0:
                    tiempo_almuerzo = 0.0

                if tiempo_almuerzo > 0:
                    if horas_fuera > 0:
                        if horas_fuera <= tiempo_almuerzo + 0.01:
                            # Todo el tiempo fuera cabe dentro del almuerzo
                            horas_trabajadas = round(horas_netas - horas_fuera, 2)
                        else:
                            # Tiene tiempo fuera adicional al almuerzo
                            horas_trabajadas = round(horas_netas - horas_fuera - tiempo_almuerzo, 2)
                    else:
                        # No tiene salidas registradas pero sí almuerzo
                        horas_trabajadas = round(horas_netas - tiempo_almuerzo, 2)
                else:
                    # Sin almuerzo: solo se descuenta el tiempo fuera
                    horas_trabajadas = round(horas_netas - horas_fuera, 2)

                if horas_trabajadas < 0:
                    horas_trabajadas = 0.0

                es_valido, mensaje_validacion = validar_registro_puntualidad({
                    "Fecha": fecha_str,
                    "Hora_Ingreso": hora_ingreso.strftime("%H:%M:%S"),
                    "Ingreso_Asignado": horario_entrada
                })
                
                total_registros_procesados += 1
                
                if es_valido:
                    registros_validos_puntualidad += 1
                    punt_entrada = calcular_puntualidad(
                        hora_ingreso.strftime("%H:%M:%S"),
                        horario_entrada
                    )
                    punt_salida = calcular_puntualidad_salida(
                        hora_salida.strftime("%H:%M:%S"),
                        horario_salida
                    )
                    
                    if punt_entrada == "tarde":
                        empleados_data[number_id]["fechas_tarde"].add(fecha)
                    elif punt_entrada == "a_tiempo":
                        empleados_data[number_id]["fechas_atiempo"].add(fecha)
                    elif punt_entrada == "temprano":
                        empleados_data[number_id]["fechas_temprano"].add(fecha)
                        
                    if punt_salida == "temprano":
                        empleados_data[number_id]["fechas_salida_temprano"].add(fecha)
                        
                    empleados_data[number_id]["fechas_analizadas"].add(fecha)
                    empleados_data[number_id]["registros_validos"] += 1
                    fechas_analizadas.add(fecha)
                else:
                    registros_invalidos_puntualidad += 1

                empleados_data[number_id]["registros_totales"] += 1

                # Horario habitual si el día fue "normal"
                if es_valido and punt_entrada != "tarde":
                    if empleados_data[number_id]["horario_habitual_entrada"] is None:
                        empleados_data[number_id]["horario_habitual_entrada"] = hora_ingreso.time()
                    if empleados_data[number_id]["horario_habitual_salida"] is None:
                        empleados_data[number_id]["horario_habitual_salida"] = hora_salida.time()

                observacion = "Asistencia normal"
                if evento_tipo == "Permiso":
                    observacion = "Permiso"
                elif evento_tipo == "Campaña":
                    observacion = "Campaña"
                elif departures == "Sí":
                    observacion = "Departures"

                resultados.append({
                    "Empleado": nombre,
                    "Cedula": number_id,
                    "Fecha": fecha_str,
                    "Ingreso_Asignado": horario_entrada,
                    "Salida_Asignada": horario_salida,
                    "Hora_Ingreso": hora_ingreso.strftime("%H:%M:%S"),
                    "Hora_Salida": hora_salida.strftime("%H:%M:%S"),
                    "Horas_Netas": horas_netas,
                    "Horas_Fuera": horas_fuera,
                    "Horas_Trabajadas": horas_trabajadas,
                    "Observaciones": observacion,
                    "Tipo_Evento": evento_tipo or "N/A",
                    "Tipo_Permiso": tipo_permiso if evento_tipo == "Permiso" else "N/A",
                    "Hora_Evento_Inicio": persona.get("hora_inicio", "N/A") if evento_tipo == "Permiso" else persona.get("campaña-hora-inicio", "N/A"),
                    "Hora_Evento_Fin": persona.get("hora_fin", "N/A") if evento_tipo == "Permiso" else persona.get("campaña-hora-fin", "N/A"),
                    "Horas_Evento": horas_evento,
                    "Vista_Previa_Permiso": url_permiso,
                    "Puntualidad_Valido": es_valido,
                    "Puntualidad_Mensaje": mensaje_validacion if not es_valido else "OK",
                    "Campaña_Entidad": persona.get("campaña(nombre entidad)", "N/A") if evento_tipo == "Campaña" else "N/A",
                    "Departures_Institucion": persona.get("institución", "N/A") if evento_tipo == "Departures" else "N/A",
                    "Departures_Razon": persona.get("razón", "N/A") if evento_tipo == "Departures" else "N/A"
                })

    # Aplicar ajuste por novedad de orden público
    print("\n🔍 Analizando posibles novedades de orden público...")
    for fecha in fechas_analizadas:
        if _detectar_novedad_orden_publico(empleados_data, fecha):
            print(f"⚠️  Detectada NOVEDAD ORDEN PÚBLICO para fecha: {fecha}")
            
            for resultado in resultados:
                if resultado["Fecha"] == fecha.strftime("%Y-%m-%d"):
                    cedula = resultado["Cedula"]
                    if cedula in empleados_data:
                        emp_data = empleados_data[cedula]
                        
                        if (emp_data["horario_habitual_entrada"] and 
                            emp_data["horario_habitual_salida"]):
                            
                            resultado["Ingreso_Asignado"] = emp_data["horario_habitual_entrada"].strftime("%H:%M:%S")
                            resultado["Salida_Asignada"] = emp_data["horario_habitual_salida"].strftime("%H:%M:%S")
                            
                            if resultado["Hora_Ingreso"] != "N/A":
                                try:
                                    hora_real = datetime.strptime(resultado["Hora_Ingreso"], "%H:%M:%S").time()
                                    hora_ajustada = emp_data["horario_habitual_entrada"]
                                    
                                    if hora_real > hora_ajustada:
                                        resultado["Observaciones"] = "NOVEDAD ORDEN PUBLICO - Llegó tarde"
                                    else:
                                        resultado["Observaciones"] = "NOVEDAD ORDEN PUBLICO - Cumplió horario"
                                except:
                                    resultado["Observaciones"] = "NOVEDAD ORDEN PUBLICO"
                            else:
                                resultado["Observaciones"] = "NOVEDAD ORDEN PUBLICO"

    print(f"\n📊 DIAGNÓSTICO FINAL DE PUNTUALIDAD:")
    print(f"   Total registros procesados: {total_registros_procesados}")
    print(f"   Registros válidos para puntualidad: {registros_validos_puntualidad}")
    print(f"   Registros inválidos para puntualidad: {registros_invalidos_puntualidad}")
    if total_registros_procesados > 0:
        tasa_validez = registros_validos_puntualidad / total_registros_procesados * 100
        print(f"   Tasa de validez: {tasa_validez:.1f}%")
    else:
        print(f"   Tasa de validez: N/A (sin registros)")

    os.makedirs(os.path.dirname(salida_path), exist_ok=True)
    with open(salida_path, "w", encoding="utf-8") as f:
        json.dump(resultados, f, ensure_ascii=False, indent=4)

    print(f"\n✅ Comparativa completada. Archivo generado en: {salida_path}")
    print(f"📄 Total registros procesados: {len(resultados)}")

    return resultados
