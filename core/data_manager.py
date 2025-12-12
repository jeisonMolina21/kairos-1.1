import json
import os
import pandas as pd
from datetime import datetime
from typing import Dict, List, Any, Optional

from .api_client import APIClient

class DataManager:
    def __init__(self):
        self.api_client = APIClient()
        self.output_dir = "outputs"
        self.temp_dir = "temp"
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.temp_dir, exist_ok=True)

    def load_all_data(self) -> Dict[str, Any]:
        """
        Carga todos los datos necesarios desde la API o cache
        """
        print("\n🔄 CARGANDO TODOS LOS DATOS...")
        
        # Cargar datos completos desde API
        complete_data = self.api_client.get_complete_data()
        
        if not complete_data.get("statistics"):
            print("❌ No se pudieron cargar datos de la API")
            return {}
        
        # Guardar datos individuales para compatibilidad
        self._save_individual_files(complete_data)
        
        return complete_data

    def _save_individual_files(self, complete_data: Dict[str, Any]):
        """Guarda archivos individuales para compatibilidad con el sistema existente"""
        # Guardar statistics.json
        stats_path = os.path.join(self.output_dir, "statistics.json")
        with open(stats_path, "w", encoding="utf-8") as f:
            json.dump(complete_data.get("statistics", []), f, ensure_ascii=False, indent=4)
        print(f"💾 statistics.json guardado ({len(complete_data.get('statistics', []))} registros)")

        # Guardar datos completos en temp
        complete_path = os.path.join(self.temp_dir, "complete_api_data.json")
        with open(complete_path, "w", encoding="utf-8") as f:
            json.dump(complete_data, f, ensure_ascii=False, indent=4)
        print(f"💾 Datos completos guardados en temp/")

    def get_carnet_mapping(self, complete_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Extrae y procesa el mapeo de carnets desde los datos completos
        """
        print("\n🔍 PROCESANDO MAPEO DE CARNETS...")
        
        statistics = complete_data.get("statistics", [])
        carnet_map = {}
        fecha_carnet_map = {}
        
        for persona in statistics:
            try:
                carnet = persona.get("carnet", "No")
                if carnet == "Sí":
                    carnet_number = str(persona.get("carnet_number", "")).strip()
                    number_id = str(persona.get("number_id", "")).strip()
                    nombre_completo = persona.get("nombre", "").strip()
                    
                    if carnet_number and carnet_number != "N/A" and number_id:
                        # Mapeo general por carnet
                        carnet_map[carnet_number] = {
                            'number_id': number_id,
                            'nombre': nombre_completo,
                            'apellido': ''  # No hay apellido separado en el JSON
                        }
                        
                        # Mapeo específico por fecha
                        año = persona.get("año")
                        mes = persona.get("mes") 
                        dia = persona.get("dia")
                        
                        if año and mes and dia:
                            try:
                                # Asegurar formato correcto de fecha
                                mes_str = str(mes).zfill(2)
                                dia_str = str(dia).zfill(2)
                                fecha_str = f"{año}-{mes_str}-{dia_str}"
                                fecha_carnet_map[(carnet_number, fecha_str)] = {
                                    'number_id': number_id,
                                    'nombre': nombre_completo,
                                    'apellido': ''
                                }
                                print(f"📋 Mapeo carnet: {carnet_number} -> {number_id} ({nombre_completo}) para {fecha_str}")
                            except Exception as e:
                                continue
            except Exception as e:
                continue
        
        print(f"✅ Mapeo de carnets procesado: {len(carnet_map)} carnets generales, {len(fecha_carnet_map)} mapeos por fecha")
        
        return {
            'carnet_map': carnet_map,
            'fecha_carnet_map': fecha_carnet_map
        }

    def save_processed_data(self, data: Dict[str, Any], filename: str):
        """Guarda datos procesados en la carpeta temp"""
        filepath = os.path.join(self.temp_dir, filename)
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        print(f"💾 Datos procesados guardados: {filename}")

    def load_processed_data(self, filename: str) -> Dict[str, Any]:
        """Carga datos procesados desde la carpeta temp"""
        filepath = os.path.join(self.temp_dir, filename)
        if os.path.exists(filepath):
            with open(filepath, "r", encoding="utf-8") as f:
                return json.load(f)
        return {}