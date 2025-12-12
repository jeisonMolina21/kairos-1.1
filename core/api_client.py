"""
Cliente para comunicación con API REST de gestión de asistencias
CON SISTEMA DE FALLBACK AUTOMÁTICO
"""
import requests
import os
from typing import Dict, List, Any

class APIClient:
    def __init__(self, email: str, password: str):
        self.email = email
        self.password = password
        self.token = None
        self.primary_url = "https://api.rh360.unihorizonte.edu.co/api"
        self.secondary_url = "http://10.0.1.14:5000/api"
        self.current_base_url = self.primary_url  # Por defecto usa la primaria
        self.timeout = 15
        self.is_using_secondary = False

    def _try_request(self, method, endpoint, **kwargs):
        """
        Intenta hacer una petición a la API primaria, si falla intenta con la secundaria
        """
        # Primero intentamos con la API primaria
        try:
            url = f"{self.primary_url}{endpoint}"
            print(f"🌐 Intentando con API PRIMARIA: {url}")
            response = requests.request(method, url, timeout=self.timeout, **kwargs)
            response.raise_for_status()
            self.current_base_url = self.primary_url
            self.is_using_secondary = False
            return response
        except requests.exceptions.RequestException as primary_error:
            print(f"❌ API PRIMARIA falló: {primary_error}")
            
            # Si falla la primaria, intentamos con la secundaria
            try:
                url = f"{self.secondary_url}{endpoint}"
                print(f"🌐 Intentando con API SECUNDARIA: {url}")
                response = requests.request(method, url, timeout=self.timeout, **kwargs)
                response.raise_for_status()
                self.current_base_url = self.secondary_url
                self.is_using_secondary = True
                print("✅ Conectado a API SECUNDARIA")
                return response
            except requests.exceptions.RequestException as secondary_error:
                print(f"❌ API SECUNDARIA también falló: {secondary_error}")
                self.is_using_secondary = False
                raise Exception(f"Ambas APIs fallaron. Primaria: {primary_error}, Secundaria: {secondary_error}")

    def login(self) -> str:
        """
        Login con sistema de fallback automático
        """
        print("🔐 Iniciando autenticación con sistema de fallback...")
        
        payload = {"mail": self.email, "password": self.password}
        headers = {"Content-Type": "application/json"}

        try:
            response = self._try_request("POST", "/auth/login", json=payload, headers=headers)
            
            data = response.json()
            self.token = data.get("token")
            
            if not self.token:
                raise Exception("No se recibió token en la respuesta")
                
            if self.is_using_secondary:
                print("🔑 Autenticación exitosa en API SECUNDARIA")
            else:
                print("🔑 Autenticación exitosa en API PRIMARIA")
                
            return self.token
            
        except Exception as e:
            raise Exception(f"Error de autenticación en ambas APIs: {e}")

    def get_statistics(self) -> List[Dict[str, Any]]:
        """
        Obtiene estadísticas de permisos y horarios desde RH360
        """
        if not self.token:
            raise Exception("No hay token válido. Ejecuta login() primero.")

        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.token}"
        }

        try:
            response = self._try_request("GET", "/permissions/statistics", headers=headers)
            
            data = response.json()
            print(f"📊 Estadísticas recibidas: {len(data)} registros")
            if self.is_using_secondary:
                print("📡 Datos obtenidos desde API SECUNDARIA")
            else:
                print("📡 Datos obtenidos desde API PRIMARIA")
            print("🔗 URLs de permisos disponibles en campo 'permiso_url'")
            
            return data
            
        except Exception as e:
            raise Exception(f"Error obteniendo estadísticas: {e}")

    def enviar_asistencias_api(self, datos: List[Dict]) -> bool:
        """
        Envía los datos de asistencia a la API con fallback automático
        """
        if not self.token:
            raise Exception("No hay token válido. Ejecuta login() primero.")

        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.token}"
        }

        try:
            response = self._try_request("POST", "/attendance-reports/bulk-create", 
                                       json=datos, headers=headers)
            
            print(f"✅ Datos enviados a la API: {len(datos)} registros")
            if self.is_using_secondary:
                print("📤 Envío realizado a API SECUNDARIA")
            else:
                print("📤 Envío realizado a API PRIMARIA")
            return True
            
        except Exception as e:
            raise Exception(f"Error enviando datos a la API: {e}")

    def enviar_excel_api(self, archivo_path: str) -> bool:
        """
        Envía el archivo Excel a la API con fallback automático
        """
        if not self.token:
            raise Exception("No hay token válido. Ejecuta login() primero.")

        headers = {
            "Authorization": f"Bearer {self.token}"
        }

        try:
            with open(archivo_path, "rb") as f:
                files = {"file": (os.path.basename(archivo_path), f, 
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
                
                response = self._try_request("POST", "/attendance-reports/upload-excel", 
                                           files=files, headers=headers)
            
            print(f"✅ Archivo Excel enviado a la API: {archivo_path}")
            if self.is_using_secondary:
                print("📤 Archivo enviado a API SECUNDARIA")
            else:
                print("📤 Archivo enviado a API PRIMARIA")
            return True
            
        except Exception as e:
            raise Exception(f"Error enviando archivo Excel a la API: {e}")

    def enviar_json_api(self, archivo_path: str) -> bool:
        """
        Envía el archivo JSON a la API con fallback automático
        """
        if not self.token:
            raise Exception("No hay token válido. Ejecuta login() primero.")

        headers = {
            "Authorization": f"Bearer {self.token}"
        }

        try:
            with open(archivo_path, "rb") as f:
                files = {"file": (os.path.basename(archivo_path), f, "application/json")}
                response = self._try_request("POST", "/attendance-reports", 
                                           files=files, headers=headers)
            
            print(f"✅ Archivo JSON enviado a la API: {archivo_path}")
            if self.is_using_secondary:
                print("📤 Archivo enviado a API SECUNDARIA")
            else:
                print("📤 Archivo enviado a API PRIMARIA")
            return True
            
        except Exception as e:
            raise Exception(f"Error enviando archivo JSON a la API: {e}")

    def get_current_api_status(self) -> str:
        """
        Retorna el estado actual de la conexión API
        """
        if self.is_using_secondary:
            return "Conectado a API SECUNDARIA (rh360)"
        else:
            return "Conectado a API PRIMARIA (local)"

    def test_both_apis(self) -> Dict[str, bool]:
        """
        Testea la conectividad de ambas APIs
        """
        results = {}
        
        # Test API primaria
        try:
            response = requests.get(f"{self.primary_url}/auth/test", timeout=10)
            results["primary"] = response.status_code == 200
        except:
            results["primary"] = False
            
        # Test API secundaria  
        try:
            response = requests.get(f"{self.secondary_url}/auth/test", timeout=10)
            results["secondary"] = response.status_code == 200
        except:
            results["secondary"] = False
            
        return results