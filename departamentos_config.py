"""
departamentos_config.py
Configuración para extraer iniciales de departamentos de los Exceles
"""

def extraer_iniciales_departamento(texto: str):
    """
    Extrae iniciales de departamento de un texto
    """
    if not texto or not isinstance(texto, str):
        return []
    
    departamentos = set()
    texto_upper = texto.upper()
    
    # Mapeo de departamentos comunes
    mapeo_departamentos = {
        'RH': ['RECURSOS HUMANOS', 'RRHH', 'HUMAN RESOURCES'],
        'TI': ['TECNOLOGÍA INFORMACIÓN', 'SISTEMAS', 'INFORMÁTICA', 'IT'],
        'FIN': ['FINANZAS', 'CONTABILIDAD', 'TESORERÍA'],
        'ADM': ['ADMINISTRACIÓN', 'ADMINISTRATIVO'],
        'VENT': ['VENTAS', 'COMERCIAL'],
        'MKT': ['MARKETING', 'MERCADEO'],
        'LOG': ['LOGÍSTICA', 'ALMACÉN'],
        'PROD': ['PRODUCCIÓN', 'OPERACIONES'],
        'CAL': ['CALIDAD'],
        'JUR': ['JURÍDICO', 'LEGAL'],
        'COM': ['COMUNICACIONES'],
        'DIR': ['DIRECCIÓN', 'GERENCIA'],
        'DOC': ['DOCENCIA', 'ACADÉMICO'],
        'BIEN': ['BIENESTAR'],
        'SEC': ['SECRETARÍA'],
        'MNT': ['MANTENIMIENTO']
    }
    
    # Buscar coincidencias
    for inicial, palabras in mapeo_departamentos.items():
        for palabra in palabras:
            if palabra in texto_upper:
                departamentos.add(inicial)
                break
    
    return list(departamentos)

def analizar_departamentos_dataframe(df):
    """
    Analiza un DataFrame para extraer departamentos
    """
    departamentos = set()
    
    # Buscar en nombres de columnas
    for col in df.columns:
        iniciales = extraer_iniciales_departamento(str(col))
        departamentos.update(iniciales)
    
    # Buscar en los primeros 50 registros de cada columna
    for col in df.columns:
        if df[col].dtype == 'object':  # Solo columnas de texto
            try:
                muestras = df[col].dropna().head(50)
                for valor in muestras:
                    iniciales = extraer_iniciales_departamento(str(valor))
                    departamentos.update(iniciales)
            except:
                continue
    
    return departamentos

def generar_resumen_departamentos(departamentos_encontrados):
    """
    Genera un resumen legible de los departamentos encontrados
    """
    if not departamentos_encontrados:
        return "No se detectaron departamentos específicos"
    
    resumen = "📊 DEPARTAMENTOS DETECTADOS:\n"
    for depto in sorted(departamentos_encontrados):
        resumen += f"  • {depto}\n"
    
    return resumen