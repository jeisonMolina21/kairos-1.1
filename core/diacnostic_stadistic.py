# diagnostic_completo.py
import json
import pandas as pd
import os

def diagnosticar_problema():
    print("🔍 DIAGNÓSTICO COMPLETO DEL PROBLEMA")
    print("=" * 60)
    
    # Verificar statistics.json
    stats_path = "outputs/statistics.json"
    if not os.path.exists(stats_path):
        print("❌ statistics.json no existe")
        return
    
    with open(stats_path, 'r', encoding='utf-8') as f:
        stats = json.load(f)
    
    print(f"✅ Statistics cargado: {len(stats)} registros")
    
    # Encontrar registros con carnet
    carnets_si = [r for r in stats if r.get('carnet') == 'Sí' and r.get('carnet_number') not in ['N/A', '', None]]
    print(f"🎫 Registros con carnet='Sí': {len(carnets_si)}")
    
    for registro in carnets_si[:5]:
        print(f"   - {registro.get('nombre')}: {registro.get('carnet_number')} - {registro.get('año')}-{registro.get('mes')}-{registro.get('dia')}")
    
    # Verificar archivos Excel
    excel_files = [f for f in os.listdir() if f.endswith('.xlsx') or f.endswith('.xls')]
    print(f"\n📊 Archivos Excel encontrados: {len(excel_files)}")
    for f in excel_files[:3]:
        print(f"   - {f}")
    
    # Leer un archivo Excel de ejemplo para ver estructura
    if excel_files:
        try:
            df_ejemplo = pd.read_excel(excel_files[0])
            print(f"\n📋 Estructura de {excel_files[0]}:")
            print(f"   Columnas: {list(df_ejemplo.columns)}")
            print(f"   Filas: {len(df_ejemplo)}")
            
            # Buscar columnas de tarjeta
            tarjeta_cols = [col for col in df_ejemplo.columns if 'tarjeta' in str(col).lower() or 'carnet' in str(col).lower() or 'card' in str(col).lower()]
            print(f"   Columnas de tarjeta: {tarjeta_cols}")
            
            if tarjeta_cols:
                print(f"   Valores únicos en {tarjeta_cols[0]}: {df_ejemplo[tarjeta_cols[0]].dropna().unique()[:10]}")
                
        except Exception as e:
            print(f"❌ Error leyendo Excel: {e}")

if __name__ == "__main__":
    diagnosticar_problema()