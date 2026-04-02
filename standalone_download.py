import os
import requests
import subprocess
import re
import shutil
import base64
import json
from datetime import datetime

# --- CONFIGURACIÓN ---
API_URL = "http://35.190.132.15:8087/WS_AB/api/Fitosanidad/ZABG_ExcelRptEvaluacionesXVariable"
AUTHORIZATION = "Basic NGdyMUJyNCFuTVM6NDg2Mjc1OTEz"

# Rclone
# SI RCLONE NO ESTÁ EN EL PATH: Reemplaza "rclone" por la ruta completa
# Ejemplo: RCLONE_PATH = r"C:\rclone\rclone.exe"
RCLONE_PATH = r"C:\Users\Cyto\Desktop\rclone-v1.72.0-windows-amd64\rclone.exe" 

RCLONE_REMOTE = "Drive"
DRIVE_FOLDER = "Pimientos/Bot"
TEMP_DIR = "temp_downloads"

# Datos Hardcodeados (Extraídos de Libro1.csv)
FUNDOS = [
    {"name": "Muchik", "code": "238"},
    {"name": "San Pedro", "code": "239"},
    {"name": "Compositan", "code": "234"},
    {"name": "Jayanca", "code": "290"}
]

CARTILLAS = [
    {"name": "PROYECCIÓN: FENOLOGÍA PIMIENTO", "code": 478},
    {"name": "FENOLOGIA VARIEDADES-ENSAYOS PIMIENTO", "code": 623},
    {"name": "FENOLOGÍA 4 HILERAS", "code": 715},
    {"name": "PESOS_COSECHA", "code": 617},
    {"name": "C5-PESOS PROYECCIONES CALIFORNIA", "code": 724},
    {"name": "C6-PESOS PROYECCIONES PIQUILLO", "code": 725},
    {"name": "EVALUACION DIAMETRO NUEVO", "code": 731},
    {"name": "PROYECCIÓN: NDVI PIMIENTO", "code": 480},
    {"name": "CARTILLA PROYECCION P.PIQUILLO", "code": 492},
    {"name": "CARTILLA PROYECCION P.CALIFORNIA 2", "code": 493},
    {"name": "PROYECCIONES: CONTEOS VARIEDADES PIQUILLO", "code": 624},
    {"name": "PROYECCIONES: CONTEOS VARIEDADES CALIFORNIA", "code": 669}
]

def sanitize_filename(name):
    """Limpia el nombre para usarlo en archivos"""
    name = name.replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u')
    name = name.replace('Á', 'A').replace('É', 'E').replace('Í', 'I').replace('Ó', 'O').replace('Ú', 'U')
    name = name.replace('ñ', 'n').replace('Ñ', 'N')
    clean = re.sub(r'[^a-zA-Z0-9_]', '_', name)
    clean = re.sub(r'_+', '_', clean)
    return clean.strip('_')[:30].upper()

def download_report(fundo, cartilla, date_str):
    """Descarga un reporte individual"""
    params = {
        'prmintCultivo': '2',
        'prmstrRUCEmpresa': '20170040938',
        'prmstrFundo': fundo['code'],
        'prmintCartilla': str(cartilla['code']),
        'prmdatFechaInicio': "2025-11-28",
        'prmdatFechaFin': "2025-11-28"
    }
    
    headers = {'Authorization': AUTHORIZATION}
    
    try:
        response = requests.get(API_URL, params=params, headers=headers, timeout=60)
        
        if response.status_code == 200:
            # 1. Obtener texto de respuesta
            response_text = response.text
            
            # 2. Intentar parsear JSON (a veces viene envuelto en {"content": "..."})
            try:
                json_response = json.loads(response_text)
                if isinstance(json_response, dict) and 'content' in json_response:
                    base64_content = json_response['content']
                else:
                    base64_content = response_text
            except json.JSONDecodeError:
                base64_content = response_text
            
            # 3. Decodificar Base64
            try:
                # Limpiar espacios/newlines
                clean_base64 = base64_content.strip().replace('\n', '').replace('\r', '').replace(' ', '')
                content = base64.b64decode(clean_base64)
                
                if len(content) < 100:
                    print(f"  ⚠️ Archivo decodificado muy pequeño para {cartilla['name']}")
                    return None
                    
                return content
            except Exception as e:
                print(f"  ❌ Error decodificando Base64: {e}")
                # Fallback: devolver contenido raw si falla
                return response.content

        else:
            print(f"  ❌ Error HTTP {response.status_code}")
            return None
            
    except Exception as e:
        print(f"  ❌ Error de conexión: {e}")
        return None

def upload_to_drive(local_path, filename):
    """Sube el archivo usando Rclone"""
    remote_path = f"{RCLONE_REMOTE}:{DRIVE_FOLDER}/{filename}"
    
    cmd = [RCLONE_PATH, "copyto", local_path, remote_path]
    
    try:
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode == 0:
            print(f"  ☁️ Subido a Drive: {remote_path}")
            return True
        else:
            print(f"  ❌ Error Rclone: {result.stderr}")
            return False
    except FileNotFoundError:
        print("  ❌ Error: rclone no está instalado o no está en el PATH")
        return False

def main():
    print("🚀 Iniciando Script de Descarga Standalone")
    
    # Preparar carpeta temporal
    if os.path.exists(TEMP_DIR):
        shutil.rmtree(TEMP_DIR)
    os.makedirs(TEMP_DIR)
    
    # FECHA MANUAL: 28/11/2025
    today = "2025-11-28"
    date_compact = "20251128"
    # today = datetime.now().strftime('%Y-%m-%d')
    # date_compact = datetime.now().strftime('%Y%m%d')
    
    print(f"📅 Fecha: {today}")
    
    for fundo in FUNDOS:
        print(f"\n🏢 Fundo: {fundo['name']}")
        
        for cartilla in CARTILLAS:
            print(f"  ⬇️ Procesando: {cartilla['name']}...")
            
            content = download_report(fundo, cartilla, today)
            
            if content:
                # Generar nombre
                abbr = sanitize_filename(cartilla['name'])
                filename = f"{abbr}_{fundo['code']}_{date_compact}.xlsx"
                local_path = os.path.join(TEMP_DIR, filename)
                
                # Guardar localmente
                with open(local_path, 'wb') as f:
                    f.write(content)
                
                # Subir a Drive
                upload_to_drive(local_path, filename)
            else:
                print("     (Saltado)")

    # Limpieza final
    print("\n🧹 Limpiando archivos temporales...")
    shutil.rmtree(TEMP_DIR)
    print("🏁 Proceso completado.")

if __name__ == "__main__":
    main()

