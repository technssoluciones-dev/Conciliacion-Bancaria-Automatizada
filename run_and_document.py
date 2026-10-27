import subprocess
import time
from pathlib import Path
import threading
import cv2
import numpy as np
from mss import mss

def record_process(stop_event, output_path):
    fps = 12.0  # 12 FPS fluidos y ligeros para demostraciones de escritorio
    frame_duration = 1.0 / fps
    scale_percent = 50  # Reduce la resolución al 50% para mantener el video en un rango de 5-10 MB
    
    with mss() as sct:
        monitor = sct.monitors[1]
        orig_width = monitor["width"]
        orig_height = monitor["height"]
        
        target_width = int(orig_width * scale_percent / 100)
        target_height = int(orig_height * scale_percent / 100)
        
        fourcc = cv2.VideoWriter_fourcc(*"mp4v")
        out = cv2.VideoWriter(str(output_path), fourcc, fps, (target_width, target_height))
        
        print("[REC] Grabación optimizada iniciada (Resolución ajustada para bajo peso)...")
        next_frame_time = time.time()
        
        while not stop_event.is_set():
            current_time = time.time()
            if current_time >= next_frame_time:
                img = np.array(sct.grab(monitor))
                frame = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
                
                # Redimensionar el fotograma para reducir drásticamente el peso del archivo
                resized_frame = cv2.resize(frame, (target_width, target_height), interpolation=cv2.INTER_AREA)
                out.write(resized_frame)
                
                next_frame_time += frame_duration
            else:
                time.sleep(0.005)
                
        out.release()
        print("[✔] Grabación optimizada finalizada con éxito.")

def update_readme():
    readme_path = Path("README.md")
    if not readme_path.exists():
        print("[✘] No se encontró el archivo README.md para actualizar.")
        return

    content = readme_path.read_text(encoding="utf-8")
    
    demo_section = """
## 🎥 Demostración y Funcionamiento del Sistema

> Sistema automatizado diseñado para optimizar, validar y acelerar el proceso de conciliación entre los libros contables y las cartolas bancarias, reduciendo drásticamente la carga operativa manual.

### 📋 ¿Cómo opera el sistema?
1. **Carga de Datos:** El usuario selecciona el formato contable (ej. Libro Diario) y adjunta la cartola bancaria en múltiples formatos compatibles (`.xlsx`, `.xls`, `.csv`, `.pdf`).
2. **Procesamiento Inteligente:** El motor interno impulsado por Python (`pandas`, `openpyxl`, `pypdf`) parsea, limpia y estructura la información financiera de forma automatizada.
3. **Generación de Reportes:** El sistema cruza las transacciones y genera reportes de conciliación listos para auditoría y gestión administrativa.

<p align="center">
  <video width="100%" controls autoplay muted loop>
    <source src="assets/demo_conciliacion.mp4" type="video/mp4">
    Tu navegador no soporta la etiqueta de video.
  </video>
</p>
"""

    if "## 🎥 Demostración y Funcionamiento del Sistema" in content:
        print("[*] La sección de video ya está integrada en el README.")
    else:
        new_content = content + "\n" + demo_section
        readme_path.write_text(new_content, encoding="utf-8")
        print("[+] ¡README.md actualizado correctamente!")

def main():
    output_path = Path("assets/demo_conciliacion.mp4")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    print("==================================================")
    print("   LANZADOR LIGERO: GRABACIÓN OPTIMIZADA (5-10MB) ")
    print("==================================================\n")
    print("[*] Preparando motor de compresión y grabación...")
    
    stop_event = threading.Event()
    
    recorder_thread = threading.Thread(target=record_process, args=(stop_event, output_path))
    recorder_thread.start()
    
    time.sleep(1.5)
    
    print("[*] Abriendo la interfaz gráfica de conciliación...")
    try:
        subprocess.run(["python", "conciliador_gui.py"], check=True)
    except Exception as e:
        print(f"[!] Error al ejecutar la interfaz: {e}")
    finally:
        print("[*] Interfaz cerrada. Compilando video optimizado...")
        stop_event.set()
        recorder_thread.join()
        print(f"\n[✔] ¡Video optimizado guardado en: {output_path}!")
        
        print("[*] Actualizando documentación en el README...")
        update_readme()
        print("\n[✔] ¡Ciclo completado con éxito! El archivo es ahora ligero y apto para GitHub.")

if __name__ == "__main__":
    main()