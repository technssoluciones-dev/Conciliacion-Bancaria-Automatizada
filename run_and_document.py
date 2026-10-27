import subprocess
import time
from pathlib import Path
import threading
import cv2
import numpy as np
from mss import mss
import imageio
import re

def record_process(stop_event, frames_list):
    fps = 10.0  # 10 FPS estables para mantener un GIF liviano y fluido
    frame_duration = 1.0 / fps
    scale_percent = 50  # Reduce la resolución al 50% para cuidar el peso del archivo
    
    with mss() as sct:
        monitor = sct.monitors[1]
        orig_width = monitor["width"]
        orig_height = monitor["height"]
        
        target_width = int(orig_width * scale_percent / 100)
        target_height = int(orig_height * scale_percent / 100)
        
        print("[REC] Grabación para GIF iniciada en segundo plano...")
        next_frame_time = time.time()
        
        while not stop_event.is_set():
            current_time = time.time()
            if current_time >= next_frame_time:
                img = np.array(sct.grab(monitor))
                # Convertir de BGRA a RGB
                frame = cv2.cvtColor(img, cv2.COLOR_BGRA2RGB)
                
                resized_frame = cv2.resize(frame, (target_width, target_height), interpolation=cv2.INTER_AREA)
                frames_list.append(resized_frame)
                
                next_frame_time += frame_duration
            else:
                time.sleep(0.005)
                
        print("[✔] Captura de fotogramas finalizada.")

def update_readme():
    readme_path = Path("README.md")
    if not readme_path.exists():
        print("[✘] No se encontró el archivo README.md para actualizar.")
        return

    content = readme_path.read_text(encoding="utf-8")
    
    # Limpiar cualquier sección de demo anterior para evitar duplicados
    content = re.sub(r'\n## 🎥 Demostración.*?(?=\n# |$)', '', content, flags=re.DOTALL)
    
    demo_section = """
## 🎥 Demostración y Funcionamiento del Sistema

> Sistema automatizado diseñado para optimizar, validar y acelerar el proceso de conciliación entre los libros contables y las cartolas bancarias, reduciendo drásticamente la carga operativa manual.

### 📋 ¿Cómo opera el sistema?
1. **Carga de Datos:** El usuario selecciona el formato contable (ej. Libro Diario) y adjunta la cartola bancaria en múltiples formatos compatibles (`.xlsx`, `.xls`, `.csv`, `.pdf`).
2. **Procesamiento Inteligente:** El motor interno impulsado por Python (`pandas`, `openpyxl`, `pypdf`) parsea, limpia y estructura la información financiera de forma automatizada.
3. **Generación de Reportes:** El sistema cruza las transacciones y genera reportes de conciliación listos para auditoría y gestión administrativa.

<p align="center">
  <img src="assets/demo_conciliacion.gif" alt="Demostración en Vivo de la Conciliación Bancaria" width="100%">
</p>
"""

    # Insertar la sección de demostración justo después del primer título principal (# )
    if content.startswith("# "):
        first_newline = content.find("\n")
        if first_newline != -1:
            new_content = content[:first_newline] + "\n" + demo_section + content[first_newline:]
        else:
            new_content = content + "\n" + demo_section
    else:
        new_content = demo_section + "\n" + content

    readme_path.write_text(new_content, encoding="utf-8")
    print("[+] ¡README.md actualizado con la demostración posicionada al principio!")

def main():
    output_path = Path("assets/demo_conciliacion.gif")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    print("==================================================")
    print("   GENERADOR DE DEMOS EN GIF PARA GITHUB          ")
    print("==================================================\n")
    
    stop_event = threading.Event()
    frames_list = []
    
    recorder_thread = threading.Thread(target=record_process, args=(stop_event, frames_list))
    recorder_thread.start()
    
    time.sleep(1.5)
    
    print("[*] Abriendo la interfaz gráfica de conciliación...")
    try:
        subprocess.run(["python", "conciliador_gui.py"], check=True)
    except Exception as e:
        print(f"[!] Error al ejecutar la interfaz: {e}")
    finally:
        print("[*] Interfaz cerrada. Compilando y optimizando el archivo GIF...")
        stop_event.set()
        recorder_thread.join()
        
        if frames_list:
            imageio.mimsave(str(output_path), frames_list, fps=10, loop=0)
            print(f"[✔] ¡GIF generado exitosamente en: {output_path}!")
        else:
            print("[!] No se capturaron fotogramas para el GIF.")
        
        print("[*] Actualizando documentación en el README...")
        update_readme()
        print("\n[✔] ¡Proceso completo! Ya puedes hacer commit y push para ver el GIF arriba en GitHub.")

if __name__ == "__main__":
    main()