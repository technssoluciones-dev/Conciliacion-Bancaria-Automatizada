import os
from pathlib import Path

def apply_improvements(root_dir="."):
    root = Path(root_dir)
    actions = []

    print("==================================================")
    print("     APLICANDO MEJORAS Y AUTOMATIZACIÓN           ")
    print("==================================================\n")

    # 1. Crear README.md si no existe
    readme_path = root / "README.md"
    if not readme_path.exists():
        readme_content = """# Conciliación Bancaria Automatizada

## Descripción
Proyecto automatizado en Python para el procesamiento, validación y conciliación de planillas bancarias y contables.

## Requisitos
- Python 3.8+
- Librerías especificadas en `requirements.txt`

## Instalación y Uso
1. Clonar el repositorio.
2. Crear y activar el entorno virtual:
   ```bash
   python -m venv venv
   source venv/bin/activate  # En Windows: venv\\Scripts\\activate