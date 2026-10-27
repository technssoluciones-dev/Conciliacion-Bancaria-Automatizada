import os
import re
from pathlib import Path

def evaluate_project(root_dir="."):
    root = Path(root_dir)
    score = 0
    report = []

    print("==================================================")
    print("      AUDITOR DE PRODUCTIVIDAD Y OPERATIVIDAD     ")
    print("==================================================\n")

    # 1. Documentación (20 pts)
    if (root / "README.md").exists():
        score += 20
        report.append("[✔] Documentación (README.md): OK (+20%)")
    else:
        report.append("[✘] Documentación (README.md): Faltante (0/20%)")

    # 2. Dependencias y Configuración (20 pts)
    has_config = (root / "requirements.txt").exists() or (root / "package.json").exists() or (root / "pyproject.toml").exists()
    if has_config:
        score += 20
        report.append("[✔] Dependencias y Configuración: OK (+20%)")
    else:
        report.append("[✘] Dependencias y Configuración: Falta gestor de dependencias (0/20%)")

    # 3. Control de Versiones y Git (20 pts)
    if (root / ".git").exists():
        score += 15
        report.append("[✔] Repositorio Git: Inicializado (+15%)")
        if (root / ".gitignore").exists():
            score += 5
            report.append("[✔] Archivo .gitignore: Presente (+5%)")
        else:
            report.append("[✘] Archivo .gitignore: Faltante (0/5%)")
    else:
        report.append("[✘] Repositorio Git: No inicializado (0/20%)")

    # 4. Estructura de Pruebas (20 pts)
    has_tests = (root / "tests").exists() or (root / "test").exists()
    if has_tests:
        score += 20
        report.append("[✔] Carpeta de Pruebas (tests/test): Presente (+20%)")
    else:
        report.append("[✘] Carpeta de Pruebas: Faltante (0/20%)")

    # 5. Calidad del Código (Deuda técnica / TODOs reales) (20 pts)
    todo_count = 0
    excluded_files = {"project_auditor.py", "project_fixer.py", "find_todos.py"}
    todo_pattern = re.compile(r"#.*?\b(TODO|FIXME)\b", re.IGNORECASE)
    
    code_files = list(root.glob("**/*.py"))
    for file_path in code_files:
        if "venv" in file_path.parts or "node_modules" in file_path.parts:
            continue
        if file_path.name in excluded_files:
            continue
        try:
            content = file_path.read_text(encoding="utf-8")
            for line in content.splitlines():
                if todo_pattern.search(line):
                    todo_count += 1
        except Exception:
            pass

    if todo_count == 0:
        score += 20
        report.append("[✔] Deuda Técnica (TODOs/FIXMEs): Sin pendientes críticos (+20%)")
    else:
        penalty = min(20, todo_count * 4)
        score += (20 - penalty)
        report.append(f"[!] Deuda Técnica: Se encontraron {todo_count} pendientes reales (Penalización: -{penalty}%)")

    # Resultado final
    print("--- REPORTE DE HALLAZGOS ---")
    for line in report:
        print(line)
    
    print("\n--------------------------------------------------")
    print(f" PRODUCTIVIDAD ACTUAL DEL PROYECTO: {score}% / 100%")
    print("--------------------------------------------------")
    
    if score == 100:
        print("¡Felicitaciones! El proyecto está 100% operativo.")
    else:
        print(f"Faltan {100 - score}% para alcanzar la operatividad total.")
    
    return score

if __name__ == "__main__":
    evaluate_project()