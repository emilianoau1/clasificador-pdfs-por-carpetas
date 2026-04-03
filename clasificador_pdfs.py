"""
clasificador_pdfs.py
---------------------
Organiza automáticamente cientos de archivos PDF en carpetas
estructuradas por Responsable y Grupo, usando un Excel como mapa de referencia.

Cada PDF se nombra con un identificador único (ej. número de unidad o folio).
El script lee el Excel, encuentra el responsable y grupo correspondiente
a cada archivo, y lo copia a la carpeta correcta automáticamente.

Características:
  - Detección automática de columnas en el Excel
  - Modo simulación (--dry-run) para revisar antes de ejecutar
  - Reporte CSV con resultados detallados
  - Resumen por responsable y por grupo al finalizar

Autor: [Tu nombre]
Uso:   python clasificador_pdfs.py [--dry-run] [--excel ruta.xlsx] [--report reporte.csv]
"""

import argparse
import re
import shutil
import sys
from pathlib import Path
from typing import Dict, Tuple, Optional
from collections import defaultdict


# ─────────────────────────────────────────────
# CONFIGURACIÓN POR DEFECTO
# ─────────────────────────────────────────────

DEFAULT_PDF_DIR    = Path(r'C:\ruta\a\tu\carpeta\PDFS')
DEFAULT_EXCEL_NAME = None   # None = autodetecta el único .xlsx en la carpeta


# ─────────────────────────────────────────────
# UTILIDADES
# ─────────────────────────────────────────────

def cargar_pandas():
    try:
        import pandas as pd
        return pd
    except ImportError:
        sys.stderr.write("ERROR: Instala pandas con: pip install pandas openpyxl\n")
        sys.exit(1)


ID_PATTERN = re.compile(r"^[A-HJ-NPR-Z0-9]{6,17}$", re.IGNORECASE)


def normalizar_id(valor: str) -> str:
    """Elimina espacios y caracteres especiales. Convierte a mayúsculas."""
    if valor is None:
        return ""
    valor = str(valor).strip().upper()
    return re.sub(r"[^A-Z0-9]", "", valor)


def nombre_carpeta_seguro(nombre: str) -> str:
    """Limpia un texto para usarlo como nombre de carpeta en Windows/Linux."""
    if not nombre:
        return "_SIN_NOMBRE"
    nombre = str(nombre).strip()
    return re.sub(r'[<>:"/\\|?*]+', "_", nombre) or "_SIN_NOMBRE"


def detectar_columna_id(columnas):
    """Detecta automáticamente la columna de identificador único."""
    mapa = {c.lower(): c for c in columnas}
    for candidato in ["id", "folio", "vin", "numero", "clave", "código"]:
        if candidato in mapa:
            return mapa[candidato]
    return None


def detectar_columna_grupo(columnas):
    """Detecta automáticamente la columna de grupo o área."""
    mapa = {c.lower(): c for c in columnas}
    for candidato in ["grupo", "group", "ecogroup", "area", "área", "departamento", "región"]:
        if candidato in mapa:
            return mapa[candidato]
    for c in columnas:
        if any(k in c.lower() for k in ["grupo", "group", "area", "region"]):
            return c
    return None


def detectar_columna_responsable(columnas):
    """Detecta automáticamente la columna de responsable o ejecutivo."""
    mapa = {c.lower(): c for c in columnas}
    for candidato in ["responsable", "ejecutivo", "asesor", "vendedor", "agente", "manager"]:
        if candidato in mapa:
            return mapa[candidato]
    for c in columnas:
        if any(k in c.lower() for k in ["ejecut", "respons", "asesor", "manager", "agent"]):
            return c
    return None


# ─────────────────────────────────────────────
# LECTURA DEL EXCEL
# ─────────────────────────────────────────────

def leer_mapeo(
    excel_path: Path,
    hoja: Optional[str],
    col_id: Optional[str],
    col_grupo: Optional[str],
    col_responsable: Optional[str],
) -> Tuple[dict, dict, set, dict]:
    """
    Lee el Excel y construye el mapa: ID -> (Grupo, Responsable).

    Retorna:
      - mapeo          : dict ID -> (grupo, responsable)
      - duplicados     : dict ID -> cantidad de veces que aparece
      - responsables   : set con todos los responsables
      - resp_a_grupos  : dict responsable -> set de grupos asignados
    """
    pd = cargar_pandas()

    try:
        xls = pd.ExcelFile(excel_path)
        hoja_usar = hoja or xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=hoja_usar)
    except Exception as e:
        sys.stderr.write(f"ERROR al leer el Excel: {e}\n")
        sys.exit(2)

    # Detección automática de columnas si no se especifican
    col_id          = col_id          or detectar_columna_id(df.columns)
    col_grupo       = col_grupo       or detectar_columna_grupo(df.columns)
    col_responsable = col_responsable or detectar_columna_responsable(df.columns)

    if not col_id or not col_grupo:
        sys.stderr.write(
            f"ERROR: No se detectaron columnas de ID y Grupo.\n"
            f"Columnas disponibles: {list(df.columns)}\n"
            f"Usa --id-col y --group-col para especificarlas.\n"
        )
        sys.exit(3)

    # Preparar dataframe
    cols_usar = [col_id, col_grupo]
    if col_responsable:
        cols_usar.append(col_responsable)

    df = df[cols_usar].copy()
    df.rename(columns={col_id: "_ID", col_grupo: "_GRUPO"}, inplace=True)
    if col_responsable:
        df.rename(columns={col_responsable: "_RESPONSABLE"}, inplace=True)
    else:
        df["_RESPONSABLE"] = "_SIN_RESPONSABLE"

    df["_ID"]           = df["_ID"].map(normalizar_id)
    df["_GRUPO"]        = df["_GRUPO"].astype(str).str.strip()
    df["_RESPONSABLE"]  = df["_RESPONSABLE"].astype(str).str.strip()
    df["_RESPONSABLE"]  = df["_RESPONSABLE"].replace({"": "_SIN_RESPONSABLE", "nan": "_SIN_RESPONSABLE"})

    # Filtrar filas inválidas
    df = df[(df["_ID"] != "") & df["_GRUPO"].notna() & (df["_GRUPO"] != "")]
    df = df[df["_ID"].str.fullmatch(ID_PATTERN)]

    duplicados = df["_ID"].value_counts().to_dict()
    df = df.drop_duplicates(subset=["_ID"], keep="first")

    mapeo = {row["_ID"]: (str(row["_GRUPO"]), str(row["_RESPONSABLE"])) for _, row in df.iterrows()}
    responsables = set(df["_RESPONSABLE"].unique())
    resp_a_grupos = defaultdict(set)
    for _, row in df.iterrows():
        resp_a_grupos[row["_RESPONSABLE"]].add(row["_GRUPO"])

    return mapeo, duplicados, responsables, resp_a_grupos


# ─────────────────────────────────────────────
# OPERACIONES DE ARCHIVO
# ─────────────────────────────────────────────

def crear_carpeta(ruta: Path, dry_run: bool):
    if not ruta.exists():
        if dry_run:
            print(f"[simulación] Crear carpeta: {ruta}")
        else:
            ruta.mkdir(parents=True, exist_ok=True)


def copiar_archivo(origen: Path, destino: Path, dry_run: bool):
    if dry_run:
        print(f"[simulación] Copiar '{origen.name}' → '{destino}'")
    else:
        shutil.copy2(str(origen), str(destino))


# ─────────────────────────────────────────────
# PROCESO PRINCIPAL
# ─────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description=(
            "Clasifica archivos PDF en carpetas por Responsable y Grupo "
            "usando un Excel como referencia."
        )
    )
    parser.add_argument("--excel",      help="Ruta al Excel con columnas ID, Grupo y Responsable.")
    parser.add_argument("--pdf-dir",    dest="pdf_dir",    help="Carpeta con los PDFs a clasificar.")
    parser.add_argument("--output-dir", dest="output_dir", help="Carpeta de salida (default: misma que --pdf-dir).")
    parser.add_argument("--sheet",      default=None,      help="Hoja del Excel (default: primera hoja).")
    parser.add_argument("--id-col",     dest="id_col",     default=None, help="Nombre de la columna de identificador.")
    parser.add_argument("--group-col",  dest="group_col",  default=None, help="Nombre de la columna de grupo.")
    parser.add_argument("--resp-col",   dest="resp_col",   default=None, help="Nombre de la columna de responsable.")
    parser.add_argument("--dry-run",    action="store_true", help="Simula el proceso sin copiar archivos.")
    parser.add_argument("--report",     default=None,      help="Ruta para guardar reporte CSV.")
    args = parser.parse_args()

    # Resolver rutas
    pdf_dir = Path(args.pdf_dir).expanduser() if args.pdf_dir else DEFAULT_PDF_DIR
    if not pdf_dir.exists():
        sys.stderr.write(f"ERROR: Carpeta de PDFs no encontrada: {pdf_dir}\n")
        sys.exit(5)

    if args.excel:
        excel_path = Path(args.excel).expanduser()
    else:
        candidatos = list(pdf_dir.glob("*.xlsx")) + list(pdf_dir.glob("*.xls"))
        if not candidatos:
            sys.stderr.write("ERROR: No se encontró ningún Excel en la carpeta. Usa --excel.\n")
            sys.exit(4)
        if len(candidatos) > 1:
            sys.stderr.write(f"ERROR: Hay varios Excel. Especifica --excel. Candidatos: {[p.name for p in candidatos]}\n")
            sys.exit(4)
        excel_path = candidatos[0]

    output_root = Path(args.output_dir).expanduser() if args.output_dir else pdf_dir

    if not excel_path.exists():
        sys.stderr.write(f"ERROR: Excel no encontrado: {excel_path}\n")
        sys.exit(4)

    print(f"Excel:   {excel_path}")
    print(f"PDFs:    {pdf_dir}")
    print(f"Salida:  {output_root}")
    if args.dry_run:
        print("[MODO SIMULACIÓN — no se copiarán archivos]\n")

    # Leer mapeo del Excel
    mapeo, duplicados, todos_responsables, resp_a_grupos = leer_mapeo(
        excel_path, args.sheet, args.id_col, args.group_col, args.resp_col
    )

    # Crear estructura de carpetas
    for responsable in sorted(todos_responsables):
        dir_resp = output_root / nombre_carpeta_seguro(responsable)
        crear_carpeta(dir_resp, args.dry_run)
        for grupo in sorted(resp_a_grupos[responsable]):
            crear_carpeta(dir_resp / nombre_carpeta_seguro(grupo), args.dry_run)

    # Clasificar y copiar PDFs
    total = copiados = sin_coincidencia = 0
    ids_en_excel  = set(mapeo.keys())
    ids_encontrados = set()
    conteo_grupo   = defaultdict(int)
    conteo_resp    = defaultdict(int)
    conteo_resp_grupo = defaultdict(lambda: defaultdict(int))

    for archivo in pdf_dir.iterdir():
        if not archivo.is_file() or archivo.suffix.lower() != ".pdf":
            continue
        total += 1
        id_archivo = normalizar_id(archivo.stem)

        if id_archivo in mapeo:
            grupo, responsable = mapeo[id_archivo]
            destino = output_root / nombre_carpeta_seguro(responsable) / nombre_carpeta_seguro(grupo) / archivo.name
            crear_carpeta(destino.parent, args.dry_run)
            copiar_archivo(archivo, destino, args.dry_run)

            copiados += 1
            ids_encontrados.add(id_archivo)
            conteo_grupo[grupo] += 1
            conteo_resp[responsable] += 1
            conteo_resp_grupo[responsable][grupo] += 1
        else:
            sin_coincidencia += 1

    ids_sin_pdf = ids_en_excel - ids_encontrados

    # Resumen
    print(f"\n{'─'*40}")
    print("Resumen del proceso")
    print(f"{'─'*40}")
    print(f"  PDFs revisados       : {total}")
    print(f"  Copiados             : {copiados}")
    print(f"  Sin coincidencia     : {sin_coincidencia}")
    print(f"  En Excel sin PDF     : {len(ids_sin_pdf)}")

    if conteo_resp:
        print("\nPor responsable:")
        for resp in sorted(conteo_resp):
            print(f"  {resp}: {conteo_resp[resp]} archivo(s)")
            for grupo in sorted(conteo_resp_grupo[resp]):
                print(f"    — {grupo}: {conteo_resp_grupo[resp][grupo]}")

    # Reporte CSV opcional
    if args.report:
        import csv
        filas = [
            {"métrica": "total_pdfs",            "valor": total},
            {"métrica": "copiados",               "valor": copiados},
            {"métrica": "sin_coincidencia",       "valor": sin_coincidencia},
            {"métrica": "en_excel_sin_pdf",       "valor": len(ids_sin_pdf)},
        ]
        for resp, n in sorted(conteo_resp.items()):
            filas.append({"métrica": "responsable", "valor": resp, "extra": n})
        for grupo, n in sorted(conteo_grupo.items()):
            filas.append({"métrica": "grupo", "valor": grupo, "extra": n})

        ruta_reporte = Path(args.report).expanduser()
        with open(ruta_reporte, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=["métrica", "valor", "extra"])
            writer.writeheader()
            writer.writerows(filas)
        print(f"\nReporte guardado en: {ruta_reporte}")

    print(f"{'─'*40}")


if __name__ == "__main__":
    main()
