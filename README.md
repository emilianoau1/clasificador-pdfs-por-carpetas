# Clasificador automático de PDFs por responsable y grupo

Herramienta de línea de comandos que organiza automáticamente cientos de
archivos PDF en carpetas estructuradas, usando un Excel como mapa de referencia.

Ideal para empresas que reciben grandes volúmenes de documentos (contratos,
reportes, facturas, cartas) y necesitan distribuirlos por área, ejecutivo o cliente.

---

## Problema que resuelve

Cuando se tienen 200, 500 o más archivos PDF en una sola carpeta y necesitan
distribuirse por responsable y grupo, hacerlo a mano implica:

- Abrir el Excel para saber a quién corresponde cada archivo
- Crear carpetas por ejecutivo y subgrupo
- Mover o copiar cada PDF al lugar correcto
- Llevar un registro de qué se procesó y qué faltó

Este script hace todo eso en segundos, con un solo comando.

---

## Qué hace el script

1. Lee un Excel con tres columnas: identificador, grupo y responsable
2. Detecta automáticamente los nombres de columnas (sin configuración manual)
3. Crea la estructura de carpetas: `Salida / Responsable / Grupo /`
4. Copia cada PDF a su carpeta correspondiente según el identificador en el nombre
5. Reporta en consola cuántos se procesaron, cuántos faltaron y el detalle por responsable
6. Opcionalmente genera un reporte CSV con el resultado completo

---

## Características técnicas

- Detección automática de columnas (no requiere configurar nombres exactos)
- Normalización de identificadores (ignora mayúsculas, espacios y caracteres especiales)
- Modo `--dry-run` para simular el proceso sin mover archivos
- Manejo de duplicados en el Excel (conserva el primero)
- Nombres de carpetas seguros para Windows y Linux
- Reporte CSV opcional con métricas detalladas

---

## Tecnologías

| Librería | Uso |
|---|---|
| `pandas` | Lectura y procesamiento del Excel |
| `pathlib` | Manejo de rutas multiplataforma |
| `shutil` | Copia de archivos |
| `argparse` | Interfaz de línea de comandos |
| `re` | Validación y normalización de identificadores |
| `csv` | Generación del reporte |

---

## Estructura esperada

### Excel de referencia

| ID | Grupo | Responsable |
|---|---|---|
| ABC123 | Zona Norte | Luis Martínez |
| DEF456 | Zona Sur | Ana González |

Los nombres de columnas son detectados automáticamente.

### Carpeta de PDFs

Los archivos deben nombrarse con el identificador:
```
PDFS/
  ABC123.pdf
  DEF456.pdf
  GHI789.pdf
```

### Resultado

```
Salida/
  Luis Martínez/
    Zona Norte/
      ABC123.pdf
  Ana González/
    Zona Sur/
      DEF456.pdf
```

---

## Cómo usarlo

### Instalación

```bash
pip install pandas openpyxl
```

### Uso básico (autodetecta el Excel en la carpeta de PDFs)

```bash
python clasificador_pdfs.py --pdf-dir "C:\mis_pdfs"
```

### Simular sin copiar archivos

```bash
python clasificador_pdfs.py --pdf-dir "C:\mis_pdfs" --dry-run
```

### Especificar Excel y carpeta de salida

```bash
python clasificador_pdfs.py \
  --pdf-dir "C:\mis_pdfs" \
  --excel "C:\datos\clientes.xlsx" \
  --output-dir "C:\clasificados"
```

### Generar reporte CSV

```bash
python clasificador_pdfs.py --pdf-dir "C:\mis_pdfs" --report reporte.csv
```

### Ejemplo de salida en consola

```
Excel:   C:\datos\clientes.xlsx
PDFs:    C:\mis_pdfs
Salida:  C:\clasificados

────────────────────────────────────────
Resumen del proceso
────────────────────────────────────────
  PDFs revisados       : 312
  Copiados             : 298
  Sin coincidencia     : 14
  En Excel sin PDF     : 6

Por responsable:
  Ana González: 145 archivo(s)
    — Zona Norte: 89
    — Zona Sur: 56
  Luis Martínez: 153 archivo(s)
    — Zona Centro: 98
    — Zona Occidente: 55
────────────────────────────────────────
```

---

## Casos de uso

- Clasificación de contratos por ejecutivo de cuenta
- Distribución de reportes de inspección por área
- Organización de facturas por proveedor o cliente
- Cualquier flujo donde un Excel define a quién pertenece cada documento

---

## Autor

Desarrollado para automatizar la distribución de documentos en empresa
del sector financiero-automotriz. Adaptable a cualquier industria o proceso.
