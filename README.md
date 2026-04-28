# Estadincho-Gen v5.1

> Convierte cualquier archivo de datos en un dashboard visual interactivo con un solo clic.

![Python](https://img.shields.io/badge/Python-3.8%2B-3776AB?style=flat-square&logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/pandas-required-150458?style=flat-square&logo=pandas&logoColor=white)
![Chart.js](https://img.shields.io/badge/Chart.js-4.4-FF6384?style=flat-square&logo=chartdotjs&logoColor=white)
![License](https://img.shields.io/badge/licencia-MIT-06d6a0?style=flat-square)
![Version](https://img.shields.io/badge/versión-5.1-7209b7?style=flat-square)
![Platform](https://img.shields.io/badge/plataforma-Windows%20%7C%20Linux%20%7C%20macOS-555?style=flat-square)

---

## ¿Qué hace?

Estadincho-Gen analiza automáticamente tu archivo de datos y genera un dashboard HTML interactivo con:

- **KPIs automáticos** — totales, promedios y score de calidad de datos
- **Gráficas inteligentes** — línea temporal, donut, barras horizontales y top categorías
- **Filtros interactivos en tiempo real** — por fecha, categoría (hasta 4 dropdowns), rango numérico (sliders) y búsqueda de texto libre
- **Resumen ejecutivo** — párrafo generado automáticamente con los hallazgos principales
- **Calidad de datos** — score por variable con detección de nulos y outliers (IQR + Z-score)
- **Matriz de correlación** — heatmap renderizado sobre canvas
- **Título editable** — haz clic en el título del dashboard para personalizarlo sin regenerar el archivo
- **Exportación a PDF** — directamente desde el navegador
- **Copiar gráficas** — cada gráfica tiene un botón para copiarla al portapapeles como imagen

---

## Demo

El siguiente ejemplo usa un archivo real de control de trabajos de grado con **1.365 registros** y **7 columnas**. El dashboard se generó en segundos sin ninguna configuración manual.

**Panel de filtros, KPIs y gráficas principales:**

![Panel superior con filtros por categoría, KPI de 1365 registros, calidad 100/100, gráfica de evolución temporal y distribución por programa](https://raw.githubusercontent.com/REKOL08/estadincho-gen-v5/main/docs/demo1.png)

**Distribución por múltiples variables — donut, barras y comparativos:**

![Gráficas de distribución por autores, autorización, URLs y tipo de contenido adjunto](https://raw.githubusercontent.com/REKOL08/estadincho-gen-v5/main/docs/demo2.png)

**Tabla de distribución con niveles de participación calculados automáticamente:**

![Tabla con programas académicos ordenados por cantidad, porcentaje y nivel Alto/Medio/Bajo](https://raw.githubusercontent.com/REKOL08/estadincho-gen-v5/main/docs/demo3.png)

> Para agregar las imágenes al repositorio, sube las capturas a una carpeta `/docs` en la raíz del repo y GitHub las servirá automáticamente en el README.

---

## Formatos soportados

| Formato | Extensión |
|--------|-----------|
| Excel | `.xlsx` `.xls` `.xlsm` |
| CSV / TSV | `.csv` `.tsv` |
| OpenDocument | `.ods` |
| SPSS | `.sav` |
| Stata | `.dta` |
| R | `.rds` `.RData` |

---

## Requisitos

```bash
pip install pandas openpyxl
```

Para archivos `.sav` / `.dta`:
```bash
pip install pyreadstat
```

Para archivos `.rds` / `.RData`:
```bash
pip install pyreadr
```

---

## Uso

### Modo directo
```bash
python generar_dashboard.py archivo.xlsx
python generar_dashboard.py datos.csv
python generar_dashboard.py encuesta.sav
```

### Con el .bat (Windows)
Arrastra tu archivo sobre `generar_dashboard (3).bat` — no necesitas abrir la terminal.

El script ejecuta estos pasos automáticamente:

1. Detecta el formato y carga el archivo
2. Solicita un título personalizado (opcional — Enter para usar el nombre del archivo)
3. Infiere el tipo de cada columna: fecha, numérico o categórico
4. Calcula KPIs, gráficas, correlaciones y calidad de datos
5. Genera `dashboard_<nombre>.html` en la misma carpeta del archivo de entrada
6. Abre el dashboard directamente en el navegador

---

## Output

El resultado es un archivo `.html` autocontenido que incluye todos los datos embebidos. No requiere servidor, funciona completamente offline y se puede compartir por correo o alojar en cualquier hosting estático.

El dashboard está optimizado para datasets de hasta **50.000 filas** en los filtros interactivos. Archivos más grandes se procesan correctamente para las gráficas estáticas, con aviso en consola.

---

## Autor

**REKOL08** · [github.com/REKOL08](https://github.com/REKOL08)
