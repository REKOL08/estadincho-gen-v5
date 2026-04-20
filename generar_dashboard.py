#!/usr/bin/env python3
"""
Estadincho-Gen v5.1 - Convierte cualquier archivo de datos en un dashboard visual interactivo.
Formatos soportados: .xlsx, .xls, .xlsm, .csv, .tsv, .ods, .sav, .dta, .rds, .RData

Uso: python generar_dashboard.py archivo.xlsx
     python generar_dashboard.py datos.csv
"""

import sys
import os
import json
import webbrowser
import re
from pathlib import Path
from datetime import datetime

try:
    import pandas as pd
except ImportError:
    print("ERROR: Falta la libreria pandas.")
    print("Ejecuta: pip install pandas openpyxl")
    input("Presiona Enter para cerrar...")
    sys.exit(1)

# ─────────────────────────────────────────────────────────────────
# UTILIDADES
# ─────────────────────────────────────────────────────────────────

def limpiar_nombre(col):
    return str(col).strip()

def es_numerica(series):
    return pd.api.types.is_numeric_dtype(series)

def es_texto_col(series):
    return pd.api.types.is_string_dtype(series) or series.dtype == object

def es_fecha(series):
    if pd.api.types.is_datetime64_any_dtype(series):
        return True
    if es_texto_col(series):
        sample = series.dropna().head(20).astype(str)
        hits = sample.str.match(r'\d{1,4}[-/]\d{1,2}[-/]\d{1,4}').sum()
        return hits > len(sample) * 0.5
    return False

def es_categorica(series, umbral=0.5):
    if es_texto_col(series):
        return True
    if pd.api.types.is_integer_dtype(series):
        ratio = series.nunique() / max(len(series), 1)
        return ratio < umbral and series.nunique() < 50
    return False

def top_n(series, n=10):
    clean = series.dropna()
    vc = clean.value_counts().head(n)
    return {"labels": [str(x) for x in vc.index.tolist()],
            "values": [int(x) for x in vc.values.tolist()]}

def serie_temporal(df, col_fecha, col_valor=None):
    df2 = df.copy()
    df2["__mes__"] = pd.to_datetime(df2[col_fecha], errors="coerce").dt.to_period("M")
    df2 = df2.dropna(subset=["__mes__"])
    if len(df2) == 0:
        return None
    if col_valor and es_numerica(df[col_valor]):
        grp = df2.groupby("__mes__")[col_valor].sum()
    else:
        grp = df2.groupby("__mes__").size()
    grp = grp.sort_index()
    return {
        "labels": [str(p) for p in grp.index],
        "values": [float(v) for v in grp.values]
    }

# ─────────────────────────────────────────────────────────────────
# CARGA DEL ARCHIVO
# ─────────────────────────────────────────────────────────────────

def cargar_archivo(ruta):
    ruta = Path(ruta)
    if not ruta.exists():
        print(f"ERROR: No se encontro el archivo: {ruta}")
        input("Presiona Enter para cerrar...")
        sys.exit(1)

    ext = ruta.suffix.lower()
    print(f"Cargando archivo: {ruta.name} ...")

    var_labels = {}
    val_labels = {}

    try:
        if ext in [".xlsx", ".xls", ".xlsm"]:
            xl = pd.ExcelFile(ruta)
            mejor = None
            for hoja in xl.sheet_names:
                try:
                    tmp = xl.parse(hoja)
                    if mejor is None or len(tmp) > len(mejor):
                        mejor = tmp
                except:
                    pass
            if mejor is None:
                raise ValueError("No se pudo leer ninguna hoja.")
            df = mejor

        elif ext == ".ods":
            df = pd.read_excel(ruta, engine="odf")

        elif ext == ".csv":
            df = None
            for sep in [",", ";", "\t", "|"]:
                for enc in ["utf-8", "latin-1"]:
                    try:
                        tmp = pd.read_csv(ruta, sep=sep, encoding=enc, on_bad_lines="skip")
                        if tmp.shape[1] > 1:
                            df = tmp
                            break
                    except:
                        continue
                if df is not None:
                    break
            if df is None:
                raise ValueError("No se pudo detectar el separador del CSV.")

        elif ext == ".tsv":
            df = None
            for enc in ["utf-8", "latin-1"]:
                try:
                    df = pd.read_csv(ruta, sep="\t", encoding=enc, on_bad_lines="skip")
                    break
                except:
                    continue
            if df is None:
                raise ValueError("No se pudo leer el archivo TSV.")

        elif ext == ".sav":
            try:
                import pyreadstat
                df, meta = pyreadstat.read_sav(str(ruta), apply_value_formats=False)
                var_labels = meta.column_labels_and_names
                val_labels = meta.variable_value_labels
                for col, mapping in val_labels.items():
                    if col in df.columns:
                        df[col] = df[col].map(lambda x: mapping.get(x, x))
            except ImportError:
                print("AVISO: pyreadstat no instalado. Ejecuta: pip install pyreadstat")
                input("Presiona Enter para cerrar...")
                sys.exit(1)

        elif ext == ".dta":
            try:
                import pyreadstat
                df, meta = pyreadstat.read_dta(str(ruta), apply_value_formats=False)
                var_labels = meta.column_labels_and_names
                val_labels = meta.variable_value_labels
                for col, mapping in val_labels.items():
                    if col in df.columns:
                        df[col] = df[col].map(lambda x: mapping.get(x, x))
            except ImportError:
                print("AVISO: pyreadstat no instalado. Ejecuta: pip install pyreadstat")
                input("Presiona Enter para cerrar...")
                sys.exit(1)

        elif ext in [".rds", ".rdata"]:
            try:
                import pyreadr
                result = pyreadr.read_r(str(ruta))
                df = max(result.values(), key=lambda x: len(x) if hasattr(x, '__len__') else 0)
            except ImportError:
                print("AVISO: pyreadr no instalado. Ejecuta: pip install pyreadr")
                input("Presiona Enter para cerrar...")
                sys.exit(1)

        else:
            print(f"ERROR: Formato no soportado: {ext}")
            print("Formatos soportados: .xlsx .xls .xlsm .csv .tsv .ods .sav .dta .rds .RData")
            input("Presiona Enter para cerrar...")
            sys.exit(1)

        if var_labels:
            rename_map = {}
            for col in df.columns:
                etiqueta = var_labels.get(col, "")
                if etiqueta and etiqueta != col:
                    rename_map[col] = etiqueta
            if rename_map:
                df = df.rename(columns=rename_map)

        df.columns = [limpiar_nombre(c) for c in df.columns]
        df = df.dropna(how="all").reset_index(drop=True)
        print(f"  -> {len(df)} filas, {len(df.columns)} columnas cargadas.")
        return df, ruta.stem

    except Exception as e:
        print(f"ERROR al leer el archivo: {e}")
        input("Presiona Enter para cerrar...")
        sys.exit(1)

# ─────────────────────────────────────────────────────────────────
# FASE 3 — CALIDAD DE DATOS
# ─────────────────────────────────────────────────────────────────

def calcular_calidad(df):
    cols_num = [c for c in df.columns if es_numerica(df[c])]
    if not cols_num:
        return [], 100

    total = len(df)
    filas = []
    scores = []

    for col in cols_num:
        s = df[col]
        nulos = int(s.isna().sum())
        pct_nulos = nulos / total * 100
        s_clean = s.dropna()
        n = len(s_clean)

        if n >= 4:
            q1 = s_clean.quantile(0.25)
            q3 = s_clean.quantile(0.75)
            iqr = q3 - q1
            out_iqr = int(((s_clean < q1 - 1.5 * iqr) | (s_clean > q3 + 1.5 * iqr)).sum())
        else:
            out_iqr = 0

        if n >= 4:
            mean = s_clean.mean()
            std = s_clean.std()
            if std > 0:
                zscores = ((s_clean - mean) / std).abs()
                out_z = int((zscores > 3).sum())
            else:
                out_z = 0
        else:
            out_z = 0

        penalizacion = pct_nulos * 0.7 + (out_iqr / max(n, 1)) * 100 * 0.3
        score = max(0, round(100 - penalizacion))
        scores.append(score)

        if score >= 80:
            estado = "Buena"
        elif score >= 50:
            estado = "Regular"
        else:
            estado = "Revisar"

        filas.append({
            "col": col,
            "nulos": nulos,
            "pct_nulos": f"{pct_nulos:.1f}%",
            "out_iqr": out_iqr,
            "out_z": out_z,
            "score": score,
            "estado": estado
        })

    score_global = round(sum(scores) / len(scores)) if scores else 100
    return filas, score_global

# ─────────────────────────────────────────────────────────────────
# FASE 3 — MATRIZ DE CORRELACIÓN
# ─────────────────────────────────────────────────────────────────

def calcular_correlacion(df):
    cols_num = [c for c in df.columns if es_numerica(df[c]) and df[c].dropna().nunique() > 1]
    if len(cols_num) < 2:
        return None
    cols_num = cols_num[:10]
    corr = df[cols_num].corr().round(2)
    matrix = []
    for _, row in corr.iterrows():
        matrix.append([None if pd.isna(v) else float(v) for v in row])
    return {
        "labels": cols_num,
        "matrix": matrix
    }

# ─────────────────────────────────────────────────────────────────
# FASE 4A — DATOS PARA FILTROS INTERACTIVOS
# ─────────────────────────────────────────────────────────────────

def exportar_dataset_json(df):
    """
    Serializa el DataFrame completo como lista de dicts para el frontend.
    Convierte fechas a string ISO y NaN a None para JSON válido.
    Limita a 50,000 filas para no romper el navegador.
    """
    MAX_FILAS = 50_000
    if len(df) > MAX_FILAS:
        print(f"  AVISO: Dataset tiene {len(df):,} filas. Se usarán las primeras {MAX_FILAS:,} para los filtros.")
        df = df.head(MAX_FILAS)

    df_copy = df.copy()
    for col in df_copy.columns:
        if pd.api.types.is_datetime64_any_dtype(df_copy[col]):
            df_copy[col] = df_copy[col].dt.strftime("%Y-%m-%d")

    records = df_copy.where(pd.notnull(df_copy), None).to_dict(orient="records")
    return json.dumps(records, ensure_ascii=False, default=str)


def metadata_filtros(df):
    """
    Retorna metadata para construir los controles de filtro en el HTML:
    - col_fecha: nombre de la columna fecha (o None)
    - cols_cat: lista de {col, valores_unicos} para dropdowns
    - cols_num: lista de {col, min, max} para sliders
    - cols_texto: lista de columnas de texto para búsqueda libre
    """
    col_fecha = None
    cols_cat = []
    cols_num = []
    cols_texto = []

    for c in df.columns:
        s = df[c].dropna()
        if len(s) == 0:
            continue
        if es_fecha(s):
            if col_fecha is None:
                col_fecha = c
        elif es_numerica(s):
            cols_num.append({
                "col": c,
                "min": float(s.min()),
                "max": float(s.max())
            })
        elif es_categorica(s):
            vals = [str(v) for v in s.value_counts().head(50).index.tolist()]
            cols_cat.append({"col": c, "valores": vals})
        if es_texto_col(df[c]):
            cols_texto.append(c)

    return {
        "col_fecha": col_fecha,
        "cols_cat": cols_cat[:4],      # máx 4 dropdowns
        "cols_num": cols_num[:3],      # máx 3 sliders
        "cols_texto": cols_texto[:4]   # máx 4 búsquedas
    }


# ─────────────────────────────────────────────────────────────────
# RESUMEN EJECUTIVO AUTOMÁTICO
# ─────────────────────────────────────────────────────────────────

def generar_resumen(df, datos, nombre_archivo):
    total = len(df)
    n_cols = len(df.columns)
    score = None
    for kpi in datos.get("kpis", []):
        if kpi.get("clase") == "calidad":
            try:
                score = int(kpi["value"].split("/")[0])
            except:
                pass

    partes = []
    partes.append(f"El archivo <strong>{nombre_archivo}</strong> contiene <strong>{total:,} registros</strong> distribuidos en {n_cols} variables.")

    # Calidad
    if score is not None:
        if score >= 80:
            partes.append(f"La calidad general de los datos es <strong>buena ({score}/100)</strong>, lo que indica que el conjunto está en condiciones adecuadas para el análisis.")
        elif score >= 50:
            partes.append(f"La calidad general de los datos es <strong>regular ({score}/100)</strong>. Se recomienda revisar los valores nulos y los valores atípicos antes de sacar conclusiones.")
        else:
            partes.append(f"La calidad general de los datos es <strong>baja ({score}/100)</strong>. Es necesario limpiar el conjunto de datos antes de analizarlo.")

    # Variables numéricas destacadas
    cols_num = [c for c in df.columns if es_numerica(df[c])]
    if cols_num:
        c = cols_num[0]
        s = df[c].dropna()
        partes.append(f"La variable numérica principal (<em>{c}</em>) presenta un total acumulado de <strong>{s.sum():,.1f}</strong> y un promedio de <strong>{s.mean():,.1f}</strong> por registro.")

    # Variable categórica principal
    cols_cat = [c for c in df.columns if es_categorica(df[c]) and not es_fecha(df[c])]
    if cols_cat:
        c = cols_cat[0]
        top_val = df[c].dropna().value_counts().idxmax()
        top_cnt = df[c].dropna().value_counts().iloc[0]
        top_pct = top_cnt / total * 100
        partes.append(f"En la variable categórica <em>{c}</em>, la categoría más frecuente es <strong>\"{top_val}\"</strong>, representando el {top_pct:.1f}% del total ({top_cnt:,} registros).")

    # Correlaciones destacadas
    corr_data = datos.get("correlacion")
    if corr_data and len(corr_data["labels"]) >= 2:
        labels = corr_data["labels"]
        matrix = corr_data["matrix"]
        max_corr = 0
        par = ("", "")
        for i in range(len(labels)):
            for j in range(i+1, len(labels)):
                v = matrix[i][j]
                if v is not None and abs(v) > abs(max_corr):
                    max_corr = v
                    par = (labels[i], labels[j])
        if abs(max_corr) >= 0.5:
            tipo = "positiva" if max_corr > 0 else "negativa"
            fuerza = "fuerte" if abs(max_corr) >= 0.75 else "moderada"
            partes.append(f"Se detectó una correlación {tipo} {fuerza} (<strong>r = {max_corr:.2f}</strong>) entre <em>{par[0]}</em> y <em>{par[1]}</em>.")

    return " ".join(partes)

# ─────────────────────────────────────────────────────────────────
# ANÁLISIS INTELIGENTE
# ─────────────────────────────────────────────────────────────────

def analizar(df):
    cols = list(df.columns)
    total_filas = len(df)
    col_fecha = None
    cols_num = []
    cols_cat = []

    for c in cols:
        s = df[c].dropna()
        if len(s) == 0:
            continue
        if es_fecha(s):
            if col_fecha is None:
                col_fecha = c
        elif es_numerica(s):
            cols_num.append(c)
        elif es_categorica(s):
            cols_cat.append(c)

    filas_calidad, score_global = calcular_calidad(df)
    correlacion = calcular_correlacion(df)

    icono_score = "🟢" if score_global >= 80 else ("🟡" if score_global >= 50 else "🔴")
    kpis = [
        {"icon": "📊", "label": "Total Registros", "value": str(total_filas), "clase": "total"},
        {"icon": icono_score, "label": "Calidad de Datos", "value": f"{score_global}/100", "clase": "calidad"}
    ]

    for c in cols_num[:3]:
        s = df[c].dropna()
        if len(s) == 0:
            continue
        val = s.sum()
        fmt = f"{val:,.0f}" if val == int(val) else f"{val:,.2f}"
        kpis.append({"icon": "💰", "label": f"Total {c}", "value": fmt, "clase": "num"})
        prom = s.mean()
        kpis.append({"icon": "📈", "label": f"Promedio {c}", "value": f"{prom:,.1f}", "clase": "promedio"})

    graficas = []

    if col_fecha:
        col_val = cols_num[0] if cols_num else None
        data_ts = serie_temporal(df, col_fecha, col_val)
        if data_ts and len(data_ts["labels"]) >= 2:
            graficas.append({
                "id": "chartFecha",
                "titulo": f"Evolución por Período ({col_fecha})",
                "tipo": "line",
                "labels": data_ts["labels"],
                "datasets": [{"label": col_val or "Registros", "data": data_ts["values"]}],
                "ancho": "half"
            })

    tipo_ciclo = ["doughnut", "bar", "doughnut", "bar", "bar", "bar"]
    for i, c in enumerate(cols_cat[:6]):
        data_cat = top_n(df[c], n=8 if i > 0 else 5)
        if len(data_cat["labels"]) < 2:
            continue
        t = tipo_ciclo[i % len(tipo_ciclo)]
        graficas.append({
            "id": f"chartCat{i}",
            "titulo": f"Distribución por {c}",
            "tipo": t,
            "labels": data_cat["labels"],
            "datasets": [{"label": c, "data": data_cat["values"]}],
            "ancho": "half"
        })

    if cols_num and cols_cat:
        try:
            c_cat = cols_cat[0]
            c_num = cols_num[0]
            agr = df.groupby(c_cat)[c_num].sum().nlargest(10).reset_index()
            if len(agr) >= 3:
                graficas.append({
                    "id": "chartTop",
                    "titulo": f"Top 10: {c_cat} por {c_num}",
                    "tipo": "bar",
                    "labels": [str(x) for x in agr[c_cat].tolist()],
                    "datasets": [{"label": c_num, "data": [float(x) for x in agr[c_num].tolist()]}],
                    "ancho": "full"
                })
        except:
            pass

    tabla = None
    if cols_cat:
        c = cols_cat[0]
        vc = df[c].value_counts().head(12)
        total = vc.sum()
        filas = []
        for val, cnt in vc.items():
            pct = cnt / total * 100
            estado = "Alto" if pct >= 10 else ("Medio" if pct >= 5 else "Bajo")
            filas.append({
                "nombre": str(val),
                "cantidad": int(cnt),
                "porcentaje": f"{pct:.1f}%",
                "estado": estado
            })
        tabla = {"columna": c, "filas": filas}

    resultado = {
        "titulo": "",
        "subtitulo": f"Análisis automático · {total_filas:,} registros · {len(df.columns)} columnas",
        "kpis": kpis[:8],
        "graficas": graficas,
        "tabla": tabla,
        "calidad": filas_calidad,
        "correlacion": correlacion,
        "generado": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "_df": df,          # referencia temporal para exportar JSON
    }

    return resultado

# ─────────────────────────────────────────────────────────────────
# GENERACIÓN HTML
# ─────────────────────────────────────────────────────────────────

COLORS = [
    "#00b4d8","#f72585","#7209b7","#4361ee",
    "#fb8500","#06d6a0","#ffd60a","#ef476f",
    "#3a86ff","#8338ec"
]

def color_datasets(datasets, tipo):
    result = []
    for i, ds in enumerate(datasets):
        c = COLORS[i % len(COLORS)]
        entry = dict(ds)
        if tipo == "line":
            entry["borderColor"] = c
            entry["_fillColor"] = c
            entry["borderWidth"] = 3
            entry["fill"] = True
            entry["tension"] = 0.4
            entry["pointBackgroundColor"] = c
            entry["pointBorderColor"] = "#fff"
            entry["pointBorderWidth"] = 2
            entry["pointRadius"] = 6
        elif tipo in ["doughnut", "pie"]:
            entry["backgroundColor"] = COLORS[:len(ds["data"])]
            entry["borderWidth"] = 0
        else:
            if len(datasets) == 1:
                entry["backgroundColor"] = COLORS[:len(ds["data"])]
            else:
                entry["backgroundColor"] = c
            entry["borderRadius"] = 8
        result.append(entry)
    return result

def grafica_js(g):
    tipo = g["tipo"]
    datasets = color_datasets(g["datasets"], tipo)

    if tipo == "line":
        opts = """{
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        y: { beginAtZero: true, grid: { color: 'rgba(255,255,255,0.05)' } },
        x: { grid: { display: false } }
      }
    }"""
    elif tipo in ["doughnut", "pie"]:
        opts = """{
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'bottom', labels: { padding: 20 } } },
      cutout: '65%'
    }"""
    else:
        max_label = max((len(str(l)) for l in g["labels"]), default=0)
        horiz = len(g["labels"]) > 4 or max_label > 12
        axis_extra = "indexAxis: 'y'," if horiz else ""
        multi_legend = "position: 'bottom', labels: { padding: 20 }" if len(datasets) > 1 else "display: false"
        opts = f"""{{
      responsive: true, maintainAspectRatio: false,
      {axis_extra}
      plugins: {{ legend: {{ {multi_legend} }} }},
      scales: {{
        x: {{ beginAtZero: true, grid: {{ color: 'rgba(255,255,255,0.05)' }} }},
        y: {{ grid: {{ display: false }} }}
      }}
    }}"""

    ds_json = json.dumps(datasets, ensure_ascii=False)
    return f"""
  (function() {{
    var ctx = document.getElementById('{g["id"]}').getContext('2d');
    window._charts = window._charts || {{}};
    window._charts['{g["id"]}'] = new Chart(ctx, {{
      type: '{tipo}',
      data: {{
        labels: {json.dumps(g["labels"], ensure_ascii=False)},
        datasets: {ds_json}
      }},
      options: {opts}
    }});
  }})();"""

def html_calidad(filas_calidad):
    if not filas_calidad:
        return ""
    filas_html = ""
    for f in filas_calidad:
        badge = ("badge-success" if f["estado"] == "Buena"
                 else "badge-warning" if f["estado"] == "Regular"
                 else "badge-danger")
        bar_color = ("#06d6a0" if f["score"] >= 80
                     else "#fb8500" if f["score"] >= 50
                     else "#f72585")
        filas_html += f"""
      <tr>
        <td>{f["col"]}</td>
        <td>{f["nulos"]} ({f["pct_nulos"]})</td>
        <td>{f["out_iqr"]}</td>
        <td>{f["out_z"]}</td>
        <td>
          <div class="score-bar-bg">
            <div class="score-bar" style="width:{f["score"]}%;background:{bar_color}"></div>
          </div>
          <span style="font-size:.85rem;color:{bar_color}">{f["score"]}</span>
        </td>
        <td><span class="badge {badge}">{f["estado"]}</span></td>
      </tr>"""
    return f"""
  <section class="chart-card full-width" style="margin-top:25px">
    <h3>Calidad de Datos por Variable</h3>
    <table class="stats-table">
      <thead>
        <tr>
          <th>Variable</th><th>Nulos</th><th>Outliers IQR</th>
          <th>Outliers Z-score</th><th>Score</th><th>Estado</th>
        </tr>
      </thead>
      <tbody>{filas_html}</tbody>
    </table>
  </section>"""

def html_correlacion(correlacion):
    if not correlacion:
        return "", ""
    labels_json = json.dumps(correlacion["labels"], ensure_ascii=False)
    matrix_json = json.dumps(correlacion["matrix"])
    n = len(correlacion["labels"])
    html = """
  <section class="chart-card full-width" style="margin-top:25px">
    <h3>Matriz de Correlación</h3>
    <div style="overflow-x:auto">
      <canvas id="canvasCorr"></canvas>
    </div>
  </section>"""
    js = f"""
  (function() {{
    var labels = {labels_json};
    var matrix = {matrix_json};
    var n = labels.length;
    var cell = 70; var pad = 130;
    var cvs = document.getElementById('canvasCorr');
    cvs.width = pad + n * cell; cvs.height = pad + n * cell;
    var ctx = cvs.getContext('2d');
    function lerp(a,b,t){{ return a+(b-a)*t; }}
    function corrColor(v) {{
      if(v===null) return '#2a2a4a';
      var t=(v+1)/2; var r,g,b;
      if(t<0.5){{ r=Math.round(lerp(220,50,t*2)); g=Math.round(lerp(50,50,t*2)); b=Math.round(lerp(50,120,t*2)); }}
      else{{ r=Math.round(lerp(50,30,(t-0.5)*2)); g=Math.round(lerp(50,80,(t-0.5)*2)); b=Math.round(lerp(120,220,(t-0.5)*2)); }}
      return 'rgb('+r+','+g+','+b+')';
    }}
    ctx.fillStyle='#0f0f23'; ctx.fillRect(0,0,cvs.width,cvs.height);
    ctx.fillStyle='#a0a0a0'; ctx.font='11px Segoe UI'; ctx.textAlign='center';
    for(var j=0;j<n;j++){{
      ctx.save(); ctx.translate(pad+j*cell+cell/2,pad-10); ctx.rotate(-Math.PI/4);
      ctx.fillText(labels[j].substring(0,14),0,0); ctx.restore();
    }}
    ctx.textAlign='right';
    for(var i=0;i<n;i++){{
      ctx.fillStyle='#a0a0a0';
      ctx.fillText(labels[i].substring(0,16),pad-8,pad+i*cell+cell/2+4);
    }}
    ctx.textAlign='center';
    for(var i=0;i<n;i++){{
      for(var j=0;j<n;j++){{
        var v=matrix[i][j]; var x=pad+j*cell; var y=pad+i*cell;
        ctx.fillStyle=corrColor(v); ctx.beginPath();
        ctx.roundRect(x+2,y+2,cell-4,cell-4,6); ctx.fill();
        if(v!==null){{
          ctx.fillStyle=(Math.abs(v)>0.5)?'#ffffff':'#cccccc';
          ctx.font='bold 12px Segoe UI'; ctx.fillText(v.toFixed(2),x+cell/2,y+cell/2+4);
        }}
      }}
    }}
  }})();"""
    return html, js

# ─────────────────────────────────────────────────────────────────
# RESUMEN EJECUTIVO HTML
# ─────────────────────────────────────────────────────────────────

def html_resumen(resumen_texto):
    return f"""
  <section class="resumen-ejecutivo full-width">
    <div class="resumen-icono">📋</div>
    <div>
      <h3>Resumen Ejecutivo</h3>
      <p>{resumen_texto}</p>
    </div>
  </section>"""

# ─────────────────────────────────────────────────────────────────
# GENERACIÓN HTML COMPLETA (v5.0)
# ─────────────────────────────────────────────────────────────────

def generar_html(datos, nombre_archivo, titulo_personalizado=""):
    titulo = titulo_personalizado or datos["titulo"] or nombre_archivo.replace("_", " ").replace("-", " ").title()
    titulo_escaped = titulo.replace('"', '&quot;').replace("'", "&#39;")

    clases_kpi = ["total","calidad","num","promedio","activos","tiempo","retraso","tasa"]
    kpi_html = ""
    for i, k in enumerate(datos["kpis"]):
        cls = k.get("clase", clases_kpi[i % len(clases_kpi)])
        kpi_html += f"""
      <div class="kpi-card {cls}">
        <div class="icon">{k["icon"]}</div>
        <div class="value" id="kpi_{i}_val">{k["value"]}</div>
        <div class="label">{k["label"]}</div>
      </div>"""

    charts_html = ""
    charts_js = ""
    for g in datos["graficas"]:
        ancho = g.get("ancho", "half")
        tall = ' tall' if ancho == "full" else ""
        charts_html += f"""
      <div class="chart-card {ancho}-width">
        <div class="chart-header">
          <h3>{g["titulo"]}</h3>
          <button class="btn-copy-chart" onclick="copiarGrafica('{g["id"]}')" title="Copiar imagen al portapapeles">
            📋 Copiar
          </button>
        </div>
        <div class="chart-container{tall}">
          <canvas id="{g["id"]}"></canvas>
        </div>
      </div>"""
        charts_js += grafica_js(g)

    tabla_html = ""
    if datos.get("tabla"):
        t = datos["tabla"]
        filas_html = ""
        for f in t["filas"]:
            badge = ("badge-success" if f["estado"] == "Alto"
                     else "badge-warning" if f["estado"] == "Medio"
                     else "badge-danger")
            filas_html += f"""
          <tr>
            <td>{f["nombre"]}</td>
            <td>{f["cantidad"]}</td>
            <td>{f["porcentaje"]}</td>
            <td><span class="badge {badge}">{f["estado"]}</span></td>
          </tr>"""
        tabla_html = f"""
      <section class="chart-card" style="margin-top:25px">
        <h3>Distribución por {t["columna"]}</h3>
        <table class="stats-table" id="tablaCategoria">
          <thead>
            <tr><th>{t["columna"]}</th><th>Cantidad</th><th>Porcentaje</th><th>Nivel</th></tr>
          </thead>
          <tbody id="tablaCategoriaCuerpo">{filas_html}</tbody>
        </table>
      </section>"""

    calidad_html = html_calidad(datos.get("calidad", []))
    corr_html, corr_js = html_correlacion(datos.get("correlacion"))
    resumen_html = html_resumen(datos.get("resumen", ""))

    # ── Panel de filtros ──────────────────────────────────────────
    meta = datos.get("meta_filtros", {})
    col_fecha = meta.get("col_fecha")
    cols_cat  = meta.get("cols_cat", [])
    cols_num  = meta.get("cols_num", [])
    cols_texto = meta.get("cols_texto", [])

    filtros_controles = ""

    # Rango de fechas
    if col_fecha:
        filtros_controles += f"""
        <div class="filtro-grupo">
          <label class="filtro-label">📅 {col_fecha}</label>
          <div class="filtro-fechas">
            <input type="date" id="filtroFechaDesde" class="filtro-input" placeholder="Desde" onchange="aplicarFiltros()">
            <span style="color:var(--text-secondary)">→</span>
            <input type="date" id="filtroFechaHasta" class="filtro-input" placeholder="Hasta" onchange="aplicarFiltros()">
          </div>
        </div>"""

    # Dropdowns de categorías
    for item in cols_cat:
        col = item["col"]
        col_id = re.sub(r'[^a-zA-Z0-9]', '_', col)
        opciones = '<option value="">Todas</option>' + "".join(
            f'<option value="{v}">{v}</option>' for v in item["valores"]
        )
        filtros_controles += f"""
        <div class="filtro-grupo">
          <label class="filtro-label">🏷️ {col}</label>
          <select id="filtrocat_{col_id}" class="filtro-select" data-col="{col}" onchange="aplicarFiltros()">
            {opciones}
          </select>
        </div>"""

    # Sliders numéricos
    for item in cols_num:
        col = item["col"]
        col_id = re.sub(r'[^a-zA-Z0-9]', '_', col)
        mn = item["min"]
        mx = item["max"]
        filtros_controles += f"""
        <div class="filtro-grupo">
          <label class="filtro-label">🔢 {col}</label>
          <div class="filtro-slider-wrap">
            <input type="range" id="filtronum_{col_id}_min" class="filtro-slider"
                   data-col="{col}" data-role="min"
                   min="{mn}" max="{mx}" step="{max((mx-mn)/100, 0.01):.4f}" value="{mn}"
                   oninput="actualizarSlider('{col_id}'); aplicarFiltros()">
            <input type="range" id="filtronum_{col_id}_max" class="filtro-slider"
                   data-col="{col}" data-role="max"
                   min="{mn}" max="{mx}" step="{max((mx-mn)/100, 0.01):.4f}" value="{mx}"
                   oninput="actualizarSlider('{col_id}'); aplicarFiltros()">
            <div class="filtro-slider-labels">
              <span id="filtronum_{col_id}_minval">{mn:,.1f}</span>
              <span id="filtronum_{col_id}_maxval">{mx:,.1f}</span>
            </div>
          </div>
        </div>"""

    # Búsqueda de texto
    if cols_texto:
        cols_texto_json = json.dumps(cols_texto, ensure_ascii=False)
        filtros_controles += f"""
        <div class="filtro-grupo filtro-busqueda">
          <label class="filtro-label">🔍 Búsqueda en tabla</label>
          <input type="text" id="filtroTexto" class="filtro-input" placeholder="Buscar en los datos..."
                 oninput="aplicarFiltros()" autocomplete="off">
          <span class="filtro-hint" id="filtroTextoHint"></span>
        </div>"""
    else:
        cols_texto_json = "[]"

    panel_filtros = f"""
    <div id="panelFiltros" class="panel-filtros">
      <div class="filtros-inner">
        <div class="filtros-titulo">
          <span>⚙️ Filtros</span>
          <span id="filtrosActivos" class="badge-filtros-activos" style="display:none"></span>
        </div>
        <div class="filtros-controles">
          {filtros_controles}
        </div>
        <button class="btn-limpiar" onclick="limpiarFiltros()">✖ Limpiar filtros</button>
        <div class="filtro-contador" id="filtroContador"></div>
      </div>
    </div>"""

    # ── JS de filtros ─────────────────────────────────────────────
    dataset_json = datos.get("dataset_json", "[]")
    meta_json    = json.dumps(meta, ensure_ascii=False)
    graficas_meta = json.dumps([
        {"id": g["id"], "tipo": g["tipo"],
         "col_x": g.get("col_x",""), "col_y": g.get("col_y","")}
        for g in datos["graficas"]
    ], ensure_ascii=False)
    kpis_meta = json.dumps([
        {"label": k["label"], "clase": k.get("clase",""), "icon": k["icon"]}
        for k in datos["kpis"]
    ], ensure_ascii=False)
    calidad_original = json.dumps(datos.get("calidad", []), ensure_ascii=False)
    tabla_col = datos["tabla"]["columna"] if datos.get("tabla") else ""

    js_filtros = f"""
  // ── FASE 4A: FILTROS INTERACTIVOS ──────────────────────────────
  var _DATASET     = {dataset_json};
  var _META        = {meta_json};
  var _KPIS_META   = {kpis_meta};
  var _CALIDAD_ORI = {calidad_original};
  var _TABLA_COL   = {json.dumps(tabla_col)};
  var _COLS_TEXTO  = {cols_texto_json};
  var _dfFiltrado  = _DATASET.slice();

  function actualizarSlider(colId) {{
    var mn = parseFloat(document.getElementById('filtronum_' + colId + '_min').value);
    var mx = parseFloat(document.getElementById('filtronum_' + colId + '_max').value);
    if (mn > mx) {{
      document.getElementById('filtronum_' + colId + '_min').value = mx;
      mn = mx;
    }}
    document.getElementById('filtronum_' + colId + '_minval').textContent = mn.toLocaleString('es-CO', {{maximumFractionDigits:1}});
    document.getElementById('filtronum_' + colId + '_maxval').textContent = mx.toLocaleString('es-CO', {{maximumFractionDigits:1}});
  }}

  function aplicarFiltros() {{
    var df = _DATASET.slice();

    // Filtro fecha
    var colFecha = _META.col_fecha;
    if (colFecha) {{
      var desde = document.getElementById('filtroFechaDesde') && document.getElementById('filtroFechaDesde').value;
      var hasta = document.getElementById('filtroFechaHasta') && document.getElementById('filtroFechaHasta').value;
      if (desde) df = df.filter(function(r) {{ return r[colFecha] && r[colFecha] >= desde; }});
      if (hasta) df = df.filter(function(r) {{ return r[colFecha] && r[colFecha] <= hasta; }});
    }}

    // Filtros categoría
    (_META.cols_cat || []).forEach(function(item) {{
      var colId = item.col.replace(/[^a-zA-Z0-9]/g, '_');
      var sel   = document.getElementById('filtrocat_' + colId);
      if (sel && sel.value) {{
        df = df.filter(function(r) {{ return String(r[item.col]) === sel.value; }});
      }}
    }});

    // Filtros numéricos
    (_META.cols_num || []).forEach(function(item) {{
      var colId = item.col.replace(/[^a-zA-Z0-9]/g, '_');
      var elMin = document.getElementById('filtronum_' + colId + '_min');
      var elMax = document.getElementById('filtronum_' + colId + '_max');
      if (elMin && elMax) {{
        var mn = parseFloat(elMin.value);
        var mx = parseFloat(elMax.value);
        df = df.filter(function(r) {{
          var v = parseFloat(r[item.col]);
          return !isNaN(v) && v >= mn && v <= mx;
        }});
      }}
    }});

    // Búsqueda texto
    var txt = document.getElementById('filtroTexto') && document.getElementById('filtroTexto').value.toLowerCase().trim();
    if (txt) {{
      df = df.filter(function(r) {{
        return _COLS_TEXTO.some(function(c) {{
          return r[c] && String(r[c]).toLowerCase().includes(txt);
        }});
      }});
      var hint = document.getElementById('filtroTextoHint');
      if (hint) hint.textContent = df.length + ' coincidencias';
    }} else {{
      var hint = document.getElementById('filtroTextoHint');
      if (hint) hint.textContent = '';
    }}

    _dfFiltrado = df;
    actualizarContador(df.length);
    actualizarKPIs(df);
    actualizarTablaCalidad(df);
    actualizarGraficas(df);
    actualizarTablaCategoria(df);
    actualizarBadgeFiltros();
  }}

  function actualizarContador(n) {{
    var el = document.getElementById('filtroContador');
    if (!el) return;
    var total = _DATASET.length;
    var pct   = total > 0 ? ((n / total) * 100).toFixed(1) : 0;
    el.innerHTML = '<strong>' + n.toLocaleString('es-CO') + '</strong> de ' + total.toLocaleString('es-CO') + ' registros (' + pct + '%)';
  }}

  function actualizarBadgeFiltros() {{
    var badge = document.getElementById('filtrosActivos');
    if (!badge) return;
    var activos = 0;
    if (_META.col_fecha) {{
      if ((document.getElementById('filtroFechaDesde') || {{}}).value) activos++;
      if ((document.getElementById('filtroFechaHasta') || {{}}).value) activos++;
    }}
    (_META.cols_cat || []).forEach(function(item) {{
      var colId = item.col.replace(/[^a-zA-Z0-9]/g, '_');
      var sel = document.getElementById('filtrocat_' + colId);
      if (sel && sel.value) activos++;
    }});
    (_META.cols_num || []).forEach(function(item) {{
      var colId = item.col.replace(/[^a-zA-Z0-9]/g, '_');
      var elMin = document.getElementById('filtronum_' + colId + '_min');
      var elMax = document.getElementById('filtronum_' + colId + '_max');
      if (elMin && elMax && (parseFloat(elMin.value) !== item.min || parseFloat(elMax.value) !== item.max)) activos++;
    }});
    var txt = document.getElementById('filtroTexto');
    if (txt && txt.value.trim()) activos++;
    if (activos > 0) {{
      badge.style.display = 'inline-block';
      badge.textContent   = activos + ' activo' + (activos > 1 ? 's' : '');
    }} else {{
      badge.style.display = 'none';
    }}
  }}

  function actualizarKPIs(df) {{
    var total = df.length;
    // KPI 0: total registros
    var el0 = document.getElementById('kpi_0_val');
    if (el0) el0.textContent = total.toLocaleString('es-CO');
    // KPIs numéricos (sum / promedio)
    var numIdx = 0;
    _KPIS_META.forEach(function(k, i) {{
      var el = document.getElementById('kpi_' + i + '_val');
      if (!el) return;
      if (k.clase === 'num') {{
        var col = k.label.replace('Total ', '');
        var vals = df.map(function(r) {{ return parseFloat(r[col]); }}).filter(function(v) {{ return !isNaN(v); }});
        var suma = vals.reduce(function(a,b){{return a+b;}}, 0);
        el.textContent = suma % 1 === 0 ? suma.toLocaleString('es-CO') : suma.toLocaleString('es-CO', {{maximumFractionDigits:2}});
      }} else if (k.clase === 'promedio') {{
        var col = k.label.replace('Promedio ', '');
        var vals = df.map(function(r) {{ return parseFloat(r[col]); }}).filter(function(v) {{ return !isNaN(v); }});
        var prom = vals.length > 0 ? vals.reduce(function(a,b){{return a+b;}}, 0) / vals.length : 0;
        el.textContent = prom.toLocaleString('es-CO', {{maximumFractionDigits:1}});
      }}
    }});
  }}

  function actualizarTablaCalidad(df) {{
    // Recalcula nulos sobre el subset filtrado
    var body = document.getElementById('cuerpoCalidad');
    if (!body || !_CALIDAD_ORI.length) return;
    var total = df.length;
    if (total === 0) return;
    var filas = body.querySelectorAll('tr');
    _CALIDAD_ORI.forEach(function(orig, idx) {{
      var tr = filas[idx];
      if (!tr) return;
      var col = orig.col;
      var nulos = df.filter(function(r) {{ return r[col] === null || r[col] === undefined || r[col] === ''; }}).length;
      var pctN  = (nulos / total * 100).toFixed(1);
      var vals  = df.map(function(r) {{ return parseFloat(r[col]); }}).filter(function(v) {{ return !isNaN(v); }});
      var n     = vals.length;
      var out_iqr = 0;
      if (n >= 4) {{
        vals.sort(function(a,b){{return a-b;}});
        var q1 = vals[Math.floor(n*0.25)], q3 = vals[Math.floor(n*0.75)];
        var iqr = q3 - q1;
        out_iqr = vals.filter(function(v){{ return v < q1-1.5*iqr || v > q3+1.5*iqr; }}).length;
      }}
      var mean = n > 0 ? vals.reduce(function(a,b){{return a+b;}},0)/n : 0;
      var std  = n > 1 ? Math.sqrt(vals.reduce(function(a,v){{return a+Math.pow(v-mean,2);}},0)/(n-1)) : 0;
      var out_z = std > 0 ? vals.filter(function(v){{ return Math.abs((v-mean)/std) > 3; }}).length : 0;
      var pen   = (nulos/total)*100*0.7 + (out_iqr/Math.max(n,1))*100*0.3;
      var score = Math.max(0, Math.round(100 - pen));
      var estado = score >= 80 ? 'Buena' : (score >= 50 ? 'Regular' : 'Revisar');
      var badgeC = score >= 80 ? 'badge-success' : (score >= 50 ? 'badge-warning' : 'badge-danger');
      var barC   = score >= 80 ? '#06d6a0' : (score >= 50 ? '#fb8500' : '#f72585');
      var tds = tr.querySelectorAll('td');
      if (tds[1]) tds[1].textContent = nulos + ' (' + pctN + '%)';
      if (tds[2]) tds[2].textContent = out_iqr;
      if (tds[3]) tds[3].textContent = out_z;
      if (tds[4]) tds[4].innerHTML   = '<div class="score-bar-bg"><div class="score-bar" style="width:'+score+'%;background:'+barC+'"></div></div><span style="font-size:.85rem;color:'+barC+'">'+score+'</span>';
      if (tds[5]) tds[5].innerHTML   = '<span class="badge '+badgeC+'">'+estado+'</span>';
    }});
  }}

  function actualizarGraficas(df) {{
    if (!window._charts) return;
    // Re-agrega datos por cada gráfica según columna original
    Object.keys(window._charts).forEach(function(chartId) {{
      var chart = window._charts[chartId];
      if (!chart) return;
      var meta = chart._estadinchoMeta;
      if (!meta) return;
      if (meta.tipo === 'line' && meta.colFecha) {{
        // Serie temporal
        var grupos = {{}};
        df.forEach(function(r) {{
          var fecha = r[meta.colFecha];
          if (!fecha) return;
          var mes = String(fecha).substring(0,7);
          if (!grupos[mes]) grupos[mes] = 0;
          grupos[mes] += meta.colVal ? (parseFloat(r[meta.colVal]) || 0) : 1;
        }});
        var keys = Object.keys(grupos).sort();
        chart.data.labels   = keys;
        chart.data.datasets[0].data = keys.map(function(k){{ return grupos[k]; }});
        chart.update();
      }} else if (meta.tipo !== 'line' && meta.colCat) {{
        // Categorías
        var conteo = {{}};
        df.forEach(function(r) {{
          var v = String(r[meta.colCat] !== null && r[meta.colCat] !== undefined ? r[meta.colCat] : '(vacío)');
          conteo[v] = (conteo[v] || 0) + 1;
        }});
        var sorted = Object.entries(conteo).sort(function(a,b){{return b[1]-a[1];}}).slice(0,10);
        chart.data.labels   = sorted.map(function(x){{return x[0];}});
        chart.data.datasets[0].data = sorted.map(function(x){{return x[1];}});
        chart.update();
      }}
    }});
  }}

  function actualizarTablaCategoria(df) {{
    var tbody = document.getElementById('tablaCategoriaCuerpo');
    if (!tbody || !_TABLA_COL) return;
    var conteo = {{}};
    df.forEach(function(r) {{
      var v = String(r[_TABLA_COL] !== null && r[_TABLA_COL] !== undefined ? r[_TABLA_COL] : '(vacío)');
      conteo[v] = (conteo[v] || 0) + 1;
    }});
    var sorted  = Object.entries(conteo).sort(function(a,b){{return b[1]-a[1];}}).slice(0,12);
    var totalF  = sorted.reduce(function(a,b){{return a+b[1];}}, 0);
    var html = '';
    sorted.forEach(function(entry) {{
      var pct    = totalF > 0 ? (entry[1]/totalF*100).toFixed(1) : '0.0';
      var estado = parseFloat(pct) >= 10 ? 'Alto' : (parseFloat(pct) >= 5 ? 'Medio' : 'Bajo');
      var badge  = estado === 'Alto' ? 'badge-success' : (estado === 'Medio' ? 'badge-warning' : 'badge-danger');
      html += '<tr><td>' + entry[0] + '</td><td>' + entry[1].toLocaleString('es-CO') + '</td><td>' + pct + '%</td><td><span class="badge ' + badge + '">' + estado + '</span></td></tr>';
    }});
    tbody.innerHTML = html;
  }}

  function limpiarFiltros() {{
    if (_META.col_fecha) {{
      var d = document.getElementById('filtroFechaDesde');
      var h = document.getElementById('filtroFechaHasta');
      if (d) d.value = '';
      if (h) h.value = '';
    }}
    (_META.cols_cat || []).forEach(function(item) {{
      var colId = item.col.replace(/[^a-zA-Z0-9]/g, '_');
      var sel   = document.getElementById('filtrocat_' + colId);
      if (sel) sel.value = '';
    }});
    (_META.cols_num || []).forEach(function(item) {{
      var colId = item.col.replace(/[^a-zA-Z0-9]/g, '_');
      var elMin = document.getElementById('filtronum_' + colId + '_min');
      var elMax = document.getElementById('filtronum_' + colId + '_max');
      if (elMin) {{ elMin.value = item.min; }}
      if (elMax) {{ elMax.value = item.max; }}
      actualizarSlider(colId);
    }});
    var txt = document.getElementById('filtroTexto');
    if (txt) txt.value = '';
    aplicarFiltros();
  }}

  // Inyectar meta en cada chart para que el filtro sepa qué recalcular
  (function inyectarMeta() {{
    if (!window._charts) {{ setTimeout(inyectarMeta, 100); return; }}
"""

    for g in datos["graficas"]:
        if g["tipo"] == "line" and col_fecha:
            col_val_js = json.dumps(cols_num[0]["col"] if cols_num else "")
            js_filtros += f"""
    if (window._charts['{g["id"]}']) window._charts['{g["id"]}']._estadinchoMeta = {{tipo:'line', colFecha:{json.dumps(col_fecha)}, colVal:{col_val_js}}};"""
        else:
            col_cat_js = ""
            titulo_g = g.get("titulo", "")
            for item in cols_cat:
                if item["col"] in titulo_g:
                    col_cat_js = item["col"]
                    break
            if not col_cat_js and cols_cat:
                col_cat_js = cols_cat[0]["col"]
            js_filtros += f"""
    if (window._charts['{g["id"]}']) window._charts['{g["id"]}']._estadinchoMeta = {{tipo:'{g["tipo"]}', colCat:{json.dumps(col_cat_js)}}};"""

    js_filtros += """
    actualizarContador(_DATASET.length);
  })();"""

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Dashboard - {titulo}</title>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
  <style>
    * {{ margin:0; padding:0; box-sizing:border-box; }}
    :root {{
      --bg-primary:#0f0f23; --bg-secondary:#1a1a2e; --bg-card:#16213e;
      --text-primary:#eaeaea; --text-secondary:#a0a0a0;
      --accent-blue:#4361ee; --accent-purple:#7209b7; --accent-pink:#f72585;
      --accent-orange:#fb8500; --accent-green:#06d6a0; --accent-cyan:#00b4d8;
      --border-color:#2a2a4a;
    }}
    body {{ font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif; background:linear-gradient(135deg,var(--bg-primary) 0%,var(--bg-secondary) 100%); color:var(--text-primary); min-height:100vh; }}
    .container {{ max-width:1600px; margin:0 auto; padding:20px; }}

    /* ── HEADER ── */
    header {{ text-align:center; padding:30px 0 20px; border-bottom:1px solid var(--border-color); margin-bottom:20px; }}
    header h1 {{ font-size:2.5rem; background:linear-gradient(90deg,var(--accent-cyan),var(--accent-pink)); -webkit-background-clip:text; -webkit-text-fill-color:transparent; background-clip:text; margin-bottom:8px; }}
    header p {{ color:var(--text-secondary); font-size:1.1rem; }}
    .titulo-wrapper {{ display:flex; align-items:center; justify-content:center; gap:10px; margin-bottom:6px; flex-wrap:wrap; }}
    #tituloDashboard {{ font-size:2.5rem; font-weight:bold; background:linear-gradient(90deg,var(--accent-cyan),var(--accent-pink)); -webkit-background-clip:text; -webkit-text-fill-color:transparent; background-clip:text; border:none; outline:none; text-align:center; min-width:200px; width:auto; background-color:transparent; cursor:text; border-bottom:2px dashed transparent; transition:border-color .2s; }}
    #tituloDashboard:focus {{ border-bottom:2px dashed var(--accent-cyan); -webkit-text-fill-color:var(--accent-cyan); }}
    .edit-hint {{ color:var(--text-secondary); font-size:0.78rem; opacity:0.6; }}

    /* ── TOOLBAR ── */
    .toolbar {{ display:flex; justify-content:flex-end; gap:10px; margin-bottom:16px; flex-wrap:wrap; }}
    .btn {{ display:inline-flex; align-items:center; gap:6px; padding:9px 18px; border:none; border-radius:10px; font-size:.9rem; font-weight:600; cursor:pointer; transition:all .2s; }}
    .btn-pdf {{ background:linear-gradient(135deg,var(--accent-pink),var(--accent-purple)); color:#fff; }}
    .btn-pdf:hover {{ opacity:.85; transform:translateY(-2px); }}

    /* ── PANEL DE FILTROS ── */
    .panel-filtros {{ background:var(--bg-card); border:1px solid var(--border-color); border-radius:16px; margin-bottom:24px; overflow:hidden; }}
    .filtros-inner {{ padding:18px 22px; }}
    .filtros-titulo {{ font-size:1rem; font-weight:700; color:var(--accent-cyan); margin-bottom:14px; display:flex; align-items:center; gap:10px; }}
    .badge-filtros-activos {{ background:rgba(247,37,133,.2); color:var(--accent-pink); padding:2px 10px; border-radius:20px; font-size:.78rem; font-weight:600; }}
    .filtros-controles {{ display:flex; flex-wrap:wrap; gap:16px; align-items:flex-end; }}
    .filtro-grupo {{ display:flex; flex-direction:column; gap:6px; min-width:180px; flex:1; max-width:280px; }}
    .filtro-busqueda {{ min-width:240px; flex:2; }}
    .filtro-label {{ font-size:.82rem; color:var(--text-secondary); font-weight:600; text-transform:uppercase; letter-spacing:.04em; }}
    .filtro-input {{ background:var(--bg-secondary); border:1px solid var(--border-color); border-radius:8px; color:var(--text-primary); padding:7px 12px; font-size:.9rem; width:100%; transition:border-color .2s; }}
    .filtro-input:focus {{ outline:none; border-color:var(--accent-cyan); }}
    .filtro-select {{ background:var(--bg-secondary); border:1px solid var(--border-color); border-radius:8px; color:var(--text-primary); padding:7px 12px; font-size:.9rem; width:100%; cursor:pointer; transition:border-color .2s; }}
    .filtro-select:focus {{ outline:none; border-color:var(--accent-cyan); }}
    .filtro-fechas {{ display:flex; align-items:center; gap:8px; }}
    .filtro-fechas .filtro-input {{ flex:1; }}
    .filtro-slider-wrap {{ display:flex; flex-direction:column; gap:4px; }}
    .filtro-slider {{ width:100%; accent-color:var(--accent-cyan); cursor:pointer; }}
    .filtro-slider-labels {{ display:flex; justify-content:space-between; font-size:.78rem; color:var(--text-secondary); }}
    .filtro-hint {{ font-size:.78rem; color:var(--accent-green); }}
    .btn-limpiar {{ margin-top:14px; background:rgba(247,37,133,.1); border:1px solid rgba(247,37,133,.3); color:var(--accent-pink); padding:6px 16px; border-radius:8px; font-size:.85rem; font-weight:600; cursor:pointer; transition:all .2s; }}
    .btn-limpiar:hover {{ background:rgba(247,37,133,.2); }}
    .filtro-contador {{ margin-top:10px; font-size:.88rem; color:var(--text-secondary); }}
    .filtro-contador strong {{ color:var(--accent-cyan); }}

    /* ── RESUMEN EJECUTIVO ── */
    .resumen-ejecutivo {{ background:linear-gradient(135deg,rgba(67,97,238,.12),rgba(0,180,216,.08)); border:1px solid rgba(0,180,216,.3); border-radius:16px; padding:22px 28px; margin-bottom:25px; display:flex; gap:18px; align-items:flex-start; }}
    .resumen-icono {{ font-size:2rem; flex-shrink:0; margin-top:2px; }}
    .resumen-ejecutivo h3 {{ font-size:1rem; color:var(--accent-cyan); margin-bottom:8px; letter-spacing:.05em; text-transform:uppercase; }}
    .resumen-ejecutivo p {{ color:var(--text-primary); line-height:1.7; font-size:.97rem; }}

    /* ── KPIs ── */
    .kpi-grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:20px; margin-bottom:30px; }}
    .kpi-card {{ background:var(--bg-card); border-radius:16px; padding:25px; text-align:center; border:1px solid var(--border-color); transition:transform .3s,box-shadow .3s; }}
    .kpi-card:hover {{ transform:translateY(-5px); box-shadow:0 10px 30px rgba(67,97,238,.2); }}
    .kpi-card .icon {{ font-size:2.5rem; margin-bottom:10px; }}
    .kpi-card .value {{ font-size:2.2rem; font-weight:bold; margin-bottom:5px; transition:color .3s; }}
    .kpi-card .label {{ color:var(--text-secondary); font-size:.95rem; }}
    .kpi-card.total {{ border-left:4px solid var(--accent-cyan); }}
    .kpi-card.calidad {{ border-left:4px solid var(--accent-green); }}
    .kpi-card.activos {{ border-left:4px solid var(--accent-orange); }}
    .kpi-card.tiempo {{ border-left:4px solid var(--accent-green); }}
    .kpi-card.retraso {{ border-left:4px solid var(--accent-pink); }}
    .kpi-card.tasa {{ border-left:4px solid var(--accent-purple); }}
    .kpi-card.promedio {{ border-left:4px solid var(--accent-blue); }}
    .kpi-card.num {{ border-left:4px solid var(--accent-cyan); }}

    /* ── GRÁFICAS ── */
    .charts-grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(500px,1fr)); gap:25px; margin-bottom:30px; }}
    .chart-card {{ background:var(--bg-card); border-radius:16px; padding:25px; border:1px solid var(--border-color); }}
    .chart-header {{ display:flex; align-items:center; justify-content:space-between; margin-bottom:20px; gap:10px; }}
    .chart-card h3 {{ color:var(--text-primary); font-size:1.2rem; display:flex; align-items:center; gap:10px; }}
    .chart-card h3::before {{ content:''; width:4px; height:20px; background:linear-gradient(180deg,var(--accent-cyan),var(--accent-pink)); border-radius:2px; flex-shrink:0; }}
    .btn-copy-chart {{ background:rgba(255,255,255,.07); border:1px solid var(--border-color); color:var(--text-secondary); padding:5px 12px; border-radius:8px; font-size:.8rem; cursor:pointer; transition:all .2s; white-space:nowrap; }}
    .btn-copy-chart:hover {{ background:rgba(0,180,216,.15); color:var(--accent-cyan); border-color:var(--accent-cyan); }}
    .chart-container {{ position:relative; height:300px; }}
    .chart-container.tall {{ height:400px; }}
    .full-width {{ grid-column:1/-1; }}
    .half-width {{ grid-column:span 1; }}

    /* ── TABLAS ── */
    .stats-table {{ width:100%; margin-top:15px; border-collapse:collapse; }}
    .stats-table th,.stats-table td {{ padding:12px 15px; text-align:left; border-bottom:1px solid var(--border-color); }}
    .stats-table th {{ background:var(--bg-secondary); color:var(--accent-cyan); font-weight:600; }}
    .stats-table tr:hover {{ background:var(--bg-secondary); }}
    .badge {{ display:inline-block; padding:4px 12px; border-radius:20px; font-size:.85rem; font-weight:500; }}
    .badge-success {{ background:rgba(6,214,160,.2); color:var(--accent-green); }}
    .badge-warning {{ background:rgba(251,133,0,.2); color:var(--accent-orange); }}
    .badge-danger {{ background:rgba(247,37,133,.2); color:var(--accent-pink); }}
    .score-bar-bg {{ display:inline-block; width:80px; height:8px; background:#2a2a4a; border-radius:4px; vertical-align:middle; margin-right:6px; }}
    .score-bar {{ height:8px; border-radius:4px; transition:width .5s; }}

    /* ── FOOTER ── */
    footer {{ text-align:center; padding:20px; margin-top:30px; border-top:1px solid var(--border-color); color:var(--text-secondary); font-size:.9rem; }}

    /* ── TOAST ── */
    #toast {{ position:fixed; bottom:28px; right:28px; background:#06d6a0; color:#0f0f23; padding:10px 22px; border-radius:10px; font-weight:600; font-size:.9rem; opacity:0; transform:translateY(12px); transition:all .3s; pointer-events:none; z-index:9999; }}
    #toast.show {{ opacity:1; transform:translateY(0); }}

    /* ── PRINT ── */
    @media print {{
      body {{ background:#fff !important; color:#000 !important; }}
      .toolbar, .btn-copy-chart, .edit-hint, .panel-filtros {{ display:none !important; }}
      .chart-card, .kpi-card, .resumen-ejecutivo {{ border:1px solid #ddd !important; background:#fff !important; break-inside:avoid; }}
      header h1, #tituloDashboard {{ -webkit-text-fill-color:#1a1a2e !important; color:#1a1a2e !important; }}
      .charts-grid {{ grid-template-columns:1fr 1fr !important; }}
      .full-width {{ grid-column:1/-1 !important; }}
    }}

    @media(max-width:1100px){{ .charts-grid{{grid-template-columns:1fr;}} .chart-container{{height:280px;}} .filtros-controles{{flex-direction:column;}} .filtro-grupo{{max-width:100%;}} }}
    @media(max-width:600px){{ .kpi-grid{{grid-template-columns:repeat(2,1fr);}} header h1{{font-size:1.8rem;}} }}
  </style>
</head>
<body>
<div class="container">

  <header>
    <div class="titulo-wrapper">
      <input id="tituloDashboard" type="text" value="{titulo_escaped}" spellcheck="false"
             oninput="this.style.width=(this.value.length+2)+'ch'"
             title="Haz clic para editar el título" />
    </div>
    <span class="edit-hint">✏️ Haz clic en el título para editarlo</span>
    <p>{datos["subtitulo"]} &nbsp;|&nbsp; Generado: {datos["generado"]}</p>
  </header>

  <div class="toolbar">
    <button class="btn btn-pdf" onclick="window.print()">🖨️ Exportar / Imprimir PDF</button>
  </div>

  {panel_filtros}

  {resumen_html}

  <section class="kpi-grid">
    {kpi_html}
  </section>

  <section class="charts-grid">
    {charts_html}
  </section>

  {tabla_html}
  {calidad_html}
  {corr_html}

  <footer>
    <p>Estadincho-Gen v5.1 &nbsp;|&nbsp; Archivo: {nombre_archivo} &nbsp;|&nbsp; {datos["generado"]}</p>
  </footer>
</div>

<div id="toast">✅ Imagen copiada al portapapeles</div>

<script>
  Chart.defaults.color = '#a0a0a0';
  Chart.defaults.borderColor = '#2a2a4a';

  (function() {{
    var inp = document.getElementById('tituloDashboard');
    if(inp) inp.style.width = (inp.value.length + 2) + 'ch';
  }})();

  function copiarGrafica(chartId) {{
    var chart = window._charts && window._charts[chartId];
    if (!chart) {{ mostrarToast('❌ No se pudo copiar'); return; }}
    chart.canvas.toBlob(function(blob) {{
      try {{
        var item = new ClipboardItem({{ 'image/png': blob }});
        navigator.clipboard.write([item]).then(function() {{
          mostrarToast('✅ Imagen copiada al portapapeles');
        }}).catch(function() {{
          mostrarToast('❌ Permiso denegado por el navegador');
        }});
      }} catch(e) {{
        mostrarToast('❌ Tu navegador no soporta esta función');
      }}
    }});
  }}

  function mostrarToast(msg) {{
    var t = document.getElementById('toast');
    t.textContent = msg;
    t.classList.add('show');
    setTimeout(function(){{ t.classList.remove('show'); }}, 2800);
  }}

  {charts_js}
  {corr_js}
  {js_filtros}
</script>
</body>
</html>"""
    return html
def main():
    if len(sys.argv) < 2:
        print("=" * 55)
        print("  ESTADINCHO-GEN v5.1")
        print("=" * 55)
        print()
        print("Arrastra tu archivo sobre este script, o ejecuta:")
        print()
        print("  python generar_dashboard.py mi_archivo.xlsx")
        print("  python generar_dashboard.py datos.csv")
        print()
        input("Presiona Enter para cerrar...")
        sys.exit(0)

    ruta_archivo = sys.argv[1]
    df, nombre = cargar_archivo(ruta_archivo)

    # Selector de título en consola
    print()
    print(f"  Título por defecto: {nombre.replace('_',' ').replace('-',' ').title()}")
    titulo_input = input("  ¿Título del dashboard? (Enter para usar el nombre del archivo): ").strip()
    titulo_personalizado = titulo_input if titulo_input else ""

    print("Analizando datos...")
    datos = analizar(df)

    print("Generando resumen ejecutivo...")
    datos["resumen"] = generar_resumen(df, datos, nombre)

    print("Preparando filtros interactivos...")
    datos["dataset_json"] = exportar_dataset_json(df)
    datos["meta_filtros"] = metadata_filtros(df)
    del datos["_df"]  # limpiar referencia interna

    print("Generando HTML...")
    html = generar_html(datos, nombre, titulo_personalizado)

    carpeta = Path(ruta_archivo).parent
    salida = (carpeta / f"dashboard_{nombre}.html").resolve()
    salida.write_text(html, encoding="utf-8")

    print(f"\n✔ Dashboard creado: {salida}")
    print("  Abriendo en el navegador...")
    webbrowser.open(salida.as_uri())
    print("\nListo.")

if __name__ == "__main__":
    main()
