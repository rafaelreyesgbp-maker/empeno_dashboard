import os
import re
import json
import unicodedata
import requests
import xlrd
from datetime import datetime

# ── Configuración ──────────────────────────────────────────────────────────────
FOLDER_ID = "1Xz2sv3PjhwzSesUjpj1XoFF37fFfS42F"
API_KEY   = os.environ["DRIVE_API_KEY"]
HTML_PATH = "index.html"

METAS = {
    1:58989, 2:58989, 3:58989,  4:58989,
    5:58989, 6:58989, 7:58989,  8:58989,
    9:58989,10:58989,11:58989, 12:58993,
}

MONTH_NAMES = {
    "enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,
    "julio":7,"agosto":8,"septiembre":9,"octubre":10,"noviembre":11,"diciembre":12,
}

MONTH_LABELS = {
    1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
    7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre",
}

# Columnas base-0: RFC=A(0), Contrib=B(1), Periodo=E(4), Rec=G(6)+M(12)
RFC_COL     = 0
CONTRIB_COL = 1
PERIODO_COL = 4
REC_G_COL   = 6
REC_M_COL   = 12


# ── Google Drive ───────────────────────────────────────────────────────────────
def drive_list_files():
    """Lista archivos en la carpeta (sin filtro MIME, acepta .xls y .xlsx)."""
    url = "https://www.googleapis.com/drive/v3/files"
    params = {
        "q": f"'{FOLDER_ID}' in parents and trashed=false",
        "fields": "files(id,name,mimeType)",
        "pageSize": 40,
        "key": API_KEY,
    }
    r = requests.get(url, params=params, timeout=30)
    r.raise_for_status()
    return r.json().get("files", [])


def drive_download(file_id):
    """Descarga el contenido binario del archivo."""
    url = f"https://www.googleapis.com/drive/v3/files/{file_id}"
    params = {"alt": "media", "key": API_KEY}
    r = requests.get(url, params=params, timeout=60)
    r.raise_for_status()
    return r.content


# ── Parser XLS ─────────────────────────────────────────────────────────────────
def normalize(s):
    """Elimina acentos y convierte a minúsculas."""
    return unicodedata.normalize("NFD", s).encode("ascii", "ignore").decode("ascii").lower()


def _find_data_start(rows, col):
    """Busca la primera fila con RFC válido (>=12 chars, no solo letras)."""
    for i in range(min(20, len(rows))):
        v = str(rows[i][col]).strip() if col < len(rows[i]) else ""
        if len(v) >= 12 and not re.match(r"^[a-zA-Z\s]+$", v):
            return i
    return -1


def _parse_xls(content):
    """Parsea un archivo .xls con xlrd y retorna lista de registros."""
    try:
        wb = xlrd.open_workbook(file_contents=content)
        sheet = wb.sheet_by_index(0)
    except Exception as e:
        print(f"  [ERROR] No se pudo abrir el archivo XLS: {e}")
        return []

    # Convertir a lista de listas
    rows = []
    for i in range(sheet.nrows):
        row = []
        for j in range(sheet.ncols):
            cell = sheet.cell(i, j)
            if cell.ctype == xlrd.XL_CELL_NUMBER:
                row.append(cell.value)
            elif cell.ctype == xlrd.XL_CELL_TEXT:
                row.append(cell.value)
            else:
                row.append("")
        rows.append(row)

    start = _find_data_start(rows, RFC_COL)
    if start < 0:
        print("  [WARN] No se encontro inicio de datos valido")
        return []

    out = []
    rec_sum = 0.0
    for i in range(start, len(rows)):
        row = rows[i]
        max_col = max(RFC_COL, CONTRIB_COL, PERIODO_COL, REC_G_COL, REC_M_COL)
        if len(row) <= max_col:
            continue

        # RFC
        rfc = str(row[RFC_COL]).strip().upper()
        if not rfc or len(rfc) < 12:
            continue

        # Periodo: float -> int -> str, validar 6 digitos exactos
        p = str(row[PERIODO_COL]).strip()
        if re.match(r"^\d+\.?\d*$", p):
            p = str(int(float(p)))
        if len(p) != 6:
            continue  # Serial de fecha Excel u otro formato incorrecto

        # Recaudacion: G + M
        try:
            g_val = float(row[REC_G_COL]) if row[REC_G_COL] != "" else 0.0
            m_val = float(row[REC_M_COL]) if row[REC_M_COL] != "" else 0.0
        except (ValueError, TypeError):
            g_val, m_val = 0.0, 0.0
        rec = g_val + m_val

        contrib = str(row[CONTRIB_COL]).strip()
        out.append({"rfc": rfc, "periodo": p, "recaudacion": rec, "contrib": contrib})
        rec_sum += rec

    print(f"  Registros parseados: {len(out)} | Suma recaudacion: ${rec_sum:,.2f}")
    return out


# ── Logica de omisos ───────────────────────────────────────────────────────────
def prev_period(p):
    p = str(p)
    y, m = int(p[:4]), int(p[4:])
    m -= 1
    if m == 0:
        m, y = 12, y - 1
    return f"{y}{str(m).zfill(2)}"


def format_period(p):
    s = str(p)
    labels = ["","Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
    try:
        return labels[int(s[4:6])] + "-" + s[2:4]
    except Exception:
        return s


def get_dominant(records):
    t = {}
    for r in records:
        if r["periodo"]:
            t[r["periodo"]] = t.get(r["periodo"], 0) + r["recaudacion"]
    if not t:
        return None
    return max(t, key=lambda k: t[k])


def get_missing_periods(paid_set, dominant, max_back, stop_before):
    out = []
    p = str(dominant)
    while p not in paid_set:
        if stop_before and int(p) < int(stop_before):
            break
        out.append(p)
        p = prev_period(p)
        if len(out) >= max_back:
            break
    return out


def compute_month(month_num, all_month_data):
    cur = all_month_data.get(month_num, [])
    dominant = get_dominant(cur)
    acumulado = sum(r["recaudacion"] for r in cur)

    # Meses de referencia: ultimos 4 meses anteriores con datos
    ref_months = [
        m for m in range(max(1, month_num - 4), month_num)
        if all_month_data.get(m)
    ]
    n_ref = len(ref_months)

    # Indice global de periodos pagados por RFC (maximo por periodo)
    global_periods = {}
    global_contrib = {}
    for m in range(1, month_num + 1):
        for r in all_month_data.get(m, []):
            if r["rfc"] not in global_periods:
                global_periods[r["rfc"]] = {}
            gp = global_periods[r["rfc"]]
            if r["periodo"] not in gp or r["recaudacion"] > gp[r["periodo"]]:
                gp[r["periodo"]] = r["recaudacion"]
            if r["rfc"] not in global_contrib and r["contrib"]:
                global_contrib[r["rfc"]] = r["contrib"]

    # Conteos en meses de referencia
    rfc_ref_count    = {}
    rfc_ref_periods  = {}
    rfc_contrib      = {}
    paid_2026_in_ref = {}

    for rm in ref_months:
        seen = set()
        for r in all_month_data.get(rm, []):
            if r["rfc"] not in rfc_ref_periods:
                rfc_ref_periods[r["rfc"]] = {}
            rp = rfc_ref_periods[r["rfc"]]
            if r["periodo"] not in rp or r["recaudacion"] > rp[r["periodo"]]:
                rp[r["periodo"]] = r["recaudacion"]
            if str(r["periodo"]).startswith("2026"):
                if r["rfc"] not in paid_2026_in_ref:
                    paid_2026_in_ref[r["rfc"]] = set()
                paid_2026_in_ref[r["rfc"]].add(str(r["periodo"]))
            if r["rfc"] not in seen:
                seen.add(r["rfc"])
                rfc_ref_count[r["rfc"]] = rfc_ref_count.get(r["rfc"], 0) + 1
            if r["rfc"] not in rfc_contrib and r["contrib"]:
                rfc_contrib[r["rfc"]] = r["contrib"]

    # Candidatos: >=2 meses de referencia O >=2 periodos 2026 pagados en ref
    candidates = set()
    for rfc, cnt in rfc_ref_count.items():
        if cnt >= 2:
            candidates.add(rfc)
    for rfc, ps in paid_2026_in_ref.items():
        if len(ps) >= 2:
            candidates.add(rfc)

    omisos = []
    for rfc in candidates:
        cnt = rfc_ref_count.get(rfc, 0)
        p26 = paid_2026_in_ref.get(rfc, set())
        if cnt < 2 and len(p26) < 2:
            continue
        paid_set = set(global_periods.get(rfc, {}).keys())
        if not dominant or dominant in paid_set:
            continue
        has_2026 = any(p.startswith("2026") for p in paid_set)
        if has_2026:
            missing = get_missing_periods(paid_set, dominant, 12, None)
        else:
            missing = get_missing_periods(paid_set, dominant, 12, "202601")
        if not missing:
            continue
        ref_amounts = list(rfc_ref_periods.get(rfc, {}).values())
        if not ref_amounts:
            continue
        avg = sum(ref_amounts) / len(ref_amounts)

        # Segmentacion
        if not has_2026:
            seg = "omisos_totales"
        elif cnt >= n_ref:
            seg = "alta"
        elif cnt >= 3:
            seg = "media"
        else:
            seg = "seguimiento"

        contrib = rfc_contrib.get(rfc) or global_contrib.get(rfc, "")
        omisos.append({
            "rfc":      rfc,
            "contrib":  contrib,
            "count":    cnt,
            "avg":      round(avg * len(missing)),
            "nMissing": len(missing),
            "pending":  [format_period(p) for p in missing],
            "seg":      seg,
        })

    omisos.sort(key=lambda o: o["avg"], reverse=True)

    esperado    = sum(o["avg"] for o in omisos if o["seg"] in ("alta", "media"))
    proyeccion  = acumulado + esperado
    meta        = METAS.get(month_num, 0)

    segments = {}
    for o in omisos:
        if o["seg"] not in segments:
            segments[o["seg"]] = {"count": 0, "monto": 0, "omisos": []}
        segments[o["seg"]]["count"]  += 1
        segments[o["seg"]]["monto"]  += o["avg"]
        segments[o["seg"]]["omisos"].append({
            "rfc": o["rfc"], "contrib": o["contrib"],
            "avg": o["avg"], "count": o["count"],
            "nMissing": o["nMissing"], "pending": o["pending"],
        })
    for s in segments.values():
        s["monto"] = round(s["monto"])
        s["omisos"].sort(key=lambda x: x["avg"], reverse=True)

    return {
        "mes_label":         MONTH_LABELS.get(month_num, str(month_num)),
        "mes_num":           month_num,
        "meta":              meta,
        "dominant_period":   int(dominant) if dominant else 0,
        "ref_months":        ref_months,
        "acumulado_real":    round(acumulado),
        "total_omisos":      len(omisos),
        "total_esperado":    round(esperado),
        "proyeccion_cierre": round(proyeccion),
        "meta_cruzada":      proyeccion >= meta,
        "pct_acumulado":     acumulado / meta * 100 if meta else 0,
        "pct_proyeccion":    proyeccion / meta * 100 if meta else 0,
        "segmentos":         segments,
        "omisos":            omisos[:5000],
    }


# ── Actualizacion del HTML ─────────────────────────────────────────────────────
def update_html(new_data, html_path):
    with open(html_path, "r", encoding="utf-8") as f:
        html = f.read()

    # Leer allData existente; filtrar claves invalidas (solo "1"-"12")
    existing = {}
    m = re.search(r"let allData\s*=\s*(\{.*?\});", html, re.DOTALL)
    if m:
        try:
            raw = json.loads(m.group(1))
            existing = {
                k: v for k, v in raw.items()
                if k.isdigit() and 1 <= int(k) <= 12
            }
        except Exception:
            pass

    # Merge: solo actualizar si el nuevo acumulado_real es mayor (proteccion)
    merged = dict(existing)
    for key, val in new_data.items():
        k = str(key)
        existing_acum = merged.get(k, {}).get("acumulado_real", 0)
        new_acum      = val.get("acumulado_real", 0)
        if k not in merged or new_acum > existing_acum:
            merged[k] = val

    now = datetime.now().strftime("%d/%m/%Y %H:%M")

    html = re.sub(
        r"let allData\s*=\s*\{.*?\};",
        f"let allData = {json.dumps(merged, ensure_ascii=False)};",
        html,
        flags=re.DOTALL,
    )
    html = re.sub(
        r"var lastUpdated\s*=\s*'[^']*';",
        f"var lastUpdated = '{now}';",
        html,
    )

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  HTML actualizado: {html_path} [{now}]")


# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("  AAFY · Casas de Empeno · Actualizador de Dashboard")
    print("=" * 60)

    # 1. Listar archivos
    print("\n[1] Listando archivos en Drive...")
    files = drive_list_files()
    print(f"  Archivos encontrados: {len(files)}")

    # 2. Detectar mes por nombre (normalizacion sin acentos)
    month_files = []
    for f in files:
        name_n = normalize(f["name"])
        mn = next((MONTH_NAMES[m] for m in MONTH_NAMES if m in name_n), None)
        if mn:
            month_files.append({**f, "num": mn})
            print(f"  -> {f['name']}  ->  mes {mn} ({MONTH_LABELS[mn]})")

    if not month_files:
        print("[WARN] No se encontraron archivos con nombre de mes. Saliendo.")
        return

    month_files.sort(key=lambda f: f["num"])

    # 3. Descargar y parsear
    print("\n[2] Descargando y parseando archivos...")
    all_month_data = {}
    for f in month_files:
        print(f"\n  Archivo: {f['name']}")
        content = drive_download(f["id"])
        records = _parse_xls(content)
        if f["num"] not in all_month_data:
            all_month_data[f["num"]] = []
        all_month_data[f["num"]].extend(records)

    # 4. Calcular proyecciones
    print("\n[3] Calculando proyecciones...")
    new_data = {}
    for num in sorted(all_month_data.keys()):
        result = compute_month(num, all_month_data)
        new_data[str(num)] = result
        print(
            f"  Mes {num:2d} ({MONTH_LABELS[num]:<12}): "
            f"acumulado=${result['acumulado_real']:>10,.0f} | "
            f"omisos={result['total_omisos']:>4}"
        )

    # 5. Actualizar HTML
    print(f"\n[4] Actualizando {HTML_PATH}...")
    update_html(new_data, HTML_PATH)

    print("\nProceso completado.")


if __name__ == "__main__":
    main()
