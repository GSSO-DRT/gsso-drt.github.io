"""
Genera un corte semanal del Reporte RF DRT a partir de la exportación
DETALLE_REQUERIMIENTOS_RF_DRT_<fecha>.xlsx.

Uso:
    python generar_corte.py DETALLE_REQUERIMIENTOS_RF_DRT_2026-11-03.xlsx 2026-11-03

Crea (o reemplaza) en la carpeta cortes/:
    resumen_<fecha>.json   resumen por contrato y RF (liviano, se guarda para el historial)
    brechas_<fecha>.json   detalle de brechas abiertas del corte
    index.json             lista de cortes disponibles (se actualiza sola)

Requiere: pip install pandas openpyxl
Reglas aplicadas (las mismas del corte 27-10-2026):
  - Solo registros con Existencia = 1.
  - % implementación = Sí / (Sí + No) a nivel de requisito; N/A y sin respuesta no cuentan.
  - Si un contrato trae varias versiones vigentes, se usa la del mes/año más reciente indicado en el nombre del archivo.
  - Finning 4600020362 y 4600028056 se agrupan (misma línea base).
"""
import json, re, sys
from pathlib import Path
import pandas as pd

F = "Cumple Evidencia (SI/NO/NA) Final"
MES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
       "septiembre": 9, "setiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}
AGRUPAR = {"FINNING CHILE S.A. - 4600028056": "FINNING CHILE S.A. - 4600020362"}
NUM_AGRUPADO = {"4600020362": "4600020362 · 4600028056"}
NOMBRES = [("MANTENCIONES Y SERVICIOS SALFA", "Salfa Mantenciones"), ("SOLUCIONES AMBIENTALES DEL NORTE", "Seche"),
           ("TECNOLOGIA EN TRANSP. DE MINERALES", "TTM"), ("TRANYMEC", "Tranymec"), ("(MISPA)", "MISPA"),
           ("COMUNICACIONES Y TECNOLOGIAS", "Comtecsa"), ("C. DE ESTUD. MEDICION", "Cesmec"),
           ("INGENIERIA CIVIL VICENTE", "Ing. Civil Vicente"), ("PETROLEOS DE CHILE COPEC", "Copec"),
           ("REPREIND", "Repreind (subc. Joy Global)"), ("INVERSIONES FARIAS", "Inv. Farías (subc. Joy Global)"),
           ("SANTA IDE", "Santa Ide (Provetec)"), ("JOY GLOBAL", "Joy Global (subc. Komatsu)")]


def fecha_txt(s):
    if not isinstance(s, str):
        return None
    m = re.search(r"(\d{1,2}) de (\w+) de (\d{4})", s)
    if not m or m.group(2).lower() not in MES:
        return None
    return pd.Timestamp(int(m.group(3)), MES[m.group(2).lower()], int(m.group(1)))


MES_ARCH = {"ENE": 1, "FEB": 2, "MAR": 3, "ABR": 4, "MAY": 5, "JUN": 6, "JUL": 7, "AGO": 8, "SEP": 9, "SET": 9, "OCT": 10, "NOV": 11, "DIC": 12}


def fecha_archivo(nombre_archivo):
    """Ordena versiones de un mismo contrato por el mes/año escrito en el nombre del archivo."""
    t = nombre_archivo.upper()
    m = re.search(r"(\d{2})[-_.]?(\d{2})[-_.]?(20\d{2})", t)
    if m:
        return (int(m.group(3)), int(m.group(2)), int(m.group(1)), t)
    for k, v in MES_ARCH.items():
        if re.search(r"\b" + k, t):
            y = re.search(r"(20\d{2})", t)
            return (int(y.group(1)) if y else 0, v, 0, t)
    return (0, 0, 0, t)


def nombre(k):
    for a, b in NOMBRES:
        if a in k:
            return b
    m = re.match(r"(.*?)\s*-\s*((?:SC[\w-]*-)?\d{10})\s*$", k)
    n = m.group(1) if m else k
    n = re.sub(r"\b(S\.?\s?P\.?\s?A\.?|S\.?\s?A\.?\s?C?\.?|LTDA\.?|LIMITADA|LIMI|CHILE)(?=\s|$|\)|,)", "", n, flags=re.I)
    return re.sub(r"\s+", " ", n).strip(" .,-").title().replace("Sub Contrato", "subc.")


def numero(k):
    m = re.search(r"((?:SC[\w-]*-)?\d{10})\s*$", k)
    n = m.group(1) if m else ""
    return NUM_AGRUPADO.get(n, n)


def txt(v, n):
    if not isinstance(v, str):
        return ""
    v = re.sub(r"\s+", " ", v.replace("_x000D_", " ")).strip()
    return v[:n]


def main(xlsx, fecha, out_dir="cortes"):
    out = Path(out_dir)
    out.mkdir(exist_ok=True)
    corte = pd.Timestamp(fecha)
    d = pd.read_excel(xlsx, dtype=str)
    d = d[d["Existencia"] != "0"].copy()
    # una versión por contrato: la del archivo más reciente (orden por nombre de archivo)
    eecc = d["Fuente"] != "CODELCO"
    ult = d[eecc].groupby("CODIGO_CONTRATO")["Archivo Origen"].agg(lambda s: max(s.unique(), key=fecha_archivo))
    d = d[~eecc | (d["Archivo Origen"] == d["CODIGO_CONTRATO"].map(ult))]

    def clave(r):
        if r["Fuente"] == "CODELCO":
            sup = r["Superintendencia / Dirección"]
            return "COD|" + r["Gerencia"] + "|" + (sup if isinstance(sup, str) and sup != "-" else "Gerencia")
        return "EC|" + r["Gerencia"] + "|" + AGRUPAR.get(r["CODIGO_CONTRATO"], r["CODIGO_CONTRATO"])

    d["u"] = d.apply(clave, axis=1)
    d = d.drop_duplicates(["u", "RF-CC-EV"])
    d["si"] = (d[F] == "SI").astype(int)
    d["no"] = (d[F] == "NO").astype(int)
    d["sr"] = (d["Tipo de Respuesta"] == "SR").astype(int)
    rfs = sorted(d[d.si + d.no > 0]["RF"].unique())

    # brechas abiertas
    b = d[(d[F] == "NO") & (d["Estatus"] != "Cerrado")].copy()
    b["X"] = b["Fecha Programación"].map(fecha_txt)
    b["plan"] = b["Plan de acción"].map(lambda v: txt(v, 220)).where(lambda s: ~s.isin(["N/A", "-"]), "")
    b["venc"] = ((b["Estatus"] == "Fuera de Plazo") | (b["X"] < corte)).astype(int)
    b["sp"] = (b["plan"] == "").astype(int)

    # semanas abierta: se hereda la fecha "desde" del corte anterior
    idx_path = out / "index.json"
    idx = json.loads(idx_path.read_text(encoding="utf-8")) if idx_path.exists() else {"cortes": []}
    prev = [c for c in sorted(idx["cortes"]) if c < fecha]
    desde_prev = {}
    if prev:
        pb = out / f"brechas_{prev[-1]}.json"
        if pb.exists():
            for r in json.loads(pb.read_text(encoding="utf-8"))["br"]:
                desde_prev[(r[0], r[2])] = r[13]

    units = []
    for u, g in d.groupby("u"):
        tipo, ger, k = u.split("|", 2)
        sup = g["Superintendencia / Dirección"].dropna()
        sup = sup[sup != "-"]
        rf = {r: [int(x.si.sum()), int(x.no.sum()), int(x.sr.sum())] for r, x in g.groupby("RF") if r in rfs}
        bb = b[b.u == u]
        if tipo == "COD":
            nm, num = "Codelco · " + (k if k != "Gerencia" else "línea base gerencia"), ""
        else:
            nm, num = nombre(k), numero(k)
        units.append([u, "Codelco" if tipo == "COD" else "EECC", ger, sup.mode().iloc[0] if len(sup) else "", nm, num, rf,
                      int(g.si.sum()), int(g.no.sum()), int(g.sr.sum()), len(bb), int(bb.sp.sum()), int(bb.venc.sum())])

    br = []
    for _, r in b.iterrows():
        desde = desde_prev.get((r.u, r["RF-CC-EV"]), fecha)
        br.append([r.u, r["RF"], r["RF-CC-EV"], txt(r["CODIGO CC"], 10), txt(r["Nombre del Control Crítico"], 90),
                   txt(r["Brecha"], 220), r["plan"], txt(r["Responsable"], 60),
                   r["X"].strftime("%Y-%m-%d") if r["X"] is not None and not pd.isna(r["X"]) else "",
                   txt(r["Estatus"], 30), int(r["venc"]), int(r["sp"]), txt(r["Pregunta"], 200), desde])

    def dump(p, obj):
        p.write_text(json.dumps(obj, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")

    dump(out / f"resumen_{fecha}.json", {"corte": fecha, "rfs": rfs, "units": units})
    dump(out / f"brechas_{fecha}.json", {"corte": fecha, "br": br})
    idx["cortes"] = sorted(set(idx["cortes"]) | {fecha})
    dump(idx_path, idx)
    ec = [x for x in units if x[1] == "EECC" and x[7] + x[8] > 0]
    print(f"Corte {fecha}: {len(units)} unidades, {len(br)} brechas abiertas, "
          f"EECC {100 * sum(x[7] for x in ec) / sum(x[7] + x[8] for x in ec):.1f}% · cortes: {idx['cortes']}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2], *(sys.argv[3:4]))
