import os
import math
import json
from datetime import datetime
import openpyxl
from flask import Flask, jsonify, send_from_directory

app = Flask(__name__, static_folder=".")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ANBIMA_BASE_FILE = os.path.join(BASE_DIR, "anbima v0.xlsx")
ANBIMA_CURR_FILE = os.path.join(BASE_DIR, "anbima data.xlsx")
BTG_FILE = os.path.join(BASE_DIR, "119452.xlsx")

ANBIMA_SHEETS = ["DI_SPREAD", "IPCA_SPREAD", "PREFIXADO", "DI_PERCENTUAL", "IGP-M"]

RECOVERY_RATE = 0.30


def safe_float(val):
    try:
        if val is None or val == "--" or val == "N/D" or val == "":
            return None
        f = float(val)
        return None if math.isnan(f) or math.isinf(f) else f
    except (ValueError, TypeError):
        return None


def parse_anbima_file(filepath):
    """Parse all sheets from an ANBIMA MSD xlsx file.
    Returns dict: {codigo: row_dict}
    Also returns the reference date.
    """
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    records = {}
    ref_date = None

    for sheet_name in ANBIMA_SHEETS:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        rows_iter = ws.iter_rows(values_only=True)

        # Find reference date (row 4, col 2)
        for i, row in enumerate(rows_iter):
            if i == 3 and ref_date is None:
                val = row[1] if len(row) > 1 else None
                if isinstance(val, datetime):
                    ref_date = val.strftime("%Y-%m-%d")
                elif val:
                    ref_date = str(val)[:10]
            # Data starts at row 10 (index 9)
            if i < 9:
                continue
            if not row or row[0] is None:
                continue
            codigo = str(row[0]).strip()
            if len(codigo) < 4:
                continue

            nome = str(row[1]).strip() if row[1] else ""
            has_star = "(*)" in nome and "(**)" not in nome
            has_double_star = "(**)" in nome
            nome_clean = nome.replace(" (**)", "").replace(" (*)", "").strip()

            venc = str(row[2]).strip() if row[2] else ""
            indice = str(row[3]).strip() if row[3] else ""

            taxa_compra = safe_float(row[4]) if len(row) > 4 else None
            taxa_venda = safe_float(row[5]) if len(row) > 5 else None
            taxa_ind = safe_float(row[6]) if len(row) > 6 else None
            desvio = safe_float(row[7]) if len(row) > 7 else None
            int_min = safe_float(row[8]) if len(row) > 8 else None
            int_max = safe_float(row[9]) if len(row) > 9 else None
            pu = safe_float(row[10]) if len(row) > 10 else None
            pu_par = safe_float(row[11]) if len(row) > 11 else None
            duration = safe_float(row[12]) if len(row) > 12 else None
            pct_reune = safe_float(row[13]) if len(row) > 13 else None

            sem_cotacao = (row[4] == "--" or row[4] is None) if len(row) > 4 else False

            bid_ask_bps = None
            if taxa_compra is not None and taxa_venda is not None:
                bid_ask_bps = round((taxa_compra - taxa_venda) * 100, 2)

            records[codigo] = {
                "codigo": codigo,
                "nome": nome_clean,
                "vencimento": venc,
                "indice_correcao": indice,
                "indexador": sheet_name,
                "taxa_compra": taxa_compra,
                "taxa_venda": taxa_venda,
                "taxa_indicativa": taxa_ind,
                "desvio_padrao": desvio,
                "intervalo_min": int_min,
                "intervalo_max": int_max,
                "pu": pu,
                "pu_par": pu_par,
                "duration": duration,
                "pct_reune": pct_reune,
                "bid_ask_bps": bid_ask_bps,
                "has_star": has_star,
                "has_double_star": has_double_star,
                "sem_cotacao": sem_cotacao,
            }

    wb.close()
    return records, ref_date


def parse_btg_sectors():
    """Parse sector info from BTG Weekly BigChart sheets.
    Returns dict: {codigo: setor}
    """
    sector_map = {}
    setor_sheets = [
        ("BigChart Deb CDI - SETOR", "DI+"),
        ("BigChart Deb incent - SETOR", "Incent"),
        ("BigChart IPCA+ - SETOR", "IPCA+"),
    ]

    try:
        wb = openpyxl.load_workbook(BTG_FILE, read_only=True, data_only=True)
    except Exception:
        return sector_map

    for sheet_name, _ in setor_sheets:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        current_sector = "Outros"
        for row in ws.iter_rows(values_only=True):
            if not row or row[1] is None:
                continue
            cell = str(row[1]).strip()
            # Sector header rows: no code-like value in col 2 (emissor col)
            # Detect: if col 2 (emissor) is text and col 1 (codigo) matches sector name
            if len(row) > 2 and row[2] is None:
                # Likely a sector header
                current_sector = cell
            elif len(row) > 2 and row[2] is not None:
                # Code row
                codigo = cell
                if len(codigo) >= 4 and codigo not in ("Código", "C\u00f3digo"):
                    sector_map[codigo] = current_sector

    wb.close()
    return sector_map


def compute_pd(taxa_indicativa, recovery=RECOVERY_RATE):
    """PD implícita via credit spread model."""
    if taxa_indicativa is None:
        return None
    spread = taxa_indicativa / 100
    pd = 1 - math.exp(-spread / (1 - recovery))
    return round(max(0, min(1, pd)) * 100, 2)


def classify_stress(row):
    """Return stress level: 0=ok, 1=atenção, 2=stress, 3=distressed."""
    taxa = row.get("taxa_indicativa")
    pu_par = row.get("pu_par")
    desvio = row.get("desvio_padrao")
    bid_ask = row.get("bid_ask_bps")
    delta = row.get("delta_taxa_bps")
    double_star = row.get("has_double_star")
    sem_cot = row.get("sem_cotacao")

    score = 0
    if taxa and taxa > 15:
        score += 3
    elif taxa and taxa > 10:
        score += 2
    elif taxa and taxa > 5:
        score += 1

    if pu_par is not None:
        if pu_par < 50:
            score += 3
        elif pu_par < 70:
            score += 2
        elif pu_par < 85:
            score += 1

    if desvio and desvio > 1.0:
        score += 2
    elif desvio and desvio > 0.5:
        score += 1

    if bid_ask and bid_ask > 300:
        score += 2
    elif bid_ask and bid_ask > 100:
        score += 1

    if delta and delta > 200:
        score += 2
    elif delta and delta > 100:
        score += 1

    if double_star:
        score += 2
    if sem_cot:
        score += 1

    if score >= 6:
        return 3
    elif score >= 3:
        return 2
    elif score >= 1:
        return 1
    return 0


_cache = None


def build_data():
    global _cache
    if _cache is not None:
        return _cache

    base_records, base_date = parse_anbima_file(ANBIMA_BASE_FILE)
    curr_records, curr_date = parse_anbima_file(ANBIMA_CURR_FILE)
    sector_map = parse_btg_sectors()

    papers = []
    for codigo, curr in curr_records.items():
        base = base_records.get(codigo, {})
        setor = sector_map.get(codigo, "Outros")

        # Deltas
        delta_taxa = None
        delta_pu = None
        delta_desvio = None

        if curr.get("taxa_indicativa") is not None and base.get("taxa_indicativa") is not None:
            delta_taxa = round((curr["taxa_indicativa"] - base["taxa_indicativa"]) * 100, 1)

        if curr.get("pu_par") is not None and base.get("pu_par") is not None:
            delta_pu = round(curr["pu_par"] - base["pu_par"], 2)

        if curr.get("desvio_padrao") is not None and base.get("desvio_padrao") is not None:
            delta_desvio = round(curr["desvio_padrao"] - base["desvio_padrao"], 4)

        pd_impl = compute_pd(curr.get("taxa_indicativa"))

        paper = {
            **curr,
            "setor": setor,
            # Flat fields for both dates (table columns)
            "taxa_v0": base.get("taxa_indicativa"),
            "taxa_v1": curr.get("taxa_indicativa"),
            "pu_par_v0": base.get("pu_par"),
            "pu_par_v1": curr.get("pu_par"),
            "desvio_v0": base.get("desvio_padrao"),
            "desvio_v1": curr.get("desvio_padrao"),
            # Deltas
            "delta_taxa_bps": delta_taxa,
            "delta_pu_par": delta_pu,
            "delta_desvio": delta_desvio,
            "pd_implicita": pd_impl,
            # History for modal charts
            "history": [
                {
                    "data": base_date,
                    "taxa_indicativa": base.get("taxa_indicativa"),
                    "pu_par": base.get("pu_par"),
                    "desvio_padrao": base.get("desvio_padrao"),
                    "bid_ask_bps": base.get("bid_ask_bps"),
                },
                {
                    "data": curr_date,
                    "taxa_indicativa": curr.get("taxa_indicativa"),
                    "pu_par": curr.get("pu_par"),
                    "desvio_padrao": curr.get("desvio_padrao"),
                    "bid_ask_bps": curr.get("bid_ask_bps"),
                },
            ],
        }
        paper["stress_level"] = classify_stress(paper)
        papers.append(paper)

    papers.sort(key=lambda x: (x.get("delta_taxa_bps") or 0), reverse=True)

    meta = {
        "base_date": base_date,
        "curr_date": curr_date,
        "total": len(papers),
        "distressed": sum(1 for p in papers if p["stress_level"] == 3),
        "stress": sum(1 for p in papers if p["stress_level"] == 2),
        "atencao": sum(1 for p in papers if p["stress_level"] == 1),
        "ok": sum(1 for p in papers if p["stress_level"] == 0),
        "setores": sorted(set(p["setor"] for p in papers)),
        "indexadores": sorted(set(p["indexador"] for p in papers)),
    }

    _cache = {"papers": papers, "meta": meta}
    return _cache


@app.route("/")
def index():
    return send_from_directory(".", "index.html")


@app.route("/api/data")
def api_data():
    data = build_data()
    return jsonify(data)


@app.route("/api/paper/<codigo>")
def api_paper(codigo):
    data = build_data()
    paper = next((p for p in data["papers"] if p["codigo"] == codigo), None)
    if not paper:
        return jsonify({"error": "Not found"}), 404
    return jsonify(paper)


if __name__ == "__main__":
    print("Carregando dados...")
    build_data()
    print("Servidor iniciado em http://localhost:5000")
    app.run(debug=False, host="0.0.0.0", port=5000)
