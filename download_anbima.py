#!/usr/bin/env python3
"""
Script de download de dados diários ANBIMA (sem Playwright — usa requests puro).

Fluxo:
  Debêntures : lê data disponível do iframe → baixa XLS direto
  CRI/CRA    : lê dataFinal da página → baixa CSV via endpoint

Salva em:
  debentures/debenture-DD-MM.xls
  cri_cra/cri_cra-DD-MM.csv

Uso:
  python download_anbima.py              # baixa dados mais recentes disponíveis
  python download_anbima.py --data 13-05 # baixa especificamente o dia 13/05
  python download_anbima.py --tentar 5   # tenta até 5 dias anteriores se hoje falhar
  python download_anbima.py --apenas debentures
  python download_anbima.py --apenas cri_cra
"""

import re
import sys
import csv
import argparse
import requests
from datetime import datetime, timedelta
from pathlib import Path

BASE_DIR       = Path(__file__).parent.resolve()
DEBENTURES_DIR = BASE_DIR / "debentures"
CRI_CRA_DIR    = BASE_DIR / "cri_cra"

PT_MONTHS = {
    1: "jan",  2: "fev",  3: "mar",  4: "abr",  5: "mai",  6: "jun",
    7: "jul",  8: "ago",  9: "set", 10: "out", 11: "nov", 12: "dez",
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.8",
}


def _session() -> requests.Session:
    s = requests.Session()
    s.headers.update(HEADERS)
    return s


def _msd_fname(dt: datetime) -> str:
    """Monta nome de arquivo MSD da ANBIMA: d26mai13.xls"""
    return f"d{dt.strftime('%y')}{PT_MONTHS[dt.month]}{dt.strftime('%d')}.xls"


def _parse_msd_fname(stem: str) -> datetime | None:
    """d26mai13 → datetime(2026, 5, 13)"""
    PT_MONTHS_REV = {v: k for k, v in PT_MONTHS.items()}
    m = re.match(r"d(\d{2})([a-z]{3})(\d{2})", stem.lower())
    if not m:
        return None
    mon = PT_MONTHS_REV.get(m.group(2))
    if not mon:
        return None
    try:
        return datetime(2000 + int(m.group(1)), mon, int(m.group(3)))
    except ValueError:
        return None


def _read_date_from_csv(content: bytes) -> datetime | None:
    """Lê a data da primeira linha de dados de um CSV CRI/CRA."""
    for enc in ("utf-8-sig", "latin-1", "cp1252"):
        try:
            text = content.decode(enc)
            reader = csv.reader(text.splitlines(), delimiter=";")
            next(reader, None)
            for row in reader:
                if row and row[0].strip():
                    try:
                        return datetime.strptime(row[0].strip(), "%d/%m/%Y")
                    except ValueError:
                        pass
            break
        except UnicodeDecodeError:
            continue
    return None


def download_debentures(
    preferred_dt: datetime | None = None,
    session: requests.Session | None = None,
) -> tuple[bytes, str] | None:
    """
    Baixa arquivo MSD de debêntures.
    Retorna (conteúdo, nome_original) ou None.

    Estratégia:
    1. Busca a data mais recente disponível no iframe ANBIMA.
    2. Se preferred_dt informado, tenta esse arquivo primeiro.
    3. Baixa o XLS via URL direta.
    """
    s = session or _session()
    s.headers.update({"Referer": "https://www.anbima.com.br/"})

    IFRAME_URL = "https://www.anbima.com.br/informacoes/merc-sec-debentures/default.asp"
    ARQS_BASE  = "https://www.anbima.com.br/informacoes/merc-sec-debentures/arqs/"

    # ── 1. Ler data mais recente disponível ─────────────────────────────────
    available_dt: datetime | None = None
    try:
        print(f"  Lendo data disponível: {IFRAME_URL}")
        r = s.get(IFRAME_URL, timeout=20)
        m = re.search(r'name="Dt_Ref"[^>]+value="(\d{2}/\d{2}/\d{4})"', r.text)
        if m:
            available_dt = datetime.strptime(m.group(1), "%d/%m/%Y")
            print(f"  Data disponível no iframe: {available_dt.strftime('%d/%m/%Y')}")
    except requests.RequestException as e:
        print(f"  Aviso: não consegui ler o iframe ({e})")

    # ── 2. Montar lista de datas para tentar ─────────────────────────────────
    candidates: list[datetime] = []
    if preferred_dt:
        candidates.append(preferred_dt)
    if available_dt and available_dt not in candidates:
        candidates.append(available_dt)
    # Fallback: últimos 5 dias
    today = datetime.now()
    for i in range(5):
        dt = today - timedelta(days=i)
        if dt not in candidates:
            candidates.append(dt)

    # ── 3. Baixar ────────────────────────────────────────────────────────────
    for dt in candidates:
        fname = _msd_fname(dt)
        url   = ARQS_BASE + fname
        print(f"  GET {url}")
        try:
            r = s.get(url, timeout=30)
            if r.status_code == 200 and len(r.content) > 50_000:
                print(f"      OK ({len(r.content):,} bytes)")
                return r.content, fname
            print(f"      HTTP {r.status_code} / {len(r.content)} bytes — pulado")
        except requests.RequestException as e:
            print(f"      Erro: {e}")

    return None


def download_cri_cra(
    preferred_dt: datetime | None = None,
    session: requests.Session | None = None,
) -> tuple[bytes, datetime] | None:
    """
    Baixa CSV de taxas CRI/CRA.
    Retorna (conteúdo, data_referência) ou None.

    Estratégia:
    1. Visita a página principal para obter cookies e dataFinal.
    2. Baixa o CSV com filtroData={dataFinal}.
    """
    s = session or _session()

    PAGE_URL = (
        "https://www.anbima.com.br/pt_br/informar/precos-e-indices"
        "/precos/taxas-de-cri-e-cra/taxas-de-cri-e-cra.htm"
    )
    EXPORT_URL = "https://www.anbima.com.br/pt_br/anbima/TaxasCriCraExport/exportarCSV"

    # ── 1. Página principal (cookies + dataFinal) ────────────────────────────
    data_final: datetime | None = None
    try:
        print(f"  Obtendo cookies e dataFinal: {PAGE_URL}")
        r = s.get(PAGE_URL, timeout=20)
        m = re.search(r'id="dataFinal"[^>]*value="(\d{2}/\d{2}/\d{4})"', r.text)
        if m:
            data_final = datetime.strptime(m.group(1), "%d/%m/%Y")
            print(f"  dataFinal na página: {data_final.strftime('%d/%m/%Y')}")
    except requests.RequestException as e:
        print(f"  Aviso: não consegui ler a página ({e})")

    # ── 2. Montar lista de datas para tentar ─────────────────────────────────
    candidates: list[datetime] = []
    if preferred_dt:
        candidates.append(preferred_dt)
    if data_final and data_final not in candidates:
        candidates.append(data_final)
    today = datetime.now()
    for i in range(5):
        dt = today - timedelta(days=i)
        if dt not in candidates:
            candidates.append(dt)

    # ── 3. Baixar ────────────────────────────────────────────────────────────
    s.headers.update({"Referer": PAGE_URL})
    for dt in candidates:
        date_str = dt.strftime("%d/%m/%Y")
        url = f"{EXPORT_URL}?filtroTermo=&filtroData={date_str}"
        print(f"  GET filtroData={date_str}")
        try:
            r = s.get(url, timeout=30)
            if r.status_code == 200 and len(r.content) > 1_000:
                print(f"      OK ({len(r.content):,} bytes)")
                actual_dt = _read_date_from_csv(r.content) or dt
                return r.content, actual_dt
            print(f"      HTTP {r.status_code} / {len(r.content)} bytes — pulado")
        except requests.RequestException as e:
            print(f"      Erro: {e}")

    return None


def save(content: bytes, dest: Path) -> bool:
    dest.parent.mkdir(parents=True, exist_ok=True)
    try:
        with open(dest, "wb") as f:
            f.write(content)
        print(f"  Salvo: {dest.name}")
        return True
    except OSError as e:
        print(f"  Erro ao salvar: {e}")
        return False


def main():
    parser = argparse.ArgumentParser(
        description="Download diário ANBIMA — debêntures e CRI/CRA (sem Playwright)"
    )
    parser.add_argument(
        "--data", metavar="DD-MM",
        help="Data preferida no formato DD-MM (ex: 13-05). Usa a mais recente disponível se omitido."
    )
    parser.add_argument(
        "--apenas", choices=["debentures", "cri_cra"],
        help="Baixa apenas um dos dois arquivos"
    )
    args = parser.parse_args()

    preferred: datetime | None = None
    if args.data:
        try:
            d, m = map(int, args.data.split("-"))
            preferred = datetime(datetime.now().year, m, d)
        except (ValueError, TypeError):
            print("Formato inválido. Use DD-MM, ex: 13-05")
            sys.exit(1)

    print("\n" + "=" * 55)
    print("  Download ANBIMA")
    if preferred:
        print(f"  Data preferida: {preferred.strftime('%d/%m/%Y')}")
    print("=" * 55)

    session = _session()
    deb_ok = cri_ok = True

    # ── Debêntures ────────────────────────────────────────────────────────────
    if args.apenas != "cri_cra":
        print("\n[Debêntures]")
        result = download_debentures(preferred, session)
        if result:
            content, orig_fname = result
            file_dt = _parse_msd_fname(Path(orig_fname).stem) or datetime.now()
            label = file_dt.strftime("%d-%m")
            deb_ok = save(content, DEBENTURES_DIR / f"debenture-{label}.xls")
        else:
            print("  FALHOU — nenhum arquivo encontrado.")
            deb_ok = False

    # ── CRI/CRA ───────────────────────────────────────────────────────────────
    if args.apenas != "debentures":
        print("\n[CRI/CRA]")
        result = download_cri_cra(preferred, session)
        if result:
            content, file_dt = result
            label = file_dt.strftime("%d-%m")
            cri_ok = save(content, CRI_CRA_DIR / f"cri_cra-{label}.csv")
        else:
            print("  FALHOU — nenhum arquivo encontrado.")
            cri_ok = False

    # ── Resumo ────────────────────────────────────────────────────────────────
    print("\n" + "=" * 55)
    if args.apenas != "cri_cra":
        print(f"  Debêntures : {'OK' if deb_ok else 'FALHOU'}")
    if args.apenas != "debentures":
        print(f"  CRI/CRA    : {'OK' if cri_ok else 'FALHOU'}")
    print("=" * 55)

    if not (deb_ok and cri_ok):
        sys.exit(1)

    print(
        "\nDica: após o download, se o servidor estiver rodando, acesse"
        "\n  http://localhost:5000/api/reload   para recarregar sem reiniciar."
    )


if __name__ == "__main__":
    main()
