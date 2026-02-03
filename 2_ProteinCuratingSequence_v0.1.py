import os
import re
import time
import json
from typing import Dict, Optional, Tuple, List

import requests
from openpyxl import load_workbook

# ========= CONFIG =========
INPUT_XLSX = r"PATH_with_GeneID.xlsx"
OUTPUT_XLSX = r"PATH_selected_and_fasta.xlsx"

COL_ACCESSION = "A"   # protein accession (NCBI)
COL_PROT_LEN = "C"    # protein length (aa)
COL_GENEID = "Y"      # GeneID
COL_FASTA = "Z"       # output FASTA (custom)
COL_FLAG = "AA"       # optional flag column for skipped duplicates

NCBI_EMAIL = "email"
NCBI_API_KEY = "API KEY"  # opcional (recomendado)
TOOL_NAME = "geneid_dedup_and_fasta_fetch"

SLEEP_SECONDS = 0.34 if not NCBI_API_KEY else 0.12

CACHE_FILE = "ncbi_protein_fasta_cache.json"
EUTILS_BASE = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils"


# ========= HELPERS =========
def infer_species_abbrev_from_filename(path: str) -> str:
    """
    Inferir abbrev: 1ª letra do género + 2 primeiras da espécie
    Ex.: Latimeria.chalumnae_... -> Lch
    """
    base = os.path.basename(path)
    base = re.sub(r"\.xlsx?$", "", base, flags=re.IGNORECASE)
    m = re.match(r"^([A-Z][a-z]+)[\._]([a-z]{2,})", base)
    if not m:
        return "Sp"
    genus, species = m.group(1), m.group(2)
    return f"{genus[0]}{species[:2]}"

def parse_int(value) -> int:
    """Converte para int (comprimento AA); se falhar devolve -1."""
    if value is None:
        return -1
    try:
        return int(float(str(value).strip()))
    except Exception:
        return -1

def wrap_seq(seq: str, width: int = 60) -> str:
    seq = re.sub(r"\s+", "", seq).strip()
    return "\n".join(seq[i:i+width] for i in range(0, len(seq), width))

def load_cache(path: str) -> Dict[str, str]:
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return {str(k): str(v) for k, v in data.items()}
        except Exception:
            return {}
    return {}

def save_cache(path: str, cache: Dict[str, str]) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)

def ncbi_req(url: str, params: dict, timeout: int = 30) -> str:
    params = dict(params)
    params["tool"] = TOOL_NAME
    params["email"] = NCBI_EMAIL
    if NCBI_API_KEY:
        params["api_key"] = NCBI_API_KEY
    r = requests.get(url, params=params, timeout=timeout)
    r.raise_for_status()
    return r.text

def fetch_protein_fasta_by_accession(accession: str) -> Optional[str]:
    """
    Busca FASTA de aminoácidos no db=protein usando o accession.
    1) efetch direto por id=accession
    2) fallback: esearch por accession -> uid -> efetch
    """
    acc = accession.strip()
    if not acc:
        return None

    # 1) tentativa direta
    try:
        txt = ncbi_req(
            f"{EUTILS_BASE}/efetch.fcgi",
            {"db": "protein", "id": acc, "rettype": "fasta", "retmode": "text"},
        )
        if txt.lstrip().startswith(">"):
            return txt
    except Exception:
        pass

    # 2) fallback esearch -> uid -> efetch
    try:
        es = ncbi_req(
            f"{EUTILS_BASE}/esearch.fcgi",
            {"db": "protein", "term": f"{acc}[Accession]", "retmode": "json", "retmax": 1},
        )
        data = json.loads(es)
        ids = data.get("esearchresult", {}).get("idlist", [])
        if not ids:
            return None
        uid = ids[0]
        txt = ncbi_req(
            f"{EUTILS_BASE}/efetch.fcgi",
            {"db": "protein", "id": uid, "rettype": "fasta", "retmode": "text"},
        )
        if txt.lstrip().startswith(">"):
            return txt
    except Exception:
        return None

    return None

def extract_sequence_from_fasta(fasta_text: str) -> str:
    lines = [ln.strip() for ln in fasta_text.splitlines() if ln.strip()]
    return "".join(ln for ln in lines if not ln.startswith(">"))


# ========= MAIN =========
def main():
    if "@" not in NCBI_EMAIL:
        raise SystemExit("Define um email válido em NCBI_EMAIL (no script ou via variável de ambiente).")

    abbrev = infer_species_abbrev_from_filename(INPUT_XLSX)

    wb = load_workbook(INPUT_XLSX)
    cache = load_cache(CACHE_FILE)

    for sheet_name in wb.sheetnames[1:]:

        ws = wb[sheet_name]
        print(f"\n==> A processar folha: {sheet_name}")
        max_row = ws.max_row

        # geneid -> lista de (row, prot_len, accession)
        gene_map: Dict[str, List[Tuple[int, int, str]]] = {}

        for row in range(1, max_row + 1):
            geneid_val = ws[f"{COL_GENEID}{row}"].value
            if geneid_val is None:
                continue
            geneid = str(geneid_val).strip()
            if not geneid or geneid.lower() in {"geneid", "gene_id"}:
                continue

            prot_len = parse_int(ws[f"{COL_PROT_LEN}{row}"].value)
            acc_val = ws[f"{COL_ACCESSION}{row}"].value
            accession = str(acc_val).strip() if acc_val is not None else ""
            gene_map.setdefault(geneid, []).append((row, prot_len, accession))

        # decidir vencedores por geneid: maior prot_len; em empate fica a primeira
        selected_rows = set()
        skipped_rows = set()

        for geneid, entries in gene_map.items():
            if len(entries) == 1:
                selected_rows.add(entries[0][0])
                continue

            # Preferir NP_ quando existem duplicados (anotadas); caso contrário manter lógica por comprimento
            np_entries = [e for e in entries if str(e[2]).startswith("NP_")]
            candidates = np_entries if np_entries else entries

            # escolhe o maior comprimento; empate -> mantém o primeiro (ordem do Excel)
            best_row, best_len, _best_acc = candidates[0]
            for r, l, acc in candidates[1:]:
                if l > best_len:
                    best_row, best_len, _best_acc = r, l, acc

            selected_rows.add(best_row)
            for r, _l, _acc in entries:
                if r != best_row:
                    skipped_rows.add(r)

        # marcar skipped (opcional)
        for r in skipped_rows:
            ws[f"{COL_FLAG}{r}"].value = "DUP_SKIP"

        # fetch FASTA e escrever em Z apenas nas selecionadas (inclui não-duplicadas)
        for row in range(1, max_row + 1):
            acc_val = ws[f"{COL_ACCESSION}{row}"].value
            if acc_val is None:
                continue

            accession = str(acc_val).strip()
            if not accession or accession.lower() in {"accession", "accession_number", "acc"}:
                continue

            geneid_val = ws[f"{COL_GENEID}{row}"].value
            geneid = str(geneid_val).strip() if geneid_val is not None else ""

            # Se tem geneid e não foi selecionada -> skip
            if geneid and (row not in selected_rows):
                continue

            # fetch (com cache)
            if accession in cache:
                fasta_raw = cache[accession]
            else:
                fasta_raw = fetch_protein_fasta_by_accession(accession)
                cache[accession] = fasta_raw or ""
                time.sleep(SLEEP_SECONDS)

            if not fasta_raw:
                ws[f"{COL_FASTA}{row}"].value = None
                continue

            seq = extract_sequence_from_fasta(fasta_raw)
            if not seq:
                ws[f"{COL_FASTA}{row}"].value = None
                continue
            # ALERTA: sequência retida não começa por Metionina (M)
            if not seq.startswith("M"):
                print(
                    f"[ALERTA N-TERM] {sheet_name} | linha {row} | acc={accession} | "
                    f"GeneID={geneid} | start with '{seq[0]}'"
                )

            header = f">{abbrev}_{accession}"
            ws[f"{COL_FASTA}{row}"].value = header + "\n" + wrap_seq(seq, width=60)

            if row % 25 == 0:
                print(f"  ... linha {row}/{max_row}")

    wb.save(OUTPUT_XLSX)
    save_cache(CACHE_FILE, cache)

    print(f"\n[OK] Guardado: {OUTPUT_XLSX}")
    print(f"[OK] Cache: {CACHE_FILE}")


if __name__ == "__main__":
    main()
