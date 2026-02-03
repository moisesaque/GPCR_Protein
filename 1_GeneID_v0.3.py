import os
import time
import json
import requests
from typing import Optional, Dict, Tuple, List
import re
from json import JSONDecodeError

from openpyxl import load_workbook

# ====== CONFIGURAÇÕES ======
INPUT_XLSX = r"PATH.xlsx"   # ajusta se necessário
OUTPUT_XLSX = r"PATH.xlsx"

ACCESSION_COL = "A"
OUTPUT_COL = "Y"

# O NCBI pede que identifiques um email válido no tool usage
NCBI_EMAIL = "email"
NCBI_API_KEY = "API KEY"  # opcional, mas recomendado se tiveres
TOOL_NAME = "accession_to_geneid_mapper"

# Rate limit: sem API key ~3 req/s (na prática, usa 0.34s). Com API key podes ir mais alto,
# mas mantém “polite usage”.
SLEEP_SECONDS = 0.34 if not NCBI_API_KEY else 0.12

CACHE_FILE = "ncbi_geneid_cache.json"

EUTILS_BASE = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils"


_CONTROL_CHARS_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")

def _clean_json_text(s: str) -> str:
    return _CONTROL_CHARS_RE.sub("", s)

def _req_json(url: str, params: dict, timeout: int = 30, retries: int = 4, backoff: float = 0.8) -> dict:
    """
    Faz GET e tenta interpretar JSON.
    Retry automático quando:
      - JSON inválido (JSONDecodeError)
      - resposta parece HTML/XML (começa por '<') ou vazia
    """
    last_err: Optional[Exception] = None

    for attempt in range(1, retries + 1):
        try:
            text = _clean_json_text(_req(url, params, timeout=timeout))
            stripped = text.lstrip()

            if not stripped:
                raise ValueError("Empty response")
            if stripped.startswith("<"):
                raise ValueError("Non-JSON (HTML/XML) response")

            return json.loads(text)

        except (JSONDecodeError, ValueError) as e:
            last_err = e
            if attempt < retries:
                time.sleep(backoff * attempt)
                continue
            raise

    raise last_err if last_err else RuntimeError("Unknown JSON parse error")



# ====== UTILITÁRIOS NCBI ======
def _req(url: str, params: dict, timeout: int = 30) -> str:
    """GET com identificação NCBI e tratamento simples."""
    # Identificação recomendada pelo NCBI
    params = dict(params)
    params["tool"] = TOOL_NAME
    params["email"] = NCBI_EMAIL
    if NCBI_API_KEY:
        params["api_key"] = NCBI_API_KEY

    r = requests.get(url, params=params, timeout=timeout)
    r.raise_for_status()
    return r.text


def esearch_uid(db: str, accession: str) -> Optional[str]:
    """
    Procura UID (entrez id) em 'protein' ou 'nuccore' a partir de um accession.
    """
    url = f"{EUTILS_BASE}/esearch.fcgi"
    # Termo robusto: accession[Accession]
    params = {"db": db, "term": f"{accession}[Accession]", "retmode": "json", "retmax": 5}
    data = _req_json(url, params)
    ids = data.get("esearchresult", {}).get("idlist", [])
    return ids[0] if ids else None


def elink_geneids(dbfrom: str, uid: str) -> List[str]:
    """
    Faz link do UID (protein/nuccore) para gene, devolvendo lista de GeneIDs.
    """
    url = f"{EUTILS_BASE}/elink.fcgi"
    params = {"dbfrom": dbfrom, "db": "gene", "id": uid, "retmode": "json"}
    data = _req_json(url, params)

    geneids: List[str] = []
    for linkset in data.get("linksets", []):
        for db in linkset.get("linksetdbs", []):
            if db.get("dbto") == "gene":
                for gid in db.get("links", []):
                    geneids.append(str(gid))
    return geneids


def accession_to_geneid(accession: str) -> Optional[str]:
    """
    Resolve um accession -> GeneID.
    Estratégia:
      1) tenta em protein (esearch->elink)
      2) se falhar, tenta em nuccore (esearch->elink)
    """
    acc = accession.strip()
    if not acc:
        return None

    # 1) protein
    uid = esearch_uid("protein", acc)
    if uid:
        gids = elink_geneids("protein", uid)
        if gids:
            return gids[0]

    # 2) nuccore
    uid = esearch_uid("nuccore", acc)
    if uid:
        gids = elink_geneids("nuccore", uid)
        if gids:
            return gids[0]

    return None


# ====== CACHE ======
def load_cache(path: str) -> Dict[str, Optional[str]]:
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            # normaliza chaves
            return {str(k): (str(v) if v is not None else None) for k, v in data.items()}
        except Exception:
            return {}
    return {}


def save_cache(path: str, cache: Dict[str, Optional[str]]) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


# ====== PROCESSAMENTO EXCEL ======
def main():
    if "@" not in NCBI_EMAIL:
        raise SystemExit("Define um email válido em NCBI_EMAIL (no script ou via variável de ambiente).")

    wb = load_workbook(INPUT_XLSX)
    cache = load_cache(CACHE_FILE)

    # registo de ocorrências para alertar quando o mesmo GeneID aparece em folhas diferentes
    geneid_hits: Dict[str, List[Tuple[str, int, str]]] = {}

    # === ALTERAÇÃO: processar todas as folhas menos a primeira ===
    for sheet_name in wb.sheetnames[1:]:
        ws = wb[sheet_name]
        print(f"\n==> A processar folha: {sheet_name}")

        # percorre linhas com base na coluna A (accession)
        max_row = ws.max_row
        for row in range(2, max_row + 1):
            cell = ws[f"{ACCESSION_COL}{row}"]
            acc = cell.value

            if acc is None:
                continue

            acc_str = str(acc).strip()
            if not acc_str:
                continue

            if acc_str in cache:
                geneid = cache[acc_str]
            else:
                try:
                    geneid = accession_to_geneid(acc_str)
                except requests.HTTPError as e:
                    print(f"[ERRO HTTP] {sheet_name} linha {row} acc={acc_str}: {e}")
                    geneid = None
                except Exception as e:
                    print(f"[ERRO] {sheet_name} linha {row} acc={acc_str}: {e}")
                    geneid = None

                cache[acc_str] = geneid
                time.sleep(SLEEP_SECONDS)

            ws[f"{OUTPUT_COL}{row}"].value = geneid

            if geneid:
                geneid_hits.setdefault(str(geneid), []).append((sheet_name, row, acc_str))


            if row % 25 == 0:
                print(f"  ... linha {row}/{max_row}")

    # alerta: mesmo GeneID em folhas diferentes (mostra linhas e accessions)
    for gid, hits in geneid_hits.items():
        sheets = {h[0] for h in hits}
        if len(sheets) > 1:
            sheets_str = ", ".join(sorted(sheets))
            print(f"[ALERTA] GeneID {gid} aparece em folhas diferentes: {sheets_str}")
            for sh, r, acc in hits:
                print(f"  - {sh} linha {r} acc={acc}")

    wb.save(OUTPUT_XLSX)
    save_cache(CACHE_FILE, cache)

    print(f"\n[OK] Guardado: {OUTPUT_XLSX}")
    print(f"[OK] Cache guardada: {CACHE_FILE}")

if __name__ == "__main__":
    main()
