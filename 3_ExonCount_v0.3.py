import os
import time
import re
from typing import Dict, Optional, Iterable, Tuple

from urllib.error import HTTPError, URLError

from openpyxl import load_workbook

# Biopython / Entrez
from Bio import Entrez, SeqIO

# ========= CONFIG =========
INPUT_XLSX = r"PATH_selected_and_fasta.xlsx"
OUTPUT_XLSX = r"PATH_with_ExonCount.xlsx"

COL_ACCESSION = "A"   # XP_...
COL_EXONCOUNT = "X"   # saída: nº exões
COL_CHROM = "AA"      # saída: cromossoma/replicão (ex: NC_..., NW_..., NZ_...)

# ========= NCBI ENTREZ CONFIG =========
# Obrigatório por políticas do NCBI:
Entrez.email = "email"
# Opcional (recomendado): acelera limites e reduz erros 429
Entrez.api_key = "API KEY"

# Respeitar rate limits (sem API key: tipicamente ~3 req/seg; com key pode ser mais)
ENTREZ_SLEEP_SECONDS = 0.34



# ========= ENTREZ ROBUSTEZ (retry/backoff) =========
# Alguns erros são transitórios (EOF / ligação fechada, timeouts, 429/5xx). Fazemos retry com backoff.
def _should_retry_exception(e: Exception) -> bool:
    msg = str(e)
    # Erros transitórios comuns reportados pelo Entrez/Biopython
    transient_markers = [
        "Read failed: EOF",
        "the other side has unexpectedly closed connection",
        "timed out",
        "timeout",
        "Temporary failure",
        "temporarily unavailable",
        "Connection reset by peer",
        "Connection aborted",
        "EOF",
    ]
    return any(m in msg for m in transient_markers)


def _entrez_read_with_retries(handle_factory, *, retries: int = 4, backoff: float = 0.8):
    """Cria um handle (via handle_factory), faz Entrez.read, com retry em erros transitórios."""
    last_exc: Optional[Exception] = None
    for attempt in range(1, retries + 1):
        h = None
        try:
            h = handle_factory()
            data = Entrez.read(h)
            return data
        except HTTPError as e:
            # 400 costuma ser erro "hard" (id inválido, parâmetros inválidos). Não vale a pena insistir.
            if getattr(e, "code", None) == 400:
                raise
            # 429/5xx: retry
            if getattr(e, "code", None) in {429, 500, 502, 503, 504} and attempt < retries:
                last_exc = e
                time.sleep(backoff * attempt)
                continue
            raise
        except (URLError, RuntimeError, OSError, Exception) as e:
            # RuntimeError: muitas vezes vem do parser/IO do Entrez
            if _should_retry_exception(e) and attempt < retries:
                last_exc = e
                time.sleep(backoff * attempt)
                continue
            raise
        finally:
            try:
                if h is not None:
                    h.close()
            except Exception:
                pass
    if last_exc:
        raise last_exc
    raise RuntimeError("Entrez.read failed")


# ========= HELPERS =========
def _sleep_polite():
    time.sleep(ENTREZ_SLEEP_SECONDS)


def fetch_protein_genbank_record(prot_acc: str):
    """
    Faz download do registo GenBank da proteína (db=protein).
    Devolve um SeqRecord (Biopython).
    """
    handle = Entrez.efetch(db="protein", id=prot_acc, rettype="gb", retmode="text")
    try:
        rec = SeqIO.read(handle, "genbank")
    finally:
        handle.close()
    _sleep_polite()
    return rec


def parse_coded_by_to_exoncount_and_chrom(coded_by: str) -> Tuple[Optional[int], Optional[str]]:
    """
    coded_by típico:
      join(NC_000001.11:123..456,NC_000001.11:789..900)
      complement(join(NC_...:...,NC_...:...))
      NC_...:123..456 (sem join -> 1 exão)

    Estratégia:
      - contar segmentos dentro de join(...)
      - cromossoma = accession antes de ':' do 1º segmento
    """
    if not coded_by:
        return None, None

    s = coded_by.strip()

    exon_count: Optional[int] = None
    chrom: Optional[str] = None

    def _first_accession(seg: str) -> Optional[str]:
        seg = seg.strip()
        if seg.startswith("complement(") and seg.endswith(")"):
            seg = seg[len("complement("):-1].strip()
        if ":" in seg:
            return seg.split(":", 1)[0].strip() or None
        return None

    if "join(" in s:
        join_start = s.find("join(") + len("join(")
        join_end = s.rfind(")")
        inner = s[join_start:join_end] if join_end > join_start else s[join_start:]

        parts = [p.strip() for p in inner.split(",") if p.strip()]
        exon_count = len(parts) if parts else None
        chrom = _first_accession(parts[0]) if parts else None
    else:
        exon_count = 1
        chrom = _first_accession(s)

    return exon_count, chrom


# ======= NOVO: protein -> gene -> esummary (para ExonCount e Chromosome) =======
def _protein_to_gene_id(prot_acc: str) -> Optional[str]:
    """
    Mapeia XP_/NP_ -> GeneID via:
      1) esearch(db=protein) para obter UID numérico
      2) elink(dbfrom=protein, db=gene) usando esse UID
    Isto evita alguns HTTP 400 quando o elink não aceita o accession diretamente.
    """
    prot = prot_acc.strip()
    if not prot:
        return None

    # 1) Protein UID
    esearch_data = _entrez_read_with_retries(
        lambda: Entrez.esearch(db="protein", term=f"{prot}[Accession]", retmax=1)
    )
    uid_list = esearch_data.get("IdList", []) if isinstance(esearch_data, dict) else []
    if not uid_list:
        return None
    uid = uid_list[0]

    # 2) UID -> GeneID
    elink_data = _entrez_read_with_retries(
        lambda: Entrez.elink(dbfrom="protein", db="gene", id=uid)
    )

    linksetdb = elink_data[0].get("LinkSetDb", []) if elink_data else []
    if not linksetdb:
        return None
    links = linksetdb[0].get("Link", [])
    if not links:
        return None
    return links[0].get("Id")


def _normalize_chromosome_number(chrom_raw) -> str:
    """
    Queremos: número do cromossoma (1..n) ou X/Y.
    Se não der (scaffold/unlocalized/unplaced/empty): 'unplaced'
    """
    if chrom_raw is None:
        return "unplaced"

    s = str(chrom_raw).strip()
    if not s:
        return "unplaced"

    u = s.upper()

    if s.isdigit():
        return s
    if u in {"X", "Y"}:
        return u

    # tenta extrair um número dentro da string (ex: "chromosome 12")
    m = re.search(r"\b(\d+)\b", s)
    if m:
        return m.group(1)

    return "unplaced"


def _gene_exoncount_and_chromosome_number(gene_id: str) -> Tuple[Optional[int], str]:
    """
    Usa Entrez.esummary(db=gene) para obter ExonCount + Chromosome.
    """
    h = Entrez.esummary(db="gene", id=gene_id, retmode="xml")
    s = Entrez.read(h)
    h.close()
    _sleep_polite()

    docset = s.get("DocumentSummarySet", {})
    docs = docset.get("DocumentSummary", [])
    if not docs:
        return None, "unplaced"

    doc = docs[0]
    genomic0 = (doc.get("GenomicInfo") or [{}])[0]

    exon_count = genomic0.get("ExonCount")
    chrom_raw = doc.get("Chromosome")

    return exon_count, _normalize_chromosome_number(chrom_raw)



def get_exoncount_and_chrom_from_protein(prot_acc: str) -> Tuple[Optional[int], Optional[str]]:
    """
    Obtém nº exões e cromossoma a partir de:
      protein -> gene -> gene esummary
    - ExonCount vem de GenomicInfo[0].ExonCount
    - Chromosome vem de DocumentSummary.Chromosome, normalizado para número/X/Y,
      e se não for possível: 'unplaced'
    """
    gene_id = _protein_to_gene_id(prot_acc)
    if not gene_id:
        return None, "unplaced"

    exon_count, chrom = _gene_exoncount_and_chromosome_number(gene_id)
    return exon_count, chrom


# ========= MAIN =========
def main():
    wb = load_workbook(INPUT_XLSX)
    sheetnames = wb.sheetnames[1:]

    # cache para não pedir ao NCBI várias vezes a mesma proteína
    cache: Dict[str, Tuple[Optional[int], Optional[str]]] = {}

    for sh in sheetnames:
        if sh not in wb.sheetnames:
            print(f"[AVISO] Folha '{sh}' não existe. A saltar.")
            continue

        ws = wb[sh]
        max_row = ws.max_row
        print(f"\n==> {sh} | linhas: {max_row}")

        for row in range(2, max_row + 1):  # linha 1 = cabeçalho
            acc_val = ws[f"{COL_ACCESSION}{row}"].value
            if acc_val is None:
                continue

            prot = str(acc_val).strip()
            if not prot:
                continue

            if prot not in cache:
                try:
                    exon_count, chrom = get_exoncount_and_chrom_from_protein(prot)
                except Exception as e:
                    print(f"[ERRO] {prot} (linha {row}): {e}")
                    exon_count, chrom = None, "unplaced"
                cache[prot] = (exon_count, chrom)

            exon_count, chrom = cache[prot]

            ws[f"{COL_EXONCOUNT}{row}"].value = exon_count
            ws[f"{COL_CHROM}{row}"].value = chrom

            if row % 25 == 0:
                print(f"  ... {row}/{max_row}")

    wb.save(OUTPUT_XLSX)
    print(f"\n[OK] Guardado: {OUTPUT_XLSX}")


if __name__ == "__main__":
    main()
