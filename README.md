# GPCR Annotation Pipeline (NCBI-based)

This repository provides a **stepwise Python pipeline** for curating GPCR-related protein sequences starting from NCBI accessions stored in Excel files.
The workflow progressively annotates **GeneIDs**, removes redundant isoforms, retrieves **protein FASTA sequences**, infers **exon counts and chromosome location**, and finally exports a **clean multi-FASTA file**.

Each script must be executed **in order**, as downstream steps depend on the outputs of previous ones.

---

## Requirements

### Software

* Python ≥ 3.9
* Excel files (`.xlsx`)

### Python dependencies

Install required packages with:

```bash
pip install requests openpyxl biopython
```

### NCBI requirements (mandatory)

NCBI requires user identification for automated queries.

Before running the scripts, edit them to include:

* a valid email address
* (recommended) an NCBI API key

Example:

```python
NCBI_EMAIL = "your.email@institution.edu"
NCBI_API_KEY = "your_ncbi_api_key"
```

---

## Input data format

* Input files must be Excel (`.xlsx`)
* **The first worksheet is ignored**
* All subsequent worksheets are processed independently
* Each worksheet must contain:

  * **Column A**: NCBI protein accession (e.g. `XP_`, `NP_`)

Additional annotation columns are added automatically by the scripts.

⚠️ This pipeline assumes **RefSeq-style accessions**.
GenBank-only, obsolete, or poorly annotated accessions may fail.

---

## Pipeline overview

### 1️⃣ `1_GeneID_v0.3.py` — Accession → GeneID mapping

**Purpose**

* Maps each protein accession to its corresponding **NCBI GeneID**
* Uses `esearch` + `elink` via NCBI E-utilities
* Implements local caching to minimize redundant API requests

**Input**

* Excel file with protein accessions in column `A`

**Output**

* New Excel file with GeneID added (column `Y`)
* Cache file: `ncbi_geneid_cache.json`

---

### 2️⃣ `2_ProteinCuratingSequence_v0.1.py` — Isoform curation and FASTA retrieval

**Purpose**

* Resolves multiple protein isoforms per GeneID
* Keeps **one representative sequence per gene** using:

  1. Preference for annotated proteins (`NP_`)
  2. Longest amino acid sequence
* Retrieves **protein FASTA sequences**
* Flags discarded duplicate isoforms
* Warns if the retained protein does not start with Methionine (`M`)

**Input**

* Excel file generated in step 1

**Output**

* Excel file with:

  * Curated FASTA sequence (column `Z`)
  * Optional duplicate flag (`DUP_SKIP`)
* Cache file: `ncbi_protein_fasta_cache.json`

---

### 3️⃣ `3_ExonCount_v0.3.py` — Exon count and chromosome inference

**Purpose**

* Infers:

  * Number of exons
  * Chromosome assignment
* Strategy:

  * protein accession → GeneID
  * GeneID → `Entrez.esummary`
  * Exon count from `GenomicInfo.ExonCount`
* Chromosome normalization:

  * `1–22`, `X`, `Y`
  * otherwise labeled as `unplaced`

**Important limitation**

* Exon counts correspond to **RefSeq gene models**
* Transcript- or isoform-specific exon structures are **not resolved**

**Input**

* Excel file generated in step 2

**Output**

* Excel file with:

  * Exon count (column `X`)
  * Chromosome (column `AA`)

---

### 4️⃣ `4_GenerateFasta_v0.1.py` — Final multi-FASTA export

**Purpose**

* Extracts curated FASTA sequences from Excel
* Appends worksheet name to each FASTA header

Example:

```
>Lch_XP_012345678_CALCR
```

* Merges all sequences into a single FASTA file

**Input**

* Excel file generated in step 3

**Output**

* Final multi-FASTA file (`.fasta`)

---

## Execution order

The scripts **must** be run in the following order:

```
1_GeneID_v0.3.py
→ 2_ProteinCuratingSequence_v0.1.py
→ 3_ExonCount_v0.3.py
→ 4_GenerateFasta_v0.1.py
```

Skipping steps will break downstream scripts.
