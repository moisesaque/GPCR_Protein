## Excel–NCBI Annotation Pipeline

**GeneID mapping → isoform curation & FASTA export → exon count from RefSeq GFF**

---

## Overview

This pipeline processes Excel files containing **NCBI protein accession numbers (XP_/NP_)** and progressively enriches them with:

1. **GeneID annotation** (NCBI Gene)
2. **Isoform curation + protein FASTA export**
3. **Exact exon count per isoform**, derived from the **RefSeq genome annotation (GFF3)**

The pipeline is designed to:

* handle **multiple isoforms per gene**
* preserve **isoform-specific exon structures**
* support **predicted RefSeq annotations (XM_/XP_)**
* avoid ambiguity introduced by Gene-level exon counts

Each step produces a new Excel file that becomes the input for the next step.

---

## Requirements

### Software

* Python ≥ 3.9
* Internet connection (NCBI Entrez access)

### Python dependencies

```bash
pip install openpyxl requests
```

### NCBI usage requirements

NCBI requires user identification for programmatic access.

NCBI_EMAIL "your.email@institution.edu"

NCBI_API_KEY "YOUR_NCBI_API_KEY"


---

## Input Excel file (initial)

* **Row 1**: header (ignored by all scripts)
* **Column A**: NCBI **protein accession number** (XP_, NP_, etc.)
* Multiple sheets allowed (e.g. `CALCR`, `CRHR`, `PTHR`, `GCGR`, `VIP SCRT`)

---

## Step 1 — Map protein accessions to GeneID

**Script:** `1_GeneID.py` 

### Purpose

* Resolves each protein accession (column A) to its corresponding **NCBI GeneID**
* Writes GeneID to **column Y**
* Uses NCBI `esearch` + `elink`
* Caches results to avoid redundant queries

### Input

* Excel file with protein accessions in **column A**

### Output

* New Excel file with:

  * **Column Y** → GeneID

### Run

```bash
python 1_GeneID.py
```

---

## Step 2 — Isoform curation and protein FASTA export

**Script:** `2_ProteinCuratingSequence.py` 

### Purpose

For each GeneID **within the same sheet**:

1. Detects duplicated GeneIDs (multiple isoforms)
2. Keeps **only the longest protein isoform**

   * based on **protein length (aa) in column C**
3. Marks discarded isoforms with `DUP_SKIP` (column AA)
4. Downloads the **protein amino acid sequence** from NCBI
5. Writes a **custom FASTA** to **column Z**

### FASTA format (per cell)

```
>Lch_XP_006010891.1
MPALIMEKKWAQFLLILSV...
```

Species abbreviation is inferred automatically from the Excel filename:

* `Latimeria.chalumnae` → `Lch`

### Input

* Excel file produced by **Step 1**
* Required columns:

  * A → protein accession
  * C → protein length (aa)
  * Y → GeneID

### Output

* New Excel file with:

  * **Column Z** → protein FASTA
  * **Column AA** → `DUP_SKIP` (optional)

### Run

```bash
python 2_ProteinCuratingSequence.py
```

---

## Step 3 — Isoform-specific exon count from RefSeq GFF

**Script:** `3_ExonCount_2.py` 

### Purpose

Computes the **exact exon count per isoform**, using the **RefSeq genome annotation (GFF3)** instead of transcript GenBank records.

This is essential because:

* Many XM_/XP_ records **do not list exon coordinates**
* Exon structure is defined in the **assembly-level GFF**
* Different isoforms of the same gene can have **different exon counts**

### Method

1. Parse the RefSeq **genomic GFF3**
2. Build mappings:

   * `protein_id (XP_) → transcript_id`
   * `transcript_id → number of exons`
3. For each protein accession in column A:

   * Find its parent transcript
   * Count `exon` features
4. Write exon count to **column X**

### Input

* Excel file produced by **Step 2**
* RefSeq **GFF3 or GFF3.GZ** file for the same assembly used to generate XP_/XM_

Example:

```
GCF_018977255.1_IMCB_Cmil_1.0_genomic.gff.gz
```

> ⚠️ The GFF **must match the RefSeq assembly version** used for the annotations, otherwise exon counts may be missing.

### Output

* Final Excel file with:

  * **Column X** → isoform-specific exon count

### Run

```bash
python 3_ExonCount_2.py
```

---

## Final Excel structure (relevant columns)

| Column | Content                                 |
| ------ | --------------------------------------- |
| A      | Protein accession (XP_/NP_)             |
| C      | Protein length (aa)                     |
| X      | **Exact exon count (isoform-specific)** |
| Y      | GeneID                                  |
| Z      | Protein FASTA                           |
| AA     | `DUP_SKIP` (discarded isoforms)         |

---

## Notes & limitations

* **WP_ accessions** may not map to a unique transcript → exon count may be `None`
* If exon count is missing:

  * assembly mismatch is the most common cause
  * check GFF version vs XP_/XM_ release
* Exon counts include **predicted exons** (RefSeq pipeline)
