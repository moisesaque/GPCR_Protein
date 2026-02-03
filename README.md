GPCR Annotation Pipeline (NCBI-based)

This repository contains a stepwise Python pipeline to curate GPCR-related sequences starting from NCBI accessions stored in Excel files.
The workflow progressively annotates GeneIDs, removes redundant isoforms, retrieves protein FASTA sequences, infers exon counts and chromosome location, and finally exports a clean multi-FASTA file.

The scripts are designed to be run sequentially, as each step depends on the output of the previous one.

Requirements
Software

Python ≥ 3.9

Microsoft Excel files (.xlsx)

Python dependencies

Install required packages with:

pip install requests openpyxl biopython

NCBI requirements (mandatory)

NCBI requires user identification for automated queries.

You must define:

a valid email address

(recommended) an NCBI API key

These are hardcoded in the scripts and should be edited before use.

Example:

NCBI_EMAIL = "your.email@institution.edu"
NCBI_API_KEY = "your_ncbi_api_key"

Input data format

Input files are Excel (.xlsx)

First worksheet is ignored

All subsequent sheets represent gene families or receptor subgroups

Each sheet must contain:

Column A: NCBI protein accession (e.g. XP_, NP_)

Additional columns are added automatically by the scripts

⚠️ The pipeline assumes RefSeq-style accessions. GenBank-only or obsolete accessions may fail.

Pipeline overview
1️⃣ 1_GeneID_v0.3.py — Map accessions to GeneID

Purpose

Resolves each protein accession (XP_, NP_) to its corresponding NCBI GeneID

Uses esearch + elink via NCBI E-utilities

Results are cached locally to reduce API load

Input

Excel file with accessions in column A

Output

New Excel file with GeneID added (column Y)

JSON cache file:

ncbi_geneid_cache.json

1_GeneID_v0.3

2️⃣ 2_ProteinCuratingSequence_v0.1.py — Isoform deduplication + FASTA retrieval

Purpose

For each GeneID:

Keeps one representative protein

Preference order:

Annotated proteins (NP_)

Longest amino acid sequence

Downloads protein FASTA sequences

Flags skipped duplicate isoforms

Warns if protein does not start with Methionine (M)

Input

Excel produced in step 1

Output

Excel file with:

Curated FASTA sequence (column Z)

Optional duplicate flag (DUP_SKIP)

JSON cache:

ncbi_protein_fasta_cache.json

2_ProteinCuratingSequence_v0.1

3️⃣ 3_ExonCount_v0.3.py — Exon count and chromosome inference

Purpose

Infers number of exons and chromosome number

Strategy:

protein accession → GeneID

GeneID → Entrez.esummary

Uses GenomicInfo.ExonCount

Chromosome is normalized:

1–22, X, Y

otherwise labeled as unplaced

Important limitation

Exon counts correspond to RefSeq gene models

Alternative transcripts and isoform-specific exon structures are not resolved

Input

Excel produced in step 2

Output

Excel file with:

Exon count (column X)

Chromosome (column AA)

3_ExonCount_v0.3

4️⃣ 4_GenerateFasta_v0.1.py — Final multi-FASTA export

Purpose

Extracts curated FASTA sequences from Excel

Appends sheet name as a suffix to each FASTA header

Example:
>Lch_XP_012345678_CALCR

Merges all sheets into one FASTA file

Input

Excel produced in step 3

Output

Final multi-FASTA file (.fasta)

4_GenerateFasta_v0.1

Recommended execution order
1_GeneID_v0.3.py
→ 2_ProteinCuratingSequence_v0.1.py
→ 3_ExonCount_v0.3.py
→ 4_GenerateFasta_v0.1.py


⚠️ Skipping steps will break downstream scripts.
