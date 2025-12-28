# usgs-pdf-table-extractor

Reproducible pipeline for layout-faithful extraction of tables from ***USGS Mineral Commodity Summary*** PDFs, which are not provided in a machine-readable format. The pipeline converts PDF → DOCX → raw Excel tables, preserving structure and artifacts and flagging superscripts and subscripts. Semantic reconstruction and normalization are intentionally out of scope and handled in a separate downstream stage.


### **1. SCOPE and NON-GOALS**

**What this project does:**

- Converts PDFs to DOCX deterministically
- Extracts tables as-is, preserving layout artifacts
- Flags superscripts and subscripts at the cell level
- Produces auditable raw Excel outputs traceable to the source PDF

**What this project does not do:**
- Infer table semantics
- Fix misaligned or hierarchical tables
- Normalize country names or reconstruct totals

**These tasks belong to a separate normalization and analysis stage.**

## **2. REPOSITORY STRUCTURE**

```text
src/docx extractor/
  pdf_converter.py       # Wrapper around pdf2docx for reproducible PDF → DOCX conversion
  docx_extractor.py      # Core extraction logic: DOCX parsing, table capture, flagging, export
  utils.py               # Utility functions (section headings, minimal cleaning, sheet naming)

scripts/
  run_extraction.py      # Batch entry point for reproducible, headless extraction
  run_gui.py             # Optional graphical interface for interactive inspection and QA
  run_extraction.bat     # Windows launcher for run_extraction.py
  run_gui.bat            # Windows launcher for run_gui.py

examples/
  input/                 # Small illustrative files only
  output/

data/                    # gitignored
  raw_pdf/
  intermediate_docx/
  raw_xlsx/

```

## **3. INSTALLATION:**
```text
pip install -r requirements.txt
```
Key dependencies include:
- pdf2docx
- python-docx
- pandas

## **4. HOW TO RUN:**

**Option A**: Batch extraction (recommended)
```text
python scripts/run_extractor.py
```

Use this for:
- full-year or multi-file processing
- reproducible data generation
- server or CI runs

**Option B:** GUI (optional, for inspection)
```text
python scripts/run_gui.py
```

Use this for:
- selecting individual files
- inspecting raw tables
- QA and debugging

**For large batches, use the batch script rather than the GUI.**

## **4. OUTPUTS**

For each input PDF, the pipeline produces:

- a DOCX file (optional, retained for audit)
- a raw Excel file containing:

  -  extracted tables
  -  superscript (_SUP) and subscript (_SUB) flag sheets

All outputs are **fully traceable back to the source PDF.**

## **5. KNOWN LIMITATIONS:**
- Tables with hierarchical or multi-line row structures are exported faithfully but may appear misaligned in Excel.
- Some values may appear in unexpected columns due to layout constraints.
- These cases are handled in a **separate normalization project.**

## **6. INTENDED AUDIENCE:**
This codebase is intended for:

- research and policy analysis
- ESG and resource governance applications
- reproducible data extraction from semi-structured PDFs

## **7. LICENSE:**

MIT License. See `LICENSE` for details.
