# SPL Appendix Generator

Generates the traceability appendices **A**, **B** and **C** from a QR_TS_SYS
Test Specification export and appends them into a **copy** of the document.
Two ways to run it: a **browser tool** (no install) and a **Python script**
(for batch / legacy `.doc` handling). Both produce identical output.

| Appendix | Content | Built from |
|---|---|---|
| **A. Requirement Traceability Matrix** | Requirement → Safety class → Test Case → Test Suite | Export + RVTM/SRS |
| **B. Test Suites Traceability Matrix** | Test Suite → Description / Execution / Test Cases | Export |
| **C. Test Cases Traceability Matrix** | Test Case → Description / Test Suites | Export |

The original file is never modified — the appendices are added to a fresh copy.

---

## Option 1 — Browser tool (recommended)

Open **`index.html`** (or `SPL_Appendix_Generator.html`) in Chrome or Edge, or
use the hosted link if GitHub Pages is enabled:

> `https://<your-username>.github.io/<your-repo>/`

**Steps**

1. Drop in the test-specification **`.docx`** export.
2. (Optional) Drop in the **RVTM / SRS** workbook (`.xlsx`) for the Appendix A safety column.
3. Choose which appendices to generate (A / B / C).
4. Click **Generate**, then click **Download Word document** in the readout panel.

Everything runs locally in your browser — **no file is ever uploaded**. It works
fully offline; hosting on GitHub Pages just gives you a convenient URL.

> **Input must be `.docx`.** A browser cannot read the legacy binary `.doc`
> format. If your export is `.doc`, open it in Word and **Save As → .docx**
> first (or use the Python script below, which converts `.doc` automatically).

---

## Option 2 — Python script

For batch runs, or to process a legacy `.doc` directly.

**Install (one time)**

```bash
pip install python-docx openpyxl
```

LibreOffice (`soffice`) is only needed if the input is a legacy `.doc`.

**Run**

```bash
# Appendices A, B, C with the safety column filled from the RVTM:
python generate_appendices.py EXPORT.docx --safety RVTM.xlsx

# .doc input is converted automatically:
python generate_appendices.py EXPORT.doc --safety RVTM.xlsx
```

Output: `EXPORT_with_appendices.docx` (override with `-o OUT.docx`).

**Options**

| Option | Purpose |
|---|---|
| `--safety FILE` | RVTM/SRS `.xlsx` or `.csv` supplying the safety classification |
| `--req-col NAME` | Requirement-ID column in the lookup (auto-detected if omitted) |
| `--safety-col NAME` | Safety-class column in the lookup (auto-detected if omitted) |
| `--appendices ABC` | Build only a subset (any of A B C) |
| `--srs-rev "11.00 [5]"` | SRS revision text in the Appendix A header |
| `--placeholder TBC` | Value used when safety class is unknown (no RVTM loaded) |
| `-o OUT.docx` | Output path |

Run `python generate_appendices.py -h` for full help.

---

## The RVTM / SRS lookup

An `.xlsx` (or `.csv`) with a **requirement-ID** column (values like `SYS_2841`)
and a **safety** column (e.g. `Safety Related` / `Not Safety Related`; also
accepts Yes/No, SR/NSR). Both tools auto-detect the right sheet and columns.

**Safety-Related handling in Appendix A**

- Requirement **found** in the RVTM → its class is written into the safety column.
- Requirement **not found** in the RVTM → it is **highlighted in the readout and
  excluded from Appendix A** (it is not listed with a placeholder).
- **No RVTM loaded** → the safety column shows the placeholder (`TBC`) and nothing
  is excluded.

---

## How it reads the export

Each `QR_TS_SYS_*` suite is parsed for its description, Execution type, its
`QR_TC_SYS_*` test cases (letter suffixes such as `…0143b` included), and the
`SYS_####` requirements referenced in its steps. A `SYS_####` that is part of a
test-case/suite reference (e.g. `QR_TC_SYS_0102`) is **not** treated as a
requirement. Requirement content held inside nested tables is excluded, matching
the established methodology.

---

## Repository contents

| File | Purpose |
|---|---|
| `index.html` / `SPL_Appendix_Generator.html` | The browser tool (single self-contained file) |
| `generate_appendices.py` | Command-line / batch version |
| `README.md` | This file |

---

## Notes

- Built for the SEQ Cross River Rail ETCS L2 SPL milestone V&V documentation workflow.
- The browser tool uses [JSZip](https://stuk.github.io/jszip/) and
  [SheetJS](https://sheetjs.com/) loaded from a CDN; an internet connection is
  needed the first time it loads those libraries (the document processing itself
  is entirely local).
