# Business Document Automation Toolkit

> **Portfolio demonstration project.** All organisations, records and content are fictional or synthetic. This project is not presented as prior client work.

A reusable data-to-document workflow that accepts **CSV or Excel (.xlsx)** input, validates records before generation, and produces consistent **editable DOCX and PDF** deliverables plus explicit processing reports.

## Business problem

Teams often maintain operational data in spreadsheets and then manually copy the same information into reports, delivery notes, summaries or client documents. That process is slow, difficult to audit and prone to copy/paste errors.

This project demonstrates a safer pattern:

**CSV / XLSX → schema checks → record validation → DOCX + PDF → manifest / rejection report**

## What it does

- accepts `.csv` and `.xlsx` source files;
- supports explicit worksheet selection for multi-sheet Excel workbooks;
- fails fast when required input columns are missing;
- validates required fields and ISO dates;
- rejects due dates earlier than creation dates;
- detects duplicate record IDs instead of silently overwriting outputs;
- validates structured `Label=Value` metrics;
- omits optional document sections when their source data is empty;
- creates native Word tables/lists and editable `.docx` files;
- creates matching `.pdf` deliverables;
- writes a generation manifest, rejection report and JSON run summary.

## Demonstration input

Two equivalent synthetic examples are provided:

- [`examples/records.csv`](examples/records.csv)
- [`examples/records.xlsx`](examples/records.xlsx), worksheet `Records`

The Excel example is deliberately presented as a normal business workbook rather than a raw machine export. It contains an instruction sheet, formatted headers, frozen panes and typed date cells.

Both datasets contain three valid records and one deliberately invalid row. The invalid row is rejected because its required `document_title` is missing.

## Example output

The repository includes three editable DOCX examples plus the generated manifest/rejection reports. The generator can also produce matching PDFs; PDF output is exercised in the automated tests and CI workflow rather than stored as duplicate binary examples.

## Run with CSV

```bash
python -m pip install -r requirements.txt
python src/generate_documents.py \
  --input examples/records.csv \
  --brand examples/brand.json \
  --output examples/generated \
  --format both
```

## Run with Excel

```bash
python src/generate_documents.py \
  --input examples/records.xlsx \
  --sheet Records \
  --brand examples/brand.json \
  --output examples/generated \
  --format both
```

`--format` accepts `docx`, `pdf` or `both`.

## Validation approach

The generator treats uncertain or malformed input as a data-quality problem rather than something to guess around. A record with invalid dates, duplicate identity or malformed metrics is isolated in the validation report and no document is generated for that record.

The automated test suite covers both input formats, generated DOCX/PDF structure, duplicate handling, date rules, strict metrics parsing, missing columns and optional sections. Generated examples are also rendered during development for visual QA.

See [`TEST_MATRIX.md`](TEST_MATRIX.md) for the acceptance matrix.

## Tools

`Python` · `CSV` · `Excel/XLSX` · `openpyxl` · `python-docx` · `ReportLab` · `pytest`

## Deliberate scope

This is a public demonstration, not a universal document platform. It intentionally does not include client-specific mappings, credentials, databases, email delivery, private workflow integrations or a reusable commercial engine.
