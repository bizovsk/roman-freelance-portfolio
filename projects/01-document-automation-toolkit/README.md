# Business Document Automation Toolkit

> **Portfolio demonstration project.** All organisations, records and content are fictional or synthetic. This project is not presented as prior client work.

A small document-automation system that turns structured CSV data into consistent, editable Microsoft Word documents while separating invalid input for review.

## What it demonstrates

This project is intentionally general rather than modelled on a particular company or job brief. It demonstrates a reusable workflow that can be adapted to many document-heavy processes:

**structured input → validation → document generation → manifest / error report**

The demo reads records from CSV, applies a lightweight brand configuration from JSON and creates one DOCX file per valid record.

## Behaviour

- required fields are validated before generation;
- invalid dates and missing required values are reported instead of guessed;
- optional sections are omitted when no data is provided;
- key/value metrics become native Word tables;
- next steps become native numbered lists;
- headers, footers and page numbers use native Word structure;
- generated files remain editable and do not rely on floating text boxes for core content.

## Demonstration data

[`examples/records.csv`](examples/records.csv) contains synthetic records with different content lengths and one deliberately invalid row.

[`examples/brand.json`](examples/brand.json) contains a fictional organisation name and simple brand settings.

The example output contains:

- three generated DOCX files;
- `generation_manifest.csv` listing successful outputs;
- `validation_report.csv` explaining why the invalid row was rejected.

## Run locally

```bash
python -m pip install -r requirements.txt
python src/generate_documents.py \
  --input examples/records.csv \
  --brand examples/brand.json \
  --output examples/generated
pytest -q
```

## Validation

Automated tests cover successful batch generation, invalid-input rejection, optional-section handling, date validation, metric parsing and native DOCX structure.

The generated examples are also rendered during development for visual QA. The same generator can be rerun with different synthetic input to verify that layout and validation behaviour are not tied to one fixed dataset.

## Why this project exists

The commercial value is not the presence of a Word file. It is the ability to turn structured information into repeatable business documents without manually copying values into each document, while making data problems visible instead of silently inventing replacements.

## Tools

`Python` · `CSV` · `JSON` · `python-docx` · `pytest` · `Microsoft Word / DOCX structure`
