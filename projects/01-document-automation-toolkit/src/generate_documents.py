from __future__ import annotations

import argparse
import csv
import json
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

REQUIRED_FIELDS = ("record_id", "document_title", "recipient", "owner", "status", "created_on", "summary")


@dataclass(frozen=True)
class Brand:
    organisation: str
    tagline: str
    accent_hex: str
    primary_hex: str
    footer: str

    @property
    def accent(self) -> RGBColor:
        return RGBColor.from_string(self.accent_hex)

    @property
    def primary(self) -> RGBColor:
        return RGBColor.from_string(self.primary_hex)


def load_brand(path: Path) -> Brand:
    data = json.loads(path.read_text(encoding="utf-8"))
    return Brand(**data)


def slug(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9]+", "_", value.strip())
    return value.strip("_") or "document"


def add_page_number(paragraph) -> None:
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, end])


def configure_document(doc: Document, brand: Brand) -> None:
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.top_margin = Mm(18)
    section.bottom_margin = Mm(18)
    section.left_margin = Mm(20)
    section.right_margin = Mm(20)

    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10.5)
    normal.paragraph_format.space_after = Pt(6)

    for name, size, colour in (
        ("Title", 24, brand.primary),
        ("Heading 1", 15, brand.primary),
        ("Heading 2", 11.5, brand.accent),
    ):
        style = doc.styles[name]
        style.font.name = "Arial"
        style.font.size = Pt(size)
        style.font.color.rgb = colour

    existing = {style.name for style in doc.styles}
    for name, size, bold, colour in (
        ("Toolkit Label", 8.5, True, RGBColor(105, 115, 125)),
        ("Toolkit Note", 9, False, RGBColor(105, 115, 125)),
    ):
        style = doc.styles[name] if name in existing else doc.styles.add_style(name, 1)
        style.font.name = "Arial"
        style.font.size = Pt(size)
        style.font.bold = bold
        style.font.color.rgb = colour

    header = section.header.paragraphs[0]
    header.text = f"{brand.organisation.upper()}  |  DOCUMENT AUTOMATION DEMO"
    header.style = doc.styles["Toolkit Label"]

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    footer.add_run(f"{brand.footer}  |  Page ")
    add_page_number(footer)


def add_metadata_table(doc: Document, row: dict[str, str]) -> None:
    pairs = [
        ("Record", row["record_id"]),
        ("Prepared for", row["recipient"]),
        ("Owner", row["owner"]),
        ("Status", row["status"]),
        ("Created", row["created_on"]),
    ]
    if row.get("due_on"):
        pairs.append(("Due", row["due_on"]))

    table = doc.add_table(rows=len(pairs), cols=2)
    table.style = "Table Grid"
    for idx, (label, value) in enumerate(pairs):
        table.cell(idx, 0).text = label
        table.cell(idx, 1).text = value
        for run in table.cell(idx, 0).paragraphs[0].runs:
            run.bold = True
    doc.add_paragraph()


def parse_metrics(value: str) -> list[tuple[str, str]]:
    pairs: list[tuple[str, str]] = []
    for item in filter(None, (part.strip() for part in value.split(";"))):
        if "=" not in item:
            pairs.append((item, ""))
        else:
            label, metric = item.split("=", 1)
            pairs.append((label.strip(), metric.strip()))
    return pairs


def parse_items(value: str, delimiter: str) -> list[str]:
    return [part.strip() for part in value.split(delimiter) if part.strip()]


def validate_row(row: dict[str, str]) -> list[str]:
    errors = [f"Missing required field: {field}" for field in REQUIRED_FIELDS if not row.get(field, "").strip()]
    for field in ("created_on", "due_on"):
        raw = row.get(field, "").strip()
        if raw:
            try:
                datetime.strptime(raw, "%Y-%m-%d")
            except ValueError:
                errors.append(f"Invalid date in {field}: expected YYYY-MM-DD")
    return errors


def build_document(row: dict[str, str], brand: Brand, output_path: Path) -> None:
    doc = Document()
    configure_document(doc, brand)

    title = doc.add_paragraph(style="Title")
    title.add_run(row["document_title"])
    subtitle = doc.add_paragraph(style="Toolkit Note")
    subtitle.add_run(brand.tagline)

    add_metadata_table(doc, row)

    doc.add_heading("Summary", level=1)
    doc.add_paragraph(row["summary"])

    metrics = parse_metrics(row.get("metrics", ""))
    if metrics:
        doc.add_heading("Key figures", level=1)
        table = doc.add_table(rows=len(metrics) + 1, cols=2)
        table.style = "Table Grid"
        table.cell(0, 0).text = "Metric"
        table.cell(0, 1).text = "Value"
        for col in range(2):
            for run in table.cell(0, col).paragraphs[0].runs:
                run.bold = True
        for index, (label, value) in enumerate(metrics, start=1):
            table.cell(index, 0).text = label
            table.cell(index, 1).text = value

    risks = parse_items(row.get("risks", ""), ";")
    if risks:
        doc.add_heading("Risks / open points", level=1)
        for risk in risks:
            doc.add_paragraph(risk, style="List Bullet")

    next_steps = parse_items(row.get("next_steps", ""), "|")
    if next_steps:
        doc.add_heading("Next steps", level=1)
        for item in next_steps:
            doc.add_paragraph(item, style="List Number")

    note = doc.add_paragraph(style="Toolkit Note")
    note.add_run("Generated from structured input. Optional empty sections are omitted automatically.")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)


def generate(input_csv: Path, brand_json: Path, output_dir: Path) -> tuple[list[dict[str, str]], list[dict[str, str]]]:
    brand = load_brand(brand_json)
    generated: list[dict[str, str]] = []
    rejected: list[dict[str, str]] = []

    with input_csv.open(newline="", encoding="utf-8") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            row = {key: (value or "").strip() for key, value in row.items()}
            errors = validate_row(row)
            if errors:
                rejected.append({"record_id": row.get("record_id", ""), "errors": " | ".join(errors)})
                continue

            filename = f"{slug(row['record_id'])}_{slug(row['document_title'])}.docx"
            output_path = output_dir / filename
            build_document(row, brand, output_path)
            generated.append({"record_id": row["record_id"], "file": filename, "status": "generated"})

    output_dir.mkdir(parents=True, exist_ok=True)
    with (output_dir / "generation_manifest.csv").open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=["record_id", "file", "status"])
        writer.writeheader()
        writer.writerows(generated)

    with (output_dir / "validation_report.csv").open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=["record_id", "errors"])
        writer.writeheader()
        writer.writerows(rejected)

    return generated, rejected


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate reusable DOCX summaries from structured CSV data.")
    parser.add_argument("--input", type=Path, default=Path("examples/records.csv"))
    parser.add_argument("--brand", type=Path, default=Path("examples/brand.json"))
    parser.add_argument("--output", type=Path, default=Path("examples/generated"))
    args = parser.parse_args()

    generated, rejected = generate(args.input, args.brand, args.output)
    print(f"Generated {len(generated)} document(s); rejected {len(rejected)} row(s).")


if __name__ == "__main__":
    main()
