from pathlib import Path
from zipfile import ZipFile

from docx import Document

from src.generate_documents import generate, parse_metrics, validate_row

ROOT = Path(__file__).resolve().parents[1]


def test_demo_generation(tmp_path):
    generated, rejected = generate(ROOT / "examples/records.csv", ROOT / "examples/brand.json", tmp_path)
    assert len(generated) == 3
    assert len(rejected) == 1
    assert (tmp_path / "generation_manifest.csv").exists()
    assert (tmp_path / "validation_report.csv").exists()


def test_generated_documents_are_native_docx(tmp_path):
    generated, _ = generate(ROOT / "examples/records.csv", ROOT / "examples/brand.json", tmp_path)
    for item in generated:
        path = tmp_path / item["file"]
        doc = Document(path)
        assert doc.paragraphs
        assert doc.tables
        with ZipFile(path) as archive:
            document_xml = archive.read("word/document.xml").decode("utf-8")
            assert "w:txbxContent" not in document_xml


def test_optional_risk_section_is_omitted(tmp_path):
    generated, _ = generate(ROOT / "examples/records.csv", ROOT / "examples/brand.json", tmp_path)
    target = next(tmp_path / item["file"] for item in generated if item["record_id"] == "PRJ-003")
    text = "\n".join(paragraph.text for paragraph in Document(target).paragraphs)
    assert "Risks / open points" not in text
    assert "Next steps" in text


def test_missing_required_field_is_rejected():
    row = {
        "record_id": "X",
        "document_title": "",
        "recipient": "Team",
        "owner": "Owner",
        "status": "Draft",
        "created_on": "2026-09-10",
        "summary": "Summary",
        "due_on": "",
    }
    errors = validate_row(row)
    assert "Missing required field: document_title" in errors


def test_invalid_date_is_reported():
    row = {
        "record_id": "X",
        "document_title": "Title",
        "recipient": "Team",
        "owner": "Owner",
        "status": "Draft",
        "created_on": "10/09/2026",
        "summary": "Summary",
        "due_on": "",
    }
    assert any("Invalid date" in error for error in validate_row(row))


def test_metrics_parser():
    assert parse_metrics("Rows=20;Errors=2") == [("Rows", "20"), ("Errors", "2")]
