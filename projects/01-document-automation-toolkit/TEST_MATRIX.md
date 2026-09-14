# Validation matrix

This project is validated as a **data-to-document workflow**, not just a collection of static files.

| Check | Expected result | Status |
|---|---|---|
| CSV input | Valid CSV records are read and normalized | Passed |
| XLSX input | Typed Excel records from an explicit worksheet are read and normalized | Passed |
| Missing columns | Missing required input columns fail fast | Passed |
| Valid batch generation | 3 valid source rows generate deliverables | Passed |
| Invalid input isolation | Invalid rows are rejected and reported | Passed |
| Duplicate identity | Duplicate `record_id` values are rejected | Passed |
| Date validation | Non-ISO dates are flagged | Passed |
| Date ordering | `due_on < created_on` is rejected | Passed |
| Metrics validation | Metrics must use `Label=Value` structure | Passed |
| Optional content | Empty optional sections are omitted | Passed |
| DOCX structure | Native paragraphs, styles, tables and lists | Passed |
| No floating text boxes | Core Word content remains editable and reflows | Passed |
| PDF generation | Matching PDF output is produced | Passed |
| Manifest | Successful outputs are listed separately | Passed |
| Rejection report | Rejected records include source row and reason | Passed |
| Run summary | Machine-readable JSON summary is created | Passed |
| Visual DOCX QA | Generated DOCX examples render cleanly | Passed |
| Visual PDF QA | Generated PDF examples render cleanly | Passed |
| Automated tests | 12 tests pass | Passed |

## Scope note

All example data, names and organisations are synthetic. The demo is deliberately generic and is not based on a specific client assignment.
