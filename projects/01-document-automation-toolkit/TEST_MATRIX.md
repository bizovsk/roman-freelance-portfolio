# Validation matrix

This demonstration is tested as a **data-to-document workflow**, not merely as a set of static Word files.

| Check | Expected result | Status |
|---|---|---|
| Valid batch generation | 3 valid input rows create 3 DOCX files | Passed |
| Invalid input isolation | Missing required fields are rejected and reported | Passed |
| Date validation | Non-ISO dates are flagged | Passed |
| Optional content | Empty optional sections are omitted | Passed |
| Structured metrics | CSV key/value metrics become native Word tables | Passed |
| Reusable structure | Native paragraphs, styles, tables and lists | Passed |
| No floating text boxes | Core content remains editable and reflows | Passed |
| Manifest | Successful outputs are listed separately | Passed |
| Visual rendering | Every generated DOCX renders cleanly | Passed |
| Automated tests | 6 tests pass | Passed |

## Scope note

All example data, names and organisations are synthetic. The demo is deliberately generic and is not based on a specific client assignment.
