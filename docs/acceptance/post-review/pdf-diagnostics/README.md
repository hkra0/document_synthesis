# N9 PDF parser diagnostics

This directory contains diagnostics derived from the immutable R0–R11 review
PDFs. The original PDFs are not rewritten or replaced. Each report records the
input SHA-256, parser versions, parser stderr, page geometry values, page-level
text fingerprints from pypdf/PyMuPDF/Poppler, page labels, and bookmarks so a
parser warning cannot be mistaken for a successful Word/PDF verification.

The reports are diagnostic evidence only. The same anonymous DOCX has now been
re-exported from the production CLI in the permitted Word automation context.
For that rerun and the two historical PDFs, the observed warnings are
classified as the known non-fatal `quartz_orphan_zero_offset_xref` pattern:
the offset-zero objects are absent and unreferenced, while pypdf, PyMuPDF and
Poppler agree on page count and page geometry. Raw stderr and qpdf diagnostics
remain preserved. Any other warning pattern, referenced missing object, or
parser/page-geometry disagreement remains `blocked`.
