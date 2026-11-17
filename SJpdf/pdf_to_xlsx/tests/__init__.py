"""Tests for FIS PDF-to-Excel conversion."""

from pdf_to_xlsx import pdf2xlsx

# The tests do not depend on the user's SJ-Listen.xlsm; tests of the lookup
# pass their own list file.
pdf2xlsx.SJ_LISTEN_FILE = None
