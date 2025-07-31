# Contributor Guide

## Overview
- Python automation scripts live in `scripts/`.
- Excel workbook `Finmodel.xlsm` is stored in `excel/`.
- Unit tests live in `tests/`.

## Setup
1. Create a virtual environment with Python 3.11.
   ```bash
   python -m venv venv
   source venv/bin/activate  # Windows: venv\Scripts\activate
   ```
2. Install dependencies.
   ```bash
   pip install -r requirements.txt
   ```
   On Linux some Windows-only packages may fail (e.g. `pywin32`), so run tests on Windows if possible.

## Validation
- Run the linter and tests before committing.
   ```bash
   ruff check .
   pytest -q
   ```

## PR instructions
- Title format: `[Finmodel] <short summary>`
- Describe any new scripts or tests in the body.
