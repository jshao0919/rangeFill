# rangeFill

Convert scanned Rifle/Carbine Table VI roll-up PDFs directly to Excel. 
`pdf_to_xlsx.py` writes the same one-lane-per-row layout as `rangeCardData.xlsx`: metadata in A–G, forty hits in H–AU, calculated hits in AV, independently read printed total in AW, and their difference in AX. This layout is accepted by the supplied `xlsx_toForm.py`. All printed lanes are retained, including zero-filled lanes. Handwritten names, ranks, and DoD IDs are left blank for entry in D–F.

## Install

Use Python 3.11 or newer (tested with Python 3.12) and the Tesseract OCR executable (tested with Tesseract 5).

```bash
# Ubuntu/Debian, if Tesseract is not installed:
sudo apt-get install tesseract-ocr
# macOS alternative: brew install tesseract

python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

On Windows, install Tesseract and add its installation directory to PATH; activate the environment with `.venv\Scripts\activate`.

## Convert

```bash
python pdf_to_xlsx.py range24.pdf -o range24.xlsx

# All PDF files in a directory, or multiple PDFs; every page is processed:
python pdf_to_xlsx.py scans/ -o range_scores.xlsx --debug-dir outputs/debug
python pdf_to_xlsx.py first.pdf second.pdf -o combined.xlsx
```

Existing output files are never overwritten. Choose a new output filename for a repeat run. `--workers 1` reduces concurrent OCR work; the default is four. `--width` controls PDF rasterization, with 3000 pixels as the tested default. Enlarging a blurry original does not recover missing detail.

The optional `--template existing.xlsx` appends records to a new copy of a compatible workbook and preserves the original. Its target headers must match the distances read from the scans. The uploaded `rangeCardData.xlsx` has different distance labels for some targets (for example target 23), so use a fresh output workbook for that file rather than appending under incorrect headers.

## Accuracy and review

The converter finds printed rules, corrects page rotation, follows local folds and curved rules, and normalizes row spacing and horizontal perspective. It identifies lane numbers from their printed labels, without assuming every consecutive lane is present. Printed scores are recognized from the 0/1 glyphs; difficult cells are compared with clear glyphs from the same page. Two glyph regions must independently prefer the same digit by sufficient margins. Neither the printed total nor the desired score is used to choose or change a hit.

`Scan Review` contains each lane's status, source PDF/page, issues, and images of its printed label and total. Unknown hits stay blank, with a highlighted cell and comment. Calculated hits and the difference stay blank until all forty hits are available. A readable total that disagrees with the extracted hits is highlighted for review; it does not cause the hits to be adjusted. Status and totals describe the extraction run and are stored as values, so manual edits require rechecking those values.

Distances are read in printed row order. `Printed Target Rows` records the actual distances for each PDF page. When distances differ between pages, the combined workbook uses generic numbered target headers. If a distance cannot be read, it remains blank and is flagged; hits and totals are still retained.

Exit codes:

- **0:** all lanes, hits, printed totals, and target distances passed the extraction checks.
- **2:** the workbook was written, with items requiring review.
- **1:** invalid input/arguments, no successfully converted pages, or the output could not be written. Page failures are recorded in the review sheet; successful pages from a mixed batch are retained with exit code 2.



## Tests

```bash
python -m pip install pytest==9.1.1
python -m pytest -q

# Include the real scan and compatibility checks when those files are available:
RANGEFILL_SAMPLE_PDF=/path/to/range24.pdf \
RANGEFILL_FORM_READER=/path/to/xlsx_toForm.py \
python -m pytest -q
```

