## 🚀 Quick Start (For Reviewers)

### 1. Install Dependencies
The project uses standard Python packages (`pymupdf`, `openpyxl`, `tqdm`, `pytest`):

```bash
pip install -r requirements.txt
```

---

### 2. Run the Assignment Extraction (Pages 10–14)
To reproduce the completed assignment artifact ([`Output_completed.xlsx`](Output_completed.xlsx)), continuing from page 10 and preserving reference rows `BAPC-01` and `BAPC-02`:

```bash
python3 extract_abstracts.py --start-page 10 --end-page 14 --keep-existing
```

---

### 3. Run Custom Page Ranges (Optional)
To run a clean extraction on any custom range without template seed rows:

```bash
python3 extract_abstracts.py --start-page 775 --end-page 785 --output Output_pages_775_785.xlsx
```

---

### 4. Run the Automated Unit Tests

```bash
pytest -v
```

---

## ⚙️ CLI Reference Table

| Argument | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `--pdf` | `str` | `ESAIC2025_Abstracts-1.pdf` | Path to source conference PDF |
| `--template` | `str` | `Output_example.xlsx` | Reference template (used when `--keep-existing` is set) |
| `--output` | `str` | `Output_completed.xlsx` | Destination Excel file path |
| `--start-page` | `int` | `10` | 1-based start page index |
| `--end-page` | `int` | `14` | 1-based end page index |
| `--keep-existing` | `flag` | `False` | Includes seed template rows (`BAPC-01`, `BAPC-02`) |
| `--debug` | `flag` | `False` | Enables verbose debug logging |

---
